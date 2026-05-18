from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

from lib.openai_translator import Translator


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
CONTENT_XML_PATTERNS = (
    "word/document.xml",
    "word/header*.xml",
    "word/footer*.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/comments.xml",
)
MARKER_RE = re.compile(r"\[\[\[(\d+)]]]([\s\S]*?)\[\[\[/\1]]]")
WORD_RE = re.compile(r"[A-Za-zÀ-ỹ]")
INVALID_XML_CHAR_RE = re.compile(
    "[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]"
)


@dataclass
class TranslationStats:
    files_scanned: int = 0
    paragraphs_found: int = 0
    paragraphs_translated: int = 0
    paragraphs_skipped: int = 0
    paragraphs_failed: int = 0


@dataclass
class TextSlot:
    node: etree._Element
    index: int
    text: str


@dataclass
class ParagraphUnit:
    id: str
    text: str
    slots: list[TextSlot]


def translate_docx_xml_folder(
    docx_folder: str | Path,
    translator: Translator,
    *,
    batch_size: int = 25,
    retries: int = 2,
    verbose: bool = False,
    min_batch_size: int = 1,
) -> TranslationStats:
    root = Path(docx_folder).expanduser().resolve()
    if not root.exists():
        raise FileNotFoundError(f"Không tìm thấy folder docx đã unzip: {root}")

    stats = TranslationStats()
    for xml_path in _iter_content_xml_files(root):
        stats.files_scanned += 1
        file_stats = _translate_xml_file(
            xml_path,
            root=root,
            translator=translator,
            batch_size=batch_size,
            retries=retries,
            verbose=verbose,
            min_batch_size=min_batch_size,
        )
        stats.paragraphs_found += file_stats.paragraphs_found
        stats.paragraphs_translated += file_stats.paragraphs_translated
        stats.paragraphs_skipped += file_stats.paragraphs_skipped
        stats.paragraphs_failed += file_stats.paragraphs_failed

    return stats


def _translate_xml_file(
    xml_path: Path,
    *,
    root: Path,
    translator: Translator,
    batch_size: int,
    retries: int,
    verbose: bool,
    min_batch_size: int,
) -> TranslationStats:
    parser = etree.XMLParser(remove_blank_text=False, recover=False)
    tree = etree.parse(str(xml_path), parser)
    units = _extract_paragraph_units(tree, xml_path.relative_to(root).as_posix())
    stats = TranslationStats(paragraphs_found=len(units))

    translatable = [unit for unit in units if _should_translate(unit.text)]
    stats.paragraphs_skipped += len(units) - len(translatable)

    changed = False
    batches = _chunks(translatable, batch_size)
    for batch_index, batch in enumerate(batches, start=1):
        if verbose:
            char_count = sum(len(unit.text) for unit in batch)
            marker_count = sum(len(unit.slots) for unit in batch)
            print(
                f"Translating {xml_path.relative_to(root)} "
                f"batch {batch_index}/{len(batches)} "
                f"({len(batch)} paragraphs, {char_count} chars, {marker_count} markers)",
                flush=True,
            )
        translations = _translate_units_with_split_retry(
            batch,
            translator=translator,
            retries=retries,
            min_batch_size=min_batch_size,
            verbose=verbose,
            label=f"{xml_path.relative_to(root)} batch {batch_index}",
        )
        failed_count = len(batch) - len(translations)
        stats.paragraphs_failed += failed_count

        for unit in batch:
            translated = translations.get(unit.id)
            if not translated:
                continue
            _apply_translation(unit, translated, verbose=verbose)
            stats.paragraphs_translated += 1
            changed = True

    if changed:
        if verbose:
            print(f"Writing translated XML: {xml_path.relative_to(root)}", flush=True)
        tree.write(
            str(xml_path),
            encoding="UTF-8",
            xml_declaration=True,
            standalone=None,
        )

    return stats


def _translate_units_with_split_retry(
    units: list[ParagraphUnit],
    *,
    translator: Translator,
    retries: int,
    min_batch_size: int,
    verbose: bool,
    label: str,
) -> dict[str, str]:
    if not units:
        return {}

    try:
        return _translate_units_with_marker_retry(
            units,
            translator=translator,
            retries=retries,
            verbose=verbose,
            label=label,
        )
    except TimeoutError:
        if len(units) <= min_batch_size:
            if verbose:
                print(f"OpenAI timeout, skipped smallest batch: {label}", flush=True)
            return {}

        midpoint = len(units) // 2
        if verbose:
            char_count = sum(len(unit.text) for unit in units)
            print(
                f"OpenAI timeout, splitting {label}: "
                f"{len(units)} paragraphs / {char_count} chars -> "
                f"{midpoint} + {len(units) - midpoint}",
                flush=True,
            )

        left = _translate_units_with_split_retry(
            units[:midpoint],
            translator=translator,
            retries=retries,
            min_batch_size=min_batch_size,
            verbose=verbose,
            label=f"{label}.1",
        )
        right = _translate_units_with_split_retry(
            units[midpoint:],
            translator=translator,
            retries=retries,
            min_batch_size=min_batch_size,
            verbose=verbose,
            label=f"{label}.2",
        )
        return left | right


def _translate_units_with_marker_retry(
    units: list[ParagraphUnit],
    *,
    translator: Translator,
    retries: int,
    verbose: bool,
    label: str,
) -> dict[str, str]:
    pending = units
    translations: dict[str, str] = {}
    for attempt in range(retries + 1):
        if verbose and attempt > 0:
            print(
                f"Retrying marker validation failures: {label}, "
                f"attempt {attempt + 1}/{retries + 1}, {len(pending)} paragraphs",
                flush=True,
            )

        raw = translator.translate_batch(
            [{"id": unit.id, "text": unit.text} for unit in pending]
        )
        failed: list[ParagraphUnit] = []
        for unit in pending:
            translated = raw.get(unit.id)
            if translated is None:
                failed.append(unit)
                continue
            if _translation_matches_slots(translated, len(unit.slots)):
                translations[unit.id] = translated
            else:
                failed.append(unit)

        if not failed:
            break
        pending = failed
        if attempt == retries and verbose:
            failed_ids = ", ".join(unit.id for unit in failed[:5])
            suffix = "..." if len(failed) > 5 else ""
            print(
                f"Skipped paragraphs after marker validation failed: "
                f"{len(failed)} ({failed_ids}{suffix})",
                flush=True,
            )

    return translations


def _iter_content_xml_files(root: Path) -> list[Path]:
    files: list[Path] = []
    for pattern in CONTENT_XML_PATTERNS:
        files.extend(root.glob(pattern))
    return sorted(path for path in files if path.is_file())


def _extract_paragraph_units(tree: etree._ElementTree, file_id: str) -> list[ParagraphUnit]:
    units: list[ParagraphUnit] = []
    for paragraph_index, paragraph in enumerate(tree.xpath("//w:p", namespaces=NS)):
        slots: list[TextSlot] = []
        parts: list[str] = []
        for text_index, text_node in enumerate(paragraph.xpath(".//w:t", namespaces=NS)):
            text = text_node.text or ""
            if text == "":
                continue
            slot = TextSlot(node=text_node, index=len(slots), text=text)
            slots.append(slot)
            parts.append(f"[[[{slot.index}]]]{text}[[[/{slot.index}]]]")

        if slots:
            units.append(
                ParagraphUnit(
                    id=f"{file_id}:p:{paragraph_index}",
                    text="".join(parts),
                    slots=slots,
                )
            )

    return units


def _should_translate(text: str) -> bool:
    plain = MARKER_RE.sub(lambda match: match.group(2), text).strip()
    if not plain:
        return False
    if not WORD_RE.search(plain):
        return False
    if plain.startswith("http://") or plain.startswith("https://"):
        return False
    return True


def _translation_matches_slots(text: str, slot_count: int) -> bool:
    matches = list(MARKER_RE.finditer(text))
    if len(matches) != slot_count:
        return False
    return [int(match.group(1)) for match in matches] == list(range(slot_count))


def _apply_translation(unit: ParagraphUnit, translated: str, *, verbose: bool = False) -> None:
    matches = list(MARKER_RE.finditer(translated))
    for slot, match in zip(unit.slots, matches, strict=True):
        text = _sanitize_xml_text(match.group(2))
        if verbose and text != match.group(2):
            print(
                f"Removed XML-incompatible control characters from {unit.id} "
                f"slot {slot.index}",
                flush=True,
            )
        slot.node.text = text


def _sanitize_xml_text(text: str) -> str:
    return INVALID_XML_CHAR_RE.sub("", text)


def _chunks(items: list[ParagraphUnit], size: int) -> list[list[ParagraphUnit]]:
    return [items[index : index + size] for index in range(0, len(items), size)]
