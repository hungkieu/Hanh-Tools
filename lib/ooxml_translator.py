from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from lxml import etree

from lib.openai_translator import Translator


class CancelledError(Exception):
    """Người dùng đã yêu cầu dừng."""


def _noop_cancel() -> bool:
    return False


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}
XML_SPACE_ATTR = f"{{{XML_NS}}}space"
CONTENT_XML_PATTERNS = (
    "word/document.xml",
    "word/header*.xml",
    "word/footer*.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/comments.xml",
)
MARKER_RE = re.compile(r"\[\[\[(\d+)]]]([\s\S]*?)\[\[\[/\1]]]")
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
    nodes: list[etree._Element]
    index: int
    text: str

    @property
    def node(self) -> etree._Element:
        return self.nodes[0]


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
    cancel_check: Callable[[], bool] = _noop_cancel,
) -> TranslationStats:
    root = Path(docx_folder).expanduser().resolve()
    if not root.exists():
        raise FileNotFoundError(f"Không tìm thấy folder docx đã unzip: {root}")

    stats = TranslationStats()
    for xml_path in _iter_content_xml_files(root):
        if cancel_check():
            raise CancelledError("Đã dừng bởi người dùng")
        stats.files_scanned += 1
        file_stats = _translate_xml_file(
            xml_path,
            root=root,
            translator=translator,
            batch_size=batch_size,
            retries=retries,
            verbose=verbose,
            min_batch_size=min_batch_size,
            cancel_check=cancel_check,
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
    cancel_check: Callable[[], bool] = _noop_cancel,
) -> TranslationStats:
    parser = etree.XMLParser(remove_blank_text=False, recover=False)
    tree = etree.parse(str(xml_path), parser)
    units = _extract_paragraph_units(tree, xml_path.relative_to(root).as_posix())
    stats = TranslationStats(paragraphs_found=len(units))

    translatable: list[ParagraphUnit] = []
    for unit in units:
        reason = _skip_reason(unit.text)
        if reason is None:
            translatable.append(unit)
        else:
            stats.paragraphs_skipped += 1
            if verbose:
                plain = MARKER_RE.sub(lambda m: m.group(2), unit.text)
                print(f"  [SKIP {reason}] {unit.id} text={_truncate(plain, 80)!r}", flush=True)

    changed = False
    batches = _chunks(translatable, batch_size)
    for batch_index, batch in enumerate(batches, start=1):
        if cancel_check():
            if changed:
                if verbose:
                    print(f"Cancelled; writing partial translation: {xml_path.relative_to(root)}", flush=True)
                tree.write(str(xml_path), encoding="UTF-8", xml_declaration=True, standalone=None)
            raise CancelledError("Đã dừng bởi người dùng")
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
                if verbose:
                    print(
                        f"  [FAIL no-response] {unit.id} "
                        f"src={_truncate(unit.text)}",
                        flush=True,
                    )
                continue
            reason = _validate_markers(translated, len(unit.slots))
            if reason is None:
                translations[unit.id] = translated
            else:
                failed.append(unit)
                if verbose:
                    print(
                        f"  [FAIL {reason}] {unit.id}\n"
                        f"    src: {_truncate(unit.text)}\n"
                        f"    out: {_truncate(translated)}",
                        flush=True,
                    )

        if not failed:
            break
        pending = failed
        if attempt == retries and verbose:
            print(
                f"Skipped {len(failed)} paragraphs after {retries + 1} attempts: "
                f"{', '.join(unit.id for unit in failed)}",
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
        raw_slots = _collect_text_nodes(paragraph)
        slots = _merge_adjacent_slots(raw_slots)
        if not slots:
            continue
        parts = [f"[[[{slot.index}]]]{slot.text}[[[/{slot.index}]]]" for slot in slots]
        units.append(
            ParagraphUnit(
                id=f"{file_id}:p:{paragraph_index}",
                text="".join(parts),
                slots=slots,
            )
        )
    return units


def _collect_text_nodes(paragraph: etree._Element) -> list[tuple[etree._Element, etree._Element]]:
    """Trả về [(run, w:t)] giữ nguyên thứ tự xuất hiện trong paragraph."""
    pairs: list[tuple[etree._Element, etree._Element]] = []
    for run in paragraph.iter(f"{{{W_NS}}}r"):
        for text_node in run.findall(f"{{{W_NS}}}t"):
            if (text_node.text or "") != "":
                pairs.append((run, text_node))
    return pairs


def _merge_adjacent_slots(
    pairs: list[tuple[etree._Element, etree._Element]],
) -> list[TextSlot]:
    """Gộp các w:t liền kề có cùng run-property (bỏ qua hint=eastAsia)."""
    slots: list[TextSlot] = []
    current_nodes: list[etree._Element] = []
    current_text: list[str] = []
    current_key: str | None = None
    prev_run: etree._Element | None = None

    def _flush() -> None:
        if current_nodes:
            slots.append(TextSlot(
                nodes=list(current_nodes),
                index=len(slots),
                text="".join(current_text),
            ))

    for run, text_node in pairs:
        key = _run_format_key(run)
        adjacent = (
            prev_run is not None
            and key == current_key
            and _runs_are_adjacent_siblings(prev_run, run)
        )
        if not adjacent:
            _flush()
            current_nodes = []
            current_text = []
            current_key = key
        current_nodes.append(text_node)
        current_text.append(text_node.text or "")
        prev_run = run
    _flush()
    return slots


def _run_format_key(run: etree._Element) -> str:
    rpr = run.find(f"{{{W_NS}}}rPr")
    if rpr is None:
        return ""
    clone = etree.fromstring(etree.tostring(rpr))
    # Bỏ <w:rFonts> chỉ chứa hint (không thay đổi hình thức)
    for rfonts in list(clone.findall(f"{{{W_NS}}}rFonts")):
        attrs = {k for k in rfonts.attrib.keys() if k != f"{{{XML_NS}}}space"}
        if attrs == {f"{{{W_NS}}}hint"} or not attrs:
            clone.remove(rfonts)
    # rPr rỗng sau khi normalize ≡ không có rPr
    if len(clone) == 0 and not clone.attrib:
        return ""
    return etree.tostring(clone, method="c14n").decode("utf-8")


def _runs_are_adjacent_siblings(a: etree._Element, b: etree._Element) -> bool:
    """True nếu giữa 2 run không có element nào khác (bookmark, field, ...)."""
    sibling = a.getnext()
    return sibling is b


def _should_translate(text: str) -> bool:
    return _skip_reason(text) is None


def _skip_reason(text: str) -> str | None:
    """None nếu nên dịch, ngược lại trả về lý do skip để log."""
    plain = MARKER_RE.sub(lambda match: match.group(2), text).strip()
    if not plain:
        return "empty"
    if not _has_letter(plain):
        return "no letters"
    if plain.startswith("http://") or plain.startswith("https://"):
        return "url"
    return None


def _has_letter(text: str) -> bool:
    """True nếu có ít nhất 1 ký tự thuộc category Unicode 'Letter' (L*)."""
    return any(unicodedata.category(ch).startswith("L") for ch in text)


def _translation_matches_slots(text: str, slot_count: int) -> bool:
    return _validate_markers(text, slot_count) is None


def _validate_markers(text: str, slot_count: int) -> str | None:
    """Trả về None nếu hợp lệ, hoặc chuỗi mô tả lỗi để debug."""
    matches = list(MARKER_RE.finditer(text))
    if len(matches) != slot_count:
        return (
            f"marker count mismatch: expected {slot_count}, got {len(matches)}"
        )
    ids = [int(match.group(1)) for match in matches]
    if ids != list(range(slot_count)):
        return f"marker ids out of order: expected {list(range(slot_count))}, got {ids}"
    return None


def _truncate(text: str, limit: int = 200) -> str:
    text = text.replace("\n", "\\n")
    if len(text) <= limit:
        return text
    return text[:limit] + f"... (+{len(text) - limit} chars)"


def _apply_translation(unit: ParagraphUnit, translated: str, *, verbose: bool = False) -> None:
    matches = list(MARKER_RE.finditer(translated))
    for slot, match in zip(unit.slots, matches, strict=True):
        raw = match.group(2)
        text = _sanitize_xml_text(raw)
        text = _restore_edge_whitespace(source=slot.text, translated=text)
        if verbose and text != raw:
            print(
                f"Adjusted text for {unit.id} slot {slot.index} "
                f"(sanitized/edge-whitespace restored)",
                flush=True,
            )
        # Ghi toàn bộ text dịch vào node đầu, xoá text các node còn lại của slot
        first = slot.nodes[0]
        first.text = text
        if _needs_preserve_space(text):
            first.set(XML_SPACE_ATTR, "preserve")
        for extra in slot.nodes[1:]:
            extra.text = ""


def _restore_edge_whitespace(*, source: str, translated: str) -> str:
    """Nếu model lỡ trim space đầu/cuối slot mà nguồn có, phục hồi để tránh dính chữ."""
    if not translated:
        return translated
    leading = _leading_space(source)
    if leading and not translated[0].isspace():
        translated = leading + translated
    trailing = _trailing_space(source)
    if trailing and not translated[-1].isspace():
        translated = translated + trailing
    return translated


def _leading_space(text: str) -> str:
    i = 0
    while i < len(text) and text[i].isspace():
        i += 1
    return text[:i]


def _trailing_space(text: str) -> str:
    i = len(text)
    while i > 0 and text[i - 1].isspace():
        i -= 1
    return text[i:]


def _needs_preserve_space(text: str) -> bool:
    return bool(text) and (text[0].isspace() or text[-1].isspace())


def _sanitize_xml_text(text: str) -> str:
    return INVALID_XML_CHAR_RE.sub("", text)


def _chunks(items: list[ParagraphUnit], size: int) -> list[list[ParagraphUnit]]:
    return [items[index : index + size] for index in range(0, len(items), size)]
