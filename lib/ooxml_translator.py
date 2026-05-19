from __future__ import annotations

import re
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from lxml import etree

from lib.openai_translator import BatchTooLargeError, Translator
from lib.translation_cache import TranslationCache


class CancelledError(Exception):
    """Người dùng đã yêu cầu dừng."""


def _noop_cancel() -> bool:
    return False


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}
XML_SPACE_ATTR = f"{{{XML_NS}}}space"
DEFAULT_CONTENT_XML_PATTERNS = ("word/document.xml",)
EXTRA_CONTENT_XML_PATTERNS = (
    "word/header*.xml",
    "word/footer*.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/comments.xml",
)
# Marker ngắn để tiết kiệm token. Dùng angle bracket toán học (U+2329/U+232A),
# rất hiếm khi xuất hiện trong text thường.
MARKER_RE = re.compile(r"〈(\d+)〉([\s\S]*?)〈/\1〉")

# Lenient pattern dùng để "sửa" marker mà model viết hơi lệch (thừa space, thiếu /).
_MARKER_REPAIR_OPEN_RE = re.compile(r"〈\s*(\d+)\s*〉")
_MARKER_REPAIR_CLOSE_RE = re.compile(r"〈\s*/\s*(\d+)\s*〉")
# Bắt cả trường hợp model dùng nhầm bracket ASCII < > vì model hay "phiên dịch"
# angle bracket toán học sang HTML-like.
_MARKER_REPAIR_OPEN_ASCII_RE = re.compile(r"<\s*(\d+)\s*>")
_MARKER_REPAIR_CLOSE_ASCII_RE = re.compile(r"<\s*/\s*(\d+)\s*>")


def _repair_markers(text: str) -> str:
    """Sửa các marker bị viết lệch nhẹ (whitespace, ASCII <> thay vì 〈〉)."""
    # Close trước để dấu / không bị nuốt bởi pattern open.
    text = _MARKER_REPAIR_CLOSE_RE.sub(lambda m: f"〈/{m.group(1)}〉", text)
    text = _MARKER_REPAIR_OPEN_RE.sub(lambda m: f"〈{m.group(1)}〉", text)
    text = _MARKER_REPAIR_CLOSE_ASCII_RE.sub(lambda m: f"〈/{m.group(1)}〉", text)
    text = _MARKER_REPAIR_OPEN_ASCII_RE.sub(lambda m: f"〈{m.group(1)}〉", text)
    text = _repair_missing_close(text)
    return text


_ANY_MARKER_RE = re.compile(r"〈(/?)(\d+)〉")


def _repair_missing_close(text: str) -> str:
    """Nếu marker cuối cùng trong text là OPEN (model quên đóng), tự append close.

    Chỉ áp dụng khi đúng 1 close bị thiếu và open chưa đóng đó nằm ở vị trí cuối —
    tránh trường hợp thiếu close ở giữa (có thể hỏng thêm).
    """
    markers = list(_ANY_MARKER_RE.finditer(text))
    if not markers:
        return text
    last = markers[-1]
    if last.group(1) == "/":  # marker cuối là close → không phải case này
        return text
    open_ids = [int(m.group(2)) for m in markers if not m.group(1)]
    close_ids = [int(m.group(2)) for m in markers if m.group(1)]
    if len(open_ids) - len(close_ids) != 1:
        return text
    open_count: dict[int, int] = {}
    for i in open_ids:
        open_count[i] = open_count.get(i, 0) + 1
    for i in close_ids:
        open_count[i] = open_count.get(i, 0) - 1
    missing = [i for i, c in open_count.items() if c > 0]
    if len(missing) == 1 and open_count[missing[0]] == 1 and missing[0] == int(last.group(2)):
        return text + f"〈/{missing[0]}〉"
    return text


def _open_marker(idx: int) -> str:
    return f"〈{idx}〉"


def _close_marker(idx: int) -> str:
    return f"〈/{idx}〉"
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
    paragraphs_fallback: int = 0  # marker fail, dùng fallback distribute vào slot 0
    cache_hits: int = 0
    cache_misses: int = 0


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
    include_extras: bool = False,
    concurrency: int = 1,
    fallback_translators: list[Translator] | None = None,
    cache: TranslationCache | None = None,
    target_language: str = "Vietnamese",
    input_token_budget: int = 2000,
    force_font: str | None = None,
    only_english: bool = True,
) -> TranslationStats:
    root = Path(docx_folder).expanduser().resolve()
    if not root.exists():
        raise FileNotFoundError(f"Không tìm thấy folder docx đã unzip: {root}")

    stats = TranslationStats()
    for xml_path in _iter_content_xml_files(root, include_extras=include_extras):
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
            concurrency=concurrency,
            fallback_translators=fallback_translators or [],
            cache=cache,
            target_language=target_language,
            input_token_budget=input_token_budget,
            force_font=force_font,
            only_english=only_english,
        )
        stats.paragraphs_found += file_stats.paragraphs_found
        stats.paragraphs_translated += file_stats.paragraphs_translated
        stats.paragraphs_skipped += file_stats.paragraphs_skipped
        stats.paragraphs_failed += file_stats.paragraphs_failed
        stats.paragraphs_fallback += file_stats.paragraphs_fallback
        stats.cache_hits += file_stats.cache_hits
        stats.cache_misses += file_stats.cache_misses

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
    concurrency: int = 1,
    fallback_translators: list[Translator] | None = None,
    cache: TranslationCache | None = None,
    target_language: str = "Vietnamese",
    input_token_budget: int = 2000,
    force_font: str | None = None,
    only_english: bool = True,
) -> TranslationStats:
    parser = etree.XMLParser(remove_blank_text=False, recover=False)
    tree = etree.parse(str(xml_path), parser)
    units = _extract_paragraph_units(tree, xml_path.relative_to(root).as_posix())
    stats = TranslationStats(paragraphs_found=len(units))

    translatable: list[ParagraphUnit] = []
    for unit in units:
        reason = _skip_reason(unit.text, only_latin=only_english)
        if reason is None:
            translatable.append(unit)
        else:
            stats.paragraphs_skipped += 1
            if verbose:
                plain = MARKER_RE.sub(lambda m: m.group(2), unit.text)
                print(f"  [SKIP {reason}] {unit.id} text={_truncate(plain, 80)!r}", flush=True)

    # Dedupe: nhiều paragraph trong file có thể có text giống nhau (header bảng,
    # ghi chú lặp). Chỉ gửi unique text lên API, áp dụng kết quả cho mọi bản sao.
    unique_units: dict[str, ParagraphUnit] = {}
    duplicate_groups: dict[str, list[ParagraphUnit]] = {}
    for unit in translatable:
        if unit.text not in unique_units:
            unique_units[unit.text] = unit
            duplicate_groups[unit.id] = [unit]
        else:
            rep_id = unique_units[unit.text].id
            duplicate_groups[rep_id].append(unit)

    dedup_units = list(unique_units.values())
    dedup_savings = len(translatable) - len(dedup_units)
    if verbose and dedup_savings > 0:
        print(
            f"Dedup: {len(translatable)} translatable -> {len(dedup_units)} unique "
            f"(saved {dedup_savings} API translations)",
            flush=True,
        )

    # Cache lookup: tách thành (cached, to_translate)
    changed = False
    cache_hits_units: list[ParagraphUnit] = []
    to_translate: list[ParagraphUnit] = []
    model_name = getattr(translator, "model", "unknown")
    # Mode khác nhau (only_english on/off) cho output khác nhau với đoạn lai → tách bucket.
    cache_lang_key = f"{target_language}|en-only" if only_english else target_language
    if cache is not None:
        for unit in dedup_units:
            cached = cache.get(model_name, cache_lang_key, unit.text)
            if cached is not None and _validate_markers(cached, len(unit.slots)) is None:
                cache_hits_units.append(unit)
                # Apply ngay
                for u in duplicate_groups[unit.id]:
                    _apply_translation(u, cached, verbose=False, force_font=force_font)
                    stats.paragraphs_translated += 1
                stats.cache_hits += 1
                changed = True
            else:
                to_translate.append(unit)
        stats.cache_misses += len(to_translate)
        if verbose and cache_hits_units:
            print(
                f"Cache: {len(cache_hits_units)} hit / {len(dedup_units)} unique "
                f"(saved {len(cache_hits_units)} API calls)",
                flush=True,
            )
    else:
        to_translate = dedup_units

    # Token-budget batching: pack đến khi đạt ngân sách input tokens.
    batches = _token_aware_chunks(
        to_translate, translator,
        input_token_budget=input_token_budget,
        hard_cap=batch_size,
    )

    def _log_batch_start(batch_index: int, batch: list[ParagraphUnit]) -> None:
        if not verbose:
            return
        char_count = sum(len(unit.text) for unit in batch)
        marker_count = sum(len(unit.slots) for unit in batch)
        print(
            f"Translating {xml_path.relative_to(root)} "
            f"batch {batch_index}/{len(batches)} "
            f"({len(batch)} paragraphs, {char_count} chars, {marker_count} markers)",
            flush=True,
        )

    def _do_batch(batch: list[ParagraphUnit], batch_index: int) -> dict[str, tuple[str, bool]]:
        return _translate_units_with_split_retry(
            batch,
            translator=translator,
            retries=retries,
            min_batch_size=min_batch_size,
            verbose=verbose,
            label=f"{xml_path.relative_to(root)} batch {batch_index}",
            fallback_translators=fallback_translators or [],
        )

    def _apply_batch(batch: list[ParagraphUnit], translations: dict[str, tuple[str, bool]]) -> None:
        nonlocal changed
        for rep_unit in batch:
            group = duplicate_groups[rep_unit.id]
            result = translations.get(rep_unit.id)
            if not result:
                stats.paragraphs_failed += len(group)
                continue
            translated, used_fallback = result
            for unit in group:
                _apply_translation(unit, translated, verbose=verbose, force_font=force_font)
                stats.paragraphs_translated += 1
                if used_fallback:
                    stats.paragraphs_fallback += 1
            changed = True
            # Ghi cache (chỉ khi không phải fallback distribute, để không "đầu độc" cache)
            if cache is not None and not used_fallback:
                plain_len = len(_strip_markers(rep_unit.text).strip())
                cache.put(model_name, cache_lang_key, rep_unit.text, translated, plain_len)

    def _check_cancel() -> None:
        if cancel_check():
            if changed:
                if verbose:
                    print(f"Cancelled; writing partial translation: {xml_path.relative_to(root)}", flush=True)
                tree.write(str(xml_path), encoding="UTF-8", xml_declaration=True, standalone=None)
            raise CancelledError("Đã dừng bởi người dùng")

    if concurrency <= 1:
        for batch_index, batch in enumerate(batches, start=1):
            _check_cancel()
            _log_batch_start(batch_index, batch)
            translations = _do_batch(batch, batch_index)
            _apply_batch(batch, translations)
    else:
        # Song song: worker chỉ gọi API; main thread sequential apply vào tree
        # (lxml không thread-safe khi ghi).
        with ThreadPoolExecutor(max_workers=concurrency) as pool:
            future_map: dict = {}
            for batch_index, batch in enumerate(batches, start=1):
                if cancel_check():
                    break
                _log_batch_start(batch_index, batch)
                future_map[pool.submit(_do_batch, batch, batch_index)] = (batch_index, batch)
            try:
                for fut in as_completed(future_map):
                    batch_index, batch = future_map[fut]
                    translations = fut.result()
                    _apply_batch(batch, translations)
                    _check_cancel()
            except CancelledError:
                # Huỷ các future chưa chạy; future đang chạy không huỷ được, đành đợi.
                for f in future_map:
                    f.cancel()
                raise

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
    fallback_translators: list[Translator] | None = None,
) -> dict[str, tuple[str, bool]]:
    if not units:
        return {}

    # Pre-emptive split: nếu ước lượng output vượt max_completion_tokens → chia đôi trước.
    is_too_large = getattr(translator, "is_batch_too_large", None)
    if is_too_large is not None and len(units) > min_batch_size:
        items = [{"id": u.id, "text": u.text} for u in units]
        if is_too_large(items):
            midpoint = max(1, len(units) // 2)
            if verbose:
                est = getattr(translator, "estimate_output_tokens", lambda _: 0)(items)
                print(
                    f"Pre-emptive split (est {est} output tokens > limit) {label}: "
                    f"{len(units)} -> {midpoint} + {len(units) - midpoint}",
                    flush=True,
                )
            left = _translate_units_with_split_retry(
                units[:midpoint], translator=translator, retries=retries,
                min_batch_size=min_batch_size, verbose=verbose, label=f"{label}.1",
                fallback_translators=fallback_translators,
            )
            right = _translate_units_with_split_retry(
                units[midpoint:], translator=translator, retries=retries,
                min_batch_size=min_batch_size, verbose=verbose, label=f"{label}.2",
                fallback_translators=fallback_translators,
            )
            return left | right

    try:
        return _translate_units_with_marker_retry(
            units,
            translator=translator,
            retries=retries,
            verbose=verbose,
            label=label,
            fallback_translators=fallback_translators or [],
        )
    except (TimeoutError, BatchTooLargeError) as exc:
        cause = "timeout" if isinstance(exc, TimeoutError) else "output too large"
        if len(units) <= min_batch_size:
            if verbose:
                print(f"OpenAI {cause}, skipped smallest batch: {label}", flush=True)
            return {}

        midpoint = len(units) // 2
        if verbose:
            char_count = sum(len(unit.text) for unit in units)
            print(
                f"OpenAI {cause}, splitting {label}: "
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
            fallback_translators=fallback_translators,
        )
        right = _translate_units_with_split_retry(
            units[midpoint:],
            translator=translator,
            retries=retries,
            min_batch_size=min_batch_size,
            verbose=verbose,
            label=f"{label}.2",
            fallback_translators=fallback_translators,
        )
        return left | right


def _strip_markers(text: str) -> str:
    """Bỏ tất cả marker, trả về plain text."""
    return re.sub(r"〈/?\d+〉", "", text)


def _fallback_to_slot_zero(translated_raw: str, expected_count: int) -> str | None:
    """Khi model phá vỡ cấu trúc marker, gom hết text dịch vào slot 0, slot khác trống.
    Mất chi tiết format từng slot nhưng giữ được nội dung dịch."""
    plain = _strip_markers(translated_raw).strip()
    if not plain:
        return None
    parts = [f"〈0〉{plain}〈/0〉"]
    parts.extend(f"〈{i}〉〈/{i}〉" for i in range(1, expected_count))
    return "".join(parts)


def _translate_units_with_marker_retry(
    units: list[ParagraphUnit],
    *,
    translator: Translator,
    retries: int,
    verbose: bool,
    label: str,
    fallback_translators: list[Translator] | None = None,
) -> dict[str, tuple[str, bool]]:
    pending = units
    # value = (translated_text, used_fallback)
    translations: dict[str, tuple[str, bool]] = {}
    last_raw: dict[str, str] = {}
    # Chain: primary trước, rồi tới các model fallback.
    translator_chain = [translator] + list(fallback_translators or [])
    current_translator = translator_chain[0]
    translator_idx = 0
    for attempt in range(retries + 1):
        if verbose and attempt > 0:
            print(
                f"Retrying marker validation failures: {label}, "
                f"attempt {attempt + 1}/{retries + 1}, {len(pending)} paragraphs",
                flush=True,
            )

        raw = current_translator.translate_batch(
            [{"id": unit.id, "text": unit.text} for unit in pending]
        )
        failed: list[ParagraphUnit] = []
        for unit in pending:
            translated = raw.get(unit.id)
            if translated is not None:
                repaired = _repair_markers(translated)
                if repaired != translated and verbose:
                    print(f"  [REPAIRED markers] {unit.id}", flush=True)
                translated = repaired
                last_raw[unit.id] = translated
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
                translations[unit.id] = (translated, False)
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
        # Còn fail và còn model mạnh hơn → escalate (đẩy lên model cao hơn).
        if attempt < retries and translator_idx + 1 < len(translator_chain):
            translator_idx += 1
            current_translator = translator_chain[translator_idx]
            if verbose:
                next_model = getattr(current_translator, "model", str(current_translator))
                print(
                    f"  Escalating to stronger model '{next_model}' for {len(pending)} "
                    f"paragraph(s) in {label}",
                    flush=True,
                )
            continue
        if attempt == retries:
            # Hết retry — thử fallback distribute cho từng unit còn fail.
            recovered: list[str] = []
            still_failed: list[str] = []
            for unit in failed:
                raw = last_raw.get(unit.id)
                if raw is None:
                    still_failed.append(unit.id)
                    continue
                fb = _fallback_to_slot_zero(raw, len(unit.slots))
                if fb is not None:
                    translations[unit.id] = (fb, True)
                    recovered.append(unit.id)
                else:
                    still_failed.append(unit.id)
            if verbose and recovered:
                print(
                    f"Fallback distribute (text → slot 0) for {len(recovered)} paragraphs: "
                    f"{', '.join(recovered)}",
                    flush=True,
                )
            if verbose and still_failed:
                print(
                    f"Skipped {len(still_failed)} paragraphs after {retries + 1} attempts: "
                    f"{', '.join(still_failed)}",
                    flush=True,
                )

    return translations


def _iter_content_xml_files(root: Path, *, include_extras: bool = False) -> list[Path]:
    patterns = list(DEFAULT_CONTENT_XML_PATTERNS)
    if include_extras:
        patterns.extend(EXTRA_CONTENT_XML_PATTERNS)
    files: list[Path] = []
    for pattern in patterns:
        files.extend(root.glob(pattern))
    return sorted(path for path in files if path.is_file())


def _extract_paragraph_units(tree: etree._ElementTree, file_id: str) -> list[ParagraphUnit]:
    units: list[ParagraphUnit] = []
    for paragraph_index, paragraph in enumerate(tree.xpath("//w:p", namespaces=NS)):
        pairs = _collect_text_nodes(paragraph)
        if not pairs:
            continue
        # Paragraph "đơn giản" (không hình ảnh, không field, không math) + cùng format
        # → gộp tất cả về 1 slot, tiết kiệm rất nhiều marker.
        if _is_simple_paragraph(paragraph) and _all_runs_same_format(pairs):
            slots = [TextSlot(
                nodes=[t for _, t in pairs],
                index=0,
                text="".join((t.text or "") for _, t in pairs),
            )]
        else:
            slots = _merge_adjacent_slots(pairs)
        if not slots:
            continue
        parts = [f"{_open_marker(s.index)}{s.text}{_close_marker(s.index)}" for s in slots]
        units.append(
            ParagraphUnit(
                id=f"{file_id}:p:{paragraph_index}",
                text="".join(parts),
                slots=slots,
            )
        )
    return units


# Các tag block aggressive-merge (giữ format chi tiết để không phá layout).
_COMPLEX_PARAGRAPH_TAGS = frozenset({
    f"{{{W_NS}}}drawing",
    f"{{{W_NS}}}pict",
    f"{{{W_NS}}}object",
    f"{{{W_NS}}}fldChar",
    f"{{{W_NS}}}instrText",
    f"{{{W_NS}}}oMath",
    f"{{{W_NS}}}oMathPara",
})


def _is_simple_paragraph(paragraph: etree._Element) -> bool:
    """Không tính descendant nằm trong <w:p> lồng (text box, ...)."""
    p_tag = f"{{{W_NS}}}p"

    def _walk(el: etree._Element) -> bool:
        for child in el:
            if child.tag == p_tag:
                continue
            if child.tag in _COMPLEX_PARAGRAPH_TAGS:
                return False
            if not _walk(child):
                return False
        return True

    return _walk(paragraph)


def _all_runs_same_format(pairs: list[tuple[etree._Element, etree._Element]]) -> bool:
    if len(pairs) <= 1:
        return True
    first_key = _run_format_key(pairs[0][0])
    return all(_run_format_key(run) == first_key for run, _ in pairs)


def _collect_text_nodes(paragraph: etree._Element) -> list[tuple[etree._Element, etree._Element]]:
    """Trả về [(run, w:t)] thuộc về paragraph này, KHÔNG đi vào các <w:p> lồng bên trong
    (text box, alternate content, ...) — chúng được liệt kê và dịch riêng."""
    pairs: list[tuple[etree._Element, etree._Element]] = []
    p_tag = f"{{{W_NS}}}p"
    r_tag = f"{{{W_NS}}}r"
    t_tag = f"{{{W_NS}}}t"

    def _walk(element: etree._Element) -> None:
        for child in element:
            if child.tag == p_tag:
                continue  # paragraph lồng — xử lý ở vòng //w:p ngoài
            if child.tag == r_tag:
                for text_node in child.findall(t_tag):
                    if (text_node.text or "") != "":
                        pairs.append((child, text_node))
                _walk(child)
            else:
                _walk(child)

    _walk(paragraph)
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


def _skip_reason(text: str, *, only_latin: bool = True) -> str | None:
    """None nếu nên dịch, ngược lại trả về lý do skip để log.

    only_latin=True: bỏ qua đoạn không chứa ký tự Latin (A-Za-z) — dùng để skip
    đoạn thuần Trung/Nhật/Hàn khi user chỉ muốn dịch tiếng Anh.
    """
    plain = MARKER_RE.sub(lambda match: match.group(2), text).strip()
    if not plain:
        return "empty"
    if not _has_letter(plain):
        return "no letters"
    if plain.startswith("http://") or plain.startswith("https://"):
        return "url"
    if only_latin and not _has_latin_letter(plain):
        return "no latin"
    return None


def _has_letter(text: str) -> bool:
    """True nếu có ít nhất 1 ký tự thuộc category Unicode 'Letter' (L*)."""
    return any(unicodedata.category(ch).startswith("L") for ch in text)


def _has_latin_letter(text: str) -> bool:
    return any(("A" <= ch <= "Z") or ("a" <= ch <= "z") for ch in text)


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


def _apply_translation(
    unit: ParagraphUnit, translated: str,
    *, verbose: bool = False, force_font: str | None = None,
) -> None:
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
        # Ép font (nếu chỉ định) cho tất cả run thuộc slot này.
        if force_font:
            for node in slot.nodes:
                parent_run = node.getparent()
                if parent_run is not None:
                    _force_run_font(parent_run, force_font)


def _force_run_font(run: etree._Element, font_name: str) -> None:
    """Đặt rPr/rFonts của 1 <w:r> dùng đúng font_name cho mọi script."""
    rpr_tag = f"{{{W_NS}}}rPr"
    rfonts_tag = f"{{{W_NS}}}rFonts"
    rpr = run.find(rpr_tag)
    if rpr is None:
        rpr = etree.Element(rpr_tag)
        run.insert(0, rpr)  # rPr phải là child đầu tiên của w:r
    for rf in rpr.findall(rfonts_tag):
        rpr.remove(rf)
    rfonts = etree.Element(rfonts_tag)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rfonts.set(f"{{{W_NS}}}{attr}", font_name)
    rpr.insert(0, rfonts)  # rFonts thường là child đầu trong rPr


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


def _token_aware_chunks(
    units: list[ParagraphUnit],
    translator: Translator,
    *,
    input_token_budget: int,
    hard_cap: int,
) -> list[list[ParagraphUnit]]:
    """Đóng batch theo ngân sách input tokens; mỗi batch tối đa hard_cap paragraphs."""
    estimate = getattr(translator, "estimate_input_tokens", None)
    if estimate is None:
        # Translator không hỗ trợ đếm token (vd DryRunTranslator) — fallback cố định.
        return _chunks(units, hard_cap)

    batches: list[list[ParagraphUnit]] = []
    current: list[ParagraphUnit] = []
    current_tokens = 0
    for unit in units:
        unit_tokens = estimate([{"id": unit.id, "text": unit.text}])
        # Nếu thêm 1 paragraph nữa vượt budget hoặc cap, đóng batch.
        if current and (
            current_tokens + unit_tokens > input_token_budget
            or len(current) >= hard_cap
        ):
            batches.append(current)
            current = []
            current_tokens = 0
        current.append(unit)
        current_tokens += unit_tokens
    if current:
        batches.append(current)
    return batches
