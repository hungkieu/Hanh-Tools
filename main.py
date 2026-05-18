import argparse
import shutil
from pathlib import Path

from dotenv import load_dotenv

from lib.doc_converter import convert_doc_to_docx
from lib.docx_package import unzip_docx, zip_docx
from lib.ooxml_translator import translate_docx_xml_folder
from lib.openai_translator import DryRunTranslator, OpenAITranslator
from lib.translation_cache import TranslationCache


def main() -> None:
    load_dotenv()
    args = _parse_args()

    input_path = Path(args.input)
    converted_dir = Path(args.converted_dir)
    work_dir = Path(args.work_dir)
    output_path = Path(args.output)

    docx_path = _prepare_docx_input(
        input_path=input_path,
        converted_dir=converted_dir,
        verbose=args.verbose,
    )

    if work_dir.exists():
        if args.verbose:
            print(f"Removing old work directory: {work_dir}", flush=True)
        shutil.rmtree(work_dir)
    unzip_docx(docx_path, work_dir, verbose=args.verbose)

    if args.dry_run:
        translator = DryRunTranslator(prefix=args.dry_run_prefix)
        fallback_translators: list = []
    else:
        translator = OpenAITranslator(
            model=args.model,
            source_language=args.source_language,
            target_language=args.target_language,
            timeout=args.openai_timeout,
            max_completion_tokens=args.max_completion_tokens,
        )
        fallback_translators = [
            OpenAITranslator(
                model=m,
                source_language=args.source_language,
                target_language=args.target_language,
                timeout=args.openai_timeout,
                max_completion_tokens=args.max_completion_tokens,
            )
            for m in args.fallback_models
        ]

    cache = TranslationCache(enabled=not args.no_cache)
    stats = translate_docx_xml_folder(
        work_dir,
        translator,
        batch_size=args.batch_size,
        retries=args.retries,
        verbose=args.verbose,
        min_batch_size=args.min_batch_size,
        include_extras=args.include_extras,
        concurrency=args.concurrency,
        fallback_translators=fallback_translators,
        cache=cache,
        target_language=args.target_language,
        input_token_budget=args.input_token_budget,
        force_font=args.force_font or None,
    )
    cache.save()
    final_docx = zip_docx(work_dir, output_path, verbose=args.verbose)

    print("Đã xử lý xong:")
    print(f"- File .docx đã convert: {docx_path}")
    print(f"- File .docx đã dịch: {final_docx}")
    print(f"- XML files scanned: {stats.files_scanned}")
    print(f"- Paragraphs found: {stats.paragraphs_found}")
    print(f"- Paragraphs translated: {stats.paragraphs_translated}")
    print(f"- Paragraphs skipped: {stats.paragraphs_skipped}")
    print(f"- Paragraphs failed: {stats.paragraphs_failed}")
    print(f"- Paragraphs via fallback (text → slot 0): {stats.paragraphs_fallback}")
    print(f"- Cache hits / misses: {stats.cache_hits} / {stats.cache_misses}")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert .doc sang .docx, dịch OOXML, rồi đóng gói lại .docx."
    )
    parser.add_argument("--input", default="input/10.doc")
    parser.add_argument("--output", default="translated/10.vi.docx")
    parser.add_argument("--converted-dir", default="converted")
    parser.add_argument("--work-dir", default="work/unzipped_docx")
    parser.add_argument("--model", default="gpt-4.1-nano")
    parser.add_argument("--fallback-models", nargs="*", default=["gpt-4.1-mini"],
                        help="Danh sách model dùng khi primary fail marker.")
    parser.add_argument("--force-font", default="Times New Roman",
                        help="Ép font cho text dịch. Rỗng = giữ font gốc.")
    parser.add_argument("--input-token-budget", type=int, default=2000,
                        help="Ngân sách input tokens cho mỗi batch (token-aware packing).")
    parser.add_argument("--no-cache", action="store_true", help="Không dùng cache persistent.")
    parser.add_argument("--source-language", default="auto")
    parser.add_argument("--target-language", default="Vietnamese")
    parser.add_argument("--batch-size", type=int, default=5)
    parser.add_argument("--min-batch-size", type=int, default=1)
    parser.add_argument("--retries", type=int, default=2)
    parser.add_argument("--openai-timeout", type=float, default=180)
    parser.add_argument("--max-completion-tokens", type=int, default=8192)
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--dry-run-prefix", default="")
    parser.add_argument("--concurrency", type=int, default=3,
                        help="Số luồng song song gọi API (1 = sequential).")
    parser.add_argument(
        "--include-extras",
        action="store_true",
        help="Dịch thêm header/footer/footnote/endnote/comment (mặc định chỉ document.xml).",
    )
    return parser.parse_args()


def _prepare_docx_input(
    *,
    input_path: Path,
    converted_dir: Path,
    verbose: bool,
) -> Path:
    if not input_path.exists():
        raise FileNotFoundError(f"Không tìm thấy file input: {input_path}")

    suffix = input_path.suffix.lower()
    if suffix == ".docx":
        if verbose:
            print(f"Input đã là .docx, bỏ qua bước convert: {input_path}", flush=True)
        return input_path

    if suffix != ".doc":
        raise ValueError("Input phải là file .doc hoặc .docx")

    converted_path = converted_dir / f"{input_path.stem}.docx"
    if converted_path.exists():
        if verbose:
            print(f"Đã có file .docx trong converted, bỏ qua convert: {converted_path}", flush=True)
        return converted_path

    return convert_doc_to_docx(
        input_path,
        converted_dir,
        backend="libreoffice",
        verbose=verbose,
    )


if __name__ == "__main__":
    main()
