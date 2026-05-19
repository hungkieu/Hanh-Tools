import os
import queue
import shutil
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from dotenv import load_dotenv

from lib.doc_converter import convert_doc_to_docx
from lib.docx_package import unzip_docx, zip_docx
from lib.ooxml_translator import CancelledError, translate_docx_xml_folder
from lib.openai_translator import DryRunTranslator, OpenAITranslator
from lib.translation_cache import TranslationCache, cache_file, open_cache_folder


MODEL_CHOICES = ["gpt-4.1-nano", "gpt-4.1-mini"]
# Default fallback chain — khi model hiện tại fail marker liên tục thì đẩy lên model mạnh hơn.
FALLBACK_CHAIN = {
    "gpt-4.1-nano": ["gpt-4.1-mini"],
    "gpt-4.1-mini": [],
}
DEFAULT_OUTPUT_FONT = "Times New Roman"


APP_TITLE = "Hạnh Tools"


class _StreamToQueue:
    def __init__(self, q: "queue.Queue[str]") -> None:
        self._q = q

    def write(self, text: str) -> None:
        if text:
            self._q.put(text)

    def flush(self) -> None:
        pass


class HanhToolsApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("820x680")
        self.root.minsize(720, 580)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.target_lang_var = tk.StringVar(value="Vietnamese")
        self.model_var = tk.StringVar(value="gpt-4.1-nano")
        self.output_font_var = tk.StringVar(value=DEFAULT_OUTPUT_FONT)
        self.dry_run_var = tk.BooleanVar(value=False)
        self.include_extras_var = tk.BooleanVar(value=False)
        self.use_cache_var = tk.BooleanVar(value=True)
        self.only_english_var = tk.BooleanVar(value=True)
        self.batch_size_var = tk.StringVar(value="20")
        self.max_tokens_var = tk.StringVar(value="8192")
        self.concurrency_var = tk.StringVar(value="3")
        self.token_budget_var = tk.StringVar(value="2000")
        self.api_key_var = tk.StringVar()
        self.show_key_var = tk.BooleanVar(value=False)
        self.save_env_var = tk.BooleanVar(value=True)

        # Tự load API key từ .env hoặc biến môi trường nếu có
        load_dotenv()
        existing_key = os.environ.get("OPENAI_API_KEY", "")
        if existing_key:
            self.api_key_var.set(existing_key)

        self._log_queue: "queue.Queue[str]" = queue.Queue()
        self._worker: threading.Thread | None = None
        self._cancel_event = threading.Event()

        self._build_ui()
        self._poll_log_queue()

    def _build_ui(self) -> None:
        pad = {"padx": 8, "pady": 4}
        frm = ttk.Frame(self.root)
        frm.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Hàng 1-3: input / output / API key ---
        ttk.Label(frm, text="File input (.doc / .docx):").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.input_var).grid(row=0, column=1, columnspan=3, sticky="we", **pad)
        ttk.Button(frm, text="Chọn...", command=self._pick_input).grid(row=0, column=4, **pad)

        ttk.Label(frm, text="Lưu file dịch (.docx):").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.output_var).grid(row=1, column=1, columnspan=3, sticky="we", **pad)
        ttk.Button(frm, text="Chọn...", command=self._pick_output).grid(row=1, column=4, **pad)

        ttk.Label(frm, text="OPENAI_API_KEY:").grid(row=2, column=0, sticky="w", **pad)
        self.api_key_entry = ttk.Entry(frm, textvariable=self.api_key_var, show="*")
        self.api_key_entry.grid(row=2, column=1, columnspan=3, sticky="we", **pad)
        ttk.Checkbutton(frm, text="Hiện", variable=self.show_key_var,
                        command=self._toggle_key_visibility).grid(row=2, column=4, sticky="w", **pad)

        # --- Hàng 4: dịch & model ---
        ttk.Label(frm, text="Ngôn ngữ đích:").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.target_lang_var).grid(row=3, column=1, sticky="we", **pad)
        ttk.Label(frm, text="Model:").grid(row=3, column=2, sticky="e", **pad)
        ttk.Combobox(frm, textvariable=self.model_var, values=MODEL_CHOICES,
                     state="readonly").grid(row=3, column=3, sticky="we", **pad)

        # --- Hàng 4b: font đầu ra ---
        ttk.Label(frm, text="Font đầu ra:").grid(row=4, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.output_font_var).grid(row=4, column=1, sticky="we", **pad)
        ttk.Label(frm, text="(rỗng = giữ font gốc)",
                  foreground="#777").grid(row=4, column=2, columnspan=2, sticky="w", **pad)

        # --- Hàng 5: tham số nâng cao (4 ô) ---
        params = ttk.LabelFrame(frm, text="Tham số nâng cao")
        params.grid(row=5, column=0, columnspan=5, sticky="we", **pad)
        for i in range(8):
            params.columnconfigure(i, weight=1)
        ttk.Label(params, text="Batch:").grid(row=0, column=0, sticky="e", **pad)
        ttk.Entry(params, textvariable=self.batch_size_var, width=6).grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(params, text="Max output tokens:").grid(row=0, column=2, sticky="e", **pad)
        ttk.Entry(params, textvariable=self.max_tokens_var, width=8).grid(row=0, column=3, sticky="w", **pad)
        ttk.Label(params, text="Luồng song song:").grid(row=0, column=4, sticky="e", **pad)
        ttk.Entry(params, textvariable=self.concurrency_var, width=5).grid(row=0, column=5, sticky="w", **pad)
        ttk.Label(params, text="Input token budget:").grid(row=0, column=6, sticky="e", **pad)
        ttk.Entry(params, textvariable=self.token_budget_var, width=7).grid(row=0, column=7, sticky="w", **pad)

        # --- Hàng 6: tuỳ chọn (checkboxes) ---
        flags = ttk.Frame(frm)
        flags.grid(row=6, column=0, columnspan=5, sticky="we", **pad)
        ttk.Checkbutton(flags, text="Dùng cache",
                        variable=self.use_cache_var).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(flags, text="Lưu key vào .env",
                        variable=self.save_env_var).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(flags, text="Dịch cả header/footer/footnote/comment",
                        variable=self.include_extras_var).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(flags, text="Chỉ dịch tiếng Anh (skip đoạn thuần Trung/Nhật/Hàn)",
                        variable=self.only_english_var).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(flags, text="Dry run (không gọi API)",
                        variable=self.dry_run_var).pack(side=tk.LEFT)

        # --- Hàng 7: nút action ---
        btns = ttk.Frame(frm)
        btns.grid(row=7, column=0, columnspan=5, sticky="we", **pad)
        self.run_btn = ttk.Button(btns, text="Bắt đầu dịch", command=self._on_run)
        self.run_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.stop_btn = ttk.Button(btns, text="Dừng", command=self._on_stop, state="disabled")
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 24))
        ttk.Button(btns, text="Mở thư mục cache",
                   command=self._open_cache_folder).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btns, text="Xoá cache",
                   command=self._clear_cache).pack(side=tk.LEFT)

        # --- Hàng 8: log ---
        ttk.Label(frm, text="Log:").grid(row=8, column=0, sticky="w", **pad)
        log_frame = ttk.Frame(frm)
        log_frame.grid(row=9, column=0, columnspan=5, sticky="nsew", **pad)
        self.log_text = tk.Text(log_frame, height=16, wrap="word", state="disabled")
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=sb.set)

        for col in (1, 2, 3):
            frm.columnconfigure(col, weight=1)
        frm.rowconfigure(9, weight=1)

    def _toggle_key_visibility(self) -> None:
        self.api_key_entry.config(show="" if self.show_key_var.get() else "*")

    def _open_cache_folder(self) -> None:
        try:
            open_cache_folder()
        except Exception as e:
            messagebox.showwarning(APP_TITLE, f"Không mở được thư mục cache: {e}")

    def _clear_cache(self) -> None:
        if not messagebox.askyesno(APP_TITLE, "Xoá toàn bộ cache?"):
            return
        try:
            cache_file().unlink(missing_ok=True)
            messagebox.showinfo(APP_TITLE, "Đã xoá cache.")
        except Exception as e:
            messagebox.showwarning(APP_TITLE, f"Lỗi xoá cache: {e}")

    def _pick_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Chọn file Word",
            filetypes=[("Word files", "*.doc *.docx"), ("All files", "*.*")],
        )
        if path:
            self.input_var.set(path)
            if not self.output_var.get():
                stem = Path(path).stem
                self.output_var.set(str(Path(path).with_name(f"{stem}.vi.docx")))

    def _pick_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Lưu file dịch",
            defaultextension=".docx",
            filetypes=[("Word docx", "*.docx")],
        )
        if path:
            self.output_var.set(path)

    def _append_log(self, text: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, text)
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def _poll_log_queue(self) -> None:
        try:
            while True:
                text = self._log_queue.get_nowait()
                self._append_log(text)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_log_queue)

    def _on_run(self) -> None:
        input_path = self.input_var.get().strip()
        output_path = self.output_var.get().strip()
        if not input_path or not output_path:
            messagebox.showwarning(APP_TITLE, "Vui lòng chọn file input và đường dẫn output.")
            return
        if self._worker and self._worker.is_alive():
            return

        self._cancel_event.clear()
        self.run_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

        api_key = self.api_key_var.get().strip()
        dry_run = self.dry_run_var.get()
        if not dry_run and not api_key:
            messagebox.showwarning(APP_TITLE, "Vui lòng nhập OPENAI_API_KEY (hoặc bật Dry run).")
            return
        try:
            batch_size = max(1, int(self.batch_size_var.get().strip()))
            max_tokens = max(512, int(self.max_tokens_var.get().strip()))
            concurrency = max(1, min(10, int(self.concurrency_var.get().strip())))
            token_budget = max(200, int(self.token_budget_var.get().strip()))
        except ValueError:
            messagebox.showwarning(APP_TITLE, "Batch / Max tokens / Luồng / Token budget phải là số nguyên.")
            return
        if api_key:
            os.environ["OPENAI_API_KEY"] = api_key
            if self.save_env_var.get():
                self._save_env_file(api_key)

        self._worker = threading.Thread(
            target=self._run_pipeline,
            args=(input_path, output_path, self.target_lang_var.get().strip() or "Vietnamese",
                  self.model_var.get().strip() or "gpt-4.1-nano", dry_run,
                  self.include_extras_var.get(), batch_size, max_tokens, concurrency,
                  token_budget, self.use_cache_var.get(),
                  self.output_font_var.get().strip(),
                  self.only_english_var.get()),
            daemon=True,
        )
        self._worker.start()

    def _on_stop(self) -> None:
        if not self._worker or not self._worker.is_alive():
            return
        self._cancel_event.set()
        self.stop_btn.config(state="disabled")
        self._log_queue.put("\n>>> Đang yêu cầu dừng... (sẽ dừng sau khi batch hiện tại xong)\n")

    def _save_env_file(self, api_key: str) -> None:
        env_path = Path(os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
                        else os.path.dirname(os.path.abspath(__file__))) / ".env"
        lines: list[str] = []
        if env_path.exists():
            for line in env_path.read_text(encoding="utf-8").splitlines():
                if not line.startswith("OPENAI_API_KEY="):
                    lines.append(line)
        lines.append(f"OPENAI_API_KEY={api_key}")
        env_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    def _run_pipeline(self, input_path: str, output_path: str, target_lang: str, model: str,
                      dry_run: bool, include_extras: bool, batch_size: int, max_tokens: int,
                      concurrency: int, token_budget: int, use_cache: bool,
                      output_font: str, only_english: bool = True) -> None:
        stream = _StreamToQueue(self._log_queue)
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = stream
        sys.stderr = stream
        try:
            load_dotenv()
            in_p = Path(input_path)
            out_p = Path(output_path)
            base_dir = out_p.parent if out_p.parent.as_posix() else Path(".")
            converted_dir = base_dir / "converted"
            work_dir = base_dir / "work" / "unzipped_docx"
            out_p.parent.mkdir(parents=True, exist_ok=True)
            converted_dir.mkdir(parents=True, exist_ok=True)

            print(f"Input: {in_p}")
            print(f"Output: {out_p}")
            print(f"Target language: {target_lang} | Model: {model} | Dry-run: {dry_run}")
            print(f"Batch size: {batch_size} | Max output tokens: {max_tokens} | Concurrency: {concurrency} | Token budget: {token_budget}")
            print(f"Include extras: {include_extras} | Cache: {'on' if use_cache else 'off'} | "
                  f"Output font: {output_font or '(giữ font gốc)'}")

            docx_path = self._prepare_docx(in_p, converted_dir)

            if work_dir.exists():
                print(f"Removing old work directory: {work_dir}")
                shutil.rmtree(work_dir)
            unzip_docx(docx_path, work_dir, verbose=True)

            if dry_run:
                translator = DryRunTranslator(prefix="")
                fallback_translators: list = []
            else:
                translator = OpenAITranslator(
                    model=model, source_language="auto", target_language=target_lang,
                    timeout=180, max_completion_tokens=max_tokens,
                    only_english=only_english,
                )
                fallback_translators = [
                    OpenAITranslator(
                        model=m, source_language="auto", target_language=target_lang,
                        timeout=180, max_completion_tokens=max_tokens,
                        only_english=only_english,
                    )
                    for m in FALLBACK_CHAIN.get(model, [])
                ]
                if fallback_translators:
                    print(f"Fallback chain: {' → '.join([model] + FALLBACK_CHAIN[model])}")
            cache = TranslationCache(enabled=use_cache)
            if use_cache:
                print(f"Cache: {cache.entry_count()} entries ({cache.file_size_kb():.1f} KB) at {cache.path}")
            stats = translate_docx_xml_folder(
                work_dir, translator,
                batch_size=batch_size, retries=2, verbose=True, min_batch_size=1,
                cancel_check=self._cancel_event.is_set,
                include_extras=include_extras,
                concurrency=concurrency,
                fallback_translators=fallback_translators,
                cache=cache,
                target_language=target_lang,
                input_token_budget=token_budget,
                force_font=output_font or None,
                only_english=only_english,
            )
            cache.save()
            final_docx = zip_docx(work_dir, out_p, verbose=True)

            print("\n=== HOÀN TẤT ===")
            print(f"- File .docx đã convert: {docx_path}")
            print(f"- File .docx đã dịch: {final_docx}")
            print(f"- XML files scanned: {stats.files_scanned}")
            print(f"- Paragraphs found: {stats.paragraphs_found}")
            print(f"- Paragraphs translated: {stats.paragraphs_translated}")
            print(f"- Paragraphs skipped: {stats.paragraphs_skipped}")
            print(f"- Paragraphs failed: {stats.paragraphs_failed}")
            print(f"- Paragraphs via fallback (text → slot 0): {stats.paragraphs_fallback}")
            print(f"- Cache hits / misses: {stats.cache_hits} / {stats.cache_misses}")
        except CancelledError:
            print("\n>>> Đã dừng theo yêu cầu. File output có thể chưa hoàn chỉnh.")
        except Exception as e:
            print(f"\n[LỖI] {e}")
        finally:
            sys.stdout, sys.stderr = old_out, old_err
            def _reset_buttons() -> None:
                self.run_btn.config(state="normal")
                self.stop_btn.config(state="disabled")
            self.root.after(0, _reset_buttons)

    def _prepare_docx(self, input_path: Path, converted_dir: Path) -> Path:
        if not input_path.exists():
            raise FileNotFoundError(f"Không tìm thấy file input: {input_path}")
        suffix = input_path.suffix.lower()
        if suffix == ".docx":
            print(f"Input đã là .docx, bỏ qua convert: {input_path}")
            return input_path
        if suffix != ".doc":
            raise ValueError("Input phải là file .doc hoặc .docx")
        converted_path = converted_dir / f"{input_path.stem}.docx"
        if converted_path.exists():
            print(f"Đã có file .docx trong converted: {converted_path}")
            return converted_path
        return convert_doc_to_docx(input_path, converted_dir, backend="libreoffice", verbose=True)


def main() -> None:
    root = tk.Tk()
    try:
        if sys.platform.startswith("win"):
            root.iconbitmap(default=os.path.join(os.path.dirname(__file__), "icon.ico"))
    except Exception:
        pass
    HanhToolsApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
