# Hạnh Tools

Công cụ dịch file Word (`.doc` / `.docx`) bằng OpenAI, giữ nguyên định dạng OOXML. Có giao diện đồ hoạ (Tkinter) và đóng gói được thành file `.exe` cho Windows.

---

## 1. Cài đặt

### A. Dùng file `.exe` (Windows — đơn giản nhất)

1. Tải `Hạnh Tools.exe` từ:
   - **GitHub Actions:** tab **Actions** → mở run mới nhất → mục **Artifacts** → tải `HanhTools-Windows.zip` → giải nén.
   - **Hoặc Release:** mục **Releases** nếu repo đã có tag (`v*`).
2. Cài [LibreOffice](https://www.libreoffice.org/download/download/) — bắt buộc nếu bạn cần convert file `.doc` cũ sang `.docx`. Nếu chỉ làm việc với `.docx` thì có thể bỏ qua.
3. Chạy đúp `Hạnh Tools.exe`. Không cần cài Python.

### B. Chạy từ source (macOS / Linux / Windows)

Yêu cầu:
- Python 3.10+
- [LibreOffice](https://www.libreoffice.org/) (chỉ cần khi xử lý file `.doc`)

```bash
git clone <repo-url>
cd hanhTool
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python hanh_tools_gui.py
```

### C. Tự build `.exe` trên Windows

```bat
build_exe.bat
```
Output: `dist\Hạnh Tools.exe`.

> Lưu ý: PyInstaller **không cross-compile**. Muốn `.exe` thì phải build trên Windows (hoặc dùng GitHub Actions — workflow đã có sẵn ở [.github/workflows/build-exe.yml](.github/workflows/build-exe.yml)).

---

## 2. Hướng dẫn sử dụng

### Lấy OpenAI API Key

1. Vào https://platform.openai.com/api-keys
2. Bấm **Create new secret key** → copy key dạng `sk-...`

### Các bước trong UI

![flow](./docs/flow.png)

1. **File input (.doc / .docx)** — bấm *Chọn...* để chọn file Word cần dịch.
2. **Lưu file dịch (.docx)** — bấm *Chọn...* để chọn nơi lưu file đầu ra (mặc định gợi ý `tên-file.vi.docx` cạnh file gốc).
3. **OPENAI_API_KEY** — dán key. Tick *Hiện* để xem key plaintext. Tick *Lưu key vào .env* để app tự ghi vào file `.env` cạnh exe, lần sau không cần nhập lại.
4. **Ngôn ngữ đích** — mặc định `Vietnamese`. Có thể đổi sang `English`, `Japanese`, ...
5. **Model** — mặc định `gpt-4.1-nano` (rẻ, nhanh). Có thể đổi `gpt-4.1-mini`, `gpt-4.1`, ...
6. **Dry run** — bật để test pipeline mà KHÔNG gọi OpenAI (không tốn tiền, không cần API key). Output sẽ là file gốc copy nguyên xi.
7. Bấm **Bắt đầu dịch**. Theo dõi tiến trình ở ô **Log** bên dưới.

### Kết quả

- File `.docx` đã dịch nằm đúng đường dẫn bạn chọn ở bước 2.
- Cạnh file output sẽ có 2 thư mục tạm:
  - `converted/` — chứa file `.docx` convert từ `.doc` (cache, không cần xoá).
  - `work/unzipped_docx/` — thư mục giải nén OOXML để dịch (có thể xoá sau khi xong).

### Mẹo & khắc phục lỗi

| Triệu chứng | Nguyên nhân / cách xử lý |
|---|---|
| `Không tìm thấy file input` | Sai đường dẫn — chọn lại file. |
| `LibreOffice not found` khi xử lý `.doc` | Cài LibreOffice, hoặc tự mở Word → Save As `.docx` rồi dùng `.docx` đó. |
| `Incorrect API key` / `401` | Key sai hoặc hết hạn. Lấy key mới tại platform.openai.com. |
| `RateLimitError` | Đã chạm rate limit — đợi vài phút hoặc đổi sang model rẻ hơn. |
| File output trống / chỉ vài đoạn được dịch | Xem ô Log: nếu có `Paragraphs failed` cao, thường do timeout — thử model mạnh hơn hoặc chạy lại. |

---

## 3. Tham khảo CLI

Nếu thích dòng lệnh, có sẵn `main.py`:

```bash
python main.py --input input/abc.doc --output translated/abc.vi.docx --target-language Vietnamese
python main.py --input file.docx --output out.docx --dry-run    # test không gọi API
```

Xem `python main.py --help` để biết toàn bộ tham số.

---

## 4. Bảo mật

- API key được lưu cục bộ trong `.env` cạnh exe — **không** push file này lên git (đã có trong `.gitignore`).
- Nội dung file Word sẽ được gửi tới OpenAI để dịch. Đừng dùng với tài liệu nhạy cảm/bảo mật trừ khi đã hiểu rõ chính sách của OpenAI.
