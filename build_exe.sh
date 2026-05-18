#!/usr/bin/env bash
# Build "Hạnh Tools" bằng PyInstaller cho hệ điều hành hiện tại.
# Trên macOS/Linux script này sẽ tạo binary tương ứng (KHÔNG phải .exe Windows).
# Để có .exe thật, chạy build_exe.bat trên máy Windows.
set -euo pipefail

python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt

rm -rf build dist "Hạnh Tools.spec"

python3 -m PyInstaller \
  --noconfirm \
  --clean \
  --onefile \
  --windowed \
  --name "Hạnh Tools" \
  --collect-all tiktoken \
  --collect-all tiktoken_ext \
  --collect-submodules openai \
  --collect-submodules lib \
  hanh_tools_gui.py

echo "=== Xong! Output trong thư mục dist/ ==="
