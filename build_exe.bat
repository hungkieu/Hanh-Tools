@echo off
REM Build "Hạnh Tools.exe" bằng PyInstaller (chạy trên Windows)
REM Yêu cầu: Python 3.10+ đã cài, đang ở thư mục dự án.

setlocal
where python >nul 2>&1 || (echo [LOI] Chua cai Python & exit /b 1)

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

REM Dọn build cũ
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "Hạnh Tools.spec" del /q "Hạnh Tools.spec"

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "Hạnh Tools" ^
  --collect-all tiktoken ^
  --collect-all tiktoken_ext ^
  --collect-submodules openai ^
  --collect-submodules lib ^
  hanh_tools_gui.py

echo.
echo === Xong! File cai dat tai: dist\Hạnh Tools.exe ===
endlocal
