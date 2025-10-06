@echo off
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
pyinstaller --noconfirm --onefile --windowed 20251003.py
if %errorlevel% neq 0 (
  echo PyInstaller 构建失败，检查输出日志。
  pause
  exit /b %errorlevel%
)

echo 构建完成。生成的 exe 位于 dist\ 目录。
pause
