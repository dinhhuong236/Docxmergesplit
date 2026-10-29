#!/bin/bash
# Đóng gói DocxTool thành ứng dụng macOS (dist/DocxTool.app)
set -e
cd "$(dirname "$0")"

PY=${PYTHON:-python3}
if ! $PY -c "import tkinter" 2>/dev/null; then
  echo "Python hiện tại thiếu Tkinter (giao diện)."
  echo "  • Nếu dùng Homebrew:  brew install python-tk"
  echo "  • Hoặc cài Python từ https://www.python.org/downloads/macos/ (có sẵn Tkinter)"
  exit 1
fi

$PY -m venv .venv
source .venv/bin/activate
pip install --upgrade pip >/dev/null
pip install -r requirements.txt pyinstaller

pyinstaller --noconfirm --clean --windowed \
  --name "DocxTool" \
  --osx-bundle-identifier "vn.docxtool.app" \
  --collect-submodules lxml \
  app.py

echo
echo "Xong! Ứng dụng nằm ở: $(pwd)/dist/DocxTool.app"
echo "Kéo vào thư mục Applications để dùng. Lần đầu mở: chuột phải > Open (vì app chưa ký số)."
