#!/bin/bash
# Nhấp đúp để chạy DocxTool mà không cần đóng gói
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  python3 -m venv .venv && .venv/bin/pip install -q -r requirements.txt
fi
.venv/bin/python app.py
