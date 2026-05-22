#!/usr/bin/env bash
set -e
python3 -m pip install -r requirements.txt
echo "服务启动中，请在浏览器打开 http://127.0.0.1:8000"
python3 -m uvicorn app:app --host 127.0.0.1 --port 8000
