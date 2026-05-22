@echo off
chcp 65001 >nul
echo 正在安装依赖...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo 依赖安装失败，请确认已安装 Python 3.11+
  pause
  exit /b 1
)
echo.
echo 服务启动中，请在浏览器打开 http://127.0.0.1:8000
echo 按 Ctrl+C 可以停止服务。
python -m uvicorn app:app --host 127.0.0.1 --port 8000
pause
