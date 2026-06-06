@echo off
cd /d "%~dp0"
python -m uvicorn app.main:app --host 0.0.0.0 --port 8000 --ws-ping-interval 20 --ws-ping-timeout 60 --log-level warning
pause
