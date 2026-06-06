#!/usr/bin/env bash
cd "$(dirname "$0")"
python3 -m uvicorn app.main:app --host 0.0.0.0 --port 8000 --ws-ping-interval 20 --ws-ping-timeout 60 --log-level warning
