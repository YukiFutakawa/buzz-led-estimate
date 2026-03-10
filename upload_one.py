#!/usr/bin/env python3
"""1件分のExcelファイルをWebSocket経由で提供するサーバー

Usage:
    python upload_one.py <excel_path>
    # → ポート番号を出力し、WebSocket接続を1回待ってから終了
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from pathlib import Path
from sfa_uploader import start_file_server, stop_file_server, get_upload_script
import time

if len(sys.argv) < 2:
    print("Usage: python upload_one.py <excel_path>")
    sys.exit(1)

file_path = Path(sys.argv[1])
if not file_path.exists():
    print(f"ERROR: {file_path}")
    sys.exit(1)

port = start_file_server(file_path)
print(f"PORT:{port}")
print(f"FILE:{file_path.name}")
sys.stdout.flush()

# Wait up to 5 minutes for the upload to complete
try:
    time.sleep(300)
except KeyboardInterrupt:
    pass
finally:
    stop_file_server()
