"""数量入力待ち案件のバッチ処理

SFAから現調写真ZIPをダウンロード → パイプライン実行 → 結果Excelを生成。

使い方:
  1. ブラウザでSFAにログイン済みの状態で実行
  2. WebSocketサーバーが起動し、ブラウザからZIPデータを受信
  3. 自動的にパイプラインを実行して見積Excelを生成
"""

from __future__ import annotations

import asyncio
import base64
import json
import logging
import os
import sys
import zipfile
import threading
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# プロジェクトルート
ROOT = Path(__file__).parent.parent
SURVEY_DIR = ROOT / "現調写真"
OUTPUT_DIR = ROOT / "output"
LINEUP_DIR = ROOT / "ラインナップ表"
TEMPLATE_DIR = ROOT / "見積りテンプレート"


def load_env():
    """`.env` ファイルからAPIキーを読み込み"""
    env_path = ROOT / ".env"
    if env_path.exists():
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, _, value = line.partition("=")
                    os.environ[key.strip()] = value.strip()


def get_download_js(file_key: str, filename: str, ws_port: int) -> str:
    """ブラウザで実行するZIPダウンロード+WebSocket転送JSを生成

    SFAの正しいダウンロードエンドポイント:
      GET /ajaxBuckets/download_bucket_file?file_key={key}&from_business_card=
    """
    return f"""
(function() {{
    var url = '/ajaxBuckets/download_bucket_file?file_' + 'key=' + encodeURIComponent('{file_key}') + '&from_business_card=';
    var xhr = new XMLHttpRequest();
    xhr.open('GET', url, true);
    xhr.responseType = 'arraybuffer';
    xhr.onload = function() {{
        if (xhr.status !== 200) {{
            window._zipDlResult = 'HTTP error: ' + xhr.status;
            return;
        }}
        var bytes = new Uint8Array(xhr.response);
        if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) {{
            window._zipDlResult = 'Not a ZIP file (header: ' + bytes[0] + ',' + bytes[1] + ')';
            return;
        }}
        // チャンク分割でBase64エンコード（大容量対応）
        var chunkSize = 32768;
        var b64parts = [];
        for (var i = 0; i < bytes.length; i += chunkSize) {{
            var chunk = bytes.subarray(i, Math.min(i + chunkSize, bytes.length));
            b64parts.push(String.fromCharCode.apply(null, chunk));
        }}
        var b64 = btoa(b64parts.join(''));
        var ws = new WebSocket('ws://127.0.0.1:{ws_port}');
        ws.onopen = function() {{
            ws.send(JSON.stringify({{name: '{filename}', data: b64}}));
        }};
        ws.onmessage = function(ev) {{
            window._zipDlResult = 'OK: ' + bytes.length + ' bytes sent';
        }};
        ws.onerror = function() {{
            window._zipDlResult = 'WebSocket error';
        }};
    }};
    xhr.onerror = function() {{
        window._zipDlResult = 'XHR error';
    }};
    xhr.send();
    window._zipDlResult = 'downloading...';
}})();
"""


def receive_zip_from_browser(port: int = 19908, timeout_sec: int = 120) -> tuple[bytes, str]:
    """ブラウザからWebSocket経由でZIPデータを受信

    Returns:
        (zip_bytes, filename)
    """
    import websockets

    result = [None, None]

    async def handler(ws):
        msg = await ws.recv()
        payload = json.loads(msg)
        result[0] = base64.b64decode(payload["data"])
        result[1] = payload["name"]
        await ws.send("OK")
        logger.info(f"受信完了: {result[1]} ({len(result[0]):,} bytes)")

    async def main():
        server = await websockets.serve(
            handler, "127.0.0.1", port,
            max_size=50 * 1024 * 1024,  # 50MB上限
        )
        actual_port = server.sockets[0].getsockname()[1]
        logger.info(f"ZIP受信サーバー起動: ws://127.0.0.1:{actual_port}")

        for _ in range(timeout_sec * 2):
            await asyncio.sleep(0.5)
            if result[0] is not None:
                break
        server.close()
        await server.wait_closed()

    asyncio.run(main())
    return result[0], result[1]


def extract_zip(zip_bytes: bytes, project_name: str) -> Path:
    """ZIPを現調写真フォルダに展開

    Returns:
        展開先フォルダパス
    """
    import io

    extract_dir = SURVEY_DIR / project_name
    extract_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        zf.extractall(extract_dir)

    # 画像ファイル数をカウント
    image_exts = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".heic"}
    images = [
        f for f in extract_dir.rglob("*")
        if f.is_file() and f.suffix.lower() in image_exts
    ]
    logger.info(f"ZIP展開完了: {extract_dir.name} ({len(images)}枚の画像)")
    return extract_dir


def run_pipeline_for_project(
    survey_dir: Path,
    template_name: str = "田村基本形",
    api_key: Optional[str] = None,
) -> Path:
    """パイプラインを実行してExcelを生成"""
    from pipeline import run_pipeline

    return run_pipeline(
        survey_dir=survey_dir,
        lineup_dir=LINEUP_DIR,
        template_dir=TEMPLATE_DIR,
        template_name=template_name,
        api_key=api_key,
    )


def process_single(
    zip_bytes: bytes,
    project_name: str,
    template_name: str = "田村基本形",
    api_key: Optional[str] = None,
) -> Path:
    """1物件を処理: ZIP展開 → パイプライン → Excel

    Returns:
        出力Excelパス
    """
    # ZIP展開
    survey_dir = extract_zip(zip_bytes, project_name)

    # パイプライン実行
    result_path = run_pipeline_for_project(
        survey_dir=survey_dir,
        template_name=template_name,
        api_key=api_key,
    )

    logger.info(f"完了: {result_path.name}")
    return result_path


# --- CLI ---

if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    load_env()

    print("=" * 50)
    print("  LED見積バッチ処理")
    print("=" * 50)
    print()
    print("ブラウザからZIPファイルを送信してください。")
    print("WebSocket ポート: 19908")
    print()

    zip_data, filename = receive_zip_from_browser(port=19908)
    if zip_data:
        # ファイル名から物件名を推定
        name = filename.replace("施工前", "").replace("【施工前】", "")
        name = name.replace(".zip", "").strip()
        name = name.strip("　 ")

        print(f"\n物件名: {name}")
        result = process_single(zip_data, name)
        print(f"\n出力: {result}")
    else:
        print("データを受信できませんでした。")
