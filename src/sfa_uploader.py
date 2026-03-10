"""ネクストSFA 履歴情報ファイルアップロード

ブラウザ自動化（Claude Code MCP）経由でSFAの履歴情報にファイルを添付する。

## 自動化フロー（Claude Code が実行する手順）

### 前提
- Chrome で SFA にログイン済み
- Claude in Chrome 拡張が接続済み

### Step-by-step

1. **Python**: `start_file_server(file_path)` → WebSocket ポート取得
2. **Browser**: `navigate` → `/projects/detail/{project_id}`
3. **Browser**: `find("履歴情報")` → click（タブ切替）
4. **Browser**: `wait(2)` → 履歴一覧が読み込まれるまで待つ
5. **Browser**: 左サイドバーの最初の履歴エントリをクリック
6. **Browser**: `find("履歴追加")` → click（フォーム表示）
7. **Browser**: `wait(2)` → フォームが読み込まれるまで待つ
8. **Browser**: `javascript_tool` → `get_upload_script(port, filename)` を実行
9. **Browser**: `javascript_tool` → `window._sfaUploadResult` をポーリング
   - "PENDING" → 待機
   - "SUCCESS:..." → 成功
   - "ERROR:..." → エラー
10. **Python**: `stop_file_server()`
"""

from __future__ import annotations

import asyncio
import base64
import logging
import threading
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# グローバルサーバー参照（停止用）
_server_thread: Optional[threading.Thread] = None
_server_stop_event: Optional[threading.Event] = None


def start_file_server(file_path: Path, port: int = 0) -> int:
    """WebSocket ファイルサーバーを起動（バックグラウンドスレッド）

    HTTPS ページからの fetch は mixed-content で blocked されるため、
    WebSocket (ws://127.0.0.1) を使用してファイルを転送する。

    Args:
        file_path: 転送するファイルのパス
        port: ポート番号（0 で自動割当）

    Returns:
        実際に使用されたポート番号
    """
    global _server_thread, _server_stop_event

    if not file_path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {file_path}")

    file_bytes = file_path.read_bytes()
    b64_data = base64.b64encode(file_bytes).decode("ascii")
    file_size = len(file_bytes)

    logger.info(f"ファイル準備完了: {file_path.name} ({file_size:,} bytes)")

    _server_stop_event = threading.Event()
    actual_port_holder = [0]

    def _run_server():
        import websockets  # type: ignore

        async def handler(ws):
            try:
                msg = await ws.recv()
                if msg == "GET_FILE":
                    await ws.send(b64_data)
                    logger.info(f"ファイル転送完了: {file_path.name}")
            except Exception as e:
                logger.warning(f"WebSocket エラー: {e}")

        async def main():
            server = await websockets.serve(handler, "127.0.0.1", port)
            actual_port_holder[0] = server.sockets[0].getsockname()[1]
            logger.info(f"WebSocket サーバー起動: ws://127.0.0.1:{actual_port_holder[0]}")
            _ready_event.set()

            # stop_event が set されるまで待つ（最大5分）
            while not _server_stop_event.is_set():
                await asyncio.sleep(0.5)
            server.close()
            await server.wait_closed()

        asyncio.run(main())

    _ready_event = threading.Event()
    _server_thread = threading.Thread(target=_run_server, daemon=True)
    _server_thread.start()
    _ready_event.wait(timeout=10)

    return actual_port_holder[0]


def stop_file_server():
    """WebSocket サーバーを停止"""
    global _server_thread, _server_stop_event
    if _server_stop_event:
        _server_stop_event.set()
    if _server_thread:
        _server_thread.join(timeout=5)
    _server_thread = None
    _server_stop_event = None
    logger.info("WebSocket サーバー停止")


def get_upload_script(ws_port: int, filename: str) -> str:
    """ブラウザで実行する JS コードを生成

    このスクリプトは「履歴追加」フォーム表示後に実行する。
    以下を自動で行う:
      1. フォームデータを jQuery.serializeArray() で取得
      2. Dropzone の URL と送信ハンドラを設定
      3. WebSocket 経由でファイルを受信
      4. Dropzone.processQueue() でアップロード実行

    Args:
        ws_port: WebSocket サーバーのポート番号
        filename: アップロードするファイル名

    Returns:
        実行用 JavaScript コード文字列
    """
    # JS内のバッククォートをエスケープ
    return f"""
(function() {{
    // Step 1: フォームデータをシリアライズ
    var $form = jQuery('#PastCompanyHistoryRegistForm');
    if (!$form.length) {{
        window._sfaUploadResult = 'ERROR: PastCompanyHistoryRegistForm not found';
        return;
    }}
    var serializedArray = $form.serializeArray();

    // Step 2: Dropzone を取得・設定
    var dz = null;
    for (var i = 0; i < Dropzone.instances.length; i++) {{
        if (Dropzone.instances[i].options.url === 'PastCompanyHistoryRegistForm') {{
            dz = Dropzone.instances[i];
            break;
        }}
    }}
    if (!dz) {{
        window._sfaUploadResult = 'ERROR: Dropzone instance not found';
        return;
    }}

    dz.options.url = '/ajaxHistories/regist_complete';

    // sending ハンドラ: フォームフィールドを FormData に追加
    dz.on('sending', function(file, xhr, formData) {{
        for (var j = 0; j < serializedArray.length; j++) {{
            formData.append(serializedArray[j].name, serializedArray[j].value);
        }}
        formData.append('data[bifurcatio]', 'none');
    }});

    dz.on('success', function(file, response) {{
        window._sfaUploadResult = 'SUCCESS:' + JSON.stringify(response).substring(0, 500);
    }});

    dz.on('error', function(file, msg) {{
        window._sfaUploadResult = 'ERROR:' + JSON.stringify(msg).substring(0, 300);
    }});

    // Step 3: WebSocket 経由でファイルを取得し Dropzone にキュー
    window._sfaUploadResult = 'PENDING';
    var ws = new WebSocket('ws://127.0.0.1:{ws_port}');
    ws.onopen = function() {{
        ws.send('GET_FILE');
    }};
    ws.onmessage = function(e) {{
        var bin = atob(e.data);
        var bytes = new Uint8Array(bin.length);
        for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);

        var file = new File([bytes], '{filename}', {{
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }});

        file.status = Dropzone.QUEUED;
        file.accepted = true;
        file.upload = {{
            uuid: Math.random().toString(36).substring(2),
            progress: 0,
            total: file.size,
            bytesSent: 0,
            filename: file.name
        }};

        dz.files.push(file);
        dz.emit('addedfile', file);

        // Step 4: アップロード実行
        dz.processQueue();
        ws.close();
    }};
    ws.onerror = function(e) {{
        window._sfaUploadResult = 'ERROR: WebSocket connection failed';
    }};
}})();
"""


# 定数: SFA の URL パターン
SFA_BASE_URL = "https://nsfa.next-cloud.jp"


def get_project_url(project_id: str) -> str:
    """物件詳細ページの URL"""
    return f"{SFA_BASE_URL}/projects/detail/{project_id}"


PHASE_AI_DONE = "AI作成済み"


def mark_phase_ai_done(project_id: str) -> str:
    """アップロード成功後にフェーズを「AI作成済み」に変更

    Returns:
        更新後のフェーズ名
    """
    from sfa_client import SFAClient

    client = SFAClient()
    updated = client.update_project_phase(project_id, PHASE_AI_DONE)
    logger.info(f"フェーズ更新: [{project_id}] {updated.name} → {PHASE_AI_DONE}")
    return updated.phase


# --- CLI テスト用 ---

if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO, format="%(message)s")

    if len(sys.argv) < 2:
        print("Usage: python sfa_uploader.py <file_path>")
        print("  WebSocket サーバーを起動し、ポート番号を表示します。")
        sys.exit(1)

    fpath = Path(sys.argv[1])
    port = start_file_server(fpath)
    print(f"\nポート: {port}")
    print(f"ファイル: {fpath.name}")
    print(f"\nJS コード（ブラウザで実行）:")
    print(get_upload_script(port, fpath.name))
    print("\nCtrl+C で停止")
    try:
        _server_stop_event.wait()
    except KeyboardInterrupt:
        stop_file_server()
