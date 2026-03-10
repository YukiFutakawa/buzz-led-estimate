"""SFA履歴情報 添付写真DL

ブラウザ自動化でSFAの履歴情報に添付されたファイル（写真）を
DLする。Chrome MCP経由で使用。
"""

from __future__ import annotations

import asyncio
import base64
import json
import logging
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


@dataclass
class HistoryFileInfo:
    """履歴添付ファイル情報"""
    file_key: str
    filename: str
    history_id: str = ""
    history_date: str = ""


def get_history_tab_click_js() -> str:
    """履歴情報タブをクリックするJavaScript"""
    return (
        "(function() {"
        "var tabs = jQuery('.tab_menu a, .detail_tab a');"
        "for (var i = 0; i < tabs.length; i++) {"
        "  if (jQuery(tabs[i]).text().trim().indexOf('\u5c65\u6b74\u60c5\u5831') >= 0) {"
        "    jQuery(tabs[i]).trigger('click');"
        "    return 'clicked';"
        "  }"
        "}"
        "return 'not_found';"
        "})();"
    )


def get_scan_history_attachments_js() -> str:
    """履歴エントリの添付ファイル一覧をDOM走査するJavaScript"""
    return """
(function() {
    var files = [];
    var links = document.querySelectorAll('a[href*="download_bucket_file"], a[href*="file_key"]');
    for (var i = 0; i < links.length; i++) {
        var href = links[i].getAttribute('href') || '';
        var m = href.match(/file_key=([^&]+)/);
        if (m) {
            var fn = links[i].textContent.trim() || links[i].getAttribute('title') || '';
            if (!fn) {
                var img = links[i].querySelector('img');
                if (img) fn = img.getAttribute('alt') || '';
                if (!fn) fn = 'file_' + i;
            }
            files.push({ file_key: m[1], filename: fn });
        }
    }
    if (files.length === 0) {
        var allLinks = document.querySelectorAll('a[onclick*="file_key"], a[onclick*="download"]');
        for (var j = 0; j < allLinks.length; j++) {
            var onclick = allLinks[j].getAttribute('onclick') || '';
            var om = onclick.match(/file_key[=:'"\\s]*['"]?([^'"&\\s,)]+)/);
            if (om) {
                files.push({ file_key: om[1], filename: allLinks[j].textContent.trim() || ('file_' + j) });
            }
        }
    }
    if (files.length === 0) {
        var imgs = document.querySelectorAll('.history_detail img, .memo_area img, .comment_area img, .scrollbox img');
        for (var k = 0; k < imgs.length; k++) {
            var src = imgs[k].getAttribute('src') || imgs[k].getAttribute('data-src') || '';
            var im = src.match(/file_key=([^&]+)/);
            if (im) {
                files.push({ file_key: im[1], filename: imgs[k].getAttribute('alt') || ('photo_' + k + '.jpg') });
            }
        }
    }
    var seen = {};
    var unique = [];
    for (var n = 0; n < files.length; n++) {
        if (!seen[files[n].file_key]) {
            seen[files[n].file_key] = true;
            unique.push(files[n]);
        }
    }
    window._sfaHistoryFiles = JSON.stringify(unique);
    return unique.length + ' files found';
})();
"""


def get_download_file_js(file_key: str, filename: str, ws_port: int) -> str:
    """個別ファイルDL用JavaScript (XHR+WebSocket)"""
    safe_fn = filename.replace("'", "\'")
    js = "(function() {"
    js += "var url = '/ajaxBuckets/download_bucket_file?file_' + 'key='"
    js += " + encodeURIComponent('" + file_key + "') + '&from_business_card=';"
    js += "var xhr = new XMLHttpRequest();"
    js += "xhr.open('GET', url, true);"
    js += "xhr.responseType = 'arraybuffer';"
    js += "xhr.onload = function() {"
    js += "  if (xhr.status !== 200) { window._fileDlResult = 'HTTP error: ' + xhr.status; return; }"
    js += "  var bytes = new Uint8Array(xhr.response);"
    js += "  var chunkSize = 32768; var b64parts = [];"
    js += "  for (var i = 0; i < bytes.length; i += chunkSize) {"
    js += "    var chunk = bytes.subarray(i, Math.min(i + chunkSize, bytes.length));"
    js += "    b64parts.push(String.fromCharCode.apply(null, chunk));"
    js += "  }"
    js += "  var b64 = btoa(b64parts.join(''));"
    js += "  var ws = new WebSocket('ws://127.0.0.1:" + str(ws_port) + "');"
    js += "  ws.onopen = function() {"
    js += "    ws.send(JSON.stringify({name: '" + safe_fn + "', data: b64}));"
    js += "  };"
    js += "  ws.onmessage = function(ev) { window._fileDlResult = 'OK: ' + bytes.length + ' bytes'; };"
    js += "  ws.onerror = function() { window._fileDlResult = 'WebSocket error'; };"
    js += "};"
    js += "xhr.onerror = function() { window._fileDlResult = 'XHR error'; };"
    js += "xhr.send();"
    js += "window._fileDlResult = 'downloading...';"
    js += "})();"
    return js


def parse_history_file_list(json_str: str) -> list[HistoryFileInfo]:
    """JSスキャン結果 -> HistoryFileInfo リスト"""
    try:
        raw = json.loads(json_str)
    except (json.JSONDecodeError, TypeError):
        return []
    return [
        HistoryFileInfo(
            file_key=item.get("file_key", ""),
            filename=item.get("filename", ""),
            history_id=item.get("history_id", ""),
            history_date=item.get("history_date", ""),
        )
        for item in raw
    ]


def receive_file_from_browser(
    port: int = 0,
    timeout_sec: int = 60,
    max_size_mb: int = 20,
) -> tuple[int, list, threading.Thread]:
    """ブラウザからWebSocket経由でファイルデータを受信

    Returns:
        (actual_port, result_holder, thread)
        result_holder[0] = file_bytes, result_holder[1] = filename
    """
    import websockets

    result = [None, None]
    actual_port_holder = [0]
    ready_event = threading.Event()

    async def handler(ws):
        msg = await ws.recv()
        payload = json.loads(msg)
        result[0] = base64.b64decode(payload["data"])
        result[1] = payload["name"]
        await ws.send("OK")
        logger.info("received: %s (%d bytes)", result[1], len(result[0]))

    async def main():
        server = await websockets.serve(
            handler, "127.0.0.1", port,
            max_size=max_size_mb * 1024 * 1024,
        )
        actual_port_holder[0] = server.sockets[0].getsockname()[1]
        logger.info("file receiver: ws://127.0.0.1:%d", actual_port_holder[0])
        ready_event.set()
        for _ in range(timeout_sec * 2):
            await asyncio.sleep(0.5)
            if result[0] is not None:
                break
        server.close()
        await server.wait_closed()

    def _run():
        asyncio.run(main())

    thread = threading.Thread(target=_run, daemon=True)
    thread.start()
    ready_event.wait(timeout=10)
    return actual_port_holder[0], result, thread


def save_history_photos(
    files: list[tuple[bytes, str]],
    project_name: str,
    base_dir: Optional[Path] = None,
) -> list[Path]:
    """ダウンロードした写真をローカルに保存"""
    if base_dir is None:
        base_dir = Path(__file__).parent.parent / "現調写真"

    save_dir = base_dir / project_name
    save_dir.mkdir(parents=True, exist_ok=True)

    saved = []
    for file_bytes, filename in files:
        if not file_bytes:
            continue
        safe_name = filename.replace("/", "_").replace("\\", "_")
        if not safe_name:
            safe_name = "photo_%d.jpg" % (len(saved) + 1)
        filepath = save_dir / safe_name
        if filepath.exists():
            stem = filepath.stem
            suffix = filepath.suffix
            for i in range(1, 100):
                filepath = save_dir / ("%s_%d%s" % (stem, i, suffix))
                if not filepath.exists():
                    break
        filepath.write_bytes(file_bytes)
        saved.append(filepath)
        logger.info("saved: %s", filepath)

    logger.info("photos saved: %d files -> %s", len(saved), save_dir)
    return saved


SFA_BASE_URL = "https://nsfa.next-cloud.jp"


def get_project_url(project_id: str) -> str:
    """物件詳細ページのURL"""
    return "%s/projects/detail/%s" % (SFA_BASE_URL, project_id)
