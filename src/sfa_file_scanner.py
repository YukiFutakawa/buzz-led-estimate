"""SFAファイルタブ走査 & ZIPダウンロード

ブラウザ自動化でSFAプロジェクトのファイルタブをスキャンし、
現調写真ZIPのダウンロードを行う。

Chrome MCP経由で使用する。
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path

from data_source_resolver import SFAFileInfo

logger = logging.getLogger(__name__)


def get_scan_file_tab_js() -> str:
    """ファイルタブDOMからファイル情報を抽出するJavaScript

    SFAのファイルタブはDOMに以下の構造でファイル一覧が表示される:
    - .bucket_list 内の各ファイル行
    - ダウンロードリンクに file_key が含まれる

    Returns:
        JavaScript code (window._sfaFileList にファイル情報を設定)
    """
    return """
(function() {
    var files = [];
    // ファイルタブのファイル一覧を走査
    var rows = document.querySelectorAll('.bucket_file_box, .file_box, tr[data-file-key]');

    if (rows.length === 0) {
        // 別のDOM構造を試す
        var links = document.querySelectorAll('a[href*="download_bucket_file"], a[href*="file_key"]');
        for (var i = 0; i < links.length; i++) {
            var href = links[i].getAttribute('href') || '';
            var m = href.match(/file_key=([^&]+)/);
            if (m) {
                files.push({
                    file_key: m[1],
                    filename: links[i].textContent.trim() || ('file_' + i)
                });
            }
        }
    }

    if (files.length === 0) {
        // さらに別の構造: .scrollbox 内のファイル
        var scrollbox = document.querySelector('.box_area.section.bucket .scrollbox');
        if (scrollbox) {
            var allLinks = scrollbox.querySelectorAll('a');
            for (var j = 0; j < allLinks.length; j++) {
                var h = allLinks[j].getAttribute('href') || '';
                var onclick = allLinks[j].getAttribute('onclick') || '';
                var fk = null;

                // href から file_key を抽出
                var hm = h.match(/file_key=([^&]+)/);
                if (hm) fk = hm[1];

                // onclick から file_key を抽出
                if (!fk) {
                    var om = onclick.match(/file_key['":\s]*['"]?([^'"&\s,)]+)/);
                    if (om) fk = om[1];
                }

                if (fk) {
                    files.push({
                        file_key: fk,
                        filename: allLinks[j].textContent.trim() || ('file_' + j)
                    });
                }
            }
        }
    }

    // 重複排除
    var seen = {};
    var unique = [];
    for (var k = 0; k < files.length; k++) {
        if (!seen[files[k].file_key]) {
            seen[files[k].file_key] = true;
            unique.push(files[k]);
        }
    }

    window._sfaFileList = JSON.stringify(unique);
    return unique.length + ' files found';
})();
"""


def get_file_tab_click_js() -> str:
    """ファイルタブをクリックするJavaScript"""
    return """
(function() {
    // タブメニューから「ファイル」を探してクリック
    var tabs = document.querySelectorAll('.tab_menu a, .detail_tab a, [role="tab"] a');
    for (var i = 0; i < tabs.length; i++) {
        if (tabs[i].textContent.trim().includes('ファイル')) {
            tabs[i].click();
            return 'clicked';
        }
    }
    // 左サイドバーの「ファイル」リンク
    var sideLinks = document.querySelectorAll('.side_menu a, nav a');
    for (var j = 0; j < sideLinks.length; j++) {
        if (sideLinks[j].textContent.trim() === 'ファイル') {
            sideLinks[j].click();
            return 'sidebar_clicked';
        }
    }
    return 'not_found';
})();
"""


def parse_file_list(
    json_str: str,
    project_id: str,
    project_name: str,
) -> list[SFAFileInfo]:
    """JSスキャン結果 → SFAFileInfo リスト"""
    import json

    try:
        raw = json.loads(json_str)
    except (json.JSONDecodeError, TypeError):
        return []

    results = []
    for item in raw:
        results.append(SFAFileInfo(
            file_key=item.get("file_key", ""),
            filename=item.get("filename", ""),
            project_id=project_id,
            project_name=project_name,
        ))
    return results
