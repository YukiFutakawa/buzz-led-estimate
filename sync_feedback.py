"""フィードバック同期スクリプト

GitHub Actions から日次で実行され、以下を行う:
  1. Google Apps Script 経由で未同期のフィードバックを取得
  2. feedback/ フォルダに JSON ファイルとして保存
  3. FeedbackAccumulator で全 JSON を読み込み
  4. led_selection_rules.json を再生成
  5. Google Apps Script 経由で synced フラグを更新

使い方:
  # GitHub Actions (環境変数から認証)
  GAS_WEBAPP_URL='https://script.google.com/macros/s/..../exec' \
  python sync_feedback.py
"""

import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).parent
FEEDBACK_DIR = ROOT / "feedback"
SRC_DIR = ROOT / "src"

# src/ をインポートパスに追加
sys.path.insert(0, str(SRC_DIR))


def main():
    from feedback_store import FeedbackStore
    from feedback_accumulator import FeedbackAccumulator

    # 1. Google Sheets に接続
    print("Google Apps Script に接続中...")
    store = FeedbackStore.from_env()

    # 2. 未同期フィードバックを取得
    unsynced = store.get_unsynced_feedback()
    print(f"未同期フィードバック: {len(unsynced)} 件")

    if not unsynced:
        print("新しいフィードバックはありません。")
        return

    # 3. JSON ファイルとして保存
    FEEDBACK_DIR.mkdir(parents=True, exist_ok=True)
    saved_ids = []

    for record in unsynced:
        feedback_id = record.get("id", "unknown")
        prop_name = record.get("property_name", "unknown")
        timestamp = record.get("timestamp", "")

        # タイムスタンプからファイル名用の日時を生成
        try:
            dt = datetime.fromisoformat(timestamp)
            ts_str = dt.strftime("%Y%m%d_%H%M%S")
        except (ValueError, TypeError):
            ts_str = datetime.now().strftime("%Y%m%d_%H%M%S")

        filename = f"feedback_{prop_name}_{ts_str}.json"
        filepath = FEEDBACK_DIR / filename

        # 既存ファイルとの重複を避ける
        if filepath.exists():
            filename = f"feedback_{prop_name}_{ts_str}_{feedback_id}.json"
            filepath = FEEDBACK_DIR / filename

        # レコードを FeedbackComparator 互換の JSON 形式に変換
        feedback_dict = _record_to_feedback_dict(record)
        filepath.write_text(
            json.dumps(feedback_dict, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"  保存: {filename}")
        saved_ids.append(feedback_id)

    # 4. FeedbackAccumulator で全 JSON を集約 → ルール生成
    print(f"\nルール再生成中...")
    accumulator = FeedbackAccumulator(FEEDBACK_DIR)
    count = accumulator.load_all()
    print(f"  フィードバック総数: {count} 件")

    if count > 0:
        rules_path = FEEDBACK_DIR / "led_selection_rules.json"
        accumulator.export_led_rules_json(rules_path)
        print(f"  ルール出力: {rules_path}")

        # ルール数を表示
        rules = json.loads(rules_path.read_text(encoding="utf-8"))
        applicable = [r for r in rules if r.get("count", 0) >= 1]
        print(f"  総ルール数: {len(rules)} / 適用対象(count>=1): {len(applicable)}")

    # 5. 同期済みフラグを更新
    store.mark_synced(saved_ids)
    print(f"\n同期完了: {len(saved_ids)} 件を同期済みにマーク")


def _record_to_feedback_dict(record: dict) -> dict:
    """Google Sheets のレコードを feedback JSON 形式に変換"""
    fixture_diffs = []
    selection_diffs = []
    header_diffs = []

    try:
        fd_json = record.get("fixture_diffs_json", "[]")
        fixture_diffs = json.loads(fd_json) if fd_json else []
    except (json.JSONDecodeError, TypeError):
        pass

    try:
        sd_json = record.get("selection_diffs_json", "[]")
        selection_diffs = json.loads(sd_json) if sd_json else []
    except (json.JSONDecodeError, TypeError):
        pass

    try:
        hd_json = record.get("header_diffs_json", "[]")
        header_diffs = json.loads(hd_json) if hd_json else []
    except (json.JSONDecodeError, TypeError):
        pass

    return {
        "property_name": record.get("property_name", ""),
        "ai_file": "",
        "correct_file": "",
        "timestamp": record.get("timestamp", ""),
        "comment": {
            "reading": record.get("comment_reading", ""),
            "selection": record.get("comment_selection", ""),
        },
        "summary": {
            "total_diffs": int(record.get("total_diffs", 0)),
            "fixture_match_rate": float(record.get("fixture_match_rate", 0)),
            "led_selection_match_rate": float(record.get("led_match_rate", 0)),
        },
        "fixture_diffs": fixture_diffs,
        "selection_diffs": selection_diffs,
        "header_diffs": header_diffs,
    }


if __name__ == "__main__":
    main()
