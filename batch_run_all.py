#!/usr/bin/env python3
"""バッチ処理: SFA ZIPダウンロード → パイプライン → Excel生成

ブラウザから WebSocket 経由で ZIP を順番に受信し、
各物件のパイプラインを自動実行する。

Usage:
    python batch_run_all.py
    # → ポート19908で待機。ブラウザ側JSが順番にZIPを送信する。
"""

from __future__ import annotations

import asyncio
import base64
import json
import logging
import os
import sys
import traceback
from pathlib import Path

# src/ をパスに追加
ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT / "src"))

from batch_processor import extract_zip, load_env

logger = logging.getLogger(__name__)

OUTPUT_DIR = ROOT / "output"
TARGETS_FILE = OUTPUT_DIR / "sfa_batch_targets.json"
RESULTS_FILE = OUTPUT_DIR / "sfa_batch_results.json"
LINEUP_DIR = ROOT / "ラインナップ表"
TEMPLATE_DIR = ROOT / "見積りテンプレート"
PORT = 19908


def load_results() -> list[dict]:
    if RESULTS_FILE.exists():
        return json.loads(RESULTS_FILE.read_text(encoding="utf-8"))
    return []


def save_results(results: list[dict]):
    RESULTS_FILE.write_text(
        json.dumps(results, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def is_already_done(name: str, results: list[dict]) -> bool:
    output = OUTPUT_DIR / f"【LED導入ｼﾐｭﾚｰｼｮﾝ】{name}.xlsx"
    if output.exists():
        return True
    for r in results:
        if r.get("name") == name and r.get("status") == "success":
            return True
    return False


def _run_pipeline_sync(survey_dir: Path, api_key: str | None = None) -> Path:
    """パイプライン実行（同期）"""
    from pipeline import run_pipeline as _run

    return _run(
        survey_dir=survey_dir,
        lineup_dir=LINEUP_DIR,
        template_dir=TEMPLATE_DIR,
        template_name="田村基本形",
        api_key=api_key,
    )


def process_one_sync(project_id: str, name: str, zip_bytes: bytes) -> dict:
    """1物件を処理（同期・スレッドで呼ばれる）"""
    try:
        survey_dir = extract_zip(zip_bytes, name)
        result_path = _run_pipeline_sync(survey_dir)
        return {
            "id": project_id,
            "name": name,
            "status": "success",
            "output": result_path.name,
        }
    except Exception as e:
        logger.error(f"失敗: {name} - {e}")
        logger.debug(traceback.format_exc())
        return {
            "id": project_id,
            "name": name,
            "status": "error",
            "error": str(e),
        }


async def main():
    import websockets

    load_env()
    results = load_results()
    count = {"success": 0, "error": 0, "skipped": 0}

    async def handler(ws):
        try:
            raw = await ws.recv()
            payload = json.loads(raw)

            pid = payload.get("id", "?")
            name = payload.get("name", "?")
            b64_data = payload.get("data", "")

            # 処理済みチェック
            if is_already_done(name, results):
                await ws.send(json.dumps({
                    "status": "skipped",
                    "message": f"{name}: 既に処理済み",
                }))
                count["skipped"] += 1
                logger.info(f"[SKIP] {name}")
                return

            zip_bytes = base64.b64decode(b64_data)
            logger.info(f"受信: {name} (ID:{pid}, {len(zip_bytes):,} bytes)")

            # パイプライン実行（スレッドで非同期化）
            result = await asyncio.to_thread(
                process_one_sync, pid, name, zip_bytes,
            )

            results.append(result)
            save_results(results)

            if result["status"] == "success":
                count["success"] += 1
                logger.info(f"[OK] {name} → {result['output']}")
            else:
                count["error"] += 1
                logger.info(f"[NG] {name} → {result['error']}")

            await ws.send(json.dumps(result))

        except Exception as e:
            logger.error(f"ハンドラエラー: {e}")
            traceback.print_exc()
            try:
                await ws.send(json.dumps({
                    "status": "error",
                    "message": str(e),
                }))
            except Exception:
                pass

    server = await websockets.serve(
        handler,
        "127.0.0.1",
        PORT,
        max_size=50 * 1024 * 1024,
    )
    actual_port = server.sockets[0].getsockname()[1]

    logger.info("=" * 60)
    logger.info(f"  LED見積バッチサーバー (port {actual_port})")
    logger.info(f"  既存結果: {len(results)}件")
    logger.info(f"  ブラウザからZIPを順番に送信してください")
    logger.info("=" * 60)

    try:
        while True:
            await asyncio.sleep(1)
    except (KeyboardInterrupt, asyncio.CancelledError):
        pass
    finally:
        server.close()
        await server.wait_closed()
        logger.info(
            f"\n完了: 成功={count['success']}, "
            f"失敗={count['error']}, "
            f"スキップ={count['skipped']}"
        )


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    asyncio.run(main())
