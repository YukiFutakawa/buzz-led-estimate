"""バッチ処理レポートのメール送信

SMTPでバッチ処理結果レポートをメール送信する。
.env に SMTP設定を追加して使用。

必要な .env 設定:
    SMTP_HOST=smtp.gmail.com       # SMTPサーバー
    SMTP_PORT=587                  # ポート（587=TLS, 465=SSL）
    SMTP_USER=your@email.com       # 送信元メールアドレス
    SMTP_PASS=your_app_password    # パスワードまたはアプリパスワード
"""

from __future__ import annotations

import html as html_module
import logging
import os
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Optional

logger = logging.getLogger(__name__)

DEFAULT_RECIPIENT = "yuki.futakawa@buzzarea.co.jp"


def _get_smtp_config() -> dict:
    """環境変数からSMTP設定を読み込み"""
    return {
        "host": os.environ.get("SMTP_HOST", ""),
        "port": int(os.environ.get("SMTP_PORT", "587")),
        "user": os.environ.get("SMTP_USER", ""),
        "password": os.environ.get("SMTP_PASS", ""),
    }


def send_report_email(
    report_text: str,
    recipient: str = DEFAULT_RECIPIENT,
    subject: Optional[str] = None,
    results: Optional[list] = None,
) -> bool:
    """バッチ処理レポートをメール送信

    Args:
        report_text: レポート本文（プレーンテキスト）
        recipient: 送信先メールアドレス
        subject: 件名（Noneなら自動生成）
        results: BatchResult リスト（サマリー生成用）

    Returns:
        True: 送信成功, False: 送信失敗（設定不足含む）
    """
    config = _get_smtp_config()

    if not config["host"] or not config["user"]:
        logger.warning(
            "SMTP設定が未設定です。.env に SMTP_HOST, SMTP_USER, SMTP_PASS を追加してください。"
        )
        return False

    # 件名の自動生成
    if subject is None:
        today = datetime.now().strftime("%Y/%m/%d")
        summary = ""
        if results:
            success = sum(
                1 for r in results
                if r.get("status") == "success"
                or getattr(r, "status", None) == "success"
            )
            total = len(results)
            summary = f" ({success}/{total}件成功)"
        subject = f"LED見積バッチ処理レポート {today}{summary}"

    # メール作成
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = config["user"]
    msg["To"] = recipient

    # プレーンテキスト版
    msg.attach(MIMEText(report_text, "plain", "utf-8"))

    # HTML版（プレーンテキストを <pre> で整形）
    html_body = _text_to_html(report_text, subject)
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    # 送信
    try:
        port = config["port"]
        if port == 465:
            with smtplib.SMTP_SSL(config["host"], port, timeout=30) as server:
                server.login(config["user"], config["password"])
                server.send_message(msg)
        else:
            with smtplib.SMTP(config["host"], port, timeout=30) as server:
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login(config["user"], config["password"])
                server.send_message(msg)

        logger.info(f"レポートメール送信成功: {recipient}")
        return True

    except smtplib.SMTPAuthenticationError as e:
        logger.error(f"SMTP認証エラー: {e}")
        logger.error("SMTP_USER / SMTP_PASS を確認してください（Gmailはアプリパスワードが必要）")
        return False
    except smtplib.SMTPException as e:
        logger.error(f"SMTP送信エラー: {e}")
        return False
    except Exception as e:
        logger.error(f"メール送信エラー: {e}")
        return False


def _text_to_html(text: str, title: str) -> str:
    """プレーンテキストレポートをHTML形式に変換"""
    escaped = html_module.escape(text)

    return f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family: 'Meiryo', 'Hiragino Sans', sans-serif; padding: 20px; background: #f5f5f5;">
  <div style="max-width: 700px; margin: 0 auto; background: #fff; padding: 24px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
    <h2 style="color: #333; border-bottom: 2px solid #4CAF50; padding-bottom: 8px;">{html_module.escape(title)}</h2>
    <pre style="font-family: 'Consolas', 'MS Gothic', monospace; font-size: 13px; line-height: 1.6; white-space: pre-wrap; color: #333;">{escaped}</pre>
    <hr style="border: none; border-top: 1px solid #ddd; margin-top: 20px;">
    <p style="font-size: 11px; color: #999;">このメールはLED見積自動作成システムから自動送信されています。</p>
  </div>
</body>
</html>"""
