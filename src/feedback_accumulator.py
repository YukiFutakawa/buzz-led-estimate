# -*- coding: utf-8 -*-
"""フィードバック蓄積・パターン分析

複数のフィードバックJSONを集約し、繰り返し発生するエラーパターンを特定する。
改善ルール（fixture_type → 正しいLED選定）を自動生成する。

Usage:
    from feedback_accumulator import FeedbackAccumulator
    acc = FeedbackAccumulator(feedback_dir)
    acc.load_all()
    print(acc.generate_improvement_report())
"""

from __future__ import annotations

import json
import logging
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


@dataclass
class ErrorPattern:
    """繰り返し発生するエラーパターン"""
    field_name: str          # "LED選定(G列)", "電球数(L列)" etc.
    severity: str            # "critical", "major", "minor"
    ai_value: str            # AIが出力した値
    correct_value: str       # 正解の値
    count: int = 1           # 出現回数
    fixture_types: list[str] = field(default_factory=list)
    source_files: list[str] = field(default_factory=list)


@dataclass
class LEDSelectionRule:
    """LED選定の修正ルール（フィードバックから学習）

    「この器具タイプにはこのLED」というマッピング。
    """
    fixture_type: str        # 照明種別
    wrong_selection: str     # AIが選んだ間違いLED
    correct_selection: str   # 正解のLED
    count: int = 1           # この修正が出現した回数
    source_files: list[str] = field(default_factory=list)


@dataclass
class QuantityErrorPattern:
    """数量エラーのパターン"""
    fixture_type: str
    ai_quantity: str
    correct_quantity: str
    count: int = 1
    source_files: list[str] = field(default_factory=list)


class FeedbackAccumulator:
    """フィードバックJSONの集約・分析エンジン"""

    def __init__(self, feedback_dir: Path):
        self.feedback_dir = feedback_dir
        self.reports: list[dict] = []

        # 集約結果
        self.led_selection_rules: list[LEDSelectionRule] = []
        self.error_patterns: list[ErrorPattern] = []
        self.quantity_errors: list[QuantityErrorPattern] = []

    def load_all(self) -> int:
        """feedbackフォルダ内の全JSONを読み込み"""
        if not self.feedback_dir.exists():
            logger.warning(f"フィードバックフォルダなし: {self.feedback_dir}")
            return 0

        self.reports = []
        for json_file in sorted(self.feedback_dir.glob("feedback_*.json")):
            try:
                data = json.loads(json_file.read_text(encoding="utf-8"))
                self.reports.append(data)
            except Exception as e:
                logger.warning(f"読込エラー {json_file.name}: {e}")

        if self.reports:
            self._analyze()
            logger.info(
                f"フィードバック{len(self.reports)}件を読み込み・分析完了")

        return len(self.reports)

    def _analyze(self) -> None:
        """全レポートを分析してパターンを抽出"""
        self._extract_led_rules()
        self._extract_error_patterns()
        self._extract_quantity_errors()

    # ----------------------------------------------------------
    # LED選定ルール抽出
    # ----------------------------------------------------------

    def _extract_led_rules(self) -> None:
        """G列（LED選定）の差分からルールを生成"""
        # (fixture_type, wrong, correct) → 出現データ
        rule_map: dict[tuple, dict] = defaultdict(
            lambda: {"count": 0, "files": []})

        for report in self.reports:
            source = report.get("ai_file", "")
            for fd in report.get("fixture_diffs", []):
                if fd["status"] != "modified":
                    continue
                for diff in fd.get("diffs", []):
                    if diff["field_name"] != "LED選定(G列)":
                        continue
                    key = (
                        fd["fixture_type_correct"] or fd["fixture_type_ai"],
                        diff["ai_value"],
                        diff["correct_value"],
                    )
                    rule_map[key]["count"] += 1
                    rule_map[key]["files"].append(source)

        self.led_selection_rules = []
        for (ft, wrong, correct), data in rule_map.items():
            if not correct:
                continue
            self.led_selection_rules.append(LEDSelectionRule(
                fixture_type=ft,
                wrong_selection=wrong,
                correct_selection=correct,
                count=data["count"],
                source_files=data["files"],
            ))

        # 出現回数でソート
        self.led_selection_rules.sort(key=lambda r: -r.count)

    # ----------------------------------------------------------
    # エラーパターン抽出
    # ----------------------------------------------------------

    def _extract_error_patterns(self) -> None:
        """全差分からエラーパターンを集約"""
        # (field_name, ai_value, correct_value) → 出現データ
        pattern_map: dict[tuple, dict] = defaultdict(
            lambda: {"count": 0, "severity": "", "fixtures": [], "files": []})

        for report in self.reports:
            source = report.get("ai_file", "")

            # ヘッダー差分
            for diff in report.get("header_diffs", []):
                key = (diff["field_name"], diff["ai_value"],
                       diff["correct_value"])
                pattern_map[key]["count"] += 1
                pattern_map[key]["severity"] = diff.get("severity", "unknown")
                pattern_map[key]["files"].append(source)

            # 器具行差分
            for fd in report.get("fixture_diffs", []):
                ft = fd.get("fixture_type_correct") or fd.get(
                    "fixture_type_ai", "")
                for diff in fd.get("diffs", []):
                    key = (diff["field_name"], diff["ai_value"],
                           diff["correct_value"])
                    pattern_map[key]["count"] += 1
                    pattern_map[key]["severity"] = diff.get("severity", "unknown")
                    pattern_map[key]["fixtures"].append(ft)
                    pattern_map[key]["files"].append(source)

            # 選定差分
            for sd in report.get("selection_diffs", []):
                for diff in sd.get("diffs", []):
                    key = (diff["field_name"], diff["ai_value"],
                           diff["correct_value"])
                    pattern_map[key]["count"] += 1
                    pattern_map[key]["severity"] = diff.get("severity", "unknown")
                    pattern_map[key]["files"].append(source)

        self.error_patterns = []
        for (field_name, ai_val, correct_val), data in pattern_map.items():
            self.error_patterns.append(ErrorPattern(
                field_name=field_name,
                severity=data["severity"],
                ai_value=ai_val,
                correct_value=correct_val,
                count=data["count"],
                fixture_types=list(set(data["fixtures"])),
                source_files=list(set(data["files"])),
            ))

        self.error_patterns.sort(key=lambda p: (-_severity_rank(p.severity),
                                                -p.count))

    # ----------------------------------------------------------
    # 数量エラー抽出
    # ----------------------------------------------------------

    def _extract_quantity_errors(self) -> None:
        """電球数(L列)のエラーパターンを抽出"""
        qty_map: dict[tuple, dict] = defaultdict(
            lambda: {"count": 0, "files": []})

        for report in self.reports:
            source = report.get("ai_file", "")
            for fd in report.get("fixture_diffs", []):
                if fd["status"] != "modified":
                    continue
                ft = fd.get("fixture_type_correct") or fd.get(
                    "fixture_type_ai", "")
                for diff in fd.get("diffs", []):
                    if "数量" not in diff["field_name"] and \
                       "電球数" not in diff["field_name"]:
                        continue
                    key = (ft, diff["ai_value"], diff["correct_value"])
                    qty_map[key]["count"] += 1
                    qty_map[key]["files"].append(source)

        self.quantity_errors = []
        for (ft, ai_qty, correct_qty), data in qty_map.items():
            self.quantity_errors.append(QuantityErrorPattern(
                fixture_type=ft,
                ai_quantity=ai_qty,
                correct_quantity=correct_qty,
                count=data["count"],
                source_files=list(set(data["files"])),
            ))
        self.quantity_errors.sort(key=lambda q: -q.count)

    # ----------------------------------------------------------
    # レポート生成
    # ----------------------------------------------------------

    def generate_improvement_report(self) -> str:
        """改善提案レポートを生成"""
        lines = []
        lines.append("=" * 60)
        lines.append("  フィードバック蓄積レポート")
        lines.append("=" * 60)
        lines.append(f"\n  分析対象: {len(self.reports)}件のフィードバック")

        # 全体統計
        total_diffs = sum(
            r.get("summary", {}).get("total_diffs", 0) for r in self.reports)
        avg_led_rate = 0
        led_rates = [
            r.get("summary", {}).get("led_selection_match_rate", 0)
            for r in self.reports
        ]
        if led_rates:
            avg_led_rate = sum(led_rates) / len(led_rates)

        lines.append(f"  累計差分セル数: {total_diffs}")
        lines.append(f"  平均LED選定一致率: {avg_led_rate:.1%}")

        # Section 1: LED選定の修正ルール
        if self.led_selection_rules:
            lines.append(f"\n{'='*60}")
            lines.append("  1. LED選定の修正ルール（頻出順）")
            lines.append(f"{'='*60}")
            for i, rule in enumerate(self.led_selection_rules[:20], 1):
                lines.append(f"\n  [{i}] {rule.fixture_type}")
                lines.append(f"      誤: {rule.wrong_selection or '(空)'}")
                lines.append(f"      正: {rule.correct_selection}")
                lines.append(f"      出現: {rule.count}回")

        # Section 2: 頻出エラーパターン（criticalのみ）
        critical_patterns = [
            p for p in self.error_patterns
            if p.severity == "critical" and p.count >= 2
        ]
        if critical_patterns:
            lines.append(f"\n{'='*60}")
            lines.append("  2. 重大エラーパターン（2回以上）")
            lines.append(f"{'='*60}")
            for p in critical_patterns[:15]:
                lines.append(f"\n  [{p.count}回] {p.field_name}")
                lines.append(f"      AI:  {p.ai_value or '(空)'}")
                lines.append(f"      正解: {p.correct_value}")
                if p.fixture_types:
                    lines.append(f"      器具: {', '.join(p.fixture_types[:3])}")

        # Section 3: 数量エラー
        if self.quantity_errors:
            lines.append(f"\n{'='*60}")
            lines.append("  3. 数量エラーパターン")
            lines.append(f"{'='*60}")
            for q in self.quantity_errors[:10]:
                lines.append(f"\n  [{q.count}回] {q.fixture_type}")
                lines.append(f"      AI数量:  {q.ai_quantity or '(空)'}")
                lines.append(f"      正解数量: {q.correct_quantity}")

        # Section 4: 物件別サマリー
        lines.append(f"\n{'='*60}")
        lines.append("  4. 物件別サマリー")
        lines.append(f"{'='*60}")
        for r in self.reports:
            summary = r.get("summary", {})
            name = r.get("property_name", "?")
            led_rate = summary.get("led_selection_match_rate", 0)
            total = summary.get("total_diffs", 0)
            lines.append(f"  {name}: LED一致率={led_rate:.0%}, 差分={total}件")

        return "\n".join(lines)

    def export_led_rules_json(self, output_path: Path) -> None:
        """LED選定ルールをJSON形式でエクスポート

        led_matcher.py の改善に直接利用可能な形式。
        """
        rules = []
        for rule in self.led_selection_rules:
            rules.append({
                "fixture_type": rule.fixture_type,
                "wrong_selection": rule.wrong_selection,
                "correct_selection": rule.correct_selection,
                "count": rule.count,
            })

        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(
            json.dumps(rules, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        logger.info(f"LED選定ルール出力: {output_path} ({len(rules)}件)")


def _severity_rank(severity: str) -> int:
    return {"critical": 3, "major": 2, "minor": 1}.get(severity, 0)


# ============================================================
# CLI テスト
# ============================================================

if __name__ == "__main__":
    import sys
    import io
    sys.stdout = io.TextIOWrapper(
        sys.stdout.buffer, encoding="utf-8", errors="replace")

    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    base = Path(__file__).parent.parent
    feedback_dir = base / "feedback"

    acc = FeedbackAccumulator(feedback_dir)
    count = acc.load_all()

    if count == 0:
        print("フィードバックデータがありません。")
        print(f"  フォルダ: {feedback_dir}")
        print("  先にフィードバック比較を実行してください。")
    else:
        print(acc.generate_improvement_report())

        # LED選定ルールをJSON出力
        rules_path = feedback_dir / "led_selection_rules.json"
        acc.export_led_rules_json(rules_path)
