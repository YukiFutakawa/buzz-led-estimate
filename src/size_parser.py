"""器具サイズ文字列のパース・比較モジュール

ラインナップ表や現調データで使用される多様な寸法表記を解析し、
LED選定時の「器具跡が残らない」判定に使用する。

対応フォーマット例:
  "150×632mm"           → 幅150, 長さ632
  "150(66)×598ｍｍ"      → 幅150, 埋込66, 長さ598
  "Φ100"                → 直径100
  "Φ113(埋込100)"       → 直径113, 埋込100
  "□125mm"              → 幅125, 長さ125
  "W143：H164：D45"      → 幅143, 高さ164
  "126×636mm(埋込穴：100×617)" → 幅126, 長さ636, 埋込幅100, 埋込長617
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class FixtureDimensions:
    """パースされた器具寸法"""
    width_mm: Optional[float] = None      # 外形幅 (W方向)
    length_mm: Optional[float] = None     # 外形長さ (L/H方向、長い方)
    diameter_mm: Optional[float] = None   # 外形直径 (丸形器具)
    mount_width_mm: Optional[float] = None    # 埋込幅
    mount_length_mm: Optional[float] = None   # 埋込長さ
    mount_diameter_mm: Optional[float] = None  # 埋込穴径
    is_square: bool = False               # 正方形 (□ 表記)
    raw: str = ""                         # 元の文字列

    @property
    def has_dimensions(self) -> bool:
        """有効な寸法を持つか"""
        return any([
            self.width_mm, self.length_mm,
            self.diameter_mm,
        ])

    @property
    def is_round(self) -> bool:
        """丸形か"""
        return self.diameter_mm is not None and self.width_mm is None

    @property
    def is_rectangular(self) -> bool:
        """四角形か"""
        return self.width_mm is not None and self.diameter_mm is None

    @property
    def footprint_width(self) -> Optional[float]:
        """器具の占有幅 (跡判定に使用)"""
        if self.diameter_mm:
            return self.diameter_mm
        if self.width_mm:
            return self.width_mm
        return None

    @property
    def footprint_length(self) -> Optional[float]:
        """器具の占有長さ"""
        if self.diameter_mm:
            return self.diameter_mm
        if self.length_mm:
            return self.length_mm
        if self.is_square and self.width_mm:
            return self.width_mm
        return None

    @property
    def mount_hole_width(self) -> Optional[float]:
        """埋込穴の幅"""
        if self.mount_diameter_mm:
            return self.mount_diameter_mm
        return self.mount_width_mm

    @property
    def mount_hole_length(self) -> Optional[float]:
        """埋込穴の長さ"""
        if self.mount_diameter_mm:
            return self.mount_diameter_mm
        return self.mount_length_mm


def _normalize(text: str) -> str:
    """全角→半角、不要な空白除去"""
    result = unicodedata.normalize('NFKC', text)
    result = result.replace('　', ' ').strip()
    return result


def _extract_number(s: str) -> Optional[float]:
    """文字列から数値を抽出"""
    m = re.search(r'(\d+(?:\.\d+)?)', s)
    if m:
        return float(m.group(1))
    return None


def parse_fixture_size(raw: str) -> FixtureDimensions:
    """器具サイズ文字列をパースして FixtureDimensions を返す"""
    if not raw:
        return FixtureDimensions(raw="")

    raw_str = str(raw).strip()
    text = _normalize(raw_str)
    dims = FixtureDimensions(raw=raw_str)

    if not text or text in ('-', '-', '0', ''):
        return dims

    # ===== 埋込情報の先行抽出 =====
    _parse_mount_info(text, dims)

    # ===== メインの外形寸法パース =====

    # パターン1: □NNN (正方形)
    m = re.search(r'□\s*(\d+(?:\.\d+)?)', text)
    if m:
        dims.width_mm = float(m.group(1))
        dims.length_mm = float(m.group(1))
        dims.is_square = True
        return dims

    # パターン2: W/H/D 形式 (W143:H164:D45)
    w_match = re.search(r'W\s*(\d+(?:\.\d+)?)', text)
    if w_match and not re.match(r'^\d+.*[×xX]', text):
        # "W" で始まるか、"W:" パターンがある場合
        dims.width_mm = float(w_match.group(1))
        # H (高さ) or L (長さ)
        h_match = re.search(r'[HL]\s*(\d+(?:\.\d+)?)', text[w_match.end():])
        if h_match:
            dims.length_mm = float(h_match.group(1))
        return dims

    # パターン3: W:L 形式 (W84:L272)
    m = re.search(r'W\s*(\d+(?:\.\d+)?)\s*[:：]\s*L\s*(\d+(?:\.\d+)?)', text)
    if m:
        dims.width_mm = float(m.group(1))
        dims.length_mm = float(m.group(2))
        return dims

    # パターン4: 直径表記 (Φ100, φ100, 100Φ, 直径100)
    phi_val = _parse_phi(text)
    if phi_val is not None:
        dims.diameter_mm = phi_val
        # カバー径チェック: Φ123(カバー:W180)
        cover_match = re.search(r'カバー\s*[:：]?\s*[WΦφ]?\s*(\d+)', text, re.IGNORECASE)
        if cover_match:
            # カバー径の方が大きい場合、外形として使う
            cover = float(cover_match.group(1))
            if cover > dims.diameter_mm:
                dims.diameter_mm = cover
        return dims

    # パターン5: NNN×NNN (幅×長さ)
    m = re.search(r'(\d+(?:\.\d+)?)\s*(?:\(\d+(?:\.\d+)?\))?\s*[×xX]\s*(\d+(?:\.\d+)?)', text)
    if m:
        v1 = float(m.group(1))
        v2 = float(m.group(2))
        # 小さい方を幅、大きい方を長さとする
        dims.width_mm = min(v1, v2)
        dims.length_mm = max(v1, v2)
        # W(mount) 抽出: 150(66)×598
        mount_match = re.search(r'(\d+)\s*\((\d+(?:\.\d+)?)\)\s*[×xX]', text)
        if mount_match and dims.mount_width_mm is None:
            dims.mount_width_mm = float(mount_match.group(2))
        return dims

    # パターン6: NNNmm×NNNmm (単位付き)
    m = re.search(r'(\d+(?:\.\d+)?)\s*mm\s*[×xX]\s*(\d+(?:\.\d+)?)', text, re.IGNORECASE)
    if m:
        v1 = float(m.group(1))
        v2 = float(m.group(2))
        dims.width_mm = min(v1, v2)
        dims.length_mm = max(v1, v2)
        return dims

    # パターン7: 単一の数値 (W200, 100mm等)
    m = re.match(r'W\s*(\d+(?:\.\d+)?)\s*$', text)
    if m:
        dims.width_mm = float(m.group(1))
        return dims

    return dims


def _parse_phi(text: str) -> Optional[float]:
    """Φ/φ 系の直径表記を解析

    複数のΦ値がある場合（例: "Φ257→Φ380"リニューアル表記）、
    最大値を返す（カバー/外形サイズが跡隠しに重要）。
    """
    # ★改善: 全てのΦ値を取得し、最大値を返す（リニューアル品対応）
    # Φ100, φ100 パターン
    all_phi = re.findall(r'[Φφ]\s*(\d+(?:\.\d+)?)', text)
    if all_phi:
        return max(float(v) for v in all_phi)
    # 100Φ, 100φ パターン
    all_phi_suffix = re.findall(r'(\d+(?:\.\d+)?)\s*[Φφ]', text)
    if all_phi_suffix:
        return max(float(v) for v in all_phi_suffix)
    # 直径 100
    m = re.search(r'直径\s*(\d+(?:\.\d+)?)', text)
    if m:
        return float(m.group(1))
    # 幅：φ100
    m = re.search(r'幅\s*[:：]\s*[Φφ]\s*(\d+(?:\.\d+)?)', text)
    if m:
        return float(m.group(1))
    return None


def _parse_mount_info(text: str, dims: FixtureDimensions) -> None:
    """埋込情報を抽出して dims に設定"""

    # パターン: (埋込100×1235mm) or (埋込穴：100×617)
    m = re.search(
        r'[（(]埋込?[穴]?\s*[:：]?\s*(\d+(?:\.\d+)?)\s*[×xX]\s*(\d+(?:\.\d+)?)',
        text,
    )
    if m:
        v1 = float(m.group(1))
        v2 = float(m.group(2))
        dims.mount_width_mm = min(v1, v2)
        dims.mount_length_mm = max(v1, v2)
        return

    # パターン: (埋込Φ150) or (埋Φ150)
    m = re.search(r'[（(]埋[込]?\s*[Φφ]\s*(\d+(?:\.\d+)?)', text)
    if m:
        dims.mount_diameter_mm = float(m.group(1))
        return

    # パターン: (埋W196:H56)
    m = re.search(r'[（(]埋\s*W\s*(\d+(?:\.\d+)?)\s*[:：]\s*H\s*(\d+(?:\.\d+)?)', text)
    if m:
        dims.mount_width_mm = float(m.group(1))
        dims.mount_length_mm = float(m.group(2))
        return

    # パターン: (埋□85)
    m = re.search(r'[（(]埋[込]?\s*□\s*(\d+(?:\.\d+)?)', text)
    if m:
        dims.mount_width_mm = float(m.group(1))
        dims.mount_length_mm = float(m.group(1))
        return

    # パターン: (埋込100) or (埋込穴100) - 括弧内に数値のみ
    m = re.search(r'[（(]埋[込]?[穴]?\s*(\d+(?:\.\d+)?)\s*(?:mm|ｍｍ)?\s*[）)]', text)
    if m:
        dims.mount_diameter_mm = float(m.group(1))
        return

    # パターン: 埋込深
    m = re.search(r'埋込深\s*[:：]?\s*(\d+(?:\.\d+)?)', text)
    if m:
        # 深さは跡には影響しないが記録しておく
        pass


def is_size_compatible(
    existing: FixtureDimensions,
    led: FixtureDimensions,
    is_recessed: bool = False,
) -> tuple[bool, str]:
    """LED器具が既存器具の跡を隠せるかを判定

    Args:
        existing: 既存器具の寸法
        led: LED器具の寸法
        is_recessed: 埋込式かどうか（True=埋込穴サイズ一致必須）

    Returns:
        (compatible, reason): 適合可否と理由文字列
    """
    # 寸法情報なしの場合
    if not existing.has_dimensions:
        return True, "既存器具の寸法不明（要確認）"
    if not led.has_dimensions:
        return True, "LED器具の寸法不明（要確認）"

    # ===== 埋込式の判定 =====
    if is_recessed:
        return _check_recessed_compatibility(existing, led)

    # ===== 通常器具の跡判定 =====
    return _check_surface_compatibility(existing, led)


def _check_recessed_compatibility(
    existing: FixtureDimensions,
    led: FixtureDimensions,
) -> tuple[bool, str]:
    """埋込式器具のサイズ適合判定（穴サイズ一致必須）"""
    ex_hole_w = existing.mount_hole_width
    led_hole_w = led.mount_hole_width

    # 両方に埋込穴情報がある場合→穴サイズで判定
    if ex_hole_w is not None and led_hole_w is not None:
        w_ok = abs(ex_hole_w - led_hole_w) <= 5
        ex_hole_l = existing.mount_hole_length
        led_hole_l = led.mount_hole_length
        l_ok = True
        if ex_hole_l and led_hole_l:
            l_ok = abs(ex_hole_l - led_hole_l) <= 5

        if w_ok and l_ok:
            return True, f"埋込穴適合（既存{ex_hole_w}mm → LED{led_hole_w}mm）"
        else:
            return False, (
                f"埋込穴不適合（既存{ex_hole_w}mm vs LED{led_hole_w}mm）"
            )

    # 埋込穴情報が片方にしかない場合→外形で近似判定
    ex_d = existing.diameter_mm or existing.footprint_width
    led_d = led.diameter_mm or led.footprint_width

    # 既存の埋込穴とLED外形を比較（LED外形が穴に合えばOK）
    if ex_hole_w and led_d:
        if abs(ex_hole_w - led_d) <= 10:
            return True, f"埋込穴≈LED外形（穴{ex_hole_w}mm ≈ LED{led_d}mm）"

    # LED埋込穴と既存外形を比較
    if led_hole_w and ex_d:
        if abs(led_hole_w - ex_d) <= 10:
            return True, f"LED埋込穴≈既存外形（既存{ex_d}mm ≈ LED穴{led_hole_w}mm）"

    # 外形同士の比較（許容範囲広め）
    if ex_d and led_d:
        if abs(ex_d - led_d) <= 10:
            return True, f"外形近似（既存{ex_d}mm ≈ LED{led_d}mm）"
        else:
            return False, f"外形不一致（既存{ex_d}mm vs LED{led_d}mm）"

    return True, "埋込寸法不明（要確認）"


def _check_surface_compatibility(
    existing: FixtureDimensions,
    led: FixtureDimensions,
) -> tuple[bool, str]:
    """通常（直付/壁付）器具の跡判定（LED >= 既存ならOK）"""
    ex_w = existing.footprint_width
    ex_l = existing.footprint_length
    led_w = led.footprint_width
    led_l = led.footprint_length

    if ex_w is None or led_w is None:
        return True, "寸法不明（要確認）"

    # 幅の比較
    w_ok = led_w >= ex_w - 5  # 5mm 許容

    # 長さの比較
    l_ok = True
    l_note = ""
    if ex_l and led_l:
        l_ok = led_l >= ex_l - 5
        l_note = f", 長さ: 既存{ex_l}mm→LED{led_l}mm"
    elif ex_l and not led_l:
        l_ok = True  # LED側に長さ情報がない場合は許容
        l_note = "（LED長さ不明）"

    if w_ok and l_ok:
        return True, f"跡隠し可（幅: 既存{ex_w}mm→LED{led_w}mm{l_note}）"
    else:
        reasons = []
        if not w_ok:
            reasons.append(f"幅不足（既存{ex_w}mm > LED{led_w}mm）")
        if not l_ok:
            reasons.append(f"長さ不足（既存{ex_l}mm > LED{led_l}mm）")
        return False, "跡残りリスク: " + ", ".join(reasons)
