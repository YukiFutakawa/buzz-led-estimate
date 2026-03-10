"""チェックシート画像の前処理モジュール

OCR精度向上のため、撮影されたチェックシート写真を前処理する。
- EXIF回転補正（スマホ撮影の向き修正）
- コントラスト強調（手書き文字を読みやすく）
- シャープ化（筆跡を鮮明に）
- 最適リサイズ（API推奨解像度に調整）
"""

from __future__ import annotations

import io
import logging
from pathlib import Path

from PIL import Image, ImageEnhance, ImageOps

logger = logging.getLogger(__name__)

# Claude Vision APIの推奨最大解像度
# 長辺1568px以下でトークン効率が最良
MAX_LONG_EDGE = 1568


class CheckSheetPreprocessor:
    """チェックシート画像の前処理パイプライン"""

    def __init__(
        self,
        contrast_factor: float = 1.5,
        sharpness_factor: float = 2.0,
        max_long_edge: int = MAX_LONG_EDGE,
    ):
        self.contrast_factor = contrast_factor
        self.sharpness_factor = sharpness_factor
        self.max_long_edge = max_long_edge

    def preprocess(self, image_path: Path) -> tuple[bytes, str]:
        """画像を前処理してバイトデータとメディアタイプを返す

        Args:
            image_path: 入力画像パス

        Returns:
            (image_bytes, media_type) のタプル
        """
        img = Image.open(image_path)
        original_size = img.size
        logger.info(
            f"前処理開始: {image_path.name} "
            f"({img.size[0]}x{img.size[1]}, mode={img.mode})"
        )

        # Step 1: EXIF回転情報の適用
        img = self._apply_exif_rotation(img)

        # Step 2: RGBに変換（RGBA等の場合）
        if img.mode != "RGB":
            img = img.convert("RGB")

        # Step 3: コントラスト強調
        img = self._enhance_contrast(img)

        # Step 4: シャープ化
        img = self._sharpen(img)

        # Step 5: 最適解像度にリサイズ
        img = self._resize_optimal(img)

        logger.info(
            f"前処理完了: {original_size[0]}x{original_size[1]} → "
            f"{img.size[0]}x{img.size[1]}"
        )

        # PNGで出力（可逆圧縮で品質維持）
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue(), "image/png"

    def _apply_exif_rotation(self, img: Image.Image) -> Image.Image:
        """EXIF情報に基づいて画像を正しい向きに回転

        スマホで撮影した写真はEXIF Orientationタグに
        回転情報が入っている。これを適用して用紙を正立させる。
        """
        try:
            img = ImageOps.exif_transpose(img)
        except Exception as e:
            logger.debug(f"EXIF回転スキップ: {e}")
        return img

    def _enhance_contrast(self, img: Image.Image) -> Image.Image:
        """コントラストを強調して手書き文字を読みやすくする

        用紙の白とインクの黒の差を広げ、
        薄い鉛筆書きやボールペンの筆跡を明瞭にする。
        """
        enhancer = ImageEnhance.Contrast(img)
        return enhancer.enhance(self.contrast_factor)

    def _sharpen(self, img: Image.Image) -> Image.Image:
        """シャープ化で筆跡の輪郭を鮮明にする

        特に「1」と「7」、「0」と「6」など
        似た数字の区別がしやすくなる。
        """
        enhancer = ImageEnhance.Sharpness(img)
        return enhancer.enhance(self.sharpness_factor)

    def _resize_optimal(self, img: Image.Image) -> Image.Image:
        """Claude Vision APIの推奨解像度にリサイズ

        API内部でリサイズされると品質が劣化する可能性があるため、
        送信前にこちらで高品質なリサイズを行う。
        長辺が MAX_LONG_EDGE を超える場合のみ縮小。
        """
        w, h = img.size
        long_edge = max(w, h)

        if long_edge <= self.max_long_edge:
            return img

        scale = self.max_long_edge / long_edge
        new_w = int(w * scale)
        new_h = int(h * scale)

        return img.resize((new_w, new_h), Image.Resampling.LANCZOS)


def create_thumbnail(
    image_path: Path,
    max_long_edge: int = 512,
    jpeg_quality: int = 70,
) -> tuple[bytes, str]:
    """軽量サムネイル生成（画像分類用）

    チェックシート vs 器具写真の分類では画像の細部は不要。
    EXIF回転 + リサイズのみ行い、JPEG出力でトークンコスト削減。

    Args:
        image_path: 入力画像パス
        max_long_edge: サムネイルの長辺最大サイズ（デフォルト512px）
        jpeg_quality: JPEG圧縮品質（デフォルト70）

    Returns:
        (image_bytes, media_type) のタプル
    """
    img = Image.open(image_path)

    # EXIF回転補正
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass

    # RGBに変換
    if img.mode != "RGB":
        img = img.convert("RGB")

    # リサイズ（長辺がmax_long_edgeを超える場合のみ）
    w, h = img.size
    long_edge = max(w, h)
    if long_edge > max_long_edge:
        scale = max_long_edge / long_edge
        new_w = int(w * scale)
        new_h = int(h * scale)
        img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

    # JPEG出力（PNG比で約1/3のサイズ）
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=jpeg_quality)

    logger.debug(
        f"サムネイル生成: {image_path.name} "
        f"({w}x{h} → {img.size[0]}x{img.size[1]}, "
        f"{len(buf.getvalue())} bytes)"
    )
    return buf.getvalue(), "image/jpeg"
