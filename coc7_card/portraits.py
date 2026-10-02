"""Bound portrait decoding and normalize images shared by import and export."""

import io
import warnings
from contextlib import contextmanager

from PIL import Image, ImageOps


MAX_PORTRAIT_PIXELS = 20_000_000
PORTRAIT_MAX_EDGE = 1600


class PortraitError(ValueError):
    pass


@contextmanager
def _open_portrait(data: bytes, allow_legacy: bool):
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("error", Image.DecompressionBombWarning)
            with Image.open(io.BytesIO(data)) as image:
                formats = {"PNG", "JPEG", "WEBP", "GIF", "BMP"} if allow_legacy else {"PNG", "JPEG", "WEBP"}
                if image.format not in formats:
                    raise PortraitError("头像图片格式不受支持，请选择 PNG、JPEG 或 WebP。")
                if image.width * image.height > MAX_PORTRAIT_PIXELS:
                    raise PortraitError("头像不能超过 2000 万像素，请缩小图片后重试。")
                yield image
    except (Image.DecompressionBombError, Image.DecompressionBombWarning) as exc:
        raise PortraitError("头像不能超过 2000 万像素，请缩小图片后重试。") from exc
    except OSError as exc:
        raise PortraitError("头像无法读取，请选择有效的图片文件。") from exc


def validate_portrait(data: bytes, *, allow_legacy: bool = False) -> None:
    """Inspect the real format and dimensions without decoding image pixels."""
    with _open_portrait(data, allow_legacy):
        pass


def prepare_portrait(data: bytes, *, allow_legacy: bool = False) -> Image.Image:
    """Return an owned RGBA image; callers close it after use."""
    with _open_portrait(data, allow_legacy) as image:
        # JPEG can downsample during decoding. Other formats decode only after
        # the pixel limit; resize before making orientation / colour copies.
        image.thumbnail((PORTRAIT_MAX_EDGE, PORTRAIT_MAX_EDGE), Image.Resampling.LANCZOS)
        with ImageOps.exif_transpose(image) as oriented:
            return oriented.convert("RGBA")
