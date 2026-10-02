"""Portrait limits precede decoding and preserve the smaller image's content."""

import io
import json
import struct
import zlib

import pytest
from fastapi.testclient import TestClient
from PIL import Image, PngImagePlugin

from app import app, get_catalog
from coc7_card.exporters.xlsx_template import TemplateWorkbook
from coc7_card.portraits import PortraitError, prepare_portrait, validate_portrait


def portrait(size=(32, 16), mode="RGBA", color=(20, 80, 120, 96), **save_options):
    output = io.BytesIO()
    with Image.new(mode, size, color) as image:
        image.save(output, **{"format": "PNG", **save_options})
    return output.getvalue()


def png_header(width, height):
    # A tiny PNG with changed IHDR dimensions lets tests detect any pixel decode.
    data = portrait()
    header = struct.pack(">II", width, height) + data[24:29]
    return data[:16] + header + struct.pack(">I", zlib.crc32(b"IHDR" + header)) + data[33:]


def test_twenty_megapixel_boundary_is_checked_without_decoding(monkeypatch):
    def unexpected_decode(*_args, **_kwargs):
        pytest.fail("Pixel buffers must not be allocated while checking dimensions")

    monkeypatch.setattr(PngImagePlugin.PngImageFile, "load", unexpected_decode)
    validate_portrait(png_header(5000, 4000))
    with pytest.raises(PortraitError, match="2000 万像素"):
        validate_portrait(png_header(5001, 4000))
    with pytest.raises(PortraitError, match="2000 万像素"):
        prepare_portrait(png_header(5001, 4000))


def test_resize_preserves_alpha_aspect_and_original_bytes():
    data = portrait((3200, 1600))
    with prepare_portrait(data) as output:
        assert output.size == (1600, 800)
        assert output.mode == "RGBA"
        assert output.getpixel((400, 400)) == pytest.approx((20, 80, 120, 96), abs=2)
    with Image.open(io.BytesIO(data)) as original:
        assert original.size == (3200, 1600)
    with prepare_portrait(portrait()) as small:
        assert small.size == (32, 16)  # Small portraits are not enlarged.


def test_jpeg_orientation_is_applied_after_resize():
    exif = Image.Exif()
    exif[274] = 6
    data = portrait((3200, 1600), "RGB", (20, 80, 120), format="JPEG", exif=exif)
    with prepare_portrait(data) as output:
        assert output.size == (800, 1600)
        assert output.getexif().get(274) is None


def test_legacy_import_formats_are_not_accepted_as_direct_uploads():
    data = portrait(mode="RGB", color=(20, 80, 120), format="BMP")
    with pytest.raises(PortraitError, match="格式不受支持"):
        validate_portrait(data)
    with prepare_portrait(data, allow_legacy=True) as image:
        assert image.size == (32, 16)


@pytest.mark.parametrize("path", ["/api/export/excel", "/api/export/pdf"])
@pytest.mark.parametrize("data,detail", [(png_header(5001, 4000), "2000 万像素"),
                                         (b"not an image", "有效的图片")])
def test_invalid_portrait_is_rejected_before_export(monkeypatch, path, data, detail):
    import app as web

    async def unexpected_export(*_args):
        pytest.fail("Invalid portraits must be rejected before entering an exporter")

    monkeypatch.setattr(web, "_run_heavy", unexpected_export)
    payload = {"identity": {"name": "头像边界测试", "age": 30}}
    with TestClient(app) as client:
        result = client.post(path, data={"draft_json": json.dumps(payload)},
                             files={"portrait": ("portrait.png", data, "image/png")})
    assert result.status_code == 422
    assert detail in result.json()["detail"]


def test_xlsx_embeds_resized_portrait_with_transparency():
    workbook = TemplateWorkbook(get_catalog().template_path)
    try:
        workbook.add_portrait(workbook.Worksheets("人物卡"), portrait((3200, 1600)))
        embedded = workbook.parts["xl/media/coc7-investigator-portrait.png"]
        with Image.open(io.BytesIO(embedded)) as image:
            assert image.size == (1600, 800)
            assert image.getextrema()[3] == (96, 96)
    finally:
        workbook.close()
