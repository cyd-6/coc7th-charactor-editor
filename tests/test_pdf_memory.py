"""Rendering closes native documents and releases their global image cache."""

import io

import fitz
import pytest
from PIL import Image

from app import FONT_PATH, get_catalog
from coc7_card.exporters.pdf import PdfExporter
from coc7_card.web import build_draft


@pytest.fixture
def native_cleanup(monkeypatch):
    calls = []
    original = fitz.TOOLS.store_shrink

    def shrink(percent):
        result = original(percent)
        calls.append(percent)
        return result

    monkeypatch.setattr(fitz.TOOLS, "store_shrink", shrink)
    return calls


def test_repeated_preview_keeps_two_pages_and_releases_native_store(native_cleanup):
    catalog = get_catalog()
    occupation = next(item for item in catalog.occupations if item.credit_min == 0)
    draft = build_draft({"identity": {"name": "预览内存验证", "age": 30},
                         "occupation_id": occupation.occupation_id}, catalog)
    result = PdfExporter(FONT_PATH).export(draft, include_preview=False)
    previous = None
    for index in range(3):
        previews = PdfExporter.render_preview_images(result.data)
        assert len(previews) == 2
        for page in previews:
            with Image.open(io.BytesIO(page)) as image:
                assert image.format == "PNG"
                assert image.width > 800 and image.height > 1000
        # This PyMuPDF build exposes no current store-size value; exercise the
        # real cleanup and assert it runs after every completed render.
        assert native_cleanup == [100] * (index + 1)
        if previous is not None:
            assert previews == previous
        previous = previews


def test_failed_preview_releases_cache_and_allows_next_render(monkeypatch, native_cleanup):
    with fitz.open() as document:
        for _ in range(2):
            page = document.new_page()
            with Image.new("RGB", (800, 800), (20, 80, 120)) as image:
                output = io.BytesIO()
                image.save(output, format="PNG")
            page.insert_image(fitz.Rect(0, 0, 100, 100), stream=output.getvalue())
        data = document.tobytes()
    original = fitz.Page.get_pixmap
    opened_documents = []

    def fail_second_page(page, *args, **kwargs):
        if page.number == 1:
            opened_documents.append(page.parent)
            raise RuntimeError("synthetic render failure")
        return original(page, *args, **kwargs)

    with monkeypatch.context() as patch:
        patch.setattr(fitz.Page, "get_pixmap", fail_second_page)
        with pytest.raises(RuntimeError, match="synthetic render failure"):
            PdfExporter.render_preview_images(data)
    assert opened_documents[0].is_closed
    assert native_cleanup == [100]
    assert len(PdfExporter.render_preview_images(data)) == 2
    assert native_cleanup == [100, 100]
