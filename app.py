from __future__ import annotations

import asyncio
import base64
import json
from dataclasses import asdict
from functools import lru_cache
from pathlib import Path
from threading import Lock
from urllib.parse import quote

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool
from starlette.datastructures import Headers, MutableHeaders
from starlette.exceptions import HTTPException
from starlette.types import ASGIApp, Message, Receive, Scope, Send

from coc7_card.catalog import TemplateCatalog
from coc7_card.models import Attributes
from coc7_card.portraits import PortraitError, validate_portrait
from coc7_card.template_config import TEMPLATE_PATH
from coc7_card.rules import FormulaError, RuleEngine
from coc7_card.web import (
    DraftPayloadError,
    build_draft,
    catalog_payload,
    report_payload,
)


ROOT = Path(__file__).resolve().parent
STATIC_PATH = ROOT / "static"
ASSETS_PATH = ROOT / "assets"
FONT_PATH = ASSETS_PATH / "fonts" / "NotoSansSC-Regular.ttf"
MAX_PORTRAIT_BYTES = 8 * 1024 * 1024
MAX_WORKBOOK_BYTES = 12 * 1024 * 1024
MAX_REQUEST_BYTES = 20 * 1024 * 1024
HEAVY_REQUEST_PATHS = {
    "/api/import/excel", "/api/import/attributes", "/api/export/pdf", "/api/export/excel",
}


class RequestLimitsMiddleware:
    def __init__(self, app: ASGIApp) -> None:
        self.app = app
        self.heavy_request = Lock()

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return

        async def send_with_headers(message: Message) -> None:
            if message["type"] == "http.response.start":
                headers = MutableHeaders(scope=message)
                headers.setdefault("Cache-Control", "no-store")
                headers.setdefault("X-Content-Type-Options", "nosniff")
                headers.setdefault("X-Frame-Options", "SAMEORIGIN")
                headers.setdefault("Referrer-Policy", "no-referrer")
                headers.setdefault(
                    "Content-Security-Policy",
                    "default-src 'self'; img-src 'self' data: blob:; style-src 'self' 'unsafe-inline'; "
                    "script-src 'self'; connect-src 'self'; object-src 'self' blob:; frame-src 'self' blob:",
                )
            await send(message)

        try:
            declared_size = int(Headers(scope=scope).get("content-length", "0"))
        except ValueError:
            declared_size = 0  # The actual byte limit still applies without a valid header.
        if declared_size > MAX_REQUEST_BYTES:
            await JSONResponse(status_code=413, content={"detail": "请求体过大。"})(scope, receive, send_with_headers)
            return

        heavy = scope["method"] == "POST" and scope["path"] in HEAVY_REQUEST_PATHS
        if heavy and not self.heavy_request.acquire(blocking=False):
            await JSONResponse(
                status_code=503,
                content={"detail": "正在处理另一项导入或导出，请稍后重试。"},
                headers={"Retry-After": "1"},
            )(scope, receive, send_with_headers)
            return

        received_size = 0

        async def bounded_receive() -> Message:
            nonlocal received_size
            message = await receive()
            if message["type"] == "http.request":
                received_size += len(message.get("body", b""))
                if received_size > MAX_REQUEST_BYTES:
                    # FastAPI preserves this status; the multipart parser closes partial files.
                    raise HTTPException(status_code=413, detail="请求体过大。")
            return message

        def release_after_worker(worker: asyncio.Task) -> None:
            if not worker.cancelled():
                worker.exception()  # Consume errors if the client cancelled its request.
            self.heavy_request.release()

        try:
            await self.app(scope, bounded_receive, send_with_headers)
        finally:
            if heavy:
                worker = scope.get("coc7_heavy_worker")
                if worker is not None and not worker.done():
                    worker.add_done_callback(release_after_worker)
                elif worker is not None:
                    release_after_worker(worker)
                else:
                    self.heavy_request.release()


async def _run_heavy(request: Request, function, *args):  # noqa: ANN001
    # Cancelling an HTTP task must not free its slot while its thread still owns memory.
    worker = asyncio.create_task(run_in_threadpool(function, *args))
    request.scope["coc7_heavy_worker"] = worker
    return await asyncio.shield(worker)


@lru_cache(maxsize=1)
def get_catalog() -> TemplateCatalog:
    return TemplateCatalog.load(TEMPLATE_PATH)


app = FastAPI(
    title="COC7 调查员车卡器",
    description="可自行部署的中文 COC7 调查员建卡工具。",
    docs_url=None,
    redoc_url=None,
)
app.mount("/static", StaticFiles(directory=STATIC_PATH), name="static")
app.mount("/assets", StaticFiles(directory=ASSETS_PATH), name="assets")
app.add_middleware(RequestLimitsMiddleware)


@app.exception_handler(DraftPayloadError)
async def invalid_draft_handler(_request: Request, exc: DraftPayloadError) -> JSONResponse:
    return JSONResponse(status_code=422, content={"detail": str(exc)})


@app.get("/", include_in_schema=False)
def index() -> FileResponse:
    return FileResponse(STATIC_PATH / "index.html")


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok", "service": "coc7-investigator-builder"}


@app.get("/api/bootstrap")
def bootstrap():
    try:
        return catalog_payload(get_catalog())
    except (FileNotFoundError, ValueError) as exc:
        return JSONResponse(status_code=503, content={"detail": str(exc)})


@app.post("/api/calculate")
def calculate(payload: dict) -> JSONResponse:
    try:
        attributes = Attributes.from_mapping(payload.get("attributes", {}))
        age = int(payload.get("age", 30))
        formula = str(payload.get("occupation_formula") or "EDU*4")
        derived = RuleEngine.calculate(attributes, age, formula, int(payload.get("san_loss", 0)))
        return JSONResponse({"derived": asdict(derived)})
    except (FormulaError, TypeError, ValueError) as exc:
        return JSONResponse(status_code=422, content={"detail": str(exc)})


@app.post("/api/validate")
def validate(payload: dict):
    try:
        draft = build_draft(payload, get_catalog())
        occupation_formula = draft.occupation.point_formula if draft.occupation else "EDU*4"
        derived = RuleEngine.calculate(draft.attributes, draft.identity.age, occupation_formula, draft.experience.san_loss)
    except (DraftPayloadError, FormulaError, TypeError, ValueError) as exc:
        return JSONResponse(status_code=422, content={"detail": str(exc)})
    except FileNotFoundError as exc:
        return JSONResponse(status_code=503, content={"detail": str(exc)})
    return {
        "validation": report_payload(RuleEngine.validate(draft)),
        "derived": asdict(derived),
        "asset_reference": RuleEngine.asset_reference(draft.assets.credit_rating),
    }


@app.post("/api/import/excel")
async def import_excel(request: Request, workbook: UploadFile = File(...)) -> JSONResponse:
    from coc7_card.importers.excel import ExcelImportError, import_investigator

    try:
        if Path(workbook.filename or "").suffix.lower() != ".xlsx":
            return JSONResponse(status_code=422, content={"detail": "请选择 COC7 调查员卡的 .xlsx 文件。"})
        data = await workbook.read(MAX_WORKBOOK_BYTES + 1)
        if len(data) > MAX_WORKBOOK_BYTES:
            return JSONResponse(status_code=413, content={"detail": "Excel 文件不能超过 12 MB。"})
        result = await _run_heavy(request, import_investigator, data, get_catalog())
        return JSONResponse(result)
    except ExcelImportError as exc:
        return JSONResponse(status_code=422, content={"detail": str(exc)})
    except FileNotFoundError as exc:
        return JSONResponse(status_code=503, content={"detail": str(exc)})
    finally:
        await workbook.close()


@app.post("/api/import/attributes")
async def import_attribute_table(request: Request, workbook: UploadFile = File(...)) -> JSONResponse:
    from coc7_card.importers import ExcelImportError, import_attributes

    try:
        if Path(workbook.filename or "").suffix.lower() not in {".xlsx", ".csv", ".tsv"}:
            return JSONResponse(status_code=422, content={"detail": "请选择 .xlsx、.csv 或 .tsv 简化属性表。"})
        data = await workbook.read(MAX_WORKBOOK_BYTES + 1)
        if len(data) > MAX_WORKBOOK_BYTES:
            return JSONResponse(status_code=413, content={"detail": "表格文件不能超过 12 MB。"})
        result = await _run_heavy(request, import_attributes, data, workbook.filename or "")
        return JSONResponse(result)
    except ExcelImportError as exc:
        return JSONResponse(status_code=422, content={"detail": str(exc)})
    finally:
        await workbook.close()


@app.post("/api/export/pdf")
async def export_pdf(
    request: Request,
    draft_json: str = Form(...),
    portrait: UploadFile | None = File(default=None),
) -> JSONResponse:
    from coc7_card.exporters.pdf import PdfExportError, PdfExporter

    try:
        payload = _decode_draft_json(draft_json)
        draft = build_draft(payload, get_catalog())
        portrait_bytes = await _read_portrait(portrait)
    except (DraftPayloadError, FormulaError, TypeError, ValueError) as exc:
        return JSONResponse(status_code=422, content={"detail": str(exc)})
    except FileNotFoundError as exc:
        return JSONResponse(status_code=503, content={"detail": str(exc)})
    try:
        result = await _run_heavy(request, PdfExporter(FONT_PATH).export, draft, portrait_bytes)
    except PdfExportError as exc:
        return JSONResponse(status_code=422, content={"detail": str(exc)})
    return JSONResponse(
        {
            "filename": result.filename,
            "page_count": result.page_count,
            "pdf_base64": base64.b64encode(result.data).decode("ascii"),
            "previews": [base64.b64encode(image).decode("ascii") for image in result.preview_images],
        }
    )


@app.post("/api/export/excel")
async def export_excel(
    request: Request,
    draft_json: str = Form(...),
    portrait: UploadFile | None = File(default=None),
) -> Response:
    from coc7_card.exporters.excel import ExcelExportError, ExcelUnavailableError
    from coc7_card.exporters.factory import excel_exporter

    try:
        payload = _decode_draft_json(draft_json)
        draft = build_draft(payload, get_catalog())
        portrait_bytes = await _read_portrait(portrait)
    except (DraftPayloadError, FormulaError, TypeError, ValueError) as exc:
        return JSONResponse(status_code=422, content={"detail": str(exc)})
    except FileNotFoundError as exc:
        return JSONResponse(status_code=503, content={"detail": str(exc)})
    try:
        result = await _run_heavy(request, excel_exporter(get_catalog()).export, draft, portrait_bytes)
    except (ExcelUnavailableError, ExcelExportError) as exc:
        return JSONResponse(status_code=422, content={"detail": str(exc)})
    encoded_name = quote(result.filename)
    return Response(
        result.data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}",
            "X-Workbook-Sheet-Count": str(result.sheet_count),
        },
    )


def _decode_draft_json(value: str) -> dict:
    try:
        payload = json.loads(value)
    except json.JSONDecodeError as exc:
        raise DraftPayloadError("调查员数据格式无效。") from exc
    if not isinstance(payload, dict):
        raise DraftPayloadError("调查员数据格式无效。")
    return payload


async def _read_portrait(upload: UploadFile | None) -> bytes | None:
    if upload is None or not upload.filename:
        return None
    allowed = {"image/png", "image/jpeg", "image/webp"}
    if upload.content_type not in allowed:
        raise DraftPayloadError("头像只支持 PNG、JPEG 或 WebP。")
    data = await upload.read(MAX_PORTRAIT_BYTES + 1)
    if len(data) > MAX_PORTRAIT_BYTES:
        raise DraftPayloadError("头像文件不能超过 8 MB。")
    try:
        validate_portrait(data)
    except PortraitError as exc:
        raise DraftPayloadError(str(exc)) from exc
    return data
