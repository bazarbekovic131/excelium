"""POST /render/registry/{вид} — Excel-реестры платежей.

Контракт совместим со старым сервисом: тело {"request": [...]},
ответ {"download_url": "..."}. Файл выдаётся через /files/{token}.
"""
import io
import json
import logging
import re
from datetime import date
from urllib.parse import quote

from fastapi import APIRouter, Body, HTTPException, Request, Response
from pydantic import BaseModel, Field

from ..config import APP_DIR
from ..renderers.registry_inner import render_inner
from ..renderers.registry_outer import render_outer
from ..renderers.registry_priority import render_priority
from ..renderers.typst_renderer import TypstError, render_typst, typst_available

TYPST_DIR = APP_DIR / "templates" / "typst"
TEMPLATE_NAME_RE = re.compile(r"^[a-z0-9_-]{1,64}$")

log = logging.getLogger(__name__)
router = APIRouter()

REGISTRY_PREFIX = "РЕЕСТР ПЛАТЕЖЕЙ №"


class RegistryPayload(BaseModel):
    request: list[dict] = Field(min_length=1)


def _sorted(entries: list[dict]) -> list[dict]:
    return sorted(entries, key=lambda x: x.get("object_name") or "")


def _regnum(entries: list[dict]) -> str:
    name = str(entries[0].get("registry_name") or "")
    return name.removeprefix(REGISTRY_PREFIX).strip()


def _sanitize(value: str) -> str:
    return re.sub(r"\s+", "_", re.sub(r"[^\w\s-]", "", value, flags=re.UNICODE)).strip("_")


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def file_response(data: bytes, filename: str, media_type: str) -> Response:
    """Ответ самим файлом — для одношагового скачивания из Doc-V.

    Имя отдаётся дважды: ASCII-версией и filename* по RFC 5987, иначе
    кириллица в заголовке ломает часть клиентов.
    """
    ascii_name = re.sub(r"[^A-Za-z0-9._-]", "_", filename) or "file"
    disposition = (f"attachment; filename=\"{ascii_name}\"; "
                   f"filename*=UTF-8''{quote(filename)}")
    return Response(content=data, media_type=media_type,
                    headers={"Content-Disposition": disposition})


def _wants_file(request: Request) -> bool:
    """?direct=1 — вернуть файл вместо ссылки на него."""
    return str(request.query_params.get("direct", "")).lower() in ("1", "true", "on", "yes")


def _payload_from_query(request: Request) -> dict:
    """JSON из адреса (?data=…) — так GET-запрос Doc-V умеет и передать
    данные, и сразу забрать файл: галка скачивания есть только у GET."""
    raw = request.query_params.get("data", "").strip()
    if not raw:
        return {}
    try:
        parsed = json.loads(raw)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=f"data — не JSON: {exc}")
    if not isinstance(parsed, dict):
        raise HTTPException(status_code=422, detail="data — должен быть объектом")
    return parsed


def _deliver(request: Request, workbook, orig_name: str):
    request.app.state.heartbeat.touch("render")
    buf = io.BytesIO()
    workbook.save(buf)
    workbook.close()
    if _wants_file(request):
        return file_response(buf.getvalue(), orig_name, XLSX_MIME)
    token = request.app.state.filestore.save_bytes(buf.getvalue(), ".xlsx", orig_name)
    return {"download_url": request.app.state.filestore.download_url(token)}


@router.post("/render/registry/inner")
def registry_inner(payload: RegistryPayload, request: Request):
    entries = _sorted(payload.request)
    workbook = render_inner(entries, request.app.state.template_inner,
                            request.app.state.signers)
    name = f"reestr_{_sanitize(_regnum(entries))}_ot_{date.today()}.xlsx"
    return _deliver(request, workbook, name)


@router.post("/render/registry/outer")
def registry_outer(payload: RegistryPayload, request: Request):
    workbook = render_outer(payload.request, request.app.state.template_outer,
                            request.app.state.banks)
    return _deliver(request, workbook, f"vneshny_reestr_ot_{date.today()}.xlsx")


@router.post("/render/registry/priority")
def registry_priority(payload: RegistryPayload, request: Request):
    entries = _sorted(payload.request)
    workbook = render_priority(entries, request.app.state.template_priority,
                               request.app.state.signers)
    name = f"reestr_prioritetov_{_sanitize(_regnum(entries))}_ot_{date.today()}.xlsx"
    return _deliver(request, workbook, name)


@router.get("/render/registry/{kind}")
def registry_get(kind: str, request: Request):
    """GET /render/registry/inner?data={…} — сразу файл, одним действием."""
    if kind not in ("inner", "outer", "priority"):
        raise HTTPException(status_code=404)
    payload = _payload_from_query(request)
    entries = payload.get("request") or []
    if not entries:
        raise HTTPException(status_code=422, detail="нужен data={\"request\": [...]}")
    state = request.app.state
    request.app.state.heartbeat.touch("render")
    if kind == "outer":
        workbook = render_outer(entries, state.template_outer, state.banks)
        name = f"vneshny_reestr_ot_{date.today()}.xlsx"
    elif kind == "priority":
        workbook = render_priority(_sorted(entries), state.template_priority, state.signers)
        name = f"reestr_prioritetov_{_sanitize(_regnum(entries))}_ot_{date.today()}.xlsx"
    else:
        workbook = render_inner(_sorted(entries), state.template_inner, state.signers)
        name = f"reestr_{_sanitize(_regnum(entries))}_ot_{date.today()}.xlsx"
    buf = io.BytesIO()
    workbook.save(buf)
    workbook.close()
    return file_response(buf.getvalue(), name, XLSX_MIME)


@router.get("/render/typst/{name}")
def render_typst_get(name: str, request: Request):
    """GET /render/typst/contract_card?data={…} — сразу PDF."""
    return _render_typst(name, request, _payload_from_query(request), force_file=True)


@router.post("/render/typst/{name}")
def render_typst_endpoint(name: str, request: Request, data: dict = Body(...)):
    return _render_typst(name, request, data, force_file=_wants_file(request))


def _render_typst(name: str, request: Request, data: dict, *, force_file: bool):
    if not TEMPLATE_NAME_RE.fullmatch(name):
        raise HTTPException(status_code=404)
    store = request.app.state.typst_store
    source = store.get(name)
    if source is None:
        raise HTTPException(status_code=404, detail="нет такого шаблона")
    if not typst_available():
        raise HTTPException(status_code=503, detail="typst не установлен на сервере")
    request.app.state.heartbeat.touch("render")
    try:
        pdf = render_typst(name, source, data, store.assets_bytes(),
                           verify_secret=request.app.state.settings.verify_secret,
                           directory=request.app.state.directory.all())
    except TypstError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    filename = f"{name}_{date.today()}.pdf"
    if force_file:
        return file_response(pdf, filename, "application/pdf")
    token = request.app.state.filestore.save_bytes(pdf, ".pdf", filename)
    return {"download_url": request.app.state.filestore.download_url(token)}
