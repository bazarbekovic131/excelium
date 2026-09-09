"""Веб-интерфейс администратора: /ui.

Server-rendered Jinja2 без внешних ресурсов. Вход — админ-токен
(GW_TOKEN_ADMIN), после входа кладётся в HttpOnly-cookie SameSite=Strict;
её проверяет security-middleware. Все страницы работают поверх тех же
внутренних объектов, что и API, — отдельной логики здесь нет.
"""
import json
import logging
import os
import shutil
from datetime import date, datetime, timedelta, timezone
from fastapi import APIRouter, Form, Request, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

from ..config import APP_DIR
from ..logging_setup import audit_log
from ..opsrunner.runner import OpsValidationError, run_operation
from ..renderers.registry_inner import render_inner
from ..renderers.registry_outer import render_outer
from ..renderers.registry_priority import render_priority
from ..renderers.typst_renderer import (TypstError, render_typst, typst_available,
                                        typst_binary)
from ..security import ADMIN_COOKIE, _match
from ..signers import parse_slot, slot_value
from .render import _deliver, _sorted

log = logging.getLogger(__name__)
router = APIRouter()
templates = Jinja2Templates(directory=APP_DIR / "gateway" / "webui")
# Стили встраиваются в страницу: отдельный статический файл не годится —
# страница входа отдаётся до проверки cookie, и запрос за CSS получил бы отказ.
DS_CSS = (APP_DIR / "gateway" / "webui" / "ds.css").read_text(encoding="utf-8")
templates.env.globals["ds_css"] = DS_CSS


def _dt(value, with_time: bool = True) -> str:
    """ISO-строка (обычно UTC) -> «28.08.2026 13:08» по Астане.

    В интерфейсе не должно оставаться машинных дат: администратор
    сверяет их с часами на стене, а не с UTC.
    """
    text = str(value or "").strip()
    if not text:
        return "—"
    try:
        moment = datetime.fromisoformat(text)
    except ValueError:
        return text
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=timezone.utc)
    moment = moment.astimezone(ALMATY)
    return moment.strftime("%d.%m.%Y %H:%M" if with_time else "%d.%m.%Y")


def _ago(value) -> str:
    """«3 минуты назад» — для колонок, где важна свежесть, а не дата."""
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        moment = datetime.fromisoformat(text)
    except ValueError:
        return ""
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=timezone.utc)
    seconds = int((datetime.now(timezone.utc) - moment).total_seconds())
    if seconds < 60:
        return "только что"
    if seconds < 3600:
        return f"{seconds // 60} мин назад"
    if seconds < 86400:
        return f"{seconds // 3600} ч назад"
    return f"{seconds // 86400} дн назад"


templates.env.filters["dt"] = _dt
templates.env.filters["ago"] = _ago

AUDIT_TAIL = 50
ALMATY = timezone(timedelta(hours=5))
EVENT_TITLES = {
    "deny_ip": "Отклонён запрос с чужого адреса",
    "deny_token": "Отклонён запрос с неверным токеном",
    "deny_ui": "Отклонён вход в панель",
    "rate_limited": "Запрос отклонён по лимиту частоты",
    "ui_login": "Вход в панель",
    "ui_login_failed": "Неверный админ-токен",
    "job_enqueued": "Задание поставлено в очередь",
    "jobs_acked": "Doc-V подтвердил задания",
    "job_many_attempts": "Задание выдаётся слишком часто",
    "ops_start": "Запущена операция",
    "ops_finish": "Операция завершена",
    "directory_replaced": "Обновлён справочник из Doc-V",
    "typst_template_saved": "Сохранён шаблон Typst",
    "typst_template_restored": "Возвращена версия шаблона",
    "typst_template_deleted": "Удалён шаблон Typst",
    "typst_asset_uploaded": "Загружена картинка",
    "typst_asset_deleted": "Удалена картинка",
    "file_to_assets": "Файл перенесён в картинки",
    "ui_file_uploaded": "Загружен файл",
    "ui_file_renamed": "Файл переименован",
    "ui_file_deleted": "Файл удалён",
    "ui_files_zip": "Файлы скачаны архивом",
    "ui_job_ack": "Задание подтверждено вручную",
    "ui_job_enqueued": "Создано тестовое задание",
}
SAMPLE = json.dumps({"request": [{
    "registry_name": "РЕЕСТР ПЛАТЕЖЕЙ №1", "organization": "ТОО «Шар-Кұрылыс»",
    "object_name": "Администрация", "counteragent": "ТОО «Пример»",
    "zatraty": "Прочее", "payment_sum": "100000.00", "payment_type": "Оплата",
    "payment_objective": "по счету", "doctype": "Заявка на оплату",
    "payment_number": 1, "status": "На исполнении",
}]}, ensure_ascii=False, indent=1)


REFRESHABLE = {"dash", "jobs"}


def _page(request: Request, template: str, page: str, **ctx) -> HTMLResponse:
    return templates.TemplateResponse(
        request, template,
        {"page": page, "refreshable": page in REFRESHABLE, **ctx})


# --- вход/выход -----------------------------------------------------------

@router.get("/ui/login")
def login_form(request: Request):
    return templates.TemplateResponse(request, "login.html", {})


@router.post("/ui/login")
def login(request: Request, token: str = Form(default="")):
    settings = request.app.state.settings
    if not settings.token_admin:
        return templates.TemplateResponse(
            request, "login.html",
            {"error": "Интерфейс выключен: задайте GW_TOKEN_ADMIN в .env"})
    if not _match(token, settings.token_admin):
        audit_log("ui_login_failed", ip=request.client.host if request.client else "")
        return templates.TemplateResponse(
            request, "login.html", {"error": "Неверный токен"})
    response = RedirectResponse("/ui", status_code=302)
    response.set_cookie(ADMIN_COOKIE, token, httponly=True, samesite="strict",
                        max_age=12 * 3600)
    audit_log("ui_login", ip=request.client.host if request.client else "")
    return response


@router.get("/ui/logout")
def logout():
    response = RedirectResponse("/ui/login", status_code=302)
    response.delete_cookie(ADMIN_COOKIE)
    return response


# --- обзор ----------------------------------------------------------------

@router.get("/ui")
def dashboard(request: Request):
    state = request.app.state
    files = state.filestore.list_files()

    # лента: audit-лог, свежие сверху; вид точки — по характеру события
    events, denies = [], 0
    audit_path = state.settings.var_dir / "audit.log"
    if audit_path.is_file():
        for line in audit_path.read_text(encoding="utf-8").splitlines()[-AUDIT_TAIL:][::-1]:
            try:
                e = json.loads(line)
            except ValueError:
                continue
            message = e.get("message", "")
            details = {k: v for k, v in e.items()
                       if k not in ("ts", "level", "logger", "message")}
            dot = ""
            if message.startswith("deny") or message == "rate_limited":
                dot, denies = "danger", denies + 1
            elif message.endswith("_failed") or details.get("ok") is False:
                dot = "warn"
            elif message.startswith("ui_") or message.startswith("job"):
                dot = "off" if message.startswith("ui_") else ""
            events.append({
                "time": e.get("ts", "")[11:19],
                "dot": dot,
                "text": EVENT_TITLES.get(message, message) + (
                    " · " + json.dumps(details, ensure_ascii=False) if details else ""),
            })

    beats = []
    for kind, info in state.heartbeat.snapshot().items():
        age = info["age_sec"]
        if age is None:
            dot, text = "off", "ещё не было"
        elif age < 120:
            dot, text = "", f"{age} с назад"
        elif age < 3600:
            dot, text = "warn", f"{age // 60} мин назад"
        else:
            dot, text = "warn", f"{age // 3600} ч назад"
        beats.append({"kind": kind, "label": info["label"], "dot": dot, "text": text,
                      "clock": (info["seen_at"] or "")[11:19] or "—"})

    now = datetime.now(ALMATY)
    return _page(request, "dashboard.html", "dash", beats=beats,
                 directories=state.directory.info(),
                 typst=typst_available(), typst_path=typst_binary(),
                 libreoffice=shutil.which("libreoffice") is not None,
                 libreoffice_path=shutil.which("libreoffice"),
                 path_env=os.environ.get("PATH", ""),
                 files_count=len(files),
                 files_mb=round(sum(f["size"] for f in files) / 1024 / 1024, 1),
                 templates_count=len(state.typst_store.list_templates()),
                 ops_count=len(state.ops), denies=denies,
                 jobs=state.jobs.stats(), events=events,
                 health_title=("Шлюз в работе, все каналы на связи"
                               if all(b["dot"] == "" for b in beats) or not beats
                               else "Шлюз в работе, есть молчащие каналы"),
                 now_line="обновлено только что",
                 today_line=now.strftime("%d.%m.%Y, %H:%M по Астане"))


# --- задания --------------------------------------------------------------

JOB_STATUS_TITLES = {"pending": "в очереди", "leased": "выдано Doc-V",
                     "acked": "подтверждено"}


PAGE_SIZE = 50


@router.get("/ui/jobs")
def jobs_page(request: Request, status: str = "", search: str = "",
              producer: str = "", limit: int = PAGE_SIZE, flash: str = ""):
    queue = request.app.state.jobs
    rows = queue.list_jobs(status=status or None, search=search.strip(),
                           producer=producer, limit=limit + 1)
    has_more = len(rows) > limit
    return _page(request, "jobs.html", "jobs", rows=rows[:limit], status_filter=status,
                 search=search, producer=producer, producers=queue.producers(),
                 status_titles=JOB_STATUS_TITLES, has_more=has_more,
                 next_limit=limit + PAGE_SIZE, flash=flash)


@router.post("/ui/jobs/ack/{job_id}")
def jobs_ack(request: Request, job_id: int):
    result = request.app.state.jobs.ack([job_id])
    audit_log("ui_job_ack", job_id=job_id, result=result)
    return RedirectResponse("/ui/jobs?flash=Задание подтверждено", status_code=302)


@router.post("/ui/jobs/new")
def jobs_new(request: Request, type: str = Form(...), payload: str = Form(...)):
    try:
        data = json.loads(payload)
        if not isinstance(data, dict):
            raise ValueError
    except ValueError:
        queue = request.app.state.jobs
        return _page(request, "jobs.html", "jobs", rows=queue.list_jobs(limit=PAGE_SIZE),
                     status_filter="", search="", producer="",
                     producers=queue.producers(), status_titles=JOB_STATUS_TITLES,
                     has_more=False, next_limit=PAGE_SIZE,
                     flash="Payload — не JSON-объект", flash_err=True)
    job_id, _ = request.app.state.jobs.enqueue(
        producer="ui", job_type=type.strip() or "тест", payload=data,
        idempotency_key=None)
    audit_log("ui_job_enqueued", job_id=job_id)
    return RedirectResponse(f"/ui/jobs?flash=Задание {job_id} в очереди", status_code=302)


# --- файлы ----------------------------------------------------------------

@router.get("/ui/files")
def files_page(request: Request, search: str = "", limit: int = PAGE_SIZE,
               flash: str = ""):
    rows = request.app.state.filestore.list_files()
    needle = search.strip().lower()
    if needle:
        rows = [f for f in rows if needle in f["orig_name"].lower()]
    total = len(rows)
    return _page(request, "files.html", "files", rows=[
                     dict(f, is_image=f["suffix"].lower() in IMAGE_SUFFIXES)
                     for f in rows[:limit]],
                 total=total, has_more=total > limit, next_limit=limit + PAGE_SIZE,
                 search=search, base_url=request.app.state.settings.base_url,
                 ttl_hours=request.app.state.settings.file_ttl_hours, flash=flash)


@router.post("/ui/files/upload")
async def files_upload(request: Request, uploads: list[UploadFile]):
    count = 0
    for upload in uploads:
        data = await upload.read()
        if not data and not upload.filename:
            continue
        name = upload.filename or "file"
        suffix = "." + name.rsplit(".", 1)[1].lower() if "." in name else ".bin"
        try:
            request.app.state.filestore.save_bytes(data, suffix, name)
        except ValueError:
            request.app.state.filestore.save_bytes(data, ".bin", name)
        audit_log("ui_file_uploaded", name=name, size=len(data))
        count += 1
    word = "Файл загружен" if count == 1 else f"Загружено файлов: {count}"
    return RedirectResponse(f"/ui/files?flash={word}", status_code=302)


@router.post("/ui/files/rename/{token}")
def files_rename(request: Request, token: str, new_name: str = Form(...)):
    store = request.app.state.filestore
    resolved = store.resolve(token)
    if resolved is None:
        return RedirectResponse("/ui/files?flash=Файл не найден", status_code=302)
    path, _ = resolved
    name = new_name.strip()
    if name and not name.lower().endswith(path.suffix.lower()):
        name += path.suffix  # расширение не теряем
    if not store.rename(token, name):
        return RedirectResponse("/ui/files?flash=Не переименован", status_code=302)
    audit_log("ui_file_renamed", token=token, name=name)
    return RedirectResponse("/ui/files?flash=Переименовано", status_code=302)


@router.post("/ui/files/download_zip")
def files_download_zip(request: Request, tokens: list[str] = Form(default=[])):
    import io as _io
    import zipfile
    from datetime import date as _date
    from fastapi import Response
    store = request.app.state.filestore
    buf = _io.BytesIO()
    used, count = set(), 0
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for token in tokens[:200]:
            resolved = store.resolve(token)
            if resolved is None:
                continue
            path, orig_name = resolved
            arcname, n = orig_name, 1
            while arcname in used:
                stem, _, ext = orig_name.rpartition(".")
                arcname = f"{stem}_{n}.{ext}" if ext else f"{orig_name}_{n}"
                n += 1
            used.add(arcname)
            zf.write(path, arcname=arcname)
            count += 1
    if not count:
        return RedirectResponse("/ui/files?flash=Ничего не выбрано", status_code=302)
    audit_log("ui_files_zip", count=count)
    fname = f"files_{_date.today()}.zip"
    return Response(content=buf.getvalue(), media_type="application/zip",
                    headers={"Content-Disposition": f'attachment; filename="{fname}"'})


@router.post("/ui/files/delete/{token}")
def files_delete(request: Request, token: str):
    ok = request.app.state.filestore.delete(token)
    audit_log("ui_file_deleted", token=token, ok=ok)
    return RedirectResponse("/ui/files?flash=Файл удалён", status_code=302)


# --- операции -------------------------------------------------------------

@router.get("/ui/ops")
def ops_page(request: Request, flash: str = ""):
    return _page(request, "ops.html", "ops", ops=request.app.state.ops,
                 files=request.app.state.filestore.list_files(), result=None,
                 flash=flash)


@router.post("/ui/ops/{name}")
async def ops_run(request: Request, name: str):
    state = request.app.state
    op = state.ops.get(name)
    if op is None:
        return RedirectResponse("/ui/ops", status_code=302)
    form = await request.form()
    params = {}
    for pname, pdef in op.params.items():
        if pdef.type == "file_list":
            values = [v for v in form.getlist(pname) if v]
            if values:
                params[pname] = values
        else:
            value = str(form.get(pname) or "").strip()
            if value:
                params[pname] = value
    try:
        ip = request.client.host if request.client else ""
        result = await run_in_threadpool(run_operation, op, params,
                                         state.filestore, client_ip=ip)
        flash, flash_err = "", False
    except OpsValidationError as exc:
        result, flash, flash_err = None, f"Параметры не приняты: {exc}", True
    return _page(request, "ops.html", "ops", ops=state.ops,
                 files=state.filestore.list_files(), result=result,
                 flash=flash, flash_err=flash_err)


# --- рендер ---------------------------------------------------------------

TYPST_SAMPLE = '{"title": "Справка", "fields": {"Поле": "Значение"}}'


@router.get("/ui/render")
def render_page(request: Request, link: str = "", link_name: str = ""):
    return _page(request, "render.html", "render", sample=SAMPLE,
                 typst=typst_available(), typst_templates=_typst_templates(request),
                 typst_data=TYPST_SAMPLE, link=link, link_name=link_name)


@router.post("/ui/render/registry")
def render_registry(request: Request, kind: str = Form(...), data: str = Form(...)):
    state = request.app.state
    try:
        entries = json.loads(data).get("request") or []
        assert entries
    except (ValueError, AttributeError, AssertionError):
        return _page(request, "render.html", "render", sample=data,
                     typst=typst_available(), typst_templates=_typst_templates(request),
                     typst_data=TYPST_SAMPLE,
                     flash="Нужен JSON вида {\"request\": [...]}", flash_err=True)
    try:
        if kind == "outer":
            workbook = render_outer(entries, state.template_outer, state.banks)
            name = f"vneshny_reestr_ot_{date.today()}.xlsx"
        elif kind == "priority":
            workbook = render_priority(_sorted(entries), state.template_priority,
                                       state.signers)
            name = f"reestr_prioritetov_ot_{date.today()}.xlsx"
        else:
            workbook = render_inner(_sorted(entries), state.template_inner,
                                    state.signers)
            name = f"reestr_ot_{date.today()}.xlsx"
    except Exception as exc:
        log.exception("ui render failed")
        return _page(request, "render.html", "render", sample=data,
                     typst=typst_available(), typst_templates=_typst_templates(request),
                     typst_data=TYPST_SAMPLE,
                     flash=f"Ошибка рендера: {exc}", flash_err=True)
    delivered = _deliver(request, workbook, name)
    return RedirectResponse(
        f"/ui/render?link={delivered['download_url']}&link_name={name}",
        status_code=302)


@router.post("/ui/render/typst")
def render_typst_ui(request: Request, template: str = Form(...), data: str = Form(...)):
    store = request.app.state.typst_store
    ctx = dict(sample=SAMPLE, typst=typst_available(),
               typst_templates=_typst_templates(request), typst_data=data)
    try:
        parsed = json.loads(data)
    except ValueError:
        return _page(request, "render.html", "render",
                     flash="Данные — не JSON", flash_err=True, **ctx)
    source = store.get(template)
    if source is None:
        return _page(request, "render.html", "render",
                     flash="Нет такого шаблона", flash_err=True, **ctx)
    try:
        pdf = render_typst(template, source, parsed, store.assets_bytes(),
                           verify_secret=request.app.state.settings.verify_secret,
                           directory=request.app.state.directory.all())
    except TypstError as exc:
        return _page(request, "render.html", "render",
                     flash=f"Ошибка компиляции: {exc}", flash_err=True, **ctx)
    name = f"{template}_{date.today()}.pdf"
    token = request.app.state.filestore.save_bytes(pdf, ".pdf", name)
    url = request.app.state.filestore.download_url(token)
    return RedirectResponse(f"/ui/render?link={url}&link_name={name}", status_code=302)


def _typst_templates(request: Request) -> list[str]:
    return [t["name"] for t in request.app.state.typst_store.list_templates()]


# --- шаблоны Typst --------------------------------------------------------

NEW_TEMPLATE = '''// Новый шаблон. Данные приходят из POST-запроса:
#let data = if "data" in sys.inputs { json(sys.inputs.data) } else { (:) }
#let meta = if "meta" in sys.inputs { json(sys.inputs.meta) } else { (:) }

#set page(paper: "a4", margin: 2cm)
#set text(font: ("Liberation Sans", "Arial", "DejaVu Sans"), size: 11pt, lang: "ru")

= #data.at("title", default: "Документ")
'''

TEST_DATA_DEFAULT = '{"title": "Проба"}'


def _typst_edit_ctx(request: Request, name: str, **extra):
    from ..renderers.typst_store import HISTORY_KEEP
    store = request.app.state.typst_store
    ctx = dict(name=name, source=store.get(name) or "",
               history=store.history(name), history_keep=HISTORY_KEEP,
               typst=typst_available(), assets=store.list_assets(),
               test_data=extra.pop("test_data", TEST_DATA_DEFAULT))
    ctx.update(extra)
    return ctx


@router.get("/ui/typst")
def typst_list(request: Request, flash: str = ""):
    store = request.app.state.typst_store
    return _page(request, "typst_list.html", "typst",
                 rows=store.list_templates(), assets=store.list_assets(), flash=flash)


@router.post("/ui/typst/create")
def typst_create(request: Request, name: str = Form(...)):
    store = request.app.state.typst_store
    name = name.strip()
    try:
        if store.get(name) is None:
            store.save(name, NEW_TEMPLATE)
    except ValueError as exc:
        return _page(request, "typst_list.html", "typst",
                     rows=store.list_templates(), assets=store.list_assets(),
                     flash=str(exc), flash_err=True)
    return RedirectResponse(f"/ui/typst/{name}", status_code=302)


@router.post("/ui/typst/assets/upload")
async def typst_asset_upload(request: Request, uploads: list[UploadFile]):
    store = request.app.state.typst_store
    count = 0
    for upload in uploads:
        if not upload.filename:
            continue
        try:
            store.save_asset(upload.filename.strip(), await upload.read())
        except ValueError as exc:
            return RedirectResponse(f"/ui/typst?flash={upload.filename}: {exc}",
                                    status_code=302)
        audit_log("typst_asset_uploaded", name=upload.filename)
        count += 1
    word = "Картинка загружена" if count == 1 else f"Загружено картинок: {count}"
    return RedirectResponse(f"/ui/typst?flash={word}", status_code=302)


@router.post("/ui/typst/assets/delete/{name}")
def typst_asset_delete(request: Request, name: str):
    request.app.state.typst_store.delete_asset(name)
    audit_log("typst_asset_deleted", name=name)
    return RedirectResponse("/ui/typst?flash=Картинка удалена", status_code=302)


@router.get("/ui/typst/{name}")
def typst_edit(request: Request, name: str, link: str = ""):
    if request.app.state.typst_store.get(name) is None:
        return RedirectResponse("/ui/typst", status_code=302)
    return _page(request, "typst_edit.html", "typst",
                 **_typst_edit_ctx(request, name, link=link))


@router.post("/ui/typst/{name}/save")
def typst_save(request: Request, name: str, source: str = Form(...)):
    store = request.app.state.typst_store # from renderers/typst_store.py;
    try:
        store.save(name, source)
    except ValueError as exc:
        return _page(request, "typst_edit.html", "typst",
                     **_typst_edit_ctx(request, name), flash=str(exc), flash_err=True)
    audit_log("typst_template_saved", name=name, size=len(source))
    return _page(request, "typst_edit.html", "typst",
                 **_typst_edit_ctx(request, name), flash="Сохранено")


@router.post("/ui/typst/{name}/test")
def typst_test(request: Request, name: str, data: str = Form(...)):
    store = request.app.state.typst_store
    source = store.get(name)
    if source is None:
        return RedirectResponse("/ui/typst", status_code=302)
    try:
        parsed = json.loads(data)
    except ValueError:
        return _page(request, "typst_edit.html", "typst",
                     **_typst_edit_ctx(request, name, test_data=data),
                     flash="Данные — не JSON", flash_err=True)
    try:
        pdf = render_typst(name, source, parsed, store.assets_bytes(),
                           verify_secret=request.app.state.settings.verify_secret,
                           directory=request.app.state.directory.all())
    except TypstError as exc:
        return _page(request, "typst_edit.html", "typst",
                     **_typst_edit_ctx(request, name, test_data=data, error=str(exc)))
    fname = f"{name}_{date.today()}.pdf"
    token = request.app.state.filestore.save_bytes(pdf, ".pdf", fname)
    url = request.app.state.filestore.download_url(token)
    return RedirectResponse(f"/ui/typst/{name}?link={url}", status_code=302)


@router.post("/ui/typst/{name}/restore/{history_id}")
def typst_restore(request: Request, name: str, history_id: int):
    request.app.state.typst_store.restore(name, history_id)
    audit_log("typst_template_restored", name=name, history_id=history_id)
    return RedirectResponse(f"/ui/typst/{name}", status_code=302)


@router.post("/ui/typst/{name}/delete")
def typst_delete(request: Request, name: str):
    request.app.state.typst_store.delete(name)
    audit_log("typst_template_deleted", name=name)
    return RedirectResponse("/ui/typst?flash=Шаблон удалён", status_code=302)


# --- картинки: отдача и перенос из «Файлов» -------------------------------

ASSET_MEDIA = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
               ".gif": "image/gif", ".svg": "image/svg+xml"}
IMAGE_SUFFIXES = set(ASSET_MEDIA)


@router.get("/ui/typst/assets/raw/{name}")
def typst_asset_raw(request: Request, name: str):
    from fastapi import Response
    data = request.app.state.typst_store.assets_bytes().get(name)
    if data is None:
        return Response(status_code=404)
    suffix = "." + name.rsplit(".", 1)[-1].lower()
    return Response(content=data, media_type=ASSET_MEDIA.get(suffix, "application/octet-stream"),
                    headers={"Cache-Control": "private, max-age=300"})


def _asset_name(orig_name: str) -> str:
    """Имя файла -> допустимое имя картинки: транслитерация кириллицы,
    пробелы и прочее -> подчёркивание."""
    import re as _re
    stem, _, ext = orig_name.rpartition(".")
    try:
        from transliterate import translit
        stem = translit(stem, "ru", reversed=True)
    except Exception:
        pass
    stem = _re.sub(r"[^A-Za-z0-9_-]+", "_", stem).strip("_") or "img"
    return f"{stem[:60]}.{ext.lower()}"


@router.post("/ui/files/to_assets/{token}")
def file_to_assets(request: Request, token: str):
    """Файл-изображение из «Файлов» -> «Картинки» (доступно шаблонам)."""
    resolved = request.app.state.filestore.resolve(token)
    if resolved is None:
        return RedirectResponse("/ui/files?flash=Файл не найден", status_code=302)
    path, orig_name = resolved
    name = _asset_name(orig_name)
    try:
        request.app.state.typst_store.save_asset(name, path.read_bytes())
    except ValueError as exc:
        return RedirectResponse(f"/ui/files?flash=Не перенесён: {exc}", status_code=302)
    audit_log("file_to_assets", token=token, name=name)
    return RedirectResponse(f"/ui/typst?flash=Картинка доступна шаблонам: assets/{name}",
                            status_code=302)


# --- конструктор операций -------------------------------------------------

def _reload_ops(request: Request, registry: dict) -> None:
    """Пишет ops.yaml и подменяет реестр в живом процессе — без рестарта."""
    from ..opsrunner.registry import dump_registry
    request.app.state.ops_path.write_text(dump_registry(registry), encoding="utf-8")
    request.app.state.ops = registry


@router.get("/ui/opsedit/new")
def ops_new(request: Request):
    return _page(request, "ops_edit.html", "ops", op=None, name="",
                 param_types=PARAM_TYPES)


@router.get("/ui/opsedit/{name}")
def ops_edit(request: Request, name: str):
    op = request.app.state.ops.get(name)
    if op is None:
        return RedirectResponse("/ui/ops", status_code=302)
    params = [{"name": p.name, "type": p.type, "required": p.required,
               "pattern": p.pattern.pattern if p.pattern else "",
               "max_items": p.max_items} for p in op.params.values()]
    return _page(request, "ops_edit.html", "ops", op=op, name=name,
                 params=params, param_types=PARAM_TYPES)


PARAM_TYPES = (("str", "строка по образцу"), ("file", "файл из хранилища"),
               ("file_list", "несколько файлов"))


@router.post("/ui/opsedit/save")
async def ops_save(request: Request):
    from ..opsrunner.registry import OpsConfigError, parse_operation
    form = await request.form()
    name = str(form.get("old_name") or "").strip()   # пусто = создаём новую
    new_name = str(form.get("new_name") or name).strip()

    spec: dict = {"argv": [a for a in form.getlist("argv") if a.strip()]}
    if str(form.get("description") or "").strip():
        spec["description"] = str(form["description"]).strip()
    if str(form.get("collect") or "").strip():
        spec["collect"] = str(form["collect"]).strip()
    for key in ("timeout_sec", "max_output_kb"):
        if str(form.get(key) or "").strip():
            spec[key] = int(str(form[key]).strip())

    params: dict = {}
    for pname, ptype, pattern, req in zip(
            form.getlist("p_name"), form.getlist("p_type"),
            form.getlist("p_pattern"), form.getlist("p_required")):
        pname = pname.strip()
        if not pname:
            continue
        entry: dict = {"type": ptype}
        if ptype == "str":
            entry["pattern"] = pattern.strip() or "^.{0,200}$"
        if req != "on":
            entry["required"] = False
        params[pname] = entry
    if params:
        spec["params"] = params

    registry = dict(request.app.state.ops)
    try:
        operation = parse_operation(new_name, spec)
    except OpsConfigError as exc:
        return _page(request, "ops_edit.html", "ops",
                     op=request.app.state.ops.get(name), name=name,
                     params=[{"name": k, "type": v.get("type", "str"),
                              "pattern": v.get("pattern", ""),
                              "required": v.get("required", True), "max_items": 50}
                             for k, v in params.items()],
                     draft=spec, param_types=PARAM_TYPES,
                     flash=str(exc), flash_err=True)
    if name and name in registry and name != new_name:
        del registry[name]
    registry[new_name] = operation
    _reload_ops(request, registry)
    audit_log("ops_saved", op=new_name, argv=operation.argv)
    return RedirectResponse(f"/ui/ops?flash=Операция {new_name} сохранена и уже доступна",
                            status_code=302)


@router.post("/ui/opsedit/delete/{name}")
def ops_delete(request: Request, name: str):
    registry = dict(request.app.state.ops)
    if registry.pop(name, None) is not None:
        _reload_ops(request, registry)
        audit_log("ops_deleted", op=name)
    return RedirectResponse("/ui/ops?flash=Операция удалена", status_code=302)


# --- API-консоль ----------------------------------------------------------

API_ENDPOINTS = [
    ("Задания", [("GET", "/jobs/pending?consumer=docv&limit=20", ""),
                 ("POST", "/jobs", '{"type": "тест", "payload": {}}'),
                 ("POST", "/jobs/ack", '{"ids": [1]}')]),
    ("Рендер", [("POST", "/render/registry/inner", '{"request": []}'),
                ("POST", "/render/typst/contract_card", "{}")]),
    ("Операции и файлы", [("GET", "/ops", ""),
                          ("POST", "/ops/translit", '{"params": {"text": "Тест"}}'),
                          ("GET", "/health", "")]),
    ("Справочники", [("POST", "/directory/structura",
                      '[{"uid": "u1", "display_name": "Тест Т.Т."}]'),
                     ("GET", "/directory", ""),
                     ("GET", "/signers?company=ТОО «Шар-Кұрылыс»", "")]),
    ("Отладка", [("POST", "/debug/echo", '{"пример": true}')]),
]
API_PREFIXES = ("/jobs", "/render", "/ops", "/directory", "/signers", "/health",
                "/debug", "/files")


@router.get("/ui/api")
def api_console(request: Request, method: str = "GET",
                path: str = "/jobs/pending?consumer=docv&limit=20", body: str = ""):
    return _page(request, "api.html", "api", groups=API_ENDPOINTS,
                 method=method, path=path, body=body, result=None,
                 log=request.app.state.apilog.items())


@router.post("/ui/api")
async def api_call(request: Request, method: str = Form("GET"),
                   path: str = Form(...), body: str = Form("")):
    """Вызывает собственную точку шлюза изнутри — с теми же токенами,
    что предъявляет Doc-V. Чужие адреса не допускаются."""
    import time

    import httpx

    settings = request.app.state.settings
    path = path.strip()
    if not path.startswith(API_PREFIXES):
        return _page(request, "api.html", "api", groups=API_ENDPOINTS, method=method,
                     path=path, body=body, result=None,
                     log=request.app.state.apilog.items(),
                     flash="Консоль вызывает только точки шлюза: "
                           + ", ".join(API_PREFIXES), flash_err=True)
    token = settings.token_ops if path.startswith("/ops") else settings.token_docv
    if path == "/jobs" and method == "POST":
        token = next(iter(settings.producer_tokens.values()), "")
    headers = {"Authorization": f"Bearer {token}"} if token else {}
    if body.strip():
        headers["Content-Type"] = "application/json"

    started = time.monotonic()
    try:
        transport = httpx.ASGITransport(app=request.app, client=("127.0.0.1", 0))
        async with httpx.AsyncClient(transport=transport, base_url="http://gateway",
                                     timeout=30) as client:
            response = await client.request(method, path, headers=headers,
                                            content=body.strip() or None)
        text, status = response.text, response.status_code
    except Exception as exc:  # ошибка вызова — тоже результат для журнала
        text, status = f"{type(exc).__name__}: {exc}", 0
    duration = int((time.monotonic() - started) * 1000)
    request.app.state.apilog.add(method, path, status, duration, len(text))

    pretty = text
    try:
        pretty = json.dumps(json.loads(text), ensure_ascii=False, indent=2)
    except ValueError:
        pass
    curl = (f"curl -X {method} 'http://{settings.base_url.split('//')[-1]}{path}'"
            + (f" \\\n  -H 'Authorization: Bearer <токен>'" if token else "")
            + (f" \\\n  -H 'Content-Type: application/json' \\\n  -d '{body.strip()}'"
               if body.strip() else ""))
    return _page(request, "api.html", "api", groups=API_ENDPOINTS, method=method,
                 path=path, body=body, log=request.app.state.apilog.items(),
                 result={"status": status, "duration_ms": duration, "size": len(text),
                         "text": pretty[:20000], "curl": curl,
                         "tone": ("success" if status and status < 300 else
                                  "warn" if status and status < 500 else "danger")})


# --- настройки ------------------------------------------------------------

def _mask(token: str) -> str:
    if not token:
        return "не задан"
    return token[:4] + "…" + token[-4:] if len(token) > 12 else "задан"


@router.get("/ui/settings")
def settings_page(request: Request, flash: str = ""):
    s = request.app.state.settings
    return _page(request, "settings.html", "settings", flash=flash,
                 values=request.app.state.settings_store.current(),
                 tokens=[("Doc-V", "GW_TOKEN_DOCV", _mask(s.token_docv)),
                         ("Операции", "GW_TOKEN_OPS", _mask(s.token_ops)),
                         ("Веб-интерфейс", "GW_TOKEN_ADMIN", _mask(s.token_admin)),
                         ("Код подлинности", "GW_VERIFY_SECRET", _mask(s.verify_secret))],
                 producers=[(name, _mask(tok)) for name, tok in s.producer_tokens.items()],
                 allowlist=s.allowlist, ui_allowlist=s.ui_allowlist,
                 base_url=s.base_url, var_dir=str(s.var_dir),
                 typst_bin=typst_binary() or s.typst_bin)


@router.post("/ui/settings")
async def settings_save(request: Request):
    form = await request.form()
    try:
        applied = request.app.state.settings_store.save(dict(form))
    except ValueError as exc:
        return RedirectResponse(f"/ui/settings?flash={exc}", status_code=302)
    audit_log("settings_saved", values=applied)
    return RedirectResponse("/ui/settings?flash=Настройки применены на лету",
                            status_code=302)


# --- подписанты -----------------------------------------------------------

def _binding_rows(store, search: str = "") -> list[dict]:
    """Привязки с уже подставленными ФИО — таблица должна читаться
    глазами, а не показывать идентификаторы."""
    people = {p["id"]: p for p in store.people()}

    def person(row, role):
        found = people.get(row[f"{role}_id"])
        if not found:
            return None
        return {"fio": found["fio"],
                "position": row[f"{role}_position"] or found["position"],
                "company": row[f"{role}_company"] or row["company"]}

    needle = search.strip().casefold()
    rows = []
    for row in store.bindings():
        if needle and needle not in f"{row['company']} {row['object_name']}".casefold():
            continue
        rows.append({**row, "left": person(row, "soglasovano"),
                     "right": person(row, "utverzhdayu")})
    return rows


@router.get("/ui/signers")
def signers_page(request: Request, search: str = "", flash: str = "", flash_err: str = ""):
    store = request.app.state.signers
    return _page(request, "signers.html", "signers", search=search,
                 bindings=_binding_rows(store, search), sets=store.sets(),
                 people=store.people(), stats=store.stats(),
                 gateway_positions=store.gateway_positions(),
                 flash=flash, flash_err=bool(flash_err))


@router.post("/ui/signers/link")
def signers_link(request: Request):
    result = request.app.state.signers.link_directory()
    audit_log("signers_linked", **result)
    if not result["staff"]:
        return RedirectResponse(
            "/ui/signers?flash=В Структуре нет записей с должностью. Выгрузите её"
            " из Doc-V: действие «HTTP-запрос» POST /directory/structura&flash_err=1",
            status_code=302)
    flash = (f"Структура: {result['staff']} сотрудников. Узнали "
             f"{result['matched']}, из них переименовано {result['renamed']}.")
    if result["unmatched"]:
        flash += (f" Не нашли в Doc-V: {result['unmatched']} — им подпись идёт"
                  " по записи справочника.")
    return RedirectResponse(f"/ui/signers?flash={flash}", status_code=302)


@router.post("/ui/signers/position")
def signers_position(request: Request, name: str = Form(default=""),
                     delete: str = Form(default="")):
    store = request.app.state.signers
    if delete:
        store.delete_position(delete)
        return RedirectResponse(f"/ui/signers?flash=Должность «{delete}» убрана из каталога",
                                status_code=302)
    added = store.add_position(name)
    if not added:
        return RedirectResponse("/ui/signers?flash=Пустое название&flash_err=1",
                                status_code=302)
    return RedirectResponse(f"/ui/signers?flash=Должность «{added}» в каталоге",
                            status_code=302)


@router.post("/ui/signers/person")
def signers_person(request: Request, fio: str = Form(...), position: str = Form(default=""),
                   person_id: str = Form(default="")):
    try:
        request.app.state.signers.save_person(int(person_id) if person_id else None,
                                              fio, position)
    except ValueError as exc:
        return RedirectResponse(f"/ui/signers?flash={exc}&flash_err=1", status_code=302)
    return RedirectResponse("/ui/signers?flash=Сохранено", status_code=302)


# Пустая привязка обязана иметь ровно те же ключи, что и строка из базы:
# форма читает их по имени, и недостающий ключ роняет страницу.
EMPTY_BINDING = {"id": None, "company": "", "object_name": "", "set_name": "",
                 **{f"{role}{suffix}": None if suffix == "_id" else ""
                    for role in ("soglasovano", "utverzhdayu")
                    for suffix in ("_id", "_position", "_company", "_ref", "_dept")}}


def _position_options(store) -> list[dict]:
    """Должности из Структуры для выпадающих списков — с теми, кто их
    сейчас занимает: по одному названию должности человека не узнать."""
    return [{**pos, "value": slot_value(pos["position"], pos["department"])}
            for pos in store.positions()]


def _staff_options(store) -> list[dict]:
    """Сотрудники из Структуры: ссылка идёт на uid, поэтому подпись
    останется за этим человеком, даже если его должность там сменится."""
    return [{**person, "value": slot_value(person["uid"])} for person in store.staff()]


@router.get("/ui/signers/binding/{binding_id}")
def signers_binding(request: Request, binding_id: str, flash: str = ""):
    store = request.app.state.signers
    binding = dict(EMPTY_BINDING)
    preview = ""
    if binding_id != "new":
        found = next((b for b in store.bindings() if str(b["id"]) == binding_id), None)
        if found is None:
            return RedirectResponse("/ui/signers?flash=Привязка не найдена&flash_err=1",
                                    status_code=302)
        binding = found
        preview = json.dumps(store.resolve(found["company"], found["object_name"]),
                             ensure_ascii=False, indent=1)
    slots = {role: slot_value(binding[f"{role}_ref"], binding[f"{role}_dept"],
                              binding[f"{role}_id"])
             for role in ("soglasovano", "utverzhdayu")}
    return _page(request, "signers_binding.html", "signers", binding=binding,
                 people=store.people(), set_names=sorted(store.sets()),
                 positions=_position_options(store), staff=_staff_options(store),
                 gateway_positions=store.gateway_positions(), slots=slots,
                 preview=preview, flash=flash, flash_err=False)


@router.post("/ui/signers/binding/save")
async def signers_binding_save(request: Request):
    form = await request.form()

    def role(name: str) -> dict | None:
        slot = parse_slot(form.get(f"{name}_slot"))
        if not slot["person_id"] and not slot["ref"]:
            return None
        return {**slot, "position": str(form.get(f"{name}_position") or ""),
                "company": str(form.get(f"{name}_company") or "")}

    raw_id = str(form.get("binding_id") or "").strip()
    try:
        request.app.state.signers.save_binding(
            company=str(form.get("company") or ""),
            object_name=str(form.get("object_name") or ""),
            set_name=str(form.get("set_name") or ""),
            soglasovano=role("soglasovano"), utverzhdayu=role("utverzhdayu"),
            binding_id=int(raw_id) if raw_id else None)
    except ValueError as exc:
        return RedirectResponse(f"/ui/signers?flash={exc}&flash_err=1", status_code=302)
    audit_log("signers_binding_saved", company=str(form.get("company") or ""))
    return RedirectResponse("/ui/signers?flash=Привязка сохранена", status_code=302)


@router.post("/ui/signers/binding/delete/{binding_id}")
def signers_binding_delete(request: Request, binding_id: int):
    request.app.state.signers.delete_binding(binding_id)
    audit_log("signers_binding_deleted", binding_id=binding_id)
    return RedirectResponse("/ui/signers?flash=Привязка удалена", status_code=302)


@router.get("/ui/signers/set/{name}")
def signers_set(request: Request, name: str, add: int = 0, flash: str = ""):
    store = request.app.state.signers
    rows = [{**dict(r), "slot": slot_value(r["position_ref"], r["dept_ref"],
                                           r["person_id"])}
            for r in store.sets().get(name, [])]
    rows += [{"person_id": None, "slot": "", "position": "", "print_company": "",
              "mark": "", "skip_expense_types": "", "default_position": ""}
             for _ in range(max(add, 1))]
    used = sum(1 for b in store.bindings() if b["set_name"] == name)
    return _page(request, "signers_set.html", "signers", name=name, rows=rows,
                 people=store.people(), positions=_position_options(store),
                 staff=_staff_options(store),
                 gateway_positions=store.gateway_positions(),
                 used=used, flash=flash, flash_err=False)


@router.post("/ui/signers/set/{name}/save")
async def signers_set_save(request: Request, name: str):
    form = await request.form()
    if form.get("add"):
        return RedirectResponse(f"/ui/signers/set/{name}?add=2", status_code=302)
    slots = form.getlist("slot")
    positions = form.getlist("position")
    companies = form.getlist("print_company")
    marks = form.getlist("mark")
    skips = form.getlist("skip_expense_types")
    entries = []
    for i, raw in enumerate(slots):
        slot = parse_slot(raw)
        if not slot["person_id"] and not slot["ref"]:
            continue
        entries.append({**slot, "position": positions[i], "company": companies[i],
                        "mark": marks[i], "skip_expense_types": skips[i]})
    try:
        request.app.state.signers.save_set(name, entries)
    except ValueError as exc:
        return RedirectResponse(f"/ui/signers?flash={exc}&flash_err=1", status_code=302)
    audit_log("signers_set_saved", name=name, count=len(entries))
    return RedirectResponse(f"/ui/signers/set/{name}?flash=Набор сохранён", status_code=302)
