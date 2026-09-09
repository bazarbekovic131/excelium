"""Doc-V Gateway: FastAPI-приложение."""
import asyncio
import contextlib
import logging
import os

from fastapi import FastAPI
from fastapi import Response
from fastapi.responses import JSONResponse, RedirectResponse

from .filestore.store import FileStore
from .jobsqueue.db import init_db
from .logging_setup import setup_logging
from .config import APP_DIR, Settings
from .apilog import ApiLog
from .directory import DirectoryStore
from .heartbeat import Heartbeat
from .renderers.registry_outer import load_banks
from .renderers.typst_store import TypstStore
from .renderers.typst_renderer import configure as configure_typst
from .renderers.typst_renderer import typst_available, typst_binary
from .opsrunner.registry import load_registry
from .settings_store import SettingsStore
from .signers import SignerStore
from .jobsqueue.service import JobQueue
from .routers import debug, directory, files, jobs, ops, render, signers, ui
from .security import SecurityMiddleware

log = logging.getLogger(__name__)

SWEEP_INTERVAL_SEC = 15 * 60


def create_app(settings: Settings | None = None) -> FastAPI:
    settings = settings or Settings()
    setup_logging(settings.var_dir)

    @contextlib.asynccontextmanager
    async def lifespan(app: FastAPI):
        init_db(settings.db_path)
        app.state.settings = settings
        app.state.filestore = FileStore(settings)
        app.state.signers = SignerStore(settings.db_path)
        app.state.banks = load_banks(APP_DIR / "data" / "banks.yaml")
        app.state.template_inner = APP_DIR / "templates" / "excel" / "template.xlsx"
        app.state.template_outer = APP_DIR / "templates" / "excel" / "template_outer.xlsx"
        app.state.template_priority = APP_DIR / "templates" / "excel" / "template_priority_registry.xlsx"
        app.state.ops_path = APP_DIR / "ops.yaml"
        app.state.ops = load_registry(app.state.ops_path)
        app.state.settings_store = SettingsStore(settings, settings.var_dir / "settings.json")
        app.state.settings_store.load()
        app.state.apilog = ApiLog()
        app.state.jobs = JobQueue(settings.db_path, lease_seconds=settings.lease_seconds,
                                  keep_days=settings.jobs_keep_days)
        app.state.heartbeat = Heartbeat(settings.db_path)
        app.state.directory = DirectoryStore(settings.db_path)
        # Разовый перенос состава подписантов из прежних источников:
        # структура — из data/signers_seed.yaml, ФИО и должности — со
        # скрытого листа шаблона. Дальше состав живёт в базе и правится
        # в /ui, а листы СПР в выдачу больше не попадают.
        if app.state.signers.is_empty():
            app.state.signers.seed(APP_DIR / "data" / "signers_seed.yaml",
                                   app.state.template_inner)
        app.state.signers.link_directory()
        app.state.typst_store = TypstStore(settings.db_path)
        # list_soglasovaniya переименован в contract_card: переносим вместе
        # с правками администратора, чтобы они не потерялись
        if app.state.typst_store.rename("list_soglasovaniya", "contract_card"):
            log.info("шаблон list_soglasovaniya переименован в contract_card")
        seeded = app.state.typst_store.seed_from_dir(APP_DIR / "templates" / "typst")
        if seeded:
            log.info("typst-шаблоны импортированы из файлов",
                     extra={"data": {"count": seeded}})
        task = asyncio.create_task(_sweep_loop(app))
        configure_typst(settings.typst_bin)
        binary = typst_binary()
        if binary:
            log.info("typst найден", extra={"data": {"path": binary}})
        else:
            log.warning("бинарь typst не найден — /render/typst будет отвечать 503; "
                        "укажите GW_TYPST_BIN или добавьте каталог в PATH юнита",
                        extra={"data": {"looked_for": settings.typst_bin,
                                        "PATH": os.environ.get("PATH", "")}})
        # Списки — в лог при старте: так видно, что прочитал ИМЕННО
        # работающий процесс (journalctl -u docv-gateway | grep конфигурация)
        log.info("конфигурация прочитана", extra={"data": {
            "allowlist_api": settings.allowlist,
            "allowlist_ui": settings.ui_allowlist,
            "base_url": settings.base_url,
            "ui_enabled": bool(settings.token_admin),
            "env_file": str(APP_DIR / ".env"),
        }})
        log.info("gateway started")
        yield
        task.cancel()
        with contextlib.suppress(asyncio.CancelledError):
            await task

    app = FastAPI(title="Doc-V Gateway", lifespan=lifespan)
    app.add_middleware(SecurityMiddleware, settings=settings)
    app.include_router(files.router)
    app.include_router(render.router)
    app.include_router(ops.router)
    app.include_router(jobs.router)
    app.include_router(ui.router)
    app.include_router(debug.router)
    app.include_router(directory.router)
    app.include_router(signers.router)

    @app.get("/", include_in_schema=False)
    def root():
        return RedirectResponse("/ui")

    @app.get("/health")
    def health():
        return {"status": "ok"}

    @app.get("/favicon.ico", include_in_schema=False)
    def favicon():
        return Response(status_code=204)

    @app.exception_handler(Exception)
    async def unhandled(request, exc):
        log.exception("unhandled error at %s", request.url.path)
        return JSONResponse({"error": "internal", "detail": str(exc)}, status_code=500)

    return app


async def _sweep_loop(app: FastAPI) -> None:
    while True:
        await asyncio.sleep(SWEEP_INTERVAL_SEC)
        try:
            await asyncio.to_thread(app.state.filestore.sweep)
            await asyncio.to_thread(app.state.jobs.sweep)
            await asyncio.to_thread(app.state.directory.sweep)
        except Exception:
            log.exception("sweep failed")


app = create_app()
