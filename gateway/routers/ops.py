"""POST /ops/{name} — выполнение именованной операции из ops.yaml.

HTTP 200 и при неуспехе команды: факт ошибки — в поле ok, чтобы
действие «HTTP-запрос» в Doc-V записало результат в поле, а не упало.
422 — только за неверные параметры, 404 — за неизвестную операцию.
"""
from fastapi import APIRouter, Body, HTTPException, Request
from fastapi.concurrency import run_in_threadpool

from ..opsrunner.runner import OpsValidationError, run_operation

router = APIRouter()


@router.get("/ops")
def list_ops(request: Request):
    return {name: op.description for name, op in request.app.state.ops.items()}


@router.post("/ops/{name}")
async def execute(name: str, request: Request, body: dict = Body(default={})):
    op = request.app.state.ops.get(name)
    if op is None:
        raise HTTPException(status_code=404, detail="нет такой операции")
    request.app.state.heartbeat.touch("ops")
    params = body.get("params") or {}
    ip = request.client.host if request.client else ""
    try:
        return await run_in_threadpool(
            run_operation, op, params, request.app.state.filestore, client_ip=ip,
            history=request.app.state.ops_history, source="api")
    except OpsValidationError as exc:
        raise HTTPException(status_code=422, detail=str(exc))
