"""Исполнение операций из реестра: валидация параметров, сборка argv
без shell, таймаут, сбор порождённых файлов в filestore, аудит."""
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from ..config import APP_DIR
from ..filestore.store import TOKEN_RE, FileStore
from ..logging_setup import audit_log
from .registry import Operation

log = logging.getLogger(__name__)


class OpsValidationError(Exception):
    """Ошибки параметров -> HTTP 422."""


def run_operation(op: Operation, params: dict, filestore: FileStore,
                  *, client_ip: str = "", history=None, source: str = "ui") -> dict:
    """history — OpsHistory: запуск записывается туда, а его номер
    возвращается в result["run_id"]. Ошибки параметров в историю не
    попадают: это не запуск."""
    values = _validate(op, params, filestore)
    audit_log("ops_start", op=op.name, params={k: str(v)[:200] for k, v in params.items()},
              ip=client_ip)

    with tempfile.TemporaryDirectory(prefix=f"ops-{op.name}-") as workdir_str:
        workdir = Path(workdir_str)
        inputs = workdir / "inputs"
        argv = _build_argv(op, values, workdir, inputs)

        started = time.monotonic()
        try:
            proc = subprocess.run(
                argv, cwd=workdir, capture_output=True, text=True,
                timeout=op.timeout_sec,
                env={"PATH": os.environ.get("PATH", "/usr/bin:/bin"),
                     "HOME": os.environ.get("HOME", str(workdir)),
                     "LANG": "ru_RU.UTF-8"},
            )
            exit_code, stdout, stderr = proc.returncode, proc.stdout, proc.stderr
            error = None
        except subprocess.TimeoutExpired as exc:
            exit_code, error = -1, "timeout"
            stdout = (exc.stdout or b"").decode("utf-8", "replace") if isinstance(exc.stdout, bytes) else (exc.stdout or "")
            stderr = (exc.stderr or b"").decode("utf-8", "replace") if isinstance(exc.stderr, bytes) else (exc.stderr or "")
        except FileNotFoundError as exc:
            exit_code, error, stdout, stderr = -1, f"команда не найдена: {exc.filename}", "", ""
        duration_ms = int((time.monotonic() - started) * 1000)

        limit = op.max_output_kb * 1024
        truncated = len(stdout) > limit or len(stderr) > limit
        stdout, stderr = stdout[:limit], stderr[:limit]

        files = []
        if error is None and exit_code == 0 and op.collect:
            for produced in sorted(workdir.glob(op.collect)):
                if produced.is_file() and not produced.is_relative_to(inputs):
                    token = filestore.save_file(produced)
                    files.append({"name": produced.name,
                                  "download_url": filestore.download_url(token)})

    ok = error is None and exit_code == 0
    audit_log("ops_finish", op=op.name, ok=ok, exit_code=exit_code,
              duration_ms=duration_ms, files=len(files))
    result = {"op": op.name, "ok": ok, "exit_code": exit_code, "stdout": stdout,
              "stderr": stderr, "duration_ms": duration_ms, "truncated": truncated,
              "files": files}
    if error:
        result["error"] = error
    if history is not None:
        result["run_id"] = history.add(result, params, ip=client_ip, source=source)
    return result


def _validate(op: Operation, params: dict, filestore: FileStore) -> dict:
    if not isinstance(params, dict):
        raise OpsValidationError("params должен быть объектом")
    if unknown := set(params) - set(op.params):
        raise OpsValidationError(f"незадекларированные параметры: {sorted(unknown)}")
    values = {}
    for pdef in op.params.values():
        if pdef.name not in params:
            if pdef.required:
                raise OpsValidationError(f"не хватает параметра {pdef.name!r}")
            continue
        raw = params[pdef.name]
        if pdef.type == "str":
            raw = str(raw)
            if not pdef.pattern.fullmatch(raw):
                raise OpsValidationError(f"параметр {pdef.name!r} не прошёл проверку формата")
            values[pdef.name] = raw
        elif pdef.type == "file":
            values[pdef.name] = _resolve_file(raw, pdef.name, filestore)
        elif pdef.type == "file_list":
            if not isinstance(raw, list) or not raw:
                raise OpsValidationError(f"параметр {pdef.name!r}: нужен непустой список токенов")
            if len(raw) > pdef.max_items:
                raise OpsValidationError(f"параметр {pdef.name!r}: больше {pdef.max_items} файлов")
            values[pdef.name] = [_resolve_file(t, pdef.name, filestore) for t in raw]
    return values


def _resolve_file(token, pname: str, filestore: FileStore) -> tuple:
    token = str(token)
    if not TOKEN_RE.fullmatch(token):
        raise OpsValidationError(f"параметр {pname!r}: это не токен файла")
    resolved = filestore.resolve(token)
    if resolved is None:
        raise OpsValidationError(f"параметр {pname!r}: файл {token} не найден или истёк")
    return resolved  # (path, orig_name)


def _build_argv(op: Operation, values: dict, workdir: Path, inputs: Path) -> list[str]:
    """Файловые параметры копируются в workdir/inputs под человеческими
    именами (скрипты не трогают хранилище), встроенные плейсхолдеры
    подставляются внутри элементов, пользовательские — целым элементом."""
    builtins = {"workdir": str(workdir), "app_dir": str(APP_DIR), "python": sys.executable}

    def materialize(resolved: tuple) -> str:
        inputs.mkdir(exist_ok=True)
        src, orig_name = resolved
        dest = inputs / orig_name
        n = 1
        while dest.exists():
            dest = inputs / f"{n}_{orig_name}"
            n += 1
        shutil.copy(src, dest)
        return str(dest)

    argv = []
    for element in op.argv:
        matched = re.fullmatch(r"\{([a-z_]+)(\.\.\.)?\}", element)
        pname = matched.group(1) if matched else None
        if pname and pname in op.params:
            pdef = op.params[pname]
            if pname not in values:
                continue  # необязательный параметр не передан
            if pdef.type == "file_list":
                argv.extend(materialize(r) for r in values[pname])
            elif pdef.type == "file":
                argv.append(materialize(values[pname]))
            else:
                argv.append(values[pname])
        else:
            for key, val in builtins.items():
                element = element.replace(f"{{{key}}}", val)
            argv.append(element)
    return argv
