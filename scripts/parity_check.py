"""Сверка старого сервиса и шлюза на одном и том же payload.

Запускать на сервере перед переключением действия Doc-V:

    python3 scripts/parity_check.py tests/model.json
    python3 scripts/parity_check.py real.json --kind inner

Оба рендера получают одни и те же данные, результаты сравниваются
ячейка в ячейку по всем листам. Время генерации (G10) пропускается —
оно заведомо разное. Расхождений нет — можно переключать.
"""
import argparse
import io
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import json  # noqa: E402

import openpyxl  # noqa: E402

TEMPLATES = ROOT / "templates" / "excel"
SKIP_CELLS = {"G10"}  # отметка времени формирования


def _cells(workbook) -> dict:
    buf = io.BytesIO()
    workbook.save(buf)
    workbook.close()
    loaded = openpyxl.load_workbook(io.BytesIO(buf.getvalue()))
    return {
        (sheet.title, cell.coordinate): cell.value
        for sheet in loaded.worksheets
        for row in sheet.iter_rows()
        for cell in row
        if cell.value is not None and cell.coordinate not in SKIP_CELLS
    }


def _cases(entries: list[dict]):
    from gateway.renderers.approvers import ApproverMatrix
    from gateway.renderers.registry_inner import render_inner
    from gateway.renderers.registry_outer import load_banks, render_outer
    from gateway.renderers.registry_priority import render_priority
    from models.inner_registry import format_excel_inner
    from models.outer_registry import format_excel_outer
    from models.priority_registry import fill_priority_registry

    matrix = ApproverMatrix(ROOT / "data" / "approvers.yaml")
    banks = load_banks(ROOT / "data" / "banks.yaml")
    ordered = sorted(entries, key=lambda x: x.get("object_name") or "")
    return {
        "inner": (lambda: format_excel_inner({"request": ordered}),
                  lambda: render_inner(ordered, TEMPLATES / "template.xlsx", matrix)),
        "outer": (lambda: format_excel_outer({"request": entries}),
                  lambda: render_outer(entries, TEMPLATES / "template_outer.xlsx", banks)),
        "priority": (lambda: fill_priority_registry({"request": ordered}),
                     lambda: render_priority(ordered, TEMPLATES / "template_priority_registry.xlsx")),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("payload", help="JSON вида {\"request\": [...]}")
    parser.add_argument("--kind", choices=("inner", "outer", "priority"),
                        action="append", help="сверить только этот реестр")
    args = parser.parse_args()

    data = json.loads(Path(args.payload).read_text(encoding="utf-8"))
    entries = data.get("request") if isinstance(data, dict) else data
    if not entries:
        print("в payload нет request[]")
        return 2

    cases = _cases(entries)
    failed = False
    for kind in args.kind or list(cases):
        old_fn, new_fn = cases[kind]
        try:
            old = _cells(old_fn())
        except Exception as exc:
            print(f"[{kind}] СТАРЫЙ не отработал: {exc!r} — сверять не с чем, "
                  f"проверьте новый вручную")
            old = None
        try:
            new = _cells(new_fn())
        except Exception as exc:
            print(f"[{kind}] НОВЫЙ не отработал: {exc!r} — переключать нельзя")
            failed = True
            continue
        if old is None:
            print(f"[{kind}] новый отработал: ячеек {len(new)}")
            continue
        diffs = sorted(set(old) | set(new))
        diffs = [key for key in diffs if old.get(key) != new.get(key)]
        mark = "совпадает" if not diffs else f"РАСХОЖДЕНИЙ {len(diffs)}"
        print(f"[{kind}] ячеек {len(new)}, {mark}")
        for key in diffs[:30]:
            print(f"    {key[0]}!{key[1]}: старый={old.get(key)!r} новый={new.get(key)!r}")
        if len(diffs) > 30:
            print(f"    … ещё {len(diffs) - 30}")
        failed = failed or bool(diffs)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
