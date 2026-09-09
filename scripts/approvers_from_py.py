"""Пересборка data/approvers.yaml из боевого utils/firmen_und_objekte.py.

Матрица подписантов живёт только на сервере и в git не попадает, поэтому
справочник для шлюза собирается на месте:

    python3 scripts/approvers_from_py.py > data/approvers.yaml

Содержимое исходного файла не меняется — он остаётся источником правды.
"""
import argparse
import importlib.util
import re
import sys
from pathlib import Path

import yaml

MATRIX_PATH = Path(__file__).resolve().parent.parent / "utils" / "firmen_und_objekte.py"


def load_matrix(path: Path):
    """Матрица подключается по пути, а не импортом пакета: сам файл лежит
    вне git, и никакого кода вокруг него в репозитории больше нет."""
    if not path.exists():
        sys.exit(f"нет файла матрицы: {path}")
    spec = importlib.util.spec_from_file_location("firmen_und_objekte", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

HEADER = """# Матрица подписантов реестров платежей.
#
# Собрана скриптом scripts/approvers_from_py.py из боевого
# utils/firmen_und_objekte.py (он в git не хранится). После правки
# матрицы на сервере пересоберите этот файл и перезапустите шлюз.
#
# approver_lists — именованные списки согласующих; число — ID строки на
# листе СПР_ПОДПИСАНТОВ шаблона template.xlsx. Порядок = порядок подписей.
# ID 3 особый: перед ним печатается строка «СОГЛАСОВАНО».
# rules — объекты компании, два директора и имя списка;
# company_fallbacks — если пара компания+объект не найдена.
"""


def main() -> None:
    parser = argparse.ArgumentParser(description="Пересборка data/approvers.yaml")
    parser.add_argument("--matrix", type=Path, default=MATRIX_PATH,
                        help=f"путь к матрице (по умолчанию {MATRIX_PATH})")
    fu = load_matrix(parser.parse_args().matrix)

    lists, by_ids = {}, {}
    for n in range(1, 10):
        values = list(getattr(fu, f"approver_ids_{n}").values())
        lists[f"list_{n}"] = values
        by_ids[tuple(values)] = f"list_{n}"

    captured: dict = {}

    def tracer(frame, event, arg):
        if event == "return" and frame.f_code.co_name == "check_company_object_pair":
            captured.update(frame.f_locals)
        return tracer

    sys.settrace(tracer)
    fu.check_company_object_pair("-", "-")
    sys.settrace(None)

    def fix(name: str) -> str:
        return re.sub(r"^TOO ", "ТОО ", name)  # латинское TOO встречается опечаткой

    rules: dict = {}
    for (company, obj), (directors, approvers) in captured["company_object_pairs"].items():
        key = (fix(company), tuple(directors), by_ids[tuple(approvers)])
        rules.setdefault(key, []).append(obj)

    out = {
        "approver_lists": lists,
        "rules": [{"company": c, "directors": list(d), "approvers": lst, "objects": objs}
                  for (c, d, lst), objs in rules.items()],
        "company_fallbacks": [
            {"company": fix(c), "directors": list(d), "approvers": by_ids[tuple(a)]}
            for c, (d, a) in captured["company_list"].items()],
        "expense_type_exclusions": {
            "expense_types": ["коммерческие расходы", "зарплата", "налоги"],
            "remove_ids": [4, 6],
        },
    }
    sys.stdout.write(HEADER + yaml.safe_dump(out, allow_unicode=True, sort_keys=False,
                                             width=100))


if __name__ == "__main__":
    main()
