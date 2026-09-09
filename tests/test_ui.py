import io
import json
import textwrap
from pathlib import Path

from gateway.opsrunner.registry import load_registry

from conftest import TOKEN_ADMIN, docv_headers

MODEL = json.loads((Path(__file__).parent / "data" / "model.json").read_text(encoding="utf-8"))


def _login(client):
    r = client.post("/ui/login", data={"token": TOKEN_ADMIN}, follow_redirects=False)
    assert r.status_code == 302
    return client


def test_ui_requires_login(client):
    r = client.get("/ui", follow_redirects=False)
    assert r.status_code == 302 and r.headers["location"] == "/ui/login"
    assert client.post("/ui/jobs/ack/1", follow_redirects=False).status_code == 403


def test_wrong_token_stays_on_login(client):
    r = client.post("/ui/login", data={"token": "мимо"})
    assert r.status_code == 200 and "Неверный токен" in r.text
    assert client.get("/ui", follow_redirects=False).status_code == 302


def test_dashboard_after_login(client):
    _login(client)
    r = client.get("/ui")
    assert r.status_code == 200 and "Обзор" in r.text


def test_jobs_roundtrip(client):
    _login(client)
    r = client.post("/ui/jobs/new",
                    data={"type": "тест", "payload": '{"a": 1}'},
                    follow_redirects=True)
    assert "в очереди" in r.text and "тест" in r.text
    r = client.post("/ui/jobs/ack/1", follow_redirects=True)
    assert "подтверждено" in r.text
    assert client.app.state.jobs.stats() == {"acked": 1}


def test_jobs_bad_payload(client):
    _login(client)
    r = client.post("/ui/jobs/new", data={"type": "т", "payload": "не json"})
    assert "не JSON-объект" in r.text


def test_files_upload_and_delete(client):
    _login(client)
    r = client.post("/ui/files/upload",
                    files={"uploads": ("устав.pdf", io.BytesIO(b"%PDF-fake"), "application/pdf")},
                    follow_redirects=True)
    assert "устав.pdf" in r.text
    token = client.app.state.filestore.list_files()[0]["token"]
    r = client.post(f"/ui/files/delete/{token}", follow_redirects=True)
    assert "устав.pdf" not in r.text
    assert client.app.state.filestore.list_files() == []


def test_ops_run_via_ui(client, tmp_path):
    _login(client)
    ops_file = tmp_path / "ops.yaml"
    ops_file.write_text(textwrap.dedent("""
    operations:
      echo:
        argv: ["/bin/echo", "{text}"]
        params:
          text: {pattern: "^[-\\\\w\\\\s.,]{1,100}$"}
    """), encoding="utf-8")
    client.app.state.ops = load_registry(ops_file)
    r = client.get("/ui/ops")
    assert "echo" in r.text
    r = client.post("/ui/ops/echo", data={"text": "привет"})
    assert r.status_code == 200 and "привет" in r.text
    r = client.post("/ui/ops/echo", data={"text": "x; rm -rf /"})
    assert "Параметры не приняты" in r.text


def test_render_registry_via_ui(client):
    _login(client)
    r = client.post("/ui/render/registry",
                    data={"kind": "inner", "data": json.dumps(MODEL, ensure_ascii=False)},
                    follow_redirects=True)
    assert r.status_code == 200 and "Скачать" in r.text  # кнопка со ссылкой на файл
    files = client.app.state.filestore.list_files()
    assert files and files[0]["suffix"] == ".xlsx"


def test_render_bad_json_flash(client):
    _login(client)
    r = client.post("/ui/render/registry", data={"kind": "inner", "data": "мусор"})
    assert "Нужен JSON" in r.text


def test_typst_pages_and_editing(client):
    _login(client)
    r = client.get("/ui/typst")
    assert r.status_code == 200 and "primer" in r.text  # засеян из templates/typst/
    r = client.post("/ui/typst/create", data={"name": "akt_sverki"}, follow_redirects=True)
    assert "akt_sverki" in r.text
    r = client.post("/ui/typst/akt_sverki/save", data={"source": "= Акт сверки v2"})
    assert "Сохранено" in r.text
    client.post("/ui/typst/akt_sverki/save", data={"source": "= Акт сверки v3"})
    hist = client.app.state.typst_store.history("akt_sverki")
    assert len(hist) == 2
    client.post(f"/ui/typst/akt_sverki/restore/{hist[0]['id']}")
    assert client.app.state.typst_store.get("akt_sverki") == "= Акт сверки v2"
    r = client.post("/ui/typst/akt_sverki/delete", follow_redirects=True)
    assert "akt_sverki" not in r.text


def test_typst_asset_upload_delete(client):
    _login(client)
    r = client.post("/ui/typst/assets/upload",
                    files={"uploads": ("logo.png", io.BytesIO(b"\x89PNG"), "image/png")},
                    follow_redirects=True)
    assert "assets/logo.png" in r.text
    r = client.post("/ui/typst/assets/upload",
                    files={"uploads": ("hack.sh", io.BytesIO(b"#!"), "text/plain")},
                    follow_redirects=True)
    assert "расширение" in r.text
    client.post("/ui/typst/assets/delete/logo.png")
    assert client.app.state.typst_store.list_assets() == []


def test_dashboard_heartbeat_card(client):
    _login(client)
    r = client.get("/ui")
    assert "Связь с Doc-V" in r.text and "ещё не было" in r.text
    client.get("/jobs/pending", headers={"Authorization": "Bearer test-docv-token"})
    r = client.get("/ui")
    assert "с назад" in r.text


def test_asset_raw_and_thumbnails(client):
    _login(client)
    client.app.state.typst_store.save_asset("logo.png", b"\x89PNG-data")
    r = client.get("/ui/typst/assets/raw/logo.png")
    assert r.status_code == 200 and r.headers["content-type"] == "image/png"
    assert client.get("/ui/typst/assets/raw/net.png").status_code == 404
    r = client.get("/ui/typst")
    assert "/ui/typst/assets/raw/logo.png" in r.text  # миниатюра в списке
    r = client.get("/ui/typst/primer")
    assert '#image(&#34;assets/logo.png&#34;)' in r.text or 'assets/logo.png' in r.text


def test_file_to_assets_one_click(client):
    _login(client)
    store = client.app.state.filestore
    token = store.save_bytes(b"\x89PNG-blank", ".png", "blank_shar.png")
    r = client.get("/ui/files")
    assert f"/ui/files/to_assets/{token}" in r.text  # кнопка есть у картинки
    r = client.post(f"/ui/files/to_assets/{token}", follow_redirects=True)
    assert "доступна шаблонам" in r.text
    assert client.app.state.typst_store.assets_bytes()["blank_shar.png"] == b"\x89PNG-blank"
    # не-картинка кнопки не имеет
    t2 = store.save_bytes(b"x", ".xlsx", "реестр.xlsx")
    assert f"/ui/files/to_assets/{t2}" not in client.get("/ui/files").text


def test_file_to_assets_translit_name(client):
    _login(client)
    token = client.app.state.filestore.save_bytes(b"\x89PNG", ".png", "логотип шар.png")
    r = client.post(f"/ui/files/to_assets/{token}", follow_redirects=True)
    assert "assets/logotip_shar.png" in r.text  # кириллица транслитерирована
    assert "logotip_shar.png" in client.app.state.typst_store.assets_bytes()


def test_files_multi_upload(client):
    _login(client)
    r = client.post("/ui/files/upload", files=[
        ("uploads", ("а.txt", io.BytesIO(b"1"), "text/plain")),
        ("uploads", ("б.txt", io.BytesIO(b"2"), "text/plain")),
    ], follow_redirects=True)
    assert "Загружено файлов: 2" in r.text
    assert len(client.app.state.filestore.list_files()) == 2


def test_files_rename_keeps_extension(client):
    _login(client)
    token = client.app.state.filestore.save_bytes(b"x", ".xlsx", "реестр.xlsx")
    r = client.post(f"/ui/files/rename/{token}", data={"new_name": "реестр август"},
                    follow_redirects=True)
    assert "Переименовано" in r.text
    _, name = client.app.state.filestore.resolve(token)
    assert name == "реестр август.xlsx"


def test_files_download_zip(client):
    import zipfile
    _login(client)
    store = client.app.state.filestore
    t1 = store.save_bytes("один".encode(), ".txt", "документ.txt")
    t2 = store.save_bytes("два".encode(), ".txt", "документ.txt")  # дубль имени
    r = client.post("/ui/files/download_zip", data={"tokens": [t1, t2]})
    assert r.status_code == 200
    assert r.headers["content-type"] == "application/zip"
    zf = zipfile.ZipFile(io.BytesIO(r.content))
    assert sorted(zf.namelist()) == ["документ.txt", "документ_1.txt"]
    assert zf.read("документ.txt") == "один".encode()
    # пустой выбор — просто возврат на страницу
    r = client.post("/ui/files/download_zip", data={}, follow_redirects=True)
    assert "Ничего не выбрано" in r.text


def test_dashboard_shows_directories(client):
    _login(client)
    assert "Пока не приходили" in client.get("/ui").text
    client.app.state.directory.replace("structura", [
        {"uid": "u1", "display_name": "Абдрахманова Х.М.", "position": "Гл. бухгалтер",
         "department": "Бухгалтерия"}])
    r = client.get("/ui")
    assert "structura" in r.text and "1 записей" in r.text  # имя и счётчик


def test_ops_builder_creates_and_reloads(client, tmp_path):
    """Операция создаётся из интерфейса и сразу доступна — без перезапуска."""
    _login(client)
    ops_file = tmp_path / "ops.yaml"
    ops_file.write_text("operations: {}\n", encoding="utf-8")
    client.app.state.ops_path = ops_file
    client.app.state.ops = {}

    r = client.post("/ui/opsedit/save", follow_redirects=True, data={
        "new_name": "echo_test", "description": "эхо", "timeout_sec": "10",
        "argv": ["/bin/echo", "{text}"],
        "p_name": ["text"], "p_type": ["str"],
        "p_pattern": [r"^[-\w\s]{1,50}$"], "p_required": ["on"]})
    assert "сохранена" in r.text
    assert "echo_test" in client.app.state.ops           # реестр перечитан на лету
    assert "echo_test" in ops_file.read_text(encoding="utf-8")  # и записан в файл

    # операция работает сразу
    r = client.post("/ui/ops/echo_test", data={"text": "привет"})
    assert "привет" in r.text

    r = client.post("/ui/opsedit/delete/echo_test", follow_redirects=True)
    assert "echo_test" not in client.app.state.ops


def test_ops_builder_rejects_dangerous_command(client, tmp_path):
    """argv[0] вне разрешённых каталогов не сохраняется."""
    _login(client)
    client.app.state.ops_path = tmp_path / "ops.yaml"
    before = dict(client.app.state.ops)
    r = client.post("/ui/opsedit/save", data={
        "new_name": "hack", "argv": ["/home/radmin/evil.sh"], "timeout_sec": "10"})
    assert "argv[0]" in r.text
    assert client.app.state.ops == before


def test_api_console_calls_own_endpoint(client):
    _login(client)
    r = client.post("/ui/api", data={"method": "GET", "path": "/health", "body": ""})
    assert r.status_code == 200
    # Jinja экранирует кавычки, поэтому сравниваем по содержимому
    assert "status" in r.text and "ok" in r.text
    assert 'class="chip success">200' in r.text
    assert "curl -X GET" in r.text
    assert client.app.state.apilog.items()[0]["path"] == "/health"


def test_api_console_refuses_foreign_host(client):
    _login(client)
    r = client.post("/ui/api", data={"method": "GET", "path": "http://evil.example.com/x"})
    assert "только точки шлюза" in r.text


def test_settings_apply_live(client):
    _login(client)
    r = client.post("/ui/settings", data={"file_ttl_hours": "48", "lease_seconds": "90"},
                    follow_redirects=True)
    assert "применены" in r.text
    assert client.settings.file_ttl_hours == 48 and client.settings.lease_seconds == 90
    assert (client.settings.var_dir / "settings.json").is_file()
    # границы соблюдаются
    r = client.post("/ui/settings", data={"file_ttl_hours": "99999"}, follow_redirects=True)
    assert "допустимо от" in r.text
    assert client.settings.file_ttl_hours == 48


def test_settings_page_masks_tokens(client):
    _login(client)
    text = client.get("/ui/settings").text
    assert "GW_TOKEN_DOCV" in text
    assert "test-docv-token" not in text  # значение не раскрывается


def test_dates_are_human_and_local(client):
    """В интерфейсе не должно оставаться ISO-дат в UTC."""
    _login(client)
    client.app.state.filestore.save_bytes(b"x", ".txt", "файл.txt")
    text = client.get("/ui/files").text
    assert "+00:00" not in text and "T08:" not in text
    import re
    assert re.search(r"\d\d\.\d\d\.20\d\d \d\d:\d\d", text)  # дд.мм.гггг чч:мм
    assert "назад" in text or "только что" in text


def test_jobs_search_and_payload_visible(client):
    _login(client)
    q = client.app.state.jobs
    q.enqueue(producer="1c", job_type="оплата", payload={"счёт": "KZ12"},
              idempotency_key=None)
    q.enqueue(producer="ui", job_type="тест", payload={"x": 1}, idempotency_key=None)

    text = client.get("/ui/jobs").text
    assert "KZ12" in text  # payload виден прямо в строке

    only = client.get("/ui/jobs?search=оплата").text
    assert "оплата" in only and ">тест<" not in only
    by_producer = client.get("/ui/jobs?producer=ui").text
    assert "тест" in by_producer and "оплата" not in by_producer
    # поиск по содержимому payload
    assert "оплата" in client.get("/ui/jobs?search=KZ12").text
    assert "Ничего не нашлось" in client.get("/ui/jobs?search=неттакого").text


def test_files_search_and_paging(client):
    _login(client)
    store = client.app.state.filestore
    for i in range(55):
        store.save_bytes(b"x", ".txt", f"отчёт-{i}.txt")
    store.save_bytes(b"x", ".txt", "реестр.txt")

    page = client.get("/ui/files").text
    assert "Показать ещё" in page              # постранично, а не всё сразу
    assert page.count("копировать ссылку") == 50
    assert "всего 56" in page

    found = client.get("/ui/files?search=реестр").text
    assert "реестр.txt" in found and "Показать ещё" not in found


def test_signers_page_lists_bindings(client):
    _login(client)
    r = client.get("/ui/signers")
    assert r.status_code == 200
    assert "Подписанты" in r.text and "Шар" in r.text
    # поиск сужает список
    narrow = client.get("/ui/signers", params={"search": "СМУ Аргон"})
    assert narrow.status_code == 200 and "СМУ Аргон" in narrow.text


def test_signers_binding_edit_roundtrip(client):
    _login(client)
    store = client.app.state.signers
    person = store.people()[0]
    r = client.post("/ui/signers/binding/save", follow_redirects=False, data={
        "binding_id": "", "company": "ТОО «Новая»", "object_name": "",
        "set_name": sorted(store.sets())[0],
        "soglasovano_slot": "", "soglasovano_position": "", "soglasovano_company": "",
        "utverzhdayu_slot": f"p:{person['id']}", "utverzhdayu_position": "Директор",
        "utverzhdayu_company": "ТОО «Новая»"})
    assert r.status_code == 302
    resolved = store.resolve("ТОО «Новая»")
    assert resolved["utverzhdayu"]["position"] == "Директор"
    assert resolved["coordinators"]

    binding = next(b for b in store.bindings() if b["company"] == "ТОО «Новая»")
    page = client.get(f"/ui/signers/binding/{binding['id']}")
    assert page.status_code == 200 and "Директор" in page.text

    client.post(f"/ui/signers/binding/delete/{binding['id']}", follow_redirects=False)
    assert store.resolve("ТОО «Новая»")["utverzhdayu"] is None


def test_signers_set_edit_changes_registry_signatures(client):
    _login(client)
    store = client.app.state.signers
    name = sorted(store.sets())[0]
    page = client.get(f"/ui/signers/set/{name}")
    assert page.status_code == 200
    person = store.people()[0]
    r = client.post(f"/ui/signers/set/{name}/save", follow_redirects=False, data={
        "slot": f"p:{person['id']}", "position": "Проверяющий",
        "print_company": "ТОО «Тест»", "mark": "", "skip_expense_types": ""})
    assert r.status_code == 302
    entries = store.sets()[name]
    assert len(entries) == 1 and entries[0]["position"] == "Проверяющий"


def test_signers_link_without_directory_says_so(client):
    _login(client)
    r = client.post("/ui/signers/link", follow_redirects=True)
    assert r.status_code == 200 and "Выгрузите её" in r.text


def test_signers_link_reports_what_it_found(client):
    _login(client)
    client.post("/directory/structura", headers=docv_headers(), json={"items": [
        {"uid": "u-1", "display_name": "Аманов Бауыржан Шарипович",
         "position": "Генеральный директор", "department": "Дирекция"},
        {"uid": "u-2", "display_name": "Иванов Иван Иванович",
         "position": "Кладовщик", "department": "Склад"}]})
    r = client.post("/ui/signers/link", follow_redirects=True)
    assert "Структура: 2 сотрудников" in r.text
    assert "Узнали 1" in r.text and "переименовано 1" in r.text
    assert "Не нашли в Doc-V" in r.text
