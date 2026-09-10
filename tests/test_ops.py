import textwrap

import pytest
from fastapi.testclient import TestClient

from gateway.main import create_app
from gateway.opsrunner.registry import OpsConfigError, load_registry

from conftest import docv_headers, ops_headers

TEST_OPS = textwrap.dedent("""
operations:
  echo:
    description: эхо
    argv: ["/bin/echo", "{text}"]
    params:
      text: {pattern: "^[-\\\\w\\\\s.,]{1,100}$"}
  slow:
    argv: ["{python}", "-c", "import time; time.sleep(30)"]
    timeout_sec: 1
  makefile:
    argv: ["{python}", "-c", "open('out.txt','w').write('готово')"]
    collect: "*.txt"
  pack:
    argv: ["{python}", "{app_dir}/scripts/zip_bundle.py", "{files...}"]
    params:
      files: {type: file_list, max_items: 3}
    collect: "*.zip"
""")


@pytest.fixture()
def ops_client(settings, tmp_path, monkeypatch):
    ops_file = tmp_path / "ops.yaml"
    ops_file.write_text(TEST_OPS, encoding="utf-8")
    app = create_app(settings)
    with TestClient(app) as c:
        c.app.state.ops = load_registry(ops_file)
        c.settings = settings
        yield c


def test_echo_ok(ops_client):
    r = ops_client.post("/ops/echo", json={"params": {"text": "привет мир"}},
                        headers=ops_headers())
    assert r.status_code == 200
    body = r.json()
    assert body["ok"] and body["exit_code"] == 0
    assert body["stdout"].strip() == "привет мир"


def test_unknown_op_404(ops_client):
    assert ops_client.post("/ops/rm_rf", json={}, headers=ops_headers()).status_code == 404


def test_injection_rejected(ops_client):
    r = ops_client.post("/ops/echo", json={"params": {"text": "x; rm -rf /"}},
                        headers=ops_headers())
    assert r.status_code == 422


def test_undeclared_param_rejected(ops_client):
    r = ops_client.post("/ops/echo", json={"params": {"text": "ok", "extra": "1"}},
                        headers=ops_headers())
    assert r.status_code == 422


def test_missing_param_rejected(ops_client):
    assert ops_client.post("/ops/echo", json={}, headers=ops_headers()).status_code == 422


def test_timeout(ops_client):
    r = ops_client.post("/ops/slow", json={}, headers=ops_headers())
    body = r.json()
    assert r.status_code == 200 and not body["ok"] and body["error"] == "timeout"


def test_collect_files(ops_client):
    r = ops_client.post("/ops/makefile", json={}, headers=ops_headers())
    body = r.json()
    assert body["ok"] and len(body["files"]) == 1
    token = body["files"][0]["download_url"].rsplit("/", 1)[1]
    assert ops_client.get(f"/files/{token}").content.decode() == "готово"


def test_file_list_param(ops_client):
    store = ops_client.app.state.filestore
    tokens = [store.save_bytes(b"a", ".txt", "устав.txt"),
              store.save_bytes(b"b", ".txt", "приказ.txt")]
    r = ops_client.post("/ops/pack", json={"params": {"files": tokens}},
                        headers=ops_headers())
    body = r.json()
    assert body["ok"], body
    assert body["files"] and body["files"][0]["name"] == "bundle.zip"
    r = ops_client.post("/ops/pack", json={"params": {"files": ["не-токен"]}},
                        headers=ops_headers())
    assert r.status_code == 422


def test_docv_token_has_no_ops_access(ops_client):
    r = ops_client.post("/ops/echo", json={"params": {"text": "x"}}, headers=docv_headers())
    assert r.status_code == 403


def test_bad_config_fails_fast(tmp_path):
    bad = tmp_path / "ops.yaml"
    bad.write_text('operations:\n  x:\n    argv: ["/bin/sh", "-c", "echo {cmd}"]\n'
                   '    params:\n      cmd: {pattern: ".*"}\n', encoding="utf-8")
    with pytest.raises(OpsConfigError):
        load_registry(bad)


def test_blank_to_png_from_pdf(client):
    """Реальная операция из ops.yaml: PDF-бланк -> фоновый PNG."""
    import pymupdf
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)  # A4
    page.draw_rect(pymupdf.Rect(0, 0, 595, 80), color=None, fill=(0, 0.6, 0.55))
    pdf_bytes = doc.tobytes()
    token = client.app.state.filestore.save_bytes(pdf_bytes, ".pdf", "бланк.pdf")
    r = client.post("/ops/blank_to_png", json={"params": {"file": token}},
                    headers=ops_headers())
    body = r.json()
    assert r.status_code == 200 and body["ok"], body
    assert body["files"] and body["files"][0]["name"] == "blank.png"
    png_token = body["files"][0]["download_url"].rsplit("/", 1)[1]
    png = client.get(f"/files/{png_token}").content
    assert png.startswith(b"\x89PNG")


def _admin(client):
    from conftest import TOKEN_ADMIN
    client.post("/ui/login", data={"token": TOKEN_ADMIN}, follow_redirects=False)


def test_ops_run_redirects_to_detail_and_repeat(ops_client):
    _admin(ops_client)
    r = ops_client.post("/ui/ops/echo", data={"text": "привет"}, follow_redirects=False)
    assert r.status_code == 302 and r.headers["location"] == "/ui/ops/run/1"
    page = ops_client.get("/ui/ops/run/1")
    assert page.status_code == 200 and "привет" in page.text and "Параметры" in page.text
    # обновление страницы запуска не запускает команду заново
    ops_client.get("/ui/ops/run/1")
    assert len(ops_client.app.state.ops_history.recent()) == 1
    r = ops_client.post("/ui/ops/run/1/repeat", follow_redirects=False)
    assert r.headers["location"] == "/ui/ops/run/2"
    runs = ops_client.app.state.ops_history.recent()
    assert len(runs) == 2 and runs[0]["params"] == {"text": "привет"}
    # список запусков и подстановка параметров в форму
    listing = ops_client.get("/ui/ops").text
    assert "Последние запуски" in listing and "последний запуск" in listing
    filled = ops_client.get("/ui/ops?prefill=2").text
    assert 'value="привет"' in filled and "заполнена параметрами запуска №2" in filled
    # запуск через API тоже попадает в историю с пометкой источника
    ops_client.post("/ops/echo", json={"params": {"text": "api"}}, headers=ops_headers())
    assert ops_client.app.state.ops_history.recent()[0]["source"] == "api"
    assert ops_client.get("/ui/ops/run/999", follow_redirects=True).status_code == 200


def test_ops_validation_error_is_not_a_run(ops_client):
    _admin(ops_client)
    r = ops_client.post("/ui/ops/echo", data={"text": "x; rm -rf /"}, follow_redirects=True)
    assert "Параметры не приняты" in r.text
    assert ops_client.app.state.ops_history.recent() == []


def test_ops_history_sweep_limits(ops_client):
    from gateway.jobsqueue.db import connect
    from gateway.opsrunner import history as h
    store = ops_client.app.state.ops_history
    result = {"op": "echo", "ok": True, "exit_code": 0, "duration_ms": 1,
              "stdout": "x" * (h.OUTPUT_LIMIT + 10), "stderr": "", "files": []}
    for _ in range(5):
        store.add(result, {"text": "t"})
    assert len(store.get(1)["stdout"]) == h.OUTPUT_LIMIT
    with connect(ops_client.settings.db_path) as conn:
        conn.execute("UPDATE ops_runs SET started_at = '2000-01-01T00:00:00+00:00' WHERE id = 1")
    old_rows = h.KEEP_ROWS
    h.KEEP_ROWS = 3
    try:
        removed = store.sweep()
    finally:
        h.KEEP_ROWS = old_rows
    assert removed == 2 and [r["id"] for r in store.recent()] == [5, 4, 3]
