from datetime import datetime, timedelta, timezone

from gateway.jobsqueue.db import connect


def test_roundtrip(client):
    store = client.app.state.filestore
    token = store.save_bytes(b"hello", ".txt", "отчет.txt")
    r = client.get(f"/files/{token}")
    assert r.status_code == 200
    assert r.content == b"hello"
    assert "otchet" in r.headers["content-disposition"] or "%D0" in r.headers["content-disposition"]


def test_unknown_token_404(client):
    assert client.get("/files/" + "a" * 24).status_code == 404
    assert client.get("/files/short").status_code == 404


def test_sweep_removes_expired(client):
    store = client.app.state.filestore
    token = store.save_bytes(b"old", ".txt", "old.txt")
    old = (datetime.now(timezone.utc) - timedelta(hours=100)).isoformat(timespec="seconds")
    with connect(client.settings.db_path) as conn:
        conn.execute("UPDATE files SET created_at = ? WHERE token = ?", (old, token))
    assert store.sweep() >= 1
    assert client.get(f"/files/{token}").status_code == 404


def test_sweep_keeps_pinned(client):
    """Закреплённый файл переживает срок хранения, обычный — нет."""
    store = client.app.state.filestore
    keep = store.save_bytes(b"keep", ".txt", "blank.txt")
    drop = store.save_bytes(b"drop", ".txt", "tmp.txt")
    old = (datetime.now(timezone.utc) - timedelta(hours=100)).isoformat(timespec="seconds")
    with connect(client.settings.db_path) as conn:
        conn.execute("UPDATE files SET created_at = ?", (old,))
    assert store.set_pinned([keep], True) == 1
    store.sweep()
    assert client.get(f"/files/{keep}").status_code == 200
    assert client.get(f"/files/{drop}").status_code == 404
    # старая база без колонки pinned получает её при init_db
    from gateway.jobsqueue.db import init_db
    init_db(client.settings.db_path)
    with connect(client.settings.db_path) as conn:
        assert "pinned" in {r["name"] for r in conn.execute("PRAGMA table_info(files)")}


def test_debug_echo(client):
    from conftest import docv_headers
    r = client.post("/debug/echo", headers=docv_headers(),
                    json={"Регистрационный номер": "12-СМР", "Сумма договора": 1000000})
    body = r.json()
    assert r.status_code == 200
    assert "Регистрационный номер" in body["received"]
    token = body["saved_as"].rsplit("/", 1)[1]
    assert "12-СМР" in client.get(f"/files/{token}").text
    # без токена — 403
    assert client.post("/debug/echo", json={}).status_code == 403
