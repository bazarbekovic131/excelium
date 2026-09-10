import sys
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from gateway.config import Settings  # noqa: E402
from gateway.main import create_app  # noqa: E402

TOKEN_DOCV = "test-docv-token"
TOKEN_OPS = "test-ops-token"
TOKEN_ADMIN = "test-admin-token"
PRODUCER_TOKENS = {"onec": "test-producer-token"}


@pytest.fixture()
def settings(tmp_path):
    return Settings(
        allowlist=["testclient", "127.0.0.1"],
        token_docv=TOKEN_DOCV,
        token_ops=TOKEN_OPS,
        token_admin=TOKEN_ADMIN,
        verify_secret="test-verify-secret",
        producer_tokens=PRODUCER_TOKENS,
        base_url="http://testserver",
        var_dir=tmp_path / "var",
        _env_file=None,
    )


SEED = Path(__file__).parent / "data" / "signers_seed.yaml"


@pytest.fixture()
def client(settings):
    app = create_app(settings)
    with TestClient(app) as c:
        c.settings = settings
        # боевой состав подписантов живёт в базе на сервере; тесты
        # получают такой же объём из своей копии старого справочника
        c.app.state.signers.seed(SEED, c.app.state.template_inner)
        yield c


def docv_headers():
    return {"Authorization": f"Bearer {TOKEN_DOCV}"}


def ops_headers():
    return {"Authorization": f"Bearer {TOKEN_OPS}"}


def producer_headers():
    return {"Authorization": f"Bearer {PRODUCER_TOKENS['onec']}"}
