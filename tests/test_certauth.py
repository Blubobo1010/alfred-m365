import json
import stat

import pytest
from alfred_m365 import certauth, settings, tools
from alfred_m365.graph import GraphError
from conftest import call


def test_keygen_writes_private_key_0600_and_refuses_overwrite(tmp_path):
    info = certauth.generate(tmp_path / "certs", "Alfred")
    key = tmp_path / "certs" / certauth.KEY_FILE
    assert stat.S_IMODE(key.stat().st_mode) == 0o600 and "PRIVATE KEY" in key.read_text()
    assert "PRIVATE" not in (tmp_path / "certs" / certauth.CERT_FILE).read_text()
    assert len(info["thumbprint"]) == 40
    with pytest.raises(GraphError):
        certauth.generate(tmp_path / "certs", "Alfred")


def test_certificate_login_without_any_secret(ctx, fake, monkeypatch, tmp_path):
    monkeypatch.delenv("MSGRAPH_CLIENT_SECRET")
    assert settings.is_configured() is False  # no secret, no certificate yet
    certauth.generate(tmp_path / "c", "Alfred")
    ctx.config["certificate_dir"] = str(tmp_path / "c")
    fake.cert_path = tmp_path / "c" / certauth.CERT_FILE
    assert settings.is_configured() is True
    out = call(tools.outlook_search_mail, {"limit": 1})
    assert out["ok"] and fake.token_calls == 1


def test_wrong_certificate_is_rejected(ctx, fake, monkeypatch, tmp_path):
    monkeypatch.delenv("MSGRAPH_CLIENT_SECRET")
    certauth.generate(tmp_path / "mine", "Alfred")
    certauth.generate(tmp_path / "other", "Alfred")
    ctx.config["certificate_dir"] = str(tmp_path / "mine")
    fake.cert_path = tmp_path / "other" / certauth.CERT_FILE  # IT uploaded a different certificate
    out = json.loads(tools.outlook_search_mail({}))
    assert out["error"]["code"] == "auth_failed" and "certificate" in out["error"]["hint"]
