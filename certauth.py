"""Certificate credential for the Entra app: the private key is generated on this machine and never
leaves it; IT uploads only the public certificate (Entra ID > App registrations > Certificates & secrets).

The token request then carries a short-lived signed JWT (client assertion, RFC 7523) instead of a
client secret, so no secret ever travels by mail, chat or phone. Uses ``cryptography``, which ships
with the Alfred runtime; imported lazily so the secret path stays stdlib-only.
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import time
import uuid
from datetime import UTC, datetime, timedelta
from pathlib import Path

from .graph import GraphError

KEY_FILE = "alfred-m365-key.pem"
CERT_FILE = "alfred-m365.cer"  # public, DER-free PEM: this is the file IT uploads
VALID_DAYS = 730


def default_dir() -> Path:
    home = os.environ.get("HERMES_HOME") or str(Path.home() / ".hermes")
    return Path(home) / "alfred-m365"


def _b64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def generate(directory: Path, common_name: str, *, overwrite: bool = False) -> dict[str, str]:
    """RSA 2048 key (0600) + self-signed certificate valid two years. Refuses to overwrite an existing key."""
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import NameOID

    directory.mkdir(parents=True, exist_ok=True)
    os.chmod(directory, 0o700)
    key_path, cert_path = directory / KEY_FILE, directory / CERT_FILE
    if key_path.exists() and not overwrite:
        raise GraphError("exists", f"a key already exists at {key_path}",
                         "Keep it (IT has its certificate) or pass --overwrite and send IT the new certificate.")
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, common_name)])
    now = datetime.now(UTC)
    cert = (x509.CertificateBuilder().subject_name(name).issuer_name(name).public_key(key.public_key())
            .serial_number(x509.random_serial_number()).not_valid_before(now - timedelta(minutes=5))
            .not_valid_after(now + timedelta(days=VALID_DAYS)).sign(key, hashes.SHA256()))
    pem_key = key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption())
    fd = os.open(key_path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "wb") as fh:
        fh.write(pem_key)
    cert_path.write_bytes(cert.public_bytes(serialization.Encoding.PEM))
    return {"key": str(key_path), "certificate": str(cert_path), "thumbprint": thumbprint(cert_path),
            "expires": (now + timedelta(days=VALID_DAYS)).date().isoformat()}


def _load(directory: Path):
    from cryptography import x509
    from cryptography.hazmat.primitives import serialization
    try:
        key = serialization.load_pem_private_key((directory / KEY_FILE).read_bytes(), password=None)
        cert = x509.load_pem_x509_certificate((directory / CERT_FILE).read_bytes())
    except FileNotFoundError:
        raise GraphError("not_configured", f"no certificate in {directory}",
                         "Run `alfred outlook keygen` and send the .cer file to IT.") from None
    return key, cert


def thumbprint(cert_path: Path) -> str:
    """SHA-1 thumbprint as Entra shows it (uppercase hex), to match the uploaded certificate."""
    from cryptography import x509
    from cryptography.hazmat.primitives import serialization
    cert = x509.load_pem_x509_certificate(cert_path.read_bytes())
    return hashlib.sha1(cert.public_bytes(serialization.Encoding.DER)).hexdigest().upper()  # noqa: S324 - Entra's x5t


def available(directory: Path) -> bool:
    return (directory / KEY_FILE).is_file() and (directory / CERT_FILE).is_file()


def client_assertion(directory: Path, client_id: str, token_url: str) -> str:
    """RS256 JWT for the token endpoint: aud = token URL, iss = sub = client id, valid 10 minutes."""
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import padding

    key, cert = _load(directory)
    der = cert.public_bytes(serialization.Encoding.DER)
    header = {"alg": "RS256", "typ": "JWT", "x5t": _b64url(hashlib.sha1(der).digest()),  # noqa: S324
              "x5t#S256": _b64url(hashlib.sha256(der).digest())}
    now = int(time.time())
    claims = {"aud": token_url, "iss": client_id, "sub": client_id, "jti": str(uuid.uuid4()),
              "nbf": now - 60, "iat": now, "exp": now + 600}
    signing_input = f"{_b64url(json.dumps(header, separators=(',', ':')).encode())}.{_b64url(json.dumps(claims, separators=(',', ':')).encode())}"
    signature = key.sign(signing_input.encode("ascii"), padding.PKCS1v15(), hashes.SHA256())
    return f"{signing_input}.{_b64url(signature)}"
