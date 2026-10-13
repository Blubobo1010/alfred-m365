"""Credentials and settings. Secrets come from Hermes' profile-scoped secret store only."""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any

from .graph import AUTHORITY_URL, GRAPH_BASE_URL, GraphClient, GraphError

REQUIRED_ENV = ("MSGRAPH_TENANT_ID", "MSGRAPH_CLIENT_ID", "MSGRAPH_CLIENT_SECRET")

_ctx: Any = None  # the PluginContext handed to register(); settings are read live from it


def bind_context(ctx: Any) -> None:
    global _ctx
    _ctx = ctx


def get_secret(name: str) -> str:
    """Profile-aware secret lookup (``agent.secret_scope``); plain environment outside Hermes (dev/tests)."""
    try:
        from agent.secret_scope import get_secret as _hermes_get_secret
    except ImportError:  # running outside a Hermes runtime (unit tests, dev/live_check.py)
        return os.environ.get(name, "") or ""
    return _hermes_get_secret(name) or ""


def setting(key: str, default: Any) -> Any:
    if _ctx is None:
        return default
    try:
        value = _ctx.get_config(key, default)
    except Exception:  # noqa: BLE001 - a bad config path must not break a tool call
        return default
    return default if value is None else value


@dataclass(frozen=True)
class Settings:
    mailbox: str = ""
    owner_name: str = ""
    timezone: str = "Europe/Rome"
    write_enabled: bool = False
    approve_drafts: bool = False
    default_limit: int = 15
    max_body_chars: int = 6000
    timeout_seconds: int = 30
    graph_base_url: str = GRAPH_BASE_URL  # dev/tests only (a local fake Graph)

    @property
    def authority_url(self) -> str:
        """A local fake Graph (dev) also serves the token endpoint; otherwise Microsoft's login."""
        if self.graph_base_url == GRAPH_BASE_URL:
            return AUTHORITY_URL
        from urllib.parse import urlsplit
        parts = urlsplit(self.graph_base_url)
        return f"{parts.scheme}://{parts.netloc}/login"


def load_settings() -> Settings:
    return Settings(
        mailbox=str(setting("mailbox", "") or "").strip().lower(),
        owner_name=str(setting("owner_name", "") or "").strip(),
        timezone=str(setting("timezone", "Europe/Rome") or "Europe/Rome"),
        write_enabled=bool(setting("write_enabled", False)),
        approve_drafts=bool(setting("approve_drafts", False)),
        default_limit=int(setting("default_limit", 15) or 15),
        max_body_chars=int(setting("max_body_chars", 6000) or 6000),
        timeout_seconds=int(setting("timeout_seconds", 30) or 30),
        graph_base_url=_local_override(str(setting("graph_base_url", "") or os.environ.get("ALFRED_M365_GRAPH_URL", ""))),
    )


def _local_override(url: str) -> str:
    """A fake Graph for dev/tests may only live on this machine: the bearer token is sent there."""
    from urllib.parse import urlsplit
    parts = urlsplit(url.strip())
    if parts.scheme in ("http", "https") and parts.hostname in ("127.0.0.1", "localhost", "::1"):
        return url.rstrip("/")
    return GRAPH_BASE_URL


def is_configured() -> bool:
    """check_fn for every tool: credentials and mailbox present; never touches the network."""
    try:
        return all(get_secret(n) for n in REQUIRED_ENV) and bool(load_settings().mailbox)
    except Exception:  # noqa: BLE001 - UnscopedSecretError & co.: treat as not configured
        return False


def write_enabled() -> bool:
    return is_configured() and load_settings().write_enabled


_clients: dict[tuple, GraphClient] = {}


def make_client(settings: Settings | None = None) -> GraphClient:
    """One client per credentials+mailbox, so the access token is reused across tool calls."""
    settings = settings or load_settings()
    values = {n: get_secret(n) for n in REQUIRED_ENV}
    missing = [n for n in REQUIRED_ENV if not values[n]] + ([] if settings.mailbox else ["settings.mailbox"])
    if missing:
        raise GraphError("not_configured", "Microsoft 365 is not configured: missing " + ", ".join(missing),
                         "IT enters the three MSGRAPH_* values on the Mac mini; set the mailbox with "
                         "`alfred config set plugins.entries.alfred-m365.settings.mailbox <address>`.")
    key = (values["MSGRAPH_TENANT_ID"], values["MSGRAPH_CLIENT_ID"], hash(values["MSGRAPH_CLIENT_SECRET"]),
           settings.mailbox, settings.timezone, settings.graph_base_url, settings.timeout_seconds)
    if key not in _clients:
        _clients.clear()
        _clients[key] = GraphClient(values["MSGRAPH_TENANT_ID"], values["MSGRAPH_CLIENT_ID"], values["MSGRAPH_CLIENT_SECRET"],
                                    settings.mailbox, timezone=settings.timezone, timeout=settings.timeout_seconds,
                                    base_url=settings.graph_base_url, authority_url=settings.authority_url)
    return _clients[key]
