"""alfred-m365 — the CEO's Outlook mail and calendar for Hermes / Alfred.

Reads mail and calendar; prepares drafts (never sends) and creates or edits appointments, each
calendar change confirmed by a human. Registration is synchronous and I/O-free; availability is
decided per call by ``check_fn`` (credentials + mailbox present, writes enabled).
"""

from __future__ import annotations

from pathlib import Path

from . import schemas, tools
from .approvals import pre_tool_call
from .cli import outlook_command, register_cli
from .settings import bind_context, is_configured, write_enabled

__version__ = "0.1.0"
TOOLSET = "outlook"
REQUIRED_ENV = ["MSGRAPH_TENANT_ID", "MSGRAPH_CLIENT_ID"]  # + certificate (keygen) or MSGRAPH_CLIENT_SECRET

_TOOLS = (
    ("outlook_search_mail", schemas.OUTLOOK_SEARCH_MAIL, tools.outlook_search_mail, "📬", is_configured),
    ("outlook_read_mail", schemas.OUTLOOK_READ_MAIL, tools.outlook_read_mail, "📬", is_configured),
    ("outlook_list_events", schemas.OUTLOOK_LIST_EVENTS, tools.outlook_list_events, "📅", is_configured),
    ("outlook_find_free_time", schemas.OUTLOOK_FIND_FREE_TIME, tools.outlook_find_free_time, "📅", is_configured),
    ("outlook_create_draft", schemas.OUTLOOK_CREATE_DRAFT, tools.outlook_create_draft, "✍️", write_enabled),
    ("outlook_update_draft", schemas.OUTLOOK_UPDATE_DRAFT, tools.outlook_update_draft, "✍️", write_enabled),
    ("outlook_create_event", schemas.OUTLOOK_CREATE_EVENT, tools.outlook_create_event, "🗓️", write_enabled),
    ("outlook_update_event", schemas.OUTLOOK_UPDATE_EVENT, tools.outlook_update_event, "🗓️", write_enabled),
)


def register(ctx) -> None:
    bind_context(ctx)
    for name, schema, handler, emoji, check in _TOOLS:
        ctx.register_tool(name=name, toolset=TOOLSET, schema=schema, handler=handler, check_fn=check,
                          requires_env=REQUIRED_ENV, emoji=emoji)
    # Calendar writes (and drafts with approve_drafts) go through the human approval gate.
    ctx.register_hook("pre_tool_call", pre_tool_call)
    ctx.register_skill("outlook", Path(__file__).parent / "skills" / "outlook" / "SKILL.md")
    ctx.register_cli_command(
        name="outlook", help="Microsoft 365 connection: keygen, test, whoami",
        setup_fn=register_cli, handler_fn=outlook_command,
        description="Check the Outlook mail and calendar connection used by the outlook_* tools.",
    )
