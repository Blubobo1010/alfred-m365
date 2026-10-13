"""Human approval for the write tools, through Hermes' ``pre_tool_call`` hook (alfred-odoo pattern).

Event writes always need the CEO's OK: they change his calendar and, with attendees, make Outlook
send invitations or updates in his name. Drafts are never sent, so they pass without a prompt
unless ``approve_drafts`` is on. The prompt shows only ``message``: it must describe the change on
its own, in Italian, from the raw arguments (no network here). ``rule_key`` hashes tool + arguments,
so "always" covers only the identical change. Any internal failure blocks the call.
"""

from __future__ import annotations

import hashlib
import json
from typing import Any

from .settings import load_settings
from .timeutil import human_range, parse_local, zone

EVENT_TOOLS = ("outlook_create_event", "outlook_update_event")
DRAFT_TOOLS = ("outlook_create_draft", "outlook_update_draft")
WRITE_TOOLS = EVENT_TOOLS + DRAFT_TOOLS
SUMMARY_MAX_CHARS = 400


def _short(text: Any, limit: int) -> str:
    text = " ".join(str(text or "").split())
    return text if len(text) <= limit else text[: max(0, limit - 1)].rstrip() + "…"


def _people(value: Any) -> str:
    if isinstance(value, str):
        value = [value]
    return ", ".join(str(v).strip() for v in value or [] if str(v).strip())


def _when(args: dict[str, Any], tz_name: str) -> str:
    tz = zone(tz_name)
    try:
        from .tools import event_times
        start, end = event_times(args, tz)
        return human_range(start, end)
    except Exception:  # noqa: BLE001 - an unparsable time is shown raw; the tool will reject it
        return _short(f"{args.get('start') or ''} {args.get('end') or ''}", 60)


def approval_summary(tool_name: str, args: dict[str, Any], tz_name: str, owner: str = "il CEO") -> str:
    if tool_name == "outlook_create_event":
        text = f'Outlook: crea "{_short(args.get("subject"), 80)}" {_when(args, tz_name)} nel calendario di {owner}'
        if args.get("location"):
            text += f", luogo: {_short(args['location'], 60)}"
        if args.get("online_meeting"):
            text += ", con link Teams"
        if args.get("attendees"):
            text += f". ⚠ Invia un invito a nome di {owner} a: {_people(args['attendees'])}"
    elif tool_name == "outlook_update_event":
        parts = []
        if args.get("start"):
            tz = zone(tz_name)
            try:
                start = parse_local(args["start"], tz, field="start")
                end = parse_local(args["end"], tz, field="end") if args.get("end") else None
                parts.append("sposta a " + (human_range(start, end) if end else human_range(start, start).split("–")[0]))
            except Exception:  # noqa: BLE001
                parts.append(f"sposta a {_short(args['start'], 30)}")
        elif args.get("end"):
            parts.append(f"nuova fine {_short(args['end'], 30)}")
        for key, label in (("subject", "titolo"), ("location", "luogo"), ("body", "note")):
            if args.get(key):
                parts.append(f'{label} "{_short(args[key], 60)}"')
        if args.get("add_attendees"):
            parts.append(f"⚠ invita a nome di {owner}: {_people(args['add_attendees'])}")
        text = (f"Outlook: modifica l'appuntamento {_short(args.get('id'), 12)} ({'; '.join(parts) or 'nessuna modifica'}). "
                "Se ha partecipanti, ricevono l'aggiornamento da Outlook")
    elif tool_name == "outlook_create_draft":
        target = "risposta" if args.get("reply_to_id") else f'nuova email "{_short(args.get("subject"), 60)}"'
        text = (f"Outlook: prepara una BOZZA ({target}) per {_people(args.get('to')) or 'i destinatari del messaggio'}: "
                f'"{_short(args.get("body"), 160)}". Non viene inviata')
    elif tool_name == "outlook_update_draft":
        text = f'Outlook: modifica la bozza {_short(args.get("id"), 12)}: "{_short(args.get("body") or args.get("subject"), 160)}"'
    else:
        text = f"Outlook: {tool_name} {_short(json.dumps(args, ensure_ascii=False, sort_keys=True, default=str), 200)}"
    return _short(text, SUMMARY_MAX_CHARS)


def rule_key(tool_name: str, args: dict[str, Any]) -> str:
    canonical = json.dumps(args or {}, sort_keys=True, ensure_ascii=False, default=str, separators=(",", ":"))
    return f"alfred-m365:{tool_name}:{hashlib.sha256(canonical.encode('utf-8')).hexdigest()[:16]}"


def pre_tool_call(tool_name: str = "", args: Any = None, **_: Any) -> dict[str, str] | None:
    """Hermes ``pre_tool_call`` hook: None for everything that is not one of our write tools."""
    if tool_name not in WRITE_TOOLS:
        return None
    try:
        call_args = dict(args) if isinstance(args, dict) else {}
        settings = load_settings()
        if not settings.write_enabled:
            return {"action": "block", "message": "BLOCKED: Outlook writes are disabled "
                    "(set write_enabled: true in the alfred-m365 plugin settings)."}
        if tool_name in DRAFT_TOOLS and not settings.approve_drafts:
            return None  # a draft is never sent: the CEO reviews it in Outlook
        owner = settings.owner_name or settings.mailbox
        return {"action": "approve", "message": approval_summary(tool_name, call_args, settings.timezone, owner),
                "rule_key": rule_key(tool_name, call_args)}
    except Exception as exc:  # noqa: BLE001 - never let a write slip through on a hook failure
        return {"action": "block", "message": f"BLOCKED: alfred-m365 approval hook failed ({exc.__class__.__name__})."}
