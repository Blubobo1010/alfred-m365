"""OpenAI-style tool schemas. Few, plain parameters: local models must call them reliably."""

from __future__ import annotations

_ADDRESSES = {"type": "array", "items": {"type": "string"},
              "description": "Email addresses (a full name is looked up among the CEO's correspondents)"}
_LOCAL = "Local time in the CEO's time zone, e.g. 2026-10-15T10:00"

OUTLOOK_SEARCH_MAIL = {
    "name": "outlook_search_mail",
    "description": ("Search the CEO's Outlook mailbox (newest first). Returns id, date, sender, subject and a preview; "
                    "open one with outlook_read_mail. Use it before drafting a reply."),
    "parameters": {"type": "object", "properties": {
        "query": {"type": "string", "description": "Words to find in subject, body or people (optional)"},
        "sender": {"type": "string", "description": "Sender email address (optional)"},
        "since": {"type": "string", "description": "Only mail received from this date, YYYY-MM-DD (optional)"},
        "folder": {"type": "string", "enum": ["inbox", "sentitems", "drafts", "archive"], "description": "Default: all folders"},
        "unread_only": {"type": "boolean"},
        "limit": {"type": "integer", "minimum": 1, "maximum": 50},
    }},
}

OUTLOOK_READ_MAIL = {
    "name": "outlook_read_mail",
    "description": "Read one email in full (text, recipients, attachment names) by the id from outlook_search_mail.",
    "parameters": {"type": "object", "properties": {"id": {"type": "string"}}, "required": ["id"]},
}

OUTLOOK_LIST_EVENTS = {
    "name": "outlook_list_events",
    "description": "List the CEO's calendar between two dates (default: today and the next 7 days), with ids for outlook_update_event.",
    "parameters": {"type": "object", "properties": {
        "start": {"type": "string", "description": "YYYY-MM-DD or 'today' (default today)"},
        "end": {"type": "string", "description": "YYYY-MM-DD, inclusive (default start + 7 days)"},
    }},
}

OUTLOOK_FIND_FREE_TIME = {
    "name": "outlook_find_free_time",
    "description": "Free slots in the CEO's calendar within his working hours. Call it before proposing or creating an appointment.",
    "parameters": {"type": "object", "properties": {
        "date_from": {"type": "string", "description": "YYYY-MM-DD or 'today'"},
        "date_to": {"type": "string", "description": "YYYY-MM-DD, inclusive (default date_from + 4 days)"},
        "duration_minutes": {"type": "integer", "minimum": 15, "maximum": 480, "description": "Default 60"},
    }, "required": ["date_from"]},
}

OUTLOOK_CREATE_DRAFT = {
    "name": "outlook_create_draft",
    "description": ("Prepare an email DRAFT in the CEO's Outlook (Drafts folder). It is never sent: the CEO reviews and "
                    "sends it himself. With reply_to_id it becomes a reply in that thread (quoted history kept). "
                    "Returns a link to open the draft."),
    "parameters": {"type": "object", "properties": {
        "to": _ADDRESSES, "cc": _ADDRESSES,
        "subject": {"type": "string", "description": "Ignored for replies (Outlook uses RE: <subject>)"},
        "body": {"type": "string", "description": "Plain text, in the CEO's voice, with his sign-off"},
        "reply_to_id": {"type": "string", "description": "Id of the email being answered (from outlook_search_mail)"},
        "reply_all": {"type": "boolean", "description": "Reply to all recipients (only with reply_to_id)"},
    }, "required": ["body"]},
}

OUTLOOK_UPDATE_DRAFT = {
    "name": "outlook_update_draft",
    "description": "Change a draft prepared earlier (only drafts can be changed). The new body replaces the text above the quoted history.",
    "parameters": {"type": "object", "properties": {
        "id": {"type": "string"}, "subject": {"type": "string"}, "body": {"type": "string"}, "to": _ADDRESSES, "cc": _ADDRESSES,
    }, "required": ["id"]},
}

OUTLOOK_CREATE_EVENT = {
    "name": "outlook_create_event",
    "description": ("Create an appointment in the CEO's calendar (he confirms it first). With attendees, Outlook sends them "
                    "an invitation from the CEO: add attendees only when he asked to invite them."),
    "parameters": {"type": "object", "properties": {
        "subject": {"type": "string"},
        "start": {"type": "string", "description": _LOCAL},
        "end": {"type": "string", "description": _LOCAL + " (or use duration_minutes)"},
        "duration_minutes": {"type": "integer", "minimum": 5, "maximum": 1440},
        "location": {"type": "string"},
        "body": {"type": "string", "description": "Notes / agenda (plain text)"},
        "attendees": _ADDRESSES,
        "online_meeting": {"type": "boolean", "description": "Add a Teams link"},
    }, "required": ["subject", "start"]},
}

OUTLOOK_UPDATE_EVENT = {
    "name": "outlook_update_event",
    "description": ("Move or change an appointment the CEO organises (he confirms it first). Attendees receive the update "
                    "from Outlook. Meetings organised by others cannot be changed: draft a reply instead."),
    "parameters": {"type": "object", "properties": {
        "id": {"type": "string", "description": "Event id from outlook_list_events"},
        "subject": {"type": "string"},
        "start": {"type": "string", "description": _LOCAL},
        "end": {"type": "string", "description": _LOCAL},
        "location": {"type": "string"},
        "body": {"type": "string"},
        "add_attendees": _ADDRESSES,
    }, "required": ["id"]},
}
