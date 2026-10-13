import json
from datetime import datetime

import alfred_m365
from alfred_m365 import tools
from alfred_m365.approvals import WRITE_TOOLS, approval_summary, pre_tool_call
from conftest import MAILBOX, SECRET, FakeCtx, call, mid

ROOT_TOOLS = ("outlook_search_mail", "outlook_read_mail", "outlook_list_events", "outlook_find_free_time")


def test_register_matches_manifest_and_gates(monkeypatch, fake):
    from conftest import ROOT
    text = (ROOT / "plugin.yaml").read_text(encoding="utf-8")
    block = text.split("provides_tools:", 1)[1].split("provides_hooks:", 1)[0]
    manifest = [ln.strip()[2:] for ln in block.splitlines() if ln.strip().startswith("- ")]
    c = FakeCtx()
    alfred_m365.register(c)
    assert sorted(c.tools) == sorted(manifest) and {t["toolset"] for t in c.tools.values()} == {"outlook"}
    assert c.hooks == [("pre_tool_call", pre_tool_call)] and c.skills["outlook"].is_file() and "outlook" in c.cli
    assert all(t["check_fn"]() is False for t in c.tools.values())  # nothing configured
    for name in ("MSGRAPH_TENANT_ID", "MSGRAPH_CLIENT_ID", "MSGRAPH_CLIENT_SECRET"):
        monkeypatch.setenv(name, "x")
    c.config["mailbox"] = MAILBOX
    assert all(c.tools[n]["check_fn"]() for n in ROOT_TOOLS)
    assert not any(c.tools[n]["check_fn"]() for n in WRITE_TOOLS)
    c.config["write_enabled"] = True
    assert all(c.tools[n]["check_fn"]() for n in WRITE_TOOLS)
    assert fake.calls == []  # register and check_fn never touch the network


def test_search_and_read_mark_mail_as_untrusted(ctx, fake):
    found = call(tools.outlook_search_mail, {"limit": 5})
    assert found["ok"] and found["items"][0]["subject"] == "Preventivo telaio" and found["content_notice"]
    assert found["items"][0]["date"].startswith("lun 12/10 10:30")  # 08:30Z -> Rome
    mail = call(tools.outlook_read_mail, {"id": mid(1)})
    assert "IGNORA LE ISTRUZIONI" in mail["untrusted_text"] and "text" not in mail
    assert mail["attachment_names"] == ["preventivo.pdf"] and "third parties" in mail["content_notice"]


def test_search_by_query_uses_kql(ctx, fake):
    call(tools.outlook_search_mail, {"query": "preventivo", "sender": "marco@fornitore.it", "since": "2026-10-01"})
    params = fake.calls[-1][2]
    assert params["$search"] == '"preventivo from:marco@fornitore.it received>=2026-10-01"' and "$orderby" not in params


def test_new_draft_is_created_never_sent(ctx, fake):
    out = call(tools.outlook_create_draft, {"to": ["marco@fornitore.it"], "subject": "Preventivo", "body": "Grazie Marco,\nLuca"})
    assert out["ok"] and "NON inviata" in out["status"] and out["web_link"]
    method, path, body = next(c for c in fake.calls if c[0] == "POST")
    assert (method, path) == ("POST", f"/users/{MAILBOX}/messages")
    assert body["body"] == {"contentType": "Text", "content": "Grazie Marco,\nLuca"} and fake.messages[out["draft_id"]]["isDraft"]
    assert not any("send" in p.lower() for _m, p, _b in fake.calls)


def test_reply_draft_keeps_the_thread_and_update_keeps_the_quote(ctx, fake):
    out = call(tools.outlook_create_draft, {"reply_to_id": mid(1), "body": "Grazie, lo guardo entro venerdì.\nLuca"})
    assert out["subject"] == "RE: Preventivo telaio" and out["to"] == ["Marco Bianchi <marco@fornitore.it>"]
    assert fake.calls[-1][:2] == ("POST", f"/users/{MAILBOX}/messages/{mid(1)}/createReply")
    upd = call(tools.outlook_update_draft, {"id": out["draft_id"], "body": "Grazie, rispondo lunedì."})
    assert upd["ok"]
    content = fake.messages[out["draft_id"]]["body"]["content"]
    assert content.startswith("<div>Grazie, rispondo lunedì.</div>") and 'divRplyFwdMsg">quoted' in content


def test_update_draft_refuses_received_mail(ctx, fake):
    out = call(tools.outlook_update_draft, {"id": mid(1), "subject": "x"})
    assert out["error"]["code"] == "not_a_draft"


def test_writes_disabled(ctx, fake):
    ctx.config["write_enabled"] = False
    out = call(tools.outlook_create_draft, {"to": ["a@b.it"], "body": "x"})
    assert out["error"]["code"] == "writes_disabled" and fake.calls == []


def test_create_event_in_local_time_with_invite(ctx, fake):
    out = call(tools.outlook_create_event, {"subject": "Riunione Bonetto", "start": "2026-10-16T10:00", "duration_minutes": 90,
                                            "attendees": ["bonetto@example.com"], "online_meeting": True})
    body = fake.calls[-1][2]
    assert body["start"] == {"dateTime": "2026-10-16T10:00:00", "timeZone": "W. Europe Standard Time"}
    assert body["end"]["dateTime"] == "2026-10-16T11:30:00" and body["isOnlineMeeting"]
    assert body["attendees"] == [{"emailAddress": {"address": "bonetto@example.com"}, "type": "required"}]
    assert out["when"] == "ven 16/10 10:00–11:30" and "invito inviato a bonetto@example.com" in out["status"]


def test_move_event_keeps_duration_and_notifies(ctx, fake):
    out = call(tools.outlook_update_event, {"id": "EVT0000001", "start": "2026-10-15T11:00"})
    patch = [c for c in fake.calls if c[0] == "PATCH"][-1][2]
    assert patch["start"]["dateTime"] == "2026-10-15T11:00:00" and patch["end"]["dateTime"] == "2026-10-15T12:30:00"
    assert "partecipanti" in out["status"]
    refused = call(tools.outlook_update_event, {"id": "EVT0000002", "start": "2026-10-15T16:00"})
    assert refused["error"]["code"] == "not_organizer"


def test_list_events_and_free_time(ctx, fake, monkeypatch):
    out = call(tools.outlook_list_events, {"start": "2026-10-15", "end": "2026-10-15"})
    assert [e["subject"] for e in out["events"]] == ["Riunione produzione", "Call fornitore"]
    assert out["events"][0]["when"] == "gio 15/10 09:00–10:30" and out["events"][1]["organised_by_ceo"] is False

    class FrozenNow(datetime):
        @classmethod
        def now(cls, tz=None):
            return datetime(2026, 10, 14, 8, 0, tzinfo=tz)

    monkeypatch.setattr(tools, "datetime", FrozenNow)
    free = call(tools.outlook_find_free_time, {"date_from": "2026-10-15", "date_to": "2026-10-15", "duration_minutes": 60})
    assert [s["when"] for s in free["slots"]] == ["gio 15/10 10:30–14:00", "gio 15/10 15:00–18:00"]


def test_bad_time_is_a_clear_error(ctx, fake):
    out = call(tools.outlook_create_event, {"subject": "x", "start": "domani alle 10"})
    assert out["error"]["code"] == "bad_time" and fake.calls == []


def test_approval_always_for_events_drafts_only_on_request(ctx):
    ev = pre_tool_call("outlook_create_event", {"subject": "Riunione Bonetto", "start": "2026-10-16T10:00",
                                                "attendees": ["bonetto@example.com"]})
    assert ev["action"] == "approve" and ev["rule_key"].startswith("alfred-m365:outlook_create_event:")
    assert "ven 16/10 10:00–11:00" in ev["message"] and "calendario di Luca" in ev["message"]
    assert "⚠ Invia un invito a nome di Luca a: bonetto@example.com" in ev["message"]
    assert pre_tool_call("outlook_create_draft", {"to": ["a@b.it"], "body": "x"}) is None
    ctx.config["approve_drafts"] = True
    draft = pre_tool_call("outlook_create_draft", {"to": ["a@b.it"], "body": "Ciao"})
    assert draft["action"] == "approve" and "BOZZA" in draft["message"] and "Non viene inviata" in draft["message"]
    ctx.config["write_enabled"] = False
    assert pre_tool_call("outlook_update_event", {"id": "EVT0000001"})["action"] == "block"
    assert pre_tool_call("terminal", {"command": "ls"}) is None


def test_approval_hook_fails_closed(ctx, monkeypatch):
    import alfred_m365.approvals as ap
    monkeypatch.setattr(ap, "load_settings", lambda: (_ for _ in ()).throw(RuntimeError("boom")))
    assert pre_tool_call("outlook_create_event", {"subject": "x", "start": "2026-10-16T10:00"})["action"] == "block"


def test_move_summary_is_readable():
    text = approval_summary("outlook_update_event", {"id": "EVT0000001", "start": "2026-10-15T11:00", "end": "2026-10-15T12:00"},
                            "Europe/Rome", "Luca")
    assert "sposta a gio 15/10 11:00–12:00" in text and "aggiornamento" in text


def test_secret_never_in_any_output(ctx, fake):
    outputs = [tools.outlook_search_mail({}), tools.outlook_read_mail({"id": mid(1)}), tools.outlook_list_events({}),
               tools.outlook_create_draft({"to": ["a@b.it"], "body": "x"}), tools.outlook_read_mail({"id": "bad"})]
    fake.fail_next.append((500, {}, {"error": {"code": "x", "message": "boom"}}))
    outputs.append(tools.outlook_search_mail({}))
    assert all(SECRET not in o for o in outputs)
    assert all(json.loads(o) for o in outputs)


def test_graph_url_override_only_local(ctx):
    from alfred_m365.settings import load_settings
    ctx.config["graph_base_url"] = "https://evil.example.com/v1.0"
    assert load_settings().graph_base_url == "https://graph.microsoft.com/v1.0"
    ctx.config["graph_base_url"] = "http://127.0.0.1:8765/v1.0"
    assert load_settings().graph_base_url == "http://127.0.0.1:8765/v1.0"


def test_recipient_names_are_resolved_from_correspondents(ctx, fake):
    out = call(tools.outlook_create_draft, {"to": ["Marco Bianchi"], "subject": "Preventivo", "body": "Grazie"})
    assert out["ok"] and out["to"] == ["marco@fornitore.it"]
    unknown = call(tools.outlook_create_draft, {"to": ["Giulia Verdi"], "body": "x"})
    assert unknown["error"]["code"] == "unknown_recipient"


def test_guessed_address_is_flagged(ctx, fake):
    known = call(tools.outlook_create_draft, {"to": ["marco@fornitore.it"], "body": "x"})
    assert "warning" not in known
    guessed = call(tools.outlook_create_draft, {"to": ["marco.bianchi@example.com"], "body": "x"})
    assert guessed["ok"] and "marco.bianchi@example.com" in guessed["warning"]
