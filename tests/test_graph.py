"""The allow-list is the "never send, move or delete" guarantee: refused calls never reach the network."""

import pytest
from alfred_m365.graph import GraphClient, GraphError, NotAllowed
from conftest import MAILBOX, SECRET, mid


def client(fake):
    return GraphClient("tenant-1", "client-1", SECRET, MAILBOX, transport=fake, sleep=lambda s: None)


@pytest.mark.parametrize("method,path", [
    ("DELETE", f"/messages/{mid(1)}"),
    ("POST", f"/messages/{mid(1)}/send"),
    ("POST", "/sendMail"),
    ("POST", f"/messages/{mid(1)}/move"),
    ("POST", f"/messages/{mid(1)}/copy"),
    ("POST", f"/messages/{mid(1)}/forward"),
    ("POST", f"/messages/{mid(1)}/reply"),  # reply (sends) vs createReply (draft)
    ("DELETE", "/events/EVT0000001"),
    ("POST", "/events/EVT0000001/cancel"),
    ("POST", "/events/EVT0000001/decline"),
    ("PATCH", "/mailboxSettings"),
    ("GET", "/../other@example.com/messages"),
    ("GET", "/mailFolders/inbox/messages/../../../x"),
])
def test_forbidden_operations_never_reach_graph(fake, method, path):
    c = client(fake)
    with pytest.raises(NotAllowed):
        c._call(method, path)
    assert fake.calls == [] and fake.token_calls == 0


def test_mailbox_is_fixed_and_validated(fake):
    with pytest.raises(GraphError):
        GraphClient("t", "c", SECRET, "not-an-address", transport=fake)
    c = client(fake)
    c.search_messages(limit=1)
    assert fake.calls[0][1] == f"/users/{MAILBOX}/messages"


def test_ids_are_validated(fake):
    with pytest.raises(GraphError) as exc:
        client(fake).get_message("../../users/other@example.com/messages")
    assert exc.value.code == "bad_id" and fake.calls == []


def test_patch_only_on_drafts(fake):
    with pytest.raises(GraphError) as exc:
        client(fake).update_draft(mid(1), {"subject": "hacked"})
    assert exc.value.code == "not_a_draft"
    assert [c[0] for c in fake.calls] == ["GET"] and fake.messages[mid(1)]["subject"] == "Preventivo telaio"


def test_event_patch_refuses_others_meetings_and_series(fake):
    c = client(fake)
    with pytest.raises(GraphError) as exc:
        c.update_event("EVT0000002", {"subject": "x"})
    assert exc.value.code == "not_organizer"
    with pytest.raises(GraphError) as exc:
        c.update_event("EVTSERIES01", {"subject": "x"})
    assert exc.value.code == "series"
    assert all(m == "GET" for m, _p, _d in fake.calls)


def test_token_is_cached_and_refreshed_on_401(fake):
    c = client(fake)
    c.search_messages(limit=1)
    c.search_messages(limit=1)
    assert fake.token_calls == 1
    fake.fail_next.append((401, {}, {"error": {"code": "InvalidAuthenticationToken", "message": "expired"}}))
    c.search_messages(limit=1)
    assert fake.token_calls == 2


def test_throttling_honours_retry_after(fake):
    slept = []
    c = GraphClient("tenant-1", "client-1", SECRET, MAILBOX, transport=fake, sleep=slept.append)
    fake.fail_next.append((429, {"Retry-After": "7"}, {"error": {"code": "TooManyRequests", "message": "slow down"}}))
    assert c.search_messages(limit=1)
    assert slept == [7.0]


def test_403_explains_rbac(fake):
    fake.fail_next.append((403, {}, {"error": {"code": "ErrorAccessDenied", "message": "Access is denied."}}))
    with pytest.raises(GraphError) as exc:
        client(fake).search_messages(limit=1)
    assert exc.value.code == "forbidden" and "RBAC" in exc.value.hint


def test_wrong_secret_is_never_echoed(fake):
    c = GraphClient("tenant-1", "client-1", "wrong-secret-value-123", MAILBOX, transport=fake)
    with pytest.raises(GraphError) as exc:
        c.search_messages(limit=1)
    assert exc.value.code == "auth_failed" and "wrong-secret-value-123" not in exc.value.message + exc.value.hint


def test_other_mailbox_probe(fake):
    assert client(fake).probe_other_mailbox("ceo2@example.com") == 403
    assert client(fake).probe_other_mailbox(MAILBOX) == 200
