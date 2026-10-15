"""A local fake Microsoft Graph (the unit tests' FakeGraph over HTTP) for end-to-end runs in Alfred.

    python dev/fake_graph.py 8765
    ALFRED_M365_GRAPH_URL=http://127.0.0.1:8765/v1.0 MSGRAPH_CLIENT_SECRET=s3cr3t-value-for-tests-only …

Every request is printed, so you can see exactly what Alfred asked Microsoft 365 to do. Besides the unit tests' data
it holds a meeting with outside people starting about half an hour after launch, and some mail with them: what
alfred-brain's meeting brief picks up (``alfred brain brief``).
"""

from __future__ import annotations

import sys
from datetime import UTC, datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tests"))
from conftest import MAILBOX, FakeGraph, mid  # noqa: E402

GRAPH = FakeGraph()
CUSTOMER = ("Mario Rossi", "mario.rossi@beta.example")
COLLEAGUE = ("Paolo Neri", "paolo.neri@beta.example")


def add_brief_demo(graph: FakeGraph, minutes: int = 32) -> datetime:
    """A meeting starting ``minutes`` from now, organised by a customer, plus one email each way with him."""
    start = (datetime.now(ZoneInfo("Europe/Rome")) + timedelta(minutes=minutes)).replace(second=0, microsecond=0)

    def at(dt: datetime) -> dict:
        return {"dateTime": dt.strftime("%Y-%m-%dT%H:%M:%S.0000000"), "timeZone": "W. Europe Standard Time"}

    def who(person: tuple[str, str]) -> dict:
        return {"emailAddress": {"name": person[0], "address": person[1]}}

    graph.events["EVTBRIEF001"] = {
        "id": "EVTBRIEF001", "subject": "Revisione Z20", "start": at(start), "end": at(start + timedelta(minutes=45)),
        "isOrganizer": False, "type": "singleInstance", "isAllDay": False, "isCancelled": False, "sensitivity": "normal",
        "responseStatus": {"response": "accepted"}, "showAs": "busy", "isOnlineMeeting": True, "webLink": "webrief",
        "organizer": who(CUSTOMER), "attendees": [
            {**who(CUSTOMER), "type": "required"}, {**who(COLLEAGUE), "type": "optional"},
            {**who(("Luca", MAILBOX)), "type": "required"}, {**who(("Sala Riunioni", "sala@example.com")), "type": "resource"}],
        "bodyPreview": "Punti: tempi di consegna e prezzi ______________________________ Riunione di Microsoft Teams "
                       "Partecipa da computer, app per dispositivi mobili o dispositivo della sala"}
    now = datetime.now(UTC)
    for n, (days, sender, to, subject) in enumerate(((1, CUSTOMER, ("Luca", MAILBOX), "Disegni cofano"),
                                                     (5, ("Luca", MAILBOX), CUSTOMER, "Offerta Z20 rev. 2")), start=901):
        when = (now - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
        graph.messages[mid(n)] = {"id": mid(n), "subject": subject, "from": who(sender), "toRecipients": [who(to)],
                                  "ccRecipients": [], "receivedDateTime": when, "sentDateTime": when, "bodyPreview": subject,
                                  "isRead": True, "isDraft": False, "hasAttachments": False, "webLink": f"w{n}",
                                  "body": {"contentType": "text", "content": subject}}
    return start


class Handler(BaseHTTPRequestHandler):
    def _serve(self) -> None:
        length = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(length) if length else None
        url = f"http://{self.headers.get('Host')}{self.path}"
        status, headers, payload = GRAPH(self.command, url, dict(self.headers.items()), body, 30)
        print(f"{self.command} {self.path} -> {status}", flush=True)
        self.send_response(status)
        for k, v in {"Content-Type": "application/json", **headers}.items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(payload)

    do_GET = do_POST = do_PATCH = do_DELETE = _serve

    def log_message(self, *args) -> None:  # the one-line print above is enough
        pass


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
    start = add_brief_demo(GRAPH)
    print(f"fake Graph on http://127.0.0.1:{port}/v1.0 (mailbox {MAILBOX}); "
          f"meeting with {CUSTOMER[0]} at {start:%H:%M}", flush=True)
    ThreadingHTTPServer(("127.0.0.1", port), Handler).serve_forever()
