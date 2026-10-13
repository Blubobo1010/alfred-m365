"""A local fake Microsoft Graph (the unit tests' FakeGraph over HTTP) for end-to-end runs in Alfred.

    python dev/fake_graph.py 8765
    ALFRED_M365_GRAPH_URL=http://127.0.0.1:8765/v1.0 MSGRAPH_CLIENT_SECRET=s3cr3t-value-for-tests-only …

Every request is printed, so you can see exactly what Alfred asked Microsoft 365 to do.
"""

from __future__ import annotations

import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tests"))
from conftest import FakeGraph  # noqa: E402

GRAPH = FakeGraph()


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
    print(f"fake Graph on http://127.0.0.1:{port}/v1.0 (mailbox luca@example.com)", flush=True)
    ThreadingHTTPServer(("127.0.0.1", port), Handler).serve_forever()
