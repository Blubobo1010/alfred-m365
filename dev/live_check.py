"""Live check against a REAL tenant (Kimera on site, or a test tenant). Creates one draft and one
event in the configured mailbox, edits both, prints their links. Deletes nothing: clean up by hand.

    MSGRAPH_TENANT_ID=… MSGRAPH_CLIENT_ID=… MSGRAPH_CLIENT_SECRET=… \\
    LIVE_MAILBOX=luca.betti@kimera-automobili.com [LIVE_OTHER=someone.else@kimera-automobili.com] \\
    uv run --no-project python dev/live_check.py
"""

from __future__ import annotations

import importlib.util
import json
import os
import sys
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("alfred_m365", ROOT / "__init__.py", submodule_search_locations=[str(ROOT)])
pkg = importlib.util.module_from_spec(spec)
sys.modules["alfred_m365"] = pkg
spec.loader.exec_module(pkg)

from alfred_m365 import settings, tools  # noqa: E402


class Ctx:
    def get_config(self, key, default=None):
        return {"mailbox": os.environ["LIVE_MAILBOX"], "write_enabled": True, "owner_name": "test"}.get(key, default)


def show(label: str, raw: str) -> dict:
    data = json.loads(raw)
    print(f"--- {label}: {'OK' if data.get('ok') else 'FAILED'}")
    print(json.dumps(data, ensure_ascii=False, indent=2)[:1500])
    if not data.get("ok"):
        sys.exit(1)
    return data


settings.bind_context(Ctx())
client = settings.make_client()
print(json.dumps(tools.ping(client, settings.load_settings(), os.environ.get("LIVE_OTHER", "")), indent=2))
show("search", tools.outlook_search_mail({"limit": 3}))
show("events", tools.outlook_list_events({}))
day = date.today() + timedelta(days=7)
show("free time", tools.outlook_find_free_time({"date_from": day.isoformat(), "duration_minutes": 30}))
draft = show("draft", tools.outlook_create_draft({"to": [os.environ["LIVE_MAILBOX"]], "subject": "[TEST Alfred] bozza",
                                                  "body": "Bozza di prova creata da alfred-m365. Si può cancellare."}))
show("draft update", tools.outlook_update_draft({"id": draft["draft_id"], "body": "Bozza di prova (modificata). Si può cancellare."}))
event = show("event", tools.outlook_create_event({"subject": "[TEST Alfred] appuntamento", "start": f"{day}T07:00",
                                                  "duration_minutes": 15, "body": "Prova di alfred-m365, si può cancellare."}))
show("event move", tools.outlook_update_event({"id": event["id"], "start": f"{day}T07:30"}))
print("\nDone. Delete the [TEST Alfred] draft and appointment by hand in Outlook.")
