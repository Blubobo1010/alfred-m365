# alfred-m365

Outlook (Microsoft 365) for [Alfred](https://github.com/Blubobo1010/Alfred-Ai-Audit) / Hermes Agent:
the CEO's mailbox and calendar, over Microsoft Graph with an app-only Entra registration limited by
Exchange **RBAC for Applications** to one mailbox.

- **Read**: `outlook_search_mail`, `outlook_read_mail` (email text marked as untrusted data),
  `outlook_list_events`, `outlook_find_free_time` (working hours from Outlook).
- **Drafts** (`write_enabled`): `outlook_create_draft` (new or reply in the thread) and
  `outlook_update_draft`. Drafts land in Outlook's Drafts folder and are **never sent**: the CEO
  reviews and sends them. Names are resolved among his correspondents; an address never seen in the
  mailbox is flagged so a guessed address does not slip through.
- **Calendar** (`write_enabled`): `outlook_create_event`, `outlook_update_event` (the CEO's own
  events, single occurrences). Every calendar change is **confirmed by a human** first (CLI prompt or
  Telegram buttons); with attendees, the prompt warns that Outlook sends the invitation in his name.

## What it can never do

Exchange has no "drafts only" or "no delete" permission: the app holds `Mail.ReadWrite` and
`Calendars.ReadWrite` on the mailbox. The limits are enforced here, in one place
(`graph.GraphClient._call`): an allow-list of (method, path) pairs on the configured mailbox. No
DELETE, no send, no move/copy, no other mailbox; message PATCH only on drafts; event PATCH only on
events the mailbox organises, never a whole series. Refused calls never reach the network.

## Install

```bash
alfred plugins install Blubobo1010/alfred-m365 --enable      # prompts for MSGRAPH_TENANT_ID / MSGRAPH_CLIENT_ID
alfred outlook keygen                                        # key stays here; send the printed .cer to IT
alfred config set plugins.entries.alfred-m365.settings.mailbox ceo@company.com
alfred config set plugins.entries.alfred-m365.settings.owner_name Luca
alfred outlook test --other someone.else@company.com         # the other mailbox must be DENIED
alfred config set plugins.entries.alfred-m365.settings.write_enabled true
alfred gateway restart
```

Credential: a **certificate** generated on the machine (`alfred outlook keygen`, RSA 2048, two
years, key file mode 600 in `$HERMES_HOME/alfred-m365/`). IT uploads only the public `.cer` in Entra
(Certificates & secrets), so no secret ever travels; `MSGRAPH_CLIENT_SECRET` still works as a fallback.
IT registers the app and assigns the Exchange roles (procedure: RBAC for Applications, no Graph
application permissions and no admin consent in Entra). Settings: `mailbox`, `owner_name`,
`timezone` (default Europe/Rome), `write_enabled`, `approve_drafts` (also confirm drafts),
`default_limit`, `max_body_chars`, `timeout_seconds`.

Note: `alfred -z` (one-shot) runs in YOLO mode and skips every approval by design; use the gateway
(Telegram), the interactive CLI or `alfred chat -q` (approvals denied) instead.

## Development

```bash
uv run --no-project --with pytest --with cryptography pytest -q tests     # fake Graph, no network
uvx ruff check .
uv run --no-project --with pytest python dev/fake_graph.py 8765   # local fake Graph for end-to-end runs
```
End to end in Alfred: settings `graph_base_url: http://127.0.0.1:8765/v1.0` (only loopback URLs are
accepted) and `MSGRAPH_CLIENT_SECRET=s3cr3t-value-for-tests-only`. Live tenant: `dev/live_check.py`.
