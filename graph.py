"""Microsoft Graph over urllib, app-only (client credentials), behind an allow-list.

Exchange RBAC for Applications grants ``Mail.ReadWrite`` and ``Calendars.ReadWrite`` on the CEO's
mailbox: Microsoft has no "drafts only" or "never delete" permission, so those limits live HERE.
Every request goes through ``GraphClient._call``, which refuses anything that is not one of the
``ALLOWED`` (method, path) pairs on the configured mailbox before touching the network: no DELETE,
no send, no move/copy, no other mailbox. Message PATCH is further limited to drafts
(``update_draft`` checks ``isDraft`` first) and event PATCH to events the mailbox organises.
"""

from __future__ import annotations

import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from collections.abc import Callable
from typing import Any

__version__ = "0.1.0"

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
AUTHORITY_URL = "https://login.microsoftonline.com"
SCOPE = "https://graph.microsoft.com/.default"
TOKEN_SKEW_SECONDS = 120
MAX_ATTEMPTS = 3

_ID = r"[A-Za-z0-9=_\-+/%]{8,600}"
_FOLDERS = "inbox|sentitems|drafts|archive"
# (method, path regex relative to /users/{mailbox}) — the complete list of what this plugin may do.
ALLOWED: tuple[tuple[str, str], ...] = (
    ("GET", r"/messages"),
    ("GET", rf"/mailFolders/(?:{_FOLDERS})/messages"),
    ("GET", rf"/messages/{_ID}"),
    ("GET", rf"/messages/{_ID}/attachments"),
    ("POST", r"/messages"),  # creates a DRAFT (Graph never sends on this endpoint)
    ("POST", rf"/messages/{_ID}/createReply(?:All)?"),  # creates a reply DRAFT
    ("PATCH", rf"/messages/{_ID}"),  # drafts only: see GraphClient.update_draft
    ("GET", r"/calendarView"),
    ("GET", rf"/events/{_ID}"),
    ("POST", r"/events"),
    ("PATCH", rf"/events/{_ID}"),  # organiser's own events only: see GraphClient.update_event
    ("POST", r"/calendar/getSchedule"),  # free/busy query, read-only despite the verb
)
_ALLOWED_RE = tuple((m, re.compile(p + r"$")) for m, p in ALLOWED)

# Windows time zone names for the zones we expect; Graph also accepts IANA names, these are safer.
WINDOWS_TZ = {"Europe/Rome": "W. Europe Standard Time", "Europe/Berlin": "W. Europe Standard Time",
              "Europe/Paris": "Romance Standard Time", "Europe/London": "GMT Standard Time", "UTC": "UTC"}

Transport = Callable[[str, str, dict[str, str], bytes | None, float], tuple[int, dict[str, str], bytes]]


class GraphError(Exception):
    def __init__(self, code: str, message: str, hint: str = "", status: int = 0, retry_after: float | None = None):
        super().__init__(message)
        self.code, self.message, self.hint, self.status, self.retry_after = code, message, hint, status, retry_after


class NotAllowed(GraphError):
    def __init__(self, method: str, path: str):
        super().__init__("not_allowed", f"refused: {method} {path} is not an allowed Outlook operation",
                         "This plugin only reads mail and calendar, creates and edits drafts, and creates or edits "
                         "events of the configured mailbox. It never sends, moves or deletes.")


def default_transport(method: str, url: str, headers: dict[str, str], body: bytes | None,
                      timeout: float) -> tuple[int, dict[str, str], bytes]:
    req = urllib.request.Request(url, data=body, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:  # noqa: S310 - fixed https Graph/login hosts
            return resp.status, dict(resp.headers.items()), resp.read()
    except urllib.error.HTTPError as exc:  # 4xx/5xx carry Graph's JSON error body
        return exc.code, dict(exc.headers.items()) if exc.headers else {}, exc.read()


def graph_tz(name: str) -> str:
    return WINDOWS_TZ.get(name, name or "UTC")


class GraphClient:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str, mailbox: str, *, timezone: str = "Europe/Rome",
                 timeout: float = 30, base_url: str = GRAPH_BASE_URL, authority_url: str = AUTHORITY_URL,
                 transport: Transport | None = None, sleep: Callable[[float], None] | None = None):
        if not re.fullmatch(r"[^@\s/]+@[^@\s/]+\.[^@\s/]+", mailbox or ""):
            raise GraphError("bad_config", f"mailbox {mailbox!r} is not an email address",
                             "Set plugins.entries.alfred-m365.settings.mailbox to the CEO's address.")
        self.mailbox = mailbox.strip().lower()
        self.timezone = timezone or "Europe/Rome"
        self._tenant, self._client_id, self._secret = tenant_id.strip(), client_id.strip(), client_secret
        self._timeout, self._base, self._authority = timeout, base_url.rstrip("/"), authority_url.rstrip("/")
        self._transport = transport or default_transport
        self._sleep = sleep or (lambda seconds: time.sleep(seconds))
        self._token: tuple[str, float] | None = None

    # ---- auth -----------------------------------------------------------------------------

    def _access_token(self, force: bool = False) -> str:
        if self._token and not force and self._token[1] - TOKEN_SKEW_SECONDS > time.time():
            return self._token[0]
        form = urllib.parse.urlencode({"client_id": self._client_id, "client_secret": self._secret,
                                       "scope": SCOPE, "grant_type": "client_credentials"}).encode()
        url = f"{self._authority}/{urllib.parse.quote(self._tenant, safe='')}/oauth2/v2.0/token"
        try:
            status, _h, raw = self._transport("POST", url, {"Content-Type": "application/x-www-form-urlencoded"},
                                              form, self._timeout)
        except (urllib.error.URLError, TimeoutError, ConnectionError) as exc:
            raise GraphError("unreachable", f"Microsoft login unreachable ({exc.__class__.__name__})", retryable_hint()) from None
        data = _json(raw)
        if status != 200 or not isinstance(data, dict) or not data.get("access_token"):
            desc = str((data or {}).get("error_description") or (data or {}).get("error") or f"HTTP {status}")
            desc = desc.replace(self._secret, "***") if self._secret else desc
            raise GraphError("auth_failed", f"Microsoft login failed: {desc.splitlines()[0][:300]}",
                             "Check MSGRAPH_TENANT_ID, MSGRAPH_CLIENT_ID and MSGRAPH_CLIENT_SECRET (expired secret?).",
                             status=status)
        self._token = (str(data["access_token"]), time.time() + float(data.get("expires_in") or 3600))
        return self._token[0]

    # ---- the single choke point -------------------------------------------------------------

    def _call(self, method: str, path: str, *, params: dict[str, Any] | None = None, body: Any = None,
              prefer: list[str] | None = None) -> Any:
        method = method.upper()
        if not any(m == method and rx.match(path) for m, rx in _ALLOWED_RE):
            raise NotAllowed(method, path)
        url = f"{self._base}/users/{urllib.parse.quote(self.mailbox, safe='@.')}{path}"
        if params:
            url += "?" + urllib.parse.urlencode(params, quote_via=urllib.parse.quote)
        return self._request(method, url, body, prefer)

    def _request(self, method: str, url: str, body: Any, prefer: list[str] | None) -> Any:
        payload = json.dumps(body).encode("utf-8") if body is not None else None
        refreshed = False
        for attempt in range(1, MAX_ATTEMPTS + 1):
            headers = {"Authorization": f"Bearer {self._access_token()}",
                       "Accept": "application/json", "User-Agent": f"alfred-m365/{__version__}"}
            if payload is not None:
                headers["Content-Type"] = "application/json"
            if prefer:
                headers["Prefer"] = ", ".join(prefer)
            try:
                status, resp_headers, raw = self._transport(method, url, headers, payload, self._timeout)
            except (urllib.error.URLError, TimeoutError, ConnectionError) as exc:
                if attempt == MAX_ATTEMPTS:
                    raise GraphError("unreachable", f"Microsoft Graph unreachable ({exc.__class__.__name__})",
                                     retryable_hint()) from None
                self._sleep(2 ** attempt)
                continue
            if status == 401 and not refreshed:
                refreshed = True
                self._token = None
                continue
            if status in (429, 500, 502, 503, 504) and attempt < MAX_ATTEMPTS:
                self._sleep(_retry_after(resp_headers) or 2 ** attempt)
                continue
            if status >= 400:
                raise _graph_error(status, _json(raw), _retry_after(resp_headers))
            return _json(raw) if raw else {}
        raise GraphError("unreachable", "Microsoft Graph did not answer", retryable_hint())

    # ---- mail ---------------------------------------------------------------------------------

    MESSAGE_FIELDS = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,bodyPreview,isRead,isDraft," \
                     "hasAttachments,webLink,conversationId,importance"

    def search_messages(self, *, query: str = "", sender: str = "", since: str = "", folder: str = "",
                        unread_only: bool = False, limit: int = 20) -> list[dict[str, Any]]:
        path = f"/mailFolders/{folder}/messages" if folder else "/messages"
        params: dict[str, Any] = {"$top": limit, "$select": self.MESSAGE_FIELDS}
        if query:  # $search cannot be combined with $orderby / most $filter: fold the rest into KQL
            kql = [query.replace('"', " ")]
            if sender:
                kql.append(f"from:{sender}")
            if since:
                kql.append(f"received>={since}")
            params["$search"] = '"' + " ".join(kql) + '"'
        else:
            filters = []
            if sender:
                filters.append(f"from/emailAddress/address eq '{sender.replace(chr(39), '')}'")
            if since:
                filters.append(f"receivedDateTime ge {since}T00:00:00Z")
            if unread_only:
                filters.append("isRead eq false")
            if filters:
                params["$filter"] = " and ".join(filters)
            params["$orderby"] = "receivedDateTime desc"
        data = self._call("GET", path, params=params, prefer=[f'outlook.timezone="{graph_tz(self.timezone)}"'])
        items = list((data or {}).get("value") or [])
        return [m for m in items if not unread_only or not m.get("isRead")]

    def get_message(self, message_id: str) -> dict[str, Any]:
        return self._call("GET", f"/messages/{_enc(message_id)}",
                          params={"$select": self.MESSAGE_FIELDS + ",body,replyTo"},
                          prefer=['outlook.body-content-type="text"', f'outlook.timezone="{graph_tz(self.timezone)}"'])

    def attachment_names(self, message_id: str) -> list[dict[str, Any]]:
        data = self._call("GET", f"/messages/{_enc(message_id)}/attachments", params={"$select": "name,size,contentType"})
        return list((data or {}).get("value") or [])

    def create_draft(self, *, subject: str, body: str, to: list[str], cc: list[str]) -> dict[str, Any]:
        return self._call("POST", "/messages", body={
            "subject": subject, "body": {"contentType": "Text", "content": body},
            "toRecipients": _recipients(to), "ccRecipients": _recipients(cc)})

    def create_reply_draft(self, message_id: str, *, body: str, reply_all: bool, to: list[str], cc: list[str]) -> dict[str, Any]:
        action = "createReplyAll" if reply_all else "createReply"
        draft = self._call("POST", f"/messages/{_enc(message_id)}/{action}", body={"comment": body})
        changes: dict[str, Any] = {}
        if to:
            changes["toRecipients"] = _recipients(to)
        if cc:
            changes["ccRecipients"] = _recipients(cc)
        if changes and draft.get("id"):
            draft = self.update_draft(str(draft["id"]), changes) or draft
        return draft

    def update_draft(self, message_id: str, changes: dict[str, Any]) -> dict[str, Any]:
        current = self._call("GET", f"/messages/{_enc(message_id)}", params={"$select": "id,isDraft,body"})
        if not current.get("isDraft"):
            raise GraphError("not_a_draft", "refused: that message is not a draft",
                             "Only drafts can be edited; received and sent mail is never modified.")
        return self._call("PATCH", f"/messages/{_enc(message_id)}", body=changes) or current

    def draft_body(self, message_id: str) -> dict[str, Any]:
        return self._call("GET", f"/messages/{_enc(message_id)}", params={"$select": "id,isDraft,body"})

    # ---- calendar -----------------------------------------------------------------------------

    EVENT_FIELDS = "id,subject,start,end,location,attendees,organizer,isOrganizer,isAllDay,isOnlineMeeting,onlineMeeting," \
                   "showAs,type,seriesMasterId,webLink,isCancelled,bodyPreview"

    def list_events(self, start: str, end: str, limit: int = 100) -> list[dict[str, Any]]:
        data = self._call("GET", "/calendarView", params={
            "startDateTime": start, "endDateTime": end, "$select": self.EVENT_FIELDS,
            "$orderby": "start/dateTime", "$top": limit}, prefer=[f'outlook.timezone="{graph_tz(self.timezone)}"'])
        return list((data or {}).get("value") or [])

    def get_event(self, event_id: str) -> dict[str, Any]:
        return self._call("GET", f"/events/{_enc(event_id)}", params={"$select": self.EVENT_FIELDS},
                          prefer=[f'outlook.timezone="{graph_tz(self.timezone)}"'])

    def create_event(self, event: dict[str, Any]) -> dict[str, Any]:
        return self._call("POST", "/events", body=event, prefer=[f'outlook.timezone="{graph_tz(self.timezone)}"'])

    def update_event(self, event_id: str, changes: dict[str, Any]) -> dict[str, Any]:
        current = self.get_event(event_id)
        if current.get("isOrganizer") is False:
            raise GraphError("not_organizer", "refused: this meeting was organised by someone else",
                             "Only the organiser can move it: propose a new time to them in a reply draft instead.")
        if current.get("type") == "seriesMaster":
            raise GraphError("series", "refused: this is a whole recurring series",
                             "Edit a single occurrence (use the occurrence id from outlook_list_events).")
        return self._call("PATCH", f"/events/{_enc(event_id)}", body=changes,
                          prefer=[f'outlook.timezone="{graph_tz(self.timezone)}"']) or current

    def schedule(self, start: str, end: str, interval: int = 30) -> dict[str, Any]:
        tz = graph_tz(self.timezone)
        data = self._call("POST", "/calendar/getSchedule", body={
            "schedules": [self.mailbox], "startTime": {"dateTime": start, "timeZone": tz},
            "endTime": {"dateTime": end, "timeZone": tz}, "availabilityViewInterval": interval},
            prefer=[f'outlook.timezone="{tz}"'])
        values = (data or {}).get("value") or [{}]
        return values[0] if values else {}

    # ---- operator checks ------------------------------------------------------------------------

    def probe_other_mailbox(self, address: str) -> int:
        """HTTP status of reading ONE message header of another mailbox: 403 is the healthy answer (the
        RBAC scope excludes it). Used only by ``alfred outlook test``; the result is never returned."""
        url = (f"{self._base}/users/{urllib.parse.quote(address.strip().lower(), safe='@.')}/messages?"
               + urllib.parse.urlencode({"$top": 1, "$select": "id"}))
        try:
            self._request("GET", url, None, None)
            return 200
        except GraphError as exc:
            return exc.status


def _enc(value: str) -> str:
    value = str(value or "").strip()
    if not re.fullmatch(_ID, value):
        raise GraphError("bad_id", "invalid Outlook id", "Use an id returned by outlook_search_mail or outlook_list_events.")
    return urllib.parse.quote(value, safe="")


def _recipients(addresses: list[str]) -> list[dict[str, Any]]:
    out = []
    for a in addresses or []:
        a = str(a).strip()
        if not re.fullmatch(r"[^@\s<>,;]+@[^@\s<>,;]+\.[^@\s<>,;]+", a):
            raise GraphError("bad_address", f"{a!r} is not an email address", "Pass plain addresses like name@company.com.")
        out.append({"emailAddress": {"address": a}})
    return out


def _json(raw: bytes) -> Any:
    try:
        return json.loads(raw.decode("utf-8")) if raw else None
    except (UnicodeDecodeError, json.JSONDecodeError):
        return None


def _retry_after(headers: dict[str, str]) -> float | None:
    value = next((v for k, v in (headers or {}).items() if k.lower() == "retry-after"), None)
    try:
        return min(float(value), 60.0) if value is not None else None
    except ValueError:
        return None


def retryable_hint() -> str:
    return "Network problem: retry in a minute."


def _graph_error(status: int, data: Any, retry_after: float | None) -> GraphError:
    err = (data or {}).get("error") if isinstance(data, dict) else None
    code = str((err or {}).get("code") or "") if isinstance(err, dict) else ""
    message = str((err or {}).get("message") or f"HTTP {status}") if isinstance(err, dict) else f"HTTP {status}"
    if status == 403 or code in ("ErrorAccessDenied", "Authorization_RequestDenied"):
        return GraphError("forbidden", f"Microsoft 365 denied access: {message[:200]}",
                          "IT must grant the app this permission on the mailbox (RBAC for Applications); "
                          "Test-ServicePrincipalAuthorization should report InScope True. Changes take up to 2 hours.",
                          status=status)
    if status == 404 or code in ("ErrorItemNotFound", "ResourceNotFound"):
        return GraphError("not_found", "not found (deleted, moved, or wrong id)", "Search again to get a fresh id.", status=status)
    if status == 429:
        return GraphError("throttled", "Microsoft 365 is throttling requests", "Retry in a minute.", status=status,
                          retry_after=retry_after)
    if status == 400:
        return GraphError("bad_request", f"Microsoft 365 rejected the request: {message[:300]}", "", status=status)
    return GraphError("graph_error", f"Microsoft Graph error {status}: {message[:200]}", retryable_hint(), status=status)
