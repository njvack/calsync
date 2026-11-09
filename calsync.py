#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "msal>=1.28",
#   "google-auth>=2.29",
#   "google-auth-oauthlib>=1.2",
#   "requests>=2.31",
# ]
# ///
"""Two-way Google Calendar <-> Outlook (Microsoft Graph) sync with dedup.

Copies events between the two calendars as attendee-less, reminder-less
"busy" copies. An event is not copied if an equivalent event already exists
natively on the other side; if one appears later, our copy is deleted.

Commands:
  calsync.py auth [--device]   sign in to Google and Microsoft (once)
  calsync.py run [--dry-run]   one sync pass (launchd calls this)
  calsync.py ours [--delete]   list (or delete) every copy this script made
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tomllib
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

CONFIG_DIR = Path(os.environ.get("CALSYNC_DIR", "~/.config/calsync")).expanduser()
CONFIG_FILE = CONFIG_DIR / "config.toml"
GOOGLE_TOKEN = CONFIG_DIR / "google_token.json"
MSAL_CACHE = CONFIG_DIR / "msal_cache.json"

GOOGLE_SCOPES = ["https://www.googleapis.com/auth/calendar.events"]
MS_SCOPES = ["Calendars.ReadWrite", "User.Read"]  # MSAL adds offline_access itself

# Marker stored on every copy we create. The value is "<side>:<source event id>".
MARK_KEY = "calsync_src"
MS_PROP_ID = "String {6f1c2a8e-4b7d-4e2a-9c51-3d8e0f7a2b94} Name " + MARK_KEY

GRAPH = "https://graph.microsoft.com/v1.0"
GCAL = "https://www.googleapis.com/calendar/v3"


def log(msg: str) -> None:
    print(f"{datetime.now().isoformat(timespec='seconds')} {msg}", flush=True)


# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------

DEFAULTS = {
    "days_back": 1,
    "days_ahead": 60,
    "settle_seconds": 60,
    "min_prefix": 4,
    "private_title": "Busy",
}


def load_config() -> dict:
    if not CONFIG_FILE.exists():
        sys.exit(f"Missing {CONFIG_FILE}; copy config.example.toml there and fill it in.")
    with CONFIG_FILE.open("rb") as f:
        cfg = tomllib.load(f)
    cfg.setdefault("sync", {})
    for k, v in DEFAULTS.items():
        cfg["sync"].setdefault(k, v)
    return cfg


def write_private(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w") as f:
        f.write(text)


# --------------------------------------------------------------------------
# Normalized event
# --------------------------------------------------------------------------

@dataclass
class Event:
    side: str                 # "google" or "outlook"
    id: str
    title: str
    all_day: bool
    start: datetime | date    # UTC datetime, or date when all_day
    end: datetime | date      # exclusive
    updated: datetime
    mark: str | None = None   # source key if this is one of our copies
    copyable: bool = True     # False for declined / free / cancelled / etc.
    private: bool = False
    tentative: bool = False
    raw: dict = field(default_factory=dict, repr=False)

    @property
    def key(self) -> str:
        return f"{self.side}:{self.id}"

    @property
    def slot(self) -> tuple:
        return (self.all_day, self.start, self.end)


_COPY_PREFIX = re.compile(r"^\s*(copy:|\[copy\])\s*", re.I)


def norm_title(t: str) -> str:
    t = _COPY_PREFIX.sub("", t or "")
    return " ".join(t.casefold().split())


def titles_match(a: str, b: str, min_prefix: int) -> bool:
    a, b = norm_title(a), norm_title(b)
    if a == b:
        return True
    short, long_ = sorted((a, b), key=len)
    return len(short) >= min_prefix and long_.startswith(short)


def same_meeting(a: Event, b: Event, min_prefix: int) -> bool:
    return a.slot == b.slot and titles_match(a.title, b.title, min_prefix)


def copy_title(src: Event, cfg: dict) -> str:
    return cfg["sync"]["private_title"] if src.private else (src.title or "(no title)")


_FRAC = re.compile(r"(\.\d{6})\d+")


def parse_ts(s: str) -> datetime:
    """Parse ISO timestamps from either API (Graph uses 7 fractional digits)."""
    s = _FRAC.sub(r"\1", s.replace("Z", "+00:00"))
    dt = datetime.fromisoformat(s)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


# --------------------------------------------------------------------------
# Google
# --------------------------------------------------------------------------

class Google:
    side = "google"

    def __init__(self, cfg: dict):
        import requests
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials

        if not GOOGLE_TOKEN.exists():
            sys.exit("Not signed in to Google; run `calsync.py auth`.")
        creds = Credentials.from_authorized_user_file(str(GOOGLE_TOKEN), GOOGLE_SCOPES)
        if not creds.valid:
            creds.refresh(Request())
            write_private(GOOGLE_TOKEN, creds.to_json())
        self.s = requests.Session()
        self.s.headers["Authorization"] = f"Bearer {creds.token}"
        self.cal = cfg.get("google", {}).get("calendar_id", "primary")
        self.base = f"{GCAL}/calendars/{self.cal}/events"

    @staticmethod
    def login(cfg: dict) -> None:
        from google_auth_oauthlib.flow import InstalledAppFlow

        secret = Path(cfg["google"]["client_secret_file"]).expanduser()
        flow = InstalledAppFlow.from_client_secrets_file(str(secret), GOOGLE_SCOPES)
        creds = flow.run_local_server(port=0, prompt="consent", access_type="offline")
        write_private(GOOGLE_TOKEN, creds.to_json())
        log("Google: signed in")

    def _check(self, r):
        if r.status_code >= 400:
            raise RuntimeError(f"Google {r.request.method} {r.url}: {r.status_code} {r.text[:500]}")
        return r

    def list(self, t0: datetime, t1: datetime) -> list[Event]:
        out, token = [], None
        while True:
            params = {
                "timeMin": t0.isoformat(), "timeMax": t1.isoformat(),
                "singleEvents": "true", "maxResults": 2500, "showDeleted": "false",
            }
            if token:
                params["pageToken"] = token
            data = self._check(self.s.get(self.base, params=params)).json()
            out += [self._parse(e) for e in data.get("items", [])]
            token = data.get("nextPageToken")
            if not token:
                return out

    def _parse(self, e: dict) -> Event:
        all_day = "date" in e["start"]
        if all_day:
            start, end = date.fromisoformat(e["start"]["date"]), date.fromisoformat(e["end"]["date"])
        else:
            start, end = parse_ts(e["start"]["dateTime"]), parse_ts(e["end"]["dateTime"])
        me = next((a for a in e.get("attendees", []) if a.get("self")), {})
        resp = me.get("responseStatus")
        copyable = (
            e.get("status") != "cancelled"
            and resp != "declined"
            and e.get("transparency") != "transparent"
            and e.get("eventType", "default") not in ("workingLocation", "birthday")
        )
        return Event(
            side=self.side, id=e["id"], title=e.get("summary", ""), all_day=all_day,
            start=start, end=end, updated=parse_ts(e["updated"]),
            mark=e.get("extendedProperties", {}).get("private", {}).get(MARK_KEY),
            copyable=copyable,
            private=e.get("visibility") in ("private", "confidential"),
            tentative=resp == "tentative", raw=e,
        )

    def _body(self, src: Event, cfg: dict) -> dict:
        if src.all_day:
            when = {"start": {"date": src.start.isoformat()}, "end": {"date": src.end.isoformat()}}
        else:
            when = {"start": {"dateTime": src.start.isoformat()}, "end": {"dateTime": src.end.isoformat()}}
        return {
            **when,
            "summary": copy_title(src, cfg),
            "description": f"Copied from {src.side} by calsync.",
            "transparency": "opaque",
            "visibility": "private" if src.private else "default",
            "reminders": {"useDefault": False, "overrides": []},
            "extendedProperties": {"private": {MARK_KEY: src.key}},
        }

    def create(self, src: Event, cfg: dict) -> None:
        self._check(self.s.post(self.base, params={"sendUpdates": "none"}, json=self._body(src, cfg)))

    def update(self, copy: Event, src: Event, cfg: dict) -> None:
        self._check(self.s.patch(f"{self.base}/{copy.id}", params={"sendUpdates": "none"},
                                 json=self._body(src, cfg)))

    def delete(self, copy: Event) -> None:
        r = self.s.delete(f"{self.base}/{copy.id}", params={"sendUpdates": "none"})
        if r.status_code not in (404, 410):
            self._check(r)


# --------------------------------------------------------------------------
# Microsoft Graph
# --------------------------------------------------------------------------

def _msal_app(cfg: dict):
    import msal

    cache = msal.SerializableTokenCache()
    if MSAL_CACHE.exists():
        cache.deserialize(MSAL_CACHE.read_text())
    ms = cfg["microsoft"]
    app = msal.PublicClientApplication(
        ms["client_id"],
        authority=f"https://login.microsoftonline.com/{ms['tenant_id']}",
        token_cache=cache,
    )
    return app, cache


def _save_cache(cache) -> None:
    if cache.has_state_changed:
        write_private(MSAL_CACHE, cache.serialize())


class Outlook:
    side = "outlook"

    def __init__(self, cfg: dict):
        import requests

        app, cache = _msal_app(cfg)
        accounts = app.get_accounts()
        result = app.acquire_token_silent(MS_SCOPES, account=accounts[0]) if accounts else None
        _save_cache(cache)
        if not result or "access_token" not in result:
            sys.exit("Microsoft sign-in expired or missing; run `calsync.py auth`.")
        self.s = requests.Session()
        self.s.headers["Authorization"] = f"Bearer {result['access_token']}"
        self.s.headers["Prefer"] = 'outlook.timezone="UTC"'
        self._mailbox_tz = None

    @staticmethod
    def login(cfg: dict, device: bool = False) -> None:
        app, cache = _msal_app(cfg)
        if device:
            flow = app.initiate_device_flow(scopes=MS_SCOPES)
            if "user_code" not in flow:
                sys.exit(f"Device flow failed: {flow}")
            print(flow["message"], flush=True)
            result = app.acquire_token_by_device_flow(flow)
        else:
            result = app.acquire_token_interactive(MS_SCOPES)
        if "access_token" not in result:
            sys.exit(f"Microsoft sign-in failed: {result.get('error')}: {result.get('error_description')}")
        _save_cache(cache)
        log("Microsoft: signed in")

    def _check(self, r):
        if r.status_code >= 400:
            raise RuntimeError(f"Graph {r.request.method} {r.url}: {r.status_code} {r.text[:500]}")
        return r

    @property
    def mailbox_tz(self) -> str:
        # Needed so all-day copies land on the right date in Outlook.
        if self._mailbox_tz is None:
            self._mailbox_tz = self._check(self.s.get(f"{GRAPH}/me/mailboxSettings/timeZone")).json()["value"]
        return self._mailbox_tz

    def list(self, t0: datetime, t1: datetime) -> list[Event]:
        url = f"{GRAPH}/me/calendarView"
        params = {
            "startDateTime": t0.isoformat(), "endDateTime": t1.isoformat(), "$top": 250,
            "$select": "id,subject,start,end,isAllDay,lastModifiedDateTime,showAs,"
                       "sensitivity,responseStatus,isCancelled,type",
            "$expand": f"singleValueExtendedProperties($filter=id eq '{MS_PROP_ID}')",
        }
        out = []
        while url:
            data = self._check(self.s.get(url, params=params)).json()
            out += [self._parse(e) for e in data.get("value", [])]
            url, params = data.get("@odata.nextLink"), None
        return out

    def _parse(self, e: dict) -> Event:
        start, end = parse_ts(e["start"]["dateTime"]), parse_ts(e["end"]["dateTime"])
        if e.get("isAllDay"):
            # Returned as UTC; shift back to local time to recover the dates.
            start, end = start.astimezone().date(), end.astimezone().date()
        mark = next((p["value"] for p in e.get("singleValueExtendedProperties", [])
                     if p["id"].lower() == MS_PROP_ID.lower()), None)
        resp = (e.get("responseStatus") or {}).get("response")
        copyable = (
            not e.get("isCancelled")
            and resp != "declined"
            and e.get("showAs") not in ("free", "workingElsewhere")
        )
        return Event(
            side=self.side, id=e["id"], title=e.get("subject") or "", all_day=bool(e.get("isAllDay")),
            start=start, end=end, updated=parse_ts(e["lastModifiedDateTime"]), mark=mark,
            copyable=copyable, private=e.get("sensitivity") in ("private", "confidential"),
            tentative=resp == "tentativelyAccepted" or e.get("showAs") == "tentative", raw=e,
        )

    def _body(self, src: Event, cfg: dict) -> dict:
        if src.all_day:
            tz = self.mailbox_tz
            when = {"isAllDay": True,
                    "start": {"dateTime": f"{src.start.isoformat()}T00:00:00", "timeZone": tz},
                    "end": {"dateTime": f"{src.end.isoformat()}T00:00:00", "timeZone": tz}}
        else:
            fmt = "%Y-%m-%dT%H:%M:%S"
            when = {"isAllDay": False,
                    "start": {"dateTime": src.start.strftime(fmt), "timeZone": "UTC"},
                    "end": {"dateTime": src.end.strftime(fmt), "timeZone": "UTC"}}
        return {
            **when,
            "subject": copy_title(src, cfg),
            "body": {"contentType": "text", "content": f"Copied from {src.side} by calsync."},
            "showAs": "tentative" if src.tentative else "busy",
            "sensitivity": "private" if src.private else "normal",
            "isReminderOn": False,
            "singleValueExtendedProperties": [{"id": MS_PROP_ID, "value": src.key}],
        }

    def create(self, src: Event, cfg: dict) -> None:
        self._check(self.s.post(f"{GRAPH}/me/events", json=self._body(src, cfg)))

    def update(self, copy: Event, src: Event, cfg: dict) -> None:
        body = self._body(src, cfg)
        body.pop("singleValueExtendedProperties")
        self._check(self.s.patch(f"{GRAPH}/me/events/{copy.id}", json=body))

    def delete(self, copy: Event) -> None:
        r = self.s.delete(f"{GRAPH}/me/events/{copy.id}")
        if r.status_code != 404:
            self._check(r)


# --------------------------------------------------------------------------
# Reconcile
# --------------------------------------------------------------------------

def plan(src_side: str, src_events: list[Event], dst_events: list[Event], cfg: dict,
         now: datetime) -> list[tuple]:
    """Return actions for copying src -> dst.

    Each action is ("create", src), ("update", copy, src) or ("delete", copy, reason).
    """
    min_prefix = cfg["sync"]["min_prefix"]
    settle = timedelta(seconds=cfg["sync"]["settle_seconds"])

    sources = [e for e in src_events if e.mark is None and e.copyable]
    natives = [e for e in dst_events if e.mark is None]
    copies: dict[str, list[Event]] = {}
    for e in dst_events:
        if e.mark and e.mark.startswith(f"{src_side}:"):
            copies.setdefault(e.mark, []).append(e)

    actions: list[tuple] = []
    seen: set[str] = set()
    for s in sources:
        seen.add(s.key)
        mine = copies.get(s.key, [])
        for extra in mine[1:]:
            actions.append(("delete", extra, "duplicate copy"))
        mine = mine[:1]
        native = any(same_meeting(s, n, min_prefix) for n in natives)
        if native:
            for c in mine:
                actions.append(("delete", c, "native event exists"))
            continue
        if mine:
            c = mine[0]
            if c.slot != s.slot or c.title != copy_title(s, cfg):
                actions.append(("update", c, s))
        elif now - s.updated >= settle:
            actions.append(("create", s))

    for key, mine in copies.items():
        if key not in seen:
            for c in mine:
                actions.append(("delete", c, "source gone, declined, or freed"))
    return actions


def describe(e: Event) -> str:
    when = e.start.isoformat() if e.all_day else e.start.strftime("%Y-%m-%d %H:%MZ")
    return f"{e.side} [{when}] {e.title!r}"


def apply(actions: list[tuple], dst, cfg: dict, dry_run: bool) -> None:
    for a in actions:
        verb = a[0]
        if verb == "create":
            log(f"create in {dst.side}: {describe(a[1])}")
            if not dry_run:
                dst.create(a[1], cfg)
        elif verb == "update":
            log(f"update in {dst.side}: {describe(a[1])} -> {describe(a[2])}")
            if not dry_run:
                dst.update(a[1], a[2], cfg)
        elif verb == "delete":
            log(f"delete in {dst.side}: {describe(a[1])} ({a[2]})")
            if not dry_run:
                dst.delete(a[1])


def window(cfg: dict) -> tuple[datetime, datetime, datetime]:
    now = datetime.now(timezone.utc)
    t0 = (now - timedelta(days=cfg["sync"]["days_back"])).replace(hour=0, minute=0, second=0, microsecond=0)
    t1 = now + timedelta(days=cfg["sync"]["days_ahead"])
    return now, t0, t1


def cmd_run(cfg: dict, dry_run: bool) -> None:
    now, t0, t1 = window(cfg)
    g, o = Google(cfg), Outlook(cfg)
    # Fetch everything before changing anything; any fetch error aborts the run.
    ge, oe = g.list(t0, t1), o.list(t0, t1)
    a1 = plan("google", ge, oe, cfg, now)
    a2 = plan("outlook", oe, ge, cfg, now)
    if dry_run:
        log(f"dry run: {len(ge)} google, {len(oe)} outlook events in window")
    apply(a1, o, cfg, dry_run)
    apply(a2, g, cfg, dry_run)
    if not (a1 or a2) and dry_run:
        log("nothing to do")


def cmd_ours(cfg: dict, delete: bool) -> None:
    _, t0, t1 = window(cfg)
    for side in (Google(cfg), Outlook(cfg)):
        for e in side.list(t0, t1):
            if e.mark:
                log(f"{'deleting' if delete else 'copy'}: {describe(e)} <- {e.mark}")
                if delete:
                    side.delete(e)


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("auth", help="sign in to Google and Microsoft")
    a.add_argument("--device", action="store_true", help="use Microsoft device-code sign-in")
    a.add_argument("--only", choices=["google", "microsoft"])
    r = sub.add_parser("run", help="one sync pass")
    r.add_argument("--dry-run", action="store_true")
    o = sub.add_parser("ours", help="list copies this script created")
    o.add_argument("--delete", action="store_true", help="delete them all (undo)")
    args = p.parse_args()

    cfg = load_config()
    if args.cmd == "auth":
        if args.only != "microsoft":
            Google.login(cfg)
        if args.only != "google":
            Outlook.login(cfg, device=args.device)
    elif args.cmd == "run":
        cmd_run(cfg, args.dry_run)
    elif args.cmd == "ours":
        cmd_ours(cfg, args.delete)


if __name__ == "__main__":
    main()
