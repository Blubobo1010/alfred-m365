"""`alfred outlook test [--other <address>] | whoami` — connection checks for the operator."""

from __future__ import annotations

import argparse
import json
import sys

from .graph import GraphError
from .settings import load_settings, make_client
from .tools import ping


def register_cli(parser: argparse.ArgumentParser) -> None:
    sub = parser.add_subparsers(dest="outlook_command")
    test = sub.add_parser("test", help="Log in, read the latest mail header and today's calendar, report write settings")
    test.add_argument("--other", default="", help="Another mailbox of the tenant: it must be DENIED (checks the RBAC scope)")
    sub.add_parser("whoami", help="Show the mailbox the tools act on")


def outlook_command(args: argparse.Namespace) -> int:
    command = getattr(args, "outlook_command", None) or "test"
    try:
        settings = load_settings()
        client = make_client(settings)
        if command == "whoami":
            print(f"{client.mailbox} ({settings.timezone}); writes {'on' if settings.write_enabled else 'off'}")
            return 0
        if command == "test":
            report = ping(client, settings, getattr(args, "other", ""))
            print(json.dumps(report, ensure_ascii=False, indent=2))
            if str(report.get("other_mailbox_check", "")).startswith("WARNING"):
                print("The app can read another mailbox: ask IT to fix the RBAC scope before going live.", file=sys.stderr)
                return 1
            print("Outlook connection OK.")
            return 0
    except GraphError as exc:
        print(f"error [{exc.code}]: {exc.message}" + (f"\n  hint: {exc.hint}" if exc.hint else ""), file=sys.stderr)
        return 1
    print(f"unknown outlook command {command!r}", file=sys.stderr)
    return 2
