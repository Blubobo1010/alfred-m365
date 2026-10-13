"""`alfred outlook test [--other <address>] | whoami` — connection checks for the operator."""

from __future__ import annotations

import argparse
import json
import sys

from . import certauth
from .graph import GraphError
from .settings import cert_dir, load_settings, make_client
from .tools import ping


def register_cli(parser: argparse.ArgumentParser) -> None:
    sub = parser.add_subparsers(dest="outlook_command")
    test = sub.add_parser("test", help="Log in, read the latest mail header and today's calendar, report write settings")
    test.add_argument("--other", default="", help="Another mailbox of the tenant: it must be DENIED (checks the RBAC scope)")
    sub.add_parser("whoami", help="Show the mailbox the tools act on")
    keygen = sub.add_parser("keygen", help="Create the app's certificate on this machine (the private key never leaves it); "
                            "send the printed .cer file to IT for Entra > Certificates & secrets")
    keygen.add_argument("--name", default="Alfred", help="Certificate common name (default Alfred)")
    keygen.add_argument("--overwrite", action="store_true", help="Replace an existing key (IT must upload the new certificate)")
    keygen.add_argument("--show", action="store_true", help="Only show the current certificate's path and thumbprint")


def outlook_command(args: argparse.Namespace) -> int:
    command = getattr(args, "outlook_command", None) or "test"
    try:
        settings = load_settings()
        if command == "keygen":
            return _keygen(args, settings)
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


def _keygen(args: argparse.Namespace, settings) -> int:
    directory = cert_dir(settings)
    if getattr(args, "show", False):
        if not certauth.available(directory):
            print(f"No certificate in {directory}: run `alfred outlook keygen`.", file=sys.stderr)
            return 1
        cert = directory / certauth.CERT_FILE
        print(f"certificate: {cert}\nthumbprint:  {certauth.thumbprint(cert)}")
        return 0
    info = certauth.generate(directory, args.name, overwrite=args.overwrite)
    print(f"private key: {info['key']} (stays on this machine, mode 600)")
    print(f"certificate: {info['certificate']}  <- send THIS file to IT (it is public)")
    print(f"thumbprint:  {info['thumbprint']} (Entra shows it after the upload)")
    print(f"expires:     {info['expires']} — renew with `alfred outlook keygen --overwrite` before then")
    return 0
