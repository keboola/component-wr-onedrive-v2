"""Interactive OAuth authorization helper for wr-onedrive-v2 test credentials.

Runs the Microsoft identity platform authorization-code flow against a local
redirect URI, exchanges the code for tokens, and merges the resulting
``refresh_token`` (plus the app credentials you entered) into ``secrets.json``
at the repo root — the file the VCR recording script reads.

Usage:
    uv run python scripts/authorize_oauth.py

Prerequisite: the Azure AD app you use must list the local redirect URI
(``http://localhost:53682/callback``) under Authentication → Platform
configurations → Web (or Mobile/Desktop). The production wr-onedrive app only
registers the Keboola OAuth broker's callback, so either add the localhost URI
to it temporarily or use a test app registration in the test tenant.

No secret value is printed; everything lands only in the gitignored
``secrets.json``.
"""

import base64
import hashlib
import json
import secrets as pysecrets
import sys
import threading
import urllib.parse
import webbrowser
from getpass import getpass
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

import requests

REDIRECT_PORT = 53682
REDIRECT_URI = f"http://localhost:{REDIRECT_PORT}/callback"
SCOPES = "offline_access User.Read Files.ReadWrite.All Sites.ReadWrite.All"
SECRETS_PATH = Path(__file__).resolve().parent.parent / "secrets.json"


class _CallbackHandler(BaseHTTPRequestHandler):
    """Catches the single OAuth redirect and stores the auth code on the server."""

    def do_GET(self):
        query = urllib.parse.urlparse(self.path).query
        params = urllib.parse.parse_qs(query)
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        if "code" in params and params.get("state", [""])[0] == self.server.expected_state:
            self.server.auth_code = params["code"][0]
            self.wfile.write(b"<h2>Authorization complete.</h2>You can close this tab.")
        else:
            self.server.auth_error = params.get("error_description", params.get("error", ["unknown"]))[0]
            self.wfile.write(b"<h2>Authorization failed.</h2>See the terminal for details.")

    def log_message(self, *args):  # silence request logging
        pass


def _prompt(existing: dict) -> dict:
    def ask(key: str, label: str, secret: bool = False) -> str:
        current = existing.get(key, "")
        hint = " [keep existing]" if current else ""
        value = (getpass if secret else input)(f"{label}{hint}: ").strip()
        return value or current

    values = {
        "appKey": ask("appKey", "Azure app client id (appKey)"),
        "#appSecret": ask("#appSecret", "Azure app client secret (#appSecret)", secret=True),
        "tenant_id": ask("tenant_id", "Tenant id (empty = 'common')"),
        "site_url": ask("site_url", "SharePoint site URL (for recording)"),
    }
    missing = [k for k in ("appKey", "#appSecret") if not values[k]]
    if missing:
        sys.exit(f"Required value(s) missing: {', '.join(missing)}")
    return values


def main() -> None:
    existing = json.loads(SECRETS_PATH.read_text()) if SECRETS_PATH.exists() else {}
    values = _prompt(existing)
    authority = values["tenant_id"] or "common"

    state = pysecrets.token_urlsafe(16)
    verifier = pysecrets.token_urlsafe(64)
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()

    authorize_url = (
        f"https://login.microsoftonline.com/{authority}/oauth2/v2.0/authorize?"
        + urllib.parse.urlencode(
            {
                "client_id": values["appKey"],
                "response_type": "code",
                "redirect_uri": REDIRECT_URI,
                "response_mode": "query",
                "scope": SCOPES,
                "state": state,
                "code_challenge": challenge,
                "code_challenge_method": "S256",
                "prompt": "select_account",
            }
        )
    )

    server = HTTPServer(("127.0.0.1", REDIRECT_PORT), _CallbackHandler)
    server.expected_state = state
    server.auth_code = None
    server.auth_error = None
    threading.Thread(target=server.serve_forever, daemon=True).start()

    print(f"\nOpening browser for consent (redirect: {REDIRECT_URI})…")
    print("If no browser opens, visit this URL manually:\n")
    print(authorize_url + "\n")
    webbrowser.open(authorize_url)

    print("Waiting for the redirect (Ctrl+C to abort)…")
    try:
        while server.auth_code is None and server.auth_error is None:
            threading.Event().wait(0.2)
    except KeyboardInterrupt:
        sys.exit("Aborted.")
    finally:
        server.shutdown()

    if server.auth_error:
        sys.exit(f"Authorization failed: {server.auth_error}")

    print("Code received — exchanging for tokens…")
    response = requests.post(
        f"https://login.microsoftonline.com/{authority}/oauth2/v2.0/token",
        data={
            "client_id": values["appKey"],
            "client_secret": values["#appSecret"],
            "grant_type": "authorization_code",
            "code": server.auth_code,
            "redirect_uri": REDIRECT_URI,
            "code_verifier": verifier,
            "scope": SCOPES,
        },
        timeout=30,
    )
    payload = response.json()
    if response.status_code != 200 or "refresh_token" not in payload:
        sys.exit(
            f"Token exchange failed ({response.status_code}): "
            f"{payload.get('error')}: {payload.get('error_description', '')[:300]}"
        )

    merged = {**existing, **values, "refresh_token": payload["refresh_token"]}
    SECRETS_PATH.write_text(json.dumps(merged, indent=2) + "\n")
    granted = payload.get("scope", "?")
    print(f"\nDone. refresh_token written to {SECRETS_PATH.name} (granted scopes: {granted}).")
    print("You can now run: uv run python scripts/record_vcr_cassettes.py")


if __name__ == "__main__":
    main()
