"""
Single shared-password auth.

Scope: one shop owner, one password. No user accounts (there's one inventory,
so per-user logins would be theatre).

Design notes:
- Password comes from the APP_PASSWORD env var. Never hardcoded, never committed.
- Login sets a signed, HttpOnly session cookie. The cookie holds no secret —
  just an expiry and a signature, so it can't be forged without SECRET_KEY.
- Password comparison uses compare_digest to avoid timing leaks.
"""
import hashlib
import hmac
import os
import secrets
import time

from fastapi import HTTPException, Request

COOKIE_NAME = "pf_session"
SESSION_MAX_AGE = 60 * 60 * 24 * 14  # 14 days

# Failing loudly beats defaulting to a guessable password.
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

# If no SECRET_KEY is set, generate one at startup. Sessions then reset on
# restart (acceptable for a prototype); set it explicitly in production.
SECRET_KEY = os.environ.get("SECRET_KEY") or secrets.token_hex(32)


def auth_configured() -> bool:
    return bool(APP_PASSWORD)


def _sign(payload: str) -> str:
    return hmac.new(SECRET_KEY.encode(), payload.encode(), hashlib.sha256).hexdigest()


def check_password(candidate: str) -> bool:
    if not APP_PASSWORD:
        return False
    return hmac.compare_digest(candidate.encode(), APP_PASSWORD.encode())


def make_session() -> str:
    """Create a signed cookie value: '<expiry>.<signature>'."""
    expiry = str(int(time.time()) + SESSION_MAX_AGE)
    return f"{expiry}.{_sign(expiry)}"


def valid_session(cookie: str | None) -> bool:
    if not cookie or "." not in cookie:
        return False
    expiry, sig = cookie.rsplit(".", 1)
    if not hmac.compare_digest(sig, _sign(expiry)):
        return False
    try:
        return int(expiry) > time.time()
    except ValueError:
        return False


def require_auth(request: Request) -> None:
    """FastAPI dependency. Raises 401 if the request isn't authenticated."""
    if not auth_configured():
        raise HTTPException(
            503,
            "Login is not configured on the server (APP_PASSWORD is not set).",
        )
    if not valid_session(request.cookies.get(COOKIE_NAME)):
        raise HTTPException(401, "Not signed in.")
