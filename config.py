"""
Central configuration — env vars, constants, and static mappings.
Import this instead of reaching for os.getenv() in individual modules.
"""

import os
from dotenv import load_dotenv

load_dotenv()

# ── Discord ───────────────────────────────────────────────────────────────────

DISCORD_BOT_TOKEN: str = os.getenv("DISCORD_BOT_TOKEN", "")

# ── NEXUS API ─────────────────────────────────────────────────────────────────

NEXUS_API_URL: str = os.getenv("NEXUS_API_URL", "http://localhost:8001")
NEXUS_API_KEY: str = os.getenv("NEXUS_API_KEY", "")

# ── Google Sheets (legacy pipeline) ──────────────────────────────────────────

def _service_email() -> str:
    """Read service account email from env var, JSON credentials, or file fallback."""
    explicit = os.getenv("SERVICE_EMAIL", "")
    if explicit:
        return explicit
    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
    if raw:
        try:
            import json
            return json.loads(raw).get("client_email", "")
        except Exception:
            pass
    try:
        import json
        with open("secrets/gspread.json") as f:
            return json.load(f).get("client_email", "")
    except Exception:
        return ""

SERVICE_EMAIL: str = _service_email()
SHEET_ID: str = os.getenv("SHEET_ID", "")
SHEET_PAGE_NAME: str = os.getenv("SHEET_PAGE_NAME", "lambot")

# ── Bot behaviour ─────────────────────────────────────────────────────────────

AUTO_CREATE_ROLES: bool = os.getenv("AUTO_CREATE_ROLES", "true").lower() == "true"
DEFAULT_ROLE_COLOR: str = os.getenv("DEFAULT_ROLE_COLOR", "light_gray")
RESET_SERVER: bool = os.getenv("RESET_SERVER", "false").lower() == "true"

# ── Cache ─────────────────────────────────────────────────────────────────────

CACHE_FILE: str = "bot_cache.json"

# ── Guild → Tournament mapping ────────────────────────────────────────────────
# Used until GET /tournaments/by-guild/{guild_id} exists in NEXUS.
# Key: Discord guild ID (str), Value: NEXUS tournament ID (int).
# TODO: replace with a live API call once the NEXUS endpoint is available.
GUILD_TOURNAMENT_MAP: dict[str, int] = {
    # "123456789": 1,  # example: The Pasture (staging) → tournament ID 1
}
