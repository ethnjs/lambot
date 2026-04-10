"""
data_router.py — NEXUS-first, Google Sheets fallback data access.

Cogs call this instead of touching NexusClient or SheetsClient directly.
Every public function resolves its data source transparently:
  1. Try NEXUS API
  2. On 404 / unavailable → fall back to the guild's connected Google Sheet

Guild → tournament resolution uses GUILD_TOURNAMENT_MAP from config until
GET /tournaments/by-guild/{guild_id} is live in NEXUS.

# TODO(sheets-fallback): remove sheet fallback paths as NEXUS endpoints mature.
"""

from __future__ import annotations

import config
from clients.nexus import NexusClient, NexusError


# ── Guild → Tournament resolution ────────────────────────────────────────────

async def resolve_tournament_id(guild_id: int) -> int | None:
    """Return the NEXUS tournament ID for a Discord guild.

    Tries the NEXUS endpoint first; falls back to the local map in config.
    Returns None if the guild has no known tournament.
    """
    # Try NEXUS (endpoint may not exist yet — returns None on 404)
    try:
        async with NexusClient() as client:
            tournament = await client.get_tournament_by_guild(guild_id)
        if tournament:
            return tournament["id"]
    except NexusError:
        pass
    except Exception as e:
        print(f"data_router: NEXUS guild lookup failed, using local map: {e}")

    # Local fallback
    return config.GUILD_TOURNAMENT_MAP.get(str(guild_id))


# ── Volunteer lookup ──────────────────────────────────────────────────────────

async def get_volunteer_by_email(
    guild_id: int,
    email: str,
    *,
    spreadsheets: dict,
    sheet_page_name: str = config.SHEET_PAGE_NAME,
) -> dict | None:
    """Look up a volunteer by email address.

    Returns a normalised dict:
        {
            "email": str,
            "first_name": str,
            "last_name": str,
            "status": str,          # "confirmed" | "interested" | ...
            "positions": list[str],
            "assigned_event_id": int | None,
            "source": "nexus" | "sheets",
        }
    or None if not found in either source.
    """
    tournament_id = await resolve_tournament_id(guild_id)

    # ── NEXUS path ────────────────────────────────────────────────────────────
    if tournament_id is not None:
        try:
            async with NexusClient() as client:
                membership = await client.get_membership_by_email(tournament_id, email)
            if membership:
                return _normalise_membership(membership)
        except NexusError as e:
            print(f"data_router: NEXUS membership lookup failed ({e}), falling back to sheets")
        except Exception as e:
            print(f"data_router: unexpected NEXUS error ({e}), falling back to sheets")

    # ── Sheets fallback ───────────────────────────────────────────────────────
    # TODO(sheets-fallback): remove when NEXUS membership lookup is stable
    return _sheets_lookup_by_email(guild_id, email, spreadsheets, sheet_page_name)


async def update_volunteer_status(
    guild_id: int,
    membership_id: int,
    payload: dict,
) -> dict | None:
    """PATCH a membership record in NEXUS.

    Returns the updated membership dict or None if the tournament/membership
    can't be resolved.
    """
    tournament_id = await resolve_tournament_id(guild_id)
    if tournament_id is None:
        print(f"data_router: no tournament ID for guild {guild_id}, cannot update membership")
        return None

    try:
        async with NexusClient() as client:
            return await client.update_membership(tournament_id, membership_id, payload)
    except NexusError as e:
        print(f"data_router: NEXUS membership update failed: {e}")
        return None


# ── Event listing ─────────────────────────────────────────────────────────────

async def list_events(
    guild_id: int,
    *,
    spreadsheets: dict,
) -> list[dict]:
    """Return events for a tournament.

    NEXUS-first; falls back to reading the Room Assignments sheet.
    # TODO(sheets-fallback): remove sheet fallback when NEXUS events are stable
    """
    tournament_id = await resolve_tournament_id(guild_id)

    if tournament_id is not None:
        try:
            async with NexusClient() as client:
                events = await client.list_events(tournament_id)
            if events:
                return events
        except NexusError as e:
            print(f"data_router: NEXUS events lookup failed ({e}), falling back to sheets")
        except Exception as e:
            print(f"data_router: unexpected NEXUS error ({e}), falling back to sheets")

    # TODO(sheets-fallback): remove when NEXUS supports this
    return _sheets_list_events(guild_id, spreadsheets)


# ── Internal helpers ──────────────────────────────────────────────────────────

def _normalise_membership(m: dict) -> dict:
    """Map a raw NEXUS membership dict to the shared volunteer schema."""
    user = m.get("user") or {}
    return {
        "email": user.get("email") or m.get("email", ""),
        "first_name": user.get("first_name", ""),
        "last_name": user.get("last_name", ""),
        "status": m.get("status", ""),
        "positions": m.get("positions", []),
        "assigned_event_id": m.get("assigned_event_id"),
        "nexus_membership_id": m.get("id"),
        "source": "nexus",
    }


def _sheets_lookup_by_email(
    guild_id: int,
    email: str,
    spreadsheets: dict,
    sheet_page_name: str,
) -> dict | None:
    """Read the main lambot worksheet and return a normalised volunteer dict."""
    ss = spreadsheets.get(guild_id)
    if not ss:
        return None
    try:
        data = ss.worksheet(sheet_page_name).get_all_records()
    except Exception as e:
        print(f"data_router: sheets lookup error for guild {guild_id}: {e}")
        return None

    email_lower = email.strip().lower()
    for row in data:
        if str(row.get("Email", "")).strip().lower() == email_lower:
            name_parts = str(row.get("Name", "")).strip().split(None, 1)
            roles_raw = str(row.get("Roles", "")).strip()
            positions = [r.strip() for r in roles_raw.split(";") if r.strip()]
            return {
                "email": email,
                "first_name": name_parts[0] if name_parts else "",
                "last_name": name_parts[1] if len(name_parts) > 1 else "",
                "status": str(row.get("Status", "")).strip().lower() or "confirmed",
                "positions": positions,
                "assigned_event_id": None,
                "nexus_membership_id": None,
                "source": "sheets",
            }
    return None


def _sheets_list_events(guild_id: int, spreadsheets: dict) -> list[dict]:
    """Read the Room Assignments worksheet and return a list of event dicts."""
    ss = spreadsheets.get(guild_id)
    if not ss:
        return []
    try:
        rows = ss.worksheet("Room Assignments").get_all_records()
    except Exception as e:
        print(f"data_router: sheets events lookup error for guild {guild_id}: {e}")
        return []

    seen = set()
    events = []
    for row in rows:
        name = str(row.get("Events", "")).strip()
        if name and name not in seen:
            seen.add(name)
            events.append({
                "name": name,
                "building": str(row.get("Building", "")).strip(),
                "room": str(row.get("Room", "")).strip(),
                "source": "sheets",
            })
    return events
