"""
nexus.py — Async HTTP client for the NEXUS API.

All requests are authenticated with X-API-Key.
Callers receive plain dicts (parsed JSON) or None on 404.
Any other non-2xx response raises NexusError.

Usage:
    async with NexusClient() as client:
        membership = await client.get_membership_by_email(tournament_id=1, email="a@b.com")
"""

import aiohttp

import config


class NexusError(Exception):
    """Raised when NEXUS returns an unexpected error status."""

    def __init__(self, status: int, message: str) -> None:
        super().__init__(f"NEXUS {status}: {message}")
        self.status = status


class NexusClient:
    """Thin async wrapper around the NEXUS REST API.

    Can be used as an async context manager or instantiated manually and
    closed with await client.close().
    """

    def __init__(self) -> None:
        self._session: aiohttp.ClientSession | None = None

    # ── Session lifecycle ─────────────────────────────────────────────────────

    async def _get_session(self) -> aiohttp.ClientSession:
        if self._session is None or self._session.closed:
            self._session = aiohttp.ClientSession(
                base_url=config.NEXUS_API_URL,
                headers={"X-API-Key": config.NEXUS_API_KEY},
            )
        return self._session

    async def close(self) -> None:
        if self._session and not self._session.closed:
            await self._session.close()

    async def __aenter__(self) -> "NexusClient":
        return self

    async def __aexit__(self, *_) -> None:
        await self.close()

    # ── Internal request helper ───────────────────────────────────────────────

    async def _request(self, method: str, path: str, **kwargs) -> dict | list | None:
        """Make a request and return parsed JSON.

        Returns None on 404. Raises NexusError on other non-2xx responses.
        """
        session = await self._get_session()
        async with session.request(method, path, **kwargs) as resp:
            if resp.status == 404:
                return None
            if not resp.ok:
                text = await resp.text()
                raise NexusError(resp.status, text)
            return await resp.json()

    # ── Tournament endpoints ──────────────────────────────────────────────────

    async def get_tournament_by_guild(self, guild_id: int) -> dict | None:
        """GET /tournaments/by-guild/{guild_id}

        Returns the tournament dict or None if not found.
        NOTE: This endpoint does not exist yet in NEXUS — see issue-tournament-discord-guild-id.md.
        """
        return await self._request("GET", f"/tournaments/by-guild/{guild_id}")

    # ── Membership endpoints ──────────────────────────────────────────────────

    async def list_memberships(self, tournament_id: int) -> list[dict]:
        """GET /tournaments/{id}/memberships/"""
        result = await self._request("GET", f"/tournaments/{tournament_id}/memberships/")
        return result or []

    async def get_membership_by_email(self, tournament_id: int, email: str) -> dict | None:
        """GET /tournaments/{id}/memberships/?email=...

        Returns the first matching membership dict or None if not found.
        """
        result = await self._request(
            "GET",
            f"/tournaments/{tournament_id}/memberships/",
            params={"email": email},
        )
        if not result:
            return None
        # API returns a list; treat empty list as not-found
        if isinstance(result, list):
            return result[0] if result else None
        return result

    async def update_membership(
        self, tournament_id: int, membership_id: int, payload: dict
    ) -> dict | None:
        """PATCH /tournaments/{id}/memberships/{id}/

        Returns the updated membership dict.
        """
        return await self._request(
            "PATCH",
            f"/tournaments/{tournament_id}/memberships/{membership_id}/",
            json=payload,
        )

    # ── Event endpoints ───────────────────────────────────────────────────────

    async def list_events(self, tournament_id: int) -> list[dict]:
        """GET /tournaments/{id}/events/"""
        result = await self._request("GET", f"/tournaments/{tournament_id}/events/")
        return result or []

    async def get_event(self, tournament_id: int, event_id: int) -> dict | None:
        """GET /tournaments/{id}/events/{id}/"""
        return await self._request("GET", f"/tournaments/{tournament_id}/events/{event_id}/")