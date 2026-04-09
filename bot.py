"""
bot.py — entry point for the refactored lambot.

Loads cogs, syncs slash commands, and runs the bot.
Business logic lives in cogs/, not here.
"""

import asyncio
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import discord
from discord.ext import commands

import config


# ── Bot class ─────────────────────────────────────────────────────────────────

COGS: list[str] = [
    # Cog module paths are added here as they are extracted from lam_bot.py.
    # e.g. "cogs.admin", "cogs.onboarding", "cogs.tickets"
]

class LamBot(commands.Bot):
    def __init__(self) -> None:
        intents = discord.Intents.default()
        intents.members = True
        intents.message_content = True
        super().__init__(command_prefix="!", intents=intents)

    async def setup_hook(self) -> None:
        """Load all cogs before the bot connects."""
        for cog in COGS:
            await self.load_extension(cog)
            print(f"  Loaded cog: {cog}")

        # Sync slash commands globally.
        await self.tree.sync()
        print("Slash commands synced.")

    async def on_ready(self) -> None:
        print(f"Logged in as {self.user} (ID: {self.user.id})")
        print(f"Active in {len(self.guilds)} guild(s):")
        for guild in self.guilds:
            print(f"  • {guild.name} (ID: {guild.id})")


# ── Health check server (Railway / Fly.io keepalive) ─────────────────────────

class _HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        self.send_response(200)
        self.send_header("Content-type", "text/plain")
        self.end_headers()
        self.wfile.write(b"OK")

    def log_message(self, format: str, *args) -> None:  # noqa: A002
        pass  # suppress HTTP access logs


def _start_health_server(port: int = 8080) -> None:
    server = HTTPServer(("0.0.0.0", port), _HealthHandler)
    print(f"Health check server listening on port {port}")
    server.serve_forever()


# ── Entry point ───────────────────────────────────────────────────────────────

async def main() -> None:
    bot = LamBot()
    async with bot:
        await bot.start(config.DISCORD_BOT_TOKEN)


if __name__ == "__main__":
    threading.Thread(target=_start_health_server, daemon=True).start()
    asyncio.run(main())
