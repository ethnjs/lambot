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
from clients.sheets import SheetsClient


# ── Bot class ─────────────────────────────────────────────────────────────────

COGS: list[str] = [
    "cogs.onboarding",
    "cogs.tickets",
]

class LamBot(commands.Bot):
    def __init__(self) -> None:
        intents = discord.Intents.default()
        intents.members = True
        intents.message_content = True
        super().__init__(command_prefix="!", intents=intents)

        # Shared state — populated by the sheets setup flow (enterfolder / cache load).
        # Cogs read and write these rather than keeping their own copies.
        self.spreadsheets: dict = {}   # guild_id (int) -> gspread.Spreadsheet
        self.sheets: dict = {}         # guild_id (int) -> gspread.Worksheet (main sheet)
        self.runner_all_access: dict = {}  # guild_id (int) -> int flag
        self.sheets_client: SheetsClient | None = None

    async def setup_hook(self) -> None:
        """Load all cogs before the bot connects."""
        # Initialise sheets client and restore any cached guild connections.
        try:
            self.sheets_client = SheetsClient()
            self.spreadsheets = self.sheets_client.load_spreadsheets_from_cache()
            self.runner_all_access = self.sheets_client.load_runner_access_from_cache()
            print(f"Sheets client ready — {len(self.spreadsheets)} cached connection(s) restored.")
        except Exception as e:
            print(f"Sheets client unavailable (no secrets/gspread.json?): {e}")

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
