"""
Usage:
  python3 guilds.py              -- list all guilds
  python3 guilds.py leave <id>   -- leave a guild and clean up its cache
"""

import os
import sys
import json
import discord
from dotenv import load_dotenv

load_dotenv()
TOKEN = os.getenv("DISCORD_TOKEN")
CACHE_FILE = "bot_cache.json"

# ── helpers ──────────────────────────────────────────────────────────────────

def _load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE) as f:
            return json.load(f)
    return {}

def _save_cache(data):
    with open(CACHE_FILE, "w") as f:
        json.dump(data, f, indent=2)

def _purge_guild_from_cache(guild_id: int):
    """Remove every trace of guild_id from bot_cache.json."""
    cache = _load_cache()
    changed = False

    gid = str(guild_id)

    if gid in cache.get("guilds", {}):
        del cache["guilds"][gid]
        changed = True
        print(f"  removed from cache['guilds']")

    runner = cache.get("runner_access_settings", {})
    if gid in runner:
        del runner[gid]
        cache["runner_access_settings"] = runner
        changed = True
        print(f"  removed from cache['runner_access_settings']")

    if changed:
        _save_cache(cache)
        print(f"  cache saved")
    else:
        print(f"  guild {guild_id} was not in the cache file")

# ── commands ─────────────────────────────────────────────────────────────────

def cmd_list():
    intents = discord.Intents.default()
    client = discord.Client(intents=intents)

    @client.event
    async def on_ready():
        guilds = client.guilds
        print(f"\nLambot is in {len(guilds)} server(s):\n")
        print(f"{'ID':<22} {'Joined':<28} {'Name'}")
        print("-" * 75)
        for guild in sorted(guilds, key=lambda g: g.name.lower()):
            joined = guild.me.joined_at.strftime("%Y-%m-%d %H:%M UTC") if guild.me.joined_at else "unknown"
            print(f"{guild.id:<22} {joined:<28} {guild.name}")
        print()
        await client.close()

    client.run(TOKEN)


def cmd_leave(guild_id: int):
    intents = discord.Intents.default()
    client = discord.Client(intents=intents)

    @client.event
    async def on_ready():
        guild = client.get_guild(guild_id)
        if guild:
            print(f"Found guild: '{guild.name}' ({guild.id})")
            await guild.leave()
            print(f"  left guild '{guild.name}'")
        else:
            print(f"Bot is not in guild {guild_id} (already left or bad ID)")

        print(f"Cleaning up cache for guild {guild_id}...")
        _purge_guild_from_cache(guild_id)

        await client.close()

    client.run(TOKEN)

# ── entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    args = sys.argv[1:]

    if not args:
        cmd_list()
    elif args[0] == "leave" and len(args) == 2:
        try:
            gid = int(args[1])
        except ValueError:
            print(f"Error: '{args[1]}' is not a valid guild ID")
            sys.exit(1)
        cmd_leave(gid)
    else:
        print(__doc__)
        sys.exit(1)
