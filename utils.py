"""
utils.py — Shared Discord helpers used across multiple cogs.
"""

import asyncio

import discord
from discord.ext import commands

import config


async def handle_rate_limit(coro, operation_name: str, max_retries: int = 3, default_delay: float = 0.1):
    """Execute a coroutine, retrying automatically on Discord 429 rate limits.

    Returns the coroutine result on success, or None if all retries are exhausted.
    Re-raises any non-rate-limit exception immediately.
    """
    retry_count = 0
    while retry_count < max_retries:
        try:
            result = await coro
            await asyncio.sleep(default_delay)
            return result
        except discord.HTTPException as e:
            msg = str(e)
            is_rate_limit = e.status == 429 or "429" in msg or "rate limit" in msg.lower() or "too many requests" in msg.lower()
            if not is_rate_limit:
                raise
            retry_count += 1
            if retry_count >= max_retries:
                print(f"Rate limited on '{operation_name}' after {max_retries} retries, giving up")
                return None
            retry_after = float(getattr(e, "retry_after", None) or 1.0)
            print(f"Rate limited on '{operation_name}', retrying in {retry_after}s ({retry_count}/{max_retries})")
            await asyncio.sleep(retry_after)
        except Exception as e:
            msg = str(e)
            if "429" in msg or "rate limit" in msg.lower() or "too many requests" in msg.lower():
                retry_count += 1
                if retry_count >= max_retries:
                    print(f"Rate limited on '{operation_name}' after {max_retries} retries, giving up")
                    return None
                print(f"Rate limited on '{operation_name}', retrying in 1s ({retry_count}/{max_retries})")
                await asyncio.sleep(1.0)
            else:
                raise
    return None


async def get_or_create_role(guild: discord.Guild, role_name: str) -> discord.Role | None:
    """Return a guild role by name, creating it if it doesn't exist.

    Returns None if auto-creation is disabled (AUTO_CREATE_ROLES=false) and
    the role doesn't already exist.
    """
    role = discord.utils.get(guild.roles, name=role_name)
    if role:
        return role

    if not config.AUTO_CREATE_ROLES:
        print(f"Role '{role_name}' not found and AUTO_CREATE_ROLES is disabled")
        return None

    # Admin role gets full permissions and a distinct colour
    if role_name == "Admin":
        return await handle_rate_limit(
            guild.create_role(
                name="Admin",
                permissions=discord.Permissions.all(),
                color=discord.Color.purple(),
                reason="Auto-created Admin role",
            ),
            f"creating Admin role in {guild.name}",
        )

    color = _color_from_name(config.DEFAULT_ROLE_COLOR)
    return await handle_rate_limit(
        guild.create_role(name=role_name, color=color, reason="Auto-created by lambot"),
        f"creating role '{role_name}' in {guild.name}",
    )


def _color_from_name(name: str) -> discord.Color:
    mapping = {
        "blue": discord.Color.blue(),
        "red": discord.Color.red(),
        "green": discord.Color.green(),
        "purple": discord.Color.purple(),
        "gold": discord.Color.gold(),
        "orange": discord.Color.orange(),
        "teal": discord.Color.teal(),
        "light_gray": discord.Color.light_gray(),
        "dark_gray": discord.Color.dark_gray(),
    }
    return mapping.get(name.lower(), discord.Color.light_gray())
