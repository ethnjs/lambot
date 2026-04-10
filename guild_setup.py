"""
guild_setup.py — Discord guild structural helpers.

These functions create and organize the standard channel/category/role
layout for a tournament guild.  They are called from cogs (admin,
on_ready equivalent in bot.py) and accept the guild and any required
shared state as parameters rather than reading globals.

Functions:
  sanitize_for_discord
  get_or_create_category
  get_or_create_channel
  sort_building_categories_alphabetically
  sort_channels_in_building_categories
  setup_building_structure
  setup_chapter_structure
  sort_chapter_channels_alphabetically
  add_runner_access
  ensure_runner_tournament_officials_access
  send_building_welcome_message
  add_role_to_building_chat
  reset_server_for_guild
  post_welcome_instructions
  post_welcome_tldr
  setup_static_channels_for_guild
  move_bot_role_to_top_for_guild
  organize_role_hierarchy_for_guild
  remove_runner_access_from_building_channels_for_guild
  give_runner_access_to_all_channels_for_guild
  setup_ezhang_admin_role
  generate_building_structures
  get_building_events
"""

from __future__ import annotations

import asyncio

import discord

import data_router
from utils import get_or_create_role, handle_rate_limit

# Role names that are treated as structural/priority rather than event roles.
PRIORITY_ROLES = frozenset([
    "Admin", "Volunteer", "Lead ES", "Social Media",
    "Photographer", "Arbitrations", "Awards", "Runner", "VIPer",
])

# Category names that are always kept at the top and never sorted as building categories.
STATIC_CATEGORIES = ["Welcome", "Tournament Officials", "Chapters", "Volunteers"]


# ── String helpers ────────────────────────────────────────────────────────────

def sanitize_for_discord(text: str) -> str:
    """Return a Discord-safe channel name derived from *text*."""
    return (
        text.lower()
        .replace(" ", "-")
        .replace("/", "-")
        .replace("\\", "-")
        .replace(":", "-")
        .replace("*", "-")
        .replace("?", "-")
        .replace('"', "")
        .replace("<", "")
        .replace(">", "")
        .replace("|", "-")
    )


# ── Category / channel primitives ─────────────────────────────────────────────

async def get_or_create_category(guild: discord.Guild, category_name: str) -> discord.CategoryChannel | None:
    """Return a guild category by name, creating it if absent."""
    category = discord.utils.get(guild.categories, name=category_name)
    if category:
        return category
    try:
        return await handle_rate_limit(
            guild.create_category(name=category_name, reason="Auto-created by LAM Bot"),
            f"creating category '{category_name}'",
        )
    except discord.Forbidden:
        print(f"No permission to create category '{category_name}'")
        return None
    except Exception as e:
        print(f"Error creating category '{category_name}': {e}")
        return None


async def get_or_create_channel(
    guild: discord.Guild,
    channel_name: str,
    category: discord.CategoryChannel | None,
    event_role: discord.Role | None = None,
    is_building_chat: bool = False,
    runner_access: dict | None = None,
) -> discord.TextChannel | None:
    """Return a text channel by name, creating it if absent.

    *runner_access* is the bot's ``runner_all_access`` dict (guild_id → flag).
    """
    runner_access = runner_access or {}

    existing = discord.utils.get(guild.text_channels, name=channel_name)
    if existing:
        return existing

    try:
        overwrites: dict = {}

        runner_role = discord.utils.get(guild.roles, name="Runner")
        guild_runner_access = runner_access.get(guild.id, 0)
        in_static = category and category.name in STATIC_CATEGORIES

        if runner_role and (guild_runner_access or in_static):
            overwrites[runner_role] = discord.PermissionOverwrite(
                read_messages=True,
                send_messages=True,
                read_message_history=True,
            )

        if event_role:
            overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)
            overwrites[event_role] = discord.PermissionOverwrite(
                read_messages=True,
                send_messages=True,
                read_message_history=True,
            )
        elif is_building_chat:
            overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)

        return await handle_rate_limit(
            guild.create_text_channel(
                name=channel_name,
                category=category,
                overwrites=overwrites,
                reason="Auto-created by LAM Bot",
            ),
            f"creating channel '{channel_name}'",
        )
    except discord.Forbidden:
        print(f"No permission to create channel '{channel_name}'")
        return None
    except Exception as e:
        print(f"Error creating channel '{channel_name}': {e}")
        return None


# ── Sorting helpers ───────────────────────────────────────────────────────────

async def sort_building_categories_alphabetically(guild: discord.Guild) -> None:
    """Place static categories first (in fixed order), then building categories alphabetically."""
    try:
        static: list[discord.CategoryChannel] = []
        building: list[discord.CategoryChannel] = []

        for cat in guild.categories:
            if cat.name in STATIC_CATEGORIES:
                static.append(cat)
            else:
                building.append(cat)

        building.sort(key=lambda c: c.name.lower())

        # Sort static categories by desired display order
        desired = ["Welcome", "Tournament Officials", "Chapters", "Volunteers"]
        ordered_static = []
        for name in desired:
            for cat in static:
                if cat.name == name:
                    ordered_static.append(cat)
                    break
        for cat in static:
            if cat not in ordered_static:
                ordered_static.append(cat)

        position = 0
        for cat in ordered_static + building:
            if cat.position != position:
                await handle_rate_limit(
                    cat.edit(position=position, reason="Organizing categories"),
                    f"moving category '{cat.name}'",
                )
            position += 1

        print("Categories organized: static first, then buildings alphabetically")
    except Exception as e:
        print(f"Error organizing categories: {e}")


async def sort_channels_in_building_categories(guild: discord.Guild) -> None:
    """Within each building category, keep the building-chat first then sort event channels."""
    try:
        for category in guild.categories:
            if category.name in STATIC_CATEGORIES:
                continue

            channels = category.text_channels
            if len(channels) <= 1:
                continue

            building_chat_name = f"{sanitize_for_discord(category.name)}-chat"
            building_chats = [c for c in channels if c.name == building_chat_name or c.name.endswith("-chat")]
            event_channels = [c for c in channels if c not in building_chats]
            event_channels.sort(key=lambda c: c.name.lower())

            for i, channel in enumerate(building_chats + event_channels):
                if channel.position != i:
                    try:
                        await handle_rate_limit(
                            channel.edit(position=i, reason=f"Sorting {category.name} channels"),
                            f"moving channel '{channel.name}'",
                        )
                    except Exception as e:
                        print(f"Error moving #{channel.name}: {e}")

        print("Finished sorting channels inside building categories")
    except Exception as e:
        print(f"Error sorting channels inside building categories: {e}")


async def sort_chapter_channels_alphabetically(guild: discord.Guild) -> None:
    """Sort chapter channels alphabetically, placing 'unaffiliated' last."""
    try:
        chapters_cat = discord.utils.get(guild.categories, name="Chapters")
        if not chapters_cat:
            return

        channels = chapters_cat.text_channels
        if len(channels) <= 1:
            return

        unaffiliated = [c for c in channels if c.name == "unaffiliated"]
        others = sorted([c for c in channels if c.name != "unaffiliated"], key=lambda c: c.name.lower())

        for i, channel in enumerate(others + unaffiliated):
            if channel.position != i:
                try:
                    await handle_rate_limit(
                        channel.edit(position=i, reason="Sorting chapter channels alphabetically"),
                        f"moving channel '{channel.name}'",
                    )
                except Exception as e:
                    print(f"Error moving #{channel.name}: {e}")

        print("Chapter channels sorted (unaffiliated at bottom)")
    except Exception as e:
        print(f"Error sorting chapter channels: {e}")


# ── Permission helpers ────────────────────────────────────────────────────────

async def add_runner_access(channel: discord.TextChannel, runner_role: discord.Role) -> None:
    """Grant *runner_role* read/send access to *channel*."""
    if not channel or not runner_role:
        return
    try:
        overwrites = channel.overwrites
        overwrites[runner_role] = discord.PermissionOverwrite(
            read_messages=True,
            send_messages=True,
            read_message_history=True,
        )
        await handle_rate_limit(
            channel.edit(overwrites=overwrites, reason=f"Added {runner_role.name} access"),
            f"editing channel '{channel.name}' permissions",
        )
    except discord.Forbidden:
        print(f"No permission to edit channel permissions for #{channel.name}")
    except Exception as e:
        print(f"Error updating permissions for #{channel.name}: {e}")


async def ensure_runner_tournament_officials_access(guild: discord.Guild, runner_role: discord.Role) -> None:
    """Make sure Runner has access to all Tournament Officials channels."""
    if not runner_role:
        return
    cat = discord.utils.get(guild.categories, name="Tournament Officials")
    if not cat:
        return
    for name in ["runner", "scoring", "awards-ceremony"]:
        ch = discord.utils.get(guild.text_channels, name=name)
        if ch and ch.category == cat:
            try:
                await add_runner_access(ch, runner_role)
            except Exception as e:
                print(f"Error adding Runner access to #{name}: {e}")


async def add_role_to_building_chat(channel: discord.TextChannel, role: discord.Role) -> None:
    """Grant *role* read/send access to a building-chat channel, hiding it from @everyone."""
    if not channel or not role:
        return
    try:
        overwrites = channel.overwrites
        overwrites[channel.guild.default_role] = discord.PermissionOverwrite(read_messages=False)
        overwrites[role] = discord.PermissionOverwrite(
            read_messages=True,
            send_messages=True,
            read_message_history=True,
        )
        await handle_rate_limit(
            channel.edit(overwrites=overwrites, reason=f"Added {role.name} to building chat"),
            f"editing building chat '{channel.name}' permissions",
        )
    except discord.Forbidden:
        print(f"No permission to edit channel permissions for #{channel.name}")
    except Exception as e:
        print(f"Error updating permissions for #{channel.name}: {e}")


# ── Data helpers ──────────────────────────────────────────────────────────────

async def get_building_events(guild_id: int, building: str, *, spreadsheets: dict) -> list[tuple[str, str]]:
    """Return a list of (event, room) tuples for every event in *building*."""
    try:
        room_data = await data_router.list_events(guild_id, spreadsheets=spreadsheets)
        result: list[tuple[str, str]] = []
        for row in room_data:
            row_building = str(row.get("building", "")).strip()
            if row_building.lower() != building.lower():
                continue
            event = str(row.get("name", "")).strip()
            room = str(row.get("room", "")).strip()
            if event and event not in PRIORITY_ROLES:
                combo = (event, room)
                if combo not in result:
                    result.append(combo)
        print(f"Found {len(result)} event(s) in building '{building}'")
        return result
    except Exception as e:
        print(f"Error looking up building events: {e}")
        return []


# ── Welcome messages ──────────────────────────────────────────────────────────

async def send_building_welcome_message(
    guild: discord.Guild,
    building_chat: discord.TextChannel,
    building: str,
    *,
    spreadsheets: dict,
) -> None:
    """Post (and pin) a welcome embed in a building-chat channel."""
    if not building_chat or not building:
        return
    try:
        building_events = await get_building_events(guild.id, building, spreadsheets=spreadsheets)
        if not building_events:
            print(f"No events found for building '{building}', skipping welcome message")
            return

        building_events.sort(key=lambda x: x[0].lower())

        embed = discord.Embed(
            title=f"Welcome to {building}!",
            description=f"This is the general chat for everyone with events in **{building}**.",
            color=discord.Color.blue(),
        )

        events_text = ""
        for event, room in building_events:
            events_text += f"• **{event}** - {room}\n" if room else f"• **{event}**\n"

        embed.add_field(name="Events in this building:", value=events_text, inline=False)
        embed.add_field(
            name="How to use this chat:",
            value=(
                "• Coordinate with other events in your building\n"
                "• Share building-specific information\n"
                "• Ask questions about the venue\n"
                "• Connect with nearby events"
            ),
            inline=False,
        )
        embed.set_footer(text="Each event also has its own dedicated channel for event-specific discussions.")

        msg = await building_chat.send(embed=embed)
        await asyncio.sleep(0.5)

        try:
            await msg.pin()
        except discord.Forbidden:
            print(f"Could not pin welcome message in #{building_chat.name}")
        except Exception as e:
            print(f"Error pinning welcome message in #{building_chat.name}: {e}")

        print(f"Sent welcome message to #{building_chat.name} for '{building}'")
    except Exception as e:
        print(f"Error sending welcome message to #{building_chat.name}: {e}")


async def post_welcome_instructions(
    welcome_channel: discord.TextChannel,
    bot_user: discord.ClientUser,
) -> None:
    """Post the full login instructions embed in *welcome_channel* (once)."""
    try:
        async for message in welcome_channel.history(limit=10):
            if message.author == bot_user and message.embeds:
                for embed in message.embeds:
                    if embed.title and "Welcome to the Science Olympiad Server" in embed.title:
                        print(f"Welcome instructions already posted in #{welcome_channel.name}")
                        return

        embed = discord.Embed(
            title="Welcome to the Science Olympiad Server!",
            description="Thank you for joining our Science Olympiad community! This server helps coordinate events, volunteers, and communication.",
            color=discord.Color.blue(),
        )
        embed.add_field(
            name="Getting Started - Login Required",
            value=(
                "**To access all channels and get your roles, you need to login:**\n\n"
                "1️⃣ Type `/login email:your@email.com password:yourpassword`\n"
                "2️⃣ Replace with your actual email address and the password you received in your volunteer info email\n"
                "3️⃣ Get instant access to your assigned channels!\n\n"
                "✅ You'll automatically receive:\n"
                "• Your assigned roles\n"
                "• Access to relevant channels\n"
                "• Your building and room information\n"
                "• Updated nickname with your event"
            ),
            inline=False,
        )
        embed.add_field(
            name="What You Can Do Right Now",
            value=(
                "Even before logging in, you can:\n"
                "• Read announcements in this channel\n"
                "• Browse volunteer channels for general info\n"
                "• Ask questions in the help forum\n"
                "• Start sobbing uncontrollably"
            ),
            inline=False,
        )
        embed.add_field(
            name="Need Help?",
            value=(
                "• **Can't find your email?** Contact an admin\n"
                "• **Questions about your assignment?** Ask in volunteer channels\n"
                "• **Technical problems?** Mention an admin or moderator"
            ),
            inline=False,
        )
        embed.add_field(
            name="Important Notes",
            value=(
                "• Your email must be in our system to login\n"
                "• Each email can only be linked to one Discord account\n"
                "• Your nickname will be updated to show your event\n"
                "• Channels will appear based on your assigned roles"
            ),
            inline=False,
        )
        embed.set_footer(text="Use /login to get started! • Questions? Ask in volunteer channels")

        await welcome_channel.send(embed=embed)
        print(f"Posted welcome instructions to #{welcome_channel.name}")
    except Exception as e:
        print(f"Error posting welcome instructions: {e}")


async def post_welcome_tldr(
    welcome_channel: discord.TextChannel,
    bot_user: discord.ClientUser,
) -> None:
    """Post the TLDR /login reminder embed in *welcome_channel* (once)."""
    try:
        async for message in welcome_channel.history(limit=10):
            if message.author == bot_user and message.embeds:
                for embed in message.embeds:
                    if embed.title and "TLDR: TYPE" in embed.title:
                        print(f"Welcome TLDR already posted in #{welcome_channel.name}")
                        return

        embed = discord.Embed(
            title="TLDR: TYPE `/login` TO GET STARTED",
            description="Read below message for more info",
            color=discord.Color.blue(),
        )
        await welcome_channel.send(embed=embed)
        print(f"Posted welcome TLDR to #{welcome_channel.name}")
    except Exception as e:
        print(f"Error posting welcome TLDR: {e}")


# ── Building / chapter structure ──────────────────────────────────────────────

async def setup_building_structure(
    guild: discord.Guild,
    building: str,
    first_event: str,
    room: str | None = None,
    *,
    runner_access: dict | None = None,
    spreadsheets: dict | None = None,
) -> None:
    """Create the category, building-chat, and event channel for one building/event pair."""
    if first_event and first_event in PRIORITY_ROLES:
        print(f"Skipping building structure for priority role '{first_event}' in {building}")
        return

    runner_access = runner_access or {}
    spreadsheets = spreadsheets or {}

    category = await get_or_create_category(guild, building)
    if not category:
        return

    building_chat_name = f"{sanitize_for_discord(building)}-chat"
    building_chat = await get_or_create_channel(
        guild, building_chat_name, category,
        is_building_chat=True,
        runner_access=runner_access,
    )

    if building_chat:
        try:
            messages = [m async for m in building_chat.history(limit=1)]
            if not messages:
                await send_building_welcome_message(guild, building_chat, building, spreadsheets=spreadsheets)
        except Exception as e:
            print(f"Error checking/sending welcome message for #{building_chat.name}: {e}")

    if first_event and first_event.lower() != "runner":
        event_role = await get_or_create_role(guild, first_event)
        if event_role and building_chat:
            await add_role_to_building_chat(building_chat, event_role)

            channel_name = (
                f"{sanitize_for_discord(first_event)}-{sanitize_for_discord(building)}-{sanitize_for_discord(room)}"
                if room
                else f"{sanitize_for_discord(first_event)}-{sanitize_for_discord(building)}"
            )
            await get_or_create_channel(
                guild, channel_name, category, event_role,
                runner_access=runner_access,
            )


async def setup_chapter_structure(
    guild: discord.Guild,
    chapter_name: str,
    chapter_role_names: set[str],
    runner_access: dict | None = None,
) -> None:
    """Create the chapter channel under the Chapters category and set its permissions."""
    chapter_role_names.add(chapter_name)
    runner_access = runner_access or {}

    chapters_cat = await get_or_create_category(guild, "Chapters")
    if not chapters_cat:
        return

    channel_name = sanitize_for_discord(chapter_name)
    chapter_channel = await get_or_create_channel(
        guild, channel_name, chapters_cat, runner_access=runner_access,
    )
    chapter_role = await get_or_create_role(guild, chapter_name)

    if chapter_channel and chapter_role:
        try:
            overwrites = chapter_channel.overwrites
            overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)
            overwrites[chapter_role] = discord.PermissionOverwrite(
                read_messages=True,
                send_messages=True,
                read_message_history=True,
            )
            await handle_rate_limit(
                chapter_channel.edit(overwrites=overwrites, reason=f"Set up {chapter_name} chapter permissions"),
                f"editing chapter channel '{channel_name}' permissions",
            )
            await sort_chapter_channels_alphabetically(guild)
        except Exception as e:
            print(f"Error setting up permissions for #{channel_name}: {e}")


# ── Static channel setup ──────────────────────────────────────────────────────

async def setup_static_channels_for_guild(
    guild: discord.Guild,
    bot_user: discord.ClientUser,
    runner_access: dict | None = None,
) -> None:
    """Create the standard Welcome / Tournament Officials / Chapters / Volunteers structure."""
    if not guild:
        return

    runner_access = runner_access or {}
    print(f"Setting up static channels for {guild.name}...")

    runner_role = await get_or_create_role(guild, "Runner")
    awards_role = await get_or_create_role(guild, "Awards")

    # Welcome
    welcome_cat = await get_or_create_category(guild, "Welcome")
    if welcome_cat:
        welcome_ch = await get_or_create_channel(guild, "welcome", welcome_cat, runner_access=runner_access)
        if welcome_ch:
            await post_welcome_tldr(welcome_ch, bot_user)
            await post_welcome_instructions(welcome_ch, bot_user)

    # Tournament Officials
    to_cat = await get_or_create_category(guild, "Tournament Officials")
    if to_cat:
        for ch_name in ["runner", "scoring", "awards-ceremony"]:
            ch = discord.utils.get(guild.text_channels, name=ch_name)
            if not ch:
                try:
                    overwrites: dict = {
                        guild.default_role: discord.PermissionOverwrite(read_messages=False),
                    }
                    if runner_role:
                        overwrites[runner_role] = discord.PermissionOverwrite(
                            read_messages=True, send_messages=True, read_message_history=True,
                        )
                    if ch_name == "awards-ceremony" and awards_role:
                        overwrites[awards_role] = discord.PermissionOverwrite(
                            read_messages=True, send_messages=True, read_message_history=True,
                        )
                    ch = await handle_rate_limit(
                        guild.create_text_channel(
                            name=ch_name,
                            category=to_cat,
                            overwrites=overwrites,
                            reason="Auto-created by LAM Bot - Tournament Officials only",
                        ),
                        f"creating channel '{ch_name}'",
                    )
                    if ch and runner_role:
                        await add_runner_access(ch, runner_role)
                    if ch and ch_name == "awards-ceremony" and awards_role:
                        await add_runner_access(ch, awards_role)
                except discord.Forbidden:
                    print(f"No permission to create channel '{ch_name}'")
                except Exception as e:
                    print(f"Error creating channel '{ch_name}': {e}")
            else:
                # Update permissions on existing channel
                try:
                    overwrites = ch.overwrites
                    overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)
                    if runner_role:
                        overwrites[runner_role] = discord.PermissionOverwrite(
                            read_messages=True, send_messages=True, read_message_history=True,
                        )
                    if ch_name == "awards-ceremony" and awards_role:
                        overwrites[awards_role] = discord.PermissionOverwrite(
                            read_messages=True, send_messages=True, read_message_history=True,
                        )
                    await handle_rate_limit(
                        ch.edit(overwrites=overwrites, reason="Updating to restrict to Runner only"),
                        f"editing channel '{ch_name}' permissions",
                    )
                except Exception as e:
                    print(f"Error updating permissions for #{ch_name}: {e}")
                if runner_role:
                    await add_runner_access(ch, runner_role)
                if ch_name == "awards-ceremony" and awards_role:
                    await add_runner_access(ch, awards_role)

    # Chapters (create category only; channels populated by /syncrooms)
    await get_or_create_category(guild, "Chapters")

    # Volunteers
    vol_cat = await get_or_create_category(guild, "Volunteers")
    if vol_cat:
        # lead-es (restricted to Lead ES role)
        lead_es_role = await get_or_create_role(guild, "Lead ES")
        if not discord.utils.get(guild.text_channels, name="lead-es"):
            try:
                ow: dict = {guild.default_role: discord.PermissionOverwrite(read_messages=False)}
                if lead_es_role:
                    ow[lead_es_role] = discord.PermissionOverwrite(
                        read_messages=True, send_messages=True, read_message_history=True,
                    )
                await handle_rate_limit(
                    guild.create_text_channel(
                        name="lead-es",
                        category=vol_cat,
                        overwrites=ow,
                        reason="Auto-created by LAM Bot for Lead ES",
                    ),
                    "creating channel 'lead-es'",
                )
            except discord.Forbidden:
                print("No permission to create channel 'lead-es'")
            except Exception as e:
                print(f"Error creating channel 'lead-es': {e}")

        # Standard volunteer text channels
        for ch_name in ["general", "useful-links", "announcements", "random"]:
            ch = await get_or_create_channel(guild, ch_name, vol_cat, runner_access=runner_access)
            if ch and ch_name in ("useful-links", "announcements"):
                try:
                    ow = ch.overwrites
                    ow[guild.default_role] = discord.PermissionOverwrite(
                        read_messages=True, send_messages=False, read_message_history=True,
                    )
                    await handle_rate_limit(
                        ch.edit(overwrites=ow, reason=f"Removed write access for default role in #{ch_name}"),
                        f"editing channel '{ch_name}' permissions",
                    )
                except discord.Forbidden:
                    print(f"No permission to edit #{ch_name}")
                except Exception as e:
                    print(f"Error updating #{ch_name}: {e}")

        # Help forum channel
        help_exists = any(
            c.name == "help" and c.type == discord.ChannelType.forum
            for c in guild.channels
        )
        if not help_exists:
            try:
                ow = {
                    guild.default_role: discord.PermissionOverwrite(
                        read_messages=True,
                        send_messages=True,
                        read_message_history=True,
                        create_public_threads=True,
                        send_messages_in_threads=True,
                        manage_threads=True,
                    ),
                }
                if runner_role:
                    ow[runner_role] = discord.PermissionOverwrite(
                        read_messages=True,
                        send_messages=True,
                        read_message_history=True,
                        create_public_threads=True,
                        send_messages_in_threads=True,
                        manage_threads=True,
                    )
                if hasattr(guild, "create_forum_channel"):
                    await handle_rate_limit(
                        guild.create_forum_channel(
                            name="help",
                            category=vol_cat,
                            overwrites=ow,
                            reason="Auto-created by LAM Bot",
                        ),
                        "creating forum channel 'help'",
                    )
                elif hasattr(guild, "create_forum"):
                    await handle_rate_limit(
                        guild.create_forum(
                            name="help",
                            category=vol_cat,
                            overwrites=ow,
                            reason="Auto-created by LAM Bot",
                        ),
                        "creating forum 'help'",
                    )
                else:
                    print("Forum creation not supported — please manually create a 'help' forum in Volunteers")
            except (discord.Forbidden, AttributeError):
                print("No permission / no support for forum channels — create 'help' forum manually")
            except Exception as e:
                print(f"Error creating forum channel 'help': {e}")

    print("Finished setting up static channels")


# ── Role hierarchy ────────────────────────────────────────────────────────────

async def move_bot_role_to_top_for_guild(guild: discord.Guild) -> None:
    """Move the bot's managed role to the highest reachable position and tint it teal."""
    if not guild:
        return
    if not guild.me.guild_permissions.manage_roles:
        print("Bot missing 'Manage Roles' permission; cannot move bot role to top")
        return

    bot_role: discord.Role | None = None
    for role in guild.roles:
        if role.managed and guild.me in role.members:
            bot_role = role
            break

    if not bot_role:
        print("Could not find bot's managed role")
        return

    higher_unmovable = [r for r in guild.roles if r.position > bot_role.position and (r.managed or r == guild.default_role)]
    max_pos = (min(r.position for r in higher_unmovable) - 1) if higher_unmovable else len(guild.roles) - 1

    if bot_role.color != discord.Color.teal():
        try:
            await handle_rate_limit(
                bot_role.edit(color=discord.Color.teal(), reason="Making bot role teal"),
                "editing bot role color",
            )
        except Exception as e:
            print(f"Could not change bot role color: {e}")

    if bot_role.position != max_pos:
        try:
            await handle_rate_limit(
                bot_role.edit(position=max_pos, reason="Moving bot role to top"),
                "moving bot role position",
            )
        except Exception as e:
            print(f"Could not move bot role to top: {e}")


async def organize_role_hierarchy_for_guild(
    guild: discord.Guild,
    chapter_role_names: set[str] | None = None,
) -> None:
    """Arrange guild roles: structural roles at top, chapters in middle, others at bottom."""
    if not guild:
        return
    if not guild.me.guild_permissions.manage_roles:
        print("Bot missing 'Manage Roles' permission; cannot organize role hierarchy")
        return

    chapter_role_names = chapter_role_names or set()

    priority_order = [
        "Volunteer", "Lead ES", "Social Media", "Photographer",
        "Arbitrations", "Awards", "Runner", "VIPer", "Admin",
    ]

    bot_role: discord.Role | None = None
    for role in guild.roles:
        if role.managed and guild.me in role.members:
            bot_role = role
            break

    if not bot_role:
        print("Could not find bot's managed role")
        return

    all_roles = [r for r in guild.roles if r.name != "@everyone" and r != bot_role]

    priority_objs: list[discord.Role] = []
    chapter_roles: list[discord.Role] = []
    other_roles: list[discord.Role] = []

    for role in all_roles:
        if role.position >= bot_role.position:
            continue  # Can't touch roles above the bot
        if role.name in priority_order:
            priority_objs.append(role)
        elif role.name == "Unaffiliated" or role.name in chapter_role_names:
            chapter_roles.append(role)
        else:
            other_roles.append(role)

    priority_objs.sort(key=lambda r: priority_order.index(r.name) if r.name in priority_order else 999)
    chapter_roles.sort(key=lambda r: r.name.lower())
    other_roles.sort(key=lambda r: r.name.lower())

    final_order = other_roles + chapter_roles + priority_objs
    positions = {role: i + 1 for i, role in enumerate(final_order)}

    try:
        await handle_rate_limit(
            guild.edit_role_positions(positions, reason="Organizing role hierarchy"),
            "editing role positions",
        )
        print("Role hierarchy organized")
    except Exception as e:
        print(f"Error organizing role hierarchy: {e}")


# ── Runner access management ──────────────────────────────────────────────────

async def remove_runner_access_from_building_channels_for_guild(guild: discord.Guild) -> None:
    """Strip Runner role permissions from all non-static category channels."""
    runner_role = discord.utils.get(guild.roles, name="Runner")
    if not runner_role:
        return

    removed = 0
    for ch in guild.text_channels:
        if ch.category and ch.category.name not in STATIC_CATEGORIES:
            try:
                overwrites = ch.overwrites
                if runner_role in overwrites:
                    del overwrites[runner_role]
                    await handle_rate_limit(
                        ch.edit(overwrites=overwrites, reason=f"Removed {runner_role.name} from building channel"),
                        f"removing access from channel '{ch.name}'",
                    )
                    removed += 1
            except Exception as e:
                print(f"Error removing Runner access from #{ch.name}: {e}")

    print(f"Removed Runner access from {removed} building/event channel(s)")


async def give_runner_access_to_all_channels_for_guild(guild: discord.Guild) -> None:
    """Grant Runner access to all channels in the three static categories (not building channels)."""
    runner_role = discord.utils.get(guild.roles, name="Runner")
    if not runner_role:
        return

    counts = {"Welcome": 0, "Tournament Officials": 0, "Volunteers": 0, "forums": 0}

    for ch in guild.text_channels:
        if ch.category and ch.category.name in ("Welcome", "Tournament Officials", "Volunteers"):
            try:
                await add_runner_access(ch, runner_role)
                counts[ch.category.name] += 1
            except Exception as e:
                print(f"Error adding Runner access to #{ch.name}: {e}")

    for ch in guild.channels:
        if ch.type == discord.ChannelType.forum and ch.category:
            if ch.category.name in ("Welcome", "Tournament Officials", "Volunteers"):
                try:
                    ow = ch.overwrites
                    ow[runner_role] = discord.PermissionOverwrite(
                        read_messages=True,
                        send_messages=True,
                        read_message_history=True,
                        create_public_threads=True,
                        send_messages_in_threads=True,
                    )
                    await handle_rate_limit(
                        ch.edit(overwrites=ow, reason=f"Added {runner_role.name} access"),
                        f"editing forum channel '{ch.name}' permissions",
                    )
                    counts["forums"] += 1
                except Exception as e:
                    print(f"Error adding Runner access to forum #{ch.name}: {e}")

    print(
        f"Added Runner access: {counts['Welcome']} Welcome, "
        f"{counts['Tournament Officials']} Tournament Officials, "
        f"{counts['Volunteers']} Volunteers, {counts['forums']} forums"
    )


# ── Special admin setup ───────────────────────────────────────────────────────

async def setup_ezhang_admin_role(guild: discord.Guild) -> None:
    """Grant Admin role to ezhang. if they are present in the guild."""
    if not guild:
        return

    ezhang = next(
        (m for m in guild.members
         if m.name.lower() == "ezhang." or (m.global_name and m.global_name.lower() == "ezhang.")),
        None,
    )
    if not ezhang:
        return

    try:
        admin_role = await get_or_create_role(guild, "Admin")
        if admin_role and admin_role not in ezhang.roles:
            await handle_rate_limit(
                ezhang.add_roles(admin_role, reason="Special admin access for ezhang."),
                f"adding Admin to {ezhang}",
            )
            print(f"Granted Admin to {ezhang} (ezhang.) in {guild.name}")
    except Exception as e:
        print(f"Could not grant Admin to ezhang. in {guild.name}: {e}")


# ── Server reset ──────────────────────────────────────────────────────────────

async def reset_server_for_guild(guild: discord.Guild) -> None:
    """DANGER: Delete all channels, categories, and non-managed roles; clear all nicknames."""
    if not guild:
        return

    print("STARTING COMPLETE SERVER RESET — 3 second grace period...")
    await asyncio.sleep(3)

    # Nicknames
    nick_count = 0
    for member in guild.members:
        if member.nick and not member.bot:
            try:
                await handle_rate_limit(
                    member.edit(nick=None, reason="Server reset"),
                    f"resetting nickname for {member}",
                )
                nick_count += 1
            except (discord.Forbidden, Exception):
                pass
    print(f"Reset {nick_count} nickname(s)")

    # Text channels
    for ch in list(guild.text_channels):
        try:
            await ch.delete(reason="Server reset")
        except Exception:
            pass

    # Voice channels
    for ch in list(guild.voice_channels):
        try:
            await ch.delete(reason="Server reset")
        except Exception:
            pass

    # Forum channels
    for ch in list(guild.channels):
        if getattr(ch, "type", None) == discord.ChannelType.forum:
            try:
                await ch.delete(reason="Server reset")
            except Exception:
                pass

    # Categories
    for cat in list(guild.categories):
        try:
            await cat.delete(reason="Server reset")
        except Exception:
            pass

    # Roles (skip @everyone, managed, and roles above bot's top role)
    role_count = 0
    for role in list(guild.roles):
        if role.name != "@everyone" and not role.managed and role < guild.me.top_role:
            try:
                await role.delete(reason="Server reset")
                role_count += 1
            except Exception:
                pass

    print(f"Server reset complete: {nick_count} nicknames, {role_count} roles removed")


# ── Top-level orchestration ───────────────────────────────────────────────────

async def generate_building_structures(
    guild: discord.Guild,
    *,
    runner_access: dict | None = None,
    spreadsheets: dict | None = None,
    force_refresh_welcome: bool = False,
) -> tuple[int, int]:
    """Create all building/event channels from the Room Assignments sheet.

    Returns (num_structures, num_buildings).
    """
    runner_access = runner_access or {}
    spreadsheets = spreadsheets or {}

    print("Generating building structures from Room Assignments...")
    room_data = await data_router.list_events(guild.id, spreadsheets=spreadsheets)

    if not room_data:
        print("No data found in Room Assignments sheet")
        return 0, 0

    building_structures: set[tuple[str, str, str]] = set()
    buildings: set[str] = set()

    for row in room_data:
        building = str(row.get("building", "")).strip()
        event = str(row.get("name", "")).strip()
        room = str(row.get("room", "")).strip()

        if building and event and event not in PRIORITY_ROLES:
            building_structures.add((building, event, room))
            buildings.add(building)

    print(f"Found {len(building_structures)} unique building/event combos across {len(buildings)} building(s)")

    for building, event, room in building_structures:
        await setup_building_structure(
            guild, building, event, room,
            runner_access=runner_access,
            spreadsheets=spreadsheets,
        )

    if force_refresh_welcome:
        for building in buildings:
            chat_name = f"{sanitize_for_discord(building)}-chat"
            chat = discord.utils.get(guild.text_channels, name=chat_name)
            if chat:
                try:
                    async for msg in chat.history(limit=50):
                        if msg.embeds and msg.embeds[0].title and f"Welcome to {building}" in msg.embeds[0].title:
                            await msg.delete()
                            await asyncio.sleep(0.5)
                            break
                except Exception as e:
                    print(f"Could not clear old welcome messages in #{chat.name}: {e}")
                await send_building_welcome_message(guild, chat, building, spreadsheets=spreadsheets)

    print("Organizing building categories and channels...")
    await sort_building_categories_alphabetically(guild)
    await sort_channels_in_building_categories(guild)

    return len(building_structures), len(buildings)
