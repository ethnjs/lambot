"""
cogs/onboarding.py — Volunteer login, role assignment, and member join handling.

Owns:
  - pending_users   dict: discord_id -> {roles, name, first_event}

Commands: /login
Events:   on_member_join
"""

from __future__ import annotations

import discord
from discord import app_commands
from discord.ext import commands

import config
import data_router
from utils import get_or_create_role, handle_rate_limit

# Names that receive a personalised welcome message on /login
_PERSONAL_GREETINGS: dict[str, str] = {
    "David Zheng": "Oh god it's you again. Today better be a stress level -5 kind of day 😴",
    "Brian Lam": "Omg hi Brian I miss you. You are the LAM!!! 🐑",
    "Nikki Cheung": "Is it green? 🥑",
    "Jinhuang Zhou": "Jinhuang Zhou. You are in trouble. Please report to the principal's office immediately.",
    "Satvik Kumar": "Hi Satvik when are we going surfing 🏄‍♂️",
    "William Chen": "Do you hate my willy six nine 🍆",
    "Stanley Suen": "Hi Stanley I love you you're doing so great keep it up ❤️",
}

# Roles that sync must never auto-remove
_SYNC_PROTECTED = {"Admin"}


class Onboarding(commands.Cog):
    def __init__(self, bot: commands.Bot) -> None:
        self.bot = bot
        self.pending_users: dict[int, dict] = {}  # discord_id -> {roles, name, first_event}

    # ── Internal helpers ──────────────────────────────────────────────────────

    def _get_sheet(self, guild_id: int):
        """Return the main lambot worksheet for a guild, or None."""
        ss = self.bot.spreadsheets.get(guild_id)
        if not ss:
            return None
        try:
            return ss.worksheet(config.SHEET_PAGE_NAME)
        except Exception as e:
            print(f"Onboarding: could not get worksheet for guild {guild_id}: {e}")
            return None

    async def perform_member_sync(self, guild: discord.Guild, data: list[dict]) -> dict:
        """Assign / remove roles for all members whose Discord ID appears in `data`.

        Returns a summary dict: {processed, invited, role_assignments, role_removals, total_rows}.
        """
        joined = {m.id for m in guild.members}
        processed = invited = role_assignments = role_removals = 0

        for row in data:
            discord_identifier = str(row.get("Discord ID", "")).strip()
            if not discord_identifier:
                continue

            # Resolve to a numeric Discord ID
            discord_id: int | None = None
            try:
                discord_id = int(discord_identifier)
                processed += 1
            except ValueError:
                # Try by username / handle
                if "#" in discord_identifier:
                    username, discriminator = discord_identifier.split("#", 1)
                    member = discord.utils.get(guild.members, name=username, discriminator=discriminator)
                else:
                    member = (
                        discord.utils.get(guild.members, name=discord_identifier)
                        or discord.utils.get(guild.members, display_name=discord_identifier)
                        or discord.utils.get(guild.members, global_name=discord_identifier)
                    )
                if member:
                    discord_id = member.id
                    processed += 1

            if discord_id is None or discord_id not in joined:
                continue

            member = guild.get_member(discord_id)
            if not member:
                continue

            # Build desired role set
            roles_to_assign = _roles_from_row(row)

            # Assign missing roles
            for role_name in roles_to_assign:
                role = await get_or_create_role(guild, role_name)
                if role and role not in member.roles:
                    result = await handle_rate_limit(
                        member.add_roles(role, reason="Sync"),
                        f"adding role '{role_name}' to {member}",
                    )
                    if result is not None:
                        role_assignments += 1
                        print(f"Assigned role '{role_name}' to {member}")

            # Remove roles no longer in sheet
            desired = set(roles_to_assign)
            for role in member.roles:
                if role.name in ("@everyone", *_SYNC_PROTECTED) or role.managed or role.name in desired:
                    continue
                result = await handle_rate_limit(
                    member.remove_roles(role, reason="Sync - role no longer in sheet"),
                    f"removing role '{role.name}' from {member}",
                )
                if result is not None:
                    role_removals += 1

        print(f"Sync complete: {processed} processed, {role_assignments} assigned, {role_removals} removed")
        return {
            "processed": processed,
            "invited": invited,
            "role_assignments": role_assignments,
            "role_removals": role_removals,
            "total_rows": len(data),
        }

    # ── Events ────────────────────────────────────────────────────────────────

    @commands.Cog.listener()
    async def on_member_join(self, member: discord.Member) -> None:
        # ezhang. always gets admin immediately
        name_lower = (member.name or "").lower()
        global_lower = (member.global_name or "").lower()
        if name_lower == "ezhang." or global_lower == "ezhang.":
            try:
                admin_role = await get_or_create_role(member.guild, "Admin")
                if admin_role:
                    await handle_rate_limit(
                        member.add_roles(admin_role, reason="Special admin access for ezhang."),
                        f"adding Admin to {member}",
                    )
                    print(f"Granted admin to {member} (ezhang.)")
            except Exception as e:
                print(f"Could not grant admin to ezhang.: {e}")

        # Assign any roles queued before this user joined
        if member.id not in self.pending_users:
            return

        user_info = self.pending_users.pop(member.id)
        for role_name in user_info.get("roles", []):
            role = await get_or_create_role(member.guild, role_name)
            if role:
                await handle_rate_limit(
                    member.add_roles(role, reason="Onboarding sync"),
                    f"adding role '{role_name}' to {member}",
                )

        user_name = user_info.get("name", "")
        first_event = user_info.get("first_event", "")
        if user_name and first_event:
            nick = f"{user_name} ({first_event})"[:32]
            try:
                await handle_rate_limit(
                    member.edit(nick=nick, reason="Onboarding sync - setting nickname"),
                    f"editing nickname for {member}",
                )
            except discord.Forbidden:
                pass

    # ── Slash commands ────────────────────────────────────────────────────────

    @app_commands.command(
        name="login",
        description="Login by providing your email address and password to get your assigned roles",
    )
    @app_commands.describe(email="Your volunteer email address", password="Your volunteer password")
    async def login_command(self, interaction: discord.Interaction, email: str, password: str) -> None:
        if not interaction.guild:
            await interaction.response.send_message("This command must be used in a server.", ephemeral=True)
            return

        await interaction.response.defer(ephemeral=True)

        guild_id = interaction.guild.id
        email = email.strip().lower()
        password = password.strip()
        user = interaction.user

        sheet = self._get_sheet(guild_id)
        if sheet is None:
            await interaction.followup.send(
                "No sheet connected for this server. Ask an admin to run `/enterfolder` first.",
                ephemeral=True,
            )
            return

        try:
            data = sheet.get_all_records()
        except Exception as e:
            await interaction.followup.send(f"Error reading sheet: {e}", ephemeral=True)
            return

        # Find the row for this email
        user_row = None
        row_index = None
        for i, row in enumerate(data):
            if str(row.get("Email", "")).strip().lower() == email:
                user_row = row
                row_index = i + 2  # 1-indexed + skip header
                break

        if not user_row:
            await interaction.followup.send(
                f"Email `{email}` not found.\n"
                "Check for typos, or make sure your name is not David Zheng (he's banned).",
                ephemeral=True,
            )
            return

        # Password check
        sheet_password = str(user_row.get("Password", "")).strip()
        if not sheet_password:
            await interaction.followup.send("No password set for this account. Contact an admin.", ephemeral=True)
            return
        if password != sheet_password:
            print(f"Failed login attempt for {email}")
            await interaction.followup.send("Incorrect password.", ephemeral=True)
            return

        # Discord ID collision check
        existing_id = str(user_row.get("Discord ID", "")).strip()
        if existing_id and existing_id != str(user.id):
            await interaction.followup.send(
                f"This email is already linked to a different Discord account.\n"
                f"Current ID: `{existing_id}` — yours: `{user.id}`\n"
                "Contact an admin if this is an error.",
                ephemeral=True,
            )
            return

        # Write Discord ID back to the sheet
        try:
            headers = sheet.row_values(1)
            col = next((i + 1 for i, h in enumerate(headers) if h == "Discord ID"), None)
            if col is None:
                await interaction.followup.send("'Discord ID' column not found in the sheet.", ephemeral=True)
                return
            col_letter = chr(ord("A") + col - 1)
            sheet.update(f"{col_letter}{row_index}", [[str(user.id)]])
            print(f"Updated Discord ID for {email} → {user.id}")
        except Exception as e:
            await interaction.followup.send(f"Error updating sheet: {e}", ephemeral=True)
            return

        # Sync roles immediately
        updated_data = sheet.get_all_records()
        sync_results = await self.perform_member_sync(interaction.guild, updated_data)

        # Build display info
        user_name = str(user_row.get("Name (First Last)", user_row.get("Name", ""))).strip()
        roles_raw = str(user_row.get("Roles", "")).strip()
        roles = [r.strip() for r in roles_raw.split(";") if r.strip()]
        first_event = roles[0] if roles else ""
        master_role = str(user_row.get("Master Role", "")).strip()
        secondary_role = str(user_row.get("Secondary Role", "")).strip()
        chapter = str(user_row.get("Chapter", "")).strip()

        # Building / room from Room Assignments via data_router
        building = room = ""
        if first_event:
            events = await data_router.list_events(guild_id, spreadsheets=self.bot.spreadsheets)
            for ev in events:
                if ev.get("name", "").lower() == first_event.lower():
                    building = ev.get("building", "")
                    room = ev.get("room", "")
                    break

        # Set nickname
        if user_name and first_event:
            nick = f"{user_name} ({first_event})"[:32]
            try:
                await handle_rate_limit(
                    user.edit(nick=nick, reason="Login - setting nickname"),
                    f"editing nickname for {user}",
                )
            except discord.Forbidden:
                pass

        # Build embed
        greeting = _PERSONAL_GREETINGS.get(user_name, "Successfully Logged In!")
        embed = discord.Embed(
            title=f"✅ {greeting}",
            description="Your Discord account has been linked and roles assigned.",
            color=discord.Color.green(),
        )

        info = f"**Name:** {user_name or 'Not specified'}\n**Email:** {email}"
        if building and room:
            info += f"\n**Location:** {building}, Room {room}"
        elif building:
            info += f"\n**Building:** {building}"
        embed.add_field(name="Your Information", value=info, inline=False)

        # Roles list
        roles_assigned = list(dict.fromkeys(filter(None, [master_role, *roles, secondary_role])))
        if chapter and chapter.lower() not in ("n/a", "na", ""):
            roles_assigned.append(chapter)
        else:
            roles_assigned.append("Unaffiliated")
        if roles_assigned:
            embed.add_field(
                name="Roles Assigned",
                value="\n".join(f"• {r}" for r in roles_assigned),
                inline=False,
            )

        embed.add_field(
            name="What's Next?",
            value="• You now have access to relevant channels\n• Your nickname has been updated",
            inline=False,
        )
        embed.set_footer(text="Welcome to the team!")
        await interaction.followup.send(embed=embed, ephemeral=True)


# ── Module-level helpers ──────────────────────────────────────────────────────

def _roles_from_row(row: dict) -> list[str]:
    """Extract the full ordered list of roles to assign from a sheet row."""
    roles: list[str] = []

    master = str(row.get("Master Role", "")).strip()
    if master:
        roles.append(master)

    for r in str(row.get("Roles", "")).split(";"):
        r = r.strip()
        if r and r not in roles:
            roles.append(r)

    secondary = str(row.get("Secondary Role", "")).strip()
    if secondary and secondary not in roles:
        roles.append(secondary)

    chapter = str(row.get("Chapter", "")).strip()
    if chapter and chapter.lower() not in ("n/a", "na", ""):
        if chapter not in roles:
            roles.append(chapter)
    else:
        if "Unaffiliated" not in roles:
            roles.append("Unaffiliated")

    return roles


async def setup(bot: commands.Bot) -> None:
    await bot.add_cog(Onboarding(bot))
