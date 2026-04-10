"""
cogs/admin.py — Admin-only utility commands.

Commands included here are standalone and don't require heavy guild-structure
helper functions. The structural commands (syncrooms, organizeroles, resetserver,
assignrunnerzones, sendallmaterials) will be added once those helpers are extracted.

Commands: /help, /gettemplate, /serviceaccount, /sheetinfo, /sync,
          /reloadcommands, /cacheinfo, /clearcache, /set_runner_all_access,
          /refreshnicknames, /msg
"""

from __future__ import annotations

import asyncio
import json
import os
from datetime import datetime

import discord
from discord import app_commands
from discord.ext import commands

import config
from utils import handle_rate_limit


def _admin_only(interaction: discord.Interaction) -> bool:
    return interaction.user.guild_permissions.administrator


class Admin(commands.Cog):
    def __init__(self, bot: commands.Bot) -> None:
        self.bot = bot
        self._lock = asyncio.Lock()  # guards multi-step admin operations

    def _spreadsheet(self, guild_id: int):
        return self.bot.spreadsheets.get(guild_id)

    def _sheet(self, guild_id: int):
        ss = self._spreadsheet(guild_id)
        if not ss:
            return None
        try:
            return ss.worksheet(config.SHEET_PAGE_NAME)
        except Exception:
            return None

    # ── /help ─────────────────────────────────────────────────────────────────

    @app_commands.command(name="help", description="Show all available bot commands and how to use them")
    async def help_command(self, interaction: discord.Interaction) -> None:
        await interaction.response.defer(ephemeral=True)
        embed = discord.Embed(
            title="Getting Started With LamBot",
            description=(
                "**For Users:**\n"
                "1. Use `/login email:your@email.com password:yourpassword` to log in and get your roles\n"
                "2. Access to channels will be granted based on your assigned roles\n\n"
                "**For Admins:**\n"
                "1. Use `/gettemplate` to get the template Google Drive folder\n"
                "2. Use `/serviceaccount` to get the bot's service account email\n"
                "3. Share your Google Drive folder with that email (Editor permissions)\n"
                "4. Get the folder link: Right-click → Share → Copy link\n"
                "5. Use `/enterfolder` with that folder link\n"
                "6. Use `/sheetinfo` to verify the connection"
            ),
            color=discord.Color.blue(),
        )
        embed.set_footer(text="Need more help? Contact your server administrator.")
        await interaction.followup.send(embed=embed, ephemeral=True)

    # ── /gettemplate ──────────────────────────────────────────────────────────

    @app_commands.command(name="gettemplate", description="Get a link to the template Google Drive folder")
    async def get_template_command(self, interaction: discord.Interaction) -> None:
        await interaction.response.defer(ephemeral=True)
        template_url = "https://drive.google.com/drive/folders/1drRK7pSdCpbqzJfaDhFtKlYUrf_uYsN8?usp=sharing"
        embed = discord.Embed(
            title="Template Google Drive Folder",
            description=f"Access all the template files here:\n[**Click here to open the template folder**]({template_url})",
            color=discord.Color.blue(),
        )
        embed.add_field(
            name="Important: Share Your Folder!",
            value=(
                f"When you create your own folder from this template, share it with:\n"
                f"`{config.SERVICE_EMAIL}`\n\n"
                "Steps:\n"
                "1. Right-click your folder in Google Drive\n"
                "2. Click 'Share'\n"
                "3. Add the email above with 'Editor' permissions\n"
                "4. Click 'Copy link' to get the folder URL\n\n"
                "Then use `/enterfolder` with that copied link!"
            ),
            inline=False,
        )
        await interaction.followup.send(embed=embed, ephemeral=True)

    # ── /serviceaccount ───────────────────────────────────────────────────────

    @app_commands.command(name="serviceaccount", description="Show the service account email for sharing Google Sheets")
    async def service_account_command(self, interaction: discord.Interaction) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        await interaction.response.defer(ephemeral=True)
        embed = discord.Embed(
            title="Service Account Information",
            description="Share your sheets/folder with this email so the bot can access them:",
            color=discord.Color.blue(),
        )
        embed.add_field(name="Service Account Email", value=f"`{config.SERVICE_EMAIL}`", inline=False)
        embed.add_field(
            name="How to Share",
            value=(
                "1. Open your Google Sheet or folder\n"
                "2. Click 'Share'\n"
                "3. Add the service account email above\n"
                "4. Set permissions to 'Editor'\n"
                "5. Click 'Send'"
            ),
            inline=False,
        )
        await interaction.followup.send(embed=embed, ephemeral=True)

    # ── /sheetinfo ────────────────────────────────────────────────────────────

    @app_commands.command(name="sheetinfo", description="Show information about the currently connected Google Sheet")
    async def sheet_info_command(self, interaction: discord.Interaction) -> None:
        await interaction.response.defer(ephemeral=True)
        guild_id = interaction.guild.id
        ss = self._spreadsheet(guild_id)

        if ss is None:
            embed = discord.Embed(
                title="No Sheet Connected",
                description="No Google Sheet is currently connected.\nUse `/enterfolder` to connect one.",
                color=discord.Color.orange(),
            )
            await interaction.followup.send(embed=embed, ephemeral=True)
            return

        try:
            sheet = ss.worksheet(config.SHEET_PAGE_NAME)
            data = sheet.get_all_records()
            embed = discord.Embed(
                title="Current Sheet Information",
                description=(
                    f"**Spreadsheet:** [{ss.title}]({ss.url})\n"
                    f"**Worksheet:** {sheet.title}\n"
                    f"**Rows:** {len(data)} users"
                ),
                color=discord.Color.green(),
            )
            worksheets = [ws.title for ws in ss.worksheets()]
            if len(worksheets) > 1:
                embed.add_field(
                    name="Available Worksheets",
                    value="\n".join(f"• {ws}" + (" ✅" if ws == sheet.title else "") for ws in worksheets),
                    inline=False,
                )
            if data:
                fields = [f"• {k}" for k, v in data[0].items() if k and v][:5]
                if fields:
                    embed.add_field(name="Available Fields", value="\n".join(fields), inline=False)
            embed.add_field(name="Sync Status", value="Syncing every minute automatically", inline=False)
            embed.set_footer(text="Use /sync to manually trigger a sync")
        except Exception as e:
            embed = discord.Embed(title="Error Reading Sheet", description=str(e), color=discord.Color.red())

        await interaction.followup.send(embed=embed, ephemeral=True)

    # ── /sync ─────────────────────────────────────────────────────────────────

    @app_commands.command(name="sync", description="Manually trigger a member sync from the current Google Sheet (admin only)")
    async def sync_command(self, interaction: discord.Interaction) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        if self._lock.locked():
            await interaction.response.send_message("Server configuration is in progress. Try again shortly.", ephemeral=True)
            return

        async with self._lock:
            await interaction.response.defer(ephemeral=True)

            if not interaction.guild:
                await interaction.followup.send("This command must be used in a server.", ephemeral=True)
                return

            guild_id = interaction.guild.id
            sheet = self._sheet(guild_id)
            if sheet is None:
                await interaction.followup.send("No sheet connected. Use `/enterfolder` first.", ephemeral=True)
                return

            try:
                data = sheet.get_all_records()
            except Exception as e:
                await interaction.followup.send(f"Could not fetch sheet data: {e}", ephemeral=True)
                return

            onboarding = self.bot.get_cog("Onboarding")
            if onboarding is None:
                await interaction.followup.send("Onboarding cog not loaded.", ephemeral=True)
                return

            results = await onboarding.perform_member_sync(interaction.guild, data)
            embed = discord.Embed(
                title="Manual Sync Complete!",
                description=(
                    f"**Processed:** {results['processed']} Discord IDs\n"
                    f"**Members:** {len(interaction.guild.members)}\n"
                    f"**Roles assigned:** {results['role_assignments']}\n"
                    f"**Roles removed:** {results['role_removals']}\n"
                    f"**Sheet rows:** {results['total_rows']}"
                ),
                color=discord.Color.green(),
            )
            await interaction.followup.send(embed=embed, ephemeral=True)

    # ── /reloadcommands ───────────────────────────────────────────────────────

    @app_commands.command(name="reloadcommands", description="Manually sync slash commands with Discord (Admin only)")
    async def reload_commands_command(self, interaction: discord.Interaction) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        await interaction.response.defer(ephemeral=True)
        try:
            synced = await self.bot.tree.sync()
            embed = discord.Embed(title="Commands Synced!", color=discord.Color.green())
            if synced:
                embed.description = f"Synced {len(synced)} slash commands."
                embed.add_field(
                    name="Commands",
                    value="\n".join(f"• `/{c.name}`" for c in synced),
                    inline=False,
                )
            await interaction.followup.send(embed=embed, ephemeral=True)
        except Exception as e:
            await interaction.followup.send(f"Error syncing commands: {e}", ephemeral=True)

    # ── /cacheinfo ────────────────────────────────────────────────────────────

    @app_commands.command(name="cacheinfo", description="Show cached spreadsheet connection info (Admin only)")
    async def cache_info_command(self, interaction: discord.Interaction) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        await interaction.response.defer(ephemeral=True)

        cache_path = config.CACHE_FILE
        if not os.path.exists(cache_path):
            await interaction.followup.send("No cache file found.", ephemeral=True)
            return

        try:
            with open(cache_path) as f:
                cache = json.load(f)

            guild_id = str(interaction.guild.id)
            guild_cache = cache.get("guilds", {}).get(guild_id)

            embed = discord.Embed(title="Cache Information", color=discord.Color.blue())
            embed.add_field(
                name="Cache File",
                value=f"`{cache_path}` ({os.path.getsize(cache_path)} bytes)",
                inline=False,
            )

            if guild_cache:
                embed.add_field(
                    name="This Server",
                    value=(
                        f"**Spreadsheet ID:** `{guild_cache.get('spreadsheet_id', 'N/A')}`\n"
                        f"**Worksheet:** {guild_cache.get('worksheet_name', 'N/A')}"
                    ),
                    inline=False,
                )
            else:
                embed.add_field(name="This Server", value="No cached connection for this server.", inline=False)

            runner_access = cache.get("runner_access_settings", {}).get(guild_id, "Not set")
            embed.add_field(name="Runner All-Access", value=str(runner_access), inline=True)

            await interaction.followup.send(embed=embed, ephemeral=True)
        except Exception as e:
            await interaction.followup.send(f"Error reading cache: {e}", ephemeral=True)

    # ── /clearcache ───────────────────────────────────────────────────────────

    @app_commands.command(name="clearcache", description="Clear the cached spreadsheet connection (Admin only)")
    async def clear_cache_command(self, interaction: discord.Interaction) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        if self._lock.locked():
            await interaction.response.send_message("Server configuration is in progress. Try again shortly.", ephemeral=True)
            return

        async with self._lock:
            await interaction.response.defer(ephemeral=True)
            guild_id = interaction.guild.id

            sc = self.bot.sheets_client
            cleared = sc.clear_guild_from_cache(guild_id) if sc else False

            self.bot.spreadsheets.pop(guild_id, None)
            self.bot.sheets.pop(guild_id, None)

            if cleared:
                embed = discord.Embed(
                    title="Cache Cleared",
                    description="Cached connection for this server has been removed.\nUse `/enterfolder` to reconnect.",
                    color=discord.Color.green(),
                )
            else:
                embed = discord.Embed(
                    title="No Cache Found",
                    description="No cached connection found for this server.",
                    color=discord.Color.orange(),
                )
            await interaction.followup.send(embed=embed, ephemeral=True)

    # ── /set_runner_all_access ────────────────────────────────────────────────

    @app_commands.command(name="set_runner_all_access", description="Set if runners get access to all building/event channels (Admin only)")
    @app_commands.describe(runner_access="1 to give Runners access to all rooms, 0 to restrict them")
    async def set_runner_all_access_command(self, interaction: discord.Interaction, runner_access: int) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        await interaction.response.defer(ephemeral=True)

        guild = interaction.guild
        guild_id = guild.id
        new_val = bool(runner_access)
        current_val = bool(self.bot.runner_all_access.get(guild_id, 0))

        if new_val == current_val:
            state = "already HAS" if new_val else "is already RESTRICTED from"
            await interaction.followup.send(f"Runner role {state} all-access.", ephemeral=True)
            return

        self.bot.runner_all_access[guild_id] = 1 if new_val else 0
        sc = self.bot.sheets_client
        if sc:
            sc.save_runner_access_to_cache(self.bot.runner_all_access)

        runner_role = discord.utils.get(guild.roles, name="Runner")
        if not runner_role:
            await interaction.followup.send("Runner role not found. Run `/enterfolder` first.", ephemeral=True)
            return

        static_categories = {"Welcome", "Tournament Officials", "Volunteers", "Chapters"}
        modified = 0
        try:
            for category in guild.categories:
                if category.name in static_categories:
                    continue
                for ch in category.channels:
                    if not isinstance(ch, (discord.TextChannel, discord.VoiceChannel)):
                        continue
                    overwrites = ch.overwrites
                    if new_val:
                        overwrites[runner_role] = discord.PermissionOverwrite(
                            read_messages=True, send_messages=True, read_message_history=True
                        )
                    elif runner_role in overwrites:
                        del overwrites[runner_role]
                    await handle_rate_limit(
                        ch.edit(overwrites=overwrites, reason=f"Runner all-access → {new_val}"),
                        f"editing '{ch.name}' permissions",
                    )
                    modified += 1

            action = "GRANTED" if new_val else "REMOVED"
            await interaction.followup.send(
                f"Successfully **{action}** runner access on {modified} channel(s).", ephemeral=True
            )
        except discord.Forbidden:
            await interaction.followup.send("Bot lacks permission to modify channel overwrites.", ephemeral=True)
        except Exception as e:
            await interaction.followup.send(f"Error updating runner access: {e}", ephemeral=True)

    # ── /refreshnicknames ─────────────────────────────────────────────────────

    @app_commands.command(name="refreshnicknames", description="Reapply nicknames for all users with a Discord ID (Admin only)")
    async def refresh_nicknames_command(self, interaction: discord.Interaction) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        await interaction.response.defer(ephemeral=True)

        guild = interaction.guild
        ss = self._spreadsheet(guild.id)
        if ss is None:
            await interaction.followup.send("No spreadsheet connected for this server.", ephemeral=True)
            return

        try:
            data = ss.worksheet(config.SHEET_PAGE_NAME).get_all_records()
        except Exception as e:
            await interaction.followup.send(f"Error reading sheet: {e}", ephemeral=True)
            return

        updated = skipped = 0
        for row in data:
            try:
                discord_id = int(str(row.get("Discord ID", "")).strip())
            except ValueError:
                continue

            member = guild.get_member(discord_id)
            if not member:
                continue

            user_name = str(row.get("Name", "")).strip()
            roles = [r.strip() for r in str(row.get("Roles", "")).split(";") if r.strip()]
            first_event = roles[0] if roles else ""

            if not user_name or not first_event:
                skipped += 1
                continue

            nick = f"{user_name} ({first_event})"[:32]
            try:
                await handle_rate_limit(
                    member.edit(nick=nick, reason="Admin nickname refresh"),
                    f"editing nickname for {member}",
                )
                updated += 1
            except discord.Forbidden:
                pass
            except Exception as e:
                print(f"Error updating nickname for {member}: {e}")

        await interaction.followup.send(
            f"Nickname refresh complete — updated: **{updated}**, skipped: **{skipped}**",
            ephemeral=True,
        )

    # ── /msg ──────────────────────────────────────────────────────────────────

    @app_commands.command(name="msg", description="Send a message as the bot (Admin only)")
    @app_commands.describe(message="The message to send", channel="Channel to send to (defaults to current)")
    async def msg_command(self, interaction: discord.Interaction, message: str, channel: discord.TextChannel = None) -> None:
        await interaction.response.defer(ephemeral=True)

        admin_role = discord.utils.get(interaction.user.roles, name="Admin")
        if not admin_role:
            await interaction.followup.send("You need the Admin role to use this command.", ephemeral=True)
            return

        target = channel or interaction.channel
        if not target.permissions_for(interaction.guild.me).send_messages:
            await interaction.followup.send(f"I can't send messages in {target.mention}.", ephemeral=True)
            return

        try:
            await target.send(message)
            suffix = f"to {target.mention}" if target != interaction.channel else ""
            await interaction.followup.send(f"Message sent{' ' + suffix if suffix else ''}!", ephemeral=True)
            print(f"{interaction.user} used /msg in {interaction.guild.name}: '{message}' → #{target.name}")
        except Exception as e:
            await interaction.followup.send(f"Error sending message: {e}", ephemeral=True)


async def setup(bot: commands.Bot) -> None:
    await bot.add_cog(Admin(bot))
