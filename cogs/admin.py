"""
cogs/admin.py — Admin-only utility commands.

Commands: /help, /gettemplate, /serviceaccount, /sheetinfo, /sync,
          /reloadcommands, /cacheinfo, /clearcache, /set_runner_all_access,
          /refreshnicknames, /msg,
          /enterfolder, /syncrooms, /organizeroles, /sortrooms,
          /assignrunnerzones, /rolereset, /resetserver,
          /sendallmaterials, /sendsingularmaterial

Events:   on_guild_join, on_guild_remove
Tasks:    auto_leave_old_guilds (24 hr)
"""

from __future__ import annotations

import asyncio
import json
import os
import random
from collections import defaultdict
from datetime import datetime, timezone

import discord
from discord import app_commands
from discord.ext import commands, tasks
from googleapiclient.discovery import build

import config
import guild_setup
from utils import handle_rate_limit


def _admin_only(interaction: discord.Interaction) -> bool:
    return interaction.user.guild_permissions.administrator


# ── K-means (pure, synchronous) ───────────────────────────────────────────────

def _run_kmeans_clustering(points: list[tuple[float, float]], k: int, max_iterations: int = 100) -> list[int]:
    """Simple K-means++ clustering on 2-D points. Returns a label per point (0-indexed)."""
    if not points:
        return []
    if k <= 0:
        return [0] * len(points)
    if k >= len(points):
        return list(range(len(points)))

    random.seed(42)
    centroids = [list(random.choice(points))]
    for _ in range(k - 1):
        max_dist = 0.0
        farthest = None
        for p in points:
            d = min((p[0] - c[0]) ** 2 + (p[1] - c[1]) ** 2 for c in centroids)
            if d > max_dist:
                max_dist = d
                farthest = p
        centroids.append(list(farthest) if farthest else list(random.choice(points)))

    def _dsq(a: list, b: list) -> float:
        return (a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2

    labels = [0] * len(points)
    for _ in range(max_iterations):
        changed = False
        for i, p in enumerate(points):
            best = min(range(k), key=lambda j: _dsq(list(p), centroids[j]))
            if labels[i] != best:
                labels[i] = best
                changed = True

        sums = [[0.0, 0.0, 0] for _ in range(k)]
        for i, p in enumerate(points):
            c = labels[i]
            sums[c][0] += p[0]
            sums[c][1] += p[1]
            sums[c][2] += 1

        for j in range(k):
            if sums[j][2] > 0:
                centroids[j][0] = sums[j][0] / sums[j][2]
                centroids[j][1] = sums[j][1] / sums[j][2]
            else:
                centroids[j] = list(random.choice(points))

        if not changed:
            break

    return labels


# ── Cog ───────────────────────────────────────────────────────────────────────

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
            embed.add_field(name="Sync Status", value="Syncing every 60 minutes automatically", inline=False)
            embed.set_footer(text="Use /sync to manually trigger a sync")
        except Exception as e:
            embed = discord.Embed(title="Error Reading Sheet", description=str(e), color=discord.Color.red())

        await interaction.followup.send(embed=embed, ephemeral=True)

    # ── /enterfolder ──────────────────────────────────────────────────────────

    @app_commands.command(name="enterfolder", description="Connect a Google Drive folder to sync users from")
    @app_commands.describe(
        folder_link="Google Drive folder link (use 'Copy link' from Share dialog)",
        main_sheet_name="Name of the main sheet (e.g., '[TEMPLATE] Socal State')",
    )
    async def enter_folder_command(
        self,
        interaction: discord.Interaction,
        folder_link: str,
        main_sheet_name: str,
    ) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        if self._lock.locked():
            await interaction.response.send_message("Server configuration is in progress. Try again shortly.", ephemeral=True)
            return

        # Extract folder ID
        if "drive.google.com/drive/folders/" not in folder_link:
            await interaction.response.send_message(
                "Please provide a valid Google Drive folder link.\n\n"
                "How to get the correct link:\n"
                "1. Right-click your folder in Google Drive\n"
                "2. Click 'Share' → 'Copy link'\n"
                "3. Paste that link here (not the address bar URL)",
                ephemeral=True,
            )
            return

        try:
            folder_id = folder_link.split("/folders/")[1].split("?")[0]
        except (IndexError, AttributeError):
            await interaction.response.send_message("Could not parse folder ID from that link.", ephemeral=True)
            return

        async with self._lock:
            await interaction.response.defer(ephemeral=True)

            sc = self.bot.sheets_client
            if sc is None:
                await interaction.followup.send("Sheets client not available — check secrets/gspread.json.", ephemeral=True)
                return

            # Find the sheet in the folder
            try:
                found_sheet = sc.find_sheet_in_folder(folder_id, main_sheet_name)
            except Exception as e:
                msg = str(e)
                if "403" in msg or "permission" in msg.lower():
                    await interaction.followup.send(
                        f"Permission error — share the folder with `{config.SERVICE_EMAIL}` (Editor).",
                        ephemeral=True,
                    )
                else:
                    await interaction.followup.send(f"Error searching for sheet: {msg}", ephemeral=True)
                return

            if found_sheet is None:
                await interaction.followup.send(
                    f"Could not find '{main_sheet_name}' in that folder.\n\n"
                    f"Make sure the sheet name matches exactly and the folder is shared with:\n`{config.SERVICE_EMAIL}`",
                    ephemeral=True,
                )
                return

            guild_id = interaction.guild.id
            guild = interaction.guild

            # Connect worksheet
            try:
                try:
                    ws = found_sheet.worksheet(config.SHEET_PAGE_NAME)
                except Exception:
                    ws = found_sheet.worksheets()[0]

                test_data = ws.get_all_records()
            except Exception as e:
                await interaction.followup.send(f"Error accessing sheet data: {e}", ephemeral=True)
                return

            # Store connection
            self.bot.spreadsheets[guild_id] = found_sheet

            # Extract chapters
            chapters: set[str] = set()
            for row in test_data:
                chapter = str(row.get("Chapter", "")).strip()
                if chapter and chapter.lower() not in ("n/a", "na", ""):
                    chapters.add(chapter)
                else:
                    chapters.add("Unaffiliated")

            # Build structural channels
            try:
                await guild_setup.generate_building_structures(
                    guild,
                    runner_access=self.bot.runner_all_access,
                    spreadsheets=self.bot.spreadsheets,
                    force_refresh_welcome=False,
                )
                for chapter in chapters:
                    await guild_setup.setup_chapter_structure(
                        guild, chapter, self.bot.chapter_role_names,
                        runner_access=self.bot.runner_all_access,
                    )
                await guild_setup.sort_chapter_channels_alphabetically(guild)
            except Exception as e:
                print(f"Warning: error creating structures during /enterfolder: {e}")

            await guild_setup.setup_ezhang_admin_role(guild)

            # Immediate member sync
            sync_results = None
            try:
                onboarding = self.bot.get_cog("Onboarding")
                if onboarding:
                    sync_results = await onboarding.perform_member_sync(guild, test_data)
            except Exception as e:
                print(f"Warning: initial sync error in /enterfolder: {e}")

            # Cache the connection
            sc.save_guild_to_cache(guild_id, found_sheet.id, ws.title)

            embed = discord.Embed(
                title="Sheet Connected & Synced!",
                description=(
                    f"Connected to: **{found_sheet.title}**\n"
                    f"Worksheet: **{ws.title}**\n"
                    f"Rows: {len(test_data)}"
                ),
                color=discord.Color.green(),
            )
            if sync_results:
                embed.add_field(
                    name="Initial Sync",
                    value=(
                        f"• {sync_results['processed']} Discord IDs processed\n"
                        f"• {sync_results['role_assignments']} roles assigned"
                    ),
                    inline=False,
                )
            embed.set_footer(text="Use /sync to manually trigger another sync")
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

    # ── /syncrooms ────────────────────────────────────────────────────────────

    @app_commands.command(name="syncrooms", description="Regenerate building/room channels from Room Assignments (Admin only)")
    async def sync_rooms_command(self, interaction: discord.Interaction) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        if self._lock.locked():
            await interaction.response.send_message("Server configuration is in progress. Try again shortly.", ephemeral=True)
            return

        async with self._lock:
            await interaction.response.defer(ephemeral=True)

            guild_id = interaction.guild.id
            if guild_id not in self.bot.spreadsheets:
                await interaction.followup.send("No spreadsheet connected. Use `/enterfolder` first.", ephemeral=True)
                return

            await interaction.followup.send("Reading Room Assignments and regenerating channels... this may take a minute.", ephemeral=True)

            try:
                num_structures, num_buildings = await guild_setup.generate_building_structures(
                    interaction.guild,
                    runner_access=self.bot.runner_all_access,
                    spreadsheets=self.bot.spreadsheets,
                    force_refresh_welcome=True,
                )
                embed = discord.Embed(
                    title="Rooms Synced!",
                    description=(
                        f"**Event channels:** {num_structures}\n"
                        f"**Buildings:** {num_buildings}\n\n"
                        "Building chats verified and welcome messages refreshed."
                    ),
                    color=discord.Color.green(),
                )
                await interaction.followup.send(embed=embed, ephemeral=True)
            except Exception as e:
                await interaction.followup.send(f"Error syncing rooms: {e}", ephemeral=True)
                print(f"Error in /syncrooms: {e}")

    # ── /organizeroles ────────────────────────────────────────────────────────

    @app_commands.command(name="organizeroles", description="Organize server roles in priority order (Admin only)")
    async def organize_roles_command(self, interaction: discord.Interaction) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        if self._lock.locked():
            await interaction.response.send_message("Server configuration is in progress. Try again shortly.", ephemeral=True)
            return

        async with self._lock:
            await interaction.response.defer(ephemeral=True)

            if not interaction.guild.me.guild_permissions.manage_roles:
                embed = discord.Embed(
                    title="Missing Permissions!",
                    description="Bot cannot organize roles — it lacks 'Manage Roles' permission.",
                    color=discord.Color.red(),
                )
                embed.add_field(
                    name="How to Fix",
                    value="1. Go to Server Settings → Roles\n2. Find the bot's role\n3. Enable 'Manage Roles'\n4. Try again",
                    inline=False,
                )
                await interaction.followup.send(embed=embed, ephemeral=True)
                return

            try:
                bot_role = next(
                    (r for r in interaction.guild.roles if r.managed and interaction.guild.me in r.members),
                    None,
                )
                await guild_setup.organize_role_hierarchy_for_guild(
                    interaction.guild,
                    chapter_role_names=self.bot.chapter_role_names,
                )
                higher_roles = [
                    r for r in interaction.guild.roles
                    if r.position >= (bot_role.position if bot_role else 0)
                    and r.name != "@everyone"
                    and r != bot_role
                ]

                if higher_roles:
                    embed = discord.Embed(
                        title="Partial Success",
                        description="Some roles organized; others couldn't be moved due to hierarchy restrictions.",
                        color=discord.Color.orange(),
                    )
                    embed.add_field(
                        name="Couldn't Move",
                        value=f"These roles are higher than the bot:\n• " + "\n• ".join(r.name for r in higher_roles[:5])
                              + (f"\n• ... and {len(higher_roles)-5} more" if len(higher_roles) > 5 else ""),
                        inline=False,
                    )
                    embed.add_field(
                        name="To Fix",
                        value=f"Drag **{bot_role.name if bot_role else 'bot role'}** to the top of Server Settings → Roles, then run again.",
                        inline=False,
                    )
                else:
                    embed = discord.Embed(
                        title="Roles Organized!",
                        description="Roles arranged: other (alpha) → chapters (alpha) → Volunteer → Lead ES → Social Media → Photographer → Arbitrations → Awards → Runner → VIPer → Admin",
                        color=discord.Color.green(),
                    )
                embed.set_footer(text="Role organization complete!")
                await interaction.followup.send(embed=embed, ephemeral=True)
            except Exception as e:
                await interaction.followup.send(f"Error organizing roles: {e}", ephemeral=True)

    # ── /sortrooms ────────────────────────────────────────────────────────────

    @app_commands.command(name="sortrooms", description="Sort building categories and event channels alphabetically (Admin only)")
    async def sort_rooms_command(self, interaction: discord.Interaction) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        if self._lock.locked():
            await interaction.response.send_message("Server configuration is in progress. Try again shortly.", ephemeral=True)
            return

        async with self._lock:
            await interaction.response.defer(ephemeral=True)

            try:
                await interaction.followup.send("Sorting building categories and inner channels...", ephemeral=True)
                await guild_setup.sort_building_categories_alphabetically(interaction.guild)
                await guild_setup.sort_channels_in_building_categories(interaction.guild)
                embed = discord.Embed(
                    title="Channels Sorted!",
                    description="All building categories and event channels sorted alphabetically.\nBuilding chats pinned to top of each category.",
                    color=discord.Color.green(),
                )
                await interaction.followup.send(embed=embed, ephemeral=True)
            except Exception as e:
                await interaction.followup.send(f"Error sorting channels: {e}", ephemeral=True)

    # ── /assignrunnerzones ────────────────────────────────────────────────────

    @app_commands.command(name="assignrunnerzones", description="Assign zone numbers in Runner Assignments using K-means (Admin only)")
    async def assign_runner_zones_command(self, interaction: discord.Interaction) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        if self._lock.locked():
            await interaction.response.send_message("Server configuration is in progress. Try again shortly.", ephemeral=True)
            return

        async with self._lock:
            await interaction.response.defer(ephemeral=True)

            guild_id = interaction.guild.id
            if guild_id not in self.bot.spreadsheets:
                await interaction.followup.send("No spreadsheet connected. Use `/enterfolder` first.", ephemeral=True)
                return

            sc = self.bot.sheets_client
            if sc is None:
                await interaction.followup.send("Sheets client unavailable.", ephemeral=True)
                return

            spreadsheet = self.bot.spreadsheets[guild_id]
            worksheet_name = "Runner Assignments"
            ws = None

            try:
                ws = spreadsheet.worksheet(worksheet_name)
            except Exception:
                # Search parent Drive folder for a separate Runner Assignments spreadsheet
                try:
                    drive = build("drive", "v3", credentials=sc.creds)
                    meta = drive.files().get(fileId=spreadsheet.id, fields="parents").execute()
                    parents = meta.get("parents", [])
                    if not parents:
                        await interaction.followup.send("Could not determine parent folder to search for Runner Assignments.", ephemeral=True)
                        return
                    q = (
                        f"'{parents[0]}' in parents"
                        " and mimeType='application/vnd.google-apps.spreadsheet'"
                        f" and name contains '{worksheet_name}'"
                    )
                    files = drive.files().list(q=q, fields="files(id, name)").execute().get("files", [])
                    if not files:
                        await interaction.followup.send(f"Could not find '{worksheet_name}' in the same Drive folder.", ephemeral=True)
                        return
                    other = sc.open_by_key(files[0]["id"])
                    try:
                        ws = other.worksheet(worksheet_name)
                    except Exception:
                        ws = other.worksheets()[0]
                except Exception as e:
                    await interaction.followup.send(f"Error locating Runner Assignments: {e}", ephemeral=True)
                    return

            try:
                headers = ws.row_values(1)
                rows = ws.get_all_records()
            except Exception as e:
                await interaction.followup.send(f"Could not read Runner Assignments: {e}", ephemeral=True)
                return

            def _col(candidates: list[str]) -> int | None:
                for i, h in enumerate(headers):
                    if h and h.strip().lower() in candidates:
                        return i + 1
                return None

            zones_col = _col(["zone number"])
            if zones_col is None:
                new_idx = len(headers) + 1
                col_letter = chr(ord("A") + new_idx - 1)
                try:
                    ws.update(f"{col_letter}1", [["zone number"]])
                    headers.append("zone number")
                    zones_col = new_idx
                except Exception as e:
                    await interaction.followup.send(f"Could not create 'Zone Number' column: {e}", ephemeral=True)
                    return

            # Parse global K
            global_k = 1
            for row in rows:
                lr = {(k.strip().lower() if isinstance(k, str) else k): v for k, v in row.items()}
                raw = lr.get("number of zones") or lr.get("zones count") or lr.get("num zones") or lr.get("k")
                if raw is not None and str(raw).strip():
                    try:
                        global_k = max(1, int(float(raw)))
                        break
                    except Exception:
                        continue

            def _float(val) -> float | None:
                try:
                    s = str(val).strip()
                    return float(s) if s else None
                except Exception:
                    return None

            # Collect all points
            all_items: list[tuple[str, int, tuple[float, float]]] = []
            for idx, row in enumerate(rows, start=2):
                lr = {(k.strip().lower() if isinstance(k, str) else k): v for k, v in row.items()}
                building = str(lr.get("building", lr.get("building 1", ""))).strip()
                if not building:
                    continue
                lat = _float(lr.get("latitude") or lr.get("lat"))
                lon = _float(lr.get("longitude") or lr.get("lon") or lr.get("lng"))
                if (lat is None or lon is None) and lr.get("coordinates"):
                    parts = str(lr["coordinates"]).split(",")
                    if len(parts) >= 2:
                        lat = lat or _float(parts[0])
                        lon = lon or _float(parts[1])
                if lat is None or lon is None:
                    continue
                all_items.append((building, idx, (lat, lon)))

            if not all_items:
                await interaction.followup.send("No valid location rows found to cluster.", ephemeral=True)
                return

            labels = _run_kmeans_clustering([p for _, _, p in all_items], global_k)
            zones_col_letter = chr(ord("A") + zones_col - 1)
            updated = 0
            for i, (_, row_idx, _) in enumerate(all_items):
                try:
                    ws.update(f"{zones_col_letter}{row_idx}", [[str(labels[i] + 1)]])
                    updated += 1
                except Exception:
                    pass

            await interaction.followup.send(
                f"Assigned {global_k} zones for {updated} rows across {len({b for b, *_ in all_items})} buildings.\n\nSending runner assignments to building channels...",
                ephemeral=True,
            )

            # Post runner embeds to building channels
            try:
                await self._post_runner_zone_assignments(interaction.guild, rows, guild_id)
            except Exception as e:
                print(f"Error posting runner zone assignments: {e}")

    async def _post_runner_zone_assignments(
        self,
        guild: discord.Guild,
        rows: list[dict],
        guild_id: int,
    ) -> None:
        """Post runner zone assignment embeds to each building chat."""
        main_sheet = self._sheet(guild_id)
        email_to_discord: dict[str, int] = {}
        if main_sheet:
            try:
                for row in main_sheet.get_all_records():
                    email = str(row.get("Email", "")).strip().lower()
                    raw_id = str(row.get("Discord ID", "")).strip()
                    if email and raw_id:
                        try:
                            email_to_discord[email] = int(raw_id)
                        except ValueError:
                            pass
            except Exception as e:
                print(f"Could not read main sheet for Discord IDs: {e}")

        building_zones: dict[str, int] = {}
        zone_runners: dict[int, list[tuple[str, int | None]]] = defaultdict(list)

        for row in rows:
            lr = {(k.strip().lower() if isinstance(k, str) else k): v for k, v in row.items()}
            building = str(lr.get("building", lr.get("building 1", ""))).strip()
            zone_raw = str(lr.get("zone number", lr.get("zone", ""))).strip()
            if building and zone_raw:
                try:
                    building_zones[building] = int(zone_raw)
                except ValueError:
                    pass
            name = str(lr.get("name", "")).strip()
            email = str(lr.get("email", "")).strip().lower()
            runner_zone_raw = str(lr.get("runner zone", "")).strip()
            if name and runner_zone_raw:
                try:
                    zone_runners[int(runner_zone_raw)].append((name, email_to_discord.get(email)))
                except ValueError:
                    pass

        messages_sent = 0
        for building, zone_num in building_zones.items():
            chat_name = f"{guild_setup.sanitize_for_discord(building)}-chat"
            ch = discord.utils.get(guild.text_channels, name=chat_name)
            if not ch:
                continue
            runners = zone_runners.get(zone_num, [])
            if not runners:
                continue
            runner_lines = [
                f"• <@{did}>" if did else f"• {name}"
                for name, did in runners
            ]
            embed = discord.Embed(
                title=f"Designated Runners for {building}",
                description="Runners assigned to help with this building:",
                color=discord.Color.orange(),
            )
            embed.add_field(name="Runners", value="\n".join(runner_lines), inline=False)
            embed.set_footer(text="If you need help, create a ticket in the #help forum! DM these runners for urgent help!")
            try:
                await ch.send(embed=embed)
                messages_sent += 1
            except Exception as e:
                print(f"Error sending runner assignment to {chat_name}: {e}")

        print(f"Sent runner assignments to {messages_sent} building channel(s)")

    # ── /rolereset ────────────────────────────────────────────────────────────

    @app_commands.command(name="rolereset", description="Delete stale roles, rebuild structures, and re-sync (Admin only)")
    async def role_reset_command(self, interaction: discord.Interaction) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        if self._lock.locked():
            await interaction.response.send_message("Server configuration is in progress. Try again shortly.", ephemeral=True)
            return

        async with self._lock:
            await interaction.response.defer(ephemeral=True)

            guild = interaction.guild
            guild_id = guild.id
            sheet = self._sheet(guild_id)
            if sheet is None:
                await interaction.followup.send("No sheet connected. Use `/enterfolder` first.", ephemeral=True)
                return

            try:
                test_data = sheet.get_all_records()
            except Exception as e:
                await interaction.followup.send(f"Could not read sheet: {e}", ephemeral=True)
                return

            # Compute which roles to keep
            event_list: set[str] = set()
            chapters: set[str] = set()
            for row in test_data:
                roles_raw = str(row.get("Roles", "")).strip()
                first_event = next((r.strip() for r in roles_raw.split(";") if r.strip()), "")
                if first_event:
                    event_list.add(first_event)
                chapter = str(row.get("Chapter", "")).strip()
                if chapter and chapter.lower() not in ("n/a", "na", ""):
                    chapters.add(chapter)
                else:
                    chapters.add("Unaffiliated")

            # Also protect roles that appear in Room Assignments
            try:
                import data_router
                room_data = await data_router.list_events(guild_id, spreadsheets=self.bot.spreadsheets)
                for row in room_data:
                    event = str(row.get("name", "")).strip()
                    if event:
                        event_list.add(event)
            except Exception as e:
                print(f"Warning: could not read Room Assignments during /rolereset: {e}")

            keep_roles = guild_setup.PRIORITY_ROLES | event_list | chapters | {"@everyone"}

            # Delete stale roles
            role_count = 0
            for role in list(guild.roles):
                if (
                    role.name not in keep_roles
                    and not role.managed
                    and role < guild.me.top_role
                ):
                    try:
                        await role.delete(reason=f"Role reset by {interaction.user}")
                        role_count += 1
                        await asyncio.sleep(0.5)
                    except discord.Forbidden:
                        pass
                    except Exception as e:
                        print(f"Error deleting role '{role.name}': {e}")

            # Rebuild structures
            try:
                await guild_setup.generate_building_structures(
                    guild,
                    runner_access=self.bot.runner_all_access,
                    spreadsheets=self.bot.spreadsheets,
                    force_refresh_welcome=True,
                )
                for chapter in chapters:
                    await guild_setup.setup_chapter_structure(
                        guild, chapter, self.bot.chapter_role_names,
                        runner_access=self.bot.runner_all_access,
                    )
                await guild_setup.sort_chapter_channels_alphabetically(guild)
            except Exception as e:
                print(f"Warning: structure error during /rolereset: {e}")

            await guild_setup.setup_ezhang_admin_role(guild)

            # Re-sync
            sync_results = None
            try:
                onboarding = self.bot.get_cog("Onboarding")
                if onboarding:
                    sync_results = await onboarding.perform_member_sync(guild, test_data)
            except Exception as e:
                print(f"Warning: sync error during /rolereset: {e}")

            embed = discord.Embed(
                title="Role Reset Complete!",
                description="Stale roles deleted, structures rebuilt, members re-synced.",
                color=discord.Color.green(),
            )
            embed.add_field(name="Roles Deleted", value=str(role_count), inline=True)
            if sync_results:
                embed.add_field(name="Roles Assigned", value=str(sync_results["role_assignments"]), inline=True)
            await interaction.followup.send(embed=embed, ephemeral=True)

    # ── /resetserver ─────────────────────────────────────────────────────────

    @app_commands.command(name="resetserver", description="DANGER: Delete all channels, roles, and categories (Admin only)")
    async def reset_server_command(self, interaction: discord.Interaction) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        if self._lock.locked():
            await interaction.response.send_message("Server configuration is in progress. Try again shortly.", ephemeral=True)
            return

        async with self._lock:
            await interaction.response.defer(ephemeral=True)
            try:
                await interaction.followup.send(
                    "Server reset starting in 3 seconds — this will delete ALL channels, categories, and non-managed roles.",
                    ephemeral=True,
                )
                await guild_setup.reset_server_for_guild(interaction.guild)
                embed = discord.Embed(
                    title="Server Reset Complete!",
                    description="All channels, categories, and non-managed roles have been deleted. Nicknames cleared.",
                    color=discord.Color.green(),
                )
                await interaction.followup.send(embed=embed, ephemeral=True)
            except Exception as e:
                await interaction.followup.send(f"Error during server reset: {e}", ephemeral=True)

    # ── /sendallmaterials ─────────────────────────────────────────────────────

    @app_commands.command(name="sendallmaterials", description="Send test materials to all event channels (Admin only)")
    async def send_all_materials_command(self, interaction: discord.Interaction) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        if self._lock.locked():
            await interaction.response.send_message("Server configuration is in progress. Try again shortly.", ephemeral=True)
            return

        async with self._lock:
            await interaction.response.defer(ephemeral=True)

            guild_id = interaction.guild.id
            if guild_id not in self.bot.spreadsheets:
                await interaction.followup.send("No spreadsheet connected. Use `/enterfolder` first.", ephemeral=True)
                return

            sc = self.bot.sheets_client
            if sc is None:
                await interaction.followup.send("Sheets client unavailable.", ephemeral=True)
                return

            guild = interaction.guild
            event_roles = [
                r.name for r in guild.roles
                if r.name != "@everyone"
                and not r.managed
                and r.name not in guild_setup.PRIORITY_ROLES
                and r.name not in self.bot.chapter_role_names
            ]

            if not event_roles:
                await interaction.followup.send("No event roles found. Run `/enterfolder` first.", ephemeral=True)
                return

            await interaction.followup.send(
                f"Processing test materials for {len(event_roles)} event(s)...\nCheck event channels for results.",
                ephemeral=True,
            )

            success = 0
            deleted = 0
            for role_name in event_roles:
                try:
                    sanitized = guild_setup.sanitize_for_discord(role_name)
                    for ch in guild.text_channels:
                        if ch.name.startswith(sanitized + "-"):
                            try:
                                for msg in await ch.pins():
                                    if msg.author == self.bot.user:
                                        await msg.delete()
                                        deleted += 1
                                        await asyncio.sleep(0.2)
                            except Exception:
                                pass
                    await self._search_and_share_test_folder(guild, role_name)
                    success += 1
                    await asyncio.sleep(0.5)
                except Exception as e:
                    print(f"Error sending materials for {role_name}: {e}")

            try:
                await self._search_and_share_useful_links(guild)
            except Exception as e:
                print(f"Error sharing useful links: {e}")

            try:
                await self._search_and_share_runner_info(guild)
            except Exception as e:
                print(f"Error sharing runner info: {e}")

            embed = discord.Embed(
                title="All Materials Sent!",
                description=f"Sent for **{success}/{len(event_roles)}** events.",
                color=discord.Color.green(),
            )
            if deleted:
                embed.add_field(name="Old Materials Cleared", value=f"Deleted {deleted} pinned message(s)", inline=False)
            await interaction.followup.send(embed=embed, ephemeral=True)

    # ── /sendsingularmaterial ─────────────────────────────────────────────────

    @app_commands.command(name="sendsingularmaterial", description="Send materials for a specific type (Admin only)")
    @app_commands.describe(
        material_type="Type of material to send",
        event_name="Event name (required when material_type is 'event')",
    )
    @app_commands.choices(material_type=[
        app_commands.Choice(name="Event Test Materials", value="event"),
        app_commands.Choice(name="Useful Links", value="useful-links"),
        app_commands.Choice(name="Runner Info", value="runner"),
    ])
    async def send_singular_material_command(
        self,
        interaction: discord.Interaction,
        material_type: app_commands.Choice[str],
        event_name: str | None = None,
    ) -> None:
        if not _admin_only(interaction):
            await interaction.response.send_message("You need administrator permissions.", ephemeral=True)
            return
        if self._lock.locked():
            await interaction.response.send_message("Server configuration is in progress. Try again shortly.", ephemeral=True)
            return

        async with self._lock:
            await interaction.response.defer(ephemeral=True)

            guild_id = interaction.guild.id
            if guild_id not in self.bot.spreadsheets:
                await interaction.followup.send("No spreadsheet connected. Use `/enterfolder` first.", ephemeral=True)
                return

            sc = self.bot.sheets_client
            if sc is None:
                await interaction.followup.send("Sheets client unavailable.", ephemeral=True)
                return

            guild = interaction.guild
            value = material_type.value

            try:
                if value == "event":
                    if not event_name:
                        await interaction.followup.send("Event name is required when sending event materials.", ephemeral=True)
                        return
                    await interaction.followup.send(f"Searching for test materials for {event_name}...", ephemeral=True)
                    await self._search_and_share_test_folder(guild, event_name)

                elif value == "useful-links":
                    await interaction.followup.send("Searching for useful links...", ephemeral=True)
                    await self._search_and_share_useful_links(guild)

                elif value == "runner":
                    await interaction.followup.send("Searching for runner info...", ephemeral=True)
                    await self._search_and_share_runner_info(guild)

                embed = discord.Embed(
                    title="Material Sent!",
                    description=f"Successfully sent {material_type.name}.",
                    color=discord.Color.green(),
                )
                await interaction.followup.send(embed=embed, ephemeral=True)
            except Exception as e:
                await interaction.followup.send(f"Error sending materials: {e}", ephemeral=True)

    # ── Drive API material helpers ─────────────────────────────────────────────

    async def _search_and_share_test_folder(self, guild: discord.Guild, role_name: str) -> None:
        """Share test materials for *role_name* from the Drive Tests folder."""
        sc = self.bot.sheets_client
        if sc is None:
            return
        guild_id = guild.id
        if guild_id not in self.bot.spreadsheets:
            return

        spreadsheet = self.bot.spreadsheets[guild_id]
        drive = build("drive", "v3", credentials=sc.creds)

        # Get parent folder
        meta = drive.files().get(fileId=spreadsheet.id, fields="parents").execute()
        parents = meta.get("parents", [])
        if not parents:
            return
        parent_id = parents[0]

        # Find Tests folder
        tests_q = f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='Tests'"
        tests = drive.files().list(q=tests_q, fields="files(id, name)").execute().get("files", [])
        if not tests:
            print(f"No 'Tests' folder found for {role_name}")
            return
        tests_id = tests[0]["id"]

        # Find event subfolder
        event_q = f"'{tests_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='{role_name}'"
        event_folders = drive.files().list(q=event_q, fields="files(id, name, webViewLink)").execute().get("files", [])
        if not event_folders:
            print(f"No test folder for '{role_name}'")
            return
        event_id = event_folders[0]["id"]

        files = drive.files().list(
            q=f"'{event_id}' in parents and trashed=false",
            fields="files(id, name, webViewLink, mimeType)",
        ).execute().get("files", [])
        if not files:
            return

        # Find event channel
        target = next(
            (ch for ch in guild.text_channels
             if role_name.lower().replace(" ", "-") in ch.name.lower()
             and ch.category and ch.category.name not in guild_setup.STATIC_CATEGORIES),
            None,
        )
        if not target:
            print(f"No channel found for event '{role_name}'")
            return

        # Check if already pinned
        pinned = await target.pins()
        if any(m.embeds and m.embeds[0].title and f"Test Materials for {role_name}" in m.embeds[0].title for m in pinned):
            return

        _MIME_EMOJI = {
            "pdf": "📄", "document": "📝", "spreadsheet": "📊",
            "presentation": "📖", "image": "🖼️", "folder": "📁",
        }
        links = [
            f"• {next((e for k, e in _MIME_EMOJI.items() if k in f.get('mimeType','')), '📎')} "
            f"[**{f['name']}**]({f['webViewLink']})"
            for f in files
        ]

        # Chunk if needed
        chunks: list[str] = []
        current = ""
        for link in links:
            if len(current + link + "\n") > 1000:
                if current:
                    chunks.append(current.strip())
                current = link + "\n"
            else:
                current += link + "\n"
        if current:
            chunks.append(current.strip())

        embed = discord.Embed(
            title=f"Test Materials for {role_name}",
            description="Access your event-specific test materials. Do NOT share with anyone outside your event.",
            color=discord.Color.green(),
        )
        embed.add_field(name="Test Materials", value=chunks[0] if chunks else "No files found", inline=False)
        msg = await target.send(embed=embed)
        try:
            await msg.pin()
        except Exception:
            pass

        for i, chunk in enumerate(chunks[1:], start=2):
            cont = discord.Embed(
                title=f"Test Materials for {role_name} (continued {i})",
                color=discord.Color.green(),
            )
            cont.add_field(name="Test Materials", value=chunk, inline=False)
            await target.send(embed=cont)
            await asyncio.sleep(0.5)

        # Scoring instructions
        if not any(m.embeds and m.embeds[0].title and "Score Input Instructions" in m.embeds[0].title for m in pinned):
            scoring = discord.Embed(
                title="Score Input Instructions",
                description="**IMPORTANT**: All Lead Event Supervisors must input scores through the official scoring portal!",
                color=discord.Color.blue(),
            )
            scoring.add_field(
                name="Scoring Portal",
                value="[Click here to access the scoring system](https://scoring.duosmium.org/login)",
                inline=False,
            )
            scoring.add_field(
                name="Instructions",
                value=(
                    "• Lead ES should have received a scoring portal invitation email\n"
                    "• Select the correct tournament and event\n"
                    "• Input all team scores accurately\n"
                    "• Contact admin if you have login issues"
                ),
                inline=False,
            )
            sm = await target.send(embed=scoring)
            try:
                await sm.pin()
            except Exception:
                pass

        print(f"Shared test materials for '{role_name}' in #{target.name}")

    async def _search_and_share_useful_links(self, guild: discord.Guild) -> None:
        """Share files from the Drive 'Useful Links' folder into #useful-links."""
        guild_id = guild.id
        if guild_id not in self.bot.spreadsheets:
            return
        sc = self.bot.sheets_client
        if sc is None:
            return

        try:
            spreadsheet = self.bot.spreadsheets[guild_id]
            drive = build("drive", "v3", credentials=sc.creds)

            meta = drive.files().get(fileId=spreadsheet.id, fields="parents").execute()
            parents = meta.get("parents", [])
            if not parents:
                return
            parent_id = parents[0]

            folders = drive.files().list(
                q=f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='Useful Links'",
                fields="files(id, name)",
            ).execute().get("files", [])
            if not folders:
                print("No 'Useful Links' folder found in Drive folder")
                return
            folder_id = folders[0]["id"]

            files = drive.files().list(
                q=f"'{folder_id}' in parents and trashed=false",
                fields="files(id, name, webViewLink, mimeType)",
            ).execute().get("files", [])
            if not files:
                return

            target = discord.utils.get(guild.text_channels, name="useful-links")
            if not target:
                print("No #useful-links channel found")
                return

            # Clear old pinned useful-links messages from the bot
            for msg in await target.pins():
                if msg.author == self.bot.user and msg.embeds:
                    if msg.embeds[0].title and "Useful Links" in msg.embeds[0].title:
                        try:
                            await msg.delete()
                            await asyncio.sleep(0.2)
                        except Exception:
                            pass

            _MIME_EMOJI = {
                "pdf": "📄", "document": "📝", "spreadsheet": "📊",
                "presentation": "📖", "image": "🖼️", "folder": "📁",
            }
            links = [
                f"• {next((e for k, e in _MIME_EMOJI.items() if k in f.get('mimeType', '')), '📎')} "
                f"[**{f['name']}**]({f['webViewLink']})"
                for f in files
            ]

            chunks: list[str] = []
            current = ""
            for link in links:
                if len(current + link + "\n") > 1000:
                    if current:
                        chunks.append(current.strip())
                    current = link + "\n"
                else:
                    current += link + "\n"
            if current:
                chunks.append(current.strip())

            embed = discord.Embed(
                title="Useful Links & Resources",
                description="Important links and resources for volunteers!",
                color=discord.Color.green(),
            )
            embed.add_field(name="Useful Links", value=chunks[0] if chunks else "No files found", inline=False)
            msg = await target.send(embed=embed)
            try:
                await msg.pin()
            except Exception:
                pass

            for i, chunk in enumerate(chunks[1:], start=2):
                cont = discord.Embed(
                    title=f"Useful Links & Resources (continued {i})",
                    color=discord.Color.green(),
                )
                cont.add_field(name="Useful Links", value=chunk, inline=False)
                await target.send(embed=cont)
                await asyncio.sleep(0.5)

            print(f"Shared useful links in #{target.name}")
        except Exception as e:
            print(f"Error sharing useful links: {e}")

    async def _search_and_share_runner_info(self, guild: discord.Guild) -> None:
        """Share files from the Drive 'Runner' folder into #runner."""
        guild_id = guild.id
        if guild_id not in self.bot.spreadsheets:
            return
        sc = self.bot.sheets_client
        if sc is None:
            return

        try:
            spreadsheet = self.bot.spreadsheets[guild_id]
            drive = build("drive", "v3", credentials=sc.creds)

            meta = drive.files().get(fileId=spreadsheet.id, fields="parents").execute()
            parents = meta.get("parents", [])
            if not parents:
                return
            parent_id = parents[0]

            folders = drive.files().list(
                q=f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='Runner'",
                fields="files(id, name)",
            ).execute().get("files", [])
            if not folders:
                print("No 'Runner' folder found in Drive folder")
                return
            folder_id = folders[0]["id"]

            files = drive.files().list(
                q=f"'{folder_id}' in parents and trashed=false",
                fields="files(id, name, webViewLink, mimeType)",
            ).execute().get("files", [])
            if not files:
                return

            target = discord.utils.get(guild.text_channels, name="runner")
            if not target:
                print("No #runner channel found")
                return

            # Clear old pinned runner info messages from the bot
            for msg in await target.pins():
                if msg.author == self.bot.user and msg.embeds:
                    if msg.embeds[0].title and "Runner Information" in msg.embeds[0].title:
                        try:
                            await msg.delete()
                            await asyncio.sleep(0.2)
                        except Exception:
                            pass

            _MIME_EMOJI = {
                "pdf": "📄", "document": "📝", "spreadsheet": "📊",
                "presentation": "📖", "image": "🖼️", "folder": "📁",
            }
            links = [
                f"• {next((e for k, e in _MIME_EMOJI.items() if k in f.get('mimeType', '')), '📎')} "
                f"[**{f['name']}**]({f['webViewLink']})"
                for f in files
            ]

            chunks: list[str] = []
            current = ""
            for link in links:
                if len(current + link + "\n") > 1000:
                    if current:
                        chunks.append(current.strip())
                    current = link + "\n"
                else:
                    current += link + "\n"
            if current:
                chunks.append(current.strip())

            embed = discord.Embed(
                title="Runner Information & Resources",
                description="Important information and resources for runners!",
                color=discord.Color.blue(),
            )
            embed.add_field(name="Runner Info", value=chunks[0] if chunks else "No files found", inline=False)
            msg = await target.send(embed=embed)
            try:
                await msg.pin()
            except Exception:
                pass

            for i, chunk in enumerate(chunks[1:], start=2):
                cont = discord.Embed(
                    title=f"Runner Information & Resources (continued {i})",
                    color=discord.Color.blue(),
                )
                cont.add_field(name="Runner Info", value=chunk, inline=False)
                await target.send(embed=cont)
                await asyncio.sleep(0.5)

            print(f"Shared runner info in #{target.name}")
        except Exception as e:
            print(f"Error sharing runner info: {e}")

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

    # ── Guild lifecycle ───────────────────────────────────────────────────────

    async def cog_load(self) -> None:
        self.auto_leave_old_guilds.start()

    async def cog_unload(self) -> None:
        self.auto_leave_old_guilds.cancel()

    def _cleanup_guild(self, guild_id: int) -> None:
        """Remove all in-memory and cached state for a guild."""
        self.bot.spreadsheets.pop(guild_id, None)
        self.bot.sheets.pop(guild_id, None)
        self.bot.runner_all_access.pop(guild_id, None)
        sc = self.bot.sheets_client
        if sc:
            sc.clear_guild_from_cache(guild_id)

    @commands.Cog.listener()
    async def on_guild_join(self, guild: discord.Guild) -> None:
        """Run standard setup when the bot joins a new guild."""
        print(f"Bot joined new guild: {guild.name} ({guild.id}) — {guild.member_count} members")

        async with self._lock:
            try:
                # Remove the default "general" channel Discord creates
                for ch in guild.channels:
                    if ch.name.lower() == "general":
                        try:
                            await ch.delete(reason="Removing default Discord channel")
                        except Exception:
                            pass

                await guild_setup.setup_static_channels_for_guild(
                    guild,
                    bot_user=self.bot.user,
                    runner_access=self.bot.runner_all_access,
                )
                await guild_setup.move_bot_role_to_top_for_guild(guild)
                await guild_setup.organize_role_hierarchy_for_guild(
                    guild, chapter_role_names=self.bot.chapter_role_names
                )
                if not self.bot.runner_all_access.get(guild.id, 0):
                    await guild_setup.remove_runner_access_from_building_channels_for_guild(guild)
                await guild_setup.give_runner_access_to_all_channels_for_guild(guild)
                await guild_setup.setup_ezhang_admin_role(guild)

                print(f"Setup complete for new guild: {guild.name}")
            except Exception as e:
                print(f"Error setting up new guild {guild.name}: {e}")

    @commands.Cog.listener()
    async def on_guild_remove(self, guild: discord.Guild) -> None:
        print(f"Removed from guild '{guild.name}' ({guild.id}) — cleaning up state")
        self._cleanup_guild(guild.id)

    @tasks.loop(hours=24)
    async def auto_leave_old_guilds(self) -> None:
        """Leave any guild the bot has been in for more than 30 days."""
        now = datetime.now(timezone.utc)
        for guild in list(self.bot.guilds):
            joined_at = guild.me.joined_at if guild.me else None
            if joined_at is None:
                continue
            age_days = (now - joined_at).days
            if age_days < 30:
                continue
            print(f"auto_leave: '{guild.name}' ({guild.id}) joined {age_days}d ago — leaving")
            try:
                await guild.leave()
                self._cleanup_guild(guild.id)
                print(f"auto_leave: left '{guild.name}'")
            except Exception as e:
                print(f"auto_leave: could not leave '{guild.name}': {e}")

    @auto_leave_old_guilds.before_loop
    async def before_auto_leave(self) -> None:
        await self.bot.wait_until_ready()


async def setup(bot: commands.Bot) -> None:
    await bot.add_cog(Admin(bot))
