"""
cogs/tickets.py — Help ticket creation, runner pinging, and re-ping tracking.

Owns:
  - active_help_tickets   dict: thread_id -> ticket metadata
  - active_burger_deliveries  dict: user_id -> delivery state

Depends on (via self.bot):
  - self.bot.spreadsheets  dict: guild_id -> gspread Spreadsheet
  - self.bot.gc            gspread client
  - self.bot.creds         oauth2client credentials (for Drive API)

Commands: /activetickets, /stopburgers, /debugzone
Events:   on_thread_create, on_message, on_reaction_add, on_thread_delete
Tasks:    check_help_tickets (every 1 min)
"""

import asyncio
import random
from datetime import datetime, timedelta

import discord
from discord import app_commands
from discord.ext import commands, tasks

import config


class Tickets(commands.Cog):
    def __init__(self, bot: commands.Bot) -> None:
        self.bot = bot
        # Ticket state owned by this cog
        self.active_help_tickets: dict = {}   # thread_id -> ticket_info
        self.active_burger_deliveries: dict = {}  # user_id -> {"stop": bool, "user": Member}

    async def cog_load(self) -> None:
        self.check_help_tickets.start()

    async def cog_unload(self) -> None:
        self.check_help_tickets.cancel()

    # ── Sheets helpers ────────────────────────────────────────────────────────

    async def _get_room_assignments(self, guild_id: int) -> list:
        if guild_id not in self.bot.spreadsheets:
            return []
        try:
            ws = self.bot.spreadsheets[guild_id].worksheet("Room Assignments")
            return ws.get_all_records()
        except Exception as e:
            print(f"'Room Assignments' worksheet not found or error: {e}")
            return []

    async def _get_runner_sheet(self, guild_id: int):
        """Return the Runner Assignments worksheet, searching Drive if needed."""
        spreadsheet = self.bot.spreadsheets.get(guild_id)
        if not spreadsheet:
            return None

        try:
            return spreadsheet.worksheet("Runner Assignments")
        except Exception:
            pass

        # Fall back to searching the parent Drive folder via sheets_client
        sc = self.bot.sheets_client
        if not sc:
            return None
        try:
            from googleapiclient.discovery import build
            drive = build("drive", "v3", credentials=sc.creds)
            meta = drive.files().get(fileId=spreadsheet.id, fields="parents").execute()
            parents = meta.get("parents", [])
            if not parents:
                return None

            q = (
                f"'{parents[0]}' in parents"
                " and mimeType='application/vnd.google-apps.spreadsheet'"
                " and name contains 'Runner Assignments'"
            )
            files = drive.files().list(q=q, fields="files(id, name)").execute().get("files", [])
            if not files:
                return None

            return sc.open_by_key(files[0]["id"]).sheet1
        except Exception as e:
            print(f"Error finding Runner Assignments spreadsheet: {e}")
            return None

    async def _get_user_event_building(self, guild_id: int, discord_id: int) -> dict | None:
        if guild_id not in self.bot.spreadsheets:
            print(f"No spreadsheet connected for guild {guild_id}")
            return None

        try:
            spreadsheet = self.bot.spreadsheets[guild_id]
            sheet = spreadsheet.worksheet(config.SHEET_PAGE_NAME)
            data = sheet.get_all_records()

            user_event = None
            user_name = None

            for row in data:
                row_discord_id = str(row.get("Discord ID", "")).strip()
                if not row_discord_id:
                    continue
                try:
                    if int(row_discord_id) == discord_id:
                        roles_raw = str(row.get("Roles", "")).strip()
                        roles = [r.strip() for r in roles_raw.split(";") if r.strip()]
                        user_event = roles[0] if roles else None
                        user_name = str(row.get("Name", "")).strip()
                        break
                except ValueError:
                    continue

            if not user_name:
                print(f"User with Discord ID {discord_id} not found in sheet")
                return None

            building = None
            room = None
            if user_event:
                room_data = await self._get_room_assignments(guild_id)
                for r_row in room_data:
                    if str(r_row.get("Events", "")).strip().lower() == user_event.lower():
                        building = str(r_row.get("Building", "")).strip()
                        room = str(r_row.get("Room", "")).strip()
                        break

            return {"event": user_event, "building": building, "room": room, "name": user_name}

        except Exception as e:
            print(f"Error looking up user event/building: {e}")
            return None

    async def _get_building_zone(self, guild_id: int, building: str) -> int | None:
        sheet = await self._get_runner_sheet(guild_id)
        if not sheet:
            return None
        try:
            for row in sheet.get_all_records():
                if str(row.get("Building", "")).strip().lower() == building.lower():
                    zone = row.get("Zone Number", "")
                    try:
                        return int(zone)
                    except (ValueError, TypeError):
                        print(f"Invalid zone value '{zone}' for building '{building}'")
                        return None
            print(f"Building '{building}' not found in Runner Assignments")
            return None
        except Exception as e:
            print(f"Error looking up building zone: {e}")
            return None

    async def _get_zone_runners(self, guild_id: int, zone: int) -> list[int]:
        sheet = await self._get_runner_sheet(guild_id)
        if not sheet:
            return []
        try:
            runner_emails = []
            for row in sheet.get_all_records():
                row_zone = row.get("Runner Zone", "")
                try:
                    if int(row_zone) == zone:
                        email = str(row.get("Email", "")).strip()
                        if email:
                            runner_emails.append(email.lower())
                except (ValueError, TypeError):
                    continue

            if not runner_emails:
                return []

            spreadsheet = self.bot.spreadsheets[guild_id]
            main_data = spreadsheet.worksheet(config.SHEET_PAGE_NAME).get_all_records()
            ids = []
            for row in main_data:
                if str(row.get("Email", "")).strip().lower() in runner_emails:
                    try:
                        ids.append(int(str(row.get("Discord ID", "")).strip()))
                    except ValueError:
                        pass
            return ids
        except Exception as e:
            print(f"Error looking up zone runners: {e}")
            return []

    async def _get_all_runners(self, guild_id: int) -> list[int]:
        sheet = await self._get_runner_sheet(guild_id)
        if not sheet:
            return []
        try:
            runner_emails = set()
            for row in sheet.get_all_records():
                if row.get("Runner Zone", ""):
                    email = str(row.get("Email", "")).strip()
                    if email:
                        runner_emails.add(email.lower())

            if not runner_emails:
                return []

            spreadsheet = self.bot.spreadsheets[guild_id]
            main_data = spreadsheet.worksheet(config.SHEET_PAGE_NAME).get_all_records()
            ids = []
            for row in main_data:
                if str(row.get("Email", "")).strip().lower() in runner_emails:
                    try:
                        ids.append(int(str(row.get("Discord ID", "")).strip()))
                    except ValueError:
                        pass
            return ids
        except Exception as e:
            print(f"Error looking up all runners: {e}")
            return []

    # ── Ticket helpers ────────────────────────────────────────────────────────

    async def _send_ticket_repings(self, thread: discord.Thread, ticket_info: dict) -> None:
        try:
            ping_count = ticket_info["ping_count"] + 1

            if ping_count >= 3:
                print(f"Final ping for ticket {thread.id} - getting ALL runners")
                runner_ids = await self._get_all_runners(thread.guild.id)
            else:
                runner_ids = ticket_info["zone_runners"]

            runner_mentions = [
                thread.guild.get_member(rid).mention
                for rid in runner_ids
                if thread.guild.get_member(rid)
            ]

            if not runner_mentions:
                print(f"No valid runners found for re-ping in ticket {thread.id}")
                return

            location_parts = [ticket_info["building"]]
            if ticket_info["room"]:
                location_parts.append(f"Room {ticket_info['room']}")
            location = ", ".join(location_parts)

            if ping_count < 3:
                embed = discord.Embed(
                    title="Still Need Help!",
                    description=f"**Event:** {ticket_info['event']}\n**Location:** {location}",
                    color=discord.Color.orange(),
                )
                embed.add_field(name="Runners Assigned", value="This ticket still needs assistance!", inline=False)
            else:
                embed = discord.Embed(
                    title="Final Call!",
                    description=f"**Event:** {ticket_info['event']}\n**Location:** {location}",
                    color=discord.Color.red(),
                )
                embed.add_field(name="ALL RUNNERS", value="This ticket still needs assistance!", inline=False)
                embed.add_field(name="\nFinal Ping", value="This is the final automatic ping. Please respond if you can help!", inline=False)

            await thread.send(content=" ".join(runner_mentions), embed=embed)
            print(f"Sent re-ping #{ping_count} for ticket {thread.id}")

        except Exception as e:
            print(f"Error sending re-ping for ticket {thread.id}: {e}")

    async def _check_for_burger_request(self, thread: discord.Thread) -> None:
        try:
            phrases = ("55 burgers", "fifty five burgers", "55 burger", "fifty five burger")
            title_lower = thread.name.lower()
            has_phrase = any(p in title_lower for p in phrases)

            if not has_phrase:
                try:
                    async for message in thread.history(limit=1, oldest_first=True):
                        content_lower = message.content.lower()
                        has_phrase = any(p in content_lower for p in phrases)
                        break
                except Exception as e:
                    print(f"Could not check initial message for burger phrase: {e}")

            if not has_phrase:
                return

            ticket_creator = thread.owner
            if not ticket_creator:
                return

            print(f"Burger request detected in ticket '{thread.name}'!")
            self.active_burger_deliveries[ticket_creator.id] = {"stop": False, "user": ticket_creator}

            try:
                for burger_num in range(1, 56):
                    delivery = self.active_burger_deliveries.get(ticket_creator.id)
                    if delivery and delivery["stop"]:
                        print(f"Burger delivery stopped for {ticket_creator} at burger {burger_num}")
                        await ticket_creator.send("Grill exploded. No more burgers for you :(")
                        self.active_burger_deliveries.pop(ticket_creator.id, None)
                        return

                    await ticket_creator.send("🍔")
                    await ticket_creator.send(f"Burger {burger_num} of 55")
                    print(f"Sent burger {burger_num} of 55 to {ticket_creator}")

                    if burger_num < 55:
                        await asyncio.sleep(random.uniform(5, 3600))

                self.active_burger_deliveries.pop(ticket_creator.id, None)
                await ticket_creator.send("Would you like any fries with that? 🍟")
                print(f"Completed sending all 55 burgers to {ticket_creator}")

            except discord.Forbidden:
                print(f"Cannot DM {ticket_creator} - they may have DMs disabled")
                self.active_burger_deliveries.pop(ticket_creator.id, None)
            except Exception as e:
                print(f"Error sending burger DM to {ticket_creator}: {e}")

        except Exception as e:
            print(f"Error in burger request check: {e}")

    # ── Background task ───────────────────────────────────────────────────────

    @tasks.loop(minutes=1)
    async def check_help_tickets(self) -> None:
        """Re-ping runners on unanswered tickets."""
        if not self.active_help_tickets:
            return

        print(f"Checking {len(self.active_help_tickets)} active help tickets...")
        current_time = datetime.now()
        to_remove = []

        for thread_id, ticket_info in self.active_help_tickets.items():
            try:
                ping_count = ticket_info["ping_count"]
                wait_time = timedelta(minutes=3) if ping_count == 1 else timedelta(minutes=1)

                if current_time - ticket_info["created_at"] < wait_time:
                    continue

                thread = None
                for guild in self.bot.guilds:
                    thread = guild.get_thread(thread_id)
                    if thread:
                        break

                if not thread:
                    print(f"Thread {thread_id} not found in any guild, removing from tracking")
                    to_remove.append(thread_id)
                    continue

                if thread.archived or thread.locked:
                    to_remove.append(thread_id)
                    continue

                if ticket_info["ping_count"] >= 3:
                    to_remove.append(thread_id)
                    continue

                await self._send_ticket_repings(thread, ticket_info)
                ticket_info["ping_count"] += 1
                ticket_info["created_at"] = current_time
                print(f"Re-pinged ticket {thread_id} (ping #{ticket_info['ping_count']})")

            except Exception as e:
                print(f"Error checking ticket {thread_id}: {e}")
                to_remove.append(thread_id)

        for thread_id in to_remove:
            self.active_help_tickets.pop(thread_id, None)

    @check_help_tickets.before_loop
    async def before_check_help_tickets(self) -> None:
        await self.bot.wait_until_ready()

    # ── Events ────────────────────────────────────────────────────────────────

    @commands.Cog.listener()
    async def on_thread_create(self, thread: discord.Thread) -> None:
        try:
            if not (
                hasattr(thread, "parent")
                and thread.parent
                and thread.parent.name == "help"
                and hasattr(thread.parent, "type")
                and thread.parent.type == discord.ChannelType.forum
            ):
                return

            print(f"New help ticket created: '{thread.name}' by {thread.owner}")
            await self._check_for_burger_request(thread)

            ticket_creator = thread.owner
            if not ticket_creator:
                print("Could not determine ticket creator")
                return

            guild_id = thread.guild.id
            user_event_info = await self._get_user_event_building(guild_id, ticket_creator.id)
            if not user_event_info:
                print(f"Could not find event/building info for user {ticket_creator}")
                return

            building = user_event_info.get("building")
            event = user_event_info.get("event")
            room = user_event_info.get("room")

            if not building:
                print(f"No building found for user {ticket_creator} (event: {event})")
                return

            zone = await self._get_building_zone(guild_id, building)
            is_fallback = False

            if not zone:
                print(f"No zone for building '{building}' — falling back to ALL runners")
                zone_runners = await self._get_all_runners(guild_id)
                is_fallback = True
            else:
                zone_runners = await self._get_zone_runners(guild_id, zone)
                if not zone_runners:
                    print(f"No runners in zone {zone}, falling back to ALL runners")
                    zone_runners = await self._get_all_runners(guild_id)
                    is_fallback = True

            if not zone_runners:
                print("No runners found in server at all!")
                return

            runner_mentions = [
                thread.guild.get_member(rid).mention
                for rid in zone_runners
                if thread.guild.get_member(rid)
            ]

            if not runner_mentions:
                print(f"No valid Discord members found among runners")
                return

            location_parts = [building]
            if room:
                location_parts.append(f"Room {room}")
            location = ", ".join(location_parts)

            embed = discord.Embed(
                title="New Help Ticket",
                description=f"**Ticket:** {thread.mention}\n**Creator:** {ticket_creator.mention}\n**Event:** {event}\n**Location:** {location}",
                color=discord.Color.yellow(),
            )
            field_name = "ALL Runners" if is_fallback else "Runners Assigned"
            field_value = (
                "No zone runners found - pinging all runners!\nPlease respond here if you can assist with this ticket!"
                if is_fallback
                else "Please respond here if you can assist with this ticket!"
            )
            embed.add_field(name=field_name, value=field_value, inline=False)

            await thread.send(content=" ".join(runner_mentions), embed=embed)
            print(f"Pinged {len(runner_mentions)} runner(s) in ticket")

            self.active_help_tickets[thread.id] = {
                "created_at": datetime.now(),
                "zone_runners": zone_runners,
                "has_response": False,
                "ping_count": 1,
                "zone": zone,
                "creator_id": ticket_creator.id,
                "building": building,
                "event": event,
                "room": room,
            }
            print(f"Added ticket {thread.id} to tracking system")

        except Exception as e:
            print(f"Error handling help ticket creation: {e}")
            import traceback
            traceback.print_exc()

    @commands.Cog.listener()
    async def on_message(self, message: discord.Message) -> None:
        if message.author.bot:
            return
        if message.channel.id not in self.active_help_tickets:
            return
        try:
            guild_id = message.guild.id if message.guild else None
            if guild_id:
                all_runner_ids = await self._get_all_runners(guild_id)
                if message.author.id in all_runner_ids:
                    self.active_help_tickets.pop(message.channel.id, None)
                    print(f"Runner {message.author} responded to ticket {message.channel.id} — removed from tracking")
        except Exception as e:
            print(f"Error handling message for ticket tracking: {e}")

    @commands.Cog.listener()
    async def on_reaction_add(self, reaction: discord.Reaction, user: discord.User) -> None:
        if user.bot:
            return
        if reaction.message.channel.id not in self.active_help_tickets:
            return
        try:
            helpful_reactions = {"👍", "✅", "🆗", "👌", "✋", "🙋", "🙋‍♂️", "🙋‍♀️"}
            if str(reaction.emoji) not in helpful_reactions:
                return
            guild_id = reaction.message.guild.id if reaction.message.guild else None
            if guild_id:
                all_runner_ids = await self._get_all_runners(guild_id)
                if user.id in all_runner_ids:
                    self.active_help_tickets.pop(reaction.message.channel.id, None)
                    print(f"Runner {user} reacted to ticket {reaction.message.channel.id} — removed from tracking")
        except Exception as e:
            print(f"Error handling reaction for ticket tracking: {e}")

    @commands.Cog.listener()
    async def on_thread_delete(self, thread: discord.Thread) -> None:
        if thread.id in self.active_help_tickets:
            self.active_help_tickets.pop(thread.id, None)
            print(f"Removed deleted ticket {thread.id} from tracking")

    # ── Slash commands ────────────────────────────────────────────────────────

    @app_commands.command(name="activetickets", description="Show all active help tickets being tracked (Admin only)")
    async def active_tickets_command(self, interaction: discord.Interaction) -> None:
        if not interaction.user.guild_permissions.administrator:
            await interaction.response.send_message("You need administrator permissions to use this command.", ephemeral=True)
            return

        await interaction.response.defer(ephemeral=True)
        try:
            if not self.active_help_tickets:
                await interaction.followup.send("No active help tickets being tracked.")
                return

            embed = discord.Embed(
                title="Active Help Tickets",
                description=f"Currently tracking {len(self.active_help_tickets)} help ticket(s)",
                color=discord.Color.blue(),
            )

            for thread_id, ticket_info in list(self.active_help_tickets.items())[:10]:
                thread = interaction.guild.get_thread(thread_id)
                thread_name = thread.name if thread else f"Thread {thread_id} (not found)"
                minutes_elapsed = int((datetime.now() - ticket_info["created_at"]).total_seconds() / 60)

                location_parts = [ticket_info["building"]]
                if ticket_info["room"]:
                    location_parts.append(f"Room {ticket_info['room']}")
                location = ", ".join(location_parts)

                embed.add_field(
                    name=f"Ticket: {thread_name}",
                    value=(
                        f"**Event:** {ticket_info['event']}\n"
                        f"**Location:** {location}\n"
                        f"**Zone:** {ticket_info['zone']}\n"
                        f"**Pings:** {ticket_info['ping_count']}\n"
                        f"**Time:** {minutes_elapsed}m ago"
                    ),
                    inline=True,
                )

            if len(self.active_help_tickets) > 10:
                embed.add_field(
                    name="Note",
                    value=f"Showing first 10 of {len(self.active_help_tickets)} active tickets",
                    inline=False,
                )

            await interaction.followup.send(embed=embed)
        except Exception as e:
            await interaction.followup.send(f"Error fetching active tickets: {e}")

    @app_commands.command(name="stopburgers", description="Stop all active burger deliveries")
    async def stop_burgers_command(self, interaction: discord.Interaction) -> None:
        await interaction.response.defer(ephemeral=True)
        try:
            if not self.active_burger_deliveries:
                await interaction.followup.send("No active burger deliveries to stop.")
                return

            count = len(self.active_burger_deliveries)
            for delivery_info in self.active_burger_deliveries.values():
                delivery_info["stop"] = True

            await interaction.followup.send(
                f"Emergency stop activated! Stopped **{count}** active burger deliver{'y' if count == 1 else 'ies'}.\n"
                "Users will receive a final 'Grill exploded' message."
            )
        except Exception as e:
            await interaction.followup.send(f"Error stopping burgers: {e}")

    @app_commands.command(name="debugzone", description="Debug zone assignment for a user (Admin only)")
    @app_commands.describe(user="The user to debug zone assignment for")
    async def debug_zone_command(self, interaction: discord.Interaction, user: discord.Member) -> None:
        if not interaction.user.guild_permissions.administrator:
            await interaction.response.send_message("You need administrator permissions to use this command.", ephemeral=True)
            return

        await interaction.response.defer(ephemeral=True)
        try:
            guild_id = interaction.guild.id
            user_event_info = await self._get_user_event_building(guild_id, user.id)
            if not user_event_info:
                await interaction.followup.send(f"Could not find event/building info for {user.mention}")
                return

            building = user_event_info.get("building")
            event = user_event_info.get("event")
            room = user_event_info.get("room")
            name = user_event_info.get("name")

            if not building:
                await interaction.followup.send(f"No building found for {user.mention} (event: {event})")
                return

            zone = await self._get_building_zone(guild_id, building)
            zone_runners_ids = await self._get_zone_runners(guild_id, zone) if zone else []
            zone_runners = [interaction.guild.get_member(rid) for rid in zone_runners_ids]
            zone_runners = [m for m in zone_runners if m]

            embed = discord.Embed(
                title=f"Zone Debug: {name}",
                description=f"**User:** {user.mention}\n**Event:** {event}\n**Building:** {building}\n**Room:** {room or 'N/A'}\n**Zone:** {zone or 'Not assigned'}",
                color=discord.Color.blue(),
            )
            if zone_runners:
                embed.add_field(
                    name=f"Zone {zone} Runners ({len(zone_runners)})",
                    value="\n".join(m.mention for m in zone_runners[:10]) or "None",
                    inline=False,
                )

            await interaction.followup.send(embed=embed)
        except Exception as e:
            await interaction.followup.send(f"Error debugging zone: {e}")
            import traceback
            traceback.print_exc()


async def setup(bot: commands.Bot) -> None:
    await bot.add_cog(Tickets(bot))
