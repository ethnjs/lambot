# lam_bot.py
import os
import asyncio
import discord
from discord.ext import commands, tasks
from discord import app_commands
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv
from datetime import datetime, timedelta
import json
import random

load_dotenv()

TOKEN         = os.getenv("DISCORD_TOKEN")
SERVICE_EMAIL = os.getenv("SERVICE_EMAIL")
SHEET_ID      = os.getenv("SHEET_ID")  # Optional - can be set via /enterfolder command
SHEET_PAGE_NAME = os.getenv("SHEET_PAGE_NAME", "lambot")  # Name of the worksheet/tab within the sheet
AUTO_CREATE_ROLES = os.getenv("AUTO_CREATE_ROLES", "true").lower() == "true"
DEFAULT_ROLE_COLOR = os.getenv("DEFAULT_ROLE_COLOR", "light_gray")  # blue, red, green, purple, etc.

# ⚠️ ⚠️ ⚠️  DANGER ZONE: COMPLETE SERVER RESET  ⚠️ ⚠️ ⚠️
# Set to True to COMPLETELY RESET the server on bot startup
# WARNING: This will permanently delete ALL channels, categories, roles, and reset all nicknames!
# This is IRREVERSIBLE! Use only for testing or complete server reset!
RESET_SERVER = os.getenv("RESET_SERVER", "false").lower() == "true"

intents = discord.Intents.default()
intents.members = True
intents.message_content = True

class LamBot(commands.Bot):
    def __init__(self):
        super().__init__(command_prefix='!', intents=intents)

bot = LamBot()

# Set up gspread client and credentials for Drive API
scope = [
    "https://www.googleapis.com/auth/spreadsheets",  # Full spreadsheet access (read & write)
    "https://www.googleapis.com/auth/drive.readonly"  # Needed to search for sheets
]
creds = ServiceAccountCredentials.from_json_keyfile_dict(json.load(open("secrets/gspread.json")), scope)
gc = gspread.authorize(creds)

# Sheet connections are now per-guild (per-server)
# Each Discord server can have its own Google Sheet connection
sheets = {}  # guild_id -> worksheet object
spreadsheets = {}  # guild_id -> spreadsheet object

# Note: SHEET_ID is still available as environment variable but won't auto-connect
# Use /enterfolder command to connect to sheets dynamically
print("📋 Bot starting - will attempt to load cached sheet connections or use /enterfolder command")

# Store pending role assignments and user info for users who haven't joined yet
pending_users = {}  # Changed from pending_roles to store more info

# Track chapter role names globally
chapter_role_names = set()

# Track active help tickets for re-pinging
active_help_tickets = {}  # thread_id -> ticket_info

# Track active burger deliveries for stopping
active_burger_deliveries = {}  # user_id -> {"stop": False, "user": user_object}

# Cache configuration
CACHE_FILE = "bot_cache.json"

# Bit to show if setup is done
admin_lock = asyncio.Lock()
rate_limit_lock = asyncio.Lock()
reset_active = False


ALLOWED_DURING_RESET = {"enterfolder"}

#Setting bits for the server (guild_id -> int)
runner_all_access = {}

async def safe_call(coro):
    async with rate_limit_lock:
        result = await coro
        await asyncio.sleep(0.5)
        return result
    
def save_cache(data):
    """Save cache data to JSON file"""
    try:
        with open(CACHE_FILE, 'w') as f:
            json.dump(data, f, indent=2)
        print(f"✅ Saved cache to {CACHE_FILE}")
    except Exception as e:
        print(f"❌ Error saving cache: {e}")

def load_cache():
    """Load cache data from JSON file"""
    try:
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, 'r') as f:
                data = json.load(f)
            print(f"✅ Loaded cache from {CACHE_FILE}")
            return data
        else:
            print(f"📄 No cache file found at {CACHE_FILE}")
            return {}
    except Exception as e:
        print(f"❌ Error loading cache: {e}")
        return {}

def clear_cache():
    """Clear the cache file"""
    try:
        if os.path.exists(CACHE_FILE):
            os.remove(CACHE_FILE)
            print(f"🗑️ Cleared cache file {CACHE_FILE}")
        else:
            print(f"📄 No cache file to clear")
    except Exception as e:
        print(f"❌ Error clearing cache: {e}")

async def load_spreadsheets_from_cache():
    """Try to load all guild spreadsheet connections from cache"""
    global sheets, spreadsheets

    cache = load_cache()
    guilds_cache = cache.get("guilds", {})

    if not guilds_cache:
        print("📋 No cached spreadsheet connections found")
        return False

    success_count = 0
    for guild_id_str, guild_cache in guilds_cache.items():
        guild_id = int(guild_id_str)
        spreadsheet_id = guild_cache.get("spreadsheet_id")
        worksheet_name = guild_cache.get("worksheet_name", SHEET_PAGE_NAME)

        if not spreadsheet_id:
            continue

        try:
            print(f"🔄 Attempting to connect to cached spreadsheet for guild {guild_id}: {spreadsheet_id}")
            spreadsheet = gc.open_by_key(spreadsheet_id)
            sheet = spreadsheet.worksheet(worksheet_name)

            # Test the connection by getting the first row
            headers = sheet.row_values(1)

            # Store in per-guild dictionaries
            spreadsheets[guild_id] = spreadsheet
            sheets[guild_id] = sheet

            print(f"✅ Successfully connected to cached spreadsheet for guild {guild_id}: '{spreadsheet.title}'")
            print(f"📊 Worksheet: '{sheet.title}' with {len(headers)} columns")
            success_count += 1

        except Exception as e:
            print(f"❌ Failed to connect to cached spreadsheet for guild {guild_id}: {e}")

    if success_count > 0:
        print(f"✅ Loaded {success_count} cached spreadsheet connection(s)")
        return True
    else:
        print("❌ No cached connections could be loaded")
        return False

def save_guild_spreadsheet_to_cache(guild_id, spreadsheet_id, worksheet_name):
    """Save a guild's spreadsheet connection to cache"""
    cache = load_cache()

    if "guilds" not in cache:
        cache["guilds"] = {}

    cache["guilds"][str(guild_id)] = {
        "spreadsheet_id": spreadsheet_id,
        "worksheet_name": worksheet_name
    }

    save_cache(cache)
    print(f"💾 Cached spreadsheet connection for guild {guild_id}")

def clear_guild_cache(guild_id):
    """Clear a specific guild's cache"""
    cache = load_cache()

    if "guilds" in cache and str(guild_id) in cache["guilds"]:
        del cache["guilds"][str(guild_id)]
        save_cache(cache)
        print(f"🧹 Cleared cache for guild {guild_id}")
        return True
    return False

async def handle_rate_limit(coro, operation_name, max_retries=3, default_delay=0.1):
    """
    Helper function to handle rate limits for Discord API calls.

    Args:
        coro: Coroutine to execute
        operation_name: Name of the operation for logging
        max_retries: Maximum number of retries (default: 3)
        default_delay: Default delay after successful operation in seconds (default: 0.1)

    Returns:
        Result of the coroutine, or None if all retries failed
    """
    retry_count = 0
    while retry_count < max_retries:
        try:
            result = await coro
            # Small delay after successful operation to avoid rate limits
            await asyncio.sleep(default_delay)
            return result
        except discord.HTTPException as e:
            error_msg = str(e)
            if e.status == 429 or "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                retry_count += 1
                if retry_count < max_retries:
                    # Extract retry_after from response if available
                    retry_after = 1.0  # Default 1 second
                    if hasattr(e, 'retry_after') and e.retry_after:
                        retry_after = float(e.retry_after)
                    elif isinstance(e.response, dict) and 'retry_after' in e.response:
                        retry_after = float(e.response['retry_after'])

                    print(f"⚠️ Rate limited on {operation_name}, waiting {retry_after}s before retry {retry_count}/{max_retries}...")
                    await asyncio.sleep(retry_after)
                else:
                    print(f"❌ Rate limited on {operation_name} after {max_retries} retries, giving up")
                    return None
            else:
                # Re-raise non-rate-limit errors
                raise
        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                retry_count += 1
                if retry_count < max_retries:
                    print(f"⚠️ Rate limited on {operation_name}, waiting 1s before retry {retry_count}/{max_retries}...")
                    await asyncio.sleep(1.0)
                else:
                    print(f"❌ Rate limited on {operation_name} after {max_retries} retries, giving up")
                    return None
            else:
                # Re-raise non-rate-limit errors
                raise

    return None

async def get_or_create_role(guild, role_name):
    """Get a role by name, or create it if it doesn't exist"""
    role = discord.utils.get(guild.roles, name=role_name)
    if role:
        #print(f"🧻 Found '{role_name}' so no longer attempting to create it")
        return role

    # Check if auto-creation is enabled
    if not AUTO_CREATE_ROLES:
        print(f"⚠️ Role '{role_name}' not found and auto-creation is disabled")
        return None

    # Role doesn't exist, create it
    try:
        # Special case: Admin role gets full permissions
        if role_name == "Admin":
            max_retries = 3
            retry_count = 0
            while retry_count < max_retries:
                try:
                    role = await guild.create_role(
                        name="Admin",
                        permissions=discord.Permissions.all(),
                        color=discord.Color.purple(),
                        reason="Auto-created Admin role for ezhang."
                    )
                    print(f"🆕 Created Admin role with full permissions")
                    await asyncio.sleep(0.5)  # Small delay to avoid rate limits
                    return role
                except discord.HTTPException as e:
                    error_msg = str(e)
                    if e.status == 429 or "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                        retry_count += 1
                        if retry_count < max_retries:
                            retry_after = 1.0
                            if hasattr(e, 'retry_after') and e.retry_after:
                                retry_after = float(e.retry_after)
                            print(f"⚠️ Rate limited creating Admin role, waiting {retry_after}s before retry {retry_count}/{max_retries}...")
                            await asyncio.sleep(retry_after)
                        else:
                            print(f"❌ Rate limited creating Admin role after {max_retries} retries, giving up")
                            return None
                    else:
                        raise  # Re-raise non-rate-limit errors
                except Exception as e:
                    error_msg = str(e)
                    if "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                        retry_count += 1
                        if retry_count < max_retries:
                            print(f"⚠️ Rate limited creating Admin role, waiting 1s before retry {retry_count}/{max_retries}...")
                            await asyncio.sleep(1.0)
                        else:
                            print(f"❌ Rate limited creating Admin role after {max_retries} retries, giving up")
                            return None
                    else:
                        raise  # Re-raise non-rate-limit errors
            return None

        # Custom color mapping for specific roles
        custom_role_colors = {
            # Team roles only
            "Runner": discord.Color.orange(),
            "Awards": discord.Color.yellow(),
            "Volunteer": discord.Color.blue(),
            "Lead ES": discord.Color.yellow(),
            "Photographer": discord.Color.red(),
            "Arbitrations": discord.Color.green(),
            "Social Media": discord.Color.magenta(),
            "VIPer": discord.Color.green(),
        }

        # Check if this role has a custom color
        if role_name in custom_role_colors:
            role_color = custom_role_colors[role_name]
            color_name = {
                discord.Color.yellow(): "yellow",
                discord.Color.blue(): "blue",
                discord.Color.green(): "green",
                discord.Color.red(): "red"
            }.get(role_color, "custom")
        elif role_name == "Unaffiliated" or role_name in chapter_role_names:
            # Chapter roles should be green
            role_color = discord.Color.green()
            color_name = "green"
        else:
            # Use default color mapping
            color_map = {
                "blue": discord.Color.blue(),
                "red": discord.Color.red(),
                "green": discord.Color.green(),
                "purple": discord.Color.purple(),
                "orange": discord.Color.orange(),
                "yellow": discord.Color.yellow(),
                "teal": discord.Color.teal(),
                "pink": discord.Color.magenta(),
                "light_gray": discord.Color.light_gray(),
                "dark_gray": discord.Color.dark_gray(),
                "black": discord.Color.from_rgb(0, 0, 0),
                "white": discord.Color.from_rgb(255, 255, 255)
            }

            role_color = color_map.get(DEFAULT_ROLE_COLOR.lower(), discord.Color.light_gray())
            color_name = DEFAULT_ROLE_COLOR

        max_retries = 3
        retry_count = 0
        while retry_count < max_retries:
            try:
                role = await guild.create_role(
                    name=role_name,
                    color=role_color,
                    reason="Auto-created by LAM Bot for onboarding"
                )
                print(f"🆕 Created new role: '{role_name}' (color: {color_name})")

                # If we just created the Runner role, ensure it has access to Tournament Officials channels
                if role_name == "Runner":
                    await ensure_runner_tournament_officials_access(guild, role)

                # Note: Test folder search is now handled in setup_building_structure after channels are created
                # to ensure the target channel exists when we try to post the message

                # Small delay after creating role to avoid rate limits
                await asyncio.sleep(0.5)
                return role
            except discord.HTTPException as e:
                error_msg = str(e)
                if e.status == 429 or "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                    retry_count += 1
                    if retry_count < max_retries:
                        retry_after = 1.0
                        if hasattr(e, 'retry_after') and e.retry_after:
                            retry_after = float(e.retry_after)
                        print(f"⚠️ Rate limited creating role '{role_name}', waiting {retry_after}s before retry {retry_count}/{max_retries}...")
                        await asyncio.sleep(retry_after)
                    else:
                        print(f"❌ Rate limited creating role '{role_name}' after {max_retries} retries, giving up")
                        return None
                else:
                    raise  # Re-raise non-rate-limit errors
            except Exception as e:
                error_msg = str(e)
                if "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
                    retry_count += 1
                    if retry_count < max_retries:
                        print(f"Rate limited creating role '{role_name}', waiting 1s before retry {retry_count}/{max_retries}...")
                        await asyncio.sleep(1.0)
                    else:
                        print(f"❌ Rate limited creating role '{role_name}' after {max_retries} retries, giving up")
                        return None
                else:
                    raise  # Re-raise non-rate-limit errors

        return None
    except discord.Forbidden:
        print(f"❌ No permission to create role '{role_name}'")
        return None
    except Exception as e:
        print(f"❌ Error creating role '{role_name}': {e}")
        return None

async def get_or_create_category(guild, category_name):
    """Get a category by name, or create it if it doesn't exist"""
    category = discord.utils.get(guild.categories, name=category_name)
    if category:
        print(f"✅ DEBUG: Found existing category: '{category_name}'")
        return category

    try:
        category = await handle_rate_limit(
            guild.create_category(
                name=category_name,
                reason="Auto-created by LAM Bot for building organization"
            ),
            f"creating category '{category_name}'"
        )
        if category:
            print(f"🏢 DEBUG: Created NEW category: '{category_name}' (ID: {category.id})")
        return category
    except discord.Forbidden:
        print(f"❌ No permission to create category '{category_name}'")
        return None
    except Exception as e:
        print(f"❌ Error creating category '{category_name}': {e}")
        return None

async def get_or_create_channel(guild, channel_name, category, event_role=None, is_building_chat=False):
    """Get a channel by name, or create it if it doesn't exist"""
    try:
        # Set up permissions
        overwrites = {}

        # Give Runner role access only to static channels (not building/event channels)
        runner_role = discord.utils.get(guild.roles, name="Runner")
        static_categories = ["Welcome", "Tournament Officials", "Volunteers"]
        
        # Look up this specific server's setting (default to 0 / restricted)
        guild_runner_access = runner_all_access.get(guild.id, 0)
        
        if runner_role and category and (guild_runner_access or category.name in static_categories):
            overwrites[runner_role] = discord.PermissionOverwrite(
                read_messages=True,
                send_messages=True,
                read_message_history=True
            )
    except discord.Forbidden:
        print(f"❌ No permission to edit runner permissions for '{channel_name}'")
        return None
    except Exception as e:
        print(f"❌ Error editing channel permissions for '{channel_name}': {e}")
        return None

    channel = discord.utils.get(guild.text_channels, name=channel_name)
    if channel:
        print(f"✅ DEBUG: Found existing channel: #{channel_name} (ID: {channel.id})")
        return channel

    print(f"🔍 DEBUG: Channel #{channel_name} not found, creating new one...")

    try:
        # Set up permissions
        overwrites = {}

        # Give Runner role access only to static channels (not building/event channels)
        runner_role = discord.utils.get(guild.roles, name="Runner")
        static_categories = ["Welcome", "Tournament Officials", "Volunteers"]
        if runner_role and category and (guild_runner_access or category.name in static_categories):
            overwrites[runner_role] = discord.PermissionOverwrite(
                read_messages=True,
                send_messages=True,
                read_message_history=True
            )

        if event_role:
            # Event-specific channel: only event role can see it (plus Runner)
            overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)
            overwrites[event_role] = discord.PermissionOverwrite(
                read_messages=True,
                send_messages=True,
                read_message_history=True
            )
        elif is_building_chat:
            # Building chat channel: restricted by default, roles will be added later
            overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)

        channel = await handle_rate_limit(
            guild.create_text_channel(
                name=channel_name,
                category=category,
                overwrites=overwrites,
                reason="Auto-created by LAM Bot for event organization"
            ),
            f"creating channel '{channel_name}'"
        )

        if channel:
            if event_role:
                print(f"📺 DEBUG: Created NEW channel: '#{channel_name}' (ID: {channel.id}, restricted to {event_role.name})")
            elif is_building_chat:
                print(f"📺 DEBUG: Created NEW building chat: '#{channel_name}' (ID: {channel.id}, restricted)")
            else:
                print(f"📺 DEBUG: Created NEW channel: '#{channel_name}' (ID: {channel.id})")
        return channel
    except discord.Forbidden:
        print(f"❌ No permission to create channel '{channel_name}'")
        return None
    except Exception as e:
        print(f"❌ Error creating channel '{channel_name}': {e}")
        return None

async def sort_building_categories_alphabetically(guild):
    """Sort all building categories alphabetically, keeping non-building categories at the top"""
    try:
        # Get all categories
        all_categories = guild.categories
        print(f"📋 DEBUG: Sorting {len(all_categories)} total categories")

        # Separate building categories from static categories
        static_categories = ["Welcome", "Tournament Officials", "Chapters", "Volunteers"]
        building_categories = []
        other_categories = []

        for category in all_categories:
            if category.name in static_categories:
                other_categories.append(category)
                print(f"📋 DEBUG: Static category: '{category.name}'")
            else:
                building_categories.append(category)
                print(f"🏢 DEBUG: Building category: '{category.name}'")

        # Sort building categories alphabetically
        building_categories.sort(key=lambda cat: cat.name.lower())
        print(f"📋 DEBUG: Sorted {len(building_categories)} building categories alphabetically")

        # Calculate positions: static categories first, then building categories
        position = 0

        # Position static categories first in the correct order
        desired_order = ["Welcome", "Tournament Officials", "Chapters", "Volunteers"]
        ordered_static_categories = []

        # Sort other_categories by desired order
        for desired_name in desired_order:
            for category in other_categories:
                if category.name == desired_name:
                    ordered_static_categories.append(category)
                    break

        # Add any remaining static categories that weren't in the desired order
        for category in other_categories:
            if category not in ordered_static_categories:
                ordered_static_categories.append(category)

        for category in ordered_static_categories:
            if category.position != position:
                result = await handle_rate_limit(
                    category.edit(position=position, reason="Organizing categories"),
                    f"moving category '{category.name}'"
                )
                if result is not None:
                    print(f"📋 Moved category '{category.name}' to position {position}")
            position += 1

        # Position building categories alphabetically after static ones
        for category in building_categories:
            if category.position != position:
                result = await handle_rate_limit(
                    category.edit(position=position, reason="Organizing building categories alphabetically"),
                    f"moving building category '{category.name}'"
                )
                if result is not None:
                    print(f"🏢 Moved building category '{category.name}' to position {position}")
            position += 1

        print("📋 Categories organized: Static categories first, then buildings alphabetically")

    except Exception as e:
        print(f"⚠️ Error organizing categories: {e}")

def sanitize_for_discord(text):
    """Sanitize text to be valid for Discord channel names"""
    # Replace spaces with hyphens and remove/replace invalid characters
    return text.lower().replace(' ', '-').replace('/', '-').replace('\\', '-').replace(':', '-').replace('*', '-').replace('?', '-').replace('"', '').replace('<', '').replace('>', '').replace('|', '-')

async def setup_building_structure(guild, building, first_event, room=None):
    """Set up category and channels for a building and event"""
    print(f"🏗️ DEBUG: Setting up building structure - Building: '{building}', Event: '{first_event}', Room: '{room}'")

    # Skip creating building chat for priority/custom roles (only create for actual event roles)
    priority_roles = ["Admin", "Volunteer", "Lead ES", "Social Media", "Photographer", "Arbitrations", "Awards", "Runner", "VIPer"]
    if first_event and first_event in priority_roles:
        print(f"⏭️ Skipping building structure creation for priority role '{first_event}' in {building} (only event roles get building structures)")
        return

    # Create or get the building category
    category_name = building
    print(f"🏢 DEBUG: Getting/creating category: '{category_name}'")
    category = await get_or_create_category(guild, category_name)
    if not category:
        return

    # Get Runner role to ensure access
    runner_role = discord.utils.get(guild.roles, name="Runner")

    # Create general building chat channel (restricted to people with events in this building)
    building_chat_name = f"{sanitize_for_discord(building)}-chat"
    print(f"📺 DEBUG: Getting/creating building chat: '{building_chat_name}'")
    building_chat = await get_or_create_channel(guild, building_chat_name, category, is_building_chat=True)

    # Check if this is a newly created building chat (no messages yet) and send welcome message
    if building_chat:
        try:
            # Check if the channel has any messages (to avoid sending duplicate welcome messages)
            messages = [message async for message in building_chat.history(limit=1)]
            if not messages:
                print(f"📝 Building chat #{building_chat.name} appears to be new, sending welcome message...")
                await send_building_welcome_message(guild, building_chat, building)
        except Exception as e:
            print(f"⚠️ Error checking/sending welcome message for #{building_chat.name}: {e}")

    # Create event-specific channel if we have the info
    if first_event:
        # Get or create the event role
        event_role = await get_or_create_role(guild, first_event)
        if event_role:
            # Skip giving building chat access to Runner role
            # (Runners use the existing "runner" channel in Tournament Officials)
            if first_event.lower() != "runner":
                # Add the event role to the building chat permissions
                await add_role_to_building_chat(building_chat, event_role)

                # Create channel name: [First Event] - [Building] [Room]
                if room:
                    channel_name = f"{sanitize_for_discord(first_event)}-{sanitize_for_discord(building)}-{sanitize_for_discord(room)}"
                else:
                    channel_name = f"{sanitize_for_discord(first_event)}-{sanitize_for_discord(building)}"

                print(f"📺 DEBUG: Getting/creating event channel: '{channel_name}' for role '{event_role.name}'")

                # Create event channel
                event_channel = await get_or_create_channel(guild, channel_name, category, event_role)

                # Note: Test materials are no longer automatically shared
                # Use /sendtestmaterials command to manually share test materials for events

async def search_and_share_test_folder(guild, role_name):
    """Search for test materials folder and share with event participants"""
    try:
        print(f"🔍 DEBUG: Starting search for test materials for event: {role_name}")

        guild_id = guild.id

        # Check if we have a connected spreadsheet to get the folder ID
        if guild_id not in spreadsheets:
            print(f"❌ DEBUG: No spreadsheet connected for guild {guild_id}, cannot search for test folder for {role_name}")
            return

        guild_spreadsheet = spreadsheets[guild_id]
        print(f"✅ DEBUG: Spreadsheet connected, ID: {guild_spreadsheet.id}")

        # Import Drive API
        from googleapiclient.discovery import build

        # Build Drive API service
        drive_service = build('drive', 'v3', credentials=creds)

        # Get the parent folder ID of the connected spreadsheet
        sheet_metadata = drive_service.files().get(fileId=guild_spreadsheet.id, fields='parents').execute()
        parent_folders = sheet_metadata.get('parents', [])

        if not parent_folders:
            print(f"❌ DEBUG: Could not find parent folder for the spreadsheet")
            return

        parent_folder_id = parent_folders[0]
        print(f"✅ DEBUG: Found parent folder ID: {parent_folder_id}")

        # Search for "Tests" folder in the parent directory
        tests_query = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='Tests'"
        tests_results = drive_service.files().list(q=tests_query, fields='files(id, name)').execute()
        tests_folders = tests_results.get('files', [])

        if not tests_folders:
            print(f"❌ DEBUG: No 'Tests' folder found in the parent directory")
            print(f"🔍 DEBUG: Searching for any folders in parent directory...")
            # List all folders in parent to help debug
            all_folders_query = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'"
            all_folders_results = drive_service.files().list(q=all_folders_query, fields='files(id, name)').execute()
            all_folders = all_folders_results.get('files', [])
            print(f"📁 DEBUG: Found folders: {[f['name'] for f in all_folders]}")
            return

        tests_folder_id = tests_folders[0]['id']
        print(f"✅ DEBUG: Found Tests folder: {tests_folder_id}")

        # Search for the event-specific folder within Tests
        event_query = f"'{tests_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='{role_name}'"
        event_results = drive_service.files().list(q=event_query, fields='files(id, name, webViewLink)').execute()
        event_folders = event_results.get('files', [])

        if not event_folders:
            print(f"❌ DEBUG: No folder found for event '{role_name}' in Tests directory")
            print(f"🔍 DEBUG: Searching for any folders in Tests directory...")
            # List all folders in Tests to help debug
            all_test_folders_query = f"'{tests_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'"
            all_test_folders_results = drive_service.files().list(q=all_test_folders_query, fields='files(id, name)').execute()
            all_test_folders = all_test_folders_results.get('files', [])
            print(f"📁 DEBUG: Found test folders: {[f['name'] for f in all_test_folders]}")
            return

        event_folder = event_folders[0]
        event_folder_id = event_folder['id']
        print(f"✅ DEBUG: Found test folder for {role_name}: {event_folder_id}")

        # Get all files in the event folder
        files_query = f"'{event_folder_id}' in parents and trashed=false"
        files_results = drive_service.files().list(q=files_query, fields='files(id, name, webViewLink, mimeType)').execute()
        files = files_results.get('files', [])

        if not files:
            print(f"❌ DEBUG: No files found in {role_name} test folder")
            return

        print(f"✅ DEBUG: Found {len(files)} files in {role_name} test folder")

        # Find the appropriate channel to post to (event-specific channel)
        target_channel = None

        print(f"🔍 DEBUG: Looking for channel containing '{role_name.lower().replace(' ', '-')}'")
        print(f"📺 DEBUG: Available channels: {[c.name for c in guild.text_channels]}")

        # Look for event-specific channels that contain the role name
        for channel in guild.text_channels:
            print(f"🔍 DEBUG: Checking channel #{channel.name} in category {channel.category.name if channel.category else 'None'}")
            if (role_name.lower().replace(' ', '-') in channel.name.lower() and
                channel.category and
                channel.category.name not in ["Welcome", "Tournament Officials", "Volunteers"]):
                target_channel = channel
                print(f"✅ DEBUG: Found target channel: #{channel.name}")
                break

        if not target_channel:
            print(f"❌ DEBUG: Could not find appropriate channel for {role_name}")
            print(f"🔍 DEBUG: Searched for channels containing: '{role_name.lower().replace(' ', '-')}'")
            return

        # Check if test materials message already exists in pinned messages
        pinned_messages = await safe_call(target_channel.pins())
        test_materials_exists = False

        for message in pinned_messages:
            if message.embeds and message.embeds[0].title and f"📚 Test Materials for {role_name}" in message.embeds[0].title:
                test_materials_exists = True
                print(f"✅ DEBUG: Test materials message already pinned in #{target_channel.name}, skipping")
                break

        if test_materials_exists:
            return

        # Create embed for the test materials
        embed = discord.Embed(
            title=f"📚 Test Materials for {role_name}",
            description=f"Access your event-specific test materials and resources!\nPlease DO NOT share these materials with ANYBODY else (not even volunteers from different events).",
            color=discord.Color.green()
        )

        # Create file links as bullet points
        file_links = []
        for file in files:
            file_name = file['name']
            file_link = file['webViewLink']

            # Determine file type emoji
            mime_type = file.get('mimeType', '')
            if 'pdf' in mime_type:
                emoji = "📄"
            elif 'document' in mime_type:
                emoji = "📝"
            elif 'spreadsheet' in mime_type:
                emoji = "📊"
            elif 'presentation' in mime_type:
                emoji = "📖"
            elif 'image' in mime_type:
                emoji = "🖼️"
            elif 'folder' in mime_type:
                emoji = "📁"
            else:
                emoji = "📎"

            file_links.append(f"• {emoji} [**{file_name}**]({file_link})")

        # Split into chunks if too long for Discord (1024 character limit per field)
        # Discord's limit is 1024 chars per field, so we use 1000 to be safe
        chunk_size = 1000
        chunks = []
        current_chunk = ""

        for link in file_links:
            # Check if adding this link would exceed the limit
            if len(current_chunk + link + "\n") > chunk_size:
                # Save current chunk and start a new one
                if current_chunk.strip():
                    chunks.append(current_chunk.strip())
                current_chunk = link + "\n"
            else:
                current_chunk += link + "\n"

        # Don't forget the last chunk
        if current_chunk.strip():
            chunks.append(current_chunk.strip())

        # Send the first message with the main embed
        embed.add_field(name="📋 Test Materials", value=chunks[0] if chunks else "No files found", inline=False)
        message = await target_channel.send(embed=embed)
        await asyncio.sleep(0.2)
        print(f"📚 Shared test materials for {role_name} in #{target_channel.name}")

        # Pin the first message
        try:
            await safe_call(message.pin())
            print(f"📌 Pinned test materials message in #{target_channel.name}")
        except discord.Forbidden:
            print(f"⚠️ No permission to pin message in #{target_channel.name}")
        except Exception as pin_error:
            print(f"⚠️ Error pinning message in #{target_channel.name}: {pin_error}")

        # Send additional messages for remaining chunks
        if len(chunks) > 1:
            for i, chunk in enumerate(chunks[1:], start=2):
                continuation_embed = discord.Embed(
                    title=f"📚 Test Materials for {role_name} (continued {i})",
                    description="",
                    color=discord.Color.green()
                )
                continuation_embed.add_field(name="📋 Test Materials", value=chunk, inline=False)
                await target_channel.send(embed=continuation_embed)
                print(f"📚 Sent continuation message {i} for {role_name}")
                await asyncio.sleep(0.5)

        # Check if scoring message already exists in pinned messages
        scoring_message_exists = False
        for msg in pinned_messages:
            if msg.embeds and msg.embeds[0].title and "📊 Score Input Instructions" in msg.embeds[0].title:
                scoring_message_exists = True
                print(f"✅ DEBUG: Scoring message already pinned in #{target_channel.name}, skipping")
                break

        if not scoring_message_exists:
            # Create scoring instructions embed
            scoring_embed = discord.Embed(
                title="📊 Score Input Instructions",
                description="**IMPORTANT**: All Lead Event Supervisors must input scores through the official scoring portal!",
                color=discord.Color.blue()
            )

            scoring_embed.add_field(
                name="🔗 Scoring Portal",
                value="[**Click here to access the scoring system**](https://scoring.duosmium.org/login)",
                inline=False
            )

            scoring_embed.add_field(
                name="📋 Instructions",
                value="• Lead Event Supervisors should have received an invitation email to the scoring portal\n• Select the correct tournament and event\n• Input all team scores accurately\n• Double-check scores before submitting\n• Contact admin if you have login issues",
                inline=False
            )

            scoring_embed.add_field(
                name="⚠️ Important Notes",
                value="• Scores must be entered promptly after each event\n• Do not share your login credentials\n• Report any technical issues immediately",
                inline=False
            )

            # Post the scoring embed
            scoring_message = await target_channel.send(embed=scoring_embed)
            print(f"📊 Shared scoring instructions for {role_name} in #{target_channel.name}")

            # Pin the scoring message
            try:
                await safe_call(scoring_message.pin())
                print(f"📌 Pinned scoring instructions message in #{target_channel.name}")
            except discord.Forbidden:
                print(f"⚠️ No permission to pin scoring message in #{target_channel.name}")
            except Exception as pin_error:
                print(f"⚠️ Error pinning scoring message in #{target_channel.name}: {pin_error}")

    except Exception as e:
        print(f"❌ Error searching for test folder for {role_name}: {e}")

async def setup_chapter_structure(guild, chapter_name):
    """Set up channels for a chapter"""
    print(f"📖 DEBUG: Setting up chapter structure - Chapter: '{chapter_name}'")

    # Add to global chapter role names set
    global chapter_role_names
    chapter_role_names.add(chapter_name)

    # Get or create the Chapters category
    chapters_category = await get_or_create_category(guild, "Chapters")
    if not chapters_category:
        return

    # Sanitize chapter name for Discord channel
    channel_name = sanitize_for_discord(chapter_name)

    # Create chapter channel
    chapter_channel = await get_or_create_channel(guild, channel_name, chapters_category)

    # Get or create the chapter role
    chapter_role = await get_or_create_role(guild, chapter_name)

    if chapter_channel and chapter_role:
        # Set up permissions so only chapter members can see the channel
        try:
            overwrites = chapter_channel.overwrites
            # Hide from @everyone
            overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)
            # Give chapter role access
            overwrites[chapter_role] = discord.PermissionOverwrite(
                read_messages=True,
                send_messages=True,
                read_message_history=True
            )

            await handle_rate_limit(
                chapter_channel.edit(overwrites=overwrites, reason=f"Set up {chapter_name} chapter permissions"),
                f"editing chapter channel '{channel_name}' permissions"
            )
            print(f"📖 Set up permissions for #{channel_name} chapter channel")

            # Sort chapter channels after creating a new one
            await sort_chapter_channels_alphabetically(guild)

        except Exception as e:
            print(f"❌ Error setting up permissions for #{channel_name}: {e}")

async def sort_chapter_channels_alphabetically(guild):
    """Sort chapter channels alphabetically with unaffiliated at the bottom"""
    try:
        # Find the Chapters category
        chapters_category = discord.utils.get(guild.categories, name="Chapters")
        if not chapters_category:
            print("⚠️ Chapters category not found")
            return

        # Get all text channels in the Chapters category
        chapter_channels = [channel for channel in chapters_category.text_channels]
        if len(chapter_channels) <= 1:
            print("📖 Not enough chapter channels to sort")
            return

        print(f"📖 Sorting {len(chapter_channels)} chapter channels alphabetically...")

        # Separate unaffiliated from other channels
        unaffiliated_channels = [ch for ch in chapter_channels if ch.name == "unaffiliated"]
        other_channels = [ch for ch in chapter_channels if ch.name != "unaffiliated"]

        # Sort other channels alphabetically
        other_channels.sort(key=lambda ch: ch.name.lower())

        # Combine: other channels first, then unaffiliated at the bottom
        final_order = other_channels + unaffiliated_channels

        # Update positions within the category
        for i, channel in enumerate(final_order):
            if channel.position != i:
                try:
                    result = await handle_rate_limit(
                        channel.edit(position=i, reason="Sorting chapter channels alphabetically"),
                        f"moving channel '{channel.name}'"
                    )
                    if result is not None:
                        print(f"📖 Moved #{channel.name} to position {i}")
                except Exception as e:
                    print(f"❌ Error moving #{channel.name}: {e}")

        print("✅ Chapter channels sorted alphabetically (unaffiliated at bottom)")

    except Exception as e:
        print(f"❌ Error sorting chapter channels: {e}")

async def search_and_share_useful_links(guild):
    """Search for Useful Links folder and share with volunteers"""
    try:
        guild_id = guild.id
        print(f"🔍 DEBUG: Searching for Useful Links folder for guild {guild_id}")

        # Check if we have a connected spreadsheet to get the folder ID
        if guild_id not in spreadsheets:
            print(f"❌ DEBUG: No spreadsheet connected for guild {guild_id}, cannot search for Useful Links folder")
            return

        guild_spreadsheet = spreadsheets[guild_id]
        print(f"✅ DEBUG: Spreadsheet connected, ID: {guild_spreadsheet.id}")

        # Import Drive API
        from googleapiclient.discovery import build

        # Build Drive API service
        drive_service = build('drive', 'v3', credentials=creds)

        # Get the parent folder ID of the connected spreadsheet
        sheet_metadata = drive_service.files().get(fileId=guild_spreadsheet.id, fields='parents').execute()
        parent_folders = sheet_metadata.get('parents', [])

        if not parent_folders:
            print(f"❌ DEBUG: Could not find parent folder for the spreadsheet")
            return

        parent_folder_id = parent_folders[0]
        print(f"✅ DEBUG: Found parent folder ID: {parent_folder_id}")

        # Search for "Useful Links" folder in the parent directory
        useful_links_query = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='Useful Links'"
        useful_links_results = drive_service.files().list(q=useful_links_query, fields='files(id, name, webViewLink)').execute()
        useful_links_folders = useful_links_results.get('files', [])

        if not useful_links_folders:
            print(f"❌ DEBUG: No 'Useful Links' folder found in the parent directory")
            return

        useful_links_folder = useful_links_folders[0]
        useful_links_folder_id = useful_links_folder['id']
        print(f"✅ DEBUG: Found Useful Links folder: {useful_links_folder_id}")

        # Get all files in the Useful Links folder
        files_query = f"'{useful_links_folder_id}' in parents and trashed=false"
        files_results = drive_service.files().list(q=files_query, fields='files(id, name, webViewLink, mimeType)').execute()
        files = files_results.get('files', [])

        if not files:
            print(f"❌ DEBUG: No files found in Useful Links folder")
            return

        print(f"✅ DEBUG: Found {len(files)} files in Useful Links folder")

        # Find the volunteers useful-links channel
        target_channel = discord.utils.get(guild.text_channels, name="useful-links")

        if not target_channel:
            print(f"❌ DEBUG: Could not find useful-links channel")
            return

        print(f"✅ DEBUG: Found target channel: #{target_channel.name}")

        # Delete old pinned useful links messages from the bot
        pinned_messages = await safe_call(target_channel.pins())
        deleted_count = 0
        
        for message in pinned_messages:
            # Check if this is a useful links message sent by the bot
            if message.author == bot.user and message.embeds:
                if message.embeds[0].title and "🔗 Useful Links & Resources" in message.embeds[0].title:
                    try:
                        await message.delete()
                        deleted_count += 1
                        print(f"🗑️ Deleted old pinned useful links message from #{target_channel.name}")
                        # Small delay to avoid rate limiting
                        await asyncio.sleep(0.2)
                    except Exception as delete_error:
                        print(f"⚠️ Could not delete useful links message in #{target_channel.name}: {delete_error}")
        
        if deleted_count > 0:
            print(f"✅ Cleaned up {deleted_count} old useful links message(s)")

        # Create embed for the useful links
        embed = discord.Embed(
            title="🔗 Useful Links & Resources",
            description="Access important links and resources for volunteers!",
            color=discord.Color.green()
        )

        # Create file links as bullet points
        file_links = []
        for file in files:
            file_name = file['name']
            file_link = file['webViewLink']

            # Determine file type emoji
            mime_type = file.get('mimeType', '')
            if 'pdf' in mime_type:
                emoji = "📄"
            elif 'document' in mime_type:
                emoji = "📝"
            elif 'spreadsheet' in mime_type:
                emoji = "📊"
            elif 'presentation' in mime_type:
                emoji = "📖"
            elif 'image' in mime_type:
                emoji = "🖼️"
            elif 'folder' in mime_type:
                emoji = "📁"
            else:
                emoji = "📎"

            file_links.append(f"• {emoji} [**{file_name}**]({file_link})")

        # Split into chunks if too long for Discord (1024 character limit per field)
        # Discord's limit is 1024 chars per field, so we use 1000 to be safe
        chunk_size = 1000
        chunks = []
        current_chunk = ""

        for link in file_links:
            # Check if adding this link would exceed the limit
            if len(current_chunk + link + "\n") > chunk_size:
                # Save current chunk and start a new one
                if current_chunk.strip():
                    chunks.append(current_chunk.strip())
                current_chunk = link + "\n"
            else:
                current_chunk += link + "\n"

        # Don't forget the last chunk
        if current_chunk.strip():
            chunks.append(current_chunk.strip())

        # Send the first message with the main embed
        embed.add_field(name="📋 Useful Links", value=chunks[0] if chunks else "No files found", inline=False)
        message = await target_channel.send(embed=embed)
        print(f"🔗 Shared useful links in #{target_channel.name}")

        # Pin the first message
        try:
            await safe_call(message.pin())
            print(f"📌 Pinned useful links message in #{target_channel.name}")
        except discord.Forbidden:
            print(f"⚠️ No permission to pin message in #{target_channel.name}")
        except Exception as pin_error:
            print(f"⚠️ Error pinning message in #{target_channel.name}: {pin_error}")

        # Send additional messages for remaining chunks
        if len(chunks) > 1:
            for i, chunk in enumerate(chunks[1:], start=2):
                continuation_embed = discord.Embed(
                    title=f"🔗 Useful Links & Resources (continued {i})",
                    description="",
                    color=discord.Color.green()
                )
                continuation_embed.add_field(name="📋 Useful Links", value=chunk, inline=False)
                await target_channel.send(embed=continuation_embed)
                print(f"🔗 Sent continuation message {i} for useful links")
                await asyncio.sleep(0.5)

    except Exception as e:
        print(f"❌ Error searching for Useful Links folder: {e}")

async def search_and_share_runner_info(guild):
    """Search for Runner folder and share with runner channel"""
    try:
        guild_id = guild.id
        print(f"🔍 DEBUG: Searching for Runner folder for guild {guild_id}")

        # Check if we have a connected spreadsheet to get the folder ID
        if guild_id not in spreadsheets:
            print(f"❌ DEBUG: No spreadsheet connected for guild {guild_id}, cannot search for Runner folder")
            return

        guild_spreadsheet = spreadsheets[guild_id]
        print(f"✅ DEBUG: Spreadsheet connected, ID: {guild_spreadsheet.id}")

        # Import Drive API
        from googleapiclient.discovery import build

        # Build Drive API service
        drive_service = build('drive', 'v3', credentials=creds)

        # Get the parent folder ID of the connected spreadsheet
        sheet_metadata = drive_service.files().get(fileId=guild_spreadsheet.id, fields='parents').execute()
        parent_folders = sheet_metadata.get('parents', [])

        if not parent_folders:
            print(f"❌ DEBUG: Could not find parent folder for the spreadsheet")
            return

        parent_folder_id = parent_folders[0]
        print(f"✅ DEBUG: Found parent folder ID: {parent_folder_id}")

        # Search for "Runner" folder in the parent directory
        runner_query = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='Runner'"
        runner_results = drive_service.files().list(q=runner_query, fields='files(id, name, webViewLink)').execute()
        runner_folders = runner_results.get('files', [])

        if not runner_folders:
            print(f"❌ DEBUG: No 'Runner' folder found in the parent directory")
            return

        runner_folder = runner_folders[0]
        runner_folder_id = runner_folder['id']
        print(f"✅ DEBUG: Found Runner folder: {runner_folder_id}")

        # Get all files in the Runner folder
        files_query = f"'{runner_folder_id}' in parents and trashed=false"
        files_results = drive_service.files().list(q=files_query, fields='files(id, name, webViewLink, mimeType)').execute()
        files = files_results.get('files', [])

        if not files:
            print(f"❌ DEBUG: No files found in Runner folder")
            return

        print(f"✅ DEBUG: Found {len(files)} files in Runner folder")

        # Find the runner channel
        target_channel = discord.utils.get(guild.text_channels, name="runner")

        if not target_channel:
            print(f"❌ DEBUG: Could not find runner channel")
            return

        print(f"✅ DEBUG: Found target channel: #{target_channel.name}")

        # Delete old pinned runner info messages from the bot
        pinned_messages = await safe_call(target_channel.pins())
        deleted_count = 0
        
        for message in pinned_messages:
            # Check if this is a runner info message sent by the bot
            if message.author == bot.user and message.embeds:
                if message.embeds[0].title and "🏃 Runner Information & Resources" in message.embeds[0].title:
                    try:
                        await message.delete()
                        deleted_count += 1
                        print(f"🗑️ Deleted old pinned runner info message from #{target_channel.name}")
                        # Small delay to avoid rate limiting
                        await asyncio.sleep(0.2)
                    except Exception as delete_error:
                        print(f"⚠️ Could not delete runner info message in #{target_channel.name}: {delete_error}")
        
        if deleted_count > 0:
            print(f"✅ Cleaned up {deleted_count} old runner info message(s)")

        # Create embed for the runner info
        embed = discord.Embed(
            title="🏃 Runner Information & Resources",
            description="Access important information and resources for runners!",
            color=discord.Color.blue()
        )

        # Create file links as bullet points
        file_links = []
        for file in files:
            file_name = file['name']
            file_link = file['webViewLink']

            # Determine file type emoji
            mime_type = file.get('mimeType', '')
            if 'pdf' in mime_type:
                emoji = "📄"
            elif 'document' in mime_type:
                emoji = "📝"
            elif 'spreadsheet' in mime_type:
                emoji = "📊"
            elif 'presentation' in mime_type:
                emoji = "📖"
            elif 'image' in mime_type:
                emoji = "🖼️"
            elif 'folder' in mime_type:
                emoji = "📁"
            else:
                emoji = "📎"

            file_links.append(f"• {emoji} [**{file_name}**]({file_link})")

        # Split into chunks if too long for Discord (1024 character limit per field)
        # Discord's limit is 1024 chars per field, so we use 1000 to be safe
        chunk_size = 1000
        chunks = []
        current_chunk = ""

        for link in file_links:
            # Check if adding this link would exceed the limit
            if len(current_chunk + link + "\n") > chunk_size:
                # Save current chunk and start a new one
                if current_chunk.strip():
                    chunks.append(current_chunk.strip())
                current_chunk = link + "\n"
            else:
                current_chunk += link + "\n"

        # Don't forget the last chunk
        if current_chunk.strip():
            chunks.append(current_chunk.strip())

        # Send the first message with the main embed
        embed.add_field(name="📋 Runner Info", value=chunks[0] if chunks else "No files found", inline=False)
        message = await target_channel.send(embed=embed)
        print(f"🏃 Shared runner info in #{target_channel.name}")

        # Pin the first message
        try:
            await safe_call(message.pin())
            print(f"📌 Pinned runner info message in #{target_channel.name}")
        except discord.Forbidden:
            print(f"⚠️ No permission to pin message in #{target_channel.name}")
        except Exception as pin_error:
            print(f"⚠️ Error pinning message in #{target_channel.name}: {pin_error}")

        # Send additional messages for remaining chunks
        if len(chunks) > 1:
            for i, chunk in enumerate(chunks[1:], start=2):
                continuation_embed = discord.Embed(
                    title=f"🏃 Runner Information & Resources (continued {i})",
                    description="",
                    color=discord.Color.blue()
                )
                continuation_embed.add_field(name="📋 Runner Info", value=chunk, inline=False)
                await target_channel.send(embed=continuation_embed)
                print(f"🏃 Sent continuation message {i} for runner info")
                await asyncio.sleep(0.5)

    except Exception as e:
        print(f"❌ Error searching for Runner folder: {e}")

async def add_runner_access(channel, runner_role):
    """Add Runner role access to a channel"""
    if not channel or not runner_role:
        return

    try:
        # Get current overwrites
        overwrites = channel.overwrites

        # Add Runner role with full permissions
        overwrites[runner_role] = discord.PermissionOverwrite(
            read_messages=True,
            send_messages=True,
            read_message_history=True
        )

        # Update channel permissions
        await handle_rate_limit(
            channel.edit(overwrites=overwrites, reason=f"Added {runner_role.name} access to all channels"),
            f"editing channel '{channel.name}' permissions"
        )
        print(f"🔑 Added {runner_role.name} access to #{channel.name}")

    except discord.Forbidden:
        print(f"❌ No permission to edit channel permissions for #{channel.name}")
    except Exception as e:
        print(f"❌ Error updating channel permissions for #{channel.name}: {e}")

async def ensure_runner_tournament_officials_access(guild, runner_role):
    """Ensure Runner role has access to Tournament Officials channels"""
    if not runner_role:
        return

    print(f"🔑 Ensuring {runner_role.name} access to Tournament Officials channels...")

    # Get Tournament Officials category
    tournament_officials_category = discord.utils.get(guild.categories, name="Tournament Officials")
    if not tournament_officials_category:
        print("⚠️ Tournament Officials category not found, skipping access setup")
        return

    # List of Tournament Officials channels
    official_channels = ["runner", "scoring", "awards-ceremony"]

    added_count = 0
    for channel_name in official_channels:
        channel = discord.utils.get(guild.text_channels, name=channel_name)
        if channel and channel.category == tournament_officials_category:
            try:
                await add_runner_access(channel, runner_role)
                added_count += 1
            except Exception as e:
                print(f"❌ Error adding Runner access to #{channel_name}: {e}")

    print(f"✅ Added {runner_role.name} access to {added_count} Tournament Officials channels")

async def send_building_welcome_message(guild, building_chat, building):
    """Send an initial welcome message to a building chat with all events in that building"""
    if not building_chat or not building:
        return

    try:
        # Get all events in this building
        building_events = await get_building_events(guild.id, building)

        if not building_events:
            print(f"⚠️ No events found for building '{building}', skipping welcome message")
            return
        
        # Sort the events alphabetically
        building_events.sort(key=lambda x: x[0].lower())

        # Create the welcome message
        embed = discord.Embed(
            title=f"🏢 Welcome to {building}!",
            description=f"This is the general chat for everyone with events in **{building}**.",
            color=discord.Color.blue()
        )

        # Add events list
        events_text = ""
        for event, room in building_events:
            if room:
                events_text += f"• **{event}** - {room}\n"
            else:
                events_text += f"• **{event}**\n"

        embed.add_field(
            name="📋 Events in this building:",
            value=events_text,
            inline=False
        )

        embed.add_field(
            name="💬 How to use this chat:",
            value="• Coordinate with other events in your building\n• Share building-specific information\n• Ask questions about the venue\n• Connect with nearby events",
            inline=False
        )

        embed.set_footer(text="Each event also has its own dedicated channel for event-specific discussions.")

        # Send the message
        message = await building_chat.send(embed=embed)
        await asyncio.sleep(0.5)
        print(f"🏢 Sent welcome message to #{building_chat.name} for building '{building}'")

        # Pin the message so it's always visible
        try:
            await safe_call(message.pin())
            print(f"📌 Pinned welcome message in #{building_chat.name}")
        except discord.Forbidden:
            print(f"⚠️ Could not pin welcome message in #{building_chat.name} (missing permissions)")
        except Exception as e:
            print(f"⚠️ Error pinning welcome message in #{building_chat.name}: {e}")

    except Exception as e:
        print(f"❌ Error sending welcome message to #{building_chat.name}: {e}")


async def add_role_to_building_chat(channel, role):
    """Add a role to a building chat channel permissions"""
    if not channel:
        print(f"❌ No channel : {channel} found")
        return
    if not role:
        print(f"❌ No role : {role} found")
        return

    try:
        # Get current overwrites
        overwrites = channel.overwrites

        # Set @everyone to not see the channel
        overwrites[channel.guild.default_role] = discord.PermissionOverwrite(read_messages=False)

        # Add the role with permissions to see and participate
        overwrites[role] = discord.PermissionOverwrite(
            read_messages=True,
            send_messages=True,
            read_message_history=True
        )

        # Update channel permissions
        await handle_rate_limit(
            channel.edit(overwrites=overwrites, reason=f"Added {role.name} to building chat access"),
            f"editing building chat '{channel.name}' permissions"
        )
        print(f"🔒 Added {role.name} access to #{channel.name}")

    except discord.Forbidden:
        print(f"❌ No permission to edit channel permissions for #{channel.name}")
    except Exception as e:
        print(f"❌ Error updating channel permissions for #{channel.name}: {e}")

async def reset_server_for_guild(guild):
    """⚠️ DANGER: Completely reset the server by deleting all channels, categories, roles, and nicknames"""
    if not guild:
        print("❌ Guild not provided!")
        return

    print("⚠️ ⚠️ ⚠️  STARTING COMPLETE SERVER RESET  ⚠️ ⚠️ ⚠️")
    print("🧨 This will delete EVERYTHING and reset all nicknames!")
    print("⏰ Starting in 3 seconds... (Ctrl+C to cancel)")

    await asyncio.sleep(3)

    print("🗑️ Starting server reset...")

    # Reset all member nicknames
    print("📝 Resetting all member nicknames...")
    nickname_count = 0
    for member in guild.members:
        if member.nick and not member.bot:  # Don't reset bot nicknames
            try:
                await handle_rate_limit(
                    member.edit(nick=None, reason="Server reset - clearing nickname"),
                    f"resetting nickname for {member}"
                )
                nickname_count += 1
                print(f"📝 Reset nickname for {member.display_name}")
            except discord.Forbidden:
                print(f"❌ No permission to reset nickname for {member.display_name}")
            except Exception as e:
                print(f"⚠️ Error resetting nickname for {member.display_name}: {e}")
    print(f"✅ Reset {nickname_count} nicknames")

    # Delete all text channels
    print("🗑️ Deleting all text channels...")
    channel_count = 0
    for channel in guild.text_channels:
        try:
            await safe_call(channel.delete(reason="Server reset"))
            channel_count += 1
            print(f"🗑️ Deleted text channel: #{channel.name}")
        except discord.Forbidden:
            print(f"❌ No permission to delete channel #{channel.name}")
        except Exception as e:
            print(f"⚠️ Error deleting channel #{channel.name}: {e}")

    # Delete all voice channels
    print("🗑️ Deleting all voice channels...")
    voice_count = 0
    for channel in guild.voice_channels:
        try:
            await safe_call(channel.delete(reason="Server reset"))
            voice_count += 1
            print(f"🗑️ Deleted voice channel: {channel.name}")
        except discord.Forbidden:
            print(f"❌ No permission to delete voice channel {channel.name}")
        except Exception as e:
            print(f"⚠️ Error deleting voice channel {channel.name}: {e}")

    # Delete all forum channels
    print("🗑️ Deleting all forum channels...")
    forum_count = 0
    for channel in guild.channels:
        if hasattr(channel, 'type') and channel.type == discord.ChannelType.forum:
            try:
                await safe_call(channel.delete(reason="Server reset"))
                forum_count += 1
                print(f"🗑️ Deleted forum channel: #{channel.name}")
            except discord.Forbidden:
                print(f"❌ No permission to delete forum #{channel.name}")
            except Exception as e:
                print(f"⚠️ Error deleting forum #{channel.name}: {e}")

    # Delete all categories
    print("🗑️ Deleting all categories...")
    category_count = 0
    for category in guild.categories:
        try:
            await safe_call(category.delete(reason="Server reset"))
            category_count += 1
            print(f"🗑️ Deleted category: {category.name}")
        except discord.Forbidden:
            print(f"❌ No permission to delete category {category.name}")
        except Exception as e:
            print(f"⚠️ Error deleting category {category.name}: {e}")

    # Delete all custom roles (keep @everyone and bot roles)
    print("🗑️ Deleting all custom roles...")
    role_count = 0
    for role in guild.roles:
        # Skip @everyone, bot roles, and roles higher than bot's highest role
        if (role.name != "@everyone" and
            not role.managed and
            role < guild.me.top_role):
            try:
                await safe_call(role.delete(reason="Server reset"))
                role_count += 1
                print(f"🗑️ Deleted role: {role.name}")
            except discord.Forbidden:
                print(f"❌ No permission to delete role {role.name}")
            except Exception as e:
                print(f"⚠️ Error deleting role {role.name}: {e}")

    print("🧨 SERVER RESET COMPLETE!")
    print(f"📊 Summary:")
    print(f"   • {nickname_count} nicknames reset")
    print(f"   • {channel_count} text channels deleted")
    print(f"   • {voice_count} voice channels deleted")
    print(f"   • {forum_count} forum channels deleted")
    print(f"   • {category_count} categories deleted")
    print(f"   • {role_count} roles deleted")
    print("🏗️ Server is now completely clean and ready for fresh setup!")

async def post_welcome_instructions(welcome_channel):
    """Post welcome instructions and login information to the welcome channel"""
    try:
        # Check if there are already messages in the channel
        async for message in welcome_channel.history(limit=10):
            # If there are messages, check if any are from the bot with the welcome embed
            if message.author == bot.user and message.embeds:
                for embed in message.embeds:
                    if embed.title and "Welcome to the Science Olympiad Server" in embed.title:
                        print(f"✅ Welcome instructions already posted in #{welcome_channel.name}")
                        return

        # Create welcome embed
        embed = discord.Embed(
            title="🎉 Welcome to the Science Olympiad Server!",
            description="Thank you for joining our Science Olympiad community! This server helps coordinate events, volunteers, and communication.",
            color=discord.Color.blue()
        )

        embed.add_field(
            name="🔐 Getting Started - Login Required",
            value="**To access all channels and get your roles, you need to login:**\n\n"
                  "1️⃣ Type `/login email:your@email.com password:yourpassword`\n"
                  "2️⃣ Replace with your actual email address and the password you received in your volunteer info email\n"
                  "3️⃣ Get instant access to your assigned channels!\n\n"
                  "✅ You'll automatically receive:\n"
                  "• Your assigned roles\n"
                  "• Access to relevant channels\n"
                  "• Your building and room information\n"
                  "• Updated nickname with your event",
            inline=False
        )

        embed.add_field(
            name="📍 What You Can Do Right Now",
            value="Even before logging in, you can:\n"
                  "• Read announcements in this channel\n"
                  "• Browse volunteer channels for general info\n"
                  "• Ask questions in the help forum\n"
                  "• Start sobbing uncontrollably",
            inline=False
        )

        embed.add_field(
            name="❓ Need Help?",
            value="• **Can't find your email?** Contact an admin\n"
                  "• **Questions about your assignment?** Ask in volunteer channels\n"
                  "• **Technical problems?** Mention an admin or moderator",
            inline=False
        )

        embed.add_field(
            name="🎯 Important Notes",
            value="• Your email must be in our system to login\n"
                  "• Each email can only be linked to one Discord account\n"
                  "• Your nickname will be updated to show your event\n"
                  "• Channels will appear based on your assigned roles",
            inline=False
        )

        embed.set_footer(text="Use /login to get started! • Questions? Ask in volunteer channels")

        # Send the welcome message
        await welcome_channel.send(embed=embed)
        print(f"📋 Posted welcome instructions to #{welcome_channel.name}")

    except Exception as e:
        print(f"❌ Error posting welcome instructions: {e}")

async def post_welcome_tldr(welcome_channel):
    """Post welcome instructions and login information to the welcome channel"""
    try:
        # Check if TLDR message already exists
        async for message in welcome_channel.history(limit=10):
            if message.author == bot.user and message.embeds:
                for embed in message.embeds:
                    if embed.title and "TLDR: TYPE" in embed.title:
                        print(f"✅ Welcome TLDR already posted in #{welcome_channel.name}")
                        return

        # Create welcome embed
        embed = discord.Embed(
            title="TLDR: TYPE `/login` TO GET STARTED",
            description="Read below message for more info",
            color=discord.Color.blue()
        )

        # Send the welcome message
        await welcome_channel.send(embed=embed)
        print(f"📋 Posted welcome tldr to #{welcome_channel.name}")

    except Exception as e:
        print(f"❌ Error posting welcome tldr: {e}")

async def setup_static_channels_for_guild(guild):
    """Create static categories and channels for Tournament Officials and Volunteers"""
    if not guild:
        print("❌ Guild not provided!")
        return

    print(f"🏗️ Setting up static channels for {guild.name}...")

    # Get or create Runner role for permissions
    runner_role = await get_or_create_role(guild, "Runner")
    # Get or create Awards role for awards-ceremony access
    awards_role = await get_or_create_role(guild, "Awards")

    # Welcome Category
    print("👋 Setting up Welcome category...")
    welcome_category = await get_or_create_category(guild, "Welcome")
    if welcome_category:
        # Create welcome channel (visible to everyone)
        welcome_channel = await get_or_create_channel(guild, "welcome", welcome_category)

        # Post welcome instructions
        if welcome_channel:
            await post_welcome_tldr(welcome_channel)
            await post_welcome_instructions(welcome_channel)

    # Tournament Officials Category
    print("📋 Setting up Tournament Officials category...")
    tournament_officials_category = await get_or_create_category(guild, "Tournament Officials")
    if tournament_officials_category:
        # Create channels in Tournament Officials category (restricted to Runner role only)
        official_channels = ["runner", "scoring", "awards-ceremony"]
        for channel_name in official_channels:
            # Check if channel already exists
            channel = discord.utils.get(guild.text_channels, name=channel_name)
            if not channel:
                # Create new channel with restricted permissions
                try:
                    overwrites = {}
                    # Hide from @everyone
                    overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)
                    # Give Runner role full access
                    if runner_role:
                        overwrites[runner_role] = discord.PermissionOverwrite(
                            read_messages=True,
                            send_messages=True,
                            read_message_history=True
                        )
                    # Give Awards role access to awards-ceremony channel
                    if channel_name == "awards-ceremony" and awards_role:
                        overwrites[awards_role] = discord.PermissionOverwrite(
                            read_messages=True,
                            send_messages=True,
                            read_message_history=True
                        )

                    channel = await handle_rate_limit(
                        guild.create_text_channel(
                            name=channel_name,
                            category=tournament_officials_category,
                            overwrites=overwrites,
                            reason="Auto-created by LAM Bot - Tournament Officials only"
                        ),
                        f"creating channel '{channel_name}'"
                    )
                    if channel_name == "awards-ceremony":
                        print(f"📺 Created restricted channel: '#{channel_name}' (Runner + Awards)")
                    else:
                        print(f"📺 Created restricted channel: '#{channel_name}' (Runner only)")

                    # Ensure Runner access is properly added after channel creation
                    if runner_role:
                        try:
                            await add_runner_access(channel, runner_role)
                            print(f"✅ Ensured Runner access to #{channel_name}")
                        except Exception as e:
                            print(f"❌ Error ensuring Runner access to #{channel_name}: {e}")

                    # Ensure Awards role access to awards-ceremony channel
                    if channel_name == "awards-ceremony" and awards_role:
                        try:
                            await add_runner_access(channel, awards_role)  # Reuse the same function
                            print(f"✅ Ensured Awards role access to #{channel_name}")
                        except Exception as e:
                            print(f"❌ Error ensuring Awards access to #{channel_name}: {e}")

                except discord.Forbidden:
                    print(f"❌ No permission to create channel '{channel_name}'")
                except Exception as e:
                    print(f"❌ Error creating channel '{channel_name}': {e}")
            else:
                # Update existing channel to be restricted
                print(f"✅ Channel '#{channel_name}' already exists, updating permissions...")
                try:
                    overwrites = channel.overwrites
                    # Hide from @everyone
                    overwrites[guild.default_role] = discord.PermissionOverwrite(read_messages=False)
                    # Give Runner role full access
                    if runner_role:
                        overwrites[runner_role] = discord.PermissionOverwrite(
                            read_messages=True,
                            send_messages=True,
                            read_message_history=True
                        )
                    # Give Awards role access to awards-ceremony channel
                    if channel_name == "awards-ceremony" and awards_role:
                        overwrites[awards_role] = discord.PermissionOverwrite(
                            read_messages=True,
                            send_messages=True,
                            read_message_history=True
                        )

                    await handle_rate_limit(
                        channel.edit(overwrites=overwrites, reason="Updated to restrict to Runner role only"),
                        f"editing channel '{channel_name}' permissions"
                    )
                    if channel_name == "awards-ceremony":
                        print(f"🔒 Updated #{channel_name} to be Runner + Awards")
                    else:
                        print(f"🔒 Updated #{channel_name} to be Runner-only")
                except Exception as e:
                    print(f"❌ Error updating permissions for #{channel_name}: {e}")

                # Ensure Runner access is properly added after channel creation/update
                if runner_role:
                    try:
                        await add_runner_access(channel, runner_role)
                        print(f"✅ Ensured Runner access to #{channel_name}")
                    except Exception as e:
                        print(f"❌ Error ensuring Runner access to #{channel_name}: {e}")

                # Ensure Awards role access to awards-ceremony channel
                if channel_name == "awards-ceremony" and awards_role:
                    try:
                        await add_runner_access(channel, awards_role)  # Reuse the same function
                        print(f"✅ Ensured Awards role access to #{channel_name}")
                    except Exception as e:
                        print(f"❌ Error ensuring Awards access to #{channel_name}: {e}")

    # Chapters Category
    print("📖 Setting up Chapters category...")
    chapters_category = await get_or_create_category(guild, "Chapters")

    # Volunteers Category
    print("🙋 Setting up Volunteers category...")
    volunteers_category = await get_or_create_category(guild, "Volunteers")
    if volunteers_category:
        # Create regular text channels
        volunteer_text_channels = ["general", "useful-links", "announcements", "random"]
        for channel_name in volunteer_text_channels:
            channel = await get_or_create_channel(guild, channel_name, volunteers_category)
            if (channel_name == "useful-links" or channel_name == "announcements"):
                try:
                    # Get current overwrites
                    overwrites = channel.overwrites

                    # Add Runner role with full permissions
                    overwrites[guild.default_role] = discord.PermissionOverwrite(
                        read_messages=True,
                        send_messages=False,
                        read_message_history=True
                    )

                    # Update channel permissions
                    await handle_rate_limit(
                        channel.edit(overwrites=overwrites, reason=f"Removed write access from '{channel.name}' for default role "),
                        f"editing channel '{channel.name}' permissions"
                    )
                    print(f"🔑 Removed write access to #{channel.name} for default role")

                except discord.Forbidden:
                    print(f"❌ No permission to edit channel permissions for #{channel.name}")
                except Exception as e:
                    print(f"❌ Error updating channel permissions for #{channel.name}: {e}")


        # Create forum channel for "help"
        help_channel = None

        # Try to find existing forum channel first
        for channel in guild.channels:
            if channel.name == "help" and hasattr(channel, 'type') and channel.type == discord.ChannelType.forum:
                help_channel = channel
                break

        if not help_channel:
            # Try to create forum channel
            try:
                overwrites = {}
                # Give Runner role access to forum
                if runner_role:
                    overwrites[runner_role] = discord.PermissionOverwrite(
                        read_messages=True,
                        send_messages=True,
                        read_message_history=True,
                        create_public_threads=True,
                        send_messages_in_threads=True
                    )

                # Try different methods to create forum
                if hasattr(guild, 'create_forum_channel'):
                    help_channel = await handle_rate_limit(
                        guild.create_forum_channel(
                            name="help",
                            category=volunteers_category,
                            overwrites=overwrites,
                            reason="Auto-created LAM Bot forum channel"
                        ),
                        "creating forum channel 'help'"
                    )
                    if help_channel:
                        print(f"📺 Created forum channel: '#{help_channel.name}' ✅")
                elif hasattr(guild, 'create_forum'):
                    help_channel = await handle_rate_limit(
                        guild.create_forum(
                            name="help",
                            category=volunteers_category,
                            overwrites=overwrites,
                            reason="Auto-created by LAM Bot - Volunteers help forum"
                        ),
                        "creating forum 'help'"
                    )
                    print(f"📺 Created forum channel: '#{help_channel.name}' ✅")
                else:
                    print("⚠️ Forum creation not supported in this py-cord version")
                    print("📝 Please manually create a forum channel named 'help' in the Volunteers category")
                    print("   1. Right-click the Volunteers category")
                    print("   2. Create Channel → Forum")
                    print("   3. Name it 'help'")
                    print("   4. The bot will add permissions automatically on next restart")

            except AttributeError:
                print("⚠️ Forum channels not supported in this py-cord version")
                print("📝 Please manually create a forum channel named 'help' in the Volunteers category")
            except discord.Forbidden:
                print(f"❌ No permission to create forum channel 'help'")
                print("📝 Please manually create a forum channel named 'help' in the Volunteers category")
            except Exception as e:
                print(f"❌ Error creating forum channel 'help': {e}")
                print("📝 Please manually create a forum channel named 'help' in the Volunteers category")
        else:
            print(f"✅ Forum channel 'help' already exists")

        # Runner access to help channel will be handled automatically by the static category logic

    print("✅ Finished setting up static channels")

async def move_bot_role_to_top_for_guild(guild):
    """Move the bot's role to the highest possible position and make it teal"""
    if not guild:
        print("❌ Guild not provided!")
        return

    # Check if bot has required permissions
    bot_member = guild.me
    if not bot_member.guild_permissions.manage_roles:
        print("❌ Bot missing 'Manage Roles' permission! Cannot move bot role to top.")
        return

    print("🤖 Moving bot role to top and making it teal...")

    try:
        # Find the bot's role
        bot_role = None
        for role in guild.roles:
            if role.managed and role.members and bot.user in role.members:
                bot_role = role
                break

        if not bot_role:
            print("⚠️ Could not find bot's role!")
            return

        print(f"🤖 Found bot role: '{bot_role.name}' (current position: {bot_role.position})")

        # Calculate the highest position the bot can reach
        # Bot can only move to positions below roles that are above it and unmovable
        max_possible_position = len(guild.roles) - 1  # Highest possible position

        # Check if there are unmovable roles above us
        higher_unmovable_roles = []
        for role in guild.roles:
            if role.position > bot_role.position and (role.managed or role == guild.default_role):
                higher_unmovable_roles.append(role)

        if higher_unmovable_roles:
            # Can't go above unmovable roles, so go just below the lowest unmovable role
            max_possible_position = min([r.position for r in higher_unmovable_roles]) - 1

        # Try to move the bot role to the highest possible position and make it teal
        changes_made = False

        # Try to update color to teal if it's not already
        if bot_role.color != discord.Color.teal():
            try:
                await handle_rate_limit(
                    bot_role.edit(color=discord.Color.teal(), reason="Making bot role teal"),
                    "editing bot role color"
                )
                print(f"🎨 Changed bot role color to teal")
                changes_made = True
            except discord.Forbidden:
                print(f"⚠️ Cannot change bot role color automatically (Discord restriction)")
                print(f"💡 To make the role teal: Go to Server Settings → Roles → {bot_role.name} → Change color to teal")
            except Exception as e:
                print(f"⚠️ Could not change bot role color: {e}")
        else:
            print(f"✅ Bot role already teal colored")

        # Try to move to highest position if not already there
        if bot_role.position != max_possible_position:
            try:
                await handle_rate_limit(
                    bot_role.edit(position=max_possible_position, reason="Moving bot role to top"),
                    "moving bot role position"
                )
                print(f"📈 Moved bot role to position {max_possible_position} (highest possible)")
                changes_made = True
            except discord.Forbidden:
                print(f"⚠️ Cannot move bot role automatically (Discord restriction)")
                print(f"💡 To move to top: Go to Server Settings → Roles → Drag {bot_role.name} to the top")
            except Exception as e:
                print(f"⚠️ Could not move bot role to top: {e}")
        else:
            print(f"✅ Bot role already at position {bot_role.position}")

        # Check final status and provide summary
        roles_above_bot = [r for r in guild.roles if r.position > bot_role.position and r.name != "@everyone"]
        is_teal_now = bot_role.color == discord.Color.teal()

        if not roles_above_bot and is_teal_now:
            print(f"✅ Bot role '{bot_role.name}' is perfectly optimized! (Top position + Teal color)")
        else:
            print(f"⚠️ Bot role '{bot_role.name}' needs manual adjustments:")

            if roles_above_bot:
                print(f"   📋 Position: {len(roles_above_bot)} roles still above the bot")
                for role in roles_above_bot[:3]:  # Show first 3
                    print(f"      • {role.name}")
                if len(roles_above_bot) > 3:
                    print(f"      • ... and {len(roles_above_bot)-3} more")
            else:
                print(f"   ✅ Position: At top (#{bot_role.position})")

            if not is_teal_now:
                print(f"   🎨 Color: Needs to be changed to teal manually")
            else:
                print(f"   ✅ Color: Already teal")

            print(f"\n💡 Manual steps needed:")
            print(f"   1. Go to Server Settings → Roles")
            if roles_above_bot:
                print(f"   2. Drag '{bot_role.name}' to the VERY TOP")
            if not is_teal_now:
                print(f"   {'3' if roles_above_bot else '2'}. Click '{bot_role.name}' → Change color to teal")
            print(f"   {'4' if (roles_above_bot and not is_teal_now) else ('3' if (roles_above_bot or not is_teal_now) else '2')}. Use /fixbotrole to verify, then /organizeroles to organize other roles")

    except Exception as e:
        print(f"❌ Error moving bot role to top: {e}")

async def organize_role_hierarchy_for_guild(guild):
    """Organize roles in priority order: lambot, Admin, Runner, Arbitrations, Photographer, Social Media, Lead ES, Volunteer, then others"""
    if not guild:
        print("❌ Guild not provided!")
        return

    # Check if bot has required permissions
    bot_member = guild.me
    if not bot_member.guild_permissions.manage_roles:
        print("❌ Bot missing 'Manage Roles' permission! Cannot organize role hierarchy.")
        print("💡 Please give the bot 'Manage Roles' permission in Server Settings → Roles")
        return

    # Define the priority order (higher index = higher priority/position)
    priority_roles = [
        "Volunteer",  # Lowest priority
        "Lead ES",
        "Social Media",
        "Photographer",
        "Arbitrations",
        "Awards",
        "Runner",
        "VIPer",
        "Admin",  # Highest priority (below bot role)
        # Bot role will be handled separately as highest priority
    ]

    print("📋 Organizing role hierarchy...")

    try:
        # Get all roles except @everyone
        all_roles = [role for role in guild.roles if role.name != "@everyone"]

        # Find the bot's role
        bot_role = None
        for role in all_roles:
            if role.managed and role.members and bot.user in role.members:
                bot_role = role
                break

        if not bot_role:
            print("⚠️ Could not find bot's role!")
            return

        print(f"🤖 Bot role: '{bot_role.name}' (current position: {bot_role.position})")

        # Separate roles into priority roles, chapter roles, and other roles
        priority_role_objects = []
        chapter_roles = []
        other_roles = []
        unmovable_roles = []

        for role in all_roles:
            if role == bot_role:
                continue  # Handle bot role separately
            elif role.position >= bot_role.position:
                # Can't move roles that are higher than or equal to bot's current position
                unmovable_roles.append(role)
                continue
            elif role.name in priority_roles:
                priority_role_objects.append(role)
            elif role.name == "Unaffiliated" or role.name in chapter_role_names:
                # This is a chapter role
                chapter_roles.append(role)
            else:
                other_roles.append(role)

        if unmovable_roles:
            print(f"⚠️ Cannot move {len(unmovable_roles)} roles (higher than bot): {', '.join([r.name for r in unmovable_roles])}")
            print("💡 Move the bot's role higher in Server Settings → Roles to manage these roles")

        # Sort priority roles according to the defined order
        priority_role_objects.sort(key=lambda r: priority_roles.index(r.name) if r.name in priority_roles else 999)

        # Sort chapter roles alphabetically
        chapter_roles.sort(key=lambda r: r.name.lower())

        # Sort other roles alphabetically
        other_roles.sort(key=lambda r: r.name.lower())

        # Build final order: other roles (lowest first) + chapter roles + priority roles (Volunteer through Admin)
        # Admin role will be at the top of priority roles since it's last in the list
        # Note: We won't try to move the bot role itself to avoid permission issues
        final_order = other_roles + chapter_roles + priority_role_objects

        # Update positions (start from position 1, @everyone stays at 0)
        position = 1
        moved_count = 0
        rate_limited_roles = []

        role_positions = {}
        for role in final_order:
            role_positions[role] = position
            position += 1
        try:
            await safe_call(guild.edit_role_positions(role_positions, reason="Organizing role hierarchy"))
            print(f"✅ Successfully moved all roles!")
        except Exception as e:
            print(f"⚠️ Unexpected error moving roles: {e}")

        #for role in final_order:
        #    if role.position != position:
        #        max_retries = 3
        #        retry_count = 0
        #        success = False
        #
        #        while retry_count < max_retries and not success:
        #            try:
        #                await safe_call(role.edit(position=position, reason="Organizing role hierarchy"))
        #                print(f"📋 Moved '{role.name}' to position {position}")
        #                moved_count += 1
        #                success = True
        #                # Small delay to avoid rate limits (Discord allows ~50 requests/second, but be conservative)
        #                await asyncio.sleep(0.5)  # 100ms delay between role moves
        #            except discord.Forbidden:
        #                print(f"❌ No permission to move role '{role.name}' (may be higher than bot)")
        #                success = True  # Don't retry permission errors
        #            except discord.HTTPException as e:
        #                error_msg = str(e)
        #                # Check for rate limit errors
        #                if e.status == 429 or "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
        #                    retry_count += 1
        #                    if retry_count < max_retries:
        #                        # Extract retry_after from response if available
        #                        retry_after = 1.0  # Default 1 second
        #                        if hasattr(e, 'retry_after') and e.retry_after:
        #                            retry_after = float(e.retry_after)
        #                        elif isinstance(e.response, dict) and 'retry_after' in e.response:
        #                            retry_after = float(e.response['retry_after'])
        #
        #                        print(f"⚠️ Rate limited moving role '{role.name}', waiting {retry_after}s before retry {retry_count}/{max_retries}...")
        #                        await asyncio.sleep(retry_after)
        #                    else:
        #                        print(f"⚠️ Rate limited moving role '{role.name}' after {max_retries} retries, skipping...")
        #                        rate_limited_roles.append(role.name)
        #                        success = True  # Give up on this role
        #                elif e.code == 50013:
        #                    print(f"❌ Missing permissions to move role '{role.name}'")
        #                    success = True  # Don't retry permission errors
        #                else:
        #                    print(f"⚠️ Error moving role '{role.name}': {e}")
        #                    success = True  # Don't retry other errors
        #            except Exception as e:
        #                error_msg = str(e)
        #                if "429" in error_msg or "rate limit" in error_msg.lower() or "too many requests" in error_msg.lower():
        #                    retry_count += 1
        #                    if retry_count < max_retries:
        #                        print(f"⚠️ Rate limited moving role '{role.name}', waiting 1s before retry {retry_count}/{max_retries}...")
        #                        await asyncio.sleep(1.0)
        #                    else:
        #                        print(f"⚠️ Rate limited moving role '{role.name}' after {max_retries} retries, skipping...")
        #                        rate_limited_roles.append(role.name)
        #                        success = True
        #                else:
        #                    print(f"⚠️ Unexpected error moving role '{role.name}': {e}")
        #                    success = True
        #    position += 1
        #
        #if rate_limited_roles:
        #    print(f"⚠️ Could not move {len(rate_limited_roles)} roles due to rate limits: {', '.join(rate_limited_roles)}")
        #    print("💡 These roles will be organized on the next sync or when you run /organizeroles again")
        #
        #if moved_count > 0:
        #    print(f"✅ Successfully moved {moved_count} roles!")
        #    print(f"📋 Organized order (bottom to top): {' → '.join([r.name for r in final_order])}")
        #else:
        #    print("ℹ️ No roles needed to be moved (already in correct positions)")

        # Final recommendation if there were permission issues
        #if unmovable_roles:
        #    print("\n💡 To fix permission issues:")
        #    print("1. Go to Server Settings → Roles")
        #    print(f"2. Drag '{bot_role.name}' role to the TOP of the role list")
        #    print("3. Run /organizeroles command again")

    except Exception as e:
        print(f"❌ Error organizing role hierarchy: {e}")
        if "50013" in str(e):
            print("💡 This is a permissions issue. Please ensure the bot has 'Manage Roles' permission and is high in the role hierarchy.")

async def remove_runner_access_from_building_channels_for_guild(guild):
    """Remove Runner role access from building/event channels"""
    if not guild:
        print("❌ Guild not provided!")
        return

    runner_role = discord.utils.get(guild.roles, name="Runner")
    if not runner_role:
        print("⚠️ Runner role not found")
        return

    print(f"🚫 Removing {runner_role.name} access from building/event channels...")

    removed_count = 0

    # Remove access from building/event channels
    for channel in guild.text_channels:
        if channel.category:
            # Remove access from channels that are NOT in static categories
            if channel.category.name not in ["Welcome", "Tournament Officials", "Volunteers"]:
                try:
                    # Check if Runner role has access to this channel
                    overwrites = channel.overwrites
                    if runner_role in overwrites:
                        # Remove the Runner role from overwrites
                        del overwrites[runner_role]
                        await handle_rate_limit(
                            channel.edit(overwrites=overwrites, reason=f"Removed {runner_role.name} access from building channel"),
                            f"removing access from channel '{channel.name}'"
                        )
                        removed_count += 1
                        print(f"🚫 Removed {runner_role.name} access from #{channel.name}")

                except Exception as e:
                    print(f"❌ Error removing Runner access from #{channel.name}: {e}")

    print(f"✅ Removed {runner_role.name} access from {removed_count} building/event channels")

async def give_runner_access_to_all_channels_for_guild(guild):
    """Give Runner role access only to static channels (not building/event channels)"""
    if not guild:
        print("❌ Guild not provided!")
        return

    runner_role = discord.utils.get(guild.roles, name="Runner")
    if not runner_role:
        print("⚠️ Runner role not found, will be created when needed")
        return

    print(f"🔑 Adding {runner_role.name} access to static channels only...")

    welcome_channels = 0
    tournament_official_channels = 0
    volunteer_channels = 0
    forum_channels = 0

    # Add access to all channels in specific categories
    for channel in guild.text_channels:
        if channel.category:
            if channel.category.name == "Welcome":
                try:
                    await add_runner_access(channel, runner_role)
                    welcome_channels += 1
                    print(f"🔑 Added {runner_role.name} access to #{channel.name} (Welcome)")
                except Exception as e:
                    print(f"❌ Error adding Runner access to #{channel.name}: {e}")

            elif channel.category.name == "Tournament Officials":
                try:
                    await add_runner_access(channel, runner_role)
                    tournament_official_channels += 1
                    print(f"🔑 Added {runner_role.name} access to #{channel.name} (Tournament Officials)")
                except Exception as e:
                    print(f"❌ Error adding Runner access to #{channel.name}: {e}")

            # elif channel.category.name == "Chapters":
            #     try:
            #         await add_runner_access(channel, runner_role)
            #         volunteer_channels += 1
            #         print(f"🔑 Added {runner_role.name} access to #{channel.name} (Chapters)")
            #     except Exception as e:
            #         print(f"❌ Error adding Runner access to #{channel.name}: {e}")

            elif channel.category.name == "Volunteers":
                try:
                    await add_runner_access(channel, runner_role)
                    volunteer_channels += 1
                    print(f"🔑 Added {runner_role.name} access to #{channel.name} (Volunteers)")
                except Exception as e:
                    print(f"❌ Error adding Runner access to #{channel.name}: {e}")

    # Add access to forum channels in static categories
    for channel in guild.channels:
        if channel.type == discord.ChannelType.forum and channel.category:
            if channel.category.name in ["Welcome", "Tournament Officials", "Volunteers"]:
                try:
                    overwrites = channel.overwrites
                    overwrites[runner_role] = discord.PermissionOverwrite(
                        read_messages=True,
                        send_messages=True,
                        read_message_history=True,
                        create_public_threads=True,
                        send_messages_in_threads=True
                    )
                    await handle_rate_limit(
                        channel.edit(overwrites=overwrites, reason=f"Added {runner_role.name} access"),
                        f"editing forum channel '{channel.name}' permissions"
                    )
                    print(f"🔑 Added {runner_role.name} access to #{channel.name} (forum in {channel.category.name})")
                    forum_channels += 1
                except Exception as e:
                    print(f"❌ Error adding Runner access to forum #{channel.name}: {e}")

    print(f"✅ Added {runner_role.name} access to:")
    print(f"   • {welcome_channels} Welcome channels")
    print(f"   • {tournament_official_channels} Tournament Officials channels")
    print(f"   • {volunteer_channels} Volunteers channels")
    print(f"   • {forum_channels} forum channels")
    print(f"🔑 Total: {welcome_channels + tournament_official_channels + volunteer_channels + forum_channels} channels with Runner access")
    print(f"🚫 Building/event channels are restricted to event participants only")

async def setup_ezhang_admin_role(guild):
    """Set up admin role for ezhang. if they're in the server"""
    if not guild:
        return

    ezhang_member = None
    for member in guild.members:
        if (member.name.lower() == "ezhang." or
            (member.global_name and member.global_name.lower() == "ezhang.")):
            ezhang_member = member
            break

    if ezhang_member:
        try:
            # Get or create Admin role
            admin_role = discord.utils.get(guild.roles, name="Admin")
            if not admin_role:
                admin_role = await handle_rate_limit(
                    guild.create_role(
                        name="Admin",
                        permissions=discord.Permissions.all(),
                        color=discord.Color.purple(),
                        reason="Created admin role for ezhang."
                    ),
                    "creating admin role for ezhang"
                )
                print(f"🆕 Created Admin role for ezhang. in {guild.name}")

            # Assign admin role if they don't have it
            if admin_role not in ezhang_member.roles:
                await handle_rate_limit(
                    ezhang_member.add_roles(admin_role, reason="Special admin access for ezhang."),
                    f"adding admin role to {ezhang_member}"
                )
                print(f"👑 Granted admin privileges to {ezhang_member} (ezhang.) in {guild.name}")
            else:
                print(f"✅ {ezhang_member} already has Admin role in {guild.name}")
        except Exception as e:
            print(f"⚠️ Could not grant admin privileges to ezhang. in {guild.name}: {e}")

async def generate_building_structures(guild, force_refresh_welcome=False):
    """Standalone function to generate building/room channels strictly from 'Room Assignments'"""
    print("🏗️ Generating building structures from Room Assignments...")
    guild_id = guild.id
    room_data = await get_room_assignments(guild_id)
    
    if not room_data:
        print("⚠️ No data found in Room Assignments sheet.")
        return 0, 0
        
    building_structures = set()
    buildings = set()
    
    for row in room_data:
        # Make column lookups case-insensitive and highly flexible
        lower_row = {str(k).strip().lower(): v for k, v in row.items()}
        
        building = str(lower_row.get("building", lower_row.get("building 1", ""))).strip()
        event = str(lower_row.get("events", lower_row.get("event", lower_row.get("first event", "")))).strip()
        room = str(lower_row.get("room", lower_row.get("room 1", ""))).strip()
        
        # Skip priority/custom roles
        priority_roles = ["Admin", "Volunteer", "Lead ES", "Social Media", "Photographer", "Arbitrations", "Awards", "Runner", "VIPer"]
        
        if building and event and event not in priority_roles:
            building_structures.add((building, event, room))
            buildings.add(building)
            
    print(f"🏗️ Found {len(building_structures)} unique building/event combinations to create across {len(buildings)} buildings.")
    
    # Create all building structures
    for building, event, room in building_structures:
        print(f"🏗️ Creating structure: {building} - {event} - {room}")
        await setup_building_structure(guild, building, event, room)
        
    # If forced refresh, delete old welcome messages and send new ones
    if force_refresh_welcome:
        for building in buildings:
            building_chat_name = f"{sanitize_for_discord(building)}-chat"
            building_chat = discord.utils.get(guild.text_channels, name=building_chat_name)
            
            if building_chat:
                # Delete old pinned welcome messages sent by the bot
                try:
                    pinned_messages = await safe_call(building_chat.pins())
                    for msg in pinned_messages:
                        if msg.author == bot.user and msg.embeds:
                            if msg.embeds[0].title and f"Welcome to {building}" in msg.embeds[0].title:
                                await msg.delete()
                                print(f"🗑️ Deleted old welcome message in #{building_chat.name}")
                                await asyncio.sleep(0.5)
                except Exception as e:
                    print(f"⚠️ Could not clear old welcome messages in #{building_chat.name}: {e}")
                    
                # Send a fresh welcome message
                print(f"📝 Sending refreshed welcome message to #{building_chat.name}")
                await send_building_welcome_message(guild, building_chat, building)

    # Sort categories once after all structures are created
    print("📋 Organizing all building categories alphabetically...")
    await sort_building_categories_alphabetically(guild)
    
    # NEW: Sort the channels inside those building categories
    print("📋 Organizing channels within building categories alphabetically...")
    await sort_channels_in_building_categories(guild)
    
    return len(building_structures), len(buildings)

@bot.event
async def on_ready():
    global runner_all_access
    # NEW: Load saved runner access settings from cache so they survive restarts!
    cache = load_cache()
    saved_runner_access = cache.get("runner_access_settings", {})
    # Convert string keys from JSON back to integer guild IDs
    runner_all_access = {int(k): v for k, v in saved_runner_access.items()}

    async with admin_lock:

        print(f"Logged in as {bot.user} (ID: {bot.user.id})")
        print(f"🌐 Bot is active in {len(bot.guilds)} guild(s):")
        for guild in bot.guilds:
            print(f"  • {guild.name} (ID: {guild.id}) - {guild.member_count} members")

        # Process each guild the bot is in
        for guild in bot.guilds:
            print(f"\n🏗️ Setting up guild: {guild.name} (ID: {guild.id})")

            # Check if server reset is enabled for this guild
            if RESET_SERVER:
                print(f"⚠️ ⚠️ ⚠️  SERVER RESET ENABLED FOR {guild.name}!  ⚠️ ⚠️ ⚠️")
                await reset_server_for_guild(guild)
                print(f"🔄 Reset complete for {guild.name}, continuing with setup...")

            try:
                print(f"🏗️ Setting up static channels for {guild.name}...")
                await setup_static_channels_for_guild(guild)
                print(f"🤖 Moving bot role to top for {guild.name}...")
                await move_bot_role_to_top_for_guild(guild)
                print(f"🎭 Organizing role hierarchy for {guild.name}...")
                await organize_role_hierarchy_for_guild(guild)
                # Check guild specific setting, default to 0
                if not runner_all_access.get(guild.id, 0):
                    print(f"🚫 Removing Runner access from building channels for {guild.name}...")
                    await remove_runner_access_from_building_channels_for_guild(guild)
                print(f"🔑 Adding Runner access to static channels for {guild.name}...")
                await give_runner_access_to_all_channels_for_guild(guild)

                # Check if ezhang. is already in this server and give them the Admin role
                await setup_ezhang_admin_role(guild)

            except Exception as e:
                print(f"❌ Error setting up guild {guild.name}: {e}")

        # Try to load spreadsheet connections from cache (per-guild)
        print("\n💾 Attempting to load cached spreadsheet connections...")
        cache_loaded = await load_spreadsheets_from_cache()
        if cache_loaded:
            print("✅ Successfully loaded spreadsheet connections from cache!")
        else:
            print("📋 No cached connections available - use /enterfolder to connect to a sheet")

        print("🔄 Starting member sync task...")
        sync_members.start()

        print("🎫 Starting help ticket monitoring task...")
        check_help_tickets.start()

@bot.event
async def on_guild_join(guild):
    """Handle setup when bot joins a new guild"""
    global runner_all_access
    async with admin_lock:

        print(f"🎉 Bot joined new guild: {guild.name} (ID: {guild.id}) - {guild.member_count} members")

        try:
            print(f"🏗️ Setting up new guild: {guild.name}")

            # Delete default Discord channels (general text and General voice)
            print("🗑️ Removing default Discord channels...")
            for channel in guild.channels:
                if channel.name.lower() == "general":
                    try:
                        await safe_call(channel.delete(reason="Removing default Discord channel"))
                        print(f"🗑️ Deleted default channel: {channel.name}")
                    except discord.Forbidden:
                        print(f"❌ No permission to delete channel {channel.name}")
                    except Exception as e:
                        print(f"⚠️ Error deleting channel {channel.name}: {e}")

            # Set up the guild with all the standard setup
            await setup_static_channels_for_guild(guild)
            await move_bot_role_to_top_for_guild(guild)
            await organize_role_hierarchy_for_guild(guild)
            # Check guild specific setting, default to 0
            if not runner_all_access.get(guild.id, 0):
                await remove_runner_access_from_building_channels_for_guild(guild)
            await give_runner_access_to_all_channels_for_guild(guild)
            await setup_ezhang_admin_role(guild)

            print(f"✅ Successfully set up new guild: {guild.name}")

        except Exception as e:
            print(f"❌ Error setting up new guild {guild.name}: {e}")

@bot.event
async def on_member_join(member):
    """Handle role assignment and nickname setting when a user joins the server"""
    # Special case: Give ezhang. admin privileges immediately upon joining
    if member.name.lower() == "ezhang." or member.global_name and member.global_name.lower() == "ezhang.":
        try:
            # Get or create Admin role
            admin_role = discord.utils.get(member.guild.roles, name="Admin")
            if not admin_role:
                admin_role = await handle_rate_limit(
                    member.guild.create_role(
                        name="Admin",
                        permissions=discord.Permissions.all(),
                        color=discord.Color.purple(),
                        reason="Created admin role for ezhang."
                    ),
                    "creating admin role for ezhang"
                )
                print(f"🆕 Created Admin role for ezhang.")

            # Assign admin role
            await handle_rate_limit(
                member.add_roles(admin_role, reason="Special admin access for ezhang."),
                f"adding admin role to {member}"
            )
            print(f"👑 Granted admin privileges to {member} (ezhang.) upon joining")



        except Exception as e:
            print(f"⚠️ Could not grant admin privileges to ezhang. upon joining: {e}")

    if member.id in pending_users:
        user_info = pending_users[member.id]
        role_names = user_info.get("roles", [])
        user_name = user_info.get("name", "")
        first_event = user_info.get("first_event", "")

        # Assign roles
        for role_name in role_names:
            role = await get_or_create_role(member.guild, role_name)
            if role:
                try:
                    result = await handle_rate_limit(
                        member.add_roles(role, reason="Onboarding sync"),
                        f"adding role '{role_name}' to {member}"
                    )
                    if result is not None:
                        print(f"✅ Assigned role {role.name} to {member}")
                except Exception as e:
                    print(f"⚠️ Could not add role {role_name} to {member}: {e}")

        # Set nickname if we have both name and first event
        if user_name and first_event:
            nickname = f"{user_name} ({first_event})"
            # Truncate to 32 characters (Discord limit)
            if len(nickname) > 32:
                nickname = nickname[:32]
            try:
                await handle_rate_limit(
                    member.edit(nick=nickname, reason="Onboarding sync - setting nickname"),
                    f"editing nickname for {member}"
                )
                print(f"📝 Set nickname for {member}: '{nickname}'")
            except discord.Forbidden:
                print(f"❌ No permission to set nickname for {member}")
            except Exception as e:
                print(f"⚠️ Could not set nickname for {member}: {e}")

        # Remove from pending users
        del pending_users[member.id]

async def get_room_assignments(guild_id):
    """Helper to safely fetch Room Assignments data"""
    if guild_id not in spreadsheets:
        return []
    try:
        ws = spreadsheets[guild_id].worksheet("Room Assignments")
        return ws.get_all_records()
    except Exception as e:
        print(f"⚠️ 'Room Assignments' worksheet not found or error: {e}")
        return []
    
# async def get_user_event_building(guild_id, discord_id):
#     """Look up a user's event and building from the main sheet"""
#     if guild_id not in spreadsheets:
#         print(f"❌ No spreadsheet connected for guild {guild_id}")
#         return None

#     try:
#         # Get the main worksheet
#         spreadsheet = spreadsheets[guild_id]
#         sheet = spreadsheet.worksheet(SHEET_PAGE_NAME)
#         data = sheet.get_all_records()

#         # Find the user by Discord ID
#         for row in data:
#             row_discord_id = str(row.get("Discord ID", "")).strip()
#             if not row_discord_id:
#                 continue

#             # Try to match Discord ID
#             try:
#                 if int(row_discord_id) == discord_id:
#                     # Found the user, extract their info
#                     event = str(row.get("First Event", "")).strip()
#                     building = str(row.get("Building 1", "")).strip()
#                     room = str(row.get("Room 1", "")).strip()

#                     return {
#                         "event": event if event else None,
#                         "building": building if building else None,
#                         "room": room if room else None,
#                         "name": str(row.get("Name", "")).strip()
#                     }
#             except ValueError:
#                 # Not a numeric Discord ID, skip
#                 continue

#         print(f"⚠️ User with Discord ID {discord_id} not found in sheet")
#         return None

#     except Exception as e:
#         print(f"❌ Error looking up user event/building: {e}")
#         return None


# async def get_building_events(guild_id, building):
#     """Get all events and rooms for a specific building from the main sheet"""
#     if guild_id not in spreadsheets:
#         print(f"❌ No spreadsheet connected for guild {guild_id}")
#         return []

#     try:
#         # Get the main worksheet
#         spreadsheet = spreadsheets[guild_id]
#         sheet = spreadsheet.worksheet(SHEET_PAGE_NAME)
#         data = sheet.get_all_records()

#         # Find all events in this building
#         building_events = []
#         for row in data:
#             row_building = str(row.get("Building 1", "")).strip()
#             if row_building.lower() == building.lower():
#                 event = str(row.get("First Event", "")).strip()
#                 room = str(row.get("Room 1", "")).strip()

#                 # Skip priority/custom roles (only include actual events)
#                 priority_roles = ["Admin", "Volunteer", "Lead ES", "Social Media", "Photographer", "Arbitrations", "Awards", "Runner", "VIPer"]
#                 if event and event not in priority_roles:
#                     # Create a tuple of (event, room) to avoid duplicates
#                     event_room_combo = (event, room if room else "")
#                     if event_room_combo not in building_events:
#                         building_events.append(event_room_combo)

#         print(f"🏢 Found {len(building_events)} events in building '{building}': {building_events}")
#         return building_events

#     except Exception as e:
#         print(f"❌ Error looking up building events: {e}")
#         return []

async def get_user_event_building(guild_id, discord_id):
    """Look up a user's event from the main sheet, then find their building/room in Room Assignments"""
    if guild_id not in spreadsheets:
        print(f"❌ No spreadsheet connected for guild {guild_id}")
        return None

    try:
        # Get the main worksheet
        spreadsheet = spreadsheets[guild_id]
        sheet = spreadsheet.worksheet(SHEET_PAGE_NAME)
        data = sheet.get_all_records()

        user_event = None
        user_name = None

        # Find the user by Discord ID
        for row in data:
            row_discord_id = str(row.get("Discord ID", "")).strip()
            if not row_discord_id:
                continue

            # Try to match Discord ID
            try:
                if int(row_discord_id) == discord_id:
                    # Found the user, extract their info
                    roles_raw = str(row.get("Roles", "")).strip()
                    roles = [r.strip() for r in roles_raw.split(";") if r.strip()]
                    user_event = roles[0] if roles else None
                    user_name = str(row.get("Name", "")).strip()
                    break
            except ValueError:
                # Not a numeric Discord ID, skip
                continue

        if not user_name:
            print(f"⚠️ User with Discord ID {discord_id} not found in sheet")
            return None

        # Look up their building and room in Room Assignments
        building = None
        room = None
        
        if user_event:
            room_data = await get_room_assignments(guild_id)
            for r_row in room_data:
                r_event = str(r_row.get("Events", "")).strip()
                if r_event.lower() == user_event.lower():
                    building = str(r_row.get("Building", "")).strip()
                    room = str(r_row.get("Room", "")).strip()
                    break

        return {
            "event": user_event,
            "building": building,
            "room": room,
            "name": user_name
        }

    except Exception as e:
        print(f"❌ Error looking up user event/building: {e}")
        return None


async def get_building_events(guild_id, building):
    """Get all events and rooms for a specific building from the Room Assignments sheet"""
    if guild_id not in spreadsheets:
        print(f"❌ No spreadsheet connected for guild {guild_id}")
        return []

    try:
        room_data = await get_room_assignments(guild_id)
        building_events = []
        
        # Find all events in this building
        for row in room_data:
            row_building = str(row.get("Building", "")).strip()
            if row_building.lower() == building.lower():
                event = str(row.get("Events", "")).strip()
                room = str(row.get("Room", "")).strip()

                # Skip priority/custom roles (only include actual events)
                priority_roles = ["Admin", "Volunteer", "Lead ES", "Social Media", "Photographer", "Arbitrations", "Awards", "Runner", "VIPer"]
                if event and event not in priority_roles:
                    # Create a tuple of (event, room) to avoid duplicates
                    event_room_combo = (event, room if room else "")
                    if event_room_combo not in building_events:
                        building_events.append(event_room_combo)

        print(f"🏢 Found {len(building_events)} events in building '{building}': {building_events}")
        return building_events

    except Exception as e:
        print(f"❌ Error looking up building events: {e}")
        return []
    
async def get_building_zone(guild_id, building):
    """Get the zone number for a building from the Runner Assignments sheet"""
    if guild_id not in spreadsheets:
        print(f"❌ No spreadsheet connected for guild {guild_id}")
        return None

    try:
        spreadsheet = spreadsheets[guild_id]

        # Try to get the Runner Assignments worksheet
        try:
            sheet = spreadsheet.worksheet("Runner Assignments")
        except Exception:
            # If not found as a worksheet, search for a separate spreadsheet
            try:
                from googleapiclient.discovery import build
                drive_service = build('drive', 'v3', credentials=creds)

                # Get parent folder of the currently connected spreadsheet
                sheet_metadata = drive_service.files().get(fileId=spreadsheet.id, fields='parents').execute()
                parent_folders = sheet_metadata.get('parents', [])
                if not parent_folders:
                    print("❌ Could not determine parent folder for Runner Assignments lookup")
                    return None

                parent_folder_id = parent_folders[0]

                # Search for Runner Assignments spreadsheet
                q = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and name contains 'Runner Assignments'"
                results = drive_service.files().list(q=q, fields='files(id, name)').execute()
                files = results.get('files', [])

                if not files:
                    print("❌ Could not find Runner Assignments spreadsheet")
                    return None

                # Open the first matching spreadsheet
                runner_spreadsheet = gc.open_by_key(files[0]['id'])
                sheet = runner_spreadsheet.sheet1  # Use first worksheet

            except Exception as e:
                print(f"❌ Error finding Runner Assignments spreadsheet: {e}")
                return None

        # Get all data from the sheet
        data = sheet.get_all_records()

        # Find the building and get its zone number
        for row in data:
            row_building = str(row.get("Building", "")).strip()
            if row_building.lower() == building.lower():
                zone = row.get("Zone Number", "")
                if zone:
                    try:
                        return int(zone)
                    except (ValueError, TypeError):
                        print(f"⚠️ Invalid zone value '{zone}' for building '{building}'")
                        return None

        print(f"⚠️ Building '{building}' not found in Runner Assignments")
        return None

    except Exception as e:
        print(f"❌ Error looking up building zone: {e}")
        return None

async def sort_channels_in_building_categories(guild):
    """Sort channels inside building categories alphabetically, keeping building-chat at the top"""
    try:
        static_categories = ["Welcome", "Tournament Officials", "Chapters", "Volunteers"]
        
        for category in guild.categories:
            # Skip static categories
            if category.name in static_categories:
                continue
                
            print(f"📋 Sorting channels in building category: '{category.name}'")
            
            # Get all text channels in this category
            channels = category.text_channels
            if len(channels) <= 1:
                continue
                
            # The expected name for the building chat
            building_chat_name = f"{sanitize_for_discord(category.name)}-chat"
            
            building_chats = []
            event_channels = []
            
            # Separate building chat from event channels
            for channel in channels:
                if channel.name == building_chat_name or channel.name.endswith("-chat"):
                    building_chats.append(channel)
                else:
                    event_channels.append(channel)
                    
            # Sort event channels alphabetically
            event_channels.sort(key=lambda c: c.name.lower())
            
            # Combine them: Building chat first, then event channels
            sorted_channels = building_chats + event_channels
            
            # Update positions within the category
            for i, channel in enumerate(sorted_channels):
                if channel.position != i:
                    try:
                        result = await handle_rate_limit(
                            channel.edit(position=i, reason=f"Sorting {category.name} channels alphabetically"),
                            f"moving channel '{channel.name}'"
                        )
                        if result is not None:
                            print(f"📺 Moved #{channel.name} to position {i} in {category.name}")
                    except Exception as e:
                        print(f"❌ Error moving #{channel.name}: {e}")

        print("✅ Finished sorting channels inside building categories")

    except Exception as e:
        print(f"⚠️ Error organizing channels inside building categories: {e}")

async def get_zone_runners(guild_id, zone):
    """Get all Discord IDs of runners assigned to a specific zone"""
    if guild_id not in spreadsheets:
        print(f"❌ No spreadsheet connected for guild {guild_id}")
        return []

    try:
        spreadsheet = spreadsheets[guild_id]

        # Try to get the Runner Assignments worksheet
        try:
            sheet = spreadsheet.worksheet("Runner Assignments")
        except Exception:
            # If not found as a worksheet, search for a separate spreadsheet
            try:
                from googleapiclient.discovery import build
                drive_service = build('drive', 'v3', credentials=creds)

                # Get parent folder of the currently connected spreadsheet
                sheet_metadata = drive_service.files().get(fileId=spreadsheet.id, fields='parents').execute()
                parent_folders = sheet_metadata.get('parents', [])
                if not parent_folders:
                    print("❌ Could not determine parent folder for zone runners lookup")
                    return []

                parent_folder_id = parent_folders[0]

                # Search for Runner Assignments spreadsheet
                q = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and name contains 'Runner Assignments'"
                results = drive_service.files().list(q=q, fields='files(id, name)').execute()
                files = results.get('files', [])

                if not files:
                    print("❌ Could not find Runner Assignments spreadsheet")
                    return []

                # Open the first matching spreadsheet
                runner_spreadsheet = gc.open_by_key(files[0]['id'])
                sheet = runner_spreadsheet.sheet1  # Use first worksheet

            except Exception as e:
                print(f"❌ Error finding Runner Assignments spreadsheet: {e}")
                return []

        # Get all data from the sheet
        data = sheet.get_all_records()

        # Find all runners in the specified zone
        runner_emails = []
        for row in data:
            row_zone = row.get("Runner Zone", "")
            if row_zone:
                try:
                    if int(row_zone) == zone:
                        # This runner is in the target zone
                        email = str(row.get("Email", "")).strip()
                        if email:
                            runner_emails.append(email.lower())
                except (ValueError, TypeError):
                    continue

        if not runner_emails:
            return []

        print(f"🔍 Found {len(runner_emails)} runner emails in zone {zone}")

        # Now cross-reference with the main sheet to get Discord IDs
        try:
            main_sheet = spreadsheet.worksheet(SHEET_PAGE_NAME)
            main_data = main_sheet.get_all_records()
        except Exception as e:
            print(f"❌ Error accessing main sheet for Discord ID lookup: {e}")
            return []

        zone_runners = []
        for row in main_data:
            email = str(row.get("Email", "")).strip().lower()
            if email in runner_emails:
                discord_id = str(row.get("Discord ID", "")).strip()
                if discord_id:
                    try:
                        zone_runners.append(int(discord_id))
                        print(f"✅ Found Discord ID {discord_id} for runner email {email}")
                    except ValueError:
                        print(f"⚠️ Invalid Discord ID '{discord_id}' for runner email {email}")

        return zone_runners

    except Exception as e:
        print(f"❌ Error looking up zone runners: {e}")
        return []


async def check_for_burger_request(thread):
    """Check if a help ticket contains '55 burgers' or 'fifty five burgers' and DM burger to creator"""
    try:
        # Check the thread title
        thread_title_lower = thread.name.lower()
        
        # Check if the phrases exist in the title
        has_burger_phrase = ("55 burgers" in thread_title_lower or 
                            "fifty five burgers" in thread_title_lower or
                            "55 burger" in thread_title_lower or
                            "fifty five burger" in thread_title_lower)
        
        # If not in title, check the initial message
        if not has_burger_phrase:
            try:
                # Get the first message in the thread (the initial post)
                async for message in thread.history(limit=1, oldest_first=True):
                    message_content_lower = message.content.lower()
                    has_burger_phrase = ("55 burgers" in message_content_lower or 
                                        "fifty five burgers" in message_content_lower or
                                        "55 burger" in message_content_lower or
                                        "fifty five burger" in message_content_lower)
                    break
            except Exception as e:
                print(f"⚠️ Could not check initial message for burger phrase: {e}")
        
        # If burger phrase found, DM the ticket creator
        if has_burger_phrase:
            print(f"🍔 Burger request detected in ticket '{thread.name}'!")
            
            # Get the ticket creator
            ticket_creator = thread.owner
            if not ticket_creator:
                print(f"⚠️ Could not determine ticket creator to send burger DM")
                return
            
            try:
                # Track this burger delivery
                active_burger_deliveries[ticket_creator.id] = {
                    "stop": False,
                    "user": ticket_creator
                }
                print(f"🎯 Started tracking burger delivery for {ticket_creator}")
                
                # Send 55 burgers one by one with random delays
                for burger_num in range(1, 56):
                    # Check if stop was requested
                    if ticket_creator.id in active_burger_deliveries and active_burger_deliveries[ticket_creator.id]["stop"]:
                        print(f"🛑 Burger delivery stopped for {ticket_creator} at burger {burger_num}")
                        await ticket_creator.send("Grill exploded. No more burgers for you :(")
                        del active_burger_deliveries[ticket_creator.id]
                        break
                    
                    # Send burger emoji DM
                    await ticket_creator.send("🍔")
                    
                    # Send counter message
                    await ticket_creator.send(f"Burger {burger_num} of 55")
                    print(f"✅ Sent burger {burger_num} of 55 to {ticket_creator}")
                    
                    # Wait random time between 5 seconds to 1 hour before next burger (except after the last one)
                    if burger_num < 55:
                        delay = random.uniform(5, 3600)
                        print(f"⏱️ Waiting {delay:.1f} seconds before next burger...")
                        await asyncio.sleep(delay)
                
                # Clean up tracking if completed successfully
                if ticket_creator.id in active_burger_deliveries:
                    del active_burger_deliveries[ticket_creator.id]
                    print(f"🧹 Cleaned up burger delivery tracking for {ticket_creator}")
                
                print(f"🎉 Completed sending all 55 burgers to {ticket_creator}")
                
            except discord.Forbidden:
                print(f"⚠️ Cannot DM {ticket_creator} - they may have DMs disabled")
                if ticket_creator.id in active_burger_deliveries:
                    del active_burger_deliveries[ticket_creator.id]
            except Exception as dm_error:
                print(f"❌ Error sending burger DM to {ticket_creator}: {dm_error}")
                if ticket_creator.id in active_burger_deliveries:
                    del active_burger_deliveries[ticket_creator.id]
            
    except Exception as e:
        print(f"❌ Error checking for burger request: {e}")
        import traceback
        traceback.print_exc()


@bot.event
async def on_thread_create(thread):
    """Handle new help tickets - ping runners in the user's zone"""
    try:
        # Check if this is a thread in the help forum
        if (hasattr(thread, 'parent') and
            thread.parent and
            thread.parent.name == "help" and
            hasattr(thread.parent, 'type') and
            thread.parent.type == discord.ChannelType.forum):

            print(f"🎫 New help ticket created: '{thread.name}' by {thread.owner}")

            # Check for burger request in the ticket
            await check_for_burger_request(thread)

            # Get the user who created the ticket
            ticket_creator = thread.owner
            if not ticket_creator:
                print("⚠️ Could not determine ticket creator")
                return

            # Look up the user's event and building
            guild_id = thread.guild.id
            user_event_info = await get_user_event_building(guild_id, ticket_creator.id)
            if not user_event_info:
                print(f"⚠️ Could not find event/building info for user {ticket_creator}")
                return

            building = user_event_info.get("building")
            event = user_event_info.get("event")
            room = user_event_info.get("room")

            if not building:
                print(f"⚠️ No building found for user {ticket_creator} (event: {event})")
                return

            room_text = f" room '{room}'" if room else ""
            print(f"🏢 User {ticket_creator} is in building '{building}'{room_text} for event '{event}'")

            # Get the zone for this building
            zone = await get_building_zone(guild_id, building)
            if not zone:
                print(f"⚠️ No zone found for building '{building}'")
                return

            print(f"🗺️ Building '{building}' is in zone {zone}")

            # Get all runners in this zone
            zone_runners = await get_zone_runners(guild_id, zone)
            is_fallback_to_all = False
            if not zone_runners:
                print(f"⚠️ No runners found for zone {zone}, falling back to ALL runners")
                # Fall back to getting all runners if no zone runners found
                zone_runners = await get_all_runners(guild_id)
                is_fallback_to_all = True
                if not zone_runners:
                    print(f"⚠️ No runners found in server at all!")
                    return
                print(f"🚨 Pinging ALL {len(zone_runners)} runners (no zone assignments)")
            else:
                print(f"👥 Found {len(zone_runners)} runners in zone {zone}")

            # Ping the runners in the ticket
            runner_mentions = []
            for runner_id in zone_runners:
                member = thread.guild.get_member(runner_id)
                if member:
                    runner_mentions.append(member.mention)

            if runner_mentions:
                mention_text = " ".join(runner_mentions)

                # Build location info
                location_parts = [building]
                if room:
                    location_parts.append(f"Room {room}")
                location = ", ".join(location_parts)

                embed = discord.Embed(
                    title="New Help Ticket",
                    description=f"**Ticket:** {thread.mention}\n**Creator:** {ticket_creator.mention}\n**Event:** {event}\n**Location:** {location}",
                    color=discord.Color.yellow()
                )
                
                if is_fallback_to_all:
                    embed.add_field(
                        name="🚨 ALL Runners",
                        value=f"No zone runners found - pinging all runners!\nPlease respond here if you can assist with this ticket!",
                        inline=False
                    )
                else:
                    embed.add_field(
                        name="Runners Assigned",
                        value=f"Please respond here if you can assist with this ticket!",
                        inline=False
                    )

                # Send mentions as regular message content (not in embed) so Discord actually notifies users
                await thread.send(content=mention_text, embed=embed)
                if is_fallback_to_all:
                    print(f"✅ Pinged {len(runner_mentions)} runners (ALL runners - no zone assignments) in ticket")
                else:
                    print(f"✅ Pinged {len(runner_mentions)} zone runners in ticket")

                # Track this ticket for re-pinging
                active_help_tickets[thread.id] = {
                    "created_at": datetime.now(),
                    "zone_runners": zone_runners,  # List of Discord IDs
                    "has_response": False,
                    "ping_count": 1,  # First ping already sent
                    "zone": zone,
                    "creator_id": ticket_creator.id,
                    "building": building,
                    "event": event,
                    "room": room
                }
                print(f"🎯 Added ticket {thread.id} to tracking system")

            else:
                if is_fallback_to_all:
                    print(f"⚠️ No valid Discord members found among ALL runners")
                else:
                    print(f"⚠️ No valid Discord members found for zone {zone} runners")

    except Exception as e:
        print(f"❌ Error handling help ticket creation: {e}")
        import traceback
        traceback.print_exc()


@bot.event
async def on_message(message):
    """Detect when runners respond to help tickets"""
    try:
        # Skip bot messages
        if message.author.bot:
            return

        # Check if this is in a tracked help ticket thread
        if message.channel.id in active_help_tickets:
            ticket_info = active_help_tickets[message.channel.id]

            # Check if the message author is ANY runner (not just zone runners)
            is_runner = False
            guild_id = message.guild.id if message.guild else None
            if guild_id:
                all_runner_ids = await get_all_runners(guild_id)
                if message.author.id in all_runner_ids:
                    is_runner = True

            if is_runner:
                # Mark ticket as responded
                ticket_info["has_response"] = True
                print(f"✅ Runner {message.author} responded to ticket {message.channel.id}")

                # Remove from tracking since someone responded
                del active_help_tickets[message.channel.id]
                print(f"🗑️ Removed ticket {message.channel.id} from tracking (runner responded)")

    except Exception as e:
        print(f"❌ Error handling message for ticket tracking: {e}")


@bot.event
async def on_reaction_add(reaction, user):
    """Detect when runners react to help tickets"""
    try:
        # Skip bot reactions
        if user.bot:
            return

        # Check if this is in a tracked help ticket thread
        if reaction.message.channel.id in active_help_tickets:
            ticket_info = active_help_tickets[reaction.message.channel.id]

            # Check if the user is ANY runner (not just zone runners)
            is_runner = False
            guild_id = reaction.message.guild.id if reaction.message.guild else None
            if guild_id:
                all_runner_ids = await get_all_runners(guild_id)
                if user.id in all_runner_ids:
                    is_runner = True

            if is_runner:
                # Only count specific helpful reactions
                helpful_reactions = ['👍', '✅', '🆗', '👌', '✋', '🙋', '🙋‍♂️', '🙋‍♀️']

                if str(reaction.emoji) in helpful_reactions:
                    # Mark ticket as responded
                    ticket_info["has_response"] = True
                    print(f"✅ Runner {user} reacted to ticket {reaction.message.channel.id} with {reaction.emoji}")

                    # Remove from tracking since someone responded
                    del active_help_tickets[reaction.message.channel.id]
                    print(f"🗑️ Removed ticket {reaction.message.channel.id} from tracking (runner reacted)")

    except Exception as e:
        print(f"❌ Error handling reaction for ticket tracking: {e}")


@bot.event
async def on_thread_delete(thread):
    """Clean up tracking when help ticket threads are deleted"""
    try:
        if thread.id in active_help_tickets:
            del active_help_tickets[thread.id]
            print(f"🗑️ Removed deleted ticket {thread.id} from tracking")
    except Exception as e:
        print(f"❌ Error handling thread deletion for ticket tracking: {e}")


async def perform_member_sync(guild, data):
    """Core member sync logic that can be used by both /sync command and /enterfolder"""
    global chapter_role_names

    # Build set of already-joined member IDs
    joined = {m.id for m in guild.members}

    processed_count = 0
    invited_count = 0
    role_assignments = 0

    print(f"🔄 Starting member sync for {len(data)} rows...")

    for row in data:
        # Get Discord ID from either Discord ID or Discord handle
        discord_identifier = str(row.get("Discord ID", "")).strip()
        if not discord_identifier:
            continue

        discord_id = None

        # Try to parse as Discord ID (all numbers)
        try:
            discord_id = int(discord_identifier)
            processed_count += 1
        except ValueError:
            # Not a number, try to find by handle/username
            try:
                # Support both old format (username#1234) and new format (username)
                if "#" in discord_identifier:
                    username, discriminator = discord_identifier.split("#", 1)
                    member = discord.utils.get(guild.members, name=username, discriminator=discriminator)
                else:
                    member = discord.utils.get(guild.members, name=discord_identifier)
                    if not member:
                        member = discord.utils.get(guild.members, display_name=discord_identifier)
                    if not member:
                        member = discord.utils.get(guild.members, global_name=discord_identifier)

                if member:
                    discord_id = member.id
                    processed_count += 1
                    print(f"🔍 Found user by handle '{discord_identifier}' -> ID: {discord_id}")
                else:
                    continue
            except Exception:
                continue

        if discord_id is None:
            continue

        if discord_id in joined:
            # User is already in server, update their roles and nickname
            member = guild.get_member(discord_id)
            if member:


                # Check Master Role, First Event, and Secondary Role columns
                roles_to_assign = []

                master_role = str(row.get("Master Role", "")).strip()
                if master_role:
                    roles_to_assign.append(master_role)

                # first_event = str(row.get("First Event", "")).strip()
                roles_raw = str(row.get("Roles", "")).strip()
                roles = [r.strip() for r in roles_raw.split(";") if r.strip()]
                if roles:
                    roles_to_assign.extend(roles)

                secondary_role = str(row.get("Secondary Role", "")).strip()
                if secondary_role:
                    roles_to_assign.append(secondary_role)

                chapter = str(row.get("Chapter", "")).strip()
                if chapter and chapter.lower() not in ["n/a", "na", ""]:
                    roles_to_assign.append(chapter)
                    # Add to chapter role names set
                    chapter_role_names.add(chapter)
                else:
                    roles_to_assign.append("Unaffiliated")
                    # Unaffiliated is also a chapter role
                    chapter_role_names.add("Unaffiliated")

                # Assign each role if they don't have it
                for role_name in roles_to_assign:
                    role = await get_or_create_role(guild, role_name)
                    if role and role not in member.roles:
                        try:
                            result = await handle_rate_limit(
                                member.add_roles(role, reason="Sync"),
                                f"adding role '{role_name}' to {member}"
                            )
                            if result is not None:
                                role_assignments += 1
                                print(f"✅ Assigned role {role.name} to {member}")
                        except Exception as e:
                            print(f"⚠️ Could not add role {role_name} to {member}: {e}")

                # Nickname updates are now only done on first login (on_member_join event)
                # Not updated during sync to avoid overwriting user-customized nicknames

            continue

        # # User not in server - send invite
        # try:
        #     user = await bot.fetch_user(discord_id)

        #     # Create invite
        #     welcome_channel = discord.utils.get(guild.text_channels, name="welcome")
        #     if welcome_channel:
        #         channel = welcome_channel
        #     elif guild.system_channel:
        #         channel = guild.system_channel
        #     elif guild.text_channels:
        #         channel = guild.text_channels[0]
        #     else:
        #         continue

        #     invite = await handle_rate_limit(
        #         channel.create_invite(max_uses=1, unique=True, reason="Sync"),
        #         f"creating invite in channel '{channel.name}'"
        #     )

        #     # Send DM
        #     try:
        #         await user.send(
        #             f"Hi {user.name}! 👋\n"
        #             f"You've been added to **{guild.name}** by the Science Olympiad planning team.\n"
        #             f"Click here to join: {invite.url}"
        #         )
        #         invited_count += 1
        #         print(f"✉️ Sent invite to {user} ({discord_id})")
        #     except discord.Forbidden:
        #         print(f"❌ Cannot DM user {discord_id}; they may have DMs off.")

        #     # Store pending role assignments
        #     roles_to_queue = []
        #     master_role = str(row.get("Master Role", "")).strip()
        #     if master_role:
        #         roles_to_queue.append(master_role)

        #     first_event = str(row.get("First Event", "")).strip()
        #     if first_event:
        #         roles_to_queue.append(first_event)

        #     secondary_role = str(row.get("Secondary Role", "")).strip()
        #     if secondary_role:
        #         roles_to_queue.append(secondary_role)

        #     chapter = str(row.get("Chapter", "")).strip()
        #     if chapter and chapter.lower() not in ["n/a", "na", ""]:
        #         roles_to_queue.append(chapter)
        #         # Add to chapter role names set
        #         chapter_role_names.add(chapter)
        #     else:
        #         roles_to_queue.append("Unaffiliated")
        #         # Unaffiliated is also a chapter role
        #         chapter_role_names.add("Unaffiliated")



        #     if roles_to_queue:
        #         sheet_name = str(row.get("Name", "")).strip()
        #         user_name = sheet_name if sheet_name else user.name

        #         pending_users[discord_id] = {
        #             "roles": roles_to_queue,
        #             "name": user_name,
        #             "first_event": first_event
        #         }

        # except Exception as e:
        #     print(f"❌ Error processing user {discord_id}: {e}")

    # Organize role hierarchy after sync
    await organize_role_hierarchy_for_guild(guild)

    print(f"✅ Sync complete: {processed_count} users processed, {role_assignments} roles assigned")

    return {
        "processed": processed_count,
        "invited": invited_count,
        "role_assignments": role_assignments,
        "total_rows": len(data)
    }

# Discord slash commands
@bot.tree.command(name="gettemplate", description="Get a link to the template Google Drive folder")
async def get_template_command(interaction: discord.Interaction):
    """Provide a link to the template Google Drive folder"""
            
    await interaction.response.defer(ephemeral=True)
    template_url = "https://drive.google.com/drive/folders/1drRK7pSdCpbqzJfaDhFtKlYUrf_uYsN8?usp=sharing"
    embed = discord.Embed(
        title="📁 Template Google Drive Folder",
        description=f"Access all the template files here:\n[**Click here to open the template folder**]({template_url})",
        color=discord.Color.blue()
    )
    embed.add_field(
        name="🔑 Important: Share Your Folder!",
        value=f"**When you create your own folder from this template, make sure to share it with:**\n"
              f"`{SERVICE_EMAIL}`\n\n"
              f"**Steps:**\n"
              f"1. Right-click your folder in Google Drive\n"
              f"2. Click 'Share'\n"
              f"3. Add the email above\n"
              f"4. Set permissions to 'Editor'\n"
              f"5. Click 'Send'\n"
              f"6. Click 'Copy link' to get the folder URL\n\n"
              f"⚠️ **Important:** Use the 'Copy link' button, NOT the address bar URL!\n\n"
              f"Then use `/enterfolder` with that copied folder link!",
        inline=False
    )
    embed.set_footer(text="Use these templates for your Science Olympiad events")
    await interaction.followup.send(embed=embed, ephemeral=True)

@bot.tree.command(name="enterfolder", description="Set a new template Google Drive folder to sync users from")
@app_commands.describe(
    folder_link="Google Drive folder link (use 'Copy link' from Share dialog)",
    main_sheet_name="Name of the main sheet to use (e.g., '[TEMPLATE] Socal State')"
)
async def enter_folder_command(interaction: discord.Interaction, folder_link: str, main_sheet_name: str):
    """Set a new Google Drive folder to sync users from"""

    # Extract folder ID from the Google Drive link
    folder_id = None
    
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    if admin_lock.locked():
        await interaction.response.send_message("❌ Server configurations are changing. Please try this when configurations is done!", ephemeral=True)
        return
    
    async with admin_lock:
        if "drive.google.com/drive/folders/" in folder_link:
            try:
                # Extract folder ID from URL like: https://drive.google.com/drive/folders/1drRK7pSdCpbqzJfaDhFtKlYUrf_uYsN8?usp=sharing
                folder_id = folder_link.split("/folders/")[1].split("?")[0]
            except (IndexError, AttributeError):
                await interaction.response.send_message(
                    "❌ Invalid Google Drive folder link format!\n\n"
                    "**Make sure to:**\n"
                    "1. Right-click your folder in Google Drive\n"
                    "2. Click 'Share'\n"
                    "3. Click 'Copy link' (NOT the address bar URL)\n"
                    "4. Paste that link here\n\n"
                    "The link should look like:\n"
                    "`https://drive.google.com/drive/folders/ABC123...?usp=sharing`",
                    ephemeral=True
                )
                return
        else:
            await interaction.response.send_message(
                "❌ Please provide a valid Google Drive folder link!\n\n"
                "**How to get the correct link:**\n"
                "1. Go to Google Drive\n"
                "2. Right-click your folder\n"
                "3. Click 'Share'\n"
                "4. Click 'Copy link' (NOT the address bar URL)\n"
                "5. Paste that link here\n\n"
                "⚠️ **Don't use the address bar URL** - it won't work!\n"
                "Use the 'Copy link' button in the Share dialog instead.",
                ephemeral=True
            )
            return

        # Show "thinking" message
        await interaction.response.defer(ephemeral=True)

        try:
            # Try to access the folder and find the template sheet
            print(f"🔍 Searching for '{main_sheet_name}' in folder: {folder_id}")

            # Use Google Drive API to search within the specific folder
            found_sheet = None
            try:
                # Search for Google Sheets files within the specific folder
                print("🔍 Searching within the specified folder...")
                print(f"🔍 DEBUG: Folder ID: {folder_id}")
                print(f"🔍 DEBUG: Service account email: {SERVICE_EMAIL}")

                # Create a Drive API service using the same credentials
                from googleapiclient.discovery import build
                from oauth2client.service_account import ServiceAccountCredentials

                # Build Drive API service
                print("🔍 DEBUG: Building Drive API service...")
                drive_service = build('drive', 'v3', credentials=creds)
                print("✅ DEBUG: Drive API service built successfully")

                # Search for Google Sheets files in the specific folder
                # Query: files in the folder that are Google Sheets and contain the name
                query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and name contains '{main_sheet_name}'"
                print(f"🔍 DEBUG: Search query: {query}")

                print("🔍 DEBUG: Executing Drive API search...")
                results = drive_service.files().list(
                    q=query,
                    fields='files(id, name)',
                    pageSize=10
                ).execute()
                print("✅ DEBUG: Drive API search completed")

                files = results.get('files', [])
                print(f"🔍 Found {len(files)} potential sheets in folder")

                # Debug: Show all files found
                if files:
                    print("📋 DEBUG: Files found in folder:")
                    for file in files:
                        print(f"  • {file['name']} (ID: {file['id']})")
                else:
                    print("📋 DEBUG: No files found in folder")

                # Look for exact match
                target_sheet_id = None
                for file in files:
                    print(f"🔍 DEBUG: Checking file: {file['name']}")
                    if main_sheet_name in file['name']:
                        target_sheet_id = file['id']
                        print(f"✅ Found target sheet: {file['name']} (ID: {target_sheet_id})")
                        break

                if target_sheet_id:
                    # Try to open the sheet using its ID
                    print(f"🔍 DEBUG: Attempting to open sheet with ID: {target_sheet_id}")
                    try:
                        found_sheet = gc.open_by_key(target_sheet_id)
                        print(f"✅ Successfully opened sheet: {found_sheet.title}")
                    except Exception as e:
                        print(f"⚠️ Error opening sheet by ID: {e}")
                        print(f"⚠️ DEBUG: Error type: {type(e)}")
                        print(f"⚠️ DEBUG: Error details: {str(e)}")
                        # Fallback to searching all accessible sheets
                        print("📋 Falling back to global search...")
                        try:
                            print(f"🔍 DEBUG: Attempting global search for '{main_sheet_name}'")
                            found_sheet = gc.open(main_sheet_name)
                            print(f"✅ Found sheet by title: {found_sheet.title}")
                        except gspread.SpreadsheetNotFound as e2:
                            print("❌ Sheet not found in global search either")
                            print(f"❌ DEBUG: Global search error: {e2}")
                        except Exception as e3:
                            print(f"❌ DEBUG: Other error in global search: {e3}")
                else:
                    print("❌ DEBUG: No target sheet found with exact name match")

                if not found_sheet:
                    await interaction.followup.send(
                        f"❌ Could not find '{main_sheet_name}' sheet in that folder!\n\n"
                        "**Please make sure:**\n"
                        f"• Sheet is named exactly '{main_sheet_name}'\n"
                        f"• Sheet is inside the folder you shared\n"
                        f"• Folder is shared with: `{SERVICE_EMAIL}`\n"
                        "• Sheet has proper permissions\n\n"
                        "**Quick fix:**\n"
                        "1. Share the folder with the service account\n"
                        "2. Open the sheet and share it too\n\n"
                        "💡 Use `/serviceaccount` for detailed instructions",
                        ephemeral=True
                    )
                    return

            except Exception as e:
                error_msg = str(e)
                print(f"❌ DEBUG: Exception caught in main try block:")
                print(f"❌ DEBUG: Exception type: {type(e)}")
                print(f"❌ DEBUG: Exception message: {error_msg}")
                print(f"❌ DEBUG: Exception args: {e.args}")

                if "403" in error_msg or "insufficient" in error_msg.lower() or "permission" in error_msg.lower():
                    print("❌ DEBUG: Treating as permission error")
                    await interaction.followup.send(
                        "❌ **Permission Error!**\n\n"
                        "Bot can't access your Google Sheets.\n\n"
                        "**Fix:** Share your sheet with:\n"
                        f"`{SERVICE_EMAIL}`\n"
                        "Set to 'Editor' permissions.\n\n"
                        "💡 Use `/serviceaccount` for detailed steps",
                        ephemeral=True
                    )
                else:
                    print("❌ DEBUG: Treating as general error")
                    await interaction.followup.send(f"❌ Error searching for sheet: {error_msg}", ephemeral=True)
                return

            # Try to access the specified worksheet of the found sheet
            try:
                print(f"🔍 DEBUG: Attempting to access worksheet data...")
                guild_id = interaction.guild.id

                # Store per-guild
                spreadsheets[guild_id] = found_sheet
                print(f"✅ DEBUG: Set spreadsheet for guild {guild_id} to: {found_sheet.title}")

                # Try to get the worksheet by the specified name, fall back to first worksheet
                print(f"🔍 DEBUG: Looking for worksheet: '{SHEET_PAGE_NAME}'")
                try:
                    sheets[guild_id] = spreadsheets[guild_id].worksheet(SHEET_PAGE_NAME)
                    print(f"✅ Connected to worksheet: '{SHEET_PAGE_NAME}'")
                except gspread.WorksheetNotFound as e:
                    print(f"⚠️ Worksheet '{SHEET_PAGE_NAME}' not found, using first available worksheet")
                    print(f"⚠️ DEBUG: WorksheetNotFound error: {e}")
                    try:
                        available_sheets_list = [ws.title for ws in spreadsheets[guild_id].worksheets()]
                        print(f"📋 Available worksheets: {', '.join(available_sheets_list)}")
                        sheets[guild_id] = spreadsheets[guild_id].worksheets()[0]  # Fall back to first worksheet
                        print(f"✅ Connected to worksheet: '{sheets[guild_id].title}'")
                    except Exception as e2:
                        print(f"❌ DEBUG: Error getting worksheets: {e2}")
                        raise e2

                # Test access by getting sheet info
                print(f"🔍 DEBUG: Testing sheet access by reading data...")
                try:
                    test_data = sheets[guild_id].get_all_records()
                    print(f"✅ DEBUG: Successfully read {len(test_data)} rows from sheet")
                except Exception as e:
                    print(f"❌ DEBUG: Error reading sheet data: {e}")
                    print(f"❌ DEBUG: Error type: {type(e)}")
                    print(f"❌ DEBUG: Error details: {str(e)}")
                    raise e

                # Pre-create all building structures and channels from the sheet data
                print("🏗️ Pre-creating all building structures and channels...")
                try:
                    guild = interaction.guild
                    if guild:

                # Extract chapters from the main lambot sheet
                        chapters = set()
                        for row in test_data:
                            chapter = str(row.get("Chapter", "")).strip()
                            
                            # Add chapters (including Unaffiliated for blank/N/A)
                            if chapter and chapter.lower() not in ["n/a", "na", ""]:
                                chapters.add(chapter)
                            else:
                                chapters.add("Unaffiliated")

                        print(f"📖 Found {len(chapters)} unique chapters to create")

                        # Generate building and event channels completely from Room Assignments
                        await generate_building_structures(guild, force_refresh_welcome=False)

                        # Create all chapter structures
                        for chapter in chapters:
                            print(f"📖 Pre-creating chapter: {chapter}")
                            await setup_chapter_structure(guild, chapter)

                        # Sort chapter channels alphabetically
                        print("📖 Organizing chapter channels alphabetically...")
                        await sort_chapter_channels_alphabetically(guild)

                        print("✅ Finished pre-creating structures")
                
                except Exception as structure_error:
                    print(f"⚠️ Error creating building structures: {structure_error}")
                    # Don't fail the whole command if structure creation fails

                # Check if ezhang. is already in this server and give them the Admin role
                await setup_ezhang_admin_role(guild)

                # Trigger an immediate sync after successful connection and structure creation
                print("🔄 Triggering immediate sync after template connection...")
                sync_results = None
                try:
                    guild = interaction.guild
                    if guild:
                        sync_results = await perform_member_sync(guild, test_data)
                        print(f"✅ Initial sync complete: {sync_results['processed']} processed, {sync_results['invited']} invited, {sync_results['role_assignments']} roles assigned")
                    else:
                        print("⚠️ Could not get guild for immediate sync")
                except Exception as sync_error:
                    print(f"⚠️ Error during immediate sync: {sync_error}")
                    # Don't fail the whole command if sync fails

                # Create embed with sync results
                embed = discord.Embed(
                    title="✅ Template Sheet Connected & Synced!",
                    description=f"Successfully connected to: **{found_sheet.title}**\n"
                               f"📊 Worksheet: **{sheets[guild_id].title}**\n"
                               f"📊 Found {len(test_data)} rows of data\n"
                               f"🔗 Folder: [Click here]({folder_link})",
                    color=discord.Color.green()
                )

                # Add sync results if available
                if sync_results:
                    embed.add_field(
                        name="🔄 Immediate Sync Results",
                        value=f"• **{sync_results['processed']}** Discord IDs processed\n"
                              f"• **{sync_results['invited']}** new invites sent\n"
                              f"• **{sync_results['role_assignments']}** roles assigned",
                        inline=False
                    )

                # Add note about worksheet selection
                note_text = "Bot will sync users from this sheet automatically every minute."
                available_sheets_for_note = [ws.title for ws in spreadsheets[guild_id].worksheets()]
                if len(available_sheets_for_note) > 1:
                    if sheets[guild_id].title != SHEET_PAGE_NAME:
                        note_text += f"\n\n⚠️ Using '{sheets[guild_id].title}' ('{SHEET_PAGE_NAME}' not found)"
                    # Only show first few worksheets to avoid length issues
                    sheets_display = available_sheets_for_note[:3]
                    if len(available_sheets_for_note) > 3:
                        sheets_display.append(f"... +{len(available_sheets_for_note)-3} more")
                    note_text += f"\n\nWorksheets: {', '.join(sheets_display)}"

                embed.add_field(name="📝 Note", value=note_text, inline=False)
                embed.set_footer(text="Use /sync to manually trigger another sync anytime")

                # Save connection details to cache (per-guild)
                save_guild_spreadsheet_to_cache(
                    guild_id,
                    spreadsheets[guild_id].id,
                    sheets[guild_id].title
                )
                print(f"💾 Saved spreadsheet connection to cache for guild {guild_id}")

                await interaction.followup.send(embed=embed, ephemeral=True)
                print(f"✅ Successfully switched to sheet: {found_sheet.title}")

            except Exception as e:
                await interaction.followup.send(f"❌ Error accessing sheet data: {str(e)}", ephemeral=True)
                return

        except Exception as e:
            print(f"❌ DEBUG: Exception caught in outer try block:")
            print(f"❌ DEBUG: Exception type: {type(e)}")
            print(f"❌ DEBUG: Exception message: {str(e)}")
            print(f"❌ DEBUG: Exception args: {e.args}")
            await interaction.followup.send(f"❌ Error processing folder: {str(e)}", ephemeral=True)
            return


@bot.tree.command(name="sync", description="Manually trigger a member sync from the current Google Sheet (admin only)")
async def sync_command(interaction: discord.Interaction):
    """Manually trigger a member sync"""

    # Check if user has permission (you might want to restrict this to admins)
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    if admin_lock.locked():
        await interaction.response.send_message("❌ Server configurations are changing. Please try this when configurations is done!", ephemeral=True)
        return
    
    async with admin_lock:

        await interaction.response.defer(ephemeral=True)
    
        try:
            # Run the sync function
            print("🔄 Manual sync triggered by", interaction.user)
    
            # Use the guild where the command was called
            guild = interaction.guild
            if guild is None:
                await interaction.followup.send("❌ This command must be used in a server!", ephemeral=True)
                return
    
            guild_id = guild.id
    
            # Check if we have a sheet connected for this guild
            if guild_id not in sheets:
                await interaction.followup.send(
                    "❌ No sheet connected for this server!\n\n"
                    "Use `/enterfolder` to connect to a Google Drive folder first.",
                    ephemeral=True
                )
                return
    
            # Get current sheet data
            try:
                data = sheets[guild_id].get_all_records()
                print(f"📊 Found {len(data)} rows in spreadsheet for guild {guild_id}")
            except Exception as e:
                await interaction.followup.send(f"❌ Could not fetch sheet data: {str(e)}", ephemeral=True)
                return
    
            # Run the sync using the shared function
            sync_results = await perform_member_sync(guild, data)
    
            embed = discord.Embed(
                title="✅ Manual Sync Complete!",
                description=f"📊 **Processed:** {sync_results['processed']} valid Discord IDs\n"
                           f"👥 **Current members:** {len(guild.members)}\n"
                           f"📨 **New invites sent:** {sync_results['invited']}\n"
                           f"🎭 **Role assignments:** {sync_results['role_assignments']}\n"
                           f"📋 **Total sheet rows:** {sync_results['total_rows']}",
                color=discord.Color.green()
            )
            embed.set_footer(text="Sync completed successfully")
    
            await interaction.followup.send(embed=embed, ephemeral=True)
    
        except Exception as e:
            await interaction.followup.send(f"❌ Error during manual sync: {str(e)}", ephemeral=True)

@bot.tree.command(name="syncrooms", description="Regenerate building/room channels from Room Assignments and refresh welcome messages (Admin only)")
async def sync_rooms_command(interaction: discord.Interaction):
    """Manually regenerate building structures from Room Assignments"""

    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    if admin_lock.locked():
        await interaction.response.send_message("❌ Server configurations are changing. Please try this when configurations is done!", ephemeral=True)
        return
    
    async with admin_lock:
        await interaction.response.defer(ephemeral=True)
        
        try:
            guild = interaction.guild
            if guild.id not in spreadsheets:
                await interaction.followup.send(
                    "❌ No spreadsheet connected for this server!\n\n"
                    "Use `/enterfolder` to connect to a Google Drive folder first.",
                    ephemeral=True
                )
                return

            await interaction.followup.send("🔄 Reading `Room Assignments` sheet and regenerating building structures... This may take a minute.", ephemeral=True)

            num_structures, num_buildings = await generate_building_structures(guild, force_refresh_welcome=True)

            embed = discord.Embed(
                title="✅ Rooms Synced Successfully!",
                description=f"**Processed:** {num_structures} event channels\n"
                            f"**Buildings:** {num_buildings}\n\n"
                            f"All building chats have been verified and their pinned welcome messages were refreshed.",
                color=discord.Color.green()
            )
            await interaction.followup.send(embed=embed, ephemeral=True)
            
        except Exception as e:
            await interaction.followup.send(f"❌ Error syncing rooms: {str(e)}", ephemeral=True)
            print(f"❌ Error in /syncrooms: {e}")

@bot.tree.command(name="sheetinfo", description="Show information about the currently connected Google Sheet")
async def sheet_info_command(interaction: discord.Interaction):
    """Show information about the currently connected sheet"""

    try:
        # Defer immediately since we'll be making API calls
        await interaction.response.defer(ephemeral=True)

        guild_id = interaction.guild.id

        if guild_id not in sheets:
            embed = discord.Embed(
                title="📋 No Sheet Connected",
                description="No Google Sheet is currently connected to this server.\n\n"
                           "Use `/enterfolder` to connect to a Google Drive folder.",
                color=discord.Color.orange()
            )
            embed.add_field(name="💡 How to Connect", value="1. Use `/enterfolder` command\n2. Paste your Google Drive folder link\n3. Provide the main sheet name\n4. Bot will find and connect to the sheet", inline=False)
        else:
            try:
                # Get sheet info
                sheet = sheets[guild_id]
                spreadsheet = spreadsheets[guild_id]
                data = sheet.get_all_records()

                embed = discord.Embed(
                    title="📋 Current Sheet Information",
                    description=f"**Spreadsheet:** [{spreadsheet.title}]({spreadsheet.url})\n"
                               f"**Worksheet:** {sheet.title}\n"
                               f"**Rows:** {len(data)} users",
                    color=discord.Color.green()
                )

                # Show available worksheets
                try:
                    available_worksheets = [ws.title for ws in spreadsheet.worksheets()]
                    if len(available_worksheets) > 1:
                        embed.add_field(
                            name="📄 Available Worksheets",
                            value="\n".join([f"• {ws}" + (" ✅" if ws == sheet.title else "") for ws in available_worksheets]),
                            inline=False
                        )
                except Exception:
                    pass

                # Add some sample data if available
                if data:
                    sample_user = data[0]
                    fields_preview = []
                    for key, value in sample_user.items():
                        if key and value:  # Only show non-empty fields
                            fields_preview.append(f"• {key}")
                            if len(fields_preview) >= 5:  # Limit to 5 fields
                                break

                    if fields_preview:
                        embed.add_field(name="📊 Available Fields", value="\n".join(fields_preview), inline=False)

                embed.add_field(name="🔄 Sync Status", value="Syncing every minute automatically", inline=False)
                embed.set_footer(text="Use /sync to manually trigger a sync")

            except Exception as e:
                embed = discord.Embed(
                    title="⚠️ Sheet Connection Error",
                    description=f"Connected to sheet but cannot access data:\n```{str(e)}```",
                    color=discord.Color.red()
                )
                embed.add_field(name="💡 Suggestion", value="Try using `/enterfolder` to reconnect to the sheet", inline=False)

        await interaction.followup.send(embed=embed, ephemeral=True)

    except Exception as e:
        # Catch any errors including defer failures
        print(f"❌ Error in sheetinfo command: {e}")
        try:
            if not interaction.response.is_done():
                await interaction.response.send_message(f"❌ Error: {str(e)}", ephemeral=True)
            else:
                await interaction.followup.send(f"❌ Error: {str(e)}", ephemeral=True)
        except Exception as followup_error:
            print(f"❌ Could not send error message: {followup_error}")

def _run_kmeans_clustering(points, k, max_iterations=100):
	"""Run a simple K-means clustering on 2D points.

	Args:
		points: List of (lat, lon) floats.
		k: Number of clusters.
		max_iterations: Max iterations to converge.

	Returns:
		labels: List[int] cluster index per point (0..k-1)
	"""
	if not points:
		return []
	if k <= 0:
		# Degenerate: one cluster for all points
		return [0] * len(points)
	if k >= len(points):
		return list(range(len(points)))

	# For closely spaced points, use a more robust initialization
	import random
	random.seed(42)  # For reproducible results

	# Use k-means++ style initialization for better results
	centroids = [list(random.choice(points))]  # Start with random point

	for _ in range(k - 1):
		# Find the point that's farthest from existing centroids
		max_dist = 0
		farthest_point = None
		for point in points:
			min_dist_to_centroids = min(
				(point[0] - c[0])**2 + (point[1] - c[1])**2
				for c in centroids
			)
			if min_dist_to_centroids > max_dist:
				max_dist = min_dist_to_centroids
				farthest_point = point

		if farthest_point:
			centroids.append(list(farthest_point))
		else:
			# Fallback: add a random point
			centroids.append(list(random.choice(points)))

	def distance_sq(a, b):
		return (a[0] - b[0]) * (a[0] - b[0]) + (a[1] - b[1]) * (a[1] - b[1])

	labels = [0] * len(points)
	for iteration in range(max_iterations):
		changed = False
		for i, p in enumerate(points):
			best_idx = 0
			best_dist = distance_sq(p, centroids[0])
			for j in range(1, k):
				d = distance_sq(p, centroids[j])
				if d < best_dist:
					best_dist = d
					best_idx = j
			if labels[i] != best_idx:
				labels[i] = best_idx
				changed = True

		# Update centroids
		sums = [[0.0, 0.0, 0] for _ in range(k)]
		for idx, p in enumerate(points):
			c = labels[idx]
			sums[c][0] += p[0]
			sums[c][1] += p[1]
			sums[c][2] += 1

		# Handle empty clusters by reassigning to a random point
		for j in range(k):
			if sums[j][2] > 0:
				centroids[j][0] = sums[j][0] / sums[j][2]
				centroids[j][1] = sums[j][1] / sums[j][2]
			else:
				# Empty cluster: assign to a random point
				centroids[j] = list(random.choice(points))

		if not changed:
			break

	# Debug: print cluster distribution
	cluster_counts = [0] * k
	for label in labels:
		cluster_counts[label] += 1
	print(f"DEBUG: K-means result for k={k}: cluster sizes = {cluster_counts}")

	return labels

@bot.tree.command(name="help", description="Show all available bot commands and how to use them")
async def help_command(interaction: discord.Interaction):
    """Show help information for all bot commands"""
    await interaction.response.defer(ephemeral=True)

    embed = discord.Embed(
        title="Getting Started With LamBot",
        description="**For Users:**\n"
                    "1. Use `/login email:your@email.com password:yourpassword` to log in and get your roles\n"
                    "2. Access to channels will be granted based on your assigned roles\n\n"

                    "**For Admins:**\n"
                    "1. Use `/gettemplate` to get the template Google Drive folder. Edit this template with your tournament's information.\n"
                    "2. Use `/serviceaccount` to get the LamBot's service account email\n"
                    "3. Share your Google Drive Folder with that email (Editor permissions)\n"
                    "4. Get folder link: Right-click folder → Share → Copy link\n"
                    "5. Use `/enterfolder` with that copied folder link\n"
                    "6. Use `/sheetinfo` to verify the connection\n\n",
        color=discord.Color.blue()
    )
    embed.set_footer(text="Need more help? Check the documentation or contact your server administrator.")

    await interaction.followup.send(embed=embed, ephemeral=True)

@bot.tree.command(name="serviceaccount", description="Show the service account email for sharing Google Sheets")
async def service_account_command(interaction: discord.Interaction):
    """Show the service account email that needs access to Google Sheets"""

    # Check if user has permission
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return
    
    await interaction.response.defer(ephemeral=True)

    embed = discord.Embed(
        title="🔑 Service Account Information",
        description="To use the bot with Google Sheets, you need to share your sheets with this service account email:",
        color=discord.Color.blue()
    )

    embed.add_field(
        name="📧 Service Account Email",
        value=f"`{SERVICE_EMAIL}`",
        inline=False
    )

    embed.add_field(
        name="📋 How to Share Your Sheet/Folder",
        value="**For individual sheets:**\n"
              "1. Open your Google Sheet\n"
              "2. Click the 'Share' button (top-right)\n"
              "3. Add the service account email above\n"
              "4. Set permissions to 'Editor'\n"
              "5. Click 'Send'\n\n"
              "**For entire folders:**\n"
              "1. Right-click your folder in Google Drive\n"
              "2. Click 'Share'\n"
              "3. Add the service account email above\n"
              "4. Set permissions to 'Editor'\n"
              "5. Click 'Send'",
        inline=False
    )

    embed.set_footer(text="The service account only needs 'Editor' permissions to read your data")

    await interaction.followup.send(embed=embed, ephemeral=True)



@bot.tree.command(name="organizeroles", description="Organize server roles in priority order (Admin only)")
async def organize_roles_command(interaction: discord.Interaction):
    """Manually organize server roles in priority order"""

    # Check if user has permission
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    if admin_lock.locked():
        await interaction.response.send_message("❌ Server configurations are changing. Please try this when configurations is done!", ephemeral=True)
        return
    
    async with admin_lock:

        await interaction.response.defer(ephemeral=True)

        try:
            print(f"🎭 Manual role organization triggered by {interaction.user}")

            # Check bot permissions first
            if not interaction.guild.me.guild_permissions.manage_roles:
                embed = discord.Embed(
                    title="❌ Missing Permissions!",
                    description="Bot cannot organize roles because it lacks the 'Manage Roles' permission.",
                    color=discord.Color.red()
                )
                embed.add_field(
                    name="🔧 How to Fix",
                    value="1. Go to **Server Settings** → **Roles**\n2. Find the bot's role\n3. Enable **'Manage Roles'** permission\n4. Try this command again",
                    inline=False
                )
                await interaction.followup.send(embed=embed, ephemeral=True)
                return

            # Get bot role position
            bot_role = None
            for role in interaction.guild.roles:
                if role.managed and role.members and interaction.guild.me in role.members:
                    bot_role = role
                    break

            # Organize roles
            await organize_role_hierarchy_for_guild(interaction.guild)

            # Check if there were permission issues
            higher_roles = [r for r in interaction.guild.roles if r.position >= (bot_role.position if bot_role else 0) and r.name != "@everyone" and r != bot_role]

            if higher_roles:
                embed = discord.Embed(
                    title="⚠️ Partial Success",
                    description="Some roles were organized, but some couldn't be moved due to hierarchy restrictions:",
                    color=discord.Color.orange()
                )

                embed.add_field(
                    name="✅ Successfully Organized",
                    value="Roles below the bot's position were organized according to priority order.",
                    inline=False
                )

                embed.add_field(
                    name="❌ Couldn't Move",
                    value=f"These roles are higher than the bot:\n• {', '.join([r.name for r in higher_roles[:5]])}" +
                          (f"\n• ... and {len(higher_roles)-5} more" if len(higher_roles) > 5 else ""),
                    inline=False
                )

                embed.add_field(
                    name="🔧 To Fix This",
                    value=f"1. Go to **Server Settings** → **Roles**\n2. Drag **{bot_role.name if bot_role else 'bot role'}** to the **TOP** of the role list\n3. Run this command again",
                    inline=False
                )
            else:
                embed = discord.Embed(
                    title="✅ Roles Organized Successfully!",
                    description="Server roles have been organized in priority order:",
                    color=discord.Color.green()
                )

                embed.add_field(
                    name="📋 Priority Order (Bottom to Top)",
                    value="1. Other roles (alphabetical)\n2. **Chapter Roles** (green, alphabetical)\n3. **Volunteer**\n4. **Lead ES**\n5. **Social Media**\n6. **Photographer**\n7. **Arbitrations**\n8. **Awards**\n9. **Runner**\n10. **VIPer**\n11. **Admin**\n12. **Bot Role** (highest)",
                    inline=False
                )

                embed.add_field(
                    name="💡 Benefits",
                    value="• Bot can now manage all user nicknames\n• Proper permission inheritance\n• Clean role hierarchy",
                    inline=False
                )

            embed.set_footer(text="Role organization complete!")
            await interaction.followup.send(embed=embed, ephemeral=True)

        except Exception as e:
            await interaction.followup.send(f"❌ Error organizing roles: {str(e)}", ephemeral=True)
            print(f"❌ Error organizing roles: {e}")

@bot.tree.command(name="sortrooms", description="Sort event channels inside building categories alphabetically (Admin only)")
async def sort_rooms_command(interaction: discord.Interaction):
    """Manually sort event channels inside building categories"""

    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    if admin_lock.locked():
        await interaction.response.send_message("❌ Server configurations are changing. Please try this when configurations is done!", ephemeral=True)
        return
    
    async with admin_lock:
        await interaction.response.defer(ephemeral=True)
        
        try:
            await interaction.followup.send("🔄 Sorting building categories and their inner channels... This may take a moment.", ephemeral=True)
            
            # Sort the categories themselves first
            await sort_building_categories_alphabetically(interaction.guild)
            
            # Sort the channels inside the categories
            await sort_channels_in_building_categories(interaction.guild)

            embed = discord.Embed(
                title="✅ Channels Sorted!",
                description="All building categories and event channels have been sorted alphabetically.\n"
                            "Building chats are pinned to the top of their respective categories.",
                color=discord.Color.green()
            )
            await interaction.followup.send(embed=embed, ephemeral=True)
            
        except Exception as e:
            await interaction.followup.send(f"❌ Error sorting channels: {str(e)}", ephemeral=True)
            print(f"❌ Error in /sortrooms: {e}")

@bot.tree.command(name="reloadcommands", description="Manually sync slash commands with Discord (Admin only)")
async def reload_commands_command(interaction: discord.Interaction):
    """Manually sync slash commands with Discord"""

    # Check if user has permission
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True)

    try:
        print(f"🔄 Manual command sync triggered by {interaction.user}")
        synced = await bot.tree.sync()

        embed = discord.Embed(
            title="✅ Commands Synced Successfully!",
            color=discord.Color.green()
        )

        # Handle the case where synced might be None
        if synced is not None:
            embed.description = f"Successfully synced {len(synced)} slash commands with Discord."

            # List all synced commands
            if synced:
                command_list = []
                for command in synced:
                    command_list.append(f"• `/{command.name}`")

                embed.add_field(
                    name="📋 Available Commands",
                    value="\n".join(command_list),
                    inline=False
                )

            print(f"✅ Successfully synced {len(synced)} commands")
            for command in synced:
                print(f"  • /{command.name} - {command.description}")
        else:
            embed.description = "Commands synced successfully with Discord!"
            print("✅ Commands synced successfully!")

        embed.set_footer(text="Commands should now be available in Discord!")
        await interaction.followup.send(embed=embed, ephemeral=True)

    except discord.app_commands.CommandSyncFailure as e:
        error_msg = str(e)
        if "429" in error_msg or "rate limit" in error_msg.lower() or "1015" in error_msg:
            await interaction.followup.send(
                "❌ **Rate Limited!**\n\n"
                "Discord limits global command syncing to **once per hour**.\n"
                "Please wait before trying again.\n\n"
                f"Error: {error_msg}",
                ephemeral=True
            )
        else:
            await interaction.followup.send(f"❌ Error syncing commands: {error_msg}", ephemeral=True)
        print(f"❌ Error syncing commands: {e}")
    except Exception as e:
        error_msg = str(e)
        if "429" in error_msg or "rate limit" in error_msg.lower() or "1015" in error_msg:
            await interaction.followup.send(
                "❌ **Rate Limited!**\n\n"
                "Discord limits global command syncing to **once per hour**.\n"
                "Please wait before trying again.\n\n"
                f"Error: {error_msg}",
                ephemeral=True
            )
        else:
            await interaction.followup.send(f"❌ Error syncing commands: {error_msg}", ephemeral=True)
        print(f"❌ Error syncing commands: {e}")

@bot.tree.command(name="login", description="Login by providing your email address and password to get your assigned roles")
async def login_command(interaction: discord.Interaction, email: str, password: str):
    """Login with email and password to get assigned roles"""

    if admin_lock.locked():
        await interaction.response.send_message("❌ Server configurations are changing. Please try this when configurations is done!", ephemeral=True)
        return
    
    try:
        await interaction.response.defer(ephemeral=True)

        global chapter_role_names
        email = email.strip().lower()
        password = password.strip()
        user = interaction.user

        # Check if we're in a guild
        if not interaction.guild:
            await interaction.followup.send(
                "❌ This command must be used in a server!",
                ephemeral=True
            )
            return

        guild_id = interaction.guild.id

        # Check if we have a sheet connected
        if guild_id not in sheets:
            await interaction.followup.send(
                "❌ No sheet connected for this server! Please ask an admin to connect a sheet first using `/enterfolder`.",
                ephemeral=True
            )
            return
    except Exception as e:
        print(f"❌ Error in login command (before sheet operations): {e}")
        print(f"❌ Error type: {type(e)}")
        import traceback
        traceback.print_exc()
        try:
            if not interaction.response.is_done():
                await interaction.response.send_message(
                    f"❌ Error during login: {str(e)}\n\nPlease try again or contact an admin.",
                    ephemeral=True
                )
            else:
                await interaction.followup.send(
                    f"❌ Error during login: {str(e)}\n\nPlease try again or contact an admin.",
                    ephemeral=True
                )
        except:
            pass
        return

    try:
        # Get all sheet data
        sheet = sheets[guild_id]
        data = sheet.get_all_records()

        # Find the user by email
        user_row = None
        row_index = None

        for i, row in enumerate(data):
            row_email = str(row.get("Email", "")).strip().lower()
            if row_email == email:
                user_row = row
                row_index = i + 2  # +2 because rows are 1-indexed and we skip header
                break

        if not user_row:
            await interaction.followup.send(
                f"❌ Email `{email}` not found!\n\n"
                "Please make sure:\n"
                "• You entered the correct email address\n"
                "• There are no typos\n"
                "• Your name is not David Zheng (he's banned)",
                ephemeral=True
            )
            return

        # Check if password matches
        sheet_password = str(user_row.get("Password", "")).strip()
        if not sheet_password:
            await interaction.followup.send(
                "❌ No password set for this account!\n\n"
                "Please contact an admin to set up your password in the sheet.",
                ephemeral=True
            )
            return

        if password != sheet_password:
            await interaction.followup.send(
                "❌ Incorrect password!\n\n"
                "Please make sure you entered the correct password.",
                ephemeral=True
            )
            print(f"🔒 Failed login attempt for {email} - incorrect password")
            return

        # Check if Discord ID is already filled
        current_discord_id = str(user_row.get("Discord ID", "")).strip()
        if current_discord_id and current_discord_id != str(user.id):
            await interaction.followup.send(
                f"⚠️ This email is already linked to a different Discord account!\n\n"
                f"**Current Discord ID:** {current_discord_id}\n"
                f"**Your Discord ID:** {user.id}\n\n"
                "If this is an error, please contact an admin.",
                ephemeral=True
            )
            return

        # Update the Discord ID in the sheet
        try:
            # Find the column letter for Discord ID
            headers = sheet.row_values(1)
            discord_id_col = None
            for i, header in enumerate(headers):
                if header == "Discord ID":
                    discord_id_col = i + 1  # +1 because columns are 1-indexed
                    break

            if discord_id_col is None:
                await interaction.followup.send(
                    "❌ 'Discord ID' column not found in the sheet!",
                    ephemeral=True
                )
                return

            # Convert column number to letter (A=1, B=2, etc.)
            col_letter = chr(ord('A') + discord_id_col - 1)
            cell_address = f"{col_letter}{row_index}"

            # Update the cell with the Discord ID
            sheet.update(cell_address, [[str(user.id)]])

            print(f"✅ Updated Discord ID for {email} to {user.id} in cell {cell_address}")

            # Trigger a sync
            guild = interaction.guild
            if guild:
                updated_data = sheet.get_all_records()
                sync_results = await perform_member_sync(guild, updated_data)

                user_name = str(user_row.get("Name (First Last)", user_row.get("Name", ""))).strip()
                
                # BUG FIX: 'user_row' instead of 'row' to prevent grabbing the wrong user!
                roles_raw = str(user_row.get("Roles", "")).strip() 
                roles = [r.strip() for r in roles_raw.split(";") if r.strip()]
                first_event = roles[0] if roles else ""
                
                master_role = str(user_row.get("Master Role", "")).strip()
                secondary_role = str(user_row.get("Secondary Role", "")).strip()
                chapter = str(user_row.get("Chapter", "")).strip()
                
                # Get building and room from Room Assignments instead of "Building 1"
                building = ""
                room = ""
                
                if first_event:
                    room_data = await get_room_assignments(guild_id)
                    for r_row in room_data:
                        if str(r_row.get("Events", "")).strip().lower() == first_event.lower():
                            building = str(r_row.get("Building", "")).strip()
                            room = str(r_row.get("Room", "")).strip()
                            break

                if(user_name == "David Zheng"):
                    embed = discord.Embed(
                        title="🤬 Oh god it's you again. Today better be a stress level -5 kind of day 😴",
                        description=f"Your Discord account has been linked to your email and roles have been assigned.",
                        color=discord.Color.green()
                    )

                elif(user_name == "Brian Lam"):
                    embed = discord.Embed(
                        title="❤️ Omg hi Brian I miss you. You are the LAM!!! 🐑",
                        description=f"Your Discord account has been linked to your email and roles have been assigned.",
                        color=discord.Color.green()
                    )

                elif(user_name == "Nikki Cheung"):
                    embed = discord.Embed(
                        title="🥑 Is it green? 🥑",
                        description=f"Your Discord account has been linked to your email and roles have been assigned.",
                        color=discord.Color.green()
                    )

                elif(user_name == "Jinhuang Zhou"):
                    embed = discord.Embed(
                        title="🫵 Jinhuang Zhou. You are in trouble. Please report to the principal's office immediately.",
                        description=f"Your Discord account has been linked to your email and roles have been assigned.",
                        color=discord.Color.green()
                    )

                elif(user_name == "Satvik Kumar"):
                    embed = discord.Embed(
                        title="🌊 Hi Satvik when are we going surfing 🏄‍♂️",
                        description=f"Your Discord account has been linked to your email and roles have been assigned.",
                        color=discord.Color.green()
                    )

                else:
                    embed = discord.Embed(
                        title="✅ Successfully Logged In!",
                        description=f"Your Discord account has been linked to your email and roles have been assigned.",
                        color=discord.Color.green()
                    )

                # Set nickname after successful login
                if user_name and first_event:
                    nickname = f"{user_name} ({first_event})"
                    # Truncate to 32 characters (Discord limit)
                    if len(nickname) > 32:
                        nickname = nickname[:32]
                    try:
                        await handle_rate_limit(
                            user.edit(nick=nickname, reason="Login - setting nickname"),
                            f"editing nickname for {user}"
                        )
                        print(f"📝 Set nickname for {user} after login: '{nickname}'")
                    except discord.Forbidden:
                        print(f"❌ No permission to set nickname for {user}")
                    except Exception as e:
                        print(f"⚠️ Could not set nickname for {user}: {e}")

                # Build your information field
                info_text = f"**Name:** {user_name or 'Not specified'}\n"
                info_text += f"**Email:** {email}"

                if building and room:
                    info_text += f"\n**Location:** {building}, Room {room}"
                elif building:
                    info_text += f"\n**Building:** {building}"
                elif room:
                    info_text += f"\n**Room:** {room}"

                embed.add_field(
                    name="👤 Your Information",
                    value=info_text,
                    inline=False
                )

                roles_assigned = []                
                # 1. Add Master Role
                if master_role:
                    roles_assigned.append(master_role)
                # 2. Add ALL events from the semicolon-delimited list
                for r in roles:
                    if r and r not in roles_assigned:
                        roles_assigned.append(r)
                # 3. Add Secondary Role
                if secondary_role and secondary_role not in roles_assigned:
                    roles_assigned.append(secondary_role)

                # Add chapter role
                if chapter and chapter.lower() not in ["n/a", "na", ""]:
                    roles_assigned.append(chapter)
                    # Add to chapter role names set
                    chapter_role_names.add(chapter)
                else:
                    roles_assigned.append("Unaffiliated")
                    # Unaffiliated is also a chapter role
                    chapter_role_names.add("Unaffiliated")

                if roles_assigned:
                    embed.add_field(
                        name="🎭 Roles Assigned",
                        value="\n".join([f"• {role}" for role in roles_assigned]),
                        inline=False
                    )

                embed.add_field(
                    name="🎉 What's Next?",
                    value="• You now have access to relevant channels\n"
                        "• Your nickname has been updated\n"
                        "• Check out the channels you can now see!",
                    inline=False
                )

                embed.set_footer(text="Welcome to the team!")

                await interaction.followup.send(embed=embed, ephemeral=True)

            else:
                await interaction.followup.send(
                    "✅ Discord ID updated successfully, but could not sync with server. Please contact an admin.",
                    ephemeral=True
                )

        except Exception as e:
            await interaction.followup.send(
                f"❌ Error updating sheet: {str(e)}",
                ephemeral=True
            )
            print(f"❌ Error updating sheet for {email}: {e}")

    except Exception as e:
        print(f"❌ Error accessing sheet in login: {e}")
        print(f"❌ Error type: {type(e)}")
        import traceback
        traceback.print_exc()
        try:
            await interaction.followup.send(
                f"❌ Error accessing sheet: {str(e)}\n\nPlease contact an admin for help.",
                ephemeral=True
            )
        except Exception as followup_error:
            print(f"❌ Could not send error message via followup: {followup_error}")


@bot.tree.command(name="assignrunnerzones", description="Assign zone numbers per building in 'Runner Assignments' using K-means (Admin only)")
async def assign_runner_zones_command(interaction: discord.Interaction):
    """Read 'Runner Assignments' worksheet, cluster by building into K zones, write labels to 'Zone Number' column."""
    # Admin only
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    if admin_lock.locked():
        await interaction.response.send_message("❌ Server configurations are changing. Please try this when configurations is done!", ephemeral=True)
        return
    
    async with admin_lock:

        await interaction.response.defer(ephemeral=True)
    
        # Verify spreadsheet connection
        guild_id = interaction.guild.id
        if guild_id not in spreadsheets:
            await interaction.followup.send(
                "❌ No spreadsheet connected for this server! Use `/enterfolder` first to connect your sheet.",
                ephemeral=True
            )
            return
    
        spreadsheet = spreadsheets[guild_id]
    
        # Open the worksheet or find a separate spreadsheet in the same Drive folder
        worksheet_name = "Runner Assignments"
        ws = None
        try:
            ws = spreadsheet.worksheet(worksheet_name)
        except Exception:
            # If not a tab in the current spreadsheet, search the parent Drive folder for a spreadsheet named like it
            try:
                from googleapiclient.discovery import build
                drive_service = build('drive', 'v3', credentials=creds)
                # Get parent folder of the currently connected spreadsheet
                sheet_metadata = drive_service.files().get(fileId=spreadsheet.id, fields='parents').execute()
                parent_folders = sheet_metadata.get('parents', [])
                if not parent_folders:
                    await interaction.followup.send("❌ Could not determine parent folder to search for 'Runner Assignments' sheet.", ephemeral=True)
                    return
                parent_folder_id = parent_folders[0]
                # Search spreadsheets in same folder
                q = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and name contains '{worksheet_name}'"
                results = drive_service.files().list(q=q, fields='files(id, name)').execute()
                files = results.get('files', [])
                if not files:
                    await interaction.followup.send(f"❌ Could not find a spreadsheet named '{worksheet_name}' in the same folder as the template.", ephemeral=True)
                    return
                target = files[0]
                other_sheet = gc.open_by_key(target['id'])
                # Prefer a worksheet named exactly worksheet_name; otherwise first tab
                try:
                    ws = other_sheet.worksheet(worksheet_name)
                except Exception:
                    ws = other_sheet.worksheets()[0]
            except Exception as e2:
                await interaction.followup.send(f"❌ Could not locate '{worksheet_name}' in the same Drive folder: {str(e2)}", ephemeral=True)
                return
    
        # Fetch data
        try:
            headers = ws.row_values(1)
            rows = ws.get_all_records()
        except Exception as e:
            await interaction.followup.send(f"❌ Could not read worksheet data: {str(e)}", ephemeral=True)
            return

    # Normalize header names for lookups
    def _find_col_index(name_candidates):
        for i, h in enumerate(headers):
            if not h:
                continue
            for cand in name_candidates:
                if h.strip().lower() == cand:
                    return i + 1  # 1-indexed
        return None

    building_col_index = _find_col_index(["building"])
    coords_col_index = _find_col_index(["coordinates"])
    lat_col_index = _find_col_index(["latitude"])
    lon_col_index = _find_col_index(["longitude"])
    num_zones_col_index = _find_col_index(["number of zones"])
    zones_col_index = _find_col_index(["zone number"])
    num_zones = 0

    # Create zones column if missing
    if zones_col_index is None:
        try:
            new_col_idx = len(headers) + 1
            # Column letter (simple A..Z mapping consistent with rest of file usage)
            col_letter = chr(ord('A') + new_col_idx - 1)
            # Use user's preferred header name
            ws.update(f"{col_letter}1", [["zone number"]])
            headers.append("zone number")
            zones_col_index = new_col_idx
        except Exception as e:
            await interaction.followup.send(f"❌ Could not create 'Zone Number' column: {str(e)}", ephemeral=True)
            return

    # Build data per building
    from collections import defaultdict
    building_points = defaultdict(list)  # building -> list of (row_idx_1_based, (lat, lon))

    # Find the global K value from any row that has it
    global_k = None
    for row in rows:
        lower_row = { (k.strip().lower() if isinstance(k, str) else k): v for k, v in row.items() }
        k_raw = lower_row.get("number of zones", lower_row.get("zones count", lower_row.get("num zones", lower_row.get("k"))))
        if k_raw is not None and str(k_raw).strip() != "":
            try:
                global_k = int(float(k_raw))
                break  # Found it, use this value for all buildings
            except Exception:
                continue

    def _parse_float(val):
        try:
            if isinstance(val, str):
                val = val.strip()
                if not val:
                    return None
            return float(val)
        except Exception:
            return None

    for idx, row in enumerate(rows, start=2):  # data starts at row 2
        # Case-insensitive row access
        lower_row = { (k.strip().lower() if isinstance(k, str) else k): v for k, v in row.items() }
        building = str(lower_row.get("building", lower_row.get("building 1", ""))).strip()
        if not building:
            continue

        lat = _parse_float(lower_row.get("latitude", lower_row.get("lat")))
        lon = _parse_float(lower_row.get("longitude", lower_row.get("lon", lower_row.get("lng"))))

        if (lat is None or lon is None) and ("coordinates" in lower_row and lower_row["coordinates"]):
            coord_str = str(lower_row["coordinates"]).strip()
            if "," in coord_str:
                parts = [p.strip() for p in coord_str.split(",")]
                if len(parts) >= 2:
                    if lat is None:
                        lat = _parse_float(parts[0])
                    if lon is None:
                        lon = _parse_float(parts[1])

        if lat is None or lon is None:
            continue

        building_points[building].append((idx, (lat, lon)))

    if not building_points:
        await interaction.followup.send("⚠️ No valid location rows found to cluster.", ephemeral=True)
        return

    # Compute clusters and prepare updates
    updates = []  # list of (row_index, zone_label_str)
    k_to_use = global_k if global_k is not None and global_k > 0 else 1

    # Collect ALL points from ALL buildings for global clustering
    all_points = []
    all_items = []  # (building, row_idx, point)

    for bldg, items in building_points.items():
        for row_idx, point in items:
            all_points.append(point)
            all_items.append((bldg, row_idx, point))

    # Run K-means on ALL points together
    labels = _run_kmeans_clustering(all_points, k_to_use)

    # Debug: count cluster distribution by building
    building_zone_counts = {}  # building -> {zone: count}
    for i, (bldg, row_idx, point) in enumerate(all_items):
        zone = labels[i] + 1  # Convert to 1-based
        if bldg not in building_zone_counts:
            building_zone_counts[bldg] = {}
        building_zone_counts[bldg][zone] = building_zone_counts[bldg].get(zone, 0) + 1

    debug_info = []
    for bldg, zone_counts in building_zone_counts.items():
        zones = sorted(zone_counts.keys())
        counts = [zone_counts[z] for z in zones]
        debug_info.append(f"{bldg}: zones {zones} (counts {counts})")

    # Create updates
    for i, (bldg, row_idx, point) in enumerate(all_items):
        zone_label = str(labels[i] + 1)  # Convert to 1-based
        updates.append((row_idx, zone_label))

    # Apply updates (per cell to minimize risk of range mistakes)
    zones_col_letter = chr(ord('A') + zones_col_index - 1)
    updated = 0
    for row_idx, value in updates:
        try:
            ws.update(f"{zones_col_letter}{row_idx}", [[value]])
            updated += 1
        except Exception:
            pass

    # Summarize K used per building (limit for brevity)
    # Send debug info first
    debug_text = "\n".join(debug_info[:5])  # Limit to first 5 buildings
    if len(debug_info) > 5:
        debug_text += f"\n... and {len(debug_info) - 5} more buildings"

    await interaction.followup.send(
        f"✅ Assigned {k_to_use} zones for {len(updates)} rows across {len(building_points)} buildings in '{worksheet_name}'.\n\n"
        f"Now sending runner assignments to building channels...",
        ephemeral=True
    )

    # Send messages to each building channel with their designated runners
    try:
        guild = interaction.guild
        
        # Get the main sheet to cross-reference Discord IDs
        main_sheet = sheets.get(guild_id)
        main_data = []
        if main_sheet:
            try:
                main_data = main_sheet.get_all_records()
            except Exception as e:
                print(f"⚠️ Could not access main sheet for Discord IDs: {e}")
        
        # Create email -> discord_id mapping from main sheet
        email_to_discord = {}
        for row in main_data:
            email = str(row.get("Email", "")).strip().lower()
            discord_id = str(row.get("Discord ID", "")).strip()
            if email and discord_id:
                try:
                    email_to_discord[email] = int(discord_id)
                except ValueError:
                    pass
        
        # First pass: find all buildings and their zones
        building_zones = {}  # building -> zone_number
        zone_buildings = defaultdict(list)  # zone -> [building names]
        
        for idx, row in enumerate(rows, start=2):
            lower_row = {(k.strip().lower() if isinstance(k, str) else k): v for k, v in row.items()}
            building = str(lower_row.get("building", lower_row.get("building 1", ""))).strip()
            zone_raw = str(lower_row.get("zone number", lower_row.get("zone", ""))).strip()
            
            # If this row has a building and zone, it's a building definition
            if building and zone_raw:
                try:
                    zone_num = int(zone_raw)
                    building_zones[building] = zone_num
                    zone_buildings[zone_num].append(building)
                    print(f"📍 Found building: {building} → Zone {zone_num}")
                except ValueError:
                    pass
        
        # Second pass: find all runners and their zones from Runner Zone column
        zone_runners = defaultdict(list)  # zone -> [(name, discord_id)]
        
        # Debug: print available columns
        if rows:
            first_row = rows[0]
            lower_first = {(k.strip().lower() if isinstance(k, str) else k): v for k, v in first_row.items()}
            print(f"🔍 DEBUG: Available columns: {list(lower_first.keys())}")
        
        for idx, row in enumerate(rows, start=2):
            lower_row = {(k.strip().lower() if isinstance(k, str) else k): v for k, v in row.items()}
            name = str(lower_row.get("name", "")).strip()
            email = str(lower_row.get("email", "")).strip().lower()
            
            # Use "Runner Zone" column specifically (NOT "Zone Number")
            runner_zone_raw = str(lower_row.get("runner zone", "")).strip()
            
            # Debug first few rows
            if name and idx <= 10:
                print(f"🔍 Row {idx}: name={name}, email={email}, runner_zone={runner_zone_raw}")
            
            # If this row has a name and runner zone, it's a runner
            if name and runner_zone_raw:
                try:
                    zone_num = int(runner_zone_raw)
                    discord_id = email_to_discord.get(email)
                    zone_runners[zone_num].append((name, discord_id))
                    print(f"🏃 Found runner: {name} (email: {email}) → Zone {zone_num}")
                except ValueError:
                    print(f"⚠️ Could not parse runner zone '{runner_zone_raw}' for {name}")
        
        # Match runners to buildings by zone
        building_runners = defaultdict(list)  # building -> [(name, discord_id)]
        
        for building, zone_num in building_zones.items():
            # Get all runners for this zone
            runners_in_zone = zone_runners.get(zone_num, [])
            building_runners[building] = runners_in_zone
            print(f"✅ Matched {len(runners_in_zone)} runners to building {building} (zone {zone_num})")
        
        # Send message to each building channel
        messages_sent = 0
        for building, runners in building_runners.items():
            # Find the building chat channel
            building_chat_name = f"{building.lower().replace(' ', '-')}-chat"
            building_channel = discord.utils.get(guild.text_channels, name=building_chat_name)
            
            if not building_channel:
                print(f"⚠️ Could not find building channel: {building_chat_name}")
                continue
            
            if not runners:
                continue
            
            # Build the message with mentions
            embed = discord.Embed(
                title=f"🏃 Designated Runners for {building}",
                description="Here are the runners assigned to help with this building:",
                color=discord.Color.orange()
            )
            
            # Create list of runner mentions/names
            runner_mentions = []
            for name, discord_id in runners:
                if discord_id:
                    runner_mentions.append(f"• <@{discord_id}>")
                else:
                    runner_mentions.append(f"• {name}")
            
            runners_text = "\n".join(runner_mentions)
            embed.add_field(
                name=f"Runners:",
                value=runners_text,
                inline=False
            )
            
            embed.set_footer(text="If you need help, create a ticket in the #help forum!\nDM these runners if you need urgent help!")
            
            try:
                await building_channel.send(embed=embed)
                messages_sent += 1
                print(f"✅ Sent runner assignments to {building_chat_name}")
            except Exception as e:
                print(f"⚠️ Error sending message to {building_chat_name}: {e}")
        
        print(f"✅ Sent runner assignments to {messages_sent} building channels")
        
    except Exception as e:
        print(f"⚠️ Error sending runner assignments to channels: {e}")

    


@tasks.loop(minutes=60)
async def sync_members():
    """Every minute, read each guild's spreadsheet and invite any new Discord IDs."""
    print("🔄 Running member sync...")

    # Check if we have any sheets connected
    if not sheets:
        print("⚠️ No sheets connected - use /enterfolder to connect to a sheet in each server")
        return

    # Sync each guild with its own sheet
    total_processed = 0
    for guild in bot.guilds:
        guild_id = guild.id

        # Skip guilds without a sheet connection
        if guild_id not in sheets:
            print(f"⚠️ No sheet connected for guild {guild.name} - skipping")
            continue

        try:
            # Fetch all rows from this guild's sheet
            data = sheets[guild_id].get_all_records()
            print(f"📊 Found {len(data)} rows in spreadsheet for {guild.name}")

            print(f"🔄 Syncing members for guild: {guild.name}")
            sync_results = await perform_member_sync(guild, data)
            total_processed += sync_results['processed']
            print(f"✅ Sync complete for {guild.name}. Processed {sync_results['processed']} valid Discord IDs.")
        except Exception as e:
            print(f"❌ Error syncing guild {guild.name}: {e}")

    print(f"✅ Total sync complete. Processed {total_processed} valid Discord IDs across {len(bot.guilds)} guilds.")


@tasks.loop(minutes=1)
async def check_help_tickets():
    """Every minute, check for unresponded help tickets and re-ping if needed."""
    if not active_help_tickets:
        return

    print(f"🎫 Checking {len(active_help_tickets)} active help tickets...")

    current_time = datetime.now()
    tickets_to_remove = []

    for thread_id, ticket_info in active_help_tickets.items():
        try:
            # Calculate time since last ping
            time_since_created = current_time - ticket_info["created_at"]

            # Determine wait time based on ping count
            # ping_count = 1 (after 1st ping): wait 3 minutes for 2nd ping
            # ping_count = 2 (after 2nd ping): wait 1 minute for 3rd ping
            ping_count = ticket_info["ping_count"]
            if ping_count == 1:
                wait_time = timedelta(minutes=3)
            elif ping_count == 2:
                wait_time = timedelta(minutes=1)
            else:
                # Should not reach here, but default to 3 minutes
                wait_time = timedelta(minutes=1)

            # Check if enough time has passed since last ping
            if time_since_created >= wait_time:
                # Find the thread across all guilds
                thread = None
                for guild in bot.guilds:
                    thread = guild.get_thread(thread_id)
                    if thread:
                        break

                if not thread:
                    print(f"⚠️ Thread {thread_id} not found in any guild, removing from tracking")
                    tickets_to_remove.append(thread_id)
                    continue

                # Check if thread is still active/not archived
                if thread.archived or thread.locked:
                    print(f"🗄️ Thread {thread_id} is archived/locked, removing from tracking")
                    tickets_to_remove.append(thread_id)
                    continue

                # Check ping count limit (max 3 pings)
                if ticket_info["ping_count"] >= 3:
                    print(f"⏹️ Thread {thread_id} reached max ping limit, removing from tracking")
                    tickets_to_remove.append(thread_id)
                    continue

                # Re-ping the runners
                await send_ticket_repings(thread, ticket_info)

                # Update ping count and reset timer
                ticket_info["ping_count"] += 1
                ticket_info["created_at"] = current_time
                print(f"🔄 Re-pinged ticket {thread_id} (ping #{ticket_info['ping_count']})")

        except Exception as e:
            print(f"❌ Error checking ticket {thread_id}: {e}")
            tickets_to_remove.append(thread_id)

    # Clean up invalid tickets
    for thread_id in tickets_to_remove:
        if thread_id in active_help_tickets:
            del active_help_tickets[thread_id]
            print(f"🗑️ Removed invalid ticket {thread_id} from tracking")


async def get_all_runners(guild_id):
    """Get Discord IDs of ALL runners from the Runner Assignments sheet"""
    if guild_id not in spreadsheets:
        print(f"❌ No spreadsheet connected for guild {guild_id}")
        return []

    try:
        spreadsheet = spreadsheets[guild_id]

        # Try to get the Runner Assignments worksheet
        try:
            sheet = spreadsheet.worksheet("Runner Assignments")
        except Exception:
            # If not found as a worksheet, search for a separate spreadsheet
            try:
                from googleapiclient.discovery import build
                drive_service = build('drive', 'v3', credentials=creds)

                # Get parent folder of the currently connected spreadsheet
                sheet_metadata = drive_service.files().get(fileId=spreadsheet.id, fields='parents').execute()
                parent_folders = sheet_metadata.get('parents', [])
                if not parent_folders:
                    print("❌ Could not determine parent folder for all runners lookup")
                    return []

                parent_folder_id = parent_folders[0]

                # Search for Runner Assignments spreadsheet
                q = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and name contains 'Runner Assignments'"
                results = drive_service.files().list(q=q, fields='files(id, name)').execute()
                files = results.get('files', [])

                if not files:
                    print("❌ Could not find Runner Assignments spreadsheet")
                    return []

                # Open the first matching spreadsheet
                runner_spreadsheet = gc.open_by_key(files[0]['id'])
                sheet = runner_spreadsheet.sheet1  # Use first worksheet

            except Exception as e:
                print(f"❌ Error finding Runner Assignments spreadsheet: {e}")
                return []

        # Get all data from the sheet
        data = sheet.get_all_records()

        # Find all runner emails (anyone with a "Runner Zone" value)
        runner_emails = []
        for row in data:
            runner_zone = row.get("Runner Zone", "")
            if runner_zone:  # Has a runner zone assigned
                email = str(row.get("Email", "")).strip()
                if email:
                    runner_emails.append(email.lower())

        if not runner_emails:
            print("⚠️ No runner emails found in Runner Assignments")
            return []

        print(f"🔍 Found {len(runner_emails)} total runner emails")

        # Now cross-reference with the main sheet to get Discord IDs
        try:
            main_sheet = spreadsheet.worksheet(SHEET_PAGE_NAME)
            main_data = main_sheet.get_all_records()
        except Exception as e:
            print(f"❌ Error accessing main sheet for Discord ID lookup: {e}")
            return []

        all_runners = []
        for row in main_data:
            email = str(row.get("Email", "")).strip().lower()
            if email in runner_emails:
                discord_id = str(row.get("Discord ID", "")).strip()
                if discord_id:
                    try:
                        all_runners.append(int(discord_id))
                    except ValueError:
                        print(f"⚠️ Invalid Discord ID '{discord_id}' for runner email {email}")

        print(f"✅ Found {len(all_runners)} total runner Discord IDs")
        return all_runners

    except Exception as e:
        print(f"❌ Error looking up all runners: {e}")
        return []


async def send_ticket_repings(thread, ticket_info):
    """Send re-ping message for a help ticket"""
    try:
        ping_count = ticket_info["ping_count"] + 1

        # For final ping (3rd ping), get ALL runners instead of just zone runners
        if ping_count >= 3:
            print(f"🚨 Final ping for ticket {thread.id} - getting ALL runners")
            guild_id = thread.guild.id
            all_runner_ids = await get_all_runners(guild_id)
            runner_mentions = []
            for runner_id in all_runner_ids:
                member = thread.guild.get_member(runner_id)
                if member:
                    runner_mentions.append(member.mention)
        else:
            # Regular ping - just zone runners
            runner_mentions = []
            for runner_id in ticket_info["zone_runners"]:
                member = thread.guild.get_member(runner_id)
                if member:
                    runner_mentions.append(member.mention)

        if not runner_mentions:
            print(f"⚠️ No valid runners found for re-ping in ticket {thread.id}")
            return

        mention_text = " ".join(runner_mentions)

        # Build location info
        location_parts = [ticket_info["building"]]
        if ticket_info["room"]:
            location_parts.append(f"Room {ticket_info['room']}")
        location = ", ".join(location_parts)

        # Different field names for final ping vs regular ping
        if ping_count < 3:
            embed = discord.Embed(
                title=f"Still Need Help!",
                description=f"**Event:** {ticket_info['event']}\n**Location:** {location}",
                color=discord.Color.orange()
            )
            embed.add_field(
                name="Runners Assigned",
                value=f"This ticket still needs assistance!",
                inline=False
            )
        else:
            embed = discord.Embed(
                title=f"Final Call!",
                description=f"**Event:** {ticket_info['event']}\n**Location:** {location}",
                color=discord.Color.red()
            )
            embed.add_field(
                name="ALL RUNNERS",
                value=f"This ticket still needs assistance!",
                inline=False
            )
            embed.add_field(
                name="\nFinal Ping",
                value="This is the final automatic ping. Please respond if you can help!",
                inline=False
            )

        # Send mentions as regular message content (not in embed) so Discord actually notifies users
        await thread.send(content=mention_text, embed=embed)
        print(f"📢 Sent re-ping #{ping_count} for ticket {thread.id}")

    except Exception as e:
        print(f"❌ Error sending re-ping for ticket {thread.id}: {e}")


@bot.tree.command(name="debugzone", description="Debug zone assignment for a user (Admin only)")
@app_commands.describe(user="The user to debug zone assignment for")
async def debug_zone_command(interaction: discord.Interaction, user: discord.Member):
    """Debug command to test zone assignment for a specific user"""
    # Admin only
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True)

    try:
        guild_id = interaction.guild.id

        # Look up the user's event and building
        user_event_info = await get_user_event_building(guild_id, user.id)
        if not user_event_info:
            await interaction.followup.send(f"❌ Could not find event/building info for {user.mention}")
            return

        building = user_event_info.get("building")
        event = user_event_info.get("event")
        room = user_event_info.get("room")
        name = user_event_info.get("name")

        if not building:
            await interaction.followup.send(f"❌ No building found for {user.mention} (event: {event})")
            return

        # Get the zone for this building
        zone = await get_building_zone(guild_id, building)
        if not zone:
            await interaction.followup.send(f"❌ No zone found for building '{building}'")
            return

        # Get all runners in this zone
        zone_runners = await get_zone_runners(guild_id, zone)

        # Create embed with debug info
        embed = discord.Embed(
            title="🐛 Zone Debug Info",
            description=f"Debug information for {user.mention}",
            color=discord.Color.blue()
        )

        # Build location info for debug display
        location_parts = [building]
        if room:
            location_parts.append(f"Room {room}")
        location = ", ".join(location_parts)

        embed.add_field(name="User Info", value=f"**Name:** {name}\n**Event:** {event}\n**Location:** {location}", inline=False)
        embed.add_field(name="Zone Assignment", value=f"**Zone:** {zone}", inline=False)

        if zone_runners:
            runner_mentions = []
            for runner_id in zone_runners:
                member = interaction.guild.get_member(runner_id)
                if member:
                    runner_mentions.append(member.mention)
                else:
                    runner_mentions.append(f"<@{runner_id}> (not in server)")

            embed.add_field(
                name=f"Zone {zone} Runners ({len(zone_runners)} total)",
                value="\n".join(runner_mentions) if runner_mentions else "No valid runners found",
                inline=False
            )
        else:
            embed.add_field(name="Zone Runners", value="No runners found for this zone", inline=False)

        await interaction.followup.send(embed=embed)

    except Exception as e:
        await interaction.followup.send(f"❌ Error during debug: {str(e)}")
        print(f"❌ Debug zone error: {e}")
        import traceback
        traceback.print_exc()


@bot.tree.command(name="activetickets", description="Show all active help tickets being tracked (Admin only)")
async def active_tickets_command(interaction: discord.Interaction):
    """Debug command to show all active help tickets"""
    # Admin only
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True)

    try:
        if not active_help_tickets:
            await interaction.followup.send("✅ No active help tickets being tracked.")
            return

        embed = discord.Embed(
            title="🎫 Active Help Tickets",
            description=f"Currently tracking {len(active_help_tickets)} help tickets",
            color=discord.Color.blue()
        )

        for thread_id, ticket_info in list(active_help_tickets.items())[:10]:  # Limit to 10 for display
            # Get thread info
            thread = interaction.guild.get_thread(thread_id)
            thread_name = thread.name if thread else f"Thread {thread_id} (not found)"

            # Calculate time since creation
            time_elapsed = datetime.now() - ticket_info["created_at"]
            minutes_elapsed = int(time_elapsed.total_seconds() / 60)

            # Build location info
            location_parts = [ticket_info["building"]]
            if ticket_info["room"]:
                location_parts.append(f"Room {ticket_info['room']}")
            location = ", ".join(location_parts)

            embed.add_field(
                name=f"🎫 {thread_name}",
                value=f"**Event:** {ticket_info['event']}\n**Location:** {location}\n**Zone:** {ticket_info['zone']}\n**Pings:** {ticket_info['ping_count']}\n**Time:** {minutes_elapsed}m ago",
                inline=True
            )

        if len(active_help_tickets) > 10:
            embed.add_field(
                name="📋 Note",
                value=f"Showing first 10 of {len(active_help_tickets)} active tickets",
                inline=False
            )

        await interaction.followup.send(embed=embed)

    except Exception as e:
        await interaction.followup.send(f"❌ Error fetching active tickets: {str(e)}")
        print(f"❌ Active tickets error: {e}")
        import traceback
        traceback.print_exc()


@bot.tree.command(name="stopburgers", description="Stop all active burger deliveries")
async def stop_burgers_command(interaction: discord.Interaction):
    """Emergency stop command for burger deliveries"""
    await interaction.response.defer(ephemeral=True)

    try:
        if not active_burger_deliveries:
            await interaction.followup.send("✅ No active burger deliveries to stop.", ephemeral=True)
            return

        stopped_count = len(active_burger_deliveries)
        print(f"🛑 Stopping {stopped_count} burger deliveries by {interaction.user}")

        # Set stop flag for all active deliveries
        for user_id, delivery_info in list(active_burger_deliveries.items()):
            delivery_info["stop"] = True
            print(f"🛑 Set stop flag for user {user_id}")

        await interaction.followup.send(
            f"🛑 Emergency stop activated!\n\n"
            f"Stopped **{stopped_count}** active burger deliver{'y' if stopped_count == 1 else 'ies'}.\n"
            f"Users will receive a final 'Grill exploded' message.",
            ephemeral=True
        )
        print(f"✅ Stopped {stopped_count} burger deliveries")

    except Exception as e:
        await interaction.followup.send(f"❌ Error stopping burgers: {str(e)}", ephemeral=True)
        print(f"❌ Stop burgers error: {e}")
        import traceback
        traceback.print_exc()


@bot.tree.command(name="sendallmaterials", description="Send all materials to all channels (Admin only)")
async def send_all_materials_command(interaction: discord.Interaction):
    """Manually trigger materials search and sharing for all events, useful links, and runner"""
    # Admin only
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    if admin_lock.locked():
        await interaction.response.send_message("❌ Server configurations are changing. Please try this when configurations is done!", ephemeral=True)
        return
    
    async with admin_lock:

        await interaction.response.defer(ephemeral=True)

        try:
            guild = interaction.guild

            print(f"📚 Manual test materials request by {interaction.user} for all events")

            # Check if we have a spreadsheet connected
            guild_id = guild.id
            if guild_id not in spreadsheets:
                await interaction.followup.send(
                    "❌ No spreadsheet connected for this server!\n\n"
                    "Use `/enterfolder` to connect to a Google Drive folder first.",
                    ephemeral=True
                )
                return

            # Get all roles in the server
            priority_roles = ["Admin", "Volunteer", "Lead ES", "Social Media", "Photographer", "Arbitrations", "Awards", "Runner", "VIPer"]

            # Find all event roles (roles that aren't priority/system roles)
            event_roles = []
            for role in guild.roles:
                # Skip @everyone, bot-managed roles, and priority roles
                if (role.name != "@everyone" and
                    not role.managed and
                    role.name not in priority_roles and
                    role.name not in chapter_role_names):  # Skip chapter roles too
                    event_roles.append(role.name)

            if not event_roles:
                await interaction.followup.send(
                    "❌ No event roles found in this server!\n\n"
                    "Make sure you've run `/enterfolder` to create event roles first.",
                    ephemeral=True
                )
                return

            # Send initial status
            await interaction.followup.send(
                f"🔄 Processing test materials for {len(event_roles)} event(s)...\n\n"
                f"• Deleting old pinned test materials\n"
                f"• Sending new test materials\n\n"
                f"This may take a while. Check the event channels for results.",
                ephemeral=True
            )

            # Loop through all event roles and send test materials
            success_count = 0
            deleted_count = 0
            for role_name in event_roles:
                try:
                    print(f"📚 Processing test materials for: {role_name}")
                    
                    # First, delete old pinned test materials from this event's channels
                    sanitized_role_name = role_name.lower().replace(' ', '-').replace('/', '-').replace('\\', '-').replace(':', '-').replace('*', '-').replace('?', '-').replace('"', '').replace('<', '').replace('>', '').replace('|', '-')
                    
                    # Find all channels that belong to this event (start with the sanitized role name)
                    event_channels = [ch for ch in guild.text_channels if ch.name.startswith(sanitized_role_name + "-")]
                    
                    for channel in event_channels:
                        try:
                            # Get all pinned messages in this channel
                            pinned_messages = await channel.pins()
                            
                            # Delete pinned messages that were sent by the bot
                            for msg in pinned_messages:
                                if msg.author == bot.user:
                                    try:
                                        await msg.delete()
                                        deleted_count += 1
                                        print(f"🗑️ Deleted old pinned test material from #{channel.name}")
                                        # Small delay to avoid rate limiting
                                        await asyncio.sleep(0.2)
                                    except Exception as delete_error:
                                        print(f"⚠️ Could not delete message in #{channel.name}: {delete_error}")
                        except Exception as pin_error:
                            print(f"⚠️ Error checking pins in #{channel.name}: {pin_error}")
                    
                    # Now send new test materials
                    print(f"📚 Searching test materials for: {role_name}")
                    await search_and_share_test_folder(guild, role_name)
                    success_count += 1
                    # Small delay to avoid rate limiting
                    await asyncio.sleep(0.5)
                except Exception as e:
                    print(f"⚠️ Error sending test materials for {role_name}: {e}")

            # Release useful links
            try:
                print("🔗 Releasing useful links...")
                await search_and_share_useful_links(guild)
                print("✅ Useful links released")
            except Exception as useful_links_error:
                print(f"⚠️ Error releasing useful links: {useful_links_error}")

            # Release runner info
            try:
                print("🏃 Releasing runner info...")
                await search_and_share_runner_info(guild)
                print("✅ Runner info released")
            except Exception as runner_info_error:
                print(f"⚠️ Error releasing runner info: {runner_info_error}")

            result_embed = discord.Embed(
                title="✅ All Test Materials Sent",
                description=f"Successfully sent test materials for **{success_count}/{len(event_roles)}** events!",
                color=discord.Color.green()
            )
            
            if deleted_count > 0:
                result_embed.add_field(
                    name="🗑️ Old Materials Cleaned",
                    value=f"Deleted **{deleted_count}** old pinned test material message(s)",
                    inline=False
                )

            await interaction.followup.send(embed=result_embed, ephemeral=True)
            print(f"✅ Test materials command completed: {success_count}/{len(event_roles)} events processed, {deleted_count} old messages deleted")

        except Exception as e:
            await interaction.followup.send(f"❌ Error sending test materials: {str(e)}", ephemeral=True)
            print(f"❌ Send test materials error: {e}")
            import traceback
            traceback.print_exc()


@bot.tree.command(name="cacheinfo", description="Show cached spreadsheet connection info (Admin only)")
async def cache_info_command(interaction: discord.Interaction):
    """Show information about the cached spreadsheet connection"""
    # Admin only
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True)

    try:
        cache = load_cache()

        if not cache:
            await interaction.followup.send("📄 No cache file found.")
            return

        embed = discord.Embed(
            title="💾 Cache Information",
            description="Current cached spreadsheet connection details",
            color=discord.Color.blue()
        )

        # Basic connection info
        if cache.get("spreadsheet_id"):
            embed.add_field(
                name="📊 Spreadsheet",
                value=f"**Title:** {cache.get('spreadsheet_title', 'Unknown')}\n**ID:** `{cache.get('spreadsheet_id')}`",
                inline=False
            )

        if cache.get("worksheet_name"):
            embed.add_field(
                name="📋 Worksheet",
                value=cache.get("worksheet_name"),
                inline=True
            )

        if cache.get("connected_at"):
            try:
                connected_time = datetime.fromisoformat(cache.get("connected_at"))
                embed.add_field(
                    name="🕐 Connected At",
                    value=connected_time.strftime("%Y-%m-%d %H:%M:%S"),
                    inline=True
                )
            except:
                embed.add_field(name="🕐 Connected At", value="Unknown format", inline=True)

        if cache.get("folder_link"):
            embed.add_field(
                name="📁 Folder Link",
                value=f"[Open in Drive]({cache.get('folder_link')})",
                inline=False
            )

        # Cache file info
        if os.path.exists(CACHE_FILE):
            file_size = os.path.getsize(CACHE_FILE)
            embed.add_field(
                name="📄 Cache File",
                value=f"**Path:** `{CACHE_FILE}`\n**Size:** {file_size} bytes",
                inline=False
            )

        await interaction.followup.send(embed=embed)

    except Exception as e:
        await interaction.followup.send(f"❌ Error reading cache: {str(e)}")
        print(f"❌ Cache info error: {e}")


@bot.tree.command(name="clearcache", description="Clear the cached spreadsheet connection (Admin only)")
async def clear_cache_command(interaction: discord.Interaction):
    """Clear the cached spreadsheet connection"""
    # Admin only
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    if admin_lock.locked():
        await interaction.response.send_message("❌ Server configurations are changing. Please try this when configurations is done!", ephemeral=True)
        return
        
    async with admin_lock:

        await interaction.response.defer(ephemeral=True)

        try:
            guild_id = interaction.guild.id

            # Clear the guild-specific cache
            cleared = clear_guild_cache(guild_id)

            # Also clear the current connection for this guild
            if guild_id in sheets:
                del sheets[guild_id]
            if guild_id in spreadsheets:
                del spreadsheets[guild_id]

            if cleared or guild_id in sheets or guild_id in spreadsheets:
                embed = discord.Embed(
                    title="🗑️ Cache Cleared",
                    description="Cached spreadsheet connection for this server has been cleared.\nUse `/enterfolder` to reconnect to a sheet.",
                    color=discord.Color.green()
                )
            else:
                embed = discord.Embed(
                    title="📄 No Cache Found",
                    description="No cached connection found for this server.",
                    color=discord.Color.orange()
                )

            await interaction.followup.send(embed=embed)
            print(f"🧹 Admin cleared the cache for guild {guild_id} via command")

        except Exception as e:
            await interaction.followup.send(f"❌ Error clearing cache: {str(e)}")
            print(f"❌ Clear cache error: {e}")

@bot.tree.command(name="sendsingularmaterial", description="Send material for a specific type (Admin only)")
@app_commands.describe(
    material_type="Type of material to send (event, useful-links, or runner)",
    event_name="Event name (only needed if material_type is 'event')"
)
@app_commands.choices(material_type=[
    app_commands.Choice(name="Event Test Materials", value="event"),
    app_commands.Choice(name="Useful Links", value="useful-links"),
    app_commands.Choice(name="Runner Info", value="runner")
])
async def send_singular_material_command(interaction: discord.Interaction, material_type: app_commands.Choice[str], event_name: str = None):
    """Send materials for a specific type (event, useful-links, or runner)"""

    # Check if user has administrator permission
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

    if admin_lock.locked():
        await interaction.response.send_message("❌ Server configurations are changing. Please try this when configurations is done!", ephemeral=True)
        return
    
    async with admin_lock:

        await interaction.response.defer(ephemeral=True)

        try:
            guild = interaction.guild
            material_value = material_type.value

            print(f"📚 Manual material request by {interaction.user} for {material_value}")

            # Check if we have a spreadsheet connected
            guild_id = guild.id
            if guild_id not in spreadsheets:
                await interaction.followup.send(
                    "❌ No spreadsheet connected for this server!\n\n"
                    "Use `/enterfolder` to connect to a Google Drive folder first.",
                    ephemeral=True
                )
                return

            # Handle event materials
            if material_value == "event":
                if not event_name:
                    await interaction.followup.send(
                        "❌ Event name is required when sending event test materials!",
                        ephemeral=True
                    )
                    return

                # Get all roles in the server
                priority_roles = ["Admin", "Volunteer", "Lead ES", "Social Media", "Photographer", "Arbitrations", "Awards", "Runner", "VIPer"]

                if (event_name not in [role.name for role in guild.roles] and
                    event_name in priority_roles and
                    event_name in chapter_role_names):
                    await interaction.followup.send(
                        "❌ This event does not exist in this server or is not an appropriate argument!",
                        ephemeral=True
                    )
                    return

                # Send initial status
                await interaction.followup.send(
                    f"🔍 Searching for test materials for {event_name}...\n\n"
                    f"This may take a while. Check the event channels for results.",
                    ephemeral=True
                )

                try:
                    print(f"📚 Searching test materials for: {event_name}")
                    await search_and_share_test_folder(guild, event_name)
                    await asyncio.sleep(0.5)
                    print(f"✅ Test materials sent for {event_name}")
                except Exception as e:
                    print(f"⚠️ Error sending test materials for {event_name}: {e}")
                    await interaction.followup.send(f"❌ Error: {str(e)}", ephemeral=True)
                    return

            # Handle useful links
            elif material_value == "useful-links":
                await interaction.followup.send(
                    "🔗 Searching for useful links...\n\n"
                    "This may take a moment.",
                    ephemeral=True
                )

                try:
                    print("🔗 Releasing useful links...")
                    await search_and_share_useful_links(guild)
                    print("✅ Useful links released")
                except Exception as e:
                    print(f"⚠️ Error releasing useful links: {e}")
                    await interaction.followup.send(f"❌ Error: {str(e)}", ephemeral=True)
                    return

            # Handle runner info
            elif material_value == "runner":
                await interaction.followup.send(
                    "🏃 Searching for runner info...\n\n"
                    "This may take a moment.",
                    ephemeral=True
                )

                try:
                    print("🏃 Releasing runner info...")
                    await search_and_share_runner_info(guild)
                    print("✅ Runner info released")
                except Exception as e:
                    print(f"⚠️ Error releasing runner info: {e}")
                    await interaction.followup.send(f"❌ Error: {str(e)}", ephemeral=True)
                    return

            result_embed = discord.Embed(
                title=f"✅ Material Sent Successfully",
                description=f"Successfully sent {material_type.name}!",
                color=discord.Color.green()
            )

            await interaction.followup.send(embed=result_embed, ephemeral=True)
            print(f"✅ Material command completed for {material_value}")

        except Exception as e:
            await interaction.followup.send(f"❌ Error sending materials: {str(e)}", ephemeral=True)
            print(f"❌ Send materials error: {e}")
            import traceback
            traceback.print_exc()

@bot.tree.command(name="set_runner_all_access", description="Set if runners get access to all building/event channels (Admin only)")
@app_commands.describe(
    runner_access="1 to give Runners access to all rooms, 0 to restrict them to static channels",
)
async def set_runner_all_access_command(interaction: discord.Interaction, runner_access: int):
    """Set Runner All Access Command"""
    global runner_all_access

    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return
    
    await interaction.response.defer(ephemeral=True)
    
    guild = interaction.guild
    guild_id = guild.id
    
    runner_access_bool = bool(runner_access)
    # Default to 0 if the server hasn't set it yet
    current_access_bool = bool(runner_all_access.get(guild_id, 0))

    if runner_access_bool == current_access_bool:
        state = "already HAS" if runner_access_bool else "is already RESTRICTED from"
        await interaction.followup.send(f"⚠️ The Runner role {state} all-access.", ephemeral=True)
        return

    # Save the setting per-guild in memory
    runner_all_access[guild_id] = 1 if runner_access_bool else 0
    
    # Save to the permanent cache file so it survives bot restarts
    cache = load_cache()
    cache["runner_access_settings"] = runner_all_access
    save_cache(cache)
    
    runner_role = discord.utils.get(guild.roles, name="Runner")
    if not runner_role:
        await interaction.followup.send("❌ Runner role not found! Please create it or run /enterfolder first.", ephemeral=True)
        return

    try:
        # Added "Chapters" to protect them from being accidentally modified
        static_categories = ["Welcome", "Tournament Officials", "Volunteers", "Chapters"]
        modified_count = 0
        
        for category in guild.categories:
            # Skip the static categories; only target building/room categories
            if category.name not in static_categories:
                for room in category.channels:
                    # Make sure we are only modifying text or voice channels
                    if isinstance(room, discord.TextChannel) or isinstance(room, discord.VoiceChannel):
                        overwrites = room.overwrites

                        if runner_access_bool:
                            # Grant access to this specific room
                            overwrites[runner_role] = discord.PermissionOverwrite(
                                read_messages=True,
                                send_messages=True,
                                read_message_history=True
                            )
                        else:
                            # Remove Runner overwrite entirely so it defaults to hidden
                            if runner_role in overwrites:
                                del overwrites[runner_role]

                        # Apply the changes via API
                        await handle_rate_limit(
                            room.edit(overwrites=overwrites, reason=f"Updated Runner all-access to {runner_access_bool}"),
                            f"editing channel '{room.name}' permissions"
                        )
                        modified_count += 1
        
        action_text = "GRANTED access to" if runner_access_bool else "REMOVED access from"
        await interaction.followup.send(f"✅ Successfully **{action_text}** {modified_count} building/event channels for the Runner role.", ephemeral=True)
        
    except discord.Forbidden:
        await interaction.followup.send("❌ Bot lacks permissions to modify channel overwrites.", ephemeral=True)
        print("❌ Error: Bot forbidden from editing channel overwrites.")
    except Exception as e:
        await interaction.followup.send(f"❌ Error updating runner access: {str(e)}", ephemeral=True)
        print(f"❌ Error in set_runner_all_access: {e}")

@bot.tree.command(name="refreshnicknames", description="Reapply nicknames for all users with a Discord ID (Admin only)")
async def refresh_nicknames_command(interaction: discord.Interaction):

    # Admin check
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message(
            "❌ You need administrator permissions to use this command!",
            ephemeral=True
        )
        return

    await interaction.response.defer(ephemeral=True)

    guild = interaction.guild
    guild_id = guild.id

    if guild_id not in spreadsheets:
        await interaction.followup.send(
            "❌ No spreadsheet connected for this server.",
            ephemeral=True
        )
        return

    try:
        spreadsheet = spreadsheets[guild_id]
        sheet = spreadsheet.worksheet(SHEET_PAGE_NAME)
        data = sheet.get_all_records()

        updated_count = 0
        skipped_count = 0

        for row in data:

            discord_identifier = str(row.get("Discord ID", "")).strip()
            if not discord_identifier:
                continue

            try:
                discord_id = int(discord_identifier)
            except ValueError:
                continue

            member = guild.get_member(discord_id)
            if not member:
                continue

            user_name = str(row.get("Name", "")).strip()
            # first_event = str(row.get("First Event", "")).strip()
            roles_raw = str(row.get("Roles", "")).strip()
            roles = [r.strip() for r in roles_raw.split(";") if r.strip()]
            first_event = roles[0] if roles else ""

            if not user_name or not first_event:
                skipped_count += 1
                continue

            # Construct nickname exactly like login_command
            nickname = f"{user_name} ({first_event})"

            if len(nickname) > 32:
                nickname = nickname[:32]

            try:
                await handle_rate_limit(
                    member.edit(nick=nickname, reason="Admin nickname refresh"),
                    f"editing nickname for {member}"
                )

                updated_count += 1
                print(f"📝 Updated nickname for {member} -> {nickname}")

            except discord.Forbidden:
                print(f"❌ No permission to change nickname for {member}")

            except Exception as e:
                print(f"⚠️ Error updating nickname for {member}: {e}")

        await interaction.followup.send(
            f"✅ Nickname refresh complete\n"
            f"• Updated: **{updated_count}**\n"
            f"• Skipped: **{skipped_count}**",
            ephemeral=True
        )

    except Exception as e:
        await interaction.followup.send(
            f"❌ Error refreshing nicknames: {e}",
            ephemeral=True
        )

@bot.tree.command(name="dummy4", description="Dummy 4 (Admin only)")
async def dummy4_command(interaction: discord.Interaction):
    """Dummy Command 4"""

    # Check if user has administrator permission
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return

@bot.tree.command(name="msg", description="Send a message as the bot (admin only)")
@app_commands.describe(
    message="The message to send",
    channel="Channel to send to (optional, defaults to current channel)"
)
async def msg_command(interaction: discord.Interaction, message: str, channel: discord.TextChannel = None):
    """Send a message as the bot - restricted to admin only"""

    # Defer immediately to prevent timeout
    await interaction.response.defer(ephemeral=True)

    # Check if user has the admin role
    sad_face_role = discord.utils.get(interaction.user.roles, name="Admin")
    if not sad_face_role:
        await interaction.followup.send("❌ You need a special role to use this command!", ephemeral=True)
        return

    # Use current channel if no channel specified
    target_channel = channel or interaction.channel

    # Check if bot has permission to send messages in the target channel
    if not target_channel.permissions_for(interaction.guild.me).send_messages:
        await interaction.followup.send(f"❌ I don't have permission to send messages in {target_channel.mention}!", ephemeral=True)
        return

    try:
        # Send the message to the target channel
        await target_channel.send(message)

        # Confirm to the user (privately)
        if target_channel == interaction.channel:
            await interaction.followup.send("✅ Message sent!", ephemeral=True)
        else:
            await interaction.followup.send(f"✅ Message sent to {target_channel.mention}!", ephemeral=True)

        # Log the action
        print(f"📢 {interaction.user} used /msg in {interaction.guild.name}: '{message}' → #{target_channel.name}")

    except discord.Forbidden:
        await interaction.followup.send(f"❌ I don't have permission to send messages in {target_channel.mention}!", ephemeral=True)
    except Exception as e:
        await interaction.followup.send(f"❌ Error sending message: {str(e)}", ephemeral=True)
        print(f"❌ Error in /msg command: {e}")

@bot.tree.command(name="rolereset", description="Reset the server - roles + nicknames (Admin only)")
async def role_reset_command(interaction: discord.Interaction):
    """Reset the roles and nicknames"""

    # Check if user has administrator permission
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return
    
    if admin_lock.locked():
        await interaction.response.send_message("❌ Server configurations are changing. Please try this when configurations is done!", ephemeral=True)
        return
    
    async with admin_lock:
        priority_roles = ["Admin", "Volunteer", "Lead ES", "Social Media", "Photographer", "Arbitrations", "Awards", "Runner", "VIPer"]

        # Defer immediately since this will take time
        await interaction.response.defer(ephemeral=True)

        try:
            guild = interaction.guild
            # Fetch all rows from this guild's sheet
            guild_id = guild.id

            print(f"🔍 DEBUG: Testing sheet access by reading data...")
            try:
                test_data = sheets[guild_id].get_all_records()
                print(f"✅ DEBUG: Successfully read {len(test_data)} rows from sheet")
            except Exception as e:
                print(f"❌ DEBUG: Error reading sheet data: {e}")
                print(f"❌ DEBUG: Error type: {type(e)}")
                print(f"❌ DEBUG: Error details: {str(e)}")
                raise e

            if guild:
                # Extract all unique chapters and events from the main sheet
                event_list = set()
                chapters = set()
                
                for row in test_data:
                    roles_raw = str(row.get("Roles", "")).strip()
                    roles = [r.strip() for r in roles_raw.split(";") if r.strip()]
                    first_event = roles[0] if roles else ""
                    chapter = str(row.get("Chapter", "")).strip()
                    
                    if first_event:
                        event_list.add(first_event)
                    
                    if chapter and chapter.lower() not in ["n/a", "na", ""]:
                        chapters.add(chapter)
                    else:
                        chapters.add("Unaffiliated")
                        
                # Ensure events from Room Assignments are also protected from deletion
                room_data = await get_room_assignments(guild_id)
                for r_row in room_data:
                    event = str(r_row.get("Events", "")).strip()
                    if event:
                        event_list.add(event)

                print(f"📖 Found {len(chapters)} unique chapters to preserve/create")
            else:
                print("⚠️ Could not get guild for structure creation")
        

        except Exception as e:
            print(f"❌ DEBUG: Error parsing sheet for build/rooms and chapter: {e}")
            raise e

        print(f"🔄 Starting role reset for {guild.name} requested by {interaction.user}")
                
        # Counters
        nickname_count = 0
        role_count = 0
        role_id = {}

        current_roles = {role.name for role in guild.roles}
        for role in guild.roles:
            role_id[role.name] = role.id
        new_roles = set(priority_roles + list(event_list) + list(chapters))
        common_roles = current_roles & new_roles
        delete_roles = (common_roles ^ current_roles) - set(priority_roles)
        add_roles = (common_roles ^ new_roles) - set(priority_roles)

        print(f"🔄 DEBUG: current_roles = {current_roles}")
        print(f"🔄 DEBUG: new_roles = {new_roles}")
        print(f"🔄 DEBUG: common_roles = {common_roles}")
        print(f"🔄 DEBUG: delete_roles = {delete_roles}")
        print(f"🔄 DEBUG: add_roles = {add_roles}")

        try:
            # Delete custom roles (keep @everyone and bot roles)
            print("🗑️ Deleting custom roles...")
            for role_str in delete_roles:
                role = guild.get_role(role_id[role_str])
                # Skip @everyone, bot roles, and roles higher than bot's highest role
                if (role.name != "@everyone" and
                    not role.managed and
                    role < guild.me.top_role):
                    try:
                        await safe_call(role.delete(reason=f"Role reset by {interaction.user}"))
                        role_count += 1
                        print(f"🗑️ Deleted role: {role.name}")
                    except discord.Forbidden:
                        print(f"❌ No permission to delete role {role.name}")
                    except Exception as e:
                        print(f"⚠️ Error deleting role {role.name}: {e}")

            # Pre-create all building structures and channels from the sheet data
            print("🏗️ Pre-creating all building structures and channels...")
            # Pre-create all building structures and channels from the sheet data
            print("🏗️ Pre-creating all building structures and channels...")
            try:
                # Call the standalone generation function, forcing the welcome messages to reset
                await generate_building_structures(guild, force_refresh_welcome=True)
                
                # Create all chapter structures upfront
                for chapter in chapters:
                    print(f"📖 Pre-creating chapter: {chapter}")
                    await setup_chapter_structure(guild, chapter)
                    
                # Sort chapter channels alphabetically
                print("📖 Organizing chapter channels alphabetically...")
                await sort_chapter_channels_alphabetically(guild)
                
            except Exception as structure_error:
                print(f"⚠️ Error creating building structures: {structure_error}")
                # Don't fail the whole command if structure creation fails
            
            # Check if ezhang. is already in this server and give them the Admin role
            await setup_ezhang_admin_role(guild)

            # Trigger an immediate sync after successful connection and structure creation
            print("🔄 Triggering immediate sync after template connection...")
            sync_results = None
            try:
                guild = interaction.guild
                if guild:
                    sync_results = await perform_member_sync(guild, test_data)
                    print(f"✅ Initial sync complete: {sync_results['processed']} processed, {sync_results['invited']} invited, {sync_results['role_assignments']} roles assigned")
                else:
                    print("⚠️ Could not get guild for immediate sync")
            except Exception as sync_error:
                print(f"⚠️ Error during immediate sync: {sync_error}")
                # Don't fail the whole command if sync fails
                ############

            result_embed = discord.Embed(
                title="✅ Role Reset Complete",
                description="The server has completely reset roles and assignments!",
                color=discord.Color.green()
            )
            result_embed.add_field(name="Nicknames Reset", value=str(nickname_count), inline=True)
            result_embed.add_field(name="Roles Deleted", value=str(role_count), inline=True)
            result_embed.set_footer(text="🏗️ Roles and nicknames have been refreshed!")

            # Try to send to user via DM first
            sent_dm = False
            try:
                await interaction.user.send(embed=result_embed)
                print(f"✅ Sent completion message to {interaction.user} via DM")
                await interaction.followup.send(embed=result_embed, ephemeral=True)
                print(f"✅ Successfully reset roles")
                sent_dm = True
            except:
                print(f"⚠️ Could not send completion message to {interaction.user} via DM")

            print(f"📊 Summary:")
            print(f"   • {nickname_count} nicknames reset")
            print(f"   • {role_count} roles deleted")

        except Exception as e:
            error_msg = f"❌ Error during role reset: {str(e)}"
            try:
                await interaction.followup.send(error_msg)
            except:
                # If followup fails, try to DM the user
                try:
                    await interaction.user.send(error_msg)
                except:
                    print(error_msg)
            print(f"❌ Role reset error: {e}")
            import traceback
            traceback.print_exc()

@bot.tree.command(name="resetserver", description="⚠️ DANGER: Completely reset the server - deletes channels, roles, categories (Admin only)")
async def reset_server_command(interaction: discord.Interaction):
    """⚠️ DANGER: Completely reset the server by deleting all channels, categories, roles, and nicknames"""

    global reset_active

    # Check if user has administrator permission
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ You need administrator permissions to use this command!", ephemeral=True)
        return
    
    if admin_lock.locked():
        await interaction.response.send_message("❌ Server configurations are changing. Please try this when configurations is done!", ephemeral=True)
        return
    
    async with admin_lock:
        reset_active = True

        # Defer immediately since this will take time
        await interaction.response.defer(ephemeral=True)

        try:
            guild = interaction.guild
            print(f"🔄 Starting complete server reset for {guild.name} requested by {interaction.user}")

            # Counters
            nickname_count = 0
            channel_count = 0
            voice_count = 0
            forum_count = 0
            category_count = 0
            role_count = 0

            # Reset all member nicknames
            print("📝 Resetting all member nicknames...")
            for member in guild.members:
                if member.nick and not member.bot:
                    try:
                        await handle_rate_limit(
                            member.edit(nick=None, reason=f"Server reset by {interaction.user}"),
                            f"resetting nickname for {member}"
                        )
                        nickname_count += 1
                        print(f"📝 Reset nickname for {member.display_name}")
                    except discord.Forbidden:
                        print(f"❌ No permission to reset nickname for {member.display_name}")
                    except Exception as e:
                        print(f"⚠️ Error resetting nickname for {member.display_name}: {e}")
            print(f"✅ Reset {nickname_count} nicknames")

            # Delete all text channels
            print("🗑️ Deleting all text channels...")
            for channel in guild.text_channels:
                try:
                    await safe_call(channel.delete(reason=f"Server reset by {interaction.user}"))
                    channel_count += 1
                    print(f"🗑️ Deleted text channel: #{channel.name}")
                except discord.Forbidden:
                    print(f"❌ No permission to delete channel #{channel.name}")
                except Exception as e:
                    print(f"⚠️ Error deleting channel #{channel.name}: {e}")

            # Delete all voice channels
            print("🗑️ Deleting all voice channels...")
            for channel in guild.voice_channels:
                try:
                    await safe_call(channel.delete(reason=f"Server reset by {interaction.user}"))
                    voice_count += 1
                    print(f"🗑️ Deleted voice channel: {channel.name}")
                except discord.Forbidden:
                    print(f"❌ No permission to delete voice channel {channel.name}")
                except Exception as e:
                    print(f"⚠️ Error deleting voice channel {channel.name}: {e}")

            # Delete all forum channels
            print("🗑️ Deleting all forum channels...")
            for channel in guild.channels:
                if hasattr(channel, 'type') and channel.type == discord.ChannelType.forum:
                    try:
                        await safe_call(channel.delete(reason=f"Server reset by {interaction.user}"))
                        forum_count += 1
                        print(f"🗑️ Deleted forum channel: #{channel.name}")
                    except discord.Forbidden:
                        print(f"❌ No permission to delete forum #{channel.name}")
                    except Exception as e:
                        print(f"⚠️ Error deleting forum #{channel.name}: {e}")

            # Delete all categories
            print("🗑️ Deleting all categories...")
            for category in guild.categories:
                try:
                    await safe_call(category.delete(reason=f"Server reset by {interaction.user}"))
                    category_count += 1
                    print(f"🗑️ Deleted category: {category.name}")
                except discord.Forbidden:
                    print(f"❌ No permission to delete category {category.name}")
                except Exception as e:
                    print(f"⚠️ Error deleting category {category.name}: {e}")

            # Delete all custom roles (keep @everyone and bot roles)
            print("🗑️ Deleting all custom roles...")
            for role in guild.roles:
                # Skip @everyone, bot roles, and roles higher than bot's highest role
                if (role.name != "@everyone" and
                    not role.managed and
                    role < guild.me.top_role):
                    try:
                        await safe_call(role.delete(reason=f"Server reset by {interaction.user}"))
                        role_count += 1
                        print(f"🗑️ Deleted role: {role.name}")
                    except discord.Forbidden:
                        print(f"❌ No permission to delete role {role.name}")
                    except Exception as e:
                        print(f"⚠️ Error deleting role {role.name}: {e}")

            # Set up static channels after everything is deleted
            print("🏗️ Setting up static channels...")
            welcome_channel = None
            try:
                await setup_static_channels_for_guild(guild)
                # Get the welcome channel that was just created
                welcome_channel = discord.utils.get(guild.text_channels, name="welcome")
                print(f"✅ Static channels setup complete")
            except Exception as e:
                print(f"⚠️ Error setting up static channels: {e}")

            # Check if ezhang. is already in this server and give them the Admin role
            await setup_ezhang_admin_role(guild)

            # Send completion message
            print("🧨 SERVER RESET COMPLETE!")
            result_embed = discord.Embed(
                title="✅ Server Reset Complete",
                description="The server has been completely reset!",
                color=discord.Color.green()
            )
            result_embed.add_field(name="Nicknames Reset", value=str(nickname_count), inline=True)
            result_embed.add_field(name="Text Channels Deleted", value=str(channel_count), inline=True)
            result_embed.add_field(name="Voice Channels Deleted", value=str(voice_count), inline=True)
            result_embed.add_field(name="Forum Channels Deleted", value=str(forum_count), inline=True)
            result_embed.add_field(name="Categories Deleted", value=str(category_count), inline=True)
            result_embed.add_field(name="Roles Deleted", value=str(role_count), inline=True)
            result_embed.set_footer(text="🏗️ Server is now completely clean and ready for fresh setup!")

            # Try to send to user via DM first
            sent_dm = False
            try:
                await interaction.user.send(embed=result_embed)
                print(f"✅ Sent completion message to {interaction.user} via DM")
                sent_dm = True
            except:
                print(f"⚠️ Could not send completion message to {interaction.user} via DM")

            # If DM failed and we created a welcome channel, send there
            if not sent_dm and welcome_channel:
                try:
                    await welcome_channel.send(f"{interaction.user.mention}", embed=result_embed)
                    print(f"✅ Sent completion message to #{welcome_channel.name}")
                except Exception as e:
                    print(f"⚠️ Could not send completion message to welcome channel: {e}")

            print(f"📊 Summary:")
            print(f"   • {nickname_count} nicknames reset")
            print(f"   • {channel_count} text channels deleted")
            print(f"   • {voice_count} voice channels deleted")
            print(f"   • {forum_count} forum channels deleted")
            print(f"   • {category_count} categories deleted")
            print(f"   • {role_count} roles deleted")

        except Exception as e:
            error_msg = f"❌ Error during server reset: {str(e)}"
            try:
                await interaction.followup.send(error_msg)
            except:
                # If followup fails, try to DM the user
                try:
                    await interaction.user.send(error_msg)
                except:
                    print(error_msg)
            print(f"❌ Server reset error: {e}")
            import traceback
            traceback.print_exc()

@bot.tree.interaction_check
async def block_commands_during_reset(interaction: discord.Interaction) -> bool:
    global reset_active

    if not reset_active:
        return True

    # Allow only enterfolder during reset
    if interaction.command and interaction.command.name in ALLOWED_DURING_RESET:
        return True

    await interaction.response.send_message(
        "🚧 Server reset in progress or done.\n"
        "Only `/enterfolder` is available right now.",
        ephemeral=True
    )
    return False

if __name__ == "__main__":
    # Start a simple HTTP health check server for Fly.io
    # This prevents auto-stop since Fly.io sees the machine as "active"
    from http.server import HTTPServer, BaseHTTPRequestHandler
    import threading

    class HealthCheckHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            """Simple health check endpoint"""
            self.send_response(200)
            self.send_header('Content-type', 'text/plain')
            self.end_headers()
            self.wfile.write(b'Bot is running!')

        def log_message(self, format, *args):
            """Suppress HTTP logs to keep console clean"""
            pass

    # Start health check server in background thread
    def run_health_server():
        server = HTTPServer(('0.0.0.0', 8080), HealthCheckHandler)
        print("🏥 Health check server running on port 8080")
        server.serve_forever()

    health_thread = threading.Thread(target=run_health_server, daemon=True)
    health_thread.start()

    # Run the Discord bot
    bot.run(TOKEN)