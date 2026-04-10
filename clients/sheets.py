"""
clients/sheets.py — Google Sheets and Drive service client.

Handles credential setup, spreadsheet access, and the per-guild
connection cache. Pure I/O — no Discord logic here.

Usage:
    client = SheetsClient()          # initializes credentials
    ss = client.open_by_key(sheet_id)
    sheet = ss.worksheet("lambot")

The bot stores one shared SheetsClient on bot.sheets_client and the
per-guild spreadsheet objects in bot.spreadsheets / bot.sheets.
"""

import json
import os

import gspread
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

import config

_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

_CREDENTIALS_PATH = "secrets/gspread.json"


class SheetsClient:
    """Thin wrapper around gspread + Drive API.

    Instantiate once and attach to the bot:
        bot.sheets_client = SheetsClient()
    """

    def __init__(self) -> None:
        with open(_CREDENTIALS_PATH) as f:
            keyfile = json.load(f)
        self.creds = ServiceAccountCredentials.from_json_keyfile_dict(keyfile, _SCOPES)
        self.gc = gspread.authorize(self.creds)

    # ── Spreadsheet access ────────────────────────────────────────────────────

    def open_by_key(self, sheet_id: str) -> gspread.Spreadsheet:
        return self.gc.open_by_key(sheet_id)

    def open_by_title(self, title: str) -> gspread.Spreadsheet:
        return self.gc.open(title)

    def find_sheet_in_folder(self, folder_id: str, name: str) -> gspread.Spreadsheet | None:
        """Search a Drive folder for a spreadsheet whose title contains `name`.

        Returns the opened Spreadsheet or None if not found.
        Falls back to a global title search if opening by ID fails.
        """
        drive = build("drive", "v3", credentials=self.creds)
        query = (
            f"'{folder_id}' in parents"
            " and mimeType='application/vnd.google-apps.spreadsheet'"
            f" and name contains '{name}'"
        )
        results = drive.files().list(q=query, fields="files(id, name)", pageSize=10).execute()
        files = results.get("files", [])

        for file in files:
            if name in file["name"]:
                try:
                    return self.gc.open_by_key(file["id"])
                except Exception as e:
                    print(f"Could not open sheet by ID {file['id']}: {e}")
                    # Fall through to global search below

        # Global fallback
        try:
            return self.gc.open(name)
        except gspread.SpreadsheetNotFound:
            return None

    # ── Per-guild connection cache ────────────────────────────────────────────

    @staticmethod
    def _load_raw_cache() -> dict:
        if os.path.exists(config.CACHE_FILE):
            with open(config.CACHE_FILE) as f:
                return json.load(f)
        return {}

    @staticmethod
    def _save_raw_cache(data: dict) -> None:
        with open(config.CACHE_FILE, "w") as f:
            json.dump(data, f, indent=2)

    def load_spreadsheets_from_cache(self) -> dict[int, gspread.Spreadsheet]:
        """Re-open every cached guild spreadsheet connection.

        Returns a dict of guild_id (int) -> gspread.Spreadsheet for
        connections that succeeded. Failures are logged and skipped.
        """
        cache = self._load_raw_cache()
        guilds_cache = cache.get("guilds", {})
        result: dict[int, gspread.Spreadsheet] = {}

        for guild_id_str, guild_cache in guilds_cache.items():
            sheet_id = guild_cache.get("spreadsheet_id")
            worksheet_name = guild_cache.get("worksheet_name", config.SHEET_PAGE_NAME)
            if not sheet_id:
                continue
            try:
                ss = self.gc.open_by_key(sheet_id)
                # Validate access — this will raise if the sheet is gone/inaccessible
                ss.worksheet(worksheet_name).row_values(1)
                result[int(guild_id_str)] = ss
                print(f"Restored sheet connection for guild {guild_id_str}: '{ss.title}'")
            except Exception as e:
                print(f"Could not restore sheet for guild {guild_id_str}: {e}")

        return result

    def save_guild_to_cache(self, guild_id: int, spreadsheet_id: str, worksheet_name: str) -> None:
        cache = self._load_raw_cache()
        cache.setdefault("guilds", {})[str(guild_id)] = {
            "spreadsheet_id": spreadsheet_id,
            "worksheet_name": worksheet_name,
        }
        self._save_raw_cache(cache)
        print(f"Saved sheet connection to cache for guild {guild_id}")

    def clear_guild_from_cache(self, guild_id: int) -> bool:
        """Remove a guild's entry from the cache. Returns True if anything was removed."""
        cache = self._load_raw_cache()
        changed = False
        gid = str(guild_id)

        if gid in cache.get("guilds", {}):
            del cache["guilds"][gid]
            changed = True

        runner = cache.get("runner_access_settings", {})
        if gid in runner:
            del runner[gid]
            cache["runner_access_settings"] = runner
            changed = True

        if changed:
            self._save_raw_cache(cache)
        return changed

    def save_runner_access_to_cache(self, runner_all_access: dict) -> None:
        """Persist the runner_all_access mapping (guild_id -> flag) to cache."""
        cache = self._load_raw_cache()
        cache["runner_access_settings"] = {str(k): v for k, v in runner_all_access.items()}
        self._save_raw_cache(cache)

    def load_runner_access_from_cache(self) -> dict[int, int]:
        """Return the saved runner_all_access mapping (guild_id -> flag)."""
        cache = self._load_raw_cache()
        return {int(k): v for k, v in cache.get("runner_access_settings", {}).items()}
