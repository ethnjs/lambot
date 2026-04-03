🤖 LamBot User Guide

LamBot is a specialized Discord bot designed to automate and manage Science Olympiad tournaments. It handles role assignments, channel creation, test material distribution, help tickets, and runner zoning based on your Google Sheets data.

This guide covers how to add LamBot to your server and operate it as a Tournament Administrator.

🚀 Phase 1: Initial Setup
Step 1: Check if LamBot is running

Before starting, ensure the bot is actively hosted and online. Check with Peter Hung to confirm the bot is currently running.

Step 2: Add LamBot to your server

Use the following link to invite LamBot to your Discord server. You must have "Manage Server" or Administrator permissions to do this:
Click here to invite LamBot

Step 3: Configure Role Hierarchy and Get the Template

For LamBot to properly assign roles and change user nicknames, its role must be at the top of the server's hierarchy.

In your server, go to Server Settings > Roles.

Find the LamBot role and click/drag it to the absolute top of the role list.

Save changes.

In any text channel, type /gettemplate. The bot will reply with a link to the template Google Drive folder.

📁 Phase 2: Preparing Your Data
Step 4: Copy and Fill Out the Template

Duplicate the provided Google Drive Template folder into your own Google Drive.
Fill out your data across the provided spreadsheets.

Important Data Rules:

lambot (Main Sheet): Used for user accounts. The Roles column must be separated by semicolons (e.g., Astronomy; Anatomy; Chem Lab).

Room Assignments: Used for generating event channels. Must have columns named Events, Building, and Room.

Runner Assignments: Used for runner zoning. Must include runner emails, building names, and coordinates.

Step 5: Give LamBot Access to Your Folder

LamBot needs permission to read your Google Drive folder.

Type /serviceaccount in your Discord server. The bot will reply with an email address.

Go to your Google Drive.

Right-click your Tournament Folder > Share.

Paste the bot's service email and grant it Editor permissions.

Click Copy link (Do not copy the URL from your browser's address bar).

🔌 Phase 3: Connecting and Generating the Server
Step 6: Link the Folder

In your Discord server, run:
/enterfolder folder_link:<paste_your_copied_link> main_sheet_name:<name_of_your_main_sheet>

What happens next?
LamBot will read your data and automatically:

Create static categories (Welcome, Tournament Officials, Volunteers, Chapters).

Generate specific role colors.

Create building categories and individual event channels based on the Room Assignments sheet.

Post welcome messages in the building chats.

Run an initial user sync.

Step 7: Organize the Roles

Now that the bot has created dozens of roles, they need to be sorted so Discord permissions work properly.
Run:
/organizeroles
(This sorts priority roles like Admin, VIPer, and Runner to the top, followed by Chapter roles, and then event roles).

🏃 Phase 4: Tournament Operations
Step 8: Assign Runner Zones

LamBot uses an AI clustering algorithm (K-means) to automatically group buildings into physical zones and assign your runners to them.
Run:
/assignrunnerzones
The bot will calculate the zones based on coordinates in your sheet, update the Google Sheet with the zone numbers, and post messages in the Building Chats telling people which runners are assigned to their zone.

Step 9: Distribute Test Materials & Links

Ensure your test materials (PDFs, docs) are placed inside the Tests/[Event Name] folders within your Google Drive.
Run:
/sendallmaterials
The bot will scan your Drive folder and automatically post and pin the relevant test materials to every single event channel. It will also post in #useful-links and #runner.

👤 Phase 5: User Onboarding

Tell your volunteers, Event Supervisors, and competitors to join the server.

To get their roles and access their hidden channels, they simply need to type:
/login email:their_email@example.com password:their_password

LamBot will automatically:

Verify their credentials against the Google Sheet.

Save their Discord ID.

Change their server nickname (e.g., John Doe (Astronomy)).

Assign them their Event, Master, and Chapter roles.

Unlock their specific building and event channels.

🛠️ Command Reference
Admin Commands
Command	Description
/gettemplate	Get the template Drive folder link.
/serviceaccount	View the bot's email address to share your Google Drive with.
/enterfolder	Connect the bot to your Google Drive folder and build the server infrastructure.
/sync	Manually pull the latest user data from the Google Sheet (auto-runs every 60 mins).
/syncrooms	Re-read the Room Assignments sheet to create missing event channels or refresh building welcome messages.
/organizeroles	Automatically sort the Discord role hierarchy.
/assignrunnerzones	Calculate and assign runner zones based on building coordinates.
/sendallmaterials	Distribute all test files and useful links to their specific channels.
/refreshnicknames	Force-update all user nicknames based on the current Google Sheet data.
/debugrooms	View exactly how the bot is reading your Room Assignments tab to troubleshoot errors.
/debugzone @user	Test and view which zone/runners are assigned to a specific user.
/activetickets	View all active help forum threads currently being tracked by the bot.
/msg	Send a message as the bot to a specific channel.
/rolereset	Safely delete all custom roles/nicknames and rebuild them from the sheet.
/resetserver	⚠️ DANGER: Deletes ALL channels, roles, and categories. Wipes the server clean.
User Commands
Command	Description
/login	Authenticate with an email and password to receive roles and channel access.
/help	View basic bot information and command usage.
🎫 How Help Tickets Work

If a user asks a question in the #help forum channel:

LamBot detects which event/building the user belongs to.

It looks up the physical zone of that building.

It automatically tags the specific Runners assigned to that zone.

If a runner does not reply or react (e.g., 👍) within 3 minutes, LamBot will ping them again.

If they still don't reply after another minute, LamBot escalates and pings ALL runners in the server.