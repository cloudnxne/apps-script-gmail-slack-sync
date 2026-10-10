# Gmail sync and Slack sync (Google Apps Script)

Two internal web apps built for a team of 50+ people.

## gmail-sync
A multi-user web app. Each person owns their clients, their trigger and their Gmail. It exports email threads by label into per-client Google Sheets on a daily trigger.
- De-duplicates on message ID.
- Captures every message in a thread, not only the last.
- Trims signature markers without cutting off reply bodies.
- Includes a self-serve sheet recovery flow.

## slack-sync
A Slack receiver plus a user dashboard. It saves Slack messages to shared Sheets from a message shortcut or an emoji reaction.
- Acknowledges Slack inside its 3 second limit, queues the work and processes it behind a script lock.
- Resolves Slack user names and builds message permalinks.

## Setup
1. Create an Apps Script project and paste in the `.gs` files.
2. Add secrets under Project Settings, Script Properties. For Slack: `SLACK_BOT_TOKEN` (scopes `users:read`, `channels:history`, `reactions:read`).
3. Set `SHARED_FOLDER_ID` in `user_dashboard.gs` to your Drive folder.
4. Keep `appsscript.json` as the manifest, then review the web app access settings before deploying.
5. Point the Slack app's interactivity and event subscription URLs at the receiver deployment.
