/**
 * USER DASHBOARD — user_dashboard.gs
 * ============================================================
 * DEPLOYMENT SETTINGS:
 *   Description:    User Dashboard
 *   Execute as:     User accessing the web app
 *   Who has access: Anyone with Google account
 * ============================================================
 */

const SHARED_FOLDER_ID = "YOUR_SHARED_DRIVE_FOLDER_ID"; // Drive folder where the shared Sheets live

// ============================================================
// SLACK HELPERS
// ============================================================

// Resolve a Slack user's @username (used for queue matching)
function getSlackUsername(userId) {
  try {
    const token = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
    if (!token || !userId || userId === "Unknown") return userId;

    const response = UrlFetchApp.fetch("https://slack.com/api/users.info?user=" + userId, {
      headers: { "Authorization": "Bearer " + token }
    });
    const data = JSON.parse(response.getContentText());
    if (data.ok && data.user) {
      return (data.user.name || userId).toLowerCase().trim();
    }
  } catch (err) {
    Logger.log("Could not get username for " + userId + ": " + err.toString());
  }
  return userId;
}

// Fetch a specific message from Slack by channel ID and timestamp
function fetchSlackMessage(channelId, messageTs, teamId) {
  try {
    const token = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
    if (!token) return null;

    const url      = "https://slack.com/api/conversations.history?channel=" + channelId
                   + "&latest=" + messageTs + "&inclusive=true&limit=1";
    const response = UrlFetchApp.fetch(url, {
      headers: { "Authorization": "Bearer " + token }
    });
    const data = JSON.parse(response.getContentText());

    if (!data.ok || !data.messages || data.messages.length === 0) {
      Logger.log("Could not fetch message: " + JSON.stringify(data));
      return null;
    }

    const msg          = data.messages[0];
    const messageText  = (msg.text || "").replace(/\+/g, " ");
    const senderUserId = msg.user || "Unknown";
    const timestamp    = msg.ts
      ? new Date(parseFloat(msg.ts) * 1000).toISOString()
      : new Date().toISOString();

    // Fetch channel name
    let channelName = channelId;
    try {
      const chanResp = UrlFetchApp.fetch("https://slack.com/api/conversations.info?channel=" + channelId, {
        headers: { "Authorization": "Bearer " + token }
      });
      const chanData = JSON.parse(chanResp.getContentText());
      if (chanData.ok && chanData.channel) channelName = chanData.channel.name || channelId;
    } catch (err) {
      Logger.log("Could not fetch channel name: " + err.toString());
    }

    const messageLink = teamId
      ? "https://slack.com/archives/" + channelId + "/p" + messageTs.replace(".", "")
      : "";

    return { channelName, senderUserId, messageText, timestamp, messageLink };

  } catch (err) {
    Logger.log("fetchSlackMessage error: " + err.toString());
    return null;
  }
}

function getSlackDisplayName(userId) {
  try {
    const token = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
    if (!token || !userId || userId === "Unknown") return userId;

    const response = UrlFetchApp.fetch("https://slack.com/api/users.info?user=" + userId, {
      headers: { "Authorization": "Bearer " + token }
    });
    const data = JSON.parse(response.getContentText());
    if (data.ok && data.user) {
      const profile = data.user.profile;
      const name = profile.display_name_normalized
        || profile.display_name
        || profile.real_name_normalized
        || profile.real_name
        || data.user.name
        || userId;
      Logger.log("Resolved " + userId + " → " + name);
      return name;
    } else {
      Logger.log("Slack API error for " + userId + ": " + JSON.stringify(data));
    }
  } catch (err) {
    Logger.log("Could not look up user " + userId + ": " + err.toString());
  }
  return userId;
}

// Forces Google to request all required OAuth scopes upfront
function _forceScopes() {
  ScriptApp.getProjectTriggers();
  DriveApp.getRootFolder();
  SpreadsheetApp.create("_test");
}

// ============================================================
// WEB UI
// ============================================================

function doGet() {
  const userProps  = PropertiesService.getUserProperties();
  const authorised = userProps.getProperty("AUTHORISED");
  const email      = Session.getActiveUser().getEmail();
  const slackUser  = userProps.getProperty("SLACK_USERNAME") || "";
  let registry     = JSON.parse(userProps.getProperty("CHANNEL_REGISTRY") || "[]");

  // ── Auto-recover registry from Script Properties if empty ──
  if (registry.length === 0 && authorised) {
    const scriptProps = PropertiesService.getScriptProperties();
    const allProps    = scriptProps.getProperties();
    Object.keys(allProps)
      .filter(k => k.startsWith("channel_id:"))
      .forEach(k => {
        const channelName = k.replace("channel_id:", "");
        const fileId      = allProps[k];
        try {
          const ss = SpreadsheetApp.openById(fileId);
          registry.push({ channel: channelName, id: fileId, url: ss.getUrl() });
        } catch(e) {
          Logger.log("Auto-recover failed for #" + channelName + ": " + e.toString());
        }
      });
    if (registry.length > 0) {
      userProps.setProperty("CHANNEL_REGISTRY", JSON.stringify(registry));
      Logger.log("Auto-recovered " + registry.length + " channel(s) into registry.");
    }
  }

  // Show any pending queue entries for this user
  const scriptProps = PropertiesService.getScriptProperties();
  const allProps    = scriptProps.getProperties();
  const myQueue     = Object.keys(allProps)
    .filter(k => k.startsWith("queue_"))
    .map(k => {
      try { return JSON.parse(allProps[k]); } catch(e) { return null; }
    })
    .filter(e => e && (e.slackUsername || "").toLowerCase() === slackUser.toLowerCase());

  const queueHtml = myQueue.length > 0 ? `
    <div style="background:#fff8e1; border-radius:8px; padding:12px 16px; margin-bottom:16px;
                border-left:4px solid #f9a825;">
      <strong style="color:#f57f17; font-size:0.85em;">
        ${myQueue.length} message(s) queued — click Process Now to save immediately
      </strong>
    </div>` : "";

  let channelListHtml = registry.map((c, index) => `
    <div style="background:#f8f9fa; padding:12px; border-radius:8px; margin-bottom:10px;
                border-left:4px solid #4A154B; display:flex; justify-content:space-between; align-items:center;">
      <div>
        <strong style="color:#333; display:block;">#${c.channel}</strong>
        <span style="font-size:0.8em; color:#888;">Saved to Slack Messages folder</span>
      </div>
      <div style="display:flex; align-items:center; gap:10px;">
        <a href="${c.url}" target="_blank"
           style="color:#4A154B; font-weight:bold; font-size:0.85em; text-decoration:none; white-space:nowrap;">
          OPEN ↗
        </a>
        <span onclick="removeSheet(${index})"
              style="color:#ccc; cursor:pointer; font-size:1.4em; font-weight:bold; line-height:1;"
              title="Remove from dashboard">&times;</span>
      </div>
    </div>
  `).join("");

  if (registry.length === 0) {
    channelListHtml = `
      <div style="text-align:center; padding:20px; color:#999; border:1px dashed #ddd; border-radius:8px;">
        No sheets yet — save a Slack message to get started.
      </div>`;
  }

  const authorisePanel = !authorised ? `
    <div style="background:#f0ebf4; border-radius:8px; padding:20px; margin-bottom:24px;">
      <p style="color:#4A154B; font-weight:bold; margin:0 0 4px;">One-time setup required</p>
      <p style="color:#666; font-size:0.85em; margin:0 0 12px;">
        Enter your Slack username exactly as it appears in Slack
        (e.g. <strong>first.last</strong>) and click Authorise.
      </p>
      <input type="text" id="slackUser" placeholder="your.name"
             style="width:100%; padding:10px; margin-bottom:12px; border:1px solid #ddd;
                    border-radius:6px; box-sizing:border-box; font-size:0.9em;">
      <button onclick="authorise()"
              style="background:#4A154B; color:white; border:none; padding:12px 28px;
                     border-radius:6px; font-weight:bold; cursor:pointer; font-size:1em; width:100%;">
        Authorise Access
      </button>
    </div>` : `
    <div style="background:#e8f5e9; border-radius:8px; padding:12px 16px; margin-bottom:16px;
                display:flex; align-items:center; gap:10px;">
      <span style="font-size:1.4em;">✅</span>
      <div style="flex-grow:1;">
        <strong style="color:#2e7d32;">Authorised</strong>
        <span style="display:block; font-size:0.82em; color:#555;">
          Slack username: <strong style="color:#333;">${slackUser || "(not set — click Change!)"}</strong>
        </span>
      </div>
      <button onclick="reauthorise()"
              style="background:none; border:1px solid #ccc; color:#666; padding:6px 12px;
                     border-radius:6px; font-size:0.78em; cursor:pointer; white-space:nowrap;">
        Change
      </button>
    </div>
    ${queueHtml}
    <div style="text-align:center; margin-bottom:20px; display:flex; gap:8px; justify-content:center;">
      <button onclick="processNow()"
              style="background:#4A154B; color:white; border:none; padding:10px 24px;
                     border-radius:6px; font-size:0.85em; font-weight:bold; cursor:pointer;">
        ▶ Process Now
      </button>
      <button onclick="recoverSheets()"
              style="background:none; border:1px solid #4A154B; color:#4A154B; padding:10px 24px;
                     border-radius:6px; font-size:0.85em; font-weight:bold; cursor:pointer;">
        🔄 Recover Sheets
      </button>
    </div>`;

  const html = `
    <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
                padding:28px; max-width:480px; margin:24px auto;
                border:1px solid #eee; border-radius:14px; background:#fff;
                box-shadow:0 8px 24px rgba(0,0,0,0.06);">

      <div style="text-align:center; margin-bottom:24px;">
        <div style="font-size:2.2em; margin-bottom:6px;">💬 → 📊</div>
        <h2 style="color:#4A154B; margin:0 0 4px;">Slack → Sheets</h2>
        <p style="color:#aaa; font-size:0.82em; margin:0;">
          Signed in as <strong style="color:#666;">${email}</strong>
        </p>
      </div>

      ${authorisePanel}

      <div>
        <h4 style="margin:0 0 12px; color:#555; text-transform:uppercase;
                   font-size:0.72em; letter-spacing:1px;">Your Channel Sheets</h4>
        ${channelListHtml}
      </div>

      <div id="status" style="margin-top:16px; text-align:center; font-size:0.85em;
                               color:#4A154B; font-weight:bold; min-height:20px;"></div>
    </div>

    <script>
      function authorise() {
        const u = document.getElementById("slackUser").value.trim();
        if (!u) return alert("Please enter your Slack username");
        document.getElementById("status").innerText = "Authorising...";
        google.script.run
          .withSuccessHandler(() => location.reload())
          .withFailureHandler(e => {
            document.getElementById("status").innerText = "Error: " + e.message;
          })
          .authoriseUser(u);
      }
      function reauthorise() {
        document.getElementById("status").innerText = "Resetting...";
        google.script.run
          .withSuccessHandler(() => location.reload())
          .withFailureHandler(e => {
            document.getElementById("status").innerText = "Error: " + e.message;
          })
          .resetAuth();
      }
      function processNow() {
        document.getElementById("status").innerText = "Processing...";
        google.script.run
          .withSuccessHandler(() => {
            document.getElementById("status").innerText = "Done! Reloading...";
            setTimeout(() => location.reload(), 2000);
          })
          .withFailureHandler(e => {
            document.getElementById("status").innerText = "Error: " + e.message;
          })
          .processMyQueue();
      }
      function removeSheet(index) {
        if (!confirm("Remove this sheet from your dashboard? The file will not be deleted.")) return;
        document.getElementById("status").innerText = "Removing...";
        google.script.run
          .withSuccessHandler(() => location.reload())
          .withFailureHandler(e => {
            document.getElementById("status").innerText = "Error: " + e.message;
          })
          .removeSheetFromDashboard(index);
      }
      function recoverSheets() {
        document.getElementById("status").innerText = "Recovering...";
        google.script.run
          .withSuccessHandler(count => {
            document.getElementById("status").innerText =
              count > 0 ? "Recovered " + count + " sheet(s)! Reloading..." : "Nothing new to recover.";
            if (count > 0) setTimeout(() => location.reload(), 2000);
          })
          .withFailureHandler(e => {
            document.getElementById("status").innerText = "Error: " + e.message;
          })
          .recoverSheets();
      }
    </script>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle("Slack → Sheets")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// AUTHORISATION
// ============================================================

function resetAuth() {
  const userProps = PropertiesService.getUserProperties();
  userProps.deleteProperty("AUTHORISED");
  userProps.deleteProperty("SLACK_USERNAME");
}

function authoriseUser(slackUsername) {
  const userProps = PropertiesService.getUserProperties();
  const cleaned   = slackUsername.toLowerCase().trim();

  userProps.setProperty("AUTHORISED", "true");
  userProps.setProperty("SLACK_USERNAME", cleaned);

  // Proactively touch all services so Google prompts for permissions now
  // rather than failing silently later when processMyQueue runs
  try { DriveApp.getRootFolder(); } catch(e) {}
  try { SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  try { ScriptApp.getProjectTriggers(); } catch(e) {}
  try { UrlFetchApp.fetch("https://www.google.com", { muteHttpExceptions: true }); } catch(e) {}

  // Remove any existing processMyQueue triggers
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "processMyQueue")
    .forEach(t => ScriptApp.deleteTrigger(t));

  Logger.log("Authorised user with Slack username: " + cleaned);
}

// ============================================================
// REMOVE SHEET FROM DASHBOARD (keeps the file)
// ============================================================

function removeSheetFromDashboard(index) {
  const userProps = PropertiesService.getUserProperties();
  const registry  = JSON.parse(userProps.getProperty("CHANNEL_REGISTRY") || "[]");
  registry.splice(index, 1);
  userProps.setProperty("CHANNEL_REGISTRY", JSON.stringify(registry));
}

// ============================================================
// RECOVER SHEETS — rebuild registry from Script Properties
// ============================================================

function recoverSheets() {
  const userProps   = PropertiesService.getUserProperties();
  const scriptProps = PropertiesService.getScriptProperties();
  const allProps    = scriptProps.getProperties();
  const existing    = JSON.parse(userProps.getProperty("CHANNEL_REGISTRY") || "[]");
  const existingIds = new Set(existing.map(c => c.id));
  let added = 0;

  Object.keys(allProps)
    .filter(k => k.startsWith("channel_id:"))
    .forEach(k => {
      const channelName = k.replace("channel_id:", "");
      const fileId      = allProps[k];
      if (existingIds.has(fileId)) return; // already in registry — skip
      try {
        const ss = SpreadsheetApp.openById(fileId);
        existing.push({ channel: channelName, id: fileId, url: ss.getUrl() });
        existingIds.add(fileId);
        added++;
        Logger.log("Recovered #" + channelName);
      } catch(e) {
        Logger.log("Could not recover #" + channelName + ": " + e.toString());
      }
    });

  if (added > 0) {
    userProps.setProperty("CHANNEL_REGISTRY", JSON.stringify(existing));
  }

  Logger.log("recoverSheets: added " + added + " channel(s).");
  return added;
}

// ============================================================
// QUEUE PROCESSOR — triggered on demand and after each save
// ============================================================

function processMyQueue() {
  // Delete this trigger immediately — it's a one-off, not a recurring poll
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "processMyQueue")
    .forEach(t => ScriptApp.deleteTrigger(t));

  const userProps     = PropertiesService.getUserProperties();
  const slackUsername = userProps.getProperty("SLACK_USERNAME");

  Logger.log("processMyQueue running for: " + slackUsername);

  if (!slackUsername) {
    Logger.log("No SLACK_USERNAME set — aborting");
    return;
  }

  const scriptProps = PropertiesService.getScriptProperties();
  const allProps    = scriptProps.getProperties();
  const queueKeys   = Object.keys(allProps).filter(k => k.startsWith("queue_"));

  Logger.log("Total queue entries found: " + queueKeys.length);

  queueKeys.forEach(key => {
    try {
      const entry         = JSON.parse(allProps[key]);

      // ── Reaction entries: process by whoever's trigger fires first ──
      if (entry.source === "reaction") {
        const lock = LockService.getScriptLock();
        if (!lock.tryLock(3000)) {
          Logger.log("Could not acquire lock for " + key + " — skipping");
          return;
        }

        try {
          // Re-check entry still exists — another user may have processed it already
          if (!scriptProps.getProperty(key)) {
            Logger.log("Entry " + key + " already processed — skipping");
            return;
          }

          const messageData = fetchSlackMessage(entry.channelId, entry.messageTs, entry.teamId);
          if (!messageData) {
            Logger.log("Could not fetch message for reaction entry: " + key);
            return;
          }

          const senderName = getSlackDisplayName(messageData.senderUserId);
          const savedBy    = getSlackDisplayName(entry.slackUserId);

          appendToChannelSheet(
            messageData.channelName,
            senderName,
            savedBy,
            messageData.messageText,
            messageData.timestamp,
            messageData.messageLink,
            ""
          );

          scriptProps.deleteProperty(key);
          Logger.log("Processed reaction entry: " + key);
        } finally {
          lock.releaseLock();
        }
        return;
      }

      // ── Shortcut entries: match on stored username as before ──
      const entryUsername = (entry.slackUsername || "").toLowerCase();

      Logger.log("Queue entry " + key + " — slackUsername: '" + entryUsername + "' vs stored: '" + slackUsername + "'");

      if (entryUsername !== slackUsername) {
        Logger.log("Skipping — username mismatch");
        return;
      }

      // Resolve display names here, not in doPost (keeps doPost fast)
      const senderName = entry.senderUserId  ? getSlackDisplayName(entry.senderUserId)  : (entry.senderName  || "Unknown");
      const savedBy    = entry.savedByUserId ? getSlackDisplayName(entry.savedByUserId) : (entry.savedBy     || "Unknown");

      appendToChannelSheet(
        entry.channelName,
        senderName,
        savedBy,
        entry.messageText,
        entry.timestamp,
        entry.messageLink,
        entry.responseUrl || ""
      );

      scriptProps.deleteProperty(key);
      Logger.log("Processed and removed: " + key);

    } catch (err) {
      Logger.log("Queue error for " + key + ": " + err.toString());
    }
  });
}

// ============================================================
// CORE LOGIC — find or create a shared Sheet per channel
// ============================================================

function appendToChannelSheet(channelName, senderName, savedBy, messageText, timestamp, messageLink, responseUrl) {
  const userProps   = PropertiesService.getUserProperties();
  const scriptProps = PropertiesService.getScriptProperties();
  let spreadsheet   = null;

  // Look up shared sheet ID from Script Properties (shared across all users)
  const storedFileId = scriptProps.getProperty("channel_id:" + channelName);

  if (storedFileId) {
    try {
      const file = DriveApp.getFileById(storedFileId);
      spreadsheet = SpreadsheetApp.open(file);
    } catch (err) {
      Logger.log("Shared sheet not found for #" + channelName + " — clearing stale ID and creating a new one.");
      scriptProps.deleteProperty("channel_id:" + channelName);
      spreadsheet = null;
    }
  }

  if (!spreadsheet) {
    // Create new sheet and move it to the shared Slack Messages folder
    spreadsheet = SpreadsheetApp.create(channelName);

    const sheet = spreadsheet.getSheets()[0];
    sheet.setName("Messages");
    sheet.appendRow(["Timestamp", "Sender", "Saved By", "Message", "Link"]);

    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4A154B");
    headerRange.setFontColor("#FFFFFF");
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 400);
    sheet.setColumnWidth(5, 200);
    sheet.setFrozenRows(1);

    // Move to the shared Slack Messages folder
    try {
      const file         = DriveApp.getFileById(spreadsheet.getId());
      const targetFolder = DriveApp.getFolderById(SHARED_FOLDER_ID);
      targetFolder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
      // Share with everyone in the domain so any team member's script can open it
      file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT);
      Logger.log("Moved and shared sheet for #" + channelName);
    } catch (err) {
      Logger.log("Could not move/share sheet: " + err.toString());
    }

    // Store in Script Properties so ALL users find the same sheet
    scriptProps.setProperty("channel_id:" + channelName, spreadsheet.getId());
    Logger.log("Created new shared sheet for #" + channelName + " — ID: " + spreadsheet.getId());
  }

  // Register in this user's dashboard if not already there
  const registry = JSON.parse(userProps.getProperty("CHANNEL_REGISTRY") || "[]");
  if (!registry.find(c => c.channel === channelName)) {
    registry.push({ channel: channelName, id: spreadsheet.getId(), url: spreadsheet.getUrl() });
    userProps.setProperty("CHANNEL_REGISTRY", JSON.stringify(registry));
  }

  const sheet = spreadsheet.getSheetByName("Messages") || spreadsheet.getSheets()[0];

  const readableTime = new Date(timestamp).toLocaleString("en-GB", {
    dateStyle: "short",
    timeStyle: "short",
    timeZone: "Europe/London"
  });

  sheet.appendRow([readableTime, senderName, savedBy, messageText, messageLink]);
  Logger.log("Appended row to shared sheet for #" + channelName);

  // Send ephemeral notification back to Slack via response_url
  if (responseUrl) {
    try {
      const sheetUrl = spreadsheet.getUrl();
      const message  = {
        response_type: "ephemeral",
        replace_original: false,
        text: "✅ Saved to *#" + channelName + "* — <" + sheetUrl + "|Open Sheet>"
      };
      UrlFetchApp.fetch(responseUrl, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(message)
      });
      Logger.log("Slack notification sent");
    } catch (err) {
      Logger.log("Could not send Slack notification: " + err.toString());
    }
  }
}