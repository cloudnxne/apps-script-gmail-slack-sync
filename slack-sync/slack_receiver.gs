/**
 * SLACK RECEIVER — slack_receiver.gs
 * ============================================================
 * DEPLOYMENT SETTINGS:
 *   Description:    Slack Receiver
 *   Execute as:     Me (admin)
 *   Who has access: Anyone
 *
 * URLs to configure in Slack app settings:
 *   Interactivity & Shortcuts → Request URL
 *   Event Subscriptions → Request URL
 *   (both point to the same Deployment 1 URL)
 *
 * Requires SLACK_BOT_TOKEN in Script Properties.
 * Bot token scopes needed: users:read, channels:history, reactions:read
 * ============================================================
 */

const SAVE_EMOJI = "notebook"; // The emoji name that triggers a save (without colons)

// Creates a one-off trigger to run processMyQueue in ~1 minute.
// Avoids duplicate triggers if one is already pending.
function scheduleProcessing() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "processMyQueue");

  if (triggers.length === 0) {
    ScriptApp.newTrigger("processMyQueue")
      .timeBased()
      .after(60 * 1000) // 1 minute from now
      .create();
    Logger.log("Scheduled processMyQueue in 1 minute");
  } else {
    Logger.log("processMyQueue trigger already scheduled — skipping");
  }
}

function doPost(e) {
  // Acknowledge Slack IMMEDIATELY — must respond within 3 seconds
  const ack = ContentService
    .createTextOutput(JSON.stringify({ response_action: "clear" }))
    .setMimeType(ContentService.MimeType.JSON);

  try {
    Logger.log("doPost called");

    let payload = {};

    if (e.postData && e.postData.contents) {
      const raw = e.postData.contents;
      try {
        payload = JSON.parse(raw);
      } catch (jsonErr) {
        const decoded = decodeURIComponent(raw);
        const match   = decoded.match(/payload=(.+)/);
        if (match) payload = JSON.parse(match[1]);
      }
    } else if (e.parameter && e.parameter.payload) {
      payload = JSON.parse(e.parameter.payload);
    }

    Logger.log("payload type: " + payload.type);

    // ── Slack URL verification ───────────────────────────────
    if (payload.type === "url_verification") {
      return ContentService
        .createTextOutput(JSON.stringify({ challenge: payload.challenge }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Message shortcut ─────────────────────────────────────
    if (payload.type === "message_action") {
      const message    = payload.message || {};
      const channel    = payload.channel || {};
      const slackUser  = payload.user    || {};
      const teamDomain = payload.team ? payload.team.domain : "";

      const senderUserId = message.user || message.username || "Unknown";
      const savedByUserId = slackUser.id || "Unknown";
      const savedByName   = (slackUser.username || slackUser.name || slackUser.id || "Unknown").toLowerCase().trim();

      const messageText = decodeURIComponent((message.text || "").replace(/\+/g, " "));
      const timestamp   = message.ts
        ? new Date(parseFloat(message.ts) * 1000).toISOString()
        : new Date().toISOString();
      const messageLink = teamDomain
        ? `https://${teamDomain}.slack.com/archives/${channel.id}/p${(message.ts || "").replace(".", "")}`
        : "";
      const channelName = channel.name || channel.id || "unknown";

      const queueKey   = "queue_" + new Date().getTime() + "_" + Math.random().toString(36).slice(2, 7);
      const queueEntry = JSON.stringify({
        source:        "shortcut",
        slackUserId:   savedByUserId,
        slackUsername: savedByName,
        channelName,
        senderUserId,
        savedByUserId,
        messageText,
        timestamp,
        messageLink,
        responseUrl:   payload.response_url || ""
      });

      PropertiesService.getScriptProperties().setProperty(queueKey, queueEntry);
      Logger.log("Shortcut queued: " + queueKey + " | channel: #" + channelName);
      scheduleProcessing();
    }

    // ── Reaction added event ─────────────────────────────────
    if (payload.type === "event_callback") {
      const event = payload.event || {};

      if (event.type === "reaction_added" && event.reaction === SAVE_EMOJI) {
        const channelId  = event.item ? event.item.channel : "";
        const messageTs  = event.item ? event.item.ts      : "";
        const savedById  = event.user || "Unknown";
        const teamDomain = payload.team_id || "";

        // Queue channel + ts — message text fetched later in processMyQueue
        const queueKey   = "queue_" + new Date().getTime() + "_" + Math.random().toString(36).slice(2, 7);
        const queueEntry = JSON.stringify({
          source:        "reaction",
          slackUserId:   savedById,
          slackUsername: savedById,   // will be resolved to username in processMyQueue
          channelId,
          messageTs,
          teamId:        payload.team_id || ""
        });

        PropertiesService.getScriptProperties().setProperty(queueKey, queueEntry);
        Logger.log("Reaction queued: " + queueKey + " | channel: " + channelId + " | ts: " + messageTs);
        scheduleProcessing();
      }
    }

  } catch (err) {
    Logger.log("doPost error: " + err.toString());
  }

  return ack;
}