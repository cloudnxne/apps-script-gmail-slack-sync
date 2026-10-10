/**
 * MULTI-CLIENT EMAIL EXPORTER (V11)
 * - Multi-user: one app, each user owns their clients, their trigger, their Gmail
 * - Deduplicates against Message IDs only (thread-date dedup removed)
 * - Captures every message in a thread, not just the last one
 * - Signature markers trimmed to avoid truncating reply bodies
 * - Self-serve sheet recovery flow per user
 */

function doGet() {
  const props = PropertiesService.getScriptProperties();
  const registry = JSON.parse(props.getProperty('CLIENT_REGISTRY') || "[]");
  const currentUser = Session.getActiveUser().getEmail();
  const lastRun = props.getProperty('LAST_RUN_' + currentUser) || 'Never';
  const lastError = props.getProperty('LAST_ERROR_' + currentUser) || '';

  const myClients = registry.filter(c => c.owner === currentUser);

  let clientListHtml = myClients.map((c) => {
    const globalIndex = registry.indexOf(c);
    return `
      <div style="background: #f8f9fa; padding: 12px; border-radius: 8px; margin-bottom: 10px; border-left: 4px solid #d93025; display: flex; justify-content: space-between; align-items: center;">
        <div style="flex-grow: 1;">
          <strong style="display:block; color:#333;">${c.name}</strong>
          <span style="font-size: 0.8em; color: #666;">Label: <b>${c.label}</b></span>
          ${c.lastError ? `<span style="font-size: 0.75em; color: #c00; display:block;">⚠️ ${c.lastError}</span>` : ''}
        </div>
        <div style="text-align: right;">
          <a href="${c.url}" target="_blank" style="color: #d93025; text-decoration: none; font-size: 0.85em; font-weight: bold; margin-right: 15px;">OPEN SHEET</a>
          <span onclick="deleteClient(${globalIndex})" style="color: #ccc; cursor: pointer; font-size: 1.2em; font-weight: bold;">&times;</span>
        </div>
      </div>
    `;
  }).join('');

  if (myClients.length === 0) {
    clientListHtml = `<div style="text-align:center; padding: 20px; color:#999; border: 1px dashed #eee; border-radius: 8px;">No active syncs yet.</div>`;
  }

  const html = `
    <div style="font-family: -apple-system, sans-serif; padding: 25px; max-width: 480px; margin: 20px auto; border: 1px solid #eee; border-radius: 12px; background: #fff; box-shadow: 0 10px 25px rgba(0,0,0,0.05);">
      <h2 style="color: #d93025; text-align: center; margin-top:0;">Multi-Sync</h2>

      <div style="background: #f0f7ff; border-radius: 6px; padding: 10px 14px; margin-bottom: 20px; font-size: 0.8em; color: #444;">
        👤 Signed in as: <strong>${currentUser}</strong><br>
        ⏱️ Last sync: <strong>${lastRun}</strong>
        ${lastError ? `<br>⚠️ <span style="color:#c00;">${lastError}</span>` : ''}
      </div>

      <div style="margin-bottom: 25px;">
        <h4 style="margin-bottom:10px; color:#555; text-transform: uppercase; font-size: 0.75em; letter-spacing: 1px;">Your Active Registrations</h4>
        ${clientListHtml}
      </div>

      <div style="padding: 20px; background: #fffcfc; border-radius: 8px; border: 1px solid #ffebeb; margin-bottom: 15px;">
        <h4 style="margin-top:0; color:#d93025;">Add New Client Sync</h4>
        <input type="text" id="client" placeholder="Client Name (e.g. Acme)" style="width:100%; padding:10px; margin-bottom:10px; border:1px solid #ddd; border-radius:4px; box-sizing: border-box;">
        <input type="text" id="label" placeholder="Gmail Label to scan" style="width:100%; padding:10px; margin-bottom:15px; border:1px solid #ddd; border-radius:4px; box-sizing: border-box;">
        <button onclick="addNew()" id="addBtn" style="width:100%; background:#d93025; color:white; border:none; padding:12px; border-radius:5px; font-weight:bold; cursor:pointer;">Activate Sync</button>
      </div>

      <!-- ── RECOVERY SECTION ─────────────────────────────────────── -->
      <div style="padding: 20px; background: #fff9f0; border-radius: 8px; border: 1px solid #ffe0b2; margin-bottom: 15px;">
        <h4 style="margin-top:0; color:#e65100;">🔄 Recover Lost Sheets</h4>
        <p style="font-size: 0.82em; color: #666; margin-top:0; margin-bottom: 14px;">
          If your syncs disappeared after an update, click below to scan your Drive for existing export sheets and re-link them.
        </p>
        <button onclick="findRecoverable()" id="recoverScanBtn" style="width:100%; background:#e65100; color:white; border:none; padding:11px; border-radius:5px; font-weight:bold; cursor:pointer;">Scan for Recoverable Sheets</button>

        <div id="recoveryResults" style="margin-top:14px;"></div>
      </div>
      <!-- ─────────────────────────────────────────────────────────── -->

      <div style="margin-top: 5px; text-align: center;">
        <button onclick="forceRun()" style="background:none; border:none; color:#d93025; font-size:0.8em; text-decoration:underline; cursor:pointer;">Manual Sync Now</button>
      </div>
      <div id="status" style="margin-top:15px; text-align:center; font-size:0.85em; color:#d93025; font-weight:bold;"></div>
    </div>

    <script>
      function addNew() {
        const c = document.getElementById("client").value.trim();
        const l = document.getElementById("label").value.trim();
        if (!c || !l) return alert("Please fill in both fields");
        document.getElementById("addBtn").disabled = true;
        document.getElementById("status").innerText = "Creating spreadsheet...";
        google.script.run
          .withSuccessHandler(() => {
            document.getElementById("status").innerText = "Done! Running first sync...";
            forceRun();
          })
          .withFailureHandler(err => {
            document.getElementById("status").innerText = "Error: " + err.message;
            document.getElementById("addBtn").disabled = false;
          })
          .registerNewClient(c, l);
      }

      function forceRun() {
        document.getElementById("status").innerText = "Syncing...";
        google.script.run
          .withSuccessHandler(() => location.reload())
          .withFailureHandler(err => {
            document.getElementById("status").innerText = "Sync error: " + err.message;
          })
          .runMasterSync();
      }

      function deleteClient(index) {
        if (confirm("Stop syncing this client?")) {
          google.script.run
            .withSuccessHandler(() => location.reload())
            .removeClient(index);
        }
      }

      function findRecoverable() {
        const btn = document.getElementById("recoverScanBtn");
        const resultsDiv = document.getElementById("recoveryResults");
        btn.disabled = true;
        btn.innerText = "Scanning Drive...";
        resultsDiv.innerHTML = '';

        google.script.run
          .withSuccessHandler(function(candidates) {
            btn.disabled = false;
            btn.innerText = "Scan for Recoverable Sheets";

            if (!candidates || candidates.length === 0) {
              resultsDiv.innerHTML = '<p style="font-size:0.82em; color:#666; text-align:center;">No unregistered export sheets found in your Drive.</p>';
              return;
            }

            let html = '<p style="font-size:0.8em; color:#555; margin-bottom:10px;">Found ' + candidates.length + ' sheet(s) not currently registered. Confirm the Gmail label for each, then click Recover.</p>';

            candidates.forEach(function(c, i) {
              html += \`
                <div style="background:#fff; border:1px solid #ddd; border-radius:6px; padding:10px; margin-bottom:8px;">
                  <strong style="font-size:0.9em; color:#333;">\${c.name}</strong>
                  <a href="\${c.url}" target="_blank" style="font-size:0.75em; color:#1a73e8; margin-left:8px;">View Sheet ↗</a>
                  <div style="margin-top:8px; display:flex; gap:6px; align-items:center;">
                    <input
                      type="text"
                      id="recoverLabel_\${i}"
                      value="\${c.guessedLabel}"
                      placeholder="Gmail label"
                      style="flex:1; padding:7px 9px; border:1px solid #ddd; border-radius:4px; font-size:0.82em;"
                    >
                    <button
                      onclick="recoverSheet('\${c.id}', '\${c.name}', \${i})"
                      style="background:#e65100; color:white; border:none; padding:7px 14px; border-radius:4px; font-size:0.82em; font-weight:bold; cursor:pointer; white-space:nowrap;"
                    >Recover</button>
                    <button
                      onclick="dismissCandidate(\${i})"
                      style="background:none; border:1px solid #ddd; color:#999; padding:7px 10px; border-radius:4px; font-size:0.82em; cursor:pointer;"
                    >Skip</button>
                  </div>
                  <div id="recoverStatus_\${i}" style="font-size:0.78em; color:#e65100; margin-top:5px; font-weight:bold;"></div>
                </div>
              \`;
            });

            resultsDiv.innerHTML = html;
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            btn.innerText = "Scan for Recoverable Sheets";
            resultsDiv.innerHTML = '<p style="font-size:0.82em; color:#c00;">Scan failed: ' + err.message + '</p>';
          })
          .findRecoverableSheets();
      }

      function recoverSheet(sheetId, sheetName, index) {
        const labelInput = document.getElementById("recoverLabel_" + index);
        const statusEl  = document.getElementById("recoverStatus_" + index);
        const label = labelInput.value.trim();

        if (!label) {
          statusEl.innerText = "Please enter the Gmail label before recovering.";
          return;
        }

        labelInput.disabled = true;
        statusEl.innerText = "Recovering...";

        google.script.run
          .withSuccessHandler(function() {
            statusEl.style.color = "#2e7d32";
            statusEl.innerText = "✓ Recovered! Reloading...";
            setTimeout(() => location.reload(), 1200);
          })
          .withFailureHandler(function(err) {
            labelInput.disabled = false;
            statusEl.innerText = "Error: " + err.message;
          })
          .recoverSheet(sheetId, sheetName, label);
      }

      function dismissCandidate(index) {
        const card = document.getElementById("recoverStatus_" + index).closest('div[style*="background:#fff"]');
        if (card) card.remove();
      }
    </script>
  `;

  return HtmlService.createHtmlOutput(html)
    .setTitle("Multi-Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ─── BACKEND ─────────────────────────────────────────────────────────────────

function registerNewClient(name, label) {
  const props = PropertiesService.getScriptProperties();
  const registry = JSON.parse(props.getProperty('CLIENT_REGISTRY') || "[]");
  const currentUser = Session.getActiveUser().getEmail();

  if (registry.some(c => c.owner === currentUser && c.label === label)) {
    throw new Error('You already have a client registered with that label.');
  }

  const newSs = SpreadsheetApp.create("Gmail Export - " + name);
  const sheet = newSs.getSheets()[0];
  sheet.getRange(1, 1, 1, 7)
    .setValues([["Date", "From", "To", "Subject", "Body", "Message ID", "Keywords"]])
    .setFontWeight("bold")
    .setBackground("#f3f3f3");
  sheet.setFrozenRows(1);

  registry.push({
    name: name,
    label: label,
    id: newSs.getId(),
    url: newSs.getUrl(),
    owner: currentUser,
    lastError: ''
  });

  props.setProperty('CLIENT_REGISTRY', JSON.stringify(registry));
  ensureMyDailyTrigger();
}


// ─── RECOVERY FUNCTIONS ───────────────────────────────────────────────────────

function findRecoverableSheets() {
  const props = PropertiesService.getScriptProperties();
  const registry = JSON.parse(props.getProperty('CLIENT_REGISTRY') || "[]");
  const currentUser = Session.getActiveUser().getEmail();

  const myRegisteredIds = new Set(
    registry.filter(c => c.owner === currentUser).map(c => c.id)
  );

  const candidates = [];
  const files = DriveApp.searchFiles('title contains "Gmail Export - " and mimeType = "application/vnd.google-apps.spreadsheet"');

  while (files.hasNext()) {
    const file = files.next();
    if (myRegisteredIds.has(file.getId())) continue;

    const rawName = file.getName();
    const clientName = rawName.replace(/^Gmail Export\s*-\s*/i, '').trim();
    const guessedLabel = clientName.toLowerCase().replace(/\s+/g, '-');

    candidates.push({
      id:           file.getId(),
      name:         clientName,
      url:          file.getUrl(),
      guessedLabel: guessedLabel
    });
  }

  return candidates;
}

function recoverSheet(sheetId, sheetName, label) {
  const props = PropertiesService.getScriptProperties();
  const registry = JSON.parse(props.getProperty('CLIENT_REGISTRY') || "[]");
  const currentUser = Session.getActiveUser().getEmail();

  if (registry.some(c => c.owner === currentUser && c.label === label)) {
    throw new Error('You already have a client registered with the label "' + label + '".');
  }

  if (registry.some(c => c.owner === currentUser && c.id === sheetId)) {
    throw new Error('This sheet is already registered to your account.');
  }

  const file = DriveApp.getFileById(sheetId);

  registry.push({
    name:      sheetName,
    label:     label,
    id:        sheetId,
    url:       file.getUrl(),
    owner:     currentUser,
    lastError: ''
  });

  props.setProperty('CLIENT_REGISTRY', JSON.stringify(registry));
  ensureMyDailyTrigger();
}


// ─── SYNC ─────────────────────────────────────────────────────────────────────

function runMasterSync() {
  const props = PropertiesService.getScriptProperties();
  const registry = JSON.parse(props.getProperty('CLIENT_REGISTRY') || "[]");
  const currentUser = Session.getActiveUser().getEmail();
  const myClients = registry.filter(c => c.owner === currentUser);

  const BATCH_SIZE = 500;

  // ── Markers trimmed to quoted/forwarded content only ──────────────────────
  const markers = [
    /\nOn .* wrote:/i,
    /\n>.*$/m,
    /\n_{3,}/m,
    /\n-{3,}\s*Original Message/i,
    /\nFrom:.*Sent:.*To:.*Subject:/is,
    /\n--\s*\n/m,
  ];

  let globalError = '';

  myClients.forEach(client => {
    const globalIndex = registry.findIndex(c => c.owner === currentUser && c.label === client.label);
    try {
      const ss = SpreadsheetApp.openById(client.id);
      const sheet = ss.getSheets()[0];
      const lastRow = sheet.getLastRow();

      // ── Message ID-only dedup ─────────────────────────────────────────────
      const existingMsgIds = new Set();

      if (lastRow > 1) {
        const sheetData = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
        sheetData.forEach(row => {
          const id = row[5];
          if (id) existingMsgIds.add(String(id));
        });
      }

      const query = `label:"${client.label}"`;
      let start = 0;
      const rows = [];

      while (true) {
        const threads = GmailApp.search(query, start, BATCH_SIZE);
        if (threads.length === 0) break;

        threads.forEach(thread => {
          thread.getMessages().forEach(msg => {
            const msgId = msg.getId();

            // Skip only if this exact message ID is already in the sheet
            if (existingMsgIds.has(msgId)) return;

            let body = msg.getPlainBody();
            for (let m of markers) {
              const idx = body.search(m);
              if (idx !== -1) { body = body.substring(0, idx); break; }
            }
            body = body.replace(/[^\x20-\x7E\n\r\t]/g, '').trim();
            const kw = body.split(/\s+/).filter(w => w.length > 7).slice(0, 5).join(", ");

            rows.push([msg.getDate(), msg.getFrom(), msg.getTo(), msg.getSubject(), body, msgId, kw]);
          });
        });

        if (threads.length < BATCH_SIZE) break;
        start += BATCH_SIZE;
      }

      if (rows.length > 0) {
        rows.sort((a, b) => new Date(b[0]) - new Date(a[0]));
        sheet.insertRowsAfter(1, rows.length);
        sheet.getRange(2, 1, rows.length, 7)
          .setValues(rows)
          .setVerticalAlignment("top")
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        sheet.autoResizeColumn(1);
      }

      registry[globalIndex].lastError = '';

    } catch (e) {
      const msg = e.message || 'Unknown error';
      if (globalIndex > -1) registry[globalIndex].lastError = msg;
      globalError = 'Last sync had errors. Check client list.';
      console.error("Error for " + client.name + ": " + msg);
    }
  });

  const now = new Date().toLocaleString('en-GB', { timeZone: 'Europe/London' });
  props.setProperty('LAST_RUN_' + currentUser, now);
  props.setProperty('LAST_ERROR_' + currentUser, globalError);
  props.setProperty('CLIENT_REGISTRY', JSON.stringify(registry));
}


// ─── UTILITIES ────────────────────────────────────────────────────────────────

function removeClient(index) {
  const props = PropertiesService.getScriptProperties();
  const registry = JSON.parse(props.getProperty('CLIENT_REGISTRY') || "[]");
  registry.splice(index, 1);
  props.setProperty('CLIENT_REGISTRY', JSON.stringify(registry));
}

function ensureMyDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runMasterSync')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('runMasterSync')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .inTimezone('Europe/London')
    .create();
}

function initialSetup() {
  ensureMyDailyTrigger();
  Logger.log("Setup complete for " + Session.getActiveUser().getEmail());
}

function migrateOwnership() {
  const props = PropertiesService.getScriptProperties();
  const registry = JSON.parse(props.getProperty('CLIENT_REGISTRY') || "[]");
  const currentUser = Session.getActiveUser().getEmail();

  let changed = 0;
  registry.forEach(c => {
    if (!c.owner) {
      c.owner = currentUser;
      changed++;
    }
  });

  props.setProperty('CLIENT_REGISTRY', JSON.stringify(registry));
  Logger.log(`Migrated ${changed} clients to owner: ${currentUser}`);
}
function cleanAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log("All triggers deleted.");
  ensureMyDailyTrigger();
  Logger.log("Fresh daily trigger created.");
}

function debugSync() {
  const props = PropertiesService.getScriptProperties();
  const registry = JSON.parse(props.getProperty('CLIENT_REGISTRY') || "[]");
  const currentUser = Session.getActiveUser().getEmail();

  Logger.log("Current user: " + currentUser);
  Logger.log("Total registry entries: " + registry.length);

  registry.forEach((c, i) => {
    Logger.log(`--- Client ${i}: ${c.name}`);
    Logger.log(`  Owner: ${c.owner}`);
    Logger.log(`  Label: ${c.label}`);
    Logger.log(`  Owner match: ${c.owner === currentUser}`);

    const threads = GmailApp.search(`label:"${c.label}"`, 0, 5);
    Logger.log(`  Threads found in label: ${threads.length}`);

    try {
      const ss = SpreadsheetApp.openById(c.id);
      const sheet = ss.getSheets()[0];
      Logger.log(`  Sheet rows: ${sheet.getLastRow()}`);
    } catch(e) {
      Logger.log(`  Sheet error: ${e.message}`);
    }
  });
}