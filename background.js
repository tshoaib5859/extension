/**
 * MV3 service worker: queue, tab lifecycle, rate limits, persistence.
 */
importScripts("utils.js");

const U = self.TitanUtils;
const TITAN_URL = "https://secureserver.titan.email/mail/";
const READY_POLL_MS = 2500;
const SEND_TIMEOUT_MS = 45000;
const DOM_RETRY_ATTEMPTS = 3;

/** In-memory: immediate stop without waiting for storage round-trip */
let stopImmediate = false;
/** Processing lock to avoid double loops */
let workerRunning = false;

function nowIso() {
  return new Date().toISOString();
}

function randomJitter(max) {
  const m = Math.max(0, max | 0);
  if (!m) return 0;
  return Math.floor(Math.random() * (2 * m + 1)) - m;
}

async function storageGet(keys) {
  return new Promise((resolve, reject) => {
    chrome.storage.local.get(keys, (data) => {
      const err = chrome.runtime.lastError;
      if (err) reject(err);
      else resolve(data);
    });
  });
}

async function storageSet(obj) {
  return new Promise((resolve, reject) => {
    chrome.storage.local.set(obj, () => {
      const err = chrome.runtime.lastError;
      if (err) reject(err);
      else resolve();
    });
  });
}

async function appendLog(level, message) {
  const data = await storageGet([U.STORAGE_KEYS.LOGS]);
  const logs = Array.isArray(data[U.STORAGE_KEYS.LOGS]) ? data[U.STORAGE_KEYS.LOGS] : [];
  logs.push({ level, message, ts: nowIso() });
  while (logs.length > U.MAX_LOGS) logs.shift();
  await storageSet({ [U.STORAGE_KEYS.LOGS]: logs });
}

function tabMatchesTitan(tab) {
  if (!tab || !tab.url) return false;
  try {
    const u = new URL(tab.url);
    return u.hostname === "webmail.titan.email" || u.hostname.endsWith(".titan.email");
  } catch {
    return false;
  }
}

async function findTitanTabId() {
  const stored = await storageGet([U.STORAGE_KEYS.TITAN_TAB_ID]);
  const id = stored[U.STORAGE_KEYS.TITAN_TAB_ID];
  if (id != null) {
    try {
      const tab = await chrome.tabs.get(id);
      if (tab && tabMatchesTitan(tab)) return tab.id;
    } catch {
      /* tab gone */
    }
  }
  const tabs = await chrome.tabs.query({ url: ["https://secureserver.titan.email/mail/*", "https://*.titan.email/*"] });
  if (tabs.length) {
    const t = tabs[0];
    await storageSet({ [U.STORAGE_KEYS.TITAN_TAB_ID]: t.id });
    return t.id;
  }
  return null;
}

async function openTitanTab() {
  const tab = await chrome.tabs.create({ url: TITAN_URL, active: true });
  await storageSet({ [U.STORAGE_KEYS.TITAN_TAB_ID]: tab.id });
  return tab.id;
}

async function ensureTitanTab() {
  let tabId = await findTitanTabId();
  if (tabId == null) tabId = await openTitanTab();
  return tabId;
}

function waitForTabComplete(tabId) {
  return new Promise((resolve) => {
    function onUpdated(id, info) {
      if (id === tabId && info.status === "complete") {
        chrome.tabs.onUpdated.removeListener(onUpdated);
        resolve();
      }
    }
    chrome.tabs.onUpdated.addListener(onUpdated);
    chrome.tabs.get(tabId, (tab) => {
      if (tab && tab.status === "complete") {
        chrome.tabs.onUpdated.removeListener(onUpdated);
        resolve();
      }
    });
  });
}

/**
 * Ask content script if Titan is ready to compose (logged in + compose UI).
 */
async function pingContentReady(tabId) {
  try {
    const res = await chrome.tabs.sendMessage(tabId, { action: "checkReady" });
    return res && typeof res === "object" ? res : { ready: false, reason: "unknown" };
  } catch (e) {
    return { ready: false, reason: "no_content", error: String(e && e.message) };
  }
}

async function waitUntilReady(tabId, onProgress) {
  const deadline = Date.now() + 10 * 60 * 1000;
  let lastReason = "unknown";
  while (Date.now() < deadline) {
    if (stopImmediate) return { ok: false, reason: "stopped" };
    const st = await storageGet([U.STORAGE_KEYS.RUN_STATUS, U.STORAGE_KEYS.PAUSE_REQUESTED]);
    if (st[U.STORAGE_KEYS.RUN_STATUS] === U.RUN_STATUS.STOPPED) {
      return { ok: false, reason: "stopped" };
    }
    if (st[U.STORAGE_KEYS.PAUSE_REQUESTED]) {
      return { ok: false, reason: "pause" };
    }

    await waitForTabComplete(tabId).catch(() => {});

    const ping = await pingContentReady(tabId);
    lastReason = ping.reason || "unknown";
    if (ping.ready) {
      if (onProgress) onProgress("ready");
      return { ok: true };
    }
    if (ping.reason === "login") {
      await storageSet({
        [U.STORAGE_KEYS.RUN_STATUS]: U.RUN_STATUS.PAUSED,
        [U.STORAGE_KEYS.PAUSE_REQUESTED]: false,
      });
      await appendLog("warn", "Titan login page detected — paused. Log in, then Resume.");
      if (onProgress) onProgress("login");
    }
    await new Promise((r) => setTimeout(r, READY_POLL_MS));
  }
  return { ok: false, reason: "timeout", detail: lastReason };
}

async function sendEmailViaContent(tabId, payload) {
  const attempt = async () => {
    return new Promise((resolve, reject) => {
      const t = setTimeout(() => {
        reject(new Error("Send timeout"));
      }, SEND_TIMEOUT_MS);
      chrome.tabs
        .sendMessage(tabId, { action: "sendEmail", payload })
        .then((res) => {
          clearTimeout(t);
          resolve(res);
        })
        .catch((e) => {
          clearTimeout(t);
          reject(e);
        });
    });
  };

  let lastErr;
  for (let i = 0; i < DOM_RETRY_ATTEMPTS; i++) {
    try {
      const res = await attempt();
      if (res && res.ok) return { ok: true };
      lastErr = res && res.error ? res.error : "unknown";
      await new Promise((r) => setTimeout(r, 800 * (i + 1)));
    } catch (e) {
      lastErr = e && e.message ? e.message : String(e);
      await new Promise((r) => setTimeout(r, 800 * (i + 1)));
    }
  }
  return { ok: false, error: lastErr || "failed" };
}

async function sendEmailWithRetry(tabId, payload) {
  let r = await sendEmailViaContent(tabId, payload);
  if (r.ok) return r;
  await appendLog("warn", `Retrying send once after failure: ${r.error}`);
  await new Promise((res) => setTimeout(res, 1200));
  r = await sendEmailViaContent(tabId, payload);
  return r;
}

/**
 * Delay between sends; aborts early on stop or pause request.
 * @returns {{ ok: true } | { ok: false, reason: 'stop' | 'pause' }}
 */
async function sleepWithStop(ms) {
  const step = 200;
  let left = ms;
  while (left > 0) {
    if (stopImmediate) return { ok: false, reason: "stop" };
    const st = await storageGet([U.STORAGE_KEYS.RUN_STATUS, U.STORAGE_KEYS.PAUSE_REQUESTED]);
    if (st[U.STORAGE_KEYS.RUN_STATUS] === U.RUN_STATUS.STOPPED) return { ok: false, reason: "stop" };
    if (st[U.STORAGE_KEYS.PAUSE_REQUESTED]) return { ok: false, reason: "pause" };
    const chunk = Math.min(step, left);
    await new Promise((r) => setTimeout(r, chunk));
    left -= chunk;
  }
  return { ok: true };
}

// ─────────────────────────────────────────────────────────────────────────────
//  Google Sheet helpers
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Fetch next batch of rows from the Apps Script web app.
 * Uses GET-only to avoid CORS preflight.
 */
async function fetchSheetBatch(scriptUrl, batchSize) {
  try {
    const url = `${scriptUrl.replace(/\/$/, "")}?action=getRows&batch=${batchSize || 5}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    return json; // { ok, rows, done, pending } or { ok: false, error }
  } catch (e) {
    return { ok: false, error: String(e && e.message) };
  }
}

/**
 * Tell the Apps Script to mark a row as sent.
 */
async function markSheetRowSent(scriptUrl, rowIndex) {
  try {
    const url = `${scriptUrl.replace(/\/$/, "")}?action=markSent&row=${rowIndex}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    return json;
  } catch (e) {
    return { ok: false, error: String(e && e.message) };
  }
}

/**
 * Convert a raw Apps Script row object to a queue item.
 * The sheet row already has exact column headers as keys.
 * We normalise the email field and preserve _rowIndex.
 */
function sheetRowToQueueItem(sheetRow) {
  const item = Object.assign({}, sheetRow);
  // Find the email value regardless of column capitalisation.
  const emailKey = Object.keys(item).find((k) => k.toLowerCase() === "email");
  const rawEmail = emailKey ? U.trimStr(item[emailKey]) : "";
  item.email = U.normalizeEmailCandidate(rawEmail);
  return item;
}

/**
 * Parallel chunk loop: fetches N rows, opens N Titan tabs simultaneously,
 * sends one email per tab in parallel, closes all tabs, then repeats.
 * N = parallelTabs setting (default 3, max 10).
 */
async function parallelChunkLoop() {
  if (workerRunning) return;
  workerRunning = true;
  try {
    while (true) {
      if (stopImmediate) break;

      const data = await storageGet([
        U.STORAGE_KEYS.RUN_STATUS,
        U.STORAGE_KEYS.PAUSE_REQUESTED,
        U.STORAGE_KEYS.QUEUE,
        U.STORAGE_KEYS.CURRENT_INDEX,
        U.STORAGE_KEYS.TEMPLATE,
        U.STORAGE_KEYS.SUBJECT_TEMPLATE,
        U.STORAGE_KEYS.DELAY_MS,
        U.STORAGE_KEYS.SENT_SET,
        U.STORAGE_KEYS.SHEET_MODE,
        U.STORAGE_KEYS.SCRIPT_URL,
        U.STORAGE_KEYS.PARALLEL_TABS,
      ]);

      if (data[U.STORAGE_KEYS.RUN_STATUS] !== U.RUN_STATUS.RUNNING) break;

      if (data[U.STORAGE_KEYS.PAUSE_REQUESTED]) {
        await storageSet({
          [U.STORAGE_KEYS.PAUSE_REQUESTED]: false,
          [U.STORAGE_KEYS.RUN_STATUS]: U.RUN_STATUS.PAUSED,
        });
        await appendLog("info", "Paused before next chunk.");
        break;
      }

      const template  = data[U.STORAGE_KEYS.TEMPLATE] || "";
      const subject   = data[U.STORAGE_KEYS.SUBJECT_TEMPLATE] || U.DEFAULT_SUBJECT;
      const sheetMode = !!data[U.STORAGE_KEYS.SHEET_MODE];
      const scriptUrl = U.trimStr(data[U.STORAGE_KEYS.SCRIPT_URL] || "");
      const parallelN = Math.max(1, Math.min((data[U.STORAGE_KEYS.PARALLEL_TABS] | 0) || U.DEFAULT_PARALLEL_TABS, 10));
      const sentSet   = new Set(data[U.STORAGE_KEYS.SENT_SET] || []);

      if (!template.trim()) {
        await appendLog("error", "Template is empty — stopping.");
        await storageSet({ [U.STORAGE_KEYS.RUN_STATUS]: U.RUN_STATUS.STOPPED });
        break;
      }

      // ── Assemble next chunk of rows ──────────────────────────────────────
      let chunk = [];

      if (sheetMode && scriptUrl) {
        await appendLog("info", `Fetching next batch of ${parallelN} rows from Google Sheet…`);
        const batchRes = await fetchSheetBatch(scriptUrl, parallelN);
        if (!batchRes.ok || batchRes.error) {
          await appendLog("error", `Sheet fetch failed: ${batchRes.error || "unknown"}`);
          await storageSet({ [U.STORAGE_KEYS.RUN_STATUS]: U.RUN_STATUS.PAUSED });
          break;
        }
        if (batchRes.done || !batchRes.rows || !batchRes.rows.length) {
          await appendLog("info", "All sheet rows processed — done.");
          await storageSet({ [U.STORAGE_KEYS.RUN_STATUS]: U.RUN_STATUS.IDLE });
          break;
        }
        for (const raw of batchRes.rows) {
          const row = sheetRowToQueueItem(raw);
          if (!U.isValidEmail(row.email)) continue;
          if (sentSet.has(row.email.toLowerCase())) {
            await appendLog("warn", `Skipping already sent: ${row.email}`);
            if (row._rowIndex) await markSheetRowSent(scriptUrl, row._rowIndex).catch(() => {});
          } else {
            chunk.push(row);
          }
        }
        if (!chunk.length) {
          await appendLog("warn", "Batch had no new rows — fetching more.");
          continue;
        }
        await appendLog("info", `Loaded ${chunk.length} row(s) for this chunk (${batchRes.pending ?? "?"} still pending).`);

      } else {
        // CSV mode: slice next parallelN rows from the stored queue.
        const queue = data[U.STORAGE_KEYS.QUEUE] || [];
        let idx = data[U.STORAGE_KEYS.CURRENT_INDEX] | 0;
        if (idx >= queue.length) {
          await appendLog("info", "Queue finished.");
          await storageSet({ [U.STORAGE_KEYS.RUN_STATUS]: U.RUN_STATUS.IDLE });
          break;
        }
        const raw = queue.slice(idx, idx + parallelN);
        // Advance the index past the whole slice immediately.
        await storageSet({ [U.STORAGE_KEYS.CURRENT_INDEX]: idx + raw.length });
        for (const row of raw) {
          const email = U.trimStr(row.email);
          if (sentSet.has(email.toLowerCase())) {
            await appendLog("warn", `Skipping already sent: ${email}`);
          } else {
            chunk.push(row);
          }
        }
        if (!chunk.length) continue;
      }

      // ── Open Titan tabs for the chunk ────────────────────────────
      await appendLog("info", `Opening ${chunk.length} Titan tab(s)…`);
      const tabIds = [];
      for (const row of chunk) {
        try {
          const tab = await chrome.tabs.create({ url: TITAN_URL, active: false });
          tabIds.push(tab.id);
        } catch (e) {
          tabIds.push(null);
          await appendLog("error", `Failed to open tab for ${row.email}: ${String(e && e.message)}`);
        }
      }

      // ── Wait for all tabs ready + send in parallel ───────────────────────
      const results = await Promise.all(chunk.map(async (row, i) => {
        const tabId = tabIds[i];
        if (!tabId) return { row, ok: false, error: "Tab failed to open" };

        const ready = await waitUntilReady(tabId, () => {});
        if (!ready.ok) return { row, ok: false, error: `Tab not ready: ${ready.reason}` };

        const bodyHtml   = U.fillTemplate(template, row, false);
        const subjectStr = U.fillTemplate(subject, row, true);
        const res = await sendEmailWithRetry(tabId, { to: row.email, subject: subjectStr, bodyHtml });
        return { row, ok: res.ok, error: res.error };
      }));

      // ── Close all tabs immediately ───────────────────────────────────────
      for (const tabId of tabIds) {
        if (tabId) { try { chrome.tabs.remove(tabId); } catch {} }
      }
      await appendLog("info", `Chunk done — closed ${tabIds.filter(Boolean).length} tab(s).`);

      // ── Persist results ──────────────────────────────────────────────────
      const freshSentList = (await storageGet([U.STORAGE_KEYS.SENT]))[U.STORAGE_KEYS.SENT] || [];
      const freshFailList = (await storageGet([U.STORAGE_KEYS.FAILED]))[U.STORAGE_KEYS.FAILED] || [];
      for (const { row, ok, error } of results) {
        if (ok) {
          sentSet.add(row.email.toLowerCase());
          freshSentList.push({ email: row.email, ts: nowIso() });
          await appendLog("info", `Sent: ${row.email}`);
          if (sheetMode && scriptUrl && row._rowIndex) {
            const mr = await markSheetRowSent(scriptUrl, row._rowIndex);
            if (!mr.ok) await appendLog("warn", `Could not mark sheet row ${row._rowIndex}: ${mr.error}`);
          }
        } else {
          freshFailList.push({ email: row.email, reason: error || "send failed", ts: nowIso() });
          await appendLog("error", `Failed: ${row.email} — ${error}`);
          // Still mark as sent in the sheet so it doesn't get re-fetched.
          if (sheetMode && scriptUrl && row._rowIndex) {
            await markSheetRowSent(scriptUrl, row._rowIndex).catch(() => {});
          }
        }
      }
      await storageSet({
        [U.STORAGE_KEYS.SENT]:     freshSentList,
        [U.STORAGE_KEYS.SENT_SET]: Array.from(sentSet),
        [U.STORAGE_KEYS.FAILED]:   freshFailList,
      });

      if (stopImmediate) break;
      const st2 = await storageGet([U.STORAGE_KEYS.RUN_STATUS, U.STORAGE_KEYS.PAUSE_REQUESTED]);
      if (st2[U.STORAGE_KEYS.RUN_STATUS] === U.RUN_STATUS.STOPPED) break;
      if (st2[U.STORAGE_KEYS.PAUSE_REQUESTED]) {
        await storageSet({
          [U.STORAGE_KEYS.PAUSE_REQUESTED]: false,
          [U.STORAGE_KEYS.RUN_STATUS]: U.RUN_STATUS.PAUSED,
        });
        await appendLog("info", "Paused after chunk.");
        break;
      }

      // Brief delay between chunks.
      const interChunkMs = Math.max(2000, data[U.STORAGE_KEYS.DELAY_MS] || U.DEFAULT_DELAY_MS);
      const slept = await sleepWithStop(interChunkMs);
      if (!slept.ok) {
        if (slept.reason === "pause") {
          await storageSet({
            [U.STORAGE_KEYS.PAUSE_REQUESTED]: false,
            [U.STORAGE_KEYS.RUN_STATUS]: U.RUN_STATUS.PAUSED,
          });
          await appendLog("info", "Paused.");
        }
        break;
      }
    }
  } finally {
    workerRunning = false;
  }
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg || !msg.type) return;

  if (msg.type === "START") {
    (async () => {
      try {
        stopImmediate = false;
        const tpl  = msg.template       != null ? String(msg.template)       : "";
        const subj = msg.subjectTemplate != null ? String(msg.subjectTemplate) : "";
        const scriptUrl = msg.scriptUrl ? U.trimStr(String(msg.scriptUrl)) : "";
        const sheetMode = !!scriptUrl;

        if (!tpl.trim()) {
          sendResponse({ ok: false, error: "Template cannot be empty." });
          return;
        }

        const delayMs     = msg.delayMs      != null ? msg.delayMs      | 0 : U.DEFAULT_DELAY_MS;
        const jitterMs    = msg.jitterMs     != null ? msg.jitterMs     | 0 : U.DEFAULT_JITTER_MS;
        const maxRun      = Math.min(msg.maxPerRun != null ? msg.maxPerRun | 0 : U.DEFAULT_MAX_PER_RUN, 10000);
        const parallelTabs = Math.max(1, Math.min(msg.parallelTabs != null ? msg.parallelTabs | 0 : U.DEFAULT_PARALLEL_TABS, 10));

        // ── Sheet mode ────────────────────────────────────────────────────
        if (sheetMode) {
          await appendLog("info", `Sheet mode — verifying connection: ${scriptUrl}`);
          const testRes = await fetchSheetBatch(scriptUrl, 1);
          if (!testRes.ok || testRes.error) {
            sendResponse({ ok: false, error: `Google Sheet error: ${testRes.error || "Could not reach Apps Script. Make sure it is deployed with 'Anyone' access."}` });
            return;
          }
          if (testRes.done || !testRes.rows || !testRes.rows.length) {
            sendResponse({ ok: false, error: "No pending rows found in the Google Sheet. All rows may already be marked as sent." });
            return;
          }
          await storageSet({
            [U.STORAGE_KEYS.TEMPLATE]:         tpl,
            [U.STORAGE_KEYS.SUBJECT_TEMPLATE]: subj || U.DEFAULT_SUBJECT,
            [U.STORAGE_KEYS.RUN_STATUS]:       U.RUN_STATUS.RUNNING,
            [U.STORAGE_KEYS.PAUSE_REQUESTED]:  false,
            [U.STORAGE_KEYS.DELAY_MS]:         delayMs,
            [U.STORAGE_KEYS.JITTER_MS]:        jitterMs,
            [U.STORAGE_KEYS.MAX_PER_RUN]:      maxRun,
            [U.STORAGE_KEYS.SHEET_MODE]:       true,
            [U.STORAGE_KEYS.SCRIPT_URL]:       scriptUrl,
            [U.STORAGE_KEYS.PARALLEL_TABS]:    parallelTabs,
          });
          await appendLog("info", `Sheet mode started — ${parallelTabs} parallel tab(s) per chunk.`);
          parallelChunkLoop();
          sendResponse({ ok: true, pending: testRes.pending });
          return;
        }

        // ── CSV mode (fallback when no script URL) ────────────────────────
        const parsed = U.parseCSV(msg.csvText || "");
        const built  = U.rowsToQueue(parsed);

        const prev    = await storageGet([U.STORAGE_KEYS.SENT_SET]);
        const sentSet = new Set(prev[U.STORAGE_KEYS.SENT_SET] || []);
        const filtered = built.rows.filter((r) => !sentSet.has(U.trimStr(r.email).toLowerCase()));
        const queue    = filtered.slice(0, maxRun);

        if (!queue.length) {
          const alreadySentCount = built.rows.length - filtered.length;
          const reasonCounts = {};
          for (const item of built.skipped) {
            const baseReason = String(item.reason || "Unknown").replace(/:.+$/, "");
            reasonCounts[baseReason] = (reasonCounts[baseReason] || 0) + 1;
          }
          const topReasons = Object.entries(reasonCounts)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 3)
            .map(([k, v]) => `${k} (${v})`)
            .join(", ");
          let reason = "No valid recipient emails were loaded.";
          if (built.rows.length > 0 && alreadySentCount === built.rows.length) {
            reason = "All valid recipient emails were already sent. Clear sent history to resend.";
          } else if (built.rows.length > 0) {
            reason = "No emails left to queue after filtering. Check max-per-run and sent history.";
          } else if (built.skipped.length) {
            reason = `No valid recipient rows found. ${built.skipped.length} row(s) skipped during validation.`;
          }
          if (topReasons) reason += ` Top reasons: ${topReasons}.`;
          await appendLog("warn", `Start blocked: ${reason}`);
          sendResponse({ ok: false, error: reason, skipped: built.skipped, alreadySentCount });
          return;
        }

        await storageSet({
          [U.STORAGE_KEYS.QUEUE]:            queue,
          [U.STORAGE_KEYS.CURRENT_INDEX]:    0,
          [U.STORAGE_KEYS.TEMPLATE]:         tpl,
          [U.STORAGE_KEYS.SUBJECT_TEMPLATE]: subj || "Message for {{Name}}",
          [U.STORAGE_KEYS.RUN_STATUS]:       U.RUN_STATUS.RUNNING,
          [U.STORAGE_KEYS.PAUSE_REQUESTED]:  false,
          [U.STORAGE_KEYS.DELAY_MS]:         delayMs,
          [U.STORAGE_KEYS.JITTER_MS]:        jitterMs,
          [U.STORAGE_KEYS.MAX_PER_RUN]:      maxRun,
          [U.STORAGE_KEYS.SHEET_MODE]:       false,
          [U.STORAGE_KEYS.SCRIPT_URL]:       "",
          [U.STORAGE_KEYS.PARALLEL_TABS]:    parallelTabs,
        });
        await appendLog("info", `CSV mode started: ${queue.length} emails queued (${parallelTabs} parallel tab(s)).`);
        parallelChunkLoop();
        sendResponse({ ok: true, queued: queue.length, skipped: built.skipped });

      } catch (e) {
        sendResponse({ ok: false, error: String(e && e.message) });
      }
    })();
    return true;
  }

  if (msg.type === "RESUME") {
    (async () => {
      stopImmediate = false;
      await storageSet({
        [U.STORAGE_KEYS.RUN_STATUS]: U.RUN_STATUS.RUNNING,
        [U.STORAGE_KEYS.PAUSE_REQUESTED]: false,
      });
      await appendLog("info", "Resumed.");
      parallelChunkLoop();
      sendResponse({ ok: true });
    })();
    return true;
  }

  if (msg.type === "PAUSE") {
    (async () => {
      await storageSet({ [U.STORAGE_KEYS.PAUSE_REQUESTED]: true });
      await appendLog("info", "Pause requested — will pause after current email.");
      sendResponse({ ok: true });
    })();
    return true;
  }

  if (msg.type === "STOP") {
    (async () => {
      stopImmediate = true;
      await storageSet({
        [U.STORAGE_KEYS.RUN_STATUS]: U.RUN_STATUS.STOPPED,
        [U.STORAGE_KEYS.PAUSE_REQUESTED]: false,
      });
      await appendLog("warn", "Stop requested — halting queue.");
      sendResponse({ ok: true });
    })();
    return true;
  }

  if (msg.type === "GET_STATE") {
    storageGet(null).then((data) => sendResponse({ ok: true, data }));
    return true;
  }

  if (msg.type === "CLEAR_LOGS") {
    storageSet({ [U.STORAGE_KEYS.LOGS]: [] }).then(() => sendResponse({ ok: true }));
    return true;
  }

  if (msg.type === "CLEAR_SENT_HISTORY") {
    (async () => {
      await storageSet({
        [U.STORAGE_KEYS.SENT]: [],
        [U.STORAGE_KEYS.SENT_SET]: [],
      });
      await appendLog("info", "Sent history cleared.");
      sendResponse({ ok: true });
    })();
    return true;
  }

  if (msg.type === "RESET_SESSION") {
    (async () => {
      stopImmediate = true;
      await storageSet({
        [U.STORAGE_KEYS.QUEUE]: [],
        [U.STORAGE_KEYS.CURRENT_INDEX]: 0,
        [U.STORAGE_KEYS.RUN_STATUS]: U.RUN_STATUS.IDLE,
        [U.STORAGE_KEYS.PAUSE_REQUESTED]: false,
      });
      await appendLog("info", "Session queue cleared (sent history kept).");
      sendResponse({ ok: true });
    })();
    return true;
  }
});

chrome.tabs.onRemoved.addListener(async (tabId) => {
  const data = await storageGet([U.STORAGE_KEYS.TITAN_TAB_ID]);
  if (data[U.STORAGE_KEYS.TITAN_TAB_ID] === tabId) {
    await storageSet({ [U.STORAGE_KEYS.TITAN_TAB_ID]: null });
    const st = await storageGet([U.STORAGE_KEYS.RUN_STATUS]);
    if (st[U.STORAGE_KEYS.RUN_STATUS] === U.RUN_STATUS.RUNNING) {
      await appendLog("warn", "Titan tab closed — opening a new tab on next action.");
      try {
        const newId = await openTitanTab();
        await appendLog("info", `Opened Titan tab ${newId}.`);
      } catch (e) {
        await appendLog("error", `Could not reopen Titan tab: ${e && e.message}`);
      }
    }
  }
});

chrome.runtime.onInstalled.addListener(() => {
  const defaults = {
    [U.STORAGE_KEYS.RUN_STATUS]: U.RUN_STATUS.IDLE,
    [U.STORAGE_KEYS.CURRENT_INDEX]: 0,
    [U.STORAGE_KEYS.QUEUE]: [],
    [U.STORAGE_KEYS.SENT]: [],
    [U.STORAGE_KEYS.FAILED]: [],
    [U.STORAGE_KEYS.SENT_SET]: [],
    [U.STORAGE_KEYS.LOGS]: [],
    [U.STORAGE_KEYS.DELAY_MS]: U.DEFAULT_DELAY_MS,
    [U.STORAGE_KEYS.JITTER_MS]: U.DEFAULT_JITTER_MS,
    [U.STORAGE_KEYS.MAX_PER_RUN]: U.DEFAULT_MAX_PER_RUN,
    [U.STORAGE_KEYS.TEMPLATE]: "",
    [U.STORAGE_KEYS.SUBJECT_TEMPLATE]: "Hello {{Name}}",
    [U.STORAGE_KEYS.PAUSE_REQUESTED]: false,
    [U.STORAGE_KEYS.TITAN_TAB_ID]: null,
  };
  chrome.storage.local.get(Object.keys(defaults), (existing) => {
    const patch = {};
    for (const k of Object.keys(defaults)) {
      if (existing[k] === undefined) patch[k] = defaults[k];
    }
    if (Object.keys(patch).length) chrome.storage.local.set(patch);
  });
});
