/**
 * Shared helpers: CSV parsing, email validation, template rendering with HTML-safe
 * variable substitution, and storage keys. Safe for service worker and popup (no DOM).
 */

(function (global) {
  "use strict";

  const STORAGE_KEYS = {
    QUEUE: "queue",
    CURRENT_INDEX: "currentIndex",
    SENT: "sentEmails",
    FAILED: "failedEmails",
    SENT_SET: "sentEmailSet",
    TEMPLATE: "template",
    SUBJECT_TEMPLATE: "subjectTemplate",
    RUN_STATUS: "runStatus",
    DELAY_MS: "delayMs",
    JITTER_MS: "jitterMs",
    MAX_PER_RUN: "maxPerRun",
    LOGS: "logs",
    TITAN_TAB_ID: "titanTabId",
    PAUSE_REQUESTED: "pauseRequested",
    SCRIPT_URL: "scriptUrl",
    SHEET_MODE: "sheetMode",
    PARALLEL_TABS: "parallelTabs",
  };

  const RUN_STATUS = {
    IDLE: "idle",
    RUNNING: "running",
    PAUSED: "paused",
    STOPPED: "stopped",
  };

  const DEFAULT_DELAY_MS = 5000;
  const DEFAULT_JITTER_MS = 1500;
  const DEFAULT_MAX_PER_RUN = 100;
  const DEFAULT_PARALLEL_TABS = 3;
  const MAX_LOGS = 200;

  const DEFAULT_SUBJECT = "ACTION REQUIRED : MANDATORY VERIFICATION APPOINTMENT";

  // Variables use exact CSV column header names so fillTemplate resolves them.
  const DEFAULT_BODY_TEMPLATE = [
    '<!DOCTYPE html>',
    '<html lang="en">',
    '<head>',
    '<meta charset="UTF-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
    '<style>',
    '  body {font-family:\'Segoe UI\', Tahoma, Geneva, Verdana, sans-serif; margin:0; padding:0; background-color:#f8f9fa; color:#333;}',
    '  .container {max-width:600px; margin:auto; background-color:#ffffff; padding:20px; border-radius:8px;}',
    '  h1, h2, h3 {margin:0 0 10px 0;}',
    '  h1 {color:#856404; font-size:20px;}',
    '  h2 {color:#002e6d; font-size:18px; font-weight:bold;}',
    '  h3 {color:#004085; font-size:16px; font-weight:bold;}',
    '  p {font-size:14px; line-height:1.5;}',
    '  .status {background-color:#d4edda; padding:10px; border-radius:5px; color:#155724;}',
    '  .appointment {background-color:#c8102e; color:white; padding:15px; border-radius:5px; text-align:center;}',
    '  .appointment-details {background-color:white; color:#333; padding:10px; border-radius:5px; margin-top:10px;}',
    '  .verification {background-color:#e7f3ff; padding:10px; border-radius:5px; margin-top:10px;}',
    '  .critical {background-color:#f8d7da; padding:10px; border-radius:5px; color:#721c24; margin-top:10px;}',
    '  .notes {background-color:#d1ecf1; padding:10px; border-radius:5px; color:#0c5460; margin-top:10px;}',
    '  .ack {background-color:#ffc107; padding:10px; font-weight:bold; margin-top:10px;}',
    '  ul {padding-left:20px;}',
    '  li {margin-bottom:5px;}',
    '  .footer {font-size:12px; color:#666; text-align:center; margin-top:20px;}',
    '  .button {display:inline-block; padding:10px 20px; background-color:#0073e6; color:white; text-decoration:none; border-radius:5px; margin:10px 0; font-weight:bold;}',
    '</style>',
    '</head>',
    '<body>',
    '<div class="container">',
    '',
    '  <!-- Header -->',
    '  <div style="background-color:#fff3cd; padding:10px; border-radius:5px;">',
    '    <h1>&#x26A0;&#xFE0F; ACTION REQUIRED: MANDATORY VERIFICATION APPOINTMENT</h1>',
    '    <p>Examining attorney appointment scheduled - failure to attend may result in abandonment</p>',
    '  </div>',
    '',
    '  <p><strong>Dear Applicant,</strong></p>',
    '',
    '  <!-- Trademark Info -->',
    '  <div style="background-color:#f8f9fa; padding:10px; border-radius:5px;">',
    '    <h2>{{Mark Literal Element}}</h2>',
    '    <p><strong>Serial Number:</strong> {{Serial Number}}</p>',
    '    <p><strong>Legal Entity Type:</strong> {{Entity Type}}</p>',
    '    <p><strong>Owner Name:</strong> {{Name}}</p>',
    '    <p><strong>Address:</strong> {{Address}}</p>',
    '    <p><strong>Assigned Office:</strong> Law Office 104</p>',
    '  </div>',
    '',
    '  <!-- Status -->',
    '  <div class="status">',
    '    <h3>&#x2713; Application Status Update</h3>',
    '    <p>Your trademark application is now <strong>live and ready to be verified, attested, and endorsed</strong> by the examining authorities. This allows the USPTO to reserve the mark with all Secretary of State(s) in the US to mark your business as federally protected.</p>',
    '    <p><strong>Your application has matured</strong> and is assigned to an examining attorney from Law Office 104.</p>',
    '  </div>',
    '',
    '  <!-- Appointment -->',
    '  <div class="appointment">',
    '    <h3>&#x1F4DE; SCHEDULED VERIFICATION APPOINTMENT</h3>',
    '    <p>You are required to call the examining attorney at the scheduled time below:</p>',
    '    <div style="font-size:24px; font-weight:bold; background-color:#28a745; padding:10px; border-radius:5px;">(571) 207-5418</div>',
    '    <div class="appointment-details">',
    '      <p><strong>Examining Attorney:</strong> Jeffery Robertson</p>',
    '      <p><strong>Date:</strong> Monday, April 20, 2026</p>',
    '      <p><strong>Appointment Time:</strong> 10:30 AM (PST)</p>',
    '      <p><strong>Appointment Number:</strong> #9892</p>',
    '      <p><strong>Assigned Office:</strong> Law Office 104</p>',
    '    </div>',
    '  </div>',
    '',
    '  <!-- Verification Process -->',
    '  <h2>Verification Process Requirements</h2>',
    '  <p>The verification procedure includes the following elements:</p>',
    '  <div class="verification">',
    '    <h3>&#x1F4CB; Required Verification Components</h3>',
    '    <p><strong>Owner\u2019s Verification:</strong> Name, serial number, address, duration of business use, goods &amp; services, and position in company. Providing EIN saves time.</p>',
    '    <p><strong>Business Information Verification:</strong> Verify owner details, serial number, assigned categories, business description, platform use.</p>',
    '    <p><strong>Prescribed Federal Obligations:</strong> Discuss any federal obligations during the call.</p>',
    '    <p><strong>Conflict or Infringement Review:</strong> Examine potential disputes or infringement claims.</p>',
    '  </div>',
    '',
    '  <!-- Critical -->',
    '  <div class="critical">',
    '    <h3>&#x26A0;&#xFE0F; CRITICAL: Appointment Compliance Requirements</h3>',
    '    <ul>',
    '      <li><strong>YOU ARE REQUIRED TO CALL</strong> the attorney directly at the scheduled time.</li>',
    '      <li>If you fail, you will get ONE FINAL OPPORTUNITY to reschedule.</li>',
    '      <li>Failing verification may result in your application being <strong>abandoned or rejected</strong>.</li>',
    '      <li>If you cannot attend, reply immediately with availability.</li>',
    '      <li>This call is crucial for USPTO processing and forwarding to publication.</li>',
    '    </ul>',
    '  </div>',
    '',
    '  <!-- Acknowledgement -->',
    '  <div class="ack">&#x1F514; ACKNOWLEDGEMENT REQUIRED: Please acknowledge this email once received</div>',
    '',
    '  <!-- Notes -->',
    '  <div class="notes">',
    '    <h3>&#x1F4A1; Important Notes</h3>',
    '    <p><strong>Interaction &amp; Monetary Commitment:</strong> This verification represents an interaction and monetary commitment documented with USPTO.</p>',
    '    <p><strong>Federal Reservation Process:</strong> Successful verification reserves the mark federally.</p>',
    '    <p><strong>Publication Pathway:</strong> Verification forwards application to publication, the path to full registration.</p>',
    '  </div>',
    '',
    '  <p><strong>Regards,</strong><br>United States Patent and Trademark Office<br>Intellectual Property Office<br>600 Dulany Street, Alexandria, Virginia 22314</p>',
    '',
    '  <!-- Footer -->',
    '  <div class="footer">',
    '    <p>United States Patent and Trademark Office (USPTO) is the federal agency for granting patents and registering trademarks.</p>',
    '    <p>It advises the President, Secretary of Commerce, and U.S. government agencies on intellectual property policy and protection.</p>',
    '  </div>',
    '',
    '</div>',
    '</body>',
    '</html>',
  ].join('\n');

  const EMAIL_REGEX =
    /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/;

  function trimStr(s) {
    return String(s == null ? "" : s).trim();
  }

  /**
   * Normalize common email cell formats into a plain address.
   * Accepts: mailto:user@x.com, Name <user@x.com>, quoted values, trailing separators.
   */
  function normalizeEmailCandidate(input) {
    let s = trimStr(input);
    if (!s) return "";

    s = s.replace(/^\uFEFF/, "");
    s = s.replace(/^mailto:/i, "").trim();

    const angled = s.match(/<\s*([^>]+?)\s*>/);
    if (angled && angled[1]) {
      s = angled[1].trim();
    }

    s = s.replace(/^['"\s]+|['"\s]+$/g, "");

    // If cell contains multiple entries, keep first token with '@'
    if (s.includes(",") || s.includes(";") || /\s/.test(s)) {
      const parts = s.split(/[;,\s]+/).map((p) => p.trim()).filter(Boolean);
      const firstEmailish = parts.find((p) => p.includes("@"));
      if (firstEmailish) s = firstEmailish;
    }

    s = s.replace(/[;,.]+$/g, "").trim();
    return s.toLowerCase();
  }

  function isValidEmail(email) {
    const e = normalizeEmailCandidate(email);
    if (!e) return false;
    return EMAIL_REGEX.test(e);
  }

  /**
   * Escape text for insertion into HTML context (variable values only).
   */
  function escapeHtml(text) {
    const s = String(text == null ? "" : text);
    return s
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  /**
   * Safe plain text for subject lines (no HTML entity noise in input fields).
   */
  function plainForSubject(text) {
    return String(text == null ? "" : text)
      .replace(/<[^>]*>/g, "")
      .replace(/[\u0000-\u001F]/g, "");
  }

  /**
   * Normalize header / variable name for lookup (trim, collapse spaces).
   */
  function normalizeKey(key) {
    return trimStr(key).replace(/\s+/g, " ");
  }

  /**
   * Build lookup map: normalized key -> value from row object.
   */
  function rowLookup(row) {
    const map = Object.create(null);
    for (const k of Object.keys(row)) {
      if (k === "email" || k === "_rowIndex") continue;
      map[normalizeKey(k)] = row[k];
    }
    map["email"] = row.email;
    return map;
  }

  /**
   * Replace {{ Name }}, {{Serial Number}}, etc. Missing vars -> empty string.
   * Escapes each substituted value for HTML.
   */
  function fillTemplate(template, row, plainText) {
    const t = template == null ? "" : String(template);
    const lookup = rowLookup(row);
    return t.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (_, rawName) => {
      const name = normalizeKey(rawName);
      const val = lookup[name];
      if (val === undefined || val === null) return "";
      return plainText ? plainForSubject(val) : escapeHtml(val);
    });
  }

  /**
   * Simple CSV parser: supports quoted fields, commas, newlines in quotes.
   */
  function parseCSV(text) {
    function detectDelimiter(input) {
      const s = String(input || "").replace(/^\uFEFF/, "");
      const firstLine = (s.split(/\r?\n/).find((ln) => trimStr(ln) !== "") || "");
      const comma = (firstLine.match(/,/g) || []).length;
      const tab = (firstLine.match(/\t/g) || []).length;
      const semicolon = (firstLine.match(/;/g) || []).length;
      if (tab > comma && tab > semicolon) return "\t";
      if (semicolon > comma && semicolon > tab) return ";";
      return ",";
    }

    const rows = [];
    let row = [];
    let cur = "";
    let i = 0;
    let inQuotes = false;
    const s = String(text || "").replace(/^\uFEFF/, "");
    const delimiter = detectDelimiter(s);

    while (i < s.length) {
      const c = s[i];
      if (inQuotes) {
        if (c === '"') {
          if (s[i + 1] === '"') {
            cur += '"';
            i += 2;
            continue;
          }
          inQuotes = false;
          i++;
          continue;
        }
        cur += c;
        i++;
        continue;
      }
      if (c === '"') {
        inQuotes = true;
        i++;
        continue;
      }
      if (c === delimiter) {
        row.push(cur);
        cur = "";
        i++;
        continue;
      }
      if (c === "\r") {
        i++;
        continue;
      }
      if (c === "\n") {
        row.push(cur);
        rows.push(row);
        row = [];
        cur = "";
        i++;
        continue;
      }
      cur += c;
      i++;
    }
    row.push(cur);
    if (row.length > 1 || trimStr(row[0]) !== "") {
      rows.push(row);
    }
    return rows;
  }

  /**
   * First row = headers. Returns { rows: QueueItem[], skipped: { reason, line }[] }
   */
  function rowsToQueue(csvRows) {
    const skipped = [];
    if (!csvRows.length) {
      return { rows: [], skipped };
    }
    const headerRaw = csvRows[0].map((h) => trimStr(h));
    const headers = headerRaw.map(normalizeKey);

    let emailCol = -1;
    for (let j = 0; j < headers.length; j++) {
      if (headers[j].toLowerCase() === "email") {
        emailCol = j;
        break;
      }
    }
    if (emailCol < 0) {
      skipped.push({ reason: "No Email column found (header must be 'Email')", line: 0 });
      return { rows: [], skipped };
    }

    const out = [];
    const seen = new Set();

    for (let r = 1; r < csvRows.length; r++) {
      const line = csvRows[r];
      const rowObj = { _rowIndex: r };
      for (let c = 0; c < headers.length; c++) {
        const key = headers[c];
        if (!key) continue;
        rowObj[key] = trimStr(line[c] != null ? line[c] : "");
      }
      const emailRaw = trimStr(line[emailCol] != null ? line[emailCol] : "");
      let emailNorm = normalizeEmailCandidate(emailRaw);

      // Fallback for malformed rows: if mapped email is invalid, scan the row for the first valid email.
      if (!isValidEmail(emailNorm)) {
        for (let c = 0; c < line.length; c++) {
          const candidate = normalizeEmailCandidate(line[c]);
          if (isValidEmail(candidate)) {
            emailNorm = candidate;
            break;
          }
        }
      }

      rowObj.email = emailNorm;

      if (!emailNorm) {
        skipped.push({ reason: "Missing email", line: r + 1 });
        continue;
      }
      if (!isValidEmail(emailNorm)) {
        skipped.push({ reason: `Invalid email: ${emailRaw}`, line: r + 1 });
        continue;
      }
      const dedupKey = emailNorm.toLowerCase();
      if (seen.has(dedupKey)) {
        skipped.push({ reason: `Duplicate email in file: ${emailNorm}`, line: r + 1 });
        continue;
      }
      seen.add(dedupKey);
      out.push(rowObj);
    }

    return { rows: out, skipped };
  }

  function hashString(str) {
    let h = 5381;
    const s = String(str);
    for (let i = 0; i < s.length; i++) {
      h = (h * 33) ^ s.charCodeAt(i);
    }
    return (h >>> 0).toString(16);
  }

  const utils = {
    STORAGE_KEYS,
    RUN_STATUS,
    DEFAULT_DELAY_MS,
    DEFAULT_JITTER_MS,
    DEFAULT_MAX_PER_RUN,
    DEFAULT_PARALLEL_TABS,
    MAX_LOGS,
    DEFAULT_SUBJECT,
    DEFAULT_BODY_TEMPLATE,
    trimStr,
    normalizeEmailCandidate,
    isValidEmail,
    escapeHtml,
    plainForSubject,
    normalizeKey,
    fillTemplate,
    parseCSV,
    rowsToQueue,
    hashString,
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = utils;
  }
  global.TitanUtils = utils;
})(typeof self !== "undefined" ? self : globalThis);
