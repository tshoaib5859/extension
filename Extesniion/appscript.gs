/**
 * ============================================================
 *  Titan Mail Sender — Google Apps Script Web App  (v2)
 * ============================================================
 *
 *  HOW TO DEPLOY (one-time setup)
 *  ───────────────────────────────────────────────────────────
 *  1. Open your Google Sheet.
 *  2. Click Extensions › Apps Script.
 *  3. Delete ALL existing code and paste this entire file.
 *  4. Click Deploy › New Deployment.
 *       Type           : Web app
 *       Execute as     : Me  (your Google account)
 *       Who has access : Anyone
 *  5. Click Deploy — authorise when prompted.
 *  6. Copy the "Web app URL" (ends with /exec).
 *  7. Paste that URL into the extension popup's
 *       "Apps Script URL" field and click Start.
 *
 *  HOW IT WORKS
 *  ───────────────────────────────────────────────────────────
 *  • The extension fetches 5 rows at a time via GET.
 *  • After each email is sent the extension calls GET again
 *    to mark that specific row as sent ("Email Sent" = "Yes").
 *  • An "Email Sent" column is created automatically if it
 *    does not already exist in your sheet.
 *  • Rows where "Email Sent" is already "Yes" are skipped.
 *
 *  SHEET FORMAT
 *  ───────────────────────────────────────────────────────────
 *  • Row 1 must be column headers.
 *  • One column MUST be named exactly:  Email
 *  • All other columns become template variables:
 *      {{Name}}, {{Serial Number}}, {{Mark Literal Element}} …
 *  • Dates are returned as YYYY-MM-DD strings.
 *
 *  ENDPOINTS  (all GET, no CORS preflight)
 *  ───────────────────────────────────────────────────────────
 *  GET ?action=getRows[&batch=5]
 *      Returns the next N rows where "Email Sent" ≠ "Yes".
 *      Response: { rows, done, pending }
 *
 *  GET ?action=markSent&row=<1-based-row-number>
 *      Sets "Email Sent" = "Yes" for that row.
 *      Response: { ok, row }
 *
 *  GET ?action=ping
 *      Health check — returns { ok: true }.
 * ============================================================
 */

// ─── Configuration ────────────────────────────────────────────────────────────
var SENT_COL_NAME  = "Email Sent";
var EMAIL_COL_NAME = "Email";
var DEFAULT_BATCH  = 5;
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Main GET handler.  All requests from the extension come here.
 */
function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var action = String(params.action || "getRows").trim();

    if (action === "ping") {
      return respond({ ok: true, message: "Titan Mail Sender script is running." });
    }

    if (action === "getRows") {
      var batch = parseInt(params.batch || String(DEFAULT_BATCH), 10);
      if (isNaN(batch) || batch < 1) batch = DEFAULT_BATCH;
      return respond(getNextRows(batch));
    }

    if (action === "markSent") {
      var rowNum = parseInt(params.row, 10);
      if (isNaN(rowNum) || rowNum < 2) {
        return respond({ ok: false, error: "Invalid row number: " + params.row });
      }
      return respond(markRowSent(rowNum));
    }

    return respond({ ok: false, error: "Unknown action: " + action });

  } catch (err) {
    return respond({ ok: false, error: "Script error: " + err.toString() });
  }
}

/**
 * Wrap any object as a JSON ContentService response.
 */
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Return the active sheet.
 */
function getActiveSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

/**
 * Read up to `batchSize` rows that have not been sent yet.
 *
 * @param  {number} batchSize
 * @returns {{ rows: object[], done: boolean, pending: number } | { error: string }}
 */
function getNextRows(batchSize) {
  var sheet   = getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) {
    return { rows: [], done: true, pending: 0 };
  }

  // ── Read header row ────────────────────────────────────────────────────────
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h == null ? "" : h).trim();
  });

  // ── Locate or create "Email Sent" column ──────────────────────────────────
  var sentIdx = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toLowerCase() === SENT_COL_NAME.toLowerCase()) {
      sentIdx = i;
      break;
    }
  }
  if (sentIdx === -1) {
    // Append the column to the right of the last column.
    sentIdx = lastCol;          // 0-based index of the new column
    lastCol  = lastCol + 1;     // total columns after adding
    sheet.getRange(1, lastCol).setValue(SENT_COL_NAME);
    headers.push(SENT_COL_NAME);
  }

  // ── Locate "Email" column ─────────────────────────────────────────────────
  var emailIdx = -1;
  for (var j = 0; j < headers.length; j++) {
    if (headers[j].toLowerCase() === EMAIL_COL_NAME.toLowerCase()) {
      emailIdx = j;
      break;
    }
  }
  if (emailIdx === -1) {
    return { ok: false, error: "No column named '" + EMAIL_COL_NAME + "' found in row 1.", rows: [], done: true, pending: 0 };
  }

  // ── Read all data rows ────────────────────────────────────────────────────
  var dataRows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var batch   = [];
  var pending = 0;
  var tz      = Session.getScriptTimeZone();

  for (var r = 0; r < dataRows.length; r++) {
    var cells    = dataRows[r];
    var emailVal = String(cells[emailIdx] || "").trim();
    var sentVal  = String(cells[sentIdx]  || "").trim().toLowerCase();

    // Skip rows that have no email or are already sent.
    if (!emailVal || sentVal === "yes" || sentVal === "sent") continue;

    pending++;

    if (batch.length < batchSize) {
      var rowObj = { _rowIndex: r + 2 }; // 1-based sheet row number
      for (var c = 0; c < headers.length; c++) {
        var raw = cells[c];
        var val;
        if (raw instanceof Date) {
          val = Utilities.formatDate(raw, tz, "yyyy-MM-dd");
        } else if (raw == null) {
          val = "";
        } else {
          val = String(raw).trim();
        }
        rowObj[headers[c]] = val;
      }
      batch.push(rowObj);
    }
  }

  return {
    ok:      true,
    rows:    batch,
    done:    batch.length === 0,
    pending: pending,
  };
}

/**
 * Write "Yes" into the "Email Sent" column for the given 1-based row number.
 *
 * @param  {number} rowIndex  1-based sheet row (2 = first data row)
 * @returns {{ ok: boolean, row?: number, error?: string }}
 */
function markRowSent(rowIndex) {
  try {
    var sheet   = getActiveSheet();
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
      return String(h == null ? "" : h).trim();
    });

    var sentIdx = -1;
    for (var i = 0; i < headers.length; i++) {
      if (headers[i].toLowerCase() === SENT_COL_NAME.toLowerCase()) {
        sentIdx = i;
        break;
      }
    }

    // Create the column if it does not exist yet.
    if (sentIdx === -1) {
      sentIdx = lastCol;
      sheet.getRange(1, lastCol + 1).setValue(SENT_COL_NAME);
    }

    // sentIdx is 0-based; sheet column is 1-based.
    sheet.getRange(rowIndex, sentIdx + 1).setValue("Yes");

    // Highlight the entire row green so sent emails are visually obvious.
    var numCols = sheet.getLastColumn();
    sheet.getRange(rowIndex, 1, 1, numCols).setBackground("#b7e1cd");

    SpreadsheetApp.flush();

    return { ok: true, row: rowIndex };

  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}
