/**
 * DISCLAIMER:
 * This script uses SpreadsheetApp.getUi().alert() to show messages.
 * ⚠️ Execution will pause until you confirm each prompt inside Google Sheets.
 * If you run this script from the Apps Script editor, make sure the sheet is open
 * and you click OK on the dialogs — otherwise the script will appear to "hang."
 */

function scrapeTables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 👇 Modify the sheet name below to your preferred sheet. If it does not exist, it will be created.
  const sheet = ss.getSheetByName("Sheet1") || ss.insertSheet("Sheet1");

  // 👇 Change this URL to the page you want to scrape
  const url = "https://www.w3schools.com/css/css_table.asp";

  let resp;
  try {
    resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0" }, // mimic a browser
    });
  } catch (err) {
    SpreadsheetApp.getUi().alert("Fetch failed: " + err.message);
    return;
  }

  const html = resp.getContentText();

  // Remove scripts, styles, and pre/code blocks for cleaner parsing
  const cleaned = html
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<pre\b[^>]*>[\s\S]*?<\/pre>/gi, "")
    .replace(/<code\b[^>]*>[\s\S]*?<\/code>/gi, "");

  // Find all <table> blocks
  const tableBlocks = Array.from(cleaned.matchAll(/<table\b[\s\S]*?<\/table>/gi)).map(m => m[0]);
  if (tableBlocks.length === 0) {
    SpreadsheetApp.getUi().alert("No <table> elements found.");
    return;
  }

  // Parse each table into structured values and filter out empty tables
  const parsed = tableBlocks
    .map(parseTableToValues)
    .filter(t => t.headers.length > 0 || t.rows.length > 0);

  // Clear sheet before writing
  sheet.clearContents();
  let startRow = 1;

  parsed.forEach((t, i) => {
    const rows = [];

    // Insert headers if found
    if (t.headers.length) rows.push(t.headers);

    // Insert data rows
    rows.push(...t.rows);

    // Skip entirely if both headers and rows are empty
    if (rows.length === 0) return;

    // Ensure all rows are same width
    const maxCols = rows.reduce((m, r) => Math.max(m, r.length), 0);
    const padded = rows.map(r => {
      while (r.length < maxCols) r.push("");
      return r;
    });

    // Label table and dump values
    sheet.getRange(startRow, 1).setValue("Table " + (i + 1));
    startRow++;
    sheet.getRange(startRow, 1, padded.length, maxCols).setValues(padded);
    startRow += padded.length + 2; // space before next table
  });

  SpreadsheetApp.getUi().alert("Done — found " + parsed.length + " table(s).");
}

function parseTableToValues(tableHtml) {
  const headers = [];
  const rows = [];

  // Look at first <tr> to see if it contains <th> cells (headers)
  const headerRowMatch = /<tr\b[^>]*>([\s\S]*?)<\/tr>/i.exec(tableHtml);
  if (headerRowMatch && /<th\b/i.test(headerRowMatch[1])) {
    headers.push(...extractCellsFromRow(headerRowMatch[1], "th"));
  }

  // Extract all <tr> rows
  const rowRegex = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi;
  let rowMatch;
  let rowIndex = 0;
  while ((rowMatch = rowRegex.exec(tableHtml)) !== null) {
    // Skip first row if it was used as headers
    if (rowIndex === 0 && headers.length) {
      rowIndex++;
      continue;
    }
    const cells = extractCellsFromRow(rowMatch[1], "td|th");
    if (cells.length) rows.push(cells);
    rowIndex++;
  }

  return {
    headers: headers.map(cleanCellContent),
    rows: rows.length ? rows.map(r => r.map(cleanCellContent)) : [],
  };
}

function extractCellsFromRow(rowHtml, type = "td|th") {
  const regex = new RegExp(`<(${type})\\b[^>]*>([\\s\\S]*?)<\\/\\1>`, "gi");
  const cells = [];
  let match;
  while ((match = regex.exec(rowHtml)) !== null) {
    cells.push(match[2]);
  }
  return cells;
}

function cleanCellContent(content) {
  return content.replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
}
