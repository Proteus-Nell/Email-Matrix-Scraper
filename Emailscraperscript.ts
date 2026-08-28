/**
 * EmailScraperScript
 *
 * Takes the raw body of a CALIBRATION email, parses the participant/case matrix,
 * unpivots it into one row per (Case, Participant), drops anything already
 * recorded, and drops the remainder onto a NEW temporary worksheet laid out in
 * the exact column order of the "MainDataset" table.
 *
 * Nothing is written to MainDataset itself. You fill in the manual columns on
 * the staging sheet, copy the block, paste it into the table, then delete the
 * staging sheet.
 *
 * Parameters (exposed in Power Automate's "Run script" action):
 *   emailBody         - required. Pass the trigger's "Body" field.
 *   sheetLabel        - optional. Suffix for the sheet name. Blank = UTC timestamp.
 *   caseColumn        - optional. Blank falls back to "Case".
 *   participantColumn - optional. Blank falls back to "Participant".
 *   resultColumn      - optional. Blank falls back to "Results".
 *
 * All four optional parameters are declared with "?" so that Power Automate
 * shows them as optional and accepts an empty box. Blank values fall back to
 * the defaults resolved at the top of main().
 */

// Every staging sheet name starts with this. Used to find sheets still pending.
const STAGING_PREFIX = "CAL_";

// 1-based row on the staging sheet that holds the header row.
const HEADER_ROW = 3;

// Set to true if a case number should only ever appear once, regardless of participant.
const MATCH_ON_CASE_ONLY = false;

const BANNER_FILL = "#FCE4D6";
const AUTO_FILL = "#E2EFDA"; // columns the script fills
const MANUAL_FILL = "#FFF2CC"; // columns you fill by hand

function main(
    workbook: ExcelScript.Workbook,
    emailBody: string,
    sheetLabel?: string,
    caseColumn?: string,
    participantColumn?: string,
    resultColumn?: string
): string {

    // Resolve blanks, empty strings and undefined to sensible defaults.
    const caseName = fallback(caseColumn, "Case");
    const partName = fallback(participantColumn, "Participant");
    const resName = fallback(resultColumn, "Results");
    const label = fallback(sheetLabel, "");

    const table = workbook.getTable("MainDataset");
    if (!table) {
        throw new Error("No table named 'MainDataset' was found in this workbook.");
    }

    const headers = table.getHeaderRowRange().getValues()[0].map(h => String(h).trim());
    const caseIdx = columnIndex(headers, caseName);
    const partIdx = columnIndex(headers, partName);
    const resIdx = columnIndex(headers, resName);

    // ---------------------------------------------------------------- 1. Parse
    const grid = parseGrid(emailBody);
    if (grid.length < 2) {
        return "No matrix found in the email body. No sheet created.";
    }

    const participants = grid[0].slice(1); // first header cell is the "X" corner
    if (participants.length === 0) {
        return "Header row had no participant names. No sheet created.";
    }

    // ------------------------------------------------------------- 2. Unpivot
    const records: string[][] = []; // [caseNo, participant, result]
    for (let r = 1; r < grid.length; r++) {
        const row = grid[r];
        const caseNo = (row[0] || "").trim();
        if (!caseNo) continue;

        for (let c = 0; c < participants.length; c++) {
            const raw = row[c + 1];
            const result = raw ? raw.trim() : "";
            if (!result) continue;
            records.push([caseNo, participants[c], normaliseResult(result)]);
        }
    }

    if (records.length === 0) {
        return "Matrix parsed but no data cells were found. No sheet created.";
    }

    // ------- 3. Keys already in MainDataset OR sitting on a pending staging sheet
    const seen = new Set<string>();
    collectTableKeys(table, caseIdx, partIdx, seen);
    collectStagingKeys(workbook, caseName, partName, seen);

    // ------------------------------------------------------ 4. Filter new rows
    const newRows: string[][] = [];
    let skipped = 0;

    for (const rec of records) {
        const key = makeKey(rec[0], rec[1]);
        if (seen.has(key)) {
            skipped++;
            continue;
        }
        seen.add(key);

        const row: string[] = new Array(headers.length).fill("");
        row[caseIdx] = rec[0];
        row[partIdx] = rec[1];
        row[resIdx] = rec[2];
        newRows.push(row);
    }

    if (newRows.length === 0) {
        return `Nothing new. All ${skipped} row(s) are already recorded or already staged.`;
    }

    // ------------------------------------------------- 5. Build staging sheet
    const sheetName = uniqueSheetName(workbook, label);
    const sheet = workbook.addWorksheet(sheetName);
    const width = headers.length;
    const autoCols = [caseIdx, partIdx, resIdx];

    // Banner across the top
    const banner = sheet.getRangeByIndexes(0, 0, 1, width);
    banner.merge(false);
    banner.setValue(
        "TEMPORARY SHEET. Fill in the yellow columns below, select rows " +
        (HEADER_ROW + 1) + " downward, copy them, and paste into the MainDataset table. " +
        "Delete this whole worksheet once the rows are in."
    );
    banner.getFormat().getFill().setColor(BANNER_FILL);
    banner.getFormat().getFont().setBold(true);
    banner.getFormat().setWrapText(true);
    banner.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    sheet.getRange("1:1").getFormat().setRowHeight(46);

    // Header row, in MainDataset's own column order
    const headerRange = sheet.getRangeByIndexes(HEADER_ROW - 1, 0, 1, width);
    headerRange.setValues([headers]);
    headerRange.getFormat().getFont().setBold(true);

    // Data rows
    sheet.getRangeByIndexes(HEADER_ROW, 0, newRows.length, width).setValues(newRows);

    // Shade the columns so it is obvious which ones need typing
    for (let c = 0; c < width; c++) {
        const isAuto = autoCols.indexOf(c) >= 0;
        const colour = isAuto ? AUTO_FILL : MANUAL_FILL;
        sheet.getRangeByIndexes(HEADER_ROW - 1, c, newRows.length + 1, 1)
            .getFormat().getFill().setColor(colour);
    }

    sheet.getFreezePanes().freezeRows(HEADER_ROW);
    sheet.getRangeByIndexes(HEADER_ROW - 1, 0, newRows.length + 1, width)
        .getFormat().autofitColumns();
    sheet.activate();

    // Threaded comment on the banner cell, in case the fill colour gets lost
    try {
        workbook.addComment(
            sheet.getRange("A1"),
            "Staged from a CALIBRATION email. Complete the yellow columns, copy rows " +
            (HEADER_ROW + 1) + "+ into MainDataset, then right-click this sheet tab and delete it."
        );
    } catch (e) {
        // Comments are not critical. The banner text carries the same instruction.
    }

    return `Staged ${newRows.length} row(s) on sheet '${sheetName}', skipped ${skipped} duplicate(s).`;
}

/* ======================================================================= */
/* Helpers                                                                 */
/* ======================================================================= */

/** Treats undefined, empty and whitespace-only values as "not supplied". */
function fallback(value: string | undefined, standIn: string): string {
    if (value === undefined || value === null) return standIn;
    const trimmed = String(value).trim();
    return trimmed.length > 0 ? trimmed : standIn;
}

function collectTableKeys(
    table: ExcelScript.Table,
    caseIdx: number,
    partIdx: number,
    seen: Set<string>
): void {
    if (table.getRowCount() === 0) return;
    const existing = table.getRangeBetweenHeaderAndTotal().getValues();
    for (const row of existing) {
        const c = String(row[caseIdx] || "");
        if (!c.trim()) continue; // blank placeholder row
        seen.add(makeKey(c, String(row[partIdx] || "")));
    }
}

/** Reads any staging sheets that have not been pasted and deleted yet. */
function collectStagingKeys(
    workbook: ExcelScript.Workbook,
    caseColumn: string,
    participantColumn: string,
    seen: Set<string>
): void {
    for (const ws of workbook.getWorksheets()) {
        if (ws.getName().indexOf(STAGING_PREFIX) !== 0) continue;

        const used = ws.getUsedRange();
        if (!used) continue;

        const values = used.getValues();
        if (values.length <= HEADER_ROW) continue;

        const hdr = values[HEADER_ROW - 1].map(h => String(h).trim().toLowerCase());
        const ci = hdr.indexOf(caseColumn.trim().toLowerCase());
        const pi = hdr.indexOf(participantColumn.trim().toLowerCase());
        if (ci < 0 || pi < 0) continue;

        for (let r = HEADER_ROW; r < values.length; r++) {
            const c = String(values[r][ci] || "");
            if (!c.trim()) continue;
            seen.add(makeKey(c, String(values[r][pi] || "")));
        }
    }
}

/** Builds a legal, unused worksheet name. */
function uniqueSheetName(workbook: ExcelScript.Workbook, label: string): string {
    let stamp = label.trim();
    if (!stamp) {
        const iso = new Date().toISOString(); // 2026-08-24T14:32:10.123Z
        stamp = iso.substring(0, 10).replace(/-/g, "") + "_" + iso.substring(11, 16).replace(":", "");
    }
    stamp = stamp.replace(/[:\\\/\?\*\[\]]/g, "-");

    const base = (STAGING_PREFIX + stamp).substring(0, 31);
    let name = base;
    let n = 2;
    while (workbook.getWorksheet(name)) {
        const suffix = "_" + n;
        name = base.substring(0, 31 - suffix.length) + suffix;
        n++;
    }
    return name;
}

function columnIndex(headers: string[], name: string): number {
    const target = name.trim().toLowerCase();
    for (let i = 0; i < headers.length; i++) {
        if (headers[i].toLowerCase() === target) return i;
    }
    throw new Error(
        `Column '${name}' was not found in MainDataset. Headers are: ${headers.join(" | ")}`
    );
}

function makeKey(caseNo: string, participant: string): string {
    const c = caseNo.replace(/\s+/g, " ").trim().toLowerCase();
    if (MATCH_ON_CASE_ONLY) return c;
    const p = participant.replace(/\s+/g, " ").trim().toLowerCase();
    return c + "||" + p;
}

/**
 * Turns the email body into a 2D grid.
 * Prefers a real HTML <table>; falls back to comma-separated plain text lines.
 */
function parseGrid(body: string): string[][] {
    const rows: string[][] = [];
    if (!body) return rows;

    const trMatches = body.match(/<tr[\s\S]*?<\/tr>/gi);

    if (trMatches && trMatches.length > 1) {
        for (const tr of trMatches) {
            const cellMatches = tr.match(/<t[dh][\s\S]*?<\/t[dh]>/gi);
            if (!cellMatches) continue;
            const cells = cellMatches.map(c => cleanCell(c));
            if (cells.some(c => c.length > 0)) rows.push(cells);
        }
        return rows;
    }

    const plain = body
        .replace(/<br\s*\/?>/gi, "\n")
        .replace(/<\/(p|div|tr)>/gi, "\n");
    const lines = stripTags(plain).split("\n");

    for (const line of lines) {
        if (line.indexOf(",") < 0) continue;
        const cells = line.split(",").map(c => decodeEntities(c).replace(/\s+/g, " ").trim());
        while (cells.length > 0 && cells[cells.length - 1] === "") cells.pop();
        if (cells.length >= 2) rows.push(cells);
    }
    return rows;
}

function cleanCell(html: string): string {
    return decodeEntities(stripTags(html)).replace(/\s+/g, " ").trim();
}

function stripTags(html: string): string {
    return html.replace(/<[^>]*>/g, "");
}

function decodeEntities(s: string): string {
    return s
        .replace(/&nbsp;/gi, " ")
        .replace(/&#160;/g, " ")
        .replace(/&#39;/g, "'")
        .replace(/&quot;/gi, '"')
        .replace(/&lt;/gi, "<")
        .replace(/&gt;/gi, ">")
        .replace(/&amp;/gi, "&"); // must stay last
}

/**
 * Tidies result text and fixes the common "Callibrated" misspelling.
 * Replace the body with `return v.trim();` to keep the raw wording.
 */
function normaliseResult(v: string): string {
    let s = v.replace(/\s+/g, " ").trim();
    s = s.replace(/calli?brated/gi, "Calibrated");
    s = s.replace(/^not\s+/i, "Not ");
    return s;
}