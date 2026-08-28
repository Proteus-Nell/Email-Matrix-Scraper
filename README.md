## EmailScraperScript

An Office Script that turns the calibration matrix inside an email into rows ready to paste into the `MainDataset` table.

Calibration results arrive as a case-by-participant grid in the body of an email whose subject contains `CALIBRATION`. `MainDataset` stores the same information one row per case-and-participant pair, alongside several columns that are filled in by hand. This script does the transposing and the duplicate checking, then hands you a staging sheet to complete manually. It never writes to `MainDataset` directly.

### What it does

Given a matrix like this in the email body:

|          | Name #01       | Name #02       | Name #03       |
| -------- | -------------- | -------------- | -------------- |
| Case #001| Not Calibrated | Not Calibrated | Not Calibrated |
| Case #002| Not Calibrated | Not Calibrated | Calibrated     |

it produces a new worksheet named `CAL_20260824_1043` containing:

| Case      | Participant | Results        | *(your manual columns)* |
| --------- | ----------- | -------------- | ----------------------- |
| Case #001 | Name #01    | Not Calibrated |                         |
| Case #001 | Name #02    | Not Calibrated |                         |
| Case #001 | Name #03    | Not Calibrated |                         |
| Case #002 | Name #01    | Not Calibrated |                         |

The header row is copied from `MainDataset` in its exact column order, so the block pastes straight in with no rearranging. Columns the script fills are shaded green; the ones you fill are shaded yellow.

### Quick start

1. Open the workbook in Excel on the web, go to **Automate → New Script**, paste in `EmailScraperScript.ts` and save it as `EmailScraperScript`.
2. In Power Automate, create an instant cloud flow with a manual trigger. **When an email is flagged (V4)** works well: you flag the calibration email in Outlook and the flow picks it up.
3. Add **Run script** (Excel Online Business), point it at the workbook and the script, and set `emailBody` to the trigger's **Body** dynamic content.
4. Run it. The script returns a summary string such as `Staged 9 row(s) on sheet 'CAL_20260824_1043', skipped 3 duplicate(s).`
5. Open the workbook, fill in the yellow columns, copy the data rows, paste them into `MainDataset`, then delete the staging sheet.

Step 5 matters. A staging sheet left in place keeps its rows in the duplicate check, so those cases stay suppressed even after you have pasted them in.

### Parameters

| Name | Required | Default | Notes |
| ---- | -------- | ------- | ----- |
| `emailBody` | Yes | | The email's **Body**, not Body Preview. Must be dynamic content, not typed text. |
| `sheetLabel` | No | UTC timestamp | Suffix for the sheet name. The `CAL_` prefix is added for you. |
| `caseColumn` | No | `Case` | Header name in `MainDataset`. Matched case-insensitively. |
| `participantColumn` | No | `Participant` | As above. |
| `resultColumn` | No | `Results` | As above. |

The optional parameters are declared with `?` so Power Automate accepts empty boxes. Blank, whitespace and undefined all fall back to the defaults above.

For a local-time sheet name, pass an expression to `sheetLabel`:

```
formatDateTime(convertFromUtc(utcNow(), 'Arabian Standard Time'), 'yyyyMMdd_HHmm')
```

### Duplicate handling

Rows are keyed on Case plus Participant, normalised for letter case and whitespace. Before staging anything the script builds the set of keys already present in:

- every row of `MainDataset`, and
- every worksheet whose name begins with `CAL_`, meaning batches staged but not yet pasted.

Matching rows are skipped. If nothing survives the filter, no worksheet is created and the return string starts with `Nothing new`. Re-running the flow on the same email is therefore harmless.

To block a case number outright, regardless of participant, set `MATCH_ON_CASE_ONLY` to `true` near the top of the script.

### Parsing behaviour

The script prefers a real HTML `<table>` in the body, which is what Outlook produces when someone pastes a grid from Excel. If it finds no table it falls back to reading comma-separated lines. Either way the first cell of the header row is treated as the empty corner and ignored, and empty result cells are skipped rather than staged blank.

Result text is tidied by `normaliseResult`, which collapses whitespace and rewrites the common `Callibrated` misspelling to `Calibrated` so the same outcome does not end up stored under two spellings. To keep the wording exactly as sent, replace that function's body with `return v.trim();`.

### Return values

| String begins with | Meaning |
| ------------------ | ------- |
| `Staged N row(s)` | Success. Sheet created, name included in the message. |
| `Nothing new` | Everything in the email was already recorded or already staged. No sheet created. |
| `No matrix found` | No table or comma-separated grid in the body. Check you passed Body, not Body Preview. |
| `Header row had no participant names` | The grid was found but the top row had only one cell. |
| `Matrix parsed but no data cells` | Header row read, but every result cell was empty. |

Use this in a Condition after **Run script** to stay quiet on uneventful runs:

```
contains(outputs('Run_script')?['body']?['result'], 'Staged')
```

### Troubleshooting

**`Column 'X' was not found in MainDataset`** — the error lists every header the script did find, which usually makes the mismatch obvious. Check for a singular `Result` or a `Case No` header and pass the exact name.

**`No table named 'MainDataset' was found`** — the script looks for a real Excel table, not a named range or a sheet called MainDataset. Select the data and use **Insert → Table**, then rename it in Table Design.

**Everything parses into one long row** — the body arrived without line breaks. This happens if you type into `emailBody` by hand rather than passing dynamic content.

**A blank row sits above the first staged batch** — Excel keeps a placeholder row in an otherwise empty table. Delete it once, manually.

**The flow fails silently** — the Run script action stops the run on error, before any notification step. Turn on **Settings → Notify me of failures** on the flow.

### Limitations

- Requires the workbook to live in OneDrive for Business or SharePoint. Office Scripts do not run against local files.
- The script cannot run usefully from Excel's own **Run** button, because `emailBody` is a single-line input and the matrix needs its line breaks. It is designed to be driven from Power Automate.
- Concurrent runs can both read the table before either writes. Set **Concurrency Control** to a degree of parallelism of 1 on the trigger.
- Sheet names are capped at Excel's 31-character limit, with `_2`, `_3` appended if two batches land in the same minute.