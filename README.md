# Excel VBA Macros

VBA macros I wrote in 2020–21 to automate daily market-risk and treasury reporting in
Excel and Outlook: opening the day's report by date, pulling figures out of it,
writing the commentary, colour-coding limits and emailing the result.

File paths, sheet names and addresses are placeholders; point them at your own
workbooks before running.

## Macros

| File | What it does |
|---|---|
| `macros/EmailDailyReport.bas` | Opens the day's report file by date, copies the summary table and builds an Outlook email (HTML body, To/CC, optional attachment) for review before sending |
| `macros/WriteMarketCommentary.bas` | Asks for a report date, opens that day's file and writes a one-line market comment from the 2Y TRY/USD/EUR yield changes (bps) and 3-month ATM volatility |
| `macros/CollectFromDatedFiles.bas` | Loops over a list of dates, opens each day's file from a month/date folder structure and pastes a value back into a summary sheet – a quick way to build a time series from daily files |
| `macros/EarlyWarningColors.bas` | Reads a status code such as `goro` and colours four early-warning-indicator cells green / orange / red |
| `macros/WeeklyTaskReminder.bas` | On the weekly reporting day, asks whether the weekly task is done and marks the tracker cell |

## Snippets

| File | What it does |
|---|---|
| `snippets/FxConvertVisibleRows.bas` | Fills a column with amounts converted at rates looked up from an FX table, down to the last row |
| `snippets/AppendConvertedValue.bas` | Appends a USD-converted figure to the next empty row of a tracking sheet |
| `snippets/PasteAsValues.bas` | Copies a range and pastes it back as values |
| `snippets/WaitThenPressEnter.bas` | Waits two seconds, then sends Enter – for dismissing a dialog during automation |

## Using a macro

In Excel press **Alt + F11**, then **File → Import File…** and pick a `.bas` file.
`EmailDailyReport` also needs **Tools → References → Microsoft Outlook Object Library**.
