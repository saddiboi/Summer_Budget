# Summer Budget Tracker

> A VBA-powered Excel budget tracker built for students managing summer finances.
> Track income, expenses, bills, and savings goals — all in one place,
> with a full season history view.

---

## Overview

Summer Budget Tracker is a fully featured personal finance tool built entirely
in Microsoft Excel and VBA. Designed for students who want a simple,
no-signup, no-subscription way to manage their money over the summer —
whether income comes from a job, family support, savings, or a mix of everything.

The tracker features two dashboards: a live monthly dashboard that updates
on demand, and a history overview that automatically tracks every month
you've entered data for — no setup required.

---

## Screenshots

<img width="2559" height="1531" alt="image" src="https://github.com/user-attachments/assets/1d247d77-eee5-401f-93ea-8bc58bd10127" />

---

## Features

### Main Dashboard
- Monthly income, expenses, and net cash flow at a glance
- Color-coded category budget status per spending group
- Upcoming bills due in the next 14 days with urgency highlighting
- Savings goal progress with deadlines
- Three auto-generated charts:
  - Expenses by category (current month, filtered to non-zero only)
  - Budget vs spent by category (filtered to budgeted categories only)
  - Income vs expenses comparison (green vs red column chart)
- Auto-refreshes every time you navigate to the Dashboard tab
- Manual refresh button for on-demand updates

### History / Season Overview
- Automatically detects every month that has income or expense data
- No manual setup — just log data and the history builds itself
- Month by month breakdown with income, expenses, net, savings rate, and status
- Summer totals — cumulative income, expenses, and net savings for the whole season
- Highlights section — best month and toughest month (only flagged if net was negative)
- Top spending categories ranked by total amount across the whole summer
- Savings goal progress with color-coded completion indicators
- Two charts: monthly income vs expenses trend, top categories bar chart
- Auto-refreshes when you navigate to the Overview tab
- Manual refresh button

### Income Tracking
- Log multiple income sources with date, source, amount, and notes
- Supports jobs, family support, side gigs, or any other source
- Auto-summed into the dashboard monthly total and history cumulative total

### Expense Tracking
- Log every purchase with date, category, description, amount, and notes
- Category dropdown linked to your approved category list — no typos
- Automatically reflected in dashboard budget tracking and history

### Category Budgets
- Predefined spending categories — column A is locked so names stay consistent
- Set a monthly budget limit per category (0 = track only, no limit)
- Set a custom alert threshold % per category
- Five status levels with color-coded badges:

| Status | Meaning | Color |
|---|---|---|
| OK | Under alert threshold | Green |
| Warning | Past alert threshold, under budget | Orange |
| AT LIMIT | Spent exactly the monthly budget | Dark orange |
| OVER BUDGET | Exceeded the monthly budget | Red |
| No budget set | Category exists, no limit set | Gray italic |

### Bill Reminders
- Log recurring bills with bill name, amount, due day, frequency, last paid date
- Auto-calculates next due date on every refresh
- Three frequency options: Monthly, Weekly, Yearly
- Bills due within 14 days appear on the dashboard
- Bills due within 3 days highlighted in red
- Bills due within 4–7 days highlighted in yellow
- Due Day column validated to accept only numbers 1–31

### Savings Goals
- Set goals with a name, target amount, current savings, and deadline
- Progress percentage auto-calculated on every refresh
- Color-coded progress badges:

| Color | Meaning |
|---|---|
| Green | 100% — goal reached |
| Yellow | 75–99% — almost there |
| Light yellow | 50–74% — halfway |
| Light orange | Under 50% — early stages |

### Welcome Screen
- Appears every time the file opens
- Three options: Yes (Continue), No (New budget), Cancel (Quit)
- "Continue" — loads saved data and refreshes dashboard
- "New budget" — double-confirmation before wiping data, keeps categories
- "Quit" — saves the file automatically then closes Excel

### Sheet Theming
- Every data sheet has a consistent color theme applied automatically
- Income → Blue, Expenses → Red, Categories → Purple, Bills → Orange, Goals → Green
- Colored header rows, zebra striping, frozen header row, auto-fit columns
- Theming re-applies on every refresh so it always looks clean

---

## How to Use

### Requirements
- Microsoft Excel for Windows (2016 or newer recommended)
- Macros must be enabled
- Mac: VBA runs but with some limitations — Windows recommended

### Getting Started

1. Download [Summer_Budget.zip](https://github.com/saddiboi/Summer_Budget/releases/download/v1.0.0/Summer_Budget.zip)
2. Open in Excel
3. Click **Enable Content** when Excel prompts about macros
4. The welcome screen appears — click **Yes** to continue
5. Go to the **Categories** sheet — set monthly budgets and alert thresholds
6. Go to the **Income** sheet — log your income entries
7. Go to the **Expenses** sheet — log purchases using the category dropdown
8. Go to the **Bills** sheet — add recurring bills with due dates
9. Go to the **Goals** sheet — set savings targets
10. Click **Refresh Dashboard** to see updated totals and charts
11. Click **View Overview** to see your full season history

### Macros Blocked? Unblock the File First

If Excel shows a yellow warning bar saying *"Microsoft has blocked macros from running because the source of this file is untrusted,"* don't worry — this happens to every macro-enabled file downloaded from the internet. Windows tags downloaded files with a "Mark of the Web" and Excel blocks their macros by default for security.

To unblock the file:

1. Close the file in Excel
2. Find `Summer_Budget.xlsm` in File Explorer (usually your Downloads folder)
3. Right-click the file and choose **Properties**
4. At the bottom of the **General** tab, find the security notice
5. Tick the **Unblock** checkbox
6. Click **Apply**, then **OK**
7. Reopen the file in Excel — macros will now work

You only need to do this once. After unblocking, the file behaves like any other Excel workbook.

> **Note for Mac users:** This block doesn't apply on macOS. Just open the file and click **Enable Macros** when prompted.

### Sheets at a Glance

| Sheet | Purpose | Editable |
|---|---|---|
| Dashboard | Auto-generated monthly summary | Read only — rebuilt on refresh |
| Overview | Auto-generated season history | Read only — rebuilt on refresh |
| Income | Log all money coming in | Yes |
| Expenses | Log every purchase or payment | Yes |
| Categories | Spending groups with budget limits | Column B and C only (budget + threshold) |
| Bills | Recurring payments with due dates | Yes |
| Goals | Savings targets with deadlines | Yes |

### Bill Frequency Options

| Value | Meaning |
|---|---|
| Monthly | Due on the same day number every month |
| Weekly | Due every 7 days from last paid date |
| Yearly | Due once a year on the same calendar date |

### Date Entry Tips

- Always enter dates in `YYYY-MM-DD` format (e.g., `2026-05-15`)
- Dates should be right-aligned in the cell — left-aligned means Excel stored it as text and it won't be counted
- The dashboard only counts current month entries — make sure the year is correct
- The history overview picks up all months automatically regardless of year

---

## Navigation

| Button | Location | What it does |
|---|---|---|
| Refresh Dashboard | Dashboard sheet | Rebuilds dashboard, recalculates all data |
| Reset All Data | Dashboard sheet | Wipes Income, Expenses, Bills, Goals (keeps Categories) |
| View Overview | Dashboard sheet | Navigates to the season history sheet |
| Refresh History | Overview sheet | Rebuilds the history dashboard |

---

## Technical Details

- **Language:** VBA (Visual Basic for Applications)
- **Platform:** Microsoft Excel (.xlsm)
- **No external dependencies** — everything runs inside Excel
- **No database** — all data stored directly in sheets
- **Two modules:** Module1 (main dashboard + core logic), Module2 (history dashboard)
- **Distribution:** single `.xlsm` file, shareable via download or email

### Module Structure

**Module1 — Core**

| Sub / Function | Purpose |
|---|---|
| `BuildDashboard` | Rebuilds the entire dashboard with current data and styling |
| `BuildExpenseChart` | Generates three charts from dashboard table data |
| `GetCategorySpend` | Calculates total spending per category for the current month |
| `UpdateBillDueDates` | Auto-calculates next due date for every bill |
| `UpdateGoalProgress` | Calculates percentage complete for each savings goal |
| `StyleAllSheets` | Applies consistent color theming to all data sheets |
| `StyleSheet` | Helper — applies theme to a single sheet |
| `ResetAllData` | Clears all data while preserving headers, formatting, and validation |
| `ResetButtonClicked` | Button handler for Reset — adds confirmation before calling ResetAllData |
| `ShowWelcome` | Displays the welcome screen on file open |
| `QuitApplication` | Saves and closes the workbook cleanly |
| `RefreshAll` | Master refresh — runs all update subs in sequence |
| `RefreshButtonClicked` | Button handler for Refresh — shows confirmation message |

**Module2 — History**

| Sub / Function | Purpose |
|---|---|
| `GoToHistory` | Navigates to Overview sheet and triggers refresh |
| `RefreshHistory` | Triggers a history dashboard rebuild |
| `RefreshHistoryButtonClicked` | Button handler — shows confirmation message |
| `BuildHistoryDashboard` | Builds the full season history view |
| `BuildHistoryCharts` | Generates two charts from visible Overview table data |

**ThisWorkbook**

| Event | Purpose |
|---|---|
| `Workbook_Open` | Calls `ShowWelcome` every time the file is opened |

**Sheet modules**

| Sheet | Event | Purpose |
|---|---|---|
| Overview | `Worksheet_Activate` | Auto-refreshes history when tab is clicked |

---

## Known Limitations

- **No auto-refresh on Dashboard tab** — a `Worksheet_Activate` event on the Dashboard caused an infinite loop (Dashboard refresh activates Dashboard → triggers activate → loops). Manual refresh button is used instead.
- **Mac compatibility** — VBA runs on Mac Excel but with some limitations. Full testing was done on Windows.
- **Excel Online** — not supported. VBA macros do not run in browser-based Excel.
- **Emoji in VBA** — emoji characters in MsgBox text display as `??` on older Excel versions. Plain text is used throughout.
- **Chart rendering** — charts are recreated on every refresh. If charts appear off-screen, adjust `Left:=` and `Top:=` values in `BuildExpenseChart` and `BuildHistoryCharts`.
- **History sheet name** — "History" is a reserved word in Excel. The season overview sheet is named "Overview" instead.

---

## Roadmap

- [ ] UserForm-based data entry (custom dialog boxes)
- [ ] Month selector on dashboard (view any previous month)
- [ ] CSV import from bank statements
- [ ] PDF export of monthly summary
- [ ] Weekly spending breakdown
- [ ] Budget rollover between months
- [ ] Dark mode theming

---

## File Structure

```
Summer_Budget.xlsm
├── Dashboard          → Live monthly summary with charts and budget status
├── Overview           → Full season history across all months
├── Income             → Log all money coming in
├── Expenses           → Log every purchase or payment
├── Categories         → Predefined spending groups with budget limits
├── Bills              → Recurring bills with auto-calculated due dates
└── Goals              → Savings targets with progress tracking
```


---

## License

MIT License — free to use, modify, and distribute.

---

## Author

Built by Sadnan Sadib Nahin.
Made for students who want to stop guessing where their money went.
