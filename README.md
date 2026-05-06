# Summer Budget Tracker

> A VBA-powered Excel budget tracker built for students managing summer finances.
> Track income, expenses, bills, and savings goals — all in one place.

---

## Overview

Summer Budget Tracker is a fully featured personal finance tool built entirely 
in Microsoft Excel and VBA. It was designed for students who want a simple, 
no-signup, no-subscription way to manage their money over the summer — whether 
income comes from a job, family support, savings, or a mix of everything.

The tracker features a color-coded dashboard that updates with one click, 
category budgets with custom alert thresholds, recurring bill reminders, 
savings goal tracking, and auto-generated charts — all driven by VBA macros 
with no external dependencies.

---

## Features

### Dashboard
- Monthly income, expenses, and net cash flow summary
- Color-coded category budget status (OK / Warning / AT LIMIT / OVER BUDGET)
- Upcoming bills due in the next 14 days with urgency highlighting
- Savings goal progress with deadlines
- Three auto-generated charts: expenses by category, budget vs spent, 
  income vs expenses

### Income Tracking
- Log multiple income sources (job, family support, side gigs, etc.)
- Date-stamped entries with notes
- Auto-summed into the dashboard monthly total

### Expense Tracking
- Log every purchase with date, category, description, and amount
- Category dropdown linked to your personal category list
- Automatically reflected in dashboard budget tracking

### Category Budgets
- Predefined spending categories (fully locked — no accidental edits)
- Set a monthly budget limit per category
- Set a custom alert threshold % per category
- Status badges update automatically on dashboard refresh

### Bill Reminders
- Log recurring bills with due day, frequency, and last paid date
- Auto-calculates next due date (Monthly / Weekly / Yearly)
- Bills due within 14 days appear on the dashboard
- Bills due within 3 days highlighted in red
- Bills due within 4-7 days highlighted in yellow

### Savings Goals
- Set goals with a target amount, current savings, and deadline
- Progress percentage auto-calculated on refresh
- All goals displayed on the dashboard with deadlines

### Welcome Screen
- Appears every time the file opens
- Three options: Continue (load saved data), New (reset all data), Quit (close)
- Double-confirmation before any data reset to prevent accidents

---

## How to Use

### Requirements
- Microsoft Excel (Windows recommended)
- Macros must be enabled
- Excel 2016 or newer recommended

### Getting Started

1. Download `Summer_Budget.xlsm` from this repository
2. Open the file in Excel
3. When prompted, click **Enable Content** to allow macros
4. The welcome screen appears — click **Yes** to continue or start fresh
5. Navigate to the **Categories** sheet and set your monthly budgets
6. Start logging income on the **Income** sheet
7. Log purchases on the **Expenses** sheet
8. Add recurring bills on the **Bills** sheet
9. Set savings targets on the **Goals** sheet
10. Click **Refresh Dashboard** anytime to see updated totals and charts

### Sheets at a Glance

| Sheet | What it does |
|---|---|
| Dashboard | Auto-generated summary — read only, rebuilt on refresh |
| Income | Log all money coming in |
| Expenses | Log every purchase or payment |
| Categories | Predefined spending groups with budget limits |
| Bills | Recurring payments with due dates and reminders |
| Goals | Savings targets with deadlines and progress tracking |

### Category Budget Status

| Status | Meaning | Color |
|---|---|---|
| OK | Spending is under the alert threshold | Green |
| Warning | Spending has passed the alert threshold | Orange |
| AT LIMIT | Spent exactly the monthly budget | Dark orange |
| OVER BUDGET | Spending has exceeded the monthly budget | Red |
| No budget set | Category exists but no limit was set | Gray |

### Bill Frequency Options

When adding a bill, the Frequency column accepts exactly one of:

| Value | Meaning |
|---|---|
| Monthly | Due on the same day every month (uses Due Day column) |
| Weekly | Due every 7 days from last paid date |
| Yearly | Due once a year on the same calendar date |

---

## Technical Details

- **Language:** VBA (Visual Basic for Applications)
- **Platform:** Microsoft Excel (.xlsm)
- **No external dependencies** — everything runs inside Excel
- **No database** — all data stored directly in Excel sheets
- **Distribution:** single file, shareable via email or download

### Key Modules

| Sub / Function | Purpose |
|---|---|
| `BuildDashboard` | Rebuilds the entire dashboard with current data and styling |
| `BuildExpenseChart` | Generates the three charts from dashboard table data |
| `GetCategorySpend` | Calculates total spending per category for the current month |
| `UpdateBillDueDates` | Auto-calculates next due date for every bill |
| `UpdateGoalProgress` | Calculates percentage complete for each savings goal |
| `StyleAllSheets` | Applies consistent color theming across all data sheets |
| `ResetAllData` | Clears all data while preserving headers and validation |
| `ShowWelcome` | Displays the welcome screen on file open |
| `RefreshAll` | Master refresh — runs all update subs in sequence |

---

## Screenshots

<img width="2559" height="1531" alt="image" src="https://github.com/user-attachments/assets/1d247d77-eee5-401f-93ea-8bc58bd10127" />



---

## Known Limitations

- **Mac compatibility:** VBA runs on Mac Excel but with some limitations. 
  Full testing was done on Windows.
- **Excel Online:** Not supported. VBA macros do not run in browser-based Excel.
- **Emoji in VBA:** Emoji characters in MsgBox text may display as `??` on 
  older Excel versions — plain text fallbacks are used throughout.
- **Chart rendering:** Charts are recreated on every refresh. If charts appear 
  off-screen, adjust the `Left:=` and `Top:=` values in `BuildExpenseChart`.

---

## Roadmap

Potential future additions:

- [ ] UserForm-based data entry (custom dialog boxes instead of typing in cells)
- [ ] Monthly history view (track spending across multiple months)
- [ ] CSV import from bank statements
- [ ] PDF export of monthly summary
- [ ] Weekly spending breakdown
- [ ] Dark mode theming

---

## License

MIT License — free to use, modify, and distribute.

---

## Author

Built by Sadnan Sadib Nahin  
Made for students who want to stop guessing where their money went.
