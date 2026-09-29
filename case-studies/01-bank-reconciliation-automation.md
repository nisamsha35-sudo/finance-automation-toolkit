# Case Study 1: Automated Bank Reconciliation & Data Cleaning

> Anonymized. Company and bank names are placeholders ("Company Alpha", "Bank A"). All figures are illustrative unless marked.

## Context
Sector: services company, UAE. Volume: roughly [X] bank lines per month across [X] accounts. Tools: Excel, Power Query, Google Sheets.

## Situation
Monthly bank statements arrived as CSV/Excel exports with inconsistent date formats, extra header rows and debit/credit split across columns. Matching to the ledger was done by hand, which was slow and error-prone at month-end.

## Task
Standardize the statement format and flag unmatched or duplicate items automatically, so the accountant only reviews exceptions.

## Action
1. **Clean:** Power Query steps to remove header noise, trim descriptions, convert text dates to real dates and combine debit/credit into one signed Amount column.
2. **Match:** rule-based matching on amount, reference and a 3-day date window (Excel `XLOOKUP` or Apps Script).
3. **Exceptions:** unmatched bank lines and unmatched ledger lines listed on a separate tab, plus a duplicate flag.
4. **Control check:** opening balance + movements = closing balance, with a difference cell that must equal zero before sign-off.

## Workflow
`Bank CSV` -> `Power Query clean` -> `Auto-match vs ledger` -> `Exceptions tab` -> `Review and post` -> `Control total = 0`

*(Add a flowchart image here: images/reconciliation-flow.png)*

## Controls and risk
- Tolerance rules documented (date window, amount must match exactly)
- Each ledger line can be matched only once
- Manual review of all exceptions before month-end sign-off

## Result
*Replace with your real numbers:*
- Manual matching time: from [X] hours to [X] hours per month
- Share of lines auto-matched: [X]%
- Reconciliation differences caught before close: [X]

## What I would improve
Add fuzzy matching on descriptions, and a monthly trend of exception types.

## Code
- [Power Query](../code/power-query/clean-statement.m)
- [Apps Script](../code/apps-script/reconcile.gs)
- [VBA](../code/vba/CleanBankStatement.bas)
- [Excel formulas](../code/excel-formulas/matching-formulas.md)
- [Sample data](../sample-data/)
