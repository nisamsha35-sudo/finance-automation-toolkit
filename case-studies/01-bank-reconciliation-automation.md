# Automated Bank Reconciliation and Data Cleaning


| | |
|---|---|
| **Role** | Accountant |
| **Environment** | UAE, VAT-registered company |
| **Tools** | Excel, Power Query, Google Sheets, Apps Script, VBA |

## The problem
Bank statements arrived as CSV and Excel exports with inconsistent date formats, extra header rows and debit and credit amounts in separate columns. Matching them to the ledger was done line by line, which was slow at month-end and made errors hard to trace.

## My approach
1. **Standardize the data.** Power Query steps remove header noise, trim descriptions, convert text dates to real dates and combine debit and credit into a single signed amount.
2. **Match automatically.** Each bank line is matched to a ledger line on amount and reference, within a three-day date window. Every ledger line can be used only once.
3. **Isolate exceptions.** Unmatched bank lines, unmatched ledger lines and possible duplicates are listed separately, so review time goes only to items that need judgment.
4. **Prove the result.** A control check confirms that opening balance plus movements equals closing balance, and the difference must be zero before sign-off.

## Workflow
Bank CSV → Power Query cleaning → Automated matching → Exceptions sheet → Review and posting → Control total check

## Controls
- Matching rules are documented: exact amount, three-day date tolerance
- No ledger entry is matched twice
- All exceptions are reviewed and cleared before month-end sign-off

## Outcome
Manual line-by-line matching was replaced with exception-based review. The process is repeatable each month, produces a clear audit trail, and supports a faster and more reliable close.

## Code and sample files
- [Power Query script](https://github.com/nisamsha35-sudo/finance-automation-toolkit/blob/main/code/power-query/clean-statement.m)
- [Google Apps Script](https://github.com/nisamsha35-sudo/finance-automation-toolkit/blob/main/code/apps-script/reconcile.gs)
- [VBA macro](https://github.com/nisamsha35-sudo/finance-automation-toolkit/blob/main/code/vba/CleanBankStatement.bas)
- [Excel formulas](https://github.com/nisamsha35-sudo/finance-automation-toolkit/blob/main/code/excel-formulas/matching-formulas.md)
- [Synthetic sample data](https://github.com/nisamsha35-sudo/finance-automation-toolkit/tree/main/sample-data)
