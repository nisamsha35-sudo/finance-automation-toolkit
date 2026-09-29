/**
 * Bank vs Ledger auto-match (Google Sheets). Synthetic data only.
 * Sheets: "Bank" and "Ledger", each with columns: Date, Description, Reference, Amount.
 * Rule: amounts equal (to the fils) and dates within TOLERANCE_DAYS. Each ledger line used once.
 * Output: Status in Bank columns E:F, and an "Exceptions" sheet for unmatched ledger lines.
 */
const TOLERANCE_DAYS = 3;

function reconcileBankToLedger() {
  const ss = SpreadsheetApp.getActive();
  const bank = ss.getSheetByName('Bank');
  const ledger = ss.getSheetByName('Ledger');
  const bankData = bank.getDataRange().getValues();
  const ledgerData = ledger.getDataRange().getValues();
  const used = new Set();
  const results = [];

  for (let i = 1; i < bankData.length; i++) {
    const bDate = new Date(bankData[i][0]);
    const bAmt = Math.round(Number(bankData[i][3]) * 100);
    let matchRow = -1;

    for (let j = 1; j < ledgerData.length; j++) {
      if (used.has(j)) continue;
      const lAmt = Math.round(Number(ledgerData[j][3]) * 100);
      const days = Math.abs((bDate - new Date(ledgerData[j][0])) / 86400000);
      if (bAmt === lAmt && days <= TOLERANCE_DAYS) { matchRow = j; break; }
    }

    if (matchRow > -1) {
      used.add(matchRow);
      results.push(['MATCHED', ledgerData[matchRow][2]]);
    } else {
      results.push(['UNMATCHED', '']);
    }
  }

  bank.getRange(1, 5, 1, 2).setValues([['Status', 'Matched Ledger Ref']]);
  if (results.length) bank.getRange(2, 5, results.length, 2).setValues(results);

  // Unmatched ledger lines go to an Exceptions sheet
  const ex = ss.getSheetByName('Exceptions') || ss.insertSheet('Exceptions');
  ex.clear();
  ex.appendRow(['Date', 'Description', 'Reference', 'Amount (unmatched ledger lines)']);
  for (let j = 1; j < ledgerData.length; j++) {
    if (!used.has(j)) ex.appendRow(ledgerData[j]);
  }

  const matched = results.filter(r => r[0] === 'MATCHED').length;
  SpreadsheetApp.getUi().alert('Matched ' + matched + ' of ' + results.length + ' bank lines.');
}
