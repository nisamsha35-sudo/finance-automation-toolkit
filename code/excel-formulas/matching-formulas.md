# Excel Matching Formulas (synthetic data)

Assumes two Excel Tables: `Bank` and `Ledger`, each with Date, Description, Reference, Amount.

## 1. Match within a 3-day window (amount + date)
```
=IFERROR(XLOOKUP(1, (Ledger[Amount]=[@Amount])*(ABS(Ledger[Date]-[@Date])<=3), Ledger[Reference], "UNMATCHED"), "UNMATCHED")
```

## 2. Duplicate flag
```
=IF(COUNTIFS(Bank[Date],[@Date],Bank[Reference],[@Reference],Bank[Amount],[@Amount])>1,"CHECK DUPLICATE","")
```

## 3. Control total (must equal zero)
```
=OpeningBalance + SUM(Bank[Amount]) - ClosingBalance
```

## 4. Matched vs unmatched count
```
=COUNTIF(Bank[Status],"MATCHED")
=COUNTIF(Bank[Status],"UNMATCHED")
```

## 5. AR ageing bucket
```
=IF([@Days]<=30,"0-30",IF([@Days]<=60,"31-60",IF([@Days]<=90,"61-90","90+")))
```
