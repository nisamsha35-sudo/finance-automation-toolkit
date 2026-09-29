Attribute VB_Name = "CleanBankStatement"
Option Explicit

' Cleans a raw bank sheet and flags possible duplicates. Synthetic data only.
' Sheet "Raw" columns: A Date, B Description, C Reference, D Amount. Flags go in column E.
Sub CleanBankStatement()
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim txt As String, parts() As String

    Set ws = ThisWorkbook.Sheets("Raw")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Range("E1").Value = "Flag"

    ' 1) Convert text dates (dd/mm/yyyy, dd.mm.yyyy, dd-mm-yyyy) to real dates
    For r = 2 To lastRow
        If VarType(ws.Cells(r, 1).Value) = vbString Then
            txt = Replace(Replace(Trim(ws.Cells(r, 1).Value), ".", "/"), "-", "/")
            parts = Split(txt, "/")
            If UBound(parts) = 2 Then
                ws.Cells(r, 1).Value = DateSerial(CLng(parts(2)), CLng(parts(1)), CLng(parts(0)))
            End If
        End If
        ws.Cells(r, 1).NumberFormat = "dd-mmm-yyyy"
        ws.Cells(r, 2).Value = Trim(ws.Cells(r, 2).Value)
    Next r

    ' 2) Flag possible duplicates (same date + reference + amount)
    For r = 2 To lastRow
        If Application.WorksheetFunction.CountIfs( _
            ws.Range("A2:A" & lastRow), ws.Cells(r, 1).Value, _
            ws.Range("C2:C" & lastRow), ws.Cells(r, 3).Value, _
            ws.Range("D2:D" & lastRow), ws.Cells(r, 4).Value) > 1 Then
            ws.Cells(r, 5).Value = "CHECK DUPLICATE"
        End If
    Next r

    MsgBox "Cleaning complete. Review rows flagged in column E.", vbInformation
End Sub
