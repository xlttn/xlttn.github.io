---
Title: Copy Visible cells in Range to a new Workbook
categories: [Excel, VBA]
tags: [copy-data]
date: 2018-09-25

---

Copies only the visible cells in a selected range to a new workbook, a A1

```vb
'==================================================================================================
' ## Copy range to a new workbook, only copies the visible cells in the selected range
'==================================================================================================
Sub RangeToNewWorkbook()
    '// Vars
    Dim shtNew As Worksheet
    Dim rngSelection As Range

    '// Optimise
    Application.ScreenUpdating = False

    '// Set the selected range
    Set rngSelection = Application.Selection.SpecialCells(xlCellTypeVisible)

    '// Create new workbook and set the new sheet
    Application.Workbooks.Add
    Set shtNew = Application.ActiveSheet

    '// Copy the selected range and autofit columns
    rngSelection.Copy
    shtNew.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
    rngSelection.Copy Destination:=shtNew.Range("A1")
    shtNew.Range("A1").Select

    '// Optimise
    Application.ScreenUpdating = True
End Sub
```
