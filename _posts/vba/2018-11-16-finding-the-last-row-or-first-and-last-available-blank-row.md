---
Title: Finding the Last Row or First and Last Available Blank Row
categories: [Excel, VBA]
tags: [dynamic-ranges]
date: 2018-11-16

---

Examples on how to first the last row, first available blank row or the last blank row

```vb
' last Row, working from the bottom of a Column upwards (therefore the Column can contain blanks & the last Row will still be returned)
MsgBox Cells(Rows.Count, 1).End(xlUp).Row
MsgBox Cells(Rows.Count, "A").End(xlUp).Row
MsgBox ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
MsgBox Range("A" & Rows.Count).End(xlUp).Row
MsgBox Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row).Rows.Count

' From the first cell in a named range to the last used cell in the named ranges furthers column to the right
    Dim colLet As String
    colLet = Split([StartingCell].Address, "$")(1)
    Range([StartingCell].Address & ":" & Range(colLet & Range(colLet & Rows.Count).End(xlUp).Row).Address).Select

' From the active cell down to the last used cell in the active cell's column
    Dim colLet As String
    colLet = Split(ActiveCell.Address, "$")(1)
    Range(ActiveCell.Address & ":" & Range(colLet & Range(colLet & Rows.Count).End(xlUp).Row).Address).Select

' for a Column Range to find the last Row before a blank & the next available blank Cell
MsgBox Range("A1:A" & Range("A1").End(xlDown).Row).Rows.Count
MsgBox Range("A1:A" & Range("A1").End(xlDown).Offset(1, 0).Row).Rows.Count

' anywhere on a Worksheet to find the last Row and the next available blank Row
MsgBox Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
MsgBox Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
MsgBox Cells.Find(What:="*", SearchDirection:=xlPrevious).Offset(1, 0).Row

' find the last Row (will stop at first blank) & the next available Row (Range or Cells methods)
MsgBox Range("A1").End(xlDown).Row
MsgBox Range("A1").End(xlDown).Offset(1, 0).Row
MsgBox Cells(1, 1).End(xlDown).Row
MsgBox Cells(1, 1).End(xlDown).Offset(1, 0).Row

' the last Row or first available Row, working from the bottom of the Column upwards
MsgBox Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
MsgBox Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).Row

' the last Row or first available Row, working from the bottom of the Column upwards, Worksheet Function
MsgBox Cells(Application.WorksheetFunction.CountA(Columns(1)) + 1, 1).Row
MsgBox Cells(Application.WorksheetFunction.CountA(Columns(1)) + 1, 1).Offset(1, 0).Row

' the last Row and first available blank Row for a Defined Name / Named Range called 'Header'
MsgBox [Header].End(xlDown).Row
MsgBox [Header].End(xlDown).Offset(1, 0).Row
MsgBox Range("Header").End(xlDown).Row
MsgBox Range("Header").End(xlDown).Offset(1, 0).Row

' using CurrentRegion & UsedRange (may not always be prefferable)
MsgBox Range("A1").CurrentRegion.Rows.Count
MsgBox ActiveSheet.UsedRange.Rows.Count

' find the last blank row in a table column
Dim myTable As ListObject
Set myTable = ActiveSheet.ListObjects(1) 'or .ListObjects("TableName")
myTable.DataBodyRange.Columns(1).Find("").Select
```
