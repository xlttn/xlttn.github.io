---
Title: Loop Through Tables in the Workbook or Active Sheet
categories: [excel, vba]
tags: [tables]
date: 2017-03-19 18:43:00

---

## Loop through all tables in the workbook

```vb
Sub LoopThroughAllTablesinWorkbook()
  Dim tbl As ListObject
  Dim sht As Worksheet

  '// Loop through each sheet and table in the workbook
  For Each sht In ThisWorkbook.Worksheets
  For Each tbl In sht.ListObjects

  'Do something to all the tables...
  tbl.ShowTotals = True

  Next tbl
  Next sht

End Sub
```

## Loop through all tables in the active sheet

```vb
Sub LoopThroughAllTablesInWorksheet()

  Dim tbl As ListObject

  '// Loop through each sheet and table in the workbook
  For Each tbl In ActiveSheet.ListObjects

  '// Do something to all the tables...
  tbl.ShowTotals = True

  Next tbl

End Sub
```
