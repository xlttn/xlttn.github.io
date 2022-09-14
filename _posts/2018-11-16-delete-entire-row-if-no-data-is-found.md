---
Title: Delete entire row if no data is found
categories: [excel, vba]
tags: [deleting-data]
date: 2018-11-16

---

These piece of code lets you delete an entire row within the used range of the sheet if the entire row contains no data.

```vb
' ======================================================================================================
' ## Deletes the entire row within the used range
'    if the ENTIRE row contains no data
'=======================================================================================================
Sub DeleteBlankRows()
    ' Vars
    Dim rngUsedRange As Range
    Dim lngRow As Long

    ' Get the used range on sheet to prepare deleting rows
    Set rngUsedRange = ActiveSheet.UsedRange

    ' Delete entire rows that contain no data
    For lngRow = rngUsedRange.Rows.Count To 1 Step -1
        If WorksheetFunction.CountA(rngUsedRange.Rows(lngRow)) = 0 Then
            rngUsedRange.Rows(lngRow).EntireRow.Delete
        End If
    Next lngRow
End Sub
```
