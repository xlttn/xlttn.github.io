---
Title: Delete Rows in a Column that are Blank
categories: [Excel, VBA]
tags: [deleting-data, Dynamic-Ranges]
date: 2018-11-16

---

Here I outline lots of ways that you can delete blanks rows in any chosen column.  

```vb
'## If you have a Column with data and blank Cells,
'   you can delete the blank Rows using this method
'   starting at Cell "A2" or for Column "A"
'   (requires error trap for xlCellTypeBlanks if none exist):

' delete any Rows containing blank Cells in Column 1, "A" (Short Notation method included)
On Error Resume Next
Columns(1).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
' [A:A].SpecialCells(xlCellTypeBlanks).EntireRow.Delete
On Error GoTo 0

' delete any Rows containing blank Cells in Column A, starting at Cell "A2" (Short Notation method included)
On Error Resume Next
Range("A2:A" & Cells(Rows.Count, "A").End(xlUp).Row).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
' [A2:A & Cells(Rows.Count, "A").End(xlUp).Row)].SpecialCells(xlCellTypeBlanks).EntireRow.Delete
ActiveSheet.UsedRange.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
On Error GoTo 0

' delete any Rows containing blank Cells in the Active Worksheet for the Used Range
On Error Resume Next
ActiveSheet.UsedRange.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
On Error GoTo 0

' ======================================================================================================
' ## Delete empty rows in the selected range only
'    works for both ranges and tables as deletes one row at a time
'=======================================================================================================
Sub DeleteEmptyRowsSelection()

' Vars
    Dim rng As Range
    Dim rngDelete As Range
    Dim RowCount As Long, ColCount As Long
    Dim EmptyTest As Boolean
    Dim RowDeleteCount As Long, ColDeleteCount As Long
    Dim X As Long
    Dim UserAnswer As Variant

' Test a range selected
    If TypeName(Selection) <> "Range" Then Exit Sub

' If the activesheet is protected
    If ActiveProtected = True Then Exit Sub

' Analyze the UsedRange
    Set rng = Application.Selection
    rng.Select

    RowCount = rng.Rows.Count
    ColCount = rng.Columns.Count
    DeleteCount = 0

' Optimize Code
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

' Loop Through Rows & Accumulate Rows to Delete
    For X = RowCount To 1 Step -1
        ' Is Row Not Empty?
        If Application.WorksheetFunction.CountA(rng.Rows(X)) = 0 Then
'       '     If StopAtData = True Then Exit For
'        Else
            If rngDelete Is Nothing Then Set rngDelete = rng.Rows(X)
            Set rngDelete = Union(rngDelete, rng.Rows(X))
            RowDeleteCount = RowDeleteCount + 1
        End If
    Next X

' Delete Rows (if necessary)
    If Not rngDelete Is Nothing Then
        rngDelete.EntireRow.Delete Shift:=xlUp
        Set rngDelete = Nothing
    End If

' Refresh UsedRange (if necessary)
    If RowDeleteCount + ColDeleteCount > 0 Then
        ActiveSheet.UsedRange
    Else
        MsgBox "No blank rows or columns were found!", Title:="Delete empty rows or columns - " & ActiveSheet.Name
    End If

' Error Handler
ExitMacro:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    rng.Cells(1, 1).Select

End Sub
```
