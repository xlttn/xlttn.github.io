---
Title: Add a row to the Bottom of the Active Table
categories: [Excel, VBA]
tags: [tables]
date: 2018-09-25

---


Add a row at the bottom of the active table.  
This will test that the active cell is in a table, then add a single row to at the bottom.

```vb
'==================================================================================================
' ## Add a row at the bottom of the active table
'    This will test that the active cell is in a table, then add a single row to at the bottom.
'==================================================================================================
Private Sub AddRow_Click()
    '// Vars
    Dim SelectedCell As Range
    Dim TableName As String
    Dim ActiveTable As ListObject

    Set SelectedCell = ActiveCell

    '// Determine if ActiveCell is inside a Table
    On Error GoTo NoTableSelected
    TableName = SelectedCell.ListObject.Name
    Set ActiveTable = ActiveSheet.ListObjects(TableName)
    On Error GoTo 0

    '// Add a row to the bottom of the ActiveTable
    ActiveTable.ListRows.Add AlwaysInsert:=True

    Exit Sub

    '// Error Handling
NoTableSelected:
    MsgBox "Select a cell in a table to insert a row at the bottom"
End Sub
```
