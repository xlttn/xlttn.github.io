---
Title: Delete the selected row in the active table
categories: [Excel, VBA]
tags: [tables, deleting-data]
date: 2018-11-16

---

Delete the selected row in the active table will prompt the user that the selected cell in a table is about to be deleted and to Yes to confirm as deleting with VBA cannot be undone.

```vb
' ======================================================================================================
' ## Delete the selected row in the active table
'    This will prompt the user that the selected cell in a table is about to be deleted and to
'    press Yes to confirm as deleting with VBA cannot be undone.
' ======================================================================================================
Private Sub DeleteRow_Click()
    ' Vars
    Dim RowNumToDelete As Long
    Dim Answer As Variant

    ' Optimise
    Application.ScreenUpdating = False

    ' Get the row number in the table from the active cell address
    '   This will then prompt the user to confirm deleting the row
    '   as it cannot be undone
    With ActiveCell.ListObject
        RowNumToDelete = ActiveCell.Row - .Range.Cells(1).Row

        Answer = MsgBox("This will delete row: " & RowNumToDelete & vbLf & _
            "There is no Undo option, Proceed?", vbYesNo + vbInformation)
        If Answer = vbYes Then
            .ListRows(RowNumToDelete).Delete
        End If
    End With

    ' Optimise
    Application.ScreenUpdating = True
End Sub
```
