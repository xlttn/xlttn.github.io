---
Title: InputBox Type 8 - select a range and write values to each cell
categories: [Excel, VBA]
tags: [inputbox]
date: 2019-03-29 18:42:00

---

InputBox Type 8 - select a range and write values to each cell


```vb
'==================================================================================================
' ## InputBox Example 3: (Type 8) - Inputbox to select a range and write values to each cell
'==================================================================================================
Sub InputBoxChangeCells()
    '// Vars
    Dim strRng As String
    Dim rngCell As Range
    Dim rngSelection As Range
    Dim FirstVal As Long

    '// Pass the selected cells address to a string
    strRng = ActiveWindow.RangeSelection.Address

    '// Briefly turn off error checking
    On Error Resume Next

    '// Pick the range using an inputbox
    '   the initial range is the currently selected cells
    Set rngSelection = Application.InputBox("Initial range is currently selected cells", _
                        "Select Range", strRng, Type:=8)

    '// Test if clicked cancel
    If rngSelection Is Nothing Then Exit Sub

    '// Turn on error checking
    On Error GoTo 0

    '// Paint each cell in the selection yellow
    rngSelection.Interior.Color = vbYellow

    '// Loop through selected cell and add a value
    '   First Value will be 1, then each cell add 1
    FirstVal = 0
    For Each rngCell In rngSelection
    FirstVal = FirstVal + 1

        rngCell.Value = FirstVal
    Next rngCell
End Sub
```
