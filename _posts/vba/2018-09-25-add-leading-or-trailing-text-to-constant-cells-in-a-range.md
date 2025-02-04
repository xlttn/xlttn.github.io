---
Title: Add leading text to all the cells in a selected range
categories: [Excel, VBA]
tags: [text-strings]
date: 2018-09-25

---

You will be prompted with an input-box for the range of cells and for the leading text.

```vb
'==================================================================================================
' ## Add leading text to all the cells in a selected range
'    You will be prompted with a input box for the range of cells
'    and for the leading text.
'==================================================================================================
Sub LeadingText()
    '// Vars
    Dim rngCell         As Range
    Dim rngSelection    As Range
    Dim rngInitial      As Range
    Dim strLeading      As Variant
    Dim strPrompt       As String

    '// Test a range selected
    If TypeName(Selection) <> "Range" Then Exit Sub

    '// Prompt the user to select a range
    strPrompt = "Select the range to add leading text" & vbDoubleLine & _
                "Click cancel to quit this task"

    On Error Resume Next
        Set rngInitial = Application.Selection
        Set rngSelection = Application.InputBox(strPrompt, "Add Leading Text to Cell Values", rngInitial.Address, Type:=8)
    On Error GoTo 0

    '// Check if user clicked cancel
    If rngSelection Is Nothing Then Exit Sub

    '// Constants only
    Set rngSelection = Intersect(rngSelection, _
                       rngSelection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))

    '// Enter the separator between each cell value, check if user clicked cancel
    strLeading = Application.InputBox(Prompt:="Enter the leading text, don't forget the space", _
                                        Title:="Add Leading Text to Cell Values", Type:=2)

    '// If User clicked cancel then exit sub
    If strLeading = False Then Exit Sub

    '// Screen updating off, Add leading text then back on
    Application.ScreenUpdating = False

    For Each rngCell In rngSelection
        If rngCell.Value <> "" Then rngCell.Value = strLeading & rngCell.Value
    Next rngCell

    Application.ScreenUpdating = True
End Sub

'==================================================================================================
' ## Add trailing text to all the cells in a selected range
'    You will be prompted with a inputbox for the range of cells
'    and for the trailing text.
'==================================================================================================
Sub TrailingText()
    '// Vars
    Dim rngCell         As Range
    Dim rngSelection    As Range
    Dim rngInitial      As Range
    Dim strTrailing      As Variant
    Dim strPrompt       As String

    '// Test a range selected
    If TypeName(Selection) <> "Range" Then Exit Sub

    '// Prompt the user to select a range
    strPrompt = "Select the range to add trailing text" & vbDoubleLine & _
                "Click cancel to quit this task"

    On Error Resume Next
        Set rngInitial = Application.Selection
        Set rngSelection = Application.InputBox(strPrompt, "Add Trailing Text to Cell Values", rngInitial.Address, Type:=8)
    On Error GoTo 0

    '// Check if user clicked cancel
    If rngSelection Is Nothing Then Exit Sub

    '// Constants only
    Set rngSelection = Intersect(rngSelection, _
                       rngSelection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))

    '// Enter the separator between each cell value, check if user clicked cancel
    strTrailing = Application.InputBox(Prompt:="Enter the trailing text, don't forget the space", _
                                        Title:="Add Trailing Text to Cell Values", Type:=2)

    '// If User clicked cancel then exit sub
    If strTrailing = False Then Exit Sub

    '// Screen updating off, Add trailing text then back on
    Application.ScreenUpdating = False

    For Each rngCell In rngSelection
        If rngCell.Value <> "" Then rngCell.Value = rngCell.Value & strTrailing
    Next rngCell

    Application.ScreenUpdating = True
End Sub
```
