---
Title: Clean up your data with VBA, using TRIM, CLEAN and SUBSTITUTE
categories: [Excel, VBA]
tags: [text-strings, practical]
date: 2018-09-25

---

What this code allows you to do is circumvent testing (ie looping) each individual cell and handling trimming (removing leading and ending spaces) and cleaning (removing unprintable characters) process for your Excel data. It's a great way to clean up your data getting exported from an outside database.

```vb
'==================================================================================================
' ## Clean up your data with TRIM, CLEAN and SUBSTITUTE using VBA
'==================================================================================================
Sub TrimCells()
    '// Vars
    Dim rngArea             As Range
    Dim rngSelection        As Range
    Dim rngInitialSelect    As Range
    Dim strPrompt           As String

    '// Test a range selected
    If TypeName(Selection) <> "Range" Then Exit Sub

    '// Prompt the user to select a range
    strPrompt = "Select the range to trim cells" & vbNewLine & _
                "Click cancel to quit this task"

    On Error Resume Next
        Set rngInitialSelect = Application.Selection
        Set rngSelection = Application.InputBox(strPrompt, "Trim Cells", rngInitialSelect.Address, Type:=8)
        Set rngSelection = Intersect(rngSelection, _
                        rngSelection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0

    '// Check if user clicked cancel
    If rngSelection Is Nothing Then Exit Sub

    '// Optimise
    Application.ScreenUpdating = False

    '// Trim Clean and use Substiture on cell values
    For Each rngArea In rngSelection.Areas
        rngArea.Value = Evaluate("IF(ROW(" & rngArea.Address & "),CLEAN(TRIM(SUBSTITUTE(" & rngArea.Address & ",CHAR(160)," & " " & "))))")
    Next rngArea

    rngSelection.Select
    '// Optimise
    Application.ScreenUpdating = True
End Sub
```
