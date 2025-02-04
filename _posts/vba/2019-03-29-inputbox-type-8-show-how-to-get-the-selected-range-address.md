---
Title: InputBox Type 8 - show how to get the selected range address
categories: [Excel, VBA]
tags: [inputbox]
date: 2019-03-29 18:43:00

---

InputBox Type 8 - show how to get the selected range address

```vb
'==================================================================================================
' ## InputBox Example 4: (Type 8) - Inputbox to show how to get the selected range address
'==================================================================================================
Sub InputBoxGetAddress()
    '// Vars
    Dim strRng As String
    Dim rngCell As Range
    Dim rngSelection As Range

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

    '// Show the selected address:
    '   Line 1: Cell Absolute address
    '   Line 2: Cell Relative address
    '   Line 3: Sheet Name and Cell Address
    '   Line 4: Workbook, Sheet Name and Cell Address
    MsgBox rngSelection.Address & vbNewLine & vbNewLine & _
        rngSelection.Address(0, 0) & vbNewLine & vbNewLine & _
        rngSelection.Parent.Name & "!" & rngSelection.Address(0, 0) & _
            vbNewLine & vbNewLine & _
        "[" & rngSelection.Parent.Parent.Name & "]" & rngSelection.Parent.Name & _
            "!" & rngSelection.Address(0, 0)
End Sub
```
