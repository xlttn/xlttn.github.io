---
Title: InputBox - Assign an array of Values to a Range
categories: [Excel, VBA]
tags: [inputbox]
date: 2019-03-29 18:47:00

---
InputBox - Assign an array of Values to a Range

```vb
'==================================================================================================
' ## InputBox Example 5: Inputbox to assign an array of values to a range
'    Tip: have your selection the same size as the array
'    Also...I have no idea how to error handle this after so many attempts.
'==================================================================================================
Sub InputBoxArray()
    '// Vars
    Dim varInput As Variant
    Dim rngSelection As Range

    '// Assign selected range
    Set rngSelection = Application.Selection

    '// Enter an array with curly brackets or select a range
    varInput = Application.InputBox(Prompt:="Enter {1,2} or select range:", Type:=64)

    '// Output
    rngSelection.Value = varInput
End Sub
```
