---
Title: InputBox Type 0 - get formula from range
categories: [Excel, VBA]
tags: [inputbox]
date: 2019-03-29 18:48:00
---

InputBox Type 0 - get formula from range

```vb
'==================================================================================================
' ## InputBox Example 6: (Type 0) - Inputbox to get formula from range
'==================================================================================================
Sub InputBoxFormula()
    '// Vars
    Dim YourFormula As Variant

    '// Get formula from cell
    YourFormula = Application.InputBox _
        (Prompt:="Get the formula", Title:="Formula example", Type:=0)

    '// Test if clicked cancel
    If YourFormula = False Then Exit Sub

    '// Output
    ActiveCell.Value = YourFormula
End Sub
```
