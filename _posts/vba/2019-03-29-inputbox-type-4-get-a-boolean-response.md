---
Title: InputBox Type 4 - get a boolean response
categories: [Excel, VBA]
tags: [inputbox]
date: 2019-03-29 18:49:00

---

InputBox Type 4 - get a boolean response

```vb
'==================================================================================================
' ## InputBox Example 7: (Type 4) - Inputbox to get a boolean response
'==================================================================================================
Sub InputBoxBoolean()
    '// Vars
    Dim blnAns As Boolean

    '// Boolean prompt, 1 = true, 0 = false
    blnAns = Application.InputBox(Prompt:="Acceptable Answers: 1/True or 0/False" & vbCr & _
            "1 = True, 0 = False", Title:="Boolean Example", Type:=4)

    '// Test for True or False
    If blnAns = True Then
        MsgBox "You have responded with 1 / True"
    Else
        MsgBox "You have responded with 0 / False"
    End If
End Sub
```
