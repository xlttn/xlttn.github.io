---
Title: Inputbox Type 1 get a number from the user
categories: [excel, vba]
tags: [inputbox]
date: 2019-03-29 18:39:00

---

InputBox Example 1: (Type 1) - Inputbox to get a number from the user.

```vb
'==================================================================================================
' ## InputBox Example 1: (Type 1) - Inputbox to get a number from the user
'==================================================================================================
Sub InputBoxNumber()
    '// Vars
    Dim AgeNum As Variant

    '// Prompt for an age
    AgeNum = Application.InputBox _
        (Prompt:="How old are you?", Title:="Numbers only", Type:=1)

    '// Test if clicked cancel
    If AgeNum = False Then Exit Sub

    '// Output
    MsgBox "You are " & AgeNum & " years old."
End Sub
```
