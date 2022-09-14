---
Title: InputBox Type 2 - get some text from the
categories: [vba]
tags: [inputbox]
date: 2019-03-29 18:40:00

---

InputBox Type 2 - get some text from the user

```vb
'==================================================================================================
' ## InputBox Example 2: (Type 2) - Inputbox to get some text from the user
'==================================================================================================
Sub InputBoxText()
    '// Vars
    Dim YourName As Variant

    '// Prompt for a name
    YourName = Application.InputBox _
        (Prompt:="What's your name?", Title:="Text only", Type:=2)

    '// Test if clicked cancel
    If YourName = False Then Exit Sub

    '// Output
    MsgBox "Your name is " & YourName
End Sub
```
