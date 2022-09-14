---
Title: Function to test if the active sheet is protected
categories: [excel, vba]
tags: [validation, developer, function]
date: 2018-11-16

---

Function to tests if the active sheet is protected.

```vb
'============================================================================================
'## Tests if the activesheet is protected
'    ' If the activesheet is protected
'    If ActiveProtected = True Then Exit Sub
'============================================================================================
Function ActiveProtected() As Boolean
    ' If the activesheet is protected then exit sub
    If ActiveSheet.ProtectContents = True Then
        MsgBox "The active sheet is protected" & vbNewLine & _
               "Cannot continue with this procedure"
        ActiveProtected = True
        Exit Function
    End If

    ' Not protected
    ActiveProtected = False
End Function
```
