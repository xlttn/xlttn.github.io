---
Title: Function to test if the workbook is protected
categories: [excel, vba]
tags: [validation, developer, function]
date: 2018-11-16

---

Function to test if the workbook is protected.   

```vb
=============================================================================================
' ## Tests if the activeworkbook is protected
'    ' Test if the workbook is protected
'    If WorkBookProtected = True Then Exit Sub
'============================================================================================
Function WorkBookProtected() As Boolean
    With ActiveWorkbook
        If .ProtectWindows Or .ProtectStructure Then
            WorkBookProtected = True
            MsgBox "This workbook is protected" & vbNewline & _
                    "Cannot continue with this procedure"
            Exit Function
        End If
    End With

    ' Not protected
    WorkBookProtected = False
End Function
```
