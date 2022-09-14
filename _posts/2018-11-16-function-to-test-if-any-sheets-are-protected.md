---
Title: Function to test if any sheets are protected
categories: [excel, vba]
tags: [validation, developer, function]
date: 2018-11-16

---
Tests if any sheets in the active workbook are protected with or without a password.  

```vb
'============================================================================================
' ## Tests if any sheets in the active workbook are protected with or without a password
'    Insert before your code to validate if any sheets are protected
'    ' Test if any sheets in the activeworkbook protected
'    If AnySheetsProtected = True Then Exit Sub
'============================================================================================
Function AnySheetsProtected() As Boolean
    ' Vars
    Dim sht As Worksheet

    ' Loop through each worksheet in the ActiveWorkbook
    For Each sht In ActiveWorkbook.Sheets

        ' Test for protected sheets
        If sht.ProtectContents = True = True Then
            MsgBox "At least one sheet in this workbook is protected" & vbDoubleLine & _
                    "Cannot continue with this procedure"

            ' Protection detected on at least 1 sheet
            '   Set to True and exit function
            AnySheetsProtected = True
            Exit Function
        End If
    Next sht

    ' No sheets currently are password protected
    AnySheetsProtected = False
End Function
```
