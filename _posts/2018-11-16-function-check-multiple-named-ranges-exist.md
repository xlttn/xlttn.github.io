---
Title: Function Check multiple named ranges exist
categories: [excel, vba]
tags: [developer, function]
date: 2018-11-16

---
Function to check if multiple named ranges exist  

```vb
'==============================================================================
' ## Function: Check multiple named ranges exist
'==============================================================================
Function RangeExists(s As String) As Boolean
    On Error GoTo myExit
    RangeExists = Range(s).Count > 0
myExit:
End Function

' Function example:
Sub CheckRanges()
    Dim vNames As Variant, v As Variant

    vNames = Array("ABC", "DEF", "GHI", "JKL", "MNO")
    For Each v In vNames
        Debug.Print v, RangeExists(CStr(v))
    Next
End Sub
```
