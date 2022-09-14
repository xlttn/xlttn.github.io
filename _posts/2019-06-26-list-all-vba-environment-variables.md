---
Title: List All VBA Environment Variables
categories: [vba]
tags: [developer]
date: 2019-06-25 18:43:00

---

This VBA Environment function grabs information about your operating system and returns the information as a string. The Environ function is useful for customizing your macros so they behave differently based on your user’s operating system configuration.

To list all the operating system environment variables on your computer, run this VBA Environ macro:

```vb
Sub AllEnvironVariables()
    Dim strEnviron As String
    Dim VarSplit As Variant
    Dim i As Long
    For i = 1 To 255
        strEnviron = Environ$(i)
        If LenB(strEnviron) = 0& Then GoTo TryNext:
        VarSplit = Split(strEnviron, "=")
        If UBound(VarSplit) > 1 Then Stop
        Range("A" & Range("A" & Rows.Count).End(xlUp).Row + 1).Value = i
        Range("B" & Range("B" & Rows.Count).End(xlUp).Row + 1).Value = VarSplit(0)
        Range("C" & Range("C" & Rows.Count).End(xlUp).Row + 1).Value = VarSplit(1)
TryNext:
    Next
End Sub
```
