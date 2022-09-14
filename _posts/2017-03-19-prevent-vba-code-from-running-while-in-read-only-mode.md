---
Title: Prevent Vba Code from Running While in Read-only Mode
categories: [vba]
tags: [developer]
date: 2017-03-19 18:43:00

---


```vb
Sub ReadOnlyCheck()
    Dim Answer As Long
    'Check to see if file is Read-Only
    If ThisWorkbook.ReadOnly = TRUE Then Answer = MsgBox("Someone Is currently editing this file And your changes will Not be saved. " _ &
    "Do you wish To proceed?", vbYesNo, "Read-Only File")
    If Answer = vbNo Then Exit Sub
  End If
End Sub
```
