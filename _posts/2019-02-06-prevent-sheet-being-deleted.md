---
Title: Worksheet event to prevent the sheet being deleted
categories: [vba]
tags: [validation, developer, function]
date: 2019-02-06

---

This prevents any worksheet from being deleted by accident using the right click --> delete method.

Pesky users!

```vb
'==========================================================================================================
' ## Prevent worksheet from being deleted without woorkbook protection
'==========================================================================================================
'// In a normal module, paste this code:
Sub UnprotectBook()
	ThisWorkbook.Unprotect
End Sub

'// Then for every worksheet, right-click, select View Code, and then paste this:
Private Sub Worksheet_Deactivate()
	ThisWorkbook.Protect , True
	Application.OnTime Now, "UnprotectBook"
End Sub
```
