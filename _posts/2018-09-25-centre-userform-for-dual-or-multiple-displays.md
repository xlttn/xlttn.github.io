---
Title: Centre the Userform on the Screen for Multiple Monitors
categories: [vba]
tags: [userform ]
date: 2018-09-25

---

Force the userform to load to the centre of the active excel window, add this code to the userform Initialize event. Perfect for when you have more than 1 active monitor

```vb
'==================================================================================================
' ## Force the userform to load to the centre of the active excel window
'    add this code to the userform Initialize event
'==================================================================================================
Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub
```
