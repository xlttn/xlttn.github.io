---
Title: 32 and 64 Bit API Declarations for VBA Developers
categories: [Excel, VBA]
tags: [developer]
date: 2018-09-25

---

A whole heap of declarations for 32 and 64 bit operating systems.  

```vb
'// for developers

'// 34 bit declarations
	Private Declare Function FindWindow Lib "User32.dll" Alias "FindWindowA" (ByVal lpszClass As String, ByVal lpszWindow As String) As Long
	Private Declare Function MoveWindow Lib "User32.dll" (ByVal HWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
	Private Declare Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long) As Long
	Private Declare Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
	Private Declare Function SetLayeredWindowAttributes Lib "User32.dll" (ByVal HWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
	Private Declare Function DrawMenuBar Lib "User32.dll" (ByVal HWnd As Long) As Long
	Private Declare Function GetSystemMetrics Lib "User32.dll" (ByVal nIndex As Long) As Long
	Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
	Private Declare Function ReleaseDC Lib "User32.dll" (ByVal HWnd As Long, ByVal hDC As Long) As Long
	Private Declare Function GetDC Lib "User32.dll" (ByVal HWnd As Long) As Long
	Private Declare Function SetTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
	Private Declare Function KillTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long) As Long
	Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFilename As String) As Long
	Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFilename As String) As Long
	Private Declare Function AddFontMemResourceEx Lib "gdi32" (ByVal pbFont As Integer, ByVal cbFont As Integer, ByVal pdv As Integer, ByRef pcFonts As Integer) As Long


'// 64bit API Declarations
	Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
	Private Declare PtrSafe Function MoveWindow Lib "User32.dll" (ByVal HWnd As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal nWidth As LongPtr, ByVal nHeight As LongPtr, ByVal bRepaint As LongPtr) As LongPtr
	Private Declare PtrSafe Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" (ByVal HWnd As LongPtr, ByVal nIndex As LongPtr) As Long
	Private Declare PtrSafe Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal HWnd As LongPtr, ByVal nIndex As LongPtr, ByVal dwNewLongPtr As LongPtr) As LongPtr
	Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "User32.dll" (ByVal HWnd As LongPtr, ByVal crKey As LongPtr, ByVal bAlpha As Byte, ByVal dwFlags As LongPtr) As LongPtr
	Private Declare PtrSafe Function DrawMenuBar Lib "User32.dll" (ByVal HWnd As LongPtr) As LongPtr
	Private Declare PtrSafe Function GetSystemMetrics Lib "USER32" (ByVal nIndex As Long) As Long
	Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
	Private Declare PtrSafe Function ReleaseDC Lib "User32.dll" (ByVal HWnd As LongPtr, ByVal hDC As LongPtr) As LongPtr
	Private Declare PtrSafe Function GetDC Lib "USER32" (ByVal HWnd As LongPtr) As Long
	Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal HWnd As LongPtr, ByVal nIDEvent As LongLong, ByVal uElapse As LongPtr, ByVal lpTimerFunc As LongPtr) As Long
	Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal HWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
	Private Declare PtrSafe Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFilename As String) As LongPtr
	Private Declare PtrSafe Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFilename As String) As LongPtr
	Private Declare PtrSafe Function AddFontMemResourceEx Lib "Gdi32.dll" (ByVal pbFont As LongPtr, ByVal cbFont As Integer, ByVal pdv As Integer, ByRef pcFonts As Integer) As LongPtr

'=================================================================================================================================
' ## Test for 64 or 32 bit
'=================================================================================================================================
#If VBA7 Then
	Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
# Else
	Private Declare Function FindWindow Lib "User32.dll" Alias "FindWindowA" (ByVal lpszClass As String, ByVal lpszWindow As String) As Long
# End if
```
