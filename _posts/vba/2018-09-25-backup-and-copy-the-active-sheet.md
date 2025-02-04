---
Title: Backup and Copy the Active Sheet
categories: [Excel, VBA]
tags: [copy-data, developer]
date: 2018-09-25
---

Good if you are running a macro that will affect cell values


```vb
'==================================================================================================
' ## Backup the active sheet. Good if you are running a macro that will affect cell values
'==================================================================================================
Function BackupSheet() As Boolean
    '// vars
    Dim OriginalSheet   As String
    Dim myResponse      As Long

    BackupSheet = False

    If ActiveWorkbook.ProtectWindows Or ActiveWorkbook.ProtectStructure Then
        myResponse = MsgBox("This operation usually creates a backup of the active sheet" & vbDoubleLine & _
                            "The workbook is currently protected and cannot add any sheets." & vbNewLine & _
                            "If you want to continue anyway, please click 'Yes'" & vbDoubleLine & _
                            "Click 'No' to cancel this operation", vbYesNo + vbQuestion, "Yes or No")
        If myResponse = vbNo Then Exit Function
    End If

    '// get the original activesheet name and optimise
    OriginalSheet = ActiveSheet.Name
    Application.ScreenUpdating = False

    '// avoid sheet naming error
    On Error Resume Next

    '// Copy the active sheet
    ActiveSheet.Copy after:=ActiveSheet
    ActiveSheet.Name = OriginalSheet

    '// Activate the original sheet
    Sheets(OriginalSheet).Activate
    BackupSheet = True

    '// optimise
    Application.ScreenUpdating = False
End Function
```
