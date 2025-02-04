---
Title: Force SaveAs UI to only accept Macro Enabled Workbook (xlsm)
categories: [Excel, VBA]
tags: [export]
date: 2018-11-16

---

Forces the SaveAs UI to only accept a XLSM Macro-Enabled Workbook.

```vb
'==========================================================================================================
' ## Forces the SaveAs UI to only accept a XLSM Macro-Enabled Workbook
'    (1) If a workbook is already saved as xlsx then this code is added then clicking save
'       will prompt the user that workbook contains VBA, then bring up the modified SaveAs UI
'    (2) A workbook already saved as a xlsm clicking save will keep as xlsm but choosing SaveAs
'       will prompt the modified SaveAS UI
'    (3) Clicking Save will ignore the below code
'    (4) Place in the ThisWorkbook module
'==========================================================================================================
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim strFileName As String
    ' Check of Save As was used, if so then proceed
    If SaveAsUI = True Then
        Cancel = True

        ' Open modified Save As dialog box with only xlam file option.
        '   Cancel out if user Cancels in the dialog box.
        strFileName = Application.GetSaveAsFilename(, "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", , "Save As XLSM file")
        If strFileName = "False" Then
            Cancel = True
            Exit Sub
        End If

        ' Save the file.
        Application.EnableEvents = False
        ThisWorkbook.SaveAs Filename:=strFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Application.EnableEvents = True

    End If
End Sub
```
