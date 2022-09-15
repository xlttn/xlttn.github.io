---
Title: Hyperlink a file or folder using the HYPERLINK formula using VBA
categories: [Excel, VBA]
tags: [files]
date: 2019-03-29 18:00:00

---

Hyperlink a file or folder using the HYPERLINK formula  in the active cell using VBA. This method pulls up a file picking dialog to choose a file or folder (depending on macro) and then creates a hyperlink formula in the active cell. This method is far superior to a standard left-click hyperlink file method as links made this way can break if the workbook is saved to another location, look up 'Excel hyperlinks break when workbook saved to another folder' or 'Excel hyperlink base'.<br/>

**Hyperlink formula Syntax:** HYPERLINK(link_location, [friendly_name])<br/>
**Link_location:** (required) the path and file name to the document to be opened.<br/>
**Friendly_name:** (optional) the jump text or numeric value that is displayed in the cell. Friendly_name is displayed in blue and is underlined. If friendly_name is omitted, the cell displays the link_location as the jump text.<br/>

```vb
'==================================================================================================
' ## Hyperlink a file or folder using the HYPERLINK formula in the active cell using VBA
'    Syntax: HYPERLINK(link_location, [friendly_name])
'    Link_location: (required) the path and file name to the document to be opened.
'    Friendly_name: (optional) the jump text or numeric value that is displayed in the cell.
'                   Friendly_name is displayed in blue and is underlined. If friendly_name is
'                   omitted, the cell displays the link_location as the jump text.
'==================================================================================================
Sub HyperlinkFile()
    '// Vars
    Dim xPickedFile As Boolean
    Dim xPickFile As FileDialog
    Dim FullFileName As String
    Dim filename As String
    Dim FileNameNoExt As String

    '// Test a range selected
    If TypeName(Selection) <> "Range" Then Exit Sub

    '// If the activesheet is protected
    If ActiveProtected = True Then Exit Sub

    '// Test if an entire row or column has been selected
    If EntireSelection = True Then Exit Sub

    '// Opens dialog box to Pick File to Hyperlink
    Set xPickFile = Application.FileDialog(msoFileDialogFilePicker)

    '// Set the title and the xPickedFile to False to handle if the user cancels
    With xPickFile
        .Title = "Select file to hyperlink" 'Set title of the dialog box
        xPickedFile = False                 'Set to False
        xPickedFile = .Show                 'Open the file picker
        If xPickedFile Then                 'xPickedFile = True so continue with macro

          '// Picked file as full file path and name
            FullFileName = .SelectedItems(1)

            '// File name with extension
            filename = Right(FullFileName, Len(FullFileName) - InStrRev(FullFileName, "\"))

            '// File name without extension
            FileNameNoExt = Left(filename, (InStr(filename, ".") - 1))

            '// Hyperlink formula for active cell, change the friendly (display) name to either
            '   the full file name, file name or file name and no extension
            ActiveCell.Formula = "=HYPERLINK(""" & FullFileName & """,""" & FileNameNoExt & """)"
        End If
    End With
End Sub


Sub HyperlinkFolder()
    '// Vars
    Dim xPickedFolder As Boolean
    Dim xPickFolder As FileDialog
    Dim myFolder As String
    Dim MyFolderName As String

    '// Test a range selected
    If TypeName(Selection) <> "Range" Then Exit Sub

    '// If the activesheet is protected
    If ActiveProtected = True Then Exit Sub

    '// Test if an entire row or column has been selected
    If EntireSelection = True Then Exit Sub

    '// Opens dialog box to Pick Folder to Hyperlink
    Set xPickFolder = Application.FileDialog(msoFileDialogFolderPicker)

    '// Set the title and the xPickedFile to False to handle if the user cancels
    With xPickFolder
        .Title = "Select folder to hyperlink"   'Set title of the dialog box
        xPickedFolder = False                   'Set to False
        xPickedFolder = .Show                   'Open the folder picker
        If xPickedFolder Then                   'xPickedFolder = True so continue with macro

            '// Picked full folder path
            myFolder = .SelectedItems(1)

            '// Picked folder name from full path
            MyFolderName = Right(myFolder, Len(myFolder) - InStrRev(myFolder, "\"))

            '// Hyperlink formula for active cell, change the friendly (display) name to eithr
            '   the full folder path or only the picked folder name
            ActiveCell.Formula = "=HYPERLINK(""" & myFolder & """,""" & MyFolderName & """)"
        End If
    End With
End Sub
```
