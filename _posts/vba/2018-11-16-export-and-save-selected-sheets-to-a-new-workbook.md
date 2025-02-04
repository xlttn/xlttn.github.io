---
Title: Export and Save Selected Sheets to a New Workbook
categories: [Excel, VBA]
tags: [export]
date: 2018-11-16

---

User can save the new workbook as file types: xlsx, xlsm, xlsb, xls, csv and txt with the following features:
- external links are broken
- formulas that reference sheets outside the sheet in the new workbook are changed to values
- VBA code within modules will not copy to the new workbook however Worksheet codes will copy if the save file type is 'xls', 'xlsm' and 'xlsb'
- option to copy as values, commented out in the code
- cannot export multiple sheets for csv or txt file types, do these individually
- loops through each selected sheet and copy to a new workbook. This way you are able to for any protected sheets in the selection. The most popular use of copying to another workbook is ActiveWindow.SelectedSheets.Copy' however this an error if any of the selected sheets contains a table. This macro bypasses that. Boom!
- Utilises 3 functions: ExtractWord, SelectedSheetNames, IsFileOpen which are documented separately

```vb
' ======================================================================================================
' ## Export Selected Sheets To A New Workbook
'    User can save the new workbook as file types: xlsx, xlsm, xlsb, xls, csv, txt
'    All external links are broken
'    Formulas that reference sheets outside the sheet in the new workbook are changed to values
'    VBA code within modules will not copy to the new workbook however Worksheet codes will copy
'      if the save file type is 'xls', 'xlsm' and 'xlsb'
'    Option to copy as values, commented out in the code
'    Cannot export multiple sheets for csv or txt file types, do these individually
'    Loops through each selected sheet and copy to a new workbook. This way you are able to
'      test for any protected sheets in the selection. The most popular use of copying
'      sheets to another workbook is 'ActiveWindow.SelectedSheets.Copy' however this
'      throws an error if any of the selected sheets contains a table.
'    Utilises 3 functions: ExtractWord, SelectedSheetNames, IsFileOpen which are documented separately
' ======================================================================================================
Sub ExportSelectedSheets()

    ' Vars
    Dim wbOriginal As Workbook
    Dim wbNew As Workbook

    Dim lngResponse As Long
    Dim x As Long
    Dim i As Long
    Dim lngFileFormat As Long
    Dim SelectedCount As Long

    Dim strFileName As String
    Dim strDialogTitle As String
    Dim strFolder As String
    Dim strFormat As String
    Dim SheetNames As String
    Dim strSaveFileName As Variant

    ' Optimize Code
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' Set original workbook
    Set wbOriginal = ActiveWorkbook

    ' Set up Save as dialog box to return correct file path string
    strDialogTitle = "Export Selected Sheets to a New Workbook"

    strSaveFileName = Application.GetSaveAsFilename(InitialFileName:="", _
    filefilter:= _
        "Excel Workbook (xlsx) (*.xlsx), *.xlsx" & _
        ",Macro Enabled Workbook (xlsm) (*.xlsm), *xlsm" & _
        ",Excel Binary Workbook (xlsb) (*.xlsb), *xlsb" & _
        ",Excel 97- Excel 2003 Workbook (xls) (*.xls), *xls" & _
        ",CSV (comma delimited) (*.csv), *csv" & _
        ",Text File (txt) (*.txt), *txt" _
        , Title:=strDialogTitle)

    ' If User Proceeds with saving the new workbook
    If strSaveFileName <> False Then

        ' Get folder path
        strFolder = Left(strSaveFileName, InStrRev(strSaveFileName, "\"))

        ' Get the File Format Number of the selected save file type
        strFormat = LCase(Right(strSaveFileName, Len(strSaveFileName) - InStrRev(strSaveFileName, ".", , 1)))
        Select Case strFormat
            Case "xls": lngFileFormat = 56
            Case "xlsx": lngFileFormat = 51
            Case "xlsm": lngFileFormat = 52
            Case "xlsb": lngFileFormat = 50
            Case "csv": lngFileFormat = 6
            Case "txt": lngFileFormat = -4158
            Case Else: lngFileFormat = 51
        End Select

        ' Test if user selected txt or csv, alert that sheets are to be inidivually exported
        '   and exit sub
        If ActiveWindow.SelectedSheets.Count > 1 Then
            If lngFileFormat = 6 Or lngFileFormat = -4158 Then
                MsgBox "You cannot export multiple sheets for CSV or TXT files as the" & vbNewLine & _
                        "data from each sheet does not get appended to the previous sheet" & vbNewLine & vbNewLine & _
                        "Export these sheets individually"
                GoTo xMyExit
            End If
        End If

        ' Check if Original workbook contains VBA code as VBA will not go to new workbook
        '   Ignore txt or csv file types
        If lngFileFormat = 51 Then
            If Val(Application.Version) >= 12 Then
                If wbOriginal.HasVBProject = True Then
                    lngResponse = MsgBox("There was VBA code found in this workbook. " & vbNewLine & _
                        "If you proceed, the VBA code from Modules will not be included" & vbNewLine & _
                        "in the xlsx new workbook." & vbNewLine & vbNewLine & _
                        "Do you wish to proceed?", vbYesNo, "Do you wish to Proceed?")

                    ' Test user cancels and exit procedure
                    If lngResponse = vbNo Then
                        GoTo xMyExit
                    End If
                End If
            End If
        End If

    SelectedCount = ActiveWindow.SelectedSheets.Count
    If SelectedCount = 1 Then
        ActiveWorkbook.ActiveSheet.Copy
        Set wbNew = ActiveWorkbook
        GoTo FinishedCopying
    End If

    Set wbOriginal = ActiveWorkbook

    ' Get the list of sheet names
    SheetNames = SelectedSheetNames

    ' Select only the active sheet
    ActiveSheet.Select

    ' Loop through each selected sheet and copy to a new workbook
    For i = 1 To SelectedCount
        If i = 1 Then
            wbOriginal.Sheets(ExtractWord(SheetNames, i)).Copy
            Set wbNew = ActiveWorkbook
        Else
            wbOriginal.Sheets(ExtractWord(SheetNames, i)).Copy After:=wbNew.Sheets(Sheets.Count)
        End If
    Next

FinishedCopying:
        ' Break external links in new workbook
        ExternalLinks = wbNew.LinkSources(Type:=xlLinkTypeExcelLinks)
        On Error Resume Next
          For x = 1 To UBound(ExternalLinks)
            wbNew.BreakLink Name:=ExternalLinks(x), Type:=xlLinkTypeExcelLinks
          Next x
        On Error GoTo 0

        ' Formulas to Values
        '    Dim sh As Worksheet
        '    For Each sh In wbNew.Worksheets
        '        sh.Select
        '        With sh.UsedRange
        '            .Cells.Copy
        '            .Cells.PasteSpecial xlPasteValues
        '            .Cells(1).Select
        '        End With
        '        Application.CutCopyMode = False
        '    Next sh

        ' Test workbook doesn't already exists AND open then
        '   Save the new workbook
        strFileName = strSaveFileName
        If IsFileOpen(strFileName) = True Then
            MsgBox "This workbook is currently open" & vbNewLine & _
             "The exported workbook will be named: Export_" & Format(Now, "yymmdd_hhmmss") & "." & strFormat

            wbNew.SaveAs strFolder & "Export_" & Format(Now, "yymmdd_hhmmss"), FileFormat:=lngFileFormat, CreateBackup:=False
            wbNew.Close
            strSaveFileName = "Export_" & Format(Now, "yymmdd_hhmmss") & "." & strFormat
        Else
            wbNew.SaveAs strSaveFileName, FileFormat:=lngFileFormat, CreateBackup:=False
            wbNew.Close
        End If

        ' Open file
        If lngFileFormat = -4158 Then
            ActiveWorkbook.FollowHyperlink (strSaveFileName)
        Else
            Workbooks.Open (strSaveFileName)
        End If


    ' Optimize Code
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True


        Exit Sub
    End If

    ' ERROR HANDLER
xMyExit:

    ' Optimize Code
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True

End Sub
' ======================================================================================================
' ## Function determines whether a the file named by FileName is
'    open by another process. The fuction returns True if the file is open
'    or False if the file is not open. If the file named by FileName does
'    not exist or if FileName is not a valid file name, the result returned
'    if equal to the value of ResultOnBadFile if that parameter is provided.
'    If ResultOnBadFile is not passed in, and FileName does not exist or
'    is an invalid file name, the result is False.
' ======================================================================================================
Public Function IsFileOpen(FileName As String, _
    Optional ResultOnBadFile As Variant) As Variant
    ' Vars
    Dim FileNum As Integer
    Dim ErrNum As Integer
    Dim V As Variant

    On Error Resume Next

    ' If we were passed in an empty string,
    '   there is no file to test so return FALSE.
    If Trim(FileName) = vbNullString Then
        If IsMissing(ResultOnBadFile) = True Then
            IsFileOpen = False
        Else
            IsFileOpen = ResultOnBadFile
        End If
        Exit Function
    End If

    ' If the file doesn't exist, it isn't open
    V = Dir(FileName, vbNormal)
    If IsError(V) = True Then
        ' syntactically bad file name
        If IsMissing(ResultOnBadFile) = True Then
            IsFileOpen = False
        Else
            IsFileOpen = ResultOnBadFile
        End If
        Exit Function
    ElseIf V = vbNullString Then
        ' file doesn't exist.
        If IsMissing(ResultOnBadFile) = True Then
            IsFileOpen = False
        Else
            IsFileOpen = ResultOnBadFile
        End If
        Exit Function
    End If

    FileNum = FreeFile()

    ' Attempt to open the file and lock it.
    Err.Clear
    Open FileName For Input Lock Read As #FileNum
        ErrNum = Err.Number

        ' Close the file.
        Close FileNum
        On Error GoTo 0

        ' Check to see which error occurred
        Select Case ErrNum
            Case 0
                ' No error occurred.
                '   File is NOT already open by another user.
                IsFileOpen = False
            Case 70
                ' Error number for "Permission Denied"
                '   File is already opened by another user
                IsFileOpen = True
            Case Else
                ' Another error occurred. Assume open
                IsFileOpen = True
        End Select
End Function

' ======================================================================================================
' ## Function: Get the nth word from a text string
'    If the word position is a negative number or exceeds the amount
'    of words in the string then the user is notified
' ======================================================================================================
Function ExtractWord(Source As String, Position As Long)
    Dim arr() As String
    arr = VBA.Split(Source, "/")
    xCount = UBound(arr)
    If xCount < 1 Or (Position - 1) > xCount Or Position < 0 Then
        ExtractWord = "You have either entered a number that is more than the total words" & vbLf & _
            "or" & vbLf & _
            "You have entered a negative number"
    Else
        ExtractWord = arr(Position - 1)
    End If
End Function

' ======================================================================================================
' ## Function: create a string from sheet names separates by a forward slash - "/"
'    as this character is not allowed to be used in a sheet name. Use this with the
'    ExtractWord function to get a list of the selected sheets in the active workbook.
'    Then iterate through each selected sheet and copy to a new workbook. This way you
'    can test for any protected sheets in the selection. The most popular use of copying
'    sheets to another workbook is 'ActiveWindow.SelectedSheets.Copy' however this
'    throws an error if any of the selected sheets contains a table.
' ======================================================================================================
Function SelectedSheetNames()
    ' Vars
    Dim SheetList As String
    Dim shtName As Worksheet

    ' Create a string with sheets joined
    For Each shtName In ActiveWindow.SelectedSheets
        SheetList = SheetList & shtName.Name & "/"
    Next shtName

    ' Output the list and take off the last forward slash
    SelectedSheetNames = Left(SheetList, Len(SheetList) - 1)
End Function
```
