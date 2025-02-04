---
Title: Function to get the selected sheet names with examples
categories: [Excel, VBA]
tags: [validation, developer, function]
date: 2018-11-16

---
Function to get the selected sheet names with examples.  

1. Create a string from sheet names separates by a forward slash - "/" this character is not allowed to be used in a sheet name.
2. Use this with the ExtractWord function to get a list of the selected sheets in the active workbook.
3. Then iterate through each selected sheet and copy to a new workbook.

Looping through each sheet means you can test for any protected sheets in the selection.
The most popular use of copying to another workbook is 'ActiveWindow.SelectedSheets.Copy' however this an error if any of the selected sheets contains a table.

```vb
'==========================================================================================================
' ## Function: create a string from sheet names separates by a forward slash - "/"
'    as this character is not allowed to be used in a sheet name. Use this with the
'    ExtractWord function to get a list of the selected sheets in the active workbook.
'    Then iterate through each selected sheet and copy to a new workbook. This way you
'    can test for any protected sheets in the selection. The most popular use of copying
'    sheets to another workbook is 'ActiveWindow.SelectedSheets.Copy' however this
'    throws an error if any of the selected sheets contains a table.
'==========================================================================================================
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
'==========================================================================================================
' # Test the 'Selected Sheet Names' Function
Sub ListSelectedSheets()
    MsgBox SelectedSheetNames
End Sub
'==========================================================================================================
' # Loop through each selected sheet
Sub LoopSelectedSheetsExample1()
    ' Vars
    Dim i As Long
    Dim arr() As String

    ' Loop through each selected sheet and msgbox each sheet name
    For i = 0 To ActiveWindow.SelectedSheets.Count - 1
        arr = VBA.Split(SelectedSheetNames, "/")
        MsgBox arr(i)
    Next
End Sub
'==========================================================================================================
' # Same as example 1 but using the ExtractWord function
Sub LoopSelectedSheetsExample2()
    ' Vars
    Dim i As Long
    Dim arr() As String

    ' Loop through each selected sheet and msgbox each sheet name
    For i = 1 To ActiveWindow.SelectedSheets.Count
        MsgBox ExtractWord(SelectedSheetNames, i)
    Next
End Sub
'==========================================================================================================
' # Loop through each selected sheet and rename them
Sub LoopSelectedSheetsExample3()
    ' Vars
    Dim i As Long
    Dim SelectedCount As Long
    Dim SheetNames As String

    ' Optimise
    Application.ScreenUpdating = False

    ' Get count of selected sheets as will have to deselect before copying
    '   If count is 1 then rename this one then exit
    SelectedCount = ActiveWindow.SelectedSheets.Count
    If SelectedCount = 1 Then
        ActiveSheet.Name = "RenamedSheet-1"
        Exit Sub
    End If

    ' Get the list of sheet names
    SheetNames = SelectedSheetNames

    ' Select only the active sheet
    ActiveSheet.Select

    ' Loop through each selected sheet and copy to a new workbook
    For i = 1 To SelectedCount
        Sheets(ExtractWord(SheetNames, i)).Name = "RenamedSheet-" & i
    Next

    ' Optimise
    Application.ScreenUpdating = True
End Sub
'==========================================================================================================
' # Loop through each selected sheet and copy to new workbook
Sub LoopSelectedSheetsExample4()
    ' Vars
    Dim i As Long
    Dim SelectedCount As Long
    Dim SheetNames As String
    Dim wbOriginal As Workbook
    Dim wbNew As Workbook

    ' Optimise
    Application.ScreenUpdating = False

    ' Get count of selected sheets as will have to deselect before copying
    '   If count is 1 then copy this one then exit
    SelectedCount = ActiveWindow.SelectedSheets.Count
    If SelectedCount = 1 Then
        ActiveWorkbook.ActiveSheet.Copy
        Set wbNew = ActiveWorkbook
        Exit Sub
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

    ' Optimise
    Application.ScreenUpdating = True
End Sub
```
