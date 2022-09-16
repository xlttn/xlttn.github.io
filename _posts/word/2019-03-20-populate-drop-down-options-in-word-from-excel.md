---
Title: Populate Drop Down options in Word from Excel
categories: [Word]
tags: [practical, copy-data]
date: 2019-03-20 21:20:00

---

Manually typing each option for a drop-down list in word ***(drop-down list content control)*** is time consuming, but there is no native way to import a list you've made in Excel, or exported from a database to a workbook... to Word.

## Getting our List Set Up
There's a couple things we have to do first


### Excel
1. get your list into an excel workbook
2. do not put a header in A1, your first drop-down option goes in A1
3. your list options only in Column A
4. save and close the workbook

### Word
1. create a combo box (drop-down) content control in Word, from the Developer tab
2. select the combo box (drop-down) content control by clicking on the 3 dots on the top left

### VBE Window (Word)
1. VBA Reference to the Excel Object Model is required, via Tools &rarr; References &rarr; Microsoft Excel...
2. rename the workbook path and sheet name to your requirements
3. run the macro

```vb
'===================================================================================================
' ## Get a list from Excel and populate into a Drop Down List control in Word.
'    Do not put a header in A1, put your first drop down option in A1 then fill down.
'    Close the source workbook before operating. It is possible to have it open but then until Word
'    is closed the workbook will open in Read-Only.
'    Select the control by clicking on the 3 dots on the top left then run the macro.
'	 Rename the workbook path and sheet name to your requirements.
'    VBA Reference to the Excel Object Model is required, via Tools|References -> Microsoft Excel...
'===================================================================================================
Sub DropDownFromExcel()

    '// Vars
    Dim xlApp As New Excel.Application, xlWkBk As Excel.Workbook
    Dim wbName As String, shtName As String, LRow As Long, i As Long


    '// Screen updating off
    Application.ScreenUpdating = False

    '// Workbook path
    wbName = "C:\Book1.xlsx"

    '// Sheet name
    shtName = "Sheet1"

    '// Test the workbook exists
    If Dir(wbName) = "" Then
      MsgBox "Cannot find the designated workbook: " & wbName, vbExclamation
      Exit Sub
    End If

    '// Get the list from Excel, column A
    With xlApp
      .Visible = False
      Set xlWkBk = .Workbooks.Open(FileName:=wbName, ReadOnly:=True, AddToMRU:=False)
      With xlWkBk
          With .Worksheets(shtName)
            '// Find the last-used row in column A.
            LRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            '// Populate the selected content control with Column A cell values.
            '   note not to include a header in the spreadsheet.
            Selection.Range.ContentControls(1).DropdownListEntries.Clear
            For i = 1 To LRow
                Selection.Range.ContentControls(1).DropdownListEntries.Add _
                Text:=Trim(.Range("A" & i))
            Next
          End With
        .Close Savechanges:=False
      End With
      .Quit
    End With
    '// Release Excel object memory
    Set xlWkBk = Nothing: Set xlApp = Nothing
    Application.ScreenUpdating = True
End Sub
```
