---
Title: Create Sheets from Cell Values
categories: [Excel, VBA]
tags: [practical]
date: 2018-11-16

---

This is one of my favourites and I use it all the time.
Type a bunch of sheet names in a list of cells, select those cells then run this macro.

- Users will be prompted to select a range of cells to create sheet names
- Error checks that the sheet name doesn't already exist - highlights those cells Yellow
- Error checks that the cell doesn't contain illegal sheet name characters \ / * [ ]  : ?
- Error checks that the cell value doesn't have more than the 31 character sheet name limit
- Highlights cells red containing illegal characters or exceeding 31 characters
- Skips cells that are blank
- Checks if the Workbook Protection is on, notifies user and exits sub

```vb
' ======================================================================================================
'## Creates sheet names in the active workbook from selected cell value
'   - Users will be prompted to select a range of cells to create sheet names
'   - Error checks that the sheet name doesn't already exist - highlights those cells Yellow
'   - Error checks that the cell doesn't contain illegal sheet name characters \ / * [ ]  : ?
'   - Error checks that the cell value doesn't have more than the 31 character sheet name limit
'   - Highlights cells red containing illegal characters or exceeding 31 characters
'   - Skips cells that are blank
'   - Checks if the Workbook Protection is on, notifies user and exits sub
'=======================================================================================================
Sub AddSheetFromString()

  ' Vars
    Dim strRng As String
    Dim rngCell As Range
    Dim rngSelection As Range
    Dim blnError As Boolean
    Dim strPrompt As String

    ' Test a range selected
    If TypeName(Selection) <> "Range" Then Exit Sub

    ' Pass the selected cells address to a string
    strRng = ActiveWindow.RangeSelection.Address

    ' Briefly turn off error checking
    On Error Resume Next

    ' Pick the range using an inputbox, the initial range is the currently selected cells
    strPrompt = "Select cells, the values will be the new sheets names" & vbDoubleLine & _
                "Any duplicates will be highlighted yellow, errors highlighted red"
    Set rngSelection = Application.InputBox(strPrompt, "Create Sheets", strRng, Type:=8)

    ' Test if clicked cancel
    If rngSelection Is Nothing Then Exit Sub

    ' Turn on error checking
    On Error GoTo 0

    ' Optimise, turn off Screen Updating
    Application.ScreenUpdating = False

    ' Loop through each cell in the selected range
    For Each rngCell In rngSelection

        ' Test for blank cells, skip these
        If rngCell.Value = "" Then GoTo NextCell

        ' Test if any sheets are named the same as any of the cell values
        '   if so then paint that cell yellow and go to the next cell
        For Each Worksheet In ActiveWorkbook.Worksheets
            If rngCell = Worksheet.Name Then
                rngCell.Interior.Color = vbYellow
                GoTo NextCell
            End If
        Next Worksheet

        ' Test no invalid characters in folder for sheet name and that the
        '   character counr does not exceed the 31 sheet name character limit
        '   Paint the cell red and go to the next cell in the selected range
        If InStr(rngCell, "\") > 0 Or InStr(rngCell, "/") > 0 Or InStr(rngCell, "*") > 0 Or _
            InStr(rngCell, "[") > 0 Or InStr(rngCell, "]") > 0 Or InStr(rngCell, ":") > 0 Or _
            InStr(rngCell, "?") > 0 Or _
            Len(rngCell) > 31 Then
            rngCell.Interior.Color = vbRed
            GoTo NextCell
        End If

        ' Add a worksheet
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(Sheets.Count)

        ' Cell value is a valid sheet name, new sheet is the cell value
        Sheets(Sheets.Count).Name = rngCell.Value
NextCell:
    Next rngCell

    ' Optimise, turn on Screen Updating
    Application.ScreenUpdating = True

End Sub
```
