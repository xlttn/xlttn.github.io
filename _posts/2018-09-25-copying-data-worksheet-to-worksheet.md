---
Title: Copy as Special Values Between Worksheets
categories: [excel, vba]
tags: [copy-data]    
date: 2018-09-25

---

This will copy formulas from one place to another and then Copy & Paste the calculated results as Special Values
There are 2 components, one is the long subroutine which does the grunt work of copying and pasting. The 1 line caller is to be in a separate routine, put here the source sheet and range and the destination sheet and range.

```vb
'==================================================================================================
' ## Copy as Special Values Between Worksheets: this is the caller line
'==================================================================================================
Call CopyAndPasteFormulaAsSpecialValues(Sheets("Sheet1"), "A1:B1", Sheets("Sheet1"), "A2:B10")

'==================================================================================================
'	CopyAndPasteFormulaAsSpecialValues, Copy Formulas from one place to another and then
'	Copy & Paste the calculated results as Special Values
'==================================================================================================
Public Sub CopyAndPasteFormulaAsSpecialValues(ByVal SourceWorksheet As Worksheet, _
                                              ByVal FormulaRangeToCopy As String, _
                                              ByVal DestinationWorksheet As Worksheet, _
                                              ByVal RangeToPasteOver As String)

    ' // ensure some error handling to restore events
    On Error GoTo Catch

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' // Copy the initial Formula from one place to another
    SourceWorksheet.Range(FormulaRangeToCopy).Copy _
            Destination:=DestinationWorksheet.Range(RangeToPasteOver)

    ' // Copy & Paste Special Values over the calculated Formula Range
    DestinationWorksheet.Range(RangeToPasteOver).Copy
    DestinationWorksheet.Range(RangeToPasteOver).PasteSpecial _
            Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Catch:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
```
