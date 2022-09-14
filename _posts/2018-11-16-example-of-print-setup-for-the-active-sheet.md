---
Title: Example of Print Setup for the active sheet
categories: [vba]
tags: [interface-formatting, printing]
date: 2018-11-16

---
Example of Page Setup for the active sheet for printing example has a header, footer, narrow margins and the Column width to the page but not the rows.

```vb
' ==================================================================
' ## Example of Page Setup for the active sheet  for printing
'	 This example has a header, footer, narrow margins and
'	 fits the Column width to the page but not the rows
' ==================================================================
Sub ActiveSheet_PrintSetup()
	With ActiveSheet.PageSetup
		.PrintArea = ActiveSheet.UsedRange.Address
		.Orientation = xlLandscape
		.Zoom = False
		.FitToPagesWide = 1
		.FitToPagesTall = False
		.RightFooter = "&9" & "PAGE &P OF &N"
		.LeftMargin = Application.InchesToPoints(0.2)
		.RightMargin = Application.InchesToPoints(0.2)
		.TopMargin = Application.InchesToPoints(0.7)
		.BottomMargin = Application.InchesToPoints(0.7)
		.CenterHorizontally = True
	End With
End Sub
```
