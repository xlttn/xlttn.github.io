---
Title: Remove Page Break Lines in All Workbooks
categories: [excel, vba]
tags: [interface-formatting, practical]
date: 2017-03-19 18:43:00

---

Loop through each sheet in all open workbooks and remove the page break lines

```vb
Sub DisablePageBreaks()

    '// vars
    Dim Wb          As Workbook
    Dim Sht         As Worksheet

    '// optimise
    Application.ScreenUpdating = FALSE

    '// loop thruogh sheets and hide page breaks
    For Each Wb In Application.Workbooks
        For Each Sht In WB.Worksheets
            Sht.DisplayPageBreaks = FALSE
        Next Sht
    Next Wb

    Application.ScreenUpdating = TRUE

End Sub
```
