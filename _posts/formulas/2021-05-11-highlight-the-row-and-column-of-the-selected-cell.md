---
Title: Highlight the Row and Column of the Selected Cell
categories: [Excel, Formulas]
tags: [interface-formatting]
date: 2021-05-11

---

Here's the simple steps to highlight the row and column of the selected cell which can be extremely useful when navigating large sets of data. Here's a little example:

![highlight-row-column](/imgs/highlight-the-row-and-column-of-the-selected-cell/highlight-the-row-and-column-of-the-selected-cell.gif)


**Download the example workbook here:** [Highlight the Row and Column of the Selected Cell.xlsx](/example-files/Highlight the Row and Column of the Selected Cell.xlsm)

1. Select the data set in which you to highlight the active row/column
2. Go to the Home tab
3. Click on Conditional Formatting and then click on New Rule
4. In the New Formatting Rule dialog box, select "Use a formula to determine which cells to format"
5. In the Rule Description field, enter one of the below formulas

```vb
' Highlight Row and Column
=OR(CELL("col")=COLUMN(),CELL("row")=ROW())

' Highlight only the Row
=CELL("row")=ROW()

' Highlight only the Column
=CELL("col")=COLUMN()
```

## A little bit of VBA
Add this code to the sheet level module so that we aren't interfering with anyone copying or cutting data.

```vb
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Application.CutCopyMode = False Then
        Application.Calculate
    End If
End Sub
```

The quickest way to get to the sheet level module is to right click the sheet name and select view code.

1. Go to the Developer tab
2. Click on Visual Basic
3. In the VB Editor, on the left, you will see the project explorer that lists all the open workbooks and the worksheets in it. If you can’t see it, use the keyboard shortcut Control + R.
4. With your workbook, double-click on the sheet name in which you have the data. In this example, the data is in Sheet 1 and Sheet 2.
5. In the code window, copy and paste the above VBA code. You’ll have to copy and paste the code for both sheets if you want this functionality in both sheets.
6. Close the VB Editor
