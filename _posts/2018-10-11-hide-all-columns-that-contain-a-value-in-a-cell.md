---
Title: Hide all Columns that Contain a Value in a Cell
categories: [vba]
tags: [practical-applications]
date: 2018-10-11

---


The following macro will hide all the column containing the word "hide" in each cell in row 1.  Here is a brief description of how the code works:

This macro loops through all the cells in Range("A1:G1") using a For Loop.
The If statement checks the cell’s value to see if it equals "hide".
If the cell value equals "hide" then the cell’s entire column is hidden.


```vb
' =================================================================================================
' ## Hide all Columns that Contain a Value in a Cell
' =================================================================================================
Sub Hide_Columns_Containing_Value()

    '// variables
    Dim c As Range

    '// loop and hide if cell value is "hide"
    For Each c In Range("A1:G1").Cells
        If c.Value = "hide" Then
            c.EntireColumn.Hidden = True ' change to false to unhide the column
        End If
    Next c

End Sub
```
---   

### Unhide all Columns in a Range
This line of code will make the columns visible.
```vb
Range("A1:G1").EntireColumn.Hidden = False
```

### Toggle the Hidden State of a Column

The following line of code will set the hidden property to the opposite of it’s current state. If the column is hidden, it will be made visible (unhidden). If it’s visible, it will be hidden.  
```vb
c.EntireColumn.Hidden = Not c.EntireColumn.Hidden
```
