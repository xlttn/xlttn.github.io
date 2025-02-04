---
Title: Function to get the column letter from the column number
categories: [Excel, VBA]
tags: [practical, function]  
date: 2018-09-25

---

Obtaining Row information is easy since Rows are always Numbers. Column Letters that can be used in a Range are a little more tricky. I have many methods to obtain a Column Letter from a Column Number here are a few of my favourites:

```vb
'==================================================================================================
' ## Function to get the column letter from the column number
'==================================================================================================
Public Function GetColumnLetter(ByVal MyColumnNumber As Integer) As String
    GetColumnLetter = Left(Cells(1, Int(MyColumnNumber)).Address(1, 0), InStr(1, Cells(1, Int(MyColumnNumber)).Address(1, 0), "$") - 1)
End Function

'==================================================================================================
' ## Examples to return column numbers
'==================================================================================================
' 1. simple inline methods for the ActiveCell or for a Column Number
MsgBox Split(ActiveCell.Address, "$")(1)
MsgBox Split(ActiveCell(1).Address(1, 0), "$")(0)
MsgBox Split(ActiveCell.Address(True, False), "$")(0)
MsgBox Mid(ActiveCell.Address, 2, InStr(2, ActiveCell.Address, "$") - 2)
MsgBox Split(Columns(16384).Address(, False), ":")(1)

'1.1 dynamically find the last Column on the ActiveSheet and convert it to a Letter
MsgBox Split(Columns(Cells.Find(What:="*", SearchDirection:=xlPrevious).Column).Address(, False), ":")(1)
MsgBox Split(Columns(Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column).Address(, False), ":")(1)

' 2. an inline method to obtain the last Column Letter in a Range ie. a Row of Headers
Dim strLastColumn as string
strLastColumn = Split(Cells(1, Range("A1").End(xlToRight).Column).Address(True, False), "$")(0)

' 3. a Function that returns the Column Letter for any Column Number
Dim strColumn as String
strColumn = GetColumnLetter(1)
```
