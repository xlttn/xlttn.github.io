---
Title: Get Column Letter and Row Number from Both Parts of a Range
categories: [Excel, VBA]
tags: [copy-data]
date: 2019-03-20 21:21:00

---
Here are some basic methods to get a Column Letter & Row Number from both parts of a Range ie. "A1:B2". The examples below will result in Column "A", Row "1", Column "B" & Row "2". This can be very useful when building dynamic Ranges:

```vb
' // select a Range for the Example
Range("A1:B2").Select

' // identify the Column for the first part of the Range following the ":"
MsgBox Split(Split(Selection.Address, ":")(0), "$")(1)

' // identify the Row for the first part of the Range following the ":"
MsgBox Split(Split(Selection.Address, ":")(0), "$")(2)

' // identify the Column for the second part of the Range following the ":"
MsgBox Split(Split(Selection.Address, ":")(1), "$")(1)

' // identify the Row for the second part of the Range following the ":"
MsgBox Split(Split(Selection.Address, ":")(1), "$")(2)
```
