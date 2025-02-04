---
Title: Copy a Formula down a Dynamic Column Range
categories: [Excel, VBA]
tags: [copy-data, dynamic-ranges]
date: 2018-09-25

---

Copy Formula down a dynamic Column Range with or without AutoFill or using FillDown
You can use the following methods to Copy Formula down a dynamic Column Range (assumes data is in Column A with a Header & Formulas are in Columns "B:D"):

```vb
Dim lngLastRow As Long
lngLastRow = Range("A" & Rows.Count).End(xlUp).Row
' lngLastRow = [A1048576].End(xlUp).Row

' Copy down Formula with AutoFill in Column B to the last Row
Range("B2").AutoFill Destination:=Range("B2:B" & lngLastRow)

' Copy down Formula with AutoFill in Columns "B:D" to the last Row
Range("B2:D2").AutoFill Destination:=Range("B2:D" & lngLastRow)

' alternative method to Copy down Formula with AutoFill in Column B to the last Row
Range("B2").AutoFill Destination:=Range("B2:B" & lngLastRow)

' Copy down Formula without AutoFill in Column B to the last Row - use the Macro Recorder to get the R1C1 Formula
'  this Formula is entered using the Code and is not already present in the Cell
Range("B2:B" & lngLastRow).FormulaR1C1 = "=ROW(R[-1])&RC[-1]"

' simple - using FillDown to Copy down whatever the Formula is in Cell "B2" down the Column Range
Range("B2:B" & lngLastRow).FillDown

' using FillDown & Formula storage in array so your Formula do not have to be present in the Cells
Dim strFormulas(1 To 2) As Variant
strFormulas(1) = "=A2*9"
strFormulas(2) = "=SUM(A2:B2)"
Range("B2:C2").Formula = strFormulas
Range("B2:C" & lngLastRow).FillDown
```
