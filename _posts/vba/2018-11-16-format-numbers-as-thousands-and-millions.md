---
Title: Format Numbers as Thousands and Millions
categories: [Excel, VBA]
tags: [practical, interface-formatting]
date: 2018-11-16

---

When you want to just show numbers as 9.1 M or 830 K

```vb
'==================================================================================================
' ## Format Millions with 1 decimal. focuses on constants and formulas
'==================================================================================================
Sub FormatMillions()
    Dim rng As Range

    Set rng = Application.Selection
    If rng.Count = 1 Then
        rng.Select
        rng.NumberFormat = "[>=100000]#,##0.0,,"" M"";[<=-100000]-#,##0.0,,"" M"";0"
        Exit Sub
    End If

    On Error GoTo checkConstantsFormulas
        Set rng = Union(rng.Cells.SpecialCells(xlCellTypeConstants), _
            rng.Cells.SpecialCells(xlCellTypeFormulas))
        rng.NumberFormat = "[>=100000]#,##0.0,,"" M"";[<=-100000]-#,##0.0,,"" M"";0"
        rng.Select
    Exit Sub

ConstantType:
    On Error GoTo checkFormulas
        Set rng = rng.Cells.SpecialCells(xlCellTypeConstants)
        rng.NumberFormat = "[>=100000]#,##0.0,,"" M"";[<=-100000]-#,##0.0,,"" M"";0"
        rng.Select
    Exit Sub

FormulasType:
        Set rng = rng.SpecialCells(xlCellTypeFormulas)
        rng.NumberFormat = "[>=100000]#,##0.0,,"" M"";[<=-100000]-#,##0.0,,"" M"";0"
        rng.Select
    ' Exit Sub
    Exit Sub

checkConstantsFormulas:
    Resume ConstantType

checkFormulas:
    Resume FormulasType
End Sub

'==================================================================================================
' ## Format Thousands with 1 decimal. focuses on constants and formulas
'==================================================================================================
Sub FormatThousands()
    Dim rng As Range

    Set rng = Application.Selection
    If rng.Count = 1 Then
        rng.Select
        rng.NumberFormat = "[>=1000]#,##0.0,"" K"";[<=-1000]-#,##0.0,"" K"";0"
        Exit Sub
    End If

    On Error GoTo checkConstantsFormulas
        Set rng = Union(rng.Cells.SpecialCells(xlCellTypeConstants), _
            rng.Cells.SpecialCells(xlCellTypeFormulas))
        rng.NumberFormat = "[>=1000]#,##0.0,"" K"";[<=-1000]-#,##0.0,"" K"";0"
        rng.Select
    Exit Sub

ConstantType:
    On Error GoTo checkFormulas
        Set rng = rng.Cells.SpecialCells(xlCellTypeConstants)
        rng.NumberFormat = "[>=1000]#,##0.0,"" K"";[<=-1000]-#,##0.0,"" K"";0"
        rng.Select
    Exit Sub

FormulasType:
        Set rng = rng.SpecialCells(xlCellTypeFormulas)
        rng.NumberFormat = "[>=1000]#,##0.0,"" K"";[<=-1000]-#,##0.0,"" K"";0"
        rng.Select
    ' Exit Sub
    Exit Sub

checkConstantsFormulas:
    Resume ConstantType

checkFormulas:
    Resume FormulasType
End Sub
```
