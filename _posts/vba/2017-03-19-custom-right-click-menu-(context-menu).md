---
Title: Custom Right Click Menu (context Menu)
categories: [Excel, VBA]
tags: [developer, userform]
date: 2017-03-19 18:43:00

---

Easily add a custom right click menu for both normal ranges and in tables! First we need to put two codes into the ThisWorkbook Module, these 2 codes call macros from the Standard Modules which reset the context menu and then add your buttons each time the workbook is activated.
In the Standard Module place the 2 codes that (1) reset the context menu and (2) add to the context menu.

**Add the following To the ThisWorkbook Module**

```vb
Private Sub Workbook_Activate()
    '//  Add to the Context Menu
    Call AddToRightClickMenu
End Sub

Private Sub Workbook_Deactivate()
    '//  Reset the Context Menu
    Call ResetRightClickMenu
End Sub
```

**Reset the context menu code**

```vb
Sub ResetRightClickMenu()
    On Error Resume Next
    '// Reset the context menu for cell ranges
    CommandBars("cell").Reset

    '// Reset the context menu for tables
    Application.CommandBars("List Range Popup").Reset
End Sub
```

**Add the following To a Standard Module**  
Codes to add and reset the context menu.  
Add your Macro buttons to the context menu code.

```vb
Sub AddToRightClickMenu()
    '// vars
    Dim cmdNew      As CommandBarButton

    '// Context Menu For Cells in a Table
    With CommandBars("List Range Popup")
        .Reset
        With .Controls.Add
            .Caption = "UPPER CASE"
            .OnAction = "UpperCase"
            .FaceId = 100
            .BeginGroup = TRUE
        End With
        With .Controls.Add
            .Caption = "Proper Case"
            .OnAction = "ProperCase"
            .FaceId = 95
        End With
        With .Controls.Add
            .Caption = "lower case"
            .OnAction = "LowerCase"
            .FaceId = 91
        End With
        With .Controls.Add
            .Caption = "Sentence case"
            .OnAction = "SentenceCase"
            .FaceId = 98
        End With
    End With

    '// Context Menu for Cells in a Normal Range
    With Application.CommandBars("Cell")
        .Reset
        With .Controls.Add
            .Caption = "UPPER CASE"
            .OnAction = "UpperCase"
            .FaceId = 100
            .BeginGroup = TRUE
        End With
        With .Controls.Add
            .Caption = "Proper Case"
            .OnAction = "ProperCase"
            .FaceId = 95
        End With
        With .Controls.Add
            .Caption = "lower case"
            .OnAction = "LowerCase"
            .FaceId = 91
        End With
        With .Controls.Add
            .Caption = "Sentence case"
            .OnAction = "SentenceCase"
            .FaceId = 98
        End With
    End With
End Sub
```

**Add the Codes that you are adding To the context menu.**  
Note that the name of Each macro MUST be the same As the Text in ‘OnAction’

```vb
Sub UpperCase()
    Dim CaseRange   As Range
    Dim CalcMode    As Long
    Dim cell        As Range
    On Error Resume Next
    Set CaseRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = FALSE
        .EnableEvents = FALSE
    End With
    For Each cell In CaseRange.Cells
        cell.Value = UCase(cell.Value)
    Next cell
    With Application
        .ScreenUpdating = TRUE
        .EnableEvents = TRUE
        .Calculation = CalcMode
    End With
End Sub

Sub LowerCase()
    Dim CaseRange   As Range
    Dim CalcMode    As Long
    Dim cell        As Range
    On Error Resume Next
    Set CaseRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = FALSE
        .EnableEvents = FALSE
    End With
    For Each cell In CaseRange.Cells
        cell.Value = LCase(cell.Value)
    Next cell
    With Application
        .ScreenUpdating = TRUE
        .EnableEvents = TRUE
        .Calculation = CalcMode
    End With
End Sub

Sub ProperCase()
    Dim CaseRange   As Range
    Dim CalcMode    As Long
    Dim cell        As Range
    On Error Resume Next
    Set CaseRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = FALSE
        .EnableEvents = FALSE
    End With
    For Each cell In CaseRange.Cells
        cell.Value = StrConv(cell.Value, vbProperCase)
    Next cell
    With Application
        .ScreenUpdating = TRUE
        .EnableEvents = TRUE
        .Calculation = CalcMode
    End With
End Sub

Sub SentenceCase()
    Dim Rng         As Range
    Dim WorkRng     As Range
    On Error Resume Next
    Set WorkRng = Application.Selection
    For Each Rng In WorkRng
        xValue = Rng.Value
        xStart = TRUE
        For I = 1 To VBA.Len(xValue)
            Ch = Mid(xValue, I, 1)
            Select Case Ch
                Case "."
                    xStart = TRUE
                Case "?"
                    xStart = TRUE
                Case "a" To "z"
                    If xStart Then
                        Ch = UCase(Ch)
                        xStart = FALSE
                    End If
                Case "A" To "Z"
                    If xStart Then
                        xStart = FALSE
                    Else
                        Ch = LCase(Ch)
                    End If
            End Select
            Mid(xValue, I, 1) = Ch
        Next
        Rng.Value = xValue
    Next
End Sub
```
