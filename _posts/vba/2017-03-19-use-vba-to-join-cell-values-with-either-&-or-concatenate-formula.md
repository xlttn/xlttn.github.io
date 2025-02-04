---
Title: Use VBA to Join Cell Values With Either & or Concatenate Formula
categories: [Excel, VBA]
tags: [text-strings]
date: 2017-03-19 18:43:00

---


Join multiple string cells using one of 2 options, the code below shows how to choose the two, it boils down to what the user finds easier to read.

=CONCATENATE(A1," ",B1," ",C1)
or
=A1&" "&B1&" "&C1

note: Concatenate formula limitations, 255 string limit and & has none
The code below uses the & option, the code points to comment out:

& Option

xRange.Formula = "=" & sArgs
Concatenate option

xRange.Formula = "=CONCATENATE(" & sArgs & ")"

& option

sArgSep = "&"
Concatenate option

sArgSep = ","

## Description

Creates a formula in the active cell using the selected cells values
Prompts the user for a range of cells, then prompts the user for an argument separator
You can select any range of cells, continuous and or random cells, then type a separator and you’re good to go!
The macro enters in a Concatenate formula to the active cell with the separator between each cell in the range you selected!

```vb
Sub JoinCellsWithVBA()

    '// Vars
    Dim rSelected   As Range
    Dim c           As Range
    Dim xRange      As Range
    Dim sArgs       As String
    Dim sArgSep     As String
    Dim sSeparator  As String
    Dim sTitle      As String
    Dim lTrim       As Long

    '// Check Cell Selected
    If TypeName(Selection) <> "Range" Then Exit Sub

    '// Set variables
    Set xRange = ActiveCell
    sSeparator = ""
    sTitle = "Concatenate With VBA"

    '// Prompt user to select cells for formula
    On Error Resume Next        'Turn off error prompts
    Set rSelected = Application.InputBox(Prompt:="Select cells To join together", Title:=sTitle, Type:=8)
    On Error GoTo 0        'Turn on error prompts

    '// Only run if cells were selected and cancel button was not pressed
    If Not rSelected Is Nothing Then

        '// Set argument separator for formula, choose for either & or Concatentate formula
        '   remember to comment out the appropriate xRange.Formula below
        '& option
        sArgSep = "&"        '& option

        'Concatenate option
        'sArgSep = ","  'Concatenate option

        '// Enter the separator between each cell value, check if user clicked cancel
        xSeparator = Application.InputBox(Prompt:="Enter separator, leave blank If none.", Title:=sTitle, Type:=2)

        If xSeparator = FALSE Then Exit Sub        'If User clicked cancel then exit sub
        sSeparator = xSeparator        'Set the separator as string

        '// Create string of cell references
        For Each c In rSelected.Cells
            sArgs = sArgs & c.Address(0, 0) & sArgSep
            If sSeparator <> "" Then
                sArgs = sArgs & Chr(34) & sSeparator & Chr(34) & sArgSep
            End If
        Next

        '// Trim extra argument separator and separator characters
        lTrim = IIf(sSeparator <> "", 4 + Len(sSeparator), 1)
        sArgs = Left(sArgs, Len(sArgs) - lTrim)

        '// Create formula to the active cell
        '& option, ensure you change sArgSep above to "&"
        xRange.Formula = "=" & sArgs

        'Concatenate option, ensure you change sArgSep above to ","
        'xRange.Formula = "=CONCATENATE(" & sArgs & ")"

    End If
End Sub
```
