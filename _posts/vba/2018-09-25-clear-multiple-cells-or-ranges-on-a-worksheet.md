---
Title: Clear data from multiple ranges
categories: [Excel, VBA]
tags: [deleting]
date: 2018-09-25

---

This joins many addresses together as a string, picks up a defined named range "sNamedRange" and clears the contents from the cells.

```vb
'==================================================================================================
' ## Clear data from multiple ranges
'==================================================================================================
Sub ClearReportData()

    ' // vars
    Dim sNamedRange As String
    sNamedRange = Range("sNamedRange").Address

    Dim strCells As String
    strCells = "E6:E8, B14:C18, E14:G18, I14:J14, L14:N14, E23:E25, " & _
               "B31:C35, E31:G35, I31:J35, L31:N35, E40:E42, " & _
               "B48:C52, E48:G52, I48:J52, L48:N52"

    ' // clear the Named Range and the string of addresses
    With Worksheets("Sheet1")
        .Range(strCells).ClearContents
        .Range(sNamedRange).ClearContents
    End With

End Sub

```
