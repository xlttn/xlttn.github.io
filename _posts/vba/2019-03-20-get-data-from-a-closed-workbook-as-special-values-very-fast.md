---
Title: Get Data from a Closed Workbook as Special Values very Fast
categories: [Excel, VBA]
tags: [copy-data]
date: 2019-03-20 21:25:00

---

## Pull data from a Closed Workbook
Pass the Filepath, Filename, Sheet Name, Range and the Sheet that you want to pull the Data into.  
Call the Function, the Range to add the Data to in your ActiveSheet is the same as the Range to pull the Data from the Closed Workbook, but you can easily tweak this.

```vb
' ## Pull data from a Closed Workbook
'    (Pass the Filepath, Filename, Sheet Name, Range and the Sheet that you want to pull the Data into):
' ## Call the Function, the Range to add the Data to in your ActiveSheet is the same as the Range to pull
'    the Data from the Closed Workbook, but you can easily tweak this.
GetValuesFromAClosedWorkbook "C:\users\Paradigm\Desktop", "BloodPressureTracker.xlsx", "Daily Record", "B10:G1000", "Sheet1"

' ## GetValuesFromAClosedWorkbook, retrieves Special Values for data in a Closed Excel Workbook
Private Sub GetValuesFromAClosedWorkbook(ByVal strFilepath As String, ByVal strFilename As String, ByVal strSheet As String, ByVal strRange As String, ByVal strActiveSheet As String)
    With Sheets(strActiveSheet).Range(strRange)
        .FormulaArray = "='" & strFilepath & "\[" & strFilename & "]" & strSheet & "'!" & strRange
        .Value = .Value
    End With
End Sub
```
