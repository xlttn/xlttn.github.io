---
Title: Import csv to excel with paramaters
categories: [excel, vba]
tags: [copy-data]
date: 2019-03-29 18:00:00

---

This uses the StartImportCSV module to call the ImportCSVFile function, with set parameters to import a csv file.
You can muck around with this and use other techinques such as looping through multiple files and append to the bottom.

```vb
'==================================================================================================
' ## Import CSV subroutine example providing parameters
'    parameters: file path, destination workbook name, destination sheet name, starting row (B=2),
'                starting column number and the delimiter
'==================================================================================================
Sub StartImportCSV()
    ImportCSVFile "C:\SourceData.csv", _
                    ActiveWorkbook.Name, _
                    "Sheet1", _
                    4, 2, ","
End Sub

'==================================================================================================
' ## Function to import CSV file data to a workbook
'==================================================================================================
Function ImportCSVFile(ByVal filePath As String, _
                       ByVal wbName As String, _
                       ByVal DestSheet As String, _
                       ByVal ImportToRow As Long, _
                       ByVal StartColumn As Long, _
                       ByVal strDelimiter As String)

    '// vars
    Dim line As String
    Dim arrayOfElements
    Dim element As Variant
    Dim fileCol As Long
    Open filePath For Input As #1                       ' Open file for input
        Do While Not EOF(1)                             ' Loop until end of file
            ImportToRow = ImportToRow + 1
            Line Input #1, line
            arrayOfElements = Split(line, strDelimiter) 'Split the line into the array.
            fileCol = StartColumn

            '// Loop thorugh every element in the array and print to Excelfile
            For Each element In arrayOfElements
                Workbooks(wbName).Sheets(DestSheet).Cells(ImportToRow, fileCol).Value = element
                fileCol = fileCol + 1
            Next
        Loop
    Close #1 ' Close file.
End Function
```
