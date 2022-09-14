---
Title: Get the nth word from a text string with VBA
categories: [excel, vba]
tags: [text-strings, function]
date: 2019-03-20 21:27:00

---

An easy user defined function (UDF) to get the nth word from cell. I often use this to get first and last names when the full name is in a single cell.


```vb
'==========================================================================================================
' ## Function: Get the nth word from a text string
'==========================================================================================================
Function ExtractNthWord(Source As String, Position As Integer, Delimiter As String) As String
    Dim arr() As String
    Dim lCount As Long

    arr = VBA.Split(Source, Delimiter)
    lCount = UBound(arr)

    If lCount < 1 Or (Position - 1) > lCount Or Position < 0 Then
        ExtractNthWord = ""
    Else
        ExtractNthWord = arr(Position - 1)
    End If

End Function
```
