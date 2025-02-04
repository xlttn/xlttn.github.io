---
Title: Function to get the unique count in a range
categories: [Excel, VBA]
tags: [unique, text-strings, function]
date: 2018-11-16

---

User Defined Function to get the unique count in a range.

```vb
'   This is an Ultra-fast UDF to use in Excel or VBA to derive a Unique Count from a Range.
'   The Code works by building a Dictionary of the Unique values. You can pass the Range as
'   a Worksheet Range or a Defined Name. An example of using the UDF in Excel would be
'   ="Unique SKUS in this Report: "&TEXT(CountUnique($B$9:$B$81880),"#,##0") or ="Unique SKUS
'   in this Report: "&TEXT(CountUnique(FilteredRange),"#,##0"). Don't believe how fast this is? ...
'   then try it for yourself. I use this on over 100,000 Rows of data and even when using the
'   AutoFilter it is still instant:

Public Function CountUnique(CellRange As Range) As Long
    On Error Resume Next

    ' you can choose to set or not set this.  if you set it, then it will fire on event handlers for Cell Selections etc.
    ' Application.Volatile

    '  turn off Screen drawing
    Application.ScreenUpdating = False

    ' Vars
    Dim lngY As Long
    Dim vntData As Variant
    '  use late binding.  uncomment the Dictionary & New Dictionary to use early binding
    Dim objDictionary As Object 'Dictionary

    ' initialise the Dictionary
    Set objDictionary = CreateObject("Scripting.Dictionary") 'New Dictionary
    objDictionary.CompareMode = BinaryCompare

    ' pickup all of the data to perform the Count
    vntData = CellRange

    ' build the Unique Count
    For lngY = 1 To UBound(vntData)
        If vntData(lngY, 1) <> "" And Not objDictionary.Exists(vntData(lngY, 1)) Then
            objDictionary.Add vntData(lngY, 1), 1
        End If
    Next lngY

    ' return the Count
    CountUnique = objDictionary.Count

    ' clean up
    Set objDictionary = Nothing
    Erase vntData

End Function
```
