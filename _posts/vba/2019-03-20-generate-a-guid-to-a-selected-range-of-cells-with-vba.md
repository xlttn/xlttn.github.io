---
Title: Generate a GUID to a selected range of cells with VBA
categories: [Excel, VBA]
tags: [unique, text-strings]
date: 2019-03-20 21:00:00 

---
A GUID (or UUID) is an acronym for 'Globally Unique Identifier' (or 'Universally Unique Identifier'). It is a 128-bit integer number used to identify resources. The term GUID is generally used by developers working with Microsoft technologies, while UUID is used everywhere else.

How unique is a GUID?

128-bits is big enough and the generation algorithm is unique enough that if 1,000,000,000 GUIDs per second were generated for 1 year the probability of a duplicate would be only 50%. Or if every human on Earth generated 600,000,000 GUIDs there would only be a 50% probability of a duplicate.

```vb
'==================================================================================================
' ## Declarations for the GUID type and for Windows API
'==================================================================================================
    Private Type GUID_TYPE
        '// Vars
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(7) As Byte
    End Type

    '// Test for 32 or 64 bit Excel
    #If VBA7 Then
        Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (guid As GUID_TYPE) As LongPtr
        Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr
    #Else
        Private Declare Function CoCreateGuid Lib "ole32.dll" (guid As GUID_TYPE) As Long
        Private Declare Function StringFromGUID2 Lib "ole32.dll" (guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As Long
    #End If
'==================================================================================================
' ## Function to call Windows API and grab a GUID
'    Using this method as July 2017 Windows 10 security update
'    throws a permission denied error trying to use: CreateObject("Scriptlet.TypeLib")
'==================================================================================================
Function CreateGuidString(Optional AddHyphens As Boolean, _
                          Optional AddBraces As Boolean) _
                          As String
    '// Vars
    Dim guid As GUID_TYPE
    Dim strGuid As String
    Dim retValue As LongPtr

    '// registry GUID format with null
    '   terminator {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}
    Const guidLength As Long = 39

    retValue = CoCreateGuid(guid)

    '// Get the raw GUID which includes braces and hyphens
    If retValue = 0 Then
        strGuid = String$(guidLength, vbNullChar)
        retValue = StringFromGUID2(guid, StrPtr(strGuid), guidLength)
        If retValue = guidLength Then
            CreateGuidString = strGuid
        End If
    End If

    '// If AddHyphens is switched from the default True to False,
    '   remove them from the GUID
    If Not AddHyphens Then
        CreateGuidString = Replace(CreateGuidString, "-", vbNullString, Compare:=vbTextCompare)
    End If

    '// If AddBraces is True from the default False to True,
    '   leave those curly braces be!
    If Not AddBraces Then
        CreateGuidString = Replace(CreateGuidString, "{", vbNullString, Compare:=vbTextCompare)
        CreateGuidString = Replace(CreateGuidString, "}", vbNullString, Compare:=vbTextCompare)
    End If
End Function
'==================================================================================================
' ## Insert a GUID to the selected cells
'    This example has both braces and hyphens set to true
'==================================================================================================
Sub GuidToCell()
    '// Vars
    Dim rngCell As Range
    Dim rngSelection As Range

    '// Loop through selection and apply GUID to each cell
    Set rngSelection = Application.Selection

    For Each rngCell In rngSelection
        rngCell.Value = CreateGuidString(True, True)
    Next rngCell
End Sub
```
