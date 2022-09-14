---
Title: How to include the document version with 'Export to Excel' from Sharepoint
categories: [vba]
tags: [sharepoint]
date: 2019-06-24 18:43:00

---

When you have a SharePoint Document Library and use the ***Export to Excel*** feature, you may have noticed that document versions will not be exported to the workbook. This is the VBA-code created which reads the version-info from your SharePoint-document library and inserts the version-info into your Excel file.

The code includes 2 main sub routines, and a third which is called:

- Sub Routine ***GetVersionInfoFromSP()*** reads the version info from SharePoint, this calls the ***GetCommandText*** routine.
- Sub Routine ***PopulateVersionInfo()*** inserts the version-info into the Excel-file

## Before we run the code...

Check that there is a reference to the MSXML.DLL. To do this, go in the VBA-editor, point to **TOOLS -> REFERENCES** and add a reference to **Microsoft XML, v.4.0.** If the reference is missing you'll get a compile error on this line:

Then the following registry keys may not exist on your machine which are required to fulfill the request.

- [HKEY_CLASSES_ROOT\Msxml2.DOMDocument.4.0]
- [HKEY_CLASSES_ROOT\Msxml2.DOMDocument.4.0\CLSID]
- [HKEY_CLASSES_ROOT\Msxml2.DOMDocument.4.0\CLSID\(Default) = {88D969C0-F192-11D4-A65F-0040963251E5}]

You'll need to download and install the MSXML 4.0 pack from the following source: [http://www.microsoft.com/downloads/en/details.aspx?FamilyID=7f6c0cb4-7a5e-4790-a7cf-9e139e6819c0](http://www.microsoft.com/downloads/en/details.aspx?FamilyID=7f6c0cb4-7a5e-4790-a7cf-9e139e6819c0)

## Let's do this

Frist routine to run is ***GetVersionInfoFromSP()***. The version-info from the SharePoint Document Library will be read into an array.

Then, run the ***PopulateVersionInfo()***. This code will insert the version-info into your Excel file.

```vb
'========================================================================================
'## Get the Version number from Sharepoint with Export to Excel
'========================================================================================

Dim sViewGUID As String
Dim sListGUID As String
Dim sListWeb As String
Dim sarValues() As String

'## Firstly, run the GetVersionInfoFromSP sub routine, which calls
'   the GetCommandText routine
Sub GetVersionInfoFromSP()
	Dim objDoc
	Dim objHTTP
	Dim sGet As String
	Dim viewNodes
	Dim i As Integer

	' get the data connection info
	GetCommandText

	' we add dummy param dt and set it to the current date/time so
	' the result will not be cached.
	sGet = sListWeb & "/owssvr.dll?Cmd=Display&List=" _
		& sListGUID & "&View=" & sViewGUID & "&XMLDATA=TRUE&dt=" & Now
	Set objDoc = CreateObject("Msxml2.DOMDocument")
	Set objHTTP = CreateObject("Msxml2.XMLHTTP")
	objHTTP.Open "GET", sGet, False

	' make the call and get the response from the server
	objHTTP.send
	Set objDoc = objHTTP.responseXML

	objDoc.SetProperty "SelectionLanguage", "XPath"
	Set viewNodes = objDoc.DocumentElement.SelectNodes("//*/*/@ows__UIVersionString")

	' get the version information
	ReDim sarValues(1 To viewNodes.Length, 1 To 1)

	For i = 1 To viewNodes.Length
		sarValues(i, 1) = viewNodes.Item(i - 1).Value
	Next

	Set objHTTP = Nothing
	Set objDoc = Nothing
End Sub

Sub GetCommandText()
	Dim sCmdText As String
	Dim objDoc As New MSXML2.DOMDocument40

	' get the view guid, list guid and url from the Connection object
	sCmdText = ActiveWorkbook.Connections(1).OLEDBConnection.CommandText

	' Set objDoc = CreateObject("Msxml2.DOMDocument.4.0")
	Set objDoc = CreateObject("Msxml2.DOMDocument")

	objDoc.LoadXML sCmdText

	' parse out the items we need to make the query
	sViewGUID = objDoc.SelectSingleNode("//*/VIEWGUID").Text
	sListGUID = objDoc.SelectSingleNode("//*/LISTNAME").Text
	sListWeb = objDoc.SelectSingleNode("//*/LISTWEB").Text

	Set objDoc = Nothing
End Sub

'## Secondly, run the PopulateVersionInfo sub routine, this appends
'   a column with the Version Number, with a date stamp in the header
Sub PopulateVersionInfo()

	Dim lcVersionColumn As ListColumn
	Dim LastColumn As Long

	' get the version information
	GetVersionInfoFromSP

	' add the column and populate
	Set lcVersionColumn = ActiveSheet.ListObjects(1).ListColumns.Add
	lcVersionColumn.DataBodyRange.Select
	lcVersionColumn.Range.Cells(1, 1).Value2 = "Version at " & Format(Now, "d/mm/yy hh:mm")
	lcVersionColumn.DataBodyRange = sarValues

	' Get the end column number
	LastColumn = ActiveSheet.ListObjects(1).ListColumns.Count

	' column fit width
	ActiveSheet.ListObjects(1).ListColumns(LastColumn).Range.Columns.AutoFit

End Sub
```
