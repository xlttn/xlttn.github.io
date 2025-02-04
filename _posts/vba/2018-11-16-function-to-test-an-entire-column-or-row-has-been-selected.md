---
Title: Function to test an entire column or row has been selected
categories: [Excel, VBA]
tags: [validation, developer, function]
date: 2018-11-16

---
Function to test an entire column or row has been selected.

Using ths function before commencing your macro is a great idea if the macro is using the selection as a variable range.  

```vb
'============================================================================================
' ## Checks whether an entire column or row has been selected and notifies the user
'	 ' Test if an entire row or column has been selected
'	 If EntireSelection = True Then Exit Sub
'============================================================================================
Function EntireSelection() As Boolean
	' Vars
	Dim blnEntireColumn As Boolean
	Dim blnEntireRow As Boolean

	' Check currently selecting a cell, if not then exit the procedure.
	With Selection
		blnEntireColumn = .Address = .EntireColumn.Address
		blnEntireRow = .Address = .EntireRow.Address
	End With

	If blnEntireColumn Then
		MsgBox "Entire column(s) selected" & vbDoubleLine & _
		"Select a specific range only as selecting" & vbNewLine & _
		"an entire column can cause an error"
		' sets the boolean ExitAll to True, macro will now alert the User
		EntireSelection = True
		Exit Function
	ElseIf blnEntireRow Then
		MsgBox "Entire row(s) selected" & vbDoubleLine & _
		"Select a specific range only as selecting" & vbNewLine & _
		"an entire row can cause an error"
		' sets the boolean ExitAll to True, macro will now alert the User
		EntireSelection = True
		Exit Function
	End If

	' Not an entire row or column selected
	EntireSelection = False
End Function
```
