---
Title: Adding a Context Menu to a Userform's Textbox
categories: [Excel, VBA]
tags: [userform]
date: 2018-09-25

---

A technique to add a custom right click menu (context menu) to a text box on a UserForm.
There is no system context menu for text boxes in a userform, and the good news is that you can do this for as many text boxes in a userform as you like!
This example adds 4 buttons: Copy, Paste, Cut, Undo

In a standard module we create 4 macros using the sendkeys method for our buttons...this way we keep the undo stack in place.

Paste in a Standard Module
```vb
'==================================================================================================
' ## Step 1: Send keys subroutines
'==================================================================================================
Private Sub myCopy()
    SendKeys "^c"
End Sub

Private Sub myPaste()
    SendKeys "^v"
End Sub

Private Sub myCut()
    SendKeys "^x"
End Sub

Private Sub myUndo()
    SendKeys "^z"
End Sub

'==================================================================================================
' ## Step 2: Next we create a sub routine that adds a temporary control to the Cell CommandBar
'==================================================================================================
Private Sub MyRightClickMenuUserForm()
	Application.CommandBars("Cell").Reset
	Dim cbc As CommandBarControl
	For Each cbc In Application.CommandBars("cell").Controls
		cbc.Visible = False
	Next cbc
	With Application.CommandBars("Cell").Controls.Add(temporary:=True)
		.Caption = "Copy"
		.OnAction = "myCopy"
		.FaceId = 19
	End With
	With Application.CommandBars("Cell").Controls.Add(temporary:=True)
		.Caption = "Paste"
		.OnAction = "myPaste"
		.FaceId = 22
	End With
	With Application.CommandBars("Cell").Controls.Add(temporary:=True)
		.Caption = "Cut"
		.OnAction = "myCut"
		.FaceId = 21
	End With
	With Application.CommandBars("Cell").Controls.Add(temporary:=True)
		.Caption = "Undo"
		.OnAction = "myUndo"
		.FaceId = 128
	End With
	Application.CommandBars("Cell").ShowPopup
End Sub
```

```vb
'==================================================================================================
' ## Step 3: CREATE A USERFORM
'	 For easy understanding I kept the default UserForm control names e.g. UserForm1, TextBox1.
'	 The userform is loaded to the center of the active screen and the text box has some text
'	 loaded just for this purpose.
'==================================================================================================
Private Sub UserForm_Initialize()
	'// Center the screen for multiple displays
	With Me
		.StartUpPosition = 0
		.Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
		.Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
	End With

	'// Start up with a value in the textbox
	Me.TextBox1.Value = "Cut me, Copy me, Paste me or Undo me"
End Sub

'==================================================================================================
' ## Step 4: CREATE A USERFORM
'	 Create a new MouseUp event for the TextBox1 and for each subsequent TextBox that you create.
'==================================================================================================
Private Sub TextBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
	'// Tests that the right clicker on the mouse was clicked
	If Button = 2 Then Run "MyRightClickMenuUserForm"
End Sub
```
