---
Title: Loop Through Control Types in a Userform
categories: [vba]
tags: [userform]
date: 2017-03-19 18:43:00

---

Snippets to loop through control types or a specific control type. Especially useful for bulk actions.


## Control Type Names

- Label
- TextBox
- ComboBox
- ListBox
- CheckBox
- OptionButton
- ToggleButton
- Frame
- CommandButton
- TabStrip
- MultiPage
- ScrollBar
- SpinButton
- Image


## Loop through all controls

```vb
Sub LoopAllControls()
    '// Vars
    Dim ctrl        As Control
    '// Loop Through each control on UserForm
    For Each ctrl In UserForm1.Controls
        '// Do something with that control type...eg make visible
        ctrl.Visible = TRUE
    Next ctrl
End Sub
```

## Loop through a specific control type

```vb
Sub LoopSpecificControl()
    '// Vars
    Dim ctrl        As Control
    Dim ctrlType    As String

    '// Choose control type to loop through
    ctrlType = "Textbox"

    '// Loop Through each control on UserForm
    For Each ctrl In UserForm1.Controls

        '// Test specific control type
        If TypeName(ctrl) = ctrlType Then
            '// Do Something With That Control Type...eg no value
            ctrl.Value = ""
        End If

    Next ctrl
End Sub
```
