---
Title: Create a Table of Contents
categories: [Excel, VBA]
tags: [practical]
date: 2018-1-16

---

This one is pretty sweet. It creates a nicely formatted TOC sheet and a back button on each sheet that links to the TOC.
If you happen to add a sheet, simply click the delete back button then click refresh buttons to simply Refresh!

```vb
' ======================================================================================================
' ## Master macro which calls all of the other Sub Routines
' ======================================================================================================
Sub MasterTOC()
    TOC_DeleteBackButton        ' Deletes Back Buttons on each sheet
    TableOfContents_Create      ' Creates a Table of Contents sheet named 'TableOfContents'
    Contents_Hyperlinks         ' Deletes Back Buttons on the sheets and creates a new button
    ContentsButtons             ' Creates the Refresh and the Delete Back Buttons on the Contents Page
End Sub

' ======================================================================================================
' ## Deletes Back Buttons on each sheet
' ======================================================================================================
Sub DeleteBackButton()
' Vars
    Dim sht As Worksheet
    Dim shp As Shape
    Dim ContentName As String
    Dim ButtonID As String

    ContentName = "TableOfContents"    'Table of Contents Worksheet Name
    ButtonID = "_ContentButton" 'ID to Track Buttons for deletion

' Loop Through Each Worksheet in Workbook
    For Each sht In ActiveWorkbook.Worksheets
        If sht.Name <> ContentName Then
        ' Delete Old Button (if necessary when refreshing)
            For Each shp In sht.Shapes
                If Right(shp.Name, Len(ButtonID)) = ButtonID Then
                    shp.Delete
                    'Exit For
                End If
            Next shp
        End If
    Next sht
End Sub

' ======================================================================================================
' ## Creates a Table of Contents sheet named 'TableOfContents'
' ======================================================================================================
Sub TableOfContents_Create()

    ' Vars
    Dim sht As Worksheet
    Dim Content_sht As Worksheet
    Dim myArray As Variant
    Dim x As Long, y As Long, z As Long
    Dim shtName1 As String, shtName2 As String
    Dim ContentName As String
    Dim shtCount As Long
    Dim ColumnCount As Variant
    Dim shtCountOne As Boolean

    ' Optimize
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ContentName = "TableOfContents"

    ' Delete Contents Sheet if it already exists
    On Error Resume Next
        Worksheets("TableOfContents").Activate
    On Error GoTo 0

    If ActiveSheet.Name = ContentName Then
        myAnswer = MsgBox("A worksheet named [" & ContentName & _
        "] has already been created, would you like to replace it?", vbYesNo)

        ' Did user select No or Cancel?
        If myAnswer <> vbYes Then GoTo ExitSub

        ' Delete old Contents Tab
        If Worksheets.Count = 1 Then
            Worksheets.Add After:=Worksheets(1)
            shtCountOne = True
        End If
        Worksheets(ContentName).Delete
    End If

    ' Count how many Visible sheets there are
    For Each sht In ActiveWorkbook.Worksheets
        If sht.Visible = True Then shtCount = shtCount + 1
    Next sht

    ' Column count, I have commented out the InputBox option
    '    ColumnCount = 3

    ColumnCount = Application.InputBox("You have " & shtCount & _
      " visible worksheets." & vbNewLine & "How many columns " & _
      "would you like to have in your Contents tab?", Type:=2)

    ' Check if user cancelled, uncomment this for the InputBox option
    If TypeName(ColumnCount) = "Boolean" Or ColumnCount < 0 Then GoTo ExitSub

    ' Create New Contents Sheet
    Worksheets.Add Before:=Worksheets(1)

    ' Set variable to Contents Sheet and Rename
    Set Content_sht = ActiveSheet
    Content_sht.Name = ContentName

    ' Create Array list with sheet names (excluding Contents)
    ReDim myArray(1 To shtCount)
    For Each sht In ActiveWorkbook.Worksheets
        If sht.Name <> ContentName And sht.Visible = True Then
            myArray(x + 1) = sht.Name
            x = x + 1
        End If
    Next sht

    ' Alphabetize Sheet Names in Array List
    For x = LBound(myArray) To UBound(myArray)
        For y = x To UBound(myArray)
            If UCase(myArray(y)) < UCase(myArray(x)) Then
                shtName1 = myArray(x)
                shtName2 = myArray(y)
                myArray(x) = shtName2
                myArray(y) = shtName1
            End If
         Next y
    Next x

    ' Create Table of Contents
    x = 1

    For y = 1 To ColumnCount
        For z = 1 To WorksheetFunction.RoundUp(shtCount / ColumnCount, 0)
            If x <= UBound(myArray) Then
                Set sht = Worksheets(myArray(x))
                sht.Activate
                With Content_sht
                    .Hyperlinks.Add .Cells(z + 2, 2 * y), "", _
                    SubAddress:="'" & sht.Name & "'!A1", _
                    TextToDisplay:=sht.Name
                End With
                x = x + 1
            End If
        Next z
    Next y

    ' Select Content Sheet and clean up a little bit
    Content_sht.Activate
    Content_sht.UsedRange.EntireColumn.AutoFit
    ActiveWindow.DisplayGridlines = False

    ' Format Contents Sheet Title
    With Content_sht.Range("B1")
        .Value = "Table of Contents"
        .Font.Bold = True
        .Font.Size = 18
        .Font.Name = "Cambria"
        .Font.Color = RGB(54, 96, 146)
    End With

    With Content_sht.Range("B1:F1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(54, 96, 146)
        .Weight = xlThin
    End With

    Content_sht.Columns("A").ColumnWidth = 3

ExitSub:
    ' Optimize Code
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' If the Table of Contents sheet is the only sheet then create a new sheet.
    '   This is to error handle as the Contents sheet must be deleted then re added
    If shtCountOne = True Then
        MsgBox "A new sheet has been created; as the Table of Contents sheet must be deleted when" & _
                "refreshing and there must always be one sheet in the workbook."
    End If

End Sub

' ======================================================================================================
' ## Deletes Back Buttons on the sheets and creates a new button
'    for each sheet.
' ======================================================================================================
Sub Contents_Hyperlinks()
' Vars
    Dim sht As Worksheet
    Dim shp As Shape
    Dim ContentName As String
    Dim ButtonID As String

    ContentName = "TableOfContents" 'Table of Contents Worksheet Name
    ButtonID = "_ContentButton" 'ID to Track Buttons for deletion

' Loop Through Each Worksheet in Workbook
    For Each sht In ActiveWorkbook.Worksheets

    If sht.Name <> ContentName Then

    ' Delete Old Button (if necessary when refreshing)
        For Each shp In sht.Shapes
            If Right(shp.Name, Len(ButtonID)) = ButtonID Then
                shp.Delete
                Exit For
            End If
        Next shp

    ' Create & Position Shape
        Set shp = sht.Shapes.AddShape(msoShapeRoundedRectangle, 4, 4, 20, 20)

    ' Format Shape
        With shp
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(91, 155, 213) 'Blue
            .Line.Visible = msoFalse

            With .TextFrame2
                .TextRange.Font.Size = 18
                .TextRange.Text = "Ë"
                .TextRange.Font.Bold = True
                .TextRange.Font.Name = "Wingdings 3"
                .TextRange.Font.Fill.ForeColor.RGB = vbWhite

                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
                .VerticalAnchor = msoAnchorMiddle

                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
                .MarginBottom = 0
            End With

    ' Track Shape Name with ID Tag
        .Name = shp.Name & ButtonID

        End With

    ' Assign Hyperlink to Shape
        sht.Hyperlinks.Add shp, "", _
          SubAddress:="'" & ContentName & "'!A1"

    End If

    Next sht

End Sub

' ======================================================================================================
' ## Creates the Refresh and the Delete Back Buttons on the Contents Page
' ======================================================================================================
Sub ContentsButtons()
Dim sht As Worksheet
Dim shp As Shape
Dim ContentName As String
Dim ButtonID As String

    ContentName = "TableOfContents" 'Table of Contents Worksheet Name
    ButtonID = "_RefreshButton" 'ID to Track Buttons for deletion

    ' Delete Old Button (if necessary when refreshing)
        For Each shp In ActiveSheet.Shapes
                shp.Delete
        Next shp

    ' Create Refresh Button
        Set shp = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 657, 37.5, 144, 15)

    ' Format Shape
        With shp
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = vbWhite
            .Line.Visible = True
            .Line.BackColor.RGB = RGB(46, 116, 180) 'Blue

            With .TextFrame2
                .TextRange.Font.Size = 11
                .TextRange.Text = "Refresh"
                .TextRange.Font.Bold = True
                .TextRange.Font.Fill.ForeColor.RGB = RGB(46, 116, 180) 'Blue

                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
                .VerticalAnchor = msoAnchorMiddle

                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
                .MarginBottom = 0
            End With

    ' Track Shape Name with ID Tag, add Macro
            .Name = shp.Name & ButtonID
            .OnAction = "MasterTOC"
        End With

    ' Create Delete Back Button
        Set shp = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 657, 67.5, 144, 15)

    ' Format Shape
        With shp
            .Fill.ForeColor.RGB = vbWhite
            .Line.Visible = True
            .Line.BackColor.RGB = RGB(46, 116, 180) 'Blue

            With .TextFrame2
                .TextRange.Font.Size = 11
                .TextRange.Text = "Delete Back Buttons"
                .TextRange.Font.Bold = True
                .TextRange.Font.Fill.ForeColor.RGB = RGB(46, 116, 180) 'Blue

                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
                .VerticalAnchor = msoAnchorMiddle

                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
                .MarginBottom = 0
            End With

    ' Track Shape Name with ID Tag
            .Name = shp.Name & ButtonID
            .OnAction = "TOC_DeleteBackButton"
        End With
End Sub
```
