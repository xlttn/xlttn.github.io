---
Title: Extract link address from (insert)hyperlink or hyperlink formula
categories: [excel, vba]
tags: [text, files]
date: 2018-11-07

---

This method that will return the hyperlink text whether it has been created by a formula, or by the Insert/Hyperlink method.

If the former, we merely have to parse the formula; if the latter, we need to iterate through the hyperlinks collection on the worksheet.

The formula will return nothing if there is no hyperlink in cell_ref; change to suit.

```vb
' ## Get the formula url from hyperlink/formula or the insert/hyperlink method
'
Function LinkLocation(rng As Range)

    ' vars
    Dim sFormula As String, sAddress As String
    Dim L As Long
    Dim sHyperlink As Hyperlink, rngHyperlink As Hyperlinks

    ' cell formula
    sFormula = rng.Formula

    ' gets starting position of the file path, also acts as a test if
    ' there is a formula
    L = InStr(1, sFormula, "HYPERLINK(""", vbBinaryCompare)

    ' tests for hyperlink formula and returns the address, if a link
    ' then returns the link location.
    If L > 0 Then
        sAddress = Mid(sFormula, L + 11)
        sAddress = Left(sAddress, InStr(sAddress, """") - 1)
    Else
        Set rngHyperlink = rng.Worksheet.Hyperlinks
        For Each sHyperlink In rngHyperlink
            If sHyperlink.Range = rng Then
                sAddress = sHyperlink.Address
            End If
        Next sHyperlink
    End If

    ' boom, got the hyperlink address
    LinkLocation = sAddress

End Function
```
