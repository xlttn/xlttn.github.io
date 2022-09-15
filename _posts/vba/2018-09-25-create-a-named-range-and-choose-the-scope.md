---
Title: Create Named Ranges and Choose the Scope
categories: [Excel, VBA]
tags: [practical]
date: 2018-09-25

---

The below 2 examples create a named range for either a workbook scope or worksheet scope.
Below the subroutines is more information on the differences and advantages of scope.

```vb
'==================================================================================================
' ## Create a Named Range with Workbook Scope
'==================================================================================================
Sub createNamedRangeWB()

    '// Vars
    Dim myWorksheet As Worksheet
    Dim myNamedRange As Range

    '// Declare vars
    Dim myRangeName As String

    '// identify ranges
    Set myWorksheet = ThisWorkbook.Worksheets("Sheet1")
    Set myNamedRange = myWorksheet.Range("A1:C10")

    '// specify defined name
    myRangeName = "namedRangeWBscope"

    '// create named range with workbook scope. Defined name and cell range are as specified
    ThisWorkbook.Names.Add Name:=myRangeName, RefersTo:=myNamedRange

End Sub

'==================================================================================================
' ## Create a Named Range with Worksheet Scope
'==================================================================================================
Sub createNamedRangeWS()

    '// Vars
    Dim myWorksheet As Worksheet
    Dim myNamedRange As Range

    '// Declare vars
    Dim myRangeName As String

    '// identify ranges
    Set myWorksheet = ThisWorkbook.Worksheets("Sheet1")
    Set myNamedRange = myWorksheet.Range("D1:E10")

    '// specify defined name
    myRangeName = "namedRangeWSscope"

    '// create named range with worksheet scope. Defined name and cell range are as specified
    myWorksheet.Names.Add Name:=myRangeName, RefersTo:=myNamedRange

End Sub
```

### Choosing the Scope
Most of my Defined Names are scoped to the Workbook. This is the default scope when defining names in Excel. So why and when do we use Defined Names scoped to the worksheet?

One of the key benefits of using the worksheet scope is the ability to clone a worksheet together with all the worksheet Defined Names. This is ideal if you wish to create a number of similar or identically structured worksheets. For example, if you have a sales worksheet for each month of the year and each worksheet is very similar in structure and nature.

### Issues with Cloning Worksheets
Cloning a worksheet is very quick and simple in Excel. However, this benefit can also create problems when you clone a worksheet which has workbook Defined Names (i.e. names referring to that worksheet but scoped to the workbook) . What happens is that Excel also clones the workbook Defined Names but converts them to worksheet Defined Names. The problem here is that most users do not realise that the workbook Defined Names have been cloned. As the newly cloned worksheet now has a worksheet Defined Name scoped to itself and the workbook also has a workbook Defined Name, the model will, most likely, not work properly. This is because Excel will prioritise the Defined Name scoped to the worksheet over a Defined Name scoped to the workbook, thus all references made to that Defined Name on the newly cloned worksheet will no longer reference the original Defined Name scoped to the workbook.

### Cloning Worksheets with Workbook Names
If however, you do need to clone a worksheet that has workbook names that refer to it, you should immediately check the list of Defined Names and delete all occurrences of worksheet names that were cloned from Defined Names scoped to the workbook.

If you have followed my suggested naming convention for Defined Names, then this is extremely easy because all workbook names will be prefixed with AA_ and all worksheet names will be prefixed with BB_. When you review the list of names in the Name Manager (or Defined Name dialogue box – Excel 2003), you will immediately see, grouped together, all the names that start with AA_ but which are scoped to the worksheet. These can easily be selected and deleted. Then there will be no risk of Excel using the wrong Defined Name.

If you have not used a naming convention, then you will have to go through the list manually and look for names that are scoped both to the workbook as well as the worksheet one by one.

Note: the use of a naming convention also makes it easy to use VBA code to remove invalid names when automatically cloning worksheets!

### When to Use Worksheet Names
There are two cases when I would recommend using worksheet names.

#### Case 1: Common Worksheet Structure
Lets say we have a model that has common characteristics on many of the worksheets which will be referenced by formulas on those worksheets. i.e.

```
Formulas on Sheet1 referencing Defined Names on Sheet1
Formulas on Sheet2 referencing Defined Names on Sheet2
Formulas on Sheet3 referencing Defined Names on Sheet3
```

Then it makes sense to use worksheet names. Examples of such names would be:
- Titles (often showing months, quarters or years)
- Column Labels (usually on the left and may include several levels)
- Standard Helper Columns or Rows (perhaps a row to calculate the year or month number or a column used for validation purposes
- Columns or Rows used by Conditional Formatting or Data validation formulas

In all of these cases, these structural names should be referenced only by formulas on the worksheet to which they are scoped. The key benefit to using Defined Names scoped to the worksheet is that the worksheet can be cloned and all the formulas will continue to work without modification as the Defined Names are cloned with the worksheet.

When I create names for a structural purpose as described above, I usually prefix them with a BBB_ rather than the BB_. This is a way of differentiating them from the worksheet names described under Case 2 below. I can then also use these Defined Names easily in VBA code.

#### Case 2: Common Data Descriptors
Lets say we have a model that contains twelve worksheets representing sales figures by month. Each worksheet contains a column for each of the following:
- Country
- Currency
- Category
- Qty

If we had just an Annual Sales Summary sheet, I would create a Named Range scoped to the workbook for each for these columns. However, with twelve monthly worksheets, I would have to create Named Ranges for each column for each of the twelve worksheets (by prefixing the Defined Name with part of the worksheet name, each name would then be unique). This would require the creation of 48 names which is not only time consuming in the first instance, but will also makes it much harder to maintain the model. What if we had 10 such columns to name?

The solution is to use Named Ranges scoped to the worksheet. This way we can create just one worksheet and set it up the way we want, then simply clone it to create the other eleven month and name them appropriately.

These Named Ranges are not part of the structure of all worksheets in the model but are specific to the data contained on these worksheets. Thus I would prefix them with the simple BB_ to show that they are worksheet names but not part of the common worksheet structure.

### Using Worksheet Names
When using a workbook Defined Name in a formula, you simply use the name exactly as it has been defined. If you use a worksheet Defined Name on the same worksheet that it is scoped to, then it also can also be used exactly as defined. However, if you use a Defined Name in a formula that is scoped to a worksheet other than the one that it is scoped to, then you need to prefix the name with the worksheet name. i.e.

| Name            | Scope    | Formula on Sheet1      |
|:----------------|:---------|:-----------------------|
| AA_TotalRevenue | Workbook | AA_TotalRevenue        |
| BB_TotalRevenue | Sheet1   | BB_TotalRevenue        |
| BB_TotalRevenue | Sheet2   | Sheet1!BB_TotalRevenue |

However, again, when cloning worksheets, it can be problematic if you use names on one worksheet that are scoped to another. Because of this, I always avoid using any worksheet name on a worksheet other than the one to which it is scoped.

The main issue with using worksheet names on a worksheet that the name is not scoped to is that it can get confusing and messy and hard to manage and control. That is what workbook names are for.
