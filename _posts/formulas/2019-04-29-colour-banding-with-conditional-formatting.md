---
Title: Colour Banding with Conditional Formatting
categories: [Excel, Formulas]
tags: [interface-formatting]
date: 2019-04-29 18:00:00

---


Use Conditional Formatting from the Format menu to apply a format to the cells. The formula used in the Conditional Formatting rule is based solely on the row number, so the formatting it applies will remain intact as you sort the rows or insert or delete rows.

The formatting techniques described here create colour bands of a fixed number of rows, regardless of the content of the cells on worksheet. When banding is applied, the cells will have alternate bands of colour, as shown below:

By using Conditional Formatting rather than directly styling a range, you can prevent the colours from getting out of order when you sort the range or insert or delete rows.

## Colour Banding every second row or column
To apply the colour to every second row or column with the colour in the first cell to be formatted, enter the following formula in the formula bar in the Conditional Formatting dialog.

```vb
'## Rows
=MOD(ROW(),1*2)+1<=1

'## Columns
=MOD(COLUMN(),1*2)+1>1
```

To apply the colour to every second row or column with the colour in the second cell to be formatted, enter the following formula in the formula bar in the Conditional Formatting dialog.

```vb
'## Rows
=MOD(ROW(),1*2)+1>1

'## Columns
=MOD(COLUMN(),1*2)+1<=1
```

## Colour Banding Groups of Rows / Columns
To apply the colour on the first group and every other group, enter the following formula in the formula bar in the Conditional Formatting dialog.

```vb
'## Rows
=MOD(ROW()-Rw,N*2)+1<=N

'## Columns
=MOD(COLUMN()-Col,N*2)+1<=N
```

Where Rw is the row number of the first cell in the rows that are to be formatted, and N is the number of rows in each banded group. For example, if the first row is 12 and you want each band to contain 3 rows, you would use the formula:

```vb
=MOD(ROW()-12,3*2)+1<=3
```
