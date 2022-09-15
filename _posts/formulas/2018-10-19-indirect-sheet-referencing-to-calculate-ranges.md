---
Title: INDIRECT Function Using Sheet References
categories: [Excel, Formulas]
tags: [tables, lookup-reference, summing-counting]
date: 2018-10-19

---

The INDIRECT function returns a reference to a range. You can use this function to create a reference that won't change if row or columns are inserted in the worksheet. Or, use it to create a reference from letters and numbers in other cells.

Download the example workbook here: [Indirect Sheet Referencing to Calculate Ranges.xlsx](https://github.com/ExcelTitan/Excel_Formulas/raw/master/indirect-sheet-referencing-to-calculate-ranges.xlsx)  

#### What does it do?
Returns a reference to a cell, or a range of cells of a sheet.

#### Formula breakdown:
```vb
INDIRECT(ref_text, [a1])

INDIRECT(Return the referenced range of a sheet, Omit if the reference is an A1 style or enter FALSE if it is a R1C1 style)
```

Each month has the same data table structure for the same for Sales People
- Hugh Raye
- Justin Thyme
- Rick O’Shea
- Jacques Strap

### Reference a Specific cell
In Cell I2 of each Month is the total Sales for that month. This formula example references the cell value to get that sheet's I2 value

| ~     | **A**     | **B**             | **C**                |
|-------|-----------|-------------------|----------------------|
| **1** | **Month** | **Cell I2 Value** | **Formula**          |
| **2** | January   | 2,718,086         | =INDIRECT(B13&"!I2") |
| **3** | February  | 2,584,131         | =INDIRECT(B14&"!I2") |
| **4** | March     | 2,829,198         | =INDIRECT(B15&"!I2") |

### Reference a Range
Now, instead of relying on cell I2 in each sheet to have already calculated the Sales for that month, include the calculation in the INDIRECT function

| ~     | **A**     | **B**           | **C**                      |
|-------|-----------|-----------------|----------------------------|
| **1** | **Month** | **Total Sales** | **Formula**                |
| **2** | January   | 2,718,086       | =SUM(INDIRECT(B19&"!D:D")) |
| **3** | February  | 2,584,131       | =SUM(INDIRECT(B20&"!D:D")) |
| **4** | March     | 2,829,198       | =SUM(INDIRECT(B21&"!D:D")) |

### Reference Table Names
You have to type the table Name and wrap the column header in square brackets like below. The easiest way to get that is by testing in a spare cell type = then select the table column to get the syntax

| ~     | **A**                        | **B**           | **C**               |
|-------|------------------------------|-----------------|---------------------|
| **1** | **Table / Column Reference** | **Total Sales** | **Formula**         |
| **2** | JanTable[Sales]              | 2,718,086       | =SUM(INDIRECT(B28)) |
| **3** | FebTable[Sales]              | 2,584,131       | =SUM(INDIRECT(B29)) |
| **4** | MarTable[Sales]              | 2,829,198       | =SUM(INDIRECT(B30)) |
