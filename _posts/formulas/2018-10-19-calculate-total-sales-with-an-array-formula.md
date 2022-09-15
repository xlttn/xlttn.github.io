---
Title: Calculate Total Sales With An Array Formula
categories: [Excel, Formulas]
tags: [array-formulas, summing-counting]
date: 2018-10-19

---

If we had to calculate Total Sales the normal way, we would have to create a 'helper column' for the Totals column and then enter a formula to Sum all the Totals.
<br>
<br>
Consider this classic example tabulated below, usually we would need a 4th column to total each row, then sum the totals. But, with the use of an array formula we can do this without those helper columns.
<br>
### Entering the Array Formula
An array formula allow us to compute multiple ranges without the need for helper columns.  
***Remember:*** instead of pressing enter after typing the formula, you need to hold down **`Ctrl + Shift`** then press **`Enter`**

```vb
' formula syntax
{=SUM(Units_Sold * Unit_Price)}

' formula
{=SUM(B2:B5*C2:C5)}

' result: total sales
$443,599
```

### Sample Data

|   ~   | **A**       | **B**          | **C**          |
|-------|-------------|----------------|----------------|
| **1** | **Product** | **Units Sold** | **Unit Price** |
| **2** | Laptops     | 87             | 788            |
| **3** | Tablets     | 268            | 651            |
| **4** | Mobiles     | 541            | 325            |
| **5** | PC          | 55             | 450            |

----
### Normal Formula:
This is what you would normally have to do:
- Extra totals column: Column D
- Totals column to multiply Units Sold by Unit Price
- Sum the Total Sales

|  ~    | **A**       | **B**          | **C**           | **D**       |
|-------|-------------|----------------|-----------------|-------------|
| **1** | **Product** | **Units Sold** | **Unit Price**  | **Total**   |
| **2** | Laptops     | 87             | 788             | 68,556      |
| **3** | Tablets     | 268            | 651             | 174,468     |
| **4** | Mobiles     | 541            | 325             | 175,825     |
| **5** | PC          | 55             | 450             | 24,750      |
| **6** |             |                | **Total Sales** | **443,599** |
