---
Title: Check a Value Exists in a List or Range With COUNTIF
categories: [Excel, Formulas]
tags: [lookup-reference]
date: 2018-09-26

---

When checking if a value exists in a range, using COUNTIF is quicker, easier and less prone to errors.  
This table shows the returning values for each formula which you can paste to C2 and D2 and drag down:  

-  C2: Countif: **`=COUNTIF($A$2:$A$10, B2)`**  
-  D2: Vlookup: **`=VLOOKUP(B2, $A$2:$A$10, 1, 0)`**

|        | A        | B                  | C           | D           |
|--------|----------|--------------------|-------------|-------------|
| **1**  | **Name** | **Names to Check** | **Countif** | **Vlookup** |
| **2**  | Beth     | Terrence           | 1           | Terrence    |
| **3**  | Mitch    | Hayley             | 1           | Hayley      |
| **4**  | Trisha   | Emma               | 0           | #N/A        |
| **5**  | Ken      |                    |             |             |
| **6**  | Nicole   |                    |             |             |
| **7**  | Hayley   |                    |             |             |
| **8**  | Terrence |                    |             |             |
| **9**  | William  |                    |             |             |
| **10** | Matthew  |                    |             |             |

### Why Use COUNTIF Instead of VLOOKUP?
When you just want to determine if a value exists in a list then I recommend using COUNTIF over VLOOKUP. It has a few advantages that make it more efficient, and also give you more insight to your data.

Here are 3 reasons to use COUNTIF instead of VLOOKUP (when you just want to see if a value exists in a range of cells):
1. The COUNTIF function only has two arguments making it really fast and easy to write the formula.  VLOOKUP has four arguments.
2. COUNTIF returns the total number of matching values in the range, so you can see if there is more than one matching value. VLOOKUP cannot do this, it only returns the first match.
3. If the value does not exist, COUNTIF will return a zero (0). You do not need to worry about a formula error. With VLOOKUP, the formula would return an error and you would use and error handling function like IFERROR to handle the error.
Here is a comparison table of the list above, just in case your boss asks why.


The COUNTIF returns a number greater than or equal to 1 if the value exists in the list. It returns a zero if the value does not exist.
The VLOOKUP formula is going to return the matching value from the list. VLOOKUP returns a #N/A error if it can’t find the value in the list.
