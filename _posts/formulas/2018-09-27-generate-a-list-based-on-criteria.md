---
Title: Generate a List Based on Criteria
categories: [Excel, Formulas]
tags: [lookup-reference]  
date: 2018-09-27

---

Generate a list of items based on a criteria that doesn't show any blanks or spaces.  
20 Fruits and Vegetables listed below, choose either a fruit or vegetable and a list will populate without blanks or spaces.  

***Formula for E2:*** _remember to use Ctrl + Shift + Enter as this is an array formula_  
```vb
{=IF(COUNTIF(A:A,$D$2)<ROWS(E2:$E$2),"",INDEX(B:B,SMALL(IF($A$2:$A$7=$D$2,ROW($A$2:$A$7)),ROW(A1))))}
```


| ~     | A         | B         | C | D                           | E                 |
|-------|-----------|-----------|---|-----------------------------|-------------------|
| **1** | **Type**  | **Item**  |   | **Choose Type to look up**' | **List of items** |
| **2** | Fruit     | Apple     |   | Vegetable                   | Broccoli          |
| **3** | Fruit     | Orange    |   |                             | Spinach           |
| **4** | Vegetable | Broccoli  |   |                             | Peas              |
| **5** | Vegetable | Spinach   |   |                             |                   |
| **6** | Fruit     | Pear      |   |                             |                   |
| **7** | Vegetable | Peas      |   | ↑ _(Vegetable or Fruit)_    | ↑ _Output List_   |

### Front Half of the Formula  
**`IF(COUNTIF(A:A,$D$2)<ROWS($E$2:E2),""`**  
COUNTIF function determines the total number of records that meet our criteria. We’re then comparing this to a ROWS function which simply returns the number of rows given in the argument.  

Note the first part of the range uses an absolute reference and will not change, while the latter part is relative and will change as the formula is copied down. Thus, in the first cell, the ROWS function evaluates to 1. The next cell, it will evaluate to 2, then 3, and so on. So, the IF statement is checking to see if the number of records returned so far (i.e., formula used) is greater than the total number of possible records. If this is true, return a blank (i.e., "").

### Back Half of the Formula - Part 1
```vb
IF($A$2:$A$10=$D$2
```
This section compares A2:A10 with our criteria given in cell D2. So, the array if A2:A10 starts off looking like this:  
```vb
{Fruit, Fruit, Vegetable, Vegetable, Fruit, Vegetable, "", "", ""}
```

When we compare it with the criteria, it becomes this:
```vb
{False, False, True, True, False, True, False, False, False}
```  

Looking at the return values in our IF function, we see that only a True result is stated, the ROW.  

### Back Half of the Formula - Part 2
```vb
ROW($A$2:$A$10)
```
So, each True value from the array above will be replaced with the corresponding Row value.

This causes the array to become this:  
**`{False, False, 4, 5, False, 7, False, False, False}`**  
Now that we have a nice array with some numbers in it, this gets fed into the SMALL function.

### Back Half of the Formula - Part 3
```vb
INDEX(B:B, SMALL(IF($A$2:$A$10=$D$2, ROW($A$2:$A$10)), ROW(A1))))
```

The ROW function at the end will serve as a type of counter.
In E2, where we initially place the formula, this will evaluate to 1, thus telling the SMALL function to return the 1st smallest number.
In E3, it will evaluate to 2, and the SMALL function will return the 2nd smallest number, and so.
So, taking the 1st smallest number from our array, we get the number 4.

### Back Half of the Formula - Part 4
```vb
INDEX(B:B, SMALL(IF($A$2:$A$10=$D$2, ROW($A$2:$A$10)), ROW(A1))))
```

Note that we need to callout the entire column, since we are plugging in row numbers.  
The 4th row in column B leads us to the value "Broccoli".  
The next formula will return the 5th row, "Spinach".  
The 3rd formula will return the 7th row, "Peas".  

This method can be adapted for use with multiple criteria. We would just need to expand the IF function logic checks so that only the correct rows are returned.  

### Back Half of the Formula - Final Part
```vb
=IF(COUNTIF(A:A,$D$2)<ROWS($E$2:E2),"", INDEX(B:B, SMALL( IF($A$2:$A$10 = $D$2, ROW($A$2:$A$10)), ROW(A1))))
```

It is just there to hide any unwanted #NUM errors after all the pertinent records have been displayed.
