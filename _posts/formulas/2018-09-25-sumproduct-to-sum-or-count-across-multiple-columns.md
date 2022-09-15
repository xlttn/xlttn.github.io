---
Title: SUMPRODUCT to Sum or Count Across Multiple Columns
categories: [Excel, Formulas]
tags: [summing-counting]
date: 2018-09-25

---

### Summary
SUMPRODUCT is an incredibly versatile function that can be used to sum and count like SUMIFS or COUNTIFS, but with more flexibility.

### Drill down
Consider the below table, I have the sale amounts for cars, buses and trains from January to April.
If I want to sum the sales of 2 or more months for a particular vehicle then we can use the SUMPRODUCT formula.

We can't use SUMIFS only works when because the sum range exists in 1 column

|   | A     | B         | C         | D         | E         |
|---|-------|-----------|-----------|-----------|-----------|
| 1 | Type  | Jan Sales | Feb Sales | Mar Sales | Apr Sales |
| 2 | Car   | 48        | 39        | 96        | 82        |
| 3 | Bus   | 70        | 52        | 37        | 72        |
| 4 | Train | 0         | 88        | 35        | 47        |

### Sum using SUMPRODUCT
Formula to get the Jan - Feb sales for a car.  
 ```
 =SUMPRODUCT((A2:A4="Car")*(B2:C4))
 ```

Formula for the Jan and March to April sales for a car, you just add the 2 SUMPRODUCT formulas together as you can't use non-contiguous ranges in the formula.
```
=SUMPRODUCT((A2:A4="Car")*(B2:B4))+SUMPRODUCT((A2:A4="Car")*(D2:E4))
```

### Counting using SUMPRODUCT
If you want to check how many months there were sales (i.e sale amount more than 0) then use this...  
```
=SUMPRODUCT((A2:A4="Car")*(B2:E4<=0))
```
