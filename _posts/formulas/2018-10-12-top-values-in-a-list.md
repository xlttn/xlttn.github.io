---
Title: Top Values in a List
categories: [Excel, Formulas]
tags: [lookup-reference]  
date: 2018-10-12

---

From a list of 30 people with an assigned value, generate a smaller list that are the top n values in that list.												

Download the example workbook here: [Top n Values in a List.xlsx](/example-files/Top-n-Values-in-a-List.xlsx)  


## How it all works
Our source data is in A1:D31 and Output begins in H1.  
All formulas have error checking to display blanks, also to check that we don't generate a list longer than the amount we want to show.  
For the purpose of this, I won't include the error handling components to the formulas below however they are in the example sheet.

#### Top n Values
Cell F2 has a drop down validation to choose from 5, 10, 15, 20, 25, 30. This indicates how many values to show.

#### Ranking formula.
We have a Ranking formula included in the source data to make way for duplicate values in Column C. This ensures a unique ranking list, meaning a simple Index + Match formula can be used to look up the Names.
Ranking formula starting in C1:
```vb
=C2+10^-6*ROWS($A2:A$2)
```

#### Getting the Top Value and Top Rank Value
To get the largest to smallest values in the list, use the Large function which only has 2 arguments: array, k (i.e. 1st, 2nd, 3rd largest).  
***Top Value:***
```vb
=LARGE($C$2:$C$31,H2)
```  

***Top Rank:***
=LARGE($D$2:$D$31,H2)

#### Getting the First and Last Names based on Value:
Both of these formulas index the Rank column and match the Rank value that has already been calcualted on the same row in column L.

***First Name:***
```vb
=INDEX($A$2:$A$31,MATCH(L2,$D$2:$D$31,0))
```  

***Last Name:***
```vb
=INDEX($B$2:$B$31,MATCH(L2,$D$2:$D$31,0))
```
