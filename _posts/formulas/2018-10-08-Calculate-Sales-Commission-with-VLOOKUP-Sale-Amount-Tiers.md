---
Title: Calculate Sales Commission with Vlookup Sale Amount Tiers
categories: [Excel, Formulas]
tags: [lookup-reference, accounting]  
date: 2018-10-08

---

Learn how to calculate commissions for a basic tiered plan and rate table using the VLOOKUP function.

This tables shows a simple commission plan with a rate table that lists the payout rate at each level of sales.

Using the Vlooup function, set the match type to TRUE (you can even leave this out if you want by closing the bracket declaring the column index number).  

Setting the match type to TRUE, vlookup will find the closest match to the lookup value that is less than or equal to the lookup amount.  This basically allows us to find a value between ranges of two numbers (tiers).

**Formula:***
```vb
=VLOOKUP(80000, A2:B13, 2, TRUE)
' or
=VLOOKUP(80000, A2:B13, 2)
```

|        |           A |         B        |                   C |
|--------|------------:|:----------------:|--------------------:|
| **1**  |   **Sales** | **% Commission** |     **Sales Tiers** |
| **2**  |          $- |        0%        |        $0 - $49,999 |
| **3**  |  $50,000.00 |        10%       |   $50,000 - $99,999 |
| **4**  | $100,000.00 |        15%       | $100,000 - $149,999 |
| **5**  | $150,000.00 |        20%       | $150,000 - $199,999 |
| **6**  | $200,000.00 |        25%       | $200,000 - $249,999 |
| **7**  | $250,000.00 |        30%       | $250,000 - $299,999 |
| **8**  | $300,000.00 |        35%       | $300,000 - $349,999 |
| **9**  | $350,000.00 |        40%       | $350,000 - $399,999 |
| **10** | $400,000.00 |        45%       | $400,000 - $449,999 |
| **11** | $450,000.00 |        50%       | $450,000 - $499,999 |
| **12** | $500,000.00 |        55%       | $500,000 - $549,999 |
| **13** | $550,000.00 |        60%       |   $550,000 and over |


### Results:
---

**Corey Ander**  
*Sales:* $100,000  
*Bonus:* 15%"  

**Sue Flay**  
*Sales:* $106,000  
*Bonus:* 15%"  

**Patty O'Ferncher**  
*Sales:* $200,500  
*Bonus:* 25%"  

**Robin Banks**  
*Sales:* $156,000  
*Bonus:* 20%"  

**Lon Moore**  
*Sales:* $45,000  
*Bonus:* 0%"  

**Rose Bush**  
*Sales:* $399,999  
*Bonus:* 40%
