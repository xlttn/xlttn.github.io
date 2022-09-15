---
Title: Multiple Criteria Lookup with Index and Match
categories: [Excel, Formulas]
tags: [lookup-reference, array-formulas]
date: 2018-09-27

---

We can expand on the Index and Match formula and use an array  formula to look up a value based on multiple criteria.
This is an array formula so need to press 'Ctrl, Shift, Enter' to put curly brackets around the formula.

### Example 1:

In the table below, we can use a first and last name as criteria to find the Country.  
First Name value to lookup: **Max**  
Last Name value to lookup: **Bradley**    
Result: **Ireland**   

***Formula:***
```vb
{=INDEX(C2:C9,MATCH("Max"&"Bradley",(A2:A9)&(B2:B9),0),1)}
```

| ~     | A              | B             | C           |
|-------|----------------|---------------|-------------|
| **1** | **First Name** | **Last Name** | **Country** |
| **2** | Joe            | Bloggs        | Australia   |
| **3** | Jane           | Smith         | Venezuela   |
| **4** | Ken            | Jones         | Spain       |
| **5** | Kylie          | Moore         | Australia   |
| **6** | Nicole         | Jennings      | England     |
| **7** | Adam           | Taylor        | England     |
| **8** | Anthony        | McDonald      | Australia   |
| **9** | Max            | Bradley       | Ireland     |  

### Example 2:

Let's say we now have 2 Max Bradleys, one from Australia and one from Ireland.
A 4th column "Rank" is introduced, to get the rank of max Bradley from Australia we can follow the above syntax and just add more paramaters in the criteria and range sections within the MATCH function

First Name value to lookup: **Max**   
Last Name value to lookup: **Bradley**    
Country to lookup: **Australia**    
Result: **7**

***Formula:***
```vb
{=INDEX(D2:D9,MATCH("Max"&"Bradley"&"Australia",(A2:A9)&(B2:B9)&(C2:C9),0),1)}
```

|       | A              | B             | C           | D        |
|-------|----------------|---------------|-------------|----------|
| **1** | **First Name** | **Last Name** | **Country** | **Rank** |
| **2** | Joe            | Bloggs        | Australia   | 1        |
| **3** | Jane           | Smith         | Venezuela   | 2        |
| **4** | Ken            | Jones         | Spain       | 3        |
| **5** | Kylie          | Moore         | Australia   | 4        |
| **6** | Nicole         | Jennings      | England     | 5        |
| **7** | Adam           | Taylor        | England     | 6        |
| **8** | Max            | Bradley       | Australia   | 7        |
| **9** | Max            | Bradley       | Ireland     | 8        |
