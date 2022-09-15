---
Title: Calculate a Running Total of a Column of Cells, i.e. Cumulative Total
categories: [Excel, Formulas]
tags: [summing-counting, accounting] 
date: 2018-09-26

---

If you have a column of numbers and you want to calculate a running total of the numbers alongside, you can use the SUM() formula combined with a clever use of absolute and relative references.  
The formula below can even be used to calculate the running total across multiple columns.  

#### Example 1: Single Column  
Column C shows the cumulative running total of the sales in column B.  

|       | A          | B        | C                 | D             |
|-------|------------|---------:|------------------:|---------------|
| **1** | **Date**   | **Sale** | **Running Total** | **Formula**   |
| **2** | 26/09/2018 | 58       | 58                | =SUM($B$2:B2) |
| **3** | 27/09/2018 | 131      | 189               | =SUM($B$2:B3) |
| **4** | 28/09/2018 | 72       | 261               | =SUM($B$2:B4) |
| **5** | 29/09/2018 | 117      | 378               | =SUM($B$2:B5) |
| **6** | 30/09/2018 | 129      | 507               | =SUM($B$2:B6) |
| **7** | 1/10/2018  | 68       | 575               | =SUM($B$2:B7) |
| **8** | 2/10/2018  | 129      | 704               | =SUM($B$2:B8) |
| **9** | 3/10/2018  | 131      | 835               | =SUM($B$2:B9) |


#### Example 2: Multiple Columns   
Column F shows the cumulative running total of the sales in from columns B to E.  

|       | A          | B      | C      | D      | E      | F             | G             |
|-------|------------|-------:|-------:|-------:|-------:|--------------:|:--------------|
| **1** | **Date**   | **Sale 1** | **Sale 2** | **Sale 3** | **Sale 4** | **Running Total** | **Formula**  |
| **2** | 26/09/2018 | 58     | 90     | 102    | 109    | 359           | =SUM($B$2:E2) |
| **3** | 27/09/2018 | 131    | 90     | 145    | 67     | 792           | =SUM($B$2:E3) |
| **4** | 28/09/2018 | 72     | 132    | 125    | 142    | 1263          | =SUM($B$2:E4) |
| **5** | 29/09/2018 | 117    | 83     | 141    | 143    | 1747          | =SUM($B$2:E5) |
| **6** | 30/09/2018 | 129    | 109    | 91     | 149    | 2225          | =SUM($B$2:E6) |
| **7** | 1/10/2018  | 68     | 73     | 92     | 126    | 2584          | =SUM($B$2:E7) |
| **8** | 2/10/2018  | 129    | 141    | 98     | 90     | 3042          | =SUM($B$2:E8) |
| **9** | 3/10/2018  | 131    | 71     | 95     | 80     | 3419          | =SUM($B$2:E9) |
