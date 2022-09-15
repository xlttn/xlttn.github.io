---
Title: Convert Decimal to Total Hours and Minutes
categories: [Excel, Formulas]
tags: [date-time]  
date: 2018-09-26

---

When converting a decimal to hours and minutes, ensure to use the correct Number Format.

***[h]:mm:ss***  
Shows the total number of hours, even if the number of hours is 24 or more.

***hh:mm:ss***  
Effectively only shows the excess hours over and above complete multiples of 24.  
For example, 105 hours is 4×24, or 4 whole days, plus 9 more hours, so it just shows the 9.  
Remember, this is only a display format – the underlying values are the same in both cases.

See the below tabled example.										

|       | A       | B     | C      | D       |
|-------|--------:|------:|-------:|---------|
| **1** | Decimal | hh:mm | [h]:mm | Formula |
| **2** | 4.3     | 04:18 | 4:18   | =A2/24  |
| **3** | 25      | 01:00 | 25:00  | =A3/24  |
| **4** | 105     | 09:00 | 105:00 | =A4/24  |
