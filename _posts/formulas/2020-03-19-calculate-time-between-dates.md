---
Title: Calculate Time Between Dates
categories: [Excel, Formulas]
tags: [date-time]
date: 2020-03-19 21:45:00

---

Download the example workbook here: [calculate-time-between-dates.xlsx](/example-files/calculate-time-between-dates.xlsx)  

# Calculating hours, minutes and seconds between two times

```vb
=IF(B2< A2, 1 + B2 - A2, B2- A2)
```

**Number Format:** h:mm:ss

![calc-time-img1](/imgs/calculate-time-between-dates/calc-time-img1.png)

# Calculate days

To calculate elapsed days is so easy, you just need to apply this formula

```vb
= B2-A2
```

**Number Format:**

```
General
```

![calc-time-img2](/imgs/calculate-time-between-dates/calc-time-img2.png)

# Calculate Days, Hours and Minutes

To calculate and display the days, hours, and minutes between two dates, you can use the TEXT function with a little help from the INT function.

```vb
=INT(B2-A2)&" days "&TEXT(B2-A2,"h"" hrs ""m"" mins """)
```

**Number Format:**

```
d/m/yyyy h:mm AM/PM
```

## How this formula works
Most of the work in this formula is done by the TEXT function, which applies a custom number format for hours and minutes to a value created by subtracting the start date from the end date.

The value for days is calculated with the INT function, which simply returns the integer portion of the end date minus the start date:

Note: Although you can use "d" in a custom number format for days, the value will reset to zero when days is greater than 31.

## Total days, hours, and minutes between dates
To get the total days, hours, and minutes between a set of start and end dates, you can adapt the formula using SUMPRODUCT like this:

```vb
=INT(SUMPRODUCT(B2:B13-A2:A13))&" days "&TEXT(SUMPRODUCT(B2:B13-A2:A13),"h"" hrs ""m"" mins """)
```

![calc-time-img3](/imgs/calculate-time-between-dates/calc-time-img3.png)

# Calculate Years, Months and Days

## Calculate Days:

```vb
=B2-A2
```

**Number Format:**

```
General
```

## Calculate Months

```vb
=DATEDIF(A2,B2,"m")
```

then format the cells as number.

## Calculate Years

```vb
=DATEDIF(A2,B2,"m")/12
```

then format the cells as number.

## Calculate Years, Months and Days

```vb
=DATEDIF(A2,B2,"Y") & " Years, " & DATEDIF(A2,B2,"YM") & " Months, " & DATEDIF(A2,B2,"MD") & " Days"
```

## Without 0 Values

```vb
=DATEDIF(A2,B2,"Y") & " Years, " & DATEDIF(A2,B2,"YM") & " Months, " & DATEDIF(A2,B2,"MD") & " Days"
```

![calc-time-img4](/imgs/calculate-time-between-dates/calc-time-img4.png)
