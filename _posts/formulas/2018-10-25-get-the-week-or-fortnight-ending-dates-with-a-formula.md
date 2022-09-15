---
Title: Get the Week or Fortnight Ending Dates with a Formula
categories: [Excel, Formulas]
tags: [date-time, lookup-reference]
date: 2018-10-25

---

These formulas are great for when you have to find the week ending date or even the fortnight ending date for a specific item. Week ending dates don't have any dependencies, however can use a cell value as a date reference instead of hard coding the date.
<br>
Calculating which fortnight a date falls in relies on you having already set up a range of dates that the fortnight falls in.
<br>
Download the example workbook here: [Week Ending and Fortnight Ending Dates.xlsx](/example-files/week-ending-and-fortnight-ending-dates.xlsx)  
<br>

## Week Ending Date Functions
The example workbook provides 2 different formulas to get the week ending date:
- Weekday - [support.office.com: weekday function](https://support.office.com/en-us/article/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a?NS=EXCEL&Version=16&SysLcid=1033&UiLcid=1033&AppVer=ZXL160&HelpId=xlmain11.chm60118&ui=en-US&rs=en-US&ad=US)
- Workday.Intl - [support.office.com: workday.int functionl](https://support.office.com/en-us/article/workday-intl-function-a378391c-9ba7-4678-8a39-39611a9bf81d?NS=EXCEL&Version=16&SysLcid=1033&UiLcid=1033&AppVer=ZXL160&HelpId=xlmain11.chm60569&ui=en-US&rs=en-US&ad=US)
<br>

### Weekday Function
My opinion is that the Weekday function is easier to use.
<br>
Looking at the below table, the formulas are referencing D2 which has the date 25/10/2018 (Thursday).  

```vb
' syntax
WEEKDAY(serial_number, [return_type])

' Formula to get the week ending date, D2 has the date as a reference.
' Sunday is last day of the week in this example
=D2+(7-WEEKDAY(D2, 2))
```

| ~     | **A**                    | **B**           | **C**                   |
|-------|--------------------------|-----------------|-------------------------|
| **1** | **Last Day of the Week** | **Week Ending** | **Formula**             |
| **2** | Monday                   | 29/10/2018      | =D2+(7-WEEKDAY(D2, 12)) |
| **3** | Tuesday                  | 30/10/2018      | =D2+(7-WEEKDAY(D2, 13)) |
| **4** | Wednesday                | 31/10/2018      | =D2+(7-WEEKDAY(D2 ,14)) |
| **5** | Thursday                 | 25/10/2018      | =D2+(7-WEEKDAY(D2, 15)) |
| **6** | Friday                   | 26/10/2018      | =D2+(7-WEEKDAY(D2 ,16)) |
| **7** | Saturday                 | 27/10/2018      | =D2+(7-WEEKDAY(D2, 17)) |
| **8** | Sunday                   | 28/10/2018      | =D2+(7-WEEKDAY(D2, 2))  |

#### Return types
The hardest part of this formula is the return type, once you get it it's easy.
<br>
This is just asking, what is the last day of the week, using the following numbers you can change the week ending day of the week:
- 12: Monday
- 13: Tuesday
- 14: Wednesday
- 15: Thursday
- 16: Friday
- 17: Saturday
- 2: Sunday

Usually I use Sunday as the last day of the week so my formula is almost always **`A1+(7-WEEKDAY(A1, 2))`**

### Workday.Intl Function
An interesting function using a binary method to indicate the last day of the week, a 1 is a working day and 0 is a non-working day.

```vb
' syntax
WORKDAY.INTL(start_date, days, [weekend], [holidays])

' formula to get the week ending date, D2 has the date as a reference. Sunday is last day of the week.
=WORKDAY.INTL(D2 - 1, 1, "1111110")
```
start_date: the cell reference D2 which is 25/10/2018
days: always 1 - don't change this
weekend: there has to be 7 digits, six 1s and one 0. The position of the 0 indicates which day is the non-working day. The first digit represents Monday and the seventh represents sunday. <br>
Therefore **`"1111110"`** means that Sunday is the last day of the week

| ~     | **A**                    | **B**           | **C**                               |
|-------|--------------------------|-----------------|-------------------------------------|
| **1** | **Last Day of the Week** | **Week Ending** | **Formula**                         |
| **2** | Monday                   | 29/10/2018      | =WORKDAY.INTL(D2 - 1, 1, "0111111") |
| **3** | Tuesday                  | 30/10/2018      | =WORKDAY.INTL(D2 - 1, 1, "1011111") |
| **4** | Wednesday                | 31/10/2018      | =WORKDAY.INTL(D2 - 1, 1, "1101111") |
| **5** | Thursday                 | 25/10/2018      | =WORKDAY.INTL(D2 - 1, 1, "1110111") |
| **6** | Friday                   | 26/10/2018      | =WORKDAY.INTL(D2 - 1, 1, "1111011") |
| **7** | Saturday                 | 27/10/2018      | =WORKDAY.INTL(D2 - 1, 1, "1111101") |
| **8** | Sunday                   | 28/10/2018      | =WORKDAY.INTL(D2 - 1, 1, "1111110") |

## Fortnight Ending Dates
So, unlike week ending dates, here we have to do a bit of a setup first.  
Just make 2 columns for the start of the fortnight and the end of the fortnight. The example here is a Monday to Sunday fortnight.

| Fortnight Start | Fortnight End |
|----------------:|--------------:|
|      15/10/2018 |    28/10/2018 |
|      29/10/2018 |    11/11/2018 |
|      12/11/2018 |    25/11/2018 |
|      26/11/2018 |     9/12/2018 |
|      10/12/2018 |    23/12/2018 |
|      24/12/2018 |      6/1/2019 |

I have made a named range of these dates as ***FN_Dates*** to simplify the formula a bit.  
We can use either an array of a Vlookup formula, I find the Vlookup much simpler and faster.

***NOTE:*** the Vlookup function won't work if the range of fortnight ending dates are descending. They must start with the smallest date at the top then get bigger i.e. ascending order.

```vb
A1 value is 23/11/2018  

' vlookup
=VLOOKUP(A1, FN_Dates, 2)

' array formula (remember ctrl + shift + enter)
{=MIN(IF(FN_Dates>=A1, FN_Dates))}
```  
### Vlookup formula
As you can see the vlookup is easy, reference the date, then reference the range of fortnights, inidicate column 2 and then omit the 4th argument (range_lookup) as TRUE is default - we won't want to find an exact match here.

### Array formula
The array formula is also easy to understand but takes longer to type, also longer to recalculate as they are pretty much volatile. If any single one of the cells it references has changed or is volatile or has been recalculated then the array formula will evaluate ALL the cells it references.  
The formula is basically going through each of the lines in the FN_Dates range and finding the one that is equal to or larger than the referenced cell. Then out of those ranges, it uses the wrapped MIN function to get the smallest one, thus the closest number to the cell reference.
