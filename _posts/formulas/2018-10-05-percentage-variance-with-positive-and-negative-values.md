---
Title: Percentage Variance With Positive and Negative Values
categories: [Excel, Formulas]
tags: [accounting]  
date: 2018-10-05

---

There are 2 formulas to calculate percentage variance (change).

#### Formula 1  
```vb
=(new value – old value) / old value
```


#### Formula 2  
```vb
=(new value / old value) – 1
```

Both of these formulas will produce the same result when the numbers are positive.  The one you use is just a matter of personal preference.

## What if the numbers are Negative?
One common way to calculate percentage change with negative numbers it to make the denominator in the formula positive.  The ABS function is used in Excel to change the sign of the number to positive, or its absolute value.  

Here is the formula that is commonly used:  
```vb
=(new value – old value) / ABS(old value)
```

This produces misleading results, here the old value is negative and the new value is positive. The greater the value change shows smaller percentage changes.  

|       | A           | B       | C       | D          | E              | F                      |
|-------|-------------|---------|---------|------------|----------------|------------------------|
| **1** | **Product** | **Old** | **New** | **Change** | **Pct Change** | **Pct Change Formula** |
| **2** | Coffee      | -10     | 50      | 60         | 600.0%         | =(C2-B2)/ABS(B2)       |
| **3** | Tea         | -20     | 50      | 70         | 350.0%         | =(C3-B3)/ABS(B3)       |
| **4** | Cookies     | -30     | 50      | 80         | 266.7%         | =(C4-B4)/ABS(B4)       |
| **5** | Bagels      | -40     | 50      | 90         | 225.0%         | =(C5-B5)/ABS(B5)       |
| **6** | Apples      | -50     | 50      | 100        | 200.0%         | =(C6-B6)/ABS(B6)       |
| **7** | Cakes       | -60     | 50      | 110        | 183.3%         | =(C7-B7)/ABS(B7)       |


## Alternate Calculations for Percentage Change with Negative Numbers

### Method 1: No Result for Negatives  
The first thing we can do is check if either number is negative, and then display some text to tell the reader a percentage change calculation could not be made.

The following formula does this with an IF function and MIN function.

```vb
=IF(MIN(old value, new value)<=0,"--",(new value/old value)-1)
```

|       | A           | B       | C       | D          | E              | F                                 |
|-------|-------------|--------:|--------:|-----------:|---------------:|-----------------------------------|
| **1** | **Product** | **Old** | **New** | **Change** | **Pct Change** | **Pct Change Formula**            |
| **2** | Coffee      | -10     | 50      | 60         | --             | =IF(MIN(B2,C2)<=0,"--",(C2/B2)-1) |
| **3** | Tea         | 5       | -50     | -55        | --             | =IF(MIN(B3,C3)<=0,"--",(C3/B3)-1) |
| **4** | Cookies     | -40     | -50     | -10        | --             | =IF(MIN(B4,C4)<=0,"--",(C4/B4)-1) |
| **5** | Bagels      | 0       | 10      | 10         | --             | =IF(MIN(B5,C5)<=0,"--",(C5/B5)-1) |
| **6** | Apples      | 25      | 30      | 5          | 20.0%          | =IF(MIN(B6,C6)<=0,"--",(C6/B6)-1) |
| **7** | Cakes       | 50      | 10      | -40        | -80.0%         | =IF(MIN(B7,C7)<=0,"--",(C7/B7)-1) |


### Method #2: Show Positive or Negative Change
The Wall Street Journal guide says that its earning reports display a “P” or “L” if there is a negative number and the company posted a profit or loss.  

We could use this same methodology to tell our readers if the change was positive (P) or negative (N) when either value is negative.  

The following formula tests for this with an additional IF function.  

```vb
=IF(MIN(old value, new value)<=0,IF((new value - old value)>0,"P","N"),(new value/old value)-1)
```

| ~     | A           |       B |       C |          D |              E | F                                                  |
|-------|-------------|--------:|--------:|-----------:|---------------:|----------------------------------------------------|
| **1** | **Product** | **Old** | **New** | **Change** | **Pct Change** | **Pct Change Formula**                             |
| **2** | Coffee      |     -10 |      50 |         60 |              P | `=IF(MIN(B2,C2)<=0,IF((C2-B2)>0,"P","N"),(C2/B2)-1)` |
| **3** | Tea         |       5 |     -50 |        -55 |              N | `=IF(MIN(B3,C3)<=0,IF((C3-B3)>0,"P","N"),(C3/B3)-1)` |
| **4** | Cookies     |     -40 |     -50 |        -10 |              N | `=IF(MIN(B4,C4)<=0,IF((C4-B4)>0,"P","N"),(C4/B4)-1)` |
| **5** | Bagels      |       0 |      10 |         10 |              P | `=IF(MIN(B5,C5)<=0,IF((C5-B5)>0,"P","N"),(C5/B5)-1)` |
| **6** | Apples      |      25 |      30 |          5 |          20.0% | `=IF(MIN(B6,C6)<=0,IF((C6-B6)>0,"P","N"),(C6/B6)-1)` |
| **7** | Cakes       |      50 |      10 |        -40 |         -80.0% | `=IF(MIN(B7,C7)<=0,IF((C7-B7)>0,"P","N"),(C7/B7)-1)` |
