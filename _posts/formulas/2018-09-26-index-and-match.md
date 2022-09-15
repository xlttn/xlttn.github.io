---
Title: INDEX and MATCH in Excel, a better alternative to VLOOKUP
categories: [Excel, Formulas]
tags: [lookup-reference]  
date: 2018-09-26

---

This tutorial demonstrates the key strengths of Excel's INDEX / MATCH function that make it superior to VLOOKUP. You will find a number of formula examples that will help you easily cope with many complex tasks when VLOOKUP fails.

### How to use INDEX MATCH function in Excel

The MATCH function determines the relative position of the lookup value in the specified range of cells. From there, the INDEX function takes that number, or numbers, and returns a value in the corresponding cell.
Still having difficulties to figure it out? Think about Excel INDEX / MATCH in this way:

```vb
=INDEX (column to return a value from, MATCH (lookup value, column to look up against, 0))  
```

I believe it's even easier to understand from an example. Suppose you have a list of national capitals like this:


|    | A    | B           | C           | D          |
|----|------|-------------|-------------|------------|
| 1  | Rank | Country     | Capital     | Population |
| 2  | 1    | China       | Beijing     | 20,693,000 |
| 3  | 2    | India       | New Delhi   | 17,838,842 |
| 4  | 3    | Japan       | Tokyo       | 13,189,000 |
| 5  | 4    | Russia      | Moscow      | 11,541,000 |
| 6  | 5    | South Korea | Seoul       | 10,528,774 |
| 7  | 6    | Indonesia   | Jakarta     | 10,187,595 |
| 8  | 7    | Iran        | Tehran      | 9,110,347  |
| 9  | 8    | Mexico      | Mexico City | 8,851,080  |
| 10 | 9    | Peru        | Lima        | 8,481,415  |

Let's find the population of some capital, say the capital of Japan, using the following Index Match formula:
```vb
=INDEX($D$2:$D$10,MATCH("Japan",$B$2:$B$10,0))
```

Now, let's analyze what each component of this formula actually does:

- The MATCH function searches for the lookup value "Japan" in column B, more precisely cells B2:B10, and returns the number 3, because "Japan" is the third in the list.
- The INDEX functions takes "3" in the second parameter (row_num), which indicates which row you want to return a value from, and turns into a simple **`=INDEX($D$2:$D$10,3)`**. Translated into plain English, the formula reads: search in cells D2 through D10 and return a value of the cell in the 3rd row, i.e. cell D4, because we start counting from the second row.

And here's the result you get in Excel: **13,189,000**

**Important!** The number of rows and columns in the INDEX array should match those in the row_num or/and column_num parameters of the MATCH functions, respectively. Otherwise, the formula will return incorrect result
In this example we could use the VLOOKUP function as the lookup value (Japan) is on the left. This is example is to just get a feel for how it works.

### How to look up from right to left with INDEX & MATCH

As stated in any VLOOKUP tutorial, this Excel function cannot look at its left. So, unless your lookup column is the left-most column in the lookup range, there's no chance that a vlookup formula will return the result you want.

Excel's INDEX MATCH function is more flexible and does not really care where the return column resides. As an example, we will use the table listing national capitals by population again. This time, let's write an INDEX MATCH formula that finds how the Russian capital, Moscow, ranks in terms of population.

As you can see in the table below, the following formula has no problem with performing a left vlookup:

```vb
=INDEX($A$2:$A$10,MATCH("Russia",$B$2:$B$10,0))
```

Naturally, you can replace the "hard-coded" lookup value (Russia) with a cell reference if you want to.

By now, you should not have any difficulties to understand how the formula works:

**First**, you write a simple MATCH formula that finds the position of Russia:  

```vb
=MATCH("Russia",$B$2:$B$10,0))
```  

**Second**, you determine the array parameter for your Index function, which is column A in our case.   **`(A2:A10)`**  

**Finally**, you assemble the two parts together and get this formula:  

```vb
=INDEX($A$2:$A$10,MATCH("Russia",$B$2:$B$10,0))
```

**Tip:** It's a good idea to always use absolute cell references in INDEX and MATCH formulas so that your lookup ranges won't get distorted when you copy the formula to other cells.

### Why Excel's INDEX MATCH is better than VLOOKUP
When deciding which formula to use for vertical lookups, the majority of Excel gurus agree that INDEX / MATCH is far better than VLOOKUP. However, many Excel users still resort to utilizing VLOOKUP because it's a simpler function. This happens because very few people fully understand all the benefits of switching from Vlookup to Index Match, and without such understanding no one is willing to invest their time to learn a more complex formula.

Below, I will try to point out the key advantages of using MATCH / INDEX in Excel, and then you decide whether you'd rather stick with Vlookup or switch to Index Match.

### 4 top benefits of using MATCH INDEX in Excel

1. **_Right to left lookup._** VLOOKUP cannot look to its left, meaning that your lookup value has to be in the left-most column of the lookup range.  
2. **_Insert or delete columns safely._** VLOOKUP formulas get broken or return incorrect results when a new column is deleted from or added to a lookup table. With VLOOKUP, any inserted or deleted column changes the results returned by your formulas because the VLOOKUP function's syntax requires specifying the entire table array and a certain number indicating which column you want to pull the data from.   
With INDEX MATCH, you can delete or insert new columns in your lookup table without distorting the results since you specify directly the column containing the value you want to get.  
3. **_No limit for a lookup value's size._** When using the VLOOKUP function, remember that the total length of your lookup criteria cannot exceed 255 characters, otherwise you will end up having the #VALUE! error. So, if your dataset contains long strings, INDEX MATCH is the only working solution.
4. **_Higher processing speed._** If your tables are relatively small, there will hardly be any significant difference in Excel's performance. But if your worksheets contain hundreds or thousands of rows, and consequently hundreds or thousands of formulas, MATCH INDEX will work much faster than VLOOKUP because Excel will have to process only the lookup and return columns rather than the entire table array.  

### Calculations with INDEX MATCH in Excel (AVERAGE, MAX, MIN)
You can nest other Excel functions within the MATCH INDEX formula, say, to find the minimum or maximum value, or the value closest to the average in the range. Here are a few formula examples for the table used in the previous sample:

| Function | Formula example                                             | Description                                                                                                                      | Returned result |
|----------|-------------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------|-----------------|
| Min      | **`=INDEX($C$2:$C$10, MATCH(MIN($D$2:I$10), $D$2:D$10, 0))`**   | Finds the min value in column D and returns a value from column C in the same row.                                               | Beijing         |
| Max      | **`=INDEX($C$2:$C$10, MATCH(MAX($D$2:I$10), $D$2:D$10, 0))`**   | Finds the max value in column D and returns a value from column C in the same row.                                               | Lima            |
| Average  | **`=INDEX($C$2:$C$10, MATCH(AVERAGE($D$2:D$10), $D$2:D$10, 1))`**| Calculates the average in range D2:D10, finds the value closest to the average, and returns a corresponding value from column C. | Moscow          |
