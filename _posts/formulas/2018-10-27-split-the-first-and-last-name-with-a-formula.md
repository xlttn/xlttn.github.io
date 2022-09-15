---
Title: Split the First and Last Name with a Formula
categories: [Excel, Formulas]
tags: [text-strings]
date: 2018-10-27

---
## How to split first and last name from full name with space

These formulas cover the most typical scenario when you have the first name and last name in one column separated by a single space character.

### Formula to get first name
The first name can be easily extracted with this generic formula:
```vb
' first name
LEFT(cell, SEARCH(" ", cell) - 1)
```
You use the SEARCH or FIND function to get the position of the space character (" ") in a cell, from which you subtract 1 to exclude the space itself. This number is supplied to the LEFT function as the number of characters to be extracted, starting on the left side of the string.
<br>
### Formula to get last name
The generic formula to extract a surname is this:
```vb
' last name
RIGHT(cell, LEN(cell) - SEARCH(" ", cell))
```

In this formula, you also use the SEARCH function to find the position of the space char, subtract that number from the total length of the string (returned by LEN), and get the RIGHT function to extract that many characters from the right side of the string.

----
With the full name in cell A2, the formulas take the following shape:  
```vb
' first name
=RIGHT(A2, LEN(A2) - SEARCH(" ", A2))

' last name
=LEFT(A2, SEARCH(" ", A2) - 2)
```

| ~     | **A**         | **B**          | **C**         |
|-------|---------------|----------------|---------------|
| **1** | **Full Name** | **First Name** | **Last Name** |
| **2** | Joe King      | Joe            | King          |
| **3** | Sue Flay      | Sue            | Flay          |
| **4** | Cory Ander    | Cory           | Ander         |

## Handle potential middle names
If some of the original names contain a middle name or middle initial, you'd need a bit more tricky formula to extract the last name:  

You replace the last space in the name with a hash sign (#) or any other character that do not appear in any name and work out the position of that char. After that, you subtract the above number from the total string length to get the length of the last name, and have the RIGHT function extract that many characters.  

So, here's how you can separate the first name and surname in Excel when some of the original names include a middle name:  
```vb
' formula in C2 and copy down
=RIGHT(A2, LEN(A2) - SEARCH("#", SUBSTITUTE(A2," ", "#", LEN(A2) - LEN(SUBSTITUTE(A2, " ", "")))))
```

| ~     | **A**              | **B**          | **C**         |
|-------|--------------------|----------------|---------------|
| **1** | **Full Name**      | **First Name** | **Last Name** |
| **2** | Barry D. Hatchett  | Barry          | Hatchett      |
| **3** | Sue Flay           | Sue            | Flay          |
| **4** | Milo Fletcher Ball | Milo           | Ball          |

## Last_Name, First_Name
If you have a column of names in the Last name, First name format, with the full name in cell A2, the formulas take the following shape:

```vb
' first name
=RIGHT(A2, LEN(A2) - SEARCH(" ", A2))

' last name
=LEFT(A2, SEARCH(" ", A2) - 2)
```

Like in the above example, you use the SEARCH function to determine the position of a space character, and then subtract it from the total string length to get the length of the first name. This number goes directly to the num_chars argument of the RIGHT function indicating how many characters to extract from the end of the string.

To get a last name, you use the LEFT SEARCH combination discussed in the previous example with the difference that you subtract 2 instead of 1 to account for two extra characters, a comma and a space.

| ~     | **A**         | **B**          | **C**         |
|-------|---------------|----------------|---------------|
| **1** | **Full Name** | **First Name** | **Last Name** |
| **2** | King, Joe     | Joe            | King          |
| **3** | Flay, Sue     | Sue            | Flay          |
| **4** | Ander, Cory   | Cory           | Ander         |

## First_Name Middle_Name Last_Name
Splitting names that include a middle name or middle initial requires slightly different approaches, depending on the name format.  
If your names are in the First name Middle name Last name format, the below formulas will work nicely:

```vb
' first name
=LEFT(A2,SEARCH(" ", A2)-1)

' middle name
=MID(A2, SEARCH(" ", A2) + 1, SEARCH(" ", A2, SEARCH(" ", A2)+1) - SEARCH(" ", A2)-1)

' last name
=RIGHT(A2,LEN(A2) - SEARCH(" ", A2, SEARCH(" ", A2,1)+1))
```

| ~     | **A**              | **B**          | **C**           | **D**         |
|-------|--------------------|----------------|-----------------|---------------|
| **1** | **Full Name**      | **First Name** | **Middle Name** | **Last Name** |
| **2** | Barry D. Hatchett  | Barry          | D.              | Hatchett      |
| **3** | Milo Fletcher Ball | Milo           | Fletcher        | Ball          |
| **4** | Les Maximus Dickus | Les            | Maximus         | Dickus        |

## Last_Name, First_Name Middle_Name
Here are the formulas when the order is Last Name, First Name Middle Name

```vb
' first name
=MID(A2, SEARCH(" ",A2) + 1, SEARCH(" ", A2, SEARCH(" ", A2) + 1) - SEARCH(" ", A2) -1)

' middle name
=RIGHT(A2, LEN(A2) - SEARCH(" ", A2, SEARCH(" ", A2, 1)+1))

' last name
=LEFT(A2, SEARCH(" ",A2,1)-2)
```

| ~     | **A**               | **B**          | **C**           | **D**         |
|-------|---------------------|----------------|-----------------|---------------|
| **1** | **Full Name**       | **First Name** | **Middle Name** | **Last Name** |
| **2** | Hatchett, Barry D.  | Barry          | D.              | Hatchett      |
| **3** | Ball, Milo Fletcher | Milo           | Fletcher        | Ball          |
| **4** | Dickus, Les Maximus | Les            | Maximus         | Dickus        |

## First_Name Last_Name, Suffix
Here's a similar approach to split names with a suffix

```vb
' first name
=LEFT(A2, SEARCH(" ",A2)-1)

' last name
=MID(A2, SEARCH(" ",A2) + 1, SEARCH(",",A2) - SEARCH(" ",A2)-1)

' suffix
=RIGHT(A2, LEN(A2) - SEARCH(" ", A2, SEARCH(" ",A2)+1))
```

| ~     | **A**           | **B**          | **C**         | **D**      |
|-------|-----------------|----------------|---------------|------------|
| **1** | **Full Name**   | **First Name** | **Last Name** | **Suffix** |
| **2** | Joe King, Jr.   | Joe            | King          | Jr.        |
| **3** | Sue Flay, Ph.d  | Sue            | Flay          | Ph.d       |
| **4** | Cory Ander, Sr. | Cory           | Ander         | Sr.        |

## Last Name, First Name Suffix

```vb
' first name
=MID(A2, SEARCH(" ", A2) + 1, SEARCH(" ", A2, SEARCH(" ", A2)+1) - SEARCH(" ", A2)-1)

' last name
=LEFT(A2, SEARCH(", ", A2) - 1)

' suffix
=RIGHT(A2, LEN(A2) - SEARCH("#", SUBSTITUTE(A2," ", "#", LEN(A2) - LEN(SUBSTITUTE(A2, " ", "")))))
```

| ~     | A                         | B              | C             | D          |
|-------|---------------------------|----------------|---------------|------------|
| **1** | **Name from Data Dource** | **First Name** | **Last Name** | **Suffix** |
| **2** | King, Joe Mr              | Joe            | King          | Mr         |
| **3** | Flay, Sue Dr              | Sue            | Flay          | Dr         |
| **4** | Ander, Cory Mr            | Cory           | Ander         | Mr         |
| **5** | Fletcher-Ball, Milo Mr    | Milo           | Fletcher-Ball | Mr         |
| **6** | Sober, Alicia Mrs         | Alicia         | Sober         | Mrs        |
| **7** | Fate, Celia Miss          | Celia          | Fate          | Miss       |
