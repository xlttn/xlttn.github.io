---
Title: Extract Nth Word from a String
categories: [Excel, Formulas]
tags: [text-strings]
date: 2018-10-29

---

If you need to get the nth word in a text string (i.e. a sentence, phrase, or paragraph) you can so with a clever (and intimidating) formula that combines 5 Excel functions: TRIM, MID, SUBSTITUTE, REPT, and LEN.

Download the example workbook here: [Extract Nth Word from a String.xlsx](https://github.com/ExcelTitan/Excel_Formulas/raw/master/extract-nth-word-from-string.xlsx)  

```vb
' syntax
=TRIM(MID(SUBSTITUTE(A1," ",REPT(" ",LEN(A1))), (Nth_Word_Number-1)*LEN(A1)+1, LEN(A1)))

' sentence in A2, Nth number is in B2
=TRIM(MID(SUBSTITUTE(A2," ",REPT(" ",LEN(A2))), (B2-1)*LEN(A2)+1, LEN(A2)))
```

With the table set up below, copy the above formula into C2 and drag down to see the results.  

| ~     | A                                                                   | B            | C               |
|-------|---------------------------------------------------------------------|:------------:|-----------------|
| **1** | **Short Quote**                                                     | **Nth Word** | **Result**      |
| **2** | Procrastination is the greatest labor saving invention of all time. | 1            | Procrastination |
| **3** | Black Holes are where God divided by zero.                          | 6            | divided         |
| **4** | A day without sunshine is like, night.                              | 4            | sunshine        |
| **5** | Fish and visitors stink after 3 days.                               | 3            | visitors        |
| **6** | I'm in shape ... round's a shape, isn't it?                         | 5            | round's         |
| **7** | He who laughs last didn't get it.                                   | 3            | laughs          |
