---
Title: Word Count from Cells
categories: [Excel, Formulas]
tags: [text-strings]
date: 2018-10-29

---

There's no built in way to count the number of words in Excel, but using the following formulas will get it done.  

## Count the number of words in a cell
```vb
' A2 as the following sentence (without quotation marks) to get the word count:
' "Procrastination is the greatest labor saving invention of all time."

=IF(LEN(TRIM(A2))0,0,LEN(TRIM(A2))-LEN(SUBSTITUTE(A2," ",""))+1)
```

## Count the number of words in a range

Use an array formula to get the count of words in a range. Remember that array formulas need Ctrl+Shift+Enter to get the curly brackets **`{}`**.
```vb
' where the range is A2:A7

{SUM(IF(LEN(TRIM(A2:A7))0,0,LEN(TRIM(A2:A7))-LEN(SUBSTITUTE(A2:A7," ",""))+1))}
```
