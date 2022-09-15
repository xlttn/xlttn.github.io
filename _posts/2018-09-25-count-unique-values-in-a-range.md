---
Title: Count the Unique Values in a Range
categories: [Excel, Formulas]
tags: [unique, array-formulas]
date: 2018-09-25

---

**Link to all xlsx Example Files:** [https://github.com/ExcelTitan/Excel_Formulas](https://github.com/ExcelTitan/Excel_Formulas)

To count the unique values in a range place the following formula in B2, this is an array formula so remember to hit Ctrl+Shift+Enter											
This formula handles and ignores empty cells in the range as well											

```vb
{=SUM(IF(FREQUENCY(IF(A2:A9<>"",MATCH(A2:A9,A2:A9,0)),ROW(A2:A9)-ROW(A2)+1),1))}
```

| Name |
|:---|
| Beth |
| Nicole |
| Adam |
| Beth |
| Michele |
| Anthony |
| Beth |
| Hayley |
