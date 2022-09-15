---
Title: Repeat Values n Times
categories: [Excel, Formulas]
tags: [tables, practical]
date: 2020-03-19 18:45:00

---

Here is a method to repeat the values in Column 'A', (n) number of times required in Column 'B' outputting the results in Column 'C'. The Formula is an Array Formula requiring you to press CTRL+SHIFT+ENTER.

So with a list in Column A and the respective 'n' times to repeat the elements of the list in Column 'B', enter the following Array Formula into Column 'C'.

You will need to extend the Ranges if you want to add more Products to the list.

Download the example workbook here: [repeat-values-n-times.xlsx](/example-files/repeat-values-n-times.xlsx)  

```vb
' enter the following array formula in Cell C2
=IFERROR(INDEX($A$2:$A$5,MATCH(TRUE,MMULT(--(ROW($A$2:$A$5)>=TRANSPOSE(ROW($A$2:$A$5))),$B$2:$B$5)>=ROWS($1:1),0)),"")

' this will look like the following in the Cell when you press CTRL+SHIFT+ENTER
{=IFERROR(INDEX($A$2:$A$5,MATCH(TRUE,MMULT(--(ROW($A$2:$A$5)>=TRANSPOSE(ROW($A$2:$A$5))),$B$2:$B$5)>=ROWS($1:1),0)),"")}
```


![repeat-values-n-times-img](/imgs/repeat-values-n-times/repeat-values-n-times.png)
