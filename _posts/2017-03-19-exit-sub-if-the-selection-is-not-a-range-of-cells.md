---
Title: Exit Sub if the Selection is Not a Range of Cells
categories: [vba]
tags: [developer]
date: 2017-03-19 18:43:00

---


This code snippet is useful to check if the selection is a range of cells.

```vb
If TypeName(Selection) <> "Range" Then Exit Sub
```
