---
Title: Convert Numbers Stored as Text when Using a Formula
categories: [Excel, Formulas]
tags: [text-strings]
date: 2018-10-11

---

There are a few ways to convert Numbers stored as Text using Formula. You can use:
- =VALUE() or =NUMBERVALUE() for floating point Numbers.
- =INT() if you just have whole Numbers
- =T() to retain formatting.

If you have a Number stored as Text in Cell A1 use any of the following:

```vb
'// converts a Number stored as Text into a Number
=VALUE(A1)

'// whole Number
=INT(A1)

' // Converts text to a number, in a locale-independent way (new for Excel 2013)
=NUMBERVALUE(A1)

'// if you have a Number stored as Text like 1,200 this will retain its comma formatting
=T(A1)
```
