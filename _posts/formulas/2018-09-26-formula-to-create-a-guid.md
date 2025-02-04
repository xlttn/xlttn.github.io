---
Title: Formula to Create a GUID (Globally Unique Identifier)
categories: [Excel, Formulas]
tags: [unique]
date: 2018-09-26

---

A GUID (or UUID) is an acronym for 'Globally Unique Identifier' (or 'Universally Unique Identifier'). It is a 128-bit integer number used to identify resources. The term GUID is generally used by developers working with Microsoft technologies,										
while UUID is used everywhere else.

**How unique is a GUID?**  
128-bits is big enough and the generation algorithm is unique enough that if 1,000,000,000 GUIDs per second were generated for										
1 year the probability of a duplicate would be only 50%. Or if every human on Earth generated 600,000,000 GUIDs there would										
only be a 50% probability of a duplicate.										

***Formula:***  
```vb
=CONCATENATE(DEC2HEX(RANDBETWEEN(0,4294967295),8),"-",DEC2HEX(RANDBETWEEN(0,42949),4),"-",DEC2HEX(RANDBETWEEN(0,42949),4),"-",DEC2HEX(RANDBETWEEN(0,42949),4),"-",DEC2HEX(RANDBETWEEN(0,4294967295),8),DEC2HEX(RANDBETWEEN(0,42949),4))
```
