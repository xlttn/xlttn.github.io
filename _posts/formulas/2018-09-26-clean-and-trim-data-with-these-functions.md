---
Title: Clean and Trim Your Data With These Functions
categories: [Excel, Formulas]
tags: [text-strings]
date: 2018-09-26

---

Combine these 3 functions to completely clean your data:
- =TRIM(***text***)  
- =CLEAN(***text***)  
- =SUBSTITUTE(***text***," ","")  

The CLEAN function removes a range of non-printing characters, including line breaks, and returns "cleaned" text.  
The TRIM function then takes over to remove extra spaces and returns the final text.  
Note that CLEAN is not able to remove all non-printing characters, notably a non-breaking space, which can be appear in Excel as CHAR(160). By adding the SUBSTITUTE function to the formula, you can remove specific characters. For example, to remove a non-breaking space, you can use the following formula										

### Formula										
```
=TRIM(CLEAN(SUBSTITUTE(text, CHAR(160), " ")))
```

___

## What each formula does:

#### TRIM:										
If you need to strip leading and trailing spaces from text in one or more cells, you can use the TRIM function.  
This is very common when you copy/paste data from the Web.  
The TRIM function is fully automatic. It removes removes both leading and trailing spaces from text, and also "normalizes" multiple spaces between words to one space character only. All you need to do is supply a reference to a cell.  
   =TRIM(***text***)  


#### CLEAN:										
The CLEAN function removes a range of non-printing characters, including line breaks, and returns "cleaned" text.										
All you need to do is supply a reference to a cell.  
   =CLEAN(***text***)  


#### SUBSTITUTE:										
Substitutes a character with another one in a cell. Just supply a reference to a cell, text to replace and text to replace it with.   
Example to remove ALL spaces from a string, you write the following formula:  
   =SUBSTITUTE(***text***, " ", "")
