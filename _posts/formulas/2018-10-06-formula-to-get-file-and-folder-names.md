---
Title: Formula to get File and Folder Names
categories: [Excel, Formulas]
tags: [text-strings, lookup-reference]
date: 2018-10-06

---

Let's image that the following path is in A1:  
***C:\Users\foo\Documents\parent.folder\foo.bar.txt***  

To get the file, parent folder and path names use the following formulas which all point to A1.  

## File Names   
### File Name with File Type    
**Result:** foo.bar.txt  
```
=MID(A1,FIND("*",SUBSTITUTE(A1,"\","*",LEN(A1)-LEN(SUBSTITUTE(A1,"\",""))))+1,LEN(A1))
```  

### File Name without File Type  
**Result:** foo.bar
```
=LEFT(MID(A1,FIND("*",SUBSTITUTE(A1,"\","*",LEN(A1)-LEN(SUBSTITUTE(A1,"\",""))))+1,LEN(A1)),FIND(CHAR(1),SUBSTITUTE(A1,".",CHAR(1),LEN(A1)-LEN(SUBSTITUTE(A1,".",""))))-1)
```  

### Find the Last Instance of "."  
**Result:** 45  
```
=FIND(CHAR(1), SUBSTITUTE(A1, ".", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, ".", ""))))
```


## Folder Names  
### Folder Path With Backslash  
**Result:** C:\Users\foo\Documents\parent.folder\   
```
=LEFT(A1, FIND(CHAR(1), SUBSTITUTE(A1, "\", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, "\", "")))))
```


### Folder Path Without Backslash  
**Result:** C:\Users\foo\Documents\parent.folder  
```
=LEFT(A1, FIND(CHAR(1), SUBSTITUTE(A1, "\", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, "\", ""))))-1)
```

### Method 1: Parent Folder Name With Backslash  
**Result:** \parent.folder\    
```
="\"&TRIM(MID(SUBSTITUTE(A1,"\",REPT(" ",500)),MAX(1,500*((LEN(A1)-LEN(SUBSTITUTE(A1,"\","")))-1)),500))&"\"
```  

### Method 1: Parent Folder Name Without Backslash  
**Result:** parent.folder  
```
=TRIM(MID(SUBSTITUTE(A1,"\",REPT(" ",500)),MAX(1,500*((LEN(A1)-LEN(SUBSTITUTE(A1,"\","")))-1)),500))
```   


### Method 2: Parent Folder Name With Backslash  
**Result:** parent.folder    
```
="\"&RIGHT(LEFT(A1, FIND(CHAR(1), SUBSTITUTE(A1, "\", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, "\", ""))))-1),LEN(LEFT(A1, FIND(CHAR(1), SUBSTITUTE(A1, "\", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, "\", ""))))-1))-FIND(CHAR(1), SUBSTITUTE(LEFT(A1, FIND(CHAR(1), SUBSTITUTE(A1, "\", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, "\", ""))))-1), "\", CHAR(1), LEN(LEFT(A1, FIND(CHAR(1), SUBSTITUTE(A1, "\", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, "\", ""))))-1))-LEN(SUBSTITUTE(LEFT(A1, FIND(CHAR(1), SUBSTITUTE(A1, "\", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, "\", ""))))-1), "\", "")))))&"\"
```  

### Method 2: Parent Folder Name Without Backslash  
**Result:** parent.folder  
```
=RIGHT(LEFT(A1, FIND(CHAR(1), SUBSTITUTE(A1, "\", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, "\", ""))))-1),LEN(LEFT(A1, FIND(CHAR(1), SUBSTITUTE(A1, "\", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, "\", ""))))-1))-FIND(CHAR(1), SUBSTITUTE(LEFT(A1, FIND(CHAR(1), SUBSTITUTE(A1, "\", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, "\", ""))))-1), "\", CHAR(1), LEN(LEFT(A1, FIND(CHAR(1), SUBSTITUTE(A1, "\", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, "\", ""))))-1))-LEN(SUBSTITUTE(LEFT(A1, FIND(CHAR(1), SUBSTITUTE(A1, "\", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, "\", ""))))-1), "\", "")))))
```   


### Find the Last Instance of Backslash
**Result:** 37  
```
=FIND(CHAR(1), SUBSTITUTE(A1, "\", CHAR(1), LEN(A1)-LEN(SUBSTITUTE(A1, "\", ""))))
```  
