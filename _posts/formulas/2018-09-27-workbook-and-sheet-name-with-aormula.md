---
Title: Get the Workbook and Sheet Name With a Formula
categories: [Excel, Formulas]
tags: [file-info]
date: 2018-09-27

---

These formulas require that the workbook is saved

### Get the file name of this workbook:
Formula: **`=CELL("filename",A1)`**  
Result: C:\[MyWorkbook.xlsx]Sheet1


### Get the file name of this workbook without the file extension:
Formula: **`=MID(CELL("filename",A1),SEARCH("[",CELL("filename",A1))+1,SEARCH(".",CELL("filename",A1))-1-SEARCH("[",CELL("filename",A1)))`**  
Result: MyWorkbook

### Get the active sheet name:  
Formula: **`=RIGHT(CELL("Filename",A1),LEN(CELL("Filename",A1))-FIND("]",CELL("Filename",A1)))`**  
Result: Sheet1



### Reference another sheet name:
Formula: **`="'"&MID(CELL("filename",A1),FIND("]",CELL("filename",A1))+1,256)&"'!"`**  
Result: 'Sheet1!
