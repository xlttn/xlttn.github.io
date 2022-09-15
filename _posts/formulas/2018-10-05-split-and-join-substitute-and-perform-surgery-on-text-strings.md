---
Title: Split, Join, Substitute and Perform Surgery on Text Strings
categories: [Excel, Formulas]
tags: [text-strings]  
date: 2018-10-05

---
#### Convert text to lower case
Example: **`=LOWER("hello")`**  
Result:	hello


#### Convert text to upper case
Example: **`=UPPER(D3)`**  
Result:	JAMES


#### Convert text to proper case (each word's first letter capitalized)
Example: **`=PROPER("this is a long sentence")`**  
Result:	This Is A Long Sentence


#### Combine different text values to one text
Example: **`=CONCATENATE(A3, " and ", A4)`**  
Result:	30 and 25


#### Combine different text values to one text
Example: **`=A3&" and "&A4`**  
Result:	30 and 25


#### Extract first few letters from a text
Example: **`=LEFT("Australia",3)`**  
Result: Aus


#### Extract last few letters from a text
Example: **`=RIGHT("New Zealand",4)`**  
Result: land


#### Extract middle portion from given text
Example: **`=MID("United States",4,5)`**  
Result: ted S


#### What is the length of given text value
Example: **`=LEN("Titan")`**  
Result:	5


#### Substitute one word with another
Example: **`=SUBSTITUTE("Microsoft Excel","cel","cellent")`**  
Result:	Microsoft Excellent


#### Replace some letters with other
Example: **`=REPLACE("XYZ123",4,3,"456")`**  
Result:	XYZ456


#### Find if a text has another text
Example: **`=FIND("soft","Microsoft Excel")`**  
Result:	6


#### Extract initials from a name
A1 contains Sue Flay  
Example: **`=LEFT(A1,1)&MID(A1,FIND(" ",A1)+1,1)`**  
Result:	SF


#### Find out how many words are in a sentence
A1 contains "This is a very long sentence with lots of words"  
Example: **`=LEN(A1)-LEN(SUBSTITUTE(A1," ",""))+1`**  
Result:	10


#### Remove unnecessary spaces from a cell
Example: **`=TRIM("  why  so    serious   ")`**  
Result	why so serious


#### Remove anything after a symbol or word
A1 contains Excel, Titan
Example: **`=LEFT(A1,FIND(",",A1)-1)`**  
Result:	Excel
