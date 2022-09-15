---
Title: Text Values, Strings, Words and Phrases
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
Example: **`=LEFT("India",3)`**  
Result: Ind


#### Extract last few letters from a text
Example: **`=RIGHT("New Zealand",4)`**  
Result: land


#### Extract middle portion from given text
Example: **`=MID("United States",4,5)`**  
Result: ted S


#### What is the length of given text value
Example: **`=LEN("Chandoo.org")`**  
Result:	11


#### Substitute one word with another
Example: **`=SUBSTITUTE("Microsoft Excel","cel","cellent")`**  
Result:	Microsoft Excellent


#### Replace some letters with other
Example: **`=REPLACE("abc@email.com",5,1,"g")`**  
Result:	abc @gmail.com


#### Find if a text has another text
Example: **`=FIND("soft","Microsoft Excel")`**  
Result:	6


#### Extract initials from a name
H1 contains Bill Jelen  
Example: **`=LEFT(H1,1)&MID(H1,FIND(" ",H1)+1,1)`**  
Result:	BJ


#### Find out how many words are in a sentence
H2 contains "This is a very long sentence with lots of words"  
Example: **`=LEN(H2)-LEN(SUBSTITUTE(H2," ",""))+1`**  
Result:	10


#### Remove unnecessary spaces from a cell
Example: **`=TRIM("  chandoo.  org   ")`**  
Result	chandoo. org


#### Remove anything after a symbol or word
H3 contains someone@ something.com  
Example: **`=LEFT(H3,FIND("@",H3)-1)`**  
Result:	someone
