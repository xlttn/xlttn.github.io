---
Title: Insert Special Characters or Symbols into a Text String
categories: [Excel, Formulas]
tags: [text-strings]  
date: 2018-10-05

---

You can always use the Insert --> Symbols method to enter characters onto your spreadsheet but you can also add them via a formula.  

#### Formula
```vb
=CHAR()
```

**Result:** Exchange Rates Chart (£), Euro (€) and Dollar ($)  
```vb
="Exchange Rates Chart (" & CHAR(163) & "), Euro (" & CHAR(128) & ") and Dollar (" & (CHAR(36) & ")")
```

Using the below table as a guide, you can see that typing Alt + 0 + (code number) also produces the special character or symbol.

You can always put this in Row 1 and drag down to Row 255 for all symbols
```vb
=CHAR(ROW())
```


|        | A          | B          | C          | D               |
|--------|------------|------------|------------|-----------------|
| **1**  | **Name**   | **Number** | **Symbol** | **Code**        |
| **2**  | Quotation  | 34         | "          | *Alt + 0 + *34  |
| **3**  | Hash       | 35         | #          | *Alt + 0 + *35  |
| **4**  | Dollar     | 36         | $          | *Alt + 0 + *36  |
| **5**  | Star       | 42         | *          | *Alt + 0 + *42  |
| **6**  | Euro       | 128        | €          | *Alt + 0 + *128 |
| **7**  | Ellipsis   | 133        | …          | *Alt + 0 + *133 |
| **8**  | Dagger     | 134        | †          | *Alt + 0 + *134 |
| **9**  | Double     | 135        | ‡          | *Alt + 0 + *135 |
| **10** | Bullet     | 149        | •          | *Alt + 0 + *149 |
| **11** | Trademark  | 153        | ™          | *Alt + 0 + *153 |
| **12** | Cents      | 162        | ¢          | *Alt + 0 + *162 |
| **13** | Pounds     | 163        | £          | *Alt + 0 + *163 |
| **14** | Yen        | 165        | ¥          | *Alt + 0 + *165 |
| **15** | Copyright  | 169        | ©          | *Alt + 0 + *169 |
| **16** | Registered | 174        | ®          | *Alt + 0 + *174 |
| **17** | Plus/Minus | 177        | ±          | *Alt + 0 + *177 |
| **18** | Paragraph  | 182        | ¶          | *Alt + 0 + *182 |
| **19** | Degree     | 176        | °          | *Alt + 0 + *176 |
