---
Title: VBA Paste Values in Excel with Keyboard Shortcut
categories: [vba]
tags: [copy-data, practical]
date: 2019-08-08 18:43:00

---

This will show how to use VBA to paste values in Excel instead of formulas. You can bind this VBA macro to the Ctrl+Shift+V keyboard shortcut to make pasting even faster. The macro is smart enough to paste unformatted text, too.

> This is one of the best tips I’ve seen in quite a while. I can’t believe I didn’t think of it! I want you all to extend a big thank you to wellsrPRO power user, Mitch, for submitting this macro to me via the wellsrPRO add-in.  
[wellsr.com](https://wellsr.com/vba/2018/excel/vba-paste-values-in-excel-with-keyboard-shortcut/)

```vb
'==================================================================================================
' ## Paste as values or unformatted text from within or outside of Excel.
' Tip: assign this to a keyboard shortcut: Ctrl+Shift+V
' Developer: Mitch
'==================================================================================================
Sub PasteValues()
    '// first test if pasting from within excel, if an error then
    ' proceed to paste as unformatted text
    On Error Resume Next

    '// Paste as values
    Selection.PasteSpecial Paste:=xlPasteValues, _
    Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    '// Paste as unformatted text
    ActiveSheet.PasteSpecial Format:="Text", Link:=False, DisplayAsIcon:=False
End Sub
```

## How the VBA Paste Values Macro Works
First we to paste the contents of the clipboard two different ways. Let’s look at an example. In this example, the range <span style="background-color: #F2F2F2">A1:B4</span> contains raw numbers and the range <span style="background-color: #F2F2F2">D1:E4</span> squares those numbers using a formula. In the example, we have the Excel cells with formulas copied to our clipboard and we want to paste the values to another range.

![paste-values-img1](/imgs/paste-values/paste-values-img1.png)

First, we’ll show what happens if you copy and paste the cells like normal, using <span style="background-color: #F2F2F2">Ctrl+c</span> and <span style="background-color: #F2F2F2">Ctrl+v</span>.

![paste-values-img2](/imgs/paste-values/paste-values-img2.png)

Excel tries to outsmart you by pasting the formulas down to the relative cells. The range <span style="background-color: #F2F2F2">D6:E9</span> now contains formulas trying to square the results in cells <span style="background-color: #F2F2F2">A6:B9</span>. That’s not what we wanted! We want to paste the values themselves in the range <span style="background-color: #F2F2F2">D6:E9</span>.
To do that, we’ll run the PasteValues macro. The first thing the macro tries to do is paste the formulas as values. That works great! It will then try to paste the same content as unformatted text, but this won’t do anything and the output of the first attempt remains. Trying to paste cell values as unformatted text would normally generate an error, but we've bypassed the errors using the <span style="background-color: #F2F2F2">On Error Resume Next</span> code, so we get these results:

![paste-values-img3](/imgs/paste-values/paste-values-img3.png)

***So why we try to paste the results as unformatted text if it doesn’t work?***

Well, this is where this small macro's brilliance really becomes evident. To show you what I mean, we’ll pretend we have a table in Microsoft Word, like the one below:

![paste-values-img4](/imgs/paste-values/paste-values-img4.png)

We want to copy and paste the values from the table into Excel. If we copy the table from word and paste it normally, using <span style="background-color: #F2F2F2">Ctrl+v</span>, Excel will again try to outsmart you by pasting the format, ugly border and all, into Excel. You’ll be left with something rather hideous, like this:

![paste-values-img5](/imgs/paste-values/paste-values-img5.png)

Now we want to do the same thing, but instead of pasting with <span style="background-color: #F2F2F2">Ctrl+v</span>, we want to paste using the macro. Just like earlier, the PasteValues macro will attempt to paste the contents of your clipboard as values only. Since the contents of your clipboard didn’t originate from Excel, this would normally produce an error.  
However, since we've bypassed the errors using the <span style="background-color: #F2F2F2">On Error Resume Next</span> code, it tries the second method of pasting. This time, the second method, which pastes with unformatted text, works perfectly. No error is generated and you’re left with clean, unformatted, values pasted into Excel:

![paste-values-img6](/imgs/paste-values/paste-values-img6.png)

## Creating your Keyboard Shortcut

The icing on the cake is when you bind the PasteValues macro to a keyboard shortcut so you can paste values into Excel using VBA with just a keystroke. I’ll show you how!

From your Developer Tab, click the "Macros" button (Alt+F8). Select "PasteValues" then click "Options."  

![paste-values-img7](/imgs/paste-values/paste-values-img7.png)

In the Macro Options screen, type a capital V. Your screen will look something like this:

![paste-values-img8](/imgs/paste-values/paste-values-img8.png)

Press OK and close the Macros screen. Now, anytime you want to paste values in Excel instead of formulas, all you have to do is press Ctrl+Shift+V on your keyboard! It’s so convenient!
