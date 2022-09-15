---
Title: Getting started with Excel Tables
categories: [Excel, Formulas]
tags: [tables]
date: 2019-08-21 18:45:00
---

## What is an Excel table?

Excel TablesTable is your way of telling excel, "look, all this data from A1 to E25 is related. The row 1 has table headers. Right now we just have 24 rows of data. But I can add more later!"

When you make a table (more on this in a sec) you can easily add more rows to it without worrying about updating formula references, formatting options, filter settings etc. Excel will take care of everything thus making you a data guru.

## How to create table from a bunch of data?
To create an excel table, all you have to do is select a range of cells and press the table button from Insert ribbon in Excel (or use the shortcut CTRL+T).

![intro-to-tables-img](/imgs/intro-to-tables/introduction-to-tables-gif1.gif)

## The most important thing after you create a table – Give it a name  
Once you have a table, go to design ribbon and give your table a name. If you don’t name it, Excel will call it Table2 or whatever. But once you name it, you can write meaningful formulas thru sweet sweet structural references feature. ***So name your tables.***

## Change table formatting without lifting a finger  
Excel has some great predefined table formatting options. Just select any cell in your table and change the table formatting by going to "format as table" button in the home ribbon.  
If you are bored with the predefined formats, you can easily define your own table formatting color schemes and apply them.

![intro-to-tables-img](/imgs/intro-to-tables/introduction-to-tables-pic1.png)

## Add Banded Rows to Tables
When you create a table, banded rows come as a bonus. And when you add new rows to the table, excel takes care of this automatically. You can turn on / off the banded rows feature from "design ribbon tab" as well.  
That means you don’t need to use conditional formatting or manually format alternative rows in different color.

![intro-to-tables-img](/imgs/intro-to-tables/introduction-to-tables-gif2.gif)

## Tables Come With Data Filters and Sort Options by Default  
Each data table comes with filters and sorting options so that you can filter and sort the data in that table independently. That also means, if a worksheet has 2 tables, they each get their own data filters (usually excel wont allow you to add more than one set of filters per sheet, but when it comes to tables, all exceptions are made, just for you)

![intro-to-tables-img](/imgs/intro-to-tables/introduction-to-tables-pic2.png)


## You can also Slice your tables with slicers  
That is right. When you have a table of data, you can insert a slicer (either from design ribbon or insert ribbon) and use that to filter your table data intuitively.

## Bye, bye cell references, welcome structured references
The most important advantage of tables is that, you can write meaningful looking formulas instead of using cell references. When you create and name the table (you can name the table from design tab), you can write formulas that look like this:

![intro-to-tables-img](/imgs/intro-to-tables/introduction-to-tables-pic3.png)

The beauty of structured references is that, when you add or remove rows, you don’t need to worry about updating the references.

## Make Calculated Columns with ease
Any tabular data will have its share of calculated columns. Excel tables make having calculated columns very easy. With structured references, all you need to know is English to make a calculated column. The beauty of calculated columns in table is that, when you write formula in one cell, excel automatically fills the formula in the rest of cells in that column. That would make you an instant data guru.

![intro-to-tables-img](/imgs/intro-to-tables/introduction-to-tables-gif3.gif)

## Total your Tables without writing one formula
The ability to summarize data with pivot tables is extended to excel tables as well. You can add total row to your table with just a click.  
What more, you can easily change the summary type from "sum" to say "average".

![intro-to-tables-img](/imgs/intro-to-tables/introduction-to-tables-pic4.png)  

![intro-to-tables-img](/imgs/intro-to-tables/introduction-to-tables-pic5.png)

## Convert table back to a range, if you ever need to
If you ever wanted to go back to a normal range of data, you can easily convert the tables back to named ranges.

Excel will take care of the formulas and change the references to cell references.

## Export Tables to Pivot Tables, Woohoo
What good is a bunch of data when you can’t analyze it? That is where Pivot tables come in to picture. Thankfully, you don’t need to do much. Just click a button and your table goes to pivot table.

![intro-to-tables-img](/imgs/intro-to-tables/introduction-to-tables-pic6.png)  

![intro-to-tables-img](/imgs/intro-to-tables/introduction-to-tables-pic7.png)

## Print Tables Alone, with out all the other stuff around
Select the table, hit CTRL+P and in settings area, select "Print Selected Table" option to print your beautifully formatted Excel table

## Change, reshape or clean your table data with Power Query
When you have data in a table, you can easily load it to Power Query (Get & Transform Data) using the "From Table" button.

## Got multiple tables? Connect them to make a multi-table pivot
When you have more than one table, you can also connect them using Excel’s relationship feature. This way, you can build multi-table pivots to create powerful analysis of your data.

[Introduction to the Excel Data Model and Relationships](/formulas/introduction-to-the-excel-data-model-&-relationships/){:target="_blank"}

![intro-to-tables-img](/imgs/intro-to-tables/introduction-to-tables-pic8.png)
