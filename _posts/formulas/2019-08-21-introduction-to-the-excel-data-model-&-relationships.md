---
Title: Introduction to the Excel Data Model & Relationships
categories: [Excel, Formulas]
tags: [tables]
date: 2019-08-21 18:43:0

---

## Have you ever been in a VLOOKUP hell?

Its what happens when you have to write a lot of vlookup formulas before you can start analyzing your data. Every day, millions of analysts and managers enter VLOOKUP hell and suffer. They connect table 1 with table 2 so that all the data needed for making that pivot report is on one place. If you are one of those, then you are going to love Excel’s data model & relationships feature.

In simple words, this feature helps you **connect one set of data with another set of data so that you can create combined pivot reports.**

## VLOOKUP hell vs. Data Model Heaven

Lets say you are looking sales data for your company. You have transaction data like below.

![data-model-img](/imgs/data-model/data-model1.png)

And you want to find out how many units you are selling by product category and customer’s gender.

Unfortunately, you only have product ID & customer ID.

### With VLOOKUP Hell

1. You first fetch all the customer and product data and place them in separate ranges.
2. Then write a vlookup formula to fetch product category, another to fetch customer gender.
3. Then fill down the formulas for entire list of transactions.
4. Now make a pivot table.

Assuming you have 30,000 transactions, you have to write 60,000 VLOOKUP formulas to create this one report.

### With Data Model heaven

1. Create relationships between Sales, Products & Customer tables
2. Create a pivot table

## Creating a relationship in the Data Model
1. First set up your data as tables. To create a table, select any cell in range and press CTRL+T. Specify a name for your table from design tab.
2. Now, go to data ribbon & click on relationships button.
3. Click New to create a new relationship.
4. Select Source table & column name. Map it to target table & column name. It does not matter which order you use here. Excel is smart enough to adjust the relationship.
5. Add more relationships as needed.

![data-model-img](/imgs/data-model/data-model2.png)

## Using relationships in Pivot reports & analysis
1. Select any table and insert a pivot table (Insert > Pivot table).
2. Make sure you check the "Add this data to data model" check box.
3. In your pivot table field list, check "ALL" instead of "ACTIVE" to see all table names.
4. Select fields from various tables to create a combined pivot report or pivot chart

![data-model-img](/imgs/data-model/data-model3.png)

## Example: Category and Area Sales Report
1. Add Category to rows labels
2. Add Area to columns labels
3. Add Quantity to values
4. and your report is ready!

![data-model-img](/imgs/data-model/data-model-pivot.png)

## Excel Recognises a Relationship Between Tables may be Needed
In the demo file, the example pivot table is a Category and Area Sales Report, however when I added the fields to the pivot table, Excel recognised that I hadn't yet made a relationship

![data-model-img](/imgs/data-model/data-model-relationship-needed.png)
![data-model-img](/imgs/data-model/data-model-relationship-created.png)
![data-model-img](/imgs/data-model/data-model-relationship-updated.png)

## Things to Keep in Mind When Using Relationships
- Same data types in both columns: Columns that you are connecting in both tables should have same data type (ie both numbers or dates or text etc.)
- One to one or One to many relationships only: Excel 2013 supports only one to many or one to one relationships. That means one of the tables must have no duplicate values on the column you are linking to. (for example products table should not have duplicate product IDs).
- You can add slicers too: You can slice these pivot tables on any field you want (just like normal pivot tables). For example, you can further slice the above report on customer’s profession or product’s SKU size.

## Download Example File
Click here to download the example file: [data-model-example.xlsx](/example-files/data-model-example.xlsx). It contains 3 different tables and a combined pivot report and a slicer to show you what is possible.
