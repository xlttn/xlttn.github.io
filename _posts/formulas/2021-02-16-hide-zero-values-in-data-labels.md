---
Title: Hide Zero (0) Values in Data Labels
categories: [Excel, Formulas]
tags: [interface-formatting]
date: 2021-02-16 18:43:00

---

So you have a 0% value on one of your data labels and want to hide it?

The quick and easy way to accomplish this is to custom format your data label.

1. Select a data label.
2. Right click and select Format Data Labels
3. Choose the Number category in the Format Data Labels dialog box.
4. Select Custom in the Category box.
5. In the format code box, enter 0%;-0%; and click Add.
6. Close out of your dialog box and your 0% labels should be gone.


![repeat-values-n-times-img](/imgs/hide-zero-values-in-data-labels/hide-zero-values-before-after.png)

![repeat-values-n-times-img](/imgs/hide-zero-values-in-data-labels/hide-zero-values-data-labels.png)
