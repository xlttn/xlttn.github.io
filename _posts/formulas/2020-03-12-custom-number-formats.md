---
Title: Custom Number Formats
categories: [Excel, Formulas]
tags: [interface-formatting]
date: 2020-03-12 18:43:00

---

<!-- For ease of reading the .md file I have added this section in as there as backslash \ characters used in the formatting rules, the backslash is also a way to escape special characters, to avoid a mess, I will repeat the rules in a comment to avoid confusion.
-->
# What Are Custom Number Formats?

Custom number formats control how numbers are look in Excel. The key benefit is that they change how a number looks without changing any data.

Before we get started, here's an awesome resource [Excel Custom Format Builder](https://customformats.com/)

This concept is used most frequently with dates.

## Excel Custom Number Format Rules - Dates

| Input       | Code       | Result     |
|-------------|------------|------------|
| 20-Apr-2020 | yyyy       | 2020       |
| 20-Apr-2020 | yy         | 20         |
| 20-Apr-2020 | mmm        | Apr        |
| 20-Apr-2020 | mmmm       | April      |
| 20-Apr-2020 | d          | 20         |
| 20-Apr-2020 | ddd        | Mon        |
| 20-Apr-2020 | dddd       | Monday     |
| 20-Apr-2020 | d/mm/yyyy  | 20/04/2020 |
| 20-Apr-2020 | dd/m/yyyy  | 20/4/2020  |
| 20-Apr-2020 | dd mm yyyy | 20 04 2020 |
| 20-Apr-2020 | General    | 43941      |
| 20-Apr-2020 | mmm 'yy    | Apr '20    |

A date is just a numerical number at it’s core, changing the number format to General format shows a date's true value.

## 4 Parts of a Number Format Rule

There are four parts or sections to a Custom Number Format rule. The **first section is required** while the additional three are optional. Each section is divided up by the use of a semi-colon **( ; )**. Here is what each part of the number format rule represents:

1.  If the number is positive then do this…
2.  If the number is negative then do this…
3.  If the number equals zero then do this…
4.  If the value is not a number then do this…   

```
Positive; Negative; Zero; Text
```

### A Few Caveats

- If only the first section has a format rule, it will be applied to all numerical values whether positive or negative (text values will be left alone)
- If only the first two sections have format rules, zero values will use the positive value format

# Using the Number Format Editor
In order to write you own custom number format rules, you will need to navigate to the rule editor. The editor resides within the Format Cells dialog box where you can modify all the properties/formats of a cell.

There are multiple ways to navigate to the Format Cells dialog box:

-   **Method 1:** Right-click on cell >> Select Format Cells…
-   **Method 2:** Home Tab >> Number Button group >> click Grey Arrow in bottom corner
-   **Method 3:** Use the Keyboard Shortcut: ctrl + 1 (PC) or cmd + 1 (Mac)

Once you have opened the Format Cells dialog box, you will want to navigate to the Number tab. This tab will show you a bunch of preset number format rules you can navigate through or if you would like to write your own rule, you can navigate all the way to the bottom of the Category Pane and click Custom.

![number-format](/imgs/custom-number-format/number-format-editor.png)

## Special Characters & What They Do

There are a few special characters you can utilise while writing a Custom Number Format rule to add even more varieties to your value’s appearance. Let’s first look at the special characters available to you and then we will get into some examples.

| Character | What It Does                                                |
|:---------:|:------------------------------------------------------------|
|     @     | A placeholder for text                                      |
|     ,     | Separates thousands                                         |
|     0     | Forces the display of a numerical value                     |
|     #     | Placeholder for an optional digit                           |
|     ?     | Used to align digits at various lengths                     |
|     _     | Add a space sized as the character immediately following it |
|     *     | Repeats character immediately following it                  |

Just by including one of these symbols, your Custom Number Format rule will automatically use it’s special ability. If you wish to include one of these symbols without their ability, see how to do so in the “Escaping” section of this article (scroll down a few sections).

### @ Symbol

The @ symbol is used to control where your text value shows up in your rule. You can place modifications to your text value before or after your text via relocating the @ symbol within the rule.

| Value   | Appearance Needed             | Format Rule  | Output                              |
|:--------|:------------------------------|:-------------|:------------------------------------|
| Car     | turn text font colour blue    | [Blue]@      | <span style="color:blue">Car</span> |
| i99603  | append 'ID' to the text       | @" - ID"     | i99603 - ID                         |
| Sparrow | prepent 'Captain' to the text | "Captain. "@ | Captain. Sparrow                    |

### Comma Symbol
The comma symbol can be used to separate your numbers thousands or to round large numbers to a specific place (Millions, Billions, etc…).

If you place a comma in front of your “ones” place, you will gain the ability to see a comma separate your value every three places. You only need to use a single comma in order to trigger this format.

If you place a comma behind your “ones” place, the value will VISUALLY lose three places (essentially dividing 1,000). This behavior continues to occur for each additional comma you add behind your 'ones' place.

|   Value  | Appearance Needed              | Format Rule   |     Output   |
|---------:|:-------------------------------|:-------------:|-------------:|
|  1608047 | add thousands separator        | #,##0         |    1,608,047 |
| -1608047 | add thousands separator        | #,##0         |   -1,608,047 |
|  1608047 | show in millions (2 decimals)  | 0,,.00        |         1.61 |
| -1608047 | add negative () with separator | #,##0;(#,##0) |  (1,608,047) |

### Zero Number
Using a zero in a number format rule will force that number place to be shown visually. If you would like all your numbers to show three digits, insert three zeros into your rule and 1 will equal 001. This can be very usual in cleaning up your numerical values to ensure they all visually align with one another.

| Value | Appearance Needed   | Format Rule | Output     |
|-------|---------------------|-------------|------------|
| 1     | two number places   | 00          | 01         |
| 1     | two number places   | 00.         | 01.        |
| 1     | three number places | 000         | 001        |
| 6824  | id string           | 00-000-00   | 00-0682-40 |

### Pound/Hash Symbol
The pound (or hash) symbol serves as an optional placeholder for digits if they exist. If you value exceeds the number of pound signs to the right of your decimal, the format rule will round your value to align with your designated amount of pound symbols. If your value has less digits than pound symbols, a zero will not populate in its place.

| Value   | Appearance Needed              | Format Rule | Output   |
|---------|--------------------------------|-------------|----------|
| 63.7915 | two number places              | ##          | 64       |
| 63.7915 | two number places with decimal | ##.##       | 63.79    |
| 63.7915 | three number places            | ##.###      | 63.792   |
| 63.7915 | id string                      | ##.##.##    | 63.79.15 |

### Question Mark Symbol

Question marks can be used to align digits when you don’t necessarily want zeros to show up as numerical placeholders. When a question mark resides in a place where no value is provided, a space will be added (shown in grey below) to maintain the alignment of the number.

| Value  | Appearance Needed         | Format Rule |                                 Output |
|--------|---------------------------|-------------|---------------------------------------:|
| 63.7   | align to 3 decimal places | 0.???       | 63.7<span style="color:Gray">00</span> |
| 63.79  | align to 3 decimal places | 0.???       | 63.79<span style="color:Gray">0</span> |
| 63.791 | align to 3 decimal places | 0.???       |                                 63.791 |
| 63     | align to 3 decimal places | 0.???       | 63.<span style="color:Gray">000</span> |

### Underscore Symbol
By using the underscore symbol you can add a single space either before or after your cell value. The character immediately following the underscore determines the size of the space. In most cases, Excel users use the underscore symbol to line up positive and negative numbers that use parenthesis.

<!--
add a 'W' sized space after number is 0_W
add two 'w' spaces in front of number is _w_w0
add a ')' sized space after positive numbers is 0_);(0)  
-->

| Value | Appearance Needed                            | Format Rule |                              Output |
|-------|----------------------------------------------|-------------|------------------------------------:|
| 1     | add a 'W' sized space after number           | 0\_W        | 1<span style="color:Gray">W</span>  |
| 1     | add a ')' sized space after positive numbers | 0\_);(0)    | 1<span style="color:Gray">)</span>  |
| 1     | add two 'w' spaces in front of number        | \_w\_w0     | <span style="color:Gray">ww</span>1 |

### Escaping Special Characters
There may be instances where you literally want to use one of the above characters instead of utilizing their special abilities. To make remove the special ability (or “escape” the ability), just place a back slash before the character. You’ll need to place a backslash before each individual symbol you wish to escape.

<!--
- prepent three * to a number format rule is \*\*\*0000
- prepend @ to number is \@0
- add _ in front of last 2 digits is #\_00
-->

| Value    | Appearance Needed               | Format Rule      |    Output     |
|----------|---------------------------------|------------------|--------------:|
| 573      | prepend three * to number       | \\\*\\\*\\\*0000 |    \*\*\*0573 |
| 6492     | prepend @ to number             | \@0              |         @6492 |
| 26514274 | add _ in front of last 2 digits | #\\\_00          |     265142_74 |

### Adding Text
There may be occasions when you would like to add text before or after your values but still would like to perform spreadsheet math with your data.

| Value   | Appearance Needed                    | Format Rule                  |   Output |
|---------|--------------------------------------|------------------------------|---------:|
| 1       | append 'bps' to number               | 0" bps"                      |    1 bps |
| 2.3763  | prepend '$' and append 'M' to number | $0.0"M"                      |    $2.4M |
| 2.3     | categorize +/- numbers               | "Positive";"Negative";"Zero" | Positive |
| 150.231 | prepend 'id' to number               | "id-"0                       |   id-150 |

### Asterisk Symbol

An asterisk symbol can be used to fill the remaining space within a cell with the character immediately following it.

| Value  | Appearance Needed                    | Format Rule | Output                                                               |
|--------|--------------------------------------|-------------|----------------------------------------------------------------------|
| 592    | repeat period in front of number     | \*.0        |  \.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.592 |
| 5.642  | repeat period in front of number     | \*.0.00     |  \.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.592.64 |
| 645826 | repeat underscore in front of number | \*\_#,##0   |  \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_645,826 |
