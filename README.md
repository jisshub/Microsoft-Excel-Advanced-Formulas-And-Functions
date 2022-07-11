# Excel Formula Syntax

## Match function

![](./screenshots/match_function.jpg)


## YEAR FUNCTION

- fetching year from a column and populating them in a new column.

![](./screenshots/year_function.jpg)

- Hetre v want to populate the birthyear field with year value of the birthdate column.
- to do that, we use YEAR() and pass the birthdate column as an argument.
Eg: =YEAR(F2), =YEAR(F3) etc.

    - F2 means column F row 2.
    - F3 means column F row 3.

- Once filling one column, we just drag till the last column and then double click on the column. OR DOUBLE CLICK ON THE RIGHT BUTTOM EDGE OF THE COLUMN.

- Result looks like this:

![](./screenshots/year_function_1.jpg)

## LEFT FUNCTION

![](./screenshots/left_fn1.jpg)

- Here in Area Code column, we have to fill it with first 3 digits of the telephone number.

### Formula:

**=LEFT(H2, 3)**

    - Fetching the left most character from a column and populating them in a new column.
    - H2 means column H row 2.
    - 3 means fetching the first 3 characters.

![](./screenshots/left_fn_solution.jpg)
### Formula:

**=LEFT(J3, FIND("@", J3)-1)**

    - Fetching the left most character from a column and populating them in a new column.
    - J3 means column J and row 3.
    - FIND("@", J3)-1 means fetching the first character before the @ character.


![](./screenshots/LEFT_FN_2.jpg)

# Reference Types

![](./screenshots/reference_types.jpg)

## Problem Statement

What is the starting balance for the given 10 years ?

![](./screenshots/reference_types_1.jpg)


Aim is to find the starting balance for the given 10 years.

### Formula:

**Sample:**

1000 + (1000 * (1 + 0.05) ^ 10)

    - 1000 is the starting balance.
    - (1 + 0.05) ^ 10 is the interest rate.
    - 10 is the number of years.


### Excel Formula:

    **=D4*(1+$C4)**

    - D4 means column D row 4.
    - C4 means column C row 4.
    - $ means the column is a reference type, that column is fixed one while computation.
![](./screenshots/reference_types_2.jpg)

- find 8% of the 1000, then add the result with 1000.
- Again find 8% of the result, then add the result of previous output.

### Final Result:

![](./screenshots/reference_types_3.jpg)

# Excel Error Types

## Type 1

Error - ######
![](./screenshots/error_type_1.jpg)


## Type 2

Error - #NAME?

![](./screenshots/error_type_2.jpg)

- This error occurs when we misspell the formula function.
- For example, we misspelled the function as LET instead of LEFT.

## Type 3

Error - #VALUE!

- This error occurs when we perform an arithmetic operation on text strings.
- For example, we perform an arithmetic operation on a string.
- For example, we perform an arithmetic operation on a date.


## Type 4 (#DIV/0!)

![](./screenshots/error_type_4.jpg)


## Type 5 (#REF!)

It is a reference error.
**#REF!** error happens when we delete any important columns or rows from our sheets.

![](./screenshots/error_type_4.jpg)


### Screenshots

![](./screenshots/error_type_5.jpg)

- We will delete the Growth Rate column and then try to perform the formula.
- As a result of this all columns will be filled with #REF! error.

![](./screenshots/error_type_6.jpg)

- All the cells linked with the Growth Rate column will be filled with #REF! error.


## Type 6 (#N/A)

![](./screenshots/error_type_7.jpg)

- This error mostly occurs when we search for a match in a record and sadly we don't find any match. This is N/A error.

# Formula Auditing: Trace Precedent & Dependents

![](./screenshots/trace_dependent_1.jpg)

# Navigating Excel Workbook with Ctrl Shortcuts.

**Ctrl + A** - Jumps to last cell in a data region, in the direction of arrow.

**Ctrl + Shift + Arrow** - Selects to the last cell in a data region, in the direction of the arrow.

**Ctrl + Home/End** - Jumps to the Home(top-left) and End(bottom-right) cell in a data region.

**Ctrl + .**  - Jumps Straight to each corner within a selected cell range.

**Ctrl + PageUP/PageDown** - Switches worksheet tabs, either to left or right.


## Examples:

**Ctrl + Shift + Down Arrow** - To select to the last cell in a column.

## Function Shortcuts

### F1
    - Launches the help window.
    - Links to the microsoft support

### F2
    - Allows you to edit the active cell.
    - Highlights cells referenced by active formula.

### F4
    - Repeats the last action takes.
    - Toggles absolute/relative cell reference with a formula.

### F9
    - Calculates all workbook formulas.
    - Evaluates each function argument within the formula bar.



