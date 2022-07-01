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

    
=LEFT(H2, 3)

    - Fetching the left most character from a column and populating them in a new column.
    - H2 means column H row 2.
    - 3 means fetching the first 3 characters.

=LEFT(J3, FIND("@", J3)-1)

    - Fetching the left most character from a column and populating them in a new column.
    - J3 means column J row 3.
    - FIND("@", J3)-1 means fetching the first character before the @ character.