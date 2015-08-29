ExcelToWiki
-----------


Use is trivial as shown below::

    from exceltowiki import exceltowiki
    e2w = excelToWiki("./test.xlsx")
    # print sheet names in the excel workbook
    print e2w.sheetnames
    # print wiki text for sheet named Sheet1
    print e2w.getSheet("Sheet1")
    # print wiki text for entire workbook
    print e2w.getWorkbook()

Options are::

    exceltowiki(excelworkbook, [list of sheet names to process], caption foreground color, caption background color)

Where caption is set from the sheet name (no way to currently modify this).

Features
--------

exceltowiki can capture:

- Font styling: bold, underline, strikethrough
- Cell styling: foregroudn color, background color
- Sheet features: merged cells are captured, sheet name is captured as caption to the wiki table


exceltowiki currently cannot capture anything more complex than the above list. Features such as 'format as table', conditional formatting, and other advanced items are not inspected or captured. For these, only the data value in the cells will be captured.

Release Notes: 0.1.10
---------------------
Bug fix: Unicode was not correctly handled. 
Packaging was not following best practice of examples within the package.

Release Notes: 0.1.9
--------------------
Minor font issues fixed. Italics and font-name were being ignored.
Some other minor items fixed.

Release Notes: 0.1.8
--------------------
Added support for hyperlinks:

- Unfortunately openpyxl does not yet support reading hyperlinks.
- The way to enter hyperlinks is to place the hyperlink and wiki display text in a cell, 
  such as: "http://yahoo.com Yahoo!"
- Any cell containing "http" will be wrapped within []s.
Bugs:
-----

- Caption was missing, added it back.
- Caption style was incorrectly set. fixed.
- unneeded parameter headerRow removed.
- unneeded cellToWiki method removed


Release Notes: 0.1.7
--------------------
Cleaner output of wiki text. 
- Common cell styles across the row are boiled up to row style.
- Common row style items are boiled up to table.


Release Notes: 0.1.6
--------------------
Minor: black was being ignored for bg color as well. Instead of only the fg color

Release Notes: 0.1.5
--------------------

* Added border as default.
* Removed font color from markup when color is black
