ExcelToWiki
-----------


Use is trivial as shown below::

    from exceltowiki import exceltowiki 
    import exceltowiki

    e2w = excelToWiki("./test.xlsx",["Sheet1"],"blue","yellow")
    print e2w.sheetnames
    print e2w.getSheet("Sheet1")
    print e2w.getWorkbook()
    e2w = excelToWiki("./test.xlsx") 
    # print sheet names in the excel workbook 
    print e2w.sheetnames 
    # print wiki text for sheet named Sheet1 
    print e2w.getSheet("Sheet1") 
    # print wiki text for entire workbook 
    print e2w.getWorkbook() 

Options are:: 

    exceltowiki(excelworkbook, [list of sheet names to process], caption foreground color, caption background color, preserve widths) 

Where caption is set from the sheet name (no way to currently modify this). 

Features 
-------- 

exceltowiki can capture: 

- Font styling: bold, underline, strikethrough 
- Cell styling: foreground color, background color 
- Sheet features: merged cells are captured, sheet name is captured as caption to the wiki table 


exceltowiki currently cannot capture anything more complex than the above list. Features such as 'format as table', conditional formatting, and other advanced items are not inspected or captured. For these, only the data value in the cells will be captured. 

Release Notes: 0.1.19
--------------------- 
Added a better handler for date format detection and conversion. This is still very hacky but it works. 
MS date format to datetime strftime is never going to be straight forward until there is a method to 
obtain the resulting formatted string by itself.

Release Notes: 0.1.18
--------------------- 
Inline double-pipe separator had a bug where a newline in the cell requires the wiki table to use a single pipe character.

Release Notes: 0.1.17
--------------------- 
For better presentation of empty cells in wiki the cell contents are replaced with '&nbsp;' this retains any spacing/sizing related nature of the table which would otherwise collapse.

Release Notes: 0.1.16
--------------------- 
Wiki text within a cell was not being formatted via wiki because of being inlined. Slight update to fix this.
The 'INLINE_FMT' flag is now deprecated.


Release Notes: 0.1.15
--------------------- 
Added preserve_width option
Update to latest openpyxl - note that newer openpyxl broke something in colwidth with 0.1.14

Release Notes: 0.1.14
--------------------- 

Minor update


Release Notes: 0.1.13
--------------------- 

Added a version from git tag
Handling numeric, date and percent values a bit better

Release Notes: 0.1.12
--------------------- 

Cleanup release


Release Notes: 0.1.11 
--------------------- 
Added support for inline format with double-pipe notation.
Removed unneeded imports.
Minor fixes to value retrieval

Release Notes: 0.1.10 
--------------------- 

Packaging was not following best practice of examples within the package. 
Bug fix: Unicode was not correctly handled. 

Release Notes: 0.1.9 
-------------------- 

Minor font issues fixed. Italics and font-name were being ignored. 
Some other minor items fixed. 

Release Notes: 0.1.8 
-------------------- 

Added support for hyperlinks: 

- Unfortunately openpyxl does not yet support reading hyperlinks. 
- The way to enter hyperlinks is to place the hyperlink and wiki display text in a cell, such as: "http://yahoo.com Yahoo!" 
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
