ExcelToWiki
-----------


Use is trivial as shown below::
	>>> from exceltowiki import exceltowiki
	>>> e2w = excelToWiki("./test.xlsx")
	>>> # print sheet names in the excel workbook
	>>> print e2w.sheetnames
	>>> # print wiki text for sheet named Sheet1
	>>> print e2w.getSheet("Sheet1")
	>>> # print wiki text for entire workbook
	>>> print e2w.getWorkbook()

Options are::
	>>> exceltowiki(excelworkbook, [list of sheet names to process], caption foreground color, caption background color)
Caption is the sheet name.

Features
--------

exceltowiki can capture:
- Font styling: bold, underline, strikethrough
- Cell styling: foregroudn color, background color
- Sheet features: merged cells are captured, sheet name is captured as caption to the wiki table


exceltowiki currently cannot capture anything more complex than the above list. Features such as 'format as table', conditional formatting, and other advanced items are not inspected or captured. For these, only the data value in the cells will be captured.

Release Notes: 0.1.5
--------------------
* Added border as default.
* Removed font color from markup when color is black
