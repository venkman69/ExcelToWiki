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
