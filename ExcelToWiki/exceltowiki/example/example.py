'''
Created on Jul 28, 2015

@author: venkman69@yahoo.com
'''

from exceltowiki import excelToWiki

e2w = excelToWiki("./test.xlsx",["Sheet1"],"blue","yellow")
print e2w.sheetnames
print e2w.getSheet("Sheet1")
print e2w.getWorkbook()