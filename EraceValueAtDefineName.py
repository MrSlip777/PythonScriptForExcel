#定義された名前のセルを全て消去するスクリプト
#このスクリプトはMITライセンスです。
#改変等々ご自由にどうぞ

import openpyxl
import tkinter as window
from tkinter import filedialog
import glob
import os

#指定した名前定義の値を消去する
def EraceValue(wb,defineName):

	names = wb.defined_names
	name = names.get(defineName)

	#シート名と範囲で分割
	str1 = name.attr_text
	str2 = str1.split('!')

	sheet = wb[str2[0]]

	CellRange = str2[1].split(':')

	for rows in sheet[CellRange[0]:CellRange[1]]:
    		for cell in rows:
        		cell.value = ''

#Main関数
def Main():
	dir = "C:"
	fld = filedialog.askdirectory(initialdir = dir) 
	for filename in glob.glob(fld + '/*.xlsx'):
		wb =openpyxl.load_workbook(os.path.abspath(filename))

		#指定した文字列の値を消去する
		EraceValue(wb,'文字')
		EraceValue(wb,'文字2')

		wb.save(os.path.abspath(filename))

Main()