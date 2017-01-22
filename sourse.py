#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2016/12/24 15:22
# @Author  : ProY
# @Site    : 
# @File    : sourse.py
# @Software: PyCharm
import os
from excelDic import *
from xlwt import *

path_dir = os.path.join(os.path.abspath('.'), 'path')
path_dir_2 = os.path.join(os.path.abspath('.'), 'newpath')
files = file_names()
for file in files:
	new_workbook = Workbook()
	workbook = xlrd.open_workbook(os.path.join(path_dir, file))

	fac_name = file.split('.')[0]
	i = 0
	for sheet in workbook.sheets():
		try:
			if '日期' in sheet.cell_value(1,0) or '时间' in sheet.cell_value(1,0):
				new_sheet = new_workbook.add_sheet(str(i))
				fist_row(new_sheet)
				get_all(sheet, new_sheet, fac_name)
			else:
				continue
		except:
			continue
	try:
		new_workbook.save(os.path.join(path_dir_2, file))
	except IndexError:
		continue