#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2016/12/24 15:22
# @Author  : ProY
# @Site    : 
# @File    : excelDic.py
# @Software: PyCharm-

import xlrd,xlwt
import os

# 返回所有excel文件名字
def file_names():
	path_dir = os.path.join(os.path.abspath('.'), 'path')
	files = [x for x in os.listdir(path_dir) if os.path.splitext(x)[1] == '.xls' or os.path.splitext(x)[1] == 'xlsx']
	files = list(map(lambda x : os.path.join(path_dir, x), files))
	return files

# 返回该文件名对应的厂名
def fac_name(whole_filename):
	filename = whole_filename.split('.')[0]
	while(filename[0].isdigit()):
		filename = filename[1:]
	return filename.replace(' ', '')

def get_value(sheet, r, c):
	if sheet.cell_value(r, c):
		return sheet.cell_value(r, c)
	else:
		return get_value(sheet, r-1, c)

def get_date(old_sheet, new_sheet, number):
	col = 0
	for row in range(2, 10000):
		try:
			if old_sheet.cell_type(row, col) == 0 or old_sheet.cell_type(row, col) == 3:
				date_value = get_value(old_sheet, row, col)
				date_value = xlrd.xldate_as_tuple(date_value, 0)
				date_value = str(date_value[0])+'-'+str(date_value[1])+'-'+str(date_value[2])
				new_sheet.write(row-1, col, date_value)
				new_sheet.write(row-1, 10, number)
			elif old_sheet.cell_type(row, col) == 1:
				break
		except:
			break

def get_name(old_sheet, new_sheet):
	for coly in range(0, 20):
		rowx = 1
		try:
			if old_sheet.cell_value(rowx, coly) == '产品名称':
				col = coly
				break
		except:
			return
	for row in range(2, 10000):
		try:
			if old_sheet.cell_type(row, col) == 1:
				name_value = old_sheet.cell_value(row, col)
				new_sheet.write(row-1, 1, name_value)
			elif old_sheet.cell_type(row, col) == 0:
				if old_sheet.cell_type(row, 0) == 1:
					break
				else:
					name_value = get_value(old_sheet, row, col)
					new_sheet.write(row - 1, 1, name_value)
		except:
			break

def get_cpxh(old_sheet, new_sheet):
	for coly in range(0, 20):
		rowx = 1
		try:
			if old_sheet.cell_value(rowx, coly) == '产品代号' or old_sheet.cell_value(rowx, coly) == '图号' or old_sheet.cell_value(rowx, coly) == '型号':
				col = coly
				break
		except:
			return

	for row in range(2, 10000):
		if old_sheet.cell_type(row, col) == 1:
			name_value = old_sheet.cell_value(row, col)
			new_sheet.write(row-1, 2, name_value)
		elif old_sheet.cell_type(row, col) == 0:
			if old_sheet.cell_type(row, 0) == 1:
				break
			else:
				name_value = get_value(old_sheet, row, col)
				new_sheet.write(row-1, 2, name_value)

def get_ggxh(old_sheet, new_sheet):
	for coly in range(0, 20):
		rowx = 1
		try:
			if old_sheet.cell_value(rowx, coly) == '规格型号' or old_sheet.cell_value(rowx, coly) == '规格':
				col = coly
				break
		except:
			return
	for row in range(2, 10000):
		try:
			if old_sheet.cell_type(row, col) == 1:
				name_value = old_sheet.cell_value(row, col)
				new_sheet.write(row-1, 3, name_value)
			elif old_sheet.cell_type(row, col) == 0:
				if old_sheet.cell_type(row, 0) == 1:
					break
				else:
					name_value = get_value(old_sheet, row, col)
					new_sheet.write(row-1, 3, name_value)
		except:
			break

def get_qj(old_sheet, new_sheet):
	for coly in range(0, 20):
		rowx = 1
		try:
			if old_sheet.cell_value(rowx, coly) == '嵌件':
				col = coly
				new_col = 4
				break
		except:
			return
	for row in range(2, 10000):
		if old_sheet.cell_type(row, col) == 1:
			name_value = old_sheet.cell_value(row, col)
			new_sheet.write(row-1, new_col, name_value)
		elif old_sheet.cell_type(row, col) == 0:
			if old_sheet.cell_type(row, 0) == 1:
				break
			else:
				name_value = get_value(old_sheet, row, col)
				new_sheet.write(row-1, new_col, name_value)

def get_sl(old_sheet, new_sheet):
	for coly in range(0, 20):
		rowx = 1
		try:
			if old_sheet.cell_value(rowx, coly) == '数量' or old_sheet.cell_value(rowx, coly) == '对账数量' or old_sheet.cell_value(rowx, coly) == '入库数量':
				col = coly
				new_col = 5
				break
		except:
			return
	for row in range(2, 10000):
		if old_sheet.cell_type(row, col) == 2 or old_sheet.cell_type(row, col) == 1:
			name_value = old_sheet.cell_value(row, col)
			new_sheet.write(row-1, new_col, name_value)
		elif old_sheet.cell_type(row, col) == 0:
			if old_sheet.cell_type(row, 0) == 1:
				break
			else:
				name_value = get_value(old_sheet, row, col)
				new_sheet.write(row-1, new_col, name_value)

def get_sfsl(old_sheet, new_sheet):
	for coly in range(0, 20):
		rowx = 1
		try:
			if old_sheet.cell_value(rowx, coly) == '实发数量':
				col = coly
				new_col = 6
				break
		except:
			return
	for row in range(2, 10000):
		try:
			if old_sheet.cell_type(row, col) == 2 or old_sheet.cell_type(row, col) == 1:
				name_value = old_sheet.cell_value(row, col)
				new_sheet.write(row-1, new_col, name_value)
			elif old_sheet.cell_type(row, col) == 0:
				if old_sheet.cell_type(row, 0) == 1:
					break
				else:
					name_value = get_value(old_sheet, row, col)
					new_sheet.write(row-1, new_col, name_value)
		except:
			break

def get_dj(old_sheet, new_sheet):
	for coly in range(0, 20):
		rowx = 1
		try:
			if old_sheet.cell_value(rowx, coly) == '单价':
				col = coly
				new_col = 7
				break
		except:
			return
	for row in range(2, 10000):
		try:
			if old_sheet.cell_type(row, col) == 2 or old_sheet.cell_type(row, col) == 1:
				name_value = old_sheet.cell_value(row, col)
				new_sheet.write(row-1, new_col, name_value)
			elif old_sheet.cell_type(row, col) == 0:
				if old_sheet.cell_type(row, 0) == 1:
					break
				else:
					name_value = get_value(old_sheet, row, col)
					if name_value == '单价':
						new_sheet.write(row-1, new_col, 0)
						continue
					new_sheet.write(row-1, new_col, name_value)
		except:
			break

def get_je(old_sheet, new_sheet):
	for coly in range(0, 20):
		rowx = 1
		try:
			if old_sheet.cell_value(rowx, coly) == '金额':
				col = coly
				new_col = 8
				break
		except:
			return
	for row in range(2, 10000):
		try:
			if old_sheet.cell_type(row, col) == 2 or old_sheet.cell_type(row, col) == 1:
				if old_sheet.cell_type(row, 0) == 1:
					break
				name_value = old_sheet.cell_value(row, col)
				new_sheet.write(row-1, new_col, name_value)
			elif old_sheet.cell_type(row, col) == 0:
				if old_sheet.cell_type(row, 0) == 1:
					break
				else:
					name_value = get_value(old_sheet, row, col)
					if name_value == '金额':
						new_sheet.write(row-1, new_col, 0)
						continue
					new_sheet.write(row-1, new_col, name_value)
		except:
			break

def get_bz(old_sheet, new_sheet):
	for coly in range(0, 20):
		rowx = 1
		try:
			if old_sheet.cell_value(rowx, coly) == '备注':
				col = coly
				break
		except:
			return
	for row in range(2, 10000):
		if old_sheet.cell_type(row, col) == 1:
			name_value = old_sheet.cell_value(row, col)
			new_sheet.write(row-1, 9, name_value)
		elif old_sheet.cell_type(row, col) == 0:
			if old_sheet.cell_type(row, 0) == 1:
				break
			else:
				name_value = get_value(old_sheet, row, col)
				new_sheet.write(row-1, 9, name_value)

def get_all(old_sheet, new_sheet, number):

	get_date(old_sheet, new_sheet, number)
	get_name(old_sheet, new_sheet)
	get_cpxh(old_sheet, new_sheet)
	get_ggxh(old_sheet, new_sheet)
	get_qj(old_sheet, new_sheet)
	get_sl(old_sheet, new_sheet)
	get_sfsl(old_sheet,new_sheet)
	get_dj(old_sheet,new_sheet)
	get_je(old_sheet,new_sheet)
	get_bz(old_sheet,new_sheet)