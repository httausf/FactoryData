#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2017/1/24 18:52
# @Author  : Aries
# @Site    : 
# @File    : test.py
# @Software: ProYan

import xlrd
import os

def file_names():
	path_dir = os.path.join(os.path.abspath('.'), 'path')
	files = [x for x in os.listdir(path_dir) if os.path.splitext(x)[1] == '.xls' or os.path.splitext(x)[1] == 'xlsx']
	files = list(map(lambda x : os.path.join(path_dir, x), files))
	return files

print(file_names())