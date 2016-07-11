# -*- coding: utf-8 -*-
import os
import re
import logging
# import openpyxl
# import xlrd


class ExcelHandler():
	def __init__(self,folder_path=None):
		pass

	def find_filenames(self,folder_path,incld_path = True,re_filter = None):
		'''
		incld_path: 输出是否包含目录
		re_filter : 文件名过滤器,用正则表达式
		# ext_list: 允许的后缀名,不用带"点""
		'''
		for root,dirs,files in os.walk(folder_path):
			results = []
			if re_filter:
				if incld_path:
					results = [os.path.join(root,f) for f in files if re.search(re_filter,f) ]
				else:
					results = [f for f in files if re.search(re_filter,f)]
			else:
				if incld_path:
					results = [os.path.join(root,f) for f in files]
				else:
					results = files
			self.files = results
			return results






# class Filter():
# 	def all_in(self,str,filter_strs):

# 	pass


if __name__ == "__main__":
	hand = ExcelHandler()
	cur_dir = os.path.dirname(os.path.abspath(__file__))
	dir = cur_dir
	dir = r"D:\kuaipan\github\2016sis"
	
	files = hand.find_filenames(dir,incld_path = True,re_filter =r'py$')
	for f in files:
		print f
	# print files