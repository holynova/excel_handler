# -*- coding: utf-8 -*-
import os
import re
import logging
import openpyxl
import xlrd
import json
import datetime


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

	def get_data_from_xls(self,sheet_filter = None,cells = []):
		for file in self.files:
			if re.search(r'.xls$',file):
				try:
					wb = xlrd.open_workbook(file)
				except:
					logging.error('fail to open ',file)
					continue
				data_wb = DataWorkbook(file)

				for sht in wb.sheets():
					if sheet_filter:
						if not re.search(sheet_filter,sht.name):
							continue
					data_sheet = DataSheet(sht.name)
					# logging.warning('sht'+sht.name)
					# data_sheet.cells = [DataCell(name,row,col,sht.cell_value(row,col)) for name,row,col in cells]
					# logging.error('in sheet loop '+sht.name)
					cell_list = []
					for name,row,col in cells: 
						# logging.error('in cell loop '+name+","+str(row)+","+str(col))
						# try:
						# 	cell_value = sht.cell_value(row,col)
						# except:
						# 	logging.error('cell(%d,%d) value error' %(row,col)) 
						# 	cell_value = None
						cell_value = "test"
						logging.error( sht.cell_type(row,col))
						
						cell_list.append(DataCell(name,row,col,cell_value))
						# data_sheet.cells.append(DataCell(name,row,col,cell_value))
						# logging.error(len(data_sheet.cells))
					data_sheet.cells = cell_list

					data_wb.sheets.append(data_sheet)
		self.data_wb = data_wb
		return data_wb

	def save_to_json(self,file_name="",obj = None):
		if file_name == "":
			file_name = "data" + self.get_datetime_str() + ".json"
		if obj == None:
			obj = self.data_wb
		json_str = json.dumps(obj,default = lambda o:o.__dict__,indent = 4)
		self.save_to_file(file_name)
	
	def save_to_file(self,result_file_name = "",content ="hello world"):
		if result_file_name == "":
			result_file_name = 'save' + self.get_datetime_str() +".txt"
		with open(result_file_name,"w") as result_file:
			result_file.write(content.encode('utf-8'))

		print "Saved to file: " + result_file_name
		return result_file_name
	
	def get_datetime_str(self):
		return datetime.datetime.now().strftime('%y%m%d_%H-%M-%S')


class DataCell():
	def __init__(self,name,row = -1,col = -1,value = None):
		"""
		self定义,注意xlrd中行和列的index是从0开始的,而openpyxl是从1开始的
		"""
		self.name = name
		self.row = row
		self.col = col
		self.value = value
	def __repr__(self):
		return 'class self:name = %s,row = %d,col = %d,value = %s' %(self.name,self.row,self.col,self.value)
	# def
class DataSheet():
	def __init__(self,name,cells = []):
		self.name = name
		self.cells = cells

class DataWorkbook():
	def __init__(self,name,sheets = []):
		self.name = name 
		self.sheets = sheets



def unit_test():
	hand = ExcelHandler()
	print hand.get_datetime_str()
	hand.save_to_file()
	hand.save_to_file("sym.txt",u"大哥回答过")

def unit_test2():
	hand = ExcelHandler()
	folder = r"E:\kuaipan\github\excel_handler\test_xls"
	files = hand.find_filenames(folder,re_filter = r'.xls$')
	# for f in files:
	# 	print f
	hand.get_data_from_xls(
		# sheet_filter = u'包',
		cells = [
		(u"num",0,19),
		(u"nkt_gm3",1,19),
		(u"num_company",2,19),
		(u"winner",3,19),
		(u"min",4,19),
		(u"max",5,19),
		(u"average",6,19),
		(u"average_no_peak",7,19),
		(u"median",8,19),
		(u"winner_price",9,19),
		(u"nkt_price",10,19)
		])
	print hand.data_wb.name
	for sht in hand.data_wb.sheets:
		print '--',sht.name
		for cell in sht.cells:
			print '----',cell.name,cell.value
	# data_wb = DataWorkbook()
	# print dir(hand.data_wb),dir(hand.data_wb.sheets[0]),dir(hand.data_wb.sheets[0].cells[0])
	# hand.save_to_json()
	# print json.dumps(hand.data_wb, default = lambda o:o.__dict__, indent = 4)
	# print json.dumps(hand.data_wb, indent = 4)
	# print hand.data_wb.sheets




if __name__ == "__main__":
	unit_test2()
