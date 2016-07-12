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
		'''
		从xls表格中读取数据
		sheet_filter 是sheet名过滤器
		cell 是单元格描述数组
		cell[i] = (name,row,col)
		'''
		data_wbs = []
		for file in self.files:
			# print file
			if re.search(r'.xls$',file):
				# logging.error(file)
				# logging.error(type(file))
				try:
					wb = xlrd.open_workbook(file)
				except:
					logging.error('fail to open ',file)
					continue
				data_wb = DataWorkbook(file)
				sht_list = []
				for sht in wb.sheets():
					if sheet_filter:
						if not re.search(sheet_filter,sht.name):
							continue
					data_sheet = DataSheet(sht.name)
					cell_list = []
					for name,row,col in cells: 
						try:
							cell_type = sht.cell_type(row,col)
							if cell_type == 5:#公式错误
								cell_value = "Error"
							else:
								cell_value = sht.cell_value(row,col)
								#滤除value中的换行符
								if isinstance(cell_value,str):
									# logging.warning(cell_value)
									# print 'new line',cell_value
									cell_value = cell_value.strip('\n')#.replace('\r',"")

						except:
							logging.error('array index out of range for cell(%s,%d,%d) in sheet %s in file ' %(name,row,col,sht.name))
							cell_value = None
							# continue
						cell_list.append(DataCell(name,row,col,cell_value))
					data_sheet.cells = cell_list
					sht_list.append(data_sheet)
				data_wb.sheets = sht_list
				data_wbs.append(data_wb)
		self.data_wbs = data_wbs
		return data_wb

	def save_to_json(self,file_name="",obj = None):
		if file_name == "":
			file_name = "data" + self.get_datetime_str() + ".json"
		if obj == None:
			obj = self.data_wbs
		json_str = json.dumps(obj,default = lambda o:o.__dict__,indent = 4,ensure_ascii = False)
		self.save_to_file(file_name,json_str)
	
	def save_to_file(self,result_file_name = "",content ="hello world"):
		cur_dir = os.path.dirname(os.path.abspath(__file__))

		if result_file_name == "":
			result_file_name = 'save' + self.get_datetime_str() +".txt"
		with open(os.path.join(cur_dir,result_file_name),"w") as result_file:
			result_file.write(content.encode('utf-8'))

		print "Saved to file: " + result_file_name
		return result_file_name
	
	def get_datetime_str(self):
		return datetime.datetime.now().strftime('%y%m%d_%H-%M-%S')

	def print_data(self):
		#输出标题行
		# title_line = "".join([str(cell.name) for cell in self.data_wbs[0].sheets[0].cells])
		# print title_line
		title_line = "file_name,"
		for cell in self.data_wbs[0].sheets[0].cells:
			title_line += cell.name+","
		print title_line


		for wb in self.data_wbs:
			for sht in wb.sheets:
				# data_line = ''
				print wb.name+",",
				for cell in sht.cells:
					print "%s," %(cell.value),
				print ""
					# data_line += str(cell.value) 
				# data_line = "".join([str(cell.value) for cell in self.data_wbs[0].sheets[0].cells])
				# print data_line
				# print "\n"




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
		self.name = name.decode('gbk')#处理中文文件名问题
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
		sheet_filter = u'包',
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

	hand.save_to_json()
def get_sheet_names(folder):
	hand = ExcelHandler()
	hand.find_filenames(folder)
	hand.get_data_from_xls()
	for wb in hand.data_wbs:
		for sht in wb.sheets:
			print wb.name,sht.name
		# print "-"*100
		# print wb.name
		# print [sht.name for sht in wb.sheets]

		# print len(wb.sheets)
		# for sht in wb.sheets:
		# 	print sht.name,
def get_bid_data():
	folder = r"C:\Users\adam\Desktop\bidding_xls"
	handler = ExcelHandler()
	handler.find_filenames(folder)

	handler.get_data_from_xls(
		sheet_filter = u"基本信息",
		cells =[
		(u'工作号',1,1),
		(u'开标时间',2,1),
		(u'销售员',3,1),
		(u'项目单位',4,1),
		(u'项目名称',5,1),
		(u'csc专员',6,1),
		])
	handler.save_to_json(u'基本信息.json')
	handler.print_data()
	handler.get_data_from_xls(
		sheet_filter = ur'[(包)(标段)(开标)(门)]',
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
	handler.save_to_json(u'包数据.json')
	handler.print_data()




if __name__ == "__main__":
	# unit_test2()
	# get_sheet_names(r"C:\Users\adam\Desktop\bidding_xls")
	get_bid_data()
