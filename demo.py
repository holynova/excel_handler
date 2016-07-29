# -*- coding: utf-8 -*-
from excel_handler import *
def find_old_qef():
	folder = r"C:\Users\adam\Desktop\fujian3\old"
	handler = ExcelHandler()
	handler.get_and_print_data(
		folder = folder,
		file_filter = r".xls$",
	 	sheet_filter = "Evaluation",
		cells=[
		(u'合同额',9,3),
		(u'gm1',16,4),
		(u'gm2',21,4),
		(u'gm3',27,4),
		(u'cm1',29,4),
		(u'spc',28,3),],
		to_print = False,
		to_save = True,
		to_json = False)

def find_new_qef():
	folder = r"C:\Users\adam\Desktop\xls_sgcc"
	handler = ExcelHandler()
	handler.get_and_print_data(
		folder = folder,
		file_filter = r".xls$",
	 	sheet_filter = "Evaluation",
		cells=[
		(u'合同额',13,2),
		(u'gm1',20,3),
		(u'gm2',24,3),
		(u'gm3',29,3),
		(u'cm',33,3),
		(u'gross_profit',36,3),
		(u'spc',28,2),],
		to_print = False,
		to_save = True,
		to_json = False)



def write_data():
	h = ExcelHandler()
	h.write_data_to_xlsx(
		folder = ur"E:\kuaipan\非投标任务\2016年7月26日 三省开标记录",
		wb_filter = ur"江苏10kv电缆开标记录.xlsx$",
		sheet_filter = r"Sheet",
		cells=[
		(u'最高价',1,5),
		(u'最低价',2,5),
		(u'去极值均价',3,5),
		(u'中位数',4,5),
		(u'安凯特价格',5,5),
		(u'厂家数',6,5),
		(u'安凯特排名',7,5),
		(u'nkt新排名',8,5),
		# (u'nkt新价格',109,2),
		# (u'=1/10000',109,3),
		# ('',104,3),
		


		
	
		('=max(c:c)',1,6),
		('=min(c:c)',2,6),
		('=trimmean(c:c,0.05)',3,6),
		('=median(c:c)',4,6),
		('=INDEX(C:C,MATCH("常州安凯特电缆有限公司",B:B,0))',5,6),
		('=max(a:a)',6,6),
		('=INDEX(a:a,MATCH("常州安凯特电缆有限公司",B:B,0))',7,6),
		# ('=RANK(C109,C:C,1)',8,6),
		])
def merge_order():
	handler = ExcelHandler()
	handler.merge_sheets(
		folder = r"C:\Users\adam\Desktop\order",
		wb_filter = r".xlsx$",
		sheet_filter = r"2016",
		key_column =1,
		header_rows = 1
		)
def get_unit_price_from_new_qef():
	folder = r"C:\Users\adam\Desktop\xls_sgcc"
	handler = ExcelHandler()
	# cells = [(u"序号"+str(i),i,6) for i in range(9,18)]
	cells = []
	for i in range(9,18):
		cells.append((u"序号"+str(i),i,1))
		cells.append((u"单价"+str(i),i,6))

	handler.get_and_print_data(
		folder = folder,
		file_filter = r".xls$",
	 	sheet_filter = "Calculation",
		cells=cells,
		to_print = False,
		to_save = True,
		to_json = False)

if __name__ == "__main__":
	# write_data()
	# find_new_qef()
	# find_old_qef()
	# merge_order()

	get_unit_price_from_new_qef()

