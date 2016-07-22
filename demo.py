# -*- coding: utf-8 -*-
from excel_handler import *
def find_old_qef():
	folder = r"C:\Users\adam\Desktop\jiangsu\old"
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
	folder = r"C:\Users\adam\Desktop\jiangsu\new"
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
		folder = ur"E:\kuaipan\nkt文件存档\CSC工作流\CSC报价\2016年7月11日 代替王莺 福建电网报价\2016年7月22日 新模拟",
		wb_filter = ur"2016年7月21日 福建居配工程-开标一览表.xlsx$",
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
		(u'nkt新价格',109,2),
		(u'=1/10000',109,3),
		('',104,3),
		


		
	
		('=max(c:c)',1,6),
		('=min(c:c)',2,6),
		('=trimmean(c:c,0.05)',3,6),
		('=median(c:c)',4,6),
		('=INDEX(C:C,MATCH("常州安凯特电缆有限公司",B:B,0))',5,6),
		('=max(a:a)',6,6),
		('=INDEX(a:a,MATCH("常州安凯特电缆有限公司",B:B,0))',7,6),
		('=RANK(C109,C:C,1)',8,6),
		])

if __name__ == "__main__":
	# write_data()
	find_new_qef()
	find_old_qef()
	pass
	# find_old()
	# find_new()
	# find_new3()
	# find_fujian()
	# find_old()
	# find_old()
	# find_new()




