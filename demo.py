# -*- coding: utf-8 -*-
from excel_handler import *
def find_old():
	folder = r"C:\Users\adam\Desktop\old2.8"
	handler = ExcelHandler()
	handler.get_filenames(folder,re_filter = r'.xls$')


	handler.get_data_from_xls(
		sheet_filter = u"Evaluation",
		cells =[
		(u'合同额',9,3),
		(u'gm1',16,4),
		(u'gm2',21,4),
		(u'gm3',27,4),
		(u'cm1',29,4),
		(u'spc',28,3),
	

	
		])
	handler.save_to_json()
	handler.print_data()
def find_new():
	folder = r"C:\Users\adam\Desktop\new2.8"
	handler = ExcelHandler()
	handler.get_filenames(folder,re_filter = r'.xls$')


	handler.get_data_from_xls(
		sheet_filter = u"Evaluation",
		cells =[
		(u'合同额',13,2),
		(u'gm1',20,3),
		(u'gm2',24,3),
		(u'gm3',29,3),
		(u'cm1',33,3),
		(u'gross_profit',36,3),
		(u'spc',28,2),
		

		])
	handler.save_to_json()
	handler.print_data()

def find_new2():
	folder = r"C:\Users\adam\Desktop\new2.8"
	handler = ExcelHandler()
	handler.get_filenames(folder,re_filter = r'.xls$')


	handler.get_data_from_xls(
		sheet_filter = u"Calculation",
		cells =[
		(u'型号1',9,2),
		(u'数量1',9,3),
		(u'单价1',9,6),
		(u'小计1',9,7),

		(u'型号2',10,2),
		(u'数量2',10,3),
		(u'单价2',10,6),
		(u'小计2',10,7),
		
		(u'型号3',11,2),
		(u'数量3',11,3),
		(u'单价3',11,6),
		(u'小计3',11,7),
		
		(u'型号4',12,2),
		(u'数量4',12,3),
		(u'单价4',12,6),
		(u'小计4',12,7),
	
		(u'包总价',8,7),
		
		

		])
	handler.save_to_json()
	handler.print_data(to_print = False,to_save = True)
def find_new3():
	h = ExcelHandler()
	folder = r"C:\Users\adam\Desktop\new2.8"
	cells = [
		(u'型号1',9,2),
		(u'数量1',9,3),
		(u'单价1',9,6),
		(u'小计1',9,7),

		(u'型号2',10,2),
		(u'数量2',10,3),
		(u'单价2',10,6),
		(u'小计2',10,7),
		
		(u'型号3',11,2),
		(u'数量3',11,3),
		(u'单价3',11,6),
		(u'小计3',11,7),
		
		(u'型号4',12,2),
		(u'数量4',12,3),
		(u'单价4',12,6),
		(u'小计4',12,7),
	
		(u'包总价',8,7),
			]
	h.get_and_print_data(
		folder = folder,
		file_filter = r".xls$",
		sheet_filter = "Calculation",
		cells = cells,
		to_print = False,
		to_save = True)
def find_fujian():
	h = ExcelHandler()
	h.write_data_to_xlsx(
		folder = ur"E:\kuaipan\非投标任务\2016年7月21日 福建开标记录",
		wb_filter = r".xlsx$",
		sheet_filter = r"Sheet",
		cells=[
		(u'最高价',1,5),
		(u'最低价',2,5),
		(u'去极值均价',3,5),
		(u'中位数',4,5),
		(u'安凯特价格',5,5),
		(u'厂家数',6,5),
		(u'安凯特排名',7,5),
	
		('=max(c:c)',1,6),
		('=min(c:c)',2,6),
		('=trimmean(c:c,0.05)',3,6),
		('=median(c:c)',4,6),
		('=INDEX(C:C,MATCH("常州安凯特电缆有限公司",B:B,0))',5,6),
		('=max(a:a)',6,6),
		('=INDEX(a:a,MATCH("常州安凯特电缆有限公司",B:B,0))',7,6),
		])

if __name__ == "__main__":
	# find_old()
	# find_new()
	# find_new3()
	find_fujian()




