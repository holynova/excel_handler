# -*- coding: utf-8 -*-
from excel_handler import *
if __name__ == "__main__":
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
	handler.save_to_json()
	handler.print_data()
	# handler.get_data_from_xls(
	# 	sheet_filter = ur'[(包)(标段)(开标)(门)]',
	# 	cells = [
	# 	(u"num",0,19),
	# 	(u"nkt_gm3",1,19),
	# 	(u"num_company",2,19),
	# 	(u"winner",3,19),
	# 	(u"min",4,19),
	# 	(u"max",5,19),
	# 	(u"average",6,19),
	# 	(u"average_no_peak",7,19),
	# 	(u"median",8,19),
	# 	(u"winner_price",9,19),
	# 	(u"nkt_price",10,19)
	# 	])
	# handler.save_to_json(u'包数据.json')
	# handler.print_data()



