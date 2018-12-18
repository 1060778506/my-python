import xlrd
import copy
import os
import xlwt
from xlutils.copy import copy
from xlrd import open_workbook


data = xlrd.open_workbook('1.xlsx')#打开
table = data.sheets()[0]
nrows = table.nrows#多少行
k=0
q = 0
while 1==1:
	try:
		cell_q7 = table.cell(k,6).value#获取G列多少数据
	except IndexError:
		break;
	else:
		k = k + 1
while k > q:
	w = 0
	rb = open_workbook('1.xlsx')#打开原版文件
	rs = rb.sheet_by_index(0)#打开
	wb = copy(rb)#复制
	ws = wb.get_sheet(0)#获取
	ws.write(q, 7, int(0))#写入
	ws.write(q, 8, int(0))
	ws.write(q, 9, int(0))
	ws.write(q, 10, int(0))
	wb.save('1.xlsx')#保存
	cell_A6 = table.cell(q,6).value
	while nrows > w:
		cell_A0 = table.cell(w,0).value
		if cell_A6 == cell_A0:
			data = xlrd.open_workbook('1.xlsx')#打开
			table = data.sheets()[0]
			cell_q7 = int(table.cell(q,7).value)#获取
			cell_q8 = int(table.cell(q,8).value)
			cell_q9 = int(table.cell(q,9).value)
			cell_q10 = int(table.cell(q,10).value)
			cell_w1 = int(table.cell(w,1).value)
			cell_w2 = int(table.cell(w,2).value)
			cell_w3 = int(table.cell(w,3).value)
			cell_w4 = int(table.cell(w,4).value)
			rb = open_workbook('1.xlsx')#打开原版文件
			rs = rb.sheet_by_index(0)#打开
			wb = copy(rb)#复制
			ws = wb.get_sheet(0)#获取
			ws.write(q, 7, int(cell_w1)+int(cell_q7))#写入
			ws.write(q, 8, int(cell_w2)+int(cell_q8))
			ws.write(q, 9, int(cell_w3)+int(cell_q9))
			ws.write(q, 10, int(cell_w4)+int(cell_q10))
			wb.save('1.xlsx')#保存
			print("一共是"+str(k*nrows)+"条数据"+"\n"+"已经找到"+"\n"+"目前处理了"+str(q*k)+"\n"+"\n")
		else:
			print("一共是"+str(k*nrows)+"条数据"+"\n"+"没有找到"+"\n"+"目前处理了"+str(q*k)+"\n"+"\n")
		w = w + 1
	q = q + 1
input()
