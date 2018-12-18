import xlrd
import copy
import os
import xlwt
from xlutils.copy import copy
from xlrd import open_workbook

data = xlrd.open_workbook('1.xls')

table = data.sheets()[0]
nrows = table.nrows#行
q = 0
while nrows > q:
	cell_A4 = table.cell(q,0).value
	w = q+1
	q = q+1
k=0
while w > k:
	cell_A1 = table.cell(k,0).value
	k=k+1
	i=0
	while nrows > i:
		cell_A2 = table.cell(i,2).value
		if cell_A1==cell_A2:
			cell_A3 = table.cell(i,3).value
			rb = open_workbook('1.xls')
			rs = rb.sheet_by_index(0)
			wb = copy(rb)
			ws = wb.get_sheet(0)
			t=k-1
			ws.write(t, 1, cell_A3)
			wb.save('1.xls')
			print("匹配到",cell_A3,"第",t,"行");
		else:
			print("无匹配第",k,"行");
		i=i+1
