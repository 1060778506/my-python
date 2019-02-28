import requests#请求内容
import re#正则
import os#操作系统库
import time#时间
import xlwt#生成xlsxl表格
import xlrd#修改xlsxl表格
import json#json数据获取
import datetime#获取时间
from datetime import date#兼容获取时间
from xlutils.copy import copy#复制xlsxl表格
from xlrd import open_workbook#打开xlsxl表格




data = xlrd.open_workbook('淘宝链接.xlsx')#打开
table = data.sheets()[0]#打开
nrows = table.nrows#多少行
u=0
k=0
while nrows>k:
	rb = open_workbook('淘宝链接.xlsx')#打开原版文件
	wb = copy(rb)#复制打开原版文件
	ws = wb.get_sheet(0)#获取打开原版文件
	cell_a = table.cell(k,0).value#获取a列数据
	cell_b = table.cell(k,1).value#获取b列数据
	pattern = re.compile(r'[0-9]+')#正则
	zes = str(pattern.findall(cell_b))#正则匹配出ID
	datas = zes.replace('[\'', '').replace('\']', '')#替换不用的等于号
	begin = str(datetime.date(2019,1,29))#开始时间这里填写月单位
	end = str(datetime.date(2019,2,27))#结束时间
	endin = begin + " " + "——" + " " + end
	data = xlrd.open_workbook('淘宝链接.xlsx')#打开
	table = data.sheets()[0]#打开
	url = 'https://sycm.taobao.com/cc/item/sale/overview.json?dateType=recent30&dateRange='+begin+'%7C'+end+'&dateType=day&device=0&itemId='+datas#链接组合month&是月；week&是周；recent30&是30天；recent7&是7天
	cookies = dict(cookie2='187de4b3724d5e9cf483a86475f17930',csg='7703e16f')#发送密钥
	null=""#
	joinspayAmt = eval(requests.post(url,cookies=cookies,).text)#组合发送并且将格式str转为dict
	joinpayAmt = joinspayAmt['data']['payAmt']['value']#获取当月销售额
	url = 'https://sycm.taobao.com/cc/item/sale/sku/list.json?dateType=recent30&dateRange='+begin+'%7C'+end+'&dateType=day&pageSize=10&page=1&order=desc&orderBy=cartCnt&device=0&itemId='+datas+'&indexCode=cartCnt%2CpayAmt%2CpayItmCnt%2CpayByrCnt&_=524819609940'#获取SKU
	cookies = dict(cookie2='187de4b3724d5e9cf483a86475f17930',csg='7703e16f')#发送密钥
	joins = eval(requests.post(url,cookies=cookies,).text)#组合发送并且将格式str转为dict
	joinss = len(joins)#剩余sku数量
	join = joins['data']['recordCount']#多少个SKU
	p = join/10#获取页数
	if p != 0 and p < 1 or p ==1:
		while 1>0:
			try:
				cell_q7 = type(table.cell(u,7).value)#获取F列多少数据
			except IndexError:
				break
			else:
				if "float" in str(cell_q7):
					u = u + 1
				else:
					break
		page0=0
		while page0 < join:
			nn = joins['data']['data'][page0]#SKU名称1
			skuName = nn['skuName']['value']#SKU名称2
			#payAmts = joins['data']['data'][page0]#SKU销量
			payAmt = nn['payAmt']['value']#SKU销量
			ws.write(page0+u, 2, endin)#写入时间
			ws.write(page0+u, 3, cell_a)#写入标题
			ws.write(page0+u, 4, cell_b)#写入连接
			ws.write(page0+u, 5, joinpayAmt)#写入当日连接销售总额
			ws.write(page0+u, 6, skuName)#写入SKU名字
			ws.write(page0+u, 7, payAmt)#写入SKU销量
			page0=page0+1
		page0=0
		wb.save('淘宝链接.xlsx')#保存
	elif p != 0 and p > 1 and p < 2 or p == 2:

		while 1>0:
			try:
				cell_q7 = type(table.cell(u,7).value)#获取F列多少数据
			except IndexError:
				break
			else:
				if "float" in str(cell_q7):
					u = u + 1
				else:
					break
		page0=0
		while page0 < 10:
			nn = joins['data']['data'][page0]#SKU名称1
			skuName = nn['skuName']['value']#SKU名称2
			#payAmts = joins['data']['data'][page0]#SKU销量
			payAmt = nn['payAmt']['value']#SKU销量
			ws.write(page0+u, 2, endin)#写入时间
			ws.write(page0+u, 3, cell_a)#写入标题
			ws.write(page0+u, 4, cell_b)#写入连接
			ws.write(page0+u, 5, joinpayAmt)#写入当日连接销售总额
			ws.write(page0+u, 6, skuName)#写入SKU名字
			ws.write(page0+u, 7, payAmt)#写入SKU销量
			page0=page0+1
		page0=0
		wb.save('淘宝链接.xlsx')#保存

		url = 'https://sycm.taobao.com/cc/item/sale/sku/list.json?dateType=recent30&dateRange='+begin+'%7C'+end+'&dateType=day&pageSize=10&page=2&order=desc&orderBy=cartCnt&device=0&itemId='+datas+'&indexCode=cartCnt%2CpayAmt%2CpayItmCnt%2CpayByrCnt&_=524819609940'#获取SKU
		cookies = dict(cookie2='187de4b3724d5e9cf483a86475f17930',csg='7703e16f')#发送密钥
		joins = eval(requests.post(url,cookies=cookies,).text)#组合发送并且将格式str转为dict
		joinss = len(joins)#剩余sku数量
		rb = open_workbook('淘宝链接.xlsx')#打开原版文件
		wb = copy(rb)#复制打开原版文件
		ws = wb.get_sheet(0)#获取打开原版文件
		data = xlrd.open_workbook('淘宝链接.xlsx')#打开
		table = data.sheets()[0]#打开

		while 1>0:
			try:
				cell_q7 = type(table.cell(u,7).value)#获取F列多少数据
			except IndexError:
				break
			else:
				if "float" in str(cell_q7):
					u = u + 1
				else:
					break
		page0=0
		while page0 < joinss-10:
			nn = joins['data']['data'][page0]#SKU名称1
			skuName = nn['skuName']['value']#SKU名称2
			#payAmts = joins['data']['data'][page0]#SKU销量
			payAmt = nn['payAmt']['value']#SKU销量
			ws.write(page0+u, 2, endin)#写入时间
			ws.write(page0+u, 3, cell_a)#写入标题
			ws.write(page0+u, 4, cell_b)#写入连接
			ws.write(page0+u, 5, joinpayAmt)#写入当日连接销售总额
			ws.write(page0+u, 6, skuName)#写入SKU名字
			ws.write(page0+u, 7, payAmt)#写入SKU销量
			page0=page0+1
		wb.save('淘宝链接.xlsx')#保存

	elif p != 0 and p > 2 and p < 3 or p == 3:
		while 1>0:
			try:
				cell_q7 = type(table.cell(u,7).value)#获取F列多少数据
			except IndexError:
				break
			else:
				if "float" in str(cell_q7):
					u = u + 1
				else:
					break
		page0=0
		while page0 < 10:
			nn = joins['data']['data'][page0]#SKU名称1
			skuName = nn['skuName']['value']#SKU名称2
			payAmt = nn['payAmt']['value']#SKU销量
			ws.write(u+page0, 2, endin)#写入时间
			ws.write(u+page0, 3, cell_a)#写入标题
			ws.write(u+page0, 4, cell_b)#写入连接
			ws.write(u+page0, 5, joinpayAmt)#写入当日连接销售总额
			ws.write(u+page0, 6, skuName)#写入SKU名字
			ws.write(u+page0, 7, payAmt)#写入SKU销量
			page0=page0+1
		wb.save('淘宝链接.xlsx')#保存
		url = 'https://sycm.taobao.com/cc/item/sale/sku/list.json?dateType=recent30&dateRange='+begin+'%7C'+end+'&dateType=day&pageSize=10&page=2&order=desc&orderBy=cartCnt&device=0&itemId='+datas+'&indexCode=cartCnt%2CpayAmt%2CpayItmCnt%2CpayByrCnt&_=524819609940'#获取SKU
		cookies = dict(cookie2='187de4b3724d5e9cf483a86475f17930',csg='7703e16f')#发送密钥
		joins = eval(requests.post(url,cookies=cookies,).text)#组合发送并且将格式str转为dict
		rb = open_workbook('淘宝链接.xlsx')#打开原版文件
		wb = copy(rb)#复制打开原版文件
		ws = wb.get_sheet(0)#获取打开原版文件
		data = xlrd.open_workbook('淘宝链接.xlsx')#打开
		table = data.sheets()[0]#打开
		while 1>0:
			try:
				cell_q7 = type(table.cell(u,7).value)#获取F列多少数据
			except IndexError:
				break
			else:
				if "float" in str(cell_q7):
					u = u + 1
				else:
					break
		page0=0
		while page0 < 10:
			nn = joins['data']['data'][page0]#SKU名称1
			skuName = nn['skuName']['value']#SKU名称2
			payAmt = nn['payAmt']['value']#SKU销量
			ws.write(u+page0, 2, endin)#写入时间
			ws.write(u+page0, 3, cell_a)#写入标题
			ws.write(u+page0, 4, cell_b)#写入连接
			ws.write(u+page0, 5, joinpayAmt)#写入当日连接销售总额
			ws.write(u+page0, 6, skuName)#写入SKU名字
			ws.write(u+page0, 7, payAmt)#写入SKU销量
			page0=page0+1
		wb.save('淘宝链接.xlsx')#保存
		url = 'https://sycm.taobao.com/cc/item/sale/sku/list.json?dateType=recent30&dateRange='+begin+'%7C'+end+'&dateType=day&pageSize=10&page=3&order=desc&orderBy=cartCnt&device=0&itemId='+datas+'&indexCode=cartCnt%2CpayAmt%2CpayItmCnt%2CpayByrCnt&_=524819609940'#获取SKU
		cookies = dict(cookie2='187de4b3724d5e9cf483a86475f17930',csg='7703e16f')#发送密钥
		joins = eval(requests.post(url,cookies=cookies,).text)#组合发送并且将格式str转为dict
		joinss = len(joins)#剩余sku数量
		rb = open_workbook('淘宝链接.xlsx')#打开原版文件
		wb = copy(rb)#复制打开原版文件
		ws = wb.get_sheet(0)#获取打开原版文件
		data = xlrd.open_workbook('淘宝链接.xlsx')#打开
		table = data.sheets()[0]#打开
		while 1>0:
			try:
				cell_q7 = type(table.cell(u,7).value)#获取F列多少数据
			except IndexError:
				break
			else:
				if "float" in str(cell_q7):
					u = u + 1
				else:
					break
		page0=0
		while page0 < joinss-20:
			nn = joins['data']['data'][page0]#SKU名称1
			skuName = nn['skuName']['value']#SKU名称2
			payAmt = nn['payAmt']['value']#SKU销量
			ws.write(u+page0, 2, endin)#写入时间
			ws.write(u+page0, 3, cell_a)#写入标题
			ws.write(u+page0, 4, cell_b)#写入连接
			ws.write(u+page0, 5, joinpayAmt)#写入当日连接销售总额
			ws.write(u+page0, 6, skuName)#写入SKU名字
			ws.write(u+page0, 7, payAmt)#写入SKU销量
			page0=page0+1
		wb.save('淘宝链接.xlsx')#保存
	else:
		while 1>0:
			try:
				cell_q7 = type(table.cell(u,7).value)#获取F列多少数据
			except IndexError:
				break
			else:
				if "float" in str(cell_q7):
					u = u + 1
				else:
					break
		ws.write(u, 2, endin)#写入时间
		ws.write(u, 3, cell_a)#写入标题
		ws.write(u, 4, cell_b)#写入连接
		ws.write(u, 5, joinpayAmt)#写入当日连接销售总额
		ws.write(u, 6, "没有SKU")#写入当日连接销售总额
		ws.write(u, 7, 0)#写入当日连接销售总额
		wb.save('淘宝链接.xlsx')#保存
	k=k+1
print("结束")
input()
