#!/usr/bin/env python
#coding = utf-8

import os, re, time, random, requests, urllib.request

from bs4 import BeautifulSoup
from lxml import etree, html
from xlwt import *

# 请求 baseUrl
URL = 'https://book.douban.com/tag/'
# 每一个excel存储的数据量：TIMES * 20
TIMES = 5
# 防止封IP策略之一
# 伪装User-Agent，装作是浏览器请求，每次爬取时换一个请求头
USER_AGENTS = [
	'Opera/9.80 (Windows NT 6.1; U; en) Presto/2.8.131 Version/11.11',
	'Mozilla/5.0 (Windows NT 6.1; rv:2.0.1) Gecko/20100101 Firefox/4.0.1',
	'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.110 Safari/537.36',
	'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-us) AppleWebKit/534.50 (KHTML, like Gecko) Version/5.1 Safari/534.50'
]
# 存储路径
FOLDER_PATH = './books/'
# 表头
TABLE_HEADER = ['序号', '书名', '评分', '评价人数', '作者', '出版社', '出版日期', '售价']

# 获取全部标签
def getTags():
	headers = { 'User-Agent': random.choice(USER_AGENTS) }
	res  = requests.get(URL, headers = headers)
	tree = etree.HTML(res.text)
	tags = tree.xpath('//table[@class="tagCol"]//a/text()')
	return tags

# 获取表格数据
def getTableData(tag, index):
	headers  = { 'User-Agent': random.choice(USER_AGENTS) }
	params   = { 'start': index, 'type': 'T' }
	res      = requests.get(URL + tag, params = params, headers = headers)
	tree     = etree.HTML(res.text)
	bookList = tree.xpath('//li[@class="subject-item"]')

	if len(bookList) == 0:
		return False

	books = []
	for book in bookList:
		item = book[1]
		if item:
			index = index + 1
			detail = replaceEmpty(getValue(item.xpath('.//div[@class="pub"]/text()'), 0)).split(' / ')
			endIndex = -1

			order     = index
			bookName  = replaceEmpty(item.xpath('.//h2/a/text()')[0])
			bookName  = validateFileName(bookName)
			rating    = replaceEmpty(getValue(item.xpath('.//span[@class="rating_nums"]/text()'), 0))
			ratingNum = getNumber(replaceEmpty(item.xpath('.//span[@class="pl"]/text()')[0]))
			author    = replaceEmpty(detail[0])
			price     = getNumber(getValue(detail, endIndex))

			if price != '未知':
				endIndex = endIndex - 1

			date      = getDate(replaceEmpty(getValue(detail, endIndex)))

			if date != '未知':
				endIndex = endIndex - 1

			product   = replaceEmpty(getValue(detail, endIndex))
			bookInfo  = [order, bookName, rating, ratingNum, author, product, date, price] # 一行数据

			books.append(bookInfo)
	return books

# 确定是否有值
def getValue(arr, index):
	val = '未知'
	try:
		return arr[index]
	except BaseException:
		return val

# 提取价格
def getNumber(num_str):
	rstr = r'\d+(\.\d+)?'
	res = re.search(rstr, num_str)
	if res == None:
		return '未知'
	return res.group()

# 判断是否日期
def getDate(date_str):
	rstr = r'\d+(\.\d+)?'
	res = re.search(rstr, date_str)
	if res == None:
		return '未知'
	return date_str

# 去除换行与前后空格
def replaceEmpty(str):
	return str.replace('\n', '').strip()

# 替换不合法的文件字符
def validateFileName(title):
	rstr = r'[\/\\\:\*\?\"\<\>\|]'
	new_title = re.sub(rstr, '_', title)
	return new_title

# 保存到excel
def saveToExcel(tableData, fileName):
	file = Workbook(encoding = 'utf-8')
	table = file.add_sheet('Book Info')
	style = getTableStyle(table)

	index = 0
	for header in TABLE_HEADER: # 写入表头
		table.write(0, index, header, style['header_style'])
		index = index + 1

	for rowIndex, rowData in enumerate(tableData): # 写入表数据
		for colIndex, colItem in enumerate(rowData):
			table.write(rowIndex + 1, colIndex, colItem, style['content_style'])

	file.save(FOLDER_PATH + fileName + '.xls')
	print('\n==========%s 表格数据写入完毕==========\n' % fileName)

# 获取表格样式
def getTableStyle(table):
	header_style       = XFStyle() # 表头样式
	header_font        = Font()
	header_font.height = 20 * 18 # 字体大小
	header_font.bold   = True # 加粗
	header_style.font  = header_font

	content_style       = XFStyle() # 表格其他样式
	content_font        = Font()
	content_font.height = 20 * 14
	content_font.bold   = False
	content_style.font  = content_font

	# 设置表头
	header_row = table.row(0)
	header_row.height = 20 * 18 # 20为字体大小的基本单位

	# 设置列宽
	name_row      = table.col(TABLE_HEADER.index('书名'))
	ratingNum_row = table.col(TABLE_HEADER.index('评价人数'))
	author_row    = table.col(TABLE_HEADER.index('作者'))
	product_row   = table.col(TABLE_HEADER.index('出版社'))
	date_row      = table.col(TABLE_HEADER.index('出版日期'))
	name_row.width      = 256 * 35 # 其中 256 代表一个字符长度
	ratingNum_row.width = 256 * 15
	author_row.width    = 256 * 35
	product_row.width   = 256 * 50
	date_row.width      = 256 * 15

	style = {
		"header_style": header_style,
		"content_style": content_style
	}
	return style

# 根据需要获取一定数量的书本信息
def getBooks(tag, maxTimes):
	times = 0
	index = 0
	books = [] # 需要保存的书籍信息
	while times < maxTimes:
		bookList = getTableData(tag, index)
		try:
			books = books + bookList
		except BaseException:
			books = books
		times = times + 1
		index = index + 20
		print('\n==========休眠==========\n')
		# 防止封IP策略之一
		time.sleep(random.choice([1, 2, 3])) # 每次爬取之间暂停时间

	print('\n==========%s 书籍获取完毕==========\n' % tag)
	return books

# 执行
tags = getTags()
# print(tags)
for tag in tags:
	bookList = getBooks(tag, TIMES)
	saveToExcel(bookList, tag)












