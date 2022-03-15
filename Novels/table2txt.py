#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import time
import pandas as pd
from docx.api import Document
from functools import wraps


def timethis(func):
	@wraps(func)
	def wrapper(*args, **kwargs):
		start = time.perf_counter()
		r = func(*args, **kwargs)
		end = time.perf_counter()
		print("{}.{} : {}".format(func.__module__, func.__name__, end - start))
		return r
	return wrapper


def savemd(path, list):
	# 写入md表头与表格间的分割行
	for i in range(0, len(list)):
		a = list[i][0]
		list[i].insert(1, ["--"] * len(a))
	
	text = str(list)
	text = text.replace("'], ['", "|\n|")
	text = text.replace("', '", " | ")
	text = text.replace("']], [['", "|\n\n|")
	text = text.replace("[[['", "|")
	text = text.replace("']]]", "|")
	path = path.replace(".docx", ".md")
	
	try:
		with open(path, "w", encoding="UTF8") as f:
			f.write(text)
		path = os.path.split(path)[1]
		print("已另存为："+ path )
	except IOError:
		print("保存失败")
	
	
def savecsv(path, text):
	text = text.replace("'], " , "\n")  #行末换行
	text = text.replace("', '" , ",")   #间隔改逗号
	text = text.replace("]" , "")
	text = text.replace("[" , "")
	text = text.replace("'" , "")
	path = path.replace(".docx", ".csv")
	
	try:
		with open(path, "w", encoding="UTF-8-sig") as f:
			f.write(text)
		path = os.path.split(path)[1]
		print("已另存为："+ path )
	except IOError:
		print("保存失败")


def savetxt(path, text):
	text = text.replace("'], ", "\n")  #行末换行
	text = text.replace("', '", "\t")  #间隔改制表符
	text = text.replace("]", "")
	text = text.replace("[", "")
	text = text.replace("'", "")
	path = path.replace(".docx", ".txt")
	
	try:
		with open(path, "w", encoding = "UTF8") as f:
			f.write(text)
		path = os.path.split(path)[1]
		print("已另存为："+ path )
	except IOError:
		print("保存失败")


def tabletolist(table, row_num,col_num):
	list = [[] for i in range(row_num)]
	for i in range(0, row_num):
		for j in range(0, col_num):
			cell = table.cell(i, j).text
			list[i].append(cell)
		print(list[i])
	print("")
	return list


def docxtable(path):
	docx = Document(path)
	tables = docx.tables  # 获取文件中的表格集
	table_num = len(tables)
	print("共有 " +str(table_num)+ " 张表格")
	
	text = "" ;biglist = [] #获取文件中的表格
	for i in range(0, table_num):
		table = tables[i]  #获取文件中的第一个表格
		row_num = len(table.rows)
		col_num = len(table.columns)
		print("表"+ str(i) + "："+ str(row_num)+"行" + str(col_num)+"列")
		
		#创建一个二维list，存放表格数据
		list = tabletolist(table, row_num, col_num)
		
		text += str(list) + "\n"*2 		#存储预处理：表格间间隔
		biglist.append(list)            #md预处理
	return(text, biglist)
	

def main(path):
	(text, list) = docxtable(path)
	savetxt(path, text)
	savecsv(path, text)
	savemd(path, list)


if __name__ == "__main__":
	path = "D:测试.docx" #这么写居然是相对路径
	main(path)
