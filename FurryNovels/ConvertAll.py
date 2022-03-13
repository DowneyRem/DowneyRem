#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import time
import shutil
from docx.api import Document
from opencc import OpenCC
from functools import wraps
cc1 = OpenCC('tw2sp')  #繁转简
cc2 = OpenCC('s2twp')  #簡轉繁


def timethis(func):
	@wraps(func)
	def wrapper(*args, **kwargs):
		start = time.perf_counter()
		r = func(*args, **kwargs)
		end = time.perf_counter()
		print('{}.{} : {}'.format(func.__module__, func.__name__, end - start))
		return r
	return wrapper

def monthnow():
	year = str(time.localtime()[0])
	month = str(time.localtime()[1])
	if len(month) == 1:
		month = "0" + month
	string = os.path.join(year, month)
	return string

def opennowdir():
	path = os.getcwd()
	path = path.replace("\小说推荐\工具", "\兽人小说\小说推荐\频道版")
	text = monthnow()
	path = os.path.join(path,text)
	print("已打开：" + path)
	os.system('start explorer '+ path)


def finddocx(path):
	for dir in os.listdir(path):
		dir = os.path.join(path, dir)
		if os.path.isdir(dir):
			finddocx(dir)
		if os.path.isfile(dir):
			(name, ext) = os.path.splitext(dir)
			if ext == ".docx":
				list.append(dir)
	return list

def opendocx(path):
	(dir, name) = os.path.split(path)  # 分离文件名和目录名
	(name, ext) = os.path.splitext(name)
	
	try:
		docx = Document(path)
		text = ""
		for para in docx.paragraphs:
			if para.style.name == "Normal Indent":  # 正文缩进
				text += "　　" + para.text + "\n"
			else:
				text += para.text + "\n"  # 除正文缩进外的其他所有
	except:
		print("【" + name + "】保存失败")
	return text

def opentext(path):
	text = ""
	try:
		with open(path,"r", encoding = "UTF8") as f:
			text = f.read()
	except UnicodeError:
		try:
			with open(path,"r", encoding = "GBK") as f:
				text = f.read()
		except UnicodeError: #Big5 似乎有奇怪的bug，不过目前似乎遇不到
			with open(path,"r", encoding = "BIG5") as f:
				text = f.read()
	finally:
		return text
	
	
def makedirs(path):
	if not os.path.exists(path):
		os.makedirs(path)

def savetext(path, text):
	(dir, name) = os.path.split(path)  # 分离文件名和目录名
	name = name.replace(".txt", "")
	
	if not os.path.exists(dir):
		os.makedirs(dir)
	try:
		with open(path, "w", encoding="UTF8") as f:
			f.write(text)
	except IOError:
		print("【" + name + "】保存失败")

def removefile(path):
	if os.path.exists(path):
		shutil.rmtree(path)
	os.makedirs(path)
	
	
def getpath(path):
	path0 = path.replace("\小说推荐", "\兽人小说\小说推荐\频道版")
	path0 = path0.replace(".docx", ".txt")
	
	dirandfile = path0.replace(sharepath + "\频道版", "")
	path1 = os.path.join(sharepath + "\简体版" + cc1.convert(dirandfile))  # 简体目录
	path2 = os.path.join(sharepath + "\繁體版" + cc2.convert(dirandfile))  # 繁体目录
	
	(filepath, name) = os.path.split(path)  # 分离文件名和目录名
	name = name.replace(".docx", "")
	return path0, path1, path2, name


def convert(list):
	for i in range(0, len(list)):
		path = list[i]
		(path0, path1, path2, name) = getpath(path)
		
		if os.path.exists(path0):
			i += 1
			print("【" + name + "】在本次运行前已转换")
		else:
			text = opendocx(path)
			savetext(path0, text)  # 不区分繁简，存频道目录
			if "會"in text or "後"in text or "來"in text or "東"in text or "電"in text or "個"in text:
				text1 = cc1.convert(text)  # 转简体
				savetext(path2, text)  # 繁体文件复制后，存繁体目录
				savetext(path1, text1)  # 繁体文件转简体，存简体目录
	
			elif "会"in text or "来"in text or "东"in text or "电"in text or "个"in text:
				text2 = cc2.convert(text)  # 转繁体
				savetext(path1, text)  # 简体文件复制后，存简体目录
				savetext(path2, text2)  # 简体文件转繁体，存繁体目录
			print("【"+ name +"】转换成功，当前进度：" + str(round(100*(i+1)/len(list),2))+"%")


def main():
	print("是否要【删除旧文件】重新转换？")
	string = input("输入【 1 】即【删除旧文件】"+"\n"*2)
	if str(1) in string:
		removefile(sharepath)
		print("【删除】旧文件")
	else:
		print("【保留】旧文件")
		
	print("文档转换开始：")
	print("-"*40)
	finddocx(path)
	convert(list)
	print("-" * 40)
	print("文档转换已完成")
	opennowdir()
	print("本月文档标签如下：")
	print("")
	os.system("python ./GetTags.py")
	os.system("pause")
	
	
if __name__ == "__main__":
	path = os.path.join(os.getcwd())
	path = path.replace("\工具","")
	sharepath = path.replace("\小说推荐", "\兽人小说\小说推荐")
	list = []
	main()
