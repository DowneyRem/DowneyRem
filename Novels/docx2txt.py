#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import time
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


def findfile(path):
	for dir in os.listdir(path):
		dir = os.path.join(path, dir)
		if os.path.isdir(dir):
			findfile(dir)
		if os.path.isfile(dir):
			(name, ext) = os.path.splitext(dir)
			if ext == ".docx":
				list.append(dir)
	return list


def opendocx(path):
	docx = Document(path)
	text = ""; j = 1
	for para in docx.paragraphs:
		j += 1
		if para.style.name == "Normal Indent":  #正文缩进
			text += "　　"+ para.text + "\n"
		else:
			text += para.text + "\n"            #除正文缩进外的其他所有
	return text


def savetext(path, text):
	path = path.replace(".docx", ".txt")
	with open(path, "w", encoding = "UTF8") as f:
		f.write(text)
		f.close()


def convert(list):
	for i in range(0 ,len(list)):
		path = list[i]
		(filepath, name) = os.path.split(path)  # 分离文件名和目录名
		name = name.replace(".docx", "")
		
		try:
			text = opendocx(path)
			savetext(path, text)
			print("【" + name + "】转换成功，当前进度："+ str(round(100*(i+1)/len(list),2))+"%")
			
		except:
			print("【" + name + "】打开失败或文件有问题")


def main():
	findfile(path)
	if len(list) == 0:
		print("将当前目录下docx转换成txt")
		print("当前目录下，没有docx文件")
		print("请把本文件放在相应目录下")
		print("\n" * 2)
		print("仅支持使用了【正文缩进】样式的文档")
		
		os.system("pause")
		
	else:
		print("转换开始：")
		print("下列文件已完成转换：")
		print("—" * 30)
		convert(list)
		print("—" * 30)
		print("所有文件均完成转换")
		# os.system("pause")


if __name__ == "__main__":
	path = os.path.join(os.getcwd())
	list = []; main()
