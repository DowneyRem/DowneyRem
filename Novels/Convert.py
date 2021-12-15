#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import time
import shutil
from docx import Document
from docx import RT
from opencc import OpenCC
from functools import wraps
cc1 = OpenCC('tw2sp')  #繁转简
cc2 = OpenCC('s2twp')  #簡轉繁
#把简体的 TXT DOCX 文件转换成繁体TXT


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
            if ext == ".txt" or ext == ".docx":
                if ("工具"not in name)and("收集整理"not in name)and("繁體版"not in name)and("Rubbish" not in name):
                    list.append(str(dir))  #原谅我吧，我实在是不会写了……
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
        
        
def savetext(path, text):
    (dir, name) = os.path.split(path) #分离文件名和目录名
    if not os.path.exists(dir):
        os.makedirs(dir)
        
    with open(path, "w", encoding = "UTF8") as f:
        f.write(text)
        f.close()
         
        
def convert(list):
    for i in range(0 ,len(list)):
        readfile = list[i]
        (name, ext) = os.path.splitext(readfile)
        
        #获取转换后的文件目录名
        (filepath, name) = os.path.split(readfile) #分离文件名和目录名
        filepath = filepath.replace(path, "")      #留下文件夹名，#新建目录
        
        name2 = cc2.convert(name) 
        filepath2 = cc2.convert(filepath) 
        path2 = os.path.join(path + filepath2, name2)  #繁体文件目录
        path2 = path2.replace("\唐门小说", "\唐门小说\繁體版")
        path2 = path2.replace(".docx", ".txt")
        #print(path2)


        text = ""
        if os.path.exists(path2):
            i += 1
        elif ext == ".txt":
            text = opentext(readfile)
        elif ext == ".docx":
            text = opendocx(readfile)
            
            
        try:
            text = cc2.convert(text)
            savetext(path2, text)
            print("【" + name + "】转换成功，当前进度："+ str(round(100*(i+1)/len(list),2))+"%")
            
        except:
            print("【" + name + "】打开失败或文件有问题")
    
    
def main():
    print("繁简转换开始：")
    print("下列文件已完成转换：")
    print("————————————————")
    
    path3 = path.replace("\唐门小说", "\唐门小说\繁體版")
    if os.path.exists(path3):
        shutil.rmtree(path3)
        print("旧版文件已经删除")
        
        
    findfile(path)
    if len(list) == 0:
        print("————————————————")
        print("没有可转换的文件")
        os.system("pause")
    else:
        convert(list)
        #os.system('start explorer '+ path)
        print("————————————————")
        print("所有文件均完成转换")
        # os.system("pause")
        
        
if __name__ == "__main__":
    path = os.path.join(os.getcwd())
    path = path.replace("\工具", "")
    
    list = []; main()