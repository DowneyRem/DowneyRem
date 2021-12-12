#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import time
import shutil
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
    
    
def findfile(path):
    for dir in os.listdir(path):
        dir = os.path.join(path, dir)
        if os.path.isdir(dir):
            findfile(dir)
        if os.path.isfile(dir):
            (name, ext) = os.path.splitext(dir)
            if ext == ".TXT" or ext == ".txt":
                list.append(dir)
    return list
    
    
def makedirs(path):
    if not os.path.exists(path):
        os.makedirs(path)
        
        
def savetext(path, text):
    with open(path, "w", encoding = "UTF8") as f:
        f.write(text)
        f.close()
        
            
def copyfile(path1, path2):
    if os.path.exists(path1) and not os.path.exists(path2):
        makedirs(path2)
        shutil.copyfile(path1, path2)
        
        
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
        
        
def convert(list):
    for i in range(0, len(list)):
        readfile = list[i]
        (filepath, name) = os.path.split(readfile) #分离文件名和目录名
        filepath = filepath.replace(path, "")      #留下文件夹名，#新建目录

        name1 = cc1.convert(name)  #转简体
        name2 = cc2.convert(name)  #轉繁體
        filepath1 = cc2.convert(filepath)  #转简体
        filepath2 = cc2.convert(filepath)  #转繁体
        path11 = os.path.join(path1 + filepath1, name1)  #简体文件目录
        path22 = os.path.join(path2 + filepath2, name2)  #繁體文件目錄
        
        
        if os.path.exists(path11) and os.path.exists(path22): 
            i += 1
        else:
            makedirs(path1 + filepath1)
            makedirs(path2 + filepath2)
            text = opentext(readfile)
            
            
            if "會" in text or "後" in text or "來" in text or "東" in text or "電"in text or "個" in text: #原文件是繁体
                savetext(path22, text)     #繁体复制一份，存繁体目录
                text = cc1.convert(text)   #繁体转简体，存简体目录
                savetext(path11, text)
                print("【繁体文档】：【" + name1.replace(".txt","") + "】转换完成，当前进度："+ str(round(100*(i+1)/len(list),2))+"%")
                
                
            elif "会" in text or "来" in text or "东" in text or "电" in text or "个" in text: #原文件是简体
                savetext(path11, text)     #简体复制一份，存简体目录
                text = cc2.convert(text)   #簡體轉繁體，存繁體目錄
                savetext(path22, text)
                print("【簡體文檔】：【" + name2.replace(".txt","") + "】轉換完成，當前進度："+ str(round(100*(i+1)/len(list),2))+"%")
                
                
def main(): 
    print("繁简转换开始：")
    print("繁簡轉換開始：")
    print("下列文件已完成转换：")
    print("下列文件已完成轉換：")
    print("————————————————")
    
    makedirs(path1)
    makedirs(path2)
    findfile(path)

    if len(list) == 0:
        print("————————————————")
        print("没有可转换的文件")
        print("沒有可轉換的文件")
        os.system("pause")
        
    else:
        convert(list)
        print("————————————————")
        print("所有文件均完成转换")
        print("所有文件均完成轉換")
        # os.system("pause")
    
    
if __name__ == '__main__':
    path = os.path.join(os.getcwd())
    path = path.replace("\小说推荐", "\兽人小说")
    path1 = os.path.join(path + "\小说转换版\简体版")
    path2 = os.path.join(path + "\小说转换版\繁體版")
    path  = os.path.join(path + "\小说推荐")
    
    list = []; main()

