#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import re
import time
from docx import Document
from functools import wraps


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
            if ext == ".docx" or ext == ".DOCX":
                list.append(dir)
    return list
    
    
def savetext(path, text):
    path = path.replace(".docx",".txt")
    with open(path, "w", encoding = "UTF8") as f:
        f.write(text)
        f.close()
        
            
def opendocx(list):
    for i in range(0 ,len(list)):
        path = list[i]
        docx = Document(path)
        text = ""; j = 1
        
        for para in docx.paragraphs:
            j += 1;
            if j <= 5 or (("第"in para.text or "序"in para.text or "终"in para.text)and "章"in para.text):
                text += para.text + "\n"
            else:
                text += "　　"+ para.text + "\n"
        
        print(text)
        savetext(path, text)
        (filepath, name) = os.path.split(path) #分离文件名和目录名
        name = name.replace(".docx", "")
        print("【" + name + "】转换成功，当前进度："+ str(round(100*(i+1)/len(list),2))+"%")
        

            
@timethis
def main():
    print("转换开始：")
    print("下列文件已完成转换：")
    print("————————————————")

    path = os.path.join(os.getcwd())
    findfile(path)
    opendocx(list)
        
    print("————————————————")
    print("所有文件均完成转换")
    # os.system("pause")
    
    
if __name__ == '__main__':
    list = []; text = ""; main()



    
    
    
