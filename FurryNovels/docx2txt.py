#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import time
from docx import Document
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
            if ext == ".docx" or ext == ".DOCX":
                list.append(dir)
    return list
    
    
def savetext(path, text):
    (dir, name) = os.path.split(path) #分离文件名和目录名
    if not os.path.exists(dir):
        os.makedirs(dir)
        
    with open(path, "w", encoding = "UTF8") as f:
        f.write(text)
        f.close()
        
            
def opendocx(list):
    for i in range(0 ,len(list)):
        path = list[i]
        textpath = path.replace("\写作", "\写作\兽人小说")
        textpath = textpath.replace(".docx", ".txt")
        
        if os.path.exists(textpath):
            i += 1
            
        else:
            (filepath, name) = os.path.split(path) #分离文件名和目录名
            name = name.replace(".docx", "")
            try:
                docx = Document(path)
                text = ""; j = 1
                for para in docx.paragraphs:
                    j += 1;
                    if para.style.name == "Title":       #标题：小说名称
                        text += para.text + "\n"
                        
                    elif para.style.name == "Heading 1":  #一级标题：卷
                        text += para.text + "\n"
                        
                    elif para.style.name == "Heading 2":  #二级标题：章
                        text += para.text + "\n"
                        
                    elif para.style.name == "Heading 3":  #三级标题：待定
                        text += para.text + "\n"
                        
                    elif para.style.name == "Normal":        #正文
                        text += para.text + "\n"
                        
                    elif para.style.name == "Normal Indent": #正文缩进
                        text += "　　"+ para.text + "\n"
                        
                savetext(textpath, text)
                print("【" + name + "】转换成功，当前进度："+ str(round(100*(i+1)/len(list),2))+"%")
                
            except:
                print("【" + name + "】打开失败或文件有问题")
                
        
def main():
    print("docx 转 txt 开始：")
    print("下列文件已完成转换：")
    print("————————————————")
    
    path = os.path.join(os.getcwd())
    findfile(path)
    opendocx(list)
    
    print("————————————————")
    # print("所有文件均完成转换")
    # os.system("pause")
    
    
if __name__ == "__main__":
    list = []; text = ""; main()

