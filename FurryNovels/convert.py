#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
from opencc import OpenCC


cc1 = OpenCC('tw2s') #繁转简
cc2 = OpenCC('s2tw') #簡轉繁
path = "D:\\兽人小说\\小说推荐"
path1 = "D:\\兽人小说\\小说转换版\\简体版"
path2 = "D:\\兽人小说\\小说转换版\\繁體版"

   
def findtxt(path):
    for dir in os.listdir(path):
        dir = os.path.join(path, dir)
        if os.path.isdir(dir):
            findtxt(dir)
        if os.path.isfile(dir):
            (name, ext) = os.path.splitext(dir)
            if ext == ".TXT" or ext == ".txt":
                list.append(dir)
    return list
    
    
def makedirs(path):
    if not os.path.exists(path):
        os.makedirs(path)
        
        
def savefile(path, text):
    if not os.path.exists(path):
        with open(path, "w", encoding = "UTF-8") as f:
            f.write(text)
            f.close()
        
        
def convert(list):
    for i in range(0,len(list)):
        readfile =list[i]
        
        (filepath, name) = os.path.split(readfile) #分离文件名和目录名
        filepath = filepath.replace(path, "")      #留下文件夹名，新建目录
        makedirs(path1 + filepath)
        makedirs(path2 + filepath)

        with open(readfile,"r", encoding = "UTF-8") as f:
            (tempath, name) = os.path.split(f.name)  #带拓展名的文件名
            text = f.read()
            
            if "會"or"後"or"來"or"東"or"電"or"個" in text: #文件是繁体
                name = cc1.convert(name) #繁体转简体，存简体目录
                text = cc1.convert(text)
                savefile(os.path.join(path1 + filepath, name), text)
                
            if "会"or"来"or"东"or"电"or"个" in text: #文件是简体
                name = cc2.convert(name) #簡體轉繁體，存繁體目錄
                text = cc2.convert(text)
                savefile(os.path.join(path2 + filepath, name), text)
                
            f.close
        print("【"+ cc1.convert(name) + "】已完成转换，当前进度："+ str(100*(i+1)/len(list))+"%")
        
        
def main(): 
    print("繁简转换开始：")
    print("————————————————")
    print("下列文件已完成转换：")
    
    makedirs(path1)
    makedirs(path2)
    findtxt(path)
    convert(list)
    
    print("————————————————")
    print("转换完成")
    os.system("pause")


if __name__ == '__main__':
    list = []
    main()
