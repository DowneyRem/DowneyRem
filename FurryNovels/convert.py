#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import time
from opencc import OpenCC
from functools import wraps
cc1 = OpenCC('tw2s') #繁转简
cc2 = OpenCC('s2tw') #簡轉繁


def timethis(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        start = time.perf_counter()
        r = func(*args, **kwargs)
        end = time.perf_counter()
        print('{}.{} : {}'.format(func.__module__, func.__name__, end - start))
        return r
    return wrapper


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
        
        
@timethis
def convert(list):
    for i in range(0, len(list)):
        readfile =list[i]
        (filepath, name) = os.path.split(readfile) #分离文件名和目录名
        filepath = filepath.replace(path, "")      #留下文件夹名，#新建目录
        
        name1 = cc1.convert(name)  #转简体
        name2 = cc2.convert(name)  #转繁体
        path11 = os.path.join(path1 + filepath, name1)  #简体目录
        path22 = os.path.join(path2 + filepath, name2)  #繁体目录
        
        if os.path.exists(path11) and os.path.exists(path22): 
            i += 1
        else:
            makedirs(path1 + filepath)
            makedirs(path2 + filepath)
            
            with open(readfile,"r", encoding = "UTF-8") as f:
                text = f.read()

                # "會"or"後"or"來"or"東"or"電"or"個" in text: #文件是繁体
                text = cc1.convert(text) #繁体转简体，存简体目录
                savefile(os.path.join(path11), text)
                
                
                # "会"or"来"or"东"or"电"or"个" in text: #文件是简体
                text = cc2.convert(text) #簡體轉繁體，存繁體目錄
                savefile(os.path.join(path22), text)
                print("【"+name + "】转换成功")
                
                f.close
        print("【"+ cc1.convert(name) + "】已完成转换，当前进度："+ str(100*(i+1)/len(list))+"%")
        
        
@timethis
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
    path = os.path.join(os.getcwd() + "\小说推荐")
    path1 = os.path.join(os.getcwd() + "\小说转换版\简体版")
    path2 = os.path.join(os.getcwd() + "\小说转换版\繁體版")
    list = []; main()
