#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import time
from docx.api import Document
from opencc import OpenCC
from functools import wraps
cc = OpenCC('s2twp')  # 簡轉繁

dict = {"R-18":"R18", "R-18G":"R18G", "SFW":"SFW",
        "BL":"Gay", "腐向け":"Gay", "腐向":"Gay",
        "GL":"Lesbian", "百合":"Lesbian", "yuri":"Lesbian",
        "BG":"General",
        "LOVE":"Romance",
        "育肥":"WeightGain", "肥满化":"WeightGain",
        "战损":"BattleDamage",
        "虐屌":"CBT",
        "虐腹":"GutTorture",
        "身体改造":"BodyModification",
        "变身":"Transformation",
        "兽化":"Transfur", "同化":"Transfur", "龙化":"Transfur",
        "兽人":"Furry", "獸人":"Furry", "kemono":"Furry", "":"Furry",
        "纯兽":"non-anthro",

        "性交方式":"",
        "阴道交":"VaginalSex",
        "肛交":"Anal", "后入":"Anal",
        "口交":"BlowJob",
        "生殖腔":"Cloaca", "泄殖腔":"Cloaca",
        "尾交":"TailJob",
        "自慰":"Masturbation",
        "翼交":"WingJob",
        "拳交":"Fisting",
        "足交":"FootJob",
        "脚爪":"Paw", "足控":"Paw",
        "袜交":"SockJob",
        "手淫":"HandJob",
        "脑交":"BrainFuck",
        "耳交":"EarFuck",
        "阴茎打脸":"CockSlapping",
        "阴茎摩擦":"Frottage",
        "尿道插入":"UrethraInsertion",
        "奸杀":"RapedWhileDying",
        "搔痒":"tickling", "挠痒":"tickling",
        "强奸":"Rape",
        "乱伦":"Uncest",
        "3p":"ThreeSome",
        "群交":"Group",
        "绿帽":"Netorare",
        "父子丼":"Oyakodon", "父子":"Oyakodon",
        "兄弟丼":"brother", "兄弟":"brother",

        "调教":"BDSM",
        "捆绑":"Bondage", "束缚":"Bondage",
        "电击":"ElectricShocks",
        "精神控制":"MindControl", "精制":"MindControl",
        "窒息":"Asphyxiation",
        "高潮禁止":"OrgasmDenial", "禁止高潮":"OrgasmDenial",
        "主奴":"Slave", "性奴":"Slave",
        "鼻吊钩":"NoseHook",
        "狗链":"Leash",
        "穿孔":"Piercing",
        "纹身":"Tattoo",
        "淫纹":"CrotchTattoo",
        "写侮辱性词汇":"BodyWriting",
        "肉便器":"PublicUse", "RBQ":"PublicUse", "rbq":"PublicUse",
        "饮尿":"PissDrinking", "圣水":"PissDrinking",
        "食粪":"Coprophagia", "黄金":"Coprophagia",
        "人体家具":"Forniphilia",
        "人体餐盒":"FoodOnBody",

        "附身":"Possession",
        "寄生":"Parasite",
        "石化":"Petrification",
        "堕落":"corruption",
        "催眠":"Hypnosis",
        "气味":"Smell",
        "迷药":"Chloroform",
        "滥用药物":"Drugs",
        "酗酒":"Drunk",
        "寻欢洞":"GloryHole",
        "摄像":"Filming",
        "性转":"GenderBender",
        "灵魂交换":"BodySwap",
        "机械奸":"machine",

        "尿布":"Diaper",
        "乳胶衣":"Latex",
        "生物衣":"LivingClothes",
        "防毒面具":"GasMusk",
        "骑士盔甲":"MetalArmor",
        "动力装甲":"PowerArmor",
        "军装":"Military",
        "紧身衣":"Leotard",
        "拘束衣":"StraitJacket",
        "泳装":"SwimSuit",
        "六尺褌":"Fundoshi",
        "袜子":"socks", "白袜":"socks",

        "肌肉丰满":"Muscles",
        "大根":"BigPenis",
        "巨根":"Hyper",
        "巨大话":"Macro",
        "包茎":"Phimosis",

        "吞食":"Vore",
        "阴茎吞噬":"CockPhagia",
        "肛门吞食":"AnalPhagia",
        "尾巴吞食":"TailPhagia",

        "出产":"Birth", "生产":"Birth",
        "产卵":"Eggs",
        "阴茎出产":"PenisBirth",
        "肛门出产":"AnalBirth",

        "熊":"bear", "熊猫":"panda",
        "马":"hourse", "牛":"bull", "羊":"Sheep",
        "猫":"cat", "狮子":"lion",
        "虎":"Tiger", "老虎":"Tiger",
        "西方龙":"dragon", "东方龙":"long",
        "狗":"dog", "狼":"Wolf", "狐狸":"fox",
        "鱼":"fish", "鲨鱼":"Shark", "海豚":"dolphin",
        "大象":"elephant",
        "袋鼠":"kangaroo",
        "猴子":"monkey",
        "鼠":"mouse", "老鼠":"mouse",
        "豹":"panther",
        "猪":"Pig",
        "兔":"Rabbit", "兔子":"Rabbit",
        "犀牛":"Rhinoceros",
        "蛇":"Snake",
        "龟":"Turtle",
        }


def monthnow():
	year = str(time.localtime()[0])
	month = str(time.localtime()[1])
	if len(month) == 1:
		month = "0" + month
	string = os.path.join(year,month)
	return string


def findfile(path):
	for dir in os.listdir(path):
		dir = os.path.join(path, dir)
		if os.path.isdir(dir):
			findfile(dir)
		if os.path.isfile(dir):
			(name, ext) = os.path.splitext(dir)
			if ext == ".docx":
				pathlist.append(dir)
	return pathlist


def opendocx(path):
	docx = Document(path)
	text = ""
	j = 1
	for para in docx.paragraphs:
		if j < 5:  # 读取前4行内容
			j += 1
			if para.style.name == "Normal Indent":  # 正文缩进
				text += "　　" + para.text + "\n"
			else:
				text += para.text + "\n"  # 除正文缩进外的其他所有
		else:
			break
	
	# 写入txt，以readlines的方式再次读取，获得段落对应的list
	list = Text2List(text)
	return list


def opentext(path):
	try:
		with open(path, "r", encoding="UTF8") as f:
			list = f.readlines()[0:4]
	except UnicodeError:
		with open(path, "r", encoding="GBK") as f:
			list = f.readlines()[0:4]
	finally:
		return list


def SettoText(set):
	text = str(set)
	text = text.replace("{'", "")
	text = text.replace("', '", " ")
	text = text.replace("'}", "")
	return text


def translate(list): #获取英文标签
	tags2 = ""
	s = set()
	for i in range(0, len(list)):
		tag = list[i].replace("\n", "")
		tag = tag.replace("#", "")
		tag = dict.get(tag)
		
		if tag != None:
			s.add("#" + tag)  #利用set去重
		else:
			tag = list[i].replace("\n", "")
			tags2 += tag + " "
	
	tags1 = SettoText(s)
	return tags1, tags2


def Text2List(tags):  # 通过readlines，获得list对象
	path = os.getcwd()
	path = os.path.join(path, "tags.txt")
	with open(path, "w", encoding="UTF8") as f:
		f.write(tags)
	with open(path, "r", encoding="UTF8") as f:
		list = f.readlines()
	try:
		os.remove(path)
	except IOError:
		print("tags.txt 删除失败")
	return list


def TextConvert(list):
	name = cc.convert(list[0])
	authro = "by #" + list[1].replace("作者：", "")
	url = list[2].replace("网址：", "")
	tags = list[3].replace("标签：", "")
	tags = tags.replace(" ", "\n")
	
	list = Text2List(tags)
	(tags1, tags2) = translate(list)
	text = name + authro + tags1 + "\n特殊：" + tags2 + "\n" + url
	print(text)
	return text


def main():
	path = os.getcwd()
	findfile(path)
	j = 0
	
	for i in range(0, len(pathlist)):
		path = pathlist[i]
		(dir, name) = os.path.split(path)
		dirstr = monthnow()  #只处理本月的文件
		
		if dirstr in dir:
			j += 1
			list1 = opendocx(path)
			TextConvert(list1)
	if j == 0:
		print("本月 " + dirstr + " 无新文档")
		

if __name__ == "__main__":
	pathlist = []
	print("")
	main()
