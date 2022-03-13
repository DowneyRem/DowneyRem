#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import time
from docx.api import Document
from opencc import OpenCC
from functools import wraps
cc1 = OpenCC('tw2sp')  #繁转简
cc2 = OpenCC('s2twp')  #簡轉繁


DICT = dict = {"txt":"txt",
        "docx":"docx",
		"完结":"Finished",
		"完稿":"Finished",
		"finished":"Finished",
		"更新":"Updating",
		"更新中":"Updating",
		"updating":"Updating",
		"停笔":"Died",
		"太监":"Died",
		"died":"Died",
		"简体中文":"zh_cn",
		"简中":"zh_cn",
		"zh_cn":"zh_cn",
		"繁体中文":"zh_tw",
		"正体中文":"zh_tw",
		"繁中":"zh_tw",
		"zh_tw":"zh_tw",
		"中文":"Chinese",
		"汉语":"Chinese",
		"中国语":"Chinese",
		"中国语注意":"Chinese",
		"chinese":"Chinese",
		"原创":"Original",
		"original":"Original",
		"同人":"Doujin",
		"doujin":"Doujin",
		"翻译":"Translated",
		"translated":"Translated",
  
		"sfw":"SFW",
		"R-18":"R18",
		"R18":"R18",
		"R-18G":"R18G",
		"R18G":"R18G",
		"BL":"Gay",
		"腐向け":"Gay",
		"腐向":"Gay",
		"同志":"Gay",
		"男同性恋":"Gay",
		"gay":"Gay",
		"GL":"Lesbian",
		"百合":"Lesbian",
		"yuri":"Lesbian",
		"女同性恋":"Lesbian",
		"Lesbian":"Lesbian",
		"BG":"General",
		"男女":"General",
		"general":"General",
		"恋爱":"Romance",
		"LOVE":"Romance",
		"romance":"Romance",
		"科幻":"ScienceFiction",
		"ScienceFiction":"ScienceFiction",
		"奇幻":"Fantasy",
		"魔幻":"Fantasy",
		"Fantasy":"Fantasy",
  
		"熊":"bear",
		"bear":"bear",
		"熊猫":"panda",
		"panda":"panda",
		"马":"hourse",
		"hourse":"hourse",
		"牛":"bull",
		"bull":"bull",
		"羊":"sheep",
		"sheep":"sheep",
		"猫":"cat",
		"cat":"cat",
		"狮":"lion",
		"狮子":"lion",
		"lion":"lion",
		"虎":"tiger",
		"老虎":"tiger",
		"tiger":"tiger",
		"西方龙":"dragon",
		"龙":"dragon",
		"龙人":"dragon",
		"dragon":"dragon",
		"东方龙":"long",
		"long":"long",
		"狗":"dog",
		"dog":"dog",
		"狼":"wolf",
		"wolf":"wolf",
		"狐狸":"fox",
		"fox":"fox",
		"鱼":"fish",
		"fish":"fish",
		"鲨鱼":"shark",
		"shark":"shark",
		"海豚":"dolphin",
		"dolphin":"dolphin",
		"大象":"elephant",
		"elephant":"elephant",
		"袋鼠":"kangaroo",
		"kangaroo":"kangaroo",
		"猴子":"monkey",
		"monkey":"monkey",
		"鼠":"mouse",
		"老鼠":"mouse",
		"mouse":"mouse",
		"豹":"panther",
		"panther":"panther",
		"猪":"pig",
		"pig":"pig",
		"兔":"rabbit",
		"兔子":"rabbit",
		"rabbit":"rabbit",
		"犀牛":"rhinoceros",
		"rhinoceros":"rhinoceros",
		"蛇":"snake",
		"snake":"snake",
		"龟":"turtle",
		"turtle":"turtle",
		"鲨狗":"sergal",
		"sergal":"sergal",
  
		"兽人":"Furry",
		"獸人":"Furry",
		"kemono":"Furry",
		"獣人":"Furry",
		"furry":"Furry",
		"纯兽":"non-anthro",
		"non-anthro":"non-anthro",
		"人类":"Human",
		"human":"Human",
		"鸟人":"Harpy",
		"harpy":"Harpy",
		"机器人":"Robot",
		"robot":"Robot",
		"史莱姆":"Slime",
		"slime":"Slime",
		"触手":"Tentacles",
		"tentacles":"Tentacles",
		"吸血鬼":"Vampire",
		"vampire":"Vampire",
		"人皮服装":"SkinSuit",
		"skinsuit":"SkinSuit",
		"怪物":"Monser",
		"monser":"Monser",
		"英雄":"Hero",
		"hero":"Hero",
		"警察":"Police",
		"特警":"Police",
		"police":"Police",
		"魔王":"DevilKing",
		"devilking":"DevilKing",
		"勇者":"Brave",
		"brave":"Brave",
		"魅魔":"Succubus",
		"succubus":"Succubus",
		"医生":"Doctor",
		"doctor":"Doctor",
  
		"阴道交":"VaginalSex",
		"阴道":"VaginalSex",
		"vaginalsex":"VaginalSex",
		"肛交":"Anal",
		"后入":"Anal",
		"anal":"Anal",
		"口交":"BlowJob",
		"blowjob":"BlowJob",
		"生殖腔":"Cloaca",
		"泄殖腔":"Cloaca",
		"cloaca":"Cloaca",
		"尾交":"TailJob",
		"tailjob":"TailJob",
		"自慰":"Masturbation",
		"masturbation":"Masturbation",
		"翼交":"WingJob",
		"wingjob":"WingJob",
		"拳交":"Fisting",
		"fisting":"Fisting",
		"足交":"FootJob",
		"footjob":"FootJob",
		"脚爪":"Paw",
		"足控":"Paw",
		"舔足":"Paw",
		"paw":"Paw",
		"袜交":"SockJob",
		"sockjob":"SockJob",
		"手淫":"HandJob",
		"handjob":"HandJob",
		"脑交":"BrainFuck",
		"brainfuck":"BrainFuck",
		"耳交":"EarFuck",
		"earfuck":"EarFuck",
		"阴茎打脸":"CockSlapping",
		"cockslapping":"CockSlapping",
		"阴茎摩擦":"Frottage",
		"frottage":"Frottage",
		"尿道插入":"UrethraInsertion",
		"urethrainsertion":"UrethraInsertion",
		"奸杀":"RapedWhileDying",
		"rapedwhiledying":"RapedWhileDying",
		"搔痒":"tickling",
		"挠痒":"tickling",
		"tickling":"tickling",
  
		"强奸":"Rape",
		"rape":"Rape",
		"乱伦":"Incest",
		"incest":"Incest",
		"群交":"Group",
		"群p":"Group",
		"轮奸":"Group",
		"group":"Group",
		"绿帽":"Cuckold",
		"netorare":"Cuckold",
		"ntr":"Cuckold",
		"cuckold":"Cuckold",
		"父子丼":"Oyakodon",
		"父子":"Oyakodon",
		"oyakodon":"Oyakodon",
		"兄弟丼":"brother",
		"兄弟":"brother",
		"brother":"brother",
  
		"调教":"BDSM",
		"bdsm":"BDSM",
		"捆绑":"Bondage",
		"束缚":"Bondage",
		"bondage":"Bondage",
		"电击":"ElectricShocks",
		"electricshocks":"ElectricShocks",
		"精神控制":"MindControl",
		"精控":"MindControl",
		"洗脑":"MindControl",
		"mindcontrol":"MindControl",
		"窒息":"Asphyxiation",
		"asphyxiation":"Asphyxiation",
		"高潮禁止":"OrgasmDenial",
		"禁止高潮":"OrgasmDenial",
		"射精控制":"OrgasmDenial",
		"orgasmdenial":"OrgasmDenial",
		"主奴":"Slave",
		"性奴":"Slave",
		"主仆":"Slave",
		"slave":"Slave",
		"鼻吊钩":"NoseHook",
		"nosehook":"NoseHook",
		"狗链":"Leash",
		"leash":"Leash",
		"穿孔":"Piercing",
		"piercing":"Piercing",
		"纹身":"Tattoo",
		"tattoo":"Tattoo",
		"淫纹":"CrotchTattoo",
		"crotchtattoo":"CrotchTattoo",
		"肉便器":"PublicUse",
		"rbq":"PublicUse",
		"publicuse":"PublicUse",
		"饮尿":"PissDrinking",
		"圣水":"PissDrinking",
		"pissdrinking":"PissDrinking",
		"食粪":"Coprophagia",
		"黄金":"Coprophagia",
		"coprophagia":"Coprophagia",
		"人体家具":"Forniphilia",
		"forniphilia":"Forniphilia",
		"人体餐盒":"FoodOnBody",
		"foodonbody":"FoodOnBody",
		"育肥":"WeightGain",
		"肥满化":"WeightGain",
		"weightgain":"WeightGain",
		"战损":"BattleDamage",
		"欠損":"BattleDamage",
		"battledamage":"BattleDamage",
		"虐屌":"CBT",
		"cbt":"CBT",
		"虐腹":"GutTorture",
		"guttorture":"GutTorture",
		"身体改造":"BodyModification",
		"肉体改造":"BodyModification",
		"bodymodification":"BodyModification",
		"变身":"Transformation",
		"transformation":"Transformation",
		"兽化":"Transfur",
		"同化":"Transfur",
		"龙化":"Transfur",
		"transfur":"Transfur",
		"TF":"Transfur",
		"附身":"Possession",
		"possession":"Possession",
		"寄生":"Parasite",
		"parasite":"Parasite",
		"石化":"Petrification",
		"petrification":"Petrification",
		"堕落":"corruption",
		"恶堕":"corruption",
		"corruption":"corruption",
		"催眠":"Hypnosis",
		"hypnosis":"Hypnosis",
		"气味":"Smell",
		"雄臭":"Smell",
		"smell":"Smell",
		"迷药":"Chloroform",
		"chloroform":"Chloroform",
		"药物":"Drugs",
		"药物催淫":"Drugs",
		"drugs":"Drugs",
		"酗酒":"Drunk",
		"drunk":"Drunk",
		"寻欢洞":"GloryHole",
		"gloryhole":"GloryHole",
		"摄像":"Filming",
		"filming":"Filming",
		"性转":"GenderBender",
		"genderbender":"GenderBender",
		"灵魂交换":"BodySwap",
		"身体交换":"BodySwap",
		"bodyswap":"BodySwap",
		"机械奸":"Machine",
		"machine":"Machine",
		"drone":"Drone",
		"尿布":"Diaper",
		"diaper":"Diaper",
  
		"乳胶衣":"Latex",
		"胶衣":"Latex",
		"乳胶":"Latex",
		"latex":"Latex",
		"胶液":"Rubber",
		"胶":"Rubber",
		"龙胶":"Rubber",
		"rubber":"Rubber",
		"生物衣":"LivingClothes",
		"livingclothes":"LivingClothes",
		"防毒面具":"GasMusk",
		"gasmusk":"GasMusk",
		"盔甲":"Armor",
		"armor":"Armor",
		"骑士盔甲":"MetalArmor",
		"metalarmor":"MetalArmor",
		"动力装甲":"PowerArmor",
		"powerarmor":"PowerArmor",
		"军装":"Military",
		"military":"Military",
		"紧身衣":"Leotard",
		"leotard":"Leotard",
		"拘束衣":"StraitJacket",
		"straitjacket":"StraitJacket",
		"泳装":"SwimSuit",
		"swimsuit":"SwimSuit",
		"六尺褌":"Fundoshi",
		"fundoshi":"Fundoshi",
		"袜子":"socks",
		"白袜":"socks",
		"socks":"socks",
  
		"肌肉":"Muscles",
		"筋肉":"Muscles",
		"muscles":"Muscles",
		"巨根":"Hyper",
		"hyper":"Hyper",
		"巨大化":"Macro",
		"macro":"Macro",
		"包茎":"Phimosis",
		"phimosis":"Phimosis",
		"吞食":"Vore",
		"丸吞":"Vore",
		"vore":"Vore",
		"阴茎吞噬":"CockPhagia",
		"cockphagia":"CockPhagia",
		"肛门吞食":"AnalPhagia",
		"analphagia":"AnalPhagia",
		"尾巴吞食":"TailPhagia",
		"tailphagia":"TailPhagia",
		"出产":"Birth",
		"生产":"Birth",
		"birth":"Birth",
		"产卵":"Eggs",
		"eggs":"Eggs",
		"阴茎出产":"PenisBirth",
		"penisbirth":"PenisBirth",
		"肛门出产":"AnalBirth",
		"analbirth":"AnalBirth",
  
		"小说":"",
        }


def monthnow():
	year = str(time.localtime()[0])
	month = str(time.localtime()[1])
	if len(month) == 1:
		month = "0" + month
	string = os.path.join(year, month)
	return string


def dict2md(path):
	text = "### 关键词标签表"
	text = text + "\n| 标签 | 关键词 | "
	text = text + "\n| -- | -- | "
	list1 = list(dict.items())
	for i in range(0, len(list1)):
		(key, value) = list1[i]
		if value != "":
			value = "#" + value
			text += "\n| " + value + " | " + key + " |"
	savetext(path, text)


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


def opendocx4(path):
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
	text = cc1.convert(text)  #转简体，只处理简体标签
	# 写入txt，以readlines的方式再次读取，获得段落对应的list
	list = text2list(text)
	return list


def opentext4(path):
	try:
		with open(path, "r", encoding="UTF8") as f:
			list = f.readlines()[0:4]
	except UnicodeError:
		with open(path, "r", encoding="GBK") as f:
			list = f.readlines()[0:4]
	finally:
		return list


def set2text(set):
	text = str(set)
	text = text.replace("{'", "")
	text = text.replace("', '", " ")
	text = text.replace("'}", "")
	return text


def translate(list):  # 获取英文标签
	tags2 = ""
	s = set()
	for i in range(0, len(list)):
		tag = list[i].replace("\n", "")
		tag = tag.replace("#", "")
		tag = tag.replace(" ", "")
		tag = tag.replace("　", "")
		tag = dict.get(tag)
		
		if tag != None:
			s.add("#" + tag)  # 利用set去重
		else:
			tag = list[i].replace("\n", "")
			tag = tag.replace(" ", "")
			tag = tag.replace("　", "")
			tags2 += tag + " "
	tags1 = set2text(s)
	return tags1, tags2


def text2list(tags):  # 通过readlines，获得list对象
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


def textconvert(list):
	name = cc2.convert(list[0])
	authro = "by #" + list[1].replace("作者：", "")
	url = list[2].replace("网址：", "")
	tags = list[3].replace("标签：", "")
	tags = tags.replace(" ", "\n")
	
	list = text2list(tags)   #通过保存txt将tag写入list
	(tags1, tags2) = translate(list)  #已翻译，未翻译的标签
	text = name + authro + tags1 + "\n特殊：" + tags2 + "\n" + url
	# print(tags2)
	print(text)
	return text,tags2


def main():
	path = os.getcwd()
	dict2md(os.path.join(path,"Tags.md"))
	path = path.replace("\工具", "")
	pathlist = findfile(path)
	dirstr = monthnow()  # 只处理本月的文件
	j = 0
	s=set()
	
	for i in range(0, len(pathlist)):
		path = pathlist[i]
		(dir, name) = os.path.split(path)
		if dirstr in dir:
			j += 1
			list1 = opendocx4(path)
			s.add(textconvert(list1)[1])
		
	text = set2text(s)
	path = "D:\\Users\\Administrator\\Desktop\\tag.txt"
	savetext(path, text)
	
	if j == 0:
		print("本月 " + dirstr + " 无新文档")


if __name__ == "__main__":
	pathlist = []
	print("")
	main()
