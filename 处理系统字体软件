import os
import time
import shutil
import re
from PIL import Image, ImageFont, ImageDraw
import win32gui, win32ui, win32con, win32api

#新建文件夹
if not os.path.exists('fonts'):
    os.mkdir('fonts')
    os.mkdir('失效')
    os.mkdir('完成')
input("建议备份需要处理的字体,请吧需要处理的文字放在fonts文件夹内,任意按键继续,剩下的事情交给我吧!")

#结束字体防止BUG
os.system("taskkill /F /IM fontview.exe")
#统计数量&文件夹内的名字
dirs = os.listdir('fonts')
#准备好字体
textfan = u"长"
textzix = u"abcdefghijklmnopqrstuvwxyz\nABCDEFGHIJKLMNOPQSTUVWXYZ\n0123456789\n.:,;(*!?)\n加长静音实心用坏包换\n买一送一304不锈钢缓冲液压\n不锈钢暗藏插销木门暗插\n用坏包换质保十年送安装螺丝"

#开始循环
for file in dirs:
	#带有字体的文件夹会有ini文件,防止BUG
	if "ini" in file:
		#跳出本次循环
		continue
	#打开字体,后面的0是后台打开字体,如果是1就是可视化打开
	win32api.ShellExecute(0, 'open', 'C:/Windows/System32/fontview.exe', 'fonts/'+file,'',0)
	#延迟防止出现BUG
	time.sleep(0.6)
	#获取父窗口句柄
	FindWindowExz=win32gui.FindWindow("FontViewWClass",None)
	#获取父标题,后面的是替换掉空格
	FindWindowExztite=win32gui.GetWindowText(FindWindowExz).replace(' ', '')
	#打印名字
	print("旧名字:"+file)
	print("新名字:"+FindWindowExztite)
	print("-------------------分割线-------------------")
	#结束字体防止BUG
	os.system("taskkill /F /IM fontview.exe")
	#新建按照旧名字命名的文件夹,如果使用新名字命名,会出现BUG容易相同的文字替换掉,
	os.mkdir('完成/'+file)
	#正则提取字体名后缀,一般有ttc,ttf,otf
	files=re.search(r"\..*",file).group()
	#移动字体到新建的文件夹内
	shutil.move("fonts/"+file, "完成/"+file+"/"+FindWindowExztite+files)
	#下面是生成字体的预览图
	#预览否是繁体,后面的数字是大小
	fontfan = ImageFont.truetype(os.path.join("完成/"+file, FindWindowExztite+files), 1100,)
	#预览字体字形,后面的数字是大小
	fontzix = ImageFont.truetype(os.path.join("完成/"+file, FindWindowExztite+files), 127,)
	#新建背景大小
	imfan = Image.new("RGB", (1920, 1080), (255, 255, 255))#背景大小颜色
	imzix = Image.new("RGB", (1920, 1080), (255, 255, 255))#背景大小颜色
	#检查报错,有的文字不支持系统,或者损坏了,无法打开,所以这种文字是不能使用的,将会移动到失效文件夹内
	try:
		#设置繁体的字形格式
		ImageDraw.Draw(imfan).text((10, 5),textfan,font=fontfan,fill="#000000",)#包含文字的位置，文字，字体，颜色
		#设置简体的字形格式
		ImageDraw.Draw(imzix).text((10, 5),textzix,font=fontzix,fill="#000000",)#包含文字的位置，文字，字体，颜色
		#生成繁体简体的预览图
		imfan.save("完成/"+file+"/"+"检查字体繁简.png")
		imzix.save("完成/"+file+"/"+"检查字体字形.png")
	except Exception:
		#吧报错的字体移动到失效的文件夹内
		shutil.move("完成/"+file,"失效")
input("已经处理完成了,分类需要人工分类很快的!,任意键结束")
