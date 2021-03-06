import os,sys
import time
import os.path
import datetime
import shutil
import win32com.client as win32
import winshell

localtime=time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
Today = datetime.datetime.now()   #获取当前日期
y = Today.year  #获取当前年份数字
m = Today.month #获取当前月份数字
d = Today.day   #获取当前日期数字
h = Today.hour  #获取当前小时数字

print (y,m)

Year_Dir = 'C:/Users/Administrator/Desktop/'+ str(y)+ '/'

def Re_Lnk():
    link_filepath = os.path.join(winshell.desktop(), "网络巡检.lnk")
    print(link_filepath)
    with winshell.shortcut(link_filepath) as link:
        link.path = Year_Dir + str(m)+ '/'  #目标文件
        #link.working_directory = "os.path.join(winshell.desktop()/str(3月)" #起始位置

Re_Lnk()