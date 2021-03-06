# -*- coding:UTF-8 -*-
import os
import sys
import time
import os.path
import datetime
import shutil
import win32com.client as win32
import winshell
from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.schedulers.background import BackgroundScheduler
import pythoncom
import logging

localtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))  # 获得当前系统时间的字符串
today = datetime.datetime.now()  # 获取当前日期
y = today.year  # 获取当前年份数字
m = today.month  # 获取当前月份数字
d = today.day  # 获取当前日期数字
h = today.hour  # 获取当前小时数字
iminute = today.minute  # 获取当前分钟数字
r_dir = 'C:/Users/Administrator/Desktop/'
new_path = r_dir + str(y) + '/' + str(m) + '月/'  # 获取文件夹名称
year_dir = r_dir + str(y) + '/'
month_dir = year_dir + str(m) + '月/'


# winshell.desktop()获取桌面路径
# os.path.join(winshell.desktop(), "网络巡检.lnk")组合路径
# 如：C:\Users\Administrator\Desktop\网络巡检.lnk


def make_dir(path):
    # path = path.strip() # 去除首位空格
    # path = path.rstrip("\") # 去除尾部 \ 符号

    isexists = os.path.exists(path)
    if not isexists:
        # 如果不存在则创建目录
        # 创建目录操作函数
        os.makedirs(path)
        print(path + '创建成功')
        return True
    else:
        # 如果目录存在则不创建，并提示目录已存在
        print(path + '目录已存在')
        return False


def re_lnk():
    link_filepath = os.path.join(winshell.desktop(), "网络巡检.lnk")
    print(link_filepath)
    with winshell.shortcut(link_filepath) as link:
        link.path = year_dir + str(m) + '月/'  # 目标文件
        # link.working_directory = "os.path.join(winshell.desktop()/str(3月)" #起始位置


def make_firstflie():
    old_dir = r_dir + str(y) + str(m - 1) + r'月/'
    os.chdir(old_dir)
    mylist = os.listdir(old_dir)
    mylist.sort(reverse=True)  # 按文件名倒序
    first_name = mylist[0]  # 倒序第一个最新文件

    new_dir = r_dir + str(y) + '/' + str(m) + r'月/'
    os.chdir(new_dir)
    second_name = '机房服务器巡检表' + localtime[:10] + ' ' + localtime[11:13] + "：00" + '.xlsx'  # 中文冒号！

    shutil.copy(old_dir + first_name, new_dir + second_name)  # 复制最新文件


def make_file():
    old_dir = r_dir + str(y) + '/' + str(m) + r'月/'
    os.chdir(old_dir)
    mylist = os.listdir(old_dir)
    mylist.sort(reverse=True)  # 按文件名倒序
    first_name = mylist[0]  # 倒序第一个最新文件
    second_name = '机房服务器巡检表' + localtime[:10] + ' ' + localtime[11:13] + "：00" + '.xlsx'  # 中文冒号！
    try:
        shutil.copy(old_dir + first_name, old_dir + second_name)  # 复制最新文件
    except shutil.SameFileError as e:
        logging.exception(e)
    # finally:
    #     pass


def write_newfile():  # 填写新文件内日期时间并保存退出
    pythoncom.CoInitialize()
    excel = win32.gencache.EnsureDispatch('Excel.Application')  # opens Excel
    now_path = r_dir + str(y) + '/' + str(m) + '月/'
    os.chdir(now_path)
    mylist = os.listdir(now_path)
    mylist.sort(reverse=True)  # 按文件名倒序
    be_write_file_name = mylist[0]  # 倒序第一个最新文件
    wb = excel.Workbooks.Open(now_path + be_write_file_name)  # opens "Test" file
    wb.Sheets(1).Select()  # select 2nd worksheet "Aisle_2"
    excel.Visible = True
    excel.Range("B2").Select()
    excel.ActiveCell.Value = "时间：" + localtime[:10]  # Fill in test data #
    excel.Range("O2").Select()
    excel.ActiveCell.Value = localtime[11:13] + "：00"  # Fill in test data #
    wb.Save()
    wb.Close()
    excel.Quit()
    pythoncom.CoUninitialize()


def job():
    if localtime[8:10] == '01' and localtime[11:13] == '01':
        make_dir(New_Path)
        make_firstflie()
        write_newfile()
        re_lnk()

    else:
        make_file()
        write_newfile()


if __name__ == '__main__':

    scheduler = BlockingScheduler()
    scheduler.add_job(job, 'cron', hour='2, 6, 8, 14, 18, 20, 22', minute=30)  # day_of_week='0-6',
    # scheduler.add_job(job, 'cron', day='1' ,hour='1')  # day_of_week='0-6'
    try:
        scheduler.start()

    except (KeyboardInterrupt, SystemExit):
        job.remove()
