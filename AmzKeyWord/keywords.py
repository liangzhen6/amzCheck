#!/usr/bin/python3
# -*- coding: utf-8 -*-
import xlrd, xlwt, os
from xlutils.copy import copy
from download import request
from bs4 import BeautifulSoup
import time
import platform
from selenium import webdriver

system = platform.system()
xlsPath = ''
#根据系统识别路径
if system == 'Darwin':#mac
    originPath = os.path.abspath('.')
    xlsPath = os.path.join(originPath,'loadData.xls')
elif system == 'Windows':
    originPath = 'C:/Users/Administrator/Desktop/amzkeyword'
    xlsPath = os.path.join(originPath,'loadData.xls')


#保存xls表格
def savexls():
    os.remove(xlsPath)
    workbookCopy.save(xlsPath)
#获取当前表格的信息
def get_sheet_mes():
    workbook = xlrd.open_workbook(xlsPath)
    workbookCopy = copy(workbook)

    sheet_name = workbook.sheet_names()[0]
    sheet_one = workbook.sheet_by_name(sheet_name)
    wordlist = sheet_one.row_values(0)
    selllist = sheet_one.col_values(1)
    return workbookCopy, wordlist, selllist

# 获取时间精确到秒s
def get_date():
    return time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))



# 初始化浏览器
browser = webdriver.Chrome()

# 获取商品信息
def getdata(keyword):
    wordurl = keyword
    
    browser.get(wordurl)

    div = browser.find_element_by_id('olp-upd-new')
#    print('666',div)
    if div != None:
        all_a = div.find_element_by_tag_name('a')
        href = all_a.get_attribute('href');
        if href != None:
            print(href)
            return href
        else:
            return None
    else:
        return None


# 获取店铺的信息
def get_store_data(path):
    browser.get(path)
    #('a-section a-spacing-double-large')
    div = browser.find_element_by_xpath(".//*[@class='a-section a-spacing-double-large']")

    all_div = div.find_elements_by_xpath(".//*[@class='a-column a-span2 olpSellerColumn']")
    
    # 所有的店铺路径拼起来
    allPath = ""
    for subdiv in all_div:
        span = subdiv.find_element_by_xpath(".//*[@class='a-size-medium a-text-bold']")
        herfa = span.find_element_by_tag_name('a')
        herf = herfa.get_attribute('href')
        allPath = allPath + '        ' + herf
#        print('90', herf)

    return allPath




def get_store_path(sheetPath):
    path = getdata(sheetPath)
    
    if path != None:
        return get_store_data(path)
    else:
        return None


#使用说明
# 1.修改xls文件的路径
# 2.设置总共需要抓去多少次
# 3.设置每次抓取的时间间隔已秒为单位

curr_count = 0      #当前计次
count_max = 3       #要抓取的总次数
sleep_time = 60*10  #每次抓取的时间间隔


while curr_count < count_max:
    workbookCopy, wordlist, link_list = get_sheet_mes()
    curr_col = len(wordlist) #从哪里列开始写入
    sheet = workbookCopy.get_sheet(0)
    
    dateNow = get_date()
    # 写入时间
    sheet.write(0, curr_col, dateNow)
    
    for x in range(1, len(link_list)):
        link = link_list[x]
        
        store_link = get_store_path(link)
        sheet.write(x, curr_col, store_link)
        #休眠2s
        time.sleep(2)
    # 保存一次数据
    savexls()
    #每次 抓取之间间隔时间
    time.sleep(sleep_time)
    #次数 加1
    curr_count = curr_count + 1
    #次数够了关闭浏览器
    if curr_count == count_max:
        browser.close()






