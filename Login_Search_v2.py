# -*- coding: UTF-8 -*-
import re
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from lxml import etree
from fake_useragent import UserAgent
from time import sleep
from bs4 import BeautifulSoup
import win32gui
import win32api
import win32con
import win32clipboard
from ctypes import *
pyPath, filename = os.path.split(__file__)

ID = "YOUR ID"
PW = "YOUR PW"
EN = "YOUR EN"
streetNum = "0056"
locNum = "1036"

chrome_options = webdriver.ChromeOptions()
# chrome_options.add_argument('--headless') 
# chrome_options.add_argument('blink-settings=imagesEnabled=false')
# chrome_options.add_argument("user-agent={}".format(ua))
driver = webdriver.Chrome(executable_path=pyPath + '/chromedriver.exe',chrome_options=chrome_options)
driver.maximize_window()
driver.get("https://bs-etw.land.nat.gov.tw:8443/")
username = driver.find_element_by_xpath('//*[@name="UserID"]')
password = driver.find_element_by_xpath('//*[@name="Password"]')
EName = driver.find_element_by_xpath('//*[@name="EName"]')
BtnSure = driver.find_element_by_xpath('//*[@name="Image1"]')

username.send_keys(ID)
password.send_keys(PW)
EName.send_keys(EN)

sleep(10)
BtnSure.click()
sleep(3)
BtnEle = driver.find_element_by_xpath('//*[@name="Image9"]')
BtnEle.click()

# driver.get("http://210.71.181.114/epLandData/epLandData.aspx")

city = Select(driver.find_element_by_name('QueryREG1$DropDownList1'))
city.select_by_index(2)
sleep(2)
area = Select(driver.find_element_by_name('QueryREG1$DropDownList2'))
area.select_by_visible_text("大園區")
sleep(2)
street = driver.find_element_by_xpath('//*[@name="QueryREG1_scListBox1SelectedValue"]')
street.send_keys(Keys.CONTROL + 'a')
street.send_keys(streetNum)
sleep(2)
dept = Select(driver.find_element_by_name('QueryREG1$DropDownList4'))
dept.select_by_index(1)
sleep(2)
inputNum = driver.find_element_by_xpath('//*[@name="QueryREG1$TextBox2"]')
inputNum.send_keys(locNum)
sleep(2)
BtnQu = driver.find_element_by_xpath('//*[@name="QueryREG1$Button1"]')
BtnQu.click()

sleep(3)

##另存 html

# 测试网址
# news_url = "http://news.youth.cn/sz/201812/t20181218_11817816.htm"

# # 打开另存为mhtml功能
# options = webdriver.ChromeOptions()
# options.add_argument('--save-page-as-mhtml')
# 设置chromedriver，并打开webdriver
# driver = webdriver.Chrome()
# driver.get(news_url)
# 模拟键盘操作
# ctrl + S
win32api.keybd_event(0x11, 0, 0, 0)
win32api.keybd_event(0x53, 0, 0, 0)
win32api.keybd_event(0x53, 0, win32con.KEYEVENTF_KEYUP, 0)
win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
sleep(2)
# sleep(2)
# print(win32gui.GetCursorPos())
# sleep(2)
# print(win32gui.GetCursorPos())
# sleep(2)
# print(win32gui.GetCursorPos())
# sleep(2)
# print(win32gui.GetCursorPos())
# 貼上檔案名稱
win32clipboard.OpenClipboard()
win32clipboard.EmptyClipboard()
win32clipboard.SetClipboardText(locNum)
win32clipboard.CloseClipboard()
win32api.keybd_event(0x11, 0, 0, 0)
win32api.keybd_event(0x56, 0, 0, 0)
win32api.keybd_event(0x56, 0, win32con.KEYEVENTF_KEYUP, 0)
win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
sleep(2)

# 點擊路徑
for i in range (0,6):
    #tab 6
    win32api.keybd_event(0x09, 0, 0, 0)
    win32api.keybd_event(0x09, 0, win32con.KEYEVENTF_KEYUP, 0)
    sleep(1)
#Enter
win32api.keybd_event(0x0D, 0, 0, 0)
win32api.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)
sleep(2)
win32clipboard.OpenClipboard()
win32clipboard.EmptyClipboard()
win32clipboard.SetClipboardText(pyPath)
win32clipboard.CloseClipboard()
win32api.keybd_event(0x11, 0, 0, 0)
win32api.keybd_event(0x56, 0, 0, 0)
win32api.keybd_event(0x56, 0, win32con.KEYEVENTF_KEYUP, 0)
win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
sleep(2)
win32api.keybd_event(0x0D, 0, 0, 0)
win32api.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)
sleep(2)
# 點擊名稱
for i in range (0,6):
    #tab 6
    win32api.keybd_event(0x09, 0, 0, 0)
    win32api.keybd_event(0x09, 0, win32con.KEYEVENTF_KEYUP, 0)
    sleep(1)
#Enter
win32api.keybd_event(0x0D, 0, 0, 0)
win32api.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)
sleep(2)
#確認覆蓋 Y
win32api.keybd_event(0x59, 0, 0, 0)
win32api.keybd_event(0x59, 0, win32con.KEYEVENTF_KEYUP, 0)
sleep(2)
# 预估下载时间，后期根据实际网速调整
sleep(5)

# soup = BeautifulSoup(driver.page_source, 'html.parser')
# f = open(pyPath + "/" + "Testt" + "_soup.txt","a", encoding='UTF-8')
# f.write(str(soup))
# f.close

# urls = soup.find_all('a',href=re.compile("gateweb"))
# print(urls)

# soup = BeautifulSoup(driver.page_source, 'lxml')
# f = open(pyPath + "/" + "Testt" + "_soup_lxml.txt","a", encoding='UTF-8')
# f.write(str(soup))
# f.close
# driver.close