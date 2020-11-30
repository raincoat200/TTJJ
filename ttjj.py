#!coding:utf-8
import os
from selenium import webdriver
import pandas as pd
from pandas import DataFrame
import datetime
import random
from time import sleep
# ui模块
import tkinter as tk
from tkinter import filedialog

application_window = tk.Tk()
# 设置文件对话框会显示的文件类型
my_filetypes = [('all files', '.*'), ('text files', '.txt')]
# 请求选择文件
filename = filedialog.askopenfilename(parent=application_window,
                                    initialdir=os.getcwd(),
                                    title="请选择你要处理的EXCEL",
                                    filetypes=my_filetypes)

chrome_path = '..\\chromedriver.exe'
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless') #增加无界面选项
chrome_options.add_argument('--disable-gpu') #如果不加这个选项 有时定位会出现问题
browser = webdriver.Chrome(executable_path=chrome_path, options=chrome_options)

try:
    sheet = pd.read_excel(filename, "Sheet1")
    for row in sheet.index.values:
        jjdm = str(sheet.iloc[row, 0]).zfill(6)
        sheet.iloc[row, 0]=jjdm
        url = "https://fund.eastmoney.com/" + jjdm + ".html"
        print(url)
        browser.get(url)
        elem = browser.find_element_by_class_name("fundDetail-tit").text
        # 基金名称
        sheet.iloc[row, 1]=elem[0:elem.find("(")]
        dwjz = browser.find_element_by_class_name("dataItem02")
        # 单位净值
        sheet.iloc[row, 2] =dwjz.find_element_by_class_name("ui-font-large").text
        # 累计净值
        ljjz = browser.find_element_by_class_name("dataItem03")
        sheet.iloc[row, 3] = ljjz.find_element_by_class_name("ui-font-large").text
        #成立来
        sheet.iloc[row, 7] = ljjz.find_element_by_xpath("./dd[3]/span[2]").text
        # 基金规模
        elem = browser.find_element_by_class_name("infoOfFund")
        sheet.iloc[row, 4]= elem.find_element_by_xpath("./table/tbody/tr/td[2]").text.replace("基金规模：", "")
        # 基金经理
        sheet.iloc[row, 5] =elem.find_element_by_xpath("./table/tbody/tr/td[3]").text.replace("基金经理：", "")
        # 成立日
        sheet.iloc[row, 6] =elem.find_element_by_xpath("./table/tbody/tr[2]/td[1]").text.replace("成 立 日：", "")
        # 阶段涨跌幅
        elem = browser.find_element_by_id('increaseAmount_stage')
        sheet.iloc[row, 8] =elem.find_element_by_xpath('./table/tbody/tr[2]/td[2]').text
        sheet.iloc[row, 9] =elem.find_element_by_xpath('./table/tbody/tr[2]/td[3]').text
        sheet.iloc[row, 10] =elem.find_element_by_xpath('./table/tbody/tr[2]/td[4]').text
        sheet.iloc[row, 11] =elem.find_element_by_xpath('./table/tbody/tr[2]/td[5]').text
        sheet.iloc[row, 12] =elem.find_element_by_xpath('./table/tbody/tr[2]/td[6]').text
        sheet.iloc[row, 13] =elem.find_element_by_xpath('./table/tbody/tr[2]/td[7]').text
        sheet.iloc[row, 14] =elem.find_element_by_xpath('./table/tbody/tr[2]/td[8]').text
        sheet.iloc[row, 15] =elem.find_element_by_xpath('./table/tbody/tr[2]/td[9]').text
except Exception as e:
    print(e)
browser.quit()
print(sheet)
DataFrame(sheet).to_excel(filename, sheet_name='Sheet1', index=False, header=True)
os.startfile(os.path.dirname(filename))