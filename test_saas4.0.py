#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author  : WangYongsheng
# @Email   : wys920811@163.com
# @Date    : 2018-11-06 16:15:58

import os
import re
import xlrd
import time
import logging
import unittest
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait


class DataIsSameTest(unittest.TestCase):

    def setUp(self):
        logging.basicConfig(level=logging.INFO,  
                    format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',  
                    datefmt='%a, %d %b %Y %H:%M:%S',  
                    filename='./test.log',  
                    filemode='w')  

        driver_dir = r'D:\安装包\chromedriver_win32\chromedriver.exe'
        options = webdriver.ChromeOptions()
        prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': os.getcwd()}
        options.add_experimental_option('prefs', prefs)
        self.browser = webdriver.Chrome(executable_path=driver_dir, chrome_options=options)
        self.browser.implicitly_wait(20)
        self.profit_url = 'http://118.31.46.211:801/#/profit'
        self.management_url = 'http://118.31.46.211:801/#/management'
        self.get_web_and_excel_data()

    def tearDown(self):
        self.browser.quit()

    def is_excel_exist(self, driver):
        l = ['业务概况.xlsx', '业务走势图.xlsx', '利润概况.xlsx', '利润趋势图.xlsx', '收支明细.xlsx', '运营概况.xlsx']
        return l == os.listdir()[2:]

    def get_web_infos(self, url):
        query_cmd = 'document.getElementsByClassName("ivu-btn ivu-btn-primary")[1].click()'
        lirungaikuang_cmd = 'return document.getElementsByClassName("ivu-card-body")[0].innerText'
        yujishouzhimingxi_cmd = 'return document.getElementsByClassName("ivu-card-body")[2].innerText'
        shijishouzhimingxi_cmd = 'return document.getElementsByClassName("ivu-card-body")[3].innerText'
        if 'profit' in url:
            web_lirungaikuang = self.browser.execute_script(lirungaikuang_cmd)
            web_yujishouzhimingxi = self.browser.execute_script(yujishouzhimingxi_cmd)
            web_shijishouzhimingxi = self.browser.execute_script(shijishouzhimingxi_cmd)
            self.web_lirungaikuang = [i for i in web_lirungaikuang.replace(' ','').splitlines() if i != '']
            self.web_yujishouzhimingxi = [i for i in web_yujishouzhimingxi.replace(' ','').splitlines() if i != '']
            self.web_shijishouzhimingxi = [i for i in web_shijishouzhimingxi.replace(' ','').splitlines() if i != '']

        else:
            web_yewugaikuang = self.browser.execute_script(lirungaikuang_cmd)
            web_yunyinggaikuang = self.browser.execute_script(yujishouzhimingxi_cmd)
            self.web_yewugaikuang = [i for i in web_yewugaikuang.replace(' ','').splitlines() if i != '']
            self.web_yunyinggaikuang = [i for i in web_yunyinggaikuang.replace(' ','').splitlines() if i != '']

    def get_web_and_excel_data(self):
        for i in os.listdir():
            if 'xlsx' in i:
                os.remove(i)

        self.browser.get(self.profit_url)
        self.get_web_infos(self.profit_url)
        download_js1 = 'document.getElementsByClassName("ivu-card-head")[0].getElementsByTagName("a")[0].click()'
        download_js2 = 'document.getElementsByClassName("ivu-card-head")[1].getElementsByTagName("a")[0].click()'
        download_js3 = 'document.getElementsByClassName("ivu-card-head")[2].getElementsByTagName("a")[0].click()'
        self.browser.execute_script(download_js1)
        self.browser.execute_script(download_js2)
        self.browser.execute_script(download_js3)
        
        self.browser.get(self.management_url)
        time.sleep(10)
        self.get_web_infos(self.management_url)
        self.browser.execute_script(download_js1)
        self.browser.execute_script(download_js2)
        self.browser.execute_script(download_js3)
        WebDriverWait(self.browser, 10, 0.5).until(self.is_excel_exist)

    # 利润概况表
    def read_lirungaikuang(self):
        work_book = xlrd.open_workbook('利润概况.xlsx')
        sheet = work_book.sheet_by_index(0)
        nrows = sheet.nrows
        ncols = sheet.ncols
        pattern = re.compile(r'^[-+]?[-0-9]\d*\.\d*|[-+]?\.?[0-9]\d*$')
        l = []
        for i in range(ncols):
            l += sheet.col_values(i)
        for i in range(len(l)):
            if pattern.match(l[i]):
                l[i] = str(round(float(l[i])))
        self.excel_lirungaikuang = l[:8] + l[16:24] + l[8:16] + l[24:]

    # 收支明细
    def read_shouzhimingxi(self):
        work_book = xlrd.open_workbook('收支明细.xlsx')
        yuji_sheet = work_book.sheet_by_index(0)
        shiji_sheet = work_book.sheet_by_index(1)
        self.excel_yujishouzhimingxi = []
        self.excel_shijishouzhimingxi = []
        for i in range(1, yuji_sheet.ncols):
            if i != 12 and i != 14:
                y = [str(round(sum(float(x) for x in yuji_sheet.col_values(i, 1) if '%' not in x) / 10000, 4))]
                self.excel_yujishouzhimingxi += yuji_sheet.col_values(i, 0, 1) + y
        for i in range(1, shiji_sheet.ncols):
            if i != 12 and i != 14:
                z = [str(round(sum(float(x) for x in shiji_sheet.col_values(i, 1) if '%' not in x) / 10000, 4))]
                self.excel_shijishouzhimingxi += shiji_sheet.col_values(i, 0, 1) + z

    # 利润概况数据一致性
    def test_lirungaikuang_data_is_the_same_as_web(self):
        self.read_lirungaikuang()
        logging.info('==============excel_lirungaikuang===================')
        logging.info(self.excel_lirungaikuang)
        logging.info('==============excel_lirungaikuang===================')
        logging.info('==============web_lirungaikuang===================')
        logging.info(self.web_lirungaikuang)
        logging.info('==============web_lirungaikuang===================')
        self.assertTrue(self.web_lirungaikuang == self.excel_lirungaikuang, '利润概况数据不一致')

    # 预计收支明细数据一致性
    def test_yujishouzhimingxi_data_is_the_same_as_web(self):
        self.read_shouzhimingxi()
        logging.info('==============excel_yujishouzhimingxi===================')
        logging.info(self.excel_yujishouzhimingxi)
        logging.info('==============excel_yujishouzhimingxi===================')
        logging.info('==============web_yujishouzhimingxi===================')
        logging.info(self.web_yujishouzhimingxi)
        logging.info('==============web_yujishouzhimingxi===================')
        self.assertTrue(self.excel_yujishouzhimingxi == self.web_yujishouzhimingxi, '预计收支明细数据不一致')

    # 实际收支明细数据一致性
    def test_实际shouzhimingxi_data_is_the_same_as_web(self):
        self.read_shouzhimingxi()
        logging.info('==============excel_shijishouzhimingxi===================')
        logging.info(self.excel_shijishouzhimingxi)
        logging.info('==============excel_shijishouzhimingxi===================')
        logging.info('==============web_shijishouzhimingxi===================')
        logging.info(self.web_shijishouzhimingxi)
        logging.info('==============web_shijishouzhimingxi===================')
        self.assertTrue(self.excel_shijishouzhimingxi == self.web_shijishouzhimingxi, '实际收支明细数据不一致')


if __name__ == '__main__':
    unittest.main(warnings='ignore')