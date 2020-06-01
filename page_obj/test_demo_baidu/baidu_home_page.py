# -*- coding: utf-8 -*-
from page_obj.common.page import Page
from selenium.webdriver.common.by import By
from time import sleep
from models.my_utils import define_action


class Baidu_Home_Page(Page):
    """百度首页"""
    search_testbox = (By.ID, "kw")  # 搜索框
    search_btn = (By.ID, "su")  # 搜索按钮
    baike_text = (By.XPATH, ".//*[contains(text(), '百度百科')]")

    def baidu_search(self, search_content):
        # 百度搜索
        self.input_text(self.find_element(*self.search_testbox), search_content)
        sleep(1)
        self.click(self.find_element(*self.search_btn))
        # 检查点自行设计，一旦出错，默认用例fail，返回错误类型AssertionError
        try:
            self.find_element(*self.baike_text)
            self.DebugLogger("Pass : Success to search %s" % search_content)    # 记录执行日志
            # self.RecordFailStatus()
        except:
            self.ErrorLogger("Fail : Unable to search %s" % search_content)     # 记录错误日志
            self.ScreenShot("jietu")    # 截图
            self.RecordFailStatus()     # 记录错误状态 继续执行

    def baidu_search_new(self, search_content):
        self.find_element(*self.search_testbox).clear()
        self.input_text(self.find_element(*self.search_testbox), search_content)
        sleep(1)
        self.click(self.find_element(*self.search_btn))
        # 检查点自行设计，一旦出错，默认用例fail，返回错误类型AssertionError
        try:
            self.find_element(*self.baike_text)
            self.DebugLogger("Pass : Success to search %s" % search_content)  # 记录执行日志
        except:
            self.ErrorLogger("Fail : Unable to search %s" % search_content)  # 记录错误日志
