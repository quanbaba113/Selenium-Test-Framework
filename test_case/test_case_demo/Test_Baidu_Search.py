# -*- coding: utf-8 -*-
from models import my_unit
from models.my_utils import define_action
from page_obj.common.page import Page
from page_obj.test_demo_baidu.baidu_home_page import Baidu_Home_Page


class TS0001Test_Demo(my_unit.MyTest):
    """class名与data文件夹名和plan表TestClassName列保持一致"""
    def TC0001Baidu_Search(self, test_data):
        # 函数名与data文件名和plan表TestCaseName列保持一致
        define_action(test_case=self, action_name="open url")   # 定义事务，下同
        Page(self).open_url()

        define_action(test_case=self, action_name="search")
        Baidu_Home_Page(self).baidu_search(test_data["searchDetail"])
        self.DebugLogger("baidu search success")    # 记录执行日志
        self.ScreenShot("search_success")   # 截图

        define_action(test_case=self, action_name="research")
        if isinstance(test_data["reSearch"], list):
            for detail in test_data["reSearch"]:
                Baidu_Home_Page(self).baidu_search(detail)
        else:
            Baidu_Home_Page(self).baidu_search(test_data["reSearch"])

        define_action(test_case=self, action_name="newSearch")
        Baidu_Home_Page(self).baidu_search_new(test_data["searchDetail"])
