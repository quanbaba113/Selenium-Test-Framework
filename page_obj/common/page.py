# -*- coding: utf-8 -*-

import datetime
import os
from time import sleep

import sys

from models import my_log
from models import my_report
from models.my_config import findElementTime, LOG_PATH
from models.my_utils import screen_shot

"""
    1、本文件是所有page文件的基类，后续所有的page类均需要继承本文件中的class page
    2、公用方法应在本文件中定义
    3、封装的公用方法都要记录日志，记录方式分两种：
        1、操作日志：self.DebugLogger(info)
        2、错误日志：self.ErrorLogger(info)
"""


class Page(object):
    """页面基础类，存放元素的公共操作方法"""
    def __init__(self, test_case):
        self.test_case = test_case
        self.test_url = test_case.environment_address
        self.driver = test_case.driver
        self.test_case_name = test_case.id().split(".")[-1]
        self.action_list = test_case.action_list

    def DebugLogger(self, log_info):
        # debug日志
        name = self.test_case.CaseNameRestore()
        file_path = os.path.join(LOG_PATH, my_report.default_start_time.strftime('%Y%m%d%H%M%S'),
                                 self.test_case.browser_type, "Run%s" % self.test_case.run_index, "%s.log" % name)
        my_log.DebugLogger(log_info, file_path)

    def ErrorLogger(self, log_info):
        # error 日志
        name = self.test_case.CaseNameRestore()
        file_path = os.path.join(LOG_PATH, my_report.default_start_time.strftime('%Y%m%d%H%M%S'),
                                 self.test_case.browser_type, "Run%s" % self.test_case.run_index, "%s.log" % name)
        my_log.ErrorLogger(log_info, self.action_list, file_path)

    def ScreenShot(self, screen_name):
        screen_shot(test_case=self.test_case, file_name=screen_name)

    def RecordFailStatus(self):
        self.test_case.RecordFailStatus(exc_info=sys.exc_info())

    def open_url(self):
        """打开网页"""
        try:
            self.driver.get(self.test_url)
            self.driver.implicitly_wait(10)
            self.DebugLogger("Pass : Success to open '%s'" % self.test_url)
        except Exception as e:
            self.ErrorLogger("Unexpected error : Unable to open '%s'" % self.test_url)
            raise e

    def find_element(self, *loc):
        """循环查找元素，当超时后报错"""
        time_start = datetime.datetime.now()
        e = False
        while (datetime.datetime.now() - time_start).seconds < int(findElementTime):
            try:
                element = self.driver.find_element(*loc)
                self.DebugLogger("Pass : Success to find element '%s'" % loc[1])
                return element
            except Exception as ex:
                e = ex
                sleep(1)
        else:
            self.ErrorLogger("Unexpected error: Unable to find element '%s'" % loc[1])
            raise e

    def find_elements(self, *loc):
        """循环查找符合条件的所有元素，当超时后报错"""
        time_start = datetime.datetime.now()
        e = False
        while (datetime.datetime.now() - time_start).seconds < int(findElementTime):
            try:
                elements = self.driver.find_elements(*loc)
                if len(elements) > 0:
                    self.DebugLogger("Pass : Success to find elements '%s'" % loc[1])
                    return elements
                sleep(1)
            except Exception as ex:
                e = ex
                sleep(1)
        else:
            self.ErrorLogger("Unexpected error: Unable to find elements '%s'" % loc[1])
            raise e

    def switch_frame(self, frame):
        """切换frame"""
        try:
            if isinstance(frame, str):
                frame_name = frame
            else:
                frame_name = frame.get_attribute("src")
        except:
            frame_name = frame
        try:
            self.driver.switch_to.frame(frame)
            self.DebugLogger("Pass : Success to switch to frame '%s'" % frame_name)
        except Exception as e:
            self.ErrorLogger("Unexpected error: Unable to switch to frame '%s'" % frame_name)
            raise e

    def switch_alert(self):
        """切换到弹出框"""
        try:
            self.driver.switch_to_alert()
            self.DebugLogger("Pass : Success to switch to alert")
        except Exception as e:
            self.ErrorLogger("Unexpected error: Unable to switch to alert")
            raise e

    def switch_content(self):
        """返回默认的父窗口"""
        try:
            self.driver.switch_to_default_content()
            self.DebugLogger("Pass : Success to switch to default content")
        except Exception as e:
            self.ErrorLogger("Unexpected error: Unable to switch to default content")
            raise e

    def refresh(self):
        """刷新页面"""
        try:
            self.driver.refresh()
            self.DebugLogger("Pass : Success to refresh page")
        except Exception as e:
            self.ErrorLogger("Unexpected error: Unable to refresh page")
            raise e

    def back(self):
        """页面后退"""
        try:
            self.driver.back()
            self.DebugLogger("Pass : Success to back page")
        except Exception as e:
            self.ErrorLogger("Unexpected error: Unable to back page")
            raise e

    def forward(self):
        """页面后退"""
        try:
            self.driver.forward()
            self.DebugLogger("Pass : Success to forward page")
        except Exception as e:
            self.ErrorLogger("Unexpected error: Unable to forward page")
            raise e

    def get_title(self):
        """获取当前页的标题"""
        try:
            title = self.driver.title
            self.DebugLogger("Pass : Success to get title '%s'" % title)
            return title
        except Exception as e:
            self.ErrorLogger("Unexpected error: Unable to get title" )
            raise e

    def get_url(self):
        """获取当前页的url"""
        try:
            url = self.driver.current_url
            self.DebugLogger("Pass : Success to get current url" )
            return url
        except Exception as e:
            self.ErrorLogger("Unexpected error: Unable to get current url" )
            raise e

    def input_text(self, element, text):
        """文本框输入"""
        try:
            element.send_keys(text)
            self.DebugLogger("Pass : Success to input '%s'" % text)
        except Exception as e:
            self.ErrorLogger("Unexpected error: Unable to input '%s'" % text)
            raise e

    def click(self, element):
        """单击"""
        time_start = datetime.datetime.now()
        e = False
        while (datetime.datetime.now() - time_start).seconds < 5:
            try:
                element.click()
                self.DebugLogger("Pass : Success to click element")
                break
            except Exception as ex:
                e = ex
                sleep(1)
        else:
            self.ErrorLogger("Unexpected error: Click Failed. Error info:%s" % str(e))
            raise e
