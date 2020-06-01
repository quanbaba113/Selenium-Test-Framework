# -*- coding: utf-8 -*-
from models import my_utils
from models import my_driver
from models import my_log
from models import my_config
from models import my_report
import unittest, os


class MyTest(unittest.TestCase):
    """抽象出来MyTest类，在编写unittest用例时不再考虑setUp() & tearDown()方法"""
    def setUp(self):
        """setUp中主要负责浏览器的开启"""
        self.action_list = []   # 创建一个列表，记录用例的Action
        try:
            self.driver = my_driver.browser(self.browser_type)
            self.driver.implicitly_wait(10)
            self.driver.maximize_window()
            self.DebugLogger('Pass : Open the %s browser' % self.browser_type)
        except:
            self.ErrorLogger('Unexpected error: Unable to open the %s browser' % self.browser_type)

    def tearDown(self):
        """tearDown中主要负责浏览器的关闭"""
        # 错误时截图
        if self._outcome.errors[-1][1] is not None and len(self._outcome.skipped) == 0:
            self.ScreenShot()
        try:
            self.driver.quit()
            self.DebugLogger('Pass : Close browser')
        except:
            self.ErrorLogger('Unexpected error: Unable to close browser')

    def doCleanups(self):
        """重写doCleanups 自定义结果处理"""
        unittest.TestCase.doCleanups(self)
        self.DoWithResult()

    def DebugLogger(self, log_info):
        # debug日志
        name = self.CaseNameRestore()
        file_path = os.path.join(my_config.LOG_PATH, my_report.default_start_time.strftime('%Y%m%d%H%M%S'),
                                 self.browser_type, "Run%s" % self.run_index, "%s.log" % name)
        my_log.DebugLogger(log_info, file_path)

    def ErrorLogger(self, log_info):
        # error日志
        name = self.CaseNameRestore()
        file_path = os.path.join(my_config.LOG_PATH, my_report.default_start_time.strftime('%Y%m%d%H%M%S'),
                                 self.browser_type, "Run%s" % self.run_index, "%s.log" % name)
        my_log.ErrorLogger(log_info, self.action_list, file_path)

    def CaseNameRestore(self):
        """用例名恢复"""
        name = self.id().split('.')[-1]
        name_list = name.split("_")
        if name_list[-2].isdigit():
            num = len(name_list[-2]) + len(name_list[-1]) + 2
        else:
            num = len(name_list[-1]) + 1
        name = name[:-num] + "_" + self.data_row_num.split("-")[0] + "_" + self.data_row_num.split("-")[1]
        return name

    def ScreenShot(self, screen_name=''):
        my_utils.screen_shot(test_case=self, file_name=screen_name)

    def RecordFailStatus(self, exc_info):
        """记录错误状态"""
        self._outcome.errors.append((self, (AssertionError, exc_info[1], exc_info[2])))
        if len(self.action_list) == 0:
            self.action_list.append(["", "", ""])
        self.action_list[-1][1] = "fail"
        self.ScreenShot()   # 错误时截图

    def DoWithResult(self):
        """结果处理"""
        if len(self.action_list) == 0:
            self.action_list.append(["", "", ""])
        if self.action_list[0][0] == "":
            self.action_list[0][0] = self.CaseNameRestore()   # 判断用例第一个Action是否有name，如果没有就默认为用例名
        if self._outcome.success is True:
            isFail = False
            error_value = None
            error_tb = None
            for index, (test, exc_info) in enumerate(self._outcome.errors):
                if exc_info is not None:
                    isFail = True
                    error_value = exc_info[1]
                    error_tb = exc_info[2]
            if isFail is True:
                self._outcome.errors.clear()
                self._outcome.errors.append((self, (AssertionError, error_value, error_tb)))
                self._outcome.success = False
            if self.action_list[-1][1] == "":   # 如果最后一个Action没有状态，则给予pass
                self.action_list[-1][1] = "pass"
        else:
            if len(self._outcome.skipped) > 0:
                self._outcome.errors.clear()
                self.action_list[-1][1] = "skip"  # 最后一步设为skip
            else:
                isError = False
                error_type = AssertionError
                error_value = None
                error_tb = None
                for index, (test, exc_info) in enumerate(self._outcome.errors):
                    if exc_info is not None:
                        if issubclass(exc_info[0], self.failureException):
                            pass
                        else:
                            error_type = exc_info[0]
                            isError = True
                        error_value = exc_info[1]
                        error_tb = exc_info[2]
                self._outcome.errors.clear()
                self._outcome.errors.append((self, (error_type, error_value, error_tb)))
                if isError is True:
                    self.action_list[-1][1] = "error"    # 最后一步设为error
                else:
                    self.action_list[-1][1] = "fail"    # 最后一步设为fail
