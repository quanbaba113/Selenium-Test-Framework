# -*- coding: utf-8 -*-
import unittest, threading
import xlrd
from models import my_utils
from models import my_report
from models import my_config
from models.my_log import DebugLogger, ErrorLogger
# 自动导入测试需要的类
plan_list = xlrd.open_workbook(my_config.PLAN_PATH)  # 获取plan file的数据表
sheet_list = plan_list.sheet_names()
unique_class_list = []
for sheet_name in sheet_list:
    test_plan_tuple = my_utils.read_test_plan(my_config.PLAN_PATH, sheet_name)
    directory_list = test_plan_tuple[0]
    class_list = test_plan_tuple[1]

    for index, class_name in enumerate(class_list):
        if class_name not in unique_class_list:
            unique_class_list.append(class_name)
            exec("from " + directory_list[index] + " import " + class_name)


class TransferParamToCase(unittest.TestCase):
    """利用staticmethod生成decorate类成员函数，并返回一个函数对象"""
    @staticmethod
    def transferParamToCase(test_class, test_case, test_description, test_data):
        def func(self):
            record_test_data = test_data  # 存储传递的测试数据 test_data是字典类型
            self.__doc__ = test_description
            eval(test_class + '.' + test_case + '(' + 'self' + ', ' + 'record_test_data' + ')')   # 调用testcase
        return func


class BatchRunAll:
    def __init__(self, suite, plan_tuple, browser_type, environment_address, run_index):
        self.suite = suite
        self.plan_tuple = plan_tuple
        self.test_data_path = my_config.DATA_PATH    # my_config.DATA_PATH
        # 读取测试计划
        self.test_class_list = []
        self.test_case_list = []
        self.test_description_list = []
        self.test_data_num_list = []
        self.test_data_list = []
        self.browser_type = browser_type
        self.environment_address = environment_address
        self.run_index = run_index

    def get_suite(self):
        """读取plan list的列表并加入到执行列表中"""
        self.test_class_list = self.plan_tuple[1]
        self.test_case_list = self.plan_tuple[2]
        self.test_description_list = self.plan_tuple[3]
        self.test_data_num_list = self.plan_tuple[4]
        # 读取测试数据
        self.test_data_list = my_utils.read_test_data(self.test_data_path, self.plan_tuple)

        for i in range(0, len(self.test_class_list)):
            test_class = self.test_class_list[i]
            test_case = self.test_case_list[i]
            test_description = self.test_description_list[i]
            data_row_num = self.test_data_num_list[i]  # 获取需要执行的数据起止行
            if data_row_num == '':
                data_row_num = "1-1"  # 如果起止行为空，默认为1-1
            start_row = min(int(data_row_num.split("-")[0]), int(data_row_num.split("-")[1]))
            end_row = max(int(data_row_num.split("-")[0]), int(data_row_num.split("-")[1]))
            if start_row == end_row:
                test_data = self.test_data_list[i][0]
                setattr(eval(test_class), test_case + '_%s' % str(start_row),
                        TransferParamToCase.transferParamToCase(test_class, test_case, test_description, test_data))
                the_case = eval(test_class)(test_case + '_%s' % str(start_row))
                the_case.browser_type = self.browser_type
                the_case.environment_address = self.environment_address
                the_case.run_index = self.run_index
                the_case.data_row_num = data_row_num
                self.suite.addTest(the_case)
            else:
                for j in range(0, end_row-start_row+1):
                    test_data = self.test_data_list[i][j]
                    setattr(eval(test_class), test_case + '_%s_%s' % (str(end_row-start_row + 1),str(start_row+j)),
                            TransferParamToCase.transferParamToCase(test_class, test_case, test_description, test_data))
                    the_case = eval(test_class)(test_case + '_%s_%s' % (str(end_row-start_row + 1),str(start_row+j)))
                    the_case.browser_type = self.browser_type
                    the_case.environment_address = self.environment_address
                    the_case.run_index = self.run_index
                    the_case.data_row_num = data_row_num
                    self.suite.addTest(the_case)
                    # 当多行数据算一行时，默认生成 用例名_行数_第几行 的方法名
        return self.suite


class RunTest(object):
    def __init__(self, plan_tuple, browser_type, environment_address, run_index):
        self.plan_tuple = plan_tuple
        self.browser_type = browser_type
        self.environment_address= environment_address
        self.run_index = run_index

    def run_test(self):
        """unittest测试套件执行测试用例"""
        # 定义测试套件suite
        suite = unittest.TestSuite()
        batch_run = BatchRunAll(suite, self.plan_tuple, self.browser_type, self.environment_address, self.run_index)
        suite = batch_run.get_suite()
        result = my_report._TestResult(verbosity=2)

        try:
            DebugLogger("Run test(Browser:%s & RunTime:run%s & ThreadName:%s)" %
                        (self.browser_type, self.run_index, threading.current_thread().name))
            suite(result)
            # 执行测试用例
            DebugLogger("Stop test(Browser:%s & RunTime:run%s & ThreadName:%s)" %
                        (self.browser_type, self.run_index, threading.current_thread().name))
        except Exception as ex:
            ErrorLogger("Failed to run test(Browser:%s & RunTime:run%s & ThreadName:%s), Error info:%s" %
                        (self.browser_type, self.run_index, threading.current_thread().name, ex))
