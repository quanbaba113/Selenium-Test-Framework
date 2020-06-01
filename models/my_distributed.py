# -*- coding: utf-8 -*-
import datetime
import threading
import multiprocessing
import sys
import os
from models import my_config
from models import my_report
from models.my_run import RunTest
from models.my_utils import read_test_plan, read_rerun_test_plan


class Distributed(object):
    """分布式"""
    def __init__(self):
        self.browser_list = []
        self.environment_list = []
        self.start_time = None
        self.duration = {}
        self.test_url = {}
        self.case_num_total = 0

    def set_multi_browser(self):
        try:
            if my_config.isMultiBrowser == "Yes":
                self.browser_list = my_config.browserType.split(";")
            else:
                self.browser_list.append(my_config.browserType.split(";")[0])
        except:
            self.browser_list.append(my_config.browserType)

    def set_multi_environment(self):
        try:
            if my_config.isSameEnvironment == "Yes":
                self.environment_list.append(my_config.environmentAdress.split(";")[0])
            else:
                self.environment_list = my_config.environmentAdress.split(";")
        except:
            self.environment_list.append(my_config.environmentAdress)

    def create_process(self):
        '''创建多进程 执行不同的浏览器'''
        processes = []
        self.set_multi_browser()
        self.set_multi_environment()
        start_time = datetime.datetime.now()
        self.start_time = start_time    # 记录开始时间
        q = multiprocessing.Queue()
        for index, browser_type in enumerate(self.browser_list):
            if len(self.environment_list) > index:
                environment_address = self.environment_list[index]
            else:
                environment_address = self.environment_list[-1]
            self.test_url[browser_type] = environment_address   # 记录不同浏览器的测试环境
            p = multiprocessing.Process(target=self.create_thread, args=(start_time, browser_type, environment_address, q))
            processes.append(p)
        for i in range(0, len(processes)):
            processes[i].start()
        for j in range(0, len(processes)):
            processes[j].join()

        for n in range(q.qsize()):
            duration_list = q.get()
            self.duration[duration_list[0]] = duration_list[1]  # 记录每个浏览器的执行时间
        stop_time = datetime.datetime.now()
        print('\nTime Elapsed: %s' % str(stop_time - start_time).split(".")[0], file=sys.stderr)

    def create_thread(self, start_time, browser_type, environment_address, q):
        '''创建多线程 提升执行效率'''
        report = my_report.TestReport(plan_path=my_config.PLAN_PATH, report_path=my_config.REPORT_PATH,
                                      browser_type=browser_type, title=my_config.report_title,  # 测试结果主标题
                                      description=f'Environment：Windows7, Browser：{browser_type}')  # 测试结果的副标题
        last_run_start_time = datetime.datetime.now()
        # 默认测试执行一次
        try:
            run_time = int(my_config.runTime)
            if run_time < 1:
                run_time = 1
        except:
            run_time = 1
        for index in range(run_time):   # rerun 机制
            current_run_start_time = datetime.datetime.now()
            my_report.default_thread_lock = threading.RLock()
            my_report.default_result = []
            threads = []
            if index == 0:
                plan_tuple = read_test_plan(my_config.PLAN_PATH, browser_type)
                self.case_num_total = len(plan_tuple[0])
            else:
                excel_result_path = os.path.join(my_config.REPORT_PATH, start_time.strftime('%Y%m%d%H%M%S'),
                                                 browser_type, "Run%s" % index,
                                                 last_run_start_time.strftime('%Y_%m_%d_%H%M') + "_Report.xls")
                plan_tuple = read_rerun_test_plan(excel_result_path, browser_type)
            case_num = len(plan_tuple[0])
            try:
                errorRatio = float(my_config.errorRatio.strip("%"))/100
                if errorRatio < 0 or errorRatio > 1:
                    errorRatio = 0.1
            except Exception as e:
                print('Param "errorRatio" has an error value, '
                      'System has chosen the default value (10%) for the "errorRatio"')
                errorRatio = 0.1
            if case_num / self.case_num_total <= errorRatio:
                break   # 判断错误率是否低于配置值，如果是，停止执行
            report.generate_report_path(start_time, current_run_start_time, index + 1)
            try:
                numOfThread = int(my_config.numOfThread)
                if numOfThread < 1:
                    numOfThread = 1
            except:
                numOfThread = 1
            case_num_thread = case_num // numOfThread
            case_num_other = case_num % numOfThread
            for i in range(numOfThread):
                if i < case_num_other:
                    start_num = (case_num_thread + 1) * i
                    end_num = (case_num_thread + 1) * (i + 1)
                else:
                    start_num = (case_num_thread + 1) * case_num_other + case_num_thread * (i - case_num_other)
                    end_num = (case_num_thread + 1) * case_num_other + case_num_thread * (i - case_num_other + 1)
                if i < numOfThread - 1:
                    plan_tuple_thread = (plan_tuple[0][start_num:end_num], plan_tuple[1][start_num:end_num],
                                         plan_tuple[2][start_num:end_num], plan_tuple[3][start_num:end_num],
                                         plan_tuple[4][start_num:end_num])
                else:
                    plan_tuple_thread = (plan_tuple[0][start_num:], plan_tuple[1][start_num:],
                                         plan_tuple[2][start_num:], plan_tuple[3][start_num:],
                                         plan_tuple[4][start_num:])
                run_thread = threading.Thread(target=RunTest(plan_tuple_thread, browser_type,
                                                             environment_address, index + 1).run_test)
                threads.append(run_thread)
            report_thread = threading.Thread(target=report.generate_report, args=(str(case_num),))
            threads.append(report_thread)
            for t in threads:
                t.start()
            for t in threads:
                t.join()
            current_run_stop_time = datetime.datetime.now()
            last_run_start_time = current_run_start_time
            print('\nTime Elapsed of %s Run%s: %s\n' % (browser_type, index+1, str(
                current_run_stop_time - current_run_start_time).split(".")[0]), file=sys.stderr)
        stop_time = datetime.datetime.now()
        q.put([browser_type, str(stop_time - start_time).split(".")[0]])    # 记录不同浏览器的执行时间
