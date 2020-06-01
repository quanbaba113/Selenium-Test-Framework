# -*- coding: utf-8 -*-
import os
import logging
import threading
from models import my_config


class My_Logger(object):
    """
    自定义日志类。日志分两类：
        1.总体日志，记录日志执行整体过程，存放在./log目录下，命名读取配置文件logFileName参数
        2.测试日志，记录测试执行过程中的详细日志。有两部分：
            （1）debug log，记录从浏览器打开到结束所有的详细过程，按./log/starttime//browser/runtime/testname记录
            （2）error log，只记录错误日志，日志将被记录在html结果文件中
    """
    def __init__(self, logger_name='Auto Test'):
        self.logger = logging.getLogger(logger_name)
        # 日志输出格式
        self.formatter = logging.Formatter("%(asctime)s - %(levelname)-4s - %(message)s")
        # 指定日志的最低输出级别，默认为WARN级别
        self.logger.setLevel(logging.INFO)

    def get_handler(self, file_path):
        # 生成文件日志的handler
        p, f = os.path.split(file_path)
        if not (os.path.exists(p)):
            os.makedirs(p)  # 判断是否存在该路径，如果不存在就新创建
        file_handler = logging.FileHandler(file_path)
        file_handler.setFormatter(self.formatter)
        return file_handler


my_logger = My_Logger()
default_log_path = os.path.join(my_config.LOG_PATH, my_config.logFileName)
my_lock = threading.RLock()


def DebugLogger(log_info, file_path=default_log_path):
    """debug日志，记录所有日志"""
    try:
        if my_lock.acquire():
            file_handler = my_logger.get_handler(file_path)
            my_logger.logger.addHandler(file_handler)
            my_logger.logger.info(log_info)
            my_logger.logger.removeHandler(file_handler)

            my_lock.release()
    except Exception as e:
        print("Failed to record debug log. Reason:\n %s" % str(e))


def ErrorLogger(log_info, action_list=[], file_path=default_log_path):
    """用以在用例内记录错误日志"""
    try:
        if my_lock.acquire():
            if len(action_list) == 0:
                action_list.append(["", "", ""])
            action_list[-1][2] = action_list[-1][2] + "\n" + log_info   # 将日志记录在在执行的用例的最后一个action中

            file_handler = my_logger.get_handler(file_path)
            my_logger.logger.addHandler(file_handler)
            my_logger.logger.error(log_info)
            my_logger.logger.removeHandler(file_handler)

            my_lock.release()
    except Exception as e:
        print("Failed to record error log. Reason:\n %s" % str(e))
