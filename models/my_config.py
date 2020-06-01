# -*- coding: utf-8 -*-
import os
from models.my_reader import IniReader

"""
读取配置。这里配置文件用的INI，也可用其他如XML,yaml等，需在file_reader中添加相应的Reader进行处理。
"""
# 通过当前文件的绝对路径，其父级目录一定是框架的base目录，然后确定各层的绝对路径。如果你的结构不同，可自行修改。
# 之前直接拼接的路径，修改了一下，这种方法可以支持linux和windows等不同的平台
# 建议大家多用os.path.split()和os.path.join()，不要直接+'\\xxx\\ss'这样
# 框架路径
BASE_PATH = os.path.split(os.path.dirname(os.path.abspath(__file__)))[0]    # 'D:\\wfm3Selenium\\frameworks'
# 配置文件路径
CONFIG_INI = os.path.join(BASE_PATH, "config", "config.ini")    # 'D:\\wfm3Selenium\\frameworks\\config\\config.ini'
CONFIG_FILE = os.path.join(BASE_PATH, 'config', 'config.yml')   # 'D:\\wfm3Selenium\\frameworks\\config\\config.yml'
# 测试日志路径
LOG_PATH = os.path.join(BASE_PATH, 'log')   # 'D:\\wfm3Selenium\\frameworks\\log'
# 测试报告路径
REPORT_PATH = os.path.join(BASE_PATH, 'test_report')    # 'D:\\wfm3Selenium\\frameworks\\test_report'
# 测试计划路径
PLAN_PATH = os.path.join(BASE_PATH, "test_plan", IniReader(CONFIG_INI, "PlanName", "planFileName").data)
# 测试数据路径
DATA_PATH = os.path.join(BASE_PATH, 'test_data', IniReader(CONFIG_INI, "TestData", "dataFolder").data)
# 邮件各参数
server = IniReader(CONFIG_INI, "MailAddress", "smtpServer").data
sender = IniReader(CONFIG_INI, "MailAddress", "senderMail").data
password = IniReader(CONFIG_INI, "MailAddress", "passWord").data
receiver = IniReader(CONFIG_INI, "MailAddress", "receiverMail").data
copyList = IniReader(CONFIG_INI, "MailAddress", "copyList").data
title = IniReader(CONFIG_INI, "MailAddress", "subjectMail").data
senderInfo = IniReader(CONFIG_INI, 'MailAddress', 'senderInfo').data
shareFolder = IniReader(CONFIG_INI, 'MailAddress', 'shareFolder').data
# 报告信息配置
report_title = IniReader(CONFIG_INI, 'ReportInfo', 'reportTitle').data
# 公共日志名称
logFileName = IniReader(CONFIG_INI, 'LogSetting', 'logFileName').data
# 元素查询超时时间
findElementTime = IniReader(CONFIG_INI, 'TimeOut', 'findElementTime').data
# 多浏览器配置
isMultiBrowser = IniReader(CONFIG_INI, 'MultiBrowser', 'isEnabled').data
isSameEnvironment = IniReader(CONFIG_INI, 'MultiBrowser', 'isSameEnvironment').data
environmentAdress = IniReader(CONFIG_INI, 'MultiBrowser', 'testUrl').data
browserType = IniReader(CONFIG_INI, 'MultiBrowser', 'browserType').data
# 分布式配置
numOfThread = IniReader(CONFIG_INI, 'MultiThread', 'numOfThread').data
enableRemoteBrowser = IniReader(CONFIG_INI, 'MultiThread', 'enableRemoteBrowser').data
browserIP = IniReader(CONFIG_INI, 'MultiThread', 'browserIP').data
# 重跑配置
runTime = IniReader(CONFIG_INI, 'ReRun', 'runTime').data
errorRatio = IniReader(CONFIG_INI, 'ReRun', 'errorRatio').data
