# -*- coding: utf-8 -*-
from selenium.webdriver import Remote
from selenium import webdriver


# 启动浏览器
def browser(browserType):
    # 配置浏览器类型，默认是谷歌
    if browserType == "Ie":
        driver = webdriver.Ie()
    elif browserType == "Firefox":
        driver = webdriver.Firefox()
    else:
        driver = webdriver.Chrome()
    return driver


# Notes:
# 1. 在运行这段代码之前，要先分别在controller和remoter的机器安装Java
# 2. 都需要下载一个selenium-server的jar包，下载地址：http://selenium-release.storage.googleapis.com/index.html
# 3. 在controller(192.168.1.247)机器上启动Selenium-Server：java -jar D:\安装包path\selenium-server-standalone-3.4.0.jar -role hub
# 或者在那个包所在文件夹下打开CMD命令行 java -jar selenium-server-standalone-3.4.0.jar -role hub
# 4. 在remoter(192.168.2.153）机器上注册一个node：
# java -jar D:\Software\Selenium\selenium-server-standalone-3.4.0.jar -role node -hub http://192.168.1.247:4444/grid/register
# 远程启动浏览器
def remotebrowser():
    # 启动到远程主机中，运行自动化测试
    host = '192.168.2.153:5555' # 运行(脚本)主机：端口号， 默认是本机的4444端口,可以查看hub里面自动分配的端口号
    driver = Remote(command_executor='http://' + host + '/wd/hub',
                    desired_capabilities={'platform': 'ANY',
                                          'browserName': 'firefox',
                                          'version': '',
                                          'javascriptEnabled': True})
    return driver