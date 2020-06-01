# -*- coding: utf-8 -*-
import yaml
import os
import configparser


class YamlReader:
    """
    文件读取。YamlReader读取yaml文件，ExcelReader读取excel。
    """
    def __init__(self, CONFIG_FILE, element, index=0):
        if os.path.exists(CONFIG_FILE):
            self.yaml_file = CONFIG_FILE
        else:
            raise FileNotFoundError('文件不存在！')
        self.element = element
        self.index = index
        self._data = None

    @property
    def data(self):
        # 如果是第一次调用data，读取yaml文档，否则直接返回之前保存的数据
        if not self._data:
            with open(self.yaml_file, 'rb') as f:
                self._data = list(yaml.safe_load_all(f))  # load后是个generator，用list组织成列表
        # yaml是可以通过'---'分节的。用YamlReader读取返回的是一个list，第一项是默认的节，如果有多个节，可以传入index来获取。
        # 这样我们其实可以把框架相关的配置放在默认节，其他的关于项目的配置放在其他节中。可以在框架中实现多个项目的测试。
        return self._data[self.index].get(self.element)


class IniReader:
    """读取ini文件中的内容,返回值：str。
    指定config_name，param_name读取对应的值: read_config("BrowserName", "browserName")
    """
    def __init__(self, CONFIG_INI, config_name, param_name):
        if os.path.exists(CONFIG_INI):
            self.ini_file = CONFIG_INI
        else:
            raise FileNotFoundError('文件不存在！')
        self.config_name = config_name
        self.param_name = param_name
        self._data = str()

    @property
    def data(self):
        config = configparser.ConfigParser()
        config.read(self.ini_file)
        self._data = config.get(self.config_name, self.param_name)
        return self._data
