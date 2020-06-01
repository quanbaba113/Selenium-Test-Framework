# -*- coding: utf-8 -*-
from models.my_mail import Email
from models import my_config
from models.my_distributed import Distributed
from models.my_log import DebugLogger, ErrorLogger
from models.my_utils import generate_summary_report, compress_report


if __name__ == "__main__":
    distributed = Distributed()
    try:
        DebugLogger("Start all test")
        distributed.create_process()    # 创建多线程并执行测试
        DebugLogger("Stop all test")
    except Exception as e:
        ErrorLogger("There are some error when running test, error info: %s" % e)
    try:
        DebugLogger("Generate summary report")
        generate_summary_report(distributed.start_time, distributed.test_url, distributed.duration)  # 整理报告
    except Exception as e:
        ErrorLogger("There are some error when Generating summary report, error info: %s" % e)
    try:
        DebugLogger("Compress report")
        compress_report(distributed.start_time)  # 压缩结果文件
    except Exception as e:
        ErrorLogger("There are some error when Compressing report, error info: %s" % e)
    try:
        DebugLogger("Send Email")
        sent_email = Email(my_config.server, my_config.sender, my_config.password, my_config.receiver,
                           my_config.copyList, my_config.title + distributed.start_time.strftime("%Y_%m_%d_%H%M"),
                           sender_info=my_config.senderInfo,
                           result_name=distributed.start_time.strftime("%Y%m%d%H%M%S"))
        sent_email.send()   # 发送邮件
    except Exception as e:
        ErrorLogger("There are some error when Sending Email, error info: %s" % e)
