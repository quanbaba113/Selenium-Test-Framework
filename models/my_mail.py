# -*- coding: utf-8 -*-
import os
import re
import pandas
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from models.my_log import DebugLogger, ErrorLogger
from models.my_config import shareFolder, REPORT_PATH


class EmailTemplate(object):
    # ------------------------------------------------------------------------
    # HTML Template

    HTML_TEMP = r"""
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
    %(stylesheet)s
</head>
<body>
    <div id='honorific'>
        <p>
            <font size="3">Dear All:</font>
        </p>
    <div>
    <br/>
    <div id='summary_text'>
        <p>
            <font size="3">本次自动化测试结果整体如下：</font>
        </p>
        %(summary_table)s
    </div>
    <br/>
    <div id='exception_text'>
        <p>
            <font size="3">测试结果异常的用例列表如下：</font>
        </p>
        %(exception_table)s
    </div>
    
    <br/>
    
    <div id='last_text'>
        <p>
            <font size="3">详细信息请参考附件. 更多测试结果数据请查看: <a href="%(link)s">%(link)s</a></font>
        </p>
        <br/>
        <p>
            <font size="3">Thanks!</font>
        </p>
        <br/>
        
    </div>
    <div>
        <p>
            <font size="3" color="#4C7430">%(sender_info)s</font>
        </p>
    </div>
</body>
</html>
"""

    STYLESHEET_TEMP = """
    <style type="text/css" media="screen">
        body        { font-family: verdana, arial, helvetica, sans-serif; font-size: 80%; }
        table       { font-size: 100%; border-collapse:collapse; border-spacing:0px; }
        td          { table-layout:fixed;word-break:break-all;border: 1px solid #777; text-align:center;}
        th          { table-layout:fixed;word-break:break-all;border: 1px solid #777; text-align:center;}
        tr          { height: 4ex;}
        thead       { font-size: 12px;background-color: #CCFFCC;}
        tbody       { font-size: 10px;}
        /* -- summary ---------------------------------------------------------------------- */
        #summary_text { 
            padding-left:40px;
            padding-right:40px;
        } 
        #summary_table {
            width: 80%;
            border: 2px solid #777;
        }

        /* -- exception ---------------------------------------------------------------------- */
        #exception_text { 
            padding-left:40px;
            padding-right:40px;
        } 
        #exception_table {
            width: 50%;
            border: 2px solid #777;
        }
        #Fail {
            background-color: #FF99CC;
        }
        #Error {
            background-color: #FF0000;
        }
        #Skip {
            background-color: #FFFF00;
        }

        /* -- last    ---------------------------------------------------------------------- */
        #last_text {
            padding-left:40px;
            padding-right:40px;
        }
    </style>
"""

    def __init__(self, result_path, sender_info, result_name):
        self.result_path = result_path
        self.sender_info = sender_info
        self.result_name = result_name

    def generate_summary_table(self, summary_sheet):
        summary_text = str(summary_sheet.to_html(header=True, index=False))
        summary_text = summary_text.replace('class="dataframe"', 'id="summary_table"')
        if 'style="text-align: right;"' in summary_text:
            summary_text = summary_text.replace('style="text-align: right;"', "")
        elif 'style="text-align: left;"' in summary_text:
            summary_text = summary_text.replace('style="text-align: left;"', "")
        return summary_text

    def generate_exception_table(self, exception_sheet):
        exception_text = str(exception_sheet.to_html(header=True, index=False))
        exception_text = exception_text.replace('class="dataframe"', 'id="exception_table"')
        exception_text = exception_text.replace("<td>Fail</td>","<td id='Fail'>Fail</td>")
        exception_text = exception_text.replace("<td>Error</td>", "<td id='Error'>Error</td>")
        exception_text = exception_text.replace("<td>Skip</td>", "<td id='Skip'>Skip</td>")
        if 'style="text-align: right;"' in exception_text:
            exception_text = exception_text.replace('style="text-align: right;"', "")
        elif 'style="text-align: left;"' in exception_text:
            exception_text = exception_text.replace('style="text-align: left;"', "")
        return exception_text

    def generate_email_html(self):
        result_excel = pandas.ExcelFile(self.result_path)
        summary_sheet = result_excel.parse(0)   # 获取结果表里summary的数据
        exception_sheet = result_excel.parse(1)     # 获取结果表里exception的数据

        summary_text = self.generate_summary_table(summary_sheet)
        exception_text = self.generate_exception_table(exception_sheet)

        try:
            info_list = self.sender_info.split("\n")
            info_detail = ''
            for i in range(0, len(info_list)):
                if i != len(info_list) - 1 and info_list[i] != '':
                    info_detail = "%s%s<br/>" % (info_detail, info_list[i])
                else:
                    info_detail = "%s%s" % (info_detail, info_list[i])
        except:
            info_detail = self.sender_info

        output = self.HTML_TEMP % dict(
            stylesheet=self.STYLESHEET_TEMP,
            summary_table=summary_text,
            exception_table=exception_text,
            link="%s\\%s" % (shareFolder, self.result_name),
            sender_info=info_detail,
        )
        return output


class Email:
    """邮件类。用来给指定用户发送邮件。可指定多个收件人，可带附件。"""

    def __init__(self, server, sender, password, receiver, copy_list, title, sender_info, result_name):
        """
        :param server: smtp服务器
        :param sender: 发件人
        :param password: 发件人密码
        :param receiver: 收件人，多收件人用“；”隔开
        :param copy_list: 抄送人，多抄送人用“；”隔开
        :param title: 邮件标题
        :param sender_info: 发件人信息
        :param result_name: 结果文件夹名
        """
        self.server = server
        self.sender = sender
        self.password = password
        self.receiver = self.str_to_list(receiver)
        self.copy_list = self.str_to_list(copy_list)
        self.title = title
        self.sender_info = sender_info
        self.result_name = result_name
        self.path = os.path.join(REPORT_PATH, result_name)

    def str_to_list(self, str_list):
        strL = []
        if ";" in str_list:
            strL = str_list.split(';')
        else:
            strL.append(str_list)
        return strL

    def _attach_file(self, att_file):
        """将单个文件添加到附件列表中"""
        att = MIMEText(open('%s' % att_file, 'rb').read(), 'plain', 'utf-8')
        att["Content-Type"] = 'application/octet-stream'
        file_name = re.split(r'[\\|/]', att_file)
        att["Content-Disposition"] = 'attachment; filename="%s"' % file_name[-1]
        self.msg.attach(att)

    def send(self):
        self.msg = MIMEMultipart('related')
        self.msg['Subject'] = self.title
        self.msg['From'] = self.sender
        self.msg['To'] = ";".join(self.receiver)
        self.msg['Cc'] = ";".join(self.copy_list)

        # 添加附件，支持多个附件
        file_list = os.listdir(self.path)  # 获取报告目录下的文件目录
        for file_name in file_list:
            file_path = os.path.join(self.path, file_name)
            if os.path.splitext(file_path)[1] in (".xls", ".zip"):
                self._attach_file(file_path)

        # 邮件正文
        for file_name in file_list:
            file_path = os.path.join(self.path, file_name)
            if os.path.splitext(file_path)[1] == ".xls":
                email_temp = EmailTemplate(file_path, self.sender_info, self.result_name)
                output = email_temp.generate_email_html()
                self.msg.attach(MIMEText(output, 'html'))

        # 连接服务器并发送
        try:
            smtp_server = smtplib.SMTP(self.server)  # 连接sever
        except:
            pass
            ErrorLogger('Failed to send email, unable to connect to SMTP server，'
                        'please check network and SMTP server')
        else:
            try:
                smtp_server.login(self.sender, self.password)  # 登录
            except:
                pass
                ErrorLogger('Failure of user and password verification')
            else:
                smtp_server.sendmail(self.sender, self.receiver + self.copy_list, self.msg.as_string())  # 发送邮件
            finally:
                smtp_server.quit()  # 断开连接
                DebugLogger('Success to send "{0}" email. Addressee:{1}. if you do not receive the email, '
                            'please check dustbin and '
                            'confirm that the direction is correct.'.format(self.title, self.receiver))
