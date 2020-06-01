# -*- coding: utf-8 -*-
import datetime
import copy
import os
import sys
import unittest
import xlrd
import xlwt
import xlutils.copy
import threading
from time import sleep
from xml.sax import saxutils

__version__ = "1.0"
default_current_run_start_time = datetime.datetime.now()
default_start_time = datetime.datetime.now()
default_stop_time = datetime.datetime.now()
default_html_report_path = ''
default_excel_report_path = ''
default_browser_type = ''
default_title = 'Unit Test Report'
default_description = 'Environment：Windows7, Browser：Chrome'
default_thread_lock = threading.RLock()
default_result = []


class HtmlTemplate(object):
    """Define a HTML template for report """
    STATUS = {
        0: 'pass',
        1: 'fail',
        2: 'error',
        3: 'skip'
    }

    # ------------------------------------------------------------------------
    # HTML Template

    HTML_TEMP = r"""
<html>
<head>
    <title>%(title)s</title>
    <meta name="generator" content="%(generator)s"/>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
    %(stylesheet)s
</head>
<body>
    <script language="javascript" type="text/javascript">
        /* level - 0:Summary; 1:Failed; 2:All */
        function showCase(level) {
            var num = /[0-9]/;
            trs = document.getElementsByTagName("tr");
            for (var i = 0; i < trs.length; i++) {
                tr = trs[i];
                id = tr.id;
                if (level == 0) {
                    if (num.test(id)) {
                        tr.className = 'hiddenRow';
                    }
                }
                if (level == 1) {
                    if (num.test(id)) {
                        var status = id.substr(0,2);
                        if (status == 'ft'||status == 'et'||status == 'fa'||status == 'ea') {
                            tr.className = '';
                        }else {
                            tr.className = 'hiddenRow';
                        }
                    }
                }
                if (level == 2) {
                    if (num.test(id)) {
                        var status = id.slice(1,2);
                        if (status == 't') {
                            tr.className = '';
                        }else {
                            tr.className = 'hiddenRow';
                        }
                    }
                }
                if (level == 3) {
                    if (num.test(id)) {
                        var status = id.slice(1,2);
                        if (status == 'l') {
                            tr.className = 'hiddenRow';
                        }else {
                            tr.className = '';
                        }
                    }
                }
            }
        }
         
        function showClassDetail(cid, count) {
            var id_list = Array(count);
            for (var i = 0; i < count; i++) {
                tid0 = 't' + cid.substr(1) + '.' + (i+1);
                tid = 'f' + tid0;
                tr = document.getElementById(tid);
                if (!tr) {
                    tid = 'e' + tid0;
                    tr = document.getElementById(tid);
                }
                if (!tr) {
                    tid = 'p' + tid0;
                    tr = document.getElementById(tid);
                }
                if (!tr) {
                    tid = 's' + tid0;
                    tr = document.getElementById(tid);
                }
                id_list[i] = tid;
            }
            var hasShow = false;
            var showCount = 0;
            for (var i = 0; i < count; i++) {
                tid = id_list[i];
                if (!document.getElementById(tid).className) {
                    hasShow = true;
                    showCount++;
                }
            }
            if (showCount == count) {
                hasShow = false;
            }
            for (var i = 0; i < count; i++) {
                tid = id_list[i];
                if (hasShow){
                    document.getElementById(tid).className = '';
                }else{
                    if(document.getElementById(tid).className) {
                        document.getElementById(tid).className = '';
                    }else {
                        document.getElementById(tid).className = 'hiddenRow';
                        for (var j = 0; j < 1000; j++) {
                            aid0 = 'a' + tid.substr(2) + '.' + (j+1);
                            aid = 'f' + aid0;
                            tr = document.getElementById(aid);
                            if (!tr) {
                                aid = 'e' + aid0;
                                tr = document.getElementById(aid);
                            }
                            if (!tr) {
                                aid = 'p' + aid0;
                                tr = document.getElementById(aid);
                            }
                            if (!tr) {
                                aid = 's' + aid0;
                                tr = document.getElementById(aid);
                            }
                            if (!tr){
                                break;
                            }
                            document.getElementById(aid).className = 'hiddenRow';
                            lid = aid.substr(0,1) + 'l' + aid.substr(2) + '_log';
                            if (document.getElementById(lid)) {
                                document.getElementById(lid).className = 'hiddenRow';
                            }
                        }
                    }
                }
            }
        }     
        
        function showCaseDetail(tid, count) {
            var id_list = Array(count);
            for (var i = 0; i < count; i++) {
                aid0 = 'a' + tid.substr(1) + '.' + (i+1);
                aid = 'f' + aid0;
                tr = document.getElementById(aid);
                if (!tr) {
                    aid = 'e' + aid0;
                    tr = document.getElementById(aid);
                }
                if (!tr) {
                    aid = 'p' + aid0;
                    tr = document.getElementById(aid);
                }
                if (!tr) {
                    aid = 's' + aid0;
                    tr = document.getElementById(aid);
                }
                id_list[i] = aid;
            }
            var hasShow = false;
            var showCount = 0;
            for (var i = 0; i < count; i++) {
                aid = id_list[i];
                if (!document.getElementById(aid).className) {
                    hasShow = true;
                    showCount++;
                }
            }
            if (showCount == count) {
                hasShow = false;
            }
            for (var i = 0; i < count; i++) {
                aid = id_list[i];
                lid = aid.substr(0,1) + 'l' + aid.substr(2) + '_log';
                if (hasShow){
                    document.getElementById(aid).className = '';
                }else{
                    if(document.getElementById(aid).className) {
                        document.getElementById(aid).className = '';
                    }else {
                        document.getElementById(aid).className = 'hiddenRow';
                        if (document.getElementById(lid)) {
                            document.getElementById(lid).className = 'hiddenRow';
                        }
                    }
                }
            }
        }
      
        function showLogDetail(aid) {
            lid0 = 'l' + aid.substr(1) + '_log';
            lid = 'f' + lid0;
            tr = document.getElementById(lid);
            if (!tr) {
                lid = 'e' + lid0;
                tr = document.getElementById(lid);
            }
            if (!tr) {
                lid = 'p' + lid0;
                tr = document.getElementById(lid);
            }
            if (!tr) {
                lid = 's' + lid0;
                tr = document.getElementById(lid);
            }
            if(document.getElementById(lid).className) {
                document.getElementById(lid).className = '';
            }else {
                document.getElementById(lid).className = 'hiddenRow';
            }
        }
    </script>
    %(heading)s
    %(table)s
    %(ending)s
</body>
</html>
"""

    STYLESHEET_TEMP = """
    <style type="text/css" media="screen">
        body        { font-family: verdana, arial, helvetica, sans-serif; font-size: 80%; }
        table       { font-size: 100%; border-collapse:collapse; border-spacing:0px; }
        pre         { }
        
        /* -- heading ---------------------------------------------------------------------- */
        h1 {
            font-size: 16pt;
            color: gray;
        }
        .heading {
            margin-top: 0ex;
            margin-bottom: 1ex;
        }
        
        .heading .attribute {
            margin-top: 1ex;
            margin-bottom: 0;
        }
        
        .heading .description {
            margin-top: 3ex;
            margin-bottom: 5ex;
        }
        
        /* -- report ------------------------------------------------------------------------ */
        #result_table {
            width: 100%;
            border: 2px solid #777;
        }
        #header_row {
            height: 4ex;
            font-size: 16px;
            font-weight: bold;
            color: white;
            background-color: #777;
        }
        #total_row  { 
            height: 4ex;
            font-size: 16px;
            font-weight: bold; 
        }
        #result_table td {
            border-left: 1px solid #777;
            border-top: 1px solid #777;
        }
        #action_table tr:first-child td{
            border-top:none;
        }
        #action_table tr td:first-child{
            border-left:none;
        }
        tr {
            height: 3ex;
        }
        td {
            padding: 0;
            table-layout:fixed;
            word-break:break-all;
        }
        
        .passClass  { background-color: #32CD32; height:4ex; }
        .failClass  { background-color: #FA8072; height:4ex; }
        .errorClass { background-color: #B22222; height:4ex; }
        .skipClass  { background-color: #FF8C00; height:4ex; }
        /*
        .passCase   { color: #32CD32; font-weight: bold; }
        .failCase   { color: #FA8072; font-weight: bold; }
        .errorCase  { color: #B22222; font-weight: bold; } 
        .skipCase   { color: #FF8C00; font-weight: bold;}
        */
        .passCell   { background-color: #32CD32; font-size: 16px;}
        .failCell   { background-color: #FA8072; font-size: 16px;}
        .errorCell  { background-color: #B22222; font-size: 16px;}
        .skipCell   { background-color: #FF8C00; font-size: 16px;}
        .passAction { background-color: #32CD32; }
        .failAction { background-color: #FA8072; }
        .errorAction{ background-color: #B22222; }
        .skipAction { background-color: #FF8C00; }
        .hiddenRow  { display: none; }
        .testcase   { margin-left: 1em; }
        .action_name{ margin-left: 1em; }
        .action_log { margin-left: 1em; margin-right: 1em;}
        
        /* -- ending ---------------------------------------------------------------------- */
        #ending {}
    </style>
"""

    # Heading -----------------------------------------------------------------------
    HEADING_TEMP = """
    <div class='heading'>
        <h1>%(title)s</h1>
        %(parameters)s
        <p class='description'>%(description)s</p>
    </div>
"""    # variables: (title, parameters, description)

    HEADING_ATTRIBUTE_TEMP = """<p class='attribute'><strong>%(name)s:</strong> %(value)s</a>"""
    # variables: (name, value)

    # Report ------------------------------------------------------------------------
    REPORT_TEMP = """
    <p>Show
        <a>&nbsp;</a>
        <a href='javascript:showCase(0)'>Summary</a>
        <a href='javascript:showCase(1)'>Failed</a>
        <a href='javascript:showCase(2)'>AllCase</a>
        <a href='javascript:showCase(3)'>AllStep</a>
    </p>
    <table id='result_table'>
        <tr id='header_row'>
            <td width='56%%' align='left'>Test Group/Test case</td>
            <td width='5.76%%' align='center'>Count</td>
            <td width='5.76%%' align='center'>Pass</td>
            <td width='5.76%%' align='center'>Fail</td>
            <td width='5.76%%' align='center'>Error</td>
            <td width='5.76%%' align='center'>Skip</td>
            <td width='7.2%%' align='center'>Status</td>
            <td width='8%%' align='center'>ScreenShot</td>
        </tr>
        %(test_list)s
        <tr id='total_row'>
            <td>Total</td>
            <td align='center'>%(count)s</td>
            <td align='center'>%(Pass)s</td>
            <td align='center'>%(fail)s</td>
            <td align='center'>%(error)s</td>
            <td align='center'>%(skip)s</td>
            <td align='center'>&nbsp;</td>
            <td align='center'>&nbsp;</td>
        </tr>
    </table>
"""  # variables: (test_list, count, Pass, fail, error, skip) # 表格模板

    REPORT_CLASS_TEMP = r"""
    <tr class='%(style)s'>
        <td><strong>%(desc)s</strong></td>
        <td align='center'><strong>%(count)s</strong></td>
        <td align='center'><strong>%(Pass)s</strong></td>
        <td align='center'><strong>%(fail)s</strong></td>
        <td align='center'><strong>%(error)s</strong></td>
        <td align='center'><strong>%(skip)s</strong></td>
        <td align='center'><a href="javascript:showClassDetail('%(cid)s',%(count)s)"><strong>Detail</strong></a></td>
        <td align='center'></td>
    </tr>
"""  # variables: (style, desc, count, Pass, fail, skip, error, cid) # 测试模块class模板

    REPORT_TEST_TEMP = r"""
    <tr id='%(tid)s' class='%(classType)s'>
        <td class='%(styleCase)s'><div class='testcase'><strong>%(desc)s</strong></div></td>
        <td colspan='6' id='action_cell'>
            <table width = '100%%', id = 'action_table'>
                <colgroup>
                    <col align='center' width='16%%'/>
                    <col align='center' width='16%%'/>
                    <col align='center' width='16%%'/>
                    <col align='center' width='16%%'/>
                    <col align='center' width='16%%'/>
                    <col align='center' width='20%%'/>
                </colgroup>
                <tr>
                    <td align='center' colspan='2'>%(startTime)s</td>
                    <td align='center' colspan='2'>%(stopTime)s</td>
                    <td align='center' colspan='1'>%(duration)s</td>
                    <td class='%(styleAction)s' align='center' colspan='1'>
                        <a href="javascript:showCaseDetail('%(tid0)s',%(count)s)">%(status)s</a></td>
                </tr>
                %(action_list)s
            </table>
        </td>
        <td align='center'>
            <a href="%(image)s">%(screenShot)s</a>
        </td>
    </tr>
"""  # variables: (tid, classType, styleCase, desc, startTime, stopTime, duration, styleAction, tid0, count,
    # status, action_list, image, screenShot) case模板

    REPORT_ACTION_TEMP = r"""
    <tr id='%(aid)s' class='%(classType)s'>
        <td colspan='1'><div class= 'action_code' align='center'><strong>%(action_code)s</strong></div></td>
        <td colspan='4'><div class= 'action_name' align='center'>%(action)s</div></td>
        <td colspan='1' class='%(styleAction)s' align='center'>
            <a href="javascript:showLogDetail('%(aid0)s')">%(status)s</a></td>
    </tr>
    %(action_log)s
"""     # variables: (aid, classType, action_code, action, styleAction, aid0, status, action_log) case的详细step模板

    REPORT_LOG_TEMP = r"""
    <tr id='%(lid)s' class='hiddenRow'>
        <td align='left' colspan='6'><div class= 'action_log'>%(log)s</div></td>
    </tr>
"""     # variables: (lid, action_log) 日志存放模板

    # ENDING-----------------------------------------------------------------------
    ENDING_TEMP = """<div id='ending'>&nbsp;</div>"""
    # Html文件尾模板
# -------------------- The end of the Template class ------------------------------


class HtmlReport(HtmlTemplate):
    """generate html report"""
    def __init__(self):
        self.success_count = 0
        self.failure_count = 0
        self.error_count = 0
        self.skipped_count = 0

    def generate_report(self, result):
        """main method"""
        report_attr = self.get_report_attributes(result)  # 处理结果数据
        generator = 'HTMLTestRunner %s' % __version__
        stylesheet = self._generate_stylesheet()    # 生成Html的CSS样式
        heading = self._generate_heading(report_attr)  # 生成Html的head
        table = self._generate_report(result)  # 生成Html的报告表格
        ending = self._generate_ending()    # 生成Html的底部说明
        output = self.HTML_TEMP % dict(
            title=saxutils.escape(default_title),
            generator=generator,
            stylesheet=stylesheet,
            heading=heading,
            table=table,
            ending=ending,
        )
        # 生成结果
        fp = open(default_html_report_path, "wb")
        fp.write(output.encode('utf8'))
        fp.close()

    def get_report_attributes(self, result):
        """统计测试结果信息"""
        start_time = str(default_current_run_start_time)[:19]
        duration = str(default_stop_time - default_current_run_start_time).split(".")[0]
        status = []
        self.success_count = self.failure_count = self.error_count = self.skipped_count = 0
        for n, t, e in result:
            if n == 0: self.success_count +=1
            if n == 1: self.failure_count +=1
            if n == 2: self.error_count +=1
            if n == 3: self.skipped_count +=1

        if self.success_count: status.append('Pass %s ' % self.success_count)
        if self.failure_count: status.append('Failure %s ' % self.failure_count)
        if self.error_count: status.append('Error %s ' % self.error_count)
        if self.skipped_count: status.append('Skip %s ' % self.skipped_count)
        if status:
            status = ''.join(status)
        else:
            status = 'none'
        return [
            ('Start Time', start_time),
            ('Duration', duration),
            ('Status', status),
        ]

    def sort_result(self, result_list):
        """根据Class为单位整合result"""
        rmap = {}
        classes = []
        for n, t, e in result_list:
            cls = t.__class__
            if cls not in rmap:
                rmap[cls] = []
                classes.append(cls)
            rmap[cls].append((n, t, e))
        r = [(cls, rmap[cls]) for cls in classes]
        return r

    def _generate_stylesheet(self):
        """生成CSS样式"""
        return self.STYLESHEET_TEMP

    def _generate_heading(self, report_attr):
        """生成head"""
        a_lines = []
        for name, value in report_attr:
            line = self.HEADING_ATTRIBUTE_TEMP % dict(
                    name=saxutils.escape(name),
                    value=saxutils.escape(value),
                )
            a_lines.append(line)
        heading = self.HEADING_TEMP % dict(
            title=saxutils.escape(default_title),
            parameters=''.join(a_lines),
            description=saxutils.escape(default_description),
        )
        return heading

    def _generate_report(self, result):
        """生成table报告"""
        rows = []
        sorted_result = self.sort_result(result)
        for cid, (cls, cls_results) in enumerate(sorted_result):
            # subtotal for a class
            np = nf = ne = ns = 0  # np代表pass个数，nf代表fail, ne代表error, ns代表skip
            for n, t, e in cls_results:
                if n == 0: np += 1
                if n == 1: nf += 1
                if n == 2: ne += 1
                if n == 3: ns += 1
            if cls.__module__ == "__main__":
                name = cls.__name__
            else:
                name = "%s.%s" % (cls.__module__, cls.__name__)
            doc = cls.__doc__ and cls.__doc__ .split("\n")[0] or ""  # 获取测试模块的描述
            desc = doc and '%s: %s' % (name, doc) or name   # 记录测试模块的name和description
            # 生成测试模块row
            row = self.REPORT_CLASS_TEMP % dict(
                style=ne > 0 and 'errorClass' or nf > 0 and 'failClass' or ns > 0 and 'skipClass' or 'passClass',
                desc=desc,
                count=np+nf+ne+ns,
                Pass=np,
                fail=nf,
                error=ne,
                skip=ns,
                cid='c%s' % (cid+1)
            )
            rows.append(row)
            # 生成每个TestCase类中所有case的测试结果
            for tid, (n, t, e) in enumerate(cls_results):
                self._generate_report_case(rows, cid, tid, n, t, e)
        table = self.REPORT_TEMP % dict(   # 生成report表格
            test_list=''.join(rows),
            count=str(self.success_count+self.failure_count+self.error_count+self.skipped_count),
            Pass=str(self.success_count),
            fail=str(self.failure_count),
            error=str(self.error_count),
            skip=str(self.skipped_count)
        )
        return table

    def _generate_report_case(self, rows, cid, tid, n, t, e):
        """生成case单元行"""
        tid = '%s.%s' % (cid+1, tid+1)
        name = t.CaseNameRestore()
        doc = t.__doc__ or ""  # 打印出用例名
        desc = doc and ('%s: %s' % (name, doc)) or name
        start_time = t.start_time.strftime('%Y-%m-%d %H:%M:%S')
        stop_time = t.stop_time.strftime('%Y-%m-%d %H:%M:%S')
        duration = str(t.stop_time - t.start_time).split(".")[0]
        # 把错误的图像放在某个路径下的，如果有错误就显示图片超链接，没有就隐藏这超链接。
        if str(e) != '':    # 判断是否有错误，有错误显示截图
            screen_shot = "screenShot"
            try:
                image_url = "./screenshot/%s/" % name
            except Exception as e:
                image_url = "./screenshot/"
        else:
            image_url = ""
            screen_shot = ""

        action_rows = []
        action_num = len(t.action_list)
        for aid, [action_name, action_status, action_log] in enumerate(t.action_list):  # 生成case的详细action
            self._generate_report_action(action_rows, tid, aid, action_name, action_status, action_log)
        row = self.REPORT_TEST_TEMP % dict(
            tid=self.STATUS[n][0:1] + 't' + tid,
            classType=n == 0 and 'hiddenRow' or '',
            styleCase=n == 2 and 'errorCase' or (n == 1 and 'failCase') or (n == 3 and 'skipCase') or 'passCase',
            desc=desc,
            startTime=start_time,
            stopTime=stop_time,
            duration=duration,
            styleAction=n == 2 and 'errorCell' or (n == 1 and 'failCell') or (n == 3 and 'skipCell') or 'passCell',
            tid0='t%s' % tid,
            count=action_num,
            status=self.STATUS[n],
            action_list=''.join(action_rows),
            image=image_url,
            screenShot=screen_shot
        )
        rows.append(row)
        return rows

    def _generate_report_action(self, action_rows, tid, aid, action_name, action_status, action_log):
        """生成case的详细信息表"""
        aid = '%s.%s' % (tid, aid + 1)
        log_row = ''
        try:
            log_list = action_log.split("\n")
            log_detail = ''
            for i in range(0, len(log_list)):
                if i != len(log_list) - 1 and log_list[i] != '':
                    log_detail = "%s%s<br/>" % (log_detail, log_list[i])
                else:
                    log_detail = "%s%s" % (log_detail, log_list[i])
        except Exception as ex:
            log_detail = action_log
        if action_log != '':
            log_row = self.REPORT_LOG_TEMP % dict(
                lid=action_status[0:1] + 'l' + aid + '_log',
                log=log_detail
            )
        action_row = self.REPORT_ACTION_TEMP % dict(
            aid=action_status[0:1] + 'a' + aid,
            classType=action_status == 'pass' and 'hiddenRow' or '',
            action_code='Step%s' % aid.split(".")[2],
            action=action_name,
            styleAction=action_status == 'fail'and 'failAction' or (action_status == 'error'and 'errorAction')
                        or (action_status == 'skip'and 'skipAction') or 'passAction',
            aid0='a%s' % aid,
            status=action_status,
            action_log=''.join(log_row)
        )
        action_rows.append(action_row)
        return action_rows

    def _generate_ending(self):
        """生成Html文件尾部说明"""
        return self.ENDING_TEMP


class ExcelReport(object):
    """generate excel report"""
    def generate_excel_report(self, result, browser_type):
        """main method"""
        plan_list = xlrd.open_workbook(default_excel_report_path)  # 获取excel report数据表
        table = plan_list.sheet_by_name(browser_type)  # 获取sheet名为browser_type的表
        col_count = table.ncols  # 获取列数
        row_count = table.nrows  # 获取行数

        work_book = xlwt.Workbook(encoding='ascii')
        work_sheet = work_book.add_sheet(browser_type)
        # 设置列宽
        work_sheet.col(0).width = 256 * 8
        work_sheet.col(1).width = 256 * 30
        work_sheet.col(2).width = 256 * 20
        work_sheet.col(3).width = 256 * 20
        work_sheet.col(4).width = 256 * 30
        work_sheet.col(5).width = 256 * 12
        work_sheet.col(6).width = 256 * 12
        work_sheet.col(7).width = 256 * 20
        work_sheet.col(8).width = 256 * 20
        work_sheet.col(9).width = 256 * 10
        work_sheet.col(10).width = 256 * 10

        # 设置对齐方式
        align_center = xlwt.Alignment()
        align_center.horz = align_center.HORZ_CENTER
        align_center.vert = align_center.VERT_CENTER

        # 设置字体加粗
        bold_font = xlwt.Font()
        bold_font.bold = True

        # 设置表头格式
        style_bold = xlwt.XFStyle()
        style_bold.font = bold_font
        style_bold.alignment = align_center

        # 设置result列的背景色及对齐方式
        style_back_red = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
        style_back_red.alignment = align_center
        style_back_rose = xlwt.easyxf('pattern: pattern solid, fore_colour rose;')
        style_back_rose.alignment = align_center
        style_back_green = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;')
        style_back_green.alignment = align_center
        style_back_yellow = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
        style_back_yellow.alignment = align_center
        style_back_gray = xlwt.easyxf('pattern: pattern solid, fore_colour gray25;')
        style_back_gray.alignment = align_center

        for j in range(0, col_count):
            work_sheet.write(0, j, table.cell_value(0, j), style_bold)

        for i in range(1, row_count):
            case_value = table.cell_value(i, 3)
            if table.cell_value(i, 6) == '':
                data_row_num = "1-1"
            else:
                data_row_num = table.cell_value(i, 6)
            start_row = min(int(data_row_num.split("-")[0]), int(data_row_num.split("-")[1]))
            end_row = max(int(data_row_num.split("-")[0]), int(data_row_num.split("-")[1]))
            if start_row == end_row:
                test_num = "_%s" % str(start_row)
            else:
                test_num = "_%s_%s" % ((end_row - start_row + 1), end_row)
            test_case_name = case_value + test_num

            for single_result in result:
                if single_result[1].id().split('.')[-1] == test_case_name:
                    start_time = single_result[1].start_time
                    stop_time = single_result[1].stop_time
                    duration = str(stop_time - start_time).split(".")[0]
                    for j in range(0, col_count):
                        if j < 7 and j != 5:
                            work_sheet.write(i, j, table.cell_value(i, j))
                        else:
                            if j == 5: work_sheet.write(i, j, "Y")
                            if j == 7: work_sheet.write(i, j, str(start_time).split(".")[0])
                            if j == 8: work_sheet.write(i, j, str(stop_time).split(".")[0])
                            if j == 9: work_sheet.write(i, j, duration)
                            if j == 10:
                                if single_result[0] == 0:
                                    result_status = 'Pass'
                                    work_sheet.write(i, j, result_status, style_back_green)
                                if single_result[0] == 1:
                                    result_status = 'Fail'
                                    work_sheet.write(i, j, result_status, style_back_rose)
                                if single_result[0] == 2:
                                    result_status = 'Error'
                                    work_sheet.write(i, j, result_status, style_back_red)
                                if single_result[0] == 3:
                                    result_status = 'Skip'
                                    work_sheet.write(i, j, result_status, style_back_yellow)
                    break
            else:
                for j in range(0, col_count):
                    if j < 7 and j != 5:
                        work_sheet.write(i, j, table.cell_value(i, j))
                    else:
                        if j == 5: work_sheet.write(i, j, "N")
                        if j == 10: work_sheet.write(i, j, "Not Run", style_back_gray)

        work_book.save(default_excel_report_path)

HtmlReport = HtmlReport()
ExcelReport = ExcelReport()


class TestReport(object):
    def __init__(self, plan_path, report_path, browser_type, title=None, description=None):
        """
        :param plan_path: 测试计划的路径
        :param report_path: 测试报告的存放路径
        :param browser_type: 测试浏览器
        :param title: 测试报告标题
        :param description: 测试报告描述
        """
        self.report_path = report_path
        self.plan_path = plan_path
        self.browser_type = browser_type
        global default_browser_type
        default_browser_type = browser_type
        if title is not None:
            global default_title
            default_title = title
        if description is not None:
            global default_description
            default_description = description

    def generate_report(self, case_num):
        '''循环检查 生成报告'''
        current_result_num = 0
        while True:
            if len(default_result) == current_result_num:
                sleep(1)
            else:
                result_copy = copy.copy(default_result)
                current_result_num = len(result_copy)
                HtmlReport.generate_report(result_copy)  # 生成html报告
                ExcelReport.generate_excel_report(result_copy, self.browser_type)   # 生成Excel报告
                if str(current_result_num) == case_num:
                    break
            if threading.active_count() == 2:
                break

    def generate_report_path(self, start_time, current_run_start_time, run_index):
        """generate report path"""
        global default_current_run_start_time
        default_current_run_start_time = current_run_start_time
        global default_start_time
        default_start_time = start_time
        start_time = default_start_time.strftime('%Y%m%d%H%M%S')
        result_path = os.path.join(self.report_path, start_time, self.browser_type, "Run%s" % run_index)
        if not (os.path.exists(result_path)):  # 判断是否存在该路径，如果不存在就新创建
            os.makedirs(result_path)
        report_name = current_run_start_time.strftime('%Y_%m_%d_%H%M')
        global default_html_report_path
        default_html_report_path = self.generate_html_report(result_path, report_name)
        global default_excel_report_path
        default_excel_report_path = self.generate_excel_report(result_path, report_name)

    def generate_html_report(self, result_path, report_name):
        """生成报告存放路径"""
        result_file = os.path.join(result_path, report_name + "_Report.html" ) # 生成Html报告
        fp = open(result_file, "wb")
        fp.close()
        return result_file

    def generate_excel_report(self, result_path, report_name):
        """生成Excel结果文件"""
        plan_list = xlrd.open_workbook(self.plan_path)  # 获取plan file的数据表
        table = plan_list.sheet_by_name(self.browser_type)  # 获取sheet名为browser_type的表
        col_count = table.ncols  # 获取列数
        wb = xlutils.copy.copy(plan_list)
        ws = wb.get_sheet(wb.sheet_index(self.browser_type))
        ws.write(0, col_count, "StartTime")
        ws.write(0, col_count + 1, "EndTime")
        ws.write(0, col_count + 2, "Duration")
        ws.write(0, col_count + 3, "Result")
        result_file = os.path.join(result_path, report_name + "_Report.xls")
        wb.save(result_file)
        return result_file


class _TestResult(unittest.TestResult):
    # note: _TestResult is a pure representation of results.
    def __init__(self, verbosity=1):
        unittest.TestResult.__init__(self)
        self.verbosity = verbosity
        # result 记录用例执行结果
        self.result = []
        # 记录上几个类似的用例以整合
        self.last_similar_test_list = []

    def startTest(self, test):
        unittest.TestResult.startTest(self, test)
        test.start_time = datetime.datetime.now()

    def stopTest(self, test):
        stop_time = datetime.datetime.now()
        test.stop_time = stop_time
        if len(self.last_similar_test_list) == 0:
            if default_thread_lock.acquire():
                global default_stop_time
                default_stop_time = stop_time
                default_result.append(self.result[-1])
                sleep(1)
                default_thread_lock.release()

    def addSuccess(self, test):
        unittest.TestResult.addSuccess(self, test)
        self.processResult(0, test)
        if len(self.last_similar_test_list) > 1:
            self.result[-1] = self.mergeResult(0, test, '')
            if int(test.id().split('.')[-1].split("_")[-2]) == len(self.last_similar_test_list):
                # 判断是否最后一个相似用例，如果是打印执行结果
                n = self.result[-1][0]
                self.printResult(n, test)
        elif len(self.last_similar_test_list) == 1:
            self.result.append((0, test, ''))
        else:
            self.result.append((0, test, ''))
            self.printResult(0, test)

    def addError(self, test, err):
        unittest.TestResult.addError(self, test, err)
        _, _exc_str = self.errors[-1]
        self.processResult(2, test)
        if len(self.last_similar_test_list) > 1:
            self.result[-1] = self.mergeResult(2, test, _exc_str)
            if int(test.id().split('.')[-1].split("_")[-2]) == len(self.last_similar_test_list):
                # 判断是否最后一个相似用例，如果是打印执行结果
                n = self.result[-1][0]
                self.printResult(n, test)
        elif len(self.last_similar_test_list) == 1:
            self.result.append((2, test, _exc_str))
        else:
            self.result.append((2, test, _exc_str))
            self.printResult(2, test)

    def addSkip(self, test, reason):
        unittest.TestResult.addSkip(self, test, reason)
        self.processResult(3, test)
        if len(self.last_similar_test_list) > 1:
            self.result[-1] = self.mergeResult(3, test, reason)
            if test.id().split('.')[-1].split("_")[-2] == str(len(self.last_similar_test_list)):
                # 判断是否最后一个相似用例，如果是打印执行结果
                n = self.result[-1][0]
                self.printResult(n, test)
        elif len(self.last_similar_test_list) == 1:
            self.result.append((3, test, reason))
        else:
            self.result.append((3, test, reason))
            self.printResult(3, test)

    def addFailure(self, test, err):
        unittest.TestResult.addFailure(self, test, err)
        _, _exc_str = self.failures[-1]
        self.processResult(1, test)
        if len(self.last_similar_test_list) > 1:
            self.result[-1] = self.mergeResult(1, test, _exc_str)
            if test.id().split('.')[-1].split("_")[-2] == str(len(self.last_similar_test_list)):
                # 判断是否最后一个相似用例，如果是打印执行结果
                n = self.result[-1][0]
                self.printResult(n, test)
        elif len(self.last_similar_test_list) == 1:
            self.result.append((1, test, _exc_str))
        else:
            self.result.append((1, test, _exc_str))
            self.printResult(1, test)

    def processResult(self, n, test):
        """整理测试结果 将相似结果整合为一个"""
        test_name = test.id().split('.')[-1]
        case_num = test_name.split("_")[-2]
        if case_num.isdigit():  # 判断用例名后缀是否为 _num_num 的格式
            self.last_similar_test_list.append(test)    # 如果上一个不符合条件，从这个开始记录

    def mergeResult(self, n, t, e):
        """将相似用例的步骤整合在一起"""
        t.start_time = self.last_similar_test_list[0].start_time    # 第一次执行的开始时间为用例的开始时间
        last_result = self.result[-1]
        last_n = last_result[0]
        # 判断用例的执行状态，一旦某一次循环出错，则默认全部错误 优先级：error>fail>skip>pass
        status = (last_n or n) == 2 and 2 or ((last_n or n) == 1 and 1) or ((last_n or n) == 3 and 3) or 0
        errors = last_result[2] + e
        last_test = self.last_similar_test_list[-2]
        # print(last_test.action_list)
        for action in t.action_list:
            last_test.action_list.append(action)
        t.action_list = last_test.action_list
        return (status, t, errors)

    def printResult(self, n, test):
        """打印结果"""
        self.last_similar_test_list.clear()    # 清空相似用例列表
        name = test.CaseNameRestore()
        if self.verbosity > 1:
            if n == 0:  sys.stderr.write('Pass  ' + name + ' ' + default_browser_type + '\n')
            if n == 1:  sys.stderr.write('Fail  ' + name + ' ' + default_browser_type + '\n')
            if n == 2:  sys.stderr.write('Error  ' + name + ' ' + default_browser_type + '\n')
            if n == 3:  sys.stderr.write('Skip  ' + name + ' ' + default_browser_type + '\n')
        else:
            if n == 0:  sys.stderr.write('.')
            if n == 1:  sys.stderr.write('F')
            if n == 2:  sys.stderr.write('E')
            if n == 3:  sys.stderr.write('S')
        sys.stderr.flush()
