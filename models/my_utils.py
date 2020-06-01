# -*- coding: utf-8 -*-
import os
import zipfile
import xlrd, xlwt
#from PIL import Image
from models import my_report
from models import my_config


# ----------------------------------------------------------------------------------------------------------------------
#   事务定义
# ----------------------------------------------------------------------------------------------------------------------
def define_action(test_case, action_name, description=""):
    status_list = [action_name, "", description]
    test_case.action_list.append(status_list)
    num = len(test_case.action_list)
    if num > 1 and test_case.action_list[-2][1] == "":
        test_case.action_list[-2][1] = 'pass'    # 记录上一个action状态为通过


# ----------------------------------------------------------------------------------------------------------------------
#   截图函数
# ----------------------------------------------------------------------------------------------------------------------
def screen_shot(test_case, file_name=''):
    start_time = my_report.default_start_time.strftime('%Y%m%d%H%M%S')
    # 统一使用my_config下面的路径:BASE_PATH, REPORT_PATH
    test_case_name = test_case.id().split(".")[-1]
    name_list = test_case_name.split("_")
    if name_list[-2].isdigit():
        num = len(name_list[-2]) + len(name_list[-1]) + 2
    else:
        num = len(name_list[-1]) + 1
    test_case_name = test_case_name[:-num] + "_" + test_case.data_row_num.split("-")[0] + "_" + \
                     test_case.data_row_num.split("-")[1]
    img_path = os.path.join(my_config.REPORT_PATH, start_time, test_case.browser_type, "Run%s" % test_case.run_index,
                            "screenshot", test_case_name)
    if not (os.path.exists(img_path)):  # 判断是否存在该路径，如果不存在就新创建一个
        os.makedirs(img_path)
    if file_name == '':
        num = 0
        while os.path.exists(os.path.join(img_path, test_case_name + "_" + str(num) + ".png")):
            num +=1
        test_case.driver.get_screenshot_as_file(os.path.join(img_path, test_case_name + "_" + str(num) + ".png"))
    else:
        if os.path.exists(os.path.join(img_path, file_name + ".png")):
            num = 1
            while os.path.exists(os.path.join(img_path, file_name + "_" + str(num) + ".png")):
                num += 1
            test_case.driver.get_screenshot_as_file(os.path.join(img_path, file_name + "_" + str(num) + ".png"))
        else:
            test_case.driver.get_screenshot_as_file(os.path.join(img_path, file_name + ".png"))
    

# ----------------------------------------------------------------------------------------------------------------------
# 读取测试计划
# plan_file_path: 测试计划存放路径
# ----------------------------------------------------------------------------------------------------------------------
def read_test_plan(plan_file_path, browser_type):
    try:
        plan_list = xlrd.open_workbook(plan_file_path)  # 获取plan file的数据表
        table = plan_list.sheet_by_name(browser_type)  # 获取sheet名为browser_type的表
        row_count = table.nrows  # 获取行数
        test_directory_list = []
        test_class_list = []
        test_case_list = []
        test_description_list = []
        test_data_num_list = []
        for i in range(1, row_count):
            planValue = str(table.cell_value(i, 5))
            if planValue == "Y":
                test_directory_list.append(str(table.cell_value(i, 1)))
                test_class_list.append(str(table.cell_value(i, 2)))
                test_case_list.append(str(table.cell_value(i, 3)))
                test_description_list.append(str(table.cell_value(i, 4)))
                test_data_num_list.append(str(table.cell_value(i, 6)))
        return (test_directory_list, test_class_list, test_case_list, test_description_list, test_data_num_list)
    except Exception as e:
        print("Unexpected error: " + str(e))


# ----------------------------------------------------------------------------------------------------------------------
# 读取reRun测试计划
# excel_result_path: 上次测试结果存放路径
# ----------------------------------------------------------------------------------------------------------------------
def read_rerun_test_plan(excel_result_path, browser_type):
    try:
        plan_list = xlrd.open_workbook(excel_result_path)  # 获取reRun plan file的数据表
        table = plan_list.sheet_by_name(browser_type)  # 获取sheet名为browser_type的表
        row_count = table.nrows  # 获取行数
        test_directory_list = []
        test_class_list = []
        test_case_list = []
        test_description_list = []
        test_data_num_list = []
        for i in range(1, row_count):
            planValue = str(table.cell_value(i, 5))
            result_status = str(table.cell_value(i, 10))
            if planValue == "Y" and (result_status == "Error" or result_status =="Fail" or result_status == "Skip"):
                test_directory_list.append(str(table.cell_value(i, 1)))
                test_class_list.append(str(table.cell_value(i, 2)))
                test_case_list.append(str(table.cell_value(i, 3)))
                test_description_list.append(str(table.cell_value(i, 4)))
                test_data_num_list.append(str(table.cell_value(i, 6)))
        return (test_directory_list, test_class_list, test_case_list, test_description_list, test_data_num_list)
    except Exception as e:
        print("Unexpected error: " + str(e))


# ----------------------------------------------------------------------------------------------------------------------
# 读取测试数据
# test_data_path: 数据存放的路径
# plan_tuple：需要读取的数据描述
#   test_class_list：存放待读数据的模块名
#   test_case_list：存放待读数据的用例名
#   test_data_num_list：存放待读数据的起止行
# test_data_list数据结构：
#       [[用例1数据],[用例2数据],[用例3数据]....]
#             ↓
#         用例1数据[{第一行数据},{第二行数据}...]
#                        ↓
#              第一行数据{key1:value1, key2:[value2.1,value2.2...],....} 有些参数涉及的步骤需要执行多次，每次不同参数，因此需要多个value
# ----------------------------------------------------------------------------------------------------------------------
def read_test_data(test_data_path, plan_tuple):
    try:
        test_data_list = []
        test_class_list = plan_tuple[1]
        test_case_list = plan_tuple[2]
        test_data_num_list = plan_tuple[4]
        for i in range(0, len(test_class_list)):
            test_class = test_class_list[i]
            test_case = test_case_list[i]
            case_path = os.path.join(test_data_path, test_class, test_case + ".xls")
            data = xlrd.open_workbook(case_path)  # 获取case data file的数据表
            table = data.sheet_by_index(0)  # 获取第一个sheet的表
            col_count = table.ncols  # 获取列数
            data_row_num = test_data_num_list[i]  # 获取需要执行的数据起止行
            if data_row_num == '':
                data_row_num = "1-1"    # 如果起止行为空，默认为1-1
            start_row = int(data_row_num.split("-")[0])
            end_row = int(data_row_num.split("-")[1])
            action_data_list = []
            for m in range(start_row, end_row+1):
                param_dict = {}
                for j in range(0, col_count):
                    param_key = str(table.cell_value(0, j))
                    param_value = str(table.cell_value(m, j))
                    if param_value[0:1] == "$":
                        sheet_name = param_value[1:]
                        sheet_table = data.sheet_by_name(sheet_name)
                        for n in range(0, sheet_table.ncols):
                            if str(sheet_table.cell_value(0, n)) == param_key:
                                param_value = sheet_table.col_values(n, start_rowx=1)
                                break
                    param_dict[param_key] = param_value
                action_data_list.append(param_dict)
                # print(action_data_list)
            test_data_list.append(action_data_list)
        return test_data_list
    except Exception as e:
        print("Unexpected error: " + str(e))


# ----------------------------------------------------------------------------------------------------------------------
# 生成总体报告
# start_time: 测试开始执行时间
# test_url：不同浏览器的测试环境
# duration：不同浏览器的执行时间
# ----------------------------------------------------------------------------------------------------------------------

def generate_summary_report(start_time, test_url, duration):
    str_start_time = start_time.strftime('%Y%m%d%H%M%S')
    report_path = os.path.join(my_config.REPORT_PATH, str_start_time)

    # 获取result文件路径
    result_dic = {}
    browser_list = os.listdir(report_path)
    for browser in browser_list:
        result_dic[browser] = {}
        runTime_list = os.listdir(os.path.join(report_path, browser))
        for runTime in runTime_list:
            result_list = os.listdir(os.path.join(report_path, browser, runTime))
            for result in result_list:
                result_path = os.path.join(report_path, browser, runTime, result)
                if os.path.splitext(result_path)[1] == '.xls':
                    excel_result_list = xlrd.open_workbook(result_path)  # 获取result.xls的数据表
                    result_table = excel_result_list.sheet_by_name(browser)  # 获取sheet名为browser_type的表
                    result_dic[browser][runTime] = result_table

    # 设置边框
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    borders.bottom_colour = 0x3A

    # 设置字体
    bold_font = xlwt.Font()
    bold_font.bold = True

    # 设置对齐方式
    align_center = xlwt.Alignment()
    align_center.horz = align_center.HORZ_CENTER
    align_center.vert = align_center.VERT_CENTER

    # 设置summary的表头样式
    style_summary_header = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;')
    style_summary_header.borders = borders
    style_summary_header.font = bold_font
    style_summary_header.alignment = align_center

    # 设置summary的表格样式
    style_summary_cell = xlwt.XFStyle()
    style_summary_cell.borders = borders
    style_summary_cell.alignment = align_center

    # 设置result的表头样式
    style_result_header = xlwt.XFStyle()
    style_result_header.borders = borders
    style_result_header.font = bold_font
    style_result_header.alignment = align_center

    # 设置result的表头样式
    style_result_cell = xlwt.XFStyle()
    style_result_cell.borders = borders

    # 设置result的result列的背景色及对齐方式
    style_back_red = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
    style_back_red.borders = borders
    style_back_red.alignment = align_center
    style_back_rose = xlwt.easyxf('pattern: pattern solid, fore_colour rose;')
    style_back_rose.borders = borders
    style_back_rose.alignment = align_center
    style_back_green = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;')
    style_back_green.borders = borders
    style_back_green.alignment = align_center
    style_back_yellow = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
    style_back_yellow.borders = borders
    style_back_yellow.alignment = align_center
    style_back_gray = xlwt.easyxf('pattern: pattern solid, fore_colour gray25;')
    style_back_gray.borders = borders
    style_back_gray.alignment = align_center

    # 新建工作表
    work_book = xlwt.Workbook()
    work_summary_sheet = work_book.add_sheet("Summary")  # 新增Summary的sheet

    # 设置Summary的列表
    work_summary_sheet.col(0).width = 256 * 15
    work_summary_sheet.col(1).width = 256 * 8
    work_summary_sheet.col(2).width = 256 * 8
    work_summary_sheet.col(3).width = 256 * 8
    work_summary_sheet.col(4).width = 256 * 8
    work_summary_sheet.col(5).width = 256 * 15
    work_summary_sheet.col(6).width = 256 * 40
    work_summary_sheet.col(7).width = 256 * 20
    work_summary_sheet.col(8).width = 256 * 15
    work_summary_sheet.col(9).width = 256 * 12

    # 设置summary的表头信息
    work_summary_sheet.write(0, 0, "Browser Type", style_summary_header)
    work_summary_sheet.write(0, 1, "Pass", style_summary_header)
    work_summary_sheet.write(0, 2, "Fail", style_summary_header)
    work_summary_sheet.write(0, 3, "Error", style_summary_header)
    work_summary_sheet.write(0, 4, "Skip", style_summary_header)
    work_summary_sheet.write(0, 5, "Total Count", style_summary_header)
    work_summary_sheet.write(0, 6, "Test Url", style_summary_header)
    work_summary_sheet.write(0, 7, "Start Time", style_summary_header)
    work_summary_sheet.write(0, 8, "Duration", style_summary_header)
    work_summary_sheet.write(0, 9, "Rerun Time", style_summary_header)

    # 新增Result(Exception)的sheet
    work_exception_sheet = work_book.add_sheet("Result(Exception)")

    # 设置Result(Exception)的列表
    work_exception_sheet.col(0).width = 256 * 15
    work_exception_sheet.col(1).width = 256 * 40
    work_exception_sheet.col(2).width = 256 * 15
    work_exception_sheet.col(3).width = 256 * 15

    # 设置Result(Exception)的表头信息
    work_exception_sheet.write(0, 0, "BrowserType", style_summary_header)  # 表头样式与summary一致
    work_exception_sheet.write(0, 1, "TestCaseName", style_summary_header)
    work_exception_sheet.write(0, 2, "TestDescription", style_summary_header)
    work_exception_sheet.write(0, 3, "ResultStatus", style_summary_header)

    exception_row_count = 1  # 记录Result(Exception)写入行数

    # 生成不同浏览器的result 和summary
    for index, browser in enumerate(browser_list):
        work_result_sheet = work_book.add_sheet("Result(%s)" % browser) # 新增Result的sheet

        plan_list = xlrd.open_workbook(my_config.PLAN_PATH)  # 获取plan file的数据表
        plan_table = plan_list.sheet_by_name(browser)  # 获取sheet名为browser_type的表
        col_count = plan_table.ncols  # 获取列数
        row_count = plan_table.nrows  # 获取行数

        # 设置result列宽
        work_result_sheet.col(0).width = 256 * 8
        work_result_sheet.col(1).width = 256 * 30
        work_result_sheet.col(2).width = 256 * 20
        work_result_sheet.col(3).width = 256 * 20
        work_result_sheet.col(4).width = 256 * 30
        work_result_sheet.col(5).width = 256 * 12
        work_result_sheet.col(6).width = 256 * 12
        work_result_sheet.col(7).width = 256 * 10

        # 填写result表头信息
        for m in range(col_count+1):
            if m < col_count:
                work_result_sheet.write(0, m, plan_table.cell_value(0, m), style_result_header)
            else:
                work_result_sheet.write(0, m, "Result", style_result_header)
        status_count_list = [0, 0, 0, 0, 0]     # [not run, pass, skip, fail, error]
        for i in range(1, row_count):
            for j in range(col_count+1):
                if j < col_count:
                    work_result_sheet.write(i, j, plan_table.cell_value(i, j), style_result_cell)
                else:
                    # 0:Not Run/1:Pass/2:Skip/3:Fail/4:Error
                    result_status = 0
                    reRun_dic = result_dic[browser]
                    for reRun_time, result_table in reRun_dic.items():
                        status = result_table.cell_value(i, j+3)
                        n = status == "Pass" and 4 or status == "Error" and 3 or status == "Fail" and 2 or \
                            status == "Skip" and 1 or status == "Not Run" and 0 or 0
                        if n > result_status:
                            result_status = n
                    if result_status == 0:
                        status_count_list[0] += 1
                        work_result_sheet.write(i, j, "Not Run", style_back_gray)
                    if result_status == 4:
                        status_count_list[1] += 1
                        work_result_sheet.write(i, j, "Pass", style_back_green)
                    if result_status == 1:
                        status_count_list[2] += 1
                        work_result_sheet.write(i, j, "Skip", style_back_yellow)
                        # 填写Result(Exception)
                        work_exception_sheet.write(exception_row_count, 0, browser, style_summary_cell)
                        work_exception_sheet.write(exception_row_count, 1, plan_table.cell_value(i, 3),
                                                   style_summary_cell)
                        work_exception_sheet.write(exception_row_count, 2, plan_table.cell_value(i, 4),
                                                   style_summary_cell)
                        work_exception_sheet.write(exception_row_count, 3, "Skip", style_back_yellow)
                        exception_row_count +=1
                    if result_status == 2:
                        status_count_list[3] += 1
                        work_result_sheet.write(i, j, "Fail", style_back_rose)
                        # 填写Result(Exception)
                        work_exception_sheet.write(exception_row_count, 0, browser, style_summary_cell)
                        work_exception_sheet.write(exception_row_count, 1, plan_table.cell_value(i, 3),
                                                   style_summary_cell)
                        work_exception_sheet.write(exception_row_count, 2, plan_table.cell_value(i, 4),
                                                   style_summary_cell)
                        work_exception_sheet.write(exception_row_count, 3, "Fail", style_back_rose)
                        exception_row_count += 1
                    if result_status == 3:
                        status_count_list[4] += 1
                        work_result_sheet.write(i, j, "Error", style_back_red)
                        # 填写Result(Exception)
                        work_exception_sheet.write(exception_row_count, 0, browser, style_summary_cell)
                        work_exception_sheet.write(exception_row_count, 1, plan_table.cell_value(i, 3),
                                                   style_summary_cell)
                        work_exception_sheet.write(exception_row_count, 2, plan_table.cell_value(i, 4),
                                                   style_summary_cell)
                        work_exception_sheet.write(exception_row_count, 3, "Error", style_back_red)
                        exception_row_count += 1
        # 填写summary
        work_summary_sheet.write(index+1, 0, browser, style_summary_cell)
        work_summary_sheet.write(index+1, 1, status_count_list[1], style_summary_cell)
        work_summary_sheet.write(index+1, 2, status_count_list[3], style_summary_cell)
        work_summary_sheet.write(index+1, 3, status_count_list[4], style_summary_cell)
        work_summary_sheet.write(index+1, 4, status_count_list[2], style_summary_cell)
        work_summary_sheet.write(index+1, 5, status_count_list[1] + status_count_list[2] + status_count_list[3]
                                 + status_count_list[4], style_summary_cell)
        work_summary_sheet.write(index+1, 6, test_url[browser], style_summary_cell)
        work_summary_sheet.write(index+1, 7, str(start_time).split(".")[0], style_summary_cell)
        work_summary_sheet.write(index+1, 8, duration[browser], style_summary_cell)
        work_summary_sheet.write(index+1, 9, len(result_dic[browser]) - 1, style_summary_cell)

    work_book.save(os.path.join(report_path, "SummaryReport.xls"))


# ----------------------------------------------------------------------------------------------------------------------
# 压缩结果文件
# start_time: 测试开始执行时间
# ----------------------------------------------------------------------------------------------------------------------
def compress_report(start_time):
    str_start_time = start_time.strftime("%Y%m%d%H%M%S")
    report_path = os.path.join(my_config.REPORT_PATH, str_start_time)
    file_list = os.listdir(report_path)  # 获取报告目录下的文件目录
    # 新建压缩包
    report_zip = zipfile.ZipFile(os.path.join(report_path, str_start_time + ".zip"), 'a', zipfile.ZIP_LZMA)
    # 将结果加进压缩包里
    for file_name in file_list:
        file_path = os.path.join(report_path, file_name)
        if os.path.splitext(file_path)[1] != ".xls":
            file_tar_list = []
            for dirpath, dirnames, filenames in os.walk(file_path):
                for filename in filenames:
                    file_tar_list.append(os.path.join(dirpath, filename))
            for tar in file_tar_list:
                report_zip.write(tar, tar[len(file_path)-len(file_name):])
    report_zip.close()
