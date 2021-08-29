# -*-coding:utf-8 -*-
# @author: DanDan
# @contact: 454273687@qq.com
# @file: xml_excel_test.py
# @time: 2021/8/29 22:48
# @desc:
from xml_excel.excel2xml import excel2xml
from xml_excel.xml2excel import xml2excel
if __name__ == '__main__':
    excel2xml('季度报表-专业分包统计表(1).xlsx','text.xml')

    xml2excel('text.xml','季度报表-专业分包统计表-导出.xlsx')
