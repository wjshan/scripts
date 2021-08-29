from openpyxl import load_workbook
from .utils import WorkbookSerializer
from lxml import etree


def excel2xml(excel_file, export_file):
    wb = load_workbook(excel_file)
    serializer = WorkbookSerializer.from_excel(wb)
    workbook_tag, style_tag = serializer.to_xml()
    with open(export_file, 'wb') as f:
        f.write(etree.tostring(workbook_tag, encoding='utf-8', pretty_print=True))
    with open(f"{export_file}.style", 'wb') as f:
        f.write(etree.tostring(style_tag, encoding='utf-8', pretty_print=True))

# if __name__ == '__main__':
#     load_workbook()
