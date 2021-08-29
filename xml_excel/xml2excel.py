from .utils import WorkbookSerializer
from lxml import etree


def xml2excel(xml_file, save_file):
    work_book_tag = etree.parse(xml_file)
    style_tag = etree.parse(f"{xml_file}.style")
    serializer = WorkbookSerializer.from_xml(work_book_tag, style_node=style_tag)
    wb = serializer.to_excel()
    wb.save(save_file)
