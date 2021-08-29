import datetime
import typing


class InvoiceBaseData(object):
    """发票基础信息"""
    def __init__(self, code: str, number: str, billing_date: typing.Union[datetime.date, str], check_num: str):
        self.code = code
        self.number = number
        self.billing_date = billing_date
        self.check_num = check_num


class InvoiceItemsData(object):
    """发票明细行"""
    pass
