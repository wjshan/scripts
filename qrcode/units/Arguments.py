# -*- coding: utf-8 -*-
# @Author : dandan
# @Email  : 454273687@qq.com
# @Desc   :
from .exceptions import CodeException
import typing


class ArgumentData(object):

    def check_error(self) -> typing.List[CodeException]:
        """检查数据异常"""
        return []
