# -*- coding: utf-8 -*-
# @Author : dandan
# @Email  : 454273687@qq.com
# @Desc   :
from ..units.Arguments import ArgumentData
import typing
from ..units.exceptions import QRCodeVersionError


class Version(ArgumentData):
    """版本号允许从1到40 版本越大则生成的二维码越大"""

    def __init__(self, version: int = 1):
        self.version = version

    def check_error(self) -> typing.List[QRCodeVersionError]:
        errors: typing.List[QRCodeVersionError] = []
        if not 1 <= self.version <= 40:
            errors.append(QRCodeVersionError('QR Code version mast be in range [1,40]'))
        return errors

    @property
    def correct_version(self) -> int:
        """符合规则的版本号"""
        return min([40, max([1, self.version])])

    @property
    def qr_size(self) -> int:
        """当前版本号下的二维码尺寸"""
        return (self.version - 1) * 4 + 21


