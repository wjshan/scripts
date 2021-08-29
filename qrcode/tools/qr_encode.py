from .qrcode_version import Version


class EncodeBase(object):

    def __init__(self, string: str, version: Version = None):
        """

        :param string: 待编码的字符
        :param version: 默认使用版本号
        """
        self.s = string
        self.v = version

    def length_prediction(self, version: Version) -> int:
        raise NotImplementedError


class DigitalEncode(EncodeBase):
    pass
