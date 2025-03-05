from typing import Any


class Encoding(object):
    @staticmethod
    def unicode() -> str:
        # In .Net Encoding.Unicode set UTF16LE encoding.
        # However, AW does not support utf-16-le encoding, seems that is a restriction of the wrapper.
        # In this case we will return utf-16
        return "utf-16"

    @staticmethod
    def utf_8() -> str:
        return "utf-8"

    @staticmethod
    def ascii() -> str:
        return "ascii"

    @staticmethod
    def get_bytes(value: Any, encoding: str) -> bytes:
        return bytes(value, encoding=encoding)

    @staticmethod
    def get_string(value: Any, encoding: str) -> str:
        return bytes(value).decode(encoding=encoding, errors='strict')
