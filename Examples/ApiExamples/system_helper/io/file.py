import pathlib


class File(object):

    @staticmethod
    def exist(path: str) -> bool:
        return pathlib.Path(path).exists()

    @staticmethod
    def read_all_bytes(path: str) -> bytes:
        return pathlib.Path(path).read_bytes()

    @staticmethod
    def read_all_text(path: str) -> str:
        return pathlib.Path(path).read_text()

    @staticmethod
    def write_all_bytes(path: str, bytes_: bytes):
        pathlib.Path(path).write_bytes(bytes_)
