import pathlib
import codecs
import os


class File(object):
    @staticmethod
    def exist(path: str) -> bool:
        return pathlib.Path(path).exists()

    @staticmethod
    def read_all_bytes(path: str) -> bytes:
        return pathlib.Path(path).read_bytes()

    @staticmethod
    def read_all_text(path: str) -> str:
        encoding = File.detect_by_bom(path, 'utf-8')
        with open(file=path, mode='r', encoding=encoding, newline=os.linesep) as f:
            return f.read()

    @staticmethod
    def write_all_bytes(path: str, bytes_: bytes):
        pathlib.Path(path).write_bytes(bytes_)

    @staticmethod
    def detect_by_bom(path: str, default: str):
        with open(path, 'rb') as f:
            raw = f.read(4)  # will read less if the file is smaller
        # BOM_UTF32_LE's start is equal to BOM_UTF16_LE so need to try the former first
        for enc, boms in \
                ('utf-8-sig', (codecs.BOM_UTF8,)), \
                ('utf-32', (codecs.BOM_UTF32_LE, codecs.BOM_UTF32_BE)), \
                ('utf-16', (codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
            if any(raw.startswith(bom) for bom in boms):
                return enc
        return default
