import pathlib


class FileInfo(object):
    def __init__(self, path: str):
        self._path = path

    def length(self) -> int:
        return pathlib.Path(self._path).stat().st_size
