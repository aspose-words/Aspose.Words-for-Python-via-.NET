import pathlib
from typing import IO
from system_helper.io.file_mode import FileMode
from system_helper.io.file_access import FileAccess
import _io


class FileStream(_io.FileIO):
    def __init__(self, file_path: str, file_mode: FileMode, file_access: FileAccess = None):

        if file_mode == FileMode.OPEN and (file_access is None or file_access == FileAccess.READ):
            super().__init__(file_path, 'rb')
            return

        if file_mode == FileMode.OPEN_OR_CREATE and file_access is None:
            super().__init__(file_path, 'w+b')
            return

        if file_mode == FileMode.OPEN and file_access == FileAccess.WRITE:
            super().__init__(file_path, 'ab')
            return
        if file_mode == FileMode.OPEN and file_access == FileAccess.READ_WRITE:
            super().__init__(file_path, 'r+b')
            return

        if file_mode == FileMode.CREATE and file_access == FileAccess.READ_WRITE:
            super().__init__(file_path, 'w+b')
            return

        if file_mode == FileMode.CREATE and file_access is None:
            super().__init__(file_path, 'w+b')
            return

        raise NotImplementedError(f"{file_mode} and {file_access} not supported")

    def __enter__(self):
        return super(_io.FileIO, self).__enter__()

    def __exit__(self, exc_type, exc_val, exc_tb):
        super(_io.FileIO, self).__exit__(exc_type, exc_val, exc_tb)
