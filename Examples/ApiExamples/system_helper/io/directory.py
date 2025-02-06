import pathlib
import shutil
from typing import NoReturn
from system_helper.io.search_option import SearchOption


class Directory(object):
    @staticmethod
    def create_directory(path: str) -> NoReturn:
        pathlib.Path(path).mkdir(parents=True, exist_ok=True)

    @staticmethod
    def get_files(path: str, search_pattern: str = '*.*', search_option: SearchOption = SearchOption.TOP_All_DIRECTORY_ONLY) -> [str]:
        value = search_pattern
        if search_option == SearchOption.All_DIRECTORIES:
            value = f'**/{search_pattern}'
        return list(map(lambda a: str(a), pathlib.Path(path).glob(value)))

    @staticmethod
    def delete(path: str, recursive: bool = False) -> NoReturn:
        if not recursive:
            pathlib.Path(path).rmdir()
            return

        shutil.rmtree(path)

    @staticmethod
    def exists(path: str) -> bool:
        return pathlib.Path(path).exists()
