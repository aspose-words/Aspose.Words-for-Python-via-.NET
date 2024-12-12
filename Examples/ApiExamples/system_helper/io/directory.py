import pathlib
from typing import NoReturn
from system_helper.io.search_option import SearchOption


class Directory(object):
    @staticmethod
    def create_directory(path: str) -> NoReturn:
        pathlib.Path(path).mkdir(parents=True, exist_ok=True)

    @staticmethod
    def get_files(path: str, search_pattern: str, search_option: SearchOption) -> [str]:
        value = ''
        if search_option == SearchOption.All_DIRECTORIES:
            value = f'**/{search_pattern}'

        return list(pathlib.Path(path).glob(value))
