import os
from system_helper.environment import SpecialFolder


class Environment(object):
    @staticmethod
    def get_folder_path(folder: SpecialFolder) -> str:
        value = ''
        if folder == SpecialFolder.USER_PROFILE:
            value = 'userProfile'

        return os.getenv(value)

    @staticmethod
    def new_line() -> str:
        return os.linesep
