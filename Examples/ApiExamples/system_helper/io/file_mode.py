from enum import Enum


class FileMode(Enum):
    CREATE = 2
    OPEN = 3
    OPEN_OR_CREATE = 4

    def __str__(self):
        return self.name
