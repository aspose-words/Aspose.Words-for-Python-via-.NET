from enum import Enum


class FileAccess(Enum):
    READ = 1
    WRITE = 2
    READ_WRITE = 4

    def __str__(self):
        return self.name
