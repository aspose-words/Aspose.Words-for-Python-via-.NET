from collections.abc import Callable
from typing import Any


class Enumerable(object):

    @staticmethod
    def of_type(cast_as: Callable, value: Any) -> Any:
        try:
            return cast_as(value)
        except Exception:
            return None
