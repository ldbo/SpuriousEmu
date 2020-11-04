from enum import Enum


class Token(str):
    class Category(Enum):
        TEXT = 1
        SPACE = 2
        END_OF_STATEMENT = 3
        END_OF_FILE = 4

    def __init__(self, value, category):
        super().__init__(value)
        self.category = category


class Lexer:
    def __init__(self, content: str) -> None:
        self.content: str = content
