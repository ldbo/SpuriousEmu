"""Define parts of AST nodes, useful for block statements."""

from abc import abstractmethod
from typing import List


class BlockElement:
    @abstractmethod
    def __init__(self):
        pass


class ForHeader(BlockElement):
    def __init__(self, counter, start, end, step=None):
        super().__init__()
        self.counter = counter
        self.start = start
        self.end = end
        self.step = step


class ForFooter(BlockElement):
    def __init__(self, counter=None):
        super().__init__()
        self.counter = counter


class PartialBlock:
    def __init__(self, elements: List[BlockElement] = [],
                 statements: List["Statement"] = []):
        super().__init__()
        self.elements = elements
        self.statements = statements
