"""Define parts of AST nodes, useful for block statements."""

from abc import abstractmethod
from typing import List

from .abstract_syntax_tree import For, Sequence


class BlockElement:
    FirstElement = False
    LastElement = False

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
    def __init__(self, elements: List[BlockElement] = None,
                 statements: List["Statement"] = None):
        super().__init__()
        self.elements = elements if elements is not None else []
        self.statements = statements if statements is not None else []

    def build_block(self):
        pass
