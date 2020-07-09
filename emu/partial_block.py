"""Define parts of AST nodes, useful for block statements."""

from abc import abstractmethod


class PartialBlock:
    @abstractmethod
    def __init__(self):
        pass


class ForHeader(PartialBlock):
    def __init__(self, counter, start, end, step=None):
        super().__init__()
        self.counter = counter
        self.start = start
        self.end = end
        self.step = step


class ForFooter(PartialBlock):
    def __init__(self, counter=None):
        super().__init__()
        self.counter = counter
