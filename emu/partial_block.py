"""Define parts of AST nodes, useful for block statements."""

from abc import abstractmethod
from typing import List

from .abstract_syntax_tree import For, ElseIf, If, Block, ProcDef, FunDef

class BlockElement:
    FirstElement = False
    LastElement = False

    @abstractmethod
    def __init__(self):
        pass


class ForHeader(BlockElement):
    FirstElement = True

    def __init__(self, counter, start, end, step=None):
        super().__init__()
        self.counter = counter
        self.start = start
        self.end = end
        self.step = step


class ForFooter(BlockElement):
    LastElement = True

    def __init__(self, counter=None):
        super().__init__()
        self.counter = counter


class IfHeader(BlockElement):
    FirstElement = True

    def __init__(self, condition):
        super().__init__()
        self.condition = condition


class ElseIfHeader(BlockElement):
    def __init__(self, condition):
        super().__init__()
        self.condition = condition


class ElseHeader(BlockElement):
    pass


class IfFooter(BlockElement):
    LastElement = True


class ProcDefHeader(BlockElement):
    FirstElement = True

    def __init__(self, name, arguments=None):
        super().__init__()
        self.name = name
        self.arguments = arguments


class ProcDefFooter(BlockElement):
    LastElement = True

class FunDefHeader(BlockElement):
    FirstElement = True

    def __init__(self, name, arguments=None):
        super().__init__()
        self.name = name
        self.arguments = arguments


class FunDefFooter(BlockElement):
    LastElement = True

class PartialBlock:
    """A PartialBlock is used in the process of building block statements,
    e.g. function definitions or conditionals. It handles statements with
    multiple subblocks, as in a if/elseif/else structure."""

    def __init__(self, elements: List[BlockElement] = None,
                 statements_blocks: List[List["Statement"]] = None):
        super().__init__()
        self.elements = elements if elements is not None else []
        self.statements_blocks = statements_blocks if statements_blocks is not None else [[]]


    # TODO add syntax check
    def build_block(self):
        header = self.elements[0]
        if isinstance(header, ForHeader) and len(self.elements) == 2:
            for_block = For(header.counter, header.start, header.end,
                            header.step, body=self.statements_blocks[0])
            return for_block

        elif isinstance(header, IfHeader):
            blocks = zip(self.statements_blocks, self.elements)
            if_body, if_header = next(blocks)
            assert(isinstance(if_header, IfHeader))
            if_block = If(condition=if_header.condition, body=if_body)

            for statements, element in blocks:
                if isinstance(element, ElseIfHeader):
                    elseif = ElseIf(condition=element.condition, body=statements)
                    if_block.elsifs.append(elseif)
                elif isinstance(element, ElseHeader):
                    else_block = Block(statements)
                    if_block.else_block = else_block
                elif isinstance(element, IfFooter):
                    break

            return if_block

        elif isinstance(header, ProcDefHeader):
            body = self.statements_blocks.pop()
            proc_def = ProcDef(header.name, header.arguments, body=body)
            return proc_def

        elif isinstance(header, FunDefHeader):
            body = self.statements_blocks.pop()
            fun_def = FunDef(header.name, header.arguments, body=body)
            return fun_def

        else:
            raise NotImplementedError()