"""Define parts of AST nodes, useful for block statements."""

from abc import abstractmethod, ABC
from typing import List, Union

from .abstract_syntax_tree import (For, ElseIf, If, Block, ProcDef, FunDef,
                                   Statement)


class BlockElement(ABC):
    """
    Structural element of a block, for example an ElseIf line in an If block,
    only used during the parsing process.
    """
    IsHeader: bool = False
    IsFooter: bool = False

    @abstractmethod
    def __init__(self) -> None:
        pass


class ForHeader(BlockElement):
    IsHeader = True

    def __init__(self, counter, start, end, step=None):
        super().__init__()
        self.counter = counter
        self.start = start
        self.end = end
        self.step = step


class ForFooter(BlockElement):
    IsFooter = True

    def __init__(self, counter=None):
        super().__init__()
        self.counter = counter


class IfHeader(BlockElement):
    IsHeader = True

    def __init__(self, condition):
        super().__init__()
        self.condition = condition


class ElseIfHeader(BlockElement):
    def __init__(self, condition):
        super().__init__()
        self.condition = condition


class ElseHeader(BlockElement):
    def __init__(self):
        super().__init__()


class IfFooter(BlockElement):
    IsFooter = True

    def __init__(self):
        super().__init__()


class ProcDefHeader(BlockElement):
    IsHeader = True

    def __init__(self, name, arguments=None):
        super().__init__()
        self.name = name
        self.arguments = arguments


class ProcDefFooter(BlockElement):
    IsFooter = True

    def __init__(self):
        super().__init__()


class FunDefHeader(BlockElement):
    IsHeader = True

    def __init__(self, name, arguments=None):
        super().__init__()
        self.name = name
        self.arguments = arguments


class FunDefFooter(BlockElement):
    IsFooter = True

    def __init__(self):
        super().__init__()


class PartialBlock:
    """
    A PartialBlock is used in the process of building block statements,
    e.g. function definitions or conditionals. It handles statements with
    multiple sub-blocks, as in a if/elseif/else structure.
    """
    elements: List[BlockElement]
    statements_blocks: List[List[Union[Statement, Block]]]

    def __init__(self, elements: List[BlockElement] = None,
                 statements_blocks: List[List[Statement]] = None) -> None:
        super().__init__()
        self.elements = elements if elements is not None else []
        self.statements_blocks = statements_blocks \
            if statements_blocks is not None \
            else [[]]  # type: ignore[assignment]

    # TODO add syntax check
    def build_block(self) -> Block:
        """
        Build the Block obtained after pulling together all the BlockElement
        and Statements it is made of.
        """
        header = self.elements[0]
        if isinstance(header, ForHeader) and len(self.elements) == 2:
            for_block = For(header.counter, header.start, header.end,
                            header.step, body=self.statements_blocks[0])
            return for_block

        elif isinstance(header, IfHeader):
            blocks = zip(self.statements_blocks, self.elements)
            if_body, if_header = next(blocks)
            assert (isinstance(if_header, IfHeader))
            if_block = If(condition=if_header.condition, body=if_body)

            for statements, element in blocks:
                if isinstance(element, ElseIfHeader):
                    elseif = ElseIf(
                        condition=element.condition, body=statements)
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
