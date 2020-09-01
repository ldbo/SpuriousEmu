from functools import wraps
from inspect import getmembers, isfunction
from pathlib import Path
from pkg_resources import resource_filename
from typing import Any, Callable
from typing import Type as tType

from lark import Lark, Token, Transformer, Tree
from lark.exceptions import UnexpectedInput
from lark.tree import Meta
from lark.visitors import v_args

from .abstract_syntax_tree import *
from .error import ParsingError
from .type import Type

Callback = Callable[[Transformer, Tree, Meta], AST]


def propagate_positions_function(callback: Callback) -> Callback:
    """Function decorator propagating the meta position of a Tree to an AST"""

    @wraps(callback)
    def decorated(self, tree: Tree, meta: Meta) -> AST:
        ast = callback(self, tree, meta)

        try:
            ast.start_line = meta.line
            ast.start_column = meta.column
            ast.end_line = meta.end_line
            ast.end_column = meta.end_column
            ast.file = self.file
        except AttributeError:
            pass

        return ast

    return decorated


def propagate_positions_class(
    transformer: tType[Transformer],
) -> tType[Transformer]:
    """Class decorator using propagate_positions_function on rule callbacks"""
    for name, value in getmembers(transformer, predicate=isfunction):
        if name.islower() and not name.startswith("_") and name != "transform":
            propagating_method = propagate_positions_function(value)
            setattr(transformer, name, propagating_method)

    return transformer


def token_type_is(token: Token, type: str) -> bool:
    """Return True if the type of token is type"""
    return isinstance(token, Token) and token.type == type


@v_args(meta=True)
@propagate_positions_class
class TreeToAST(Transformer):
    """Transformer used to build an AST from a Tree"""

    file: Optional[str]

    def __init__(self, file) -> None:
        super().__init__()
        self.file = Path(file).name

    # Lexical tokens
    def COMMENT(self, token) -> Comment:
        if token.startswith("'"):
            prefix_length = 1
        elif token.startswith("REM"):
            prefix_length = 3
        else:
            raise NotImplemented(f"Unrecognized comment {token}")

        return token.update(value=token[prefix_length + 1 :])

    def eol(self, tree, meta) -> Union[List[Any], Comment]:
        if len(tree) == 0:
            return tree
        elif tree[0].type == "COMMENT":
            return Comment(content=tree[0])
        else:
            raise NotImplemented(f"eol {tree}")

    # Statements
    def variable_declaration(self, tree, meta) -> VarDec:
        tree.reverse()
        new, declared_type, value = False, None, None
        try:
            identifier = tree.pop()

            if token_type_is(tree[-1], "AS"):
                tree.pop()
                if token_type_is(tree[-1], "NEW"):
                    new = True
                    tree.pop()

                declared_type = tree.pop()

            if token_type_is(tree[-1], "EQ"):
                tree.pop()
                value = tree.pop()
        except IndexError:
            pass

        return VarDec(identifier, declared_type, value, new)

    def let_statement(self, tree, meta) -> Let:
        return Let(variable=tree[0], value=tree[1])

    # Control
    def for_statement(self, tree, meta) -> For:
        assert len(tree) == 4
        header, eol, block, footer = tree
        assert len(header.children) in (3, 4)
        assert len(footer.children) in (0, 1)

        counter, start, end = header.children[:3]
        step = header.children[3] if len(header.children) == 4 else None
        footer_counter = (
            footer.children[0] if len(footer.children) == 1 else None
        )

        return For(
            counter=counter,
            start=start,
            end=end,
            step=step,
            footer_counter=footer_counter,
            body=block.body,
        )

    def if_statement(self, tree, meta) -> If:
        header, eol, block = tree[:3]
        subblocks = tree[3:-1]

        condition = header.children[0]
        body = block.body

        elseifs = []
        else_block = None
        for position, subblock in enumerate(subblocks):
            if isinstance(subblock, ElseIf):
                elseifs.append(subblock)
            elif subblock.data == "else_block":
                else_block = subblock.children[1]

                if position != len(subblocks) - 1:
                    raise NotImplementedError(f"if_statement {tree}")

                break


        return If(
            condition=condition,
            body=body,
            elsifs=elseifs,
            else_block=else_block,
        )

    def elseif_block(self, tree, meta) -> ElseIf:
        condition = tree[0]

        if len(tree) == 3:
            block = tree[2].body
        elif len(tree) == 2:
            block = [tree[1]]
        else:
            raise NotImplementedError(f"elseif_block {tree}")

        return ElseIf(body=block, condition=condition)

    def single_line_if_statement(self, tree, meta) -> If:
        return tree[0]

    def if_with_non_empty_then(self, tree, meta) -> If:
        condition = tree[0]
        body = [tree[1]]
        else_block = tree[2].children if len(tree) == 3 else None

        return If(condition=condition, body=body, else_block=else_block)

    def if_with_empty_then(self, tree, meta) -> If:
        condition = tree[0]
        else_block = tree[1].children

        return If(condition=condition, else_block=else_block)

    # Expression
    def transform_binop(self, tree, meta) -> BinOp:
        children = tree.children
        return BinOp(
            operator=children[1].value,
            left=self.transform_expression(children[0], meta),
            right=self.transform_expression(children[2], meta),
        )

    def transform_unop(self, tree, meta) -> UnOp:
        children = tree.children
        return UnOp(
            operator=children[0].value,
            argument=self.transform_expression(children[1], meta),
        )

    def transform_expression(self, tree, meta) -> Expression:

        parts = tree.children
        if len(parts) == 1:
            if isinstance(parts[0], Tree):
                return self.transform_expression(parts[0], meta)
            else:
                return parts[0]
        elif len(parts) == 2:
            return self.transform_unop(tree, meta)
        elif len(parts) == 3:
            return self.transform_binop(tree, meta)
        else:
            raise NotImplementedError(f"Don't support expression {tree}")

    def expression(self, tree, meta) -> Expression:
        return self.transform_expression(tree[0], meta)

    # Literals
    def INTEGER(self, token) -> Token:
        return token.update(value=int(token))

    def BOOLEAN(self, token) -> Token:
        if token.value == "True":
            return token.update(value=True)
        elif token.value == "False":
            return token.update(value=False)
        else:
            raise NotADirectoryError(f"Unsupported boolean {token}")

    def STRING(self, token) -> Token:
        return token.update(value=token[1:-1].replace('""', '"'))

    def literal(self, tree, meta) -> Literal:
        token = tree[0]
        if token.type == "STRING":
            vba_type = Type.String
        elif token.type == "INTEGER":
            vba_type = Type.Integer
        elif token.type == "BOOLEAN":
            vba_type = Type.Boolean
        else:
            raise NotImplementedError(
                f"Literal type {token.type} not supported"
            )

        return Literal(vba_type=vba_type, value=token.value)

    # Left expression
    def simple_name_expression(self, tree, meta) -> Identifier:
        return Identifier(name=tree[0].value)

    def member_access_expression(self, tree, meta) -> MemberAccess:
        return MemberAccess(parent=tree[0], child=tree[1])

    def index_expression(self, tree, meta) -> IndexExpr:
        return IndexExpr(expression=tree[0], arguments=tree[1])

    # Argument list
    def argument_list(self, tree, meta) -> ArgList:
        if len(tree) == 1 and isinstance(tree[0], ArgList):
            return tree[0]
        else:
            return ArgList(arguments=tree)

    def positional_argument_list(self, tree, meta) -> ArgList:
        return ArgList(arguments=tree)

    def mixed_argument_list(self, tree, meta) -> ArgList:
        return ArgList(arguments=tree)

    def named_argument_list(self, tree, meta) -> ArgList:
        return ArgList(arguments=tree)

    def optional_positional_argument(self, tree, meta) -> Optional[Arg]:
        if len(tree) == 0:
            return Arg()
        elif len(tree) == 1:
            return tree[0]
        else:
            raise NotImplementedError(f"optional_positional_argument {tree}")

    def named_argument(self, tree, meta) -> Arg:
        arg = tree[1]
        arg.name = tree[0]
        return arg

    def argument_expression(self, tree, meta) -> Arg:
        if len(tree) == 1:
            return Arg(value=tree[0])
        elif len(tree) == 2:
            return Arg(value=tree[1], byval=True)
        else:
            raise NotImplementedError(f"Argument {tree} not supported")

    def statement_block(self, tree, meta) -> Block:
        """Remove _NEWLINE"""
        return Block(body=list(filter(lambda t: t != [], tree)))


class Parser:
    def __init__(self) -> None:
        """Load the grammar from 'grammar.lark'."""
        grammar_path = resource_filename("emu", "grammar.lark")
        with open(grammar_path) as f:
            grammar = f.read()

        self.parser = Lark(
            grammar,
            parser="lalr",
            start="program",
            propagate_positions=True,
        )

    def parse(self, content: str, filename: Optional[str] = None) -> AST:
        """
        Parse the content of a file into an abstract syntax tree.

        :arg content: str content of the file
        :arg file_name: Optional name of the file
        :return: An AST representing the syntax of the file
        """
        try:
            tree = self.parser.parse(content)
            # return tree
            ast = TreeToAST(filename).transform(tree)
            return ast
        except UnexpectedInput as e:
            msg = str(e) + "\n" + e.get_context(content)
            raise ParsingError(filename, e.line, e.column, msg)

    def parse_file(self, filename: str) -> AST:
        """
        Parse a file into an abstract syntax tree.

        :arg path: Path of the file
        :return: An AST representing the syntax of the file
        """
        with open(filename) as f:
            content = f.read()

        return self.parse(content, filename)
