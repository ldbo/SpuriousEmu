from dataclasses import dataclass, field, InitVar
from enum import Enum
from typing import Tuple, Union

from .data import Variable
from .lexer import Token
from .utils import Visitable, FilePosition


@dataclass(frozen=True)
class AST(Visitable):
    """
    Base class of the nodes of the abstract syntax tree. It conforms to the
    Visitable interface, which is the main way it is processed. Thus, this is a
    pure data class, which does not perform any processing.

    Args:
      token_ast_or_position: Token or AST to extract the position of the node
                             from, or directly the expected position.
    """

    token_ast_or_position: InitVar[Union[Token, FilePosition, "AST"]]
    position: FilePosition = field(init=False, repr=False)
    """Position of the node in the source code"""

    def __post_init__(
        self, token_ast_or_position: Union[Token, FilePosition]
    ) -> None:
        if isinstance(token_ast_or_position, FilePosition):
            object.__setattr__(self, "position", token_ast_or_position)
        elif isinstance(token_ast_or_position, (Token, AST)):
            object.__setattr__(self, "position", token_ast_or_position.position)
        else:
            msg = f"Can't initialize AST with {type(token_ast_or_position)}"
            raise TypeError(msg)


@dataclass(frozen=True)
class Block(AST):
    stmts: Tuple[AST]


# Primary expressions


@dataclass(frozen=True)
class Expr(AST):
    pass


@dataclass(frozen=True)
class ParenExpr(Expr):
    expr: Expr


@dataclass(frozen=True)
class Literal(Expr):
    variable: Variable


@dataclass(frozen=True)
class Name(Expr):
    name: str


@dataclass(frozen=True)
class DictAccess(Expr):
    expr: Expr
    field: Name


@dataclass(frozen=True)
class MemberAccess(Expr):
    expr: Expr
    field: Name


@dataclass(frozen=True)
class WithDictAccess(Expr):
    field: Name


@dataclass(frozen=True)
class WithMemberAccess(Expr):
    field: Name


@dataclass(frozen=True)
class Arg(AST):
    expr: Expr
    byval: bool = False
    addressof: bool = False


@dataclass(frozen=True)
class ArgList(Expr):
    args: Tuple[Arg, ...]


@dataclass(frozen=True)
class IndexExpr(Expr):
    expr: Expr
    args: ArgList


@dataclass(frozen=True)
class Operator(AST):
    class Type(Enum):
        LPAREN: str = "LPAREN"
        RPAREN: str = "RPAREN"
        IMP: str = "IMP"
        EQV: str = "EQV"
        XOR: str = "XOR"
        OR: str = "OR"
        AND: str = "AND"
        NOT: str = "NOT"
        IS: str = "IS"
        LIKE: str = "LIKE"
        GE: str = "GE"
        GT: str = "GT"
        LE: str = "LE"
        LT: str = "LT"
        EQ: str = "EQ"
        NEQ: str = "NEQ"
        CONCAT: str = "CONCAT"
        MINUS: str = "MINUS"
        PLUS: str = "PLUS"
        MOD: str = "MOD"
        INT_DIV: str = "INT_DIV"
        DIV: str = "DIV"
        MULT: str = "MULT"
        UNARY_MINUS: str = "UNARY_MINUS"
        EXP: str = "EXP"
        DOT: str = "DOT"
        EXCLAMATION: str = "EXCLAMATION"
        COMMA: str = "COMMA"
        INDEX: str = "INDEX"

    op: "Operator.Type"


@dataclass(frozen=True)
class Operation(Expr):
    operator: Operator
    operands: Tuple[Expr, ...]
