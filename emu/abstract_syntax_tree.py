from dataclasses import InitVar, dataclass, field
from enum import Enum
from typing import Optional, Tuple, Union

from .data import Variable
from .lexer import Token
from .utils import FilePosition, Visitable


@dataclass(frozen=True)
class AST(Visitable):
    """
    Base class of the nodes of the abstract syntax tree. It contains most of the
    static semantic of VBA, but doesn't necessarily lexicographically represent
    it. For example, whether a for loop ends with ``Next`` or ``Next i`` is not
    stored in it. However, this can be retrieved using the :attr:`position`
    field.

    It conforms to the
    Visitable interface, which is the main way it is processed. Thus, this is a
    pure data class, which does not perform any processing.

    Args:
      token_ast_or_position: Token or AST to extract the position of the node
                             from, or directly the expected position.
    """

    token_ast_or_position: InitVar[Union[Token, FilePosition, "AST"]]
    position: FilePosition = field(init=False, repr=False)
    """Position of the node in the source code, can be a real or a virtual
    position. A virtual one means that the node was added during the
    analysis."""

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


# Expressions


@dataclass(frozen=True)
class Expression(AST):
    pass


@dataclass(frozen=True)
class ParenExpr(Expression):
    expr: Expression


@dataclass(frozen=True)
class Literal(Expression):
    variable: Variable


@dataclass(frozen=True)
class Name(Expression):
    name: str


@dataclass(frozen=True)
class DictAccess(Expression):
    expr: Expression
    field: Name


@dataclass(frozen=True)
class MemberAccess(Expression):
    expr: Expression
    field: Name


@dataclass(frozen=True)
class WithDictAccess(Expression):
    field: Name


@dataclass(frozen=True)
class WithMemberAccess(Expression):
    field: Name


@dataclass(frozen=True)
class Arg(AST):
    expr: Expression
    byval: bool = False
    addressof: bool = False


@dataclass(frozen=True)
class ArgList(Expression):
    args: Tuple[Arg, ...]


@dataclass(frozen=True)
class IndexExpr(Expression):
    expr: Expression
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
class Operation(Expression):
    operator: Operator
    operands: Tuple[Expression, ...]


# Statements


@dataclass(frozen=True)
class Statement(AST):
    pass


StmtLabel = Union[Name, Literal]
"""Type of a statement label, either a name or an integer literal"""


# Error statements


@dataclass(frozen=True)
class OnError(Statement):
    goto_label: Optional[StmtLabel] = None
    """If ``None``, represent a ``On Error Resume Next`` statement, else a
    ``On Error Goto ...``"""


@dataclass(frozen=True)
class Resume(Statement):
    stmt_label: Optional[StmtLabel]
    """If ``None``, represent a ``Resume Next`` statement, else a
    ``Resume {label}``"""


@dataclass(frozen=True)
class Error(Statement):
    error_number: Expression  #: Integer expression of the error


# File statements


@dataclass(frozen=True)
class Open(Statement):
    class Mode(Enum):
        APPEND = 0
        BINARY = 1
        INPUT = 2
        OUTPUT = 3
        RANDOM = 4

    class Access(Enum):
        READ = 0
        WRITE = 1
        READ_WRITE = 2

    class Lock(Enum):
        SHARED = 1
        LOCK_READ = 2
        LOCK_WRITE = 3
        LOCK_READ_WRITE = 4

    path_name: Expression
    file_number: Expression
    mode: "Open.Mode"
    access: "Open.Access"
    lock: "Open.Lock"
    length: Expression


@dataclass(frozen=True)
class Close(Statement):
    file_numbers: Optional[Tuple[Expression, ...]]
    """If ``None, corresponds to a Reset, else a Close [file-number-list]"""


@dataclass(frozen=True)
class Seek(Statement):
    file_number: Expression
    seek_position: Expression


@dataclass(frozen=True)
class Lock(Statement):
    file_number: Expression
    start_record_number: Expression
    end_record_number: Optional[Expression]


@dataclass(frozen=True)
class Unlock(Statement):
    file_number: Expression
    start_record_number: Expression
    end_record_number: Optional[Expression]


@dataclass(frozen=True)
class LineInput(Statement):
    file_number: Expression
    variable_name: Expression


@dataclass(frozen=True)
class Width(Statement):
    file_number: Expression
    line_width: Expression


@dataclass(frozen=True)
class OutputItem(AST):
    class Clause(Enum):
        SPC = 0
        TAB = 1

    class CharPosition(Enum):
        SEMICOLON = 0
        COMMA = 1

    clause_title: Optional["OutputItem.Clause"]
    clause_argument: Optional[Expression]
    char_position: "OutputItem.CharPosition"


@dataclass(frozen=True)
class Print(Statement):
    file_number: Expression
    outputs: Tuple[OutputItem, ...]


@dataclass(frozen=True)
class Write(Statement):
    file_number: Expression
    outputs: Tuple[OutputItem, ...]


@dataclass(frozen=True)
class Input(Statement):
    file_number: Expression
    inputs: Tuple[Expression, ...]


@dataclass(frozen=True)
class Put(Statement):
    file_number: Expression
    record_number: Optional[Expression]
    data: Expression


@dataclass(frozen=True)
class Get(Statement):
    file_number: Expression
    record_number: Optional[Expression]
    variable: Optional[Expression]


# Module level elements


@dataclass(frozen=True)
class StatementBlock(AST):
    statements: Tuple[Statement, ...]
