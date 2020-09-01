"""Definition of the nodes of an abstract syntax tree."""

from dataclasses import dataclass, field
from abc import ABC, abstractmethod
from typing import List, Optional, Dict, Any, Union, ClassVar

from .type import Type
from .visitor import Visitable


class Statement:
    pass


##################
#  Base classes  #
##################


@dataclass
class AST(Visitable):
    """Base class of all the nodes of the tree."""

    __ast_nodes_number: ClassVar[int] = 0
    __hash: int = field(init=False, repr=False)

    file: Optional[str] = field(default=None, repr=False)
    start_line: Optional[int] = field(default=None, repr=False)
    start_column: Optional[int] = field(default=None, repr=False)
    end_line: Optional[int] = field(default=None, repr=False)
    end_column: Optional[int] = field(default=None, repr=False)

    def __post_init__(self) -> None:
        self.__hash = AST.__ast_nodes_number
        AST.__ast_nodes_number += 1

    def to_dict(self) -> Dict[str, Any]:
        """
        Convert a tree to a recursive dict, using all the public members of
        the nodes as keys. A node has a '_type' key containing its type.
        """
        d: Dict[str, Any]

        def attribute_filter(attribute: str) -> bool:
            return not attribute.startswith("_") and not callable(
                getattr(self, attribute)
            )

        d = {"_type": type(self).__name__}
        for attr_name in filter(attribute_filter, dir(self)):
            attr = getattr(self, attr_name)

            if isinstance(attr, AST):
                d[attr_name] = attr.to_dict()
            elif isinstance(attr, (list, tuple)):
                d[attr_name] = []
                for elt in attr:
                    try:
                        d[attr_name].append(elt.to_dict())
                    except AttributeError:
                        d[attr_name].append(elt)
            else:
                d[attr_name] = str(attr)

        return d

    def __hash__(self) -> int:
        return self.__hash


@dataclass
class Block(AST):
    """
    Sequence of statements, used as is to describe the main scope of a file,
    or inherited to implement block statements (loops, conditionals,
    functions definition, ...).
    """

    body: List[AST] = field(default_factory=list)


@dataclass
class Comment(AST):
    content: str = ""


############################
#  Declarative statements  #
############################


class VarDec(AST):
    """
    Single variable declaration, corresponding to a Dim or Const statement
    with a one-member variable list.
    """

    identifier: "Identifier"
    type: Optional[Union[Type, "Identifier"]]
    value: Optional["Expression"]
    new: bool

    def __init__(
        self,
        identifier: "Identifier",
        type: Optional[Union[Type, "Identifier"]] = None,
        value: Optional["Expression"] = None,
        new: bool = False,
        **kwargs,
    ) -> None:
        super().__init__(**kwargs)
        self.identifier = identifier
        self.type = type
        self.value = value
        self.new = new


# TODO implement
class MultipleVarDec(AST):
    """
    Variable declaration corresponding to a generic Dim or Const statement
    with potentially multiple variables.
    """

    declarations: List[VarDec]

    def __init__(self, declarations: List[VarDec], **kwargs) -> None:
        super().__init__(**kwargs)
        self.declarations = declarations


@dataclass
class Let(AST):
    """Let assignment. variable must be an l-value"""

    variable: "Expression" = field(default_factory="Expression")
    value: "Expression" = field(default_factory="Expression")


class VarAssign(AST):
    """Variable assignment."""

    variable: Union["Get", "Identifier"]
    value: "Expression"

    def __init__(self, variable: "Get", value: "Expression", **kwargs) -> None:
        super().__init__(**kwargs)
        self.variable = variable
        self.value = value


class FunDef(Block):
    """Function definition, corresponding to the Function keyword."""

    name: "Identifier"
    arguments: "ArgListDef"

    def __init__(
        self,
        name: "Identifier",
        arguments: Optional["ArgListDef"] = None,
        **kwargs,
    ) -> None:
        super().__init__(**kwargs)
        self.name = name
        self.arguments = arguments if arguments is not None else ArgListDef([])


class ProcDef(Block):
    """Procedure definition, corresponding to the Sub keyword."""

    name: "Identifier"
    arguments: "ArgListDef"

    def __init__(
        self,
        name: "Identifier",
        arguments: Optional["ArgListDef"] = None,
        **kwargs,
    ) -> None:
        super().__init__(**kwargs)
        self.name = name
        self.arguments = arguments if arguments is not None else ArgListDef([])


###########################
#  Executable statements  #
###########################


@dataclass
class Expression(AST):
    """
    Root of an expression tree, made of binary and unary operators, with
    literals, identifiers and function calls as leafs.
    """

    pass


@dataclass
class Identifier(Expression):
    """Identifier of a variable or function."""

    name: str = ""

    # def __init__(self, name: str, **kwargs) -> None:
    #     super().__init__(**kwargs)
    #     self.name = name


@dataclass
class MemberAccess(Expression):
    """Recursive member access (.) operation"""

    parent: Expression = field(default_factory=Expression)
    child: Identifier = field(default_factory=Identifier)


class Get(Expression):
    """Recursive node corresponding to the . operator."""

    parent: Union["Get", Identifier, "FunCall"]
    child: Identifier

    def __init__(
        self,
        parent: Union["Get", Identifier, "FunCall"],
        child: Identifier,
        **kwargs,
    ) -> None:
        super().__init__(**kwargs)
        self.parent = parent
        self.child = child


@dataclass
class Literal(Expression):
    """Literal value : integer, double, boolean, string, ..."""

    vba_type: Type = field(default_factory=Type.Integer)
    value: Union[int, float, bool, str] = field(default=0)

    @staticmethod
    def from_value(value) -> "Literal":
        return Literal(value.base_type, value.value)


@dataclass
class Arg(AST):
    """
    Argument in a list of arguments, be it named or anonymous, can also be
    empty.
    """

    value: Optional[Expression] = None
    name: Optional[str] = None
    byval: bool = False


@dataclass
class ArgList(AST):
    """Arguments list"""

    arguments: List[Arg] = field(default_factory=[])


class ArgListCall(AST):
    """List of arguments, used by function calls"""

    args: List["Expression"]

    def __init__(self, args: List["Expression"], **kwargs) -> None:
        super().__init__(**kwargs)
        self.args = args


class ArgListDef(AST):
    """List of arguments, used by function declarations"""

    args: List[Identifier]

    def __init__(self, args: List["Identifier"], **kwargs) -> None:
        super().__init__(**kwargs)
        self.args = args


class FunCall(Expression):
    """Function call"""

    function: Union[Get, Identifier]
    arguments: ArgListCall

    def __init__(
        self, function: Union[Get, Identifier], arguments: ArgListCall, **kwargs
    ) -> None:
        super().__init__(**kwargs)
        self.function = function
        self.arguments = arguments


@dataclass
class IndexExpr(Expression):
    """Ue of subscripting operator ( )"""

    expression: Expression = field(default_factory=Expression)
    arguments: ArgList = field(default_factory=ArgList)


@dataclass
class UnOp(Expression):
    """Unary operator"""

    operator: Optional[str] = None
    argument: Expression = field(default_factory=Expression)


@dataclass
class BinOp(Expression):
    """Binary operator"""

    operator: Optional[str] = None
    left: Expression = field(default_factory=Expression)
    right: Expression = field(default_factory=Expression)


##########################
#  Loop and conditional  #
##########################


@dataclass
class ElseIf(Block):
    """Single condition/action block, used internally by If"""

    condition: Expression = field(default_factory=Expression)


@dataclass
class If(Block):
    """If statement"""

    condition: Expression = field(default_factory=Expression)
    elsifs: List[ElseIf] = field(default_factory=list)
    else_block: Optional[Block] = None


@dataclass
class For(Block):
    """For statement with a counter"""

    counter: Expression = field(default_factory=Expression)
    start: Expression = field(default_factory=Expression)
    end: Expression = field(default_factory=Expression)
    step: Optional[Expression] = None
    footer_counter: Optional[Expression] = None


####################
#  Error handling  #
####################


class OnError(AST):
    """
    On Error statement. If goto is None, corresponds to a Resume Next policy,
    else to a GoTo policy.
    """

    goto: Optional[Union[Literal, Identifier]]

    def __init__(
        self, goto: Optional[Union[int, Identifier]] = None, **kwargs
    ) -> None:
        super().__init__(**kwargs)
        self.goto = goto


class Resume(AST):
    """
    Resume statement. If goto is None, corresponds to Resume or Resume Next,
    else to a GoTo form.
    """

    goto: Optional[Union[Literal, Identifier]]

    def __init__(
        self, goto: Optional[Union[int, Identifier]] = None, **kwargs
    ) -> None:
        super().__init__(**kwargs)
        self.goto = goto


class ErrorStatement(AST):
    """Error statement."""

    number: Literal

    def __init__(self, number: Literal, **kwargs) -> None:
        super().__init__(**kwargs)
        self.number = number
