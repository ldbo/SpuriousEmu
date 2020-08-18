"""Definition of the nodes of an abstract syntax tree."""

from abc import ABC, abstractmethod
from typing import List, Optional, Dict, Any, Union

from .operator import BinaryOperator
from .type import Type
from .visitor import Visitable


##################
#  Base classes  #
##################

class AST(Visitable, ABC):
    """Base class of all the nodes of the tree."""
    __ast_nodes_number: int = 0
    __hash: int

    @abstractmethod
    def __init__(self) -> None:
        self.__hash = AST.__ast_nodes_number
        AST.__ast_nodes_number += 1

    def to_dict(self) -> Dict[str, Any]:
        """
        Convert a tree to a recursive dict, using all the public members of
        the nodes as keys. A node has a '_type' key containing its type.
        """
        d: Dict[str, Any]

        def attribute_filter(attribute: str) -> bool:
            return not attribute.startswith('_') \
                and not callable(getattr(self, attribute))

        d = {'_type': type(self).__name__}
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


class Statement(AST):
    """
    Generic statement, usually corresponding to a simple instruction, or the
    declaration of e.g. a function.
    """
    file: str
    line_number: int

    @abstractmethod
    def __init__(self, file: str = "", line_number: int = 0) -> None:
        """ Initialize a statement, saving its context. Must be called by
        child classes.

        :arg file: Name of the source file
        :arg line_number: Line number of the statement
        """
        super().__init__()
        self.file = file
        self.line_number = line_number


class Block(AST):
    """
    Sequence of statements, used as is to describe the main scope of a file,
    or inherited to implement block statements (loops, conditionals,
    functions definition, ...).
    """
    body: Optional[List[Union[Statement, "Block"]]]

    def __init__(self, body: Optional[List[Union[Statement, "Block"]]] = None)\
            -> None:
        super().__init__()
        self.body = body if body is not None else []


############################
#  Declarative statements  #
############################


class VarDec(Statement):
    """
    Single variable declaration, corresponding to a Dim or Const statement
    with a one-member variable list.
    """
    identifier: "Identifier"
    type: Optional[Union[Type, "Identifier"]]
    value: Optional["Expression"]
    new: bool

    def __init__(self, identifier: "Identifier",
                 type: Optional[Union[Type, "Identifier"]] = None,
                 value: Optional["Expression"] = None,
                 new: bool = False,
                 **kwargs) -> None:
        super().__init__(**kwargs)
        self.identifier = identifier
        self.type = type
        self.value = value
        self.new = new


# TODO implement
class MultipleVarDec(Statement):
    """
    Variable declaration corresponding to a generic Dim or Const statement
    with potentially multiple variables.
    """
    declarations: List[VarDec]

    def __init__(self, declarations: List[VarDec], **kwargs) -> None:
        super().__init__(**kwargs)
        self.declarations = declarations


class VarAssign(Statement):
    """Variable assignment."""
    variable: Union["Get", "Identifier"]
    value: "Expression"

    def __init__(self, variable: "Get", value: "Expression",
                 **kwargs) -> None:
        super().__init__(**kwargs)
        self.variable = variable
        self.value = value


class FunDef(Block):
    """Function definition, corresponding to the Function keyword."""
    name: "Identifier"
    arguments: "ArgListDef"

    def __init__(self, name: "Identifier", arguments: "ArgListDef", **kwargs) \
            -> None:
        super().__init__(**kwargs)
        self.name = name
        self.arguments = arguments


class ProcDef(Block):
    """Procedure definition, corresponding to the Sub keyword."""
    name: "Identifier"
    arguments: "ArgListDef"

    def __init__(self, name: "Identifier", arguments: "ArgListDef", **kwargs) \
            -> None:
        super().__init__(**kwargs)
        self.name = name
        self.arguments = arguments


###########################
#  Executable statements  #
###########################


class Expression(Statement):
    """
    Root of an expression tree, made of binary and unary operators, with
    literals, identifiers and function calls as leafs.
    """

    @abstractmethod
    def __init__(self, **kwargs) -> None:
        super().__init__(**kwargs)


class Identifier(Expression):
    """Identifier of a variable or function."""
    name: str

    def __init__(self, name: str, **kwargs) -> None:
        super().__init__(**kwargs)
        self.name = name


class Get(Expression):
    """Recursive node corresponding to the . operator."""
    parent: Union["Get", Identifier, "FunCall"]
    child: Identifier

    def __init__(self, parent: Union["Get", Identifier, "FunCall"],
                 child: Identifier, **kwargs) -> None:
        super().__init__(**kwargs)
        self.parent = parent
        self.child = child


class Literal(Expression):
    """Literal value : integer, double, boolean, string, ..."""
    type: Type
    value: Union[int, float, bool, str]

    def __init__(self, type: Type, value: Union[int, float, bool, str],
                 **kwargs) -> None:
        super().__init__(**kwargs)
        self.value = value
        self.type = type

    @staticmethod
    def from_value(value) -> "Literal":
        return Literal(value.base_type, value.value)


class ArgListCall(Statement):
    """List of arguments, used by function calls"""
    args: List["Expression"]

    def __init__(self, args: List["Expression"], **kwargs) -> None:
        super().__init__(**kwargs)
        self.args = args


class ArgListDef(Statement):
    """List of arguments, used by function declarations"""
    args: List[Identifier]

    def __init__(self, args: List["Identifier"], **kwargs) -> None:
        super().__init__(**kwargs)
        self.args = args


class FunCall(Expression):
    """Function call"""
    function: Union[Get, Identifier]
    arguments: ArgListCall

    def __init__(self, function: Union[Get, Identifier],
                 arguments: ArgListCall, **kwargs) -> None:
        super().__init__(**kwargs)
        self.function = function
        self.arguments = arguments


class UnOp(Expression):
    """Unary operator"""
    operator: str
    argument: Expression

    def __init__(self, operator: str, argument: Expression, **kwargs) -> None:
        super().__init__(**kwargs)
        self.operator = operator
        self.argument = argument


class BinOp(Expression):
    """Binary operator"""
    operator: BinaryOperator
    left: Expression
    right: Expression

    def __init__(self, operator: BinaryOperator, left: Expression,
                 right: Expression, **kwargs) -> None:
        super().__init__(**kwargs)
        self.operator = operator
        self.left = left
        self.right = right


##########################
#  Loop and conditional  #
##########################

class ElseIf(Block):
    """Single condition/action block, used internally by If"""
    condition: Expression

    def __init__(self, condition: Expression, **kwargs) -> None:
        super().__init__(**kwargs)
        self.condition = condition


class If(Block):
    """If statement"""
    condition: Expression
    elsifs: List[ElseIf]
    else_block: Optional[Block]

    def __init__(self, condition: Expression,
                 elsifs: Optional[List[ElseIf]] = None,
                 else_block: Optional[Block] = None, **kwargs) -> None:
        super().__init__(**kwargs)
        self.condition = condition
        self.elsifs = elsifs if elsifs is not None else []
        self.else_block = else_block


class For(Block):
    """For statement with a counter"""
    counter: Identifier
    start: Expression
    end: Expression
    step: Optional[Expression]

    def __init__(self, counter: Identifier, start: Expression, end: Expression,
                 step: Optional[Expression] = None, **kwargs) -> None:
        super().__init__(**kwargs)
        self.counter = counter
        self.start = start
        self.end = end
        self.step = step

####################
#  Error handling  #
####################


class OnError(Statement):
    """
    On Error statement. If goto is None, corresponds to a Resume Next policy,
    else to a GoTo policy.
    """
    goto: Optional[Union[Literal, Identifier]]

    def __init__(self, goto: Optional[Union[int, Identifier]] = None,
                 **kwargs) -> None:
        super().__init__(**kwargs)
        self.goto = goto


class Resume(Statement):
    """
    Resume statement. If goto is None, corresponds to Resume or Resume Next,
    else to a GoTo form.
    """
    goto: Optional[Union[Literal, Identifier]]

    def __init__(self, goto: Optional[Union[int, Identifier]] = None,
                 **kwargs) -> None:
        super().__init__(**kwargs)
        self.goto = goto


class ErrorStatement(Statement):
    """Error statement."""
    number: Literal

    def __init__(self, number: Literal, **kwargs) -> None:
        super().__init__(**kwargs)
        self.number = number
