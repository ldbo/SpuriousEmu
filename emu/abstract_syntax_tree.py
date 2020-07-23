"""Definition of the nodes of an abstract syntax tree."""

from abc import ABC, abstractmethod
from typing import List, Optional, Dict, Any, Union

from .operator import BinaryOperator
from .type import Type


##################
#  Base classes  #
##################

class AST(ABC):
    """Base class of all the nodes of the tree."""

    @abstractmethod
    def __init__(self) -> None:
        pass

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
    type: Optional[Type]
    value: Optional["Expression"]

    def __init__(self, identifier: "Identifier", type: Optional[Type] = None,
                 value: Optional["Expression"] = None, **kwargs) -> None:
        super().__init__(**kwargs)
        self.identifier = identifier
        self.type = type
        self.value = value


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
    variable: "Identifier"
    value: "Expression"

    def __init__(self, variable: "Identifier", value: "Expression", **kwargs) \
            -> None:
        super().__init__(**kwargs)
        self.variable = variable
        self.value = value


class FunDef(Block):
    """Function definition, corresponding to the Function keyword."""
    name: "Identifier"
    arguments: "ArgList"

    def __init__(self, name: "Identifier", arguments: "ArgList", **kwargs) \
            -> None:
        super().__init__(**kwargs)
        self.name = name
        self.arguments = arguments


class ProcDef(Block):
    """Procedure definition, corresponding to the Sub keyword."""
    name: "Identifier"
    arguments: "ArgList"

    def __init__(self, name: "Identifier", arguments: "ArgList", **kwargs) \
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


class Literal(Expression):
    """Literal value : integer, double, boolean, string, ..."""
    type: Type
    value: Union[int, float, bool, str]

    def __init__(self, type: Type, value: Union[int, float, bool, str],
                 **kwargs) -> None:
        super().__init__(**kwargs)
        self.value = value
        self.type = type


# TODO differentiate between ArgList for calls and definitions
class ArgList(Statement):
    """List of arguments, used by function calls and declarations"""
    args: List["Expression"]

    def __init__(self, args: List["Expression"], **kwargs) -> None:
        super().__init__(**kwargs)
        self.args = args


class FunCall(Expression):
    """Function call"""
    function: Identifier
    arguments: ArgList

    def __init__(self, function: Identifier, arguments: ArgList, **kwargs) \
            -> None:
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


###########
#  Blocs  #
###########

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
