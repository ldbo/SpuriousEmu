"""Definition of the nodes of an abstract syntax tree."""

from .type import Type

from abc import ABC, abstractmethod
from collections.abc import Iterable
from typing import List


##################
#  Base classes  #
##################

class AST(ABC):
    """Base class of all the nodes of the tree."""

    @abstractmethod
    def __init__(self):
        pass

    def to_dict(self):
        def attribute_filter(attribute):
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

    @abstractmethod
    def __init__(self, file: str = "", line_number: int = 0):
        """ Initialize a statement, saving its context. Must be called by
        child classes.

        :arg file: Name of the source file
        :arg line_number: Line number of the statement
        """
        super().__init__()
        self.file = file
        self.line_number = line_number


class Sequence(AST):
    """
    Sequence of statements, often the base node of an AST describing a file.
    """

    def __init__(self, statements: List[Statement]):
        super().__init__()
        self.statements = statements


############################
#  Declarative statements  #
############################


class VarDec(Statement):
    """
    Single variable declaration, corresponding to a Dim or Const statement
    with a one-member variable list.
    """

    def __init__(self, identifier, type=None, value=None, **kwargs):
        super().__init__(**kwargs)
        self.identifier = identifier
        self.type = type
        self.value = value


class MultipleVarDec(Statement):
    """
    Variable declaration corresponding to a generic Dim or Const statement
    with potentially multiple variables.
    """

    def __init__(self, declarations: List[VarDec], **kwargs):
        super().__init__(**kwargs)
        self.declarations = declarations


class VarAssign(Statement):
    """Variable assignment."""

    def __init__(self, variable, value: "Expression", **kwargs):
        super().__init__(**kwargs)
        self.variable = variable
        self.value = value


class FunDef(Statement):
    """Function definition, corresponding to the Function keyword."""

    def __init__(self, name, arguments, body: Sequence, **kwargs):
        super().__init__(**kwargs)
        self.name = name
        self.arguments = arguments
        self.body = body


class ProcDef(Statement):
    """Procedure definition, corresponding to the Sub keyword."""

    def __init__(self, name, arguments, body: Sequence, **kwargs):
        super().__init__(**kwargs)
        self.name = name
        self.arguments = arguments
        self.body = body

###########################
#  Executable statements  #
###########################


class Expression(Statement):
    @abstractmethod
    def __init__(self, **kwargs):
        super().__init__(**kwargs)


class Identifier(Expression):
    def __init__(self, name, **kwargs):
        super().__init__(**kwargs)
        self.name = name


class Literal(Expression):
    # TODO value smells fishy, maybe add child classes for different types or
    # add a type member
    def __init__(self, type: Type, value, **kwargs):
        super().__init__(**kwargs)
        self.value = value
        self.type = type


class ArgList(Statement):
    def __init__(self, args: List["Expression"], **kwargs):
        super().__init__(**kwargs)
        self.args = args


class FunCall(Expression):
    def __init__(self, function: Identifier, arguments: ArgList, **kwargs):
        super().__init__(**kwargs)
        self.function = function
        self.arguments = arguments


class UnOp(Expression):
    def __init__(self, operator, argument, **kwargs):
        super().__init__(**kwargs)
        self.operator = operator
        self.argument = argument


class BinOp(Expression):
    def __init__(self, operator, left, right, **kwargs):
        super().__init__(**kwargs)
        self.operator = operator
        self.left = left
        self.right = right

###########
#  Blocs  #
###########


class If(Statement):
    """If statement."""

    def __init__(self,
                 if_conditions: List[Expression], if_actions: List[Sequence],
                 else_action: Sequence,
                 **kwargs):
        super().__init__(**kwargs)
        self.if_conditions = if_conditions
        self.if_actions = if_actions
        self.else_action = else_action


class For(Statement):
    """For statement with a counter."""

    def __init__(self, counter: Identifier, start: Expression, end: Expression,
                 body: Sequence):
        super().__init__()
        self.counter = counter
        self.start = start
        self.end = end
        self.body = body
