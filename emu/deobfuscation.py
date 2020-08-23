"""Deobfuscation tools."""

from enum import Enum
from typing import Any, Dict

from .abstract_syntax_tree import *
from .compiler import Program
from .error import DeobfuscationError
from .interpreter import Resolver, Interpreter
from .reference import *
from .visitor import Visitor


class Deobfuscator(Visitor[AST]):
    """
    Class performing AST de-obfuscation.

    You can configure it using:
        - evaluation_level: see EvaluationLevel
        - rename_symbols: if True, rename symbols that appear to be obfuscated

    You can clean an AST with the deobfuscate method. To add support for a new
    AST node type you just need to implement the corresponding visit_ method,
    which should return an equivalent de-obfuscated node.
    """

    class EvaluationLevel(Enum):
        """
        Used to specify what expressions should be evaluated : none of them,
        literal expressions or pure functions.
        """

        NONE = 0
        LITERAL = 1
        FUNCTION = 2

    __new_names: Dict[Reference, str]
    __resolver: Resolver
    __interpreter: Interpreter

    evaluation_level: "Deobfuscator.EvaluationLevel"
    rename_symbols: bool

    def __init__(
        self,
        program: Program,
        evaluation_level: Union["Deobfuscator.EvaluationLevel", int] = 1,
        rename_symbols: bool = False,
    ) -> None:
        self.__new_names = dict()
        self.__interpreter = Interpreter(program)
        self.__resolver = Resolver(self.__interpreter, program)

        self.evaluation_level = self.EvaluationLevel(evaluation_level)
        self.rename_symbols = rename_symbols

    def deobfuscate(self, ast: AST) -> AST:
        """
        Return the de-obfuscated version of an AST.
        """
        clean_ast = self.__deobfuscate(ast)
        assert isinstance(clean_ast, AST)
        return clean_ast

    def __deobfuscate(self, elt: Any) -> Any:
        """
        Internal de-obfuscation method, which can handle AST, lists (whose
        elements are individually de-obfuscated) and any other type, which is
        returned as-is.
        """
        # list
        if isinstance(elt, list):
            return list(map(self.__deobfuscate, elt))

        # non AST
        if not isinstance(elt, AST):
            return elt

        try:
            # visitable AST
            return_value = self.visit(elt)
            return return_value
        except NotImplementedError:
            # non visitable AST: deep de-obfuscated copy
            return self.__deobfuscated_copy(elt)

    def __deobfuscated_copy(self, ast: AST, **new_attributes) -> AST:
        """
        Performs a deep copy of an AST node, de-obfuscating each of its public
        fields.

        :arg new_attributes: Use it to force the value of some fields
        :returns: An equivalent de-obfuscated AST
        """
        for attribute, value in ast.__dict__.items():
            if attribute.startswith("_") or attribute in new_attributes:
                continue

            new_attributes[attribute] = self.__deobfuscate(value)

        return type(ast)(**new_attributes)

    # Expressions evaluation
    def __is_evaluable(self, expression: Expression) -> bool:
        """
        Check if an expression can be evaluated by the deobfuscator, depending
        on the value of self.evaluation_level.
        """
        if self.evaluation_level == self.EvaluationLevel.NONE:
            return False
        elif self.evaluation_level == self.EvaluationLevel.LITERAL:
            return isinstance(expression, Literal)
        elif self.evaluation_level == self.EvaluationLevel.FUNCTION:
            return isinstance(expression, (Literal, FunCall))
        else:
            msg = f"Invalid evaluation level: {self.evaluation_level}"
            raise DeobfuscationError(msg)

    def __unroll_expression(
        self, expression: Expression, operator: Optional[str] = None
    ) -> List[Expression]:
        """
        Used to transform associative operator tree to list of expressions. In
        other words, remove parenthesis from an expression. For example, if used
        with (a + (b + c)) + (a + (b * c)), it would return [a, b, c, a, b*c].

        :arg expression: Expression to unroll
        :arg operator: Select the operator to unroll relatively to
        :returns: The recursively unrolled expression relatively to operator
        """
        if not isinstance(expression, BinOp):
            return [expression]
        elif operator is not None and operator != expression.operator:
            return [expression]
        else:
            unrolled_left = self.__unroll_expression(
                expression.left, expression.operator
            )
            unrolled_right = self.__unroll_expression(
                expression.right, expression.operator
            )

            return unrolled_left + unrolled_right

    def __evaluate_unrolled_expression(
        self, operator: str, exprs: List[Expression]
    ) -> Expression:
        """
        Evaluate consecutive evaluable elements of an unrolled expression with
        an associative operator. Return an equivalent expression.
        """
        # TODO work with tuples instead of list + copy
        # TODO add commutativity support
        assert len(exprs) >= 1
        if len(exprs) == 1:
            expr = exprs[0]
            return expr
        else:
            # Copy exprs
            exprs = list(exprs)
            right = exprs.pop()
            left = exprs[-1]

            left_evaluable = self.__is_evaluable(left)
            right_evaluable = self.__is_evaluable(right)

            if left_evaluable and right_evaluable:
                tmp_op = BinOp(operator, left, right)
                value = self.__interpreter.evaluate(tmp_op)
                exprs[-1] = Literal.from_value(value)
                return self.__evaluate_unrolled_expression(operator, exprs)
            else:
                new_left = self.__evaluate_unrolled_expression(operator, exprs)
                return BinOp(operator, new_left, right)

    @classmethod
    def __is_associative(cls, operator: str) -> bool:
        """Check if an operator is associative."""
        # TODO move to operator.py
        if operator in ("+", "*", "&", "And", "Or", "Eqv"):
            return True
        else:
            return False

    # visit_ methods

    def visit_BinOp(self, bin_op: BinOp) -> AST:
        clean_bin_op = self.__deobfuscated_copy(bin_op)

        operator = clean_bin_op.operator
        if not self.__is_associative(operator):
            return clean_bin_op

        unrolled_bin_op = self.__unroll_expression(clean_bin_op)

        return self.__evaluate_unrolled_expression(operator, unrolled_bin_op)
