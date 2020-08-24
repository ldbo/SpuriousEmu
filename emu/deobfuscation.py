"""Deobfuscation tools."""

from enum import Enum
from typing import Any, Dict
from typing import Type as tType

from .abstract_syntax_tree import *
from .compiler import Program
from .error import DeobfuscationError
from .interpreter import Resolver, Interpreter
from .reference import *
from .visitor import Visitor

# TODO add pure function evaluation
# TODO sharing variable names between functions causes issues because of how
#      the resolver works

class Deobfuscator(Visitor[AST]):
    """
    Class performing AST de-obfuscation.

    You can configure it using:
        - evaluation_level: see EvaluationLevel
        - rename_symbols: if True, rename symbols that appear to be obfuscated

    You can clean an AST with the deobfuscate method. To add support for a new
    AST node type you just need to implement the corresponding visit_ method,
    which should return an equivalent de-obfuscated node. If it returns None,
    standard deep de-obfuscated copy is performed.
    """

    class EvaluationLevel(Enum):
        """
        Used to specify what expressions should be evaluated : none of them,
        literal expressions or pure functions.
        """

        NONE = 0
        LITERAL = 1
        FUNCTION = 2

    REFERENCE_PREFIXES = {
        Environment: "Environment",
        Project: "Project",
        ClassModule: "Class",
        ProceduralModule: "Module",
        Variable: "var",
    }

    CALLABLE_PREFIXES = {FunDef: "Function", ProcDef: "Procedure"}

    __program: Program
    __resolver: Resolver
    __interpreter: Interpreter

    __new_names: Dict[tType[Reference], str]
    __references_count: Dict[tType[Reference], int]
    __callable_count: Dict[Union[tType[FunDef], ProcDef], int]

    __evaluation_level: "Deobfuscator.EvaluationLevel"
    rename_symbols: bool

    # Configuration

    def __init__(
        self,
        program: Program,
        evaluation_level: Union["Deobfuscator.EvaluationLevel", int] = 1,
        rename_symbols: bool = False,
    ) -> None:
        self.__program = program
        self.__interpreter = Interpreter(program)
        self.__resolver = Resolver(self.__interpreter, program)

        self.__new_names = dict()
        self.__references_count = {
            reference_type: 0
            for reference_type in Deobfuscator.REFERENCE_PREFIXES.keys()
        }
        self.__callable_count = {
            callable_type: 0
            for callable_type in Deobfuscator.CALLABLE_PREFIXES.keys()
        }

        self.evaluation_level = evaluation_level
        self.rename_symbols = rename_symbols

    @property
    def evaluation_level(self) -> "Deobfuscator.EvaluationLevel":
        return self.__evaluation_level

    @evaluation_level.setter
    def evaluation_level(
        self, level: Union["Deobfuscator.EvaluationLevel", int]
    ) -> None:
        if level == Deobfuscator.EvaluationLevel.FUNCTION:
            msg = "Pure functions deobfuscation is not supported yet"
            raise NotImplemented(msg)
        self.__evaluation_level = Deobfuscator.EvaluationLevel(level)

    # API

    def deobfuscate(self, ast: AST) -> AST:
        """
        Return the de-obfuscated version of an AST.
        """
        clean_ast = self.__deobfuscate(ast)
        assert isinstance(clean_ast, AST)

        return clean_ast

    def deobfuscate_module(self, name: str) -> AST:
        """
        Return the de-obfuscated version of a module.
        """
        try:
            ast = self.__program.asts[name]
        except KeyError:
            msg = f"Can't find module {name}"
            raise DeobfuscationError(msg)

        self.__resolver.jump(self.__resolver.resolve(Identifier(name)))
        clean_ast = self.__deobfuscate(ast)
        assert isinstance(clean_ast, AST)

        return clean_ast

    # General utility

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

            if return_value is None:
                raise NotImplementedError("Back to default de-obfuscation.")

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

    # Symbol demangling
    def __craft_name(self, reference: Reference) -> str:
        """
        Craft a new name for the given reference, using the format
        {ReferenceType}_XX. See Deobfuscator.REFERENCE_PREFIXES.
        """
        count = self.__references_count[type(reference)] + 1
        self.__references_count[type(reference)] = count
        prefix = self.REFERENCE_PREFIXES[type(reference)]
        new_name = f"{prefix}_{count}"

        return new_name

    def __craft_arg(self, argument_count) -> str:
        """
        Craft a new argument name, using the format arg_XX.
        """
        return f"arg_{argument_count}"

    def __craft_callable_name(
        self, reference: FunctionReference, callable: Union[FunDef, ProcDef]
    ) -> str:
        """Craft a new function or procedure name, using the format
        [Function|Procedure]_XX. See Deobfuscator.CALLABLE_PREFIXES.
        """
        t = type(callable)
        count = self.__callable_count[t] + 1
        self.__callable_count[t] = count

        return f"{self.CALLABLE_PREFIXES[t]}_{count}"

    def __rename_identifier(self, identifier: Identifier) -> Identifier:
        """
        Craft a new name for the identifier using __craft_name
        """
        try:
            resolution = self.__resolver.resolve(identifier)
        except ResolutionError:
            return Identifier(identifier.name)

        try:
            new_name = self.__new_names[resolution]
        except KeyError:
            new_name = self.__craft_name(resolution)
            self.__new_names[resolution] = new_name

        return self.__deobfuscated_copy(identifier, name=new_name)

    def __rename_get(self, get: Get) -> Identifier:
        """Only rename the top parent node."""
        # TODO Rename recursively
        new_child = self.__deobfuscated_copy(get.child, name=get.child.name)

        if isinstance(get.parent, (FunCall, Get)):
            return self.__deobfuscated_copy(get, child=new_child)
        else:
            new_parent = self.visit(get.parent)
            return self.__deobfuscated_copy(
                get, parent=new_parent, child=new_child
            )

    def __rename_callable_definition(
        self, definition: Union[ProcDef, FunDef]
    ) -> Union[ProcDef, FunDef]:
        """
        Handle first the function name and arguments, to correctly rename them,
        and then make resolver jump to function.
        """
        # TODO Use generic type to ensure type(definition) is type(return)
        function_reference = self.__resolver.resolve(definition.name)
        new_name = self.__craft_callable_name(function_reference, definition)
        self.__new_names[function_reference] = new_name
        arguments = self.visit(definition.arguments)

        # Update return variable
        if isinstance(definition, FunDef):
            return_reference = function_reference.get_child(
                function_reference.name
            )
            self.__new_names[return_reference] = new_name

        self.__references_count[Variable] = 0
        self.__resolver.jump(function_reference)
        clean_definition = self.__deobfuscated_copy(
            definition, name=Identifier(new_name), arguments=arguments
        )
        self.__resolver.jump_back()

        return clean_definition

    def __rename_arg_list_def(self, arg_list: ArgListDef) -> ArgListDef:
        """
        Argument and local variable references do not use distinct types, so
        this enforces the use of __craft_arg instead of __craft_name.
        """
        arguments_count = 0
        args = []
        for arg_identifier in arg_list.args:
            arguments_count += 1
            new_name = self.__craft_arg(arguments_count)
            reference = self.__resolver.resolve(arg_identifier)
            self.__new_names[reference] = new_name
            args.append(Identifier(new_name))

        return self.__deobfuscated_copy(arg_list, args=args)

    # visit_ methods

    def visit_BinOp(self, bin_op: BinOp) -> AST:
        clean_bin_op = self.__deobfuscated_copy(bin_op)

        operator = clean_bin_op.operator
        if not self.__is_associative(operator):
            return clean_bin_op

        unrolled_bin_op = self.__unroll_expression(clean_bin_op)

        return self.__evaluate_unrolled_expression(operator, unrolled_bin_op)

    def visit_Identifier(self, identifier: Identifier) -> AST:
        if self.rename_symbols:
            return self.__rename_identifier(identifier)

    def visit_Get(self, get: Get) -> AST:
        if self.rename_symbols:
            return self.__rename_get(get)

    def visit_ArgListDef(self, arg_list: ArgListDef) -> AST:
        if self.rename_symbols:
            return self.__rename_arg_list_def(arg_list)

    def visit_FunDef(self, fun_def: FunDef) -> AST:
        if self.rename_symbols:
            return self.__rename_callable_definition(fun_def)

    def visit_ProcDef(self, proc_def: ProcDef) -> AST:
        if self.rename_symbols:
            return self.__rename_callable_definition(proc_def)
