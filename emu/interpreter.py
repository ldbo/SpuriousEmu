"""Provide the Interpreter class, that is used to execute VBA programs."""

from typing import List, Optional

from .function import Function, InternalFunction, ExternalFunction
from .abstract_syntax_tree import *
from .compiler import Program
from .error import InterpretationError, ResolutionError
from .reference import Reference, Environment, Variable, FunctionReference
from .side_effect import Memory, OutsideWorld
from .operator import OPERATORS_MAP
from .value import Value, Integer, Object
from .visitor import Visitor


class Resolver(Visitor):
    """Used to tell the interpreter what each name referes to"""
    _interpreter: "Interpreter"
    _program: Program
    _current_reference: Reference
    _previous_references: List[Reference]
    _resolution: Optional[Reference]

    def __init__(self, interpreter, program=None) -> None:
        self._interpreter = interpreter
        self._program = program if program is not None \
            else Program(Memory(), Environment())
        self._current_reference = self._program.environment
        self._previous_references = []
        self._resolution = None

    def visit_Identifier(self, identifier) -> None:
        self._resolution = Resolver.resolve_from(self._current_reference,
                                                 identifier.name)

    def visit_MemberAccess(self, member_access: MemberAccess) -> None:
        root_element = member_access.parts[0]
        self.visit(root_element)
        root = self._resolution
        self._resolution = None

        identifiers = list(map(lambda part: part if type(part) is Identifier
                               else part.function.part[0],
                               member_access.parts))

        position = 0
        identifier = identifiers[position]
        reference = self.resolve(identifier)

        # First part of a MemberAccess : only structural elements, until a
        # variable of function call
        while True:
            if isinstance(reference, FunctionReference):
                if position == len(member_access.parts) - 1:
                    self._resolution = reference
                    return
                else:
                    assert(isinstance(member_access.parts[position], FunCall))
                    # TODO reference = self._interpreter.call_with_ref
                    # (reference, FunCall)
            elif isinstance(reference, Variable):
                break
            else:
                position += 1
                identifier = identifiers[position]
                reference = reference.get_child(identifier.name)

        # Now we're only dealing with objects

        self._resolution = root

    def resolve(self, name: Union[Identifier, MemberAccess]) -> Reference:
        """
        Translate a given simple or complexe identifier into the corresponding
        reference.

        :raises ResolutionError: If the resolution can't be done
        """
        self.visit(name)
        return self._resolution

    @staticmethod
    def resolve_from(reference: Reference, name: str, go_down: bool = False,
                     exclude: Reference = None) -> Reference:
        # TODO efficient search
        try:
            return reference.search_child(name)
        except ResolutionError:
            pass

        if reference.name == name:
            return reference

        return Resolver.resolve_from(reference.parent, name,
                                     go_down=True, exclude=reference)

    def find_function(self, function_name) -> List[Function]:
        functions = []
        for project in self._program.environment.children:
            for module in project.children:
                try:
                    functions.append(module.get_function(function_name))
                except ResolutionError:
                    pass

        return functions

    def jump(self, reference) -> None:
        """Change  to a new scope"""
        self._previous_references.append(self._current_reference)
        self._current_reference = reference

    def jump_back(self) -> None:
        """Go back to the last scope"""
        self._current_reference = self._previous_references.pop()


class Interpreter(Visitor):
    """Class used to interprete a program which has already been compiled."""
    _memory: Memory
    _resolver: Resolver
    _outside_world: OutsideWorld
    _evaluation: Optional[Value]

    # Main API
    def __init__(self, program: Optional[Program] = None):
        """
        Create an interpreter used to run a precompiled program, or simply to
        run an AST starting without a state.
        """
        if program is None:
            program = Program(Memory(), Environment())

        self._memory = program.memory
        self._resolver = Resolver(self, program)
        self._outside_world = OutsideWorld()
        self._evaluation = None

    def run(self, function_name: str, args: Optional[List[Value]] = None) \
            -> None:
        """
        Search for a function entry point in the program references, and
        execute it. You should use it most of the time.
        """
        if args is None:
            args = []

        functions = self._resolver.find_function(function_name)
        for function in functions:
            self._resolver._current_reference = function
            self.call_function(self._memory.get_function(str(function)), args)

    def print_stdout(self, state: bool) -> None:
        """Enable or disable stdout forwarding"""
        if state:
            hook = print
        else:
            def hook(*args, **kwargs):
                pass

        self._outside_world.add_hook(OutsideWorld.EventType.STDOUT, hook)

    # Internal interface

    def evaluate(self, expression: Expression) -> Value:
        """Return the value of an expression"""
        assert(isinstance(expression, Expression))

        self.visit(expression)
        value = self._evaluation
        self._evaluation = None
        return value

    def visit_Block(self, block: Block) -> None:
        for statement in block.body:
            try:
                self.visit(statement)
            except InterpretationError as e:
                raise e
            except Exception as e:
                msg = f"{statement.file}:{statement.line_number}: {e}"
                raise InterpretationError(msg)

    def visit_VarDec(self, var_dec: VarDec) -> None:
        # TODO
        # Add initial value
        pass

    def visit_VarAssign(self, var_assign: VarAssign) -> None:
        value = self.evaluate(var_assign.value)
        self._memory.set_variable(var_assign.variable.name, value)

    def visit_Literal(self, literal: Literal) -> None:
        self._evaluation = Value.from_literal(literal)

    def visit_Identifier(self, identifier: Identifier) -> None:
        self._evaluation = self._memory.get_variable(identifier.name)

    def visit_MemberAccess(self, member_access: MemberAccess) -> None:
        reference = self._resolver.resolve(member_access)
        assert(isinstance(reference, Variable))

        self._evaluation = self._memory.get_variable(reference.name)

    def visit_BinOp(self, bin_op: BinOp) -> None:
        left_value = self.evaluate(bin_op.left)
        right_value = self.evaluate(bin_op.right)

        op = bin_op.operator
        self._evaluation = op.operate(left_value, right_value)

    def visit_FunCall(self, fun_call: FunCall) -> None:
        # Resolution
        function_reference = self._resolver.resolve(fun_call.function)
        assert(isinstance(function_reference, FunctionReference))
        function_name = str(function_reference)
        function_object = self._memory.get_function(function_name)

        # Arguments handling
        arg_list = fun_call.arguments.args
        arg_values = list(map(self.evaluate, arg_list))

        # Move to the function
        self._resolver.jump(function_reference)
        return_value = self.call_function(function_object, arg_values)
        self._resolver.jump_back()

        self._evaluation = return_value

    def visit_If(self, conditional: If) -> None:
        done = False

        # If
        condition_value = self.evaluate(conditional.condition)
        if condition_value.convert_to(Type.Boolean).value:
            self.visit_Block(conditional)
            done = True

        # Else ifs
        if not done:
            for elseif in conditional.elsifs:
                condition_value = elseif.evaluate(elseif.condition)
                if condition_value.convert_to(Type.Boolean).value:
                    self.visit_Block(elseif)
                    done = True

        # Else
        if not done:
            self.visit_Block(conditional.else_block)

    def visit_For(self, loop: For) -> None:
        # Evaluation of parameters
        # TODO check if Let-coercible to Double, raise E13 if not
        start_value = self.evaluate(loop.start)
        end_value = self.evaluate(loop.end)
        if loop.step is None:
            step_value = Integer(1)
        else:
            step_value = self.evaluate(loop.step)
        step_literal = Literal.from_value(step_value)
        counter_name = loop.counter.name

        # Comparison function
        if step_value.value >= 0:
            def keep_going():
                counter_value = self._memory.get_variable(counter_name)
                return counter_value.value <= end_value.value
        else:
            def keep_going():
                counter_value = self._memory.get_variable(counter_name)
                return counter_value.value > end_value.value

        # Loop
        self._memory.set_variable(counter_name, start_value)
        while keep_going():
            self.visit_Block(loop)
            add_expression = BinOp(OPERATORS_MAP['+'],
                                   loop.counter, step_literal)
            new_counter_value = self.evaluate(add_expression)
            self._memory.set_variable(counter_name, new_counter_value)

    def visit_OnError(self, on_error: OnError) -> None:
        # TODO
        pass

    def visit_Resume(self, resume: Resume) -> None:
        # TODO
        pass

    def visit_ErrorStatement(self, error: ErrorStatement) -> None:
        # TODO
        pass

    def call_function(self, function: Function,
                      arguments_values: List[Value]) -> Optional[Value]:
        if type(function) is ExternalFunction:
            return function.external_function(self, arguments_values)
        elif type(function) is InternalFunction:
            # Init memory and load arguments
            self._memory.new_locals()
            for name, value in zip(function.arguments_names, arguments_values):
                self._memory.set_variable(name, value)

            # Execute function
            self.visit_Block(function)

            # Get return value
            return_value: Optional[Value]
            try:
                return_value = self._memory.get_variable(function.name)
            except KeyError:
                return_value = None

            # Clean memory
            self._memory.discard_locals()

            return return_value
        else:
            msg = f"{type(function)} is not handled yet by the interpreter"
            raise NotImplemented(msg)

    def create_object(self, class_name: str) -> Object:
        """
        Instanciate a new object with None values for each of its
        attributes.

        :param class_name: Absolute name of the VBA class
        """
        vba_class = self._memory.classes[class_name]
        variables = {name: None for name in vba_class.variables}
        obj = Object(vba_class.class_reference, variables)
        return obj

    # External functions interface
    def add_stdout(self, content: str) -> None:
        self._outside_world.add_event(OutsideWorld.EventType.STDOUT, content)

    def add_command_execution(self, command: str) -> None:
        self._outside_world.add_event(
            OutsideWorld.EventType.EXECUTION,
            command
        )

    def add_file_event(self, *args, **kwargs) -> None:
        # TODO
        pass

    def add_network_event(self, *args, **kwargs) -> None:
        # TODO
        pass

    @staticmethod
    def run_program(program: Program) -> None:
        interpreter = Interpreter(program)
        interpreter.run('Main')
