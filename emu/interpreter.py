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

    def visit_Identifier(self, identifier: Identifier) -> None:
        self._resolution = Resolver.resolve_from(
            self._current_reference, identifier.name)

    def visit_Get(self, get: Get) -> None:
        """Resolve a composite symbol."""
        parent_resolution = self.resolve(get.parent)
        self._resolution = parent_resolution.get_child(get.child.name)

    def resolve(self, symbol: Union[Identifier, Get]) -> Reference:
        """
        Translate a given simple identifier into the corresponding reference.

        :raises ResolutionError: If the resolution can't be done
        """
        self.visit(symbol)
        resolution = self._resolution
        self._resolution = None
        return resolution

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

        if reference.parent is None:
            msg = f"Symbol {name} is not found in {reference}"
            raise ResolutionError(msg)

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
            except AssertionError as e:
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
        resolution = self._resolver.resolve(identifier)
        if isinstance(resolution, Variable):
            self._evaluation = self._memory.get_variable(str(resolution))
        elif isinstance(resolution, FunctionReference):
            self._evaluation = self._memory.functions[str(resolution)]
        else:
            msg = f"Can't interprete identifier {identifier.name}"
            raise InterpretationError(msg)

    def visit_Get(self, get: Get) -> None:
        """
        Compute the value of a Get expression.
        """
        child_name = get.child.name

        # Get parent value
        if isinstance(get.parent, Get):
            try:
                parent = self._resolver.resolve(get.parent)
            except ResolutionError:
                parent = self.evaluate(get.parent)
        elif isinstance(get.parent, Identifier):
            parent = self._resolver.resolve(get.parent)
        elif isinstance(get.parent, FunCall):
            parent = self.evaluate(get.parent)
        else:
            msg = f"Get node with {type(parent)} parent is not supported"
            raise InterpretationError(msg)

        # Access child
        if isinstance(parent, Value):
            self.__access_member(parent, child_name)
        elif isinstance(parent, Reference):
            if isinstance(parent, Variable):
                parent_value = self._memory.get_variable(str(parent))
                self.__access_member(parent_value, child_name)
            elif parent.category is Reference.Category.Structural:

                child = parent.get_child(child_name)
                assert(isinstance(child, (Variable, FunctionReference)))
                if isinstance(child, Variable):
                    self._evaluation = self._memory.get_variable(str(child))
                elif isinstance(child, FunctionReference):
                    self._evaluation = self._memory.functions[str(child)]
        else:
            msg = f"Evaluation error, with parent {get.parent} and " \
                + f"child {get.child}"
            raise InterpretationError(msg)

    def __access_member(self, object_value: Value, child_name: str) \
            -> None:
        """
        Helper function to access a variable/method from an object, putting its
        value in self._evaluation.
        """
        assert(isinstance(object_value, Object))
        try:
            variable = object_value.variables[child_name]
            self._evaluation = variable
        except KeyError:
            class_ref = object_value.class_reference
            method_reference = class_ref.get_child(child_name)
            unbound_method = self._memory.functions[str(method_reference)]
            bound_method = unbound_method.create_bound_method(object_value)
            self._evaluation = bound_method

    def visit_BinOp(self, bin_op: BinOp) -> None:
        left_value = self.evaluate(bin_op.left)
        right_value = self.evaluate(bin_op.right)

        op = bin_op.operator
        self._evaluation = op.operate(left_value, right_value)

    def visit_FunCall(self, fun_call: FunCall) -> None:
        # TODO handle method call
        # Resolution
        function = self.evaluate(fun_call.function)
        assert(isinstance(function, Function))

        # Arguments handling
        arg_list = fun_call.arguments.args
        arg_values = list(map(self.evaluate, arg_list))

        # Move to the function
        return_value = self.call_function(function, arg_values)

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
        done = done or conditional.else_block is None
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
        # TODO handle attribute access for internal class method
        if type(function) is ExternalFunction:
            if function.parent_object is not None:
                arguments_values = [function.parent_object] + arguments_values
            python_value = function.external_function(self, arguments_values)
            return Value.from_python_base_type(python_value)
        elif type(function) is InternalFunction:
            # Init memory and load arguments
            self._resolver.jump(function.reference)
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
            self._resolver.jump_back()

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

    def add_file_event(self, event_type: str, path: str,
                       data: Optional[str] = None) -> None:
        event = {'type': event_type, 'path': path}
        if data is not None:
            event['data'] = data
        self._outside_world.add_event(OutsideWorld.EventType.FILE, event)

    def add_network_event(self, *args, **kwargs) -> None:
        # TODO
        pass

    @staticmethod
    def run_program(program: Program, entry_point: str = "Main") \
            -> OutsideWorld:
        """
        Execute the a program using the given entry_point.
        """
        interpreter = Interpreter(program)
        interpreter.run(entry_point)

        return interpreter._outside_world
