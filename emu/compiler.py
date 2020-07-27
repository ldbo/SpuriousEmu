"""Pseudo-compilation stage: allow to extract symbols and function body."""

from dataclasses import dataclass
from importlib.util import spec_from_file_location, module_from_spec
from inspect import getmembers, isfunction
from pathlib import Path
from types import ModuleType
from typing import Optional, Dict, Any, List

from .abstract_syntax_tree import *
from .function import ExternalFunction, Function, InternalFunction
from .memory import Memory
from .symbol import Symbol
from .visitor import Visitor
from . import syntax


@dataclass
class Program:
    """
    Represent a statically analysed program : a set of symbols, with a memory
    containing already initialized values.
    """
    symbols: Symbol
    memory: Memory

    def to_dict(self) -> Dict[str, Any]:
        d = dict()
        d['symbols'] = list(map(lambda t: t.full_name(), self.symbols))
        d['memory'] = {'functions': list(self.memory._functions.keys())}
        return d


class Compiler(Visitor):
    """
    Class used for symbols extraction. You can analyse several modules in a row
    by using analyse_module multiple times, and can then retrieve the built
    symbols and memory by accessing the corresponding properties. To start a
    new analysis, use reset.
    """
    __root_symbol: Symbol
    __current_node: Symbol
    __memory: Memory

    def __init__(self) -> None:
        self.reset()

    def reset(self) -> None:
        """Reset the state of the compiler, erasing symbols and memory."""
        self.__root_symbol = Symbol.build_root()
        self.__current_node = self.__root_symbol
        self.__memory = Memory()

    def analyse_module(self, ast: AST, module_name: str = None) -> None:
        """Analyse a single module, given his AST."""
        self.__current_node = self.__root_symbol.add_child(
            module_name, Symbol.Type.Module)

        # self.__parse_ast(ast)
        self.visit(ast)

    def add_builtin(self,
                    function: Union[InternalFunction,
                                    ExternalFunction.Signature],
                    name: Optional[str] = None) -> None:
        """Link a builtin function to the program."""
        typed_function: Function
        if type(function) is InternalFunction:
            typed_function = function
        else:
            typed_function = ExternalFunction.from_function(function)

        if name is None:
            name = typed_function.name

        symbol = self.__root_symbol.add_child(name, Symbol.Type.Function)
        self.__memory.add_function(symbol.full_name(), typed_function)

    def load_host_project(self, project_path: str) -> None:
        """Load a Python package as a VBA host project."""
        path = Path(project_path)
        assert(path.is_dir())
        project_name = path.name

        for module_path in path.glob("*.py"):
            module_name = f"{project_name}.{module_path.name[:-3]}"
            module_spec = spec_from_file_location(module_name,
                                                  module_path.absolute())
            module = module_from_spec(module_spec)
            module_spec.loader.exec_module(module)

            self.__load_host_module(module)

    def __load_host_module(self, module: ModuleType) -> None:
        """
        Extract objects defined in an already loaded Python module and add them
        to the current program.
        """
        def locally_defined(predicate):
            def decorated_predicate(member):
                return predicate(member) \
                    and member.__module__ == module.__name__

            return decorated_predicate

        for name, function in getmembers(module, locally_defined(isfunction)):
            self.add_builtin(function)

    @property
    def program(self):
        """Return the compiled program."""
        return Program(self.__root_symbol, self.__memory)

    def __parse_callable(self, ast: Union[FunDef, ProcDef]) -> None:
        # Extract information
        name = ast.name.name
        if ast.arguments is None:
            args = []
        else:
            args = list(map(lambda t: t.name, ast.arguments.args))
        body = ast.body

        # Build symbol and memory representation
        fct_symbol = self.__current_node.add_child(name,
                                                   Symbol.Type.Function)
        fct_object = InternalFunction(name, args, body)

        # Add them
        previous_node = self.__current_node
        self.__current_node = fct_symbol
        self.__memory.add_function(fct_symbol.full_name(), fct_object)

        for arg in args:
            self.__current_node.add_child(arg, Symbol.Type.Variable)

        if type(ast) is FunDef:
            self.__current_node.add_child(name, Symbol.Type.Variable)

        # Continue
        self.visit_Block(ast)

        self.__current_node = previous_node

    def visit_Block(self, block: Block) -> None:
        for statement in block.body:
            self.visit(statement)

    def visit_VarDec(self, var_dec: VarDec) -> None:
        self.__current_node.add_child(var_dec.identifier.name,
                                      Symbol.Type.Variable)

    def visit_VarAssign(self, var_assign: VarAssign) -> None:
        if var_assign.variable.name not in self.__current_node:
            self.__current_node.add_child(var_assign.variable.name,
                                          Symbol.Type.Variable)

    def visit_FunDef(self, fun_def: FunDef) -> None:
        self.__parse_callable(fun_def)

    def visit_ProcDef(self, proc_def: ProcDef) -> None:
        self.__parse_callable(proc_def)

    def visit_FunCall(self, fun_call: FunCall) -> None:
        pass

    def visit_If(self, if_block: If) -> None:
        pass

    def visit_For(self, for_loop: For) -> None:
        if for_loop.counter.name not in self.__current_node:
            self.__current_node.add_child(for_loop.counter.name,
                                          Symbol.Type.Variable)
        self.visit_Block(for_loop)

    def visit_OnError(self, on_error: OnError) -> None:
        pass

    def visit_Resume(self, resume: Resume) -> None:
        pass

    def visit_ErrorStatement(self, error: ErrorStatement) -> None:
        pass


def compile_file(path: str) -> Program:
    """Parse a file and then compile it."""
    return compile_files([path])


def compile_files(paths: List[str]) -> Program:
    """Parse and compile a list of files."""
    compiler = Compiler()

    for path in paths:
        ast = syntax.parse_file(path)
        compiler.analyse_module(ast, Path(path).name.split('.')[0])

    return compiler.program
