"""Pseudo-compilation stage: allow to extract symbols and function body."""

from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Dict, Any, List

from .abstract_syntax_tree import *
from .function import ExternalFunction, Function, InternalFunction
from .memory import Memory
from .symbol import Symbol
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


class Compiler:
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

        self.__parse_ast(ast)

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

    @property
    def program(self):
        """Return the compiled program."""
        return Program(self.__root_symbol, self.__memory)

    def __parse_ast(self, ast: AST) -> None:
        """Recursively parse an AST, used by analyse_module."""
        def type_test(node_type) -> bool:
            return type(ast) is node_type

        if type_test(Block):
            for statement in ast.body:
                self.__parse_ast(statement)
        elif type_test(VarDec):
            self.__current_node.add_child(ast.identifier.name,
                                          Symbol.Type.Variable)
        elif type_test(VarAssign):
            if ast.variable.name not in self.__current_node:
                self.__current_node.add_child(ast.variable.name,
                                              Symbol.Type.Variable)
        elif type_test(FunDef) or type_test(ProcDef):
            self.__parse_callable(ast)
        elif type_test(For):
            if ast.counter.name not in self.__current_node:
                self.__current_node.add_child(ast.counter.name,
                                              Symbol.Type.Variable)
            self.__parse_ast(ast.body)

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
        self.__parse_ast(Block(body))

        self.__current_node = previous_node


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
