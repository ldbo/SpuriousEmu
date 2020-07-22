"""Pseudo-compilation stage: search all the defined objects in a solution."""

from typing import Optional

from .abstract_syntax_tree import *
from .function import ExternalFunction, InternalFunction
from .symbol import Symbol
from .memory import Memory


class Compiler:
    __root_symbol: Optional[Symbol]
    __current_node: Optional[Symbol]
    __memory: Optional[Memory]

    def __init__(self) -> None:
        self.reset()

    def reset(self) -> None:
        self.__root_symbol = Symbol.build_root()
        self.__current_node = None
        self.__memory = Memory()

    def analyse_module(self, ast: AST, module_name: str = None) -> None:
        self.__current_node = self.__root_symbol.add_child(
            module_name, Symbol.Type.Module)

        self.__parse_ast(ast)

    def add_builtin(self,
                    function: Union[InternalFunction,
                                    ExternalFunction.Signature],
                    name: Optional[str] = None) -> None:
        if type(function) is not InternalFunction:
            function = ExternalFunction.from_function(function)

        if name is None:
            name = function.name

        symbol = self.__root_symbol.add_child(name, Symbol.Type.Function)
        self.__memory.add_function(symbol.full_name(), function)

    @property
    def symbols(self):
        return self.__root_symbol

    @property
    def memory(self):
        return self.__memory

    def __parse_ast(self, ast):
        def type_test(node_type):
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
            name = ast.name.name

            if ast.arguments is None:
                args = []
            else:
                args = list(map(lambda t: t.name, ast.arguments.args))

            body = ast.body
            fct_symbol = self.__current_node.add_child(name,
                                                       Symbol.Type.Function)
            fct_object = InternalFunction(name, args, body)

            self.__current_node = fct_symbol
            self.__memory.add_function(fct_symbol.full_name(), fct_object)

            self.__parse_ast(Block(body))
