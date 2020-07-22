"""Pseudo-compilation stage: search all the defined objects in a solution."""

from dataclasses import dataclass

from .abstract_syntax_tree import *
from .symbol import Symbol


@dataclass
class Function:
    body: Block
    args: ArgList


FunctionsTable = Dict[str, Function]


class Compiler:
    __current_node: Symbol
    __functions_table: FunctionsTable

    def __init__(self) -> None:
        pass

    def extract_symbols(self, ast: AST, module_name: str = None) -> None:
        root = Symbol.build_root()
        self.__current_node = root.add_child(module_name, Symbol.Type.Module)
        self.__functions_table = dict()

        self.__parse_ast(ast)

        return root, self.__functions_table

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
            args = ast.arguments
            body = ast.body

            fct_symbol = self.__current_node.add_child(name,
                                                       Symbol.Type.Function)
            fct_object = Function(body, args)

            self.__current_node = fct_symbol
            self.__functions_table[fct_symbol.full_name()] = fct_object

            self.__parse_ast(Block(body))
