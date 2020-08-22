"""Converting AST to VBA."""

from dataclasses import dataclass, field

from .abstract_syntax_tree import *
from .type import Type
from .visitor import Visitor


@dataclass
class Formatter(Visitor):
    """
    Allow to format an AST to VBA code. You can call format_ast several times,
    the object will accumulate the formatted results which you can get with the
    output field. To start a new output, use reset.

    You can configure the identation character with  the indentation field.
    """
    indentation: str = " " * 4

    __formatted_node: Optional[str] = field(init=False, repr=None, default=None)
    __indentation_level: int = field(init=False, repr=None, default=0)

    def format_ast(self, ast: AST) -> str:
        """
        Format an AST to VBA, add it to the current output, and return the
        formatted AST.
        """
        previous_formatted_node = self.__formatted_node
        self.__formatted_node = ""
        self.visit(ast)
        formatted_node = self.__formatted_node
        self.__formatted_node = previous_formatted_node
        return formatted_node

    def __newline(self, line: str) -> None:
        self.__formatted_node += self.indentation * self.__indentation_level
        self.__formatted_node += line
        self.__formatted_node += "\n"

    # visit_ methods
    def visit_Block(self, block: Block) -> None:
        for statement in block.body:
            self.__newline(self.format_ast(statement))

    def visit_VarDec(self, var_dec: VarDec) -> None:
        self.__formatted_node = f"Dim " + self.format_ast(var_dec.identifier)

        if var_dec.type is not None:
            self.__formatted_node += " As "
            if isinstance(var_dec.type, Identifier):
                self.__formatted_node += self.format_ast(var_dec.type)
            else:
                self.__formatted_node += str(var_dec.type)

        if var_dec.value is not None:
            self.__formatted_node += " = " + self.format_ast(var_dec.value)

    def visit_VarAssign(self, var_assign: VarAssign) -> None:
        self.__formatted_node = self.format_ast(var_assign.variable)
        self.__formatted_node += " = "
        self.__formatted_node += self.format_ast(var_assign.value)

    def visit_FunDef(self, fun_def: FunDef) -> None:
        self.__formatted_node = "Function " + self.format_ast(fun_def.name)
        self.__formatted_node += self.format_ast(fun_def.arguments) + "\n"
        self.__indentation_level += 1
        self.visit_Block(fun_def)
        self.__indentation_level -= 1
        self.__newline("End Function")

    def visit_ProcDef(self, proc_def: ProcDef) -> None:
        self.__formatted_node = "Sub " + self.format_ast(proc_def.name)
        self.__formatted_node += self.format_ast(proc_def.arguments) + "\n"
        self.__indentation_level += 1
        self.visit_Block(proc_def)
        self.__indentation_level -= 1
        self.__newline("End Sub")

    def visit_Identifier(self, identifier: Identifier) -> None:
        self.__formatted_node = identifier.name

    def visit_Get(self, get: Get) -> None:
        self.visit(get.parent)
        self.__formatted_node += "." + self.format_ast(get.child)

    def visit_Literal(self, literal: Literal) -> None:
        if literal.type is Type.String:
            escaped_string = literal.value.replace('"', '""')
            self.__formatted_node = f'"{escaped_string}"'
        else:
            self.__formatted_node = str(literal.value)

    def visit_ArgListCall(self, arg_list_call: ArgListCall) -> None:
        self.__formatted_node = "("
        self.__formatted_node += ", ".join(
            self.format_ast(arg) for arg in arg_list_call.args
        )
        self.__formatted_node += ")"

    def visit_ArgListDef(self, arg_list_def: ArgListDef) -> None:
        self.__formatted_node = "("
        self.__formatted_node += ", ".join(
            self.format_ast(arg) for arg in arg_list_def.args
        )
        self.__formatted_node += ")"

    def visit_FunCall(self, fun_call: FunCall) -> None:
        self.__formatted_node = self.format_ast(fun_call.function)
        self.__formatted_node += self.format_ast(fun_call.arguments)

    def visit_UnOp(self, un_op: UnOp) -> None:
        self.__formatted_node = un_op.operator
        if un_op.operator not in ("+", "-"):
            self.__formatted_node += " "
        self.__formatted_node += self.format_ast(un_op.argument)

    def visit_BinOp(self, bin_op: BinOp) -> None:
        self.__formatted_node = "(" + self.format_ast(bin_op.left)
        self.__formatted_node += " " + bin_op.operator + " "
        self.__formatted_node += self.format_ast(bin_op.right) + ")"

    def visit_ElseIf(self, else_if: ElseIf) -> None:
        self.__formatted_node = "ElseIf " + self.format_ast(else_if.condition)
        self.__formatted_node += " Then\n"
        self.__indentation_level += 1
        self.visit_Block(else_if)
        self.__indentation_level -= 1

    def visit_If(self, if_block: If) -> None:
        self.__formatted_node = "If " + self.format_ast(if_block.condition)
        self.__formatted_node += " Then\n"
        self.__indentation_level += 1
        self.visit_Block(if_block)
        self.__indentation_level -= 1

        self.__formatted_node += "\n".join(
            self.format_ast(else_if) for else_if in if_block.elsifs
        )

        if if_block.else_block is not None:
            self.__formatted_node += "Else\n"
            self.__indentation_level += 1
            self.visit_Block(if_block.else_block)
            self.__indentation_level -= 1

        self.__newline("EndIf")

    def visit_For(self, for_loop: For) -> None:
        header = "For " + self.format_ast(for_loop.counter) + " = "
        header += self.format_ast(for_loop.start)
        header += " To " + self.format_ast(for_loop.end) + "\n"
        footer = "Next " + self.format_ast(for_loop.counter) + "\n"

        if for_loop.step is not None:
            header += " Step " + self.format_ast(for_loop.step)

        self.__formatted_node = header
        self.__indentation_level += 1
        self.visit_Block(for_loop)
        self.__indentation_level -= 1
        self.__newline(footer)
