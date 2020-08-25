"""Converting AST to VBA."""

from dataclasses import dataclass, field

from .abstract_syntax_tree import *
from .type import Type
from .visitor import Visitor


@dataclass
class Formatter(Visitor[str]):
    """
    Allow to format an AST to VBA code. You can call format_ast several times,
    the object will accumulate the formatted results which you can get with the
    output field. To start a new output, use reset.

    You can configure the identation character with the indentation field, and
    the end-of-line with eol.
    """

    indentation: str = " " * 4
    eol: str = "\n"

    __formatted_node: Optional[str] = field(init=False, repr=None, default=None)
    __indentation_level: int = field(init=False, repr=None, default=0)
    __newline: str = field(init=False, repr=None)

    def __post_init__(self) -> None:
        self.__newline = self.eol * 2

    def format_ast(self, ast: AST) -> str:
        """
        Format an AST to VBA, add it to the current output, and return the
        formatted AST.
        """
        output = self.visit(ast)
        assert self.__indentation_level == 0
        return output

    def __indent(self) -> str:
        """Return the current indentation string."""
        return self.__indentation_level * self.indentation

    # visit_ methods
    def visit(self, *args, **kwargs) -> str:
        tmp = self.__indentation_level
        ret = super().visit(*args, **kwargs)
        assert tmp == self.__indentation_level
        return ret

    def visit_Block(self, block: Block) -> str:
        output = ""
        for statement in block.body:
            if isinstance(statement, FunCall):
                output += self.__indent() + self.visit(statement) + self.eol
            else:
                output += self.visit(statement)

        return output

    def visit_VarDec(self, var_dec: VarDec) -> str:
        output = self.__indent() + f"Dim " + self.visit(var_dec.identifier)

        if var_dec.type is not None:
            output += " As "
            if isinstance(var_dec.type, Identifier):
                output += self.visit(var_dec.type)
            else:
                output += str(var_dec.type)

        if var_dec.value is not None:
            output += " = " + self.visit(var_dec.value)

        return output + self.eol

    def visit_VarAssign(self, var_assign: VarAssign) -> str:
        output = self.__indent() + self.visit(var_assign.variable)
        output += " = "
        output += self.visit(var_assign.value)

        return output + self.eol

    def visit_FunDef(self, fun_def: FunDef) -> str:
        output = self.__indent() + "Function " + self.visit(fun_def.name)
        output += self.visit(fun_def.arguments) + self.eol
        self.__indentation_level += 1
        output += self.visit_Block(fun_def)
        self.__indentation_level -= 1
        output += self.__indent() + "End Function" + self.__newline

        return output

    def visit_ProcDef(self, proc_def: ProcDef) -> str:
        output = self.__indent() + "Sub " + self.visit(proc_def.name)
        output += self.visit(proc_def.arguments) + self.eol
        self.__indentation_level += 1
        output += self.visit_Block(proc_def)
        self.__indentation_level -= 1
        output += self.__indent() + "End Sub" + self.__newline

        return output

    def visit_Identifier(self, identifier: Identifier) -> str:
        return identifier.name

    def visit_Get(self, get: Get) -> str:
        return self.visit(get.parent) + "." + self.visit(get.child)

    def visit_Literal(self, literal: Literal) -> str:
        if literal.type is Type.String:
            escaped_string = literal.value.replace('"', '""')
            return f'"{escaped_string}"'
        else:
            return str(literal.value)

    def visit_ArgListCall(self, arg_list_call: ArgListCall) -> str:
        output = "("
        output += ", ".join(self.visit(arg) for arg in arg_list_call.args)
        output += ")"

        return output

    def visit_ArgListDef(self, arg_list_def: ArgListDef) -> str:
        output = "("
        output += ", ".join(self.visit(arg) for arg in arg_list_def.args)
        output += ")"

        return output

    def visit_FunCall(self, fun_call: FunCall) -> str:
        return self.visit(fun_call.function) + self.visit(fun_call.arguments)

    def visit_UnOp(self, un_op: UnOp) -> str:
        output = un_op.operator
        if un_op.operator not in ("+", "-"):
            output += " "
        output += self.visit(un_op.argument)

        return output

    def visit_BinOp(self, bin_op: BinOp) -> str:
        def parenthesize(expr: Expression) -> str:
            if isinstance(expr, BinOp):
                return "(" + self.visit(expr) + ")"
            else:
                return self.visit(expr)

        output = parenthesize(bin_op.left)
        output += " " + bin_op.operator + " "
        output += parenthesize(bin_op.right)

        return output

    def visit_ElseIf(self, else_if: ElseIf) -> str:
        output = self.__indent() + "ElseIf " + self.visit(else_if.condition)
        output += " Then" + self.eol
        self.__indentation_level += 1
        output += self.visit_Block(else_if)
        self.__indentation_level -= 1

        return output

    def visit_If(self, if_block: If) -> str:
        output = self.eol + self.__indent() + "If "
        output += self.visit(if_block.condition) + " Then" + self.eol
        self.__indentation_level += 1
        output += self.visit_Block(if_block)
        self.__indentation_level -= 1

        output += "".join(self.visit(else_if) for else_if in if_block.elsifs)

        if if_block.else_block is not None:
            output += self.__indent() + "Else" + self.eol
            self.__indentation_level += 1
            output += self.visit_Block(if_block.else_block)
            self.__indentation_level -= 1

        output += self.__indent() + "End If" + self.__newline

        return output

    def visit_For(self, for_loop: For) -> str:
        output = self.eol + self.__indent() + "For "
        output += self.visit(for_loop.counter)
        output += " = " + self.visit(for_loop.start)
        output += " To " + self.visit(for_loop.end)

        if for_loop.step is not None:
            output += " Step " + self.visit(for_loop.step)

        output += self.eol

        self.__indentation_level += 1
        output += self.visit_Block(for_loop)
        self.__indentation_level -= 1

        output += self.__indent() + "Next " + self.visit(for_loop.counter)

        return output + self.__newline

    def visit_OnError(self, on_error: OnError) -> str:
        output = self.__indent() + "On Error "
        if on_error.goto is None:
            output += "Resume Next"
        else:
            output += "Goto " + self.visit(on_error.goto)

        return output + self.eol

    def visit_Resume(self, resume: Resume) -> str:
        output = self.__indent() + "Resume "
        if resume.goto is None:
            output += "Next"
        else:
            output += self.visit(resume.goto)

        return output + self.eol

    def visit_ErrorStatement(self, error: ErrorStatement) -> str:
        return self.__indent() + "Error " + self.visit(error.number) + self.eol
