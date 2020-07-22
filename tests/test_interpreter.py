from emu import syntax, interpreter, compiler
from tests.test import (assert_correct_function, SourceFile, Result)

from tests.test import run_function


def interpreting(vbs: SourceFile) -> Result:
    ast = syntax.parse_file(vbs)
    comp = compiler.Compiler()
    comp.analyse_module(ast, "main_module")
    comp.add_builtin(lambda var1: print(f"msgBox {var1}"), name="msgBox")

    interp = interpreter.Interpreter(comp.symbols, comp.memory)

    for main in comp.symbols.find('Main'):
        interp.call_procedure(comp.memory.get_function(main.full_name()))

    return []


def evaluate_expressions(vbs: SourceFile) -> Result:
    ast = syntax.parse_file(vbs)
    interp = interpreter.Interpreter()
    expressions = []
    for expression in ast.body:
        expressions.append(str(interp.evaluate_expression(expression)))
    return {'expressions': expressions}


def test_types():
    assert_correct_function("types_evaluation", evaluate_expressions)


def test_literal_expressions():
    assert_correct_function('literal_expressions_interpreter',
                            evaluate_expressions)


def test_draf():
    run_function('interpreter_01', interpreting)
