from emu import syntax, interpreter, compiler
from tests.test import (assert_correct_function, SourceFile, Result)

from tests.test import run_function


def interpreting(vbs: SourceFile) -> Result:
    ast = syntax.parse_file(vbs)
    comp = compiler.Compiler()
    comp.analyse_module(ast, "main_module")
    comp.add_builtin(
        lambda interp, args: print(f"msgBox {args[0].value}"),
        name="msgBox")

    interp = interpreter.Interpreter(comp.program.symbols, comp.program.memory)
    interp.run("Main")

    return []


def evaluate_expressions(vbs: SourceFile) -> Result:
    ast = syntax.parse_file(vbs)
    interp = interpreter.Interpreter()
    expressions = []
    for expression in ast.body:
        expressions.append(str(interp.evaluate(expression)))
    return {'expressions': expressions}


def test_types():
    assert_correct_function("types_evaluation", evaluate_expressions)


def test_literal_expressions():
    assert_correct_function('literal_expressions_interpreter',
                            evaluate_expressions)


def test_draf():
    run_function('interpreter_01', interpreting)


def test_draft():
    run_function("draft", interpreting)
