from emu import syntax, interpreter
from tests.test import (assert_correct_function, SourceFile, Result)
from tests.test import run_function, export_result


def interpreting(vbs: SourceFile) -> Result:
    ast = syntax.parse_file(vbs)
    interp = interpreter.Interpreter()
    return interp.evaluate_expression(ast.body[0])


def evaluate_expressions(vbs: SourceFile) -> Result:
    ast = syntax.parse_file(vbs)
    interp = interpreter.Interpreter()
    ret = []
    for expression in ast.body:
        ret.append(str(interp.evaluate_expression(expression)))
    return ret


def test_types():
    assert_correct_function("types_evaluation", evaluate_expressions)


def test_literal_expressions():
    run_function('interpreter_01', evaluate_expressions)
