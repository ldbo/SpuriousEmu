from emu import syntax, interpreter
from tests.test import (assert_correct_function, SourceFile, Result)


def interpreting(vbs: SourceFile) -> Result:
    ast = syntax.parse_file(vbs)
    interp = interpreter.Interpreter()
    return interp.evaluate_expression(ast.body[0])


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
