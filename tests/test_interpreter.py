from emu import Parser, interpreter, Compiler, reference
from tests.test import (assert_correct_function, SourceFile, Result)


def interpreting(vbs: SourceFile) -> Result:
    ast = Parser.parse_file(vbs)
    comp = Compiler()
    comp.add_module(ast, reference.ProceduralModule, "main_module")
    comp.load_host_project("./lib/VBA")
    comp.load_host_project("./lib/WSH")

    interp = interpreter.Interpreter(comp.program)
    interp.run("Main")

    return interp._outside_world.to_dict()


def evaluate_expressions(vbs: SourceFile) -> Result:
    ast = Parser.parse_file(vbs)
    interp = interpreter.Interpreter()
    expressions = []
    for expression in ast.body:
        expressions.append(str(interp.evaluate(expression)))
    return {'expressions': expressions}


def test_type():
    assert_correct_function("interpreter_type", evaluate_expressions)


def test_literal_expression():
    assert_correct_function('interpreter_literal_expression',
                            evaluate_expressions)


def test_01():
    assert_correct_function('interpreter_01', interpreting)


def test_02():
    assert_correct_function("interpreter_02", interpreting)
