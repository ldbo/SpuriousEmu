from emu import syntax, interpreter, compiler
from tests.test import (assert_correct_function, SourceFile, Result)

from tests.test import run_function
from tests.test_preprocessor import preprocess
from tests.test_syntax import parsing


def extract_symbols(vbs: SourceFile) -> Result:
    ast = syntax.parse_file(vbs)
    cpl = compiler.Compiler()
    cpl.analyse_module(ast, "test_module")

    symbols_list = list(map(lambda t: str(t), cpl.symbols))
    return symbols_list


def test_block():
    run_function("compiler_block", extract_symbols)


def test_function():
    run_function("compiler_function", extract_symbols)


def test_draft():
    # run_function("draft", preprocess)
    # run_function("draft", parsing)
    # run_function("draft", compile_module)
    pass
