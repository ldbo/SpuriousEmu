from emu import syntax, interpreter, compiler
from tests.test import (assert_correct_function, SourceFile, Result)

from tests.test import run_function
from tests.test_preprocessor import preprocess
from tests.test_syntax import parsing


def compile_module(vbs: SourceFile) -> Result:
    ast = syntax.parse_file(vbs)
    cpl = compiler.Compiler()
    symbols = cpl.extract_symbols(ast, "test_module")
    return list(map(lambda t: str(t), symbols))


def test_block():
    # run_function("compiler_block", compile_module)
    pass


def test_function():
    # run_function("compiler_function", compile_module)
    pass


def test_gamaredon():
    # run_function("first_stage_gamaredon_improved", compile_module)
    pass


def test_draft():
    # run_function("draft", preprocess)
    # run_function("draft", parsing)
    # run_function("draft", compile_module)
    pass
