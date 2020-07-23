from emu import syntax, interpreter, compiler
from tests.test import (assert_correct_function, SourceFile, Result)

from tests.test import run_function
from tests.test_syntax import parsing


def extract_symbols(vbs: SourceFile) -> Result:
    ast = syntax.parse_file(vbs)
    cpl = compiler.Compiler()
    cpl.analyse_module(ast, "test_module")
    return cpl.program.to_dict()


def test_block():
    # run_function("compiler_block", extract_symbols)
    pass


def test_function():
    # run_function("compiler_function", extract_symbols)
    pass


def test_draft():
    # run_function("first_stage_gamaredon_improved", parsing)
    run_function("first_stage_gamaredon_improved", extract_symbols)
    pass
