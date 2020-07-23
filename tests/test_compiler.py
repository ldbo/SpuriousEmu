from emu import compiler
from tests.test import (assert_correct_function, SourceFile, Result)


def compile_single_file(vbs: SourceFile) -> Result:
    return compiler.compile_file(vbs).to_dict()


def test_block():
    assert_correct_function("compiler_block", compile_single_file)


def test_function():
    assert_correct_function("compiler_function", compile_single_file)
