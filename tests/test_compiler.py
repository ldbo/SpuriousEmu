from emu import compiler, syntax
from tests.test import (assert_correct_function, SourceFile, Result)


def compile_single_file(vbs: SourceFile) -> Result:
    return compiler.compile_file(vbs).to_dict()


def compile_with_standard_library(vbs: SourceFile) -> Result:
    cpl = compiler.Compiler()
    cpl.load_host_project("./emu/VBA")
    ast = syntax.parse_file(vbs)
    cpl.analyse_module(ast, "main_module")

    return cpl.program.to_dict()


def test_block():
    assert_correct_function("compiler_block", compile_single_file)


def test_function():
    assert_correct_function("compiler_function", compile_single_file)


def test_standard_library():
    assert_correct_function("compiler_standard_library",
                            compile_with_standard_library)
