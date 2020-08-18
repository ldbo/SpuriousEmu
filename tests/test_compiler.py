from emu import Parser, Compiler, reference
from tests.test import (assert_correct_function, SourceFile, Result)


def compile_single_file(vbs: SourceFile) -> Result:
    return Compiler.compile_file(vbs).to_dict()


def compile_with_standard_library(vbs: SourceFile) -> Result:
    cpl = Compiler()
    cpl.load_host_project("./lib/VBA")
    ast = Parser.parse_file(vbs)
    cpl.add_module(ast, reference.ProceduralModule, "main_module")

    return cpl.program.to_dict()


def compile_project(project: SourceFile) -> Result:
    return Compiler.compile_project(project).to_dict()


def test_block():
    assert_correct_function("compiler_block", compile_single_file)


def test_function():
    assert_correct_function("compiler_function", compile_single_file)


def test_standard_library():
    assert_correct_function("compiler_standard_library",
                            compile_with_standard_library)


def test_project():
    assert_correct_function("compiler_project", compile_project)
