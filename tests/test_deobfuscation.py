from emu import Deobfuscator, Compiler
from tests.test import assert_correct_function, Result, SourceFile


def deobfuscate_literal(vbs: SourceFile) -> Result:
    program = Compiler.compile_file(vbs)
    deobfuscator = Deobfuscator(program)
    deobfuscator.evaluation_level = deobfuscator.EvaluationLevel.LITERAL
    deobfuscator.rename_symbols = True
    cleans_asts = {
        name: deobfuscator.deobfuscate(ast)
        for name, ast in program.asts.items()
    }

    for name in cleans_asts:
        cleans_asts[name] = cleans_asts[name].to_dict()

    return cleans_asts


def test_literal():
    assert_correct_function("deobfuscation_literal", deobfuscate_literal)
