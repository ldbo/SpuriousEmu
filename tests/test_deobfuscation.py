from typing import Dict, Optional

from emu import (
    AST,
    Compiler,
    Deobfuscator,
    Formatter,
    ManglingClassifier,
    Serializer,
)
from tests.test import assert_correct_function, Result, SourceFile


def deobfuscate(
    vbs: SourceFile,
    level: int,
    rename_symbols: bool,
    classifier: Optional[ManglingClassifier] = None,
) -> Dict[str, AST]:
    program = Compiler.compile_file(vbs)
    deobfuscator = Deobfuscator(
        program,
        evaluation_level=level,
        rename_symbols=rename_symbols,
        mangling_classifier=classifier,
    )
    clean_asts = {
        module: deobfuscator.deobfuscate(ast)
        for module, ast in program.asts.items()
    }

    return clean_asts


def format_deobfuscate(
    vbs: SourceFile,
    level: int,
    rename_symbols: bool,
    classifier: Optional[ManglingClassifier] = None,
) -> None:
    """
    Use this for debugging new deobfuscation features: it displays the
    de-obfuscated code instead of returning ASTs.
    """
    clean_ast = deobfuscate(vbs, level, rename_symbols, classifier)
    formatter = Formatter()

    for module, ast in clean_ast.items():
        print(f"Module: {module}")
        print(20 * "=" + "")
        print(formatter.format_ast(ast))
        print(20 * "=")
        print(2 * "\n")


def json_deobfuscate(
    vbs: SourceFile,
    level: int,
    rename_symbols: bool,
    classifier: Optional[ManglingClassifier] = None,
) -> Result:
    """
    Use this for production tests : it outputs the de-obfuscated AST Result.
    """
    clean_asts = deobfuscate(vbs, level, rename_symbols, classifier)

    return {module: ast.to_dict() for module, ast in clean_asts.items()}


def deobfuscate_none(vbs: SourceFile) -> Result:
    return json_deobfuscate(vbs, Deobfuscator.EvaluationLevel.NONE, False)


def deobfuscate_literal(vbs: SourceFile) -> Result:
    return json_deobfuscate(vbs, Deobfuscator.EvaluationLevel.LITERAL, False)


def deobfuscate_symbol(vbs: SourceFile) -> Result:
    return json_deobfuscate(vbs, Deobfuscator.EvaluationLevel.NONE, True)


def deobfuscate_symbol_markov(vbs: SourceFile) -> Result:
    classifier = Serializer.load("./emu/resources/english_n1.spemu-mancl")

    return json_deobfuscate(
        vbs, Deobfuscator.EvaluationLevel.NONE, True, classifier
    )


def test_none():
    assert_correct_function("deobfuscation_none", deobfuscate_none)


def test_literal():
    assert_correct_function("deobfuscation_literal", deobfuscate_literal)


def test_symbol():
    assert_correct_function("deobfuscation_symbol", deobfuscate_symbol)


def test_symbol_markov():
    assert_correct_function(
        "deobfuscation_symbol_markov", deobfuscate_symbol_markov
    )


def test_mangling_classifier():
    with open("./tests/words_alpha.txt") as f:
        trainig_data = f.read().split("\n")

    trainig_data = trainig_data[::100]
    classifier = ManglingClassifier(
        n=1, training_data=trainig_data, false_negative_rate=0.01
    )

    assert "uBTZUpQmXdk" in classifier
