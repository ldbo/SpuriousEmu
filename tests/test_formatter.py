from emu import Parser, Formatter
from tests.test import assert_correct_function, SourceFile, Result


def format_file(vbs: SourceFile) -> Result:
    ast = Parser.parse_file(vbs)
    formatter = Formatter()
    format_output = formatter.format_ast(ast)
    return {"code": format_output}


def test_formatter():
    assert_correct_function("formatter", format_file)
