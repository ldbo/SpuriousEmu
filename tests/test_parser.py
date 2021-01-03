from nose.tools import assert_equals, assert_raises

from emu.abstract_syntax_tree import Literal, Name
from emu.data import (
    Boolean,
    Currency,
    Double,
    Empty,
    Null,
    ObjectReference,
    Single,
    String,
)
from emu.error import ParserError
from emu.lexer import Lexer
from emu.parser import Parser
from emu.utils import to_dict

from .test import assert_result, export_result, load_source


def test_api():
    """parser: API"""
    with assert_raises(RuntimeError):
        lexer = Lexer("")
        parser = Parser(lexer)
        parser.parse("i_dont_exist")


def test_expression():
    """parser: operator precedence parser"""
    codes = (
        "a * b * (c + ab) + -c",
        "- a ^ b.c.d + Me",
        "f(a.b.c + 5 - 12, b, c, d)",
        "g((a), (b), (c))",
        "h(a, b)",
        "a(1, 2)",
        "a!d!b!eb",
        "a(arg1, arg2).b",
        "a + b * c ^ d And e",
        "(a + b)(b)",
        "a.b(c)",
        "a(b).c",
        "a + b - c",
        "- a ^ b",
        "- a * b",
        "f()",
        "a().b().c",
        "a.f(b + c, d + e)",
        "(a)",
        "((((a))))",
    )

    asts = []
    for code in codes:
        lexer = Lexer(code)
        parser = Parser(lexer)
        asts.append(parser.parse("expression"))

    assert_result(
        "parser_expression", to_dict(asts, excluded_fields={"position"})
    )

    lexer = Lexer("")
    parser = Parser(lexer)
    assert_equals(None, parser.parse("expression"))


def test_statements():
    """parser: statements"""
    code = load_source("statements.vbs")
    lexer = Lexer(code)
    parser = Parser(lexer)
    parsed_block = parser.statement_block()

    assert_result(
        "statements", to_dict(parsed_block, excluded_fields={"position"})
    )


def test_module():
    """parser: module structure"""
    test_names = ("procedural_module", "class_module")

    for test_name in test_names:
        code = load_source(test_name + ".vbs")
        module = Parser(Lexer(code)).module()

        export_result(test_name, to_dict(module, excluded_fields={"position"}))
        assert_result(test_name, to_dict(module, excluded_fields={"position"}))


def test_module_exp():
    code = """Attribute VB_Name = "Zbop"
Attribute VB_Globalnamespace = False
Attribute VB_Exposed = True
    """
    lexer = Lexer(code)
    parser = Parser(lexer)

    print("\n")
    try:
        module = parser.module()
    except ParserError as e:
        error = e
        print(f"Error: {error}")
        print(f"Next token: {lexer.peek_token()}")
        import ipdb

        ipdb.set_trace()

    from pprint import pprint

    pprint(to_dict(module, excluded_fields={"position"}))

    if lexer.peek_token().category != lexer.peek_token().Category.END_OF_FILE:
        print(f"Next token: {lexer.peek_token()}")

        import ipdb

        ipdb.set_trace()


def test_string():
    """parser: string literals"""
    codes = '"hey"', '"""ho"""', '"let""s go"'
    values = "hey", '"ho"', 'let"s go'

    for code, value in zip(codes, values):
        lexer = Lexer(code)
        parser = Parser(lexer)

        ast = parser.parse("literal")
        assert_equals(type(ast), Literal)
        assert_equals(ast.variable.declared_type, String)
        assert_equals(type(ast.variable.value), String)
        assert_equals(ast.variable.value.value, value.encode("utf-16"))


def test_integer():
    """parser: integers"""
    codes = {
        "&hffff%": -1,
        "&H8000%": -32768,
        "&h7fff%": 0x7FFF,
        "&H0%": 0,
        "32767%": 32767,
        "0%": 0,
        "&o100000%": -32768,
        "&hffffffff&": -1,
        "&h0& ": 0,
        "12345&": 12345,
        "&hffff&": 0xFFFF,
        "&h7fff&": 0x7FFF,
        "&h5fff^": 0x5FFF,
        "&hffff^": 0xFFFF,
        "&h8fff^": 0x8FFF,
        "&h12FFFFFFF^": 0x1_2FFF_FFFF,
        "&h8000000000000000^": -0x8000_0000_0000_0000,
        "1234": 1234,
        "&h7788": 0x7788,
        "&h8700": -0x7900,
        "&h80000000": -0x8000_0000,
        "4886718345": 4886718345.0,
    }

    for code, value in codes.items():
        lexer = Lexer(code)
        parser = Parser(lexer)

        tree = parser.parse("literal")
        assert_equals(type(tree), Literal)
        assert_equals(tree.variable.value.value, value)

    error_codes = (
        "32768%",
        "&h10000%",
        "2147483648&",
        "&h100000000&",
        "9223372036854775808^",
        "&H10000000000000000^",
        "&h0123456789ABCDEF",
    )

    for code in error_codes:
        lexer = Lexer(code)
        parser = Parser(lexer)

        with assert_raises(ParserError):
            parser.parse("literal")


def test_float():
    """parser: float literals"""
    codes = ".12", ".13e-5!", "20D10#", "3.12e5@"
    types = Double, Single, Double, Currency
    values = 0.12, 1.3e-6, 2e11, 312e7

    for code, value_type, value in zip(codes, types, values):
        lexer = Lexer(code)
        parser = Parser(lexer)

        ast = parser.parse("literal")
        assert_equals(type(ast), Literal)
        assert_equals(ast.variable.declared_type, value_type)
        assert_equals(type(ast.variable.value), value_type)
        assert_equals(ast.variable.value.value, value)


def test_identifier():
    """parser: identifiers"""
    codes = "Plop", "z1246_bcb", "NothIng", "Null", "Empty", "fAlSe", "TrUe"
    node_types = Name, Name, Literal, Literal, Literal, Literal, Literal
    declared_types = None, None, ObjectReference, Null, Empty, Boolean, Boolean

    for code, node_type, decl_type in zip(codes, node_types, declared_types):
        lexer = Lexer(code)
        parser = Parser(lexer)

        ast = parser.parse("primary")
        assert_equals(type(ast), node_type)
        if node_type == Literal:
            assert_equals(ast.variable.declared_type, decl_type)
