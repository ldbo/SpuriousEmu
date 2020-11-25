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

from .test import assert_result


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
    code = """
    Dim Shared variable
    Dim a As Integer, b, c%
    Static abcd As Long, def As String
    Const a = 1, b As Long = 12, c! = 30
    ReDim MyArray()
    ReDim Array1(), Array2(), Array3()
    ReDim Preserve Array1(), Array2(), Array3()
    ReDim arr(1, 2 To 12, 3)
    ReDim arr(12) As Integer
    ReDim array(12 To 15) As New MyClass
    Erase a
    Erase a, b, c
    Mid(str, 12, 3) = "abc"
    MidB$(name, 0) = str + "12"
    LSet b = "Hey"
    RSet a = "abcd"
    Let a = b.c() + 12
    Set a.b(12, 14) = 12
    On Error Goto abcd : On Error Resume Next
    On Error Goto 123 :
    Resume Next : Resume 12 : Resume abcd : Error abcd : Error 123
    Open hey As 12
    Open "hi my name is" For Binary Access Read Lock Read As (12 + "abcd") _
    Len = 1
    Open s Access Write Lock Read Write As f
    Open s Access Read Write Lock Read As f
    Open s Lock Write As #f
    Open s Shared As # f
    Open s As f Len=12
    Open s As f Len = 15
    Reset Close hey
    Reset
    Close abcd + 5, abcd(), efg
    Close
    Seek ab, 123 + 5
    Lock c
    Lock d, To b
    Lock f, a To 12 + b
    Unlock g, b
    Line Input #123, ab
    Width #145, abcd
    Print #1,
    Print #1, "This is a test"
    Print #1, "Zone 1",Tab; "Zone 2"
    Print #1, "Hello" ;" " , "World"
    Print #1, Spc(5); "5 leading spaces "
    Print #1, Tab(10) ; "Hello"
    Write #1, "Hello World", 234
    Write #1,
    Write #1, MyBool ; " is a Boolean value"
    Write #1, MyDate ; " is a date"
    Write #1, MyNull ; " is a null value"
    Write #1, MyError ; " is an error value"
    Input #1, MyString, MyNumber
    Put #1, RecordNumber, MyRecord
    Get #4,,FileBuffer
    Get #1, Position, MyRecord
    """
    lexer = Lexer(code)

    parser = Parser(lexer)

    print()
    block = parser.statement_block()

    from pprint import pprint

    for statement in block.statements:
        print(statement.position.body())
        pprint(to_dict(statement, excluded_fields={"position"}))
        print()

    if lexer.peek_token().category != lexer.peek_token().Category.END_OF_FILE:

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
