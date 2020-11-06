from .test import load_source, assert_result

from emu.error import LexerError
from emu.lexer import Lexer
from emu.utils import to_dict

from nose.tools import assert_equals, assert_false, assert_raises, raises

CODE = load_source("lexer.vbs")


def test_token_operators():
    """
    lexer: token operators
    """
    lexer = Lexer("Hey Hey")
    token1 = lexer.pop_token()
    lexer.pop_token()
    token2 = lexer.pop_token()

    assert_equals(repr(token1), "<'Hey', IDENTIFIER>")
    assert_false(token1 == 1)
    assert_equals(token1, token2)

    token2.category = token2.Category.KEYWORD
    assert_false(token1 == token2)

    assert_equals(token1 + " Ho", "Hey Ho")

    with assert_raises(RuntimeError):
        token1 + token2

    token3 = Lexer("Hey").pop_token()

    with assert_raises(RuntimeError):
        token1 + token3


def test_tokens():
    """
    lexer: tokens

    Ensures the output tokens are as expected
    """
    lexer = Lexer(CODE, "code")
    tokens = to_dict(list(lexer.tokens()))

    assert_result("lexer", tokens)


def test_peek():
    """
    lexer: peek

    peek should not interfere with tokens
    """
    lexer = Lexer(CODE)
    peeked_tokens = []
    distance = 0

    while True:
        token = lexer.peek_token(distance)
        peeked_tokens.append(token)
        distance += 1

        if token.category == token.Category.END_OF_FILE:
            break

    with assert_raises(IndexError):
        lexer.peek_token(-1)

    tokens = list(lexer.tokens())

    assert_equals(peeked_tokens, tokens)


def test_pop():
    """
    lexer: pop

    pop should return the same tokens as tokens, and leave none for it
    """
    lexer = Lexer(CODE)
    poped_tokens = []

    while True:
        token = lexer.pop_token()
        poped_tokens.append(token)

        if token.category == token.Category.END_OF_FILE:
            break

    tokens = list(lexer.tokens())
    assert_equals(len(tokens), 1)
    assert_equals(tokens[0].category, tokens[0].Category.END_OF_FILE)

    lexer = Lexer(CODE)
    tokens = list(lexer.tokens())
    assert_equals(tokens, poped_tokens)


@raises(LexerError)
def test_error():
    """
    lexer: error handling
    """
    lexer = Lexer('Hey\r  "Non ending string\nI do not exist', "code")
    try:
        for token in lexer.tokens():
            pass
    except Exception as exception:
        msg = "code:2:3: Can't scan this line\n  \"Non ending string\n  ^\n"
        assert_equals(str(exception), msg)

    lexer = Lexer('"Non ending string', "code")
    for token in lexer.tokens():
        pass
