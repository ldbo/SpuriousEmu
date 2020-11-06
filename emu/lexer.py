from collections import deque
from enum import Enum
from hashlib import md5
from typing import Callable, Deque, Generator, List, Set, Tuple, Union

import re

from .error import LexerError
from .utils import FilePosition


class Token(str):
    """
    Token returned by the :py:class:`Lexer` class. It behaves as an improved
    ``str``, with added information about the token category and its position
    in the scanned text stream.

    Attributes:
      category(:class:`Token.Category`): Category of the token
      position(:class:`emu.utils.FilePosition`): Position of the token in the
                                                 source stream

    Arguments:
      value: Object used to initialize the internal ``str``
      category: Type of the token
      position: Position of the token in the input stream
    """

    category: "Token.Category"  #: Category of the token
    position: FilePosition  #: Position of the token in the stream

    __slots__ = ("category", "position")

    class Category(Enum):
        """Enumeration of the different token types"""

        # First pass tokens
        START_OF_FILE = 0  #: Start of the file, used with empty content
        END_OF_FILE = 1  #: End of file, used with empty content
        END_OF_STATEMENT = 2  #: End of a statement
        BLANK = 3  #: Blank character, excluding end of statement
        SYMBOL = 4  #: General one or two-character token (``!``, ``<=``, ...)
        COMMENT = 5  #: Comment, including ``Rem`` or ``'``
        INTEGER = 6  #: Integer number
        FLOAT = 7  #: Floating point number
        STRING = 8  #: String, including enclosing quotes
        IDENTIFIER = 9  #: Identifier

        # Second pass tokens
        KEYWORD = 10  #: Reserved word
        OPERATOR = 11  #: Operator
        BOOLEAN = 12  #: Boolean constant (``True``, ``False``)
        VARIANT = 13  #: Variant (``Empty``, ``Null``)
        OBJECT = 14  #: Object (``Nothing``)

    def __new__(
        cls, value: object, category: "Token.Category", position: FilePosition
    ) -> "Token":
        self = super(Token, cls).__new__(cls, value)  # type: ignore [call-arg]
        self.category = category
        self.position = position

        return self

    def __repr__(self) -> str:
        return f"<{super().__repr__()}, {self.category.name}>"

    def __eq__(self, token: object) -> bool:
        """
        Lowercase extended comparison

        Returns:
          ``True`` only if token is a :py:class:`str` with the same
          case-insensitive value as the token, or if it is a :py:class:`Token`
          that has the same case-insensitive value and category
        """
        if not isinstance(token, str):
            return False

        if isinstance(token, Token) and token.category != self.category:
            return False

        return self.lower() == token.lower()

    def __hash__(self) -> int:
        """
        Returns:
          Hash of the lowercase string
        """
        return self.lower().__hash__()

    def __add__(self, token: str) -> Union["Token", str]:
        """
        Concatenate a token with an adjacent string or token.

        Returns:
          :py:class:`Token`: If ``token`` is a :py:class:`Token`, concatenate
          the two token and keep the first ``category`` field.
          str: Else, just the concatenated string

        Raises:
          :py:exc:`RuntimeError`: If ``token`` is a :py:class:`Token` not
          located just after the current token, in the same file
        """
        if not isinstance(token, Token):
            return str(self) + token

        if self.position.file_name != token.position.file_name:
            raise RuntimeError("Can't concatenate tokens from different files")

        if self.position.end_index != token.position.start_index:
            raise RuntimeError("Can't concatenate non-adjacent tokens")

        value = str(self) + token

        category = self.category

        file_name = self.position.file_name
        file_content = self.position.file_content
        start_index = self.position.start_index
        end_index = token.position.end_index
        start_line = self.position.start_line
        end_line = token.position.end_line
        start_column = self.position.start_column
        end_column = token.position.end_column
        position = FilePosition(
            file_name,
            file_content,
            start_index,
            end_index,
            start_line,
            end_line,
            start_column,
            end_column,
        )

        return Token(value, category, position)


_TC = Token.Category

# First-pass tokenization

# Type of a token scanner: it takes a stream and an index, and return the index
# of the end of the starting token
Scanner = Callable[[str, int], int]


def _build_scanner(regex: str) -> Scanner:
    """Build a regex-based ``Scanner``.

    Args:
      - regex: Pattern matched by the scanner
    """
    pattern = re.compile(regex, re.DOTALL | re.IGNORECASE)

    def scanner(stream: str, index: int) -> int:
        match = pattern.match(stream, index)

        if match is None:
            raise RuntimeError("Failed to scan regex")
        else:
            return match.span()[1]

    return scanner


# Regular expressions used to recognize the first-pass tokens, case-insensitive
TOKEN_REGEXS: Tuple[Tuple[str, _TC], ...] = (
    (r"$", _TC.END_OF_FILE),
    (
        "(\t| |\x19|\u3000)+"
        "(_(\t| |\x19|\u3000)*(\r\n|\t|\n|\u2028|\u2029)(\t| |\x19|\u3000)*)?",
        _TC.BLANK,
    ),
    (r"[0-9]+[!#@]", _TC.FLOAT),
    (r"[0-9]+(\.[0-9]*)?[de][+-]?[0-9]+[!#@]?", _TC.FLOAT),
    (r"\.[0-9]+([de][+-]?[0-9]+)?[!#@]?", _TC.FLOAT),
    (r"([0-9]+|&[o]?[0-7]+|&[h][0-9a-fA-F]+)[%&\^]?", _TC.INTEGER),
    (r'"([^"]|"")*"', _TC.STRING),
    (
        r"(,|\.|!|\?|#|\(|\)|\[|]|:|;|%|&|\^|@|\$|:=|\+|-|\*|/|\\|"
        r"<=|=<|>=|=>|<|>|=|<>|><)",
        _TC.SYMBOL,
    ),
    ("(\r\n|\r|\n|\u2028|\u2029|:)", _TC.END_OF_STATEMENT),
    (r"Rem( [^\n]*)?(?=\n)", _TC.COMMENT),
    (r"[a-zA-Z][a-zA-Z0-9_]*", _TC.IDENTIFIER),
    (r"'[^\n]*(?=\n)", _TC.COMMENT),
)

# First-pass scanners, used by Lexer.__first_pass
_TOKEN_SCANNERS: List[Tuple[Scanner, _TC]] = list(
    map(lambda t: (_build_scanner(t[0]), t[1]), TOKEN_REGEXS)
)


# Second-pass tokenization (all the sets must be lowercase)

# Set of VBA keywords
KEYWORDS: Set[str] = {
    "access",
    "addressof",
    "alias",
    "and",
    "any",
    "append",
    "as",
    "attribute",
    "base",
    "binary",
    "boolean",
    "byref",
    "byte",
    "byval",
    "call",
    "case",
    "class_initialize",
    "class_terminate",
    "close",
    "compare",
    "const",
    "currency",
    "date",
    "declare",
    "defbool",
    "defbyte",
    "defcur",
    "defdate",
    "defdbl",
    "defint",
    "deflng",
    "deflnglng",
    "deflngptr",
    "defobj",
    "defsng",
    "defstr",
    "defvar",
    "dim",
    "do",
    "double",
    "each",
    "else",
    "elseif",
    "empty",
    "end",
    "endif",
    "enum",
    "eqv",
    "erase",
    "error",
    "event",
    "exit",
    "explicit",
    "false",
    "for",
    "friend",
    "function",
    "get",
    "global",
    "go",
    "gosub",
    "goto",
    "if",
    "imp",
    "implements",
    "in",
    "input",
    "integer",
    "is",
    "len",
    "let",
    "lib",
    "like",
    "line",
    "lineinput",
    "lock",
    "long",
    "longlong",
    "longptr",
    "loop",
    "lset",
    "me",
    "mid",
    "midb",
    "mod",
    "module",
    "new",
    "next",
    "not",
    "nothing",
    "null",
    "object",
    "on",
    "open",
    "option",
    "optional",
    "or",
    "output",
    "paramarray",
    "preserve",
    "print",
    "private",
    "property",
    "ptrsafe",
    "public",
    "put",
    "raiseevent",
    "random",
    "read",
    "redim",
    "rem",
    "reset",
    "resume",
    "return",
    "rset",
    "seek",
    "select",
    "set",
    "shared",
    "single",
    "spc",
    "static",
    "step",
    "stop",
    "string",
    "sub",
    "tab",
    "text",
    "then",
    "to",
    "true",
    "type",
    "typeof",
    "unlock",
    "until",
    "variant",
    "vb_base",
    "vb_control",
    "vb_creatable",
    "vb_customizable",
    "vb_description",
    "vb_exposed",
    "vb_ext_key",
    "vb_globalnamespace",
    "vb_helpid",
    "vb_invoke_func",
    "vb_invoke_property",
    "vb_invoke_propertyput",
    "vb_invoke_propertyputref",
    "vb_memberflags",
    "vb_name",
    "vb_predeclaredid",
    "vb_procdata",
    "vb_templatederived",
    "vb_usermemid",
    "vb_vardescription",
    "vb_varhelpid",
    "vb_varmemberflags",
    "vb_varprocdata",
    "vb_varusermemid",
    "wend",
    "while",
    "width",
    "with",
    "withevents",
    "write",
    "xor",
}

# New category of some keywords
SPECIALIZED_KEYWORDS = (
    ({"like", "is", "not", "and", "or", "xor", "eqv", "imp"}, _TC.OPERATOR),
    ({"false", "true"}, _TC.BOOLEAN),
    ({"empty", "null"}, _TC.VARIANT),
    ({"nothing"}, _TC.OBJECT),
)

# Symbols that are also operators
SYMBOL_OPERATORS = {
    "+",
    "-",
    "*",
    "/",
    "\\",
    "^",
    "&",
    "=",
    "<>",
    "><",
    "<",
    ">",
    "<=",
    "=<",
    ">=",
    "=>",
    "(",
    ")",
    ".",
    "!",
    ",",
}


class Lexer:
    """
    Two-pass VBA lexer.

    Args:
      - stream: Input file
      - stream_name: Name of the file, replaced by the md5 sum of its content
                     if empty

    Todo:
      - Hide internal attributes
    """

    stream: str
    stream_name: str
    next_tokens: Deque[Token]
    last_token: Token
    token_start: int
    index: int
    line: int
    column: int

    def __init__(self, stream: str, stream_name: str = "") -> None:
        self.stream = stream
        self.stream_name = stream_name
        if stream_name == "":
            self.stream_name = md5(stream.encode("utf8")).hexdigest()

        self.next_tokens = deque()
        last_position = FilePosition(stream_name, stream, 0, 0, 1, 1, 0, 1)
        self.last_token = Token("", _TC.START_OF_FILE, last_position)
        self.token_start = 0
        self.index = 0
        self.line = 1
        self.column = 1

    def tokens(self) -> Generator[Token, None, None]:
        """
        Yields:
          The tokens of the stream, including a unique
          :py:attr:`Token.Category.END_OF_FILE` one at the end.

        Raises:
          :py:exc:`LexerError`: If the file is malformed
        """
        stop = False

        while not stop:
            token = self.pop_token()
            yield token
            stop = token.category == _TC.END_OF_FILE

    def pop_token(self) -> Token:
        """
        Returns:
          The next token scanned from the input stream, and progress in the
          stream. Once the end is reached, returns continuously an empty
          :py:attr:`Token.Category.END_OF_FILE` token.

        Raises:
          :py:exc:`LexerError`: If the next token is malformed
        """
        if len(self.next_tokens) == 0:
            self.__second_pass()

        return self.next_tokens.popleft()

    def peek_token(self, distance: int = 0) -> Token:
        """
        Returns:
          The token ``distance`` ahead from the current position in the stream,
          without consuming it. For tokens after the end of the stream, return
          an empty :py:attr:`Token.Category.END_OF_FILE` token.

        Args:
          distance: Index of the token to peek with respect to the current
                    index in the stream, must be non-negative

        Raises:
          :py:exc:`LexerError`: If one of the ``distance`` next tokens is
                                  malformed
          :py:exc:`RuntimeError`: If ``distance < 0``
        """
        if distance < 0:
            raise IndexError("Can't peek a token at a negative distance.")

        while len(self.next_tokens) < distance + 1:
            self.__second_pass()

        return self.next_tokens[distance]

    def __second_pass(self) -> None:
        """
        Set the correct category of the current token output by
        ``__first_pass.``, and update :py:attr:`self.next_tokens`, adding one
        or two tokens.

        Raises:
          :py:exc:`LexerError`: If the next token is malformed
        """
        token = self.__first_pass()

        if token.category == _TC.IDENTIFIER and token in KEYWORDS:
            matched = False
            for keywords, category in SPECIALIZED_KEYWORDS:
                if token in keywords:
                    token.category = category
                    matched = True
                    break

            if matched:
                pass
            elif token in ("Mid", "MidB"):
                next_token = self.__first_pass()
                if next_token == "$":
                    token = token + next_token  # type: ignore [assignment]
                    self.next_tokens.append(token)
                else:
                    self.next_tokens.append(token)
                    self.next_tokens.append(next_token)
                return
            else:
                token.category = _TC.KEYWORD
        elif token.category == _TC.SYMBOL and token in SYMBOL_OPERATORS:
            token.category = _TC.OPERATOR

        self.next_tokens.append(token)

    def __first_pass(self) -> Token:
        """
        Try the different regex-based scanners found in _TOKEN_SCANNERS until
        one of them succeeds in scanning the stream at ``self.index``.

        Returns:
          The first matched token, using the category of the associated
          succeeding scanner

        Raises:
          :py:exc:`LexerError`: If the next token is malformed
        """
        for scan, category in _TOKEN_SCANNERS:
            try:
                # This first line may raise a RuntimeError
                end_index = scan(self.stream, self.index)

                position = FilePosition.from_indices(
                    self.stream_name,
                    self.stream,
                    self.index,
                    end_index,
                    self.line,
                    self.column,
                )
                token = Token(
                    self.stream[self.index : end_index], category, position
                )
                self.index = position.end_index
                self.line = position.end_line
                if token == "" or token[-1] not in position.IEOLS:
                    self.column = position.end_column
                else:
                    self.column = 1
                return token
            except RuntimeError:
                continue

        name = self.stream_name
        stream = self.stream
        start_index = self.index
        end_index = start_index + 1
        start_line = self.line
        start_column = self.column
        position = FilePosition.from_indices(
            name, stream, start_index, end_index, start_line, start_column,
        )

        raise LexerError("Can't scan this line", position)
