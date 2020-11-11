from typing import Callable, Iterable, Optional, Tuple

from .abstract_syntax_tree import *
from .data import *
from .error import ParserError
from .lexer import Lexer, Token


_TC = Token.Category
_OT = Operator.Type


class Parser:
    """
    VBA hybrid parser, based on a recursive descent for most of the grammar,
    except for expressions which are parsed using an operator precedence parser.

    Args:
      lexer: Lexer used as a token source by the parser
    """

    __lexer: Lexer  #: Lexer used for tokenization
    """
    Tokens that have already been consumed from  ``__source_tokens`` for
    peeking purpose.
    """
    __file_name: str
    """Name of the source code, found by peeking at the first token"""
    __file_content: str  #: Source code, found by peeking at the first token

    __last_operatorand: Optional[Union[Operator, Expr]]
    """Hold the last expression token seen, either an operator or an operand.
    Is ``None`` when not parsing an expression."""

    __FLOAT_TYPES = {
        None: Double,
        "!": Single,
        "#": Double,
        "@": Currency,
    }  #: Mapping between float suffix and value type

    __OPERATORS_PRECEDENCES = {
        _OT.RPAREN: -2,
        _OT.INDEX: -1,
        _OT.COMMA: -0,
        _OT.IMP: 1,
        _OT.EQV: 2,
        _OT.XOR: 3,
        _OT.OR: 4,
        _OT.AND: 5,
        _OT.NOT: 6,
        _OT.IS: 7,
        _OT.LIKE: 7,
        _OT.GE: 7,
        _OT.GT: 7,
        _OT.LT: 7,
        _OT.LE: 7,
        _OT.EQ: 7,
        _OT.NEQ: 7,
        _OT.CONCAT: 8,
        _OT.PLUS: 9,
        _OT.MINUS: 9,
        _OT.MOD: 10,
        _OT.INT_DIV: 11,
        _OT.MULT: 12,
        _OT.DIV: 12,
        _OT.UNARY_MINUS: 13,
        _OT.EXP: 14,
        _OT.DOT: 15,
        _OT.EXCLAMATION: 15,
        _OT.LPAREN: 15,
    }  #: Mapping between operators and precedence

    __OP_TYPE = {
        ")": _OT.RPAREN,
        "imp": _OT.IMP,
        "eqv": _OT.EQV,
        "xor": _OT.XOR,
        "or": _OT.OR,
        "and": _OT.AND,
        "not": _OT.NOT,
        "is": _OT.IS,
        "like": _OT.LIKE,
        "=>": _OT.GE,
        ">=": _OT.GE,
        "=<": _OT.LE,
        "<=": _OT.LE,
        ">": _OT.GT,
        "<": _OT.LT,
        "><": _OT.NEQ,
        "<>": _OT.NEQ,
        "=": _OT.EQ,
        "&": _OT.CONCAT,
        "+": _OT.PLUS,
        "mod": _OT.MOD,
        "\\": _OT.INT_DIV,
        "/": _OT.DIV,
        "*": _OT.MULT,
        "^": _OT.EXP,
        ".": _OT.DOT,
        "!": _OT.EXCLAMATION,
        ",": _OT.COMMA,
    }  #: Mapping between single-arity operators and their symbols

    __OP_ARITY_TYPE = {
        "(": {1: _OT.LPAREN, 2: _OT.INDEX},
        "-": {1: _OT.UNARY_MINUS, 2: _OT.MINUS},
    }  #: Mapping between multiple-arity operators symbols and their operator

    # API

    def __init__(self, lexer: Lexer) -> None:
        self.__lexer = lexer
        first_token = self.__peek_token(0)
        self.__file_name = first_token.position.file_name
        self.__file_content = first_token.position.file_content
        self.__last_operatorand = None

    def parse(self, rule: str) -> AST:
        """
        Performs a recursive descent on the tokens given on the constructor,
        starting on a given rule.

        Args:
          rule: Name of a parsing method

        Returns:
          An abstract syntax tree whose type corresponds to ``rule``

        Raises:
          :py:exc:`emu.error.ParserError`: If there is a syntax error in the
                                           source code
          :py:exc:`RuntimeError`: If ``rule`` does not correspond to an accepted
                                  rule

        Todo:
          * Add an attribute containing all the supported starting rules
        """
        try:
            parser = getattr(self, rule)
        except (AttributeError, TypeError):
            raise RuntimeError(f"Rule {rule} is not supported")

        try:
            result = parser()
        except ParserError:
            raise
        except Exception as e:
            msg = f"Got an unexpected error while parsing {rule}:\n"
            msg += repr(e)
            raise ParserError(msg)

        return result

    # Helper functions

    def __pop_token(self) -> Token:
        return self.__lexer.pop_token()

    def __peek_token(self, distance: int = 0) -> Token:
        return self.__lexer.peek_token(distance)

    def __category_match(self, category: _TC, distance: int = 0) -> bool:
        """
        Args:
          category: Category to match against
          distance: Peeking distance

        Returns:
          If the category of the token peeked at ``distance`` is ``category``

        Raises:
          :py:exc:`IndexError`: If ``distance < 0``
        """
        return self.__peek_token(distance).category == category

    def __categories_match(
        self, categories: Iterable[_TC], distance: int = 0
    ) -> bool:
        """
        Args:
          categories: Categories to match against
          distance: Peeking distance

        Returns:
          If the category of the token peeked at ``distance`` is in
          ``categories``

        Raises:
          :py:exc:`IndexError`: If ``distance < 0``
        """
        return any(
            self.__category_match(category, distance) for category in categories
        )

    def __try_rules(self, *rules: Callable[[], Optional[AST]]) -> Optional[AST]:
        """
        Returns:
          The first non-``None`` return value of one of the rules, or ``None``.
        """
        for rule in rules:
            tree = rule()
            if tree is not None:
                return tree

        return None

    # Rules

    def expression(self) -> Optional[Expr]:
        """Expression parser using the Shunting-Yard algorithm"""
        # Operator stack, is never empty until parsing is done, which is
        # obtained by surrounding the expression in parentheses
        operators: List[Operator] = [Operator(self.__peek_token(), _OT.LPAREN)]
        # Operands, begins empty but stays non-empty during all parsing
        operands: List[Expr] = []
        self.__last_operatorand = operators[-1]

        if self.__categories_match(
            (_TC.END_OF_STATEMENT, _TC.END_OF_FILE, _TC.SYMBOL)
        ):
            return None

        while True:
            if self.__categories_match(
                (_TC.END_OF_STATEMENT, _TC.END_OF_FILE, _TC.SYMBOL)
            ):
                self.__expression_closing_parenthesis(
                    operators, operands, self.__peek_token().position
                )
                break

            if self.BLANK():
                continue

            # Encounter operator
            operator = self.operator()
            if operator is not None:
                self.__expression_operator(operator, operators, operands)
                self.__last_operatorand = operator
                continue

            # Encounter primary expression
            operand = self.primary()
            if operand is None:
                if operands == []:
                    position = None
                else:
                    position = operands[-1].position
                msg = "Can't parse the expression, "
                msg += "expecting a primary expression"
                raise ParserError(msg, position)

            operands.append(operand)
            self.__last_operatorand = operand

        self.__last_operatorand = None
        if len(operands) != 1:
            msg = "Error while parsing the expression"
            if operands == []:
                position = self.__peek_token().position
            else:
                position = operands[0].position + operands[-1].position
            raise ParserError(msg, position)

        # We are left with a single parenthesized expression at the end
        paren_operand = operands.pop()
        assert isinstance(paren_operand, ParenExpr)
        return paren_operand.expr

    def __expression_closing_parenthesis(
        self, operators, operands, rparen_position: FilePosition
    ):
        """Handle a closing parenthesis during expression parsing: it is either
        to close an index expression or just after a parenthesized expression"""
        while operators[-1].op not in (_OT.LPAREN, _OT.INDEX):
            Parser.__expression_reduce_stacks(operators, operands)

        current_operands: Tuple[Expr, ...]
        if operators[-1].op == _OT.LPAREN:
            current_operands = (operands.pop(),)
        else:  # Index expression
            last_opand = self.__last_operatorand
            # Check for empty arguments list
            if isinstance(last_opand, Operator) and last_opand.op == _OT.INDEX:
                args = ArgList(last_opand, tuple())
            else:
                args = operands.pop()

            current_operands = (operands.pop(), args)

        lparen_or_index = operators.pop()
        new_operand = Operation(
            lparen_or_index.position + rparen_position,
            lparen_or_index,
            current_operands,
        )

        operands.append(self.__expression_post_process_operation(new_operand))

    def __expression_operator(
        self,
        operator: Operator,
        operators: List[Operator],
        operands: List[Expr],
    ) -> None:
        """Handle an operator during expression parsing"""
        # operators is never empty
        if operator.op == _OT.RPAREN:
            self.__expression_closing_parenthesis(
                operators, operands, operator.position
            )
        elif Parser.__op_eq(operators[-1], operator):
            operators.append(operator)
        elif Parser.__op_precedence(operators[-1]) < Parser.__op_precedence(
            operator
        ):
            operators.append(operator)
        elif operator.op in (_OT.NOT, _OT.UNARY_MINUS):
            operators.append(operator)
        else:
            while (
                Parser.__op_precedence(operators[-1])
                >= Parser.__op_precedence(operator)
                and operators[-1].op != _OT.LPAREN
            ):
                Parser.__expression_reduce_stacks(operators, operands)
            operators.append(operator)

    @staticmethod
    def __expression_reduce_stacks(
        operators: List[Operator], operands_stack: List[Expr]
    ) -> None:
        """
        Perform stacks reduction, consuming one multi-operand operation, and
        updating the operand stack. If there is a left parenthesis, skip it.
        """
        operands = [operands_stack.pop()]
        operator = operators[-1]

        if operator.op in (_OT.UNARY_MINUS, _OT.NOT):  # Local unary operator
            while Parser.__op_eq(operators[-1], operator):
                new_operand = Operation(
                    operators.pop().position + operands[0].position,
                    operator,
                    tuple(operands),
                )
                operands = [new_operand]
        elif operator.op == _OT.LPAREN:
            operands_stack.append(operands[0])

            return
        else:  # Binary operator
            while Parser.__op_eq(operators[-1], operator):
                operators.pop()
                operands.append(operands_stack.pop())

            new_operand = Operation(
                operands[0].position + operands[-1].position,
                operator,
                tuple(reversed(operands)),
            )

        pp_operand = Parser.__expression_post_process_operation(new_operand)
        operands_stack.append(pp_operand)

    @staticmethod
    def __expression_post_process_operation(operation: Operation) -> Expr:
        """Returns:
        An expression representing the same operation, but formatted in a
        more semantic way, easing later processing.
        """
        op_type = operation.operator.op
        if op_type in (_OT.DOT, _OT.EXCLAMATION):  # Access operators
            access_type = {_OT.DOT: MemberAccess, _OT.EXCLAMATION: DictAccess}[
                op_type
            ]

            expr = operation.operands[0]
            for field in operation.operands[1:]:
                if not isinstance(field, Name):
                    msg = "The last element of a member access expression must "
                    msg += "be a name"
                    raise ParserError(msg, field.position)

                access = access_type(
                    expr.position + field.position, expr, field
                )
                expr = access

            return access
        elif op_type == _OT.COMMA:  # Argument list
            args = tuple(Arg(expr, expr) for expr in operation.operands)
            return ArgList(operation, args)
        elif op_type == _OT.INDEX:  # Index expression
            assert len(operation.operands) == 2
            u_arg_list = operation.operands[1]
            if isinstance(u_arg_list, ArgList):
                arg_list = u_arg_list
            else:
                arg_list = ArgList(u_arg_list, (Arg(u_arg_list, u_arg_list),))
            return IndexExpr(operation, operation.operands[0], arg_list)
        elif op_type == _OT.LPAREN:  # Parenthesized expression
            assert len(operation.operands) == 1
            return ParenExpr(operation, operation.operands[0])
        else:
            return operation

    @staticmethod
    def __op_eq(operator1: Operator, operator2: Operator):
        """Returns:
        ``True`` if both operators share the same name and precedence.
        """
        return operator1.op == operator2.op

    @staticmethod
    def __op_precedence(operator: Operator) -> int:
        """Returns:
        The precedence of the operator, the greater the more precedence it has
        """
        return Parser.__OPERATORS_PRECEDENCES[operator.op]

    def operator(self) -> Optional[Operator]:
        """Parse an operator, discriminating between unary and binary operators.

        Warning:
          Should only be called during expression parsing
        """
        if self.__category_match(_TC.OPERATOR):
            token = self.__pop_token()

            try:
                op = self.__OP_TYPE[token]
            except KeyError:  # Happens only for binary-and-unary operators
                if self.__last_operatorand is None:
                    arity = 1
                elif isinstance(self.__last_operatorand, Expr):
                    arity = 2
                elif self.__last_operatorand.op == _OT.RPAREN:
                    arity = 2
                else:  # Operator which is not a right parenthesis
                    arity = 1
                op = self.__OP_ARITY_TYPE[token][arity]

            return Operator(token, op)

        return None

    def primary(self) -> Optional[Expr]:
        return self.__try_rules(  # type: ignore [return-value]
            self.literal, self.NAME, self.ME
        )

    def literal(self) -> Optional[Literal]:
        return self.__try_rules(  # type: ignore [return-value]
            self.INTEGER, self.FLOAT, self.STRING, self.literal_identifier
        )

    def literal_identifier(self) -> Optional[Literal]:
        parsers = (self.TRUE, self.FALSE, self.EMPTY, self.NULL, self.NOTHING)
        for parser in parsers:
            terminal = parser()
            if terminal is not None:
                return terminal

        return None

    def l_expression(
        self,
    ) -> Optional[
        Union[
            Name,
            DictAccess,
            MemberAccess,
            IndexExpr,
            WithDictAccess,
            WithMemberAccess,
        ]
    ]:
        expression = self.expression()

        if expression is None:
            return None
        else:
            if isinstance(
                expression,
                (
                    Name,
                    DictAccess,
                    MemberAccess,
                    IndexExpr,
                    WithDictAccess,
                    WithMemberAccess,
                ),
            ):
                return expression
            else:
                raise ParserError(
                    "Expecting a l-expression", expression.position
                )

    # Literals

    def BLANK(self) -> bool:
        if self.__category_match(_TC.BLANK):
            self.__pop_token()
            return True
        else:
            return False

    def INTEGER(self) -> Optional[Literal]:
        """[MS-VBA]_ 3.3.2: Number tokens"""
        if not self.__category_match(_TC.INTEGER):
            return None

        token = self.__pop_token()
        base, raw_number, suffix = self.__split_integer(token)

        try:
            variable = _INTEGER_BUILDERS[suffix](base == 10, raw_number)
        except ValueError as e:
            raise ParserError(e.args[0], token.position)

        literal = Literal(token, variable)

        return literal

    @staticmethod
    def __split_integer(token: Token) -> Tuple[int, int, Optional[str]]:
        """Returns:
        3-tuple containing

        - the base of the integer, either ``8``, ``10`` or ``16``
        - the raw value of the number, before processing the eventual sign bit
        - the optional suffix of the integer, either ``None``, ``"%"``,
          ``"&"`` or ``"^"``
        """
        if token[0] == "&":
            if token[1] in "oO":
                base = 0o10
            else:  # token[1] in 'hH'
                base = 0x10
            digits = token[2:]
        else:
            base = 10
            digits = token

        suffix: Optional[str]
        if digits[-1] in "%&^":
            suffix = digits[-1]
            digits = digits[:-1]
        else:
            suffix = None

        raw_number = int(digits, base)

        return base, raw_number, suffix

    @staticmethod
    def __integer_value_type(decimal: bool, raw_number: int) -> Type[Value]:
        """Returns:
        The value type corresponding to a given raw number
        """
        decimal_shift = 1 if decimal else 0
        if raw_number <= (0xFFFF >> decimal_shift):  # 16-bit
            return Integer
        elif raw_number <= (0xFFFF_FFFF >> decimal_shift):  # 32-bit
            return Long
        else:  # 64-bit
            return LongLong

    @staticmethod
    def __build_number_no_suffix(decimal: bool, raw_number: int) -> Variable:
        """Determine integer width, or if double, with value"""
        value_type = Parser.__integer_value_type(decimal, raw_number)

        if value_type == LongLong:  # Either decimal double or error
            if decimal:
                value_type = Double
                number = float(raw_number)
            else:
                msg = "Hexadecimal or octal literal integer without suffix "
                msg += "needs a 32-bit value"
                raise ValueError(msg)
        else:  # Check if negative value or not
            if decimal:
                number = raw_number
            elif value_type == Integer and raw_number & (1 << 15):
                number = raw_number - 0x1_0000
            elif value_type == Long and raw_number & (1 << 31):
                number = raw_number - 0x1_0000_0000
            else:
                number = raw_number

        return Variable(value_type, value_type(number))

    @staticmethod
    def __build_integer(decimal: bool, raw_number: int) -> Variable:
        """Declared 16-bit integer: % suffix"""
        if decimal and raw_number > 0x7FFF:
            msg = "Decimal literal integer with % suffix needs 15-bit value"
            raise ValueError(msg)
        elif raw_number > 0xFFFF:
            msg = "Hexadecimal or octal literal integer with % suffix "
            msg += "needs 16-bit value"
            raise ValueError(msg)

        negative = raw_number & (1 << 15)
        if negative:
            raw_number = raw_number - 0x1_0000

        return Variable(Integer, Integer(raw_number))

    @staticmethod
    def __build_long(decimal: bool, raw_number: int) -> Variable:
        """Declared 32-bit integer: & suffix"""
        if decimal and raw_number > 0x7FFF_FFFF:
            msg = "Decimal literal integer with & suffix needs 31-bit value"
            raise ValueError(msg)
        elif raw_number > 0xFFFF_FFFF:
            msg = "Hexadecimal or octal literal integer with & suffix "
            msg += "needs 32-bit value"
            raise ValueError(msg)

        value_type: Type[Value]
        if raw_number <= 0x7FFF:
            value_type = Integer
        else:
            value_type = Long

        negative = raw_number & (1 << 31)
        if negative:
            raw_number = raw_number - 0x1_0000_0000

        value = value_type(raw_number)

        return Variable(Long, value)

    @staticmethod
    def __build_longlong(decimal: bool, raw_number: int) -> Variable:
        """Declared 64-bit integer: ^ suffix"""
        if decimal and raw_number > 0x7FFF_FFFF_FFFF_FFFF:
            msg = "Decimal literal integer with ^ suffix needs 63-bit value"
            raise ValueError(msg)
        elif raw_number > 0xFFFF_FFFF_FFFF_FFFF:
            msg = "Hexadecimal or octal literal integer with ^ suffix "
            msg += "needs 64-bit value"
            raise ValueError(msg)

        value_type: Type[Value]
        if raw_number <= 0x7FFF:
            value_type = Integer
        elif raw_number <= 0x7FFF_FFFF:
            value_type = Long
        else:
            value_type = LongLong

        negative = raw_number & (1 << 63)
        if negative:
            raw_number = raw_number - 0x1_0000_0000_0000_0000

        value = value_type(raw_number)

        return Variable(LongLong, value)

    def FLOAT(self) -> Optional[Literal]:
        """[MS-VBA]_ 3.3.2: number tokens"""
        if not self.__category_match(_TC.FLOAT):
            return None

        token = self.__pop_token()
        number, suffix = self.__split_float(token)

        float_type = self.__FLOAT_TYPES[suffix]
        value: Value
        if float_type == Currency:
            value = Currency(int(number * 10 ** 4))
        else:
            value = float_type(number)  # type: ignore [arg-type]

        return Literal(token, Variable(float_type, value))

    @staticmethod
    def __split_float(token: Token) -> Tuple[float, Optional[str]]:
        """Returns:
        A 2-tuple containing

         - The float value of the number
         - The suffix, among ``!``, ``#``, ``@`` and ``None``
        """
        if token[-1] in "!#@":
            suffix = token[-1]  # type: Optional[str]
            number = token[:-1]
        else:
            suffix = None
            number = str(token)

        exp_separator_position = max(number.find(sep) for sep in "eEdD")
        if exp_separator_position == -1:
            exponent = 0
        else:
            exponent = int(number[exp_separator_position + 1 :], 10)
            number = number[:exp_separator_position]

        return float(number) * 10 ** exponent, suffix

    def STRING(self) -> Optional[Literal]:
        """Escape the quotes in the token"""
        if not self.__category_match(_TC.STRING):
            return None

        token = self.__pop_token()
        string = token[1:-1].replace('""', '"')

        return Literal(token, Variable(String, String.from_str(string)))

    def TRUE(self) -> Optional[Literal]:
        if not self.__category_match(_TC.BOOLEAN):
            return None

        token = self.__peek_token()
        if token == "True":
            self.__pop_token()
            return Literal(token, Variable(Boolean, Boolean(True)))
        else:
            return None

    def FALSE(self) -> Optional[Literal]:
        if not self.__category_match(_TC.BOOLEAN):
            return None

        token = self.__peek_token()
        if token == "False":
            self.__pop_token()
            return Literal(token, Variable(Boolean, Boolean(False)))
        else:
            return None

    def EMPTY(self) -> Optional[Literal]:
        if not self.__category_match(_TC.VARIANT):
            return None

        token = self.__peek_token()
        if token == "Empty":
            self.__pop_token()
            return Literal(token, Variable(Empty))

        return None

    def NULL(self) -> Optional[Literal]:
        if not self.__category_match(_TC.VARIANT):
            return None

        token = self.__peek_token()
        if token == "Null":
            self.__pop_token()
            return Literal(token, Variable(Null))

        return None

    def NOTHING(self) -> Optional[Literal]:
        if not self.__category_match(_TC.OBJECT):
            return None

        token = self.__peek_token()
        if token == "Nothing":
            self.__pop_token()
            return Literal(token, Variable(ObjectReference))

        return None

    def NAME(self) -> Optional[Name]:
        if self.__category_match(_TC.IDENTIFIER):
            token = self.__pop_token()
            return Name(token, str(token))

        return None

    def ME(self) -> Optional[Name]:
        token = self.__peek_token()
        if token.category == _TC.KEYWORD and token == "Me":
            self.__pop_token()
            return Name(token, str(token))

        return None


_INTEGER_BUILDERS = {
    None: Parser._Parser__build_number_no_suffix,  # type: ignore [attr-defined]
    "%": Parser._Parser__build_integer,  # type: ignore [attr-defined]
    "&": Parser._Parser__build_long,  # type: ignore [attr-defined]
    "^": Parser._Parser__build_longlong,  # type: ignore [attr-defined]
}  #: Mapping of integer suffix and integer builder, based on the declared type
