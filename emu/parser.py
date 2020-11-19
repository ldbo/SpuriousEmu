from functools import wraps
from typing import (
    Callable,
    Iterable,
    List,
    Optional,
    Tuple,
    Type,
    TypeVar,
    Union,
)

from .abstract_syntax_tree import (
    AST,
    Arg,
    ArgList,
    Close,
    DictAccess,
    Error,
    Expression,
    Get,
    IndexExpr,
    Input,
    LineInput,
    Literal,
    Lock,
    MemberAccess,
    Name,
    OnError,
    Open,
    Operation,
    Operator,
    OutputItem,
    ParenExpr,
    Print,
    Put,
    Resume,
    Seek,
    Statement,
    StatementBlock,
    StmtLabel,
    Unlock,
    Width,
    WithDictAccess,
    WithMemberAccess,
    Write,
)
from .data import (
    Boolean,
    Currency,
    Double,
    Empty,
    Integer,
    Long,
    LongLong,
    Null,
    ObjectReference,
    Single,
    String,
    Value,
    Variable,
)
from .error import ParserError
from .lexer import Lexer, Token
from .utils import FilePosition

_TC = Token.Category
_OT = Operator.Type
_T = TypeVar("_T")
_Rule = Callable[["Parser"], _T]  #: Type of a parsing rule


def _with_backtracking(rule: _Rule) -> _Rule:
    """
    Decorator allowing to have a backtrack-enabled rule: if it fails, the lexer
    state remains unchanged. Should be used as little as possible, ie. only
    when using lookaheads is not an acceptable solution.

    Todo:
      * Improve typehint to `Optional[_T]` where T is a `Type[AST]`
    """

    @wraps(rule)
    def wrapper(self: "Parser") -> _T:
        lexer = self._Parser__lexer  # type: ignore [attr-defined]
        # See https://github.com/python/mypy/issues/8267
        lexer.save_backtracking_point()

        try:
            result = rule(self)
        except ParserError:
            lexer.backtrack()
            raise

        if result is None:
            lexer.backtrack()
        else:
            lexer.dump_backtracking_point()

        return result

    return wrapper


class Parser:
    """
    VBA hybrid parser, based on a recursive descent for most of the grammar,
    except for expressions which are parsed using an operator precedence parser.

    Rules and terminals are parsed using recursive methods which must conform
    to the following interface:

      - don't accept any argument
      - if parsing is successful, return an instance of a subclass of
        :class:`AST`, and the lexer must be reading the token following the
        last token of the rule or terminal
      - if parsing fails, return ``None``, and lexer must be reading the first
        token of the rule or terminal
      - if parsing fails because of a source code error, a ParserError must be
        raised

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

    __last_operatorand: Optional[Union[Operator, Expression]]
    """Hold the last expression token seen, either an operator or an operand.
    Is ``None`` when not parsing an expression."""
    __index_expressions: int
    """Only used during expression parsing, store the number of nested index
    expressions"""
    __paren_expressions: int

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
        self.__index_expressions = 0
        self.__paren_expressions = 0

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

    def __pop_tokens(self, number: int) -> Tuple[Token, ...]:
        return tuple(self.__pop_token() for _ in range(number))

    def __peek_token(self, distance: int = 0) -> Token:
        return self.__lexer.peek_token(distance)

    def __peek_tokens(self, *distances: int) -> Tuple[Token, ...]:
        return tuple(self.__peek_token(distance) for distance in distances)

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

    def __category_matches(self, category: _TC, *distances: int) -> bool:
        return all(
            token.category == category
            for token in self.__peek_tokens(*distances)
        )

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

        Todo:
          * Improve typehints to something like Callable[Tx], ... ->
            Union[Tx, ...]
        """
        for rule in rules:
            tree = rule()
            if tree is not None:
                return tree

        return None

    def __craft_error(self, msg: str) -> ParserError:
        return ParserError(msg, self.__peek_token().position)

    # Rules

    # Expression

    @_with_backtracking
    def expression(self) -> Optional[Expression]:  # TODO handle errors
        """Expression parser using the Shunting-Yard algorithm"""
        # Operator stack, is never empty until parsing is done, which is
        # obtained by surrounding the expression in parentheses
        operators: List[Operator] = [Operator(self.__peek_token(), _OT.LPAREN)]
        # Operands, begins empty but stays non-empty during all parsing
        operands: List[Expression] = []
        self.__last_operatorand = operators[-1]
        self.__index_expressions = 0
        self.__paren_expressions = 0

        self.__expression_shunting_yard(operators, operands)

        self.__last_operatorand = None

        if len(operands) == 0:
            return None
        elif len(operands) > 1:
            msg = "Error while parsing the expression"
            position = operands[0].position + operands[-1].position
            raise ParserError(msg, position)

        # We are left with a single parenthesized expression at the end
        paren_operand = operands.pop()
        assert isinstance(paren_operand, ParenExpr)
        return paren_operand.expr

    def __expression_shunting_yard(
        self, operators: List[Operator], operands: List[Expression]
    ) -> None:
        while True:
            if self.__categories_match(
                (_TC.END_OF_STATEMENT, _TC.END_OF_FILE, _TC.SYMBOL)
            ):
                break
            elif self.BLANK():
                continue

            # Encounter operator
            operator = self.operator()
            if operator is not None:
                self.__expression_operator(operator, operators, operands)
                self.__last_operatorand = operator
                continue

            # Encounter primary expression
            if (len(operators), len(operands)) == (1, 1):
                break

            operand = self.primary()
            if operand is not None:
                operands.append(operand)
                self.__last_operatorand = operand
                continue

            # Can't parse anymore: stop
            break

        self.__expression_closing_parenthesis(
            operators, operands, self.__peek_token().position
        )

    def __expression_closing_parenthesis(
        self, operators, operands, rparen_position: FilePosition
    ):
        """Handle a closing parenthesis during expression parsing: it is either
        to close an index expression or just after a parenthesized expression"""
        if len(operands) == 0:
            return

        while operators[-1].op not in (_OT.LPAREN, _OT.INDEX):
            Parser.__expression_reduce_stacks(operators, operands)

        current_operands: Tuple[Expression, ...]
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
            self.__index_expressions -= 1

        self.__paren_expressions -= 1

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
        operands: List[Expression],
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
        operators: List[Operator], operands_stack: List[Expression]
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
    def __expression_post_process_operation(operation: Operation) -> Expression:
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
            if self.__peek_token() == "," and self.__index_expressions == 0:
                return None

            if self.__peek_token() == ")" and self.__paren_expressions == 0:
                return None

            token = self.__pop_token()

            try:
                op = self.__OP_TYPE[token]
            except KeyError:
                op = self.__operator_multi_ary(token)

            if op == _OT.INDEX:
                self.__index_expressions += 1
                self.__paren_expressions += 1
            elif op == _OT.LPAREN:
                self.__paren_expressions += 1

            return Operator(token, op)

        return None

    def __operator_multi_ary(self, token: Token) -> Operator.Type:
        """Determine the arity of a binary-and-unary operator."""
        if self.__last_operatorand is None:
            arity = 1
        elif isinstance(self.__last_operatorand, Expression):
            arity = 2
        elif self.__last_operatorand.op == _OT.RPAREN:
            arity = 2
        else:  # Operator which is not a right parenthesis
            arity = 1

        return self.__OP_ARITY_TYPE[token][arity]

    def primary(self) -> Optional[Expression]:
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

        if expression is None or not isinstance(
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
            return None
        else:
            return expression

    def integer_expression(self) -> Optional[Expression]:
        return self.expression()

    def variable_expression(self) -> Optional[Expression]:
        return self.expression()

    def bound_variable_expression(self) -> Optional[Expression]:
        return self.expression()

    # Statements (individual statements do not include the END_OF_STATEMENT
    # token)

    def statement_block(self) -> Optional[StatementBlock]:
        ok = True
        statements = []
        self.EOS()  # Skip newlines at the beginning
        while ok:
            statement = self.block_statement()
            if statement is None:
                ok = False
            else:
                statements.append(statement)

        if statements == []:
            position = self.__peek_token().position
        else:
            position = statements[0].position + statements[-1].position

        return StatementBlock(position, tuple(statements))

    def block_statement(self) -> Optional[Statement]:  # TODO
        return self.__try_rules(self.statement)  # type: ignore

    def statement(self) -> Optional[Statement]:  # TODO
        statement = self.__try_rules(
            self.error_handling_statement, self.file_statement
        )
        if statement is None or not self.EOS():
            return None

        return statement  # type: ignore [return-value]

    def statement_label(self) -> Optional[StmtLabel]:
        return self.__try_rules(self.NAME, self.literal)  # type: ignore

    # Error statements

    def error_handling_statement(self) -> Optional[OnError]:  # TODO
        return self.__try_rules(  # type: ignore [return-value]
            self.on_error, self.resume, self.error_statement
        )

    @_with_backtracking
    def on_error(self) -> Optional[OnError]:
        start_position = self.__peek_token().position
        if not self.__pop_tokens(4)[::2] == ("On", "Error"):
            return None

        action = self.__pop_tokens(2)[0]
        if action == "Resume":
            if self.__peek_token() == "Next":
                label = None
                end_position = self.__pop_token().position
            else:
                msg = "On Error Resume statement must be followed by Next"
                position = self.__pop_token().position
                raise ParserError(msg, position)
        elif action == "Goto":
            label = self.statement_label()
            if label is None:
                msg = "On Error Goto must be followed by a literal or a name"
                position = self.__peek_token().position
                raise ParserError(msg, position)

            end_position = label.position
        else:
            return None

        return OnError(start_position + end_position, label)

    @_with_backtracking
    def resume(self) -> Optional[Resume]:
        if self.__peek_token() == "Resume":
            start_position = self.__pop_tokens(2)[0].position
        else:
            return None

        if self.__peek_token() == "Next":
            label = None
            end_position = self.__pop_token().position
        else:
            label = self.statement_label()
            if label is None:
                msg = "Resume statement must be followed by Next "
                msg += "or a statement label"
                position = self.__peek_token().position
                raise ParserError(msg, position)

            end_position = label.position

        return Resume(start_position + end_position, label)

    @_with_backtracking
    def error_statement(self) -> Optional[Error]:
        if self.__peek_token() == "Error":
            start_position = self.__pop_tokens(2)[0].position
        else:
            return None

        error_number = self.integer_expression()
        if error_number is None:
            msg = "Error statement must be followed by an integer expression"
            raise self.__craft_error(msg)

        return Error(start_position + error_number.position, error_number)

    # File statements
    def file_statement(self) -> Optional[Union[Open, Close]]:
        return self.__try_rules(  # type: ignore
            self.open,
            self.close,
            self.seek,
            self.lock,
            self.unlock,
            self.line_input,
            self.width,
            self.print,
            self.write,
            self.input,
            self.put,
            self.get,
        )

    @_with_backtracking
    def open(self) -> Optional[Open]:
        if self.__peek_token() == "Open":
            start_position = self.__pop_tokens(2)[0].position
        else:
            return None

        try:
            path_name = self.expression()
            if path_name is None:
                raise ParserError("dummy")
        except ParserError:
            msg = "Open statement requires a valid path name"
            raise self.__craft_error(msg)

        mode = self.__mode_clause()
        access = self.__access_clause(mode)
        lock = self.__lock()

        if self.__peek_token() != "As":
            raise self.__craft_error("Open statement needs a file number")
        else:
            self.__pop_tokens(2)

        file_number = self.file_number()

        if file_number is None:
            raise self.__craft_error("Open statement needs a valid file number")

        length = self.__length_clause()

        if length is not None:
            end_position = length.position
        else:
            end_position = file_number.position

        return Open(
            start_position + end_position,
            path_name,
            file_number,
            mode,
            access,
            lock,
            length,
        )

    def __mode_clause(self) -> Open.Mode:
        # "For" mode
        mode_token: Optional[Token]
        if self.__peek_token() != "For":
            return Open.Mode.RANDOM

        mode_token = self.__peek_token(2)

        try:
            mode = {
                "append": Open.Mode.APPEND,
                "binary": Open.Mode.BINARY,
                "input": Open.Mode.INPUT,
                "output": Open.Mode.OUTPUT,
                "random": Open.Mode.RANDOM,
            }[mode_token]
        except KeyError:
            position = self.__peek_token().position
            raise ParserError(
                "Open statement needs a valid mode clause", position
            )

        self.__pop_tokens(4)
        return mode

    def __access_clause(self, mode: Open.Mode) -> Open.Access:
        # "Access" access
        if self.__peek_token() != "Access":
            return {
                Open.Mode.APPEND: Open.Access.READ_WRITE,
                Open.Mode.BINARY: Open.Access.READ_WRITE,
                Open.Mode.INPUT: Open.Access.READ,
                Open.Mode.OUTPUT: Open.Access.WRITE,
                Open.Mode.RANDOM: Open.Access.READ_WRITE,
            }[mode]

        if self.__peek_tokens(2, 4) == ("Read", "Write"):
            access = Open.Access.READ_WRITE
            pop = 6
        elif self.__peek_token(2) == "Read":
            access = Open.Access.READ
            pop = 4
        elif self.__peek_token(2) == "Write":
            access = Open.Access.WRITE
            pop = 4
        else:
            position = self.__peek_token().position
            raise ParserError(
                "Open statement needs a valid access clause", position
            )

        self.__pop_tokens(pop)
        return access

    def __lock(self) -> Open.Lock:
        if self.__peek_token() == "Shared":
            lock = Open.Lock.SHARED
            pop = 2
        elif self.__peek_tokens(0, 2, 4) == ("Lock", "Read", "Write"):
            lock = Open.Lock.LOCK_READ_WRITE
            pop = 6
        elif self.__peek_tokens(0, 2) == ("Lock", "Read"):
            lock = Open.Lock.LOCK_READ
            pop = 4
        elif self.__peek_tokens(0, 2) == ("Lock", "Write"):
            lock = Open.Lock.LOCK_WRITE
            pop = 4
        else:
            return Open.Lock.SHARED

        self.__pop_tokens(pop)
        return lock

    def __length_clause(self) -> Expression:
        if self.__peek_token() != "Len":
            return Literal(
                FilePosition.build_virtual(), Variable(Integer, Integer(0))
            )

        if self.__peek_token(1) == "=":
            self.__pop_tokens(2)
        elif self.__peek_token(2) == "=":
            self.__pop_tokens(3)
        else:
            msg = "Open statement needs a valid Len clause"
            raise self.__craft_error(msg)

        return self.expression()

    def file_number(self) -> Optional[Expression]:
        if self.__peek_token() == "#":
            self.__pop_token()

        return self.expression()

    def __file_number_list(self) -> Tuple[Expression, ...]:
        numbers = []
        while True:
            number = self.file_number()
            if number is None:
                break
            else:
                numbers.append(number)

            if self.__peek_token() == ",":
                self.__pop_token()

        if numbers == []:
            for i in range(1, 512):
                numbers.append(
                    Literal(
                        FilePosition.build_virtual(),
                        Variable(Integer, Integer(i)),
                    )
                )

        return tuple(numbers)

    @_with_backtracking
    def close(self) -> Optional[Close]:
        if self.__peek_token() == "Reset":
            position = self.__pop_token().position
            file_numbers = None
        elif self.__peek_token() == "Close":
            position_start = self.__pop_token().position
            file_numbers = self.__file_number_list()

            if len(file_numbers) > 0:
                position = position_start + file_numbers[-1].position
            else:
                position = position_start
        else:
            return None

        return Close(position, file_numbers)

    @_with_backtracking
    def seek(self) -> Optional[Seek]:
        if self.__peek_token() != "Seek":
            return None
        else:
            start_position = self.__pop_tokens(2)[0].position

        file_number = self.file_number()

        if file_number is None or self.__peek_token() != ",":
            msg = "Seek statement needs a position after the file number"
            raise self.__craft_error(msg)

        self.__pop_token()
        seek_position = self.expression()
        if seek_position is None:
            raise self.__craft_error("Seek statement needs valid position")

        return Seek(
            start_position + seek_position.position, file_number, seek_position
        )

    def __lock_unlock(self, lock: bool) -> Optional[Union[Lock, Unlock]]:
        node_type: Union[Type[Lock], Type[Unlock]]
        if lock:
            keyword, node_type = "Lock", Lock
        else:
            keyword, node_type = "Unlock", Unlock

        if self.__peek_token() != keyword:
            return None
        else:
            start_position = self.__pop_tokens(2)[0].position

        file_number = self.file_number()
        if file_number is None:
            raise self.__craft_error(f"{keyword} statement needs a file number")

        if self.__peek_token() == ",":
            self.__pop_token()

            start_record_number = self.expression()
            if start_record_number is None:
                self.BLANK()
                start_record_number = Literal(
                    FilePosition.build_virtual(), Variable(Integer, Integer(1))
                )

            if self.__peek_token() == "To":
                self.__pop_token()
                end_record_number = self.expression()

                if end_record_number is None:
                    msg = keyword + " statement needs a valid end record number"
                    raise self.__craft_error(msg)
                else:
                    end_position = end_record_number.position
            else:
                end_position = start_record_number.position
                end_record_number = None
        else:
            start_record_number = None
            end_record_number = None
            end_position = file_number.position

        assert node_type in (Lock, Unlock)
        return node_type(  # type: ignore [call-arg]
            start_position + end_position,
            file_number,
            start_record_number,
            end_record_number,
        )

    @_with_backtracking
    def lock(self) -> Optional[Lock]:
        return self.__lock_unlock(lock=True)  # type: ignore [return-value]

    @_with_backtracking
    def unlock(self) -> Optional[Unlock]:
        return self.__lock_unlock(lock=False)  # type: ignore [return-value]

    @_with_backtracking
    def line_input(self) -> Optional[LineInput]:
        if self.__peek_tokens(0, 2) != ("Line", "Input"):
            return None

        start_position = self.__pop_tokens(4)[0].position
        file_number = self.file_number()
        if file_number is None:
            raise self.__craft_error("Line Input statement needs file number")
        if self.__pop_token() != ",":
            raise self.__craft_error("Line Input statement needs variable name")

        variable_name = self.variable_expression()
        if variable_name is None:
            raise self.__craft_error("Line Input statement needs variable name")

        return LineInput(
            start_position + variable_name.position,
            file_number,
            variable_name,
        )

    @_with_backtracking
    def width(self) -> Optional[Width]:
        if self.__peek_token() != "Width":
            return None

        start_position = self.__pop_tokens(2)[0].position
        file_number = self.file_number()
        if file_number is None:
            raise self.__craft_error("Line Input statement needs file number")
        if self.__pop_token() != ",":
            raise self.__craft_error("Line Input statement needs variable name")

        line_width = self.expression()
        if line_width is None:
            raise self.__craft_error("Line Input statement needs variable name")

        return Width(
            start_position + line_width.position,
            file_number,
            line_width,
        )

    @_with_backtracking
    def output_item(self) -> Optional[OutputItem]:
        clause_title = None
        clause_argument = None
        char_position = None

        if self.__category_match(_TC.BLANK):
            self.__pop_token()

        position = self.__peek_token().position
        if self.__peek_token() in (",", ";"):
            char_position = {
                ",": OutputItem.CharPosition.COMMA,
                ";": OutputItem.CharPosition.SEMICOLON,
            }[self.__pop_token()]
        else:
            clause_title, clause_argument = self.__output_clause()

            if self.__peek_token() in (",", ";"):
                position = position + self.__peek_token().position
                char_position = {
                    ",": OutputItem.CharPosition.COMMA,
                    ";": OutputItem.CharPosition.SEMICOLON,
                }[self.__pop_token()]

            if self.__category_match(_TC.BLANK):
                self.__pop_token()

        if (clause_title, clause_argument, char_position) == (None, None, None):
            return None

        if char_position is None:
            char_position = OutputItem.CharPosition.SEMICOLON

        return OutputItem(
            position, clause_title, clause_argument, char_position
        )

    def __output_clause(
        self,
    ) -> Tuple[Optional[OutputItem.Clause], Optional[Expression]]:
        clause_title: Optional[OutputItem.Clause] = None
        clause_argument = None

        if self.__category_match(_TC.BLANK):
            clause_title = OutputItem.Clause.TAB

        if self.__peek_token() == "Tab":
            clause_title = OutputItem.Clause.TAB
            clause_argument = self.__output_clause_tab()
        elif self.__peek_token() == "Spc":
            clause_title = OutputItem.Clause.SPC
            clause_argument = self.__output_clause_spc()
        else:
            clause_title = None
            clause_argument = self.expression()

        return clause_title, clause_argument

    def __output_clause_tab(self) -> Optional[Expression]:
        self.__pop_token()
        if self.__peek_token() == "(":
            self.__pop_token()
            clause_argument = self.expression()
            if clause_argument is None:
                msg = "Tab clause needs valid expression"
                raise self.__craft_error(msg)
            if not self.__pop_token() == ")":
                msg = "Tab clause needs ending parenthesis"
                raise self.__craft_error(msg)
        else:
            clause_argument = None

        return clause_argument

    def __output_clause_spc(self) -> Expression:
        self.__pop_token()
        if not self.__pop_token() == "(":
            raise self.__craft_error("Spc clause needs opening parenthesis")

        clause_argument = self.expression()
        if clause_argument is None:
            raise self.__craft_error("Spc clause needs valid expression")

        if not self.__pop_token() == ")":
            raise self.__craft_error("Spc clause needs closing parenthesis")

        return clause_argument

    def __output_list(self) -> Tuple[OutputItem, ...]:
        items = []
        while True:
            item = self.output_item()
            if item is None:
                break
            items.append(item)

        return tuple(items)

    def __write_print(self, write: bool) -> Optional[Union[Write, Print]]:
        node_type: Union[Type[Write], Type[Print]]
        if write:
            keyword = "Write"
            node_type = Write
        else:
            keyword = "Print"
            node_type = Print

        if self.__peek_token() != keyword:
            return None

        start_position = self.__pop_tokens(2)[0].position
        file_number = self.file_number()
        if file_number is None:
            raise self.__craft_error(keyword + " statement needs a file number")
        if self.__peek_token() != ",":
            msg = keyword + " statement needs an output list"
            raise self.__craft_error(msg)
        position = start_position + self.__pop_token().position

        output_list = self.__output_list()
        if len(output_list) > 1:
            position = position + output_list[-1].position

        if len(output_list) == 0 and write:
            output_list = (
                OutputItem(
                    FilePosition.build_virtual(),
                    None,
                    Literal(
                        FilePosition.build_virtual(), Variable(String, String())
                    ),
                    OutputItem.CharPosition.COMMA,
                ),
            )

        return node_type(position, file_number, output_list)

    @_with_backtracking
    def print(self) -> Optional[Write]:
        return self.__write_print(write=False)  # type: ignore [return-value]

    @_with_backtracking
    def write(self) -> Optional[Write]:
        return self.__write_print(write=True)  # type: ignore [return-value]

    @_with_backtracking
    def input(self) -> Optional[Input]:
        if not self.__peek_token() == "Input":
            return None

        start_position = self.__pop_tokens(2)[0].position
        file_number = self.file_number()
        if file_number is None:
            raise self.__craft_error("Input statement needs valid file number")

        if self.__peek_token() != ",":
            raise self.__craft_error("Input statement needs inputs")
        self.__pop_token()

        input_list = []
        while True:
            input_variable = self.bound_variable_expression()
            if input_variable is None:
                break
            input_list.append(input_variable)

            if not self.__peek_token() == ",":
                break
            self.__pop_token()

        return Input(
            start_position + input_list[-1].position,
            file_number,
            tuple(input_list),
        )

    def __put_get(self, put: bool) -> Optional[Union[Put, Get]]:
        name, node_type = {True: ("Put", Put), False: ("Get", Get)}[put]

        if not self.__peek_token() == name:
            return None

        start_position = self.__pop_tokens(2)[0].position
        file_number = self.file_number()
        if file_number is None:
            msg = f"{name} statement needs a valid file number"
            raise self.__craft_error(msg)

        if self.__pop_token() != ",":
            raise self.__craft_error("Missing comma")

        record_number = self.expression()
        if self.__pop_token() != ",":
            raise self.__craft_error("Malformed record number")

        if put:
            data_variable = self.expression()
            if data_variable is None:
                raise self.__craft_error("Put statement needs data")
        else:
            data_variable = self.variable_expression()
            if data_variable is None:
                raise self.__craft_error("Get statement needs a variable")

        return node_type(  # type: ignore [return-value]
            start_position + data_variable.position,
            file_number,
            record_number,
            data_variable,
        )

    @_with_backtracking
    def put(self) -> Optional[Put]:
        return self.__put_get(put=True)  # type: ignore [return-value]

    @_with_backtracking
    def get(self) -> Optional[Get]:
        return self.__put_get(put=False)  # type: ignore [return-value]

    # Literals

    def EOS(self) -> bool:
        ok = False
        while self.__categories_match((_TC.BLANK, _TC.END_OF_STATEMENT)):
            ok = True
            self.__pop_token()

        if not ok and self.__category_match(_TC.END_OF_FILE):
            ok = True

        return ok

    def EOF(self) -> bool:
        return self.__category_match(_TC.END_OF_FILE)

    def BLANK(self) -> bool:
        ok = False
        while self.__category_match(_TC.BLANK):
            ok = True
            self.__pop_token()

        return ok

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
