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
    ArrayDimensions,
    Call,
    CaseClause,
    Close,
    ConstItem,
    ControlStatement,
    DataManipulationStatement,
    DictAccess,
    Do,
    Erase,
    Error,
    ErrorHandlingStatement,
    ExitDo,
    ExitFor,
    ExitFunction,
    ExitProperty,
    ExitSub,
    Expression,
    FileStatement,
    For,
    ForEach,
    Get,
    GoSub,
    GoTo,
    If,
    IfClause,
    IndexExpr,
    Input,
    Let,
    LExpression,
    LineInput,
    Literal,
    LocalConstantDeclaration,
    LocalVariableDeclaration,
    Lock,
    LSet,
    MemberAccess,
    Mid,
    Name,
    OnError,
    OnGoSub,
    OnGoTo,
    Open,
    Operation,
    Operator,
    OutputItem,
    ParenExpr,
    Print,
    Put,
    RaiseEvent,
    RangeClause,
    ReDim,
    Resume,
    Return,
    RSet,
    Seek,
    SelectCase,
    Set,
    Statement,
    StatementBlock,
    StatementLabel,
    StmtLabel,
    Stop,
    Unlock,
    VariableDeclaration,
    While,
    Width,
    With,
    Write,
)
from .data import (
    Boolean,
    Byte,
    Currency,
    Date,
    DeclaredType,
    Double,
    Empty,
    Integer,
    Long,
    LongLong,
    Null,
    ObjectReference,
    Single,
    String,
    StringN,
    Value,
    Variable,
    Variant,
)
from .error import ParserError
from .lexer import Lexer, Token
from .utils import VIRTUAL_POSITION, FilePosition

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
    def backtracking_wrapper(self: "Parser") -> _T:
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

    return backtracking_wrapper


def _add_rule_position(rule: _Rule) -> _Rule:
    """
    Decorator adding position information to a rule: if it returns not ``None``,
    it replaces the node position with the position range starting at the first
    parsed token and stopping before the first non-parsed one.
    """

    @wraps(rule)
    def position_wrapper(self: "Parser") -> _T:
        lexer = self._Parser__lexer  # type: ignore [attr-defined]
        start_position = lexer.peek_token().position

        try:
            node = rule(self)
        except ParserError:
            raise

        if node is None:
            return node  # type: ignore [return-value]

        end_position = lexer.peek_token().position
        new_position = start_position - end_position

        # This is the only place you are allowed to do that !
        object.__setattr__(node, "position", new_position)

        return node

    return position_wrapper


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

    __TYPED_NAME_SUFFIXES = {
        "%": Integer,
        "&": Long,
        "^": LongLong,
        "!": Single,
        "#": Double,
        "@": Currency,
        "$": String,
    }  #: Mapping between typed names suffix and their declared type

    __BUILTIN_TYPES = {
        "boolean": Boolean,
        "byte": Byte,
        "currency": Currency,
        "date": Date,
        "double": Double,
        "integer": Integer,
        "long": Long,
        "longlong": LongLong,
        "single": Single,
        "string": String,
        "variant": Variant,
    }  #: Mapping between built-in type name and corresponding declared type

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

    __LEXPRESSION_OPERATORS = {
        _OT.INDEX,
        _OT.EXCLAMATION,
        _OT.DOT,
        _OT.RPAREN,
        _OT.LPAREN,
        _OT.COMMA,
    }  #: Operators allowed in an l-expression

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
    def expression(self) -> Optional[Expression]:
        return self.__expression(l_expression=False, arglist=False)

    @_with_backtracking
    def l_expression(self) -> Optional[LExpression]:
        return self.__expression(  # type: ignore [return-value]
            l_expression=True, arglist=False
        )

    def __expression(
        self, l_expression: bool, arglist: bool
    ) -> Optional[Expression]:
        """
        Expression parser using the Shunting-Yard algorithm. Can be used to
        parse any expression, only an l_expression (if l_expression is True),
        or only an arguments list (if arglist is True).
        """
        # Operator stack, is never empty until parsing is done, which is
        # obtained by surrounding the expression in parentheses
        operators: List[Operator] = [Operator(self.__peek_token(), _OT.LPAREN)]
        # Operands, begins empty but stays non-empty during all parsing
        operands: List[Expression] = []
        self.__index_expressions = 0
        self.__paren_expressions = 0

        if arglist:
            operators.append(Operator(self.__peek_token(), _OT.INDEX))
            operands.append(Expression(VIRTUAL_POSITION))
            self.__index_expressions = 1

        self.__last_operatorand = operators[-1]

        try:
            self.__expression_shunting_yard(
                operators, operands, l_expression, arglist
            )
        except Exception as e:
            msg = "Encountered an unexpected exception during expression "
            msg += f"parsing: {e}"
            raise self.__craft_error(msg)

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
        self,
        operators: List[Operator],
        operands: List[Expression],
        l_expression: bool,
        arglist: bool,
    ) -> None:
        while True:
            if self.__categories_match(
                (_TC.END_OF_STATEMENT, _TC.END_OF_FILE, _TC.SYMBOL)
            ):
                break
            elif self.BLANK():
                continue

            # Encounter operator
            if not l_expression:
                operator = self.operator()
            else:
                operator = self.__expression_loperator()

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

        # Close the virtual parenthesis
        closing_parentheses = 1 if not arglist else 2
        for _ in range(closing_parentheses):
            self.__expression_closing_parenthesis(
                operators, operands, self.__peek_token().position
            )

    @_with_backtracking
    def __expression_loperator(self) -> Optional[Operator]:
        operator = self.operator()
        if (
            operator is not None
            and operator.op not in self.__LEXPRESSION_OPERATORS
        ):

            return None

        return operator

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
        """Handle an operator during expression parsing: close and open
        parenthesis, simply push on operator stack or reduce the stacks
        depending on its precedence."""
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
        # TODO improve doc here
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

    def constant_expression(self) -> Optional[Expression]:
        return self.expression()

    def integer_expression(self) -> Optional[Expression]:
        return self.expression()

    def boolean_expression(self) -> Optional[Expression]:
        return self.expression()

    def variable_expression(self) -> Optional[Expression]:
        return self.expression()

    def bound_variable_expression(self) -> Optional[Expression]:
        return self.l_expression()

    def defined_type_expression(self) -> Optional[Expression]:
        return self.l_expression()

    # Statements (individual statements do not include the END_OF_STATEMENT
    # token)

    @_add_rule_position
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
        self.EOS()  # Skip trailing newlines

        return StatementBlock(VIRTUAL_POSITION, tuple(statements))

    def block_statement(self) -> Optional[Statement]:  # TODO
        return self.__try_rules(self.statement)  # type: ignore

    def statement(self) -> Optional[Statement]:
        statement = self.__try_rules(
            self.control_statement,
            self.error_handling_statement,
            self.file_statement,
            self.data_manipulation_statement,
        )
        if statement is None or not self.EOS():
            return None

        return statement  # type: ignore [return-value]

    def statement_label(self) -> Optional[StmtLabel]:
        return self.__try_rules(self.NAME, self.literal)  # type: ignore

    # Control statements

    def control_statement(self) -> Optional[ControlStatement]:
        return self.__try_rules(  # type: ignore [return-value]
            self.if_, self.control_statement_except_multiline_if
        )

    def control_statement_except_multiline_if(
        self,
    ) -> Optional[ControlStatement]:
        return self.__try_rules(  # type: ignore [return-value]
            self.call,
            self.while_,
            self.for_,
            self.for_each,
            self.exit_for,
            self.do,
            self.exit_do,
            self.single_line_if,
            self.select_case,
            self.stop,
            self.go_to,
            self.on_go_to_sub,
            self.go_sub,
            self.return_,
            self.exit_sub,
            self.exit_function,
            self.exit_property,
            self.raise_event,
            self.with_,
        )

    @_add_rule_position
    @_with_backtracking
    def call(self) -> Optional[Call]:
        if self.__peek_token() == "Call":
            self.__pop_token()

            expr = self.expression()
            if expr is None:
                raise self.__craft_error("Call statement needs a valid callee")

            if not isinstance(expr, IndexExpr):
                expr = IndexExpr(
                    VIRTUAL_POSITION, expr, ArgList(VIRTUAL_POSITION, tuple())
                )
        else:
            if self.__category_match(_TC.KEYWORD):
                return None

            callee = self.expression()
            if callee is None:
                return None

            arg_list = self.__expression(l_expression=False, arglist=True)
            if arg_list is None or not isinstance(arg_list, IndexExpr):
                return None

            arguments = arg_list.args
            expr = IndexExpr(VIRTUAL_POSITION, callee, arguments)

        return Call(VIRTUAL_POSITION, expr)

    @_add_rule_position
    @_with_backtracking
    def while_(self) -> Optional[While]:
        # Header
        if self.__pop_token() != "While":
            return None

        condition = self.boolean_expression()
        if condition is None:
            raise self.__craft_error("While loop needs valid condition")

        # Body
        self.BLANK()
        if self.__pop_token().category != _TC.END_OF_STATEMENT:
            msg = "While loop header must be followed by a newline"
            raise self.__craft_error(msg)

        body = self.statement_block()
        if body is None:
            raise self.__craft_error("While loop needs a valid body")

        # Footer
        if self.__pop_token() != "Wend":
            raise self.__craft_error("While loop must end with Wend footer")

        return While(VIRTUAL_POSITION, body, condition)

    @_add_rule_position
    @_with_backtracking
    def for_(self) -> Optional[For]:
        # Header
        if self.__pop_token() != "For":
            return None

        self.BLANK()
        header = self.__for_header()
        if header is None:
            return None
        iterator, start_value, end_value, step_value = header

        # Body
        self.BLANK()
        if self.__pop_token().category != _TC.END_OF_STATEMENT:
            msg = "For loop header must be followed by a newline"
            raise self.__craft_error(msg)

        body = self.statement_block()
        if body is None:
            raise self.__craft_error("For loop needs a valid body")

        # Footer
        if self.__pop_token() != "Next":
            raise self.__craft_error("For loop needs a Next footer")

        self.bound_variable_expression()

        return For(
            VIRTUAL_POSITION, body, iterator, start_value, end_value, step_value
        )

    def __for_header(
        self,
    ) -> Optional[
        Tuple[Expression, Expression, Expression, Optional[Expression]]
    ]:
        if self.__peek_token() == "Each":
            return None

        iterator = self.bound_variable_expression()
        if iterator is None:
            raise self.__craft_error("For loop needs a valid iterator")

        if self.__pop_token() != "=":
            raise self.__craft_error("Iterator needs an initial value")

        start_value = self.expression()
        if start_value is None:
            raise self.__craft_error("Iterator needs a valid initial value")

        if self.__pop_token() != "To":
            raise self.__craft_error("Iterator needs a final value")

        end_value = self.expression()
        if end_value is None:
            raise self.__craft_error("Iterator needs a valid final value")

        if self.__peek_token() == "Step":
            self.__pop_token()
            step_value = self.expression()
            if step_value is None:
                raise self.__craft_error("Iterator needs a valid step value")
        else:
            step_value = None

        return iterator, start_value, end_value, step_value

    @_add_rule_position
    @_with_backtracking
    def for_each(self) -> Optional[ForEach]:
        # Header
        if self.__peek_tokens(0, 2) != ("For", "Each"):
            return None
        self.__pop_tokens(3)

        iterator = self.bound_variable_expression()
        if iterator is None:
            raise self.__craft_error("For Each loop needs a valid iterator")

        if self.__pop_token() != "In":
            raise self.__craft_error("For Each loop needs a collection")

        collection = self.expression()
        if collection is None:
            raise self.__craft_error("For Each loop needs a valid collection")

        # Body
        self.BLANK()
        if self.__pop_token().category != _TC.END_OF_STATEMENT:
            msg = "For Each loop header must be followed by a newline"
            raise self.__craft_error(msg)

        body = self.statement_block()
        if body is None:
            raise self.__craft_error("For Each loop needs a valid body")

        # Footer
        if self.__pop_token() != "Next":
            raise self.__craft_error("For Each loop needs a Next footer")

        self.bound_variable_expression()

        return ForEach(VIRTUAL_POSITION, body, iterator, collection)

    @_add_rule_position
    def exit_for(self) -> Optional[ExitFor]:
        if self.__peek_tokens(0, 2) != ("Exit", "For"):
            return None

        self.__pop_tokens(3)
        return ExitFor(VIRTUAL_POSITION)

    @_add_rule_position
    @_with_backtracking
    def do(self) -> Optional[Do]:
        # Header
        if self.__pop_token() != "Do":
            return None

        first_condition = self.__do_condition()

        # Body
        self.BLANK()
        if self.__pop_token().category != _TC.END_OF_STATEMENT:
            raise self.__craft_error("Do loop header must end with a newline")
        body = self.statement_block()

        # Footer
        if self.__pop_token() != "Loop":
            raise self.__craft_error("Do loop needs a Loop footer")

        second_condition = self.__do_condition()

        # Extract condition
        condition_type: Optional[Do.ConditionType]
        condition: Optional[Expression]
        if (first_condition is not None) and (second_condition is not None):
            msg = "Do loop needs at most one condition, located either at its "
            msg += "header or footer"
            raise self.__craft_error(msg)
        elif first_condition is not None:
            condition_type, condition = first_condition
        elif second_condition is not None:
            condition_type, condition = second_condition
        else:
            condition_type, condition = None, None

        return Do(VIRTUAL_POSITION, body, condition_type, condition)

    def __do_condition(self) -> Optional[Tuple[Do.ConditionType, Expression]]:
        self.BLANK()
        keyword = self.__peek_token()
        if keyword == "While":
            condition_type = Do.ConditionType.WHILE
        elif keyword == "Until":
            condition_type = Do.ConditionType.UNTIL
        else:
            return None
        self.__pop_token()

        condition = self.boolean_expression()
        if condition is None:
            msg = "Do condition needs a valid boolean expression"
            raise self.__craft_error(msg)

        return condition_type, condition

    @_add_rule_position
    def exit_do(self) -> Optional[ExitDo]:
        if self.__peek_tokens(0, 2) != ("Exit", "Do"):
            return None

        self.__pop_tokens(3)
        return ExitDo(VIRTUAL_POSITION)

    @_add_rule_position
    @_with_backtracking
    def if_(self) -> Optional[If]:
        clauses = []
        # First If block
        if_clause = self.__if_clause()
        if if_clause is None:
            return None
        clauses.append(if_clause)

        # Optional following ElseIf blocks
        while True:
            elif_clause = self.__else_if_clause()
            if elif_clause is None:
                break
            else:
                clauses.append(elif_clause)

        # Optional Else blocks
        if self.__peek_token() == "Else":
            self.__pop_token()
            else_block = self.statement_block()
            if else_block is None:
                raise self.__craft_error("Else block needs a valid body")
        else:
            else_block = None

        # Footer
        if self.__peek_token() == "EndIf":
            self.__pop_token()
        elif self.__peek_tokens(0, 2) == ("End", "If"):
            self.__pop_tokens(3)
        else:
            raise self.__craft_error("If block needs an EndIf or End If footer")

        return If(VIRTUAL_POSITION, tuple(clauses), else_block)

    @_add_rule_position
    @_with_backtracking
    def __if_clause(self) -> Optional[IfClause]:
        return self.__if_elseif_clause(else_if=False)

    @_add_rule_position
    @_with_backtracking
    def __else_if_clause(self) -> Optional[IfClause]:
        return self.__if_elseif_clause(else_if=True)

    def __if_elseif_clause(self, else_if: bool) -> Optional[IfClause]:
        # Header
        keyword = {True: "ElseIf", False: "If"}
        if self.__peek_token() != keyword[else_if]:
            return None

        self.__pop_token()
        condition = self.boolean_expression()
        if condition is None:
            raise self.__craft_error(f"{keyword} block needs a valid condition")

        if self.__pop_token() != "Then":
            raise self.__craft_error("If header needs a Then keyword")

        # Body
        self.BLANK()
        if self.__pop_token().category != _TC.END_OF_STATEMENT:
            if else_if:
                msg = f"{keyword} block header must be followed by a newline"
                raise self.__craft_error(msg)
            else:
                return None  # Exception might be raised by single_line_if

        body = self.statement_block()
        if body is None:
            raise self.__craft_error(f"{keyword} block needs a valid body")

        return IfClause(VIRTUAL_POSITION, body, condition)

    @_add_rule_position
    @_with_backtracking
    def single_line_if(self) -> Optional[If]:
        # First list of statements (if block)
        if_clause = self.__single_line_if_clause()
        if if_clause is None:
            return None

        # Second list of statements (else block)
        if self.__peek_token() != "Else":
            if len(if_clause.statements) == 0:
                msg = "If statement needs to trigger at least one statement"
                raise self.__craft_error(msg)
            else_block = None
        else:
            self.__pop_token()
            else_block = self.__single_line_if_list_or_label()
            if else_block is None:
                msg = "If statement needs a valid Else block"
                raise self.__craft_error(msg)

        return If(VIRTUAL_POSITION, tuple([if_clause]), else_block)

    @_add_rule_position
    @_with_backtracking
    def __single_line_if_clause(self) -> Optional[IfClause]:
        # Condition
        if self.__pop_token() != "If":
            return None

        condition = self.boolean_expression()
        if condition is None:
            raise self.__craft_error("If statement needs a valid condition")

        if self.__pop_token() != "Then":
            raise self.__craft_error("If condition must be followed by Then")

        # List of statements
        if_statements = self.__single_line_if_list_or_label()
        if if_statements is None:
            raise self.__craft_error("If statement needs a valid If block")
        if_clause = IfClause(
            VIRTUAL_POSITION, if_statements.statements, condition
        )

        return if_clause

    @_add_rule_position
    @_with_backtracking
    def __single_line_if_list_or_label(self) -> Optional[StatementBlock]:
        self.BLANK()
        statements = []

        # First element: try any statement first, then label
        first_statement = self.same_line_statement()
        if first_statement is None:
            label = self.statement_label()
            if label is None:
                return None

            statement_label = StatementLabel(
                VIRTUAL_POSITION, label  # type: ignore
            )
            first_statement = GoTo(label.position, statement_label)

        statements.append(first_statement)

        # Subsequent colon-separated elements
        while self.__single_line_if_colon():
            statement = self.same_line_statement()
            if statement is None:
                break
            statements.append(statement)

        self.__single_line_if_colon()

        return StatementBlock(VIRTUAL_POSITION, tuple(statements))

    def __single_line_if_colon(self) -> bool:
        """
        Consumes colons, without backtracking needed.

        Returns:
          ``True`` if a BLANK-separated list of at least one colon is
          encountered, else ``False``
        """
        seen_one = False
        while True:
            self.BLANK()
            if self.__peek_token() == ":":
                self.__pop_token()
                seen_one = True
            else:
                return seen_one

            self.BLANK()

    def same_line_statement(self) -> Optional[Statement]:
        return self.__try_rules(  # type: ignore [return-value]
            self.file_statement,
            self.error_handling_statement,
            self.data_manipulation_statement,
            self.control_statement_except_multiline_if,
        )

    @_add_rule_position
    @_with_backtracking
    def select_case(self) -> Optional[SelectCase]:
        # Header
        if self.__peek_token() == "Select":
            if self.__peek_token(2) != "Case":
                raise self.__craft_error("Select must be followed by Case")
        else:
            return None
        self.__pop_tokens(3)

        select_expression = self.expression()
        if select_expression is None:
            msg = "Select Case statement needs a valid expression"
            raise self.__craft_error(msg)

        if not self.EOS():
            msg = "Select Case statement header must end with a newline"
            raise self.__craft_error(msg)

        # Clauses
        clauses = []
        while True:
            self.EOS()
            clause = self.case_clause()
            if clause is None:
                break
            clauses.append(clause)

        self.EOS()
        case_else = self.case_else_clause()

        self.EOS()
        if self.__peek_tokens(0, 2) != ("End", "Select"):
            msg = "Select Case statement must end with End Select"
            raise self.__craft_error(msg)
        self.__pop_tokens(3)

        return SelectCase(
            VIRTUAL_POSITION, select_expression, tuple(clauses), case_else
        )

    @_add_rule_position
    @_with_backtracking
    def case_clause(self) -> Optional[CaseClause]:
        if self.__pop_token() != "Case":
            return None
        self.BLANK()

        first_range = self.range_clause()
        if first_range is None:
            return None  # It can be a Case Else clause

        subsequent_ranges = []
        while True:
            if self.__peek_token() == ",":
                self.__pop_token()
                self.BLANK()
                clause = self.range_clause()
            else:
                break
            subsequent_ranges.append(clause)

        if not self.EOS():
            raise self.__craft_error("Case clause must end with a newline")

        body = self.statement_block()

        return CaseClause(
            VIRTUAL_POSITION, body, first_range, tuple(subsequent_ranges)
        )

    def case_else_clause(self) -> Optional[StatementBlock]:
        if self.__peek_token() == "Case":
            if self.__peek_token(2) != "Else":
                msg = "Case keyword must be followed by Else or a "
                msg += "valid expression"
                raise self.__craft_error(msg)
            self.__pop_tokens(3)

            if not self.EOS():
                msg = "Case Else clause must end with a newline"
                raise self.__craft_error(msg)

            case_else = self.statement_block()
            if case_else is None:
                msg = "Case Else clause needs a valid statements block"
                raise self.__craft_error(msg)
        else:
            case_else = None

        return case_else

    @_add_rule_position
    @_with_backtracking
    def range_clause(self) -> Optional[RangeClause]:
        if self.__peek_token() == "Is":
            self.__pop_token()
            self.BLANK()
            comparison_operator = self.__pop_token()
            valid_operators = (
                "=",
                "<>",
                "><",
                "<",
                ">",
                ">=",
                "<=",
                "=<",
                "=>",
            )

            if comparison_operator not in valid_operators:
                msg = "Range clause comparison operator needs to be one of "
                msg += ", ".join(valid_operators)
                raise self.__craft_error(msg)

            operator_type = self.__OP_TYPE[comparison_operator]
            operator = Operator(comparison_operator, operator_type)

            expression = self.expression()

            return RangeClause(VIRTUAL_POSITION, expression, None, operator)

        first_expression = self.expression()
        if first_expression is None:
            return None

        if self.__peek_token() != "To":
            return RangeClause(VIRTUAL_POSITION, first_expression, None, None)

        self.__pop_token()
        second_expression = self.expression()

        if second_expression is None:
            msg = "Range clause needs To to be followed by a valid expression"
            raise self.__craft_error(msg)

        return RangeClause(
            VIRTUAL_POSITION, first_expression, second_expression, None
        )

    @_add_rule_position
    def stop(self) -> Optional[Stop]:
        if self.__peek_token() != "Stop":
            return None

        self.__pop_token()
        return Stop(VIRTUAL_POSITION)

    @_add_rule_position
    @_with_backtracking
    def go_to(self) -> Optional[GoTo]:
        if self.__peek_token() == "GoTo":
            self.__pop_token()
        elif self.__peek_tokens(0, 2) == ("Go", "To"):
            self.__pop_tokens(3)
        else:
            return None
        self.BLANK()

        label = self.statement_label()
        if label is None:
            raise self.__craft_error("GoTo statement needs a valid label")

        return GoTo(VIRTUAL_POSITION, label)  # type: ignore

    @_add_rule_position
    @_with_backtracking
    def on_go_to_sub(self) -> Optional[Union[OnGoTo, OnGoSub]]:
        if self.__pop_token() != "On":
            return None

        expression = self.expression()
        if expression is None:
            return None

        keyword = self.__pop_token()
        ast_type: Union[Type[OnGoTo], Type[OnGoSub]]
        if keyword == "GoTo":
            ast_type = OnGoTo
        elif keyword == "GoSub":
            ast_type = OnGoSub
        else:
            msg = "On keyword must be used with GoTo or GoSub"
            raise self.__craft_error(msg)

        labels: List[StatementLabel] = []
        while True:
            self.BLANK()
            label = self.statement_label()
            if label is None:
                break
            labels.append(label)  # type: ignore

            self.BLANK()
            if self.__peek_token() != ",":
                break

            self.__pop_token()

        return ast_type(VIRTUAL_POSITION, expression, tuple(labels))

    @_add_rule_position
    @_with_backtracking
    def go_sub(self) -> Optional[GoTo]:
        if self.__peek_token() == "GoSub":
            self.__pop_token()
        elif self.__peek_tokens(0, 2) == ("Go", "Sub"):
            self.__pop_tokens(3)
        else:
            return None
        self.BLANK()

        label = self.statement_label()
        if label is None:
            raise self.__craft_error("GoSub statement needs a valid label")

        return GoSub(VIRTUAL_POSITION, label)  # type: ignore

    @_add_rule_position
    def return_(self) -> Optional[Return]:
        if self.__peek_token() != "Return":
            return None

        self.__pop_token()
        return Return(VIRTUAL_POSITION)

    @_add_rule_position
    @_with_backtracking
    def exit_sub(self) -> Optional[ExitSub]:
        if self.__peek_tokens(0, 2) != ("Exit", "Sub"):
            return None
        self.__pop_tokens(3)

        return ExitSub(VIRTUAL_POSITION)

    @_add_rule_position
    @_with_backtracking
    def exit_function(self) -> Optional[ExitFunction]:
        if self.__peek_tokens(0, 2) != ("Exit", "Function"):
            return None
        self.__pop_tokens(3)

        return ExitFunction(VIRTUAL_POSITION)

    @_add_rule_position
    @_with_backtracking
    def exit_property(self) -> Optional[ExitProperty]:
        if self.__peek_tokens(0, 2) != ("Exit", "Property"):
            return None
        self.__pop_tokens(3)

        return ExitProperty(VIRTUAL_POSITION)

    def raise_event(self) -> Optional[RaiseEvent]:
        return None

    def with_(self) -> Optional[With]:
        return None

    # Data manipulation statements

    def data_manipulation_statement(
        self,
    ) -> Optional[DataManipulationStatement]:
        return self.__try_rules(  # type: ignore [return-value]
            self.local_variable_declaration,
            self.local_constant_declaration,
            self.re_dim,
            self.erase,
            self.mid,
            self.lset,
            self.rset,
            self.let,
            self.set,
        )

    @_add_rule_position
    @_with_backtracking
    def array_dimensions(self) -> Optional[ArrayDimensions]:
        if not self.__peek_token() == "(":
            return None
        self.__pop_token()

        dimensions = []
        while True:
            if self.__peek_token() == ")":
                self.__pop_token()
                break

            bound1 = self.constant_expression()
            if bound1 is None:
                msg = "Array dimension needs at least one bound"
                raise self.__craft_error(msg)

            if self.__peek_token() == "To":
                self.__pop_token()
                bound2 = self.constant_expression()
                if bound2 is None:
                    msg = "Array dimension needs a constant expression as "
                    msg += "upper bound"
                    raise self.__craft_error(msg)
            else:
                bound2 = None

            dimensions.append((bound1, bound2))

            if self.__peek_token() == ",":
                self.__pop_token()

        return ArrayDimensions(VIRTUAL_POSITION, tuple(dimensions))

    @_add_rule_position
    @_with_backtracking
    def variable_declaration(self) -> Optional[VariableDeclaration]:
        # TODO handle declared type
        self.BLANK()
        name_node = self.NAME()
        if name_node is None:
            return None
        name = name_node.name

        new = False
        declared_type: Optional[Union[Expression, DeclaredType]] = None
        if self.__peek_token() in self.__TYPED_NAME_SUFFIXES:
            suffix = self.__pop_token()
            declared_type = self.__TYPED_NAME_SUFFIXES[suffix]
            array_dimensions = self.array_dimensions()
        else:
            array_dimensions = self.array_dimensions()
            self.BLANK()
            as_clause = self.__as_clause_type()
            if as_clause is not None:
                new, declared_type = as_clause

        return VariableDeclaration(
            VIRTUAL_POSITION, name, array_dimensions, new, declared_type
        )

    def __as_clause_type(
        self,
    ) -> Optional[Tuple[bool, Optional[Union[Expression, DeclaredType]]]]:
        if self.__peek_token() != "As":
            return None

        self.__pop_token()
        self.BLANK()

        new = False
        declared_type: Optional[Union[Expression, DeclaredType]]
        if self.__peek_token() == "New":  # Auto object: As New ...
            self.__pop_token()
            new = True
            declared_type = self.defined_type_expression()
            if declared_type is None:
                msg = "Auto object needs a valid class name"
        elif self.__peek_tokens(0, 1) == ("String", "*"):
            self.__pop_tokens(2)
            length = self.constant_expression()
            if length is None or not isinstance(length, Name):
                msg = "Fixed length string needs a valid constant length"
                raise self.__craft_error(msg)

            declared_type = StringN(length.name)
        elif self.__category_match(_TC.KEYWORD):
            type_name = self.__pop_token()
            try:
                declared_type = self.__BUILTIN_TYPES[type_name]
            except KeyError:
                msg = f"{str(type_name)} is not a supported type"
                raise self.__craft_error(msg)
        else:
            declared_type = self.l_expression()

        return (new, declared_type)

    def __variable_declarations(self) -> Tuple[VariableDeclaration, ...]:
        declarations = []
        while True:
            declaration = self.variable_declaration()
            if declaration is None:
                msg = "Malformed declaration"
                raise self.__craft_error(msg)
            declarations.append(declaration)

            self.BLANK()
            if self.__peek_token() == ",":
                self.__pop_token()
            else:
                break

        return tuple(declarations)

    @_add_rule_position
    @_with_backtracking
    def local_variable_declaration(self) -> Optional[LocalVariableDeclaration]:
        if self.__peek_token() == "Dim":
            self.__pop_token()
            self.BLANK()
            if self.__peek_token() == "Shared":
                self.__pop_token()
                self.BLANK()
            static = False
        elif self.__peek_token() == "Static":
            self.__pop_token()
            self.BLANK()
            static = True
        else:
            return None

        variable_declarations = self.__variable_declarations()

        return LocalVariableDeclaration(
            VIRTUAL_POSITION, static, variable_declarations
        )

    def __const_item_list(self) -> Tuple[ConstItem, ...]:
        items = []
        while True:
            self.BLANK()
            item = self.const_item()
            if item is None:
                raise self.__craft_error("Const item lists expects an item")

            items.append(item)

            self.BLANK()
            if self.__peek_token() != ",":
                break
            self.__pop_token()

        return tuple(items)

    @_add_rule_position
    def const_item(self) -> Optional[ConstItem]:
        name_node = self.NAME()
        if name_node is None:
            return None
        name = name_node.name

        declared_type: Optional[DeclaredType]
        if self.__peek_token() in self.__TYPED_NAME_SUFFIXES:
            suffix = self.__pop_token()
            declared_type = self.__TYPED_NAME_SUFFIXES[suffix]
        else:
            self.BLANK()
            if self.__peek_token() == "As":
                self.__pop_token()
                self.BLANK()
                builtin_type = self.__pop_token()
                if builtin_type.category != _TC.KEYWORD:
                    msg = "Constant must have a builtin type value"
                    raise self.__craft_error(msg)

                declared_type = self.__BUILTIN_TYPES[builtin_type]
            else:
                declared_type = None

        self.BLANK()
        if self.__pop_token() != "=":
            raise self.__craft_error("Constant item must have a declared value")

        value = self.constant_expression()
        if value is None:
            raise self.__craft_error("Constant item must have a declared value")

        return ConstItem(VIRTUAL_POSITION, name, declared_type, value)

    @_add_rule_position
    @_with_backtracking
    def local_constant_declaration(self) -> Optional[LocalConstantDeclaration]:
        if not self.__peek_token() == "Const":
            return None
        self.__pop_token()
        const_items = self.__const_item_list()

        return LocalConstantDeclaration(VIRTUAL_POSITION, const_items)

    @_add_rule_position
    @_with_backtracking
    def re_dim(self) -> Optional[ReDim]:
        if not self.__peek_token() == "ReDim":
            return None

        self.__pop_token()
        self.BLANK()
        if self.__peek_token() == "Preserve":
            preserve = True
            self.__pop_token()
        else:
            preserve = False

        variable_declarations = self.__variable_declarations()

        return ReDim(VIRTUAL_POSITION, preserve, variable_declarations)

    def __erase_list(self) -> Tuple[Expression, ...]:
        erase_elements = []
        while True:
            erase_element = self.l_expression()
            if erase_element is None:
                raise self.__craft_error("Erase statement needs a valid list")

            erase_elements.append(erase_element)
            if self.__peek_token() == ",":
                self.__pop_token()
            else:
                break

        return tuple(erase_elements)

    @_add_rule_position
    @_with_backtracking
    def erase(self) -> Optional[Erase]:
        if self.__peek_token() != "Erase":
            return None

        self.__pop_token()
        elements = self.__erase_list()
        return Erase(VIRTUAL_POSITION, elements)

    @_add_rule_position
    @_with_backtracking
    def mid(self) -> Optional[Mid]:
        mode_specifier = self.__peek_token()

        if mode_specifier in ("Mid", "Mid$"):
            byte_level = False
        elif mode_specifier in ("MidB", "MidB$"):
            byte_level = True
        else:
            return None
        self.__pop_token()

        self.BLANK()
        string_argument, start, length = self.__mid_arguments()
        self.BLANK()

        if self.__pop_token() != "=":
            raise self.__craft_error("Mid statement needs a assigned value")

        value = self.expression()
        if value is None:
            raise self.__craft_error("Mid statement needs a assigned value")

        return Mid(
            VIRTUAL_POSITION, byte_level, string_argument, start, length, value
        )

    def __mid_arguments(
        self,
    ) -> Tuple[Expression, Expression, Optional[Expression]]:
        if self.__pop_token() != "(":
            raise self.__craft_error("Mid statement needs an argument")

        string_argument = self.bound_variable_expression()
        if string_argument is None:
            msg = "Mid statement needs a valid bound variable string argument"
            raise self.__craft_error(msg)

        if self.__peek_token() == ",":
            self.__pop_token()
            start = self.integer_expression()
            if start is None:
                msg = "Mid statement start argument must be a valid integer "
                msg += "expression"
                raise self.__craft_error(msg)

            if self.__peek_token() == ",":
                self.__pop_token()
                length = self.integer_expression()
                if length is None:
                    msg = "Mid statement length argument must be a valid "
                    msg += "integer expression"
                    raise self.__craft_error(msg)
            else:
                length = None
        else:
            raise self.__craft_error("Mid statement needs a start argument")

        if self.__pop_token() != ")":
            msg = "Mid statement misses a closing parenthesis"
            raise self.__craft_error(msg)

        return string_argument, start, length

    def __lr_set(self, lset: bool) -> Optional[Union[LSet, RSet]]:
        node_type: Union[Type[LSet], Type[RSet]]
        if lset:
            kw = "LSet"
            node_type = LSet
        else:
            kw = "RSet"
            node_type = RSet

        if self.__peek_token() != kw:
            return None

        self.__pop_token()
        variable = self.bound_variable_expression()
        if variable is None:
            raise self.__craft_error(f"{kw} statement needs a bound variable")

        if self.__pop_token() != "=":
            raise self.__craft_error(f"{kw} statement needs a valid value")

        value = self.expression()
        if value is None:
            raise self.__craft_error(f"{kw} statement needs a valid value")

        return node_type(VIRTUAL_POSITION, variable, value)

    @_add_rule_position
    @_with_backtracking
    def lset(self) -> Optional[RSet]:
        return self.__lr_set(True)  # type: ignore [return-value]

    @_add_rule_position
    @_with_backtracking
    def rset(self) -> Optional[RSet]:
        return self.__lr_set(False)  # type: ignore [return-value]

    @_add_rule_position
    @_with_backtracking
    def let(self) -> Optional[Let]:
        if self.__peek_token() == "Let":
            self.__pop_token()

        l_expression = self.l_expression()
        if l_expression is None:
            return None

        if self.__pop_token() != "=":
            return None

        value = self.expression()
        if value is None:
            raise self.__craft_error("Let statement needs valid value")

        return Let(VIRTUAL_POSITION, l_expression, value)

    @_add_rule_position
    @_with_backtracking
    def set(self) -> Optional[Set]:
        if not self.__peek_token() == "Set":
            return None

        self.__pop_tokens(2)
        l_expression = self.l_expression()
        if l_expression is None:
            raise self.__craft_error("Set statement needs a valid l-expression")

        if self.__pop_token() != "=":
            raise self.__craft_error("Set statement needs assignment")

        value = self.expression()
        if value is None:
            raise self.__craft_error("Set statement needs a valid value")

        return Set(VIRTUAL_POSITION, l_expression, value)

    # Error statements

    def error_handling_statement(self) -> Optional[ErrorHandlingStatement]:
        return self.__try_rules(  # type: ignore [return-value]
            self.on_error, self.resume, self.error_statement
        )

    @_add_rule_position
    @_with_backtracking
    def on_error(self) -> Optional[OnError]:
        if not self.__pop_tokens(4)[::2] == ("On", "Error"):
            return None

        action = self.__pop_tokens(2)[0]
        if action == "Resume":
            if self.__peek_token() == "Next":
                label = None
                self.__pop_token()
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
        else:
            return None

        return OnError(VIRTUAL_POSITION, label)

    @_add_rule_position
    @_with_backtracking
    def resume(self) -> Optional[Resume]:
        if self.__peek_token() != "Resume":
            return None

        self.__pop_tokens(2)

        if self.__peek_token() == "Next":
            label = None
            self.__pop_token()
        else:
            label = self.statement_label()
            if label is None:
                msg = "Resume statement must be followed by Next "
                msg += "or a statement label"
                position = self.__peek_token().position
                raise ParserError(msg, position)

        return Resume(VIRTUAL_POSITION, label)

    @_add_rule_position
    @_with_backtracking
    def error_statement(self) -> Optional[Error]:
        if self.__peek_token() != "Error":
            return None

        self.__pop_tokens(2)

        error_number = self.integer_expression()
        if error_number is None:
            msg = "Error statement must be followed by an integer expression"
            raise self.__craft_error(msg)

        return Error(VIRTUAL_POSITION, error_number)

    # File statements
    def file_statement(self) -> Optional[FileStatement]:
        return self.__try_rules(  # type: ignore [return-value]
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

    @_add_rule_position
    @_with_backtracking
    def open(self) -> Optional[Open]:
        if self.__peek_token() != "Open":
            return None

        self.__pop_tokens(2)

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

        return Open(
            VIRTUAL_POSITION,
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

        return tuple(numbers)

    @_add_rule_position
    @_with_backtracking
    def close(self) -> Optional[Close]:
        if self.__peek_token() == "Reset":
            self.__pop_token()
            file_numbers = None
        elif self.__peek_token() == "Close":
            self.__pop_token()
            file_numbers = self.__file_number_list()
        else:
            return None

        return Close(VIRTUAL_POSITION, file_numbers)

    @_add_rule_position
    @_with_backtracking
    def seek(self) -> Optional[Seek]:
        if self.__peek_token() != "Seek":
            return None

        self.__pop_tokens(2)

        file_number = self.file_number()

        if file_number is None or self.__peek_token() != ",":
            msg = "Seek statement needs a position after the file number"
            raise self.__craft_error(msg)

        self.__pop_token()
        seek_position = self.expression()
        if seek_position is None:
            raise self.__craft_error("Seek statement needs valid position")

        return Seek(VIRTUAL_POSITION, file_number, seek_position)

    def __lock_unlock(self, lock: bool) -> Optional[Union[Lock, Unlock]]:
        node_type: Union[Type[Lock], Type[Unlock]]
        if lock:
            keyword, node_type = "Lock", Lock
        else:
            keyword, node_type = "Unlock", Unlock

        if self.__peek_token() != keyword:
            return None

        self.__pop_tokens(2)

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
                end_record_number = None
        else:
            start_record_number = None
            end_record_number = None

        assert node_type in (Lock, Unlock)
        return node_type(  # type: ignore [call-arg]
            VIRTUAL_POSITION,
            file_number,
            start_record_number,
            end_record_number,
        )

    @_add_rule_position
    @_with_backtracking
    def lock(self) -> Optional[Lock]:
        return self.__lock_unlock(lock=True)  # type: ignore [return-value]

    @_add_rule_position
    @_with_backtracking
    def unlock(self) -> Optional[Unlock]:
        return self.__lock_unlock(lock=False)  # type: ignore [return-value]

    @_add_rule_position
    @_with_backtracking
    def line_input(self) -> Optional[LineInput]:
        if self.__peek_tokens(0, 2) != ("Line", "Input"):
            return None

        self.__pop_tokens(4)
        file_number = self.file_number()
        if file_number is None:
            raise self.__craft_error("Line Input statement needs file number")
        if self.__pop_token() != ",":
            raise self.__craft_error("Line Input statement needs variable name")

        variable_name = self.variable_expression()
        if variable_name is None:
            raise self.__craft_error("Line Input statement needs variable name")

        return LineInput(VIRTUAL_POSITION, file_number, variable_name)

    @_add_rule_position
    @_with_backtracking
    def width(self) -> Optional[Width]:
        if self.__peek_token() != "Width":
            return None

        self.__pop_tokens(2)
        file_number = self.file_number()
        if file_number is None:
            raise self.__craft_error("Line Input statement needs file number")
        if self.__pop_token() != ",":
            raise self.__craft_error("Line Input statement needs variable name")

        line_width = self.expression()
        if line_width is None:
            raise self.__craft_error("Line Input statement needs variable name")

        return Width(VIRTUAL_POSITION, file_number, line_width)

    @_add_rule_position
    @_with_backtracking
    def output_item(self) -> Optional[OutputItem]:
        clause_title = None
        clause_argument = None
        char_position = None

        self.BLANK()
        if self.__peek_token() in (",", ";"):
            char_position = {
                ",": OutputItem.CharPosition.COMMA,
                ";": OutputItem.CharPosition.SEMICOLON,
            }[self.__pop_token()]
        else:
            clause_title, clause_argument = self.__output_clause()

            if self.__peek_token() in (",", ";"):
                char_position = {
                    ",": OutputItem.CharPosition.COMMA,
                    ";": OutputItem.CharPosition.SEMICOLON,
                }[self.__pop_token()]
            self.BLANK()

        if (clause_title, clause_argument, char_position) == (None, None, None):
            return None

        if char_position is None:
            char_position = OutputItem.CharPosition.SEMICOLON

        return OutputItem(
            VIRTUAL_POSITION, clause_title, clause_argument, char_position
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

        self.__pop_tokens(2)

        file_number = self.file_number()
        if file_number is None:
            raise self.__craft_error(keyword + " statement needs a file number")
        if self.__peek_token() != ",":
            msg = keyword + " statement needs an output list"
            raise self.__craft_error(msg)

        self.__pop_token()

        output_list = self.__output_list()

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

        return node_type(VIRTUAL_POSITION, file_number, output_list)

    @_add_rule_position
    @_with_backtracking
    def print(self) -> Optional[Write]:
        return self.__write_print(write=False)  # type: ignore [return-value]

    @_add_rule_position
    @_with_backtracking
    def write(self) -> Optional[Write]:
        return self.__write_print(write=True)  # type: ignore [return-value]

    @_add_rule_position
    @_with_backtracking
    def input(self) -> Optional[Input]:
        if not self.__peek_token() == "Input":
            return None

        self.__pop_tokens(2)
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

        return Input(VIRTUAL_POSITION, file_number, tuple(input_list))

    def __put_get(self, put: bool) -> Optional[Union[Put, Get]]:
        name, node_type = {True: ("Put", Put), False: ("Get", Get)}[put]

        if not self.__peek_token() == name:
            return None

        self.__pop_tokens(2)
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
            VIRTUAL_POSITION, file_number, record_number, data_variable
        )

    @_add_rule_position
    @_with_backtracking
    def put(self) -> Optional[Put]:
        return self.__put_get(put=True)  # type: ignore [return-value]

    @_add_rule_position
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
