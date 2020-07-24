"""Define the grammar of a VBA source file and implement a syntactic parser."""

from typing import List, Union

from pyparsing import (Forward, ParseException, ParserElement, QuotedString,
                       Regex, Suppress, StringEnd, StringStart, Word,
                       delimitedList, infixNotation, nums, oneOf, Keyword)
from pyparsing import Optional as pOptional

from .abstract_syntax_tree import *
from .partial_block import *
from .preprocessor import Instruction, Preprocessor
from .type import types
from .operator import *
from .error import ParsingError

#############
#  Grammar  #
#############

expression = Forward().setName("expression")
statement = Forward().setName("statement")

# Punctuation
lparen = Suppress("(")
rparen = Suppress(")")

# Keywords
as_kw = Suppress('As').setName('as')
dim_kw = Suppress('Dim').setName('dim')
else_kw = Suppress('Else').setName('else')
elseif_kw = Suppress('ElseIf').setName('elseif')
end_kw = Suppress('End').setName('end')
for_kw = Suppress('For').setName('for')
if_kw = Suppress('If').setName('if')
next_kw = Suppress('Next').setName('next')
set_kw = Suppress('Set').setName('set')
step_kw = Suppress('Step').setName('step')
sub_kw = Suppress('Sub').setName('sub')
then_kw = Suppress('Then').setName('then')
to_kw = Suppress('To').setName('to')
function_kw = Suppress('Function').setName('function')
true_kw = Keyword('True')
false_kw = Keyword('False')

reserved = as_kw | dim_kw | else_kw | elseif_kw | end_kw | for_kw | if_kw \
    | next_kw | set_kw | step_kw | sub_kw | then_kw | to_kw | function_kw \
    | true_kw | false_kw

# Literal
integer = Word(nums).setName("integer") \
    .setParseAction(lambda r: Literal(Type.Integer, r[0]))
boolean = (true_kw | false_kw).setName("boolean") \
    .setParseAction(lambda r: Literal(Type.Boolean, r[0]))
string = QuotedString(quoteChar='"', escQuote='""').setName("string") \
    .setParseAction(lambda r: Literal(Type.String, r[0]))
literal = (integer | boolean | string).setName("literal")

# Identifier
element_regex = r"(?:[a-zA-Z]|_[a-zA-Z])[a-zA-Z0-9_]*"
identifier_regex = rf"{element_regex}(?:\.{element_regex})*"
identifier = (~reserved + Regex(identifier_regex)).setName("identifier") \
    .setParseAction(lambda r: Identifier(r[0]))

# Types
variable_type = oneOf(types)


#################
#  Expressions  #
#################

# Operator
def __build_binary_operator(expr, pos, result):
    tokens = result[0].asList()
    operator_symbols = tokens[1::2]
    operands = tokens[0::2]

    tree = operands[0]
    for operator_symbol, operand in zip(operator_symbols, operands[1:]):
        operator = BinaryOperator.build_operator(operator_symbol)
        tree = BinOp(operator, tree, operand)

    return tree


binary_operators = OPERATORS_MAP.get_precedence_list(__build_binary_operator)

# Function call
arguments_list = pOptional(delimitedList(expression)) \
    .setName("arguments_list") \
    .setParseAction(lambda r: ArgList(list(r)))
function_call_paren = (identifier + lparen + arguments_list + rparen) \
    .setName("function_call_paren") \
    .setParseAction(lambda r: FunCall(r[0], r[1]))
function_call_no_paren = (StringStart() + identifier + arguments_list
                          + StringEnd()) \
    .setName("function_call_no_paren") \
    .setParseAction(lambda r: FunCall(r[0], r[1]))

terminal = (literal | function_call_paren | identifier)
expression << infixNotation(
    terminal, binary_operators, lpar=lparen, rpar=rparen) \
    .setName('expression')
expression_statement = (function_call_no_paren | expression) \
    .setName("expression_statement")

##################
#  Declarations  #
##################

# Variable
variable_declaration = (dim_kw + identifier
                        + pOptional(as_kw + variable_type
                                    + pOptional(Suppress('=') + expression))) \
    .setParseAction(lambda r: VarDec(*r)) \
    .setName("var dec")

variable_assignment = (set_kw + identifier + Suppress('=') + expression) \
    .setParseAction(lambda r: VarAssign(*r)) \
    .setName("var assign")

# Function

procedure_header = (sub_kw + identifier
                    + pOptional(lparen + arguments_list + rparen)) \
    .setParseAction(lambda r: ProcDefHeader(*r)) \
    .setName("proc header")
procedure_footer = (end_kw + sub_kw) \
    .setParseAction(lambda r: ProcDefFooter()) \
    .setName("proc footer")

function_header = (function_kw + identifier
                   + pOptional(lparen + arguments_list + rparen)) \
    .setParseAction(lambda r: FunDefHeader(*r)) \
    .setName("function header")
function_footer = (end_kw + function_kw) \
    .setParseAction(lambda r: FunDefFooter()) \
    .setName("function footer")


declarative_statement = (
    variable_declaration | variable_assignment
    | procedure_header | procedure_footer
    | function_header | function_footer) \
    .setName("declaration")

############################
#  Loops and conditionals  #
############################

# For
for_header = (for_kw + identifier + Suppress('=') + expression
              + to_kw + expression + pOptional(step_kw + expression)) \
    .setParseAction(lambda r: ForHeader(*r))
for_footer = (next_kw + pOptional(identifier)) \
    .setParseAction(lambda r: ForFooter(*r))

loop_statement = for_header | for_footer

# If
if_header = (if_kw + expression + pOptional(then_kw)) \
    .setParseAction(lambda r: IfHeader(*r))
elseif_header = (elseif_kw + expression + pOptional(then_kw)) \
    .setParseAction(lambda r: ElseIfHeader(*r))
else_header = else_kw.setParseAction(lambda r: ElseHeader())
if_footer = (end_kw + if_kw).setParseAction(lambda r: IfFooter())

conditional_statement = if_header | elseif_header | else_header | if_footer

# Wrap up
statement <<= declarative_statement | loop_statement | conditional_statement \
    | expression_statement

############
#  Parser  #
############

# Packrat parsing cache parsing results, improving speed
ParserElement.enablePackrat(cache_size_limit=128)


class Parser:
    """
    Syntactic parser used to transform a list of instructions into an abstract
    syntax tree.
    """
    __nested_blocks: List[PartialBlock]

    def __init__(self) -> None:
        pass

    def build_ast(self, instructions: List[Instruction], file_name: str = "") \
            -> AST:
        """
        Main method of the Parser. Parses each of the instructions using a
        pyparsing grammar, and then process the results to handle multiline
        structures such as loops, conditionals, etc.
        """
        self.__nested_blocks = [PartialBlock()]
        for instruction in instructions:
            if instruction.instruction.strip() != '':
                parsed_instruction = self.__parse_instruction(instruction)
                if isinstance(parsed_instruction, Statement):
                    self.__handle_statement(parsed_instruction)
                elif isinstance(parsed_instruction, BlockElement):
                    self.__handle_block_element(parsed_instruction)
                else:
                    assert (False)

        top_level_block = self.__nested_blocks.pop()
        main_sequence = Block(top_level_block.statements_blocks.pop())
        return main_sequence

    def __parse_instruction(self, instruction: Instruction)\
            -> Union[Statement, BlockElement]:
        """Use the pyparsing grammar to parse a single instruction."""
        try:
            parse_results = statement.parseString(
                instruction.instruction, parseAll=True)
        except ParseException as e:
            raise ParsingError(instruction.file_name, instruction.line_number,
                               str(e))

        return parse_results[0]

    def __handle_statement(self, statement: Statement) -> None:
        """
        Add a single-line statement to the last statements block of the
        innermost nest block.
        """
        current_block = self.__nested_blocks[-1]
        current_block.statements_blocks[-1].append(statement)

    def __handle_block_element(self, block_element: BlockElement) -> None:
        """
        Add a block element to the list of elements of the innermost partial
        block. Create it or add it to the statements of the parent block if
        needed.
        """
        if block_element.IsHeader:
            new_block = PartialBlock([block_element])
            self.__nested_blocks.append(new_block)
        else:
            self.__nested_blocks[-1].elements.append(block_element)
            if not block_element.IsFooter:
                self.__nested_blocks[-1].statements_blocks.append([])

        if block_element.IsFooter:
            complete_block = self.__nested_blocks.pop().build_block()
            parent_block = self.__nested_blocks[-1]
            parent_block.statements_blocks[-1].append(complete_block)


def parse_file(path: str) -> AST:
    """
    Parse a file into an abstract syntax tree.

    :arg path: Path of the file
    :return: An AST representing the syntax of the file
    """
    preprocessor = Preprocessor()
    parser = Parser()

    instructions = preprocessor.extract_instructions_from_file(path)
    tree = parser.build_ast(instructions)

    return tree
