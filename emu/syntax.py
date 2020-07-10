"""Define the grammar of a VBA source file and implement a syntactic parser."""

from typing import List, Union

from pyparsing import (Forward, Optional, ParseException, ParserElement,
                       Regex, Suppress, StringEnd, StringStart, Word,
                       delimitedList, infixNotation, nums, oneOf, opAssoc)

from .abstract_syntax_tree import *
from .partial_block import *
from .partial_block import PartialBlock
from .preprocessor import Instruction, Preprocessor
from .type import types

#############
#  Grammar  #
#############

expression = Forward().setName("expression")
statement = Forward().setName("statement")

# Punctuation
lparen = Suppress("(")
rparen = Suppress(")")

# Keywords
as_kw = Suppress('As')
dim_kw = Suppress('Dim')
for_kw = Suppress('For')
next_kw = Suppress('Next')
set_kw = Suppress('Set')
step_kw = Suppress('Step')
to_kw = Suppress('To')

# Literal
integer = Word(nums).setName("integer") \
    .setParseAction(lambda r: Literal(Type.Integer, r[0]))
boolean = oneOf("True False").setName("boolean") \
    .setParseAction(lambda r: Literal(Type.Boolean, r[0]))
literal = (integer | boolean).setName("literal")

# Identifier
element_regex = r"(?:[a-zA-Z]|_[a-zA-Z])[a-zA-Z0-9_]*"
identifier_regex = rf"{element_regex}(?:\.{element_regex})*"
identifier = Regex(identifier_regex).setName("identifier") \
    .setParseAction(lambda r: Identifier(r[0]))

# Types
variable_type = oneOf(types)


#################
#  Expressions  #
#################

# Operator


def __build_binary_operator(expr, pos, result):
    tokens = result[0].asList()
    operators = tokens[1::2]
    operands = tokens[0::2]

    tree = operands[0]
    for operator, operand in zip(operators, operands[1:]):
        tree = BinOp(operator, tree, operand)

    return tree


binary_operators = [
    ("^", 2, opAssoc.LEFT, __build_binary_operator),
    (oneOf("* /"), 2, opAssoc.LEFT, __build_binary_operator),
    ("\\", 2, opAssoc.LEFT, __build_binary_operator),
    ("Mod", 2, opAssoc.LEFT, __build_binary_operator),
    (oneOf("+ -"), 2, opAssoc.LEFT, __build_binary_operator),
    ("&", 2, opAssoc.LEFT, __build_binary_operator)
]

# Function call
arguments_list = Optional(delimitedList(expression)).setName("arguments_list") \
    .setParseAction(lambda r: ArgList(list(r)))
function_call_paren = (identifier + lparen + arguments_list + rparen) \
    .setName("function_call_paren") \
    .setParseAction(lambda r: FunCall(r[0], r[1]))
function_call_no_paren = (StringStart() + identifier + arguments_list
                          + StringEnd()) \
    .setName("function_call_no_paren") \
    .setParseAction(lambda r: FunCall(r[0], r[1]))

terminal = (function_call_paren | literal | identifier)
expression << infixNotation(
    terminal, binary_operators, lpar=lparen, rpar=rparen)
expression_statement = function_call_no_paren ^ expression

##################
#  Declarations  #
##################

# Variable
variable_declaration = (dim_kw + identifier
                        + Optional(as_kw + variable_type
                                   + Optional(Suppress('=') + expression))) \
    .setParseAction(lambda r: VarDec(*r))

variable_assignment = (set_kw + identifier + Suppress('=') + expression) \
    .setParseAction(lambda r: VarAssign(*r))

declarative_statement = variable_declaration | variable_assignment

############################
#  Loops and conditionals  #
############################

# For
for_header = (for_kw + identifier + Suppress('=') + expression
              + to_kw + expression + Optional(step_kw + expression)) \
    .setParseAction(lambda r: ForHeader(*r))
for_footer = (next_kw + Optional(identifier)) \
    .setParseAction(lambda r: ForFooter(*r))

loop_statement = for_header | for_footer

# Wrap up
statement <<= declarative_statement | loop_statement | expression_statement

############
#  Parser  #
############

# Packrat parsing cache parsing results, improving speed
ParserElement.enablePackrat(cache_size_limit=128)


class ParsingError(Exception):
    """
    Error raised during parsing.
    """

    def __init__(self, file_name: str, line_number: int, message: str):
        self.file_name = file_name
        self.line_number = line_number
        self.message = message

    def __str__(self) -> str:
        return f"{self.file_name}:{self.line_number}: {self.message}"


class Parser:
    __nested_blocks: List[PartialBlock]

    def __init__(self):
        pass

    def build_ast(self, instructions: List[Instruction], file_name: str = "") \
            -> AST:
        self.__nested_blocks = [PartialBlock()]
        for instruction in instructions:
            if instruction.instruction.strip() != '':
                parsed_instruction = self.__parse_instruction(instruction)
                if isinstance(parsed_instruction, Statement):
                    self.__handle_statement(parsed_instruction)
                elif isinstance(parsed_instruction, BlockElement):
                    self.__handle_block_element(parsed_instruction)
                else:
                    assert(False)

        top_level_block = self.__nested_blocks.pop()
        main_sequence = Block(top_level_block.statements)
        return main_sequence

    def __parse_instruction(self, instruction: Instruction) -> Union[Statement, BlockElement]:
        try:
            parse_results = statement.parseString(
                instruction.instruction, parseAll=True)
        except ParseException as e:
            raise ParsingError(instruction.file_name, instruction.line_number,
                               str(e))

        return parse_results[0]

    def __handle_statement(self, statement: Statement) -> None:
        current_block = self.__nested_blocks[-1]
        current_block.statements.append(statement)

    def __handle_block_element(self, block_element: BlockElement) -> None:
        if block_element.FirstElement:
            new_block = PartialBlock([block_element])
            self.__nested_blocks.append(new_block)
        else:
            self.__nested_blocks[-1].elements.append(block_element)

        if block_element.LastElement:
            complete_block = self.__nested_blocks.pop().build_block()
            parent_block = self.__nested_blocks[-1]
            parent_block.statements.append(complete_block)


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
