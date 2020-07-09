"""Define the grammar of a VBA source file and implement a syntactic parser."""

from .type import Type, types

from .abstract_syntax_tree import *
from .preprocessor import Instruction, Preprocessor

from pyparsing import (Forward, Optional, Or, ParseException, ParserElement,
                       Regex, Suppress, StringEnd, StringStart, Word,
                       delimitedList, infixNotation, nums, oneOf, opAssoc)

from typing import List

#############
#  Grammar  #
#############

expression = Forward().setName("expression")
statement = Forward().setName("statement")
terminal = Forward().setName("terminal")

# Punctuation
lparen = Suppress("(")
rparen = Suppress(")")

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

##################
#  Declarations  #
##################

# Variable
variable_declaration = (Suppress('Dim') + identifier
                        + Optional(Suppress("As") + variable_type
                                   + Optional(Suppress('=') + expression))) \
    .setParseAction(lambda r: VarDec(*r))

statement <<= function_call_no_paren ^ expression

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
    def __init__(self):
        pass

    def build_ast(self, instructions: List[Instruction], file_name: str = "") \
            -> AST:
        nodes = []
        for instruction in instructions:
            if instruction.instruction.strip() != '':
                nodes.append(self.__parse_instruction(instruction))

        return Sequence(nodes)

    def __parse_instruction(self, instruction: Instruction) -> Statement:
        try:
            parse_results = statement.parseString(
                instruction.instruction, parseAll=True)
        except ParseException as e:
            raise ParsingError(instruction.file_name, instruction.line_number,
                               str(e))

        return parse_results[0]


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


if __name__ == "__main__":
    from json import load
    from pprint import pprint

    file = "tests/basic_01.vbs"
    with open('tests/basic_01.json') as f:
        expected_result = load(f)

    preprocessor = Preprocessor()
    parser = Parser()

    instructions = preprocessor.extract_instructions_from_file(file)
    tree = parser.build_ast(instructions)

    assert(expected_result == tree.to_dict())
    print(f'Test: {file} OK')
