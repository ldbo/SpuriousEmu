"""Define the grammar of a VBA source file and implement a syntactic parser."""

from type import Type

from abstract_syntax_tree import *
from preprocessor import Instruction, Preprocessor

from pyparsing import (Forward, Optional, ParseException, Regex, Suppress, Word,
                       delimitedList, infixNotation, nums, oneOf, opAssoc)

from typing import List

#############
#  Grammar  #
#############

expression = Forward().setName("expression")
statement = Forward().setName("statement")
terminal = Forward().setName("terminal")

# Literal
integer = Word(nums).setParseAction(lambda r: Literal(Type.Integer, r[0]))
boolean = oneOf("True False").setParseAction(
    lambda r: Literal(Type.Boolean, r[0]))
literal = (integer | boolean)

# Identifier
element_regex = r"(?:[a-zA-Z]|_[a-zA-Z])[a-zA-Z0-9_]*"
identifier_regex = rf"{element_regex}(?:\.{element_regex})*"
identifier = Regex(identifier_regex).setParseAction(lambda r: Identifier(r[0]))

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

terminal = (literal | identifier)
expression << infixNotation(
    terminal, binary_operators, lpar=Suppress('('), rpar=Suppress(')'))


###########
#  Parser #
###########


class SyntaxError(Exception):
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
            parse_results = expression.parseString(
                instruction.instruction, parseAll=True)
        except ParseException as e:
            raise SyntaxError(instruction.file_name, instruction.line_number,
                              str(e))

        return parse_results[0]


if __name__ == "__main__":
    from pprint import pprint

    file = "tests/basic_01.vbs"

    preprocessor = Preprocessor()
    parser = Parser()

    instructions = preprocessor.extract_instructions_from_file(file)
    tree = parser.build_ast(instructions)

    pprint(tree.to_dict())
