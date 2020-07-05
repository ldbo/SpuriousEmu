"""
Preprocessing of the VBS file : handling of multiple instructions per line,
comments, etc.
"""

# Specifications reference:
# https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/statements
#
# Multiple instructions/lines:
#  - Separation with colon DONE
#  - Line ending in underscore DONE
#  - Implicit continuation TODO:
#       - Comma, concatenation operator DONE
#       - Parenthesis, brackets TODO
#       - Embedded expression TODO
#       - Assignment operators TODO
#       - Binary operators within an expression TODO
#       - Is and IsNot operators TODO
#       - Member qualifier character TODO
#       - XML property qualifier TODO
#       - < and > when specifying an attribute TODO
#       - before and after query operations TODO
#       - after IN in a For Each TODO
#       - after From in a collection initializer TODO
#
# Comments:
#  - ' and REM comments removed from instruction DONE
#  - Comments added in Instruction pretty-print TODO
#

from typing import List
from re import split


class PreprocessorError(Exception):
    """
    Error raised during preprocessing.
    """

    def __init__(self, file_name: str, line_number: int, message: str):
        self.file_name = file_name
        self.line_number = line_number
        self.message = message

    def __str__(self) -> str:
        return f"{self.file_name}:{self.line_number}: {self.message}"


class Instruction:
    """A single instruction, stored with its context."""

    def __init__(self, instruction: str, multiline: str, single: bool,
                 file_name: str, line_number: int):
        """
        :arg instruction: Unique instruction
        :arg multiline: Group of continuing lines the instruction is part of
        :arg single: Specify if the instruction is the only one in the
        multiline
        :arg file_name: Name of the file
        :arg line_number: Line of the instruction in the file
        """
        self.instruction = instruction
        self.multiline = multiline
        self.single = single
        self.file_name = file_name
        self.line_number = line_number

    def __str__(self) -> str:
        indent = "    "
        s = f"{self.file_name}:{self.line_number}:{indent}{self.instruction}"
        if not self.single:
            s += indent + "in\n" + indent
            s += self.multiline.replace('\n', '\n' + indent)

        return s

    def __repr__(self) -> str:
        return f"Instruction({self.instruction})"


class Preprocessor:
    """
    Used to extract single instructions from the content of a source file.
    """
    CONTINUING_LINE_DELIMITER = " _"
    CONTINUING_LINE_OPERATORS = (",", "&")

    def __init__(self):
        pass

    def extract_instructions(self, file_name, file_content) -> List[Instruction]:
        """
        Extract all the individual instructions from a file.

        :arg file_name: Name of the file
        :arg file_content: Content of the file
        :returns: The list of instructions, in the same order as in the file
        """
        self.__instructions: List[Instruction] = []
        self.__line_number = 1
        self.__instruction_line = 1
        self.__continuing_line = False
        self.__multiline = ""
        self.__concatenated_line = ""
        self.__file_name = file_name

        for line in file_content.splitlines():
            self.__handle_line(line)

        return self.__instructions

    def __handle_line(self, line: str) -> None:
        # Handle comments
        commentless_line = self.__remove_comments(line).strip()

        # Concatenate multi-line instructions
        if self.__continuing_line:
            self.__multiline += "\n"
        else:
            self.__concatenated_line = ""
            self.__instruction_line = self.__line_number
        self.__concatenated_line += commentless_line
        self.__multiline += line

        # If the instruction has not come to its end yet
        if commentless_line.endswith(self.CONTINUING_LINE_DELIMITER):
            self.__continuing_line = True
            del_chars = len(self.CONTINUING_LINE_DELIMITER)
            self.__concatenated_line = self.__concatenated_line[:-del_chars]
        elif any(commentless_line.strip().endswith(op)
                 for op in self.CONTINUING_LINE_OPERATORS):
            self.__continuing_line = True
            # If the instruction is complete
        else:
            instructions = self.__concatenated_line.split(':')
            single = len(instructions) == 1 \
                and "\n" not in self.__multiline
            for instruction in instructions:
                self.__add_instruction(instruction, single)

            self.__continuing_line = False
            self.__concatenated_line = ""
            self.__multiline = ""

        self.__line_number += 1

    def __add_instruction(self, instruction: str, single: bool) -> None:
        self.__instructions.append(Instruction(
            instruction=instruction.strip(),
            multiline=self.__multiline,
            single=single,
            file_name=self.__file_name,
            line_number=self.__instruction_line
        ))

    @classmethod
    def __remove_comments(cls, line: str) -> str:
        return split("REM|'", line)[0]
