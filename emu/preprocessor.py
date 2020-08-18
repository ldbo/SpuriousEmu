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
#  - Comments added in Instruction pretty-print TODO

from typing import Any, Dict, List, Optional

from .error import PreprocessorError


class Instruction:
    """A single instruction, stored with its context."""
    instruction: str
    multiline: str
    single: bool
    file_name: str
    line_number: int

    def __init__(self, instruction: str, multiline: str, single: bool,
                 file_name: str, line_number: int) -> None:
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

    def to_dict(self) -> Dict[str, Any]:
        return self.__dict__


class Preprocessor:
    """
    Used to extract single instructions from the content of a source file.
    """
    CONTINUING_LINE_DELIMITER = " _"
    CONTINUING_LINE_OPERATORS = (",", "&")
    COMMENT_DELIMITERS = ("REM", "'")
    MULTIPLE_INSTRUCTIONS_DELIMITER = ":"

    def __init__(self) -> None:
        pass

    def extract_instructions(self, file_name: str, file_content: str) \
            -> List[Instruction]:
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
            instructions = self.__split_line(self.__concatenated_line)
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
        """
        Search for the first REM or ' character not enclosed in a string or in
        another token
        """
        in_string = False
        position = 0

        if any(line.startswith(f"{d} ") for d in cls.COMMENT_DELIMITERS):
            return ''

        while position < len(line):
            char = line[position]
            remaining_line = line[position:]

            if in_string:
                if remaining_line.startswith('""'):
                    position += 1
                elif char == '"':
                    in_string = False
            else:
                if any(remaining_line.startswith(f" {d} ")
                       for d in cls.COMMENT_DELIMITERS):
                    return line[:position]
                elif char == '"':
                    in_string = True

            position += 1

        return line

    @classmethod
    def __split_line(cls, line: str) -> List[str]:
        """
        Split a line containing no comments into its different instructions,
        ignoring separators inside strings.
        """
        in_string = False
        instruction_start = 0
        position = 0
        delim = cls.MULTIPLE_INSTRUCTIONS_DELIMITER
        instructions = []

        while position < len(line):
            char = line[position]
            remaining_line = line[position:]

            if in_string:
                if remaining_line.startswith('""'):
                    position += 1
                elif char == '"':
                    in_string = False
            else:
                if remaining_line.startswith(delim):
                    instructions.append(line[instruction_start:position])
                    instruction_start = position + 1
                elif char == '"':
                    in_string = True

            position += 1

        instructions.append(line[instruction_start:position])

        return instructions

    @staticmethod
    def preprocess(file_content: str, file_name: Optional[str] = None) \
            -> List[Instruction]:
        """
        Extract the executable instructions of a file, returning them as a
        list.
        """
        preprocessor = Preprocessor()
        return preprocessor.extract_instructions(file_name, file_content)

    @staticmethod
    def preprocess_file(file_path: str) -> List[Instruction]:
        with open(file_path, 'r') as f:
            file_content = f.read()

        return Preprocessor.preprocess(file_content, file_path)
