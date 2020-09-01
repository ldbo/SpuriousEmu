from typing import Optional


class PreprocessorError(Exception):
    """
    Error raised during preprocessing.
    """

    file_name: str
    line_number: int
    message: str

    def __init__(self, file_name: str, line_number: int, message: str) -> None:
        self.file_name = file_name
        self.line_number = line_number
        self.message = message

    def __str__(self) -> str:
        return f"{self.file_name}:{self.line_number}: {self.message}"


class ParsingError(Exception):
    """
    Error raised during parsing.
    """

    def __init__(
        self,
        file_name: str,
        line: int,
        column: int,
        message: str,
        end_line: Optional[int] = None,
        end_column: Optional[int] = None,
    ) -> None:
        self.file_name = file_name
        self.line = line
        self.column = column
        self.message = message
        self.end_line = end_line
        self.end_column = end_column

    def __str__(self) -> str:
        return f"{self.file_name}:{self.line}: {self.message}"

    def pretty(self, file_content: str) -> str:
        content = str(self) + "\n"
        lines = file_content.split("\n")
        end_line = self.end_line if self.end_line is not None else self.line
        content += "\n".join(lines[self.line : end_line + 1])
        return content


class InterpretationError(Exception):
    pass


class CompilationError(Exception):
    pass


class DeobfuscationError(Exception):
    pass


class OperatorError(Exception):
    pass


class ConversionError(Exception):
    pass


class ResolutionError(Exception):
    pass


class SerializationError(Exception):
    pass
