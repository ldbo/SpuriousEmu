
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

    def __init__(self, file_name: str, line_number: int, message: str):
        self.file_name = file_name
        self.line_number = line_number
        self.message = message

    def __str__(self) -> str:
        return f"{self.file_name}:{self.line_number}: {self.message}"


class InterpretationError(Exception):
    pass


class CompilationError(Exception):
    pass


class OperatorError(Exception):
    pass


class ConversionError(Exception):
    pass


class ResolutionError(Exception):
    pass
