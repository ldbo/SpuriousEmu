from .utils import FilePosition


class LexerError(Exception):
    """Error raised during lexing, usually because of a malformed document."""

    message: str  #: Error message
    position: FilePosition  #: Position of the token causing the error

    def __init__(self, message: str, position: FilePosition) -> None:
        self.message = message
        self.position = position

    def __str__(self) -> str:
        header = self.position.header() + " " + self.message
        lines = self.position.lines()
        caret_pos = self.position.start_column - 1
        underline_len = self.position.end_index - self.position.start_index
        error_position = " " * caret_pos + "^" + "~" * (underline_len - 1)

        return header + "\n" + lines + "\n" + error_position + "\n"
