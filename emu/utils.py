from abc import abstractmethod
from dataclasses import dataclass
from typing import Generic, TypeVar


@dataclass
class FilePosition:
    """
    Position of a slice in a code chunk.
    """

    IEOLS = ("\r", "\n", "\u2028", "\u2029")  #: Individual line terminators
    EOLS = ("\r\n", "\r", "\n", "\u2028", "\u2029")  #: Line terminators

    file_name: str  #: Content of the whole file
    file_content: str  #: Name of the file
    start_index: int  #: Start index in the file, starting at 0, inclusive
    end_index: int
    """End index in the file, starting at 0, exclusive, must be
    ``>= start_index``"""
    start_line: int  #: Start line, starting at 1, inclusive
    end_line: int  #: End line, starting at 1, inclusive
    start_column: int  #: Start column, starting at 1, inclusive
    end_column: int  #: End column, starting at 1, exclusive

    def __post_init__(self) -> None:
        if not self.start_index <= self.end_index:
            msg = "Can't create a FilePosition with end_index < start_index"
            raise RuntimeError(msg)

    def start_of_line_index(self) -> int:
        """
        Returns:
          The index of the start of the first line of the position.
        """
        index = self.start_index
        while index > 0:
            if self.file_content[index] in self.IEOLS:
                index += 1
                break
            index -= 1

        return max(index, 0)

    def end_of_line_index(self) -> int:
        """
        Returns:
          The index of the character after the end of the last line of the
          position.
        """
        index = max(self.end_index - 1, self.start_index)
        while index < len(self.file_content):
            if self.file_content[index] in self.IEOLS:
                break
            index += 1

        return index

    def header(self) -> str:
        """
        Returns:
           The ``{file_name}:{start_line}:{start_column}:`` string
        """
        return "{}:{}:{}:".format(
            self.file_name, self.start_line, self.start_column
        )

    def body(self) -> str:
        """
        Returns:
          The text between :py:attr:`start_index` and :py:attr:`end_index`
        """
        return self.file_content[self.start_index : self.end_index]

    def lines(self) -> str:
        """Return the whole lines that contain the position"""
        end_index = self.end_of_line_index()
        start_index = self.start_of_line_index()
        return self.file_content[start_index:end_index]

    @classmethod
    def from_indices(
        cls,
        file_name: str,
        file_content: str,
        start_index: int,
        end_index: int,
        start_line: int,
        start_column: int,
    ) -> "FilePosition":
        """Build a position without giving the end line and column"""
        # Find end_line: count new lines

        index = start_index
        end_line = start_line
        last_line_index = 0
        while index < end_index:
            line_found = False
            for eol in FilePosition.EOLS:
                if file_content[index:].startswith(eol):
                    index += len(eol)
                    end_line += 1
                    last_line_index = index
                    line_found = True
                    break

            if not line_found:
                index += 1

        if end_line == start_line:
            end_column = start_column + end_index - start_index
        else:
            end_column = end_index - last_line_index + 1

        return FilePosition(
            file_name,
            file_content,
            start_index,
            end_index,
            start_line,
            end_line,
            start_column,
            end_column,
        )


T = TypeVar("T")  #: Output type of the Visitor DP algorithm


class Visitable:
    """Base class for Visitable classes, must simply be inherited."""

    @abstractmethod
    def __init__(self):
        pass

    def accept(self, visitor: "Visitor[T]") -> T:
        """
        Method called by a visiting visitor, should not be overriden.

        Raises:
          :py:exc:`NotImplementedError`: If the visitor does not implement the
                                         expected ``visit_`` method
        """
        assert isinstance(visitor, Visitor)

        visiter = "visit_" + type(self).__qualname__.replace(".", "_")
        try:
            visiter_function = getattr(visitor, visiter)
        except AttributeError:
            msg = (
                f"Visitor {type(visitor).__qualname__} doesn't handle "
                + f"Visitable type {type(self).__qualname__}"
            )
            raise NotImplementedError(msg)

        return visiter_function(self)


class Visitor(Generic[T]):
    """
    Base class for Visitor classes. For a subclass to be able to handle the
    :py:class:`Visitable` type ``U``, it must implement a method
    ``visit_U(self, visitable : U) -> T``. Then, to apply the algorithm to a
    :py:class:`Visitable` object ``u``, simply call ``visitor.visit(u)``. Use
    the class argument ``T`` to specify the output type of the ``visit_``
    methods.
    """

    @abstractmethod
    def __init__(self):
        pass

    def visit(self, visitable: Visitable) -> T:
        """
        Returns:
          The output of the algorithm in ``visit_U``

        Raises:
          :py:exc:`NotImplementedError`: If the ``visit_U`` method is not
                                         implemented
        """
        assert isinstance(visitable, Visitable)

        return visitable.accept(self)
