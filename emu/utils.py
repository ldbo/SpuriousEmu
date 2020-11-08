from abc import abstractmethod
from dataclasses import dataclass
from enum import Enum
from typing import Any, Dict, Generic, Iterable, Optional, Set, TypeVar, cast


@dataclass(frozen=True)
class FilePosition:
    """
    Position of a slice in a code chunk.
    """

    EXCLUDED = {"file_content"}

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

    def __add__(self, position: "FilePosition") -> "FilePosition":
        """
        Returns:
          The smallest position enclosing ``self`` and ``position``

        Raises:
          :py:exc:`RuntimeError`: If the two positions are not in the same file
        """
        if self.file_name != position.file_name:
            msg = "Can't add positions that are not in the same file"
            raise RuntimeError(msg)

        file_name = self.file_name
        file_content = self.file_content

        if self.start_index <= position.start_index:
            start_index = self.start_index
            start_column = self.start_column
            start_line = self.start_line
        else:
            start_index = position.start_index
            start_column = position.start_column
            start_line = position.start_line

        if self.end_index >= position.start_index:
            end_index = self.end_index
            end_column = self.end_column
            end_line = self.end_line
        else:
            end_index = position.end_index
            end_column = position.end_column
            end_line = position.end_line

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
        Method called by a visiting visitor, should not be overridden.

        Raises:
          :py:exc:`NotImplementedError`: If the visitor does not implement the
                                         expected ``visit_`` method

        Todo:
          * Search in the inheritance tree for a supported ``visit_`` method
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


_BASE_TYPES = {bool, int, float, str, bytes, type(None)}
MissingOptional = cast(None, object())
"""Can be used instead of ``None`` for default value of e.g. class attributes"""


def to_dict(obj: object, excluded_fields: Optional[Set[str]] = None) -> Any:
    """
    Recursively convert an arbitrary object to a dictionary.

    Arguments:
      obj: Any object

    Returns:
      If the type of ``obj`` is

        - One of :attr:`emu.utils._BASE_TYPES`: don't convert the object
        - :class:`tuple`, :class:`list` or :class:`dict`: recurse into the
          object, returning an instance of the same type
        - :class:`type`: the name of the type
        - an :class:`enum.Enum` member: the string ``<Class.VALUE>``
        - another type: a dictionary with field ``TYPE = type(obj)``, and which
          contains all the public non static attributes of the object (ie. those
          that don't start with ``_`` and are not uppercase) that are not in
          :attr:`obj.EXCLUDED`. If this type inherits a base data type, add a
          ``VALUE`` field containing its value.
    """
    t = type(obj)
    if t in (list, tuple):
        return t(to_dict(elt, excluded_fields) for elt in obj)  # type: ignore
    elif t == dict:
        return t(  # type: ignore [call-arg]
            (to_dict(key, excluded_fields), to_dict(value, excluded_fields))
            for key, value in obj.items()  # type: ignore [attr-defined]
        )
    elif t in _BASE_TYPES:
        if t == bytes:
            return obj.decode("utf-16")  # type: ignore [attr-defined]
        return obj
    elif t == type:
        return obj.__name__  # type: ignore [attr-defined]
    elif isinstance(obj, Enum):
        return f"<{type(obj).__name__}.{obj.name}>"
    else:
        d = _object_to_dict(obj, excluded_fields)
        _update_base_type_dict(obj, d)
        return d


def _object_to_dict(obj: object, excluded_fields: Optional[Set[str]]) -> Any:
    try:
        excluded = set(getattr(obj, "EXCLUDED"))
        excluded = excluded | {"EXCLUDED"}
    except AttributeError:
        excluded = set()

    if excluded_fields is not None:
        excluded = excluded | excluded_fields

    try:
        attributes = list(vars(obj).keys())
        used_dirs = False
    except TypeError:
        attributes = dir(obj)
        used_dirs = True

    items = [("TYPE", type(obj).__name__)]
    for field in attributes:
        if field.startswith("_") or field == field.upper():
            continue

        if field in excluded:
            continue

        value = getattr(obj, field)
        if (used_dirs and callable(value)) or value is MissingOptional:
            continue

        items.append((field, to_dict(value, excluded_fields)))

    return dict(items)


def _update_base_type_dict(obj: object, temp_dict: Dict[str, Any]) -> None:
    for t in _BASE_TYPES:
        if isinstance(obj, t):
            temp_dict["VALUE"] = t(obj)
