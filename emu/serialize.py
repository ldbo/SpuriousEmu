"""Allow serialization of SpuriousEmu objects"""

import pickle

from dataclasses import dataclass
from pathlib import Path
from typing import List, Type, Union

from .compiler import Program
from .error import SerializationError
from .side_effect import OutsideWorld


class Serializer:
    """
    Set of class methods allowing to serialize, deserialize, save and load
    SpuriousEmu objects.

    The specifications of the different supported formats, ie. their magic
    string, corresponding Python classes, and optional extensions, are stored
    in Serialize.FORMATS.
    """
    @dataclass
    class FileFormat:
        object_type: type
        magic_string: bytes
        extension: str

    FORMATS: List[FileFormat] = [
        FileFormat(Program, b'SpuriousEmuProgram', ".spemu-com"),
        FileFormat(OutsideWorld, b'SpuriousEmuOutsideWorld', ".spemu-out")
    ]

    SerializableType = Union[Program, OutsideWorld]

    @classmethod
    def serialize(cls, obj: SerializableType) -> bytes:
        if not isinstance(obj, Serializer.SerializableType.__args__):
            msg = f"Serialization of {type(obj)} is not supported"
            raise SerializationError(msg)

        body = pickle.dumps(obj)
        magic = cls.magic(type(obj))

        return magic + body

    @classmethod
    def save(cls, obj: SerializableType, file_path: str) -> None:
        """
        Serialize a Python object and save it to a file, potentially overriding
        it. If the save path has no extension, add the corresponding one.
        """
        path = Path(file_path)
        print(path.suffix)
        if path.suffix == "":
            save_path = path.parent / (path.name + cls.extension(type(obj)))
        else:
            save_path = path

        content = cls.serialize(obj)

        with open(save_path, 'wb') as f:
            f.write(content)

    @classmethod
    def deserialize(cls, content: bytes) -> SerializableType:
        for fmt in cls.FORMATS:
            magic = fmt.magic_string
            if content.startswith(magic):
                body = content[len(magic):]
                return pickle.loads(body)

        raise SerializationError("Unsupported format")

    @classmethod
    def load(cls, path: str) -> SerializableType:
        """
        Try to deserialize the content of the given file, based on its magic
        string.
        """
        with open(path, 'rb') as f:
            file_content = f.read()

        return cls.deserialize(file_content)

    @classmethod
    def magic(cls, object_type: type) -> bytes:
        for fmt in cls.FORMATS:
            if fmt.object_type == object_type:
                return fmt.magic_string

        msg = f"Serialization of {object_type} is not supported"
        raise SerializationError(msg)

    @classmethod
    def extension(cls, object_type: type) -> str:
        for fmt in cls.FORMATS:
            if fmt.object_type == object_type:
                return fmt.extension

        msg = f"Serialization of {object_type} is not supported"
        raise SerializationError(msg)

    @classmethod
    def type(cls, magic: bytes) -> "Serializer.SerializableType":
        for fmt in cls.FORMATS:
            if fmt.magic_string == magic:
                return fmt.object_type

        msg = f"Magic string {magic.decode('utf-8')} is not supported"
        raise SerializationError(msg)
