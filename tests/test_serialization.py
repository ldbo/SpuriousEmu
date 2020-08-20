# TODO add infile tests

from nose.tools import assert_equals

from emu import Serializer, Compiler
from tests.test import source_path


def test_serialization_deserialization():
    program = Compiler.compile_file(source_path("interpreter_01"))
    serial = Serializer.serialize(program)
    deserial = Serializer.deserialize(serial)
    assert_equals(type(program), type(deserial))
    assert_equals(program.to_dict(), deserial.to_dict())
