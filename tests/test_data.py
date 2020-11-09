from nose.tools import assert_raises, assert_equals

from emu.data import *
from emu.utils import to_dict

from pprint import pprint


def test_init():
    """
    data: value initialization
    """
    assert_equals(Variable(Boolean).value, Boolean(False))
    assert_equals(Variable(Long).value, Long(0))
    assert_equals(Variable(Date).value, Date(0.0))
    assert_equals(Variable(Variant()).value, Empty())
    assert_equals(Variable(String).value, String(b""))
    assert_equals(Variable(StringN(12)).value, String(b"\x00" * 24))
    assert_equals(
        Variable(FVArray(((1, 2), (2, 4)))).value.bounds, ((1, 2), (2, 4))
    )
    assert_equals(Variable(RVArray()).value.value, [])

    udt_name = UDTName(
        {
            "a": Currency,
            "b": Variant(),
            "c": StringN(3),
            "d": UDTName({"i": Integer}, "nested"),
        },
        "parent",
    )
    assert_equals(
        to_dict(Variable(udt_name)),
        {
            "TYPE": "Variable",
            "declared_type": {
                "TYPE": "UDTName",
                "element_types": {
                    "a": "Currency",
                    "b": {"TYPE": "Variant"},
                    "c": {"TYPE": "StringN", "n": 3},
                    "d": {
                        "TYPE": "UDTName",
                        "element_types": {"i": "Integer"},
                        "name": "nested",
                    },
                },
                "name": "parent",
            },
            "value": {
                "TYPE": "UDT",
                "name": "parent",
                "value": {
                    "a": {"TYPE": "Currency", "value": 0},
                    "b": {"TYPE": "Empty"},
                    "c": {"TYPE": "String", "value": "\x00\x00\x00"},
                    "d": {
                        "TYPE": "UDT",
                        "name": "nested",
                        "value": {"i": {"TYPE": "Integer", "value": 0}},
                    },
                },
            },
        },
    )

    with assert_raises(TypeError):
        Variable(list)

    with assert_raises(TypeError):
        Variable(FArray(Variant, ((1, 2), (3, 4))))


def test_values():
    """
    data: values
    """
    with assert_raises(ValueError):
        Byte(-1)

    assert_equals(Empty().value, None)

    with assert_raises(TypeError):
        Variable(Value)

    assert_equals(Variable(Boolean).value, Boolean(False))
    assert_equals(Variable(Double).value, Double(0.0))
    assert_equals(Variable(StringN(5)).value, String(b"\x00" * 10))

    array = Array.from_bounds(((1, 4), (2, 3)), Variant())
    assert_equals(len(array.value), 4)
    assert_equals(len(array.value[0]), 2)

    array = Array([[1, 2, 3], [4, 5, 6], [7, 8, 9]], [(1, 3), (5, 7)])
    assert_equals(array[(1, 5)], 1)
    assert_equals(array[(1, 7)], 3)
    assert_equals(array[(3, 7)], 9)
    assert_equals(array[(3, 5)], 7)

    with assert_raises(IndexError):
        array[(2, 9)]
