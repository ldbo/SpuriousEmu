from dataclasses import dataclass
from typing import List, Dict

from nose.tools import assert_equals, assert_raises

from emu.utils import FilePosition, Visitable, Visitor, to_dict


def test_position():
    """utils: file position"""
    content = "line1\nline2"

    with assert_raises(RuntimeError):
        FilePosition("f", content, 10, 1, 1, 0, 1, 1)

    p1 = FilePosition("f", content, 0, 4, 1, 1, 1, 5)
    p2 = FilePosition("f2", content, 2, 4, 1, 1, 3, 5)
    p3 = FilePosition("f", content, 7, 9, 2, 2, 2, 4)
    p4 = FilePosition("f", content, 0, 11, 1, 2, 1, 6)

    assert_equals(p1.lines(), "line1")
    assert_equals(p2.lines(), "line1")
    assert_equals(p3.lines(), "line2")
    assert_equals(p4.lines(), content)

    assert_equals(p2.body(), "ne")

    with assert_raises(RuntimeError):
        p1 + p2

    assert_equals((p1 + p3).start_index, 0)
    assert_equals((p1 + p3).end_index, 9)
    assert_equals((p4 + p1), p4)

@dataclass
class U0(Visitable):
    name: str


class U1(U0):
    pass


class U2(U0):
    pass


class V(Visitor[str]):
    def visit_U0(self, u: U0) -> str:
        return f"U0: {u}"

    def visit_U1(self, u: U1) -> str:
        return f"U1: {u}"


def test_visitor():
    """utils: visitor"""
    u0 = U0("0")
    u1 = U1("1")
    u2 = U2("2")
    v = V()

    assert_equals(v.visit(u0), "U0: U0(name='0')")
    assert_equals(v.visit(u1), "U1: U1(name='1')")

    assert_equals(v.visit(u0), u0.accept(v))
    assert_equals(v.visit(u0), v.visit_U0(u0))

    assert_equals(v.visit(u1), u1.accept(v))
    assert_equals(v.visit(u1), v.visit_U1(u1))

    with assert_raises(NotImplementedError):
        v.visit(u2)

    with assert_raises(NotImplementedError):
        u2.accept(v)


@dataclass
class A:
    EXCLUDED = {"shy"}
    shy: str
    a: int
    b: List["B"]


@dataclass
class B:
    a: Dict[str, int]
    b: str


def test_to_dict():
    """utils: to_dict"""
    b1 = B({"hey": 1, "ho": 2}, "let's go")
    b2 = B(dict(), "All empty")
    a = A("shy", 42, [b1, b2])

    assert_equals(
        to_dict(a),
        {
            "TYPE": "A",
            "a": 42,
            "b": [
                {"TYPE": "B", "a": {"hey": 1, "ho": 2}, "b": "let's go"},
                {"TYPE": "B", "a": {}, "b": "All empty"},
            ],
        },
    )

    assert_equals(
        to_dict(a, excluded_fields={"a"}),
        {
            "TYPE": "A",
            "b": [
                {"TYPE": "B", "b": "let's go"},
                {"TYPE": "B", "b": "All empty"},
            ],
        },
    )
