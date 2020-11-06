from dataclasses import dataclass

from nose.tools import assert_equals, assert_raises

from emu.utils import Visitable, Visitor


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
    """
    utils: visitor
    """
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
