from emu import syntax

from nose.tools import assert_equals

import json


def load_result(test):
    result_file = f"tests/{test}.json"
    with open(result_file) as f:
        return json.load(f)


def vbs_path(test):
    return f"tests/{test}.vbs"


def load_vbs(test):
    vbs = vbs_path(test)
    with open(vbs) as f:
        return f.read()


def assert_correct_function(test, function):
    result = function(test)
    expected_result = load_result(test)
    assert_equals(result, expected_result)


def parsing(test):
    ast = syntax.parse_file(vbs_path(test))
    return ast.to_dict()


def test_expressions():
    assert_correct_function("basic_01", parsing)