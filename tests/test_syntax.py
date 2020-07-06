from emu import syntax

from nose.tools import assert_equals

import json


def assert_expected_result(test_name, result):
    with open(test_name + ".json") as f:
        expected = json.load(f)

    assert_equals(expected, result)


def assert_correct_parsing(test):
    path = "tests/" + test
    ast = syntax.parse_file(path + '.vbs')
    assert_expected_result(path, ast.to_dict())


def test_expressions():
    assert_correct_parsing("basic_01")
