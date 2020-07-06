from emu import syntax

import json


def assert_expected_result(test_name, result):
    with open(test_name + ".json") as f:
        expected = json.load(f)

    assert(expected == result)


def assert_correct_parsing(test_name):
    ast = syntax.parse_file(test + '.vbs')
    assert_expected_result(test, ast.to_dict())


def test_expressions():
    assert_correct_parsing("basic_01")
