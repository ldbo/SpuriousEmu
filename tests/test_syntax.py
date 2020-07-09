import json
from pprint import pprint
from typing import Callable, Dict, Any

from nose.tools import assert_equals

from emu import syntax

Result = Dict[str, Any]
SourceFile = str


def result_path(test: str) -> str:
    """Path of the JSON file containing the test result"""
    return f"tests/{test}.json"


def load_result(test: str) -> Result:
    """Return the dict containing the test result"""
    with open(result_path(test)) as f:
        return json.load(f)


def export_result(test: str, function: Callable[[SourceFile], Result]) -> None:
    """Export the result of the test to JSON"""
    with open(result_path(test), 'w') as f:
        json.dump(function(vbs_path(test)), f)


def vbs_path(test: str) -> str:
    """Return the path of the VBS source file of the test"""
    return f"tests/{test}.vbs"


def assert_correct_function(test: str, function: Callable[[SourceFile], Result]) -> None:
    """Load the expected result from JSON and compare to the function result"""
    result = function(vbs_path(test))
    expected_result = load_result(test)
    assert_equals(result, expected_result)


def run_function(test : str, function: Callable[[SourceFile], Result]) -> None:
    """Display the result of a function call. Used during development until an expected result has been produced"""
    result = function(vbs_path(test))
    pprint(result)


def parsing(vbs: SourceFile) -> Result:
    ast = syntax.parse_file(vbs)
    return ast.to_dict()


def test_expressions():
    assert_correct_function("basic_01", parsing)


def test_inline_declarations():
    assert_correct_function("basic_02", parsing)