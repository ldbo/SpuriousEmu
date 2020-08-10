import json
from pathlib import Path
from pprint import pprint
from typing import Callable, Dict, Any

from nose.tools import assert_equals

Result = Dict[str, Any]
SourceFile = str


def result_path(test: str) -> str:
    """Path of the JSON file containing the test result"""
    return f"tests/result/{test}.json"


def load_result(test: str) -> Result:
    """Return the dict containing the test result"""
    with open(result_path(test)) as f:
        return json.load(f)


def export_result(test: str, function: Callable[[SourceFile], Result]) -> None:
    """Export the result of the test to JSON"""
    with open(result_path(test), 'w') as f:
        json.dump(function(source_path(test)), f, indent=4)


def source_path(test: str) -> str:
    """
    Return the path of the VBS source file, or the project directory, of the
    test.
    """
    root_path = f"tests/source/{test}"
    if Path(root_path).is_dir():
        return root_path
    else:
        path = root_path + ".vbs"
        assert(Path(path).exists())
        return path


def assert_correct_function(test: str,
                            function: Callable[[SourceFile], Result]) -> None:
    """Load the expected result from JSON and compare to the function result"""
    result = function(source_path(test))
    expected_result = load_result(test)
    assert_equals(result, expected_result)


def run_function(test: str, function: Callable[[SourceFile], Result]) -> None:
    """
    Display the result of a function call. Used during development until an
    expected result has been produced
    """
    result = function(source_path(test))
    pprint(result)
