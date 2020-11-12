"""
Testing framework helper functions.

Source files, ie. files that are being processed during a test, are stored
in tests/sources. They are generally VBA source codes.

Results are stored in JSON in tests/results.

The ``test`` parameter of the helper functions refers to the name of the current
test file.
"""

from nose.tools import assert_equals
from pathlib import Path
from typing import Any, Dict, List, Union

import json

# Tweak allowing to have arbitrarily long diffs
assert_equals.__self__.maxDiff = None

Result = Union[Dict[Any, Any], List[Any]]  #: Generic type of a JSONable result


def result_path(test: str) -> str:
    """
    Returns:
      The path of the JSON file containing the test result"""
    return f"tests/results/{test}.json"


def load_result(test: str) -> Result:
    """
    Returns:
      The ``dict`` or ``list`` containing the test result
    """
    with open(result_path(test)) as f:
        return json.load(f)


def export_result(test: str, result: Result) -> None:
    """
    Export the result of the test to a file, using the JSON format.
    """
    with open(result_path(test), "w") as f:
        json.dump(result, f, indent=2, sort_keys=True)


def assert_result(test: str, result: Result) -> None:
    """
    Raises:
      AssertionError: If the stored result of ``test`` is different than
                      ``result``
    """
    assert_equals(load_result(test), json.loads(json.dumps(result)))


def source_path(test: str) -> str:
    """
    Returns:
      The path of the source of the test, found searching for test* in the
      source directory.
    """
    glob = list(Path("tests/sources/").glob(test + "*"))
    assert len(glob) == 1
    return str(glob[0].absolute())


def load_source(test: str) -> str:
    """
    Returns:
      The file content of the corresponding test source file
    """
    with open(source_path(test)) as f:
        return f.read()
