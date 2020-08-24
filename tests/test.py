import json
import subprocess

from os import listdir
from pathlib import Path
from pprint import pprint
from typing import Any, Callable, Dict, Optional

from nose.tools import assert_equals

Result = Dict[str, Any]
SourceFile = str


SAMPLES = sorted(
    list(filter(lambda filename: "." not in filename, listdir("tests/samples")))
)

assert_equals.__self__.maxDiff = None


def result_path(test: str) -> str:
    """Path of the JSON file containing the test result"""
    return f"tests/result/{test}.json"


def load_result(test: str) -> Result:
    """Return the dict containing the test result"""
    with open(result_path(test)) as f:
        return json.load(f)


def export_result(test: str, function: Callable[[SourceFile], Result]) -> None:
    """Export the result of the test to JSON"""
    with open(result_path(test), "w") as f:
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
        assert Path(path).exists()
        return path


def output_path(test: str) -> str:
    """
    Build the path of a saved command output: tests/result/{text}.output.txt
    """
    return f"tests/result/{test}.output.txt"


def command_output(command: str, sample_number: Optional[int] = None) -> str:
    """
    Return the output of the command, with a optional sample as additional
    argument.
    """
    command_list = command.split()

    if sample_number is not None:
        command_list.append("tests/samples/" + SAMPLES[sample_number])

    proc = subprocess.run(command_list, capture_output=True, encoding="utf-8")
    output = f"""
Return code: {proc.returncode}
Output:
{proc.stdout}

Error:
{proc.stderr}
"""

    return output


def export_output(
    test: str, command: str, sample_number: Optional[int] = None
) -> None:
    """Export the output of command to the corresponding output_path."""
    output = command_output(command, sample_number)

    with open(output_path(test), "w") as f:
        f.write(output)


def assert_correct_output(
    test: str, command: str, sample_number: Optional[int] = None
) -> None:
    """Assert calling command output the expected result"""
    with open(output_path(test), "r") as f:
        expected_output = f.read()

    output = command_output(command, sample_number)

    assert_equals(expected_output, output)


def assert_correct_function(
    test: str, function: Callable[[SourceFile], Result]
) -> None:
    """Load the expected result from JSON and compare to the function result"""
    result = function(source_path(test))
    expected_result = load_result(test)

    assert_equals(expected_result, result)


def run_function(test: str, function: Callable[[SourceFile], Result]) -> None:
    """
    Display the result of a function call. Used during development until an
    expected result has been produced
    """
    result = function(source_path(test))
    pprint(result)
