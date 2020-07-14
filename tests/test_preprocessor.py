from emu import preprocessor
from tests.test import (assert_correct_function, SourceFile, Result)


def preprocess(vbs: SourceFile) -> Result:
    pp = preprocessor.Preprocessor()
    instructions = pp.extract_instructions_from_file(vbs)
    instructions_dicts = [i.to_dict() for i in instructions]
    return {'instructions': instructions_dicts}


def test_comments():
    assert_correct_function("comments", preprocess)
