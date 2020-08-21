from emu import __version__
from tests.test import assert_correct_output


def test_version():
    assert __version__ == "0.3.1"
    assert_correct_output("emu_version", "emu -v")


def test_basic_static():
    assert_correct_output("emu_basic_static", "emu static", 0)


def test_basic_dynamic():
    assert_correct_output(
        "emu_basic_dynamic", "emu dynamic -e Document_Close", 0
    )


def test_load_program_dynamic():
    assert_correct_output(
        "emu_load_program_dynamic",
        "emu dynamic -e Document_Close "
        "tests/samples/gamaredon_first_stage.spemu-com",
    )
