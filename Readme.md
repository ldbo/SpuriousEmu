# SpuriousEmu

![Travis (.com)](https://img.shields.io/travis/com/ldbo/SpuriousEmu)
![GitHub tag (latest SemVer)](https://img.shields.io/github/v/tag/ldbo/SpuriousEmu)
![PyPI - Downloads](https://img.shields.io/pypi/v/spurious-emu)
![Coveralls github](https://img.shields.io/coveralls/github/ldbo/SpuriousEmu)
![PyPI - Python Version](https://img.shields.io/pypi/pyversions/spurious-emu)
![Github - License](https://img.shields.io/github/license/ldbo/SpuriousEmu)

Visual Basic for Applications tools allowing to parse VBA files, interpret them and extract behaviour information for malware analysis purpose.

## Installation

SpuriousEmu is available on PyPI, so you can install it using

```bash
pip install spurious-emu
```

## Usage

SpuriousEmu can work with VBA source files, or directly with Office documents. For the later case, it relies on olevba to extract macros from the files. All of the command use a final positional argument to specify the input file to work with.

If you work with VBA source files, the following convention is used:
    - procedural modules have `.bas` extension
    - class modules have `.cls` extension
    - standalone script files have `.vbs` extension

SpuriousEmu uses different subcommands for its different operating modes.

### Static analysis

Static analysis is performed using the `static` subcommand.

Usually, the first step is to determine the different functions and classes defined, in order to understand the structure of the program. You can for example use it to determine the entry point prior to dynamic analysis. It is the default behaviour when using no flag:

```bash
emu static document.xlsm
```

Additionally, for large files, you can use the `-o` flag to serialize the information compiled during static analysis into a binary file that you will be able to use later with the `report` command for example:

```bash
emu static -o document.spurious-com document.xlsm
```

You can also de-obfuscate a file by using the `-d` flag, which specifies the de-obfuscation level. You can output the whole file, or a single function or module using the `-e` flag. The result can be sent to standard output or written to a file specified with the `-o` file:

```bash
emu static -d3 -e VBAEnv.Default.Main -o Main.bas document.xlsm
```

### Dynamic analysis

You can trigger dynamic analysis with the `dynamic` subcommand.

Once you have found the entry-point you want to use with the `static` subcommand, you can execute a file by specifying it with the `-e` flag. For example, to launch the `Main` function found in `doc.xlsm`, use

```bash
emu dynamic -e Main doc.xlsm
```

This will display a report of the execution of the program. Additionally, if you want to save the files created during execution, you can use the `-o` flag: it specifies a directory to save files to. Each created file is then stored in a file with its md5 sum as file name, and a `{hash}.filename.txt` file contains its original name. You can also save a report of the dynamic analysis using the `-r` flag. For example:

```bash
emu dynamic -o extract_files -r report.spemu-out doc.xlsm
```

### Report production

You can work with `.spemu-out` and `.spemu-com` file with the `report` command.

The `report` commands can have three mutually exclusive flags: `--json`, `--csv` and `--table`, which change the way reports are displayed.

Similarly to the default `static` output, you can use the `--symbols` flag with a `.spemu-com` file to get the list of functions and classes. For example, to have them in a JSON dump, you can use

```bash
emu report --symbols --json program.spemu-com
```

You can extract the files generated by the execution of a program using the `--extract-files` flag, which behaves like the `-o` flag with the `dynamic` command:

```bash
emu report --extract-files files program.spemu-out
```

A timeline of the events can be produced with the `--timeline` flag. It can be made easier to read with the `--shorten` and `--skip-streaks` commands, as in

```bash
emu report --timeline --table --shorten --skip-streaks 10 program.spemu-out
```

## Dependencies

Python 3.8 is used, and SpuriousEmu mainly relies on [PyParsing](https://github.com/pyparsing/pyparsing) for VBA grammar parsing, and [oletools](http://www.decalage.info/python/oletools) to extract VBA macros from Office documents. Report tables are generated using [PrettyTable](https://github.com/jazzband/prettytable).

[nose](https://nose.readthedocs.io/en/latest/man.html) is used as testing framework, and [mypy](http://mypy-lang.org/) to perform static code analysis. `lxml` and `coverage` are used to produce test reports.

## Tests

To set a development environment up, use `poetry`:

```bash
poetry install
```

Then, use nose to run the test suite:

```bash
poetry run nosetests
```

All test files are in `tests`, including:
    - Python test scripts, starting with `test_`
    - VBA scripts used to test the different stages of the tools, with `vbs` extensions, stored in `source`
    - expected test results, stored as JSON dumps in `result`


You can use mypy to perform code static analysis:

```bash
poetry run mypy emu/*.py
```

Both commands produce HTML reports stored in `tests/report`.
