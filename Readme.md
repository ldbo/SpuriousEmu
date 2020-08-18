# SpuriousEmu

Visual Basic for Applications tools allowing to parse VBA files, interpret them and extract behaviour information for malware analysis purpose.

## Usage

SpuriousEmu can work with VBA source files, or directly with Office documents. For the later case, it relies on olevba to extract macros from the files. For each of the commands, use the `-i` flag to specify the input file to work with, whatever its format.

If you work with VBA source files, the following convention is used:
    - procedural modules have `.bas` extension
    - class modules have `.cls` extension
    - standalone script files have `.vbs` extension

SpuriousEmu uses different subcommands for its different operating modes.

### Static analysis

Static analysis is performed using the `static` subcommand.

Usually, the first step is to determine the different functions and classes defined, in order to understand the structure of the program. You can for example use it to determine the entry point prior to dynamic analysis. It is the default behaviour when using no other flag than `-i`:

```bash
./emu.py static -i document.xlsm
```

Additionally, for large files, you can use the `-o` flag to serialize the information compiled during static analysis into a binary file that you will be able to use later with the `-i` flag:

```bash
./emu.py static -i document.xlsm -o document.spurious-com
```

You can also de-obfuscate a file by using the `-d` flag, which specifies the de-obfuscation level. You can output the whole file, or a single function or module using the `-e` flag. The result can be sent to standard output or written to a file specified with the `-o` file:

```bash
./emu.py static -i document.xlsm -d3 -e VBAEnv.Default.Main -o Main.bas
```

### Dynamic analysis

You can trigger dynamic analysis with the `dynamic` subcommand.

Once you have found the entry-point you want to use with the `static` subcommand, you can execute a file by specifying it with the `-e` flag. For example, to launch the `Main` function found in `doc.xlsm`, use

```bash
./emu.py dynamic -i doc.xlsm -e Main
```

This will display a report of the execution of the program. Additionally, if you want to save the files created during execution, you can use the `-o` flag: it specifies a directory to save files to. Each created file is then stored in a file with its md5 sum as title, and a `{hash}.filename.txt` file contains its original name.

## Dependencies

Python 3.8 is used, and SpuriousEmu mainly relies on `pyparsing` for VBA grammar parsing, and `oletools` to extract VBA macros from Office documents.

`nose` is used as testing framework, and `mypy` to perform static code analysis. `lxml` and `coverage` are used to produce test reports.

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

Both commands produce HTML reports stored in 'tests/report'.
