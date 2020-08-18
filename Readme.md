# SpuriousEmu

Visual Basic for Applications tools allowing to parse VBA files, interpret them and extract behaviour information for malware analysis purpose.

## Requirements

Python 3.8 is used, and the main dependency is `pyparsing`, which allows to parse the VBA grammar. Additionally, `nose` is needed to run the tests. You can create a conda development environment using `environment.yml`:

```bash
conda env create -f environment.yml
```

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

Once you have found the entry-point you want to use using the `static` subcommand, you can execute a file by specifying it with the `-e` flag. For example, to launch the `Main` function found in `doc.xlsm`, use

```bash
./emu.py dynamic -i doc.xlsm -e Main
```

## Tests

All test files are in `tests`, including:
    - Python test scripts, starting with `test_`
    - VBA scripts used to test the different stages of the tools, with `vbs` extensions
    - expected test results, stored as JSON dumps

The testing framework used is nose, which you can run against the test suite with:

```bash
nosetests tests/test_*.py
```
