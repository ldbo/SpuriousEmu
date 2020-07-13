# SpuriousEmu

Visual Basic for Applications tools allowing to parse VBA files, interpret them and extract behaviour information for malware analysis purpose.

## Requirements

Python 3.8 is used, and the main dependency is `pyparsing`, which allows to parse the VBA grammar. Additionally, `nose` is needed to run the tests. You can create a conda development environment using `environment.yml`:

```bash
conda env create -f environment.yml
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
