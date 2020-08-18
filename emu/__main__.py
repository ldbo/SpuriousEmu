#! /usr/bin/env python

from argparse import ArgumentParser
from hashlib import md5
from pathlib import Path
from pickle import dumps, loads
from pprint import pprint
from sys import exit
from typing import Dict

from magic import from_buffer as magic_from_buffer

from oletools.olevba import VBA_Parser

from prettytable import PrettyTable

from emu import Program, Compiler, Interpreter, Unit


def build_argparser():
    parser = ArgumentParser()
    subparsers = parser.add_subparsers(
        dest="mode",
        title="Mode",
        description="You must choose a mode for SpuriousEmu to operate in.",
        required=True)

    # Static analysis
    static_parser = subparsers.add_parser(
        'static',
        help="Static analysis, allows symbols preview, compilation and "
             "deobfuscation")
    static_parser.add_argument(
        "-i", "--input",
        required=True,
        help="Input file, either a VBA source file or project, an Office "
             "document or a SpuriousEmu compiled file")
    static_parser.add_argument(
        '-o', '--output',
        help="Output file, used to save the compilation result or the "
             "deobfuscated macros in a file")
    static_parser.add_argument(
        "-d", "--deobfuscate",
        metavar="LEVEL",
        type=int,
        help="Enable deobfuscation and specify the deobfuscation level. See "
             "-e to only process a specific symbol.")
    static_parser.add_argument(
        "-e", "--entry",
        help="Absolute path of the class or function to deobfuscate.")
    static_parser.set_defaults(func=static_analysis)

    # Dynamic analysis
    dynamic_parser = subparsers.add_parser(
        'dynamic',
        help="Dynamic analysis used to perform different kind of sandboxing.")
    dynamic_parser.add_argument(
        "-i", "--input",
        help="Input file")
    dynamic_parser.add_argument(
        "-e", "--entry",
        default='Main',
        help="Entry point of the dynamic analysis. Must be specified if no "
             "Main function is defined. Must use the absolute path of the "
             "symbol.")
    dynamic_parser.add_argument(
        "-o", "--output",
        help="Directory to save the created files to.")
    dynamic_parser.set_defaults(func=dynamic_analysis)

    return parser


def compile_input_file(input_file: str) -> Program:
    path = Path(input_file)

    if path.is_dir():
        # A directory must be a project containing VBA source files
        return Compiler.compile_project(input_file)

    with open(input_file, 'rb') as f:
        content = f.read()

    if content.startswith(b'SpuriousEmuProgram'):
        program = content[len(b'SpuriousEmuProgram'):]
        return loads(program)

    units = []
    if magic_from_buffer(content, mime=True) == "text/plain":
        if not input_file.endswith('.vbs'):
            input_file += '.vbs'
        units.append(Unit.from_content(content.decode('utf-8'), input_file))
    else:
        vba_parser = VBA_Parser(input_file, data=content)

        for _, _, vba_filename, vba_code in vba_parser.extract_all_macros():
            units.append(Unit.from_content(vba_code, vba_filename))

    return Compiler.compile_units(units)


def save_compiled_program(program: Program, path: str) -> None:
    serialized_program = dumps(program)

    with open(path, 'wb') as f:
        f.write(b'SpuriousEmuProgram')
        f.write(serialized_program)


def display_functions(program: Program) -> None:
    functions = PrettyTable()
    functions.add_column('Functions', program.to_dict()['memory']['functions'],
                         align='l')

    classes = PrettyTable()
    classes.add_column('Classes', program.to_dict()['memory']['classes'],
                       align='l')

    print(f"{functions}\n\n{classes}")


def static_analysis(arguments):
    # Check arguments validity
    if not Path(arguments.input).exists():
        print(f"Error: input file {arguments.input} does not exist.")
        return 1

    if arguments.output is not None and Path(arguments.output).exists():
        print(f"Error: output file {arguments.output} already exists.")
        return 1

    if arguments.deobfuscate is not None:
        print(f"Error: de-obfuscation is not handled yet.")
        return 1

    if (arguments.entry is not None) and arguments.deobfuscate is None:
        print(f"Error: the -e flag requires the -d flag.")
        return 1

    # Compile program
    program = compile_input_file(arguments.input)
    display_functions(program)

    if arguments.output is not None:
        save_compiled_program(program, arguments.output)

    return 0


def execute_program(program: Program, entry_point: str) -> None:
    libraries = list(map(str, Path('./lib/').glob('*')))
    linked_program = Compiler.link_libraries(program, libraries)
    return Interpreter.run_program(linked_program, entry_point)


def save_files(directory: str, files: Dict[str, str]) -> None:
    directory_path = Path(directory)
    directory_path.mkdir()

    for name, content in files.items():
        hasher = md5()
        hasher.update(content.encode('utf-8'))
        digest = hasher.hexdigest()
        content_path = directory_path.joinpath(digest)

        with open(content_path.absolute(), 'w') as f:
            f.write(content)

        name_path = directory_path.joinpath(f"{digest}.filename.txt")

        with open(name_path.absolute(), "w") as f:
            f.write(name)


def dynamic_analysis(arguments):
    # Check arguments validity
    if not Path(arguments.input).exists():
        print(f"Error: input file {arguments.input} does not exist.")
        return 1

    if arguments.output is not None and Path(arguments.output).exists():
        print(f"Error: {arguments.output} exists.")
        return 1

    # Load and execute
    program = compile_input_file(arguments.input)
    report = execute_program(program, arguments.entry)
    clean_report = {key: value for key, value in report.to_dict().items()
                    if key != 'file'}
    clean_report.update({'file': list(report.files.keys())})

    pprint(clean_report)

    if arguments.output is not None:
        save_files(arguments.output, report.files)

    return 0


def main():
    parser = build_argparser()
    args = parser.parse_args()

    return args.func(args)


if __name__ == "__main__":
    exit(main())
