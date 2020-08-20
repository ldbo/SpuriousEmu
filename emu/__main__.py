#! /usr/bin/env python

from argparse import ArgumentParser
from hashlib import md5
from pathlib import Path
from pprint import pprint
from sys import exit
from typing import Dict

from magic import from_buffer as magic_from_buffer

from oletools.olevba import VBA_Parser

from prettytable import PrettyTable

from emu import (Program, Compiler, Interpreter, Unit, __version__, Serializer,
                 SerializationError)


def build_argparser():
    parser = ArgumentParser(
        description="""VBA static and dynamic analysis tool for malware
                       analysts."""
    )
    parser.add_argument(
        "-v", "--version",
        action='version',
        version=f"SpuriousEmu v{__version__}")
    subparsers = parser.add_subparsers(
        dest="mode",
        title="Operating mode",
        required=True)

    # Static analysis
    static_parser = subparsers.add_parser(
        'static',
        help="Static analysis, allows symbols preview, compilation and "
             "deobfuscation.")
    static_parser.add_argument(
        '-o', '--output',
        help="Output file, used to save the compilation result or the "
             "deobfuscated macros")
    static_parser.add_argument(
        "-d", "--deobfuscate",
        metavar="LEVEL",
        type=int,
        help="Enable deobfuscation and specify the deobfuscation level. See "
             "-e to only process a specific symbol")
    static_parser.add_argument(
        "-e", "--entry",
        help="Absolute path of the class or function to deobfuscate")
    static_parser.add_argument(
        "input",
        help="""Input file, can be an Office document, some VBA source code, or
                a SpuriousEmu compiled program""")
    static_parser.set_defaults(func=static_analysis)

    # Dynamic analysis
    dynamic_parser = subparsers.add_parser(
        'dynamic',
        help="Dynamic analysis used to perform different kind of emulation.")
    dynamic_parser.add_argument(
        "-e", "--entry",
        default='Main',
        help="""Entry point of the dynamic analysis. Must be specified if no
                Main function is defined. Must use the absolute path of the
                symbol""")
    dynamic_parser.add_argument(
        "-r", "--results",
        help="Save the analysis results to REPORT")
    dynamic_parser.add_argument(
        "-o", "--output",
        help="Directory to save the created files to")
    dynamic_parser.add_argument(
        "input",
        help="""Input file, can be an Office document, some VBA source code, or
                a SpuriousEmu compiled program""")
    dynamic_parser.set_defaults(func=dynamic_analysis)

    # Report generation
    report_parser = subparsers.add_parser(
        'report',
        help="""Extract information from analysis results generated by other
                commands."""
    )
    report_action = report_parser.add_mutually_exclusive_group(required=True)
    report_action.add_argument(
        "-x", "--extract-files",
        metavar="DIR",
        dest="destination_directory",
        help="""Extract files from a dynamic analysis report to DIR. Defaults
                to --table""")
    report_action.add_argument(
        "-t", "--timeline",
        action="store_true",
        help="Display the event timeline. Defaults to --table")
    report_action.add_argument(
        "-c", "--category",
        action="store_true",
        help="Display all the events, grouped by category. Defaults to --json")
    report_format = report_parser.add_mutually_exclusive_group()
    report_format.add_argument(
        "--json",
        action="store_true",
        help="Use JSON output")
    report_format.add_argument(
        "--csv",
        action="store_true",
        help="Use CSV output")
    report_format.add_argument(
        "--table",
        action="store_true",
        help="Use human-friendly output")
    report_parser.add_argument(
        "input",
        help="Dynamic result file to use")
    report_format.set_defaults(func=generate_report)

    return parser


def compile_input_file(input_file: str) -> Program:
    path = Path(input_file)

    if path.is_dir():
        # A directory must be a project containing VBA source files
        return Compiler.compile_project(input_file)

    with open(input_file, 'rb') as f:
        content = f.read()

    # Try to deserialize an already compiled program
    try:
        serializer = Serializer()
        obj = serializer.deserialize(content)

        if type(obj) == Program:
            return obj
    except SerializationError:
        pass

    # Compile an Office document or VBA source files
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
    serializer = Serializer()
    serializer.save(program, path)


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
    clean_report = report.to_dict()
    clean_report['file'] = [event for event in clean_report['file']
                            if event['type'] != 'Write']

    pprint(clean_report)

    if arguments.output is not None:
        save_files(arguments.output, report.files)

    return 0


def generate_report(arguments):
    print("report command is not handled yet.")
    return 1


def main():
    parser = build_argparser()
    args = parser.parse_args()

    return args.func(args)


if __name__ == "__main__":
    exit(main())
