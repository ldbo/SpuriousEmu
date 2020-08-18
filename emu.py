#! /usr/bin/env python

from argparse import ArgumentParser
from sys import exit


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
        help="Entry point of the dynamic analysis. Must be specified if no "
             "Main function is defined. Must use the absolute path of the "
             "symbol.")
    dynamic_parser.set_defaults(func=dynamic_analysis)

    return parser


def static_analysis(arguments):
    return 0


def dynamic_analysis(arguments):
    return 0


def main():
    parser = build_argparser()
    args = parser.parse_args()

    return args.func(args)


if __name__ == "__main__":
    exit(main())
