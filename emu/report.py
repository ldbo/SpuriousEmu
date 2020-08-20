"""Tools to generate reports based on analysis results."""

import json

from dataclasses import dataclass, field
from enum import Enum
from hashlib import md5
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from prettytable import PrettyTable

from .side_effect import OutsideWorld
from .compiler import Program


def _needs_outside_world(method):
    error_msg = "ReportGenerator needs an OutsideWorld object to call " \
        f"{method.__name__}"

    def decorated(*args, **kwargs):
        if args[0].outside_world is None:
            raise RuntimeError(error_msg)

        return method(*args, **kwargs)

    return decorated


def _needs_program(method):
    error_msg = "ReportGenerator needs a Program object to call " \
        f"{method.__name__}"

    def decorated(*args, **kwargs):
        if args[0].program is None:
            raise RuntimeError(error_msg)

        return method(*args, **kwargs)

    return decorated


@dataclass
class ReportGenerator:
    """
    Class used to extract information from an OutsideWorld object. You can
    get different kind of reports (timeline of the events, hierarchical output
    of all the events, ...) using different formats, specified in the Format
    enumeration. You can also use it to extract generated files and save them
    to you file system.
    """
    class Format(Enum):
        """Display formats supported by the ReportGenerator."""
        JSON = 'json'
        CSV = 'csv'
        TABLE = 'table'

    EventLine = Tuple[int, float, str, str, Any]

    program: Optional[Program] = None
    outside_world: Optional[OutsideWorld] = None
    output_format: "ReportGenerator.Format" = field(default=Format.JSON)
    indent: int = 4
    reproducible: bool = False
    skip_identical: bool = False
    hash_algorithm = md5

    # Utility methods

    def hash_file(self, content: str) -> str:
        """
        Return the hex digest of the file content encoded in UTF8, using
        the self.hash algorithm.
        """
        hasher = self.hash_algorithm()
        hasher.update(content.encode('utf-8'))

        return hasher.hexdigest()

    def to_json(self, report: Any) -> str:
        """Return the JSON dump of the report."""
        return json.dumps(report, indent=self.indent, sort_keys=True)

    # Program reports

    @_needs_program
    def extract_symbols(self) -> Dict[str, List[str]]:
        """Returns a dictionnary with 'functions' and 'classes' keys."""
        memory_dict = self.program.to_dict()['memory']
        functions = memory_dict['functions']
        classes = memory_dict['classes']

        return {'functions': functions, 'classes': classes}

    @_needs_program
    def produce_symbols(self) -> str:
        """Returns the formatted symbols : classes and functions."""
        symbols = self.extract_symbols()

        if self.output_format is ReportGenerator.Format.JSON:
            return self.to_json(symbols)
        elif self.output_format is ReportGenerator.Format.TABLE:
            classes = PrettyTable()
            classes.add_column("Classes", symbols['classes'], align="l")

            functions = PrettyTable()
            functions.add_column("Functions", symbols['functions'], align='l')

            return f"{classes}\n\n{functions}"

    # OutsideWorld reports

    @_needs_outside_world
    def extract_timeline(self) -> List["ReportGenerator.EventLine"]:
        """
        Return the chronological list of events, as tuples with elements
        identifier, time, category, context and data.
        """
        def event_to_tuple(event: OutsideWorld.Event) \
                -> "ReportGenerator.EventLine":
            d = event.to_dict(self.reproducible)
            m = map(lambda attr: d.get(attr, 0.0),
                    ('identifier', 'time', 'category', 'context', 'data'))
            return tuple(m)

        timeline = list(map(event_to_tuple, self.outside_world.events))
        timeline.sort()

        return timeline

    @_needs_outside_world
    def produce_timeline(self) -> str:
        """Return the formatted timeline."""
        timeline = self.extract_timeline()

        if self.output_format is ReportGenerator.Format.JSON:
            return json.dumps(timeline, indent=self.indent, sort_keys=True)
        elif self.output_format is ReportGenerator.Format.TABLE:
            fields = ('ID', 'Time', 'Category', 'Context', 'Data')
            table = PrettyTable(fields)
            table.align = 'l'
            for event_tuple in timeline:
                table.add_row(event_tuple)

            return table.get_string()

    @_needs_outside_world
    def extract_files(self, save_directory: str) -> None:
        """
        Extract the content added to the OutsideWorld files, and write it to
        the actual filesystem. It uses the following convention: for each
        file that exists in OutsideWorld, with md5 sum `hash`, write to
        `hash` the content of the file, and to `hash.filename.txt` its name. In
        case of filename conflict, overrides the file.

        :arg save_directory: Output directory, created if it does not exist.
        """
        path = Path(save_directory)
        path.mkdir(exist_ok=True)

        for name, content in self.outside_world.files.items():
            content_hash = self.hash_file(content)
            content_path = path.joinpath(content_hash)
            filename_path = path.joinpath(f'{content_hash}.filename.txt')

            with open(content_path.absolute(), 'w') as f:
                f.write(content)

            with open(filename_path.absolute(), 'w') as f:
                f.write(name)
