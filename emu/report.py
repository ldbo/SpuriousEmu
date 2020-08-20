"""Tools to generate reports based on analysis results."""

import csv
import json

from dataclasses import dataclass, field
from enum import Enum
from io import StringIO
from hashlib import md5
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from prettytable import PrettyTable

from .compiler import Program
from .side_effect import OutsideWorld
from .serialize import Serializer


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

    You can customize the generated reports using the object attributes:
        - output_format : the output format
        - indent : Number of spaces for JSON indentation
        - reproducible: If True, don't display event time
        - skip_similar: If > 0, skip series of similar events in TABLE output
        - hash_algorithm: hashing algorithm used to name files
        - shorten: If True, shorten the context and data fields in TABLE output
    """
    class Format(Enum):
        """Display formats supported by the ReportGenerator."""
        JSON = 'json'
        CSV = 'csv'
        TABLE = 'table'

    EventLine = Tuple[int, str, str, str, Any]

    program: Optional[Program] = None
    outside_world: Optional[OutsideWorld] = None
    output_format: "ReportGenerator.Format" = field(default=Format.JSON)
    indent: int = 4
    reproducible: bool = False
    skip_similar: int = 0
    hash_algorithm = md5
    shorten = False

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

    def event_to_tuple(self, event: OutsideWorld.Event) \
            -> "ReportGenerator.EventLine":
        """
        Convert an event to a tuple.
        """
        identifier = event.identifier
        time = f"{event.time:.3f}"
        category = event.category.value
        context = event.context
        data = event.data

        return (identifier, time, category, context, data)

    def shorten_tuple(self, event: "ReportGenerator.EventLine") \
            -> "ReportGenerator.EventLine":
        """
        Shorten a tuple if self.shorten is True, ie. only keep context end
        symbol and discard data['data'] if possible.
        """
        if not self.shorten:
            return event

        identifier = event[0]
        time = event[1]
        category = event[2]
        context = event[3].split('.')[-1]

        if isinstance(event[4], dict):
            data = {key: value for key, value in event[4].items()
                    if key != "data"}
        else:
            data = event[4]

        return (identifier, time, category, context, data)

    def similar_events(self, event1: "ReportGenerator.EventLine",
                       event2: "ReportGenerator.EventLine") -> bool:
        """
        Compare two events, looking for similarities. Two events are similar
        if:
            - they have the same category and
            - they have the same context
            - they have dict data field and
                - they have data['data'] and data['type'] fields
                - their data['type'] fields are the same
        """
        category1 = event1[2]
        category2 = event2[2]
        if category1 != category2:
            return False

        context1 = event1[3]
        context2 = event2[3]
        if context1 != context2:
            return False

        data1 = event1[4]
        data2 = event2[4]

        if not isinstance(data1, dict):
            return True

        if 'data' not in data1 or 'data' not in data2:
            return False

        if 'type' not in data1 or 'type' not in data2:
            return False

        return data1['type'] == data2['type']

    def events_to_table(self, events: List[OutsideWorld.Event]) -> PrettyTable:
        """
        Build a table containing a list of events. The table has the columns
        ('ID', 'Time (s)', 'Category', 'Context', 'Data') and an event by row.

        Depending on shorten and skip_similar, series of similar events can be
        skiped.
        """
        fields = ('ID', 'Time (s)', 'Category', 'Context', 'Data')
        table = PrettyTable(fields)
        table.align = 'l'

        similar_events_streak = 0
        previous_event_tuple = None

        for event_tuple in events:
            if previous_event_tuple is not None:
                if self.similar_events(previous_event_tuple, event_tuple):
                    similar_events_streak += 1
                else:
                    if similar_events_streak > self.skip_similar + 1 \
                       and self.shorten:
                        table.add_row(('...', ) * 5)

                    if similar_events_streak > 1:
                        table.add_row(
                            self.shorten_tuple(previous_event_tuple))
                    similar_events_streak = 0

            if similar_events_streak <= self.skip_similar \
               or not self.shorten:
                table.add_row(self.shorten_tuple(event_tuple))

            previous_event_tuple = event_tuple

        if similar_events_streak > self.skip_similar + 1 \
           and self.shorten:
            table.add_row(('...', ) * 5)

        if similar_events_streak > 1:
            table.add_row(
                self.shorten_tuple(previous_event_tuple))

        return table

    def events_to_csv(self, events: List["ReportGenerator.EventLine"]) -> str:
        """Return a CSV formatted representation of a list of events."""
        stream = StringIO()
        writer = csv.writer(stream,
                            delimiter=";",
                            quoting=csv.QUOTE_NONNUMERIC)

        writer.writerow(('ID', 'Time (s)', 'Category', 'Context', 'Data'))

        for event in events:
            writer.writerow(event)

        return stream.getvalue()

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
        if self.output_format is ReportGenerator.Format.CSV:
            return "CSV not supported"
        elif self.output_format is ReportGenerator.Format.TABLE:
            classes = PrettyTable()
            classes.add_column("Classes", symbols['classes'], align="l")

            functions = PrettyTable()
            functions.add_column("Functions", symbols['functions'], align='l')

            return f"{classes}\n\n{functions}"

    @_needs_program
    def save_program(self, save_path: str) -> None:
        """Save the program to save_path."""
        Serializer.save(self.program, save_path)

    # OutsideWorld reports

    # TODO return a list of Events, to ease JSON output
    @_needs_outside_world
    def extract_timeline(self) -> List["ReportGenerator.EventLine"]:
        """
        Return the chronological list of events, as tuples with elements
        identifier, time, category, context and data.
        """

        timeline = list(map(self.event_to_tuple, self.outside_world.events))
        timeline.sort()

        return timeline

    @_needs_outside_world
    def produce_timeline(self) -> str:
        """Return the formatted timeline."""
        timeline = self.extract_timeline()

        if self.output_format is ReportGenerator.Format.JSON:
            return json.dumps(timeline, indent=self.indent, sort_keys=True)
        elif self.output_format is ReportGenerator.Format.CSV:
            return self.events_to_csv(timeline)
        elif self.output_format is ReportGenerator.Format.TABLE:
            table = self.events_to_table(timeline)

            if self.shorten:
                return table.get_string(
                    fields=('ID', 'Category', 'Context', 'Data'))

            return table.get_string()

    @_needs_outside_world
    def produce_organized_events(self) -> str:
        """Return the events, organized by category."""
        report = dict()

        for event in self.outside_world.events:
            category = event.category.value
            event_dict = {
                'identifier': event.identifier,
                'time': event.time,
                'context': event.context,
                'data': event.data
            }

            category_list = report.get(event.category.value, [])
            report[category] = category_list + [event_dict]

        if self.output_format == ReportGenerator.Format.JSON:
            return self.to_json(report)
        elif self.output_format == ReportGenerator.Format.CSV:
            return "CSV not supported yet"
        elif self.output_format == ReportGenerator.Format.TABLE:
            output = ""

            for category in report:
                output += f"{category}:\n"
                events = [self.event_to_tuple(event)
                          for event in self.outside_world.events
                          if event.category.value == category]
                table = self.events_to_table(events)

                if self.shorten:
                    fields = ('ID', 'Context', 'Data')
                else:
                    fields = ('ID', 'Time (s)', 'Context', 'Data')

                output += table.get_string(fields=fields)
                output += "\n\n"

            return output

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

    @_needs_outside_world
    def save_outside_world(self, save_path: str) -> None:
        """Save the outside world to save_path."""
        Serializer.save(self.outside_world, save_path)
