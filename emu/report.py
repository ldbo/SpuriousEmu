"""Tools to generate reports based on analysis results."""

import json

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, List, Tuple

from prettytable import PrettyTable

from .side_effect import OutsideWorld


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

    outside_world: OutsideWorld
    output_format: "ReportGenerator.Format" = field(default=Format.JSON)
    indent: int = 4
    reproducible: bool = False
    skip_identical: bool = False

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

    def produce_timeline(self) -> Any:
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
