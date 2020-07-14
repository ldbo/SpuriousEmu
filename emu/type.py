"""Define the different VBA types."""

# TODO add all the supported types

from enum import Enum

Type = Enum("Type", "Integer Boolean")
types = tuple(Type.__members__.keys())
