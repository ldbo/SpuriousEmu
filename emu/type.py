"""Define the different VBA types."""

# TODO add all the supported types

from enum import Enum

types = (
    'Integer',
    'Boolean'
)
Type = Enum("Type", ' '.join(types))
