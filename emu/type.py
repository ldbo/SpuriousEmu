"""Define the different VBA types."""

# To implement a new type, you must :
#  - add it to the list of Type fields
#  - add a matching grammar rule in syntax.py if needed
#  - update value.py to allow evaluation
#
# TODO:
#  - arrays
#  - Byte
#  - Char
#  - Date
#  - Decimal
#  - Double
#  - Long
#  - Object
#  - SByte
#  - Short
#  - Single
#  - String
#  - UInteger
#  - ULong
#  - User-Defined
#  - UShort

from enum import Enum

Type = Enum("Type", "Integer Boolean String Object Function")
types = tuple(Type.__members__.keys())
