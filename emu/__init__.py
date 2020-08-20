"""SpuriousEmu interface"""

from .compiler import Compiler, Unit, Program
from .error import *
from .interpreter import Interpreter
from .preprocessor import Preprocessor
from .report import ReportGenerator
from .serialize import Serializer
from .syntax import Parser, AST
from .side_effect import OutsideWorld

__version__ = "0.2.0"
