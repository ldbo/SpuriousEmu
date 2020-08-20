"""SpuriousEmu interface"""

from .compiler import Compiler, Unit, Program
from .error import *
from .interpreter import Interpreter
from .preprocessor import Preprocessor
from .report import ReportGenerator
from .serialize import Serializer
from .syntax import Parser, AST

__version__ = "0.1.1"
