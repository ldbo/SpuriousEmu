"""SpuriousEmu interface"""

from .code import Formatter
from .compiler import Compiler, Unit, Program
from .deobfuscation import Deobfuscator, ManglingClassifier
from .error import *
from .interpreter import Interpreter
from .preprocessor import Preprocessor
from .report import ReportGenerator
from .serialize import Serializer
from .syntax import Parser, AST
from .side_effect import OutsideWorld

__version__ = "0.4.1"
