Lexer and parser
================

SpuriousEmu now [#]_ relies on a hand-written two-pass lexer and a recursive
descent parser, which provide a good compromise between runtime speed and
development time, and allow an easy handling of errors.


Lexer
^^^^^

First, the input code source is split into :py:class:`Token <emu.lexer.Token>`
using the :py:class:`Lexer <emu.lexer.Lexer>` class. It implements a two-pass
lexer:

 - the first stage is based on regular expressions, it correctly classifies most
   of the tokens but not keywords;
 - the second stage distinguishes between keywords and identifiers using fast
   set membership test.

This strategy allows to have short lexing time, despite the huge number of
keywords in the VBA language.

.. autoclass:: emu.lexer.Lexer
   :members:

.. autoclass:: emu.lexer.Token
   :members:


.. [#] Different kind of parsers have been used throughout the development of
       SpuriousEmu, namely:

        - first a PEG parser generated with
          `PyParsing <https://github.com/pyparsing/pyparsing>`_, which was easy to
          maintain but too slow for VBA;
        - then an LALR parser thanks to `Lark <https://github.com/lark-parser/lark>`_,
          which was faster but required too much development time to be efficient.


