Utilities and errors
====================

Visitor design pattern
^^^^^^^^^^^^^^^^^^^^^^

Compilers and interpreters strongly rely on the use of the
`Visitor design pattern <https://www.wikiwand.com/en/Visitor_pattern>`_, which
is notably used for almost all of the processing of the
:py:class:`AST <emu.abstract_syntax_tree.AST>`.

.. autoclass:: emu.utils.Visitor
   :members:

.. autoclass:: emu.utils.Visitable
   :members:

Other
^^^^^

.. autofunction:: emu.utils.to_dict

.. autoclass:: emu.utils.FilePosition
   :members:

Errors
^^^^^^

.. automodule:: emu.error
   :members: