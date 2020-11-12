Testing
=======

The test suite is fully automated, you just need to use ``nosetest`` to run it:

.. code:: bash

    $ poetry run nosetest

You can add profiling information and visualize it with `Snakeviz <https://jiffyclub.github.io/snakeviz/>`_ with:

.. code:: bash

    $ poetry run nosetest --with-cprofile --cprofile-stats-file=stats.prof
    $ poetry run snakeviz stats.prof

To develop your own tests, you can use the helper module :py:mod:`tests.test`:

:py:mod:`tests.test`
^^^^^^^^^^^^^^^^^^^^

.. automodule:: tests.test
   :members: