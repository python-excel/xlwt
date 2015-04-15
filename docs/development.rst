Development
===========

.. highlight:: bash

This package is developed using continuous integration which can be
found here:

https://travis-ci.org/python-excel/xlwt

If you wish to contribute to this project, then you should fork the
repository found here:

https://github.com/python-excel/xlwt

Once that has been done and you have a checkout, you can follow these
instructions to perform various development tasks:

Setting up a virtualenv
-----------------------

The recommended way to set up a development environment is to turn
your checkout into a virtualenv and then install the package in
editable form as follows::

  $ virtualenv .
  $ bin/pip install -Ur requirements.txt

Running the tests
-----------------

Once you've set up a virtualenv, the tests can be run as follows::

  $ bin/nosetests

Building the documentation
--------------------------

The Sphinx documentation is built by doing the following, having activated
the virtualenv above, from the directory containing setup.py::

  $ cd docs
  $ make html

