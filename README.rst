==========
MSWord-CLI
==========
-------------------------------------------
A Command Line Interface for Microsoft Word
-------------------------------------------

.. contents:: Table of Contents
   :backlinks: top

Summary
=======

.. default-role:: code

MSWord-CLI allows you to control Microsoft Word from the command line and/or automate
it from batch or PowerShell scripts. Among other things, you may create, open, print,
export, save, and close Word documents. Note that MSWord-CLI does not actually edit the 
content of any documents. That is beyond the scope of this project.

.. warning::
	This is **Pre-Alpha** software which is in active development. The various subcommands, options,
	and arguments are subject to change without notice and may not be sufficiently tested.

	All development has been done on Windows 7 with MS Office 2013. `Bug reports`_ are welcome on any 
	system, although the reporter may need to do much of the work on other systems and/or versions.

.. _`Bug reports`: https://github.com/waylan/msword-cli/issues/new

Installation
============

Installing the latest development code
--------------------------------------

If you would like to use the bleeding edge, you can install directly from the the Github Repo. 
However, be aware that this code is not guaranteed to be stable or even run. It is recommended
that stable releases be installed instead. Proceed at your own risk.

From the command line execute the following commands as an Administrator:

.. code:: bash

	> git clone https://github.com/waylan/msword-cli.git
	> cd msword-cli
	> python setup.py install

These instructions assume that `Git for Windows`_, `Python`_ and `Setuptools`_ are already installed 
on your system. 

.. _`Git for Windows`: http://www.git-scm.com/downloads

Making the `msw` command available on your PATH
-----------------------------------------------

Todo...

Basic Usage
===========

To open an existing document:

.. code:: bash

	> msw open mydocument.docx

To print the active (focused) document:

.. code:: bash

	> msw print

To view a list of all open documents:

.. code:: bash

	> msw docs

	Open Documents:

	* [1] mydocument.docx
	  [2] otherdoc.docx

Notice that the asterisk ('`*`') indicates that the document at index 1 (`mydocument.docx`) is the 
currently active document. To change the focus to `otherdoc.docx` (at index 2):

.. code:: bash

	> msw activate 2
	> msw docs

	Open Documents:

	  [1] mydocument.docx
	* [2] otherdoc.docx

Unless otherwise specified all subcommands work on the active document.  

For a complete list of commands and options, run `msw --help` from the command line. For help
with a specific subcommand, run `msw <subcommand> --help`.

Chaining
--------

Subcommands can be chained together. For example, to open a document, print two copies of 
pages 2, 3, 4, and 6 of that document, and then close the document, the following single 
command is all that is needed:

.. code:: bash

	> msw open somedoc.docx print --count 2 --pages "2-4, 6" close

Note that if any options are specified for a subcommand, those options must be specified after
the relevant subcommand and before the next subcommand in the chain. For instance, in the above 
example, `somedoc.docx` is an argument of the `open` subcommand, `--count 2 --pages "2-4, 6"` 
are options for the `print` subcommand and the `close` subcommand has no options or arguments 
defined.

Without command chaining, three separate commands would need to be issued:

.. code:: bash

	> msw open somedoc.docx
	> msw print --count 2 --pages "2-4, 6"
	> msw close

Either method will accomplish the same end result. However, chaining should run a little faster
as the utility only needs to be loaded once for all commands rather than for each command.

Chaining also allows you to run different variations of the same command when that command's
options are mutually exclusive. For example, the `export` subcommand can only accept either
the `--pdf` or the `--xps` flag. If you want to export to both formats, you can chain two
`export` subcommands together :

.. code:: bash

    > msw export --pdf . export --xps .

Note that the dot ('`.`') in the above example specifies the current working directory as the 
export path. All of the common command line paradigms should work out-of-the-box.

Plugins
-------

MSWord-CLI includes support for third-party plugins. A plugin can add additional subcommands
which can be included in a chain. For example, one might desire to have the ability to import
some data to fill a form (perhaps content controls). While it would be unrealistic to try to
include such a script with MSWord-CLI that could meet everyone's needs, there is no reason
why an individual user could not develop a special purpose script to meet her specific needs.

While the script could be written as a standalone script, it would also be convenient to be
able to include the call within a chain. That way, the document could be opened, the data imported,
and then the document could be printed and closed -- all from a single command.

All commands need to be defined as `Click`_ commands. Create a new python file named `msw_import.py`
and define your command:

.. code:: python

    import click

    @click.command('import')
    def imprt():
        ''' Import data. '''
        click.echo('Data is being imported...')

Note that while the command is labeled 'import' (which will be used from the command line), the 
function is named `imprt` so as not to clash with Python's `import` statement. Currently, the 
new command only prints a mesage to the console and exits. Before developing the new command's 
functionality, tell MSWord-CLI about the new subcommand and verify that it can be called. 
To do that, create a second python file named `setup.py` and include a setup script:

.. code:: python

	from setuptools import setup

	setup(
	    name='MSWImportPlugin',
	    version="1",
	    description="Import plugin for MSWord_CLI",
	    py_modules=['msw_import'],
	    entry_points="""
	        [msw.plugin]
	        import=msw_import:imprt
	    """
	)

The key is in the `entry_points`. An entry point was added to the `msw.plugin` group named 'import'
which points to the `imprt` function at its path (`msw_import:imprt`). Additional commands could
be defined from the same Python module. Simply add an additional line to the `entry_points` for 
each one.

Finally, for MSWord-CLI to find the new plugin, it needs to be installed.

.. code:: bash

	> python setup.py install

The above command will do the trick. However, as the plugin isn't finished yet, is would be helpful
to use a special development mode which sets up the path to run the plugin from the source file 
rather than Python's `site-packages` directory. That way, any changes made to the file will 
immediately take effect with no need to reinstall the plugin.

.. code:: bash

	> python setup.py develop

Now that the plugin is installed, test the script:

.. code:: bash
	
	> msw --help

You should find the `import` subcommand listed among the default subcommands in the help messge. 
To ensure that the new subcommand works, try running it:

.. code:: bash
	
	> msw import
	Data is being imported...

As the message was printed to the console, the new `import` subcommand is being called. Now 
the functionally of the `import` subcommand can be fleshed out, which is left as an exercise 
for the reader.

Dependencies
============

MSWord-CLI is built on Python_ and requires that Python version 2.7 or greater be installed
on the system. In addition to the python packages listed below, you must also have a working 
copy of Microsoft Word installed on your system.

Python Packages:

* PyWin32_
* Click_ (version 3)
* Setuptools_

.. _Python: http://python.org/
.. _PyWin32: http://sf.net/projects/pywin32
.. _Click: http://click.pocoo.org/
.. _Setuptools: https://pypi.python.org/pypi/setuptools

License
=======

MSWord-CLI is licensed under the `BSD License`_ as defined in `LICENSE.txt`.

.. _`BSD License`: http://opensource.org/licenses/BSD-2-Clause
