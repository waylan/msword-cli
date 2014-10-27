==========
MSWord-CLI
==========
-------------------------------------------
A Command Line Interface for Microsoft Word
-------------------------------------------

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

Basic Usage
-----------

To open an existing document:

.. code:: bash

	$ msw open mydocument.docx

To print the active (focused) document:

.. code:: bash

	$ msw print

To view a list of all open documents:

.. code:: bash

	$ msw docs

	Open Documents:

	* [1] mydocument.docx
	  [2] otherdoc.docx

Notice that the asterisk ('`*`') indicates that the document at index 1 (`mydocument.docx`) is the 
currently active document. To change the focus to `otherdoc.docx` (at index 2):

.. code:: bash

	$ msw activate 2
	$ msw docs

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

	$ msw open somedoc.docx print --count 2 --pages "2-4, 6" close

Note that if any options are specified for a subcommand, those options must be specified after
the relevant subcommand and before the next subcommand in the chain. For instance, in the above 
example, `somedoc.docx` is an argument of the `open` subcommand, `--count 2 --pages "2-4, 6"` 
are options for the `print` subcommand and the `close` subcommand has no options or arguments 
defined.

Without command chaining, three separate commands would need to be issued:

.. code:: bash

	$ msw open somedoc.docx
	$ msw print --count 2 --pages "2-4, 6"
	$ msw close

Either method will accomplish the same end result. However, chaining should run a little faster
as the utility only needs to be loaded once for all commands rather than for each command.

Chaining also allows you to run different variations of the same command when that command's
options are mutually exclusive. For example, the `export` subcommand can only accept either
the `--pdf` or the `--xps` flag. If you want to export to both formats, you can chain two
`export` subcommands together :

.. code:: bash

    $ msw export --pdf . export --xps .

Note that the dot ('`.`') in the above example specifies the current working directory as the 
export path. All of the common command line paradigms should work out-of-the-box.

Dependencies
============

MSWord-CLI is built on Python_ and requires that Python version 2.7 or greater be installed
on the system. In addition to the python packages listed below, you must also have a working 
copy of Microsoft Word installed on your system.

Python Packages:

* PyWin32_
* Click_ >= 3

.. _Python: http://python.org/
.. _PyWin32: http://sf.net/projects/pywin32
.. _Click: http://click.pocoo.org/

License
=======

MSWord-CLI is licensed under the `BSD License`_ as defined in `LICENSE.txt`.

.. _`BSD License`: http://opensource.org/licenses/BSD-2-Clause
