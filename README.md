I Sheet You Not
===============

Create auto-updating [Alfred 3][alfred] workflows from Excel worksheets.


![I Sheet You Not demonstration][demo]

I Sheet You Not is a workflow generator/template for [Alfred 3][alfred].
It reads data from an Excel workbook and displays it in Alfred. You can specify
which rows and columns the data are read from, and changes to the data are
picked up automatically by the workflow.


## Download and installation ##

Download the workflow from [GitHub releases][gh-releases] and double-click
the downloaded `I-Sheet-You-Not.X.X.X.alfredworkflow` file to install in
Alfred.

## Usage ##

See [the documentation][doc] for instructions.


## Licencing, thanks ##

This workflow is released under the [MIT licence][mit].

It is based on the [xlrd library for Python][xlrd], released under a [BSD-style licence][xlrd-licence].

The workflow icon is from [IconArchive.com][icon].


[alfred]: https://www.alfredapp.com/
[doc]: http://www.deanishe.net/i-sheet-you-not/
<!-- [demo]: http://www.deanishe.net/i-sheet-you-not/_images/demo.gif -->
[demo]: doc/_static/demo.gif
[gh-releases]: https://github.com/deanishe/i-sheet-you-not/releases
[mit]: ./src/LICENCE.txt
[icon]: http://www.iconarchive.com/show/google-jfk-icons-by-carlosjj/spreadsheets-icon.html
[xlrd]: http://www.python-excel.org
[xlrd-licence]: https://github.com/python-excel/xlrd/blob/master/docs/licenses.rst
