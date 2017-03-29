# encoding: utf-8
#
# Copyright (c) 2016 Dean Jackson <deanishe@deanishe.net>
#
# MIT Licence. See http://opensource.org/licenses/MIT
#
# Created on 2016-05-21
#

"""
I Sheet You Not
---------------

**Search Excel data in Alfred 3**

This package implements an Alfred workflow that generates *other*
Alfred workflows that pull data from Excel files.

The core program can both search for and in Excel files, and works
by making a copy of itself with different settings (and stripped
of the workflow-generating elements in Alfred's UI).

"""

from __future__ import print_function, unicode_literals, absolute_import

from .core import version as __version__
from .core import (
    ConfigError,
    HELP_URL,
    cache_data,
    cache_key,
    cached_data,
    read_data,
    tilde,
)


__all__ = [
    'ConfigError',
    'HELP_URL',
    '__version__',
    'cache_data',
    'cache_key',
    'cached_data',
    'read_data',
    'tilde',
]
