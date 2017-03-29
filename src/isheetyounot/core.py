# encoding: utf-8
#
# Copyright (c) 2016 Dean Jackson <deanishe@deanishe.net>
#
# MIT Licence. See http://opensource.org/licenses/MIT
#
# Created on 2016-05-21
#

"""
core
^^^^

Loading and caching of Excel files, and helper functions.

"""


# I Sheet You Not. Search Excel data in Alfred 3.

# Pass this script the path to an Excel file via the -p option or the
# DOC_PATH environment variable.

# By default, the script reads the rows of the first worksheet in the
# workbook and generates Alfred JSON results.

# It reads the first three columns, treating the first as the result title,
# the second as its subtitle and the third as its value (arg).

from __future__ import print_function, unicode_literals, absolute_import

import hashlib
import os
import time

from .aw3 import av, human_time, log, make_item

version = '0.2.3'

# Fallback/default values
BUNDLE_ID = 'net.deanishe.alfred-i-sheet-you-not'
CACHE_DIR = os.path.join(os.path.expanduser('~/Library/Caches'), BUNDLE_ID)

# Link to GitHub issues. Output by rescue() on error.
HELP_URL = 'https://github.com/deanishe/i-sheet-you-not/issues'


class ConfigError(Exception):
    """Raised if a configuration value is not given or invalid.

    Typically, this will be a bad sheet number or name.

    If the program can't read the Excel data for other reasons,
    there'll be an exception from the underlying ``xlrd`` library.

    """

    pass


# dP                dP
# 88                88
# 88d888b. .d8888b. 88 88d888b. .d8888b. 88d888b. .d8888b.
# 88'  `88 88ooood8 88 88'  `88 88ooood8 88'  `88 Y8ooooo.
# 88    88 88.  ... 88 88.  .88 88.  ... 88             88
# dP    dP `88888P' dP 88Y888P' `88888P' dP       `88888P'
#                      88
#                      dP

def tilde(path):
    """Replace user's home directory in ``path`` with ~.

    Args:
        path (unicode): A filepath.

    Returns:
        unicode: Shortened filepath.
    """
    return path.replace(os.getenv('HOME'), '~')


#                            dP       oo
#                            88
# .d8888b. .d8888b. .d8888b. 88d888b. dP 88d888b. .d8888b.
# 88'  `"" 88'  `88 88'  `"" 88'  `88 88 88'  `88 88'  `88
# 88.  ... 88.  .88 88.  ... 88    88 88 88    88 88.  .88
# `88888P' `88888P8 `88888P' dP    dP dP dP    dP `8888P88
#                                                      .88
#                                                  d8888P

def cache_key(o):
    """Generate unique, deterministic key based on program options.

    Args:
        o (argparse.Namespace): Program's configuration object.

    Returns:
        str: MD5 hex digest of options.
    """
    # Cache key of full path and *all* variables to ensure uniqueness
    p = os.path.abspath(o.docpath)
    v = '-'.join([
        '{}={}'.format(k, v) for
        k, v in sorted(o.variables.items())
    ])

    tpl = ('{p}-{o.sheet}-{o.start_row}-{o.title_col}-'
           '{o.subtitle_col}-{o.value_col}-{v}')

    n = tpl.format(p=p, o=o, v=v)
    return hashlib.md5(n.encode('utf-8')).hexdigest()


def _cache_path(key):
    """Path for cached JSON based on key and workflow's cache directory.

    Args:
        key (str): Unique key from `cache_key()`.

    Returns:
        unicode: Filepath in cache directory with ".json" extension.
    """
    root = av.get('workflow_cache', CACHE_DIR)
    # log('cache_dir=%r', root)
    par = [key[:3], key[3:6], key[6:9]]
    dp = os.path.join(root, *par)

    # log('cache_dir=%r', tilde(dp))

    try:
        os.makedirs(dp, 0700)
    except OSError:
        pass

    p = os.path.join(dp, '{}.json'.format(key))

    log('cache_path=%r', tilde(p))

    return p


def cached_data(key, max_age=0):
    """Returned data cached for ``key`` or ``None``.

    Returns ``None`` if no data are cached for ``key`` or the age
    of the cached data exceeds `max_age` (if ``max_age`` is non-zero).

    Args:
        key (str): Cache key from :func:`cache_key`.
        max_age (int, optional): Maximum permissible age of cached data
            in seconds.

    Returns:
        str: The contents of the cache file, or ``None``.
    """
    p = _cache_path(key)

    if not os.path.exists(p):
        return None

    if max_age:
        age = time.time() - os.path.getmtime(p)
        log('cache_age=%s', human_time(age))

        if age > max_age:
            return None

    with open(p) as fp:
        return fp.read()


def cache_data(key, data):
    """Store ``data`` in cache under name ``key``.

    Args:
        key (str): Cache key from :func:``cache_key``.
        data (str): Data to write to file.
    """
    p = _cache_path(key)

    with open(p, 'wb') as fp:
        fp.write(data)


#                                     dP
#                                     88
# .d8888b. dP.  .dP .d8888b. .d8888b. 88
# 88ooood8  `8bd8'  88'  `"" 88ooood8 88
# 88.  ...  .d88b.  88.  ... 88.  ... 88
# `88888P' dP'  `dP `88888P' `88888P' dP

def read_data(path, sheet, cols, start_row=1, variables=None):
    """Read the specified cells from an Excel file.

    Args:
        path (unicode): Path of XLSX file to read data from.
        sheet (unicode): Number or name of sheet to read data from.
        cols (list): The three columns to read title, subtitle and
            value from respectively.
        start_row (int, optional): The row on which to start reading data.
        variables (dict, optional): name->col mapping of columns to read
            into result variables with the corresponding names.

    Returns:
        list: Sequence of Alfred 3 result dictionaries.

    Raises:
        ConfigError: Raised if an argument is invalid, e.g. non-existent
            sheet name.
    """
    from xlrd import open_workbook

    variables = variables or {}

    wb = open_workbook(path)

    if sheet.isdigit():
        s = wb.sheets()[int(sheet) - 1]
    else:  # Name
        for s in wb.sheets():
            if s.name == sheet:
                break
        else:
            raise ConfigError("Couldn't find sheet: {}".format(sheet))

    log('Opened worksheet "%s" of %s', s.name, tilde(path))

    start_row -= 1
    cols = [i - 1 for i in cols]

    items = []
    invalid = 0

    i = start_row

    while i < s.nrows:
        vars = {}
        sub = arg = ''
        tit = s.cell(i, cols[0]).value
        if cols[1] > -1:
            sub = s.cell(i, cols[1]).value
        if cols[2] > -1:
            arg = s.cell(i, cols[2]).value
        for k, j in variables.items():
            v = s.cell(i, j - 1).value
            if v:
                vars[k] = v

        i += 1

        # log('tit=%r, sub=%r, arg=%r', tit, sub, arg)

        if not tit:  # Invalid
            invalid += 1
            continue

        items.append(make_item(tit, sub, arg, **vars))

    log('Read %d rows from worksheet "%s"', len(items), s.name)

    return items
