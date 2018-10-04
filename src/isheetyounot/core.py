#!/usr/bin/env python
# encoding: utf-8
#
# Copyright (c) 2016 Dean Jackson <deanishe@deanishe.net>
#
# MIT Licence. See http://opensource.org/licenses/MIT
#
# Created on 2016-05-21
#

"""I Sheet You Not. Search Excel data in Alfred 3.

Pass this script the path to an Excel file via the -p option or the
DOC_PATH environment variable.

By default, the script reads the rows of the first worksheet in the
workbook and generates Alfred JSON results.

It reads the first three columns, treating the first as the result title,
the second as its subtitle and the third as its value (arg).

"""

from __future__ import print_function, unicode_literals, absolute_import

import hashlib
import os
import time

from .aw3 import av, human_time, log, make_item

from xlrd import (
    XL_CELL_EMPTY as TYPE_EMPTY,
    XL_CELL_TEXT as TYPE_TEXT,
    XL_CELL_NUMBER as TYPE_NUMBER,
    XL_CELL_DATE as TYPE_DATE,
    XL_CELL_BOOLEAN as TYPE_BOOLEAN,
    XL_CELL_ERROR as TYPE_ERROR,
    XL_CELL_BLANK as TYPE_BLANK,
)
from xlrd.xldate import xldate_as_datetime

# Workflow version number
version = '0.3.2'

# Fallback/default values
BUNDLE_ID = 'net.deanishe.alfred-i-sheet-you-not'
CACHE_DIR = os.path.join(os.path.expanduser('~/Library/Caches'), BUNDLE_ID)

# Link to GitHub issues. Output by rescue() on error.
HELP_URL = 'https://github.com/deanishe/i-sheet-you-not/issues'

# Excel's start date + 1 day (Jan 0 doesn't exist in Python)
# START_DATE = date(1900, 1, 1)
DEFAULT_DATE_FORMAT = '%Y-%m-%d'
DATE_FORMAT = os.getenv('DATE_FORMAT') or DEFAULT_DATE_FORMAT


class ConfigError(Exception):
    """Raised if a configuration value is not given or invalid.

    Typically, this will be a bad sheet number or name.

    If the program can't read the Excel data for other reasons,
    there'll be an exception from the underlying `xlrd` library.

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
    """Replace user's home directory in `path` with ~.

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
        k, v in sorted(o.variables.items() + o.formats.items())
    ])

    tpl = ('{p}-{o.sheet}-{o.start_row}-{o.title_col}-'
           '{o.subtitle_col}-{o.value_col}-{o.match}-{v}')

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
    """Returned data cached for `key` or `None`.

    Returns `None` if no data are cached for `key` or the age
    of the cached data exceeds `max_age` (if `max_age` is non-zero).

    Args:
        key (str): Cache key from `cache_key()`.
        max_age (int, optional): Maximum permissible age of cached data
            in seconds.

    Returns:
        str: The contents of the cache file, or `None`.
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
    """Store `data` in cache under name `key`.

    Args:
        key (str): Cache key from `cache_key()`.
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


def cell_type(cell):
    """Return type of cell.

    Args:
        cell (xlrd.sheet.Cell): Excel cell

    Returns:
        str: Type of cell as text
    """
    if cell.ctype == TYPE_BLANK:
        return 'blank'
    if cell.ctype == TYPE_BOOLEAN:
        return 'boolean'
    if cell.ctype == TYPE_DATE:
        return 'date'
    if cell.ctype == TYPE_EMPTY:
        return 'empty'
    if cell.ctype == TYPE_ERROR:
        return 'error'
    if cell.ctype == TYPE_NUMBER:
        return 'number'
    if cell.ctype == TYPE_TEXT:
        return 'text'


class Formatter(object):
    """Format Excel values according to column-specific format strings.

    Format strings should be sprintf- or strftime-style (for date columns)
    patterns.

    Attributes:
        datemode (int): Date mode of sheet this formatter is for
        formats (dict): Column -> format string mapping

    """

    def __init__(self, datemode, formats=None):
        self.datemode = datemode
        self.formats = {}
        formats = formats or {}
        for col, pat in formats.items():
            self.set(col, pat)

    def get(self, col):
        """Get format pattern (or None) for a specific column.

        Args:
            col (int): Column index (1-indexed)

        Returns:
            str: Format pattern or None
        """
        return self.formats.get(col)

    def set(self, col, pat):
        """Set format pattern for column.

        Args:
            col (int): Column index (1-indexed)
            pat (str): Format pattern
        """
        if not pat:
            return

        self.formats[col] = pat

    def format(self, col, cell):
        """Format a value with the pattern set for column.

        If no format pattern is set for column, value is returned
        unchanged.

        Args:
            col (int): Column number
            cell (xlrd.sheet.Cell): Excel cell

        Returns:
            str: Formatted value

        """
        pat = self.get(col)
        log('col=%r, pat=%r, cell=%r', col, pat, cell)
        if not pat or cell.ctype in (TYPE_BOOLEAN, TYPE_ERROR, TYPE_EMPTY):
            return self._format_default(cell)

        if cell.ctype == TYPE_DATE:
            dt = xldate_as_datetime(cell.value, self.datemode)
            formatted = dt.strftime(pat)

        else:
            try:
                formatted = pat % cell.value
            except Exception:  # Try new-style formatting
                try:
                    formatted = pat.format(cell.value)
                except Exception:
                    formatted = cell.value

        # log('pat=%r, %r  -->  %r', pat, cell.value, formatted)
        return formatted

    def _format_default(self, cell):
        """Return cell value with default formatting.

        Args:
            cell (xlrd.sheet.Cell): Excel cell

        Returns:
            str: Formatted cell value

        """
        if cell.ctype == TYPE_BOOLEAN:
            if cell.value:
                return 'yes'
            else:
                return 'no'

        if cell.ctype == TYPE_ERROR:
            return '<error>'

        if cell.ctype == TYPE_EMPTY:
            return ''

        if cell.ctype == TYPE_DATE:
            dt = xldate_as_datetime(cell.value, self.datemode)
            return dt.strftime(DATE_FORMAT)

        return cell.value


def read_data(path, sheet, cols, start_row=1, variables=None,
              formats=None, match=None):
    """Read the specified cells from an Excel file.

    Args:
        path (unicode): Path of XLSX file to read data from.
        sheet (unicode): Number or name of sheet to read data from.
        cols (list): The three columns to read title, subtitle and
            value from respectively.
        start_row (int, optional): The row on which to start reading data.
        variables (dict, optional): name->col mapping of columns to read
            into result variables with the corresponding names.
        formats (dict, optional): index->format mapping of sprintf-style
            format strings for columns.
        match (str, optional): ``sprintf``-style format string for match
            field.

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
    fmt = Formatter(wb.datemode, formats)
    # cols = [i - 1 for i in cols]

    items = []
    invalid = 0

    i = start_row

    while i < s.nrows:
        evars = {}
        match_data = None
        sub = arg = ''
        cell = s.cell(i, cols[0] - 1)
        tit = fmt.format(cols[0], cell)
        log('[title] i=%d, cell=%r, value=%r', i, cell, tit)
        if cols[1] > -1:
            cell = s.cell(i, cols[1] - 1)
            sub = fmt.format(cols[1], cell)
            log('[subtitle] i=%d, cell=%r, value=%r', i, cell, sub)
        if cols[2] > -1:
            cell = s.cell(i, cols[2] - 1)
            arg = fmt.format(cols[2], cell)
            log('[value] i=%d, cell=%r, value=%r', i, cell, arg)

        for k, j in variables.items():
            value = None
            cell = s.cell(i, j - 1)
            value = fmt.format(j, cell)
            evars[k] = value
            log('[var:%s] i=%d, cell=%r, type=%s, value=%r', k, i, cell,
                cell_type(cell), value)

        if match:
            try:
                match_data = match % evars
                log('[match] match=%s, evars=%r, match_data=%s',
                    match, evars, match_data)
            except Exception as err:
                log('[match] error formatting "%s" with %r: %s',
                    match, evars, err)

        i += 1

        log('formats=%r, cols=%r, tit=%r, sub=%r, arg=%r, match=%r', formats,
            cols, tit, sub, arg, match_data)

        if not tit:  # Invalid
            invalid += 1
            continue

        items.append(make_item(tit, sub, arg, match=match_data, **evars))

    log('Read %d rows from worksheet "%s"', len(items), s.name)

    return items
