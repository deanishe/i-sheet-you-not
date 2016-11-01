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

import argparse
import time
import os

from .core import (
    BUNDLE_ID,
    HELP_URL,
    ConfigError,
    cache_data,
    cache_key,
    cached_data,
    read_data,
    version,
)
from .aw3 import (
    Feedback,
    av,
    change_bundle_id,
    human_time,
    log,
    random_bundle_id,
)


def parse_args():
    """Read program options from the environment and command line.

    Returns:
        argparse.Namespace: Program configuration.

    """
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('-p', '--docpath',
                   metavar='FILE', type=str,
                   default=os.getenv('DOC_PATH') or './Demo.xlsx',
                   help="Excel file to read data from. "
                   "Envvar: DOC_PATH")
    p.add_argument('-n', '--sheet',
                   metavar='N', type=str,
                   default=os.getenv('SHEET') or '1',
                   help="Number or name of worksheet to read data from. "
                   "Default is the first sheet in the workbook. "
                   "Envvar: SHEET")
    p.add_argument('-r', '--row',
                   dest='start_row',
                   metavar='N', type=str,
                   default=os.getenv('START_ROW') or '1',
                   help="Number of first row to read data from. "
                   "Default is 1, i.e the first row. "
                   "Use --row 2 to ignore a title row, for example. "
                   "Envvar: START_ROW")
    p.add_argument('-t', '--title',
                   dest='title_col',
                   metavar='N', type=str,
                   default=os.getenv('TITLE_COL') or '1',
                   help="Number of column to read titles from. "
                   "Default is the first column. "
                   "Envvar: TITLE_COL")
    p.add_argument('-s', '--subtitle',
                   dest='subtitle_col',
                   metavar='N', type=str,
                   default=os.getenv('SUBTITLE_COL'),
                   help="Number of column to read subtitles from. "
                   "Default is the column after the title column. "
                   "Set to 0 if there is no subtitle column. "
                   "Envvar: SUBTITLE_COL")
    p.add_argument('-v', '--value',
                   dest='value_col',
                   metavar='N', type=str,
                   default=os.getenv('VALUE_COL'),
                   help="Number of column to read values from. "
                   "Default is the second column after the title column. "
                   "Set to 0 if there is no value column. "
                   "Envvar: VALUE_COL")
    p.add_argument('--version', action='version', version=version,
                   help="Show workflow version number and exit.")

    args = p.parse_args()
    args.docpath = os.path.expanduser(args.docpath)

    # Read VAR_ABC= values from the environment
    vars = {}
    for k in os.environ:
        if not k.startswith('VAR_') and k != 'VAR_':
            continue
        v = os.environ[k]
        if v and v.isdigit():
            vars[k[4:]] = int(v)
        else:
            log('Invalid value for `%s`: %r', k, v)

    args.variables = vars
    # args.alfred = alfred_vars()

    return args


def main():
    """Run workflow script."""
    o = parse_args()

    log('options=%r', o)

    if not o.docpath:
        raise ConfigError("You must set DOC_PATH in the workflow "
                          "configuration sheet.")

    if not os.path.exists(o.docpath):
        raise ConfigError("File does not exist : {}".format(o.docpath))

    # ---------------------------------------------------------
    # Ensure the bundle ID is *not* the default (so we can have
    # lots of copies of the workflow)
    #
    # TODO: Replace this when the workflow can create copies of itself.

    log('------ alfred env vars -------')
    for k, v in sorted(av.items()):
        log('%s=%r', k, v)
    log('------------------------------')

    if av.get('workflow_bundleid', '') == BUNDLE_ID and not os.getenv('DEV'):
        newid = random_bundle_id(BUNDLE_ID + '.')
        log('Changing bundle ID to %r ...', newid)
        change_bundle_id(newid)
        av['workflow_bundleid'] = newid

    # ---------------------------------------------------------
    # Check for valid cached data

    key = cache_key(o)
    doc_age = time.time() - os.path.getmtime(o.docpath)
    log('doc_age=%s', human_time(doc_age))
    cached = cached_data(key, max_age=doc_age)
    if cached:
        log('Using cached data.')
        print(cached)
        return 0

    # ---------------------------------------------------------
    # Data coordinates

    start_row = int(o.start_row)
    t = int(o.title_col)

    if not o.subtitle_col:
        s = t + 1
    else:
        s = int(o.subtitle_col)

    if not o.value_col:
        v = t + 2
    else:
        v = int(o.value_col)

    cols = [t, s, v]

    log('sheet=%r, start_row=%d, cols=%r, vars=%r', o.sheet, start_row,
        cols, o.variables)

    # ---------------------------------------------------------
    # Generate and cache output

    s = time.time()
    items = read_data(o.docpath, o.sheet, cols, start_row, o.variables)
    js = str(Feedback(items))
    cache_data(key, js)
    print(js)
    d = time.time() - s
    log('Updated cache in %s', human_time(d))

    return 0

if __name__ == '__main__':
    from .aw3 import rescue
    rescue(main, HELP_URL)
