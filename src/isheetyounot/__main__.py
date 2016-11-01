# encoding: utf-8
#
# Copyright (c) 2016 Dean Jackson <deanishe@deanishe.net>
#
# MIT Licence. See http://opensource.org/licenses/MIT
#
# Created on 2016-05-21
#

"""
Run the I Sheet You Not program.
"""

from __future__ import print_function, unicode_literals, absolute_import

from . import cli
from .aw3 import rescue
from .core import HELP_URL

if __name__ == '__main__':
    rescue(cli.main, HELP_URL)
