#!/usr/bin/env python3
"""Convert yWriter project to odt or ods and vice versa. 

Version @release

Copyright (c) 2021 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import sys
from urllib.parse import unquote
import platform

from openoffice.cnv_open import CnvOpen

from pywriter.converter.universal_file_factory import UniversalFileFactory
from pywriter.ui.ui_tk import UiTk


def run(sourcePath, suffix=None):
    ui = UiTk('yWriter import/export (Python version ' +
              str(platform.python_version()) + ')')
    converter = CnvOpen()
    converter.ui = ui
    converter.fileFactory = UniversalFileFactory()
    converter.run(sourcePath, suffix)
    ui.start()


if __name__ == '__main__':

    try:
        sourcePath = unquote(sys.argv[1].replace('file:///', ''))

    except:
        sourcePath = ''

    fileName, FileExtension = os.path.splitext(sourcePath)

    if not FileExtension in CnvOpen.YW_EXTENSIONS:
        # Source file is not a yWriter project
        suffix = None

    else:
        # Source file is a yWriter project; suffix matters

        try:
            suffix = sys.argv[2]

        except:
            suffix = ''

    run(sourcePath, suffix)
