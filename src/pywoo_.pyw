"""Convert yWriter project to odt or ods and vice versa. 

Version @release

Copyright (c) 2021 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import sys
import platform
from urllib.parse import unquote

from pywriter.converter.universal_converter import UniversalConverter
from pywriter.ui.ui_mb import UiMb

YW_EXTENSIONS = ['.yw7', '.yw6', '.yw5']


def run(sourcePath, suffix=None):
    converter = UniversalConverter()
    converter.ui = UiMb('yWriter import/export (Python version ' +
                        str(platform.python_version()) + ')')
    kwargs = {'suffix': suffix}
    converter.run(sourcePath, **kwargs)


if __name__ == '__main__':

    try:
        sourcePath = unquote(sys.argv[1].replace('file:///', ''))

    except:
        sourcePath = ''

    fileName, FileExtension = os.path.splitext(sourcePath)

    if not FileExtension in YW_EXTENSIONS:
        # Source file is not a yWriter project.
        suffix = None

    else:
        # Source file is a yWriter project; suffix matters.

        try:
            suffix = sys.argv[2]

        except:
            suffix = ''

    run(sourcePath, suffix)
