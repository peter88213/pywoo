"""Convert yWriter project to odt or csv and vice versa. 

Version @release

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import sys

from urllib.parse import unquote
from openoffice.yw_cnv_oo import YwCnvOo


if __name__ == '__main__':

    try:
        sourcePath = unquote(sys.argv[1].replace('file:///', ''))

    except:
        sourcePath = ''

    fileName, FileExtension = os.path.splitext(sourcePath)

    if not FileExtension in YwCnvOo.YW_EXTENSIONS:
        # Source file is not a yWriter project
        suffix = None

    else:
        # Source file is a yWriter project; suffix matters

        try:
            suffix = sys.argv[2]

        except:
            suffix = ''

    converter = YwCnvOo(sourcePath, suffix, False)
