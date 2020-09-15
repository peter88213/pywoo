"""Convert yWriter project to odt or csv and vice versa. 

Version @release

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
from urllib.parse import unquote
from tkinter import *

from pywriter.converter.file_factory import FileFactory
from openoffice.yw_cnv_oo import YwCnvOO


def run(sourcePath, suffix, silentMode):
    converter = YwCnvOO(sourcePath, suffix, silentMode)


if __name__ == '__main__':

    try:
        sourcePath = sys.argv[1]

    except:
        sourcePath = ''

    fileName, FileExtension = os.path.splitext(sourcePath)

    if FileExtension in FileFactory.YW_EXTENSIONS:

        try:
            suffix = sys.argv[2]

        except:
            suffix = ''

    else:
        sourcePath = unquote(sourcePath.replace('file:///', ''))
        suffix = None

    run(sourcePath, suffix, False)
