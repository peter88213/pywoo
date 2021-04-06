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
from pywriter.converter.universal_file_factory import UniversalFileFactory
from pywriter.converter.yw_cnv_tk import YwCnvTk
from openoffice.yw_cnv_oo import YwCnvOo


class Converter(YwCnvOo):
    """yWriter converter with a simple tkinter GUI. 
    Handles temporary files created by OpenOffice.
    Calls OpenOffice to edit the conversion result.
    """

    def __init__(self, silentMode=False):
        YwCnvOo.__init__(self, silentMode)
        self.fileFactory = UniversalFileFactory()

    def export_from_yw(self, sourceFile, targetFile):
        """Method for conversion from yw to other.
        """
        YwCnvTk.export_from_yw(self, sourceFile, targetFile)

        if self.success:
            self._newFile = targetFile.filePath
            self.edit()


if __name__ == '__main__':

    try:
        sourcePath = unquote(sys.argv[1].replace('file:///', ''))

    except:
        sourcePath = ''

    fileName, FileExtension = os.path.splitext(sourcePath)

    if not FileExtension in Converter.YW_EXTENSIONS:
        # Source file is not a yWriter project
        suffix = None

    else:
        # Source file is a yWriter project; suffix matters

        try:
            suffix = sys.argv[2]

        except:
            suffix = ''

    Converter().run(sourcePath, suffix)
