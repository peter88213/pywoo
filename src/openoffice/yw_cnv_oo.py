#!/usr/bin/env python3
"""Import and export yWriter data (OpenOffice 3/4 variant).

Copyright (c) 2021 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import sys
import platform

from pywriter.converter.ui_tk import UiTk
from pywriter.converter.yw_cnv_tk import YwCnvTk


class YwCnvOo(YwCnvTk):
    """yWriter converter with a simple tkinter GUI. 
    Handles temporary files created by OpenOffice.
    Can call OpenOffice to edit the conversion result.
    """

    def __init__(self, silentMode=False):
        self.userInterface = UiTk(
            'yWriter import/export (Python version ' + str(platform.python_version()) + ")")
        self.success = False
        self._newFile = None
        self.fileFactory = None

    def finish(self, sourcePath):
        self.delete_tempfile(sourcePath)
        self.userInterface.finish()

    def edit(self):
        os.startfile(self._newFile)
        sys.exit(0)
