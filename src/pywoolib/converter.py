"""Provide a converter class for universal import and export. 

Copyright (c) 2021 Peter Triesberger
For further information see https://github.com/peter88213/pywoo
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import platform
from pywriter.converter.universal_converter import UniversalConverter
from pywriter.ui.ui_tk_open import UiTkOpen


class Converter(UniversalConverter):
    """Extend the Super class. 
    Open the new file after conversion from yw.
    """

    def __init__(self):
        """Extend the super class method."""
        UniversalConverter.__init__(self)
        self.ui = UiTkOpen('yWriter import/export (Python version ' +
                           str(platform.python_version()) + ')')

    def export_from_yw(self, sourceFile, targetFile):
        """Extend the super class method."""
        UniversalConverter.export_from_yw(self, sourceFile, targetFile)

        if self.newFile:
            self.open_newFile()
