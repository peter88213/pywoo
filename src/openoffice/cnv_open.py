"""Import and export yWriter data (OpenOffice 3/4 variant).

Copyright (c) 2021 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import platform
from pywriter.converter.universal_converter import UniversalConverter
from pywriter.ui.ui_tk_open import UiTkOpen


class CnvOpen(UniversalConverter):
    """yWriter converter with a simple tkinter GUI. 
    Open the new file after conversion from yw.
    """

    def __init__(self):
        """Define instance variables.

        ui -- user interface object; instance of Ui or a Ui subclass.
        """
        UniversalConverter.__init__(self)
        self.ui = UiTkOpen('yWriter import/export (Python version ' +
                           str(platform.python_version()) + ')')

    def export_from_yw(self, sourceFile, targetFile):
        """Method for conversion from yw to other.
        """
        UniversalConverter.export_from_yw(self, sourceFile, targetFile)

        if self.newFile:
            self.open_newFile()
