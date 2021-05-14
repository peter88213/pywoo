"""Import and export yWriter data (OpenOffice 3/4 variant).

Copyright (c) 2021 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.converter.universal_exporter import UniversalExporter
from pywriter.ui.ui_tk_open import UiTkOpen


class PywooExporter(UniversalExporter):
    """yWriter converter with a simple tkinter GUI. 
    Show 'Open' button after conversion from yw.
    """

    def __init__(self):
        """Define instance variables.

        ui -- user interface object; instance of Ui or a Ui subclass.
        """
        UniversalExporter.__init__(self)
        self.ui = UiTkOpen('yWriter import/export')

    def export_from_yw(self, sourceFile, targetFile):
        """Method for conversion from yw to other.
        """
        UniversalExporter.export_from_yw(self, sourceFile, targetFile)

        if self.newFile:
            self.ui.show_open_button(self.open_newFile)
