"""Import and export yWriter data (OpenOffice 3/4 variant).

Copyright (c) 2021 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.converter.yw_cnv_ui import YwCnvUi


class CnvOpen(YwCnvUi):
    """yWriter converter with a simple tkinter GUI. 
    Open the new file after conversion from yw.
    """

    def export_from_yw(self, sourceFile, targetFile):
        """Method for conversion from yw to other.
        """
        YwCnvUi.export_from_yw(self, sourceFile, targetFile)

        if self.newFile:
            self.open_newFile()
