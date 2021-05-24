"""Provide a converter class for universal import and export. 

Copyright (c) 2021 Peter Triesberger
For further information see https://github.com/peter88213/pywoo
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.converter.universal_converter import UniversalConverter


class Converter(UniversalConverter):
    """Override the export_from_yw() method. 
    Open the new file after conversion from yw.
    """

    def export_from_yw(self, sourceFile, targetFile):
        """Override the super class method."""
        message = self.convert(sourceFile, targetFile)

        if message.startswith('SUCCESS'):
            self.newFile = targetFile.filePath
            self.open_newFile()

        else:
            self.newFile = None
            self.ui.set_info_how(message)
