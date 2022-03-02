"""Provide a converter class for universal import and export. 

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/pywoo
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import ERROR
from pywriter.converter.yw7_converter import Yw7Converter


class Converter(Yw7Converter):
    """A converter for universal import and export.
    
    Support yWriter 7 projects and most of the Novel subclasses 
    that can be read or written by OpenOffice/LibreOffice.
    - No message in case of success when converting from yWriter.
    - Open the new file after conversion from yWriter.
    """

    def export_from_yw(self, source, target):
        """Method for conversion from yw to other.
        
        Overrides the super class method.
        Open the new file.
        Show only error messages.
        """
        message = self.convert(source, target)
        if message.startswith(ERROR):
            self.newFile = None
            self.ui.set_info_how(message)
        else:
            self.newFile = target.filePath
            self._open_newFile()
