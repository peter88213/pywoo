"""Provide a converter class for universal import and export. 

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/pywoo
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.converter.yw7_converter import Yw7Converter

from pywriter.model.novel import Novel


class Converter(Yw7Converter):
    """A converter for universal import and export.
    
    Support yWriter 7 projects and most of the Novel subclasses 
    that can be read or written by OpenOffice/LibreOffice.
    - No message in case of success when converting from yWriter.
    - Open the new file after conversion from yWriter.
    """

    def export_from_yw(self, source, target):
        """Convert from yWriter project to other file format.

        Positional arguments:
            source -- YwFile subclass instance.
            target -- Any Novel subclass instance.
        
        Open the new file.
        Show only error messages.
        Overrides the super class method.
        """
        try:
            self.check(source, target)
            source.novel = Novel()
            source.read()
            target.novel = source.novel
            target.write()
        except Exception as ex:
            self.newFile = None
            self.ui.set_info_how(f'!{str(ex)}')
        else:
            self.newFile = target.filePath
            self._open_newFile()
