"""Import and export yWriter data (OpenOffice 3/4 variant).

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import subprocess
from tkinter import *

from pywriter.converter.ui import Ui
from pywriter.converter.ui_tk import UiTk
from pywriter.converter.yw_cnv_tk import YwCnvTk
from pywriter.converter.file_factory import FileFactory


class YwCnvOO(YwCnvTk):
    """yWriter converter with a simple tkinter GUI. 
    Handles temporary files created by OpenOffice.
    Can call OpenOffice to edit the conversion result.
    """

    def __init__(self, sourcePath, suffix=None, silentMode=False):
        """Run the converter with a GUI. """

        if silentMode:
            self.UserInterface = Ui('')

        else:
            self.UserInterface = UiTk('yWriter import/export')

        self.fileFactory = FileFactory()

        # Run the converter.

        self.success = False
        self.run_conversion(sourcePath, suffix)

        self._newFile = None

    def export_from_yw(self, sourceFile, targetFile):
        """Method for conversion from yw to other.
        """
        message = self.convert(sourceFile, targetFile)

        if message.startswith('SUCCESS'):
            self.success = True
            self.edit(targetFile.filePath)

        else:
            self.UserInterface.set_info_how(message)

    def create_yw7(self, sourceFile, targetFile):
        """TMethod for creation of a new yw7 project.
        """

        if targetFile.file_exists():
            self.UserInterface.set_info_how(
                'ERROR: "' + os.path.normpath(targetFile._filePath) + '" already exists.')

        else:
            message = self.convert(sourceFile, targetFile)
            self.UserInterface.set_info_how(message)

            if message.startswith('SUCCESS'):
                self.success = True

        self.delete_tempfile(sourceFile)

    def import_to_yw(self, sourceFile, targetFile):
        """Method for conversion from other to yw.
        """
        message = self.convert(sourceFile, targetFile)
        self.UserInterface.set_info_how(message)

        if message.startswith('SUCCESS'):
            self.success = True

        self.delete_tempfile(sourceFile)

    def delete_tempfile(self, sourceFile):

        if sourceFile.EXTENSION == '.html':

            if os.path.isfile(sourceFile.filePath.replace('.html', '.odt')):

                try:
                    os.remove(sourceFile.filePath)
                except:
                    pass

        elif sourceFile.EXTENSION == '.csv':

            if os.path.isfile(sourceFile.filePath.replace('.csv', '.ods')):

                try:
                    os.remove(sourceFile.filePath)
                except:
                    pass

    def edit(self, newFile):

        OPENOFFICE = ['c:/Program Files/OpenOffice.org 3/program/swriter.exe',
                      'c:/Program Files (x86)/OpenOffice.org 3/program/swriter.exe',
                      'c:/Program Files/OpenOffice 4/program/swriter.exe',
                      'c:/Program Files (x86)/OpenOffice 4/program/swriter.exe']

        for office in OPENOFFICE:

            if os.path.isfile(office):

                if newFile.endswith('.csv'):
                    office = office.replace('swriter', 'scalc')

                subprocess.Popen([os.path.normpath(office),
                                  os.path.normpath(newFile)])
                sys.exit(0)
