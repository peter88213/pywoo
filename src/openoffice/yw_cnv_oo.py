"""Import and export yWriter data (OpenOffice 3/4 variant).

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import subprocess
from tkinter import *

from pywriter.converter.yw_cnv_tk import YwCnvTk
from pywriter.converter.file_factory import FileFactory


class YwCnvOO(YwCnvTk):
    """yWriter converter with a simple tkinter GUI. 
    Handles temporary files created by OpenOffice.
    Can call OpenOffice to edit the conversion result.
    """

    def convert(self, sourceFile, targetFile):
        YwCnvTk.convert(self, sourceFile, targetFile)
        self._newFile = None

        if self.success and sourceFile.EXTENSION in FileFactory.YW_EXTENSIONS:
            self._newFile = targetFile.filePath
            self.root.editButton = Button(
                text="Edit", command=self.edit)
            self.root.editButton.config(height=1, width=10)
            self.root.editButton.pack(padx=5, pady=5)

        elif sourceFile.EXTENSION == '.html':

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

    def edit(self):

        OPENOFFICE = ['c:/Program Files/OpenOffice.org 3/program/swriter.exe',
                      'c:/Program Files (x86)/OpenOffice.org 3/program/swriter.exe',
                      'c:/Program Files/OpenOffice 4/program/swriter.exe',
                      'c:/Program Files (x86)/OpenOffice 4/program/swriter.exe']

        for office in OPENOFFICE:

            if os.path.isfile(office):

                if self._newFile.endswith('.csv'):
                    office = office.replace('swriter', 'scalc')

                subprocess.Popen([os.path.normpath(office),
                                  os.path.normpath(self._newFile)])
                sys.exit(0)
