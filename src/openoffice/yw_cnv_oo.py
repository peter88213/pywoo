"""Import and export yWriter data (OpenOffice 3/4 variant).

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import sys
import subprocess

from pywriter.converter.ui_tk import UiTk
from pywriter.converter.yw_cnv_tk import YwCnvTk


class YwCnvOo(YwCnvTk):
    """yWriter converter with a simple tkinter GUI. 
    Handles temporary files created by OpenOffice.
    Can call OpenOffice to edit the conversion result.
    """

    def __init__(self, silentMode=False):
        self.userInterface = UiTk('yWriter import/export')
        self.success = False
        self._newFile = None
        self.fileFactory = None

    def run(self, sourcePath, suffix=None):
        YwCnvTk.run(self, sourcePath, suffix)

        if self.success:
            self.delete_tempfile(sourcePath)

        self.userInterface.finish()

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
