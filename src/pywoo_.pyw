"""Convert yWriter project to odt or csv and vice versa. 

Version @release

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import subprocess
from tkinter import *

from pywriter.converter.yw_cnv_tk import YwCnvTk
from urllib.parse import unquote

OPENOFFICE = ['c:/Program Files/OpenOffice.org 3/program/swriter.exe',
              'c:/Program Files (x86)/OpenOffice.org 3/program/swriter.exe',
              'c:/Program Files/OpenOffice 4/program/swriter.exe',
              'c:/Program Files (x86)/OpenOffice 4/program/swriter.exe']


class Converter(YwCnvTk):

    def convert(self, sourceFile, targetFile):
        YwCnvTk.convert(self, sourceFile, targetFile)
        self._newFile = None

        if self.success and sourceFile.EXTENSION in ['.yw5', '.yw6', '.yw7']:
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

        for office in OPENOFFICE:

            if os.path.isfile(office):

                if self._newFile.endswith('.csv'):
                    office = office.replace('swriter', 'scalc')

                subprocess.Popen([os.path.normpath(office),
                                  os.path.normpath(self._newFile)])
                sys.exit(0)


def run(sourcePath, suffix, silentMode):
    converter = Converter(sourcePath, suffix, silentMode)


if __name__ == '__main__':

    try:
        sourcePath = sys.argv[1]

    except:
        sourcePath = ''

    fileName, FileExtension = os.path.splitext(sourcePath)

    if FileExtension in ['.yw5', '.yw6', '.yw7']:

        try:
            suffix = sys.argv[2]

        except:
            suffix = ''

    else:
        sourcePath = unquote(sourcePath.replace('file:///', ''))
        suffix = None

    run(sourcePath, suffix, False)
