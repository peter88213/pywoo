"""Convert yw7 to odt or csv. 

Input file format: yw7
Output file format: odt (with visible or invisible chapter and scene tags) or csv.

Depends on the PyWriter library v1.5

Copyright (c) 2020 Peter Triesberger.
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import subprocess
from tkinter import *

from pywriter.odt.odt_proof import OdtProof
from pywriter.odt.odt_manuscript import OdtManuscript
from pywriter.odt.odt_scenedesc import OdtSceneDesc
from pywriter.odt.odt_chapterdesc import OdtChapterDesc
from pywriter.odt.odt_partdesc import OdtPartDesc
from pywriter.csv.csv_scenelist import CsvSceneList
from pywriter.csv.csv_plotlist import CsvPlotList
from pywriter.odt.odt_file import OdtFile
from pywriter.converter.cnv_runner import CnvRunner

OPENOFFICE = ['c:/Program Files/OpenOffice.org 3/program/swriter.exe',
              'c:/Program Files (x86)/OpenOffice.org 3/program/swriter.exe',
              'c:/Program Files/OpenOffice 4/program/swriter.exe',
              'c:/Program Files (x86)/OpenOffice 4/program/swriter.exe']


class Converter(CnvRunner):
    """Standalone yWriter 7 converter with a simple GUI. """

    def convert(self, sourcePath,
                document,
                extension,
                suffix):

        CnvRunner.convert(self, sourcePath,
                          document,
                          extension,
                          suffix)

        if self._success and sourcePath.endswith('.yw7'):
            self._newFile = document.filePath
            self.root.editButton = Button(
                text="Edit", command=self.edit)
            self.root.editButton.config(height=1, width=10)
            self.root.editButton.pack(padx=5, pady=5)

    def edit(self):
        for office in OPENOFFICE:

            if os.path.isfile(office):

                if self._newFile.endswith('.csv'):
                    office = office.replace('swriter', 'scalc')

                subprocess.Popen([os.path.normpath(office),
                                  os.path.normpath(self._newFile)])
                sys.exit(0)


def run(sourcePath, suffix):

    if suffix == '_proof':
        extension = 'odt'
        targetDoc = OdtProof(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_manuscript':
        extension = 'odt'
        targetDoc = OdtManuscript(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_scenes':
        extension = 'odt'
        targetDoc = OdtSceneDesc(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_chapters':
        extension = 'odt'
        targetDoc = OdtChapterDesc(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_parts':
        extension = 'odt'
        targetDoc = OdtPartDesc(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_scenelist':
        extension = 'csv'
        targetDoc = CsvSceneList(sourcePath.split('.yw7')[0] + suffix + '.csv')

    elif suffix == '_plotlist':
        extension = 'csv'
        targetDoc = CsvPlotList(sourcePath.split('.yw7')[0] + suffix + '.csv')

    else:
        extension = 'odt'
        targetDoc = OdtFile(sourcePath.split('.yw7')[0] + '.odt')

    converter = Converter(sourcePath, targetDoc,
                          extension, False, suffix)


if __name__ == '__main__':
    try:
        sourcePath = sys.argv[1]
    except:
        sourcePath = ''

    if sourcePath.endswith('.yw7'):
        try:
            suffix = sys.argv[2]
        except:
            suffix = ''

        print(run(sourcePath, suffix))

    else:
        print('ERROR: File is not an yWriter 7 project.')
