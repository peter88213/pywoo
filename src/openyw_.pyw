"""Convert yWriter project to odt or csv. 

Input file format: yWriter
Output file format: odt (with visible or invisible chapter and scene tags) or csv.

Version @release

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import subprocess
from tkinter import *

from pywriter.globals import *
from pywriter.odt.odt_proof import OdtProof
from pywriter.odt.odt_manuscript import OdtManuscript
from pywriter.odt.odt_scenedesc import OdtSceneDesc
from pywriter.odt.odt_chapterdesc import OdtChapterDesc
from pywriter.odt.odt_partdesc import OdtPartDesc
from pywriter.csv.csv_scenelist import CsvSceneList
from pywriter.csv.csv_plotlist import CsvPlotList
from pywriter.csv.csv_charlist import CsvCharList
from pywriter.csv.csv_loclist import CsvLocList
from pywriter.csv.csv_itemlist import CsvItemList
from pywriter.odt.odt_export import OdtExport
from pywriter.converter.yw_cnv_gui import YwCnvGui
from pywriter.odt.odt_characters import OdtCharacters
from pywriter.odt.odt_items import OdtItems
from pywriter.odt.odt_locations import OdtLocations

OPENOFFICE = ['c:/Program Files/OpenOffice.org 3/program/swriter.exe',
              'c:/Program Files (x86)/OpenOffice.org 3/program/swriter.exe',
              'c:/Program Files/OpenOffice 4/program/swriter.exe',
              'c:/Program Files (x86)/OpenOffice 4/program/swriter.exe']


class Converter(YwCnvGui):
    """Standalone yWriter 7 converter with a simple GUI. """

    def convert(self, sourcePath,
                document,
                extension,
                suffix):

        YwCnvGui.convert(self, sourcePath,
                         document,
                         extension,
                         suffix)

        fileName, FileExtension = os.path.splitext(sourcePath)

        if self._success and FileExtension in ['.yw5', '.yw6', '.yw7']:
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

    fileName, FileExtension = os.path.splitext(sourcePath)

    if suffix == OdtProof.SUFFIX:
        extension = 'odt'
        targetDoc = OdtProof(fileName + suffix + '.odt')

    elif suffix == OdtManuscript.SUFFIX:
        extension = 'odt'
        targetDoc = OdtManuscript(fileName + suffix + '.odt')

    elif suffix == OdtSceneDesc.SUFFIX:
        extension = 'odt'
        targetDoc = OdtSceneDesc(fileName + suffix + '.odt')

    elif suffix == OdtChapterDesc.SUFFIX:
        extension = 'odt'
        targetDoc = OdtChapterDesc(fileName + suffix + '.odt')

    elif suffix == OdtPartDesc.SUFFIX:
        extension = 'odt'
        targetDoc = OdtPartDesc(fileName + suffix + '.odt')

    elif suffix == OdtCharacters.SUFFIX:
        extension = 'odt'
        targetDoc = OdtCharacters(fileName + suffix + '.odt')

    elif suffix == OdtLocations.SUFFIX:
        extension = 'odt'
        targetDoc = OdtLocations(fileName + suffix + '.odt')

    elif suffix == OdtItems.SUFFIX:
        extension = 'odt'
        targetDoc = OdtItems(fileName + suffix + '.odt')

    elif suffix == CsvSceneList.SUFFIX:
        extension = 'csv'
        targetDoc = CsvSceneList(fileName + suffix + '.csv')

    elif suffix == CsvPlotList.SUFFIX:
        extension = 'csv'
        targetDoc = CsvPlotList(fileName + suffix + '.csv')

    elif suffix == CsvCharList.SUFFIX:
        extension = 'csv'
        targetDoc = CsvCharList(fileName + suffix + '.csv')

    elif suffix == CsvLocList.SUFFIX:
        extension = 'csv'
        targetDoc = CsvLocList(fileName + suffix + '.csv')

    elif suffix == CsvItemList.SUFFIX:
        extension = 'csv'
        targetDoc = CsvItemList(fileName + suffix + '.csv')

    else:
        extension = 'odt'
        targetDoc = OdtExport(fileName + '.odt')

    converter = Converter(sourcePath, targetDoc,
                          extension, False, suffix)


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

        print(run(sourcePath, suffix))

    else:
        print('ERROR: File is not an yWriter project.')
