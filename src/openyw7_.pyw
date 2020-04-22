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

from pywriter.fileop.odt_proof_writer import OdtProofWriter
from pywriter.fileop.odt_manuscript_writer import OdtManuscriptWriter
from pywriter.fileop.odt_scenedesc_writer import OdtSceneDescWriter
from pywriter.fileop.odt_chapterdesc_writer import OdtChapterDescWriter
from pywriter.fileop.odt_partdesc_writer import OdtPartDescWriter
from pywriter.fileop.scenelist import SceneList
from pywriter.plot.plotlist import PlotList
from pywriter.fileop.odt_file_writer import OdtFileWriter
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

        if sourcePath.endswith('.yw7'):
            self._newFile = document.filePath
            self.root.editButton = Button(
                text="Edit", command=self.edit)
            self.root.editButton.config(height=1, width=10)
            self.root.editButton.pack(padx=5, pady=5)

    def edit(self):
        for office in OPENOFFICE:

            if os.path.isfile(office):
                subprocess.Popen([os.path.normpath(office),
                                  os.path.normpath(self._newFile)])
                sys.exit(0)


def run(sourcePath, suffix):

    if suffix == '_proof':
        extension = 'odt'
        targetDoc = OdtProofWriter(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_manuscript':
        extension = 'odt'
        targetDoc = OdtManuscriptWriter(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_scenes':
        extension = 'odt'
        targetDoc = OdtSceneDescWriter(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_chapters':
        extension = 'odt'
        targetDoc = OdtChapterDescWriter(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_parts':
        extension = 'odt'
        targetDoc = OdtPartDescWriter(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_scenelist':
        extension = 'csv'
        targetDoc = SceneList(sourcePath.split('.yw7')[0] + suffix + '.csv')

    elif suffix == '_plotlist':
        extension = 'csv'
        targetDoc = PlotList(sourcePath.split('.yw7')[0] + suffix + '.csv')

    else:
        extension = 'odt'
        targetDoc = OdtFileWriter(sourcePath.split('.yw7')[0] + '.odt')

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
