"""Convert html or csv to yWriter project. 

Input file format: html (with visible or invisible chapter and scene tags).

Depends on the PyWriter library v1.9

Copyright (c) 2020, peter88213
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import sys
import os

from pywriter.globals import *
from pywriter.html.html_proof import HtmlProof
from pywriter.html.html_manuscript import HtmlManuscript
from pywriter.html.html_scenedesc import HtmlSceneDesc
from pywriter.html.html_chapterdesc import HtmlChapterDesc
from pywriter.html.html_import import HtmlImport
from pywriter.csv.csv_scenelist import CsvSceneList
from pywriter.csv.csv_plotlist import CsvPlotList
from pywriter.csv.csv_charlist import CsvCharList
from pywriter.csv.csv_loclist import CsvLocList
from pywriter.csv.csv_itemlist import CsvItemList
from pywriter.converter.yw_cnv_gui import YwCnvGui


class Converter(YwCnvGui):
    """Deletes temporary html or csv file after conversion. """

    def convert(self, sourcePath,
                document,
                extension,
                suffix):

        YwCnvGui.convert(self, sourcePath,
                         document,
                         extension,
                         suffix)

        # If an Office file exists, delete the temporary file.

        if sourcePath.endswith('.html'):

            if os.path.isfile(sourcePath.replace('.html', '.odt')):
                try:
                    os.remove(sourcePath)
                except:
                    pass

        elif sourcePath.endswith('.csv'):

            if os.path.isfile(sourcePath.replace('.csv', '.ods')):
                try:
                    os.remove(sourcePath)
                except:
                    pass


def run(sourcePath):
    sourcePath = sourcePath.replace('file:///', '').replace('%20', ' ')

    if sourcePath.endswith(PROOF_SUFFIX + '.html'):
        suffix = PROOF_SUFFIX
        extension = 'html'
        sourceDoc = HtmlProof(sourcePath)

    elif sourcePath.endswith(MANUSCRIPT_SUFFIX + '.html'):
        suffix = MANUSCRIPT_SUFFIX
        extension = 'html'
        sourceDoc = HtmlManuscript(sourcePath)

    elif sourcePath.endswith(SCENEDESC_SUFFIX + '.html'):
        suffix = SCENEDESC_SUFFIX
        extension = 'html'
        sourceDoc = HtmlSceneDesc(sourcePath)

    elif sourcePath.endswith(CHAPTERDESC_SUFFIX + '.html'):
        suffix = CHAPTERDESC_SUFFIX
        extension = 'html'
        sourceDoc = HtmlChapterDesc(sourcePath)

    elif sourcePath.endswith(PARTDESC_SUFFIX + '.html'):
        suffix = PARTDESC_SUFFIX
        extension = 'html'
        sourceDoc = HtmlChapterDesc(sourcePath)

    elif sourcePath.endswith('.html'):
        suffix = ''
        extension = 'html'
        sourceDoc = HtmlImport(sourcePath)

    elif sourcePath.endswith(SCENELIST_SUFFIX + '.csv'):
        suffix = SCENELIST_SUFFIX
        extension = 'csv'
        sourceDoc = CsvSceneList(sourcePath)

    elif sourcePath.endswith(PLOTLIST_SUFFIX + '.csv'):
        suffix = PLOTLIST_SUFFIX
        extension = 'csv'
        sourceDoc = CsvPlotList(sourcePath)

    elif sourcePath.endswith(CHARLIST_SUFFIX + '.csv'):
        suffix = CHARLIST_SUFFIX
        extension = 'csv'
        sourceDoc = CsvCharList(sourcePath)

    elif sourcePath.endswith(LOCLIST_SUFFIX + '.csv'):
        suffix = LOCLIST_SUFFIX
        extension = 'csv'
        sourceDoc = CsvLocList(sourcePath)

    elif sourcePath.endswith(ITEMLIST_SUFFIX + '.csv'):
        suffix = ITEMLIST_SUFFIX
        extension = 'csv'
        sourceDoc = CsvItemList(sourcePath)

    else:
        suffix = ''
        extension = ''
        sourceDoc = None

    converter = Converter(sourcePath, sourceDoc,
                          extension, False, suffix)


if __name__ == '__main__':
    try:
        sourcePath = sys.argv[1]
    except:
        sourcePath = ''

    run(sourcePath)
