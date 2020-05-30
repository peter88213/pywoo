"""Convert html or csv to yWriter project. 

Input file format: html (with visible or invisible chapter and scene tags).

Depends on the PyWriter library v1.6

Copyright (c) 2020, peter88213
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import sys
import os

from pywriter.globals import (PROOF_HTML, MANUSCRIPT_HTML, SCENEDESC_HTML, CHAPTERDESC_HTML,
                              PARTDESC_HTML, SCENELIST_CSV, PLOTLIST_CSV, CHARLIST_CSV, LOCLIST_CSV, ITEMLIST_CSV)
from pywriter.globals import (PROOF_SUFFIX, MANUSCRIPT_SUFFIX, SCENEDESC_SUFFIX, CHAPTERDESC_SUFFIX,
                              PARTDESC_SUFFIX, SCENELIST_SUFFIX, PLOTLIST_SUFFIX, CHARLIST_SUFFIX, LOCLIST_SUFFIX, ITEMLIST_SUFFIX)
from pywriter.html.html_proof import HtmlProof
from pywriter.html.html_manuscript import HtmlManuscript
from pywriter.html.html_scenedesc import HtmlSceneDesc
from pywriter.html.html_chapterdesc import HtmlChapterDesc
from pywriter.csv.csv_scenelist import CsvSceneList
from pywriter.csv.csv_plotlist import CsvPlotList
from pywriter.csv.csv_charlist import CsvCharList
from pywriter.csv.csv_loclist import CsvLocList
from pywriter.csv.csv_itemlist import CsvItemList
from pywriter.converter.yw_cnv_gui import YwCnvGui

TAILS = [PROOF_HTML, MANUSCRIPT_HTML, SCENEDESC_HTML,
         CHAPTERDESC_HTML, PARTDESC_HTML, SCENELIST_CSV,
         PLOTLIST_CSV, CHARLIST_CSV, LOCLIST_CSV, ITEMLIST_CSV]

YW_EXTENSIONS = ['.yw7', '.yw6', '.yw5']


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

    if sourcePath.endswith(PROOF_HTML):
        suffix = PROOF_SUFFIX
        extension = 'html'
        sourceDoc = HtmlProof(sourcePath)

    elif sourcePath.endswith(MANUSCRIPT_HTML):
        suffix = MANUSCRIPT_SUFFIX
        extension = 'html'
        sourceDoc = HtmlManuscript(sourcePath)

    elif sourcePath.endswith(SCENEDESC_HTML):
        suffix = SCENEDESC_SUFFIX
        extension = 'html'
        sourceDoc = HtmlSceneDesc(sourcePath)

    elif sourcePath.endswith(CHAPTERDESC_HTML):
        suffix = CHAPTERDESC_SUFFIX
        extension = 'html'
        sourceDoc = HtmlChapterDesc(sourcePath)

    elif sourcePath.endswith(PARTDESC_HTML):
        suffix = PARTDESC_SUFFIX
        extension = 'html'
        sourceDoc = HtmlChapterDesc(sourcePath)

    elif sourcePath.endswith(SCENELIST_CSV):
        suffix = SCENELIST_SUFFIX
        extension = 'csv'
        sourceDoc = CsvSceneList(sourcePath)

    elif sourcePath.endswith(PLOTLIST_CSV):
        suffix = PLOTLIST_SUFFIX
        extension = 'csv'
        sourceDoc = CsvPlotList(sourcePath)

    elif sourcePath.endswith(CHARLIST_CSV):
        suffix = CHARLIST_SUFFIX
        extension = 'csv'
        sourceDoc = CsvCharList(sourcePath)

    elif sourcePath.endswith(LOCLIST_CSV):
        suffix = LOCLIST_SUFFIX
        extension = 'csv'
        sourceDoc = CsvLocList(sourcePath)

    elif sourcePath.endswith(ITEMLIST_CSV):
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
