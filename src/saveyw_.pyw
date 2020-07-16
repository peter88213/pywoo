"""Convert html or csv to yWriter project. 

Input file format: html (with visible or invisible chapter and scene tags).

Version @release

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import sys
import os

from urllib.parse import unquote

from pywriter.html.html_proof import HtmlProof
from pywriter.html.html_manuscript import HtmlManuscript
from pywriter.html.html_scenedesc import HtmlSceneDesc
from pywriter.html.html_chapterdesc import HtmlChapterDesc
from pywriter.html.html_partdesc import HtmlPartDesc
from pywriter.html.html_characters import HtmlCharacters
from pywriter.html.html_locations import HtmlLocations
from pywriter.html.html_items import HtmlItems
from pywriter.html.html_import import HtmlImport
from pywriter.html.html_outline import HtmlOutline
from pywriter.csv.csv_scenelist import CsvSceneList
from pywriter.csv.csv_plotlist import CsvPlotList
from pywriter.csv.csv_charlist import CsvCharList
from pywriter.csv.csv_loclist import CsvLocList
from pywriter.csv.csv_itemlist import CsvItemList
from pywriter.converter.yw_cnv_gui import YwCnvGui
from pywriter.html.html_form import *


class Converter(YwCnvGui):
    """Deletes temporary html or csv file after conversion. """

    def convert(self, sourcePath, document):

        YwCnvGui.convert(self, sourcePath, document)

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
    sourcePath = unquote(sourcePath.replace('file:///', ''))

    if sourcePath.endswith(HtmlProof.SUFFIX + HtmlProof.EXTENSION):
        sourceDoc = HtmlProof(sourcePath)

    elif sourcePath.endswith(HtmlManuscript.SUFFIX + HtmlManuscript.EXTENSION):
        sourceDoc = HtmlManuscript(sourcePath)

    elif sourcePath.endswith(HtmlSceneDesc.SUFFIX + HtmlSceneDesc.EXTENSION):
        sourceDoc = HtmlSceneDesc(sourcePath)

    elif sourcePath.endswith(HtmlChapterDesc.SUFFIX + HtmlChapterDesc.EXTENSION):
        sourceDoc = HtmlChapterDesc(sourcePath)

    elif sourcePath.endswith(HtmlPartDesc.SUFFIX + HtmlPartDesc.EXTENSION):
        sourceDoc = HtmlPartDesc(sourcePath)

    elif sourcePath.endswith(HtmlCharacters.SUFFIX + HtmlCharacters.EXTENSION):
        sourceDoc = HtmlCharacters(sourcePath)

    elif sourcePath.endswith(HtmlLocations.SUFFIX + HtmlLocations.EXTENSION):
        sourceDoc = HtmlLocations(sourcePath)

    elif sourcePath.endswith(HtmlItems.SUFFIX + HtmlItems.EXTENSION):
        sourceDoc = HtmlItems(sourcePath)

    elif sourcePath.endswith('.html'):
        result = read_html_file(sourcePath)

        if 'SUCCESS' in result[0]:

            if "<h3" in result[1].lower():
                sourceDoc = HtmlOutline(sourcePath)

            else:
                sourceDoc = HtmlImport(sourcePath)

        else:
            sourceDoc = None

    elif sourcePath.endswith(CsvSceneList.SUFFIX + CsvSceneList.EXTENSION):
        sourceDoc = CsvSceneList(sourcePath)

    elif sourcePath.endswith(CsvPlotList.SUFFIX + CsvPlotList.EXTENSION):
        sourceDoc = CsvPlotList(sourcePath)

    elif sourcePath.endswith(CsvCharList.SUFFIX + CsvCharList.EXTENSION):
        sourceDoc = CsvCharList(sourcePath)

    elif sourcePath.endswith(CsvLocList.SUFFIX + CsvLocList.EXTENSION):
        sourceDoc = CsvLocList(sourcePath)

    elif sourcePath.endswith(CsvItemList.SUFFIX + CsvItemList.EXTENSION):
        sourceDoc = CsvItemList(sourcePath)

    else:
        sourceDoc = None

    converter = Converter(sourcePath, sourceDoc, False)


if __name__ == '__main__':
    try:
        sourcePath = sys.argv[1]
    except:
        sourcePath = ''

    run(sourcePath)
