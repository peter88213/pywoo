"""Convert html or csv to yWriter project. 

Input file format: html (with visible or invisible chapter and scene tags).

Depends on the PyWriter library v1.6

Copyright (c) 2020, peter88213
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import sys
import os

from pywriter.html.html_proof import HtmlProof
from pywriter.html.html_manuscript import HtmlManuscript
from pywriter.html.html_scenedesc import HtmlSceneDesc
from pywriter.html.html_chapterdesc import HtmlChapterDesc
from pywriter.csv.csv_scenelist import CsvSceneList
from pywriter.csv.csv_plotlist import CsvPlotList
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

    if sourcePath.endswith('_proof.html'):
        suffix = '_proof'
        extension = 'html'
        sourceDoc = HtmlProof(sourcePath)

    elif sourcePath.endswith('_manuscript.html'):
        suffix = '_manuscript'
        extension = 'html'
        sourceDoc = HtmlManuscript(sourcePath)

    elif sourcePath.endswith('_scenes.html'):
        suffix = '_scenes'
        extension = 'html'
        sourceDoc = HtmlSceneDesc(sourcePath)

    elif sourcePath.endswith('_chapters.html'):
        suffix = '_chapters'
        extension = 'html'
        sourceDoc = HtmlChapterDesc(sourcePath)

    elif sourcePath.endswith('_parts.html'):
        suffix = '_parts'
        extension = 'html'
        sourceDoc = HtmlChapterDesc(sourcePath)

    elif sourcePath.endswith('_scenelist.csv'):
        suffix = '_scenelist'
        extension = 'csv'
        sourceDoc = CsvSceneList(sourcePath)

    elif sourcePath.endswith('_plotlist.csv'):
        suffix = '_plotlist'
        extension = 'csv'
        sourceDoc = CsvPlotList(sourcePath)

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
