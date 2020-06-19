"""Import an outline. 

Convert html with chapter and scene headings/descriptions to yWriter format.

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""

import sys
import os

from pywriter.html.html_outline import HtmlOutline
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


def run(sourcePath):

    sourcePath = sourcePath.replace('file:///', '').replace('%20', ' ')
    fileName, FileExtension = os.path.splitext(sourcePath)

    if FileExtension == '.html':
        document = HtmlOutline('')
        extension = 'html'

    else:
        sys.exit('ERROR: File type is not supported.')

    converter = Converter(sourcePath, document,
                          extension, False, '')


if __name__ == '__main__':
    try:
        sourcePath = sys.argv[1]
    except:
        sourcePath = ''

    run(sourcePath)
