"""Convert yWriter project to odt or csv and vice versa. 

Version @release

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import sys

from urllib.parse import unquote

from openoffice.yw_cnv_oo import YwCnvOO


def run(sourcePath, suffix, silentMode):
    converter = YwCnvOO(sourcePath, suffix, silentMode)

    if converter.success:
        delete_tempfile(sourcePath)
        return True

    else:
        return False


def delete_tempfile(filePath):
    """If an Office file exists, delete the temporary file."""

    if filePath.endswith('.html'):

        if os.path.isfile(filePath.replace('.html', '.odt')):

            try:
                os.remove(filePath)

            except:
                pass

    elif filePath.endswith('.csv'):

        if os.path.isfile(filePath.replace('.csv', '.ods')):

            try:
                os.remove(filePath)

            except:
                pass


if __name__ == '__main__':

    try:
        sourcePath = unquote(sys.argv[1].replace('file:///', ''))

    except:
        sourcePath = ''

    fileName, FileExtension = os.path.splitext(sourcePath)

    if not FileExtension in YwCnvOO.YW_EXTENSIONS:
        # Source file is not a yWriter project
        suffix = None

    else:
        # Source file is a yWriter project; suffix matters

        try:
            suffix = sys.argv[2]

        except:
            suffix = ''

    result = run(sourcePath, suffix, False)
