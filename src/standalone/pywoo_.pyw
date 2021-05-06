"""Convert yWriter project to odt or ods. 

Version @release

Copyright (c) 2021 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import sys
from urllib.parse import unquote

from openoffice.cnv_button import CnvButton

from pywriter.converter.export_file_factory import ExportFileFactory
from pywriter.ui.ui_tk_open import UiTkOpen


def run(sourcePath, suffix=None):
    ui = UiTkOpen('yWriter import/export')
    converter = CnvButton()
    converter.ui = ui
    converter.fileFactory = ExportFileFactory()
    converter.run(sourcePath, suffix)
    ui.start()


if __name__ == '__main__':

    try:
        sourcePath = unquote(sys.argv[1].replace('file:///', ''))

    except:
        sourcePath = ''

    fileName, FileExtension = os.path.splitext(sourcePath)

    if not FileExtension in CnvButton.YW_EXTENSIONS:
        # Source file is not a yWriter project
        suffix = None

    else:
        # Source file is a yWriter project; suffix matters

        try:
            suffix = sys.argv[2]

        except:
            suffix = ''

    run(sourcePath, suffix)
