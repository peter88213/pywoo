"""Convert yw7 to odt/ods, or html/csv to yw7. 

Version @release
Requires Python 3.6+
Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/pywoo
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import sys
import platform
from pywoolib.converter import Converter
from pywriter.ui.ui_mb import UiMb

YW_EXTENSIONS = ['.yw7']


def run(sourcePath, suffix=None):
    converter = Converter()
    converter.ui = UiMb(f'yWriter import/export (Python version {platform.python_version()})')
    kwargs = {'suffix': suffix}
    converter.run(sourcePath, **kwargs)


if __name__ == '__main__':
    # Enable this for debugging unhandled exceptions:
    # sys.stderr = open(os.path.join(os.getenv('TEMP'), f'stderr-{os.path.basename(sys.argv[0])}.txt'), 'w')
    try:
        sourcePath = sys.argv[1]
    except:
        sourcePath = ''
    fileName, FileExtension = os.path.splitext(sourcePath)
    if not FileExtension in YW_EXTENSIONS:
        # Source file is not a yWriter project.
        suffix = None
    else:
        # Source file is a yWriter project; suffix matters.
        try:
            suffix = sys.argv[2]
        except:
            suffix = ''
    run(sourcePath, suffix)
