""" Build python script for the OpenOffice "convert yWriter" script.
        
In order to distribute single scripts without dependencies, 
this script "inlines" all modules imported from the pywriter package.

For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import inliner

SRC = '../src/standalone/'
BUILD = '../build/'
SOURCE_FILE = SRC + 'pywoo_.pyw'
TARGET_FILE = BUILD + 'pywoo.pyw'


def main():
    # os.chdir(SRC)

    try:
        os.remove(TARGET_FILE)

    except:
        pass

    inliner.run(SOURCE_FILE,
                TARGET_FILE, 'openoffice', '../src/')
    inliner.run(TARGET_FILE,
                TARGET_FILE, 'pywriter', '../../PyWriter/src/')
    print('Done.')


if __name__ == '__main__':
    main()
