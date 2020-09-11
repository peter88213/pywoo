""" Build python script for the OpenOffice "yWriter import7export" script.
        
In order to distributea single script without dependencies, 
this script "inlines" all modules imported from the pywriter package.

For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import inliner

SRC = '../src/'
BUILD = '../test/'


def main():
    os.chdir(SRC)
    inliner.run('pywoo_.pyw', BUILD + 'pywoo.pyw', 'pywriter')
    print('Done.')


if __name__ == '__main__':
    main()
