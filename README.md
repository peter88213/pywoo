# The pywoo extension for OpenOffice: Import and export yWriter 6/7 projects

![Screenshot: Menu in LibreOffice](https://raw.githubusercontent.com/peter88213/pywoo/master/docs/Screenshots/lo_menu.png)

## Features (a Python 3 installation is required)

* Generate a "standard manuscript" formatted  __ODF text document (ODT)__  from an yWriter 6/7 project.

* Load yWriter 6/7 chapters and scenes into an OpenDocument file with chapter and scene markers for  __proof reading__  and writing back. 

* Generate a  __character list__  that can be edited in Office Calc and written back to yWriter format.

* Generate a  __location list__  that can be edited in Office Calc and written back to yWriter format.

* Generate an  __item list__  that can be edited in Office Calc and written back to yWriter format.

* Generate an OpenDocument text file containing navigable  __cross references__ , such as scenes per character, characters per tag, etc.

* Generate a new yWriter 7 project from a  __work in progress__  or an  __outline__ .

You can find more information in the [help text](https://raw.githubusercontent.com/peter88213/pywoo/master/oxt/help/help.html).

## Download and install

[Download page](https://github.com/peter88213/pywoo/releases/latest)

* Download the `OXT` file.

* Install it using the OpenOffice extension manager.

* After installation (and Office restart) you find a new "yWriter Import/Export" submenu in the "Files" menu.

* If no additional "yWriter Import/Export" submenu shows up in the "Files" menu, please look at the "Tools" > "Extensions" menu.

## Requirements

* Windows.
* OpenOffice 3.0 or more recent. 
* Python 3.4 or more recent will work. However, Python 3.7 or above is highly recommended.
* Java Runtime Environment (OpenOffice might need it for macro execution).

## Credits

[yWriter](http://spacejock.com/yWriter7.html) by Simon Haynes.

[OpenOffice Extension Compiler](https://wiki.openoffice.org/wiki/Extensions_Packager#Extension_Compiler) by Bernard Marcelly.

## License

This extension is distributed under the [MIT License](http://www.opensource.org/licenses/mit-license.php).
