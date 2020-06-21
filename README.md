# The pywoo tools for OpenOffice: Import and export yWriter 6/7 projects with Python

## Features

### yWriter5 style proof reading

#### Load yWriter 6/7 chapters and scenes into an OpenDocument file with chapter and scene markers. 

* Please consider the following conventions:
    * Text markup: Bold and italics are supported. Other highlighting such as underline and strikethrough are lost.
    * All chapters and scenes will be exported, whether "used" or "unused".
    
* Back up your yWriter project and close yWriter before.
* The proof read document is placed in the same folder as the yWriter project.
* Document's filename: `<yW project name>_proof.odt`.
* The document contains chapter `[ChID:x]` and scene `[ScID:y]` markers according to yWriter 5 standard.  __Do not touch lines containing the markers__  if you want to be able to reimport the document into yWriter.

#### Write back the proofread scenes to the yWriter 6/7 project file.

* The yWriter 6/7 project to rewrite must exist in the same folder as the document.
* If both yw6 and yw7 project files exist, yw7 is rewritten. 

### Export to odt from yw6/7 project 

Generate a "standard manuscript" formatted OpenDocument textfile from an yWriter 6/7 project.

* The document is placed in the same folder as the yWriter project.
* Document's filename: `<yW project name>.odt`.


### Create a new yw7 project 

Generate a new yWriter 7 project from a work in progress or an outline.

* The new yWriter project is placed in the same folder as the document.
* yWriter project's filename: `<document name>.yw7`.
* Existing yWriter 7 projects will not be overwritten.


#### Formatting a work in progress

A work in progress has no third level heading.

* _Heading 1_  -->  New chapter title (beginning a new section).
* _Heading 2_  -->  New chapter title.
* `* * *`  -->  Scene divider (not needed for the first scenes in a chapter).
* All other text is considered to be scene content.

#### Formatting an outline

An outline has at least one third level heading.

* _Heading 1_  -->  New chapter title (beginning a new section).
* _Heading 2_  -->  New chapter title.
* _Heading 3_  -->  New scene title.
* All other text is considered to be chapter/scene description.


## Download and install

[Download page](https://github.com/peter88213/pywoo/releases/latest)

1. Download `PyWriter_OO_<version number>.zip` . 

2. Unzip `PyWriter_OO_<version number>.zip` within your user profile.

3. Move into the `PyWriter_OO_<version number>` folder and run `Install.bat` (double click). This will copy all needed files to the right places, install an OpenOffice extension. You may be asked for approval to modify the Windows registry. Please accept in order to install Explorer context menu entries for yWriter7 files.

## How to use

### Proof reading

1. Write your novel with yWriter. Please consider the conventions descripted above. Backup entire project and close yWriter.

2. Move into your yWriter project folder, and right-click your .yw6 or .yw7 project file. In the context menu, select `Proof read with OpenOffice`. 
   
3. If everything goes well, you find an OpenDocument file named `<your yWriter project>_proof.odt`. Open it (double click) for proof reading. The proof reading document contains Chapter `[ChID:x]` and scene `[ScID:y]` markers according to yWriter 5 standard.  __Do not touch lines containing the markers__  if you want to be able to reimport the document into yWriter. 

4. In order to write back the proofread scenes to the yWriter project, select the menu item `File > Export to yWriter`.

### Exporting from yw6/7 project 

1. Write your novel with yWriter. Backup entire project and close yWriter.

2. Move into your yWriter project folder, and right-click your .yw6 or .yw7 project file. In the context menu, select `Export to OpenOffice`. 
   
3. If everything goes well, you find an OpenDocument file named `<your yWriter project>.odt`. 


### Creating a new yw7 project 

1. Open your work in progress or outline in OpenOffice writer. Please consider the conventions descripted above. 

2. Select the menu item `File > Export to yWriter`.

3. Close OpenOffice Writer. If everything goes well, you find a new yWriter project named `<your document>.yw7`.

## Requirements

* Windows
* Python 3.4 or more recent.
* OpenOffice 3.0 or more recent.
* Java Runtime Environment (OpenOffice needs it for macro execution).
* yWriter 7. 

The  _yw-cnv_  extension is distributed under the [MIT License](http://www.opensource.org/licenses/mit-license.php).
