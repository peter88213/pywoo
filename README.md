# The pywoo tools for OpenOffice: Import and export yWriter 6/7 projects with Python

## Features

### yWriter5 style proof reading

#### Load yWriter 6/7 chapters and scenes into an OpenDocument file with chapter and scene markers. 

* The proof read document is placed in the same folder as the yWriter project.
* Document's filename: `<yW project name>_proof.odt`.

#### Write back the proofread scenes to the yWriter 6/7 project file.

* The yWriter 6/7 project to rewrite must exist in the document folder.
* If both yw6 and yw7 project files exist, yw7 is rewritten. 

### Import from yw6/7 project 

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
 
 
## Requirements

* Windows
* Python 3.4 or more recent.
* OpenOffice 3.0 or more recent.
* Java Runtime Environment (OpenOffice needs it for macro execution).
* yWriter 7. 

The  _yw-cnv_  extension is distributed under the [MIT License](http://www.opensource.org/licenses/mit-license.php).
