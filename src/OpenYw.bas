REM  *****  BASIC  *****

' OpenYw - OpenOffice macro for yWriter file import 
' 
' Copyright (c) 2020 Peter Triesberger
' For further information see https://github.com/peter88213/pywoo
' Published under the MIT License (https://opensource.org/licenses/mit-license.php)

sub import_yw
open_yw7("", ".odt")
end sub


sub proof_yw
open_yw7("_proof", ".odt")
end sub



sub get_manuscript
open_yw7("_manuscript", ".odt")
end sub



sub get_scenedesc
open_yw7("_scenes", ".odt")
end sub


sub get_chapterdesc
open_yw7("_chapters", ".odt")
end sub


sub get_partdesc
open_yw7("_parts", ".odt")
end sub


sub get_chardesc
open_yw7("_characters", ".odt")
end sub


sub get_locdesc
open_yw7("_locations", ".odt")
end sub


sub get_itemdesc
open_yw7("_items", ".odt")
end sub


sub get_scenelist
open_yw7("_scenelist", ".csv")
end sub


sub get_charlist
open_yw7("_charlist", ".csv")
end sub


sub get_loclist
open_yw7("_loclist", ".csv")
end sub


sub get_itemlist
open_yw7("_itemlist", ".csv")
end sub


sub get_plotlist
open_yw7("_plotlist", ".csv")
end sub


sub open_yw(suffix As String, newExt As String)

' Set last opened yWriter project as default (if existing).


' Ask for yWriter 6 or 7 project to open:


' Store selected yWriter project as "last opened".


' Check if import file is already open in OpenOffice.


' Open yWriter project and convert data.

end sub




