REM  *****  BASIC  *****

' OpenYw - OpenOffice macro for yWriter file import
'
' Copyright (c) 2020 Peter Triesberger
' For further information see https://github.com/peter88213/pywoo
' Published under the MIT License (https://opensource.org/licenses/mit-license.php)

Sub import_yw
	open_yw("", ".odt")
End Sub

Sub proof_yw
    open_yw("_proof", ".odt")
End Sub

Sub get_manuscript
    open_yw("_manuscript", ".odt")
End Sub

Sub get_scenedesc
    open_yw("_scenes", ".odt")
End Sub

Sub get_chapterdesc
    open_yw("_chapters", ".odt")
End Sub

Sub get_partdesc
    open_yw("_parts", ".odt")
End Sub

Sub get_chardesc
    open_yw("_characters", ".odt")
End Sub

Sub get_locdesc
    open_yw("_locations", ".odt")
End Sub

Sub get_itemdesc
    open_yw("_items", ".odt")
End Sub

Sub get_scenelist
    open_yw("_scenelist", ".csv")
End Sub

Sub get_charlist
    open_yw("_charlist", ".csv")
End Sub

Sub get_loclist
    open_yw("_loclist", ".csv")
End Sub

Sub get_itemlist
    open_yw("_itemlist", ".csv")
End Sub

Sub get_plotlist
    open_yw("_plotlist", ".csv")
End Sub


Sub open_yw(suffix As String, new_ext As String)
    ' ----------------------------------------------------------------------
    ' Set last opened yWriter project As default (if existing).
    ' ----------------------------------------------------------------------
	oPackageInfoProvider = GetDefaultContext.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
	sPackageLocation = oPackageInfoProvider.getPackageLocation("org.peter88213.pywoo")

    Dim file_no As Integer
	Dim script_dir, ini_file, yw_last_open, default_file As String

	script_dir = ConvertFromURL(sPackageLocation)
	ini_file = script_dir + "\python\pywoo.ini"

    If FileExists(ini_file) Then
        file_no = FreeFile
        Open ini_file For Input As #file_no
        Line Input #file_no, yw_last_open
        Close #file_no

        If FileExists(yw_last_open) Then
            default_file = yw_last_open
        End If

    End If

    ' ----------------------------------------------------------------------
    ' Ask for yWriter 6 or 7 project to open:
    ' ----------------------------------------------------------------------
    Dim file_dialog As Object
    Dim file_path As String

    GlobalScope.BasicLibraries.LoadLibrary("Tools")
    file_dialog = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")


    Dim filterNames(2) As String

    filterNames(0) = "*.yw7"
    filterNames(1) = "*.yw6"

    AddFiltersToDialog(FilterNames(), file_dialog)
    file_dialog.SetDisplayDirectory(default_file)
    open_status = file_dialog.Execute()

    If open_status = 1 Then
        file_path = file_dialog.Files(0)
	    ' ----------------------------------------------------------------------
	    ' Store selected yWriter project as "last opened".
	    ' ----------------------------------------------------------------------
	    file_no = FreeFile
		Open ini_file For Output As #file_no
		Print #file_no, file_path
		Close #file_no
    End If

    file_dialog.Dispose()



    ' ----------------------------------------------------------------------
    ' Check whether import file is already open in OpenOffice.
    ' ----------------------------------------------------------------------
    Dim lock_file, source_dir, source_name As String
    file_path = ConvertFromURL(file_path)
    source_dir = DirectoryNameoutofPath(file_path,"\")
    source_name = FileNameoutofPath(file_path,"\")
    source_name = REPLACE(source_name, ".yw7", suffix + new_ext)
    source_name = REPLACE(source_name, ".yw6", suffix + new_ext)
	
	lock_file = source_dir  + "\.~lock." + source_name + "#"
    
    If FileExists(lock_file) Then
    	msgbox("Please close '" + source_name + "' first.")
    	Exit Sub
    	
    End If

    ' ----------------------------------------------------------------------
    ' Open yWriter project and convert data.
    ' ----------------------------------------------------------------------
    Dim batch_file As String



    batch_file = script_dir + "\python\openyw.bat"
	batch_cmd = script_dir + "\python\openyw.pyw %1 %2"
    file_no = FreeFile
	Open batch_file For Output As #file_no
	Print #file_no, "cd " + script_dir
	Print #file_no, batch_cmd	
	Close #file_no

	batch_file = ConvertToURL(batch_file)

	shell(batch_file, 2, file_path + " " + suffix, false)

End Sub




