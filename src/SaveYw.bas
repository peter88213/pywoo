REM  *****  BASIC  *****

' SaveYw - OpenOffice macro for yWriter file export 
' 
' Copyright (c) 2020 Peter Triesberger
' For further information see https://github.com/peter88213/pywoo
' Published under the MIT License (https://opensource.org/licenses/mit-license.php)

Sub OdsYw
' Save an ods file as csv and call an external Python 3 script
' via batch file to convert the csv file to yWriter format.

	Dim document   As object
	Dim dispatcher As object
	Dim ods_path, csv_path As string
	
	On Error goto DontSave
	
	' ----------------------------------------------------------------------
	' Save document in csv format
	' ----------------------------------------------------------------------
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	ods_path = ThisComponent.getURL()
	ods_path = REPLACE(ods_path,".csv",".ods")
	csv_path = REPLACE(ods_path,".ods",".csv")
	' Use path and name of input file -- just change file extension
	
	Dim args1(2) As new com.sun.star.beans.PropertyValue
	args1(0).Name = "URL"
	args1(0).Value = csv_path
	args1(1).Name = "FilterName"
	args1(1).Value = "Text - txt - csv (StarCalc)"
	args1(2).Name = "FilterOptions"
	args1(2).Value = "124,34,76,1,,0,false,true,true"
	dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())
	
	' ----------------------------------------------------------------------
	' Save document in OpenDocument format
	' ----------------------------------------------------------------------
	args1(0).Value = ods_path
	args1(1).Value = "calc8"
	dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())
	
	' ----------------------------------------------------------------------
	' Call a batch script to convert to yw7
	' ----------------------------------------------------------------------
	oPackageInfoProvider = GetDefaultContext.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
	sPackageLocation = oPackageInfoProvider.getPackageLocation("org.peter88213.pywoo")

	Dim file_no As Integer
	Dim script_dir, batch_file, batch_cmd As String
		
	script_dir = ConvertFromURL(sPackageLocation)
	batch_file = script_dir + "\python\saveyw.bat"
	batch_cmd = script_dir + "\python\saveyw.pyw %1"

	file_no = FreeFile
	
	Open batch_file For Output As #file_no
	Print #file_no, batch_cmd
	Close #file_no
	
	batch_file = ConvertToURL(batch_file)
	
	shell(batch_file, 2, csv_path, false)

DontSave:

end Sub


Sub OdtYw
' Save an odt file as html and call an external Python 3 script
' via batch file to convert the html file to yWriter format.

	Dim document   As object
	Dim dispatcher As object
	Dim odt_path, html_path As string
	
	On Error goto DontSave
	
	' ----------------------------------------------------------------------
	' Save document in HTML format
	' ----------------------------------------------------------------------
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	odt_path = ThisComponent.getURL()
	odt_path = REPLACE(odt_path,".html",".odt")
	html_path = REPLACE(odt_path,".odt",".html")
	' Use path and name of input file -- just change file extension
	
	Dim args1(1) As new com.sun.star.beans.PropertyValue
	args1(0).Name = "URL"
	args1(0).Value = html_path
	args1(1).Name = "FilterName"
	args1(1).Value = "HTML (StarWriter)"
	dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())
	
	' ----------------------------------------------------------------------
	' Save document in OpenDocument format
	' ----------------------------------------------------------------------
	args1(0).Value = odt_path
	args1(1).Value = "writer8"
	dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())
	
	' ----------------------------------------------------------------------
	' Call a batch script to convert to yw7
	' ----------------------------------------------------------------------
	oPackageInfoProvider = GetDefaultContext.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
	sPackageLocation = oPackageInfoProvider.getPackageLocation("org.peter88213.pywoo")

	Dim file_no As Integer
	Dim script_dir, batch_file, batch_cmd As String
		
	script_dir = ConvertFromURL(sPackageLocation)
	batch_file = script_dir + "\python\saveyw.bat"
	batch_cmd = script_dir + "\python\saveyw.pyw %1"

	file_no = FreeFile
	
	Open batch_file For Output As #file_no
	Print #file_no, "cd " + script_dir
	Print #file_no, batch_cmd
	Close #file_no
	
	batch_file = ConvertToURL(batch_file)
	
	shell(batch_file, 2, html_path, false)
	

DontSave:

end Sub
