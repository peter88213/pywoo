REM  *****  BASIC  *****

' SaveYw - OpenOffice macro for yWriter file export 
' 
' Copyright (c) 2020 Peter Triesberger
' For further information see https://github.com/peter88213/pywoo
' Published under the MIT License (https://opensource.org/licenses/mit-license.php)

Sub OdsYw
' Save an ods file as csv and call an external Python 3 script
' via batch file to convert the csv file to yWriter format.

	dim document   as object
	dim dispatcher as object
	dim DocPath as string
	
	On Error goto DontSave
	
	' ----------------------------------------------------------------------
	' Save document in csv format
	' ----------------------------------------------------------------------
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	OdsPath = ThisComponent.getURL()
	OdsPath = REPLACE(OdsPath,".csv",".ods")
	CsvPath = REPLACE(OdsPath,".ods",".csv")
	' Use path and name of input file -- just change file extension
	
	dim args1(2) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "URL"
	args1(0).Value = CsvPath
	args1(1).Name = "FilterName"
	args1(1).Value = "Text - txt - csv (StarCalc)"
	args1(2).Name = "FilterOptions"
	args1(2).Value = "124,34,76,1,,0,false,true,true"
	dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())
	
	' ----------------------------------------------------------------------
	' Save document in OpenDocument format
	' ----------------------------------------------------------------------
	args1(0).Value = OdsPath
	args1(1).Value = "calc8"
	dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())
	
	' ----------------------------------------------------------------------
	' Call a batch script to convert to yw7
	' ----------------------------------------------------------------------
	oPackageInfoProvider = GetDefaultContext.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
	sPackageLocation = oPackageInfoProvider.getPackageLocation("org.peter88213.pywoo")

	Dim FileNo As Integer
	Dim BatchDir, Filename, Batchcommand As String
		
	BatchDir = ConvertFromURL(sPackageLocation)
	Filename = BatchDir + "\python\saveyw.bat"
	Batchcommand = BatchDir + "\python\saveyw.pyw %1"

	FileNo = FreeFile               		' Establish free file handle
	
	Open Filename For Output As #FileNo     ' Open file (writing mode)
	Print #FileNo, Batchcommand      		' Save line
	Close #FileNo                  			' Close file
	
	
	Dim ScriptPath as String
	ScriptPath = ConvertToURL(Filename)
	
	shell(ScriptPath, 2, CsvPath, false)

DontSave:

end Sub


Sub OdtYw
' Save an odt file as html and call an external Python 3 script
' via batch file to convert the html file to yWriter format.

	dim document   as object
	dim dispatcher as object
	dim DocPath as string
	
	On Error goto DontSave
	
	' ----------------------------------------------------------------------
	' Save document in HTML format
	' ----------------------------------------------------------------------
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	OdtPath = ThisComponent.getURL()
	OdtPath = REPLACE(OdtPath,".html",".odt")
	HtmlPath = REPLACE(OdtPath,".odt",".html")
	' Use path and name of input file -- just change file extension
	
	dim args1(1) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "URL"
	args1(0).Value = HtmlPath
	args1(1).Name = "FilterName"
	args1(1).Value = "HTML (StarWriter)"
	dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())
	
	' ----------------------------------------------------------------------
	' Save document in OpenDocument format
	' ----------------------------------------------------------------------
	args1(0).Value = OdtPath
	args1(1).Value = "writer8"
	dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())
	
	' ----------------------------------------------------------------------
	' Call a batch script to convert to yw7
	' ----------------------------------------------------------------------
	oPackageInfoProvider = GetDefaultContext.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
	sPackageLocation = oPackageInfoProvider.getPackageLocation("org.peter88213.pywoo")

	Dim FileNo As Integer
	Dim BatchDir, Filename, Batchcommand As String
		
	BatchDir = ConvertFromURL(sPackageLocation)
	Filename = BatchDir + "\python\saveyw.bat"
	Batchcommand = BatchDir + "\python\saveyw.pyw %1"

	FileNo = FreeFile               		' Establish free file handle
	
	Open Filename For Output As #FileNo     ' Open file (writing mode)
	Print #FileNo, Batchcommand      		' Save line
	Close #FileNo                  			' Close file
	
	
	Dim ScriptPath as String
	ScriptPath = ConvertToURL(Filename)
	
	shell(ScriptPath, 2, HtmlPath, false)
	

DontSave:

end Sub
