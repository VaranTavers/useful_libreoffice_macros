Sub ExportAllSheetsAsPDFs

' Export options. Now they are minimal. To adjust the PDF format accurately, 
' see https://wiki.openoffice.org/wiki/API/Tutorials/PDF_export

Dim ExportArgs(1) as new com.sun.star.beans.PropertyValue
Dim dataRange(0) as new com.sun.star.beans.PropertyValue

Dim oSheets As Variant, oSheet As Variant, oCursor As Variant
Dim sPath As String, sFileName As String
Dim i As Long
	GlobalScope.BasicLibraries.LoadLibrary("Tools")
	sFileName = ThisComponent.URL
	If sFileName = "" Then Exit Sub	' This spreadsheet wasn't saved, so it hasn't path and name '
    ' Folder to store PDF - same as source file
	sPath = ConvertFromURL(DirectoryNameoutofPath(sFileName,"/"))
	If Right(sPath,1) <> GetPathSeparator() Then sPath = sPath + GetPathSeparator()	
	ExportArgs(0).Name = "FilterName"
	ExportArgs(0).Value = "calc_pdf_Export"
	ExportArgs(1).Name = "FilterData"
	dataRange(0).Name = "Selection"
	oSheets = ThisComponent.getSheets()
	For i = 0 To oSheets.getCount()-1
		oSheet = oSheets.getByIndex(i)
		If oSheet.isVisible Then ' Skip hidden sheets '
            ' Export range with data only (skip first empty columns and rows)
			oCursor = oSheet.createCursor()
			oCursor.gotoStartOfUsedArea(False)
			oCursor.gotoEndOfUsedArea(True)
            ' Set range with data as param of filter export
			dataRange(0).Value = oCursor
			ExportArgs(1).Value = dataRange()
            ' Create filename by sheet name
			sFileName = sPath + oSheet.getName()+".pdf"
            ' Rem Export
			ThisComponent.StoreToURL(ConvertToURL(sFileName),ExportArgs())
		EndIf 
	Next i
End Sub