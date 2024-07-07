Sub CalcToLaTeXTable()
    Dim oSheet As Object
    Dim oTable As Object
    Dim oCell As Object
    Dim sLaTeX As String
    Dim iRow As Integer
    Dim iCol As Integer
    
        ' Get the current controller
    Dim oController As Object
    oController = ThisComponent.CurrentController
    
    ' Get the current selection
    Dim oSelection As Object
    oSelection = oController.getSelection()
    
    ' Check if the selection is a cell range
    If oSelection.supportsService("com.sun.star.sheet.SheetCellRange") Then
        ' Use the selected range as the table range
        oTable = oSelection
    Else
        ' Set the default table range (modify as needed)
        oTable = ThisComponent.Sheets(0).getCellRangeByName("A1:F10")
    End If
    
    ' Initialize LaTeX string
    sLaTeX = "\begin{tabular}{"
    
    ' Add column alignment (assuming all columns are centered)
    For iCol = 0 To oTable.getColumns().getCount() - 1
        sLaTeX = sLaTeX & "c"
    Next iCol
    
    sLaTeX = sLaTeX & "}" & Chr(10)  
    
    ' Loop through each row
    For iRow = 0 To oTable.getRows().getCount() - 1
        sLaTeX = sLaTeX & "    "
        
        ' Loop through each column
        For iCol = 0 To oTable.getColumns().getCount() - 1
            ' Get cell value
            oCell = oTable.getCellByPosition(iCol, iRow)
            Dim cellType As Integer
            cellType = oCell.getType()
           
           If cellType = 1 Then
             sLaTeX = sLaTeX & "$ " &  Replace(oCell.getString(), ",", ".") & " $ & "
           Else
             sLaTeX = sLaTeX & Replace(oCell.getString(), "_", "\_") & " & "
           End If
            
        Next iCol
        
        ' Remove the trailing "& "
        sLaTeX = Left(sLaTeX, Len(sLaTeX) - 2) & " \\ " & Chr(10) 

    Next iRow
    
    ' Add the table footer
    sLaTeX = sLaTeX & "\end{tabular}"
    
	CreateTextBox sLaTeX
    
    ' Inform the user
    MsgBox "LaTeX code has been copied to the clipboard.", 64, "LaTeX Table"
End Sub

Sub CreateTextBox(sText As String)
    ' Used to work with copying to the clipboard, but EasyMacro did not work.
    'app = createUnoService("net.elmau.zaz.EasyMacro")
    'app.set_clipboard(sText)
    
    Dim oSheet as Object
    Dim oCell as Object
    oSheet = ThisComponent.CurrentController.getActiveSheet()
    'oCell = oSheet.getCellRangeByName("A1")
    'oCell.setString(sText)

    Dim oPosition As New com.sun.star.awt.Point
    oPosition.X = 1000
    oPosition.Y = 1000

    Dim oSize As New com.sun.star.awt.Size
    oSize.Width = 10000
    oSize.Height = 5000

    Dim oTextShape As Object
    oTextShape = ThisComponent.createInstance("com.sun.star.drawing.TextShape")

    oTextShape.setPosition(oPosition)
    oTextShape.setSize(oSize)
    oTextShape.setPropertyValue("FillStyle", "SOLID")
    oTextShape.Visible = 1

    ' Give it a name so you can find it again when you want to delete it
    oTextShape.setPropertyValue("Name", "Thingy")

    Dim oDrawPage As Object
    oDrawPage = oSheet.getDrawPage()

    oDrawPage.add(oTextShape)

    ' Set the string of the text shape AFTER adding it to the
    ' draw page, otherwise the text will not be set.
    oTextShape.setString(sText)
End Sub
