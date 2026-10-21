Sub Main () 
	
	Dim invDoc As PartDocument = ThisDoc.Document ' Documento activo

	Dim inventorParamList As UserParameters = invDoc.ComponentDefinition.Parameters.UserParameters
	Dim noPDS As Integer = inventorParamList.Count
	Dim noListPDS As Integer = noPDS - 1

'-----------------INPUT----------------------

    Dim writeParamsList(noListPDS) As Object

'---------------------------------------------

    Dim i As Integer

    For i = 0 To writeParamsList.Length - 1
        writeParamsList(i) = inventorParamList.Item(i + 1)
    Next


' Comms with data hub excel

	Dim file As String
    Dim tab As String = InputBox("Enter the excel worksheet name", "Worksheet to Export")

    If tab = "" Then
		Exit Sub
	End If

    Try
		
        file = invDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
	
    Catch

	    MsgBox("Communications with Data Hub are not set. Please try again after setting them")
	    Exit Sub

    End Try


	'Desired row and column input in given worksheet
	colParameterName = "A"   ' Column in which parameter names are listed
	colParameterValue = "B"  ' Column in which parameter values are listed
	colParameterUnits = "C"  ' Column in which parameter units are listed
	colParameterSource = "D" ' Column in which source file is listed
	

	'Row Counter Definition
	Dim RowCounter As Integer 
	
	If GoExcel.CellValue(file, tab, colParameterName & 2) <> "" Then
	    RowCounter = findLastDataRow(file, tab)
	Else
		RowCounter = 2
	End If
	
	
	For Each param As UserParameter In writeParamsList

        If param.Comment.Contains("Profile Parameter") Then

            Dim targetRow As Integer

            'Buscar si el parámetro ya existe
            targetRow = findParameterRow(file, tab, param.Name)

            'Si no existe, escribir al final
            If targetRow = 0 Then
                targetRow = RowCounter
                RowCounter = RowCounter + 1
            End If

            cf = unitsEval(param.Units)

            GoExcel.CellValue(file, tab, colParameterName & targetRow) = param.Name
            GoExcel.CellValue(file, tab, colParameterValue & targetRow) = param.Value * cf
            GoExcel.CellValue(file, tab, colParameterUnits & targetRow) = param.Units
            GoExcel.CellValue(file, tab, colParameterSource & targetRow) = invDoc.DisplayName

        End If

    Next

    GoExcel.Save

End Sub


Function unitsEval(units As String)

    If units = "in" Then
	    cf = 1 / 2.54
	Else If units = "ul" Then
		cf = 1
	Else
		cf = 10
	End If

	Return cf

End Function


Function findLastDataRow(file As String, tab As String)

	Dim cellVal As Object 
	Dim range As Integer = 1000
	Dim lastDataRow As Integer = 0

    For rowNum As Integer = 1 To range

		cellVal = GoExcel.CellValue(file, tab, "A" & rowNum)

	    If cellVal Is Nothing Then

			If (rowNum > range) Then Exit For

	    Else

			lastDataRow = rowNum + 1

		End If

	Next

	Return lastDataRow

End Function


Function findParameterRow(file As String, tab As String, parameterName As String) As Integer

    Dim cellVal As Object
    Dim range As Integer = 1000

    For rowNum As Integer = 2 To range

        cellVal = GoExcel.CellValue(file, tab, "A" & rowNum)

        If cellVal Is Nothing Then
            Exit For
        End If

        If cellVal.ToString = parameterName Then
            Return rowNum
        End If

    Next

    Return 0

End Function
