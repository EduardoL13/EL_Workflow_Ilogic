Sub Main ()

	Dim currentDoc As AssemblyDocument = ThisDoc.Document

	' Read Datahub / PropsSheet from User Defined Properties
	Dim sFile As String = GetCustomProp(currentDoc, "Datahub")
	Dim tab As String = GetCustomProp(currentDoc, "PropsSheet")

	If sFile = "" Or tab = "" Then
		MsgBox("Datahub and/or PropsSheet are not defined." & vbCrLf & "Run the setup rule first.", MsgBoxStyle.Exclamation, "Missing properties")
		Return
	End If

	If Not ConfirmSettings(sFile, tab) Then Return

	Dim count As Integer = 0
	Dim leafOccurrences As ComponentOccurrencesEnumerator = currentDoc.ComponentDefinition.Occurrences.AllLeafOccurrences

	Dim listkeyStrings As New List(Of String)

	Dim rowCounter = 2
	
	Dim excelNames As New List(Of String)

	If GoExcel.CellValue(sFile, tab, "A2") <> "" Then

        Dim lastRow As Integer = findLastDataRow(sFile, tab)

    	For i As Integer = 2 To lastRow - 1
        	excelNames.Add(GoExcel.CellValue(sFile, tab, "A" & i).ToString())
    	Next

    	rowCounter = lastRow

	Else

    	rowCounter = 2

	End If
	
	For Each compOccurrence As ComponentOccurrence In leafOccurrences
	
    	If compOccurrence.BOMStructure.ToString = "kNormalBOMStructure" And compOccurrence.Suppressed = False Then	
			Dim occDoc As PartDocument = compOccurrence.Definition.Document
			nameToCheck = compOccurrence.Name.Substring(0, compOccurrence.Name.LastIndexOf(":"))
	
		If listkeyStrings.Contains(nameToCheck) Then

		Else

    	listkeyStrings.Add(nameToCheck)

    	Dim partName As String = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))

    	' Only write if it does not already exist in Excel
    	If Not excelNames.Contains(partName) Then

        	If rowCounter = 2 Then
            	GoExcel.CellValue(sFile, tab, "A" & rowCounter) = partName
        	Else
            	GoExcel.CellValue("A" & rowCounter) = partName
        	End If

        	excelNames.Add(partName)
        	rowCounter = rowCounter + 1

    	End If

End If
	End If
Next

GoExcel.Save

End Sub

Function findLastDataRow(file As String, tab As String)

    Dim cellVal As Object
    Dim range As Integer = 1000
    Dim lastDataRow As Integer = 0

    For rowNum As Integer = 1 To range

        cellVal = GoExcel.CellValue(file, tab, "A" & rowNum)

        If cellVal Is Nothing Then
            If rowNum > range Then Exit For
        Else
            lastDataRow = rowNum + 1
        End If

    Next

    Return lastDataRow

End Function

Function GetCustomProp(doc As Document, propName As String) As String
    Dim userProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
    For Each prop As Inventor.Property In userProps
        If prop.Name = propName Then
            If prop.Value Is Nothing Then Return ""
            Return prop.Value.ToString()
        End If
    Next
    Return ""
End Function

Function ConfirmSettings(datahub As String, sheet As String) As Boolean
    Dim result As MsgBoxResult = MsgBox("Current settings:" & vbCrLf & vbCrLf & _
                                        "Datahub = " & datahub & vbCrLf & _
                                        "PropsSheet = " & sheet & vbCrLf & vbCrLf & _
                                        "Do you want to continue with these values?", _
                                        MsgBoxStyle.OkCancel + MsgBoxStyle.Question, "Confirm Datahub settings")
    Return (result = MsgBoxResult.Ok)
End Function
