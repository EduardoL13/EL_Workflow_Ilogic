Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document

' Read Datahub / PropsSheet from User Defined Properties
Dim sFile As String = GetCustomProp(currentDoc, "Datahub")
Dim tabDS As String = GetCustomProp(currentDoc, "PropsSheet")

If sFile = "" Or tabDS = "" Then
	MsgBox("Datahub and/or PropsSheet are not defined." & vbCrLf & "Run the setup rule first.", MsgBoxStyle.Exclamation, "Missing properties")
	Return
End If

If Not ConfirmSettings(sFile, tabDS) Then Return

' Resolve column letters from the header names in row 1
Dim headerCols As Dictionary(Of String, String) = BuildHeaderMap(sFile, tabDS)

Dim requiredHeaders As String() = {"PART NO", "STOCK NUMBER", "DESCRIPTION", "MATERIAL INV", "FINISH", "BOM PROP"}
Dim missingHeaders As String = ""
For Each h As String In requiredHeaders
	If Not headerCols.ContainsKey(h) Then
		missingHeaders &= h & vbCrLf
	End If
Next

If missingHeaders <> "" Then
	MsgBox("The following headers were not found in row 1 of worksheet '" & tabDS & "':" & vbCrLf & vbCrLf & _
	       missingHeaders & vbCrLf & "Nothing was changed.", MsgBoxStyle.Exclamation, "Missing headers")
	Return
End If
	
Dim lastDataValue As Integer = findLastDataRow(sFile, tabDS)

setPropertiesAndMat(currentDoc, sFile, tabDS, lastDataValue, headerCols) 
currentDoc.Update2
MsgBox("Properties updated successfully")
	
End Sub	


Function findLastDataRow(file As String, tab As String)
	Dim cellVal As Object 
	Dim range As Integer = 1000
	Dim lastDataRow As Integer = 0
    For rowNum As Integer = 1 To range
		cellVal = GoExcel.CellValue(file, tab, "A" & rowNum)
	    If cellVal Is Nothing  Then 
			If (rowNum > range) Then Exit For
	    Else
			lastDataRow = rowNum + 1
		End If
	Next
	Return lastDataRow
End Function


Function BuildHeaderMap(file As String, tab As String) As Dictionary(Of String, String)
	' Reads row 1 and maps each header text (trimmed, upper case) to its column letter
	Dim map As New Dictionary(Of String, String)
	Dim maxCols As Integer = 50
	Dim headerVal As Object
	Dim key As String

	For colNum As Integer = 1 To maxCols
		Dim letter As String = ColumnLetter(colNum)
		headerVal = GoExcel.CellValue(file, tab, letter & "1")
		If headerVal IsNot Nothing Then
			key = headerVal.ToString().Trim().ToUpper()
			If key <> "" AndAlso Not map.ContainsKey(key) Then
				map.Add(key, letter)
			End If
		End If
	Next

	Return map
End Function


Function ColumnLetter(colIndex As Integer) As String
	' Converts a 1-based column index into an Excel column letter (1 = A, 27 = AA, ...)
	Dim result As String = ""
	Dim n As Integer = colIndex
	While n > 0
		Dim remainder As Integer = (n - 1) Mod 26
		result = Chr(65 + remainder) & result
		n = (n - 1) \ 26
	End While
	Return result
End Function


Sub setPropertiesAndMat(assembComp As AssemblyDocument, file As String, tab As String, lastValue As Integer, headerCols As Dictionary(Of String, String)) 

	' Column letters resolved from the headers in row 1
	Dim colPartNo As String = headerCols("PART NO")
	Dim colStockNumber As String = headerCols("STOCK NUMBER")
	Dim colDescription As String = headerCols("DESCRIPTION")
	Dim colBOMprop As String = headerCols("BOM PROP")
	Dim colMaterial As String = headerCols("MATERIAL INV")
	Dim colFinish As String = headerCols("FINISH")
	Dim colPartID As String = ""
	If headerCols.ContainsKey("PART ID") Then colPartID = headerCols("PART ID") ' Optional, currently not assigned
	
	Dim listkeyStrings As New List(Of String)
	Dim listIncompleteParts As New List(Of String)
	missingPropsInParts = ""

    For Each compOcc As ComponentOccurrence In assembComp.ComponentDefinition.Occurrences.AllLeafOccurrences
	
		If compOcc.Suppressed Then
			
		Else	
	        If  TypeOf compOcc.Definition.Document IsNot PartDocument Or compOcc.BOMStructure.Equals(kNormalBOMStructure) <> True Then
	            'MsgBox(compOcc.Name)
	        Else	
		        Dim occDoc As PartDocument = compOcc.Definition.Document ' Access to the occurrence's part document
	        	nameToCompare = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))
		   
		    	If listkeyStrings.Contains(nameToCompare) Then
				
		    	Else
		        	'MsgBox(nameToCheck)
			     	listkeyStrings.Add(nameToCompare)
		
	             	For rowCounter = 2 To lastValue
				
		                Dim nameOccDS As String = GoExcel.CellValue(file, tab, "A" & rowCounter) 
				       'Dim occDoc As PartDocument = compOcc.Definition.Document ' Access to the occurrence's part document
					 	Dim occProps As PropertySet = occDoc.PropertySets.Item("Design Tracking Properties")
				
					 	Dim oCurrentScope As MemberEditScopeEnum
					 	Dim oFactoryDoc As PartDocument
				
				
				    	If occDoc.ComponentDefinition.IsModelStateMember Then
					
					        oFactoryDoc = occDoc.ComponentDefinition.FactoryDocument
					     	oCurrentScope = oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope
		
				            If oCurrentScope = MemberEditScopeEnum.kEditActiveMember Then
					            oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = MemberEditScopeEnum.kEditAllMembers
				    	 	End If
				    	 	occDoc = oFactoryDoc
					 	End If			
					 
						If occDoc.PropertySets.PropertySetExists("Database Properties") Then
						 
					    Else
						    occDoc.PropertySets.Add("Database Properties")
						 	occDoc.PropertySets("Database Properties").Add("Part ID")
					 	End If
				
						If occDoc.PropertySets.PropertySetExists("Finish Properties") Then
						 
					    Else
						    occDoc.PropertySets.Add("Finish Properties")
						 	occDoc.PropertySets("Finish Properties").Add("Finish")
					 	End If				
					 
				
						'nameToCompare = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))
				
				
			        	'If compOcc.Name = nameOccDS Then
                 	 	If nameToCompare = nameOccDS Then 
				        	'MsgBox(nameToCompare)
                    	    'propsAssigner(nameToCompare & ".ipt", file, tab, rowCounter, colPartNo, colStockNumber, colDescription, colMaterial, colFinish)    'compOcc.Name
						Try
                        	iProperties.Expression(nameToCompare & ".ipt", "Project", "Part Number") = GoExcel.CellValue(file, tab, colPartNo & rowCounter)
							iProperties.Expression(nameToCompare & ".ipt", "Project", "Stock Number") = GoExcel.CellValue(file, tab, colStockNumber & rowCounter)
							iProperties.Expression(nameToCompare & ".ipt", "Project", "Description") = GoExcel.CellValue(file, tab, colDescription & rowCounter)
							iProperties.Expression(nameToCompare & ".ipt", "Custom", "Finish") = GoExcel.CellValue(file, tab, colFinish & rowCounter)
							iProperties.Expression(nameToCompare & ".ipt", "Custom", "BOM Prop") = GoExcel.CellValue(file, tab, colBOMprop & rowCounter)
							iProperties.Material(nameToCompare & ".ipt") = GoExcel.CellValue(file, tab, colMaterial & rowCounter)
						Catch
							'Msgbox(nameToCompare)
							listIncompleteParts.Add(nameToCompare)
		
						End Try

					
					
	                	If oFactoryDoc IsNot Nothing Then
					    	oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = oCurrentScope
						End If
	              
			         End If
			     Next 
	
		     End If
	    End If
	End If
    Next

    If listIncompleteParts.Count > 0 Then
        For Each name In listIncompleteParts
            missingPropsInParts &= name & vbCrLf
        Next
        MsgBox("Check properties in worksheet for: " & vbCrLf & vbCrLf & missingPropsInParts)
    End If
	
End Sub

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
