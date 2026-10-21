Sub Main()

    Dim currentDoc As AssemblyDocument = ThisDoc.Document
    Dim leafOccurrences As ComponentOccurrencesEnumerator = currentDoc.ComponentDefinition.Occurrences.AllLeafOccurrences

    Dim oDlg As New System.Windows.Forms.FolderBrowserDialog
    Dim dlgResult As New System.Windows.Forms.DialogResult

    With oDlg
        .ShowNewFolderButton = True
        .InitialDirectory = "C:\Temp"
        If .ShowDialog = dlgResult.Cancel Then
            Exit Sub
        End If
    End With

    Dim targetLocation As String = oDlg.SelectedPath
    Dim listkeyStrings As New List(Of String)

    For Each compOccurrence As ComponentOccurrence In leafOccurrences
        If compOccurrence.Suppressed Then

        Else
            If TypeOf compOccurrence.Definition.Document IsNot PartDocument Then

            Else

                Dim occDoc As PartDocument = compOccurrence.Definition.Document

                ' Check for sheet metal
                If occDoc.ComponentDefinition.Type = kSheetMetalComponentDefinitionObject Then

                    ' Check if part is a replica of an existing one
                    Dim nameToCheck As String = compOccurrence.Name.Substring(0, compOccurrence.Name.LastIndexOf(":"))

                    If listkeyStrings.Contains(nameToCheck) Then

                    Else
                        listkeyStrings.Add(nameToCheck)

                        Dim smCompDef As SheetMetalComponentDefinition = occDoc.ComponentDefinition

                        ' Verifies if there is a flat pattern and creates one if there is not
                        If Not smCompDef.HasFlatPattern Then
                            smCompDef.Unfold()
                        End If

                        ' Gets flat pattern
                        Dim flatPattern As FlatPattern = smCompDef.FlatPattern

                        ' Creates or modifies the custom property "NameCutFile" using an expression
                        Dim resolvedName As String = ""
                        Try
                            resolvedName = SetCustomPropExpression(occDoc, "NameCutFile", "=<Stock Number>_<Stock Material>_<pThk>in_<Cost Center>") 'AQUí SE CAMBIA EL NOMBRE DE OUTPUT DEL FILE EN LA CARPETA
                        Catch ex As Exception
                            MsgBox("Could not set NameCutFile in part " & nameToCheck & ": " & ex.Message)
                            Exit For
                        End Try

                        ' Validates that the expression was resolved
                        If resolvedName = "" OrElse resolvedName.Contains("<") OrElse resolvedName.StartsWith("=") Then
                            MsgBox("NameCutFile could not be resolved in part " & nameToCheck & ". Resulting value: '" & resolvedName & "'. Check that Stock Number, Stock Material and pThk exist in the part.")
                            Exit For
                        End If

                        ' Removes characters that are invalid in file names
                        Dim safeName As String = resolvedName
                        For Each invalidChar As Char In System.IO.Path.GetInvalidFileNameChars()
                            safeName = safeName.Replace(invalidChar, "_"c)
                        Next

                        Dim fileName As String = System.IO.Path.Combine(targetLocation, safeName & ".dxf")

                        ' Export to DXF
                        Try
                            flatPattern.DataIO.WriteDataToFile("FLAT PATTERN DXF?AcadVersion=2018", fileName)
                        Catch
                            MsgBox("Invalid file name for " & nameToCheck & " cut file. Make sure that format is correct and that the name does not contain invalid symbols")
                        End Try

                    End If
                End If
            End If
        End If
    Next
    MsgBox("All dxf files have been added to the chosen location")
End Sub

' Creates the custom property if it does not exist, sets its expression,
' and returns the evaluated value (what Inventor displays).
Function SetCustomPropExpression(doc As Document, propName As String, expr As String) As String
    Dim customProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
    Dim prop As Inventor.Property = Nothing

    Try
        prop = customProps.Item(propName)
    Catch
        prop = customProps.Add("", propName)
    End Try

    prop.Expression = expr
    doc.Update()

    Return CStr(prop.Value).Trim()
End Function
