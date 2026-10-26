Sub Main()

    '------------------------------------------------------
    ' Settings
    '------------------------------------------------------
    Dim assyViewsFormatName As String = "AssyViews"
    Dim assyBomFormatName As String = "AssyBOM2"
    Dim partFormatName As String = "PartFab2"
    Dim flatFormatName As String = "FlatPattern"
    Dim coverSheetName As String = "Cover"

    Dim viewsSuffix As String = " Views"
    Dim bomSuffix As String = " BOM"
    Dim partSuffix As String = " Fabrication dwg"
    Dim flatSuffix As String = " Flat Pattern"

    Dim app As Inventor.Application = ThisApplication

    '------------------------------------------------------
    ' Active document must be a saved assembly
    '------------------------------------------------------
    Dim asmDoc As AssemblyDocument = TryCast(app.ActiveDocument, AssemblyDocument)

    If asmDoc Is Nothing Then
        MessageBox.Show("The active document must be an assembly.", "Invalid document")
        Exit Sub
    End If

    If String.IsNullOrEmpty(asmDoc.FullFileName) OrElse Not System.IO.File.Exists(asmDoc.FullFileName) Then
        MessageBox.Show("Please save the assembly before running this rule.", "Assembly not saved")
        Exit Sub
    End If

    '------------------------------------------------------
    ' Resolve templates path (active project, then default)
    '------------------------------------------------------
    Dim templatesPath As String = ""

    Try
        templatesPath = app.DesignProjectManager.ActiveDesignProject.TemplatesPath
    Catch
        templatesPath = ""
    End Try

    If String.IsNullOrEmpty(templatesPath) OrElse Not System.IO.Directory.Exists(templatesPath) Then
        Try
            templatesPath = app.FileLocations.TemplatesPath
        Catch
            templatesPath = ""
        End Try
    End If

    '------------------------------------------------------
    ' Let the user pick the IDW template
    '------------------------------------------------------
    Dim openDlg As New System.Windows.Forms.OpenFileDialog()
    openDlg.Title = "Select an IDW template"
    openDlg.Filter = "Inventor drawing templates (*.idw)|*.idw"
    openDlg.CheckFileExists = True

    If Not String.IsNullOrEmpty(templatesPath) AndAlso System.IO.Directory.Exists(templatesPath) Then
        openDlg.InitialDirectory = templatesPath
    End If

    If openDlg.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then Exit Sub

    '------------------------------------------------------
    ' Create the new drawing from the template
    '------------------------------------------------------
    Dim dwgDoc As DrawingDocument = Nothing

    Try
        dwgDoc = CType(app.Documents.Add( _
            DocumentTypeEnum.kDrawingDocumentObject, _
            openDlg.FileName, _
            True), DrawingDocument)
    Catch ex As Exception
        MessageBox.Show("Could not create the drawing from the template:" & vbCrLf & ex.Message)
        Exit Sub
    End Try

    '------------------------------------------------------
    ' Find all sheet formats before creating anything
    '------------------------------------------------------
    Dim assyViewsFmt As SheetFormat = FindSheetFormat(dwgDoc, assyViewsFormatName)
    Dim assyBomFmt As SheetFormat = FindSheetFormat(dwgDoc, assyBomFormatName)
    Dim partFmt As SheetFormat = FindSheetFormat(dwgDoc, partFormatName)
    Dim flatFmt As SheetFormat = FindSheetFormat(dwgDoc, flatFormatName)

    If assyViewsFmt Is Nothing OrElse assyBomFmt Is Nothing OrElse _
       partFmt Is Nothing OrElse flatFmt Is Nothing Then

        Dim missing As String = ""
        If assyViewsFmt Is Nothing Then missing &= vbCrLf & " - " & assyViewsFormatName
        If assyBomFmt Is Nothing Then missing &= vbCrLf & " - " & assyBomFormatName
        If partFmt Is Nothing Then missing &= vbCrLf & " - " & partFormatName
        If flatFmt Is Nothing Then missing &= vbCrLf & " - " & flatFormatName

        MessageBox.Show("Sheet format(s) not found in the selected template:" & missing, "Missing format")

        'Close the new drawing without saving
        dwgDoc.Close(True)
        Exit Sub

    End If

    '------------------------------------------------------
    ' Keep the template's default sheet and rename it
    '------------------------------------------------------
    Dim errors As String = ""
    Dim created As Integer = 0
    Dim err As String = ""
    Dim newSheet As Sheet = Nothing

    Try
        dwgDoc.Sheets.Item(1).Name = coverSheetName
    Catch ex As Exception
        errors &= vbCrLf & "Could not rename default sheet: " & ex.Message
    End Try

    '------------------------------------------------------
    ' Collect documents to draw (Normal BOM, not suppressed)
    '------------------------------------------------------
    Dim processed As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
    processed.Add(asmDoc.FullFileName)

    Dim docsToDraw As New System.Collections.Generic.List(Of Document)
    CollectDocuments(asmDoc.ComponentDefinition.Occurrences, processed, docsToDraw)

    '------------------------------------------------------
    ' Top level assembly: AssyViews + AssyBOM2
    '------------------------------------------------------
    Dim topPN As String = GetPartNumber(asmDoc)

    err = AddSheet(dwgDoc, assyViewsFmt, asmDoc, topPN & viewsSuffix, newSheet)
    If err = "" Then created += 1 Else errors &= vbCrLf & err

    err = AddSheet(dwgDoc, assyBomFmt, asmDoc, topPN & bomSuffix, newSheet)
    If err = "" Then created += 1 Else errors &= vbCrLf & err

    '------------------------------------------------------
    ' Subassemblies: AssyViews + AssyBOM2
    ' Parts: PartFab2 (+ FlatPattern if sheet metal)
    '------------------------------------------------------
    For Each doc As Document In docsToDraw

        Dim pn As String = GetPartNumber(doc)

        If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then

            err = AddSheet(dwgDoc, assyViewsFmt, doc, pn & viewsSuffix, newSheet)
            If err = "" Then created += 1 Else errors &= vbCrLf & err

            err = AddSheet(dwgDoc, assyBomFmt, doc, pn & bomSuffix, newSheet)
            If err = "" Then created += 1 Else errors &= vbCrLf & err

        Else

            err = AddSheet(dwgDoc, partFmt, doc, pn & partSuffix, newSheet)
            If err = "" Then created += 1 Else errors &= vbCrLf & err

            '--------------------------------------------------
            ' Sheet metal part: extra sheet with the flat pattern
            '--------------------------------------------------
            Dim smDef As SheetMetalComponentDefinition = GetSheetMetalDefinition(doc)

            If smDef IsNot Nothing Then

                If Not smDef.HasFlatPattern Then

                    errors &= vbCrLf & pn & ": sheet metal part has no flat pattern, flat sheet skipped"

                Else

                    err = AddSheet(dwgDoc, flatFmt, doc, pn & flatSuffix, newSheet)

                    If err <> "" Then

                        errors &= vbCrLf & err

                    Else

                        created += 1

                        'Make sure the views show the flat pattern
                        If EnsureFlatPatternViews(newSheet) = 0 Then
                            errors &= vbCrLf & pn & ": no view is in flat pattern (views on sheet: " & _
                                      newSheet.DrawingViews.Count & ")"
                        End If

                    End If

                End If

            End If

        End If

    Next

    '------------------------------------------------------
    ' Let the user choose where to save
    '------------------------------------------------------
    Dim saveDlg As New System.Windows.Forms.SaveFileDialog()
    saveDlg.Title = "Save drawing"
    saveDlg.Filter = "Inventor drawing (*.idw)|*.idw"
    saveDlg.DefaultExt = "idw"
    saveDlg.OverwritePrompt = True
    saveDlg.InitialDirectory = System.IO.Path.GetDirectoryName(asmDoc.FullFileName)
    saveDlg.FileName = System.IO.Path.GetFileNameWithoutExtension(asmDoc.FullFileName) & ".idw"

    If saveDlg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        Try
            dwgDoc.SaveAs(saveDlg.FileName, False)
        Catch ex As Exception
            errors &= vbCrLf & "Could not save drawing: " & ex.Message
        End Try
    End If

    '------------------------------------------------------
    ' Summary
    '------------------------------------------------------
    Dim msg As String = "Finished." & vbCrLf & "Sheets created: " & created

    If errors <> "" Then msg &= vbCrLf & vbCrLf & "Issues:" & errors

    MessageBox.Show(msg)

End Sub


'----------------------------------------------------------
' Add a sheet using a format. Returns "" on success or an
' error description on failure. The new sheet is returned
' through newSheet.
'----------------------------------------------------------
Function AddSheet(ByVal dwgDoc As DrawingDocument, _
                  ByVal fmt As SheetFormat, _
                  ByVal model As Document, _
                  ByVal sheetName As String, _
                  ByRef newSheet As Sheet) As String

    newSheet = Nothing

    Try
        newSheet = dwgDoc.Sheets.AddUsingSheetFormat(fmt, model, sheetName)
        Return ""
    Catch ex As Exception
        Return sheetName & ": " & ex.Message
    End Try

End Function


'----------------------------------------------------------
' Returns the sheet metal definition of a document, or
' Nothing if it is not a sheet metal part
'----------------------------------------------------------
Function GetSheetMetalDefinition(ByVal doc As Document) As SheetMetalComponentDefinition

    Dim partDoc As PartDocument = TryCast(doc, PartDocument)
    If partDoc Is Nothing Then Return Nothing

    Return TryCast(partDoc.ComponentDefinition, SheetMetalComponentDefinition)

End Function


'----------------------------------------------------------
' Switch the base views of a sheet from folded model to
' flat pattern. Returns the number of views switched.
'----------------------------------------------------------
Function EnsureFlatPatternViews(ByVal sh As Sheet) As Integer

    Dim flatCount As Integer = 0

    If sh Is Nothing Then Return 0

    For Each dv As DrawingView In sh.DrawingViews

        Try
            'Switch only if the view currently shows the folded model
            If dv.SheetMetalFoldedModel Then
                dv.SheetMetalFoldedModel = False
            End If

            'Count the view if it ends up as flat pattern
            If Not dv.SheetMetalFoldedModel Then flatCount += 1

        Catch
            'View does not support the option, ignore it
        End Try

    Next

    Return flatCount

End Function


'----------------------------------------------------------
' Recursively collect documents with Normal BOM structure.
' Phantom assemblies are traversed (their children are
' promoted in the BOM) but do not get a sheet themselves.
'----------------------------------------------------------
Sub CollectDocuments(ByVal occs As ComponentOccurrences, _
                     ByVal processed As System.Collections.Generic.HashSet(Of String), _
                     ByVal result As System.Collections.Generic.List(Of Document))

    For Each occ As ComponentOccurrence In occs

        'Skip suppressed occurrences
        If occ.Suppressed Then Continue For

        'Skip virtual components (no referenced file)
        If TypeOf occ.Definition Is VirtualComponentDefinition Then Continue For

        Dim bom As BOMStructureEnum = occ.Definition.BOMStructure

        'Only Normal and Phantom are relevant
        If bom <> BOMStructureEnum.kNormalBOMStructure AndAlso _
           bom <> BOMStructureEnum.kPhantomBOMStructure Then
            Continue For
        End If

        If bom = BOMStructureEnum.kNormalBOMStructure Then

            Dim doc As Document = occ.Definition.Document

            'Already handled (children were collected the first time)
            If processed.Contains(doc.FullFileName) Then Continue For

            processed.Add(doc.FullFileName)
            result.Add(doc)

        End If

        'Go deeper into subassemblies
        If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            CollectDocuments(occ.SubOccurrences, processed, result)
        End If

    Next

End Sub


'----------------------------------------------------------
' Find a sheet format by name (case-insensitive)
'----------------------------------------------------------
Function FindSheetFormat(ByVal dwgDoc As DrawingDocument, ByVal fmtName As String) As SheetFormat

    For Each fmt As SheetFormat In dwgDoc.SheetFormats
        If String.Equals(fmt.Name, fmtName, System.StringComparison.OrdinalIgnoreCase) Then
            Return fmt
        End If
    Next

    Return Nothing

End Function


'----------------------------------------------------------
' Part Number, falling back to the file name if empty
'----------------------------------------------------------
Function GetPartNumber(ByVal doc As Document) As String

    Dim pn As String = ""

    Try
        pn = CStr(doc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value)
    Catch
        pn = ""
    End Try

    If String.IsNullOrWhiteSpace(pn) Then
        pn = System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName)
    End If

    Return pn

End Function
