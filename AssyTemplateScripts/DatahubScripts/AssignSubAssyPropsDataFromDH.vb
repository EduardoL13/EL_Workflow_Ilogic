Sub Main()

    Dim currentDoc As AssemblyDocument = ThisDoc.Document

    ' Max number of worksheet rows scanned for sub-assembly names
    Const MAX_ROWS As Integer = 1000

    ' Read Datahub / SubAssySheet from User Defined Properties
    Dim sFile As String = GetCustomProp(currentDoc, "Datahub")
    Dim tabSA As String = GetCustomProp(currentDoc, "SubAssySheet")

    If sFile = "" Or tabSA = "" Then
        MsgBox("Datahub and/or SubAssySheet are not defined." & vbCrLf & "Run the setup rule first.", MsgBoxStyle.Exclamation, "Missing properties")
        Return
    End If

    If Not ConfirmSettings(sFile, tabSA) Then Return

    ' Resolve column letters from the header names in row 1 (fails if the worksheet does not exist)
    Dim headerCols As Dictionary(Of String, String)
    Try
        headerCols = BuildHeaderMap(sFile, tabSA)
    Catch ex As Exception
        MsgBox("Worksheet '" & tabSA & "' could not be read in:" & vbCrLf & sFile & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Worksheet error")
        Return
    End Try

    Dim requiredHeaders As String() = {"SUBASSEMBLY", "PART NO", "DESCRIPTION"}
    Dim missingHeaders As String = ""
    For Each h As String In requiredHeaders
        If Not headerCols.ContainsKey(h) Then
            missingHeaders &= h & vbCrLf
        End If
    Next

    If missingHeaders <> "" Then
        MsgBox("The following headers were not found in row 1 of worksheet '" & tabSA & "':" & vbCrLf & vbCrLf & _
               missingHeaders & vbCrLf & "Nothing was changed.", MsgBoxStyle.Exclamation, "Missing headers")
        Return
    End If

    Dim colSub As String = headerCols("SUBASSEMBLY")
    Dim colPartNo As String = headerCols("PART NO")
    Dim colDescription As String = headerCols("DESCRIPTION")

    ' Map each sub-assembly name in the worksheet to the row holding its data
    Dim sheetRows As Dictionary(Of String, Integer) = BuildSheetIndex(sFile, tabSA, colSub, colPartNo, MAX_ROWS)

    ' Collect unique sub-assemblies (normal BOM, not suppressed), including nested ones
    Dim subAssemblies As New Dictionary(Of String, AssemblyDocument)(StringComparer.OrdinalIgnoreCase)
    CollectSubAssemblyDocs(currentDoc.ComponentDefinition.Occurrences, subAssemblies)

    Dim updatedCount As Integer = 0
    Dim listNotFound As New List(Of String)
    Dim listIncomplete As New List(Of String)
    Dim listFailed As New List(Of String)

    For Each kvp As KeyValuePair(Of String, AssemblyDocument) In subAssemblies

        Dim asmName As String = kvp.Key

        If Not sheetRows.ContainsKey(asmName) Then
            listNotFound.Add(asmName)
            Continue For
        End If

        Dim r As Integer = sheetRows(asmName)
        Dim partNo As String = CellText(sFile, tabSA, colPartNo & r)
        Dim descr As String = CellText(sFile, tabSA, colDescription & r)

        If partNo = "" Or descr = "" Then listIncomplete.Add(asmName)
        If partNo = "" And descr = "" Then Continue For

        If WriteSubAssyProps(kvp.Value, partNo, descr) Then
            updatedCount += 1
        Else
            listFailed.Add(asmName)
        End If

    Next

    currentDoc.Update2

    Dim summary As String = "Sub-assemblies updated: " & updatedCount

    If listIncomplete.Count > 0 Then
        summary &= vbCrLf & vbCrLf & "PART NO and/or DESCRIPTION empty in worksheet for:" & vbCrLf & String.Join(vbCrLf, listIncomplete.ToArray())
    End If
    If listNotFound.Count > 0 Then
        summary &= vbCrLf & vbCrLf & "Not found under SUBASSEMBLY in worksheet:" & vbCrLf & String.Join(vbCrLf, listNotFound.ToArray())
    End If
    If listFailed.Count > 0 Then
        summary &= vbCrLf & vbCrLf & "Could not write properties (read-only file?):" & vbCrLf & String.Join(vbCrLf, listFailed.ToArray())
    End If

    MsgBox(summary, MsgBoxStyle.Information, "Done")

End Sub

Sub CollectSubAssemblyDocs(occs As System.Collections.IEnumerable, found As Dictionary(Of String, AssemblyDocument))

    For Each occ As ComponentOccurrence In occs

        If occ.Suppressed Then Continue For
        If occ.DefinitionDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Continue For

        If occ.BOMStructure = BOMStructureEnum.kNormalBOMStructure Then
            Dim asmDoc As AssemblyDocument = occ.Definition.Document
            Dim nm As String = CleanName(asmDoc.DisplayName)
            If Not found.ContainsKey(nm) Then found.Add(nm, asmDoc)
        End If

        ' Always go deeper, so normal sub-assemblies inside phantom/reference ones are found too
        CollectSubAssemblyDocs(occ.SubOccurrences, found)

    Next

End Sub

' Writes Part Number and Description (only the non-empty values) to the sub-assembly document
Function WriteSubAssyProps(asmDoc As AssemblyDocument, partNo As String, descr As String) As Boolean

    Dim targetDoc As AssemblyDocument = asmDoc
    Dim factoryDoc As AssemblyDocument = Nothing
    Dim prevScope As MemberEditScopeEnum = MemberEditScopeEnum.kEditActiveMember

    Try
        ' If the sub-assembly is a model state member, edit the properties on all members
        If asmDoc.ComponentDefinition.IsModelStateMember Then
            factoryDoc = asmDoc.ComponentDefinition.FactoryDocument
            prevScope = factoryDoc.ComponentDefinition.ModelStates.MemberEditScope
            factoryDoc.ComponentDefinition.ModelStates.MemberEditScope = MemberEditScopeEnum.kEditAllMembers
            targetDoc = factoryDoc
        End If

        Dim props As PropertySet = targetDoc.PropertySets.Item("Design Tracking Properties")
        If partNo <> "" Then props.Item("Part Number").Value = partNo
        If descr <> "" Then props.Item("Description").Value = descr

        Return True

    Catch
        Return False

    Finally
        If factoryDoc IsNot Nothing Then
            factoryDoc.ComponentDefinition.ModelStates.MemberEditScope = prevScope
        End If
    End Try

End Function

' Maps each SUBASSEMBLY name to its row. If a name appears in several rows
' (same sub-assembly under different parents), the first row with PART NO filled wins.
Function BuildSheetIndex(file As String, tab As String, colSub As String, colPartNo As String, maxRow As Integer) As Dictionary(Of String, Integer)

    Dim index As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

    For rowNum As Integer = 2 To maxRow
        Dim nm As String = CellText(file, tab, colSub & rowNum)
        If nm <> "" Then
            If Not index.ContainsKey(nm) Then
                index.Add(nm, rowNum)
            ElseIf CellText(file, tab, colPartNo & index(nm)) = "" And CellText(file, tab, colPartNo & rowNum) <> "" Then
                index(nm) = rowNum
            End If
        End If
    Next

    Return index

End Function

Function BuildHeaderMap(file As String, tab As String) As Dictionary(Of String, String)
    ' Reads row 1 and maps each header text (trimmed, upper case) to its column letter
    Dim map As New Dictionary(Of String, String)
    Dim maxCols As Integer = 50

    For colNum As Integer = 1 To maxCols
        Dim letter As String = ColumnLetter(colNum)
        Dim key As String = CellText(file, tab, letter & "1").ToUpper()
        If key <> "" AndAlso Not map.ContainsKey(key) Then
            map.Add(key, letter)
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

Function CellText(file As String, tab As String, address As String) As String
    Dim v As Object = GoExcel.CellValue(file, tab, address)
    If v Is Nothing Then Return ""
    Return v.ToString().Trim()
End Function

Function CleanName(displayName As String) As String
    Dim idx As Integer = displayName.LastIndexOf(".")
    If idx > 0 Then Return displayName.Substring(0, idx)
    Return displayName
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
                                        "SubAssySheet = " & sheet & vbCrLf & vbCrLf & _
                                        "Do you want to continue with these values?", _
                                        MsgBoxStyle.OkCancel + MsgBoxStyle.Question, "Confirm Datahub settings")
    Return (result = MsgBoxResult.Ok)
End Function
