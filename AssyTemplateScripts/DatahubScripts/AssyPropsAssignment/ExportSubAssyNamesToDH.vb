Sub Main()

    Dim currentDoc As AssemblyDocument = ThisDoc.Document

    ' Header names used to locate the columns (row 1)
    Const HEADER_ROW As Integer = 1
    Const HEADER_SUBASSEMBLY As String = "SUBASSEMBLY"
    Const HEADER_PARENT As String = "PARENT"

    ' Read the setup from Inventor Defined User Properties
    Dim sFile As String = GetUserProperty(currentDoc, "Datahub")
    Dim tab As String = GetUserProperty(currentDoc, "SubAssySheet")

    If sFile = "" Or tab = "" Then
        Dim missing As String = ""
        If sFile = "" Then missing = missing & vbCrLf & " - Datahub (path of the Excel file)"
        If tab = "" Then missing = missing & vbCrLf & " - SubAssySheet (name of the worksheet)"
        MessageBox.Show("Setup required. The following custom iProperties are missing or empty:" & missing & vbCrLf & vbCrLf & "Please run the setup for both properties and try again.", "Setup required")
        Exit Sub
    End If

    ' Open the workbook on the requested worksheet (fails if it does not exist)
    Try
        GoExcel.Open(sFile, tab)
        Dim testValue As Object = GoExcel.CellValue("A1")
    Catch ex As Exception
        MessageBox.Show("Worksheet '" & tab & "' could not be opened in:" & vbCrLf & sFile & vbCrLf & vbCrLf & ex.Message, "Error")
        Exit Sub
    End Try

    ' Locate columns by header name
    Dim subCol As Integer = FindColumn(HEADER_SUBASSEMBLY, HEADER_ROW)
    Dim parentCol As Integer = FindColumn(HEADER_PARENT, HEADER_ROW)

    If subCol = 0 Or parentCol = 0 Then
        Dim missingHeaders As String = ""
        If subCol = 0 Then missingHeaders = missingHeaders & "'" & HEADER_SUBASSEMBLY & "' "
        If parentCol = 0 Then missingHeaders = missingHeaders & "'" & HEADER_PARENT & "' "
        MessageBox.Show("Column header(s) not found in row " & HEADER_ROW & " of worksheet '" & tab & "': " & missingHeaders, "Error")
        GoExcel.Close
        Exit Sub
    End If

    Dim subLetter As String = ColumnLetter(subCol)
    Dim parentLetter As String = ColumnLetter(parentCol)

    ' Read existing rows (SUBASSEMBLY and PARENT) to avoid duplicates
    Dim seen As New HashSet(Of String)
    Dim rowCounter As Integer = HEADER_ROW + 1

    While CellText(subLetter & rowCounter) <> "" And rowCounter < 5000
        Dim key As String = CellText(subLetter & rowCounter) & "|" & CellText(parentLetter & rowCounter)
        seen.Add(key)
        rowCounter = rowCounter + 1
    End While

    ' Collect sub-assemblies recursively
    Dim results As New List(Of String())
    CollectSubAssemblies(currentDoc.ComponentDefinition.Occurrences, "", results, seen)

    ' Write results: name under SUBASSEMBLY, immediate parent (if any) under PARENT
    For Each row As String() In results
        GoExcel.CellValue(subLetter & rowCounter) = row(0)
        If row(1) <> "" Then
            GoExcel.CellValue(parentLetter & rowCounter) = row(1)
        End If
        rowCounter = rowCounter + 1
    Next

    GoExcel.Save
    GoExcel.Close

End Sub

Sub CollectSubAssemblies(occs As System.Collections.IEnumerable, parentName As String, results As List(Of String()), seen As HashSet(Of String))

    For Each occ As ComponentOccurrence In occs

        If occ.Suppressed Then Continue For
        If occ.DefinitionDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Continue For

        Dim nextParent As String = parentName

        If occ.BOMStructure.ToString = "kNormalBOMStructure" Then

            Dim asmName As String = CleanName(occ.Definition.Document.DisplayName)

            ' Row = own name + immediate parent (empty if it sits directly under the top level)
            Dim key As String = asmName & "|" & parentName
            If Not seen.Contains(key) Then
                seen.Add(key)
                results.Add(New String() {asmName, parentName})
            End If

            ' This sub-assembly becomes the immediate parent for its children
            nextParent = asmName

        End If

        CollectSubAssemblies(occ.SubOccurrences, nextParent, results, seen)

    Next

End Sub

' Returns the value of a custom iProperty, or "" if it does not exist or is empty
Function GetUserProperty(doc As Document, propName As String) As String
    Try
        Dim prop As Inventor.Property = doc.PropertySets.Item("Inventor User Defined Properties").Item(propName)
        If prop.Value Is Nothing Then Return ""
        Return prop.Value.ToString().Trim()
    Catch
        Return ""
    End Try
End Function

' Returns the 1-based index of the column whose header matches headerName (case-insensitive), or 0 if not found
Function FindColumn(headerName As String, headerRow As Integer) As Integer
    Dim maxCols As Integer = 100
    For c As Integer = 1 To maxCols
        Dim txt As String = CellText(ColumnLetter(c) & headerRow)
        If txt.ToUpper() = headerName.ToUpper() Then Return c
    Next
    Return 0
End Function

Function CleanName(displayName As String) As String
    Dim idx As Integer = displayName.LastIndexOf(".")
    If idx > 0 Then Return displayName.Substring(0, idx)
    Return displayName
End Function

Function CellText(address As String) As String
    Dim v As Object = GoExcel.CellValue(address)
    If v Is Nothing Then Return ""
    Return v.ToString().Trim()
End Function

Function ColumnLetter(index As Integer) As String
    Dim result As String = ""
    While index > 0
        Dim remainder As Integer = (index - 1) Mod 26
        result = Chr(65 + remainder) & result
        index = (index - 1) \ 26
    End While
    Return result
End Function
