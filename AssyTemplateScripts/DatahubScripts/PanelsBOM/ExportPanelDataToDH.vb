Sub Main()

    'NOTE: Dimension values are in cm by default

    Dim esteDoc As AssemblyDocument = ThisDoc.Document

    Dim leafOccurrences As ComponentOccurrencesEnumerator = _
        esteDoc.ComponentDefinition.Occurrences.AllLeafOccurrences

    '----------------------------------------
    'Read Datahub / CutlistSheet from User Defined Properties
    '----------------------------------------

    Dim file As String = GetCustomProp(esteDoc, "Datahub")
    Dim tab As String = GetCustomProp(esteDoc, "PanelsSheet")

    If file = "" Or tab = "" Then
        MsgBox("Datahub and/or CutlistSheet are not defined." & vbCrLf & _
               "Run the setup rule first.", _
               MsgBoxStyle.Exclamation, "Missing properties")
        Return
    End If

    If Not ConfirmSettings(file, tab) Then Return

    '----------------------------------------
    'Excel Columns
    '----------------------------------------

    Dim partTypeColumn As String = "A"

    Dim partLengthColumn As String = "B"

    Dim partHeightColumn As String = "C"

    Dim partQtyColumn As String = "D"

    Dim partIDColumn As String = "F"

    Dim partBOMprop As String = "G"

    ' Dim partThkColumn As String = "H"

    ' Dim partMaterialColumn As String = "I"

    '----------------------------------------
    'Build Excel Index
    '----------------------------------------

    Dim startRow As Integer

    Dim nextAvailableRow As Integer

    Dim rowDictionary As Dictionary(Of String, Integer)

    rowDictionary = BuildExcelIndex( _
                        file, _
                        tab, _
                        partTypeColumn, _
                        "P", _
                        partIDColumn, _
                        startRow, _
                        nextAvailableRow)

    '----------------------------------------
    'Avoid duplicate parts
    '----------------------------------------

    Dim processedParts As New HashSet(Of String)

    '----------------------------------------
    'Process Assembly
    '----------------------------------------

    For Each compOccurrence As ComponentOccurrence In leafOccurrences

        If compOccurrence.Suppressed Then Continue For

        If compOccurrence.BOMStructure <> _
            BOMStructureEnum.kNormalBOMStructure Then Continue For

        If TypeOf compOccurrence.Definition.Document IsNot PartDocument Then Continue For

        Dim occDoc As PartDocument = compOccurrence.Definition.Document

        If Not processedParts.Add(occDoc.FullFileName) Then Continue For

        If occDoc.ComponentDefinition.Type.ToString <> _
            "kSheetMetalComponentDefinitionObject" Then Continue For

        If Not occDoc.ComponentDefinition.HasFlatPattern Then Continue For

        Dim partID As String

        partID = occDoc.PropertySets.Item( _
            "Design Tracking Properties").Item( _
            "Stock Number").Value

        Dim currentRow As Integer

        If rowDictionary.ContainsKey(partID) Then

            currentRow = rowDictionary(partID)

        Else

            currentRow = nextAvailableRow

            rowDictionary.Add(partID, currentRow)

            nextAvailableRow += 1

        End If

        propsPrinter( _
            occDoc, _
            file, _
            tab, _
            partTypeColumn, _
            partIDColumn, _
            partQtyColumn, _
            partLengthColumn, _
            partHeightColumn, _
            partBOMprop, _
            currentRow)

    Next

    GoExcel.Save()

    MsgBox("Export Done")

End Sub



Sub propsPrinter( _
    currentPart As PartDocument, _
    file As String, _
    tab As String, _
    partTypeColumn As String, _
    partIDColumn As String, _
    partQtyColumn As String, _
    partLengthColumn As String, _
    partHeightColumn As String, _
    partBOMprop As String, _
    currentRow As Integer)

    Dim ConvFactor As Double = 1 / 2.54

    Try

        'Type
        GoExcel.CellValue(file, tab, partTypeColumn & currentRow) = "P"

        'Part ID
        GoExcel.CellValue(file, tab, partIDColumn & currentRow) = _
            currentPart.PropertySets.Item( _
            "Design Tracking Properties").Item( _
            "Stock Number").Value

        'Quantity
        GoExcel.CellValue(file, tab, partQtyColumn & currentRow) = _
            Single.Parse(currentPart.PropertySets.Item( _
            "Inventor User Defined Properties").Item( _
            "Part Qty").Value)

        'Length
        GoExcel.CellValue(file, tab, partLengthColumn & currentRow) = _
            currentPart.ComponentDefinition.FlatPattern. _
            ModelAnnotations.ModelDimensions.Item("length").ModelValue * ConvFactor

        'Height
        GoExcel.CellValue(file, tab, partHeightColumn & currentRow) = _
            currentPart.ComponentDefinition.FlatPattern. _
            ModelAnnotations.ModelDimensions.Item("height").ModelValue * ConvFactor

        'BOM prop
        GoExcel.CellValue(file, tab, partBOMprop & currentRow) = _
            currentPart.PropertySets.Item( _
            "Inventor User Defined Properties").Item( _
            "BOM Prop").Value


        'Thickness
 '       GoExcel.CellValue(file, tab, partThkColumn & currentRow) = _
 '           currentPart.ComponentDefinition.FlatPattern. _
 '           ModelAnnotations.ModelDimensions.Item("thk").ModelValue * ConvFactor

        'Material
 '       GoExcel.CellValue(file, tab, partMaterialColumn & currentRow) = _
 '           currentPart.PropertySets.Item( _
 '           "Design Tracking Properties").Item( _
 '           "Material").Value

    Catch

        MsgBox(currentPart.DisplayName & _
            " has missing properties or properties that are not in the valid format")

    End Try

End Sub



Function BuildExcelIndex( _
    file As String, _
    tab As String, _
    partTypeColumn As String, _
    partType As String, _
    partIDColumn As String, _
    ByRef startRow As Integer, _
    ByRef nextAvailableRow As Integer) _
    As Dictionary(Of String, Integer)

    Dim rowDictionary As New Dictionary(Of String, Integer)

    Dim searchRange As Integer = 1000

    Dim firstEmptyRow As Integer = 0
    Dim firstPartRow As Integer = 0
    Dim lastPartRow As Integer = 0

    For rowNum As Integer = 2 To searchRange

        Dim typeValue As Object

        typeValue = GoExcel.CellValue(file, tab, partTypeColumn & rowNum)

        '------------------------------------------
        'Store the first empty row
        '------------------------------------------

        If typeValue Is Nothing Then

            If firstEmptyRow = 0 Then
                firstEmptyRow = rowNum
            End If

            Continue For

        End If

        '------------------------------------------
        'Skip rows that belong to another table
        '------------------------------------------

        If typeValue.ToString.Trim.ToUpper <> partType.ToUpper Then Continue For

        '------------------------------------------
        'Store first and last row of this section
        '------------------------------------------

        If firstPartRow = 0 Then

            firstPartRow = rowNum

        End If

        lastPartRow = rowNum

        '------------------------------------------
        'Read Part ID
        '------------------------------------------

        Dim idValue As Object

        idValue = GoExcel.CellValue(file, tab, partIDColumn & rowNum)

        If idValue Is Nothing Then Continue For

        Dim partID As String

        partID = idValue.ToString.Trim

        If Not rowDictionary.ContainsKey(partID) Then

            rowDictionary.Add(partID, rowNum)

        End If

    Next

    '------------------------------------------
    'Determine write positions
    '------------------------------------------

    If firstPartRow = 0 Then

        startRow = firstEmptyRow
        nextAvailableRow = firstEmptyRow

    Else

        startRow = firstPartRow
        nextAvailableRow = lastPartRow + 1

    End If

    Return rowDictionary

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
                                        "CutlistSheet = " & sheet & vbCrLf & vbCrLf & _
                                        "Do you want to continue with these values?", _
                                        MsgBoxStyle.OkCancel + MsgBoxStyle.Question, "Confirm Datahub settings")
    Return (result = MsgBoxResult.Ok)
End Function
