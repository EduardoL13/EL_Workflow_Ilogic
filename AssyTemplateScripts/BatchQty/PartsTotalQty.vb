Sub Main()

    Dim currentDoc As AssemblyDocument = ThisDoc.Document

    Dim prodQty As String = InputBox("Enter the desired total qty of products", "TOTAL QTY") ' Multiplier based on the required number of products

    If prodQty = "" Then
        Exit Sub
    End If

    setQty(currentDoc, CDbl(prodQty))

    MsgBox("Done")

End Sub

Sub setQty(assembComp As AssemblyDocument, pQty As Double)

    Dim listKeyStrings As New List(Of String) ' List of unique part names already processed

    For Each compOcc As ComponentOccurrence In assembComp.ComponentDefinition.Occurrences.AllLeafOccurrences

        If compOcc.Suppressed Then Continue For

        ' Only count parts that are in Normal BOM structure
        If TypeOf compOcc.Definition.Document IsNot PartDocument Or compOcc.BOMStructure <> BOMStructureEnum.kNormalBOMStructure Then Continue For

        Dim occDoc As PartDocument = compOcc.Definition.Document ' Access to the part document of the occurrence
        Dim nameToCompare As String = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))

        Dim oCurrentScope As MemberEditScopeEnum
        Dim oFactoryDoc As PartDocument = Nothing

        ' If the part is a model state member, edit the factory document (all members)
        If occDoc.ComponentDefinition.IsModelStateMember Then
            oFactoryDoc = occDoc.ComponentDefinition.FactoryDocument
            oCurrentScope = oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope

            If oCurrentScope = MemberEditScopeEnum.kEditActiveMember Then
                oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = MemberEditScopeEnum.kEditAllMembers
            End If
            occDoc = oFactoryDoc
        End If

        If listKeyStrings.Contains(nameToCompare) Then
            ' Part already counted: add pQty to the existing value
            Dim newValue As Double = GetCustomQty(occDoc) + pQty
            SetCustomQty(occDoc, newValue)
        Else
            ' First time this part is found: initialize (overwrite) the value
            listKeyStrings.Add(nameToCompare)
            SetCustomQty(occDoc, pQty)
        End If

        ' Restore the original edit scope
        If oFactoryDoc IsNot Nothing Then
            oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = oCurrentScope
        End If

    Next

End Sub

' Returns the current value of the "Part Qty" custom property (0 if it does not exist or is not numeric)
Function GetCustomQty(doc As Document) As Double

    Dim customProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")

    For Each prop As Inventor.Property In customProps
        If prop.Name = "Part Qty" Then
            Dim result As Double = 0
            If Double.TryParse(CStr(prop.Value), result) Then
                Return result
            End If
            Return 0
        End If
    Next

    Return 0

End Function

' Creates the "Part Qty" custom property if it does not exist, otherwise updates its value
Sub SetCustomQty(doc As Document, qty As Double)

    Dim customProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")

    For Each prop As Inventor.Property In customProps
        If prop.Name = "Part Qty" Then
            prop.Value = qty.ToString
            Return
        End If
    Next

    customProps.Add(qty.ToString, "Part Qty")

End Sub
