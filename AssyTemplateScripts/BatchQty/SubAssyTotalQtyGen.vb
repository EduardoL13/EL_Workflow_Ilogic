Sub Main()

    Dim currentDoc As AssemblyDocument = ThisDoc.Document

    Dim prodQty As String = InputBox("Enter the desired total qty of products", "TOTAL QTY") ' Multiplier based on the required number of products

    If prodQty = "" Then
        Exit Sub
    End If

    ' Dictionaries keyed by the full document name of each unique subassembly
    Dim qtyByDoc As New Dictionary(Of String, Double)
    Dim docByName As New Dictionary(Of String, AssemblyDocument)

    ' Walk the whole tree and count every normal, non-suppressed subassembly occurrence
    CountSubAssemblies(currentDoc.ComponentDefinition.Occurrences, CDbl(prodQty), qtyByDoc, docByName)

    ' Write the accumulated quantities (overwrites any previous value)
    For Each key As String In qtyByDoc.Keys
        WriteQty(docByName(key), qtyByDoc(key))
    Next

    MsgBox("Done. Subassemblies updated: " & qtyByDoc.Count)

End Sub

Sub CountSubAssemblies(occs As ComponentOccurrences, pQty As Double, qtyByDoc As Dictionary(Of String, Double), docByName As Dictionary(Of String, AssemblyDocument))

    For Each compOcc As ComponentOccurrence In occs

        If compOcc.Suppressed Then Continue For

        ' Only assemblies are relevant here (parts are ignored)
        If Not TypeOf compOcc.Definition.Document Is AssemblyDocument Then Continue For

        Dim subDoc As AssemblyDocument = compOcc.Definition.Document

        Select Case compOcc.BOMStructure

            Case BOMStructureEnum.kNormalBOMStructure
                ' Resolve the factory document if the occurrence is a model state member
                Dim targetDoc As AssemblyDocument = subDoc
                If subDoc.ComponentDefinition.IsModelStateMember Then
                    targetDoc = subDoc.ComponentDefinition.FactoryDocument
                End If

                Dim key As String = targetDoc.FullDocumentName
                If qtyByDoc.ContainsKey(key) Then
                    qtyByDoc(key) += pQty
                Else
                    qtyByDoc.Add(key, pQty)
                    docByName.Add(key, targetDoc)
                End If

                ' Go inside: every occurrence path adds pQty, so nested subassemblies multiply correctly
                CountSubAssemblies(compOcc.SubOccurrences, pQty, qtyByDoc, docByName)

            Case BOMStructureEnum.kPhantomBOMStructure
                ' A phantom is not counted, but its children appear in the BOM, so keep going inside
                CountSubAssemblies(compOcc.SubOccurrences, pQty, qtyByDoc, docByName)

            Case Else
                ' Reference, Purchased, Inseparable, etc.: neither counted nor traversed

        End Select

    Next

End Sub

' Writes the quantity to the subassembly document, handling the model state edit scope
Sub WriteQty(doc As AssemblyDocument, qty As Double)

    Dim oCurrentScope As MemberEditScopeEnum
    Dim isMember As Boolean = doc.ComponentDefinition.IsModelStateMember

    If isMember Then
        oCurrentScope = doc.ComponentDefinition.ModelStates.MemberEditScope
        doc.ComponentDefinition.ModelStates.MemberEditScope = MemberEditScopeEnum.kEditAllMembers
    End If

    SetCustomQty(doc, qty)

    If isMember Then
        doc.ComponentDefinition.ModelStates.MemberEditScope = oCurrentScope
    End If

End Sub

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
