Sub Main()

    Dim currentDoc As AssemblyDocument = ThisDoc.Document

    ' Start a single transaction so the whole script is one undo step
    Dim trans As Inventor.Transaction = ThisApplication.TransactionManager.StartTransaction(currentDoc, "Origin Plane Constraints")

    Try
        AssemblyOriginCons(currentDoc)
        trans.End() ' Commit: everything above becomes ONE entry in the undo stack
    Catch ex As Exception
        trans.Abort() ' On error, roll back everything done so far
        MessageBox.Show("Script aborted, changes rolled back:" & vbCrLf & ex.Message, "Error")
    End Try

End Sub


Sub AssemblyOriginCons(assemDef As AssemblyDocument)

    ' Assembly origin planes declaration
    Dim PlaneE1 As WorkPlane = assemDef.ComponentDefinition.WorkPlanes.Item(1)
    Dim PlaneE2 As WorkPlane = assemDef.ComponentDefinition.WorkPlanes.Item(2)
    Dim PlaneE3 As WorkPlane = assemDef.ComponentDefinition.WorkPlanes.Item(3)

    ' Loop over occurrences and constrain each one to the origin planes
    For Each compOcc As ComponentOccurrence In assemDef.ComponentDefinition.Occurrences
        If compOcc.Constraints.Count = 0 And compOcc.PatternElement Is Nothing Then

            ' Occurrence origin planes declaration
            Dim Plane1 As WorkPlane = compOcc.Definition.WorkPlanes.Item(1)
            Dim Plane2 As WorkPlane = compOcc.Definition.WorkPlanes.Item(2)
            Dim Plane3 As WorkPlane = compOcc.Definition.WorkPlanes.Item(3)

            ' Geometry proxy creation prior to constraint imposition
            Dim APlane1 As WorkPlaneProxy = Nothing
            compOcc.CreateGeometryProxy(Plane1, APlane1)
            Dim APlane2 As WorkPlaneProxy = Nothing
            compOcc.CreateGeometryProxy(Plane2, APlane2)
            Dim APlane3 As WorkPlaneProxy = Nothing
            compOcc.CreateGeometryProxy(Plane3, APlane3)

            ' Origin planes constraint imposition
            assemDef.ComponentDefinition.Constraints.AddFlushConstraint(APlane1, PlaneE1, 0)
            assemDef.ComponentDefinition.Constraints.AddFlushConstraint(APlane2, PlaneE2, 0)
            assemDef.ComponentDefinition.Constraints.AddFlushConstraint(APlane3, PlaneE3, 0)

        End If
    Next

End Sub
