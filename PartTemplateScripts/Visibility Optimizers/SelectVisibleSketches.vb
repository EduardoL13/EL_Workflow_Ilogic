Sub Main()
    Dim doc As PartDocument = ThisDoc.Document
    Dim compDef As PartComponentDefinition = doc.ComponentDefinition

    ' 1. Ask the user how many sketches to select (1 to 3)
    Dim qtyInput As String = InputBox("Enter desired sketches (1, 2 or 3):", "Quantity of Sketches", "1")

    ' Validate user input
    Dim numSketches As Integer
    If Not Integer.TryParse(qtyInput, numSketches) OrElse numSketches < 1 OrElse numSketches > 3 Then
        MessageBox.Show("Invalid input. Number of sketches must be between 1 and 3.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Return
    End If

    ' List to store the names of the selected sketches
    Dim selectedNames As New List(Of String)

    ' 2. Loop to pick the requested number of sketches (done before the transaction starts)
    For i As Integer = 1 To numSketches
        Dim selectedSketch As Object
        Try
            selectedSketch = ThisApplication.CommandManager.Pick( _
                SelectionFilterEnum.kSketchObjectFilter, _
                "Select Sketch " & i & " of " & numSketches & ":")
        Catch
            ' The user pressed ESC or cancelled the selection
            Return
        End Try

        ' Exit if nothing was selected
        If selectedSketch Is Nothing Then Return

        ' Store the name of the selected sketch
        selectedNames.Add(selectedSketch.Name)
    Next

    ' 3. Open a transaction so the whole operation is a single undo step (Ctrl+Z)
    Dim trans As Transaction = ThisApplication.TransactionManager.StartTransaction( _
        CType(doc, Inventor.Document), "Create Sketches Design View")

    Try
        ' 4. Create a new design view representation with a unique name and activate it
        Dim repName As String = GetUniqueViewName(compDef, "Sketches View")
        Dim newRep As DesignViewRepresentation = _
            compDef.RepresentationsManager.DesignViewRepresentations.Add(repName)
        newRep.Activate()

        ' 5. Hide all work planes (stored in the new view)
        For Each wp As WorkPlane In compDef.WorkPlanes
            wp.Visible = False
        Next

        ' 6. Process 2D sketches (PlanarSketches)
        For Each sk2d As PlanarSketch In compDef.Sketches
            If selectedNames.Contains(sk2d.Name) Then
                sk2d.Visible = True ' Keep visible if selected
            Else
                sk2d.Visible = False ' Hide if not selected
            End If
        Next

        ' 7. Process 3D sketches (Sketches3D)
        For Each sk3d As Sketch3D In compDef.Sketches3D
            If selectedNames.Contains(sk3d.Name) Then
                sk3d.Visible = True ' Keep visible if selected
            Else
                sk3d.Visible = False ' Hide if not selected
            End If
        Next

        trans.End()
    Catch ex As Exception
        ' Roll back everything if something fails
        trans.Abort()
        MessageBox.Show("Operation failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Return
    End Try

    ' Refresh the Inventor view
    ThisApplication.ActiveView.Update()
End Sub

' Returns a view name that does not already exist in the part
Function GetUniqueViewName(compDef As PartComponentDefinition, baseName As String) As String
    Dim reps As DesignViewRepresentations = compDef.RepresentationsManager.DesignViewRepresentations
    Dim candidate As String = baseName
    Dim counter As Integer = 1

    While True
        Dim exists As Boolean = False
        For Each r As DesignViewRepresentation In reps
            If r.Name = candidate Then
                exists = True
                Exit For
            End If
        Next
        If Not exists Then Return candidate
        counter += 1
        candidate = baseName & " " & counter
    End While

    Return candidate
End Function
