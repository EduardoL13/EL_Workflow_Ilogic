Sub Main()
    Dim doc As PartDocument = ThisDoc.Document
    Dim compDef As PartComponentDefinition = doc.ComponentDefinition

    ' 1. Ask the user how many planes to select (1 to 3)
    Dim qtyInput As String = InputBox("Enter the number of planes to select (1, 2 or 3):", "Number of Planes", "1")

    ' Validate user input
    Dim numPlanes As Integer
    If Not Integer.TryParse(qtyInput, numPlanes) OrElse numPlanes < 1 OrElse numPlanes > 3 Then
        MessageBox.Show("Invalid input. You must enter a number between 1 and 3.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Return
    End If

    ' List to store the names of the selected planes
    Dim selectedPlanesNames As New List(Of String)

    ' 2. Loop to pick the requested number of planes (done before the transaction starts)
    For i As Integer = 1 To numPlanes
        Dim selectedPlane As Object
        Try
            ' Specific filter for WorkPlanes (work planes and origin planes)
            selectedPlane = ThisApplication.CommandManager.Pick( _
                SelectionFilterEnum.kWorkPlaneFilter, _
                "Select Plane " & i & " of " & numPlanes & ":")
        Catch
            ' The user pressed ESC or cancelled the selection
            Return
        End Try

        ' Exit if nothing was selected
        If selectedPlane Is Nothing Then Return

        ' Store the name of the selected plane
        selectedPlanesNames.Add(selectedPlane.Name)
    Next

    ' 3. Open a transaction so the whole operation is a single undo step (Ctrl+Z)
    Dim trans As Transaction = ThisApplication.TransactionManager.StartTransaction( _
        CType(doc, Inventor.Document), "Create Planes Design View")

    Try
        ' 4. Create a new design view representation with a unique name and activate it
        Dim repName As String = GetUniqueViewName(compDef, "Planes View")
        Dim newRep As DesignViewRepresentation = _
            compDef.RepresentationsManager.DesignViewRepresentations.Add(repName)
        newRep.Activate()

        ' 5. Hide all work planes except the selected ones (stored in the new view)
        For Each wp As WorkPlane In compDef.WorkPlanes
            If selectedPlanesNames.Contains(wp.Name) Then
                wp.Visible = True ' Keep visible if selected
            Else
                wp.Visible = False ' Hide if not selected
            End If
        Next

        ' NOTE: Sketch visibility is intentionally left untouched.

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
