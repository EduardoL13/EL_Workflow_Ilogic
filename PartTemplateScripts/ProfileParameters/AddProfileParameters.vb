' ---------------------------------------------------------------------------
' Body size parameters
'
' Asks the body type first:
'   0 = Profile      -> creates pLength1<suffix>, pLength2<suffix>, pThk<suffix>
'   1 = Sheet/Plate  -> creates only pThk<suffix>
'
' Units follow the document length units (mm, in, ...).
' If a parameter already exists, it is left untouched (no duplicates, no
' overwritten values).
' ---------------------------------------------------------------------------

Sub Main()

    ' ---- Document checks -------------------------------------------------
    Dim partDoc As PartDocument = TryCast(ThisDoc.Document, PartDocument)
    If partDoc Is Nothing Then
        MsgBox("This rule only works in a part document.", MsgBoxStyle.Exclamation, "Body Parameters")
        Exit Sub
    End If

    Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition
    Dim userParams As UserParameters = partDef.Parameters.UserParameters

    ' ---- Document length unit (e.g. "mm", "in") ----------------------------
    Dim uom As UnitsOfMeasure = partDoc.UnitsOfMeasure
    Dim lenUnit As String = uom.GetStringFromType(uom.LengthUnits)

    ' ---- 1) Body type: 0 = profile, 1 = sheet/plate --------------------------
    Dim typeText As String = Trim(InputBox("Enter body type:" & vbCrLf & _
                                           "0 = Profile" & vbCrLf & _
                                           "1 = Sheet / Plate", "Body Type", "0"))
    If typeText = "" Then Exit Sub

    If typeText <> "0" AndAlso typeText <> "1" Then
        MsgBox("Please enter 0 (Profile) or 1 (Sheet/Plate).", MsgBoxStyle.Exclamation, "Body Parameters")
        Exit Sub
    End If
    Dim isProfile As Boolean = (typeText = "0")

    ' ---- 2) Suffix -----------------------------------------------------------
    Dim sufix As String = InputBox("Enter body name (no spaces)", "Body Name", "CoreTube")
    sufix = Trim(sufix).Replace(" ", "_")
    If sufix = "" Then Exit Sub

    If Not Char.IsLetter(sufix(0)) AndAlso sufix(0) <> "_"c Then
        MsgBox("The name must start with a letter or underscore.", MsgBoxStyle.Exclamation, "Body Parameters")
        Exit Sub
    End If

    ' ---- 3) Create the parameters (single undo step) --------------------------
    Dim trans As Transaction = ThisApplication.TransactionManager.StartTransaction( _
        ThisApplication.ActiveDocument, "Create body parameters")

    Try
        Dim commentText As String = "Profile Parameter"
        'If Not isProfile Then commentText = "Sheet/Plate Parameter"

        If isProfile Then
            CreateIfMissing(userParams, "pLength1" & sufix, lenUnit, commentText)
            CreateIfMissing(userParams, "pLength2" & sufix, lenUnit, commentText)
        End If
        CreateIfMissing(userParams, "pThk" & sufix, lenUnit, commentText)

        trans.End()
    Catch ex As Exception
        trans.Abort()
        MsgBox("Error creating parameters: " & ex.Message, MsgBoxStyle.Critical, "Body Parameters")
        Exit Sub
    End Try

End Sub

' Creates a user parameter with a default value of 1 (in document units) and a
' comment. If a parameter with that name already exists, it is not modified.
Sub CreateIfMissing(ByVal userParams As UserParameters, ByVal name As String, _
                    ByVal units As String, ByVal commentText As String)

    For Each up As UserParameter In userParams
        If up.Name = name Then Exit Sub
    Next

    ' Expression with explicit unit avoids database-unit ambiguity of AddByValue
    Dim newParam As UserParameter = userParams.AddByExpression(name, "1 " & units, units)
    newParam.Comment = commentText

End Sub
