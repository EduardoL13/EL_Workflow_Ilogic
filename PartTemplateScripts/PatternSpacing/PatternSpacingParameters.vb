' ---------------------------------------------------------------------------
' Pattern spacing parameters (unified rule)
'
' Creates the user parameters that let a linear pattern self-adjust while
' keeping a fixed spacing:
'   spacingValue<suffix>   : desired spacing (asked via InputBox)
'   lengthPattern<suffix>  : user parameter pointing to a driven dimension
'                            (picked by the user)
'   rawNoElements<suffix>  : lengthPattern / spacingValue
'   NoElements<suffix>     : number of elements for the pattern
'   widthEffective<suffix> : real distance covered by the pattern
'   distEnds<suffix>       : offset from each end (use as pattern offset)
'
' Units follow the document length units (mm, in, ...).
' If a parameter already exists, its expression is updated (no duplicates).
' Everything runs inside one transaction -> a single Undo reverts it.
' ---------------------------------------------------------------------------

Sub Main()

    ' ---- Document checks -------------------------------------------------
    Dim partDoc As PartDocument = TryCast(ThisDoc.Document, PartDocument)
    If partDoc Is Nothing Then
        MsgBox("This rule only works in a part document.", MsgBoxStyle.Exclamation, "Pattern Spacing")
        Exit Sub
    End If

    Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition
    Dim userParams As UserParameters = partDef.Parameters.UserParameters

    ' ---- Document length unit (e.g. "mm", "in") ----------------------------
    Dim uom As UnitsOfMeasure = partDoc.UnitsOfMeasure
    Dim lenUnit As String = uom.GetStringFromType(uom.LengthUnits)

    ' Suggested default spacing depending on the unit
    Dim defaultSpacing As String
    Select Case lenUnit
        Case "mm" : defaultSpacing = "10"
        Case "in" : defaultSpacing = "0.5"
        Case Else : defaultSpacing = "1"
    End Select

    ' ---- 1) Suffix ---------------------------------------------------------
    Dim sufix As String = InputBox("Enter pattern feature name (no spaces)", "Pattern Feature Name", "Face1holes")
    sufix = Trim(sufix).Replace(" ", "_")
    If sufix = "" Then Exit Sub

    If Not Char.IsLetter(sufix(0)) AndAlso sufix(0) <> "_"c Then
        MsgBox("The name must start with a letter or underscore.", MsgBoxStyle.Exclamation, "Pattern Spacing")
        Exit Sub
    End If

    ' ---- 2) Spacing value --------------------------------------------------
    Dim spacingText As String = InputBox("Enter desired spacing (" & lenUnit & ")", "Spacing Value", defaultSpacing)
    If Trim(spacingText) = "" Then Exit Sub

    Dim spacingNum As Double
    ' Accept both "0.5" and "0,5"
    If Not Double.TryParse(spacingText.Replace(",", "."), _
                           System.Globalization.NumberStyles.Float, _
                           System.Globalization.CultureInfo.InvariantCulture, spacingNum) _
       OrElse spacingNum <= 0 Then
        MsgBox("Spacing must be a number greater than 0.", MsgBoxStyle.Exclamation, "Pattern Spacing")
        Exit Sub
    End If

    ' ---- 3) Pick the driven dimension ---------------------------------------
    ' The sketch containing the driven dimension must be in edit mode.
    Dim drivenName As String = ""
    Dim picked As Object = ThisApplication.CommandManager.Pick( _
        SelectionFilterEnum.kAllEntitiesFilter, _
        "Select the driven dimension that measures the pattern length (Esc to type a parameter name instead)")

    If Not picked Is Nothing Then
        Dim dimC As DimensionConstraint = TryCast(picked, DimensionConstraint)
        If dimC Is Nothing Then
            MsgBox("The selected object is not a dimension.", MsgBoxStyle.Exclamation, "Pattern Spacing")
            Exit Sub
        End If
        drivenName = dimC.Parameter.Name
    Else
        ' Fallback: type the parameter name (e.g. d12)
        drivenName = Trim(InputBox("Enter the name of the driven parameter (e.g. d12)", "Driven Parameter", ""))
        If drivenName = "" Then Exit Sub
    End If

    ' Make sure that parameter really exists
    Dim testParam As Parameter = Nothing
    Try
        testParam = partDef.Parameters.Item(drivenName)
    Catch
        MsgBox("Parameter '" & drivenName & "' was not found.", MsgBoxStyle.Exclamation, "Pattern Spacing")
        Exit Sub
    End Try

    ' ---- 4) Create / update the parameters (single undo step) ----------------
    Dim trans As Transaction = ThisApplication.TransactionManager.StartTransaction( _
        ThisApplication.ActiveDocument, "Create pattern spacing parameters")

    Try
        Dim spacingName As String = "spacingValue" & sufix
        Dim lengthName As String = "lengthPattern" & sufix
        Dim rawName As String = "rawNoElements" & sufix
        Dim noName As String = "NoElements" & sufix
        Dim widthName As String = "widthEffective" & sufix
        Dim endsName As String = "distEnds" & sufix

        ' Small tolerance so exact multiples (e.g. 10 / 2.5) are not lost to
        ' floating point rounding in floor()
        Const tol As String = "0.000001"

        ' Use expressions with explicit units to avoid database-unit ambiguity
        SetOrCreate(userParams, spacingName, _
                    spacingNum.ToString(System.Globalization.CultureInfo.InvariantCulture) & " " & lenUnit, lenUnit)

        ' Separate parameter pointing to the driven dimension. If the driven
        ' dimension is deleted, just edit this expression to another one.
        SetOrCreate(userParams, lengthName, drivenName, lenUnit)

        SetOrCreate(userParams, rawName, lengthName & " / " & spacingName, "ul")
        SetOrCreate(userParams, noName, "floor(" & rawName & " + " & tol & ") + 1", "ul")
        SetOrCreate(userParams, widthName, "(" & noName & " - 1) * " & spacingName, lenUnit)
        SetOrCreate(userParams, endsName, "(" & lengthName & " - " & widthName & ") / 2", lenUnit)

        ' Mark the main driving/result parameters as Key Parameters
        userParams.Item(spacingName).IsKey = True
        userParams.Item(noName).IsKey = True

        trans.End()
    Catch ex As Exception
        trans.Abort()
        MsgBox("Error creating parameters: " & ex.Message, MsgBoxStyle.Critical, "Pattern Spacing")
        Exit Sub
    End Try

    partDoc.Update()

    MsgBox("Parameters created for '" & sufix & "'." & vbCrLf & _
           "Use NoElements" & sufix & " and distEnds" & sufix & " in the pattern dialog.", _
           MsgBoxStyle.Information, "Pattern Spacing")

End Sub

' Creates the user parameter, or updates its expression if it already exists
Sub SetOrCreate(ByVal userParams As UserParameters, ByVal name As String, _
                ByVal expr As String, ByVal units As String)

    Dim existing As UserParameter = Nothing
    For Each up As UserParameter In userParams
        If up.Name = name Then
            existing = up
            Exit For
        End If
    Next

    If existing Is Nothing Then
        userParams.AddByExpression(name, expr, units)
    Else
        existing.Expression = expr
    End If

End Sub
