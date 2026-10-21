' ---------------------------------------------------------------------------
' Body parameter string
'
' Asks the body type:
'   0 = Profile      -> looks for pLength1, pLength2, pThk
'                       pParamString = "=<pLength1>x<pLength2>x<pThk>"
'   1 = Sheet/Plate  -> looks for pThk
'                       pParamString = "=<pThk>"
'
' Steps:
'   1) For every expected user parameter that exists, tick "Export Parameter"
'      (Inventor then creates the linked custom iProperty with the same name)
'      and set its format: no trailing zeros, no unit string.
'   2) Create (or update) the custom iProperty "pParamString" with the
'      expression above.
'   3) Report any expected parameter that was not found.
'
' No suffix is used. Everything runs in one transaction (single Undo).
' ---------------------------------------------------------------------------

Sub Main()

    ' ---- Document check ----------------------------------------------------
    Dim partDoc As PartDocument = TryCast(ThisDoc.Document, PartDocument)
    If partDoc Is Nothing Then
        MsgBox("This rule only works in a part document.", MsgBoxStyle.Exclamation, "Body Param String")
        Exit Sub
    End If

    Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition
    Dim userParams As UserParameters = partDef.Parameters.UserParameters

    ' ---- 1) Body type: 0 = profile, 1 = sheet/plate --------------------------
    Dim typeText As String = Trim(InputBox("Enter body type:" & vbCrLf & _
                                           "0 = Profile" & vbCrLf & _
                                           "1 = Sheet / Plate", "Body Type", "0"))
    If typeText = "" Then Exit Sub

    If typeText <> "0" AndAlso typeText <> "1" Then
        MsgBox("Please enter 0 (Profile) or 1 (Sheet/Plate).", MsgBoxStyle.Exclamation, "Body Param String")
        Exit Sub
    End If

    Dim expectedNames As String()
    Dim paramString As String

    If typeText = "0" Then
        expectedNames = New String() {"pLength1", "pLength2", "pThk"}
        paramString = "=<pLength1>x<pLength2>x<pThk>"
    Else
        expectedNames = New String() {"pThk"}
        paramString = "=<pThk>"
    End If

    ' ---- 2) Export matching parameters, then create pParamString -------------
    Dim missing As New System.Collections.Generic.List(Of String)
    Dim exported As New System.Collections.Generic.List(Of String)
    Dim formatErrors As New System.Collections.Generic.List(Of String)
    Dim propError As String = ""

    Dim trans As Transaction = ThisApplication.TransactionManager.StartTransaction( _
        ThisApplication.ActiveDocument, "Body param string")

    Try
        ' Step A: tick "Export Parameter" on the matching user parameters
        For Each pName As String In expectedNames
            Dim found As UserParameter = Nothing
            For Each up As UserParameter In userParams
                If up.Name = pName Then
                    found = up
                    Exit For
                End If
            Next

            If found Is Nothing Then
                missing.Add(pName)
            Else
                If Not found.ExposedAsProperty Then found.ExposedAsProperty = True
                ApplyExportFormat(found, formatErrors)
                exported.Add(pName)
            End If
        Next

        ' Step B: create or update the custom property pParamString
        Dim customSet As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")

        Dim paramStringProp As Inventor.Property = Nothing
        For Each prop As Inventor.Property In customSet
            If prop.Name = "pParamString" Then
                paramStringProp = prop
                Exit For
            End If
        Next

        If paramStringProp Is Nothing Then
            paramStringProp = customSet.Add("", "pParamString")
        End If

        Try
            paramStringProp.Expression = paramString
        Catch ex As Exception
            propError = ex.Message
        End Try

        trans.End()
    Catch ex As Exception
        trans.Abort()
        MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical, "Body Param String")
        Exit Sub
    End Try

    partDoc.Update()

    ' ---- 3) Report ---------------------------------------------------------------
    Dim msg As String = "pParamString = " & paramString & vbCrLf & vbCrLf

    If exported.Count > 0 Then
        msg &= "Exported parameters: " & String.Join(", ", exported.ToArray()) & vbCrLf
    End If

    If missing.Count > 0 Then
        msg &= "Parameters NOT found (create them to complete the link): " & _
               String.Join(", ", missing.ToArray()) & vbCrLf
    End If

    If propError <> "" Then
        msg &= vbCrLf & "Warning: could not set the property expression: " & propError & vbCrLf
    End If

    If formatErrors.Count > 0 Then
        msg &= vbCrLf & "Warning: could not apply the export format:" & vbCrLf & _
               String.Join(vbCrLf, formatErrors.ToArray())
    End If

    MsgBox(msg, MsgBoxStyle.Information, "Body Param String")

End Sub

' Sets the custom property format of an exported parameter:
'   - trailing zeros OFF (1.000 in -> 1)
'   - unit string OFF    (no "in" / "mm" next to the value)
' Late binding + one Try per setting, so a failure in one member does not
' block the other and is reported instead of stopping the rule.
Sub ApplyExportFormat(ByVal prm As UserParameter, ByVal errors As System.Collections.Generic.List(Of String))

    Dim fmt As Object = Nothing
    Try
        fmt = prm.CustomPropertyFormat
    Catch ex As Exception
        errors.Add(prm.Name & " (CustomPropertyFormat): " & ex.Message)
        Exit Sub
    End Try

    Try
        fmt.ShowTrailingZeros = False
    Catch ex As Exception
        errors.Add(prm.Name & " (ShowTrailingZeros): " & ex.Message)
    End Try

    Try
        fmt.ShowUnitsString = False
    Catch ex As Exception
        errors.Add(prm.Name & " (ShowUnitsString): " & ex.Message)
    End Try

End Sub
