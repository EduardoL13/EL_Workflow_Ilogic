Imports System.Globalization

Sub Main()
    Const IN_TO_CM As Double = 2.54

    Dim partDoc As PartDocument = ThisDoc.Document
    Dim userParams As UserParameters = partDoc.ComponentDefinition.Parameters.UserParameters

    ' Read control parameters (must exist as user parameters)
    Dim holeType As String
    Dim holeSize As String
    Try
        holeType = userParams.Item("HoleClearanceType").Value
        holeSize = userParams.Item("HoleTableImp").Value
    Catch
        MsgBox("User parameters 'HoleClearanceType' and/or 'HoleTableImp' not found.")
        Exit Sub
    End Try

    ' Feature suffix
    Dim suffix As String = InputBox("Hole Generator", "Feature Name", "ex:AtCover")
    If suffix = "" Then Exit Sub
    Dim paramName As String = "diaHole" & suffix

    ' Avoid duplicates
    For Each p As UserParameter In userParams
        If p.Name = paramName Then
            MsgBox("Parameter '" & paramName & "' already exists.")
            Exit Sub
        End If
    Next

    ' Table: size ; fastener dia ; tapped ; clear ; offset   (all in inches)
    Dim rows() As String = { _
        "#0;0.06;0.052;0.07;0.07", _
        "#1;0.073;0.0625;0.081;0.081", _
        "#2;0.086;0.073;0.096;0.096", _
        "#3;0.099;0.086;0.11;0.11", _
        "#4;0.112;0.096;0.1285;0.1285", _
        "#5;0.125;0.1094;0.1285;0.1285", _
        "#6;0.138;0.116;0.1495;0.1495", _
        "#8;0.164;0.144;0.177;0.177", _
        "#10;0.19;0.161;0.201;0.201", _
        "#12;0.216;0.189;0.228;0.228", _
        "1/4;0.25;0.2188;0.266;0.375", _
        "5/16;0.3125;0.277;0.332;0.457", _
        "3/8;0.375;0.332;0.397;0.5", _
        "7/16;0.4375;0.3096;0.4687;0.5625", _
        "1/2;0.5;0.4531;0.5312;0.625", _
        "9/16;0.5625;0.5156;0.5938;0.6875", _
        "5/8;0.625;0.5625;0.6562;0.75", _
        "11/16;0.6875;0.6562;0.6562;0.8125", _
        "3/4;0.75;0.6875;0.7812;0.875", _
        "13/16;0.8125;0.7812;0.8438;0.9375", _
        "7/8;0.875;0.7969;0.9062;1.0", _
        "15/16;0.9375;0.9062;0.9688;1.0625", _
        "1;1.0;0.9219;1.0313;1.125"}

    ' Column index depending on hole type
    Dim col As Integer
    Select Case holeType
        Case "Tapped" : col = 2
        Case "Clear" : col = 3
        Case "Offset" : col = 4
        Case Else
            MsgBox("Unknown HoleClearanceType: " & holeType)
            Exit Sub
    End Select

    ' Look up size and create the parameter
    For Each r As String In rows
        Dim f() As String = r.Split(";"c)
        If f(0) = holeSize Then
            Dim fastener As Double = Double.Parse(f(1), CultureInfo.InvariantCulture)
            Dim dia As Double = Double.Parse(f(col), CultureInfo.InvariantCulture)

            userParams.AddByValue(paramName, dia * IN_TO_CM, "in")
            userParams.Item(paramName).Comment = holeType & " hole for " & fastener & "in dia fastener (" & holeSize & ")"

            MsgBox("Parameter '" & paramName & "' has been generated.")
            Exit Sub
        End If
    Next

    MsgBox("Fastener size not found in table: " & holeSize)
End Sub
