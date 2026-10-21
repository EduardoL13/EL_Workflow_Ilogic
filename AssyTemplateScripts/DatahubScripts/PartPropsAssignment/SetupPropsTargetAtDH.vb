Sub Main()

    Dim currentDoc As Document = ThisDoc.Document
    Const PROP_DATAHUB As String = "Datahub"
    Const PROP_SHEET As String = "PropsSheet"

    ' Current values (used as defaults if they already exist)
    Dim currentDatahub As String = GetCustomProp(currentDoc, PROP_DATAHUB)
    Dim currentSheet As String = GetCustomProp(currentDoc, PROP_SHEET)

    ' Step 1: browse for the Excel file
    Dim oFileDlg As Inventor.FileDialog = Nothing
    ThisApplication.CreateFileDialog(oFileDlg)
    oFileDlg.Filter = "Excel Files (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm|All Files (*.*)|*.*"
    oFileDlg.DialogTitle = "Select the Datahub Excel file"
    oFileDlg.CancelError = True

    If currentDatahub <> "" AndAlso System.IO.File.Exists(currentDatahub) Then
        oFileDlg.InitialDirectory = System.IO.Path.GetDirectoryName(currentDatahub)
    Else
        oFileDlg.InitialDirectory = ThisDoc.Path
    End If

    Dim newDatahub As String = ""
    Try
        oFileDlg.ShowOpen()
        newDatahub = oFileDlg.FileName
    Catch
        MsgBox("Setup canceled. No properties were changed.", MsgBoxStyle.Information, "Datahub Setup")
        Return
    End Try

    If newDatahub = "" Then
        MsgBox("No file selected. No properties were changed.", MsgBoxStyle.Information, "Datahub Setup")
        Return
    End If

    ' Step 2: read the worksheet names available in the selected file
    Dim sheetNames As List(Of String) = GetSheetNames(newDatahub)

    ' Step 3: ask for the worksheet name (validated against the file)
    Dim newSheet As String = ""

    If sheetNames Is Nothing Then
        ' Could not open the file to validate (Excel not available, file locked, etc.)
        Dim answer As MsgBoxResult = MsgBox("The worksheet names could not be read from the selected file." & vbCrLf & _
                                            "The worksheet name will NOT be validated." & vbCrLf & vbCrLf & _
                                            "Do you want to continue anyway?", _
                                            MsgBoxStyle.OkCancel + MsgBoxStyle.Exclamation, "Datahub Setup")
        If answer <> MsgBoxResult.Ok Then
            MsgBox("Setup canceled. No properties were changed.", MsgBoxStyle.Information, "Datahub Setup")
            Return
        End If

        newSheet = InputBox("Enter the worksheet name (PropsSheet):", "Datahub Setup", currentSheet).Trim()
        If newSheet = "" Then
            MsgBox("No worksheet name entered. No properties were changed.", MsgBoxStyle.Information, "Datahub Setup")
            Return
        End If
    Else
        Dim defaultSheet As String = currentSheet
        Do
            Dim entry As String = InputBox("Enter the worksheet name (PropsSheet)." & vbCrLf & vbCrLf & _
                                           "Available sheets: " & String.Join(", ", sheetNames), _
                                           "Datahub Setup", defaultSheet).Trim()

            If entry = "" Then
                MsgBox("No worksheet name entered. No properties were changed.", MsgBoxStyle.Information, "Datahub Setup")
                Return
            End If

            ' Case-insensitive match; keep the exact name as it is in the Excel file
            Dim matchedName As String = ""
            For Each sName As String In sheetNames
                If String.Equals(sName, entry, StringComparison.OrdinalIgnoreCase) Then
                    matchedName = sName
                    Exit For
                End If
            Next

            If matchedName <> "" Then
                newSheet = matchedName
                Exit Do
            End If

            Dim retry As MsgBoxResult = MsgBox("The worksheet '" & entry & "' does not exist in the selected file." & vbCrLf & vbCrLf & _
                                               "Available sheets:" & vbCrLf & String.Join(vbCrLf, sheetNames) & vbCrLf & vbCrLf & _
                                               "Do you want to try again?", _
                                               MsgBoxStyle.OkCancel + MsgBoxStyle.Exclamation, "Datahub Setup")
            If retry <> MsgBoxResult.Ok Then
                MsgBox("Setup canceled. No properties were changed.", MsgBoxStyle.Information, "Datahub Setup")
                Return
            End If

            defaultSheet = entry
        Loop
    End If

    ' Step 4: create/update the User Defined Properties
    SetCustomProp(currentDoc, PROP_DATAHUB, newDatahub)
    SetCustomProp(currentDoc, PROP_SHEET, newSheet)

    ' Step 5: show final values (read back from the document)
    MsgBox("Properties assigned successfully:" & vbCrLf & vbCrLf & _
           PROP_DATAHUB & " = " & GetCustomProp(currentDoc, PROP_DATAHUB) & vbCrLf & _
           PROP_SHEET & " = " & GetCustomProp(currentDoc, PROP_SHEET), _
           MsgBoxStyle.Information, "Datahub Setup")

End Sub

Function GetSheetNames(filePath As String) As List(Of String)
    ' Opens the workbook read-only in a hidden Excel instance and returns its worksheet names.
    ' Returns Nothing if the file could not be read.
    Dim names As New List(Of String)
    Dim xlApp As Object = Nothing
    Dim wb As Object = Nothing

    Try
        xlApp = CreateObject("Excel.Application")
        xlApp.Visible = False
        xlApp.DisplayAlerts = False
        wb = xlApp.Workbooks.Open(filePath, 0, True) ' UpdateLinks = 0, ReadOnly = True

        For i As Integer = 1 To wb.Worksheets.Count
            names.Add(wb.Worksheets(i).Name.ToString())
        Next

        Return names
    Catch
        Return Nothing
    Finally
        Try
            If wb IsNot Nothing Then wb.Close(False)
        Catch
        End Try
        Try
            If xlApp IsNot Nothing Then xlApp.Quit()
        Catch
        End Try
        If wb IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(wb)
        If xlApp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
    End Try
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

Sub SetCustomProp(doc As Document, propName As String, propValue As String)
    Dim userProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
    For Each prop As Inventor.Property In userProps
        If prop.Name = propName Then
            prop.Value = propValue
            Return
        End If
    Next
    userProps.Add(propValue, propName)
End Sub
