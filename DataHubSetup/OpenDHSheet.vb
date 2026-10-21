Sub Main()

    Dim currentDoc As Document = ThisDoc.Document

    ' Read Datahub / PropsSheet from User Defined Properties
    Dim datahub As String = GetCustomProp(currentDoc, "Datahub")
    Dim currentSheet As String = GetCustomProp(currentDoc, "PropsSheet")

    If datahub = "" Then
        MsgBox("Datahub is not defined." & vbCrLf & "Please run the setup rule first.", MsgBoxStyle.Exclamation, "Missing property")
        Return
    End If

    If Not System.IO.File.Exists(datahub) Then
        MsgBox("The Datahub file was not found:" & vbCrLf & datahub & vbCrLf & vbCrLf & _
               "Please run the setup rule again.", MsgBoxStyle.Exclamation, "File not found")
        Return
    End If

    ' Read the worksheet names available in the file
    Dim sheetNames As List(Of String) = GetSheetNames(datahub)
    Dim targetSheet As String = ""

    If sheetNames Is Nothing Then
        Dim answer As MsgBoxResult = MsgBox("The worksheet names could not be read from the Datahub file." & vbCrLf & _
                                            "The worksheet name will NOT be validated." & vbCrLf & vbCrLf & _
                                            "Do you want to continue anyway?", _
                                            MsgBoxStyle.OkCancel + MsgBoxStyle.Exclamation, "Open Datahub")
        If answer <> MsgBoxResult.Ok Then Return

        targetSheet = InputBox("Enter the worksheet name to open:", "Open Datahub", currentSheet).Trim()
        If targetSheet = "" Then Return
    Else
        Dim defaultSheet As String = currentSheet
        Do
            Dim entry As String = InputBox("Enter the worksheet name to open." & vbCrLf & vbCrLf & _
                                           "Available sheets: " & String.Join(", ", sheetNames), _
                                           "Open Datahub", defaultSheet).Trim()

            If entry = "" Then Return ' Canceled or empty

            ' Case-insensitive match; keep the exact name as it is in the Excel file
            Dim matchedName As String = ""
            For Each sName As String In sheetNames
                If String.Equals(sName, entry, StringComparison.OrdinalIgnoreCase) Then
                    matchedName = sName
                    Exit For
                End If
            Next

            If matchedName <> "" Then
                targetSheet = matchedName
                Exit Do
            End If

            Dim retry As MsgBoxResult = MsgBox("The worksheet '" & entry & "' does not exist in the Datahub file." & vbCrLf & vbCrLf & _
                                               "Available sheets:" & vbCrLf & String.Join(vbCrLf, sheetNames) & vbCrLf & vbCrLf & _
                                               "Do you want to try again?", _
                                               MsgBoxStyle.OkCancel + MsgBoxStyle.Exclamation, "Open Datahub")
            If retry <> MsgBoxResult.Ok Then Return

            defaultSheet = entry
        Loop
    End If

    ' Confirm before opening
    Dim confirm As MsgBoxResult = MsgBox("Excel will be opened with:" & vbCrLf & vbCrLf & _
                                         "Datahub = " & datahub & vbCrLf & _
                                         "Worksheet = " & targetSheet & vbCrLf & vbCrLf & _
                                         "Do you want to continue?", _
                                         MsgBoxStyle.OkCancel + MsgBoxStyle.Question, "Open Datahub")
    If confirm <> MsgBoxResult.Ok Then Return

    OpenWorkbookAtSheet(datahub, targetSheet)

End Sub

Sub OpenWorkbookAtSheet(filePath As String, sheetName As String)
    ' Opens the workbook in a visible Excel instance (reusing a running one if possible)
    ' and activates the requested worksheet. Excel is left open for the user.
    Dim xlApp As Object = Nothing
    Dim wb As Object = Nothing
    Dim ws As Object = Nothing
    Dim failedStep As String = "starting"

    Try
        ' Step 1: get an Excel instance
        failedStep = "attaching to a running Excel instance"
        Try
            xlApp = Microsoft.VisualBasic.Interaction.GetObject(Nothing, "Excel.Application")
        Catch
            xlApp = Nothing
        End Try

        If xlApp Is Nothing Then
            failedStep = "creating a new Excel instance"
            xlApp = CreateObject("Excel.Application")
        End If

        If xlApp Is Nothing Then
            MsgBox("Excel could not be started.", MsgBoxStyle.Exclamation, "Open Datahub")
            Return
        End If

        failedStep = "making Excel visible"
        xlApp.Visible = True

        ' Step 2: reuse the workbook if it is already open
        failedStep = "searching for an already open workbook"
        Dim wbCount As Integer = CInt(xlApp.Workbooks.Count)
        For i As Integer = 1 To wbCount
            Dim candidate As Object = xlApp.Workbooks.Item(i)
            If candidate IsNot Nothing Then
                If String.Equals(candidate.FullName.ToString(), filePath, StringComparison.OrdinalIgnoreCase) Then
                    wb = candidate
                    Exit For
                End If
            End If
        Next

        ' Step 3: open it if it was not open
        If wb Is Nothing Then
            failedStep = "opening the workbook"
            wb = xlApp.Workbooks.Open(filePath)
        End If

        If wb Is Nothing Then
            MsgBox("The workbook could not be opened:" & vbCrLf & filePath, MsgBoxStyle.Exclamation, "Open Datahub")
            Return
        End If

        ' Step 4: activate the workbook and the worksheet
        failedStep = "activating the workbook"
        wb.Activate()

        failedStep = "getting the worksheet '" & sheetName & "'"
        ws = wb.Worksheets.Item(sheetName)

        If ws Is Nothing Then
            MsgBox("The worksheet '" & sheetName & "' was not found.", MsgBoxStyle.Exclamation, "Open Datahub")
            Return
        End If

        failedStep = "activating the worksheet '" & sheetName & "'"
        ws.Activate()

    Catch ex As Exception
        MsgBox("Failed while: " & failedStep & vbCrLf & vbCrLf & _
               "Details: " & ex.Message & vbCrLf & vbCrLf & _
               "Note: hidden sheets cannot be activated.", MsgBoxStyle.Exclamation, "Open Datahub")
    Finally
        If ws IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(ws)
        If wb IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(wb)
        If xlApp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
    End Try
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
