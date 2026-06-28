Option Explicit

Public Sub GenerateDrawingPlaceholders(ByVal swApp As Object, ByVal partPaths As Object, ByVal assemblyPath As String)
    On Error GoTo EH

    Dim key As Variant
    Dim partPath As String
    Dim drawingPath As String
    Dim made As Long

    If swApp Is Nothing Then Exit Sub

    For Each key In partPaths.Keys
        partPath = CStr(partPaths(key))
        If Len(partPath) > 0 Then
            drawingPath = JoinPath(DrawingDir(), Replace(GetFileName(partPath), ".SLDPRT", ".SLDDRW"))
            If CreateSimpleDrawing(swApp, partPath, drawingPath, CStr(key)) Then made = made + 1
        End If
    Next key

    If Len(assemblyPath) > 0 Then
        CreateSimpleDrawing swApp, assemblyPath, JoinPath(DrawingDir(), "ASM-000_AutoFeeder_Modular.SLDDRW"), "Top assembly layout drawing placeholder"
    End If

    AppendLog "Drawing placeholders generated: " & CStr(made)
    Exit Sub

EH:
    AppendLog "GenerateDrawingPlaceholders failed: " & Err.Description
End Sub

Public Function CreateSimpleDrawing(ByVal swApp As Object, ByVal modelPath As String, ByVal drawingPath As String, ByVal titleText As String) As Boolean
    On Error GoTo EH

    Dim templatePath As String
    Dim drawing As Object
    Dim errors As Long
    Dim warnings As Long

    templatePath = GetDefaultDrawingTemplate(swApp)
    If Len(templatePath) = 0 Then
        AppendLog "No default drawing template. Drawing skipped: " & drawingPath
        CreateSimpleDrawing = False
        Exit Function
    End If

    Set drawing = swApp.NewDocument(templatePath, 0, 0#, 0#)
    If drawing Is Nothing Then
        CreateSimpleDrawing = False
        Exit Function
    End If

    drawing.CreateDrawViewFromModelView3 modelPath, "*Front", 0.12, 0.18, 0#
    drawing.CreateDrawViewFromModelView3 modelPath, "*Top", 0.32, 0.18, 0#
    drawing.CreateDrawViewFromModelView3 modelPath, "*Right", 0.12, 0.08, 0#
    drawing.CreateDrawViewFromModelView3 modelPath, "*Isometric", 0.32, 0.08, 0#
    On Error Resume Next
    drawing.InsertNote titleText & vbCrLf & STATUS_PHOTO_ESTIMATE
    On Error GoTo EH

    CreateSimpleDrawing = drawing.Extension.SaveAs(drawingPath, 0, 1, Nothing, errors, warnings)
    swApp.CloseDoc drawing.GetTitle
    Exit Function

EH:
    AppendLog "CreateSimpleDrawing failed: " & drawingPath & " - " & Err.Description
    CreateSimpleDrawing = False
End Function

Private Function GetDefaultDrawingTemplate(ByVal swApp As Object) As String
    On Error Resume Next
    GetDefaultDrawingTemplate = swApp.GetUserPreferenceStringValue(3)
    If Len(GetDefaultDrawingTemplate) = 0 Then GetDefaultDrawingTemplate = swApp.GetUserPreferenceStringValue(2)
    On Error GoTo 0
End Function

Private Function GetFileName(ByVal filePath As String) As String
    GetFileName = Mid$(filePath, InStrRev(filePath, "\") + 1)
End Function