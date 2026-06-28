Option Explicit

Public Function NewAssembly(ByVal swApp As Object) As Object
    On Error Resume Next
    Set NewAssembly = swApp.NewAssembly
    On Error GoTo 0
End Function

Public Function AddAssemblyComponent(ByVal assy As Object, ByVal componentPath As String, ByVal xMm As Double, ByVal yMm As Double, ByVal zMm As Double) As Object
    On Error Resume Next
    Set AddAssemblyComponent = assy.AddComponent5(componentPath, 0, "", False, "", Mm(xMm), Mm(yMm), Mm(zMm))
    If AddAssemblyComponent Is Nothing Then
        Set AddAssemblyComponent = assy.AddComponent4(componentPath, "", Mm(xMm), Mm(yMm), Mm(zMm))
    End If
    On Error GoTo 0
End Function

Public Function SaveAssembly(ByVal model As Object, ByVal filePath As String) As Boolean
    On Error GoTo EH

    Dim errors As Long
    Dim warnings As Long

    model.ViewZoomtofit2
    SaveAssembly = model.Extension.SaveAs(filePath, 0, 1, Nothing, errors, warnings)
    If Not SaveAssembly Then AppendLog "SaveAssembly failed: " & filePath & " errors=" & CStr(errors) & " warnings=" & CStr(warnings)
    Exit Function

EH:
    AppendLog "SaveAssembly exception: " & filePath & " - " & Err.Description
    SaveAssembly = False
End Function