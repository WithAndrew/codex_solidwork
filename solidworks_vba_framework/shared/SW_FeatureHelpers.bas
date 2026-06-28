Option Explicit

Public Function NewPart(ByVal swApp As Object) As Object
    On Error Resume Next
    Set NewPart = swApp.NewPart
    On Error GoTo 0
End Function

Public Function SaveModel(ByVal swApp As Object, ByVal model As Object, ByVal filePath As String) As Boolean
    On Error GoTo EH

    Dim errors As Long
    Dim warnings As Long

    SaveModel = model.Extension.SaveAs(filePath, 0, 1, Nothing, errors, warnings)
    If Not SaveModel Then AppendLog "SaveModel failed: " & filePath & " errors=" & CStr(errors) & " warnings=" & CStr(warnings)
    Exit Function

EH:
    AppendLog "SaveModel exception: " & filePath & " - " & Err.Description
    SaveModel = False
End Function

Public Sub CreateBox(ByVal model As Object, ByVal lengthMm As Double, ByVal widthMm As Double, ByVal heightMm As Double)
    model.Extension.SelectByID2 "Front Plane", "PLANE", 0#, 0#, 0#, False, 0, Nothing, 0
    model.SketchManager.InsertSketch True
    model.SketchManager.CreateCenterRectangle 0#, 0#, 0#, Mm(lengthMm / 2#), Mm(widthMm / 2#), 0#
    model.SketchManager.InsertSketch True
    model.FeatureManager.FeatureExtrusion2 True, False, False, 0, 0, Mm(heightMm), 0#, False, False, False, False, 0#, 0#, False, False, False, False, True, True, True, 0, 0, False
End Sub

Public Sub CreateCylinder(ByVal model As Object, ByVal diameterMm As Double, ByVal heightMm As Double)
    model.Extension.SelectByID2 "Top Plane", "PLANE", 0#, 0#, 0#, False, 0, Nothing, 0
    model.SketchManager.InsertSketch True
    model.SketchManager.CreateCircleByRadius 0#, 0#, 0#, Mm(diameterMm / 2#)
    model.SketchManager.InsertSketch True
    model.FeatureManager.FeatureExtrusion2 True, False, False, 0, 0, Mm(heightMm), 0#, False, False, False, False, 0#, 0#, False, False, False, False, True, True, True, 0, 0, False
End Sub

Public Sub CutHole(ByVal model As Object, ByVal faceName As String, ByVal xMm As Double, ByVal yMm As Double, ByVal diameterMm As Double, ByVal depthMm As Double)
    AppendLog "CutHole template called for " & model.GetTitle & ". Add face-specific SolidWorks feature code in the part module before machining release."
End Sub

Public Sub CutCounterbore(ByVal model As Object, ByVal faceName As String, ByVal xMm As Double, ByVal yMm As Double, ByVal holeDiameterMm As Double, ByVal counterboreDiameterMm As Double, ByVal counterboreDepthMm As Double, ByVal throughDepthMm As Double)
    AppendLog "CutCounterbore template called for " & model.GetTitle & ". Add verified SolidWorks feature code in the part module before machining release."
End Sub

Public Sub CutCountersink(ByVal model As Object, ByVal faceName As String, ByVal xMm As Double, ByVal yMm As Double, ByVal holeDiameterMm As Double, ByVal sinkDiameterMm As Double, ByVal sinkAngleDeg As Double, ByVal throughDepthMm As Double)
    AppendLog "CutCountersink template called for " & model.GetTitle & ". Add verified SolidWorks feature code in the part module before machining release."
End Sub

Public Sub CreateTappedHole(ByVal model As Object, ByVal faceName As String, ByVal xMm As Double, ByVal yMm As Double, ByVal threadText As String, ByVal depthMm As Double)
    AddCustomProperty model, "TappedHole_Note", threadText & " depth " & CStr(depthMm) & " mm at " & faceName & " (" & CStr(xMm) & "," & CStr(yMm) & ")"
End Sub

Public Sub CutSlot(ByVal model As Object, ByVal faceName As String, ByVal xMm As Double, ByVal yMm As Double, ByVal lengthMm As Double, ByVal widthMm As Double, ByVal depthMm As Double)
    AppendLog "CutSlot template called for " & model.GetTitle & ". Add verified SolidWorks slot feature code in the part module before machining release."
End Sub

Public Sub AddChamfer(ByVal model As Object, ByVal chamferMm As Double)
    AppendLog "AddChamfer template called for " & model.GetTitle & ". Select target edges in the part module for production geometry."
End Sub

Public Sub AddFillet(ByVal model As Object, ByVal filletMm As Double)
    AppendLog "AddFillet template called for " & model.GetTitle & ". Select target edges in the part module for production geometry."
End Sub
