Option Explicit

Public Sub AddCustomProperty(ByVal model As Object, ByVal propertyName As String, ByVal propertyValue As String)
    On Error Resume Next
    model.CustomInfo2("", propertyName) = propertyValue
    On Error GoTo 0
End Sub

Public Sub AddEstimatedDimensionProperties(ByVal model As Object, ByVal lengthMm As Double, ByVal widthMm As Double, ByVal heightMm As Double)
    AddCustomProperty model, "DimensionSource", STATUS_PHOTO_ESTIMATE
    AddCustomProperty model, "Estimated_L_mm", CStr(lengthMm)
    AddCustomProperty model, "Estimated_W_mm", CStr(widthMm)
    AddCustomProperty model, "Estimated_H_mm", CStr(heightMm)
End Sub

Public Sub SetMaterial(ByVal model As Object, ByVal materialName As String)
    On Error Resume Next
    AddCustomProperty model, "Material", materialName
    model.SetMaterialPropertyName2 "", "", materialName
    On Error GoTo 0
End Sub