Option Explicit

Public Function Build_SM_001_OptionalGuardPanel(ByVal swApp As Object, ByVal outputDir As String) As String
    On Error GoTo Fail

    Dim filePath As String
    Dim model As Object

    ' Basic parameters
    Dim partNo As String
    Dim partName As String
    Dim fileName As String
    Dim lengthMm As Double
    Dim widthMm As Double
    Dim heightMm As Double
    Dim quantity As Double
    Dim materialName As String
    Dim photoReference As String
    Dim categoryName As String

    partNo = "SM-001"
    partName = "Optional transparent guard panel"
    fileName = "SM-001_OptionalGuardPanel_900x6x420.SLDPRT"
    lengthMm = 900#
    widthMm = 6#
    heightMm = 420#
    quantity = 1#
    materialName = "TO_BE_CONFIRMED"
    photoReference = "53,54"
    categoryName = "SheetMetal"

    ' Create base body
    Set model = NewPart(swApp)
    If model Is Nothing Then GoTo Fail
    CreateBox model, lengthMm, widthMm, heightMm

    ' Machining features: holes / slots / threads
    ' Add this part's verified feature code here.
    ' Examples:
    ' CutHole model, "Top", 50#, 50#, 6#, heightMm
    ' CutCounterbore model, "Top", 120#, 50#, 6.6, 11#, 6#, heightMm
    ' CutCountersink model, "Top", 220#, 50#, 6.6, 11#, 90#, heightMm
    ' CreateTappedHole model, "Top", 320#, 50#, "M6x1", 15#

    ' Chamfer / fillet
    ' AddChamfer model, 1#
    ' AddFillet model, 2#

    ' Material and custom properties
    SetMaterial model, materialName
    AddCustomProperty model, "PartNo", partNo
    AddCustomProperty model, "PartName", partName
    AddCustomProperty model, "Category", categoryName
    AddCustomProperty model, "PhotoReference", photoReference
    AddCustomProperty model, "Quantity", CStr(quantity)
    AddEstimatedDimensionProperties model, lengthMm, widthMm, heightMm

    ' Save file
    filePath = JoinPath(outputDir, fileName)
    If SaveModel(swApp, model, filePath) Then
        Build_SM_001_OptionalGuardPanel = filePath
        AppendLog fileName & " saved"
    Else
        Build_SM_001_OptionalGuardPanel = ""
        AppendLog fileName & " failed"
    End If

    swApp.CloseDoc model.GetTitle
    Exit Function

Fail:
    AppendLog "Build_SM_001_OptionalGuardPanel failed: " & Err.Description
    On Error Resume Next
    If Not model Is Nothing Then swApp.CloseDoc model.GetTitle
    Build_SM_001_OptionalGuardPanel = ""
End Function