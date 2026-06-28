Option Explicit

Public Function Build_CUS_021_GuideRodClamp(ByVal swApp As Object, ByVal outputDir As String) As String
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

    partNo = "CUS-021"
    partName = "Guide rod clamp estimated from 6 cm mark"
    fileName = "CUS-021_GuideRodClamp_60x40x40.SLDPRT"
    lengthMm = 60#
    widthMm = 40#
    heightMm = 40#
    quantity = 1#
    materialName = "TO_BE_CONFIRMED"
    photoReference = "28"
    categoryName = "Machined"

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
        Build_CUS_021_GuideRodClamp = filePath
        AppendLog fileName & " saved"
    Else
        Build_CUS_021_GuideRodClamp = ""
        AppendLog fileName & " failed"
    End If

    swApp.CloseDoc model.GetTitle
    Exit Function

Fail:
    AppendLog "Build_CUS_021_GuideRodClamp failed: " & Err.Description
    On Error Resume Next
    If Not model Is Nothing Then swApp.CloseDoc model.GetTitle
    Build_CUS_021_GuideRodClamp = ""
End Function