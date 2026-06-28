Option Explicit

Public Function Build_CUS_022_GuideBearingBlock(ByVal swApp As Object, ByVal outputDir As String) As String
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

    partNo = "CUS-022"
    partName = "Guide bearing block 7x6x5 cm"
    fileName = "CUS-022_GuideBearingBlock_70x60x50.SLDPRT"
    lengthMm = 70#
    widthMm = 60#
    heightMm = 50#
    quantity = 1#
    materialName = "TO_BE_CONFIRMED"
    photoReference = "29"
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
        Build_CUS_022_GuideBearingBlock = filePath
        AppendLog fileName & " saved"
    Else
        Build_CUS_022_GuideBearingBlock = ""
        AppendLog fileName & " failed"
    End If

    swApp.CloseDoc model.GetTitle
    Exit Function

Fail:
    AppendLog "Build_CUS_022_GuideBearingBlock failed: " & Err.Description
    On Error Resume Next
    If Not model Is Nothing Then swApp.CloseDoc model.GetTitle
    Build_CUS_022_GuideBearingBlock = ""
End Function