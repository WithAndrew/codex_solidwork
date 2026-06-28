Option Explicit

Public Function Build_PUR_005_MainCylinder(ByVal swApp As Object, ByVal outputDir As String) As String
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

    partNo = "PUR-005"
    partName = "Main pusher cylinder envelope"
    fileName = "PUR-005_MainCylinder.SLDPRT"
    lengthMm = 420#
    widthMm = 70#
    heightMm = 70#
    quantity = 1#
    materialName = "TO_BE_CONFIRMED"
    photoReference = "48,49,50,52"
    categoryName = "Purchased"

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
        Build_PUR_005_MainCylinder = filePath
        AppendLog fileName & " saved"
    Else
        Build_PUR_005_MainCylinder = ""
        AppendLog fileName & " failed"
    End If

    swApp.CloseDoc model.GetTitle
    Exit Function

Fail:
    AppendLog "Build_PUR_005_MainCylinder failed: " & Err.Description
    On Error Resume Next
    If Not model Is Nothing Then swApp.CloseDoc model.GetTitle
    Build_PUR_005_MainCylinder = ""
End Function