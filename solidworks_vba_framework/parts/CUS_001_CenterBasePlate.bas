'
' @Author: WithAndrew 1032286603@qq.com
' @Date: 2026-06-28 14:33:48
' @LastEditors: WithAndrew 1032286603@qq.com
' @LastEditTime: 2026-06-28 15:41:34
' @FilePath: \6.27\solidworks_vba_framework\parts\CUS_001_CenterBasePlate.bas
' @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
'
Option Explicit

Public Function Build_CUS_001_CenterBasePlate(ByVal swApp As Object, ByVal outputDir As String) As String
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

    partNo = "CUS-001"
    partName = "Center base plate from 70x60 cm mark"
    fileName = "CUS-001_CenterBasePlate_700x600x20.SLDPRT"
    lengthMm = 700#
    widthMm = 600#
    heightMm = 20#
    quantity = 1#
    materialName = "TO_BE_CONFIRMED"
    photoReference = "7"
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
        Build_CUS_001_CenterBasePlate = filePath
        AppendLog fileName & " saved"
    Else
        Build_CUS_001_CenterBasePlate = ""
        AppendLog fileName & " failed"
    End If

    swApp.CloseDoc model.GetTitle
    Exit Function

Fail:
    AppendLog "Build_CUS_001_CenterBasePlate failed: " & Err.Description
    On Error Resume Next
    If Not model Is Nothing Then swApp.CloseDoc model.GetTitle
    Build_CUS_001_CenterBasePlate = ""
End Function