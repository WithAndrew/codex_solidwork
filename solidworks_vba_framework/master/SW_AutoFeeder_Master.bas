Option Explicit

Public Sub Main()
    On Error GoTo EH

    Dim swApp As Object
    Dim partPaths As Object
    Dim assemblyPath As String
    Dim doneMsg As String

    EnsureProjectFolders
    AppendLog "Modular macro started"

    Set swApp = GetSolidWorksApp()
    If swApp Is Nothing Then
        AppendLog "SolidWorks application not found. Macro stopped."
        MsgBox "SolidWorks application not found.", vbCritical, PROJECT_NAME
        Exit Sub
    End If

    Set partPaths = CreateObject("Scripting.Dictionary")

    RegisterBuiltPart partPaths, "PUR-001", "PUR-001_LeftBowl.SLDPRT", Build_PUR_001_LeftBowl(swApp, PartDir())
    RegisterBuiltPart partPaths, "PUR-002", "PUR-002_RightBowl.SLDPRT", Build_PUR_002_RightBowl(swApp, PartDir())
    RegisterBuiltPart partPaths, "PUR-003", "PUR-003_LeftLinearFeeder.SLDPRT", Build_PUR_003_LeftLinearFeeder(swApp, PartDir())
    RegisterBuiltPart partPaths, "PUR-004", "PUR-004_RightLinearFeeder.SLDPRT", Build_PUR_004_RightLinearFeeder(swApp, PartDir())
    RegisterBuiltPart partPaths, "PUR-005", "PUR-005_MainCylinder.SLDPRT", Build_PUR_005_MainCylinder(swApp, PartDir())
    RegisterBuiltPart partPaths, "PUR-006", "PUR-006_StopperCylinder.SLDPRT", Build_PUR_006_StopperCylinder(swApp, PartDir())
    RegisterBuiltPart partPaths, "STD-001", "STD-001_GuideRodBearingSet.SLDPRT", Build_STD_001_GuideRodBearingSet(swApp, PartDir())
    RegisterBuiltPart partPaths, "FRM-001", "FRM-001_MachineFrameEnvelope.SLDPRT", Build_FRM_001_MachineFrameEnvelope(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-001", "CUS-001_CenterBasePlate_700x600x20.SLDPRT", Build_CUS_001_CenterBasePlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-002", "CUS-002_FeedSideRail_150x70x20.SLDPRT", Build_CUS_002_FeedSideRail(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-003", "CUS-003_TopGuideBlock_70x50x20.SLDPRT", Build_CUS_003_TopGuideBlock(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-004", "CUS-004_SmallGuideBlock_50x40x20.SLDPRT", Build_CUS_004_SmallGuideBlock(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-005", "CUS-005_VerticalStopPlate_140x70x20.SLDPRT", Build_CUS_005_VerticalStopPlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-006", "CUS-006_ClampBlock_100x70x20.SLDPRT", Build_CUS_006_ClampBlock(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-007", "CUS-007_LongSlideBase_240x70x20.SLDPRT", Build_CUS_007_LongSlideBase(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-008", "CUS-008_UpperBridgePlate_250x50x20.SLDPRT", Build_CUS_008_UpperBridgePlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-009", "CUS-009_ShortPressPlate_150x50x20.SLDPRT", Build_CUS_009_ShortPressPlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-010", "CUS-010_TallSidePlate_300x70x20.SLDPRT", Build_CUS_010_TallSidePlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-011", "CUS-011_EntryTransitionBlock_100x70x20.SLDPRT", Build_CUS_011_EntryTransitionBlock(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-012", "CUS-012_CylinderMountPlate_280x70x20.SLDPRT", Build_CUS_012_CylinderMountPlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-013", "CUS-013_SidePostPlate_60x50x20.SLDPRT", Build_CUS_013_SidePostPlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-014", "CUS-014_VerticalNarrowPlate_100x40x10.SLDPRT", Build_CUS_014_VerticalNarrowPlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-015", "CUS-015_SmallMountPlate_50x50x20.SLDPRT", Build_CUS_015_SmallMountPlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-016", "CUS-016_SmallClampPlate_50x25x10.SLDPRT", Build_CUS_016_SmallClampPlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-017", "CUS-017_PusherTallBlock_210x70x20.SLDPRT", Build_CUS_017_PusherTallBlock(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-018", "CUS-018_EndSupportBlock_60x60x50.SLDPRT", Build_CUS_018_EndSupportBlock(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-019", "CUS-019_FootClampBlock_70x70x20.SLDPRT", Build_CUS_019_FootClampBlock(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-020", "CUS-020_SlideTopBlock_145x80x20.SLDPRT", Build_CUS_020_SlideTopBlock(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-021", "CUS-021_GuideRodClamp_60x40x40.SLDPRT", Build_CUS_021_GuideRodClamp(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-022", "CUS-022_GuideBearingBlock_70x60x50.SLDPRT", Build_CUS_022_GuideBearingBlock(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-023", "CUS-023_LongGuideSupport_165x70x50.SLDPRT", Build_CUS_023_LongGuideSupport(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-024", "CUS-024_EndCylinderPlate_250x70x20.SLDPRT", Build_CUS_024_EndCylinderPlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-025", "CUS-025_RodSeatBlock_D22.SLDPRT", Build_CUS_025_RodSeatBlock_D22(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-026", "CUS-026_PositionPlate_70x50x20.SLDPRT", Build_CUS_026_PositionPlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-027", "CUS-027_LongLimitRail_300x70x30.SLDPRT", Build_CUS_027_LongLimitRail(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-028", "CUS-028_SensorLightVerticalPlate_310x60x20.SLDPRT", Build_CUS_028_SensorLightVerticalPlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-029", "CUS-029_BackPanel_1080x30x20.SLDPRT", Build_CUS_029_BackPanel(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-030", "CUS-030_LongVerticalSupport_250x80x20.SLDPRT", Build_CUS_030_LongVerticalSupport(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-031", "CUS-031_SensorAngleBlock_100x80x20.SLDPRT", Build_CUS_031_SensorAngleBlock(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-032", "CUS-032_TubeClamp_70x60x20.SLDPRT", Build_CUS_032_TubeClamp(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-033", "CUS-033_SmallAdjuster_80x60x20.SLDPRT", Build_CUS_033_SmallAdjuster(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-034", "CUS-034_LinearFeederInsertPlate_500x70x20.SLDPRT", Build_CUS_034_LinearFeederInsertPlate(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-035", "CUS-035_StopperCylinderBracket_120x60x20.SLDPRT", Build_CUS_035_StopperCylinderBracket(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-036", "CUS-036_MainPusherStrokeRail_300x60x20.SLDPRT", Build_CUS_036_MainPusherStrokeRail(swApp, PartDir())
    RegisterBuiltPart partPaths, "CUS-037", "CUS-037_MainPusherRodSupport_100x70x50.SLDPRT", Build_CUS_037_MainPusherRodSupport(swApp, PartDir())
    RegisterBuiltPart partPaths, "SM-001", "SM-001_OptionalGuardPanel_900x6x420.SLDPRT", Build_SM_001_OptionalGuardPanel(swApp, PartDir())

    WriteBuildBom partPaths
    assemblyPath = GenerateAutoFeederAssembly(swApp, partPaths)
    GenerateDrawingPlaceholders swApp, partPaths, assemblyPath

    AppendLog "Modular macro completed"
    doneMsg = "Modular parts, assembly and drawing placeholders generated." & vbCrLf & "Output: " & OutputDir()
    MsgBox doneMsg, vbInformation, PROJECT_NAME
    Exit Sub

EH:
    AppendLog "Modular macro failed: " & Err.Description
    MsgBox "Modular macro failed: " & Err.Description, vbCritical, PROJECT_NAME
End Sub

Private Sub RegisterBuiltPart(ByVal partPaths As Object, ByVal partNo As String, ByVal expectedFileName As String, ByVal builtPath As String)
    partPaths(partNo) = builtPath
    If Len(builtPath) = 0 Then
        AppendLog expectedFileName & " build failed"
    Else
        AppendLog expectedFileName & " registered: " & builtPath
    End If
End Sub

Private Sub WriteBuildBom(ByVal partPaths As Object)
    Dim lines As String
    Dim key As Variant
    Dim statusText As String
    Dim pathText As String

    lines = "PartNo,GeneratedPath,Status" & vbCrLf
    For Each key In partPaths.Keys
        pathText = CStr(partPaths(key))
        If Len(pathText) > 0 Then
            statusText = "GENERATED"
        Else
            statusText = "FAILED"
        End If
        lines = lines & CsvCell(CStr(key)) & "," & CsvCell(pathText) & "," & CsvCell(statusText) & vbCrLf
    Next key

    WriteTextFile JoinPath(BomDir(), "Modular_Build_BOM.csv"), lines
End Sub