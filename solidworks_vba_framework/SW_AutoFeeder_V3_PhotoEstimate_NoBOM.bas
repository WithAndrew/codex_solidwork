Option Explicit

Private Const PROJECT_NAME As String = "AutoFeeder Photo Estimate V3"
Private Const PROJECT_ROOT As String = "D:\working\Temp\6.27"
Private Const STATUS_PHOTO_ESTIMATE As String = "PHOTO_ESTIMATE_VERIFY_BEFORE_MACHINING"
Private Const MM_TO_M As Double = 0.001

Public Sub Main()
    On Error GoTo EH

    Dim rows As Variant
    Dim doneMsg As String

    EnsureProjectFolders
    AppendLog "Macro V3 started"

    rows = BuildPartRows()
    WritePhotoDimensionRegister rows
    GeneratePhotoEstimateParts rows
    GenerateV3SubAssemblies rows
    GenerateV3TopAssembly rows
    GenerateV3Drawings rows

    AppendLog "Macro V3 completed"
    doneMsg = "V3 photo-estimate parts, assemblies and drawing placeholders generated."
    doneMsg = doneMsg & vbCrLf & "Output: " & OutputDir()
    MsgBox doneMsg, vbInformation, PROJECT_NAME
    Exit Sub

EH:
    MsgBox "Macro V3 failed: " & Err.Description, vbCritical, PROJECT_NAME
End Sub

Private Function BuildPartRows() As Variant
    Dim rows(0 To 45) As Variant

    rows(0) = Array("PUR-001", "PUR-001_LeftBowl.SLDPRT", "Left bowl feeder simplified envelope", 1, 480#, 480#, 280#, -520#, 0#, 160#, "53,54", "Purchased")
    rows(1) = Array("PUR-002", "PUR-002_RightBowl.SLDPRT", "Right bowl feeder simplified envelope", 1, 480#, 480#, 280#, 520#, 0#, 160#, "53,54", "Purchased")
    rows(2) = Array("PUR-003", "PUR-003_LeftLinearFeeder.SLDPRT", "Left linear feeder body", 1, 520#, 80#, 55#, -250#, 130#, 225#, "41,42,44", "Purchased")
    rows(3) = Array("PUR-004", "PUR-004_RightLinearFeeder.SLDPRT", "Right linear feeder body", 1, 520#, 80#, 55#, 250#, 130#, 225#, "41,42,44", "Purchased")
    rows(4) = Array("PUR-005", "PUR-005_MainCylinder.SLDPRT", "Main pusher cylinder envelope", 1, 420#, 70#, 70#, 0#, -250#, 210#, "48,49,50,52", "Purchased")
    rows(5) = Array("PUR-006", "PUR-006_StopperCylinder.SLDPRT", "Stopper indexing cylinder envelope", 1, 160#, 45#, 45#, 210#, 25#, 245#, "45,46,47", "Purchased")
    rows(6) = Array("STD-001", "STD-001_GuideRodBearingSet.SLDPRT", "Guide rod and bearing envelope", 1, 360#, 45#, 35#, 0#, -120#, 165#, "28-34", "Standard")
    rows(7) = Array("FRM-001", "FRM-001_MachineFrameEnvelope.SLDPRT", "Machine frame envelope estimated from overview", 1, 1500#, 850#, 60#, 0#, 0#, 30#, "53,54", "Weldment")
    rows(8) = Array("CUS-001", "CUS-001_CenterBasePlate_700x600x20.SLDPRT", "Center base plate from 70x60 cm mark", 1, 700#, 600#, 20#, 0#, 120#, 130#, "7", "Machined")
    rows(9) = Array("CUS-002", "CUS-002_FeedSideRail_150x70x20.SLDPRT", "Feed side rail 15x7x2 cm", 2, 150#, 70#, 20#, -125#, 105#, 170#, "1", "Machined")
    rows(10) = Array("CUS-003", "CUS-003_TopGuideBlock_70x50x20.SLDPRT", "Top guide block 7x5x2 cm", 1, 70#, 50#, 20#, -70#, 115#, 205#, "2", "Machined")
    rows(11) = Array("CUS-004", "CUS-004_SmallGuideBlock_50x40x20.SLDPRT", "Small guide block 5x4x2 cm", 2, 50#, 40#, 20#, 75#, 115#, 205#, "3", "Machined")
    rows(12) = Array("CUS-005", "CUS-005_VerticalStopPlate_140x70x20.SLDPRT", "Vertical stop plate 14x7x2 cm", 2, 140#, 70#, 20#, -170#, 70#, 225#, "4", "Machined")
    rows(13) = Array("CUS-006", "CUS-006_ClampBlock_100x70x20.SLDPRT", "Clamp block 10x7x2 cm", 2, 100#, 70#, 20#, 30#, 75#, 225#, "5", "Machined")
    rows(14) = Array("CUS-007", "CUS-007_LongSlideBase_240x70x20.SLDPRT", "Long slide base 24x7x2 cm", 2, 240#, 70#, 20#, 0#, 45#, 145#, "6", "Machined")
    rows(15) = Array("CUS-008", "CUS-008_UpperBridgePlate_250x50x20.SLDPRT", "Upper bridge plate 25x5x2 cm", 1, 250#, 50#, 20#, 0#, 185#, 205#, "8,9", "Machined")
    rows(16) = Array("CUS-009", "CUS-009_ShortPressPlate_150x50x20.SLDPRT", "Short press plate 15x5x2 cm", 2, 150#, 50#, 20#, -60#, 175#, 230#, "10", "Machined")
    rows(17) = Array("CUS-010", "CUS-010_TallSidePlate_300x70x20.SLDPRT", "Tall side plate 30x7x2 cm", 1, 300#, 70#, 20#, 95#, 180#, 230#, "11", "Machined")
    rows(18) = Array("CUS-011", "CUS-011_EntryTransitionBlock_100x70x20.SLDPRT", "Entry transition block 10x7x2 cm", 1, 100#, 70#, 20#, -210#, 160#, 185#, "13,14", "Machined")
    rows(19) = Array("CUS-012", "CUS-012_CylinderMountPlate_280x70x20.SLDPRT", "Cylinder mount plate 28x7x2 cm", 1, 280#, 70#, 20#, 205#, 100#, 220#, "15", "Machined")
    rows(20) = Array("CUS-013", "CUS-013_SidePostPlate_60x50x20.SLDPRT", "Side post plate 6x5x2 cm", 1, 60#, 50#, 20#, 280#, 120#, 260#, "17", "Machined")
    rows(21) = Array("CUS-014", "CUS-014_VerticalNarrowPlate_100x40x10.SLDPRT", "Vertical narrow plate 10x4x1 cm", 1, 100#, 40#, 10#, 315#, 140#, 275#, "19", "Machined")
    rows(22) = Array("CUS-015", "CUS-015_SmallMountPlate_50x50x20.SLDPRT", "Small mount plate 5x5x2 cm", 1, 50#, 50#, 20#, 350#, 130#, 245#, "20", "Machined")
    rows(23) = Array("CUS-016", "CUS-016_SmallClampPlate_50x25x10.SLDPRT", "Small clamp plate 5x2.5x1 cm", 2, 50#, 25#, 10#, 360#, 100#, 255#, "21", "Machined")
    rows(24) = Array("CUS-017", "CUS-017_PusherTallBlock_210x70x20.SLDPRT", "Pusher tall block 21x7x2 cm", 1, 210#, 70#, 20#, -110#, -15#, 230#, "23", "Machined")
    rows(25) = Array("CUS-018", "CUS-018_EndSupportBlock_60x60x50.SLDPRT", "End support block 6x6x5 cm", 1, 60#, 60#, 50#, -235#, -45#, 215#, "24", "Machined")
    rows(26) = Array("CUS-019", "CUS-019_FootClampBlock_70x70x20.SLDPRT", "Foot clamp block 7x7x2 cm", 1, 70#, 70#, 20#, -285#, -80#, 160#, "25", "Machined")
    rows(27) = Array("CUS-020", "CUS-020_SlideTopBlock_145x80x20.SLDPRT", "Slide top block 14.5x8x2 cm", 1, 145#, 80#, 20#, 0#, -20#, 220#, "27", "Machined")
    rows(28) = Array("CUS-021", "CUS-021_GuideRodClamp_60x40x40.SLDPRT", "Guide rod clamp estimated from 6 cm mark", 1, 60#, 40#, 40#, -175#, -130#, 170#, "28", "Machined")
    rows(29) = Array("CUS-022", "CUS-022_GuideBearingBlock_70x60x50.SLDPRT", "Guide bearing block 7x6x5 cm", 1, 70#, 60#, 50#, -95#, -145#, 170#, "29", "Machined")
    rows(30) = Array("CUS-023", "CUS-023_LongGuideSupport_165x70x50.SLDPRT", "Long guide support 16.5x7x5 cm", 1, 165#, 70#, 50#, 45#, -145#, 170#, "30", "Machined")
    rows(31) = Array("CUS-024", "CUS-024_EndCylinderPlate_250x70x20.SLDPRT", "End cylinder plate 25x7x2 cm", 1, 250#, 70#, 20#, 210#, -145#, 185#, "31", "Machined")
    rows(32) = Array("CUS-025", "CUS-025_RodSeatBlock_D22.SLDPRT", "Rod seat block for 22 diameter rod mark", 1, 90#, 70#, 60#, 280#, -115#, 190#, "32", "Machined")
    rows(33) = Array("CUS-026", "CUS-026_PositionPlate_70x50x20.SLDPRT", "Position plate 7x5x2 cm", 4, 70#, 50#, 20#, 150#, -50#, 200#, "33", "Machined")
    rows(34) = Array("CUS-027", "CUS-027_LongLimitRail_300x70x30.SLDPRT", "Long limit rail 30x7x3 cm", 1, 300#, 70#, 30#, 0#, -205#, 165#, "34", "Machined")
    rows(35) = Array("CUS-028", "CUS-028_SensorLightVerticalPlate_310x60x20.SLDPRT", "Sensor light vertical plate 31x6x2 cm", 1, 310#, 60#, 20#, 0#, 285#, 300#, "35", "Machined")
    rows(36) = Array("CUS-029", "CUS-029_BackPanel_1080x30x20.SLDPRT", "Back panel estimated from 108x3 cm mark", 1, 1080#, 30#, 20#, 0#, 360#, 300#, "36", "Machined")
    rows(37) = Array("CUS-030", "CUS-030_LongVerticalSupport_250x80x20.SLDPRT", "Long vertical support 25x8x2 cm", 1, 250#, 80#, 20#, -320#, 250#, 260#, "37", "Machined")
    rows(38) = Array("CUS-031", "CUS-031_SensorAngleBlock_100x80x20.SLDPRT", "Sensor angle block 10x8x2 cm", 1, 100#, 80#, 20#, -260#, 295#, 255#, "38", "Machined")
    rows(39) = Array("CUS-032", "CUS-032_TubeClamp_70x60x20.SLDPRT", "Tube clamp 7x6x2 cm", 1, 70#, 60#, 20#, -210#, 300#, 235#, "39", "Machined")
    rows(40) = Array("CUS-033", "CUS-033_SmallAdjuster_80x60x20.SLDPRT", "Small adjuster 8x6x2 cm", 1, 80#, 60#, 20#, -150#, 300#, 235#, "40", "Machined")
    rows(41) = Array("CUS-034", "CUS-034_LinearFeederInsertPlate_500x70x20.SLDPRT", "Linear feeder insertion plate, 5 cm insertion allowance", 2, 500#, 70#, 20#, 0#, 155#, 190#, "41,42,44", "Machined")
    rows(42) = Array("CUS-035", "CUS-035_StopperCylinderBracket_120x60x20.SLDPRT", "Stopper cylinder bracket, stroke marks 2.5/0.5/7.5", 1, 120#, 60#, 20#, 245#, 25#, 260#, "45,46,47", "Machined")
    rows(43) = Array("CUS-036", "CUS-036_MainPusherStrokeRail_300x60x20.SLDPRT", "Main pusher stroke rail, 7.5 and 5 cm marks", 1, 300#, 60#, 20#, 0#, -285#, 190#, "48,49,50", "Machined")
    rows(44) = Array("CUS-037", "CUS-037_MainPusherRodSupport_100x70x50.SLDPRT", "Main pusher rod support, 50 stroke mark to verify", 1, 100#, 70#, 50#, -260#, -285#, 185#, "51,52", "Machined")
    rows(45) = Array("SM-001", "SM-001_OptionalGuardPanel_900x6x420.SLDPRT", "Optional transparent guard panel", 1, 900#, 6#, 420#, 0#, 430#, 260#, "53,54", "SheetMetal")

    BuildPartRows = rows
End Function

Private Sub GeneratePhotoEstimateParts(ByVal rows As Variant)
    Dim swApp As Object
    Dim i As Long
    Dim r As Variant
    Dim ok As Boolean

    Set swApp = GetSolidWorksApp()
    If swApp Is Nothing Then
        AppendLog "SolidWorks application not found. Parts skipped."
        Exit Sub
    End If

    For i = LBound(rows) To UBound(rows)
        r = rows(i)
        If Left$(CStr(r(0)), 3) = "PUR" And InStr(1, CStr(r(1)), "Bowl", vbTextCompare) > 0 Then
            ok = CreateCylinderPart(swApp, JoinPath(PartDir(), CStr(r(1))), CDbl(r(4)), CDbl(r(6)), CStr(r(2)), CStr(r(0)), CStr(r(10)))
        Else
            ok = CreateBoxPart(swApp, JoinPath(PartDir(), CStr(r(1))), CDbl(r(4)), CDbl(r(5)), CDbl(r(6)), CStr(r(2)), CStr(r(0)), CStr(r(10)))
        End If

        If ok Then
            AppendLog CStr(r(1)) & " saved"
        Else
            AppendLog CStr(r(1)) & " failed"
        End If
    Next i
End Sub

Private Sub GenerateV3SubAssemblies(ByVal rows As Variant)
    GenerateGroupAssembly rows, "ASM-100_CenterFeedAndIndexing_V3.SLDASM", "CUS-001,CUS-002,CUS-003,CUS-004,CUS-005,CUS-006,CUS-007,CUS-008,CUS-009,CUS-010,CUS-011,CUS-012,CUS-017,CUS-018,CUS-020,CUS-034"
    GenerateGroupAssembly rows, "ASM-200_GuideAndPusher_V3.SLDASM", "PUR-005,STD-001,CUS-021,CUS-022,CUS-023,CUS-024,CUS-025,CUS-026,CUS-027,CUS-036,CUS-037"
    GenerateGroupAssembly rows, "ASM-300_StopperSensorBracket_V3.SLDASM", "PUR-006,CUS-013,CUS-014,CUS-015,CUS-016,CUS-028,CUS-029,CUS-030,CUS-031,CUS-032,CUS-033,CUS-035"
    GenerateGroupAssembly rows, "ASM-400_FeederBowlLayout_V3.SLDASM", "PUR-001,PUR-002,PUR-003,PUR-004,FRM-001,SM-001"
End Sub

Private Sub GenerateGroupAssembly(ByVal rows As Variant, ByVal asmName As String, ByVal partCodesCsv As String)
    On Error GoTo EH

    Dim swApp As Object
    Dim model As Object
    Dim assy As Object
    Dim i As Long
    Dim r As Variant
    Dim errors As Long
    Dim warnings As Long

    Set swApp = GetSolidWorksApp()
    If swApp Is Nothing Then Exit Sub

    Set model = swApp.NewAssembly
    If model Is Nothing Then Exit Sub
    Set assy = model

    For i = LBound(rows) To UBound(rows)
        r = rows(i)
        If InStr(1, "," & partCodesCsv & ",", "," & CStr(r(0)) & ",", vbTextCompare) > 0 Then
            AddAssemblyComponent assy, JoinPath(PartDir(), CStr(r(1))), CDbl(r(7)), CDbl(r(8)), CDbl(r(9))
        End If
    Next i

    model.ViewZoomtofit2
    model.Extension.SaveAs JoinPath(AssemblyDir(), asmName), 0, 1, Nothing, errors, warnings
    AppendLog asmName & " saved"
    Exit Sub

EH:
    AppendLog asmName & " failed: " & Err.Description
End Sub

Private Sub GenerateV3TopAssembly(ByVal rows As Variant)
    On Error GoTo EH

    Dim swApp As Object
    Dim model As Object
    Dim assy As Object
    Dim i As Long
    Dim r As Variant
    Dim errors As Long
    Dim warnings As Long

    Set swApp = GetSolidWorksApp()
    If swApp Is Nothing Then
        AppendLog "SolidWorks application not found. Top assembly skipped."
        Exit Sub
    End If

    Set model = swApp.NewAssembly
    If model Is Nothing Then
        AppendLog "NewAssembly failed. Top assembly skipped."
        Exit Sub
    End If
    Set assy = model

    For i = LBound(rows) To UBound(rows)
        r = rows(i)
        AddAssemblyComponent assy, JoinPath(PartDir(), CStr(r(1))), CDbl(r(7)), CDbl(r(8)), CDbl(r(9))
    Next i

    model.ViewZoomtofit2
    model.Extension.SaveAs JoinPath(AssemblyDir(), "ASM-000_AutoFeeder_PhotoEstimate_V3.SLDASM"), 0, 1, Nothing, errors, warnings
    AppendLog "ASM-000_AutoFeeder_PhotoEstimate_V3.SLDASM saved"
    Exit Sub

EH:
    AppendLog "GenerateV3TopAssembly failed: " & Err.Description
End Sub

Private Sub GenerateV3Drawings(ByVal rows As Variant)
    Dim swApp As Object
    Dim i As Long
    Dim r As Variant
    Dim made As Long

    Set swApp = GetSolidWorksApp()
    If swApp Is Nothing Then Exit Sub

    For i = LBound(rows) To UBound(rows)
        r = rows(i)
        If CreateSimpleDrawing(swApp, JoinPath(PartDir(), CStr(r(1))), JoinPath(DrawingDir(), Replace(CStr(r(1)), ".SLDPRT", ".SLDDRW")), CStr(r(2))) Then
            made = made + 1
        End If
    Next i

    CreateSimpleDrawing swApp, JoinPath(AssemblyDir(), "ASM-000_AutoFeeder_PhotoEstimate_V3.SLDASM"), JoinPath(DrawingDir(), "ASM-000_AutoFeeder_PhotoEstimate_V3.SLDDRW"), "Top assembly layout drawing placeholder"
    AppendLog "Drawing placeholders generated: " & CStr(made)
End Sub

Private Function CreateSimpleDrawing(ByVal swApp As Object, ByVal modelPath As String, ByVal drawingPath As String, ByVal titleText As String) As Boolean
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

Private Function CreateBoxPart(ByVal swApp As Object, ByVal filePath As String, ByVal lengthMm As Double, ByVal widthMm As Double, ByVal heightMm As Double, ByVal titleText As String, ByVal partNo As String, ByVal photoRef As String) As Boolean
    On Error GoTo EH

    Dim model As Object
    Dim errors As Long
    Dim warnings As Long

    Set model = swApp.NewPart
    If model Is Nothing Then GoTo EH

    model.Extension.SelectByID2 "Front Plane", "PLANE", 0#, 0#, 0#, False, 0, Nothing, 0
    model.SketchManager.InsertSketch True
    model.SketchManager.CreateCenterRectangle 0#, 0#, 0#, Mm(lengthMm / 2#), Mm(widthMm / 2#), 0#
    model.SketchManager.InsertSketch True
    model.FeatureManager.FeatureExtrusion2 True, False, False, 0, 0, Mm(heightMm), 0#, False, False, False, False, 0#, 0#, False, False, False, False, True, True, True, 0, 0, False

    AddPartProperties model, titleText, partNo, photoRef, lengthMm, widthMm, heightMm
    model.ViewZoomtofit2
    CreateBoxPart = model.Extension.SaveAs(filePath, 0, 1, Nothing, errors, warnings)
    swApp.CloseDoc model.GetTitle
    Exit Function

EH:
    CreateBoxPart = False
End Function

Private Function CreateCylinderPart(ByVal swApp As Object, ByVal filePath As String, ByVal diameterMm As Double, ByVal heightMm As Double, ByVal titleText As String, ByVal partNo As String, ByVal photoRef As String) As Boolean
    On Error GoTo EH

    Dim model As Object
    Dim errors As Long
    Dim warnings As Long

    Set model = swApp.NewPart
    If model Is Nothing Then GoTo EH

    model.Extension.SelectByID2 "Top Plane", "PLANE", 0#, 0#, 0#, False, 0, Nothing, 0
    model.SketchManager.InsertSketch True
    model.SketchManager.CreateCircleByRadius 0#, 0#, 0#, Mm(diameterMm / 2#)
    model.SketchManager.InsertSketch True
    model.FeatureManager.FeatureExtrusion2 True, False, False, 0, 0, Mm(heightMm), 0#, False, False, False, False, 0#, 0#, False, False, False, False, True, True, True, 0, 0, False

    AddPartProperties model, titleText, partNo, photoRef, diameterMm, diameterMm, heightMm
    model.ViewZoomtofit2
    CreateCylinderPart = model.Extension.SaveAs(filePath, 0, 1, Nothing, errors, warnings)
    swApp.CloseDoc model.GetTitle
    Exit Function

EH:
    CreateCylinderPart = False
End Function

Private Sub AddPartProperties(ByVal model As Object, ByVal titleText As String, ByVal partNo As String, ByVal photoRef As String, ByVal lengthMm As Double, ByVal widthMm As Double, ByVal heightMm As Double)
    On Error Resume Next
    model.CustomInfo2("", "PartNo") = partNo
    model.CustomInfo2("", "PartName") = titleText
    model.CustomInfo2("", "PhotoReference") = photoRef
    model.CustomInfo2("", "DimensionSource") = STATUS_PHOTO_ESTIMATE
    model.CustomInfo2("", "Estimated_L_mm") = CStr(lengthMm)
    model.CustomInfo2("", "Estimated_W_mm") = CStr(widthMm)
    model.CustomInfo2("", "Estimated_H_mm") = CStr(heightMm)
    On Error GoTo 0
End Sub

Private Function AddAssemblyComponent(ByVal assy As Object, ByVal componentPath As String, ByVal xMm As Double, ByVal yMm As Double, ByVal zMm As Double) As Object
    On Error Resume Next
    Set AddAssemblyComponent = assy.AddComponent5(componentPath, 0, "", False, "", Mm(xMm), Mm(yMm), Mm(zMm))
    If AddAssemblyComponent Is Nothing Then
        Set AddAssemblyComponent = assy.AddComponent4(componentPath, "", Mm(xMm), Mm(yMm), Mm(zMm))
    End If
    On Error GoTo 0
End Function

Private Function GetSolidWorksApp() As Object
    On Error Resume Next
    Set GetSolidWorksApp = Application.SldWorks
    If GetSolidWorksApp Is Nothing Then
        Set GetSolidWorksApp = GetObject(, "SldWorks.Application")
    End If
    On Error GoTo 0
End Function

Private Function Mm(ByVal valueMm As Double) As Double
    Mm = valueMm * MM_TO_M
End Function

Private Sub WritePhotoDimensionRegister(ByVal rows As Variant)
    Dim lines As String
    Dim i As Long
    Dim r As Variant

    lines = "PartNo,FileName,Description,Qty,L_mm,W_mm,H_mm,AsmX_mm,AsmY_mm,AsmZ_mm,PhotoReference,Category,Status" & vbCrLf
    For i = LBound(rows) To UBound(rows)
        r = rows(i)
        lines = lines & CsvCell(CStr(r(0))) & "," & CsvCell(CStr(r(1))) & "," & CsvCell(CStr(r(2))) & "," & CStr(r(3)) & "," & CStr(r(4)) & "," & CStr(r(5)) & "," & CStr(r(6)) & "," & CStr(r(7)) & "," & CStr(r(8)) & "," & CStr(r(9)) & "," & CsvCell(CStr(r(10))) & "," & CsvCell(CStr(r(11))) & "," & CsvCell(STATUS_PHOTO_ESTIMATE) & vbCrLf
    Next i

    WriteTextFile JoinPath(BomDir(), "V3_Photo_Estimated_Dimensions.csv"), lines
End Sub

Private Function JoinPath(ByVal leftPath As String, ByVal rightPath As String) As String
    If Len(leftPath) = 0 Then
        JoinPath = rightPath
    ElseIf Right$(leftPath, 1) = "\" Then
        JoinPath = leftPath & rightPath
    Else
        JoinPath = leftPath & "\" & rightPath
    End If
End Function

Private Sub EnsureProjectFolders()
    EnsureFolder OutputDir()
    EnsureFolder PartDir()
    EnsureFolder AssemblyDir()
    EnsureFolder DrawingDir()
    EnsureFolder ExportDir()
    EnsureFolder BomDir()
    EnsureFolder LogDir()
End Sub

Private Sub EnsureFolder(ByVal folderPath As String)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If
End Sub

Private Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    Dim fileNo As Integer
    fileNo = FreeFile
    Open filePath For Output As #fileNo
    Print #fileNo, content
    Close #fileNo
End Sub

Private Function CsvCell(ByVal value As String) As String
    CsvCell = """" & Replace(value, """", """""") & """"
End Function

Private Sub AppendLog(ByVal message As String)
    Dim fileNo As Integer
    Dim logPath As String
    logPath = JoinPath(LogDir(), "macro_run_v3.log")
    fileNo = FreeFile
    Open logPath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & message
    Close #fileNo
End Sub

Private Function OutputDir() As String
    OutputDir = JoinPath(PROJECT_ROOT, "SW_Output")
End Function

Private Function PartDir() As String
    PartDir = JoinPath(OutputDir(), "01_Parts")
End Function

Private Function AssemblyDir() As String
    AssemblyDir = JoinPath(OutputDir(), "02_Assemblies")
End Function

Private Function DrawingDir() As String
    DrawingDir = JoinPath(OutputDir(), "03_Drawings")
End Function

Private Function ExportDir() As String
    ExportDir = JoinPath(OutputDir(), "04_Exports")
End Function

Private Function BomDir() As String
    BomDir = JoinPath(OutputDir(), "05_BOM")
End Function

Private Function LogDir() As String
    LogDir = JoinPath(OutputDir(), "99_Logs")
End Function
