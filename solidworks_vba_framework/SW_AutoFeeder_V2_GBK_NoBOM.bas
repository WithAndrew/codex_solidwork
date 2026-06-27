Option Explicit

Private Const PROJECT_NAME As String = "双振动盘自动上料分料设备"
Private Const STATUS_PENDING As String = "待实测确认"
Private Const PROJECT_ROOT As String = "D:\working\Temp\6.27"
Private Const MM_TO_M As Double = 0.001

Public Sub Main()
    On Error GoTo EH

    Dim msg As String

    EnsureProjectFolders
    AppendLog "Macro V2 started: " & PROJECT_NAME

    WriteParameterTemplate
    WriteBomTemplate
    GenerateAllParts
    GenerateTopAssembly
    GenerateAllDrawings
    ExportAllDeliverables

    AppendLog "Macro V2 completed"
    msg = "SolidWorks VBA V2占位布局宏已运行。" & vbCrLf
    msg = msg & "输出目录：" & OutputDir() & vbCrLf
    msg = msg & "关键尺寸状态：" & STATUS_PENDING
    MsgBox msg, vbInformation, PROJECT_NAME
    Exit Sub

EH:
    MsgBox "宏运行失败：" & Err.Description, vbCritical, PROJECT_NAME
End Sub

Private Function TemplateDir() As String
    TemplateDir = JoinPath(PROJECT_ROOT, "solidworks_vba_framework\templates")
End Function

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

Private Sub EnsureProjectFolders()
    EnsureFolder OutputDir()
    EnsureFolder PartDir()
    EnsureFolder AssemblyDir()
    EnsureFolder DrawingDir()
    EnsureFolder ExportDir()
    EnsureFolder BomDir()
    EnsureFolder LogDir()
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

    logPath = JoinPath(LogDir(), "macro_run.log")
    fileNo = FreeFile
    Open logPath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & message
    Close #fileNo
End Sub

Private Sub WriteParameterTemplate()
    Dim lines As String

    lines = "ParamName,ParamName_CN,Value,Unit,CriticalDimensionStatus,PhotoReference,Remark" & vbCrLf
    AddParamLine lines, "Product_L", "产品长度", "", "mm", "", "产品实际长度"
    AddParamLine lines, "Product_W", "产品宽度", "", "mm", "", "产品实际宽度"
    AddParamLine lines, "Product_H", "产品高度", "", "mm", "", "产品实际高度"
    AddParamLine lines, "Product_M", "单件重量", "", "g", "", "产品单件重量"
    AddParamLine lines, "Bowl_D", "振动盘直径", "", "mm", "53,54", "左右振动盘分别确认"
    AddParamLine lines, "Rail_L", "直线送料轨道长度", "", "mm", "41,42,44", "左右直振分别确认"
    AddParamLine lines, "Rail_W", "送料槽宽度", "", "mm", "41,42,44", "与产品接触尺寸"
    AddParamLine lines, "Rail_H", "送料槽高度", "", "mm", "41,42,44", "含槽深和盖板间隙"
    AddParamLine lines, "Stroke_Main", "主推料气缸行程", "", "mm", "49,50,52", "照片标注仅作线索"
    AddParamLine lines, "Stroke_Stopper", "挡料/错位气缸行程", "", "mm", "45,46,47", "照片标注仅作线索"
    AddParamLine lines, "Overall_L", "整机长度", "", "mm", "53,54", "整机外形尺寸"
    AddParamLine lines, "Overall_W", "整机宽度", "", "mm", "53,54", "整机外形尺寸"
    AddParamLine lines, "Overall_H", "整机高度", "", "mm", "53,54", "整机外形尺寸"
    AddParamLine lines, "Plate_T", "主要板件厚度", "", "mm", "1-52", "各板件需逐件确认"
    AddParamLine lines, "Hole_D", "安装孔直径", "", "mm", "1-52", "孔径、孔距、沉孔、螺纹孔需逐项确认"
    AddParamLine lines, "Rail_Insert_Depth", "直振插入中间机构深度", "", "mm", "41,42,44", "照片中约5公分仅作线索"

    WriteTextFile JoinPath(BomDir(), "Parameters_Template.csv"), lines
End Sub

Private Sub AddParamLine(ByRef lines As String, ByVal nameEN As String, ByVal nameCN As String, ByVal valueText As String, ByVal unitText As String, ByVal photoReference As String, ByVal remark As String)
    lines = lines & CsvCell(nameEN) & "," & CsvCell(nameCN) & "," & CsvCell(valueText) & "," & CsvCell(unitText) & "," & CsvCell(STATUS_PENDING) & "," & CsvCell(photoReference) & "," & CsvCell(remark) & vbCrLf
End Sub

Private Sub WriteBomTemplate()
    Dim lines As String

    lines = "Item,PartNo,PartName_CN,Qty,Material,Process,SurfaceTreatment,StandardOrCustom,DrawingRequired,PhotoReference,CriticalDimensionStatus,Remark" & vbCrLf
    AddBomLine lines, 1, "ASM-000", "总装", 1, "-", "装配", "-", "Custom", "No", "1-54", PROJECT_NAME
    AddBomLine lines, 2, "PUR-001", "左振动盘", 1, "SUS304", "外购", "原厂", "Purchased", "No", "53,54", "简化外购件模型"
    AddBomLine lines, 3, "PUR-002", "右振动盘", 1, "SUS304", "外购", "原厂", "Purchased", "No", "53,54", "简化外购件模型"
    AddBomLine lines, 4, "PUR-003", "左直振送料器", 1, "-", "外购", "原厂", "Purchased", "No", "41,42,44", "安装孔位待确认"
    AddBomLine lines, 5, "PUR-004", "右直振送料器", 1, "-", "外购", "原厂", "Purchased", "No", "41,42,44", "安装孔位待确认"
    AddBomLine lines, 6, "CUS-001", "中心安装底板", 1, "Q235/45#", "机加工", "发黑/镀镍/喷塑待定", "Custom", "Yes", "7-15,41-44", "基准件"
    AddBomLine lines, 7, "CUS-002", "中心送料槽块", 1, "SUS304/45#", "机加工", "抛光/镀硬铬待定", "Custom", "Yes", "41,42,44", "产品接触件"
    AddBomLine lines, 8, "CUS-003", "送料槽上压板", 2, "SUS304/45#", "机加工", "抛光/发黑待定", "Custom", "Yes", "1-6,41-44", "与产品间隙待确认"
    AddBomLine lines, 9, "CUS-004", "侧向限位块", 2, "SUS304/45#", "机加工", "抛光/发黑待定", "Custom", "Yes", "1-6,12-15", "左右可能不完全相同"
    AddBomLine lines, 10, "CUS-005", "挡料块", 2, "SKD11/SUS304", "机加工", "淬火/发黑待定", "Custom", "Yes", "45-47", "产品接触件"
    AddBomLine lines, 11, "CUS-006", "错位滑块", 1, "SKD11/45#", "机加工", "淬火/发黑待定", "Custom", "Yes", "45-47", "需建立伸缩配置"
    AddBomLine lines, 12, "CUS-007", "主推料推块", 1, "SKD11/SUS304", "机加工", "淬火/镀硬铬待定", "Custom", "Yes", "29,49,50,52", "产品接触件"
    AddBomLine lines, 13, "CUS-008", "主气缸安装座", 1, "Q235/45#", "机加工", "发黑/镀镍待定", "Custom", "Yes", "49-52", "气缸型号待确认"
    AddBomLine lines, 14, "CUS-009", "导杆支撑块", 2, "45#/6061", "机加工", "发黑/本色氧化待定", "Custom", "Yes", "28-34", "导杆直径和孔距待确认"
    AddBomLine lines, 15, "CUS-010", "机械限位块", 2, "45#/SUS304", "机加工", "发黑待定", "Custom", "Yes", "28-34", "限位位置待确认"
    AddBomLine lines, 16, "WLD-001", "机架焊接件", 1, "Q235方管", "焊接", "喷塑", "Custom", "Yes", "53,54", "也可改4040/4080铝型材"
    AddBomLine lines, 17, "PUR-005", "主推料气缸", 1, "-", "外购", "原厂", "Purchased", "No", "49,50,52", "缸径、行程、安装方式待确认"
    AddBomLine lines, 18, "PUR-006", "挡料/错位气缸", 1, "-", "外购", "原厂", "Purchased", "No", "45-47", "缸径、行程、安装方式待确认"
    AddBomLine lines, 19, "STD-001", "导杆/直线轴承", 1, "标准件", "外购", "原厂", "Standard", "No", "28-34", "规格待确认"
    AddBomLine lines, 20, "CUS-011", "传感器/光源支架", 1, "6061/Q235", "机加工/钣金", "氧化/喷塑待定", "Custom", "Yes", "35-40,53,54", "检测位置待确认"

    WriteTextFile JoinPath(BomDir(), "BOM_Template.csv"), lines
End Sub

Private Sub AddBomLine(ByRef lines As String, ByVal itemNo As Long, ByVal partNo As String, ByVal partNameCN As String, ByVal qty As Double, ByVal material As String, ByVal processText As String, ByVal surfaceText As String, ByVal standardOrCustom As String, ByVal drawingRequired As String, ByVal photoReference As String, ByVal remark As String)
    lines = lines & CStr(itemNo) & "," & CsvCell(partNo) & "," & CsvCell(partNameCN) & "," & CStr(qty) & "," & CsvCell(material) & "," & CsvCell(processText) & "," & CsvCell(surfaceText) & "," & CsvCell(standardOrCustom) & "," & CsvCell(drawingRequired) & "," & CsvCell(photoReference) & "," & CsvCell(STATUS_PENDING) & "," & CsvCell(remark) & vbCrLf
End Sub

Private Sub GenerateAllParts()
    Dim swApp As Object

    AppendLog "GenerateAllParts started"
    Set swApp = GetSolidWorksApp()
    If swApp Is Nothing Then
        AppendLog "SolidWorks application not found. Part placeholders skipped."
        Exit Sub
    End If

    SaveCylinder swApp, "PUR-001_LeftBowl.SLDPRT", 450, 260, "左振动盘-占位"
    SaveCylinder swApp, "PUR-002_RightBowl.SLDPRT", 450, 260, "右振动盘-占位"
    SaveBox swApp, "PUR-003_LeftLinearFeeder.SLDPRT", 520, 80, 55, "左直振送料器-占位"
    SaveBox swApp, "PUR-004_RightLinearFeeder.SLDPRT", 520, 80, 55, "右直振送料器-占位"
    SaveBox swApp, "PUR-005_MainCylinder.SLDPRT", 260, 70, 70, "主推料气缸-占位"
    SaveBox swApp, "PUR-006_StopperCylinder.SLDPRT", 120, 45, 45, "挡料错位气缸-占位"
    SaveBox swApp, "STD-001_GuideRodBearingSet.SLDPRT", 360, 45, 35, "导杆直线轴承组-占位"
    SaveBox swApp, "CUS-001_CenterBasePlate.SLDPRT", 620, 260, 20, "中心安装底板-占位"
    SaveBox swApp, "CUS-002_CenterFeedGroove.SLDPRT", 360, 55, 28, "中心送料槽块-占位"
    SaveBox swApp, "CUS-003_TopCoverPlate.SLDPRT", 280, 35, 12, "送料槽上压板-占位"
    SaveBox swApp, "CUS-004_SideGuideBlock.SLDPRT", 260, 28, 28, "侧向限位块-占位"
    SaveBox swApp, "CUS-005_StopperBlock.SLDPRT", 55, 25, 35, "挡料块-占位"
    SaveBox swApp, "CUS-006_IndexingSlide.SLDPRT", 100, 55, 30, "错位滑块-占位"
    SaveBox swApp, "CUS-007_MainPusher.SLDPRT", 85, 45, 35, "主推料推块-占位"
    SaveBox swApp, "CUS-008_MainCylinderMount.SLDPRT", 140, 95, 25, "主气缸安装座-占位"
    SaveBox swApp, "CUS-009_GuideRodSupport.SLDPRT", 70, 60, 50, "导杆支撑块-占位"
    SaveBox swApp, "CUS-010_MechanicalStop.SLDPRT", 55, 35, 35, "机械限位块-占位"
    SaveBox swApp, "CUS-011_SensorLightBracket.SLDPRT", 260, 30, 180, "传感器光源支架-占位"
    SaveBox swApp, "SM-001_OptionalGuardPanel.SLDPRT", 900, 6, 420, "透明防护罩板-占位"
    SaveBox swApp, "WLD-001_FrameEnvelope.SLDPRT", 1500, 850, 60, "机架外形包络-占位"

    AppendLog "GenerateAllParts completed"
End Sub

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

Private Sub SaveBox(ByVal swApp As Object, ByVal fileName As String, ByVal lMm As Double, ByVal wMm As Double, ByVal hMm As Double, ByVal titleText As String)
    Dim ok As Boolean

    ok = CreateBoxPart(swApp, JoinPath(PartDir(), fileName), lMm, wMm, hMm, titleText)
    If ok Then
        AppendLog fileName & " saved"
    Else
        AppendLog fileName & " failed"
    End If
End Sub

Private Sub SaveCylinder(ByVal swApp As Object, ByVal fileName As String, ByVal dMm As Double, ByVal hMm As Double, ByVal titleText As String)
    Dim ok As Boolean

    ok = CreateCylinderPart(swApp, JoinPath(PartDir(), fileName), dMm, hMm, titleText)
    If ok Then
        AppendLog fileName & " saved"
    Else
        AppendLog fileName & " failed"
    End If
End Sub

Private Function CreateBoxPart(ByVal swApp As Object, ByVal filePath As String, ByVal lengthMm As Double, ByVal widthMm As Double, ByVal heightMm As Double, ByVal titleText As String) As Boolean
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

    AddPendingCustomProperties model, titleText
    model.ViewZoomtofit2
    CreateBoxPart = model.Extension.SaveAs(filePath, 0, 1, Nothing, errors, warnings)
    swApp.CloseDoc model.GetTitle
    Exit Function

EH:
    CreateBoxPart = False
End Function

Private Function CreateCylinderPart(ByVal swApp As Object, ByVal filePath As String, ByVal diameterMm As Double, ByVal heightMm As Double, ByVal titleText As String) As Boolean
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

    AddPendingCustomProperties model, titleText
    model.ViewZoomtofit2
    CreateCylinderPart = model.Extension.SaveAs(filePath, 0, 1, Nothing, errors, warnings)
    swApp.CloseDoc model.GetTitle
    Exit Function

EH:
    CreateCylinderPart = False
End Function

Private Sub AddPendingCustomProperties(ByVal model As Object, ByVal titleText As String)
    On Error Resume Next
    model.CustomInfo2("", "PartName_CN") = titleText
    model.CustomInfo2("", "CriticalDimensionStatus") = STATUS_PENDING
    model.CustomInfo2("", "Remark") = "V2参数化占位模型，仅用于布局和测绘辅助，不能直接加工"
    On Error GoTo 0
End Sub

Private Sub GenerateTopAssembly()
    On Error GoTo EH

    Dim swApp As Object
    Dim model As Object
    Dim assy As Object
    Dim errors As Long
    Dim warnings As Long
    Dim assyPath As String

    AppendLog "GenerateTopAssembly started"

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

    AddAssemblyComponent assy, JoinPath(PartDir(), "WLD-001_FrameEnvelope.SLDPRT"), 0, 0, -60
    AddAssemblyComponent assy, JoinPath(PartDir(), "PUR-001_LeftBowl.SLDPRT"), -480, 0, 0
    AddAssemblyComponent assy, JoinPath(PartDir(), "PUR-002_RightBowl.SLDPRT"), 480, 0, 0
    AddAssemblyComponent assy, JoinPath(PartDir(), "PUR-003_LeftLinearFeeder.SLDPRT"), -250, 130, 165
    AddAssemblyComponent assy, JoinPath(PartDir(), "PUR-004_RightLinearFeeder.SLDPRT"), 250, 130, 165
    AddAssemblyComponent assy, JoinPath(PartDir(), "CUS-001_CenterBasePlate.SLDPRT"), 0, 130, 120
    AddAssemblyComponent assy, JoinPath(PartDir(), "CUS-002_CenterFeedGroove.SLDPRT"), 0, 130, 155
    AddAssemblyComponent assy, JoinPath(PartDir(), "CUS-003_TopCoverPlate.SLDPRT"), 0, 130, 195
    AddAssemblyComponent assy, JoinPath(PartDir(), "CUS-004_SideGuideBlock.SLDPRT"), 0, 90, 180
    AddAssemblyComponent assy, JoinPath(PartDir(), "CUS-005_StopperBlock.SLDPRT"), -60, 70, 205
    AddAssemblyComponent assy, JoinPath(PartDir(), "CUS-006_IndexingSlide.SLDPRT"), 80, 70, 180
    AddAssemblyComponent assy, JoinPath(PartDir(), "PUR-005_MainCylinder.SLDPRT"), 0, -170, 160
    AddAssemblyComponent assy, JoinPath(PartDir(), "CUS-007_MainPusher.SLDPRT"), 0, -40, 180
    AddAssemblyComponent assy, JoinPath(PartDir(), "CUS-008_MainCylinderMount.SLDPRT"), 0, -250, 145
    AddAssemblyComponent assy, JoinPath(PartDir(), "STD-001_GuideRodBearingSet.SLDPRT"), 0, -110, 125
    AddAssemblyComponent assy, JoinPath(PartDir(), "PUR-006_StopperCylinder.SLDPRT"), 190, 30, 200
    AddAssemblyComponent assy, JoinPath(PartDir(), "CUS-011_SensorLightBracket.SLDPRT"), 0, 280, 250
    AddAssemblyComponent assy, JoinPath(PartDir(), "SM-001_OptionalGuardPanel.SLDPRT"), 0, 430, 260

    model.ViewZoomtofit2
    assyPath = JoinPath(AssemblyDir(), "ASM-000_AutoFeeder_Layout_V2.SLDASM")
    model.Extension.SaveAs assyPath, 0, 1, Nothing, errors, warnings
    AppendLog "Top assembly saved: " & assyPath
    Exit Sub

EH:
    AppendLog "GenerateTopAssembly failed: " & Err.Description
End Sub

Private Function AddAssemblyComponent(ByVal assy As Object, ByVal componentPath As String, ByVal xMm As Double, ByVal yMm As Double, ByVal zMm As Double) As Object
    On Error Resume Next
    Set AddAssemblyComponent = assy.AddComponent5(componentPath, 0, "", False, "", Mm(xMm), Mm(yMm), Mm(zMm))
    If AddAssemblyComponent Is Nothing Then
        Set AddAssemblyComponent = assy.AddComponent4(componentPath, "", Mm(xMm), Mm(yMm), Mm(zMm))
    End If
    On Error GoTo 0
End Function

Private Sub GenerateAllDrawings()
    AppendLog "GenerateAllDrawings: framework placeholder"
    AppendLog "Drawing views, dimensions and technical notes pending confirmed part geometry"
End Sub

Private Sub ExportAllDeliverables()
    AppendLog "ExportAllDeliverables: framework placeholder"
    AppendLog "STEP/PDF/DXF/DWG export pending generated SolidWorks documents"
End Sub

