$ErrorActionPreference = 'Stop'

$root = 'D:\working\Temp\6.27'
$framework = Join-Path $root 'solidworks_vba_framework'
$srcPath = Join-Path $framework 'SW_AutoFeeder_V3_PhotoEstimate_NoBOM.bas'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Write-NoBomFile([string]$Path, [string]$Content) {
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Parse-VbArrayArgs([string]$s) {
    $items = New-Object System.Collections.Generic.List[string]
    $buf = New-Object System.Text.StringBuilder
    $inQuote = $false

    for ($i = 0; $i -lt $s.Length; $i++) {
        $ch = $s[$i]
        if ($ch -eq '"') {
            if ($inQuote -and $i + 1 -lt $s.Length -and $s[$i + 1] -eq '"') {
                [void]$buf.Append('"')
                $i++
            } else {
                $inQuote = -not $inQuote
            }
        } elseif ($ch -eq ',' -and -not $inQuote) {
            $items.Add($buf.ToString().Trim())
            [void]$buf.Clear()
        } else {
            [void]$buf.Append($ch)
        }
    }
    $items.Add($buf.ToString().Trim())

    $out = @()
    foreach ($item in $items) {
        $v = $item.Trim()
        if ($v.StartsWith('"') -and $v.EndsWith('"')) {
            $v = $v.Substring(1, $v.Length - 2)
        }
        $v = $v -replace '#$', ''
        $out += $v
    }
    return ,$out
}

function VbStr([string]$s) {
    return '"' + ($s -replace '"', '""') + '"'
}

function VbNum([string]$s) {
    return ($s.TrimEnd('#') + '#')
}

function Stem-FromFileName([string]$fileName) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    return ($base -replace '_\d+(?:\.\d+)?x\d+(?:\.\d+)?x\d+(?:\.\d+)?$', '')
}

function FunctionName-FromStem([string]$stem) {
    $fn = $stem -replace '[^A-Za-z0-9]+', '_'
    return 'Build_' + $fn.Trim('_')
}

function ModuleName-FromStem([string]$stem) {
    $name = $stem -replace '[^A-Za-z0-9_]+', '_'
    $name = $name -replace 'Vertical', 'Vert'
    if ($name.Length -gt 31) {
        $name = $name.Substring(0, 31)
    }
    return $name.Trim('_')
}

$src = Get-Content $srcPath -Raw
$rows = New-Object System.Collections.Generic.List[object]

foreach ($line in ($src -split "`r?`n")) {
    if ($line -match 'rows\((\d+)\)\s*=\s*Array\((.*)\)') {
        $args = Parse-VbArrayArgs $matches[2]
        $stem = Stem-FromFileName $args[1]
        $rows.Add([pscustomobject]@{
            Index = [int]$matches[1]
            PartNo = $args[0]
            FileName = $args[1]
            Description = $args[2]
            Qty = $args[3]
            L = $args[4]
            W = $args[5]
            H = $args[6]
            X = $args[7]
            Y = $args[8]
            Z = $args[9]
            PhotoRef = $args[10]
            Category = $args[11]
            Stem = $stem
            FunctionName = FunctionName-FromStem $stem
        }) | Out-Null
    }
}

if ($rows.Count -eq 0) {
    throw 'No rows parsed from V3 macro.'
}

foreach ($d in @('master', 'shared', 'parts', 'assembly')) {
    New-Item -ItemType Directory -Path (Join-Path $framework $d) -Force | Out-Null
}

$common = @'
Option Explicit

Public Const PROJECT_NAME As String = "AutoFeeder Modular VBA"
Public Const PROJECT_ROOT As String = "D:\working\Temp\6.27"
Public Const STATUS_PHOTO_ESTIMATE As String = "PHOTO_ESTIMATE_VERIFY_BEFORE_MACHINING"
Public Const MM_TO_M As Double = 0.001

Public Function GetSolidWorksApp() As Object
    On Error Resume Next
    Set GetSolidWorksApp = Application.SldWorks
    If GetSolidWorksApp Is Nothing Then
        Set GetSolidWorksApp = GetObject(, "SldWorks.Application")
    End If
    On Error GoTo 0
End Function

Public Function Mm(ByVal valueMm As Double) As Double
    Mm = valueMm * MM_TO_M
End Function

Public Function JoinPath(ByVal leftPath As String, ByVal rightPath As String) As String
    If Len(leftPath) = 0 Then
        JoinPath = rightPath
    ElseIf Right$(leftPath, 1) = "\" Then
        JoinPath = leftPath & rightPath
    Else
        JoinPath = leftPath & "\" & rightPath
    End If
End Function

Public Sub EnsureProjectFolders()
    EnsureFolder OutputDir()
    EnsureFolder PartDir()
    EnsureFolder AssemblyDir()
    EnsureFolder DrawingDir()
    EnsureFolder ExportDir()
    EnsureFolder BomDir()
    EnsureFolder LogDir()
End Sub

Public Sub EnsureFolder(ByVal folderPath As String)
    Dim parentPath As String

    If Len(folderPath) = 0 Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    parentPath = Left$(folderPath, InStrRev(folderPath, "\") - 1)
    If Len(parentPath) > 0 And Len(Dir$(parentPath, vbDirectory)) = 0 Then
        EnsureFolder parentPath
    End If
    MkDir folderPath
End Sub

Public Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    Dim fileNo As Integer
    fileNo = FreeFile
    Open filePath For Output As #fileNo
    Print #fileNo, content
    Close #fileNo
End Sub

Public Function CsvCell(ByVal value As String) As String
    CsvCell = """" & Replace(value, """", """""") & """"
End Function

Public Sub AppendLog(ByVal message As String)
    Dim fileNo As Integer
    Dim logPath As String

    EnsureFolder LogDir()
    logPath = JoinPath(LogDir(), "macro_run_modular.log")
    fileNo = FreeFile
    Open logPath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & message
    Close #fileNo
End Sub

Public Function OutputDir() As String
    OutputDir = JoinPath(PROJECT_ROOT, "SW_Output")
End Function

Public Function PartDir() As String
    PartDir = JoinPath(OutputDir(), "01_Parts")
End Function

Public Function AssemblyDir() As String
    AssemblyDir = JoinPath(OutputDir(), "02_Assemblies")
End Function

Public Function DrawingDir() As String
    DrawingDir = JoinPath(OutputDir(), "03_Drawings")
End Function

Public Function ExportDir() As String
    ExportDir = JoinPath(OutputDir(), "04_Exports")
End Function

Public Function BomDir() As String
    BomDir = JoinPath(OutputDir(), "05_BOM")
End Function

Public Function LogDir() As String
    LogDir = JoinPath(OutputDir(), "99_Logs")
End Function
'@
Write-NoBomFile (Join-Path $framework 'shared\SW_Common.bas') $common

$propertyHelpers = @'
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
'@
Write-NoBomFile (Join-Path $framework 'shared\SW_PropertyHelpers.bas') $propertyHelpers

$featureHelpers = @'
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
'@
Write-NoBomFile (Join-Path $framework 'shared\SW_FeatureHelpers.bas') $featureHelpers

$assemblyHelpers = @'
Option Explicit

Public Function NewAssembly(ByVal swApp As Object) As Object
    On Error Resume Next
    Set NewAssembly = swApp.NewAssembly
    On Error GoTo 0
End Function

Public Function AddAssemblyComponent(ByVal assy As Object, ByVal componentPath As String, ByVal xMm As Double, ByVal yMm As Double, ByVal zMm As Double) As Object
    On Error Resume Next
    Set AddAssemblyComponent = assy.AddComponent5(componentPath, 0, "", False, "", Mm(xMm), Mm(yMm), Mm(zMm))
    If AddAssemblyComponent Is Nothing Then
        Set AddAssemblyComponent = assy.AddComponent4(componentPath, "", Mm(xMm), Mm(yMm), Mm(zMm))
    End If
    On Error GoTo 0
End Function

Public Function SaveAssembly(ByVal model As Object, ByVal filePath As String) As Boolean
    On Error GoTo EH

    Dim errors As Long
    Dim warnings As Long

    model.ViewZoomtofit2
    SaveAssembly = model.Extension.SaveAs(filePath, 0, 1, Nothing, errors, warnings)
    If Not SaveAssembly Then AppendLog "SaveAssembly failed: " & filePath & " errors=" & CStr(errors) & " warnings=" & CStr(warnings)
    Exit Function

EH:
    AppendLog "SaveAssembly exception: " & filePath & " - " & Err.Description
    SaveAssembly = False
End Function
'@
Write-NoBomFile (Join-Path $framework 'shared\SW_AssemblyHelpers.bas') $assemblyHelpers

$drawingHelpers = @'
Option Explicit

Public Sub GenerateDrawingPlaceholders(ByVal swApp As Object, ByVal partPaths As Object, ByVal assemblyPath As String)
    On Error GoTo EH

    Dim key As Variant
    Dim partPath As String
    Dim drawingPath As String
    Dim made As Long

    If swApp Is Nothing Then Exit Sub

    For Each key In partPaths.Keys
        partPath = CStr(partPaths(key))
        If Len(partPath) > 0 Then
            drawingPath = JoinPath(DrawingDir(), Replace(GetFileName(partPath), ".SLDPRT", ".SLDDRW"))
            If CreateSimpleDrawing(swApp, partPath, drawingPath, CStr(key)) Then made = made + 1
        End If
    Next key

    If Len(assemblyPath) > 0 Then
        CreateSimpleDrawing swApp, assemblyPath, JoinPath(DrawingDir(), "ASM-000_AutoFeeder_Modular.SLDDRW"), "Top assembly layout drawing placeholder"
    End If

    AppendLog "Drawing placeholders generated: " & CStr(made)
    Exit Sub

EH:
    AppendLog "GenerateDrawingPlaceholders failed: " & Err.Description
End Sub

Public Function CreateSimpleDrawing(ByVal swApp As Object, ByVal modelPath As String, ByVal drawingPath As String, ByVal titleText As String) As Boolean
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

Private Function GetFileName(ByVal filePath As String) As String
    GetFileName = Mid$(filePath, InStrRev(filePath, "\") + 1)
End Function
'@
Write-NoBomFile (Join-Path $framework 'shared\SW_DrawingHelpers.bas') $drawingHelpers

foreach ($r in $rows) {
    $moduleName = ModuleName-FromStem $r.Stem
    $partPath = Join-Path (Join-Path $framework 'parts') ($moduleName + '.bas')
    $builder = $r.FunctionName
    $isCylinder = ($r.PartNo -like 'PUR-*' -and $r.FileName -match 'Bowl')
    $baseCall = if ($isCylinder) { '    CreateCylinder model, lengthMm, heightMm' } else { '    CreateBox model, lengthMm, widthMm, heightMm' }
    $content = @"
Option Explicit

Public Function $builder(ByVal swApp As Object, ByVal outputDir As String) As String
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

    partNo = $(VbStr $r.PartNo)
    partName = $(VbStr $r.Description)
    fileName = $(VbStr $r.FileName)
    lengthMm = $(VbNum $r.L)
    widthMm = $(VbNum $r.W)
    heightMm = $(VbNum $r.H)
    quantity = $(VbNum $r.Qty)
    materialName = "TO_BE_CONFIRMED"
    photoReference = $(VbStr $r.PhotoRef)
    categoryName = $(VbStr $r.Category)

    ' Create base body
    Set model = NewPart(swApp)
    If model Is Nothing Then GoTo Fail
$baseCall

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
        $builder = filePath
        AppendLog fileName & " saved"
    Else
        $builder = ""
        AppendLog fileName & " failed"
    End If

    swApp.CloseDoc model.GetTitle
    Exit Function

Fail:
    AppendLog "$builder failed: " & Err.Description
    On Error Resume Next
    If Not model Is Nothing Then swApp.CloseDoc model.GetTitle
    $builder = ""
End Function
"@
    Write-NoBomFile $partPath $content
}

$placeLines = New-Object System.Text.StringBuilder
foreach ($r in $rows) {
    [void]$placeLines.AppendLine("    PlacePart assy, partPaths, " + (VbStr $r.PartNo) + ", " + (VbNum $r.X) + ", " + (VbNum $r.Y) + ", " + (VbNum $r.Z))
}

$assembly = @"
Option Explicit

Public Function GenerateAutoFeederAssembly(ByVal swApp As Object, ByVal partPaths As Object) As String
    On Error GoTo EH

    Dim model As Object
    Dim assy As Object
    Dim asmPath As String

    If swApp Is Nothing Then
        AppendLog "SolidWorks application not found. Assembly skipped."
        GenerateAutoFeederAssembly = ""
        Exit Function
    End If

    Set model = NewAssembly(swApp)
    If model Is Nothing Then
        AppendLog "NewAssembly failed. Assembly skipped."
        GenerateAutoFeederAssembly = ""
        Exit Function
    End If
    Set assy = model

$($placeLines.ToString())
    asmPath = JoinPath(AssemblyDir(), "ASM-000_AutoFeeder_Modular.SLDASM")
    If SaveAssembly(model, asmPath) Then
        GenerateAutoFeederAssembly = asmPath
        AppendLog "ASM-000_AutoFeeder_Modular.SLDASM saved"
    Else
        GenerateAutoFeederAssembly = ""
    End If
    Exit Function

EH:
    AppendLog "GenerateAutoFeederAssembly failed: " & Err.Description
    GenerateAutoFeederAssembly = ""
End Function

Private Sub PlacePart(ByVal assy As Object, ByVal partPaths As Object, ByVal partNo As String, ByVal xMm As Double, ByVal yMm As Double, ByVal zMm As Double)
    Dim component As Object

    If Not partPaths.Exists(partNo) Then
        AppendLog "Assembly skipped missing part key: " & partNo
        Exit Sub
    End If
    If Len(CStr(partPaths(partNo))) = 0 Then
        AppendLog "Assembly skipped failed part: " & partNo
        Exit Sub
    End If

    Set component = AddAssemblyComponent(assy, CStr(partPaths(partNo)), xMm, yMm, zMm)
    If component Is Nothing Then
        AppendLog "Assembly insert failed: " & partNo
    End If
End Sub
"@
Write-NoBomFile (Join-Path $framework 'assembly\SW_AutoFeeder_AssemblyLayout.bas') $assembly

$callLines = New-Object System.Text.StringBuilder
foreach ($r in $rows) {
    [void]$callLines.AppendLine("    RegisterBuiltPart partPaths, " + (VbStr $r.PartNo) + ", " + (VbStr $r.FileName) + ", " + $r.FunctionName + "(swApp, PartDir())")
}

$master = @"
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

$($callLines.ToString())
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
"@
Write-NoBomFile (Join-Path $framework 'master\SW_AutoFeeder_Master.bas') $master

$migrationLines = New-Object System.Text.StringBuilder
foreach ($r in $rows) {
    $moduleName = ModuleName-FromStem $r.Stem
    [void]$migrationLines.AppendLine("- " + $r.PartNo + " -> solidworks_vba_framework\parts\" + $moduleName + ".bas")
}

$doc = @"
# Modular SolidWorks VBA Macro Structure

This folder contains the refactored modular source layout. The old V2/V3 macros are kept for reference; the new workflow imports the shared modules, all part modules, the assembly layout module, and the master module into a clean SolidWorks macro project.

## Import Order

1. ``shared\SW_Common.bas``
2. ``shared\SW_PropertyHelpers.bas``
3. ``shared\SW_FeatureHelpers.bas``
4. ``shared\SW_AssemblyHelpers.bas``
5. ``shared\SW_DrawingHelpers.bas``
6. every file under ``parts\*.bas``
7. ``assembly\SW_AutoFeeder_AssemblyLayout.bas``
8. ``master\SW_AutoFeeder_Master.bas``

Run ``Main`` from ``SW_AutoFeeder_Master.bas``.

## Ownership Rule

- Part geometry and machining features live in each part's own VBA file under ``parts``.
- The master macro only calls ``Build_xxx`` functions and records generated paths.
- The assembly macro only places generated parts by coordinate.
- CSV output is only a run report, not the source of dimensions.

## Outputs

- ``SW_Output\01_Parts``
- ``SW_Output\02_Assemblies\ASM-000_AutoFeeder_Modular.SLDASM``
- ``SW_Output\03_Drawings``
- ``SW_Output\05_BOM\Modular_Build_BOM.csv``
- ``SW_Output\99_Logs\macro_run_modular.log``

## Migration Map

$($migrationLines.ToString())
"@
Write-NoBomFile (Join-Path $framework 'MODULAR_REFACTOR_README.md') $doc

Write-Host "Generated modular VBA structure with $($rows.Count) part modules."
