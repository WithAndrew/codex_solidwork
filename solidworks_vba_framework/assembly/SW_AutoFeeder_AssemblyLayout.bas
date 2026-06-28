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

    PlacePart assy, partPaths, "PUR-001", -520#, 0#, 160#
    PlacePart assy, partPaths, "PUR-002", 520#, 0#, 160#
    PlacePart assy, partPaths, "PUR-003", -250#, 130#, 225#
    PlacePart assy, partPaths, "PUR-004", 250#, 130#, 225#
    PlacePart assy, partPaths, "PUR-005", 0#, -250#, 210#
    PlacePart assy, partPaths, "PUR-006", 210#, 25#, 245#
    PlacePart assy, partPaths, "STD-001", 0#, -120#, 165#
    PlacePart assy, partPaths, "FRM-001", 0#, 0#, 30#
    PlacePart assy, partPaths, "CUS-001", 0#, 120#, 130#
    PlacePart assy, partPaths, "CUS-002", -125#, 105#, 170#
    PlacePart assy, partPaths, "CUS-003", -70#, 115#, 205#
    PlacePart assy, partPaths, "CUS-004", 75#, 115#, 205#
    PlacePart assy, partPaths, "CUS-005", -170#, 70#, 225#
    PlacePart assy, partPaths, "CUS-006", 30#, 75#, 225#
    PlacePart assy, partPaths, "CUS-007", 0#, 45#, 145#
    PlacePart assy, partPaths, "CUS-008", 0#, 185#, 205#
    PlacePart assy, partPaths, "CUS-009", -60#, 175#, 230#
    PlacePart assy, partPaths, "CUS-010", 95#, 180#, 230#
    PlacePart assy, partPaths, "CUS-011", -210#, 160#, 185#
    PlacePart assy, partPaths, "CUS-012", 205#, 100#, 220#
    PlacePart assy, partPaths, "CUS-013", 280#, 120#, 260#
    PlacePart assy, partPaths, "CUS-014", 315#, 140#, 275#
    PlacePart assy, partPaths, "CUS-015", 350#, 130#, 245#
    PlacePart assy, partPaths, "CUS-016", 360#, 100#, 255#
    PlacePart assy, partPaths, "CUS-017", -110#, -15#, 230#
    PlacePart assy, partPaths, "CUS-018", -235#, -45#, 215#
    PlacePart assy, partPaths, "CUS-019", -285#, -80#, 160#
    PlacePart assy, partPaths, "CUS-020", 0#, -20#, 220#
    PlacePart assy, partPaths, "CUS-021", -175#, -130#, 170#
    PlacePart assy, partPaths, "CUS-022", -95#, -145#, 170#
    PlacePart assy, partPaths, "CUS-023", 45#, -145#, 170#
    PlacePart assy, partPaths, "CUS-024", 210#, -145#, 185#
    PlacePart assy, partPaths, "CUS-025", 280#, -115#, 190#
    PlacePart assy, partPaths, "CUS-026", 150#, -50#, 200#
    PlacePart assy, partPaths, "CUS-027", 0#, -205#, 165#
    PlacePart assy, partPaths, "CUS-028", 0#, 285#, 300#
    PlacePart assy, partPaths, "CUS-029", 0#, 360#, 300#
    PlacePart assy, partPaths, "CUS-030", -320#, 250#, 260#
    PlacePart assy, partPaths, "CUS-031", -260#, 295#, 255#
    PlacePart assy, partPaths, "CUS-032", -210#, 300#, 235#
    PlacePart assy, partPaths, "CUS-033", -150#, 300#, 235#
    PlacePart assy, partPaths, "CUS-034", 0#, 155#, 190#
    PlacePart assy, partPaths, "CUS-035", 245#, 25#, 260#
    PlacePart assy, partPaths, "CUS-036", 0#, -285#, 190#
    PlacePart assy, partPaths, "CUS-037", -260#, -285#, 185#
    PlacePart assy, partPaths, "SM-001", 0#, 430#, 260#

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
