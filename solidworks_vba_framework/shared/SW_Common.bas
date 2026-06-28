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