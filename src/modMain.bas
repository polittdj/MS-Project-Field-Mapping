Attribute VB_Name = "modMain"
'==============================================================================
' modMain
' ----------------------------------------------------------------------------
' Public entry points for the MS Project Field Mapper. Macros surfaced here
' appear in the user's Alt+F8 macro list inside Microsoft Project.
'
' Phase 0 surface:
'   ShowMain       - opens frmMain (the user-facing control panel)
'   RunDiagnostics - writes a diagnostics .txt without launching the UI
'
' Later phases extend modMain with the full pipeline orchestration.
'==============================================================================
Option Explicit

'==============================================================================
' ShowMain
' ----------------------------------------------------------------------------
' Primary entry point. Validates the host first; on success, shows frmMain.
' Macro is the one users bind to a Quick Access Toolbar button or run via
' Alt+F8.
'==============================================================================
Public Sub ShowMain()
    On Error GoTo Fail

    Dim why As String
    If Not AssertHostSupported(why) Then
        MsgBox "Cannot run: " & why & vbCrLf & vbCrLf & _
               "This tool requires Microsoft Project Desktop 2016 or later. " & _
               "Project for the Web and Project Online are not supported.", _
               vbCritical, APP_NAME
        Exit Sub
    End If

    frmMain.Show
    Exit Sub

Fail:
    MsgBox "Unhandled error in ShowMain: " & Err.Number & " - " & _
           Err.Description, vbCritical, APP_NAME
End Sub

'==============================================================================
' RunDiagnostics
' ----------------------------------------------------------------------------
' Generates a diagnostics report and writes it to a .txt file. Invoked from
' frmMain's Diagnostics button (and available standalone for support cases
' where the UI itself fails to load).
'
' Returns the absolute path of the written file via ByRef so the caller can
' reveal it in Explorer or display it. Returns empty string on failure.
'==============================================================================
Public Function RunDiagnostics(Optional ByVal outputFolder As String = "") As String
    On Error GoTo Fail

    Dim ctx As clsRunContext
    Set ctx = New clsRunContext
    ctx.Init Application
    If LenB(outputFolder) > 0 Then ctx.OutputFolder = outputFolder

    Dim path As String
    path = ctx.DefaultDiagnosticsPath()

    Dim report As String
    report = DiagnosticsReport()

    Dim fnum As Integer
    fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, report
    Close #fnum

    RunDiagnostics = path
    Exit Function

Fail:
    Dim n As Long, d As String
    n = Err.Number: d = Err.Description
    On Error Resume Next
    If fnum <> 0 Then Close #fnum
    On Error GoTo 0
    MsgBox "RunDiagnostics failed: " & n & " - " & d, vbCritical, APP_NAME
    RunDiagnostics = ""
End Function

'==============================================================================
' RunMapper (placeholder)
' ----------------------------------------------------------------------------
' Reserved for the full pipeline orchestration in Phases 2+. Today this is a
' clear "not yet implemented" stub so the macro is visible in the Alt+F8 list
' and users discover it when the UI is wired up.
'==============================================================================
Public Sub RunMapper()
    MsgBox "RunMapper is not implemented yet. " & vbCrLf & _
           "Use ShowMain to launch the UI, RunDiagnostics for a " & _
           "diagnostics dump, or Phase1_TestScan to scan a single .mpp.", _
           vbInformation, APP_NAME
End Sub

'==============================================================================
' Phase1_TestScan
' ----------------------------------------------------------------------------
' Phase 1 verification entry point. Prompts for a .mpp path, runs the
' single-file scan (open/scan/close), writes the result + log to disk, and
' offers to open the output folder.
'
' Macro is exposed so the user can run it via Alt+F8 without touching the
' Immediate Window.
'==============================================================================
Public Sub Phase1_TestScan()
    Dim path As String
    path = InputBox("Path to .mpp file to scan:", APP_NAME)
    If LenB(Trim$(path)) = 0 Then Exit Sub

    Dim outFolder As String
    outFolder = InputBox("Output folder for scan dump (leave blank for %TEMP%):", _
                         APP_NAME)

    Dim resultPath As String
    resultPath = Phase1_ScanToFile(Trim$(path), Trim$(outFolder))

    If LenB(resultPath) = 0 Then Exit Sub

    Dim ans As VbMsgBoxResult
    ans = MsgBox("Scan complete. Results written to:" & vbCrLf & resultPath & _
                 vbCrLf & vbCrLf & "Open the containing folder?", _
                 vbInformation Or vbYesNo, APP_NAME)
    If ans = vbYes Then
        On Error Resume Next
        Shell "explorer.exe /select,""" & resultPath & """", vbNormalFocus
        On Error GoTo 0
    End If
End Sub

'==============================================================================
' Phase1_ScanToFile
' ----------------------------------------------------------------------------
' Programmatic single-file scan. Returns the path of the dump file written,
' or empty string on failure.
'
' Suitable for Immediate-Window invocation:
'     ?modMain.Phase1_ScanToFile("C:\Projects\PlantA.mpp")
'
' Output file contains:
'   - the clsProjectScan.ToDebugString dump (field-by-field listing)
'   - the full session log appended at the end
'==============================================================================
Public Function Phase1_ScanToFile(ByVal filePath As String, _
                                  Optional ByVal outputFolder As String = "") _
                                  As String
    On Error GoTo Fail

    Dim ctx As clsRunContext
    Set ctx = New clsRunContext
    ctx.Init Application
    If LenB(outputFolder) > 0 Then ctx.OutputFolder = outputFolder
    ctx.Logger.MinSeverity = sevDebug
    ctx.Logger.LogPath = ctx.DefaultLogPath()

    Dim scan As clsProjectScan
    Set scan = ScanProjectByPath(filePath, ctx)

    Dim dump As String
    If scan Is Nothing Then
        dump = "SCAN FAILED for: " & filePath & vbCrLf & _
               "(See log entries below for details.)" & vbCrLf
    Else
        dump = scan.ToDebugString()
    End If

    Dim resultPath As String
    Dim folder As String
    If LenB(ctx.OutputFolder) > 0 Then
        folder = ctx.OutputFolder
    Else
        folder = Environ$("TEMP")
        If LenB(folder) = 0 Then folder = "C:\Temp"
        If Right$(folder, 1) <> "\" Then folder = folder & "\"
    End If
    resultPath = folder & "Phase1_Scan_" & ctx.SessionId & EXT_LOG

    Dim fnum As Integer
    fnum = FreeFile
    Open resultPath For Output As #fnum
    Print #fnum, "MS Project Field Mapper - Phase 1 scan dump"
    Print #fnum, "Generated: " & FormatIsoUtcNow()
    Print #fnum, "Session:   " & ctx.SessionId
    Print #fnum, "File:      " & filePath
    Print #fnum, String$(60, "=")
    Print #fnum, dump
    Close #fnum

    ctx.Logger.Flush

    ' Append the log into the same file so the user has one self-contained dump.
    AppendFile resultPath, ctx.Logger.LogPath

    Phase1_ScanToFile = resultPath
    Exit Function

Fail:
    Dim n As Long, d As String
    n = Err.Number: d = Err.Description
    On Error Resume Next
    If fnum <> 0 Then Close #fnum
    On Error GoTo 0
    MsgBox "Phase1_ScanToFile failed: " & n & " - " & d, vbCritical, APP_NAME
    Phase1_ScanToFile = ""
End Function

'==============================================================================
' AppendFile (private)
' ----------------------------------------------------------------------------
' Append the contents of `srcPath` to `dstPath`. Used to fold the session log
' into the Phase 1 dump.
'==============================================================================
Private Sub AppendFile(ByVal dstPath As String, ByVal srcPath As String)
    On Error GoTo Done
    If LenB(srcPath) = 0 Then Exit Sub
    If Dir$(srcPath) = "" Then Exit Sub

    Dim dstNum As Integer, srcNum As Integer
    dstNum = FreeFile
    Open dstPath For Append As #dstNum
    Print #dstNum, String$(60, "=")
    Print #dstNum, "Session log:"
    Print #dstNum, String$(60, "-")

    srcNum = FreeFile
    Open srcPath For Input As #srcNum
    Dim line As String
    Do Until EOF(srcNum)
        Line Input #srcNum, line
        Print #dstNum, line
    Loop
    Close #srcNum
    Close #dstNum
Done:
    On Error GoTo 0
End Sub
