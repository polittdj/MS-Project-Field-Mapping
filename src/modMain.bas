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
    MsgBox "RunMapper is not implemented in Phase 0. " & vbCrLf & _
           "Use ShowMain to launch the UI, or RunDiagnostics for a " & _
           "diagnostics dump.", vbInformation, APP_NAME
End Sub
