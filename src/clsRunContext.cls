VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRunContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsRunContext
' ----------------------------------------------------------------------------
' Single object passed to every entry point. Replaces the global-state pattern
' typical of VBA codebases by bundling: the Project Application reference, a
' shared logger, the user's run settings (output paths, mode flags), and
' co-operative cancel state.
'
' Phase 0 surface is intentionally narrow -- only the fields required for
' diagnostics are wired up. Later phases extend this class as orchestration
' grows.
'==============================================================================
Option Explicit

Private mApp                As Object       ' MSProject.Application (late-set)
Private mLogger             As clsLogger
Private mOutputFolder       As String
Private mMode               As RunMode
Private mCancelRequested    As Boolean
Private mScans              As Object       ' Scripting.Dictionary, key = file path
Private mSessionStartedUtc  As String
Private mSessionId          As String

Private Sub Class_Initialize()
    Set mLogger = New clsLogger
    Set mScans = CreateObject("Scripting.Dictionary")
    mScans.CompareMode = 1   ' TextCompare; case-insensitive paths on Windows
    mMode = modeExtractOnly
    mCancelRequested = False
    mSessionStartedUtc = FormatIsoUtcNow()
    mSessionId = GenerateSessionId()
End Sub

'==============================================================================
' Init
' ----------------------------------------------------------------------------
' Wires the Project Application reference. Called once at session start by
' modMain. Kept separate from Class_Initialize so the class can be unit-tested
' without a live host.
'==============================================================================
Public Sub Init(ByVal appRef As Object)
    Set mApp = appRef
End Sub

'==============================================================================
' Properties
'==============================================================================
Public Property Get App() As Object
    Set App = mApp
End Property

Public Property Get Logger() As clsLogger
    Set Logger = mLogger
End Property

Public Property Get OutputFolder() As String
    OutputFolder = mOutputFolder
End Property
Public Property Let OutputFolder(ByVal v As String)
    mOutputFolder = NormalizeFolder(v)
End Property

Public Property Get Mode() As RunMode
    Mode = mMode
End Property
Public Property Let Mode(ByVal v As RunMode)
    mMode = v
End Property

Public Property Get CancelRequested() As Boolean
    CancelRequested = mCancelRequested
End Property
Public Sub RequestCancel()
    mCancelRequested = True
    If Not mLogger Is Nothing Then mLogger.LogWarn "Cancel requested by user.", "Run"
End Sub

Public Property Get Scans() As Object
    Set Scans = mScans
End Property

Public Property Get SessionStartedUtc() As String
    SessionStartedUtc = mSessionStartedUtc
End Property

Public Property Get SessionId() As String
    SessionId = mSessionId
End Property

'==============================================================================
' DefaultLogPath
' ----------------------------------------------------------------------------
' Compute the per-session log path from OutputFolder + session id. Falls back
' to %TEMP% if OutputFolder is unset (e.g. during a Diagnostics-only run from
' frmMain before the user has chosen an output folder).
'==============================================================================
Public Function DefaultLogPath() As String
    Dim folder As String
    If LenB(mOutputFolder) > 0 Then
        folder = mOutputFolder
    Else
        folder = Environ$("TEMP")
        If LenB(folder) = 0 Then folder = "C:\Temp"
        folder = NormalizeFolder(folder)
    End If
    DefaultLogPath = folder & "FieldMapper_" & mSessionId & EXT_LOG
End Function

'==============================================================================
' DefaultDiagnosticsPath
' ----------------------------------------------------------------------------
' Used by Phase 0 RunDiagnostics. Lives next to the log if OutputFolder is
' set, otherwise in %TEMP%.
'==============================================================================
Public Function DefaultDiagnosticsPath() As String
    Dim folder As String
    If LenB(mOutputFolder) > 0 Then
        folder = mOutputFolder
    Else
        folder = Environ$("TEMP")
        If LenB(folder) = 0 Then folder = "C:\Temp"
        folder = NormalizeFolder(folder)
    End If
    DefaultDiagnosticsPath = folder & "Diagnostics_" & mSessionId & EXT_LOG
End Function

'==============================================================================
' NormalizeFolder (private)
' ----------------------------------------------------------------------------
' Ensures a trailing path separator. Accepts both "\" and "/" but emits "\"
' for consistency with Project's usual conventions on Windows.
'==============================================================================
Private Function NormalizeFolder(ByVal path As String) As String
    If LenB(path) = 0 Then Exit Function
    Dim last As String
    last = Right$(path, 1)
    If last = "\" Or last = "/" Then
        NormalizeFolder = path
    Else
        NormalizeFolder = path & "\"
    End If
End Function

'==============================================================================
' GenerateSessionId (private)
' ----------------------------------------------------------------------------
' Compact unique-ish id for filename use. Format: yyyymmdd_hhnnss_RRR where
' RRR is a 0-999 random suffix to disambiguate sub-second invocations.
'==============================================================================
Private Function GenerateSessionId() As String
    Randomize
    GenerateSessionId = Format$(Now(), "yyyymmdd_hhnnss") & "_" & _
                        Format$(Int(Rnd() * 1000), "000")
End Function
