VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsLogEntry
' ----------------------------------------------------------------------------
' Single immutable-ish log record. Exposes properties for serialization.
' Created via clsLogger.Log* methods; not normally instantiated directly.
'==============================================================================
Option Explicit

Public Timestamp As Date
Public Severity  As LogSeverity
Public Phase     As String
Public FilePath  As String
Public FieldId   As String
Public Message   As String

'==============================================================================
' Init
' ----------------------------------------------------------------------------
' Convenience initialiser. VBA classes have no parameterised constructors, so
' the canonical pattern is `Set e = New clsLogEntry: e.Init args`.
'==============================================================================
Public Sub Init(ByVal sev As LogSeverity, _
                ByVal msg As String, _
                Optional ByVal phaseName As String = "", _
                Optional ByVal fpath As String = "", _
                Optional ByVal fid As String = "")
    Timestamp = Now()
    Severity = sev
    Message = msg
    Phase = phaseName
    FilePath = fpath
    FieldId = fid
End Sub

'==============================================================================
' AsLine
' ----------------------------------------------------------------------------
' One-line plain-text representation suitable for the .txt session log.
' Example: "14:02:11 INFO  [Scan] PlantA.mpp Text2  -- Opened file"
'==============================================================================
Public Function AsLine() As String
    Dim s As String
    s = Format$(Timestamp, "hh:nn:ss") & " " & _
        Left$(SeverityName(Severity) & "     ", 5) & " "
    If LenB(Phase) > 0    Then s = s & "[" & Phase & "] "
    If LenB(FilePath) > 0 Then s = s & FilePath & " "
    If LenB(FieldId) > 0  Then s = s & FieldId & " "
    s = s & "-- " & Message
    AsLine = s
End Function
