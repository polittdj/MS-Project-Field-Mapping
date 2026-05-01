VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsLogger
' ----------------------------------------------------------------------------
' Buffered session log. Holds clsLogEntry instances in memory and flushes to
' a .txt file on demand (typically at session end and on each phase boundary
' for crash safety).
'
' Severity filter: entries below MinSeverity are silently discarded -- caller
' can switch to sevDebug for development without code changes elsewhere.
'
' OnEntryAdded is a hook the progress UI can subscribe to in order to stream
' entries live. Not raised as a real VBA event (avoids RaiseEvent boilerplate
' across class boundaries); instead the consumer assigns a delegate object
' that exposes a "Notify" sub. See SubscribeNotify below.
'==============================================================================
Option Explicit

Private mEntries     As Collection
Private mMinSeverity As LogSeverity
Private mLogPath     As String
Private mNotifySink  As Object   ' optional; must expose Notify(entry)

Private Sub Class_Initialize()
    Set mEntries = New Collection
    mMinSeverity = sevInfo
End Sub

'==============================================================================
' Properties
'==============================================================================
Public Property Get MinSeverity() As LogSeverity
    MinSeverity = mMinSeverity
End Property
Public Property Let MinSeverity(ByVal v As LogSeverity)
    mMinSeverity = v
End Property

Public Property Get LogPath() As String
    LogPath = mLogPath
End Property
Public Property Let LogPath(ByVal v As String)
    mLogPath = v
End Property

Public Property Get Count() As Long
    Count = mEntries.Count
End Property

'==============================================================================
' SubscribeNotify
' ----------------------------------------------------------------------------
' Register an object (typically frmProgress) that exposes a `Notify` sub
' taking a clsLogEntry. Each accepted entry will be passed to it. Pass Nothing
' to unsubscribe.
'==============================================================================
Public Sub SubscribeNotify(ByVal sink As Object)
    Set mNotifySink = sink
End Sub

'==============================================================================
' Log* convenience wrappers
'==============================================================================
Public Sub LogDebug(ByVal msg As String, Optional ByVal phase As String = "", _
                    Optional ByVal fpath As String = "", _
                    Optional ByVal fid As String = "")
    AddEntry sevDebug, msg, phase, fpath, fid
End Sub

Public Sub LogInfo(ByVal msg As String, Optional ByVal phase As String = "", _
                   Optional ByVal fpath As String = "", _
                   Optional ByVal fid As String = "")
    AddEntry sevInfo, msg, phase, fpath, fid
End Sub

Public Sub LogWarn(ByVal msg As String, Optional ByVal phase As String = "", _
                   Optional ByVal fpath As String = "", _
                   Optional ByVal fid As String = "")
    AddEntry sevWarn, msg, phase, fpath, fid
End Sub

Public Sub LogError(ByVal msg As String, Optional ByVal phase As String = "", _
                    Optional ByVal fpath As String = "", _
                    Optional ByVal fid As String = "")
    AddEntry sevError, msg, phase, fpath, fid
End Sub

'==============================================================================
' AddEntry (private)
'==============================================================================
Private Sub AddEntry(ByVal sev As LogSeverity, ByVal msg As String, _
                     ByVal phase As String, ByVal fpath As String, _
                     ByVal fid As String)
    If sev < mMinSeverity Then Exit Sub

    Dim e As clsLogEntry
    Set e = New clsLogEntry
    e.Init sev, msg, phase, fpath, fid
    mEntries.Add e

    If Not mNotifySink Is Nothing Then
        On Error Resume Next
        mNotifySink.Notify e
        On Error GoTo 0
    End If
End Sub

'==============================================================================
' EntryAt
' ----------------------------------------------------------------------------
' 1-based access for callers that want to render the buffer.
'==============================================================================
Public Function EntryAt(ByVal index As Long) As clsLogEntry
    Set EntryAt = mEntries.Item(index)
End Function

'==============================================================================
' Flush
' ----------------------------------------------------------------------------
' Append all buffered entries to LogPath. After a successful flush the buffer
' is cleared so subsequent writes don't double-up. If LogPath is empty,
' raises an error -- the caller should set it up-front.
'==============================================================================
Public Sub Flush()
    If LenB(mLogPath) = 0 Then
        Err.Raise vbObjectError + 5001, "clsLogger.Flush", _
                  "LogPath not set. Cannot flush log entries."
    End If
    If mEntries.Count = 0 Then Exit Sub

    Dim fnum As Integer
    fnum = FreeFile

    On Error GoTo Fail
    Open mLogPath For Append As #fnum

    Dim i As Long, e As clsLogEntry
    For i = 1 To mEntries.Count
        Set e = mEntries.Item(i)
        Print #fnum, e.AsLine()
    Next i

    Close #fnum

    Set mEntries = New Collection
    Exit Sub

Fail:
    Dim n As Long, d As String
    n = Err.Number: d = Err.Description
    On Error Resume Next
    Close #fnum
    Err.Raise n, "clsLogger.Flush", d
End Sub

'==============================================================================
' Class_Terminate
' ----------------------------------------------------------------------------
' Best-effort flush on cleanup. Errors are swallowed -- the caller has had
' every opportunity to flush deliberately, this is just a safety net.
'==============================================================================
Private Sub Class_Terminate()
    On Error Resume Next
    If LenB(mLogPath) > 0 And mEntries.Count > 0 Then Flush
    On Error GoTo 0
End Sub
