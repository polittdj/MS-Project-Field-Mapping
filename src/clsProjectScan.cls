VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProjectScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsProjectScan
' ----------------------------------------------------------------------------
' In-memory snapshot of one .mpp file's relevant metadata: file fingerprint,
' source Project version, and the full set of scanned custom fields keyed by
' "scope:internalName" (e.g. "Task:Text2", "Resource:Number5").
'
' One instance per .mpp; modFileScanner produces it, the auto-mapper consumes
' it, and the persistence layer serialises it.
'==============================================================================
Option Explicit

Private mFilePath              As String
Private mFingerprint           As String
Private mProjectVersionString  As String
Private mScannedUtc            As String
Private mTaskCount             As Long
Private mResourceCount         As Long
Private mFields                As Object   ' Scripting.Dictionary, key = field key
Private mDuplicateAliases      As Collection

Private Sub Class_Initialize()
    Set mFields = CreateObject("Scripting.Dictionary")
    mFields.CompareMode = 1            ' TextCompare; case-insensitive keys
    Set mDuplicateAliases = New Collection
End Sub

'==============================================================================
' Properties
'==============================================================================
Public Property Get FilePath() As String:             FilePath = mFilePath:                      End Property
Public Property Let FilePath(ByVal v As String):      mFilePath = v:                              End Property

Public Property Get Fingerprint() As String:          Fingerprint = mFingerprint:                End Property
Public Property Let Fingerprint(ByVal v As String):   mFingerprint = v:                           End Property

Public Property Get ProjectVersionString() As String: ProjectVersionString = mProjectVersionString: End Property
Public Property Let ProjectVersionString(ByVal v As String): mProjectVersionString = v:           End Property

Public Property Get ScannedUtc() As String:           ScannedUtc = mScannedUtc:                  End Property
Public Property Let ScannedUtc(ByVal v As String):    mScannedUtc = v:                            End Property

Public Property Get TaskCount() As Long:              TaskCount = mTaskCount:                    End Property
Public Property Let TaskCount(ByVal v As Long):       mTaskCount = v:                             End Property

Public Property Get ResourceCount() As Long:          ResourceCount = mResourceCount:            End Property
Public Property Let ResourceCount(ByVal v As Long):   mResourceCount = v:                         End Property

Public Property Get Fields() As Object:               Set Fields = mFields:                       End Property
Public Property Get FieldCount() As Long:             FieldCount = mFields.Count:                 End Property

'==============================================================================
' AddField
' ----------------------------------------------------------------------------
' Register a clsFieldMetadata. Keyed by its composite Key. Replaces any
' existing entry with the same key (defensive -- caller shouldn't double-add
' but masking the bug is worse than a silent overwrite).
'==============================================================================
Public Sub AddField(ByVal f As clsFieldMetadata)
    If f Is Nothing Then Exit Sub
    Dim k As String
    k = f.Key
    If mFields.Exists(k) Then mFields.Remove k
    mFields.Add k, f
End Sub

'==============================================================================
' GetField
' ----------------------------------------------------------------------------
' Returns the field for a key, or Nothing if absent.
'==============================================================================
Public Function GetField(ByVal key As String) As clsFieldMetadata
    If mFields.Exists(key) Then Set GetField = mFields.Item(key)
End Function

'==============================================================================
' GetFieldByInternalName
' ----------------------------------------------------------------------------
' Convenience lookup combining scope + internal name.
'==============================================================================
Public Function GetFieldByInternalName(ByVal scope As FieldScope, _
                                       ByVal internalName As String) As clsFieldMetadata
    Set GetFieldByInternalName = GetField(FieldScopeName(scope) & ":" & internalName)
End Function

'==============================================================================
' FieldKeysSnapshot
' ----------------------------------------------------------------------------
' Return the field keys as a stable Variant array. Iterating .Keys directly
' on a Dictionary is fine for one-shot consumers, but hot loops should call
' this once and index the result.
'==============================================================================
Public Function FieldKeysSnapshot() As Variant
    FieldKeysSnapshot = mFields.Keys
End Function

'==============================================================================
' DetectDuplicateAliases
' ----------------------------------------------------------------------------
' Walks the field collection, groups by (scope, dataType, lowercase alias),
' and returns a Collection of "Scope/DataType/alias" strings for any group
' with more than one member. Result is also cached internally and accessible
' via DuplicateAliases.
'
' Aliases are compared case-insensitively after trimming. Groups with empty
' alias are excluded (an unaliased field is not a "duplicate alias" case --
' that's a different scenario handled by the auto-mapper).
'==============================================================================
Public Function DetectDuplicateAliases() As Collection
    Set mDuplicateAliases = New Collection

    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = 1       ' TextCompare for grouping by alias

    Dim keys As Variant, i As Long, k As String, f As clsFieldMetadata
    keys = mFields.Keys
    For i = LBound(keys) To UBound(keys)
        k = CStr(keys(i))
        Set f = mFields.Item(k)
        If LenB(Trim$(f.Alias)) > 0 Then
            Dim gkey As String
            gkey = FieldScopeName(f.Scope) & "/" & DataTypeName(f.DataType) & _
                   "/" & LCase$(Trim$(f.Alias))
            If groups.Exists(gkey) Then
                Dim col As Collection
                Set col = groups.Item(gkey)
                col.Add f
            Else
                Dim newCol As Collection
                Set newCol = New Collection
                newCol.Add f
                groups.Add gkey, newCol
            End If
        End If
    Next i

    Dim gkeys As Variant
    gkeys = groups.Keys
    For i = LBound(gkeys) To UBound(gkeys)
        Dim members As Collection
        Set members = groups.Item(CStr(gkeys(i)))
        If members.Count > 1 Then mDuplicateAliases.Add CStr(gkeys(i))
    Next i

    Set DetectDuplicateAliases = mDuplicateAliases
End Function

Public Property Get DuplicateAliases() As Collection
    Set DuplicateAliases = mDuplicateAliases
End Property

'==============================================================================
' ToDebugString
' ----------------------------------------------------------------------------
' Multi-line dump suitable for Debug.Print or a session log entry. Phase 1's
' primary verification path: scan a known sample, dump, eyeball.
'==============================================================================
Public Function ToDebugString() As String
    Dim s As String
    s = s & "=== ProjectScan ===" & vbCrLf
    s = s & "File:        " & mFilePath & vbCrLf
    s = s & "Fingerprint: " & mFingerprint & vbCrLf
    s = s & "ProjectVer:  " & mProjectVersionString & vbCrLf
    s = s & "ScannedUtc:  " & mScannedUtc & vbCrLf
    s = s & "Tasks:       " & mTaskCount & vbCrLf
    s = s & "Resources:   " & mResourceCount & vbCrLf
    s = s & "Fields:      " & mFields.Count & vbCrLf

    If mDuplicateAliases.Count > 0 Then
        s = s & "Duplicate aliases:" & vbCrLf
        Dim j As Long
        For j = 1 To mDuplicateAliases.Count
            s = s & "  " & mDuplicateAliases.Item(j) & vbCrLf
        Next j
    End If

    s = s & String$(60, "-") & vbCrLf

    Dim keys As Variant, i As Long, f As clsFieldMetadata
    keys = mFields.Keys
    For i = LBound(keys) To UBound(keys)
        Set f = mFields.Item(CStr(keys(i)))
        s = s & f.ToDebugString() & vbCrLf
    Next i

    ToDebugString = s
End Function
