VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFieldMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsFieldMetadata
' ----------------------------------------------------------------------------
' One scanned custom-field record. Captures everything the auto-mapper needs
' to score a candidate mapping: the underlying internal name (e.g. "Text2"),
' the user-assigned alias (e.g. "Location"), data type, slot index, scope
' (Task or Resource), a sample of populated values, and a populated-ratio.
'
' Sample storage uses a Collection rather than a typed array so we can append
' mid-scan without ReDim Preserve cost. Capped at SAMPLE_VALUES_PER_FIELD per
' the perf plan; further values are silently dropped.
'==============================================================================
Option Explicit

Private mInternalName    As String       ' "Text1", "Number5", ...
Private mAlias           As String       ' "Room", "Location", or "" if unaliased
Private mDataType        As FieldDataType
Private mSlotIndex       As Long         ' 1 for Text1, 5 for Number5
Private mScope           As FieldScope
Private mFieldConstant   As Long         ' MS Project pjCustomField* numeric constant
Private mSamples         As Collection
Private mPopulatedCount  As Long         ' how many tasks/resources had a value
Private mTotalCount      As Long         ' how many tasks/resources we visited
Private mIsReserved      As Boolean      ' True for master fields already in use

Private Sub Class_Initialize()
    Set mSamples = New Collection
    mDataType = fdtUnknown
    mScope = fsTask
End Sub

'==============================================================================
' Init
' ----------------------------------------------------------------------------
' Bulk initialiser called by modFileScanner. Avoids 6+ property assignments
' at every scan site.
'==============================================================================
Public Sub Init(ByVal internalName As String, _
                ByVal aliasName As String, _
                ByVal dataType As FieldDataType, _
                ByVal slotIndex As Long, _
                ByVal scope As FieldScope, _
                ByVal fieldConstant As Long)
    mInternalName = internalName
    mAlias = aliasName
    mDataType = dataType
    mSlotIndex = slotIndex
    mScope = scope
    mFieldConstant = fieldConstant
End Sub

'==============================================================================
' Properties
'==============================================================================
Public Property Get InternalName() As String:    InternalName = mInternalName: End Property
Public Property Get Alias() As String:           Alias = mAlias:               End Property
Public Property Get DataType() As FieldDataType: DataType = mDataType:         End Property
Public Property Get SlotIndex() As Long:         SlotIndex = mSlotIndex:       End Property
Public Property Get Scope() As FieldScope:       Scope = mScope:               End Property
Public Property Get FieldConstant() As Long:     FieldConstant = mFieldConstant: End Property
Public Property Get TotalCount() As Long:        TotalCount = mTotalCount:     End Property
Public Property Get PopulatedCount() As Long:    PopulatedCount = mPopulatedCount: End Property
Public Property Get SampleCount() As Long:       SampleCount = mSamples.Count: End Property

Public Property Get IsReserved() As Boolean:     IsReserved = mIsReserved: End Property
Public Property Let IsReserved(ByVal v As Boolean): mIsReserved = v: End Property

'==============================================================================
' Key
' ----------------------------------------------------------------------------
' Composite key suitable for a Dictionary lookup. "Task:Text1",
' "Resource:Number5". Scope-aware so task and resource Text1 don't collide.
'==============================================================================
Public Property Get Key() As String
    Key = FieldScopeName(mScope) & ":" & mInternalName
End Property

'==============================================================================
' PopulatedRatio
' ----------------------------------------------------------------------------
' Fraction of visited rows that had a non-empty value. 0.0 if no rows were
' visited (degenerate case for a project with zero tasks).
'==============================================================================
Public Property Get PopulatedRatio() As Double
    If mTotalCount <= 0 Then
        PopulatedRatio = 0#
    Else
        PopulatedRatio = mPopulatedCount / mTotalCount
    End If
End Property

'==============================================================================
' RegisterRow
' ----------------------------------------------------------------------------
' Called once per task/resource visited during scan. Tracks the denominator
' for populatedRatio independently of sample storage (samples are capped, but
' totals must reflect the whole project for an honest ratio).
'==============================================================================
Public Sub RegisterRow(ByVal hadValue As Boolean)
    mTotalCount = mTotalCount + 1
    If hadValue Then mPopulatedCount = mPopulatedCount + 1
End Sub

'==============================================================================
' AddSample
' ----------------------------------------------------------------------------
' Append a sample value. Caller is responsible for not passing empties --
' but we double-check here. Capped at SAMPLE_VALUES_PER_FIELD; further calls
' are silent no-ops to bound memory on huge projects.
'==============================================================================
Public Sub AddSample(ByVal v As Variant)
    If IsNull(v) Then Exit Sub
    If IsEmpty(v) Then Exit Sub
    If VarType(v) = vbString Then
        If LenB(CStr(v)) = 0 Then Exit Sub
    End If
    If mSamples.Count >= SAMPLE_VALUES_PER_FIELD Then Exit Sub
    mSamples.Add v
End Sub

'==============================================================================
' SampleAt
' ----------------------------------------------------------------------------
' 1-based access to stored samples.
'==============================================================================
Public Function SampleAt(ByVal index As Long) As Variant
    If index < 1 Or index > mSamples.Count Then
        SampleAt = Empty
        Exit Function
    End If
    Dim v As Variant
    v = mSamples.Item(index)
    SampleAt = v
End Function

'==============================================================================
' SamplesAsString
' ----------------------------------------------------------------------------
' Pipe-joined first N samples for log/debug output. N defaults to 5.
'==============================================================================
Public Function SamplesAsString(Optional ByVal maxItems As Long = 5) As String
    Dim s As String, i As Long, n As Long
    n = mSamples.Count
    If maxItems < n Then n = maxItems
    For i = 1 To n
        If i > 1 Then s = s & " | "
        s = s & CStr(mSamples.Item(i))
    Next i
    If maxItems < mSamples.Count Then s = s & " | ..."
    SamplesAsString = s
End Function

'==============================================================================
' ToDebugString
' ----------------------------------------------------------------------------
' One-line dump for debug output / Phase 1 verification. Matches the format
' shown in the manual mapping UI's listbox rows.
'==============================================================================
Public Function ToDebugString() As String
    ToDebugString = mInternalName & vbTab & _
                    DataTypeName(mDataType) & vbTab & _
                    FieldScopeName(mScope) & vbTab & _
                    "alias=""" & mAlias & """" & vbTab & _
                    "pop=" & Format$(PopulatedRatio() * 100, "0") & "%" & _
                    " (" & mPopulatedCount & "/" & mTotalCount & ")" & vbTab & _
                    "samples=[" & SamplesAsString() & "]" & _
                    IIf(mIsReserved, vbTab & "RESERVED", "")
End Function
