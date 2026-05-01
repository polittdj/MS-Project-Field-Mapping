Attribute VB_Name = "modConstants"
'==============================================================================
' modConstants
' ----------------------------------------------------------------------------
' Central definitions of enums, public constants, and slot tables for the MS
' Project Field Mapper. Every other module pulls its magic numbers from here.
'==============================================================================
Option Explicit

'--- Application identity ----------------------------------------------------
Public Const APP_NAME              As String = "MSProjectFieldMapper"
Public Const APP_VERSION           As String = "0.9.0"
Public Const APP_SCHEMA_VERSION    As Long = 1
Public Const APP_MIN_PROJECT_MAJOR As Long = 16   ' Project 2016+

'--- File-handling tunables --------------------------------------------------
Public Const MAX_SOURCE_FILES_SOFT As Long = 10
Public Const MAX_SOURCE_FILES_HARD As Long = 25
Public Const SAMPLE_VALUES_PER_FIELD As Long = 50
Public Const WAIT_FOR_IDLE_TIMEOUT_MS As Long = 2000
Public Const FILE_LOCK_RETRY_DELAY_MS As Long = 2000
Public Const FINGERPRINT_HEAD_BYTES As Long = 4096

'--- Auto-mapper thresholds and weights --------------------------------------
Public Const SCORE_AUTO_MAP    As Long = 70
Public Const SCORE_MANUAL_CONFIRM As Long = 40

Public Const W_ALIAS_EXACT     As Long = 50
Public Const W_ALIAS_CI        As Long = 45
Public Const W_ALIAS_NORM      As Long = 40
Public Const W_ALIAS_LEV       As Long = 30   ' multiplied by similarity ratio
Public Const W_ALIAS_JW        As Long = 5
Public Const W_TYPE_MATCH      As Long = 20
Public Const W_TYPE_MISMATCH   As Long = -25
Public Const W_PATTERN_MATCH   As Long = 15
Public Const W_TOKEN_OVERLAP   As Long = 10
Public Const W_SLOT_INDEX      As Long = 5

Public Const LEV_SIMILARITY_FLOOR As Double = 0.75
Public Const JW_SIMILARITY_FLOOR  As Double = 0.85
Public Const JACCARD_FLOOR        As Double = 0.3
Public Const MIN_SAMPLES_FOR_PATTERN As Long = 3
Public Const SHORT_ALIAS_THRESHOLD   As Long = 4   ' min(len) <= 4 forces exact-or-prefix

'--- Empty-field rule (when populatedRatio = 0 on either side) ---------------
Public Const EMPTY_FIELD_ALIAS_FLOOR As Long = 45  ' must reach case-insensitive exact

'--- Collision drop-rank weights ---------------------------------------------
Public Const COLLISION_W_POPULATED  As Double = 0.6
Public Const COLLISION_W_CONFIDENCE As Double = 0.4

'--- File extensions ---------------------------------------------------------
Public Const EXT_FMAP As String = ".fmap"
Public Const EXT_LOG  As String = ".txt"
Public Const EXT_CSV  As String = ".csv"
Public Const EXT_XLSX As String = ".xlsx"
Public Const EXT_MPP  As String = ".mpp"
Public Const EXT_MPT  As String = ".mpt"

'--- Severity levels for clsLogger ------------------------------------------
Public Enum LogSeverity
    sevDebug = 1
    sevInfo = 2
    sevWarn = 3
    sevError = 4
End Enum

'--- Mapping origin for clsMapping ------------------------------------------
Public Enum MapMethod
    methodAuto = 1
    methodManualConfirmed = 2
    methodManualMapped = 3
    methodSkippedUnique = 4
    methodCollisionDropped = 5
End Enum

'--- Master kind ------------------------------------------------------------
Public Enum MasterKind
    masterFromImported = 1
    masterExternalTemplate = 2
End Enum

'--- Run mode ---------------------------------------------------------------
Public Enum RunMode
    modeExtractOnly = 1
    modeExtractAndMerge = 2
End Enum

'--- Custom field data type (mapped to Project's pjCustomField* groupings) --
Public Enum FieldDataType
    fdtUnknown = 0
    fdtText = 1
    fdtNumber = 2
    fdtDate = 3
    fdtStart = 4
    fdtFinish = 5
    fdtDuration = 6
    fdtCost = 7
    fdtFlag = 8
    fdtOutlineCode = 9
End Enum

'--- Field scope (Task vs Resource custom field) ----------------------------
' Task and Resource custom fields are separate slot pools in Project; Text1
' on tasks is a different field from Text1 on resources. The mapper treats
' them as distinct universes -- mappings only happen within a scope.
'
' Numeric values mirror MS Project's PjFieldType enum:
'   pjTask = 0, pjResource = 1
Public Enum FieldScope
    fsTask = 0
    fsResource = 1
End Enum

'==============================================================================
' SlotCountForType
' ----------------------------------------------------------------------------
' Returns the number of custom field slots Microsoft Project provides for a
' given data type. Source: documented MS Project object model.
'==============================================================================
Public Function SlotCountForType(ByVal t As FieldDataType) As Long
    Select Case t
        Case fdtText:        SlotCountForType = 30
        Case fdtNumber:      SlotCountForType = 20
        Case fdtDate:        SlotCountForType = 10
        Case fdtStart:       SlotCountForType = 10
        Case fdtFinish:      SlotCountForType = 10
        Case fdtDuration:    SlotCountForType = 10
        Case fdtCost:        SlotCountForType = 10
        Case fdtFlag:        SlotCountForType = 20
        Case fdtOutlineCode: SlotCountForType = 10
        Case Else:           SlotCountForType = 0
    End Select
End Function

'==============================================================================
' DataTypeName
' ----------------------------------------------------------------------------
' Stable string name of a data type. Used in logs and persistence schema.
'==============================================================================
Public Function DataTypeName(ByVal t As FieldDataType) As String
    Select Case t
        Case fdtText:        DataTypeName = "Text"
        Case fdtNumber:      DataTypeName = "Number"
        Case fdtDate:        DataTypeName = "Date"
        Case fdtStart:       DataTypeName = "Start"
        Case fdtFinish:      DataTypeName = "Finish"
        Case fdtDuration:    DataTypeName = "Duration"
        Case fdtCost:        DataTypeName = "Cost"
        Case fdtFlag:        DataTypeName = "Flag"
        Case fdtOutlineCode: DataTypeName = "OutlineCode"
        Case Else:           DataTypeName = "Unknown"
    End Select
End Function

'==============================================================================
' SeverityName
'==============================================================================
Public Function SeverityName(ByVal s As LogSeverity) As String
    Select Case s
        Case sevDebug: SeverityName = "DEBUG"
        Case sevInfo:  SeverityName = "INFO"
        Case sevWarn:  SeverityName = "WARN"
        Case sevError: SeverityName = "ERROR"
        Case Else:     SeverityName = "INFO"
    End Select
End Function

'==============================================================================
' FieldScopeName
'==============================================================================
Public Function FieldScopeName(ByVal s As FieldScope) As String
    Select Case s
        Case fsTask:     FieldScopeName = "Task"
        Case fsResource: FieldScopeName = "Resource"
        Case Else:       FieldScopeName = "Unknown"
    End Select
End Function

'==============================================================================
' MapMethodName
'==============================================================================
Public Function MapMethodName(ByVal m As MapMethod) As String
    Select Case m
        Case methodAuto:             MapMethodName = "auto"
        Case methodManualConfirmed:  MapMethodName = "manualConfirmed"
        Case methodManualMapped:     MapMethodName = "manualMapped"
        Case methodSkippedUnique:    MapMethodName = "skippedUnique"
        Case methodCollisionDropped: MapMethodName = "collisionDropped"
        Case Else:                   MapMethodName = "unknown"
    End Select
End Function
