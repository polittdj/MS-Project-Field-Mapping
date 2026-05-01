Attribute VB_Name = "modFileScanner"
'==============================================================================
' modFileScanner
' ----------------------------------------------------------------------------
' Open/scan/close pattern for a single .mpp. Phase 1 surface: one file at a
' time. Phase 2 will introduce the multi-file orchestrator that composes
' ScanProjectByPath in a loop with cancellation and continue-on-error.
'
' Two public entry points:
'   ScanProjectByPath - given a path, opens the file read-only, scans it,
'                       and closes it. Wraps the full lifecycle including
'                       Application-state save/restore and force-close on
'                       error.
'   ScanProject       - given an already-open Project COM object, populates
'                       a clsProjectScan. Useful for tests and for Phase 2
'                       reuse.
'
' Late-bound throughout: no compile-time reference to MSProject is required,
' so the bundle imports cleanly into a fresh VBE before any host wiring.
'==============================================================================
Option Explicit

' MS Project enum values used late-bound (numeric values are stable):
Private Const PJ_DO_NOT_SAVE      As Long = 0    ' pjDoNotSave
Private Const PJ_MANUAL           As Long = 0    ' pjManual (Calculation mode)
Private Const PJ_AUTOMATIC        As Long = 1    ' pjAutomatic
Private Const PJ_TASK_SCOPE       As Long = 0    ' pjTask
Private Const PJ_RESOURCE_SCOPE   As Long = 1    ' pjResource
Private Const PJ_AUTOMATION_LOW   As Long = 1    ' msoAutomationSecurityLow
Private Const PJ_AUTOMATION_BLOCK As Long = 3    ' msoAutomationSecurityForceDisable

' MS Project file-open errors we recognise:
Private Const ERR_FILE_LOCKED     As Long = 1101

'==============================================================================
' ScanProjectByPath
' ----------------------------------------------------------------------------
' Opens a single .mpp read-only, scans its custom fields, and closes it.
' Returns the populated clsProjectScan on success; Nothing on failure (with
' the reason logged into ctx.logger).
'
' Application state is captured before mutation and restored unconditionally
' in the cleanup label, even on early exit. The cleanup also force-closes
' any project we opened that's still open at exit -- guards against a
' partial-success state that would leak into the next file.
'==============================================================================
Public Function ScanProjectByPath(ByVal filePath As String, _
                                  ByVal ctx As clsRunContext) As clsProjectScan
    Dim openedHere     As Boolean
    Dim baselineCount  As Long
    Dim priorCalc      As Long
    Dim priorAutoSec   As Long
    Dim priorScreen    As Boolean
    Dim priorAlerts    As Boolean
    Dim stateCaptured  As Boolean
    Dim scan           As clsProjectScan

    On Error GoTo Cleanup

    If LenB(filePath) = 0 Then
        ctx.Logger.LogError "Empty file path passed to ScanProjectByPath.", "Scan"
        GoTo Cleanup
    End If
    If Dir$(filePath) = "" Then
        ctx.Logger.LogWarn "File not found: " & filePath, "Scan", filePath
        GoTo Cleanup
    End If

    ' --- Capture and configure Application state ---
    Dim app As Object
    Set app = Application

    priorCalc     = SafeReadCalculation(app)
    priorAutoSec  = SafeReadAutomationSecurity(app)
    priorScreen   = app.ScreenUpdating
    priorAlerts   = app.DisplayAlerts
    stateCaptured = True

    app.DisplayAlerts        = False
    app.ScreenUpdating       = False
    On Error Resume Next
    app.AutomationSecurity   = PJ_AUTOMATION_BLOCK
    app.Calculation          = PJ_MANUAL
    On Error GoTo Cleanup

    baselineCount = app.Projects.Count

    ' --- Open the file ---
    ctx.Logger.LogInfo "Opening (read-only): " & filePath, "Scan", filePath
    On Error GoTo HandleOpenError
    app.FileOpenEx Name:=filePath, ReadOnly:=True
    On Error GoTo Cleanup

    If app.Projects.Count <= baselineCount Then
        ctx.Logger.LogError "FileOpenEx returned without registering a new project.", _
                            "Scan", filePath
        GoTo Cleanup
    End If
    openedHere = True

    Dim proj As Object
    Set proj = app.ActiveProject

    ' --- Scan ---
    Set scan = ScanProject(proj, ctx)
    scan.FilePath = filePath
    scan.Fingerprint = FileFingerprint(filePath)
    ctx.Logger.LogInfo "Scanned " & scan.FieldCount & " custom fields (Tasks=" & _
                       scan.TaskCount & ", Resources=" & scan.ResourceCount & ")", _
                       "Scan", filePath

    ' --- Close ---
    On Error Resume Next
    app.FileCloseEx PJ_DO_NOT_SAVE
    On Error GoTo Cleanup
    openedHere = False
    ctx.Logger.LogInfo "Closed: " & filePath, "Scan", filePath

    WaitForIdle

    ' --- Restore Application state ---
    RestoreAppState app, stateCaptured, priorCalc, priorAutoSec, priorScreen, priorAlerts

    Set ScanProjectByPath = scan
    Exit Function

HandleOpenError:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.Description
    Select Case errNum
        Case ERR_FILE_LOCKED
            ctx.Logger.LogWarn "File locked by another process: " & filePath, _
                               "Scan", filePath
        Case Else
            ctx.Logger.LogError "FileOpenEx failed (" & errNum & "): " & errDesc, _
                                "Scan", filePath
    End Select
    Resume Cleanup

Cleanup:
    Dim cleanupErr As Long, cleanupDesc As String
    cleanupErr = Err.Number: cleanupDesc = Err.Description

    If openedHere Then
        ctx.Logger.LogWarn "Force-closing opened project after error: " & filePath, _
                           "Scan", filePath
        On Error Resume Next
        app.FileCloseEx PJ_DO_NOT_SAVE
        On Error GoTo 0
    End If

    RestoreAppState app, stateCaptured, priorCalc, priorAutoSec, priorScreen, priorAlerts

    If cleanupErr <> 0 Then
        ctx.Logger.LogError "Cleanup after error " & cleanupErr & ": " & cleanupDesc, _
                            "Scan", filePath
        Err.Clear
    End If

    Set ScanProjectByPath = Nothing
End Function

'==============================================================================
' ScanProject
' ----------------------------------------------------------------------------
' Pure scan over an already-open Project. Builds the per-scope custom-field
' metadata, then walks tasks and resources to populate samples and
' populated-ratios.
'
' Fields with empty alias AND zero populated values are not added to the scan
' (per Deliverable 8: scan only fields whose alias is set OR whose values are
' populated). This keeps the data model focused on fields the user actually
' uses.
'==============================================================================
Public Function ScanProject(ByVal proj As Object, _
                            ByVal ctx As clsRunContext) As clsProjectScan
    Dim scan As clsProjectScan
    Set scan = New clsProjectScan
    scan.ProjectVersionString = SafeReadAppVersion()
    scan.ScannedUtc = FormatIsoUtcNow()
    On Error Resume Next
    scan.TaskCount = proj.Tasks.Count
    scan.ResourceCount = proj.Resources.Count
    On Error GoTo 0

    Dim taskFields As Collection
    Set taskFields = BuildFieldMetadataList(proj, fsTask, ctx)

    Dim resFields As Collection
    Set resFields = BuildFieldMetadataList(proj, fsResource, ctx)

    ScanRows proj.Tasks, taskFields, ctx, fsTask
    If scan.ResourceCount > 0 Then
        ScanRows proj.Resources, resFields, ctx, fsResource
    End If

    AddNonEmptyFields scan, taskFields
    AddNonEmptyFields scan, resFields

    scan.DetectDuplicateAliases
    Set ScanProject = scan
End Function

'==============================================================================
' WaitForIdle
' ----------------------------------------------------------------------------
' Cooperative yield. Replaces fixed Sleep with a bounded DoEvents pump --
' enough to let Project drain its message queue after a file close, but
' bounded so a misbehaving callback can't block indefinitely.
'==============================================================================
Public Sub WaitForIdle(Optional ByVal maxLoops As Long = 100)
    Dim i As Long
    For i = 1 To maxLoops
        DoEvents
    Next i
End Sub

'==============================================================================
' BuildFieldMetadataList (private)
' ----------------------------------------------------------------------------
' For one scope (Task or Resource), enumerate every custom-field slot and
' produce a Collection of clsFieldMetadata seeded with internal name, alias,
' data type, and field constant. Sample values are added later by ScanRows.
'==============================================================================
Private Function BuildFieldMetadataList(ByVal proj As Object, _
                                        ByVal scope As FieldScope, _
                                        ByVal ctx As clsRunContext) As Collection
    Dim col As Collection
    Set col = New Collection

    AppendSlots col, proj, scope, fdtText,        "Text",         SlotCountForType(fdtText), ctx
    AppendSlots col, proj, scope, fdtNumber,      "Number",       SlotCountForType(fdtNumber), ctx
    AppendSlots col, proj, scope, fdtDate,        "Date",         SlotCountForType(fdtDate), ctx
    AppendSlots col, proj, scope, fdtStart,       "Start",        SlotCountForType(fdtStart), ctx
    AppendSlots col, proj, scope, fdtFinish,      "Finish",       SlotCountForType(fdtFinish), ctx
    AppendSlots col, proj, scope, fdtDuration,    "Duration",     SlotCountForType(fdtDuration), ctx
    AppendSlots col, proj, scope, fdtCost,        "Cost",         SlotCountForType(fdtCost), ctx
    AppendSlots col, proj, scope, fdtFlag,        "Flag",         SlotCountForType(fdtFlag), ctx
    AppendSlots col, proj, scope, fdtOutlineCode, "Outline Code", SlotCountForType(fdtOutlineCode), ctx

    Set BuildFieldMetadataList = col
End Function

'==============================================================================
' AppendSlots (private)
' ----------------------------------------------------------------------------
' Append all N slots of one (scope, dataType) combination to the collection.
' Each call resolves the field constant via FieldNameToFieldConstant and the
' alias via CustomFieldGetName. Failures on individual slots are logged at
' debug level and skipped -- one bad slot doesn't kill the scan.
'==============================================================================
Private Sub AppendSlots(ByVal col As Collection, _
                        ByVal proj As Object, _
                        ByVal scope As FieldScope, _
                        ByVal dataType As FieldDataType, _
                        ByVal namePrefix As String, _
                        ByVal slotCount As Long, _
                        ByVal ctx As clsRunContext)
    Dim app As Object
    Set app = Application

    Dim i As Long
    For i = 1 To slotCount
        Dim internalName As String
        internalName = namePrefix & i

        Dim fc As Long
        Dim aliasName As String

        On Error GoTo SlotFail
        fc = app.FieldNameToFieldConstant(internalName, scope)
        aliasName = ""
        aliasName = proj.CustomFieldGetName(fc)
        On Error GoTo 0

        Dim fm As clsFieldMetadata
        Set fm = New clsFieldMetadata
        fm.Init internalName, NzString(aliasName), dataType, i, scope, fc
        col.Add fm
        GoTo NextSlot

SlotFail:
        ctx.Logger.LogDebug "Slot resolve failed: " & FieldScopeName(scope) & "/" & _
                            internalName & " -- " & Err.Number & " " & Err.Description, _
                            "Scan"
        Err.Clear
NextSlot:
    Next i
End Sub

'==============================================================================
' ScanRows (private)
' ----------------------------------------------------------------------------
' Iterate a Tasks or Resources collection and pull each field's value via
' GetField, registering populated-ratio and accumulating samples.
'
' Loop invariants:
' - row may be Nothing for deleted/summary placeholders -- skip silently.
' - Flag fields are always considered "populated" (boolean has no empty
'   state in Project's model); only True values are sampled.
' - Date-typed empties show up as the OLE epoch (1899-12-30); we treat any
'   date <= 0 as empty.
'==============================================================================
Private Sub ScanRows(ByVal rowsCol As Object, _
                     ByVal fields As Collection, _
                     ByVal ctx As clsRunContext, _
                     ByVal scope As FieldScope)
    Dim row As Object
    Dim n As Long
    n = 0
    Dim startedAt As Single
    startedAt = Timer

    For Each row In rowsCol
        If Not row Is Nothing Then
            n = n + 1
            Dim fm As clsFieldMetadata
            For Each fm In fields
                Dim v As Variant
                Dim hadValue As Boolean
                Dim ok As Boolean
                ok = TryGetField(row, fm.FieldConstant, v)
                If ok Then
                    hadValue = ValueIsPopulated(v, fm.DataType)
                Else
                    hadValue = False
                End If
                fm.RegisterRow hadValue
                If hadValue And fm.DataType <> fdtFlag Then
                    fm.AddSample v
                ElseIf hadValue And fm.DataType = fdtFlag Then
                    fm.AddSample True
                End If
            Next fm

            If ctx.CancelRequested Then Exit For
        End If
    Next row

    ctx.Logger.LogDebug "ScanRows " & FieldScopeName(scope) & " visited " & n & _
                        " rows in " & Format$(Timer - startedAt, "0.00") & "s", "Scan"
End Sub

'==============================================================================
' AddNonEmptyFields (private)
' ----------------------------------------------------------------------------
' Move fields from the working Collection into the clsProjectScan, filtering
' out the truly-empty unaliased ones.
'==============================================================================
Private Sub AddNonEmptyFields(ByVal scan As clsProjectScan, _
                              ByVal fields As Collection)
    Dim fm As clsFieldMetadata
    For Each fm In fields
        If LenB(Trim$(fm.Alias)) > 0 Or fm.PopulatedCount > 0 Then
            scan.AddField fm
        End If
    Next fm
End Sub

'==============================================================================
' TryGetField (private)
' ----------------------------------------------------------------------------
' Wraps task.GetField / resource.GetField with a guard. Sets value ByRef and
' returns True on success.
'==============================================================================
Private Function TryGetField(ByVal row As Object, _
                             ByVal fieldConstant As Long, _
                             ByRef value As Variant) As Boolean
    On Error GoTo Fail
    value = row.GetField(fieldConstant)
    TryGetField = True
    Exit Function
Fail:
    value = Empty
    TryGetField = False
    Err.Clear
End Function

'==============================================================================
' ValueIsPopulated (private)
' ----------------------------------------------------------------------------
' Determines whether a value pulled from GetField counts as "populated" for
' populatedRatio purposes. The rules differ by data type:
'   Text / OutlineCode: non-empty string after trim
'   Date / Start / Finish: not Empty/Null and > OLE epoch (1899-12-30)
'   Number / Cost: non-zero (Project returns 0 for unset numerics; we accept
'                  the false-negative on intentionally-zero values as a
'                  reasonable trade-off)
'   Duration: > 0 (Project returns durations in minutes as a Long-ish value)
'   Flag: True is populated; False is the field's default and counts as not
'         populated for ratio purposes (the auto-mapper handles flag fields
'         via alias signal, not sample density)
'==============================================================================
Private Function ValueIsPopulated(ByRef v As Variant, _
                                  ByVal t As FieldDataType) As Boolean
    If IsNull(v) Or IsEmpty(v) Then Exit Function

    Select Case t
        Case fdtText, fdtOutlineCode
            ValueIsPopulated = (LenB(Trim$(CStr(v))) > 0)
        Case fdtDate, fdtStart, fdtFinish
            On Error Resume Next
            Dim d As Date
            d = CDate(v)
            ValueIsPopulated = (d > CDate(0))
            On Error GoTo 0
        Case fdtNumber, fdtCost, fdtDuration
            On Error Resume Next
            Dim dbl As Double
            dbl = CDbl(v)
            ValueIsPopulated = (dbl <> 0#)
            On Error GoTo 0
        Case fdtFlag
            On Error Resume Next
            ValueIsPopulated = CBool(v)
            On Error GoTo 0
        Case Else
            ValueIsPopulated = False
    End Select
End Function

'==============================================================================
' RestoreAppState (private)
'==============================================================================
Private Sub RestoreAppState(ByVal app As Object, _
                            ByVal captured As Boolean, _
                            ByVal priorCalc As Long, _
                            ByVal priorAutoSec As Long, _
                            ByVal priorScreen As Boolean, _
                            ByVal priorAlerts As Boolean)
    If Not captured Then Exit Sub
    On Error Resume Next
    app.Calculation        = priorCalc
    app.AutomationSecurity = priorAutoSec
    app.ScreenUpdating     = priorScreen
    app.DisplayAlerts      = priorAlerts
    On Error GoTo 0
End Sub

'==============================================================================
' SafeRead* helpers
' ----------------------------------------------------------------------------
' Some Application properties raise on certain Project builds. SafeRead
' wrappers return a sane default rather than blowing up the scan.
'==============================================================================
Private Function SafeReadCalculation(ByVal app As Object) As Long
    On Error Resume Next
    SafeReadCalculation = app.Calculation
    If Err.Number <> 0 Then SafeReadCalculation = PJ_AUTOMATIC
    Err.Clear
End Function

Private Function SafeReadAutomationSecurity(ByVal app As Object) As Long
    On Error Resume Next
    SafeReadAutomationSecurity = app.AutomationSecurity
    If Err.Number <> 0 Then SafeReadAutomationSecurity = PJ_AUTOMATION_LOW
    Err.Clear
End Function

Private Function SafeReadAppVersion() As String
    On Error Resume Next
    SafeReadAppVersion = Application.Version
    If Err.Number <> 0 Then SafeReadAppVersion = "unknown"
    Err.Clear
End Function

'==============================================================================
' NzString (private)
' ----------------------------------------------------------------------------
' Null-coalescing helper for late-bound calls that may return Null.
'==============================================================================
Private Function NzString(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        NzString = ""
    Else
        NzString = CStr(v)
    End If
End Function
