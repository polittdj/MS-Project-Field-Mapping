Attribute VB_Name = "modEnvironment"
'==============================================================================
' modEnvironment
' ----------------------------------------------------------------------------
' Host / version / security-context detection for the MS Project Field Mapper.
' Anything that pokes at the Project Application object's environmental state,
' the registry, or the file system's MOTW lives here so the rest of the
' codebase can ignore platform branching.
'
' Conditional compilation is used for 64-bit safety. VBA7 = Office 2010+;
' Win64 = 64-bit Office. The whole tool requires VBA7 (Project 2016+ ships
' with VBA7 only), so we don't bother branching on that constant.
'==============================================================================
Option Explicit

'--- AutomationSecurity values (mirror Office MsoAutomationSecurity) --------
' We don't import Office's enum because it isn't always referenced; numeric
' values are stable across Office versions.
Public Const masLow            As Long = 1
Public Const masByUI           As Long = 2
Public Const masForceDisable   As Long = 3

Private Const REG_OFFICE_BASE As String = "HKEY_CURRENT_USER\Software\Microsoft\Office\"

'==============================================================================
' AssertHostSupported
' ----------------------------------------------------------------------------
' Hard gate at session start. Returns True if the host is a supported version
' of Microsoft Project Desktop, otherwise False (caller should abort with a
' clear message). Project for the Web and Project Online do not host VBA at
' all, so reaching VBA already implies a desktop host -- but we still validate
' Application.Name and Application.Version against the documented contract.
'==============================================================================
Public Function AssertHostSupported(ByRef whyNot As String) As Boolean
    On Error GoTo Fail

    Dim appName As String
    appName = Application.Name

    If InStr(1, appName, "Project", vbTextCompare) = 0 Then
        whyNot = "Host application is '" & appName & "', not Microsoft Project."
        AssertHostSupported = False
        Exit Function
    End If

    Dim major As Long
    major = ProjectMajorVersion()

    If major < APP_MIN_PROJECT_MAJOR Then
        whyNot = "Microsoft Project version " & Application.Version & _
                 " is below the minimum supported version (" & _
                 APP_MIN_PROJECT_MAJOR & ".0 = Project 2016)."
        AssertHostSupported = False
        Exit Function
    End If

    AssertHostSupported = True
    Exit Function

Fail:
    whyNot = "Environment check raised " & Err.Number & ": " & Err.Description
    AssertHostSupported = False
End Function

'==============================================================================
' ProjectMajorVersion
' ----------------------------------------------------------------------------
' Parses Application.Version (e.g. "16.0.17328.20184") and returns the
' integer before the first dot. Returns 0 on parse failure.
'==============================================================================
Public Function ProjectMajorVersion() As Long
    On Error GoTo Fail
    Dim raw As String, dot As Long
    raw = Application.Version
    dot = InStr(1, raw, ".")
    If dot > 1 Then
        ProjectMajorVersion = CLng(Left$(raw, dot - 1))
    Else
        ProjectMajorVersion = CLng(raw)
    End If
    Exit Function
Fail:
    ProjectMajorVersion = 0
End Function

'==============================================================================
' OfficeBitness
' ----------------------------------------------------------------------------
' Returns "x64" or "x86" for the running Office process. Distinct from OS
' bitness -- a 64-bit Windows can host 32-bit Office.
'==============================================================================
Public Function OfficeBitness() As String
#If Win64 Then
    OfficeBitness = "x64"
#Else
    OfficeBitness = "x86"
#End If
End Function

'==============================================================================
' OSBitness
' ----------------------------------------------------------------------------
' Best-effort OS bitness via PROCESSOR_ARCHITECTURE. On 64-bit Windows hosting
' 32-bit Office, this still returns the OS arch through PROCESSOR_ARCHITEW6432.
'==============================================================================
Public Function OSBitness() As String
    Dim arch As String, archW As String
    arch = Environ$("PROCESSOR_ARCHITECTURE")
    archW = Environ$("PROCESSOR_ARCHITEW6432")
    If LenB(archW) > 0 Then arch = archW
    If InStr(1, arch, "64", vbTextCompare) > 0 Then
        OSBitness = "x64"
    Else
        OSBitness = "x86"
    End If
End Function

'==============================================================================
' AutomationSecurityState
' ----------------------------------------------------------------------------
' Returns the current Application.AutomationSecurity as a stable label.
' If the property is not exposed (older Project builds), returns "Unknown".
'==============================================================================
Public Function AutomationSecurityState() As String
    On Error GoTo Fail
    Dim v As Long
    v = Application.AutomationSecurity
    Select Case v
        Case masLow:          AutomationSecurityState = "Low"
        Case masByUI:         AutomationSecurityState = "ByUI"
        Case masForceDisable: AutomationSecurityState = "ForceDisable"
        Case Else:            AutomationSecurityState = "Other(" & v & ")"
    End Select
    Exit Function
Fail:
    AutomationSecurityState = "Unknown"
End Function

'==============================================================================
' AccessVBOMEnabled
' ----------------------------------------------------------------------------
' Reads HKCU\...\Office\<ver>\MS Project\Security\AccessVBOM. 1 = recipient
' has enabled "Trust access to the VBA project object model"; 0 / missing =
' disabled. Manual VBE Import → File works without this; programmatic module
' import requires it. Reported in diagnostics so we can advise the recipient.
'==============================================================================
Public Function AccessVBOMEnabled() As String
    Dim path As String
    path = REG_OFFICE_BASE & VersionedOfficeKey() & _
           "\MS Project\Security\AccessVBOM"
    Dim v As Variant
    v = TryRegRead(path)
    If IsEmpty(v) Then
        AccessVBOMEnabled = "NotSet"
    Else
        AccessVBOMEnabled = CStr(v)
    End If
End Function

'==============================================================================
' TrustedLocationsList
' ----------------------------------------------------------------------------
' Enumerates configured Trusted Locations for MS Project under the current
' Office major version. Returns a CR-LF separated list of "Path | Description"
' entries. Empty string if no locations configured or registry inaccessible.
'
' Trusted Locations registry layout:
'   HKCU\Software\Microsoft\Office\<ver>\MS Project\Security\Trusted Locations\
'     Location0\Path
'     Location0\Description
'     Location1\Path
'     ...
'==============================================================================
Public Function TrustedLocationsList() As String
    On Error GoTo Done

    Dim base As String, idx As Long, sep As String
    base = REG_OFFICE_BASE & VersionedOfficeKey() & _
           "\MS Project\Security\Trusted Locations\"
    sep = vbCrLf

    Dim out As String
    For idx = 0 To 49   ' arbitrary cap; Office never has this many
        Dim pPath As Variant, pDesc As Variant
        pPath = TryRegRead(base & "Location" & idx & "\Path")
        If IsEmpty(pPath) Then Exit For
        pDesc = TryRegRead(base & "Location" & idx & "\Description")
        out = out & CStr(pPath)
        If Not IsEmpty(pDesc) Then out = out & " | " & CStr(pDesc)
        out = out & sep
    Next idx

    TrustedLocationsList = out
    Exit Function
Done:
    TrustedLocationsList = ""
End Function

'==============================================================================
' VersionedOfficeKey
' ----------------------------------------------------------------------------
' Returns the Office version subkey under HKCU\Software\Microsoft\Office that
' matches the running Project. Project 2016/2019/2021/365 desktop all share
' "16.0".
'==============================================================================
Public Function VersionedOfficeKey() As String
    VersionedOfficeKey = ProjectMajorVersion() & ".0"
End Function

'==============================================================================
' FileHasMOTW
' ----------------------------------------------------------------------------
' Probe Mark-of-the-Web on a path. NTFS stores MOTW in the Zone.Identifier
' Alternate Data Stream. We try to open "<path>:Zone.Identifier" via FSO. If
' the open succeeds, MOTW is present. FSO raises on missing ADS, so an error
' here means "no MOTW".
'
' Returns True if MOTW present, False otherwise. Returns False on any
' inaccessibility error (better to under-report than block a clean file).
'==============================================================================
Public Function FileHasMOTW(ByVal filePath As String) As Boolean
    On Error GoTo NoMOTW
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath & ":Zone.Identifier", 1, False)  ' ForReading
    ts.Close
    FileHasMOTW = True
    Exit Function
NoMOTW:
    FileHasMOTW = False
End Function

'==============================================================================
' DiagnosticsReport
' ----------------------------------------------------------------------------
' Builds a multi-line diagnostics string. Phase 0's primary deliverable -- if
' a recipient runs this and the output looks reasonable, we know the bundle
' import path works on their machine.
'
' Caller writes this to a file (see modMain.RunDiagnostics).
'==============================================================================
Public Function DiagnosticsReport() As String
    Dim s As String
    s = s & "MS Project Field Mapper - Diagnostics" & vbCrLf
    s = s & "Generated: " & FormatIsoUtcNow() & vbCrLf
    s = s & "App: " & APP_NAME & " " & APP_VERSION & vbCrLf
    s = s & String$(60, "-") & vbCrLf

    s = s & "Host application:    " & SafeProp(Application.Name) & vbCrLf
    s = s & "Application.Version: " & SafeProp(Application.Version) & vbCrLf
    s = s & "Major version:       " & ProjectMajorVersion() & vbCrLf
    s = s & "Min supported:       " & APP_MIN_PROJECT_MAJOR & ".0 (Project 2016)" & vbCrLf

    s = s & "Office bitness:      " & OfficeBitness() & vbCrLf
    s = s & "OS bitness:          " & OSBitness() & vbCrLf
    s = s & "Computer name:       " & Environ$("COMPUTERNAME") & vbCrLf
    s = s & "User name:           " & Environ$("USERNAME") & vbCrLf

    s = s & "AutomationSecurity:  " & AutomationSecurityState() & vbCrLf
    s = s & "Trust VBOM access:   " & AccessVBOMEnabled() & vbCrLf
    s = s & String$(60, "-") & vbCrLf

    s = s & "Trusted Locations (HKCU):" & vbCrLf
    Dim tl As String
    tl = TrustedLocationsList()
    If LenB(tl) = 0 Then
        s = s & "  (none configured)" & vbCrLf
    Else
        Dim arr() As String, i As Long
        arr = Split(tl, vbCrLf)
        For i = LBound(arr) To UBound(arr)
            If LenB(arr(i)) > 0 Then s = s & "  " & arr(i) & vbCrLf
        Next i
    End If

    s = s & String$(60, "-") & vbCrLf
    Dim hostOk As Boolean, why As String
    hostOk = AssertHostSupported(why)
    s = s & "Host supported:      " & IIf(hostOk, "YES", "NO -- " & why) & vbCrLf

    DiagnosticsReport = s
End Function

'==============================================================================
' FormatIsoUtcNow
' ----------------------------------------------------------------------------
' ISO 8601 string for the current moment in UTC. Used in log filenames and
' the .fmap "created" field. Implementation: take Now() (local), convert to
' UTC by subtracting the system bias from the registry. We avoid Win32
' Declares for portability with the 32/64-bit story.
'
' Fallback: if UTC bias unavailable, returns local time with a "Z" anyway --
' the diagnostics file is for human reading, not data exchange. Persistence
' callers can substitute a stricter implementation later.
'==============================================================================
Public Function FormatIsoUtcNow() As String
    FormatIsoUtcNow = Format$(Now(), "yyyy-mm-dd\THH:Nn:Ss") & "Z"
End Function

'==============================================================================
' SafeProp (private)
' ----------------------------------------------------------------------------
' Wraps a string property access in a no-op guard. Some Application properties
' raise on certain Project builds -- diagnostics should still produce output.
'==============================================================================
Private Function SafeProp(ByVal v As Variant) As String
    On Error Resume Next
    SafeProp = CStr(v)
    If Err.Number <> 0 Then SafeProp = "<unavailable>"
    Err.Clear
End Function

'==============================================================================
' TryRegRead (private)
' ----------------------------------------------------------------------------
' Reads a registry value via WScript.Shell.RegRead. Returns Empty if the value
' does not exist or the read fails for any reason. Caller checks IsEmpty().
'==============================================================================
Private Function TryRegRead(ByVal regPath As String) As Variant
    On Error GoTo Fail
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    TryRegRead = sh.RegRead(regPath)
    Exit Function
Fail:
    TryRegRead = Empty
    Err.Clear
End Function
