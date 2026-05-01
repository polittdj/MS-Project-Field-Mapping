VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain
   Caption         =   "MS Project Field Mapper"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   1
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblTitle
      Caption         =   "MS Project Field Mapper"
      ForeColor       =   -2147483630
      Height          =   270
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   6300
   End
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblPhase
      Caption         =   "Phase 0 (Foundation) - diagnostics only. Full mapper UI arrives in later phases."
      ForeColor       =   -2147483640
      Height          =   240
      Left            =   180
      TabIndex        =   1
      Top             =   420
      Width           =   6300
   End
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblHint
      Caption         =   "Click Run Diagnostics to write a .txt with host, version and macro-security details."
      Height          =   480
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   6300
   End
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblOutputCaption
      Caption         =   "Output folder (optional - defaults to %TEMP%):"
      Height          =   240
      Left            =   180
      TabIndex        =   3
      Top             =   1320
      Width           =   6300
   End
   Begin {8BD21D10-EC42-11CE-9E0D-00AA006002F3} txtOutputFolder
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   1560
      Width           =   5340
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnBrowseOutput
      Caption         =   "Browse..."
      Height          =   285
      Left            =   5580
      TabIndex        =   5
      Top             =   1560
      Width           =   900
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnDiagnostics
      Caption         =   "Run Diagnostics..."
      Height          =   330
      Left            =   180
      TabIndex        =   6
      Top             =   2520
      Width           =   1755
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnClose
      Caption         =   "Close"
      Height          =   330
      Left            =   5580
      TabIndex        =   7
      Top             =   2520
      Width           =   900
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmMain
' ----------------------------------------------------------------------------
' Phase 0 control panel for the MS Project Field Mapper. The full multi-step
' UI (file picker, master selection, mode, mapping config) is wired up in
' later phases. For now, only the Diagnostics path is functional -- enough to
' verify that the bundle imports cleanly into a target VBE.
'
' If this .frm fails to import on the target machine, paste frmMain.code.bas
' into a manually-built UserForm per INSTALL.md.
'==============================================================================
Option Explicit

'==============================================================================
' UserForm_Initialize
'==============================================================================
Private Sub UserForm_Initialize()
    Me.Caption = APP_NAME & " " & APP_VERSION & " - Phase 0"

    Dim why As String
    If Not AssertHostSupported(why) Then
        lblHint.Caption = "WARNING: Host check failed: " & why
        btnDiagnostics.Enabled = True   ' still allow diagnostics for triage
    End If
End Sub

'==============================================================================
' btnBrowseOutput_Click
' ----------------------------------------------------------------------------
' Late-bound Shell folder picker. Avoids early reference to Shell32.
'==============================================================================
Private Sub btnBrowseOutput_Click()
    Dim folder As String
    folder = PickFolder("Select output folder for diagnostics file")
    If LenB(folder) > 0 Then txtOutputFolder.Text = folder
End Sub

'==============================================================================
' btnDiagnostics_Click
'==============================================================================
Private Sub btnDiagnostics_Click()
    On Error GoTo Fail

    Me.MousePointer = 11   ' fmMousePointerHourGlass
    btnDiagnostics.Enabled = False

    Dim folder As String
    folder = Trim$(txtOutputFolder.Text)

    Dim path As String
    path = RunDiagnostics(folder)

    Me.MousePointer = 0
    btnDiagnostics.Enabled = True

    If LenB(path) > 0 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Diagnostics written to:" & vbCrLf & path & vbCrLf & vbCrLf & _
                     "Open the containing folder?", _
                     vbInformation Or vbYesNo, APP_NAME)
        If ans = vbYes Then
            On Error Resume Next
            Shell "explorer.exe /select,""" & path & """", vbNormalFocus
            On Error GoTo 0
        End If
    End If
    Exit Sub

Fail:
    Me.MousePointer = 0
    btnDiagnostics.Enabled = True
    MsgBox "Diagnostics failed: " & Err.Number & " - " & Err.Description, _
           vbCritical, APP_NAME
End Sub

'==============================================================================
' btnClose_Click
'==============================================================================
Private Sub btnClose_Click()
    Unload Me
End Sub

'==============================================================================
' PickFolder (private)
' ----------------------------------------------------------------------------
' Late-bound BrowseForFolder via Shell.Application. Returns empty string on
' cancel.
'==============================================================================
Private Function PickFolder(ByVal title As String) As String
    On Error GoTo Fail
    Dim shell As Object, folder As Object
    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, title, 0)
    If folder Is Nothing Then Exit Function
    PickFolder = folder.Self.Path
    Exit Function
Fail:
    PickFolder = ""
End Function
