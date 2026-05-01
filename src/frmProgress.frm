VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress
   Caption         =   "Working..."
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   1
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblPhase
      Caption         =   "Phase: ..."
      ForeColor       =   -2147483630
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   6840
   End
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblPercent
      Alignment       =   1
      Caption         =   "0%"
      Height          =   240
      Left            =   6480
      TabIndex        =   1
      Top             =   480
      Width           =   540
   End
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} pgrFrame
      BorderStyle     =   1
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   480
      Width           =   6240
   End
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} pgrFill
      BackColor       =   12947505
      Height          =   210
      Left            =   195
      TabIndex        =   3
      Top             =   495
      Width           =   30
   End
   Begin {8BD21D40-EC42-11CE-9E0D-00AA006002F3} lstLog
      Height          =   2700
      IntegralHeight  =   0
      Left            =   180
      TabIndex        =   4
      Top             =   840
      Width           =   6840
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCancel
      Caption         =   "Cancel"
      Height          =   330
      Left            =   180
      TabIndex        =   5
      Top             =   3960
      Width           =   1095
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnClose
      Caption         =   "Close"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5925
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmProgress
' ----------------------------------------------------------------------------
' Progress / live-log display. Subscribes to a clsLogger via SubscribeNotify
' and renders accepted entries into the listbox in real time. Exposes a
' Cancel button that flips a flag the orchestrator polls between work units.
'
' Phase 1 surface: SetPhase, SetProgress, Notify, Show/Done. Phase 2 wires
' it to multi-file orchestration; Phases 3-7 hand it off to other phases.
'==============================================================================
Option Explicit

Private mCtx As clsRunContext
Private Const PROGRESS_TRACK_WIDTH As Long = 6210   ' twips inside pgrFrame
Private Const PROGRESS_TRACK_LEFT  As Long = 195

'==============================================================================
' AttachContext
' ----------------------------------------------------------------------------
' Wires the form to a clsRunContext. After this call:
'   - the form subscribes its Notify callback to ctx.Logger so log entries
'     stream live into the listbox
'   - the Cancel button calls ctx.RequestCancel
'==============================================================================
Public Sub AttachContext(ByVal ctx As clsRunContext)
    Set mCtx = ctx
    If Not mCtx Is Nothing Then mCtx.Logger.SubscribeNotify Me
End Sub

'==============================================================================
' SetPhase
' ----------------------------------------------------------------------------
' Update the top status label. Use for human-readable phase hints like
' "Scanning files (3 of 7)".
'==============================================================================
Public Sub SetPhase(ByVal phaseText As String)
    lblPhase.Caption = phaseText
    DoEvents
End Sub

'==============================================================================
' SetProgress
' ----------------------------------------------------------------------------
' Update the visual progress bar. percent is clamped to 0..100. Implemented
' as a coloured fill rectangle whose Width is a fraction of the track width;
' avoids a ProgressBar control dependency on MSCOMCTL.
'==============================================================================
Public Sub SetProgress(ByVal percent As Double)
    If percent < 0# Then percent = 0#
    If percent > 100# Then percent = 100#
    pgrFill.Width = CLng(PROGRESS_TRACK_WIDTH * percent / 100#)
    pgrFill.Left = PROGRESS_TRACK_LEFT
    lblPercent.Caption = Format$(percent, "0") & "%"
    DoEvents
End Sub

'==============================================================================
' Notify
' ----------------------------------------------------------------------------
' Logger sink. Called by clsLogger.AddEntry for every accepted entry.
' Appends the rendered line to the listbox and auto-scrolls to the bottom.
'==============================================================================
Public Sub Notify(ByVal entry As clsLogEntry)
    On Error Resume Next
    lstLog.AddItem entry.AsLine()
    If lstLog.ListCount > 0 Then lstLog.TopIndex = lstLog.ListCount - 1
End Sub

'==============================================================================
' Done
' ----------------------------------------------------------------------------
' Called by the orchestrator when work completes (success or failure).
' Disables Cancel, enables Close, leaves the log visible for review.
'==============================================================================
Public Sub Done()
    btnCancel.Enabled = False
    btnClose.Enabled = True
    lblPhase.Caption = "Done."
End Sub

'==============================================================================
' UserForm_Initialize
'==============================================================================
Private Sub UserForm_Initialize()
    Me.Caption = APP_NAME & " " & APP_VERSION & " - Working..."
    SetProgress 0
End Sub

'==============================================================================
' btnCancel_Click
'==============================================================================
Private Sub btnCancel_Click()
    If Not mCtx Is Nothing Then mCtx.RequestCancel
    btnCancel.Enabled = False
    lblPhase.Caption = "Cancel requested - waiting for current step to finish..."
End Sub

'==============================================================================
' btnClose_Click
'==============================================================================
Private Sub btnClose_Click()
    Unload Me
End Sub

'==============================================================================
' UserForm_Terminate
' ----------------------------------------------------------------------------
' Unsubscribe so the logger doesn't try to push entries into a dead form.
'==============================================================================
Private Sub UserForm_Terminate()
    On Error Resume Next
    If Not mCtx Is Nothing Then mCtx.Logger.SubscribeNotify Nothing
    Set mCtx = Nothing
End Sub
