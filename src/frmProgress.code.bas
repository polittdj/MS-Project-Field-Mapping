'==============================================================================
' frmProgress.code.bas - CODE-ONLY FALLBACK
' ----------------------------------------------------------------------------
' Use only if `frmProgress.frm` fails to import on your VBE. Otherwise ignore
' this file. Manual construction:
'
'   1. VBE: Insert -> UserForm. Set Name = `frmProgress`,
'      Caption = "Working...", Width ~ 480, Height ~ 320.
'   2. Add controls (set Name property exactly as listed):
'        Label  `lblPhase`     full-width across top
'        Label  `lblPercent`   top-right, ~40 wide, right-aligned
'        Label  `pgrFrame`     BorderStyle = 1 (single), thin track for the bar
'        Label  `pgrFill`      BackColor = 12947505 (any contrast colour),
'                              positioned inside pgrFrame, starts width = 30
'        ListBox `lstLog`      large central area, IntegralHeight = False
'        CommandButton `btnCancel` bottom-left, Caption = "Cancel"
'        CommandButton `btnClose`  bottom-right, Caption = "Close",
'                              Enabled = False
'   3. Double-click the form, replace the entire code-behind with the block
'      between BEGIN PASTE and END PASTE below.
'
' ============================== BEGIN PASTE ===================================
Option Explicit

Private mCtx As clsRunContext
Private Const PROGRESS_TRACK_WIDTH As Long = 6210
Private Const PROGRESS_TRACK_LEFT  As Long = 195

Public Sub AttachContext(ByVal ctx As clsRunContext)
    Set mCtx = ctx
    If Not mCtx Is Nothing Then mCtx.Logger.SubscribeNotify Me
End Sub

Public Sub SetPhase(ByVal phaseText As String)
    lblPhase.Caption = phaseText
    DoEvents
End Sub

Public Sub SetProgress(ByVal percent As Double)
    If percent < 0# Then percent = 0#
    If percent > 100# Then percent = 100#
    pgrFill.Width = CLng(PROGRESS_TRACK_WIDTH * percent / 100#)
    pgrFill.Left = PROGRESS_TRACK_LEFT
    lblPercent.Caption = Format$(percent, "0") & "%"
    DoEvents
End Sub

Public Sub Notify(ByVal entry As clsLogEntry)
    On Error Resume Next
    lstLog.AddItem entry.AsLine()
    If lstLog.ListCount > 0 Then lstLog.TopIndex = lstLog.ListCount - 1
End Sub

Public Sub Done()
    btnCancel.Enabled = False
    btnClose.Enabled = True
    lblPhase.Caption = "Done."
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = APP_NAME & " " & APP_VERSION & " - Working..."
    SetProgress 0
End Sub

Private Sub btnCancel_Click()
    If Not mCtx Is Nothing Then mCtx.RequestCancel
    btnCancel.Enabled = False
    lblPhase.Caption = "Cancel requested - waiting for current step to finish..."
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    If Not mCtx Is Nothing Then mCtx.Logger.SubscribeNotify Nothing
    Set mCtx = Nothing
End Sub
'================================ END PASTE ==================================
