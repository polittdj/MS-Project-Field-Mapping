'==============================================================================
' frmMain.code.bas - CODE-ONLY FALLBACK
' ----------------------------------------------------------------------------
' This file is NOT meant to be imported on its own. Use it only if
' `frmMain.frm` fails to import into your VBE (e.g. across mixed Office
' versions or in environments that strip the OleObjectBlob reference).
'
' Recovery procedure:
'   1. In the VBE: Insert -> UserForm. Name it `frmMain` (Properties -> Name).
'   2. Set form properties:
'      Caption = "MS Project Field Mapper"
'      Width   ~ 450
'      Height  ~ 230
'      StartUpPosition = 1 (CenterOwner)
'   3. From the Toolbox add and rename the following controls:
'      - Label   `lblTitle`           top, full width
'      - Label   `lblPhase`           below title
'      - Label   `lblHint`            below phase
'      - Label   `lblOutputCaption`   above textbox
'      - TextBox `txtOutputFolder`    main row, leaves room for Browse...
'      - Button  `btnBrowseOutput`    caption "Browse..."
'      - Button  `btnDiagnostics`     bottom-left, caption "Run Diagnostics..."
'      - Button  `btnClose`           bottom-right, caption "Close"
'   4. Double-click the form to enter the code-behind.
'   5. Replace the entire code with the block between BEGIN / END below.
'      Save the project. The form is now functionally identical to frmMain.frm.
'
' ============================== BEGIN PASTE ===================================
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = APP_NAME & " " & APP_VERSION & " - Phase 0"
    lblTitle.Caption = "MS Project Field Mapper"
    lblPhase.Caption = "Phase 0 (Foundation) - diagnostics only. Full mapper UI arrives in later phases."
    lblHint.Caption = "Click Run Diagnostics to write a .txt with host, version and macro-security details."
    lblOutputCaption.Caption = "Output folder (optional - defaults to %TEMP%):"
    btnBrowseOutput.Caption = "Browse..."
    btnDiagnostics.Caption = "Run Diagnostics..."
    btnClose.Caption = "Close"

    Dim why As String
    If Not AssertHostSupported(why) Then
        lblHint.Caption = "WARNING: Host check failed: " & why
    End If
End Sub

Private Sub btnBrowseOutput_Click()
    Dim folder As String
    folder = PickFolder("Select output folder for diagnostics file")
    If LenB(folder) > 0 Then txtOutputFolder.Text = folder
End Sub

Private Sub btnDiagnostics_Click()
    On Error GoTo Fail
    Me.MousePointer = 11
    btnDiagnostics.Enabled = False

    Dim folder As String, path As String
    folder = Trim$(txtOutputFolder.Text)
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

Private Sub btnClose_Click()
    Unload Me
End Sub

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
'================================ END PASTE ==================================
