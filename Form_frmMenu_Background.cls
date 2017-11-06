VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Background"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Background"

'VGC 10/11/2012: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Save Size:        S {cmdSaveSize}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Form:       X {cmdCloseForm}

' ** See gblnDev_NoAppBackground.

Private blnShowDev As Boolean, blnOpenNext As Boolean
Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnOpenNext = False

160           blnShowDev = False
170           If blnShowDev = True Then
180             .cmdSaveSize.Visible = True
190             .top_left_hline.Visible = True
200             .top_left_vline.Visible = True
210             .top_right_hline.Visible = True
220             .top_right_vline.Visible = True
230             .bot_left_hline.Visible = True
240             .bot_left_vline.Visible = True
250             .bot_right_hline.Visible = True
260             .bot_right_vline.Visible = True
270           End If

280           HelpBar_Close  ' ** Module Function: modWindowFunctions.

              '.top_left_hline
              '.top_left_vline
              '.top_right_hline
              '.top_right_vline
              '.bot_right_hline
              '.bot_right_vline
              '.bot_left_hline
              '.bot_left_vline

              ' ** Variables are fed empty, then populated ByRef.
290           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

300         Else
310           Cancel = -1
320         End If
330       Else
340         Cancel = -1
350       End If
360     End With

370     If Cancel = -1 Then
380       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
390     End If

EXITP:
400     Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strUsername As String
        Dim intRetVal As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Only        X {Close}

        ' ** Ctrl-Shift keys.
550     If intCtrlDown And (Not intAltDown) And intShiftDown Then
560       Select Case intRetVal
          Case vbKeyX
            ' ** Close form without moving to another.
570         strUsername = GetUserName  ' ** Module Function: modFileUtilities.
580         If strUsername = gstrDevUserName Then
590           intRetVal = 0
600           SetOption_Dev  ' ** Module Function: modUtilities.
610           DoCmd.Close acForm, THIS_NAME
620         End If
630       End Select
640     End If

EXITP:
650     KeyCode = intRetVal
660     Exit Sub

ERRH:
670     intRetVal = 0
680     Select Case ERR.Number
        Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub Form_Timer()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngXBorders As Long, lngYBorders As Long
        Dim lngNewWidth As Long, lngNewHeight As Long
        Dim strFormName As String
        Dim lngTmp02 As Long

810     With Me
820       .TimerInterval = 0&
830       If gblnSetFocus = True Then
840         gblnSetFocus = False
850         lngTmp02 = 0&
860         strFormName = vbNullString
870   On Error Resume Next
880         Do Until strFormName = THIS_NAME
890           lngTmp02 = lngTmp02 + 1&
900           DoCmd.SelectObject acForm, THIS_NAME, False
910           strFormName = Screen.ActiveForm.Name
920           If lngTmp02 > 100& Then Exit Do
930         Loop
940   On Error GoTo ERRH
950         DoEvents
960         CmdBars_Hide True  ' ** Module Procedure: modWindowFunctions.
970         DoEvents

980         GetAppDimensions lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Procedure: modWindowFunctions.
990         GetAppBorders lngXBorders, lngYBorders  ' ** Module Procedure: modWindowFunctions.

            ' ** These aren't adjusted for TwipsPerPixel!
1000        lngNewWidth = ((lngWidth - lngXBorders) - 60&)
1010        lngNewHeight = ((lngHeight - lngYBorders) - 750&)

            'DoCmd.MoveSize 0&, 0&, lngNewWidth&, lngNewHeight
            'DoCmd.MoveSize 0&, 0&, 21540&, 11970&
1020        DoCmd.Maximize  ' ** This leaves a bar at top!
1030        DoEvents

1040        blnOpenNext = True
1050        .TimerInterval = 250&

1060      ElseIf blnOpenNext = True Then
1070        DoCmd.Restore
1080        gblnSetFocus = True
1090        DoCmd.OpenForm "frmMenu_Post", , , , , , "frmMenu_Main"
1100      End If

1110    End With

EXITP:
1120    Exit Sub

ERRH:
1130    Select Case ERR.Number
        Case Else
1140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1150    End Select
1160    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1210    If strCallingForm = "frmMenu_Title" Then
1220      Beep
1230    End If

EXITP:
1240    Exit Sub

ERRH:
1250    Select Case ERR.Number
        Case Else
1260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1270    End Select
1280    Resume EXITP

End Sub

Private Sub cmdSaveSize_Click()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSaveSize_Click"

        ' ** Variables are fed empty, then populated ByRef.
1310    GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

1320    Debug.Print "'lngFrm_Top = " & CStr(lngFrm_Top)
1330    Debug.Print "'lngFrm_Left = " & CStr(lngFrm_Left)
1340    Debug.Print "'lngFrm_Width = " & CStr(lngFrm_Width)
1350    Debug.Print "'lngFrm_Height = " & CStr(lngFrm_Height)

EXITP:
1360    Exit Sub

ERRH:
1370    Select Case ERR.Number
        Case Else
1380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1390    End Select
1400    Resume EXITP

End Sub
