VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmVersion_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmVersion_Main"

'VGC 10/27/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Begin Convert:    B {cmdConvert}
' **   Cancel:           C {cmdCancel}
' **   Print:            P {cmdPrintReport}

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private intConvertTimerStep As Integer
Private strCallingProc As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String
        Dim intX As Integer
        Dim arr_varRetVal As Variant
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingProc = Trim(.OpenArgs)

150           DoCmd.Hourglass False
160           Beep

170           intConvertTimerStep = 0
180           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

190           .cmdPrintReport.Visible = False
200           .cmdPrintReport_raised_img.Visible = False
210           .cmdPrintReport_raised_semifocus_dots_img.Visible = False
220           .cmdPrintReport_raised_focus_img.Visible = False
230           .cmdPrintReport_raised_focus_dots_img.Visible = False
240           .cmdPrintReport_sunken_focus_dots_img.Visible = False
250           .cmdPrintReport_raised_img_dis.Visible = False

260           arr_varRetVal = Version_DataXFer("Get", "PathFile")  ' ** Module Function: modVersionConvertFuncs2.
270           If Left(arr_varRetVal(0, 0), 1) <> "#" Then  ' ** Covers both #ERROR and #EMPTY.
280             strTmp01 = arr_varRetVal(0, 0)
290             intLen = Len(strTmp01)
300             If intLen > 90 Then
                  ' ** Break the line after 90 characters, with 96 characters max.
310               .PathFile_TrustData.Height = 510&
320               blnFound = False
330               For intX = 80 To 91
340                 If Mid(strTmp01, intX, 1) = LNK_SEP Then
350                   blnFound = True
360                   strTmp01 = Left(strTmp01, intX) & vbCrLf & Mid(strTmp01, (intX + 1))
370                   Exit For
380                 End If
390               Next
400               If blnFound = False Then
410                 For intX = 80 To 1 Step -1
420                   If Mid(strTmp01, intX, 1) = LNK_SEP Then
430                     blnFound = True
440                     strTmp01 = Left(strTmp01, intX) & vbCrLf & Mid(strTmp01, (intX + 1))
450                     Exit For
460                   End If
470                 Next
480               End If
490             Else
500               .PathFile_TrustData.Height = 285&
510             End If
520             .PathFile_TrustData = strTmp01
530             strTmp01 = arr_varRetVal(1, 0)
540             intLen = Len(strTmp01)
550             If intLen > 90 Then
560               .PathFile_TrustArchive.Height = 510&
570               blnFound = False
580               For intX = 80 To 91
590                 If Mid(strTmp01, intX, 1) = LNK_SEP Then
600                   blnFound = True
610                   strTmp01 = Left(strTmp01, intX) & vbCrLf & Mid(strTmp01, (intX + 1))
620                   Exit For
630                 End If
640               Next
650               If blnFound = False Then
660                 For intX = 80 To 1 Step -1
670                   If Mid(strTmp01, intX, 1) = LNK_SEP Then
680                     blnFound = True
690                     strTmp01 = Left(strTmp01, intX) & vbCrLf & Mid(strTmp01, (intX + 1))
700                     Exit For
710                   End If
720                 Next
730               End If
740             Else
750               .PathFile_TrustArchive.Height = 285&
760             End If

                ' ** Make sure MouseHook.dll is present and accounted for.
770             blnRetVal = MouseWheelOFF(False)  ' ** Module Function: modMouseWheel.
                ' ** Default to mouse wheel on.
780             If blnRetVal = True Then
790               MouseWheelON  ' ** Module Function: modMouseWheel.
800             Else
                  ' ** This seems to be giving us some trouble since some Microsoft update.
                  ' ** Just let it pass for now.
                  ' ** Somewhere, the return of a False isn't caught, and the progress bar
                  ' ** tries to update based on Steps, but the array hasn't been initialized
                  ' ** yet, so it errors with 'Subscript out of range'.
810               blnRetVal = True
820             End If

830             .PathFile_TrustArchive = strTmp01
840             .TAVer_Old = arr_varRetVal(2, 0)
850             .TAVer_Old_RelDate = arr_varRetVal(3, 0)
860             .TAVer_New = AppVersion_Get2  ' ** Module Function: modAppVersionFuncs.
870             .TAVer_New_RelDate = Format(CDate(AppDate_Get), "mm/dd/yyyy")  ' ** Module Function: modAppVersionFuncs.

880             .cmdConvert.SetFocus

890           Else
900             Cancel = -1
910             gblnMessage = False
920           End If

930         Else
940           Cancel = -1
950           gblnMessage = False
960         End If
970       Else
980         Cancel = -1
990         Beep
1000        MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1010      End If
1020    End With

EXITP:
1030    Exit Sub

ERRH:
1040    Select Case ERR.Number
        Case Else
1050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1060    End Select
1070    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1120    intShiftDown = (Shift And acShiftMask) > 0
1130    intAltDown = (Shift And acAltMask) > 0
1140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyEscape
1170        intRetVal = 0
1180        With Me
1190          If .cmdCancel.Enabled = False Then
1200            .cmdCancel.Enabled = True
1210            .cmdCancel.Caption = "&Close"
1220            .cmdCancel.SetFocus
1230            If .cmdCancel_lbl.Visible = True Then
                  ' ** Normal: Arial 10, Bold, Right-Aligned.
1240              .cmdCancel_lbl.FontSize = 8
1250              .cmdCancel_lbl.Caption = "If an error has occurred, click 'Close' and contact Delta Data, Inc."
1260            End If
1270          End If
1280        End With
1290      End Select
1300    End If

        ' ** Ctrl-Shift keys.
1310    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1320      Select Case intRetVal
          Case vbKeyF
1330        With Me
1340          intRetVal = 0
1350          .FocusHolder.SetFocus
1360        End With
1370      End Select
1380    End If

EXITP:
1390    KeyCode = intRetVal
1400    Exit Sub

ERRH:
1410    intRetVal = 0
1420    Select Case ERR.Number
        Case Else
1430      Beep
1440      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & CStr(Erl), vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1450    End Select
1460    Resume EXITP

End Sub

Private Sub Form_Timer()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String
        Dim lngTmp01 As Long

1510    With Me
1520      If gblnSetFocus = True Then
1530        gblnSetFocus = False
1540        strFormName = vbNullString
1550  On Error Resume Next
1560        Do Until strFormName = THIS_NAME
1570          DoCmd.SelectObject acForm, THIS_NAME, False
1580          strFormName = Screen.ActiveForm.Name
1590        Loop
1600  On Error GoTo ERRH
1610      End If
1620      Select Case intConvertTimerStep
          Case 1
1630        .TimerInterval = 0
1640        If IsLoaded("frmMenu_Title", acForm) = True Then  ' ** Module Function: modFileUtilities.
1650          Forms("frmMenu_Title").Visible = False
1660        End If
1670        If .Status3.Visible = False Then .Status3.Visible = True
1680        .Status3.SetFocus
1690        .Status3.SelStart = 0
1700        .Status3.SelLength = 0
1710        intConvertTimerStep = 2
1720        .TimerInterval = 2000
1730      Case 2
1740        .TimerInterval = 0&
1750        If IsLoaded("frmMenu_Title", acForm) = True Then  ' ** Module Function: modFileUtilities.
1760          Forms("frmMenu_Title").Visible = False
1770        End If
1780        .ProgBar_lbl1.Visible = False
1790        .ProgBar_lbl2.Caption = "FINISHED!"
1800        .ProgBar_lbl2.TextAlign = acTextAlignCenter
1810        intConvertTimerStep = 3
1820        .TimerInterval = 350&
1830      Case 99
1840        .TimerInterval = 0&
1850        .cmdCancel.SetFocus
1860      Case Else
1870        If IsLoaded("frmMenu_Title", acForm) = True Then  ' ** Module Function: modFileUtilities.
1880          Forms("frmMenu_Title").Visible = False
1890        End If
1900        If intConvertTimerStep < 8 Then
1910          intConvertTimerStep = intConvertTimerStep + 1
1920          .ProgBar_lbl2.Visible = Not (.ProgBar_lbl2.Visible)
1930        Else
1940          If .ProgBar_lbl2.Visible = False Then
1950            .TimerInterval = 0
1960            .ProgBar_lbl2.Visible = True
1970            lngTmp01 = .cmdCancel_lbl.Left   ' ** Label will span left margin to current, closed position.
1980            .cmdCancel_lbl.Left = 0&         ' ** Move label all the way to the margin.
1990            .cmdCancel_lbl.Width = lngTmp01  ' ** Widen it up to its original position.
2000            .cmdCancel_lbl.Caption = "'Click 'Continue' To Enter Trust Accountant"  ' ** It's right-aligned.
2010            .cmdCancel_lbl.FontSize = 10
2020            .cmdCancel_lbl.Visible = True
2030            intConvertTimerStep = 99
2040            gblnSetFocus = True
2050            .TimerInterval = 100&
2060          Else
2070            .TimerInterval = 0&
2080            .ProgBar_lbl2.Visible = False
2090          End If
2100        End If
2110      End Select
2120    End With

EXITP:
2130    Exit Sub

ERRH:
2140    Select Case ERR.Number
        Case Else
2150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2160    End Select
2170    Resume EXITP

End Sub

Private Sub Form_Close()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Close"

2210    With Me
2220      If .IntroMsg_lbl_img.Visible = False Then  ' ** Simplified way of telling whether this is intro or status.
2230        gstrReturningForm = THIS_NAME
2240        If gblnDev_Debug = False Then
2250          Forms("frmMenu_Title").Visible = True
              ' ** intConverResponse is set by ConversionCheck_Response, below, which
              ' ** in turn is called at the end of ConversionCheck() in modVersionConvertFuncs1.
2260          If gintConvertResponse < 0 Then
2270            Forms("frmMenu_Title").OpenContinues False  ' ** Form Procedure: frmMenu_Title.
2280          Else
2290            Forms("frmMenu_Title").OpenContinues True  ' ** Form Procedure: frmMenu_Title.
2300          End If
2310        Else
2320          DoCmd.OpenForm "frmPleaseWait", , , , , , THIS_NAME
              'Version_Convert_RunAllStart glngInstance  ' ** Module Function: modVersionDocFuncs.
2330        End If
2340      End If
2350    End With

EXITP:
2360    Exit Sub

ERRH:
2370    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Ignore.
2380    Case Else
2390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2400    End Select
2410    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2510    With Me
2520      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
2530        Select Case blnPrintReport_Focus
            Case True
2540          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2550          .cmdPrintReport_raised_img.Visible = False
2560        Case False
2570          .cmdPrintReport_raised_img.Visible = True
2580          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2590        End Select
2600        .cmdPrintReport_raised_focus_img.Visible = False
2610        .cmdPrintReport_raised_focus_dots_img.Visible = False
2620        .cmdPrintReport_sunken_focus_dots_img.Visible = False
2630        .cmdPrintReport_raised_img_dis.Visible = False
2640      End If
2650    End With

EXITP:
2660    Exit Sub

ERRH:
2670    Select Case ERR.Number
        Case Else
2680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2690    End Select
2700    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

2810    With Me
2820      gblnMessage = False
2830      .Visible = False
2840      DoEvents
2850      DoCmd.Close acForm, THIS_NAME
2860    End With

EXITP:
2870    Exit Sub

ERRH:
2880    Select Case ERR.Number
        Case Else
2890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2900    End Select
2910    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3020    intShiftDown = (Shift And acShiftMask) > 0
3030    intAltDown = (Shift And acAltMask) > 0
3040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3060      Select Case intRetVal
          Case vbKeyTab
3070        intRetVal = 0
3080        Me.PathFile_TrustData.SetFocus
3090      End Select
3100    End If

EXITP:
3110    KeyCode = intRetVal
3120    Exit Sub

ERRH:
3130    intRetVal = 0
3140    Select Case ERR.Number
        Case Else
3150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3160    End Select
3170    Resume EXITP

End Sub

Private Sub cmdConvert_Click()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdConvert_Click"

        ' ** Close with gblnMessage still True.
3210    DoCmd.Close acForm, THIS_NAME

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub cmdConvert_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdConvert_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3320    intShiftDown = (Shift And acShiftMask) > 0
3330    intAltDown = (Shift And acAltMask) > 0
3340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shift keys.
3350    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3360      Select Case intRetVal
          Case vbKeyTab
3370        intRetVal = 0
3380        Me.PathFile_TrustArchive.SetFocus
3390      End Select
3400    End If

EXITP:
3410    KeyCode = intRetVal
3420    Exit Sub

ERRH:
3430    intRetVal = 0
3440    Select Case ERR.Number
        Case Else
3450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3460    End Select
3470    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

3510    With Me
3520      blnPrintReport_Focus = True
3530      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3540      .cmdPrintReport_raised_img.Visible = False
3550      .cmdPrintReport_raised_focus_img.Visible = False
3560      .cmdPrintReport_raised_focus_dots_img.Visible = False
3570      .cmdPrintReport_sunken_focus_dots_img.Visible = False
3580      .cmdPrintReport_raised_img_dis.Visible = False
3590    End With

EXITP:
3600    Exit Sub

ERRH:
3610    Select Case ERR.Number
        Case Else
3620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3630    End Select
3640    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

3710    With Me
3720      blnPrintReport_MouseDown = True
3730      .cmdPrintReport_sunken_focus_dots_img.Visible = True
3740      .cmdPrintReport_raised_img.Visible = False
3750      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3760      .cmdPrintReport_raised_focus_img.Visible = False
3770      .cmdPrintReport_raised_focus_dots_img.Visible = False
3780      .cmdPrintReport_raised_img_dis.Visible = False
3790    End With

EXITP:
3800    Exit Sub

ERRH:
3810    Select Case ERR.Number
        Case Else
3820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3830    End Select
3840    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

3910    With Me
3920      strDocName = "rptVersion_Convert"
3930      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
3940        DoCmd.OpenReport strDocName, acViewPreview, , "[vercnv_id] = " & CStr(.vercnv_id)
3950        DoCmd.Maximize
3960        DoCmd.RunCommand acCmdFitToWindow
3970      Else
            '##GTR_Ref: rptVersion_Convert
3980        DoCmd.OpenReport strDocName, acViewNormal, , "[vercnv_id] = " & CStr(.vercnv_id)
3990      End If
4000    End With

EXITP:
4010    Exit Sub

ERRH:
4020    Select Case ERR.Number
        Case Else
4030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4040    End Select
4050    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

4110    With Me
4120      If blnPrintReport_MouseDown = False Then
4130        Select Case blnPrintReport_Focus
            Case True
4140          .cmdPrintReport_raised_focus_dots_img.Visible = True
4150          .cmdPrintReport_raised_focus_img.Visible = False
4160        Case False
4170          .cmdPrintReport_raised_focus_img.Visible = True
4180          .cmdPrintReport_raised_focus_dots_img.Visible = False
4190        End Select
4200        .cmdPrintReport_raised_img.Visible = False
4210        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4220        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4230        .cmdPrintReport_raised_img_dis.Visible = False
4240      End If
4250    End With

EXITP:
4260    Exit Sub

ERRH:
4270    Select Case ERR.Number
        Case Else
4280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4290    End Select
4300    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

4410    With Me
4420      .cmdPrintReport_raised_focus_dots_img.Visible = True
4430      .cmdPrintReport_raised_img.Visible = False
4440      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4450      .cmdPrintReport_raised_focus_img.Visible = False
4460      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4470      .cmdPrintReport_raised_img_dis.Visible = False
4480      blnPrintReport_MouseDown = False
4490    End With

EXITP:
4500    Exit Sub

ERRH:
4510    Select Case ERR.Number
        Case Else
4520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4530    End Select
4540    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

4610    With Me
4620      .cmdPrintReport_raised_img.Visible = True
4630      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4640      .cmdPrintReport_raised_focus_img.Visible = False
4650      .cmdPrintReport_raised_focus_dots_img.Visible = False
4660      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4670      .cmdPrintReport_raised_img_dis.Visible = False
4680      blnPrintReport_Focus = False
4690    End With

EXITP:
4700    Exit Sub

ERRH:
4710    Select Case ERR.Number
        Case Else
4720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4730    End Select
4740    Resume EXITP

End Sub

Private Sub PathFile_TrustData_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "PathFile_TrustData_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4820    intShiftDown = (Shift And acShiftMask) > 0
4830    intAltDown = (Shift And acAltMask) > 0
4840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shift keys.
4850    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4870        intRetVal = 0
4880        Me.cmdCancel.SetFocus
4890      End Select
4900    End If

EXITP:
4910    KeyCode = intRetVal
4920    Exit Sub

ERRH:
4930    intRetVal = 0
4940    Select Case ERR.Number
        Case Else
4950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4960    End Select
4970    Resume EXITP

End Sub

Private Sub Status2_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "Status2_lbl_MouseMove"

5010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5020    Exit Sub

ERRH:
5030    Select Case ERR.Number
        Case Else
5040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5050    End Select
5060    Resume EXITP

End Sub

Private Sub Status3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "Status3_MouseMove"

5110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5120    Exit Sub

ERRH:
5130    Select Case ERR.Number
        Case Else
5140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5150    End Select
5160    Resume EXITP

End Sub

Private Sub PathFile_TrustArchive_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "PathFile_TrustArchive_box_MouseMove"

5210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5220    Exit Sub

ERRH:
5230    Select Case ERR.Number
        Case Else
5240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5250    End Select
5260    Resume EXITP

End Sub

Private Sub PathFile_TrustArchive_KeyDown(KeyCode As Integer, Shift As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "PathFile_TrustArchive_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5320    intShiftDown = (Shift And acShiftMask) > 0
5330    intAltDown = (Shift And acAltMask) > 0
5340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5370        intRetVal = 0
5380        Me.cmdConvert.SetFocus
5390      End Select
5400    End If

EXITP:
5410    KeyCode = intRetVal
5420    Exit Sub

ERRH:
5430    intRetVal = 0
5440    Select Case ERR.Number
        Case Else
5450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5460    End Select
5470    Resume EXITP

End Sub

Public Sub ConversionCheck_Response(msgResponse As Integer)
' ** Called at the end of ConversionCheck() in modVersionConvertFuncs1.

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "ConversionCheck_Response"

5510    With Me
5520      gintConvertResponse = msgResponse
          ' ** Once conversion is over, we want to sit on this window till they close it.
          ' ** The title screen, however, needs to finish its OnOpen event.
          ' ** I tried to make it invisible there, but that didn't work,
          ' ** so let's see how this works!
5530      intConvertTimerStep = 1
5540      .TimerInterval = 100&
5550    End With

EXITP:
5560    Exit Sub

ERRH:
5570    Select Case ERR.Number
        Case Else
5580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5590    End Select
5600    Resume EXITP

End Sub
