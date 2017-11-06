VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSiteMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmSiteMap"

'VGC 07/23/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Print Report:     P {cmdPrintReport}
' **   Exit:             X {cmdClose}
' **   Save Size/Pos:    Z {chkSaveSizePos}

' ** Shortcut Ctrl keys responsive from this form:
' **   Journal Map:      J {Level02_02_01_s02_arrow on frmSiteMap_Sub}

' *******************************************************
' ** This form can be left open while navigating menus.
' *******************************************************

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngDetail_Height As Long, lngSub_Width As Long, lngSub_Height As Long
Private lngClose_Left As Long, lngSizable_Offset As Long, lngMin_Width As Long, lngMin_Height As Long
Private lngFrm_Left_Orig As Long, lngFrm_Top_Orig As Long, lngFrm_Width_Orig As Long, lngFrm_Height_Orig As Long
Private lngSaveSizePos_Top As Long, lngOptLbl_Offset As Long, lngDetailHLine01_Top As Long

Private blnGoingToReport As Boolean
Private lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean
        Dim lngTmp01 As Long

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

170           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
180           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

200           .form_width_line.Left = 0&  ' ** Just get these out of the way.
210           .detail_height_line.Top = 0&

220           lngForm_Width = .Width
230           lngDetail_Height = .Detail.Height
240           lngSub_Width = .frmSiteMap_Sub.Width
250           lngSub_Height = .frmSiteMap_Sub.Height
260           lngClose_Left = .cmdClose.Left
270           lngSizable_Offset = (.Width - .Sizable_lbl1.Left)
280           lngSaveSizePos_Top = .chkSaveSizePos.Top
290           lngOptLbl_Offset = (.chkSaveSizePos_lbl.Left - .chkSaveSizePos.Left)
300           lngDetailHLine01_Top = .Detail_hline01.Top

310           lngMin_Width = (3& * 1440&)  ' ** 1440 Twips per inch is independent of Twips per pixel.
320           lngMin_Height = (2& * 1440&)  ' ** This is Detail height.

330           If .PopUp = False Then
340             If IsLoaded(strCallingForm, acForm) = True Then
350               Forms(strCallingForm).Visible = False
360             End If
370           End If

              ' ** Variables are fed empty, then populated ByRef.
380           GetFormDimensions Me, lngFrm_Left_Orig, lngFrm_Top_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig  ' ** Module Function: modWindowFunctions.

390           FrmDimSave THIS_NAME, lngFrm_Left_Orig, lngFrm_Top_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig  ' ** Module Function: modMonitorFuncs.

              ' ** Last saved position and size.
400           If .frm_left = 0& And .frm_top = 0& And .frm_width = 0& And .frm_height = 0& Then
                ' ** Don't resize.
410           ElseIf .frm_height > 0 Then
420             If .frm_left < 0 Or .frm_top < 0 Then
430               .frm_left = 0&
440               .frm_top = 0&
450               DoCmd.Maximize
460             Else
470               lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
480               lngMonitorNum = 1&: lngTmp01 = 0&
490               EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
500               If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
510               If lngMonitorNum = 1& Then lngTmp01 = .frm_top
520               DoCmd.SelectObject acForm, THIS_NAME, False
530               DoEvents
540               DoCmd.MoveSize .frm_left, lngTmp01, .frm_width, .frm_height  '.frm_top
550               If lngMonitorNum > 1& Then
560                 LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
570               End If
580             End If
590           End If

600           chkSaveSizePos_AfterUpdate  ' ** Procedure: Below.

610           blnIsOpen = False

620         Else
630           Cancel = -1
640         End If
650       Else
660         Cancel = -1
670       End If
680     End With

690     If Cancel = -1 Then
700       Beep
710       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
720     End If

EXITP:
730     Exit Sub

ERRH:
740     Select Case ERR.Number
        Case Else
750       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
760     End Select
770     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim Cancel As Integer
        Dim intRetVal As Integer

810     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
820     intShiftDown = (Shift And acShiftMask) > 0
830     intAltDown = (Shift And acAltMask) > 0
840     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Journal Map:      J {Level02_02_01_s02_arrow on frmSiteMap_Sub}

        ' ** Plain keys.
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyEscape
870         intRetVal = 0
880         cmdClose_Click  ' ** Procedure: Below.
890       End Select
900     End If

        ' ** Alt keys.
910     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
920       Select Case intRetVal
          Case vbKeyX
930         intRetVal = 0
940         cmdClose_Click  ' ** Procedure: Below.
950       End Select
960     End If

        ' ** Ctrl keys.
970     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
980       Select Case intRetVal
          Case vbKeyJ
990         With Me
1000          intRetVal = 0
1010          Cancel = 0
1020          .frmSiteMap_Sub.SetFocus
1030          .frmSiteMap_Sub.Form.Level02_02_01_s02_arrow_DblClick Cancel  ' ** Form Procedure: frmSiteMap_Sub.
1040        End With
1050      End Select
1060    End If

EXITP:
1070    KeyCode = intRetVal
1080    Exit Sub

ERRH:
1090    intRetVal = 0
1100    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1110    Case Else
1120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1130    End Select
1140    Resume EXITP

End Sub

Private Sub Form_Resize()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Resize"

        Dim lngDiff_Width As Long, lngDiff_Height As Long
        Dim lngTmp01 As Long, lngTmp02 As Long

1210    With Me

          ' ** Variables are fed empty, then populated ByRef.
1220      GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

1230      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
1240        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
1250      End If

1260      If lngFrm_Width_Orig > 0& Then

            ' ** Width.
1270        If lngFrm_Width > lngFrm_Width_Orig Then
1280          lngDiff_Width = (lngFrm_Width - lngFrm_Width_Orig)
1290          .Width = (lngForm_Width + lngDiff_Width)
              ' ** Error: 6
              ' ** Overflow.
1300          .frmSiteMap_Sub.Width = (lngSub_Width + lngDiff_Width)
1310          .frmSiteMap_Sub_box.Width = (.frmSiteMap_Sub.Width + (2& * lngTpp))
1320          .cmdClose.Left = (lngClose_Left + lngDiff_Width)
1330          .Sizable_lbl1.Left = ((lngForm_Width + lngDiff_Width) - lngSizable_Offset)
1340          .Sizable_lbl2.Left = ((lngForm_Width + lngDiff_Width) - lngSizable_Offset)
1350          .Header_vline01.Left = (lngForm_Width + lngDiff_Width)
1360          .Header_vline02.Left = (lngForm_Width + lngDiff_Width)
1370          .Detail_vline01.Left = (lngForm_Width + lngDiff_Width)
1380          .Detail_vline02.Left = (lngForm_Width + lngDiff_Width)
1390          .Footer_vline01.Left = (lngForm_Width + lngDiff_Width)
1400          .Footer_vline02.Left = (lngForm_Width + lngDiff_Width)
1410          .Header_hline01.Width = (lngForm_Width + lngDiff_Width)
1420          .Header_hline02.Width = (lngForm_Width + lngDiff_Width)
1430          .Detail_hline01.Width = (lngForm_Width + lngDiff_Width)
1440          .Detail_hline02.Width = (lngForm_Width + lngDiff_Width)
1450          .Footer_hline01.Width = (lngForm_Width + lngDiff_Width)
1460          .Footer_hline02.Width = (lngForm_Width + lngDiff_Width)
1470        ElseIf lngFrm_Width < lngFrm_Width_Orig Then
1480          lngDiff_Width = (lngFrm_Width_Orig - lngFrm_Width)
1490          If (lngForm_Width - lngDiff_Width) >= lngMin_Width Then
1500            .frmSiteMap_Sub.Width = (lngSub_Width - lngDiff_Width)
1510            .frmSiteMap_Sub_box.Width = (.frmSiteMap_Sub.Width + (2& * lngTpp))
1520            .cmdClose.Left = (lngClose_Left - lngDiff_Width)
1530            .Sizable_lbl1.Left = ((lngForm_Width - lngDiff_Width) - lngSizable_Offset)
1540            .Sizable_lbl2.Left = ((lngForm_Width - lngDiff_Width) - lngSizable_Offset)
1550            .Header_vline01.Left = (lngForm_Width - lngDiff_Width)
1560            .Header_vline02.Left = (lngForm_Width - lngDiff_Width)
1570            .Detail_vline01.Left = (lngForm_Width - lngDiff_Width)
1580            .Detail_vline02.Left = (lngForm_Width - lngDiff_Width)
1590            .Footer_vline01.Left = (lngForm_Width - lngDiff_Width)
1600            .Footer_vline02.Left = (lngForm_Width - lngDiff_Width)
1610            .Header_hline01.Width = (lngForm_Width - lngDiff_Width)
1620            .Header_hline02.Width = (lngForm_Width - lngDiff_Width)
1630            .Detail_hline01.Width = (lngForm_Width - lngDiff_Width)
1640            .Detail_hline02.Width = (lngForm_Width - lngDiff_Width)
1650            .Footer_hline01.Width = (lngForm_Width - lngDiff_Width)
1660            .Footer_hline02.Width = (lngForm_Width - lngDiff_Width)
1670            .Width = (lngForm_Width - lngDiff_Width)
1680          End If
1690        Else
1700          .Width = lngForm_Width
1710          .frmSiteMap_Sub.Width = lngSub_Width
1720          .frmSiteMap_Sub_box.Width = (.frmSiteMap_Sub.Width + (2& * lngTpp))
1730          .cmdClose.Left = lngClose_Left
1740          .Sizable_lbl1.Left = (lngForm_Width - lngSizable_Offset)
1750          .Sizable_lbl2.Left = (lngForm_Width - lngSizable_Offset)
1760          .Header_vline01.Left = lngForm_Width
1770          .Header_vline02.Left = lngForm_Width
1780          .Detail_vline01.Left = lngForm_Width
1790          .Detail_vline02.Left = lngForm_Width
1800          .Footer_vline01.Left = lngForm_Width
1810          .Footer_vline02.Left = lngForm_Width
1820          .Header_hline01.Width = lngForm_Width
1830          .Header_hline02.Width = lngForm_Width
1840          .Detail_hline01.Width = lngForm_Width
1850          .Detail_hline02.Width = lngForm_Width
1860          .Footer_hline01.Width = lngForm_Width
1870          .Footer_hline02.Width = lngForm_Width
1880          .Width = lngForm_Width
1890        End If

            ' ** Height.
1900        If lngFrm_Height > lngFrm_Height_Orig Then
1910          lngDiff_Height = (lngFrm_Height - lngFrm_Height_Orig)
1920          .Detail.Height = (lngDetail_Height + lngDiff_Height)
1930          .chkSaveSizePos.Top = (lngSaveSizePos_Top + lngDiff_Height)
1940          .chkSaveSizePos_lbl.Top = (.chkSaveSizePos.Top - (2& * lngTpp))
1950          .Detail_hline01.Top = (lngDetailHLine01_Top + lngDiff_Height)
1960          .Detail_hline02.Top = (.Detail_hline01.Top + lngTpp)
1970          .Detail_vline01.Top = .Detail_hline01.Top
1980          .Detail_vline02.Top = .Detail_hline01.Top
1990          .frmSiteMap_Sub.Height = (lngSub_Height + lngDiff_Height)
2000          .frmSiteMap_Sub_box.Height = (.frmSiteMap_Sub.Height + (2& * lngTpp))
2010        ElseIf lngFrm_Height < lngFrm_Height_Orig Then
2020          lngDiff_Height = (lngFrm_Height_Orig - lngFrm_Height)
2030          If (lngDetail_Height - lngDiff_Height) >= lngMin_Height Then
2040            .frmSiteMap_Sub.Height = (lngSub_Height - lngDiff_Height)
2050            .frmSiteMap_Sub_box.Height = (.frmSiteMap_Sub.Height + (2& * lngTpp))
2060            .Detail_hline01.Top = (lngDetailHLine01_Top - lngDiff_Height)
2070            .Detail_hline02.Top = (.Detail_hline01.Top + lngTpp)
2080            .Detail_vline01.Top = .Detail_hline01.Top
2090            .Detail_vline02.Top = .Detail_hline01.Top
2100            If ((lngSaveSizePos_Top - lngDiff_Height) + .chkSaveSizePos.Height) > .Detail.Height Then
                  ' ** It really shouldn't be a problem!
2110              lngTmp01 = ((lngSaveSizePos_Top - lngDiff_Height) + .chkSaveSizePos.Height)
2120              lngTmp02 = (lngTmp01 - Detail.Height)
2130              .chkSaveSizePos.Top = ((lngSaveSizePos_Top - lngDiff_Height) - lngTmp02)
2140            Else
2150              .chkSaveSizePos.Top = (lngSaveSizePos_Top - lngDiff_Height)
2160            End If
2170            .chkSaveSizePos_lbl.Top = (.chkSaveSizePos.Top - (2& * lngTpp))
2180            .Detail.Height = (lngDetail_Height - lngDiff_Height)
2190          End If
2200        Else
2210          .Detail.Height = lngDetail_Height
2220          .chkSaveSizePos.Top = lngSaveSizePos_Top
2230          .chkSaveSizePos_lbl.Top = (.chkSaveSizePos.Top - (2& * lngTpp))
2240          .Detail_hline01.Top = lngDetailHLine01_Top
2250          .Detail_hline02.Top = (.Detail_hline01.Top + lngTpp)
2260          .Detail_vline01.Top = .Detail_hline01.Top
2270          .Detail_vline02.Top = .Detail_hline01.Top
2280          .frmSiteMap_Sub.Height = lngSub_Height
2290          .frmSiteMap_Sub_box.Height = (.frmSiteMap_Sub.Height + (2& * lngTpp))
2300          .Detail.Height = lngDetail_Height
2310        End If

2320      End If

2330    End With

EXITP:
2340    Exit Sub

ERRH:
2350    Select Case ERR.Number
        Case Else
2360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2370    End Select
2380    Resume EXITP

End Sub

Private Sub Form_Timer()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2410    With Me
2420      .TimerInterval = 0&
2430      If gblnSetFocus = True Then
2440        gblnSetFocus = False
2450        strFormName = vbNullString
2460  On Error Resume Next
2470        Do Until strFormName = THIS_NAME
2480          DoCmd.SelectObject acForm, THIS_NAME, False
2490          strFormName = Screen.ActiveForm.Name
2500        Loop
2510  On Error GoTo ERRH
2520        .frmSiteMap_Sub.SetFocus
2530      End If
2540      If gblnGoToReport = True Then
2550        Select Case blnGoingToReport
            Case True
2560          DoCmd.Hourglass False  ' ** Should already be off.
2570          gblnGoToReport = False
2580          blnGoingToReport = False
2590          .GoToReport_arw_pr_img.Visible = False
2600        Case False
2610          blnGoingToReport = True
2620          DoCmd.Hourglass True  ' ** Make sure it's still running.
2630          DoEvents
2640          .GoToReport_arw_pr_img.Visible = True
2650          DoEvents
2660          Forms(strCallingForm).GTRSite_Off
2670          DoCmd.SelectObject acForm, THIS_NAME
2680          Beep
2690          DoCmd.Hourglass False
2700          DoEvents
2710          .TimerInterval = (GTR_WAIT * 2)
2720        End Select
2730      End If
2740    End With

EXITP:
2750    Exit Sub

ERRH:
2760    Select Case ERR.Number
        Case Else
2770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2780    End Select
2790    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String
        Dim lngFrms As Long
        Dim lngX As Long

2810    With Me

2820      If Reports.Count > 0 Then
2830        Do While Reports.Count > 0
2840          DoCmd.Close acReport, Reports(0).Name
2850          DoEvents
2860        Loop
2870      End If

2880      Select Case .chkSaveSizePos
          Case True

            ' ** Variables are fed empty, then populated ByRef.
2890        GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

2900        .frm_top = lngFrm_Top
2910        .frm_left = lngFrm_Left
2920        .frm_width = lngFrm_Width
2930        .frm_height = lngFrm_Height

2940      Case False
2950        .frm_top = 0&
2960        .frm_left = 0&
2970        .frm_width = 0&
2980        .frm_height = 0&
2990        .frmj_top = 0&
3000        .frmj_left = 0&
3010        .frmj_width = 0&
3020        .frmj_height = 0&
3030      End Select

3040      strDocName = "frmSiteMap_Journal"
3050      If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
3060        DoCmd.Close acForm, strDocName
3070      End If
3080      DoEvents

3090      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3100      Select Case strCallingForm
          Case "frmMenu_Main"
3110        gstrReturningForm = THIS_NAME
3120        gblnSetFocus = True
3130        Select Case IsLoaded(strCallingForm, acForm)  ' ** Module Function: modFileUtilities.
            Case True
3140          If Forms(strCallingForm).Visible = False Then
3150            Forms(strCallingForm).Visible = True
3160          End If
3170          DoCmd.SelectObject acForm, strCallingForm, False
3180          Forms(strCallingForm).TimerInterval = 100&
3190        Case False
              ' ** Since this form can be left open, just return to the lastest form.
3200          lngFrms = Forms.Count
3210          strDocName = vbNullString
3220          For lngX = (lngFrms - 1&) To 0&
3230            If Forms(lngX).Name <> THIS_NAME And Forms(lngX).Name <> "frmMenu_Background" Then
3240              strDocName = Forms(lngX).Name
3250              Exit For
3260            End If
3270          Next
3280          If strDocName <> vbNullString Then
3290            DoCmd.SelectObject acForm, strDocName, False
3300            Forms(strDocName).TimerInterval = 100&
3310          Else
3320            DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Title"
3330          End If
3340        End Select
3350      End Select

3360    End With

EXITP:
3370    Exit Sub

ERRH:
3380    Select Case ERR.Number
        Case Else
3390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3400    End Select
3410    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3510    With Me
3520      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
3530        Select Case blnPrintReport_Focus
            Case True
3540          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3550          .cmdPrintReport_raised_img.Visible = False
3560        Case False
3570          .cmdPrintReport_raised_img.Visible = True
3580          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3590        End Select
3600        .cmdPrintReport_raised_focus_img.Visible = False
3610        .cmdPrintReport_raised_focus_dots_img.Visible = False
3620        .cmdPrintReport_sunken_focus_dots_img.Visible = False
3630        .cmdPrintReport_raised_img_dis.Visible = False
3640      End If

3650    End With

EXITP:
3660    Exit Sub

ERRH:
3670    Select Case ERR.Number
        Case Else
3680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3690    End Select
3700    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Public Sub cmdClose_Click()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3910    DoCmd.Hourglass True
3920    DoEvents

3930    DoCmd.Close acForm, THIS_NAME

EXITP:
3940    Exit Sub

ERRH:
3950    Select Case ERR.Number
        Case Else
3960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3970    End Select
3980    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4020    intShiftDown = (Shift And acShiftMask) > 0
4030    intAltDown = (Shift And acAltMask) > 0
4040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4060      Select Case intRetVal
          Case vbKeyTab
4070        With Me
4080          intRetVal = 0
4090          .frmSiteMap_Sub.SetFocus
4100          .frmSiteMap_Sub.Form.FocusHolder01.SetFocus
4110        End With
4120      End Select
4130    End If

        ' ** Shift keys.
4140    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4150      Select Case intRetVal
          Case vbKeyTab
4160        With Me
4170          intRetVal = 0
4180          If .cmdPrintReport.Enabled = True Then
4190            .cmdPrintReport.SetFocus
4200          Else
4210            .chkSaveSizePos.SetFocus
4220          End If
4230        End With
4240      End Select
4250    End If

EXITP:
4260    KeyCode = intRetVal
4270    Exit Sub

ERRH:
4280    intRetVal = 0
4290    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4300    Case Else
4310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4320    End Select
4330    Resume EXITP

End Sub

Private Sub frmSiteMap_Sub_Enter()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "frmSiteMap_Sub_Enter"

        Dim Button As Integer, Shift As Integer, X As Single, Y As Single

4410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

4510    With Me
4520      blnPrintReport_Focus = True
4530      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
4540      .cmdPrintReport_raised_img.Visible = False
4550      .cmdPrintReport_raised_focus_img.Visible = False
4560      .cmdPrintReport_raised_focus_dots_img.Visible = False
4570      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4580      .cmdPrintReport_raised_img_dis.Visible = False
4590    End With

EXITP:
4600    Exit Sub

ERRH:
4610    Select Case ERR.Number
        Case Else
4620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4630    End Select
4640    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

4710    With Me
4720      blnPrintReport_MouseDown = True
4730      .cmdPrintReport_sunken_focus_dots_img.Visible = True
4740      .cmdPrintReport_raised_img.Visible = False
4750      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4760      .cmdPrintReport_raised_focus_img.Visible = False
4770      .cmdPrintReport_raised_focus_dots_img.Visible = False
4780      .cmdPrintReport_raised_img_dis.Visible = False
4790    End With

EXITP:
4800    Exit Sub

ERRH:
4810    Select Case ERR.Number
        Case Else
4820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4830    End Select
4840    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

4910    With Me
4920      blnIsOpen = False
4930      strDocName = "rptSiteMap"
4940      If gblnDev_Debug Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
4950        DoCmd.OpenReport strDocName & "_01", acViewPreview, , , , THIS_NAME & "~" & strCallingForm
4960        DoCmd.OpenReport strDocName & "_02", acViewPreview, , , , THIS_NAME & "~" & strCallingForm
4970        DoCmd.OpenReport strDocName & "_03", acViewPreview, , , , THIS_NAME & "~" & strCallingForm
4980        DoCmd.OpenReport strDocName & "_04", acViewPreview, , , , THIS_NAME & "~" & strCallingForm
4990        If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
5000          blnIsOpen = True
5010          Forms(strCallingForm).Visible = False
5020        End If
5030        .Visible = False
5040        strDocName = "rptSiteMap_Journal"
5050        DoCmd.OpenReport strDocName & "_01", acViewPreview, , , , THIS_NAME & "~" & strCallingForm
5060        DoCmd.OpenReport strDocName & "_02", acViewPreview, , , , THIS_NAME & "~" & strCallingForm
5070      Else
            '##GTR_Ref: rptSiteMap
5080        DoCmd.OpenReport strDocName & "_01", acViewNormal, , , , THIS_NAME & "~" & strCallingForm
5090        DoCmd.OpenReport strDocName & "_02", acViewNormal, , , , THIS_NAME & "~" & strCallingForm
5100        DoCmd.OpenReport strDocName & "_03", acViewNormal, , , , THIS_NAME & "~" & strCallingForm
5110        DoCmd.OpenReport strDocName & "_04", acViewNormal, , , , THIS_NAME & "~" & strCallingForm
5120        strDocName = "rptSiteMap_Journal"
            '##GTR_Ref: rptSiteMap_Journal
5130        DoCmd.OpenReport strDocName & "_01", acViewNormal, , , , THIS_NAME & "~" & strCallingForm
5140        DoCmd.OpenReport strDocName & "_02", acViewNormal, , , , THIS_NAME & "~" & strCallingForm
5150      End If
5160    End With

EXITP:
5170    Exit Sub

ERRH:
5180    Me.Visible = True
5190    If blnIsOpen = True Then
5200      blnIsOpen = False
5210      Forms(strCallingForm).Visible = True
5220    End If
5230    Select Case ERR.Number
        Case Else
5240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5250    End Select
5260    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

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
          Case vbKeyTab
5370        With Me
5380          intRetVal = 0
5390          .cmdClose.SetFocus
5400        End With
5410      End Select
5420    End If

        ' ** Shift keys.
5430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5440      Select Case intRetVal
          Case vbKeyTab
5450        With Me
5460          intRetVal = 0
5470          .chkSaveSizePos.SetFocus
5480        End With
5490      End Select
5500    End If

EXITP:
5510    KeyCode = intRetVal
5520    Exit Sub

ERRH:
5530    intRetVal = 0
5540    Select Case ERR.Number
        Case Else
5550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5560    End Select
5570    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

5610    With Me
5620      If blnPrintReport_MouseDown = False Then
5630        Select Case blnPrintReport_Focus
            Case True
5640          .cmdPrintReport_raised_focus_dots_img.Visible = True
5650          .cmdPrintReport_raised_focus_img.Visible = False
5660        Case False
5670          .cmdPrintReport_raised_focus_img.Visible = True
5680          .cmdPrintReport_raised_focus_dots_img.Visible = False
5690        End Select
5700        .cmdPrintReport_raised_img.Visible = False
5710        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5720        .cmdPrintReport_sunken_focus_dots_img.Visible = False
5730        .cmdPrintReport_raised_img_dis.Visible = False
5740      End If
5750    End With

EXITP:
5760    Exit Sub

ERRH:
5770    Select Case ERR.Number
        Case Else
5780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5790    End Select
5800    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

5910    With Me
5920      .cmdPrintReport_raised_focus_dots_img.Visible = True
5930      .cmdPrintReport_raised_img.Visible = False
5940      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5950      .cmdPrintReport_raised_focus_img.Visible = False
5960      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5970      .cmdPrintReport_raised_img_dis.Visible = False
5980      blnPrintReport_MouseDown = False
5990    End With

EXITP:
6000    Exit Sub

ERRH:
6010    Select Case ERR.Number
        Case Else
6020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6030    End Select
6040    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

6110    With Me
6120      .cmdPrintReport_raised_img.Visible = True
6130      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6140      .cmdPrintReport_raised_focus_img.Visible = False
6150      .cmdPrintReport_raised_focus_dots_img.Visible = False
6160      .cmdPrintReport_sunken_focus_dots_img.Visible = False
6170      .cmdPrintReport_raised_img_dis.Visible = False
6180      blnPrintReport_Focus = False
6190    End With

EXITP:
6200    Exit Sub

ERRH:
6210    Select Case ERR.Number
        Case Else
6220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6230    End Select
6240    Resume EXITP

End Sub

Private Sub chkSaveSizePos_AfterUpdate()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkSaveSizePos_AfterUpdate"

6310    With Me
6320      Select Case .chkSaveSizePos
          Case True
6330        .chkSaveSizePos_lbl.FontBold = True
6340      Case False
6350        .chkSaveSizePos_lbl.FontBold = False
6360      End Select
6370    End With

EXITP:
6380    Exit Sub

ERRH:
6390    Select Case ERR.Number
        Case Else
6400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6410    End Select
6420    Resume EXITP

End Sub

Private Sub chkSaveSizePos_KeyDown(KeyCode As Integer, Shift As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkSaveSizePos_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6520    intShiftDown = (Shift And acShiftMask) > 0
6530    intAltDown = (Shift And acAltMask) > 0
6540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6560      Select Case intRetVal
          Case vbKeyTab
6570        With Me
6580          intRetVal = 0
6590          If .cmdPrintReport.Enabled = True Then
6600            .cmdPrintReport.SetFocus
6610          Else
6620            .cmdClose.SetFocus
6630          End If
6640        End With
6650      End Select
6660    End If

        ' ** Shift keys.
6670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6680      Select Case intRetVal
          Case vbKeyTab
6690        With Me
6700          intRetVal = 0
6710          .frmSiteMap_Sub.SetFocus
6720        End With
6730      End Select
6740    End If

EXITP:
6750    KeyCode = intRetVal
6760    Exit Sub

ERRH:
6770    intRetVal = 0
6780    Select Case ERR.Number
        Case Else
6790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6800    End Select
6810    Resume EXITP

End Sub
