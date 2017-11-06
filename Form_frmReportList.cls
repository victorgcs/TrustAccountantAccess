VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReportList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmReportList"

'VGC 04/15/2017: CHANGES!

'POSSIBLE ADD'L ARROWS:
'HOLDINGS - DETAILED 'Detail' ARROW?
'ACCT REV BY ADMIN

'HOLD CTL ON ARCHIVE TO GO DIRECTLY TO REPORTS! DO NOT PASS GO!

'#########################################################
'SOME BEEP WHEN ARROW GOES OUT, SOME WHEN ARROW APPEARS!
'#########################################################

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Go To Report:     G {cmdLocationGo}
' **   Conversion Log:   L {cmdConvLog}
' **   Reports:          R {lbxReports}
' **   Exit:             X {cmdClose}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Array: arr_varRpt().
Private lngRpts As Long, arr_varRpt As Variant
Private Const R_RLRID   As Integer = 0
'Private Const R_RLRAID  As Integer = 1
'Private Const R_RLRACID As Integer = 2
Private Const R_RLID    As Integer = 3
'Private Const R_DID     As Integer = 4
'Private Const R_RID     As Integer = 5
Private Const R_RNAM    As Integer = 6
Private Const R_RCAP    As Integer = 7
'Private Const R_RORD    As Integer = 8
'Private Const R_RLCID1  As Integer = 9
'rptlistrpt_id
'rptlistrptalt_id
'rptlistrptaltctl_id
'rptlist_id
'dbs_id
'rpt_id
'rpt_name
'rpt_caption
'rptlistrpt_order
'rptlistctl_id1
'Private Const R_TYP1    As Integer = 10
'Private Const R_LVL1    As Integer = 11
'Private Const R_ALT1    As Integer = 12
Private Const R_FID1    As Integer = 13
Private Const R_FNAM1   As Integer = 14
'Private Const R_FCAP1   As Integer = 15
Private Const R_CID1    As Integer = 16
Private Const R_CNAM1   As Integer = 17
Private Const R_CCAP1   As Integer = 18
'Private Const R_INV1    As Integer = 19
'rptlistctltype_type1
'rptlistctl_level1
'rptlistctl_alt1
'frm_id1
'frm_name1
'frm_caption1
'ctl_id1
'ctl_name1
'ctl_caption1
'ctl_invoke1
'Private Const R_FOC1    As Integer = 20
'Private Const R_RLCID2  As Integer = 21
Private Const R_TYP2    As Integer = 22
'Private Const R_LVL2    As Integer = 23
'Private Const R_ALT2    As Integer = 24
Private Const R_FID2    As Integer = 25
Private Const R_FNAM2   As Integer = 26
'Private Const R_FCAP2   As Integer = 27
Private Const R_CID2    As Integer = 28
'Private Const R_CNAM2   As Integer = 29
'ctl_focus1
'rptlistctl_id2
'rptlistctltype_type2
'rptlistctl_level2
'rptlistctl_alt2
'frm_id2
'frm_name2
'frm_caption2
'ctl_id2
'ctl_name2
Private Const R_CCAP2   As Integer = 30
'Private Const R_INV2    As Integer = 31
'Private Const R_FOC2    As Integer = 32
'Private Const R_RLCID3  As Integer = 33
Private Const R_TYP3    As Integer = 34
'Private Const R_LVL3    As Integer = 35
'Private Const R_ALT3    As Integer = 36
Private Const R_FID3    As Integer = 37
Private Const R_FNAM3   As Integer = 38
'Private Const R_FCAP3   As Integer = 39
'ctl_caption2
'ctl_invoke2
'ctl_focus2
'rptlistctl_id3
'rptlistctltype_type3
'rptlistctl_level3
'rptlistctl_alt3
'frm_id3
'frm_name3
'frm_caption3
Private Const R_CID3    As Integer = 40
Private Const R_CNAM3   As Integer = 41
Private Const R_CCAP3   As Integer = 42
'Private Const R_INV3    As Integer = 43
'Private Const R_FOC3    As Integer = 44
'Private Const R_RLCID4  As Integer = 45
Private Const R_TYP4    As Integer = 46
'Private Const R_LVL4    As Integer = 47
'Private Const R_ALT4    As Integer = 48
Private Const R_FID4    As Integer = 49
'ctl_id3
'ctl_name3
'ctl_caption3
'ctl_invoke3
'ctl_focus3
'rptlistctl_id4
'rptlistctltype_type4
'rptlistctl_level4
'rptlistctl_alt4
'frm_id4
'Private Const R_FNAM4   As Integer = 50
'Private Const R_FCAP4   As Integer = 51
Private Const R_CID4    As Integer = 52
'Private Const R_CNAM4   As Integer = 53
Private Const R_CCAP4   As Integer = 54
'Private Const R_INV4    As Integer = 55
'Private Const R_FOC4    As Integer = 56
'Private Const R_RLCID5  As Integer = 57
Private Const R_TYP5    As Integer = 58
'Private Const R_LVL5    As Integer = 59
'frm_name4
'frm_caption4
'ctl_id4
'ctl_name4
'ctl_caption4
'ctl_invoke4
'ctl_focus4
'rptlistctl_id5
'rptlistctltype_type5
'rptlistctl_level5
'Private Const R_ALT5    As Integer = 60
Private Const R_FID5    As Integer = 61
'Private Const R_FNAM5   As Integer = 62
'Private Const R_FCAP5   As Integer = 63
Private Const R_CID5    As Integer = 64
'Private Const R_CNAM5   As Integer = 65
'Private Const R_CCAP5   As Integer = 66
'Private Const R_INV5    As Integer = 67
'Private Const R_FOC5    As Integer = 68
'Private Const R_RLCID6  As Integer = 69
'rptlistctl_alt5
'frm_id5
'frm_name5
'frm_caption5
'ctl_id5
'ctl_name5
'ctl_caption5
'ctl_invoke5
'ctl_focus5
'rptlistctl_id6
'Private Const R_TYP6    As Integer = 70
'Private Const R_LVL6    As Integer = 71
'Private Const R_ALT6    As Integer = 72
'Private Const R_FID6    As Integer = 73
'Private Const R_FNAM6   As Integer = 74
'Private Const R_FCAP6   As Integer = 75
'Private Const R_CID6    As Integer = 76
'Private Const R_CNAM6   As Integer = 77
'Private Const R_CCAP6   As Integer = 78
'Private Const R_INV6    As Integer = 79
'rptlistctltype_type6
'rptlistctl_level6
'rptlistctl_alt6
'frm_id6
'frm_name6
'frm_caption6
'ctl_id6
'ctl_name6
'ctl_caption6
'ctl_invoke6
'Private Const R_FOC6    As Integer = 80
Private Const R_CNT     As Integer = 81
'Private Const R_SID     As Integer = 82
'ctl_focus6
'cnt
'rptlistsort_id
'QRY: qryReport_List_89c  FLDS: 83
'DONE!  Tbl_Fld_List()

' ** List box column constants: lbxReports.
Private Const LBX_RPT_RLRID As Integer = 0  ' ** rptlistrpt_id
Private Const LBX_RPT_RORD  As Integer = 1  ' ** rptlistrpt_order
Private Const LBX_RPT_RCAP  As Integer = 2  ' ** rpt_caption
Private Const LBX_RPT_RID   As Integer = 3  ' ** rpt_id
Private Const LBX_RPT_RNAM  As Integer = 4  ' ** rpt_name

' ** Report List Control Type constants:
Private taReport As Long
Private taMenu1 As Long
Private taMenu2 As Long
Private taMenu3 As Long
Private taVia1 As Long
Private taVia2 As Long
Private taVia3 As Long
Private taVia4 As Long

' ** Array: arr_varCtl().
Private lngCtls As Long, arr_varCtl() As Variant
Private Const C_ELEMS As Integer = 1  ' ** Array's first-element UBound().
Private Const C_CNAM As Integer = 0
Private Const C_TOP  As Integer = 1

Private blnNoGo As Boolean
Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long, blnConvLog As Boolean
Private lngOpt2_Top As Long, lngOpt2lbl_Top As Long, lngOptOffset_Top As Long
Private strLogPathFile As String, lngTpp As Long
Private lngRptListRptID As Long, lngRptListRpts As Long
Private THAT_PROC As String, That_Erl As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean, blnShowFlds As Boolean

110     blnRetVal = True

120     With Me
130       If IsNull(.OpenArgs) = False Then
140         If Trim(.OpenArgs) <> vbNullString Then

150           strCallingForm = Trim(.OpenArgs)

160           CLR_DISABLED_FG = CLR_DKGRY
170           CLR_DISABLED_BG = CLR_LTTEAL

180           strLogPathFile = vbNullString
190           gblnGoToReport = False: blnGTR_Emblem = False: lngGTR_Stat = 0&: blnConvLog = False: blnNoGo = False

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
200           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

210           ReDim garr_varGoToReport(GTR_ELEMS)

220           blnRetVal = RptArray_Load  ' ** Function: Below.

230           Select Case blnRetVal
              Case True

                ' ** This will set the gblnPricingAllowed variable.
240             blnRetVal = Security_PricingChk  ' ** Module Function: modSecurityFunctions.

250             blnShowFlds = False
260             If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
                  'blnShowFlds = True
                  '.rptlist_id.Visible = True
270             End If

280             Select Case blnShowFlds
                Case True
290               .Multi.Visible = True
300               .frm1.Visible = True
310               .frm2.Visible = True
320               .frm3.Visible = True
330               .frm4.Visible = True
340               .ctl1.Visible = True
350               .ctl2.Visible = True
360               .ctl3.Visible = True
370               .ctl4.Visible = True
380             Case False
390               .Multi.Visible = False
400               .frm1.Visible = False
410               .frm2.Visible = False
420               .frm3.Visible = False
430               .frm4.Visible = False
440               .ctl1.Visible = False
450               .ctl2.Visible = False
460               .ctl3.Visible = False
470               .ctl4.Visible = False
480             End Select

490             .cmdConvLog.Enabled = False
500             .cmdConvLog.Visible = False

510             lngOpt2_Top = .opgGTRMulti_opt2.Top
520             lngOpt2lbl_Top = .opgGTRMulti_opt2_lbl.Top
530             lngOptOffset_Top = .opgGTRMulti_opt2.Top - .opgGTRMulti_opt1.Top

540           Case False
550             Cancel = -1
560           End Select
570         Else
580           Cancel = -1
590         End If
600       Else
610         Cancel = -1
620       End If
630     End With

640     If Cancel = -1 Then
650       Select Case blnRetVal
          Case True
660         Beep
670         MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
680       Case False
690         Beep
700         MsgBox "There is a problem with the Report List.", vbCritical + vbOKOnly, "Report List Error"
710       End Select

720     End If

EXITP:
730     Exit Sub

ERRH:
740     DoCmd.Hourglass False
750     Select Case ERR.Number
        Case Else
760       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
770     End Select
780     Resume EXITP

End Sub

Private Sub Form_Load()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim ctl As Access.Control
        Dim lngCnt As Long
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

810     With Me

820       blnRetVal = RptArray_Load  ' ** Function: Below.

830       If blnRetVal = True Then

840         lngCtls = 0&
850         ReDim arr_varCtl(C_ELEMS, 0)

860         For Each ctl In .Controls
870           With ctl
880             If Left(.Name, 13) = "cmbViaButton1" Or Left(.Name, 11) = "cmbViaForm1" Or _
                    .Name = "Location_arrow4a" Or .Name = "Location_arrow4c" Then
890               lngCtls = lngCtls + 1&
900               lngE = lngCtls - 1&
910               ReDim Preserve arr_varCtl(C_ELEMS, lngE)
920               arr_varCtl(C_CNAM, lngE) = .Name
930               arr_varCtl(C_TOP, lngE) = .Top
940             End If
950           End With
960         Next

970         .lbxReports.SetFocus
980         .lbxReports = .lbxReports.Column(0, 0)
990         lbxReports_AfterUpdate  ' ** Procedure: Below.
1000        .ReportsCnt = .lbxReports.ListCount

1010        Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
1020        DoEvents

1030        If IsNull(.lbxReports) = True Then
1040          .lbxReports = .lbxReports.Column(0, 0)
1050        End If

1060        lbxReports_AfterUpdate  ' ** Procedure: Below.
1070        opgGTRMulti_AfterUpdate  ' ** Procedure: Below.

            ' ** Position selection at top of 16-line window instead of bottom.
1080        lngRptListRptID = .lbxReports.Column(LBX_RPT_RLRID)
1090        lngX = .lbxReports.Column(LBX_RPT_RORD)
1100        lngCnt = .lbxReports.ListCount
1110        If lngX < 16& Then
1120          .lbxReports = .lbxReports.Column(LBX_RPT_RLRID, 0)
1130          .lbxReports = lngRptListRptID
1140        ElseIf lngCnt - lngX < 16& Then
1150          .lbxReports = .lbxReports.Column(LBX_RPT_RLRID, (lngCnt - 1))
1160          .lbxReports = lngRptListRptID
1170        Else
1180          .lbxReports = .lbxReports.Column(LBX_RPT_RLRID, (lngX + 15))
1190          .lbxReports = lngRptListRptID
1200        End If

1210      End If  ' ** blnRetVal.

1220      DoCmd.Hourglass False

1230    End With

EXITP:
1240    Set ctl = Nothing
1250    Exit Sub

ERRH:
1260    DoCmd.Hourglass False
1270    Select Case ERR.Number
        Case Else
1280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1290    End Select
1300    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1420    intShiftDown = (Shift And acShiftMask) > 0
1430    intAltDown = (Shift And acAltMask) > 0
1440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyEscape
1470        intRetVal = 0
1480        cmdClose_Click  ' ** Procedure: Below.
1490      End Select
1500    End If

        ' ** Alt keys.
1510    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1520      Select Case intRetVal
          Case vbKeyX
1530        intRetVal = 0
1540        cmdClose_Click  ' ** Procedure: Below.
1550      End Select
1560    End If

        ' ** Ctrl-Shift keys.
1570    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1580      Select Case intRetVal
          Case vbKeyF
1590        With Me
1600          intRetVal = 0
1610          .FocusHolder.SetFocus
1620        End With
1630      End Select
1640    End If

EXITP:
1650    KeyCode = intRetVal
1660    Exit Sub

ERRH:
1670    intRetVal = 0
1680    Select Case ERR.Number
        Case Else
1690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1700    End Select
1710    Resume EXITP

End Sub

Private Sub Form_Timer()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Static lngTmp01 As Long, lngTmp02 As Long

1810    With Me
1820      .TimerInterval = 0&
1830      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
1840        lngGTR_Stat = lngGTR_Stat + 1&
1850        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1860        lngTmp01 = lngTmp01 + 1&
1870        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1880        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1890        If lngTmp01 > 1& Then
1900          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1910        End If
1920        If lngTmp01 = 24& Then
1930          lngTmp01 = 0&
1940          lngTmp02 = lngTmp02 + 1&
1950          If lngTmp02 < 5& Then
1960            .TimerInterval = 50&
1970          Else
1980            blnGTR_Emblem = False
1990            GTREmblem_Off  ' ** Procedure: Below.
2000          End If
2010        Else
2020          .TimerInterval = 50&
2030        End If
2040      End If  ' ** blnGTR_Emblem.
2050      If gblnGoToReport = True Then
2060        If lngGTR_Stat < 25& Then
              ' ** Just let the emblem go.
2070        ElseIf lngGTR_Stat = 25& Then
2080          .TimerInterval = 50&
2090          cmdLocationGo_Click  ' ** Procedure: Below.
2100        ElseIf lngGTR_Stat > 25& And lngGTR_Stat < 72& Then  '36& Then
              ' ** Let it go some more.
2110        Else
2120          .TimerInterval = 50&
2130          Select Case blnGoingToReport
              Case True
2140            cmdConvLog_Click  ' ** Procedure: Below.
2150          Case False
2160            Select Case blnConvLog
                Case True
                  ' ** This must be some kind of stop-gap measure, because they both get set at the same time below!
2170              blnConvLog = False
2180              lngGTR_Stat = 0&
2190              lngTmp01 = 0&: lngTmp02 = 0&
2200            Case False
2210              cmdClose_Click  ' ** Procedure: Below.
2220            End Select
2230          End Select
2240        End If
2250      End If
2260    End With

EXITP:
2270    Exit Sub

ERRH:
2280    Select Case ERR.Number
        Case Else
2290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2300    End Select
2310    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2410    If DataErr <> 0 Then
2420      Me.TimerInterval = 0&
2430      Debug.Print "'THAT_PROC: " & THAT_PROC & "  That_Erl: " & That_Erl
2440      Select Case DataErr
          Case 2115       ' ** The macro or function set to the BeforeUpdate or ValidationRule property for
            ' ** Ignore.  ' ** this field is preventing Microsoft Access from saving the data in the field.
2450        Response = acDataErrContinue
2460      Case Else
2470        zErrorHandler THIS_NAME, THIS_PROC, DataErr, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2480      End Select
2490    End If

        ' ** AcDataError enumeration:
        ' **   0  acDataErrContinue  Ignore the error and continue without displaying the default Microsoft Access
        ' **                         error message. A custom error message may be displayed in place of the default
        ' **                         error message.
        ' **   1  acDataErrDisplay   Display the default Microsoft Access error message. (Default)
        ' **   2  acDataErrAdded     Don't display the default Microsoft Access error message. The entry may be
        ' **                         added to the combo box list in the NotInList event procedure. After the entry
        ' **                         is added, Microsoft Access updates the list by requerying the combo box.
        ' **                         Microsoft Access then rechecks the string against the combo box list, and saves
        ' **                         the value in the NewData argument in the field the combo box is bound to. If
        ' **                         the string is not in the list, then Microsoft Access displays an error message.

EXITP:
2500    Exit Sub

ERRH:
2510    Select Case ERR.Number
        Case Else
2520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2530    End Select
2540    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2610    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** 1  rptAccountProfile
        ' ** 2  rptTransaction_Hidden
        ' ** 3  rptChecks_Blank
        ' ** 4  rptListOfTaxLots
        ' ** 5  rptMap_Dividend
        ' ** 6  rptMap_Interest
        ' ** 7  rptMap_Received
        ' ** 8  rptMap_Split
        ' ** 9  rptMap_Reinvest_Div
        ' ** 10 rptMap_Reinvest_Int
        ' ** 11 rptMap_Reinvest_Rec

2620    If IsLoaded("frmVersion_Log", acForm) = True Then
2630      DoCmd.Close acForm, "frmVersion_Log"
2640    End If

2650    gstrFormQuerySpec = vbNullString
2660    gblnSetFocus = True
2670    gstrReturningForm = THIS_NAME
2680    Select Case strCallingForm
        Case "frmMenu_Report"
2690      If gblnGoToReport = True Then
2700        DoCmd.Hourglass True  ' ** Make sure it's still running.
2710        DoEvents
2720      End If
2730      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2740      If gblnGoToReport = True Then
            ' ** Since frmMenu_Report wasn't closed, reinitialize the timer.
2750        Forms(strCallingForm).TimerInterval = 100&  ' ** Currently set to 100.
2760      End If
2770    End Select

EXITP:
2780    Exit Sub

ERRH:
2790    DoCmd.Hourglass False
2800    Select Case ERR.Number
        Case Else
2810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2820    End Select
2830    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2910    DoCmd.Close acForm, THIS_NAME

EXITP:
2920    Exit Sub

ERRH:
2930    Select Case ERR.Number
        Case Else
2940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2950    End Select
2960    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
3070        With Me
3080          intRetVal = 0
3090          .lbxReports.SetFocus
3100        End With
3110      End Select
3120    End If

        ' ** Shift keys.
3130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3140      Select Case intRetVal
          Case vbKeyTab
3150        With Me
3160          intRetVal = 0
3170          If .cmdConvLog.Visible = True And .cmdConvLog.Enabled = True Then
3180            .cmdConvLog.SetFocus
3190          ElseIf .cmdLocationGo.Enabled = True Then
3200            .cmdLocationGo.SetFocus
3210          Else
3220            .lbxReports.SetFocus
3230          End If
3240        End With
3250      End Select
3260    End If

EXITP:
3270    KeyCode = intRetVal
3280    Exit Sub

ERRH:
3290    intRetVal = 0
3300    Select Case ERR.Number
        Case Else
3310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3320    End Select
3330    Resume EXITP

End Sub

Private Sub lbxReports_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxReports_DblClick"

3410    THAT_PROC = THIS_PROC
3420    That_Erl = 3420
3430    cmdLocationGo_Click  ' ** Procedure: Below.

EXITP:
3440    Exit Sub

ERRH:
3450    Select Case ERR.Number
        Case Else
3460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3470    End Select
3480    Resume EXITP

End Sub

Private Sub lbxReports_AfterUpdate()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxReports_AfterUpdate"

        Dim blnLevel2 As Boolean, blnLevel3 As Boolean, blnVia1 As Boolean, blnVia2 As Boolean
        Dim varAssets As Variant
        Dim blnContinue As Boolean, blnSpecialHandling As Boolean
        Dim varTmp00 As Variant, strTmp01 As String, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long
        Dim lngX As Long, lngY As Long, lngE As Long

3510    blnContinue = True
3520    blnLevel2 = False: blnLevel3 = False: blnVia1 = False: blnVia2 = False
3530    blnGoingToReport = False: blnNoGo = False

3540    With Me

          'CHECK SPECIFIED REPORT NAMES AND CAPTIONS!
          'CHECK AVAILABILITY OF SPECIAL MAINTENANCE!!!
          'CHECK PRESENCE OF FOREIGN CURRENCIES!

3550      .cmbButton1.Locked = True
3560      .cmbButton1.ForeColor = CLR_BLK
3570      .cmbButton1.BorderColor = CLR_LTBLU2
3580      .cmbButton1_lbl_dim_hi.Visible = False
3590      .cmbButton2.Locked = True
3600      .cmbButton2.ForeColor = CLR_BLK
3610      .cmbButton2.BorderColor = CLR_LTBLU2
3620      .cmbButton2_lbl_dim_hi.Visible = False
3630      .cmbViaButton1.Locked = True
3640      .cmbViaButton1.BorderColor = CLR_LTBLU2
3650      .cmbViaButton1_lbl_dim_hi.Visible = False
3660      .cmdLocationGo.Enabled = True

3670      .cmbMenu2.Visible = False
3680      .cmbButton2.Visible = False
3690      .cmbMenu3.Visible = False
3700      .cmbButton3.Visible = False
3710      .cmbViaForm1.Visible = False
3720      .cmbViaButton1.Visible = False
3730      .cmbViaForm2.Visible = False
3740      .cmbViaButton2.Visible = False
3750      .cmbMenu1.ForeColor = CLR_BLK
3760      .cmbMenu1.BackColor = CLR_WHT
3770      .cmbMenu2.ForeColor = CLR_BLK
3780      .cmbMenu2.BackColor = CLR_WHT
3790      .cmbMenu3.ForeColor = CLR_BLK
3800      .cmbMenu3.BackColor = CLR_WHT
3810      .cmbViaForm1.ForeColor = CLR_BLK

3820      .cmbViaForm1.BackColor = CLR_WHT
3830      .cmbViaForm2.ForeColor = CLR_BLK
3840      .cmbViaForm2.BackColor = CLR_WHT
3850      .cmbButton1.BackStyle = acBackStyleNormal
3860      .cmbButton2.BackStyle = acBackStyleNormal
3870      .cmbButton3.BackStyle = acBackStyleNormal
3880      .cmbViaButton1.BackStyle = acBackStyleNormal
3890      .cmbViaButton2.BackStyle = acBackStyleNormal
3900      .cmbViaButton1.ForeColor = CLR_BLK
3910      .cmbViaButton1.BackColor = CLR_WHT
3920      .cmbViaButton2.ForeColor = CLR_BLK
3930      .cmbViaButton2.BackColor = CLR_WHT
3940      .cmbViaButton1_lbl.Caption = "Button"
3950      .cmbViaButton2_lbl.Caption = "Button"

3960      .cmbButton1_cmd.Enabled = False
3970      .cmbButton1_cmd.Visible = False
3980      .cmbButton2_cmd.Enabled = False
3990      .cmbButton2_cmd.Visible = False
4000      .cmbButton3_cmd.Enabled = False
4010      .cmbButton3_cmd.Visible = False
4020      .cmbViaButton1_cmd.Enabled = False
4030      .cmbViaButton1_cmd.Visible = False

4040      .cmbMenu1_box1.Visible = False
4050      .cmbMenu1_box2.Visible = False
4060      .cmbMenu1_box3.Visible = False
4070      .cmbMenu2_box1.Visible = False
4080      .cmbMenu2_box2.Visible = False
4090      .cmbMenu3_box1.Visible = False
4100      .cmbMenu3_box2.Visible = False

4110      .opgGTRMulti.Visible = False
4120      .opgGTRMulti_opt1.Enabled = True
4130      .opgGTRMulti_opt3.Visible = True  ' ** Though Option Group as a whole isn't visible.
4140      .opgGTRMulti_opt4.Visible = True
4150      .opgGTRMulti_opt1_cmd.Visible = False
4160      .opgGTRMulti_opt1_cmd.Enabled = True
4170      .opgGTRMulti_opt2_cmd.Visible = False
4180      .opgGTRMulti_opt3_cmd.Visible = False
4190      .opgGTRMulti_opt4_cmd.Visible = False

4200      .opgGTRMulti_opt1.Top = (lngOpt2_Top - lngOptOffset_Top)
4210      .opgGTRMulti_opt1_lbl.Top = (lngOpt2lbl_Top - lngOptOffset_Top)
4220      .opgGTRMulti_opt1_cmd.Top = .cmbButton1.Top
4230      .opgGTRMulti_opt1_cmd.Left = .cmbButton1.Left
4240      .opgGTRMulti_opt2.Top = lngOpt2_Top
4250      .opgGTRMulti_opt2_lbl.Top = lngOpt2lbl_Top
4260      .opgGTRMulti_opt2_cmd.Top = .cmbButton6.Top
4270      .opgGTRMulti_opt2_cmd.Left = .cmbButton6.Left

4280      .cmbButton6.Visible = False
4290      .cmbButton7.Visible = False
4300      .cmbButton8.Visible = False
4310      For lngX = 1& To 15&
4320        .Controls("cmbButton6_hline" & Right("00" & CStr(lngX), 2)).Visible = False
4330        .Controls("cmbButton7_hline" & Right("00" & CStr(lngX), 2)).Visible = False
4340        .Controls("cmbButton8_hline" & Right("00" & CStr(lngX), 2)).Visible = False
4350      Next
4360      .Location_arrow6b.Visible = False
4370      .Location_arrow6d.Visible = False
4380      .cmbMenu6.Visible = False
4390      .cmbMenu6_box1.Visible = False
4400      .cmbMenu6_box2.Visible = False
4410      .Location_arrow7a.Visible = False
4420      .Location_arrow7b.Visible = False
4430      .Location_arrow7c.Visible = False
4440      .Location_arrow7d.Visible = False
4450      .Location_arrow8a.Visible = False
4460      .Location_arrow8d.Visible = False
4470      .Location_arrow8e.Visible = False
4480      For lngX = 2& To 3&
4490        For lngY = 1& To 15&
4500          strTmp01 = "cmbButton" & CStr(lngX) & "_hline"
4510          strTmp01 = strTmp01 & Right("00" & CStr(lngY), 2)
4520          .Controls(strTmp01).Visible = False
4530        Next
4540      Next
4550      For lngX = 1& To 2&
4560        For lngY = 1& To 15&
4570          strTmp01 = "cmbViaButton" & CStr(lngX) & "_hline"
4580          strTmp01 = strTmp01 & Right("00" & CStr(lngY), 2)
4590          .Controls(strTmp01).Visible = False
4600        Next
4610      Next
4620      If .Controls(arr_varCtl(C_CNAM, 0)).Top <> arr_varCtl(C_TOP, 0) Then
4630        For lngX = 0& To (lngCtls - 1&)
4640          .Controls(arr_varCtl(C_CNAM, lngX)).Top = arr_varCtl(C_TOP, lngX)
4650        Next
4660      End If

4670      lngRptListRptID = .lbxReports
4680      blnContinue = False: lngE = -1&
4690      For lngX = 0& To (lngRpts - 1&)
4700        If .lbxReports.Column(LBX_RPT_RNAM) = "rptArchivedTransactions" Then
4710          If arr_varRpt(R_RLRID, lngX) = lngRptListRptID And arr_varRpt(R_CNAM1, lngX) = "cmdArchiveTransactions" Then
4720            blnContinue = True
4730            lngE = lngX
4740            lngRptListRpts = arr_varRpt(R_CNT, lngX)
4750            Exit For
4760          End If
4770        ElseIf .lbxReports.Column(LBX_RPT_RNAM) = "rptVersion_Convert" Then
4780          If arr_varRpt(R_RLRID, lngX) = lngRptListRptID And arr_varRpt(R_CNAM1, lngX) = "chkConversionCheck" Then
                ' ** This is the one we want.
4790            blnContinue = True
4800            lngE = lngX
4810            lngRptListRpts = arr_varRpt(R_CNT, lngX)
4820            Exit For
4830          Else
                ' ** Keep looking for the other one.
4840          End If
4850        Else
4860          If arr_varRpt(R_RLRID, lngX) = lngRptListRptID Then
4870            blnContinue = True
4880            lngE = lngX
4890            lngRptListRpts = arr_varRpt(R_CNT, lngX)
4900            Exit For
4910          End If
4920        End If
4930      Next

4940      If blnContinue = True Then

4950        blnSpecialHandling = False
4960        If lngRptListRpts > 1 Then

4970          If Left(arr_varRpt(R_RNAM, lngE), 17) = "rptAccountProfile" Then
4980            blnSpecialHandling = True
4990            If arr_varRpt(R_FNAM1, lngE) = "frmMenu_Report" Then
5000              If arr_varRpt(R_CNAM1, lngE) = "cmdAccountProfile" Then
5010                For lngX = 0& To (lngRpts - 1&)
5020                  If lngX <> lngE Then
5030                    If Left(arr_varRpt(R_RNAM, lngX), 17) = "rptAccountProfile" And arr_varRpt(R_FNAM1, lngX) = "frmMenu_Report" And _
                            arr_varRpt(R_CNAM1, lngX) = "cmdAccountProfile" Then
                          ' ** Switch the primary to Account Profile.
5040                      lngE = lngX
5050                      Exit For
5060                    End If
5070                  End If
5080                Next
5090              End If
5100            End If
5110          ElseIf arr_varRpt(R_RNAM, lngE) = "rptChecks_Blank" Then
5120            blnSpecialHandling = True
5130            For lngX = 0& To (lngRpts - 1&)
5140              If arr_varRpt(R_RNAM, lngX) = "rptChecks_Blank" Then
5150                If IsNull(arr_varRpt(R_FNAM2, lngX)) = False And IsNull(arr_varRpt(R_FNAM2, lngE)) = False Then
5160                  lngTmp02 = lngE
5170                  lngE = lngX
5180                Else
5190                  lngTmp02 = lngX
5200                End If
5210                Exit For
5220              End If
5230            Next
5240          ElseIf arr_varRpt(R_RNAM, lngE) = "rptVersion_Convert" Then
5250            blnSpecialHandling = True
                'For lngX = 0& To (lngRpts - 1&)
                '  If arr_varRpt(R_RNAM, lngX) = "rptVersion_Convert" Then
                '    If IsNull(arr_varRpt(R_FNAM2, lngX)) = True And IsNull(arr_varRpt(R_FNAM2, lngE)) = False Then
                '      lngTmp02 = lngE
                '      lngE = lngX
                '    Else
                '      lngTmp02 = lngX
                '    End If
                '    Exit For
                '  End If
                'Next
5260          ElseIf arr_varRpt(R_RNAM, lngE) = "rptTransaction_Hidden" Then
5270            blnSpecialHandling = True
5280            For lngX = 0& To (lngRpts - 1&)
5290              If arr_varRpt(R_RNAM, lngX) = "rptTransaction_Hidden" Then
5300                If IsNull(arr_varRpt(R_CNAM3, lngX)) = False Then
                      ' ** Wait for the next one.
5310                Else
5320                  lngTmp02 = lngE
5330                  lngE = lngX
5340                  Exit For
5350                End If
5360              End If
5370            Next
5380          ElseIf arr_varRpt(R_RNAM, lngE) = "rptListOfTaxLots" Then
5390            blnSpecialHandling = True
5400            For lngX = 0& To (lngRpts - 1&)
5410              If arr_varRpt(R_RNAM, lngX) = "rptListOfTaxLots" Then
5420                If arr_varRpt(R_FNAM2, lngX) <> arr_varRpt(R_FNAM2, lngE) Then
5430                  lngTmp02 = lngX
5440                  Exit For
5450                End If
5460              End If
5470            Next
5480          ElseIf arr_varRpt(R_RNAM, lngE) = "rptMap_Dividend" Then
5490            blnSpecialHandling = True
5500            For lngX = 0& To (lngRpts - 1&)
5510              If arr_varRpt(R_RNAM, lngX) = "rptMap_Dividend" Then
5520                If arr_varRpt(R_FNAM2, lngX) <> arr_varRpt(R_FNAM2, lngE) Then
5530                  If arr_varRpt(R_FNAM2, lngE) = "frmJournal_Columns" Then
                        ' ** Switch the primary to Journal Classic.
5540                    lngTmp02 = lngE
5550                    lngE = lngX
5560                  Else
5570                    lngTmp02 = lngX
5580                  End If
5590                  Exit For
5600                End If
5610              End If
5620            Next
5630          ElseIf arr_varRpt(R_RNAM, lngE) = "rptMap_Interest" Then
5640            blnSpecialHandling = True
5650            For lngX = 0& To (lngRpts - 1&)
5660              If arr_varRpt(R_RNAM, lngX) = "rptMap_Interest" Then
5670                If arr_varRpt(R_FNAM2, lngX) <> arr_varRpt(R_FNAM2, lngE) Then
5680                  If arr_varRpt(R_FNAM2, lngE) = "frmJournal_Columns" Then
                        ' ** Switch the primary to Journal Classic.
5690                    lngTmp02 = lngE
5700                    lngE = lngX
5710                  Else
5720                    lngTmp02 = lngX
5730                  End If
5740                  Exit For
5750                End If
5760              End If
5770            Next
5780          ElseIf arr_varRpt(R_RNAM, lngE) = "rptMap_Received" Then
5790            blnSpecialHandling = True
5800            For lngX = 0& To (lngRpts - 1&)
5810              If arr_varRpt(R_RNAM, lngX) = "rptMap_Received" Then
5820                If arr_varRpt(R_FNAM2, lngX) <> arr_varRpt(R_FNAM2, lngE) Then
5830                  If arr_varRpt(R_FNAM2, lngE) = "frmJournal_Columns" Then
                        ' ** Switch the primary to Journal Classic.
5840                    lngTmp02 = lngE
5850                    lngE = lngX
5860                  Else
5870                    lngTmp02 = lngX
5880                  End If
5890                  Exit For
5900                End If
5910              End If
5920            Next
5930          ElseIf arr_varRpt(R_RNAM, lngE) = "rptMap_Split" Then
5940            blnSpecialHandling = True
5950            For lngX = 0& To (lngRpts - 1&)
5960              If arr_varRpt(R_RNAM, lngX) = "rptMap_Split" Then
5970                If arr_varRpt(R_FNAM2, lngX) <> arr_varRpt(R_FNAM2, lngE) Then
5980                  If arr_varRpt(R_FNAM2, lngE) = "frmJournal_Columns" Then
                        ' ** Switch the primary to Journal Classic.
5990                    lngTmp02 = lngE
6000                    lngE = lngX
6010                  Else
6020                    lngTmp02 = lngX
6030                  End If
6040                  Exit For
6050                End If
6060              End If
6070            Next
6080          ElseIf arr_varRpt(R_RNAM, lngE) = "rptMap_Reinvest_Div" Then
6090            blnSpecialHandling = True  ' ** There are 4!
6100            For lngX = 0& To (lngRpts - 1&)
6110              If arr_varRpt(R_RNAM, lngX) = "rptMap_Reinvest_Div" Then
6120                If arr_varRpt(R_FNAM2, lngX) <> arr_varRpt(R_FNAM2, lngE) Then
6130                  If arr_varRpt(R_FNAM2, lngE) = "frmJournal_Columns" Then
                        ' ** Switch the primary to Journal Classic.
6140                    lngTmp02 = lngE
6150                    lngE = lngX
6160                  Else
6170                    lngTmp02 = lngX
6180                  End If
6190                  Exit For
6200                End If
6210              End If
6220            Next
6230          ElseIf arr_varRpt(R_RNAM, lngE) = "rptMap_Reinvest_Int" Then
6240            blnSpecialHandling = True
6250            For lngX = 0& To (lngRpts - 1&)
6260              If arr_varRpt(R_RNAM, lngX) = "rptMap_Reinvest_Int" Then
6270                If arr_varRpt(R_FNAM2, lngX) <> arr_varRpt(R_FNAM2, lngE) Then
6280                  If arr_varRpt(R_FNAM2, lngE) = "frmJournal_Columns" Then
                        ' ** Switch the primary to Journal Classic.
6290                    lngTmp02 = lngE
6300                    lngE = lngX
6310                  Else
6320                    lngTmp02 = lngX
6330                  End If
6340                  Exit For
6350                End If
6360              End If
6370            Next
6380          ElseIf arr_varRpt(R_RNAM, lngE) = "rptMap_Reinvest_Rec" Then
6390            blnSpecialHandling = True
6400            For lngX = 0& To (lngRpts - 1&)
6410              If arr_varRpt(R_RNAM, lngX) = "rptMap_Reinvest_Rec" Then
6420                If arr_varRpt(R_FNAM2, lngX) <> arr_varRpt(R_FNAM2, lngE) Then
6430                  If arr_varRpt(R_FNAM2, lngE) = "frmJournal_Columns" Then
                        ' ** Switch the primary to Journal Classic.
6440                    lngTmp02 = lngE
6450                    lngE = lngX
6460                  Else
6470                    lngTmp02 = lngX
6480                  End If
6490                  Exit For
6500                End If
6510              End If
6520            Next
6530          ElseIf arr_varRpt(R_RNAM, lngE) = "rptMap_Misc_LTCL" Then  'ARE THESE NECESSARY?
6540            blnSpecialHandling = True
6550          ElseIf arr_varRpt(R_RNAM, lngE) = "rptMap_Misc_STCG" Then
6560            blnSpecialHandling = True
6570          ElseIf arr_varRpt(R_RNAM, lngE) = "rptMap_Misc_STCL" Then
6580            blnSpecialHandling = True
6590          End If

6600          .Multi = lngRptListRpts

6610          lngTmp02 = -1&: lngTmp03 = -1&: lngTmp04 = -1&

6620          For lngX = 1& To lngRptListRpts
6630            Select Case lngX
                Case 1&
6640              .Controls("frm" & CStr(lngX)) = arr_varRpt(R_FNAM1, lngE)
6650              .Controls("ctl" & CStr(lngX)) = arr_varRpt(R_CNAM1, lngE)
6660            Case 2&
6670              For lngY = 0& To (lngRpts - 1)
6680                If arr_varRpt(R_RLRID, lngY) = arr_varRpt(R_RLRID, lngE) And lngY <> lngE Then
6690                  .Controls("frm" & CStr(lngX)) = arr_varRpt(R_FNAM1, lngY)
6700                  .Controls("ctl" & CStr(lngX)) = arr_varRpt(R_CNAM1, lngY)
6710                  lngTmp02 = lngY
6720                  Exit For
6730                End If
6740              Next
6750            Case 3&
6760              For lngY = 0& To (lngRpts - 1)
6770                If arr_varRpt(R_RLRID, lngY) = arr_varRpt(R_RLRID, lngE) And lngY <> lngE And lngY <> lngTmp02 Then
6780                  .Controls("frm" & CStr(lngX)) = arr_varRpt(R_FNAM1, lngY)
6790                  .Controls("ctl" & CStr(lngX)) = arr_varRpt(R_CNAM1, lngY)
6800                  lngTmp03 = lngY
6810                  Exit For
6820                End If
6830              Next
6840            Case 4&
6850              For lngY = 0& To (lngRpts - 1)
6860                If arr_varRpt(R_RLRID, lngY) = arr_varRpt(R_RLRID, lngE) And lngY <> lngE And lngY <> lngTmp02 And lngY <> lngTmp03 Then
6870                  .Controls("frm" & CStr(lngX)) = arr_varRpt(R_FNAM1, lngY)
6880                  .Controls("ctl" & CStr(lngX)) = arr_varRpt(R_CNAM1, lngY)
6890                  lngTmp04 = lngY
6900                  Exit For
6910                End If
6920              Next
6930            End Select
6940            If lngX = 4 Then Exit For
6950          Next

6960          If blnSpecialHandling = True Then

6970            If lngTmp02 >= 0 Then

6980              If Left(arr_varRpt(R_RNAM, lngE), 17) = "rptAccountProfile" Then
6990                .cmbMenu6 = arr_varRpt(R_FID1, lngTmp03)
7000                .cmbMenu6.ForeColor = CLR_WHT
7010                .cmbMenu6.BackColor = CLR_DKBLU
7020                .cmbMenu6.Visible = True
7030                Select Case arr_varRpt(R_FNAM1, lngTmp03)
                    Case "frmMenu_Account"
7040                  .cmbMenu6_box1.BorderColor = .Menu_Account_Line1.BackColor
7050                  .cmbMenu6_box2.BorderColor = .Menu_Account_Line2.BackColor
7060                Case "frmMenu_Post"
7070                  .cmbMenu6_box1.BorderColor = .Menu_Post_line1.BackColor
7080                  .cmbMenu6_box2.BorderColor = .Menu_Post_line2.BackColor
7090                Case "frmMenu_Report"
7100                  .cmbMenu6_box1.BorderColor = .Menu_Report_line1.BackColor
7110                  .cmbMenu6_box2.BorderColor = .Menu_Report_line2.BackColor
7120                Case "frmMenu_Asset"
7130                  .cmbMenu6_box1.BorderColor = .Menu_Asset_line1.BackColor
7140                  .cmbMenu6_box2.BorderColor = .Menu_Asset_line2.BackColor
7150                Case "frmMenu_Utility"
7160                  .cmbMenu6_box1.BorderColor = .Menu_Utility_line1.BackColor
7170                  .cmbMenu6_box2.BorderColor = .Menu_Utility_line2.BackColor
7180                Case "frmMenu_Title"
                      ' ** Not right now.
7190                End Select
7200                .cmbMenu6_box1.Visible = True
7210                .cmbMenu6_box2.Visible = True
7220                DoEvents
7230                .cmbButton6 = arr_varRpt(R_CID1, lngTmp02)
7240                .cmbButton6.BackStyle = acBackStyleTransparent
7250                .cmbButton6.Visible = True
7260                For lngX = 1& To 15&
7270                  .Controls("cmbButton6_hline" & Right("00" & CStr(lngX), 2)).Visible = True
7280                Next
7290                .Location_arrow6b.Visible = True
7300                .Location_arrow6d.Visible = True
7310                DoEvents
7320                .cmbButton7 = arr_varRpt(R_CID1, lngTmp03)
7330                .cmbButton7.BackStyle = acBackStyleTransparent
7340                .cmbButton7.Visible = True
7350                For lngX = 1& To 15&
7360                  .Controls("cmbButton7_hline" & Right("00" & CStr(lngX), 2)).Visible = True
7370                Next
7380                .Location_arrow7a.Visible = True
7390                .Location_arrow7c.Visible = True
7400                DoEvents
7410                .cmbButton8 = arr_varRpt(R_CID1, lngTmp04)
7420                .cmbButton8.BackStyle = acBackStyleTransparent
7430                .cmbButton8.Visible = True
7440                For lngX = 1& To 15&
7450                  .Controls("cmbButton8_hline" & Right("00" & CStr(lngX), 2)).Visible = True
7460                Next
7470                .Location_arrow7b.Visible = True
7480                .Location_arrow7d.Visible = True
7490                DoEvents
7500                .opgGTRMulti = .opgGTRMulti_AcctProfile
7510                .opgGTRMulti.Visible = True
7520                .opgGTRMulti_opt1_cmd.Visible = True  ' ** Though they're transparent.
7530                .opgGTRMulti_opt2_cmd.Visible = True
7540                .opgGTRMulti_opt3_cmd.Visible = True
7550                .opgGTRMulti_opt4_cmd.Visible = True
7560                If .opgGTRMulti_opt2_cmd.Top <> .cmbButton6.Top Then
7570                  .opgGTRMulti_opt2_cmd.Top = .cmbButton6.Top
7580                  .opgGTRMulti_opt2_cmd.Left = .cmbButton6.Left
7590                  .opgGTRMulti_opt2_cmd.Visible = True
7600                End If
7610                If .opgGTRMulti_opt2.Top <> (.cmbButton6.Top + (2& * lngTpp)) Then
7620                  .opgGTRMulti_opt2.Top = (.cmbButton6.Top + (2& * lngTpp))
7630                  .opgGTRMulti_opt2_lbl.Top = .opgGTRMulti_opt2.Top
7640                End If
7650              ElseIf arr_varRpt(R_RNAM, lngE) = "rptChecks_Blank" Then
7660                DoEvents
7670                .opgGTRMulti = .opgGTRMulti_Checks
7680                .opgGTRMulti.Visible = True
7690                .opgGTRMulti_opt3.Visible = False
7700                .opgGTRMulti_opt4.Visible = False
7710                .opgGTRMulti_opt1_cmd.Visible = True
7720                .opgGTRMulti_opt2.Top = (.cmbViaButton1.Top + (2& * lngTpp))
7730                .opgGTRMulti_opt2_lbl.Top = .opgGTRMulti_opt2.Top
7740                .opgGTRMulti_opt2_cmd.Top = .cmbViaButton1.Top
7750                .opgGTRMulti_opt2_cmd.Left = .cmbViaButton1.Left
7760                .opgGTRMulti_opt2_cmd.Visible = True
7770              ElseIf arr_varRpt(R_RNAM, lngE) = "rptVersion_Convert" Then
7780                .cmbMenu6 = arr_varRpt(R_FID1, lngTmp02)
7790                .cmbMenu6.ForeColor = CLR_WHT
7800                .cmbMenu6.BackColor = CLR_DKBLU
7810                .cmbMenu6.Visible = True
7820                .cmbMenu6_box1.BorderColor = .Menu_Report_line1.BackColor
7830                .cmbMenu6_box2.BorderColor = .Menu_Report_line2.BackColor
7840                .cmbMenu6_box1.Visible = True
7850                .cmbMenu6_box2.Visible = True
7860                .cmbButton7 = arr_varRpt(R_CID1, lngTmp02)
7870                .cmbButton7.BackStyle = acBackStyleTransparent
7880                .cmbButton7.Visible = True
7890                For lngX = 1& To 15&
7900                  .Controls("cmbButton7_hline" & Right("00" & CStr(lngX), 2)).Visible = True
7910                Next
7920                .Location_arrow7a.Visible = True
7930                .Location_arrow7c.Visible = True
7940                .cmbButton1.Locked = False  ' ** To make it look disabled.
7950                .cmbButton1.ForeColor = WIN_CLR_DISF
7960                .cmbButton1.BorderColor = WIN_CLR_DISR
7970                .cmbButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP1, lngE)
7980                .cmbButton1_lbl_dim_hi.Visible = True
7990                .Location_arrow8a.Visible = True
8000                .Location_arrow8d.Visible = True
8010                .Location_arrow8e.Visible = True
8020                DoEvents
8030                .opgGTRMulti = 2
8040                .opgGTRMulti.Visible = True
8050                .opgGTRMulti_opt1.Enabled = False
8060                .opgGTRMulti_opt1_cmd.Visible = True
8070                .opgGTRMulti_opt1_cmd.Enabled = False
8080                .opgGTRMulti_opt2.Top = .opgGTRMulti_opt3.Top
8090                .opgGTRMulti_opt2_lbl.Top = .opgGTRMulti_opt3_lbl.Top
8100                .opgGTRMulti_opt2_cmd.Top = .opgGTRMulti_opt3_cmd.Top
8110                .opgGTRMulti_opt2_cmd.Visible = True
8120                .opgGTRMulti_opt3.Visible = False
8130                .opgGTRMulti_opt4.Visible = False
8140              ElseIf arr_varRpt(R_RNAM, lngE) = "rptTransaction_Hidden" Then
8150                DoEvents
8160                .opgGTRMulti = .opgGTRMulti_TransHide
8170                .opgGTRMulti.Visible = True
8180                .opgGTRMulti_opt1.Top = (.cmbViaButton1.Top + (2& * lngTpp))
8190                .opgGTRMulti_opt1_lbl.Top = .opgGTRMulti_opt1.Top
8200                .opgGTRMulti_opt1_cmd.Top = .cmbViaButton1.Top
8210                .opgGTRMulti_opt1_cmd.Left = .cmbViaButton1.Left
8220                .opgGTRMulti_opt1_cmd.Visible = True
8230                .opgGTRMulti_opt2.Top = (.cmbViaButton2.Top + (2& * lngTpp))
8240                .opgGTRMulti_opt2_lbl.Top = .opgGTRMulti_opt2.Top
8250                .opgGTRMulti_opt2_cmd.Top = .cmbViaButton2.Top
8260                .opgGTRMulti_opt2_cmd.Left = .cmbViaButton2.Left
8270                .opgGTRMulti_opt2_cmd.Visible = True
8280                .opgGTRMulti_opt3.Visible = False
8290                .opgGTRMulti_opt4.Visible = False
8300              ElseIf arr_varRpt(R_RNAM, lngE) = "rptListOfTaxLots" Then
8310                .cmbViaForm2 = arr_varRpt(R_FID2, lngTmp02)
8320                .cmbViaButton2 = arr_varRpt(R_CID2, lngTmp02)
8330                .cmbViaForm2.ForeColor = CLR_WHT
8340                .cmbViaForm2.BackColor = CLR_DKBLU
8350                DoEvents
8360                .cmbViaForm2.Visible = True
8370                DoEvents
8380                .Location_arrow5a.Visible = True
8390                .Location_arrow5c.Visible = True
8400                DoEvents
8410                .cmbViaButton2_lbl.Caption = "Field"
8420                .cmbViaButton2.ForeColor = CLR_WHT
8430                .cmbViaButton2.BackColor = CLR_DKBLU
8440                .cmbViaButton2.Visible = True
8450                DoEvents
8460                .opgGTRMulti = .opgGTRMulti_TaxLots
8470                .opgGTRMulti.Visible = True
8480                .opgGTRMulti_opt3.Visible = False
8490                .opgGTRMulti_opt4.Visible = False
8500                .opgGTRMulti_opt1_cmd.Top = .cmbViaButton1.Top
8510                .opgGTRMulti_opt1_cmd.Left = .cmbViaButton1.Left
8520                .opgGTRMulti_opt1_cmd.Visible = True
8530                .opgGTRMulti_opt1.Top = (.cmbViaButton1.Top + (2& * lngTpp))
8540                .opgGTRMulti_opt1_lbl.Top = .opgGTRMulti_opt1.Top
8550                .opgGTRMulti_opt2_cmd.Top = .cmbViaButton2.Top
8560                .opgGTRMulti_opt2_cmd.Left = .cmbViaButton2.Left
8570                .opgGTRMulti_opt2_cmd.Visible = True
8580                .opgGTRMulti_opt2.Top = (.cmbViaButton2.Top + (2& * lngTpp))
8590                .opgGTRMulti_opt2_lbl.Top = .opgGTRMulti_opt2.Top
8600              ElseIf arr_varRpt(R_RNAM, lngE) = "rptMap_Dividend" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Interest" Or _
                      arr_varRpt(R_RNAM, lngE) = "rptMap_Received" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Split" Or _
                      arr_varRpt(R_RNAM, lngE) = "rptMap_Reinvest_Div" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Reinvest_Int" Or _
                      arr_varRpt(R_RNAM, lngE) = "rptMap_Reinvest_Rec" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Misc_LTCL" Or _
                      arr_varRpt(R_RNAM, lngE) = "rptMap_Misc_STCG" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Misc_STCL" Then
8610                .cmbViaForm2 = arr_varRpt(R_FID2, lngTmp02)
8620                .cmbViaButton2 = arr_varRpt(R_CID2, lngTmp02)
8630                .cmbViaForm2.ForeColor = CLR_WHT
8640                .cmbViaForm2.BackColor = CLR_DKBLU
8650                DoEvents
8660                .cmbViaForm2.Visible = True
8670                DoEvents
8680                .Location_arrow5a.Visible = True
8690                .Location_arrow5c.Visible = True
8700                DoEvents
8710                .cmbViaButton2.Visible = True
8720                DoEvents
8730                For lngX = 1& To 15&
8740                  strTmp01 = "cmbViaButton2_hline"
8750                  strTmp01 = strTmp01 & Right("00" & CStr(lngX), 2)
8760                  .Controls(strTmp01).Visible = True
8770                Next
8780                .cmbViaButton2.BackStyle = acBackStyleTransparent
8790                DoEvents
8800                Select Case arr_varRpt(R_RNAM, lngE)
                    Case "rptMap_Dividend"
8810                  .opgGTRMulti = .opgGTRMulti_Div
8820                Case "rptMap_Interest"
8830                  .opgGTRMulti = .opgGTRMulti_Int
8840                Case "rptMap_Received"
8850                  .opgGTRMulti = .opgGTRMulti_Rec
8860                Case "rptMap_Split"
8870                  .opgGTRMulti = .opgGTRMulti_Split
8880                Case "rptMap_Reinvest_Div"
8890                  .opgGTRMulti = .opgGTRMulti_DivReinvest
8900                Case "rptMap_Reinvest_Int"
8910                  .opgGTRMulti = .opgGTRMulti_IntReinvest
8920                Case "rptMap_Reinvest_Rec"
8930                  .opgGTRMulti = .opgGTRMulti_RecReinvest
8940                Case "rptMap_Misc_LTCL"
8950                  .opgGTRMulti = .opgGTRMulti_LTCL
8960                Case "rptMap_Misc_STCG"
8970                  .opgGTRMulti = .opgGTRMulti_STCG
8980                Case "rptMap_Misc_STCL"
8990                  .opgGTRMulti = .opgGTRMulti_STCL
9000                End Select
9010                .opgGTRMulti.Visible = True
9020                .opgGTRMulti_opt3.Visible = False
9030                .opgGTRMulti_opt4.Visible = False
9040                .opgGTRMulti_opt1_cmd.Top = .cmbViaButton1.Top
9050                .opgGTRMulti_opt1_cmd.Left = .cmbViaButton1.Left
9060                .opgGTRMulti_opt1_cmd.Visible = True
9070                .opgGTRMulti_opt1.Top = (.cmbViaButton1.Top + (2& * lngTpp))
9080                .opgGTRMulti_opt1_lbl.Top = .opgGTRMulti_opt1.Top
9090                .opgGTRMulti_opt2_cmd.Top = .cmbViaButton2.Top
9100                .opgGTRMulti_opt2_cmd.Left = .cmbViaButton2.Left
9110                .opgGTRMulti_opt2_cmd.Visible = True
9120                .opgGTRMulti_opt2.Top = (.cmbViaButton2.Top + (2& * lngTpp))
9130                .opgGTRMulti_opt2_lbl.Top = .opgGTRMulti_opt2.Top
9140              End If
9150            End If  ' ** lngTmp02.

9160            If lngTmp03 >= 0& Then
9170              If arr_varRpt(R_FNAM1, lngTmp03) = arr_varRpt(R_FNAM1, lngTmp02) Then
9180                .cmbButton8 = arr_varRpt(R_CID1, lngTmp03)
9190                .cmbButton8.BackStyle = acBackStyleTransparent
9200                .cmbButton8.Visible = True
9210                For lngX = 1& To 15&
9220                  .Controls("cmbButton8_hline" & Right("00" & CStr(lngX), 2)).Visible = True
9230                Next
9240                .Location_arrow7b.Visible = True
9250                .Location_arrow7d.Visible = True
9260              End If
9270            End If

9280            If lngTmp04 >= 0& Then
9290              If arr_varRpt(R_FNAM1, lngTmp04) = arr_varRpt(R_FNAM1, lngE) Then
9300                .cmbButton6 = arr_varRpt(R_CID1, lngTmp04)
9310                .cmbButton6.BackStyle = acBackStyleTransparent
9320                .cmbButton6.Visible = True
9330                For lngX = 1& To 15&
9340                  .Controls("cmbButton6_hline" & Right("00" & CStr(lngX), 2)).Visible = True
9350                Next
9360                .Location_arrow6b.Visible = True
9370                .Location_arrow6d.Visible = True
9380              End If
9390            End If

9400          End If  ' ** blnSpecialHandling.

9410        Else
9420          .Multi = Null
9430          .frm1 = Null
9440          .frm2 = Null
9450          .frm3 = Null
9460          .frm4 = Null
9470          .ctl1 = Null
9480          .ctl2 = Null
9490          .ctl3 = Null
9500          .ctl4 = Null
9510          .cmbButton6.Visible = False
9520          .cmbButton7.Visible = False
9530          .cmbButton8.Visible = False
9540          For lngX = 1& To 15&
9550            .Controls("cmbButton6_hline" & Right("00" & CStr(lngX), 2)).Visible = False
9560            .Controls("cmbButton7_hline" & Right("00" & CStr(lngX), 2)).Visible = False
9570            .Controls("cmbButton8_hline" & Right("00" & CStr(lngX), 2)).Visible = False
9580          Next
9590          .Location_arrow6b.Visible = False
9600          .Location_arrow6d.Visible = False
9610          .cmbMenu6.Visible = False
9620          .cmbMenu6_box1.Visible = False
9630          .cmbMenu6_box2.Visible = False
9640          .Location_arrow7a.Visible = False
9650          .Location_arrow7b.Visible = False
9660          .Location_arrow7c.Visible = False
9670          .Location_arrow7d.Visible = False
9680        End If

9690        .cmbMenu1 = arr_varRpt(R_FID1, lngE)  '.lbxReports.Column(LBX_RPT_UNQID)
9700        .cmbButton1 = arr_varRpt(R_CID1, lngE) '.lbxReports.Column(LBX_RPT_UNQID)
9710        .cmbMenu1.ForeColor = CLR_WHT
9720        .cmbMenu1.BackColor = CLR_DKBLU
9730        .cmbButton1.BackStyle = acBackStyleTransparent
9740        DoEvents

9750        Select Case arr_varRpt(R_FNAM1, lngE)  '.lbxReports.Column(LBX_RPT_MFCAP1)
            Case "frmMenu_Account"
9760          .cmbMenu1_box1.BorderColor = .Menu_Account_Line1.BackColor
9770          .cmbMenu1_box2.BorderColor = .Menu_Account_Line2.BackColor
9780        Case "frmMenu_Post"
9790          .cmbMenu1_box1.BorderColor = .Menu_Post_line1.BackColor
9800          .cmbMenu1_box2.BorderColor = .Menu_Post_line2.BackColor
9810        Case "frmMenu_Report"
9820          .cmbMenu1_box1.BorderColor = .Menu_Report_line1.BackColor
9830          .cmbMenu1_box2.BorderColor = .Menu_Report_line2.BackColor
9840        Case "frmMenu_Asset"
9850          .cmbMenu1_box1.BorderColor = .Menu_Asset_line1.BackColor
9860          .cmbMenu1_box2.BorderColor = .Menu_Asset_line2.BackColor
9870        Case "frmMenu_Utility"
9880          .cmbMenu1_box1.BorderColor = .Menu_Utility_line1.BackColor
9890          .cmbMenu1_box2.BorderColor = .Menu_Utility_line2.BackColor
9900        Case "frmMenu_Title"
9910          .cmbMenu1_box1.BorderColor = .Menu_Title_Line1.BackColor
9920          .cmbMenu1_box2.BorderColor = .Menu_Title_Line2.BackColor
9930          .cmbMenu1_box3.Visible = True
9940        End Select
9950        .cmbMenu1_box1.Visible = True
9960        .cmbMenu1_box2.Visible = True

9970        If IsNull(arr_varRpt(R_FNAM2, lngE)) = False Then '.lbxReports.Column(LBX_RPT_MFCAP2)) = False Then
9980          If IsMenu(arr_varRpt(R_TYP2, lngE)) = True Then  ' ** Function: Below.
9990            Select Case arr_varRpt(R_FNAM2, lngE) '.lbxReports.Column(LBX_RPT_MFCAP2)
                Case "frmMenu_CourtReport"
10000             .cmbMenu2_box1.BorderColor = .Menu_Court_line1.BackColor
10010             .cmbMenu2_box2.BorderColor = .Menu_Court_line2.BackColor
10020           Case "frmMenu_Maintenance"
10030             .cmbMenu2_box1.BorderColor = .Menu_Maint_line1.BackColor
10040             .cmbMenu2_box2.BorderColor = .Menu_Maint_line2.BackColor
10050           End Select
10060         End If
10070       End If

10080       If IsNull(arr_varRpt(R_FNAM3, lngE)) = False Then '.lbxReports.Column(LBX_RPT_MFCAP3)) = False Then
10090         If IsMenu(arr_varRpt(R_TYP3, lngE)) = True Then  ' ** Function: Below.
10100           Select Case arr_varRpt(R_FNAM3, lngE) '.lbxReports.Column(LBX_RPT_MFCAP3)
                Case "frmMenu_SystemTable"
10110             .cmbMenu3_box1.BorderColor = .Menu_Other_line1.BackColor
10120             .cmbMenu3_box2.BorderColor = .Menu_Other_line2.BackColor
10130           Case "frmMenu_Admin"
10140             .cmbMenu3_box1.BorderColor = .Menu_Admin_line1.BackColor
10150             .cmbMenu3_box2.BorderColor = .Menu_Admin_line2.BackColor
10160           End Select
10170         End If
10180       End If

10190       If IsMenu(arr_varRpt(R_TYP2, lngE)) = True Then  'Menu Button
10200         blnLevel2 = True
10210       End If

10220       Select Case blnLevel2
            Case True
10230         .cmbMenu2 = arr_varRpt(R_FID2, lngE)
10240         .cmbButton2 = arr_varRpt(R_CID2, lngE)
10250         .cmbMenu2.ForeColor = CLR_WHT
10260         .cmbMenu2.BackColor = CLR_DKBLU
10270         DoEvents
10280         .cmbMenu2.Visible = True
10290         .cmbMenu2_box1.Visible = True
10300         .cmbMenu2_box2.Visible = True
10310         .cmbButton2.Visible = True
10320         DoEvents
10330         .Location_arrow1d.Visible = True
10340         .Location_arrow1e.Visible = True
10350         .Location_arrow2a.Visible = True
10360         .Location_arrow2c.Visible = True
10370         DoEvents
10380         For lngX = 1& To 15&
10390           strTmp01 = "cmbButton2_hline"
10400           strTmp01 = strTmp01 & Right("00" & CStr(lngX), 2)
10410           .Controls(strTmp01).Visible = True
10420         Next
10430         .cmbButton2.BackStyle = acBackStyleTransparent
10440         DoEvents
10450       Case False
10460         .cmbMenu2.Visible = False
10470         .cmbMenu2_box1.Visible = False
10480         .cmbMenu2_box2.Visible = False
10490         .cmbButton2.Visible = False
10500         .Location_arrow1d.Visible = False
10510         .Location_arrow1e.Visible = False
10520         .Location_arrow2a.Visible = False
10530         .Location_arrow2c.Visible = False
10540       End Select

10550       If IsMenu(arr_varRpt(R_TYP3, lngE)) = True Then  ' ** Function: Below.
10560         blnLevel3 = True
10570       End If

10580       Select Case blnLevel3
            Case True
10590         .cmbMenu3 = arr_varRpt(R_FID3, lngE)
10600         .cmbButton3 = arr_varRpt(R_CID3, lngE)
10610         .cmbMenu3.ForeColor = CLR_WHT
10620         .cmbMenu3.BackColor = CLR_DKBLU
10630         DoEvents
10640         .cmbMenu3.Visible = True
10650         .cmbMenu3_box1.Visible = True
10660         .cmbMenu3_box2.Visible = True
10670         .cmbButton3.Visible = True
10680         DoEvents
10690         .Location_arrow2d.Visible = True
10700         .Location_arrow2e.Visible = True
10710         .Location_arrow3a.Visible = True
10720         .Location_arrow3c.Visible = True
10730         DoEvents
10740         For lngX = 1& To 15&
10750           strTmp01 = "cmbButton3_hline"
10760           strTmp01 = strTmp01 & Right("00" & CStr(lngX), 2)
10770           .Controls(strTmp01).Visible = True
10780         Next
10790         .cmbButton3.BackStyle = acBackStyleTransparent
10800         DoEvents
10810       Case False
10820         .cmbMenu3.Visible = False
10830         .cmbMenu3_box1.Visible = False
10840         .cmbMenu3_box2.Visible = False
10850         .cmbButton3.Visible = False
10860         .Location_arrow2d.Visible = False
10870         .Location_arrow2e.Visible = False
10880         .Location_arrow3a.Visible = False
10890         .Location_arrow3c.Visible = False
10900       End Select

10910       If (arr_varRpt(R_TYP2, lngE) = taVia1) Or _
                (arr_varRpt(R_TYP3, lngE) = taVia1) Or _
                (arr_varRpt(R_TYP4, lngE) = taVia1) Then
10920         blnVia1 = True
10930       ElseIf lngTmp02 > 0& And (arr_varRpt(R_RNAM, lngE) = "rptChecks_Blank" Or arr_varRpt(R_RNAM, lngE) = "rptTransaction_Hidden") Then
10940         If (arr_varRpt(R_TYP2, lngTmp02) = taVia1) Or _
                  (arr_varRpt(R_TYP3, lngTmp02) = taVia1) Or _
                  (arr_varRpt(R_TYP4, lngTmp02) = taVia1) Then
10950           blnVia1 = True
10960         End If
10970       End If

10980       Select Case blnVia1
            Case True
10990         If (arr_varRpt(R_TYP2, lngE) = taVia1) Then
11000           .cmbViaForm1 = arr_varRpt(R_FID2, lngE)
11010           .cmbViaButton1 = arr_varRpt(R_CID2, lngE)
11020         ElseIf lngTmp02 > 0& Then
                'THIS SEEMS TO BE PUTTING THE WRONG ONE INTO VIA1!  SHOULDN'T VIA1 BE THE ONE I CHOSE ABOVE WITH lngE?
11030           If (arr_varRpt(R_TYP2, lngTmp02) = taVia1) Then
11040             .cmbViaForm1 = arr_varRpt(R_FID2, lngTmp02)
11050             .cmbViaButton1 = arr_varRpt(R_CID2, lngTmp02)
11060           End If
11070         ElseIf (arr_varRpt(R_TYP3, lngE) = taVia1) Then
11080           .cmbViaForm1 = arr_varRpt(R_FID3, lngE)
11090           .cmbViaButton1 = arr_varRpt(R_CID3, lngE)
11100         ElseIf (arr_varRpt(R_TYP4, lngE) = taVia1) Then
11110           .cmbViaForm1 = arr_varRpt(R_FID4, lngE)
11120           .cmbViaButton1 = arr_varRpt(R_CID4, lngE)
11130         End If
11140         If arr_varRpt(R_RNAM, lngE) = "rptTransaction_Hidden" Then
11150           .cmbViaForm1.ForeColor = CLR_DKGRY3 'CLR_VDKGRY
11160           .cmbViaForm1.BackColor = CLR_LTBLU '16314086
11170         Else
11180           .cmbViaForm1.ForeColor = CLR_WHT
11190           .cmbViaForm1.BackColor = CLR_DKBLU
11200         End If
11210         DoEvents
11220         .cmbViaForm1.Visible = True
11230         DoEvents
11240         .Location_arrow4a.Visible = True
11250         .Location_arrow4c.Visible = True
11260         DoEvents
11270         If .cmbViaButton1.Column(1) = "Cost" Then '"Principal Cash" Then
11280           .cmbViaButton1_lbl.Caption = "Field"
11290           .cmbViaButton1.ForeColor = CLR_WHT
11300           .cmbViaButton1.BackColor = CLR_DKBLU
11310           .cmbViaButton1.Visible = True
11320         Else
11330           .cmbViaButton1.Visible = True
11340           DoEvents
11350           For lngX = 1& To 15&
11360             strTmp01 = "cmbViaButton1_hline"
11370             strTmp01 = strTmp01 & Right("00" & CStr(lngX), 2)
11380             .Controls(strTmp01).Visible = True
11390           Next
11400           .cmbViaButton1.BackStyle = acBackStyleTransparent
11410         End If
11420         DoEvents
11430       Case False
11440         .cmbViaForm1.Visible = False
11450         .cmbViaButton1.Visible = False
11460         .Location_arrow4a.Visible = False
11470         .Location_arrow4c.Visible = False
11480         DoEvents
11490       End Select

11500       If arr_varRpt(R_RNAM, lngE) = "rptCurrencyRateHistory" Then
11510         For lngX = 0& To (lngCtls - 1&)
11520           .Controls(arr_varCtl(C_CNAM, lngX)).Top = (.Controls(arr_varCtl(C_CNAM, lngX)).Top + (96& * lngTpp))
11530         Next
11540       End If

            'Location_arrow4a
            'Location_arrow4c
            'cmbViaForm1
            'cmbViaForm1_lbl
            'cmbViaButton1_hline01
            'cmbViaButton1_hline02
            'cmbViaButton1_hline03
            'cmbViaButton1_hline04
            'cmbViaButton1_hline05
            'cmbViaButton1_hline06
            'cmbViaButton1_hline07
            'cmbViaButton1_hline08
            'cmbViaButton1_hline09
            'cmbViaButton1_hline10
            'cmbViaButton1_hline11
            'cmbViaButton1_hline12
            'cmbViaButton1_hline13
            'cmbViaButton1_hline14
            'cmbViaButton1_hline15
            'cmbViaButton1_lbl_dim_hi
            'cmbViaButton1
            'cmbViaButton1_lbl
            'cmbViaButton1_cmd

11550       If (arr_varRpt(R_TYP3, lngE) = taVia2) Or _
                (arr_varRpt(R_TYP4, lngE) = taVia2) Or _
                (arr_varRpt(R_TYP5, lngE) = taVia2) Then
11560         blnVia2 = True
11570       ElseIf arr_varRpt(R_RNAM, lngE) = "rptTransaction_Hidden" Then
11580         blnVia2 = True
11590       End If

11600       Select Case blnVia2
            Case True
11610         If arr_varRpt(R_RNAM, lngE) = "rptMap_Dividend" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Interest" Or _
                  arr_varRpt(R_RNAM, lngE) = "rptMap_Received" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Split" Or _
                  arr_varRpt(R_RNAM, lngE) = "rptMap_Reinvest_Div" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Reinvest_Int" Or _
                  arr_varRpt(R_RNAM, lngE) = "rptMap_Reinvest_Rec" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Misc_LTCL" Or _
                  arr_varRpt(R_RNAM, lngE) = "rptMap_Misc_STCG" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Misc_STCL" Then
                ' ** Don't undo what we did above.
11620         Else
11630           If (arr_varRpt(R_TYP3, lngE) = taVia2) Then
11640             .cmbViaForm2 = arr_varRpt(R_FID3, lngE)
11650             .cmbViaButton2 = arr_varRpt(R_CID3, lngE)
11660           ElseIf arr_varRpt(R_RNAM, lngE) = "rptTransaction_Hidden" Then
11670             .cmbViaForm2 = arr_varRpt(R_FID3, lngTmp02)
11680             .cmbViaButton2 = arr_varRpt(R_CID3, lngTmp02)
11690           ElseIf (arr_varRpt(R_TYP4, lngE) = taVia2) Then
11700             .cmbViaForm2 = arr_varRpt(R_FID4, lngE)
11710             .cmbViaButton2 = arr_varRpt(R_CID4, lngE)
11720           ElseIf (arr_varRpt(R_TYP5, lngE) = taVia2) Then
11730             .cmbViaForm2 = arr_varRpt(R_FID5, lngE)
11740             .cmbViaButton2 = arr_varRpt(R_CID5, lngE)
11750           End If
11760           .cmbViaForm2.ForeColor = CLR_WHT
11770           .cmbViaForm2.BackColor = CLR_DKBLU
11780           DoEvents
11790           .cmbViaForm2.Visible = True
11800           DoEvents
11810           .Location_arrow5a.Visible = True
11820           .Location_arrow5c.Visible = True
11830           DoEvents
11840           If .cmbViaButton2.Column(1) = "Cost" Then '"Principal Cash" Then
11850             .cmbViaButton2_lbl.Caption = "Field"
11860             .cmbViaButton2.ForeColor = CLR_WHT
11870             .cmbViaButton2.BackColor = CLR_DKBLU
11880             .cmbViaButton2.Visible = True
11890           Else
11900             .cmbViaButton2.Visible = True
11910             DoEvents
11920             For lngX = 1& To 15&
11930               strTmp01 = "cmbViaButton2_hline"
11940               strTmp01 = strTmp01 & Right("00" & CStr(lngX), 2)
11950               .Controls(strTmp01).Visible = True
11960             Next
11970             .cmbViaButton2.BackStyle = acBackStyleTransparent
11980           End If
11990           DoEvents
12000         End If
12010       Case False
12020         If arr_varRpt(R_RNAM, lngE) = "rptListOfTaxLots" And lngTmp02 >= 0 Then
                ' ** Don't turn off what we turned on above.
12030         Else
12040           .cmbViaForm2.Visible = False
12050           .cmbViaButton2.Visible = False
12060           .Location_arrow5a.Visible = False
12070           .Location_arrow5c.Visible = False
12080           DoEvents
12090         End If
12100       End Select

12110       .ArchivedTrans_lbl1.Visible = False       ' ** No Archived
12120       .ArchivedTrans_lbl2.Visible = False       ' ** Transactions
12130       .AssetPricing_lbl.Visible = False         ' ** No Assets Held
12140       .AssetPricingHistory_lbl.Visible = False  ' ** No Pricing History
12150       .ChecksTest_lbl.Visible = False           ' ** Test Check
12160       .Conversion_lbl1.Visible = False          ' ** Only Available
12170       .Conversion_lbl2.Visible = False          ' ** During Conversion
12180       .CourtReportNY_lbl.Visible = False        ' ** Currently Unavailable
12190       .ForEx_lbl.Visible = False                ' ** No Foreign Currency Assets Are Listed
12200       .HideTrans2_Select_lbl.Visible = False    ' ** Select-Existing Option
12210       .JournalDividend_lbl.Visible = False      ' ** Dividend Tab
12220       .JournalInterest_lbl.Visible = False      ' ** Interest Tab
12230       .JournalPurchase_lbl.Visible = False      ' ** Miscellaneous Tab
12240       .JournalSold_lbl.Visible = False          ' ** Purchase/Deposit Tab
12250       .JournalMisc_lbl.Visible = False          ' ** Sold/Withdrawn Tab
12260       .PostingJournal_lbl.Visible = False       ' ** Columnar Only
12270       .Pricing_lbl.Visible = False              ' ** Pricing Unavailable
12280       .TransactionAudit_lbl.Visible = False     ' ** No Transactions

12290       .cmdConvLog.Enabled = False
12300       .cmdConvLog.Visible = False

12310       varAssets = DCount("*", "masterAsset", "[shareface] > 0")
12320       If IsNull(varAssets) = True Then varAssets = CLng(0)

            ' ** Since it's possible one of these may change, we'll check 'em all.

            'HOLD CTL ON ARCHIVE TO GO DIRECTLY TO REPORTS! DO NOT PASS GO!

12330       If (arr_varRpt(R_RNAM, lngE) = "rptAssetHistory" Or arr_varRpt(R_RCAP, lngE) = "Asset History") Or _
                (arr_varRpt(R_RNAM, lngE) = "rptAssetHistory_All" Or arr_varRpt(R_RCAP, lngE) = "Asset History - All Accounts") Then
              ' ** Asset History.
12340         Select Case varAssets
              Case 0
                ' ** Button disabled.
12350           .AssetPricing_lbl.Visible = True
12360           blnNoGo = True
12370           If .cmbButton1.Visible = True Then
12380             .cmbButton1.Locked = False  ' ** To make it look disabled.
12390             .cmbButton1.ForeColor = WIN_CLR_DISF
12400             .cmbButton1.BorderColor = WIN_CLR_DISR
12410             .cmbButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP1, lngE)
12420             .cmbButton1_lbl_dim_hi.Visible = True
12430           End If
12440           .cmdLocationGo.Enabled = False
12450         Case Else
12460           .AssetPricing_lbl.Visible = False
12470           .cmbButton1.Locked = True
12480           .cmbButton1.ForeColor = CLR_BLK
12490           .cmbButton1.BorderColor = CLR_LTBLU2
12500           .cmbButton1_lbl_dim_hi.Visible = False
12510           .cmdLocationGo.Enabled = True
12520         End Select
12530       ElseIf arr_varRpt(R_RNAM, lngE) = "rptAssetPricing" Or arr_varRpt(R_RCAP, lngE) = "Asset Pricing Report" Or _
                arr_varRpt(R_RCAP, lngE) = "Asset Pricing" Then
              ' ** Asset Pricing.
12540         Select Case varAssets
              Case 0
                ' ** Button disabled.
12550           .AssetPricing_lbl.Visible = True
12560           blnNoGo = True
12570           If .cmbButton1.Visible = True Then
12580             .cmbButton1.Locked = False  ' ** To make it look disabled.
12590             .cmbButton1.ForeColor = WIN_CLR_DISF
12600             .cmbButton1.BorderColor = WIN_CLR_DISR
12610             .cmbButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP1, lngE)
12620             .cmbButton1_lbl_dim_hi.Visible = True
12630           End If
12640           .cmdLocationGo.Enabled = False
12650         Case Else
12660           .AssetPricing_lbl.Visible = False
12670           .cmbButton1.Locked = True
12680           .cmbButton1.ForeColor = CLR_BLK
12690           .cmbButton1.BorderColor = CLR_LTBLU2
12700           .cmbButton1_lbl_dim_hi.Visible = False
12710           .cmdLocationGo.Enabled = True
12720         End Select
12730       ElseIf arr_varRpt(R_RNAM, lngE) = "rptPricingUpdateCusips_prn" Or _
                arr_varRpt(R_RCAP, lngE) = "Asset Pricing Update" Then
              ' ** Asset Pricing Import.
12740         Select Case varAssets
              Case 0
                ' ** Button disabled.
12750           .AssetPricing_lbl.Visible = True
12760           blnNoGo = True
12770           If .cmbButton1.Visible = True Then
12780             .cmbButton1.Locked = False  ' ** To make it look disabled.
12790             .cmbButton1.ForeColor = WIN_CLR_DISF
12800             .cmbButton1.BorderColor = WIN_CLR_DISR
12810             .cmbButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP1, lngE)
12820             .cmbButton1_lbl_dim_hi.Visible = True
12830           End If
12840           .cmdLocationGo.Enabled = False
12850         Case Else
12860           .AssetPricing_lbl.Visible = False
12870           .cmbButton1.Locked = True
12880           .cmbButton1.ForeColor = CLR_BLK
12890           .cmbButton1.BorderColor = CLR_LTBLU2
12900           .cmbButton1_lbl_dim_hi.Visible = False
12910           .cmdLocationGo.Enabled = True
12920         End Select
12930         Select Case gblnPricingAllowed
              Case True
12940           Select Case varAssets
                Case 0
                  ' ** Button disabled.
12950             .Pricing_lbl.Visible = True
12960             blnNoGo = True
12970             If .cmbViaButton1.Visible = True Then
12980               .cmbViaButton1.Locked = False  ' ** To make it look disabled.
12990               .cmbViaButton1.ForeColor = WIN_CLR_DISF
13000               .cmbViaButton1.BorderColor = WIN_CLR_DISR
13010               If (arr_varRpt(R_TYP2, lngE) = taVia1) Then
13020                 .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP2, lngE)
13030               ElseIf (arr_varRpt(R_TYP3, lngE) = taVia1) Then
13040                 .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP3, lngE)
13050               ElseIf (arr_varRpt(R_TYP4, lngE) = taVia1) Then
13060                 .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP4, lngE)
13070               End If
13080               .cmbViaButton1_lbl_dim_hi.Visible = True  ' ** Has 33 backcolor, and 28 forecolor.
13090             End If
13100             .cmdLocationGo.Enabled = False
13110           Case Else
13120             .cmbViaButton1.Locked = True
13130             .cmbViaButton1.ForeColor = CLR_BLK
13140             .cmbViaButton1.BorderColor = CLR_LTBLU2
13150             .cmbViaButton1_lbl_dim_hi.Visible = False
13160             .cmdLocationGo.Enabled = True
13170           End Select
13180         Case False
                ' ** Button disabled.
13190           .Pricing_lbl.Visible = True
13200           blnNoGo = True
13210           If .cmbViaButton1.Visible = True Then
13220             .cmbViaButton1.Locked = False  ' ** To make it look disabled.
13230             .cmbViaButton1.ForeColor = WIN_CLR_DISF
13240             .cmbViaButton1.BorderColor = WIN_CLR_DISR
13250             If (arr_varRpt(R_TYP2, lngE) = taVia1) Then
13260               .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP2, lngE)
13270             ElseIf (arr_varRpt(R_TYP3, lngE) = taVia1) Then
13280               .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP3, lngE)
13290             ElseIf (arr_varRpt(R_TYP4, lngE) = taVia1) Then
13300               .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP4, lngE)
13310             End If
13320             .cmbViaButton1_lbl_dim_hi.Visible = True  ' ** Has 33 backcolor, and 28 forecolor.
13330           End If
13340           .cmdLocationGo.Enabled = False
13350         End Select
13360       ElseIf arr_varRpt(R_RNAM, lngE) = "rptAssetPricing_History" Or _
                arr_varRpt(R_RCAP, lngE) = "Asset Pricing History" Then
              ' ** Pricing History.
13370         Select Case varAssets
              Case 0
                ' ** Button disabled.
13380           .AssetPricing_lbl.Visible = True
13390           blnNoGo = True
13400           If .cmbButton1.Visible = True Then
13410             .cmbButton1.Locked = False  ' ** To make it look disabled.
13420             .cmbButton1.ForeColor = WIN_CLR_DISF
13430             .cmbButton1.BorderColor = WIN_CLR_DISR
13440             .cmbButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP1, lngE)
13450             .cmbButton1_lbl_dim_hi.Visible = True
13460           End If
13470           .cmdLocationGo.Enabled = False
13480         Case Else
13490           .AssetPricing_lbl.Visible = False
13500           .cmbButton1.Locked = True
13510           .cmbButton1.ForeColor = CLR_BLK
13520           .cmbButton1.BorderColor = CLR_LTBLU2
13530           .cmbButton1_lbl_dim_hi.Visible = False
13540           .cmdLocationGo.Enabled = True
13550         End Select
13560         varTmp00 = DCount("*", "tblPricing_MasterAsset_History")
13570         Select Case IsNull(varTmp00)
              Case True
                ' ** Button disabled.
13580           .AssetPricingHistory_lbl.Visible = True
13590           blnNoGo = True
13600           If .cmbViaButton1.Visible = True Then
13610             .cmbViaButton1.Locked = False  ' ** To make it look disabled.
13620             .cmbViaButton1.ForeColor = WIN_CLR_DISF
13630             .cmbViaButton1.BorderColor = WIN_CLR_DISR
13640             If (arr_varRpt(R_TYP2, lngE) = taVia1) Then
13650               .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP2, lngE)
13660             ElseIf (arr_varRpt(R_TYP3, lngE) = taVia1) Then
13670               .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP3, lngE)
13680             ElseIf (arr_varRpt(R_TYP4, lngE) = taVia1) Then
13690               .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP4, lngE)
13700             End If
13710             .cmbViaButton1_lbl_dim_hi.Visible = True  ' ** Has 33 backcolor, and 28 forecolor.
13720           End If
13730           .cmdLocationGo.Enabled = False
13740         Case False
13750           Select Case varTmp00
                Case 0
                  ' ** Button disabled.
13760             .AssetPricingHistory_lbl.Visible = True
13770             blnNoGo = True
13780             If .cmbViaButton1.Visible = True Then
13790               .cmbViaButton1.Locked = False  ' ** To make it look disabled.
13800               .cmbViaButton1.ForeColor = WIN_CLR_DISF
13810               .cmbViaButton1.BorderColor = WIN_CLR_DISR
13820               If (arr_varRpt(R_TYP2, lngE) = taVia1) Then
13830                 .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP2, lngE)
13840               ElseIf (arr_varRpt(R_TYP3, lngE) = taVia1) Then
13850                 .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP3, lngE)
13860               ElseIf (arr_varRpt(R_TYP4, lngE) = taVia1) Then
13870                 .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP4, lngE)
13880               End If
13890               .cmbViaButton1_lbl_dim_hi.Visible = True  ' ** Has 33 backcolor, and 28 forecolor.
13900             End If
13910             .cmdLocationGo.Enabled = False
13920           Case Else
13930             .AssetPricingHistory_lbl.Visible = False
13940             .cmbViaButton1.Locked = True
13950             .cmbViaButton1.ForeColor = CLR_BLK
13960             .cmbViaButton1.BorderColor = CLR_LTBLU2
13970             .cmbViaButton1_lbl_dim_hi.Visible = False
13980             .cmdLocationGo.Enabled = True
13990           End Select
14000         End Select
14010       ElseIf arr_varRpt(R_RNAM, lngE) = "rptArchivedTransactions" Or _
                arr_varRpt(R_RCAP, lngE) = "Archived Transaction Statement" Then
              ' ** Archived Transactions.
14020         varTmp00 = DCount("*", "LedgerArchive")
14030         Select Case IsNull(varTmp00)
              Case True
                ' ** Button disabled.
14040           .ArchivedTrans_lbl1.Visible = True
14050           .ArchivedTrans_lbl2.Visible = True
14060           blnNoGo = True
14070           If .cmbViaButton1.Visible = True Then
14080             .cmbViaButton1.Locked = False  ' ** To make it look disabled.
14090             .cmbViaButton1.ForeColor = WIN_CLR_DISF
14100             .cmbViaButton1.BorderColor = WIN_CLR_DISR
14110             If (arr_varRpt(R_TYP2, lngE) = taVia1) Then
14120               .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP2, lngE)
14130             ElseIf (arr_varRpt(R_TYP3, lngE) = taVia1) Then
14140               .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP3, lngE)
14150             ElseIf (arr_varRpt(R_TYP4, lngE) = taVia1) Then
14160               .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP4, lngE)
14170             End If
14180             .cmbViaButton1_lbl_dim_hi.Visible = True  ' ** Has 33 backcolor, and 28 forecolor.
14190           End If
14200           .cmdLocationGo.Enabled = False
14210         Case False
14220           Select Case varTmp00
                Case 0
                  ' ** Button disabled.
14230             .ArchivedTrans_lbl1.Visible = True
14240             .ArchivedTrans_lbl2.Visible = True
14250             blnNoGo = True
14260             If .cmbViaButton1.Visible = True Then
14270               .cmbViaButton1.Locked = False  ' ** To make it look disabled.
14280               .cmbViaButton1.ForeColor = WIN_CLR_DISF
14290               .cmbViaButton1.BorderColor = WIN_CLR_DISR
14300               If (arr_varRpt(R_TYP2, lngE) = taVia1) Then
14310                 .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP2, lngE)
14320               ElseIf (arr_varRpt(R_TYP3, lngE) = taVia1) Then
14330                 .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP3, lngE)
14340               ElseIf (arr_varRpt(R_TYP4, lngE) = taVia1) Then
14350                 .cmbViaButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP4, lngE)
14360               End If
14370               .cmbViaButton1_lbl_dim_hi.Visible = True  ' ** Has 33 backcolor, and 28 forecolor.
14380             End If
14390             .cmdLocationGo.Enabled = False
14400           Case Else
14410             .ArchivedTrans_lbl1.Visible = False
14420             .ArchivedTrans_lbl2.Visible = False
14430             .cmbViaButton1.Locked = True
14440             .cmbViaButton1.ForeColor = CLR_BLK
14450             .cmbViaButton1.BorderColor = CLR_LTBLU2
14460             .cmbViaButton1_lbl_dim_hi.Visible = False
14470             .cmdLocationGo.Enabled = True
14480           End Select
14490         End Select
14500       ElseIf arr_varRpt(R_RNAM, lngE) = "rptTransaction_Audit_01" Or _
                arr_varRpt(R_RCAP, lngE) = "Transaction Audit" Then
              ' ** Transaction Audit.
14510         .cmbButton1.ForeColor = CLR_BLK 'CLR_DKBLU
14520         varTmp00 = DCount("*", "ledger")
14530         Select Case IsNull(varTmp00)
              Case True
                ' ** Button disabled.
14540           .TransactionAudit_lbl.Visible = True
14550           blnNoGo = True
14560           If .cmbButton2.Visible = True Then
14570             .cmbButton2.Locked = False  ' ** To make it look disabled.
14580             .cmbButton2.ForeColor = WIN_CLR_DISF
14590             .cmbButton2.BorderColor = WIN_CLR_DISR
14600             .cmbButton2_lbl_dim_hi.Caption = arr_varRpt(R_CCAP2, lngE)
14610             .cmbButton2_lbl_dim_hi.Visible = True
14620           End If
14630           .cmdLocationGo.Enabled = False
14640         Case False
14650           Select Case varTmp00
                Case 0
                  ' ** Button disabled.
14660             .TransactionAudit_lbl.Visible = True
14670             blnNoGo = True
14680             If .cmbButton2.Visible = True Then
14690               .cmbButton2.Locked = False  ' ** To make it look disabled.
14700               .cmbButton2.ForeColor = WIN_CLR_DISF
14710               .cmbButton2.BorderColor = WIN_CLR_DISR
14720               .cmbButton2_lbl_dim_hi.Caption = arr_varRpt(R_CCAP2, lngE)
14730               .cmbButton2_lbl_dim_hi.Visible = True
14740             End If
14750             .cmdLocationGo.Enabled = False
14760           Case Else
14770             .TransactionAudit_lbl.Visible = False
14780             .cmbButton2.Locked = True
14790             .cmbButton2.ForeColor = CLR_BLK 'CLR_DKBLU
14800             .cmbButton2.BorderColor = CLR_LTBLU2
14810             .cmbButton2_lbl_dim_hi.Visible = False
14820             .cmdLocationGo.Enabled = True
14830           End Select
14840         End Select
14850       ElseIf arr_varRpt(R_RNAM, lngE) = "rptAccountSummary_ForEx" Or arr_varRpt(R_RNAM, lngE) = "rptAccountSummary_ForEx" Or _
                arr_varRpt(R_RNAM, lngE) = "rptAssetList_ForEx" Or arr_varRpt(R_RNAM, lngE) = "rptAssetPricing_ForEx" Or _
                arr_varRpt(R_RNAM, lngE) = "rptPostingJournal_Classic_ForEx" Or _
                arr_varRpt(R_RNAM, lngE) = "rptTransaction_Statement_ForEx_SortType" Then
14860         Select Case gblnHasForEx
              Case True
                ' ** ForEx_lbl remains off; proceed normally.
14870         Case False
                ' ** Button disabled.
14880           .ForEx_lbl.Visible = True
14890           blnNoGo = True
14900           If .cmbButton1.Visible = True Then
14910             .cmbButton1.Locked = False  ' ** To make it look disabled.
14920             .cmbButton1.ForeColor = WIN_CLR_DISF
14930             .cmbButton1.BorderColor = WIN_CLR_DISR
14940             .cmbButton1_lbl_dim_hi.Caption = arr_varRpt(R_CCAP1, lngE)
14950             .cmbButton1_lbl_dim_hi.Visible = True
14960           End If
14970           .cmdLocationGo.Enabled = False
14980         End Select
14990       ElseIf arr_varRpt(R_RNAM, lngE) = "rptPostingJournal_Column" Or _
                arr_varRpt(R_RCAP, lngE) = "Post Journal - Including Uncommitted Entries" Then
15000         .PostingJournal_lbl.Visible = True
15010       ElseIf (arr_varRpt(R_RNAM, lngE) = "rptMap_Dividend" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Reinvest_Div") And _
                .cmbViaForm1.Column(2) = "frmJournal" Then
15020         .JournalDividend_lbl.Visible = True
15030       ElseIf (arr_varRpt(R_RNAM, lngE) = "rptMap_Interest" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Reinvest_Int") And _
                .cmbViaForm1.Column(2) = "frmJournal" Then
15040         .JournalInterest_lbl.Visible = True
15050       ElseIf (arr_varRpt(R_RNAM, lngE) = "rptMap_Received" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Reinvest_Rec" Or _
                arr_varRpt(R_RNAM, lngE) = "rptMap_Misc_LTCL" Or arr_varRpt(R_RNAM, lngE) = "rptMap_Misc_STCG" Or _
                arr_varRpt(R_RNAM, lngE) = "rptMap_Misc_STCL") And .cmbViaForm1.Column(2) = "frmJournal" Then
15060         .JournalMisc_lbl.Visible = True
15070       ElseIf arr_varRpt(R_RNAM, lngE) = "rptMap_Split" And .cmbViaForm1.Column(2) = "frmJournal" Then
15080         .JournalPurchase_lbl.Visible = True
15090       ElseIf arr_varRpt(R_RNAM, lngE) = "rptListOfTaxLots" And .cmbViaForm1.Column(2) = "frmJournal" Then
15100         .JournalSold_lbl.Visible = True
15110       ElseIf arr_varRpt(R_RNAM, lngE) = "rptChecks_Blank" Then
15120         .ChecksTest_lbl.Visible = True
15130       ElseIf arr_varRpt(R_RNAM, lngE) = "rptVersion_Convert" Then
15140         .Conversion_lbl1.Visible = True
15150         .Conversion_lbl2.Visible = True
15160         If ConvLogChk = True Then  ' ** Function: Below.
15170           .cmdConvLog.Enabled = True
15180           .cmdConvLog.Visible = True
15190           .cmdLocationGo.Enabled = True
15200         Else
15210           .cmdLocationGo.Enabled = False
15220         End If
15230       ElseIf .cmbButton1.Column(1) = "Special Maintenance" Then
15240         .cmbButton1.ForeColor = CLR_DKBLU
15250         .cmbButton2.ForeColor = CLR_BLK 'CLR_DKBLU
15260       ElseIf arr_varRpt(R_RNAM, lngE) = "rptTransaction_Hidden" Then
15270         .HideTrans2_Select_lbl.Visible = True
15280       ElseIf arr_varRpt(R_RNAM, lngE) = "rptCourtRptNY_13" Then
15290         .CourtReportNY_lbl.Visible = True
15300         If .cmbButton2.Visible = True Then
15310           .cmbButton2.Locked = False  ' ** To make it look disabled.
15320           .cmbButton2.ForeColor = WIN_CLR_DISF
15330           .cmbButton2.BorderColor = WIN_CLR_DISR
15340           .cmbButton2_lbl_dim_hi.Caption = arr_varRpt(R_CCAP2, lngE)
15350           .cmbButton2_lbl_dim_hi.Visible = True
15360         End If
15370         .cmdLocationGo.Enabled = False
15380       Else
              ' ** Nothing else yet.
15390       End If

15400       Select Case arr_varRpt(R_RNAM, lngE)
            Case "rptArchivedTransactions", "rptAssetPricing_History", "rptPricingUpdateCusips_prn", _
                "rptListOfRelatedAccounts", "rptStatement_Qualified", "rptStatement_Balance"
              ' ** cmbViaButton2 is never used without the option buttons.
15410         .cmbViaButton1_cmd.Enabled = True
15420         .cmbViaButton1_cmd.Visible = True
15430       Case "rptVersion_Convert"
              ' ** Just the option command buttons.
15440       Case "rptChecks_Blank", "rptTransaction_Hidden", "rptListOfTaxLots", "rptListOfTaxLots"
              ' ** Just the option command buttons.
15450       Case "rptMap_Dividend", "rptMap_Reinvest_Div", "rptMap_Interest", "rptMap_Reinvest_Int", "rptMap_Received", _
                "rptMap_Reinvest_Rec", "rptMap_Split", "rptMap_Misc_LTCL", "rptMap_Misc_STCG", "rptMap_Misc_STCL"
              ' ** Just the option command buttons.
15460       Case "rptAccountProfile_SepChkAcct", "rptAccountProfile_SngChkAcct"
              ' ** This is the only one using cmbButton6, cmbButton7, and cmbButton8,
              ' ** so just the option command buttons.
15470       Case Else
              ' ** Only the last in the series is enabled.
15480         If .cmbButton3.Visible Then
15490           .cmbButton3_cmd.Enabled = True
15500           .cmbButton3_cmd.Visible = True
15510         ElseIf .cmbButton2.Visible = True Then
15520           .cmbButton2_cmd.Enabled = True
15530           .cmbButton2_cmd.Visible = True
15540         ElseIf .cmbButton1.Visible = True Then
15550           .cmbButton1_cmd.Enabled = True
15560           .cmbButton1_cmd.Visible = True
15570         End If
15580       End Select

15590       .rptlist_id = arr_varRpt(R_RLID, lngE)

15600     Else
15610       Beep
15620       MsgBox "Problem...", vbCritical + vbOKOnly, "Problem!"
15630     End If  ' ** blnContinue.

15640   End With

EXITP:
15650   Exit Sub

ERRH:
15660   Select Case ERR.Number
        Case Else
15670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15680   End Select
15690   Resume EXITP

End Sub

Private Sub lbxReports_KeyDown(KeyCode As Integer, Shift As Integer)

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "lbxReports_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15720   intShiftDown = (Shift And acShiftMask) > 0
15730   intAltDown = (Shift And acAltMask) > 0
15740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15770       With Me
15780         intRetVal = 0
15790         If .cmdLocationGo.Enabled = True Then
15800           .cmdLocationGo.SetFocus
15810         ElseIf .cmdConvLog.Visible = True And .cmdConvLog.Enabled = True Then
15820           .cmdConvLog.SetFocus
15830         Else
15840           .cmdClose.SetFocus
15850         End If
15860       End With
15870     End Select
15880   End If

        ' ** Shift keys.
15890   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15900     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15910       With Me
15920         intRetVal = 0
15930         .cmdClose.SetFocus
15940       End With
15950     End Select
15960   End If

EXITP:
15970   KeyCode = intRetVal
15980   Exit Sub

ERRH:
15990   intRetVal = 0
16000   Select Case ERR.Number
        Case Else
16010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16020   End Select
16030   Resume EXITP

End Sub

Private Sub cmbButton1_cmd_DblClick(Cancel As Integer)

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbButton1_cmd_DblClick"

16110   cmdLocationGo_Click  ' ** Procedure: Below.

EXITP:
16120   Exit Sub

ERRH:
16130   Select Case ERR.Number
        Case Else
16140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16150   End Select
16160   Resume EXITP

End Sub

Private Sub cmbButton2_cmd_DblClick(Cancel As Integer)

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbButton2_cmd_DblClick"

16210   cmdLocationGo_Click  ' ** Procedure: Below.

EXITP:
16220   Exit Sub

ERRH:
16230   Select Case ERR.Number
        Case Else
16240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16250   End Select
16260   Resume EXITP

End Sub

Private Sub cmbButton3_cmd_DblClick(Cancel As Integer)

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbButton3_cmd_DblClick"

16310   cmdLocationGo_Click  ' ** Procedure: Below.

EXITP:
16320   Exit Sub

ERRH:
16330   Select Case ERR.Number
        Case Else
16340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16350   End Select
16360   Resume EXITP

End Sub

Private Sub cmbViaButton1_cmd_DblClick(Cancel As Integer)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbViaButton1_cmd_DblClick"

16410   cmdLocationGo_Click  ' ** Procedure: Below.

EXITP:
16420   Exit Sub

ERRH:
16430   Select Case ERR.Number
        Case Else
16440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16450   End Select
16460   Resume EXITP

End Sub

Private Sub opgGTRMulti_AfterUpdate()

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGTRMulti_AfterUpdate"

        Dim varTmp00 As Variant

16510   With Me

16520     Select Case .opgGTRMulti
          Case .opgGTRMulti_opt1.OptionValue
16530       .opgGTRMulti_opt1_lbl.FontBold = True
16540       .opgGTRMulti_opt2_lbl.FontBold = False
16550       .opgGTRMulti_opt3_lbl.FontBold = False
16560       .opgGTRMulti_opt4_lbl.FontBold = False
16570     Case .opgGTRMulti_opt2.OptionValue
16580       .opgGTRMulti_opt1_lbl.FontBold = False
16590       .opgGTRMulti_opt2_lbl.FontBold = True
16600       .opgGTRMulti_opt3_lbl.FontBold = False
16610       .opgGTRMulti_opt4_lbl.FontBold = False
16620     Case .opgGTRMulti_opt3.OptionValue
16630       .opgGTRMulti_opt1_lbl.FontBold = False
16640       .opgGTRMulti_opt2_lbl.FontBold = False
16650       .opgGTRMulti_opt3_lbl.FontBold = True
16660       .opgGTRMulti_opt4_lbl.FontBold = False
16670     Case .opgGTRMulti_opt4.OptionValue
16680       .opgGTRMulti_opt1_lbl.FontBold = False
16690       .opgGTRMulti_opt2_lbl.FontBold = False
16700       .opgGTRMulti_opt3_lbl.FontBold = False
16710       .opgGTRMulti_opt4_lbl.FontBold = True
16720     End Select

16730     varTmp00 = .lbxReports.Column(LBX_RPT_RNAM)
16740     If IsNull(varTmp00) = False Then

            ' ** Save their separate choices for Prefs.
16750       Select Case .lbxReports.Column(LBX_RPT_RNAM)
            Case "rptAccountProfile_SepChkAcct"
16760         .opgGTRMulti_AcctProfile = .opgGTRMulti
16770       Case "rptTransaction_Hidden"
16780         .opgGTRMulti_TransHide = .opgGTRMulti
16790       Case "rptChecks_Blank"
16800         .opgGTRMulti_Checks = .opgGTRMulti
16810       Case "rptListOfTaxLots"
16820         .opgGTRMulti_TaxLots = .opgGTRMulti
16830       Case "rptMap_Dividend"
16840         .opgGTRMulti_Div = .opgGTRMulti
16850       Case "rptMap_Interest"
16860         .opgGTRMulti_Int = .opgGTRMulti
16870       Case "rptMap_Received"
16880         .opgGTRMulti_Rec = .opgGTRMulti
16890       Case "rptMap_Split"
16900         .opgGTRMulti_Split = .opgGTRMulti
16910       Case "rptMap_Misc_LTCL"
16920         .opgGTRMulti_LTCL = .opgGTRMulti
16930       Case "rptMap_Misc_STCG"
16940         .opgGTRMulti_STCG = .opgGTRMulti
16950       Case "rptMap_Misc_STCL"
16960         .opgGTRMulti_STCL = .opgGTRMulti
16970       Case "rptMap_Reinvest_Div"
16980         .opgGTRMulti_DivReinvest = .opgGTRMulti
16990       Case "rptMap_Reinvest_Int"
17000         .opgGTRMulti_IntReinvest = .opgGTRMulti
17010       Case "rptMap_Reinvest_Rec"
17020         .opgGTRMulti_RecReinvest = .opgGTRMulti
17030       End Select

17040     End If

17050   End With

EXITP:
17060   Exit Sub

ERRH:
17070   Select Case ERR.Number
        Case Else
17080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17090   End Select
17100   Resume EXITP

End Sub

Private Sub opgGTRMulti_opt1_cmd_Click()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGTRMulti_opt1_cmd_Click"

17210   With Me
17220     .opgGTRMulti = .opgGTRMulti_opt1.OptionValue
17230     opgGTRMulti_AfterUpdate  ' ** Procedure: Above.
17240   End With

EXITP:
17250   Exit Sub

ERRH:
17260   Select Case ERR.Number
        Case Else
17270     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17280   End Select
17290   Resume EXITP

End Sub

Private Sub opgGTRMulti_opt1_cmd_DblClick(Cancel As Integer)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGTRMulti_opt1_cmd_DblClick"

17310   cmdLocationGo_Click  ' ** Procedure: Below.

EXITP:
17320   Exit Sub

ERRH:
17330   Select Case ERR.Number
        Case Else
17340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17350   End Select
17360   Resume EXITP

End Sub

Private Sub opgGTRMulti_opt2_cmd_Click()

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGTRMulti_opt2_cmd_Click"

17410   With Me
17420     .opgGTRMulti = .opgGTRMulti_opt2.OptionValue
17430     opgGTRMulti_AfterUpdate  ' ** Procedure: Above.
17440   End With

EXITP:
17450   Exit Sub

ERRH:
17460   Select Case ERR.Number
        Case Else
17470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17480   End Select
17490   Resume EXITP

End Sub

Private Sub opgGTRMulti_opt2_cmd_DblClick(Cancel As Integer)

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGTRMulti_opt2_cmd_DblClick"

17510   cmdLocationGo_Click  ' ** Procedure: Below.

EXITP:
17520   Exit Sub

ERRH:
17530   Select Case ERR.Number
        Case Else
17540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17550   End Select
17560   Resume EXITP

End Sub

Private Sub opgGTRMulti_opt3_cmd_Click()

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGTRMulti_opt3_cmd_Click"

17610   With Me
17620     .opgGTRMulti = .opgGTRMulti_opt3.OptionValue
17630     opgGTRMulti_AfterUpdate  ' ** Procedure: Above.
17640   End With

EXITP:
17650   Exit Sub

ERRH:
17660   Select Case ERR.Number
        Case Else
17670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17680   End Select
17690   Resume EXITP

End Sub

Private Sub opgGTRMulti_opt3_cmd_DblClick(Cancel As Integer)

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGTRMulti_opt3_cmd_DblClick"

17710   cmdLocationGo_Click  ' ** Procedure: Below.

EXITP:
17720   Exit Sub

ERRH:
17730   Select Case ERR.Number
        Case Else
17740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17750   End Select
17760   Resume EXITP

End Sub

Private Sub opgGTRMulti_opt4_cmd_Click()

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGTRMulti_opt4_cmd_Click"

17810   With Me
17820     .opgGTRMulti = .opgGTRMulti_opt4.OptionValue
17830     opgGTRMulti_AfterUpdate  ' ** Procedure: Above.
17840   End With

EXITP:
17850   Exit Sub

ERRH:
17860   Select Case ERR.Number
        Case Else
17870     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17880   End Select
17890   Resume EXITP

End Sub

Private Sub opgGTRMulti_opt4_cmd_DblClick(Cancel As Integer)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgGTRMulti_opt4_cmd_DblClick"

17910   cmdLocationGo_Click  ' ** Procedure: Below.

EXITP:
17920   Exit Sub

ERRH:
17930   Select Case ERR.Number
        Case Else
17940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17950   End Select
17960   Resume EXITP

End Sub

Private Sub cmdLocationGo_Click()
' ** Right now, qualification is with lbxReports_AfterUpdate().

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLocationGo_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long, lngMultiOpt As Long
        Dim lngX As Long

18010   With Me

18020     If blnNoGo = False Then

18030       THAT_PROC = THIS_PROC
18040       That_Erl = 17940  'WHERE WAS THIS SUPPOSED TO BE?
18050       DoCmd.Hourglass True
18060       DoEvents

18070       If blnGTR_Emblem = False Then

18080         gblnGoToReport = True
18090         blnGTR_Emblem = True

18100         That_Erl = 18000
18110         garr_varGoToReport(GTR_RLRID) = .lbxReports.Column(LBX_RPT_RLRID)
18120         garr_varGoToReport(GTR_RID) = .lbxReports.Column(LBX_RPT_RID)
18130         garr_varGoToReport(GTR_RNAM) = .lbxReports.Column(LBX_RPT_RNAM)
18140         garr_varGoToReport(GTR_RCAP) = .lbxReports.Column(LBX_RPT_RCAP)
18150         garr_varGoToReport(GTR_ORD) = .lbxReports.Column(LBX_RPT_RORD)

18160         .TimerInterval = 50&

18170       Else

18180         That_Erl = 18080
18190         lngMultiOpt = .opgGTRMulti

18200         Set dbs = CurrentDb
18210         With dbs
                ' ** qryReport_List_89 (xx), by specified [rlrid].
18220           Set qdf = .QueryDefs("qryReport_List_89a")
18230           With qdf.Parameters
18240             ![rlrid] = garr_varGoToReport(GTR_RLRID)
18250           End With
18260           Set rst = qdf.OpenRecordset
18270           With rst
18280             If .BOF = True And .EOF = True Then
18290               gblnGoToReport = False
18300             Else
18310               .MoveLast
18320               lngRecs = .RecordCount
18330               .MoveFirst
                    ' ******************************************************
                    ' ** Array: garr_varGoToReport()
                    ' **
                    ' **   Field  Element  Name                Constant
                    ' **   =====  =======  ==================  ===========
                    ' **     1       0     rptlistrpt_id       GTR_RLRID
                    ' **     2       1     rpt_id              GTR_RID
                    ' **     3       2     rpt_name            GTR_RNAM
                    ' **     4       3     rpt_caption         GTR_RCAP
                    ' **     5       4     rptlistrpt_order    GTR_ORD
                    ' **     6       5     frm_name1           GTR_FRM1
                    ' **     7       6     ctl_name1           GTR_CTL1
                    ' **     8       7     ctl_invoke1         GTR_INV1
                    ' **     9       8     ctl_focus1          GTR_FOC1
                    ' **    10       9     frm_name2           GTR_FRM2
                    ' **    11      10     ctl_name2           GTR_CTL2
                    ' **    12      11     ctl_invoke2         GTR_INV2
                    ' **    13      12     ctl_focus2          GTR_FOC2
                    ' **    14      13     frm_name3           GTR_FRM3
                    ' **    15      14     ctl_name3           GTR_CTL3
                    ' **    16      15     ctl_invoke3         GTR_INV3
                    ' **    17      16     ctl_focus3          GTR_FOC3
                    ' **    18      17     frm_name4           GTR_FRM4
                    ' **    19      18     ctl_name4           GTR_CTL4
                    ' **    20      19     ctl_invoke4         GTR_INV4
                    ' **    21      20     ctl_focus4          GTR_FOC4
                    ' **    22      21     frm_name5           GTR_FRM5
                    ' **    23      22     ctl_name5           GTR_CTL5
                    ' **    24      23     ctl_invoke5         GTR_INV5
                    ' **    25      24     ctl_focus5          GTR_FOC5
                    ' **    26      25     frm_name6           GTR_FRM6
                    ' **    27      26     ctl_name6           GTR_CTL6
                    ' **    28      27     ctl_invoke6         GTR_INV6
                    ' **    29      28     ctl_focus6          GTR_FOC6
                    ' **
                    ' ******************************************************
18340               If lngRecs > 1& Then
                      ' ** Since the query wants to come up in a random order, we've got to check the choice.
18350                 If (![frm_name2] = "frmJournal" Or ![frm_name2] = "frmJournal_Columns") And _
                          lngRecs = 2& And Left(![rpt_name], 6) = "rptMap" Then
                        ' ** The options are always frmJournal 1st and frmJournal_Columns 2nd.
18360                   Select Case lngMultiOpt
                        Case 1
18370                     If ![frm_name2] = "frmJournal" Then
                            ' ** That's correct, so do nothing
18380                     ElseIf ![frm_name2] = "frmJournal_Columns" Then
                            ' ** Move to the other one.
18390                       .MoveNext
18400                     End If
18410                   Case 2
18420                     If ![frm_name2] = "frmJournal_Columns" Then
                            ' ** That's correct, so do nothing
18430                     ElseIf ![frm_name2] = "frmJournal" Then
                            ' ** Move to the other one.
18440                       .MoveNext
18450                     End If
18460                   End Select
18470                 Else
18480                   For lngX = 2& To lngMultiOpt
18490                     .MoveNext
18500                   Next
18510                 End If
18520               End If
18530               garr_varGoToReport(GTR_FRM1) = ![frm_name1]
18540               garr_varGoToReport(GTR_CTL1) = ![ctl_name1]
18550               garr_varGoToReport(GTR_INV1) = ![ctl_invoke1]
18560               garr_varGoToReport(GTR_FOC1) = ![ctl_focus1]
18570               garr_varGoToReport(GTR_FRM2) = ![frm_name2]
18580               garr_varGoToReport(GTR_CTL2) = ![ctl_name2]
18590               garr_varGoToReport(GTR_INV2) = ![ctl_invoke2]
18600               garr_varGoToReport(GTR_FOC2) = ![ctl_focus2]
18610               garr_varGoToReport(GTR_FRM3) = ![frm_name3]
18620               garr_varGoToReport(GTR_CTL3) = ![ctl_name3]
18630               garr_varGoToReport(GTR_INV3) = ![ctl_invoke3]
18640               garr_varGoToReport(GTR_FOC3) = ![ctl_focus3]
18650               garr_varGoToReport(GTR_FRM4) = ![frm_name4]
18660               garr_varGoToReport(GTR_CTL4) = ![ctl_name4]
18670               garr_varGoToReport(GTR_INV4) = ![ctl_invoke4]
18680               garr_varGoToReport(GTR_FOC4) = ![ctl_focus4]
18690               garr_varGoToReport(GTR_FRM5) = ![frm_name5]
18700               garr_varGoToReport(GTR_CTL5) = ![ctl_name5]
18710               garr_varGoToReport(GTR_INV5) = ![ctl_invoke5]
18720               garr_varGoToReport(GTR_FOC5) = ![ctl_focus5]
18730               garr_varGoToReport(GTR_FRM6) = ![frm_name6]
18740               garr_varGoToReport(GTR_CTL6) = ![ctl_name6]
18750               garr_varGoToReport(GTR_INV6) = ![ctl_invoke6]
18760               garr_varGoToReport(GTR_FOC6) = ![ctl_focus6]
18770             End If
18780             .Close
18790           End With
18800           .Close
18810         End With

18820         If garr_varGoToReport(GTR_RNAM) = "rptVersion_Convert" Then
18830           blnGoingToReport = True
18840           .GoToReport_arw_rpt_img.Visible = True
18850           .cmdConvLog.SetFocus
18860           DoEvents
18870           blnConvLog = True
18880           .TimerInterval = 50&
18890         Else

18900           Select Case gblnGoToReport
                Case True
18910             .TimerInterval = 50&
18920           Case False
18930             blnGTR_Emblem = False
18940             GTREmblem_Off  ' ** Procedure: Below.
18950             DoCmd.Hourglass False
18960           End Select

18970         End If

18980       End If  ' ** blnGTR_Emblem.

18990     End If  ' ** blnNoGo.

19000   End With

EXITP:
19010   Set rst = Nothing
19020   Set qdf = Nothing
19030   Set dbs = Nothing
19040   Exit Sub

ERRH:
19050   DoCmd.Hourglass False
19060   Select Case ERR.Number
        Case Else
19070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19080   End Select
19090   Resume EXITP

End Sub

Private Sub cmdLocationGo_KeyDown(KeyCode As Integer, Shift As Integer)

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLocationGo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19120   intShiftDown = (Shift And acShiftMask) > 0
19130   intAltDown = (Shift And acAltMask) > 0
19140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19160     Select Case intRetVal
          Case vbKeyTab
19170       With Me
19180         intRetVal = 0
19190         If .cmdConvLog.Visible = True And .cmdConvLog.Enabled = True Then
19200           .cmdConvLog.SetFocus
19210         Else
19220           .cmdClose.SetFocus
19230         End If
19240       End With
19250     End Select
19260   End If

        ' ** Shift keys.
19270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19280     Select Case intRetVal
          Case vbKeyTab
19290       With Me
19300         intRetVal = 0
19310         .lbxReports.SetFocus
19320       End With
19330     End Select
19340   End If

EXITP:
19350   KeyCode = intRetVal
19360   Exit Sub

ERRH:
19370   intRetVal = 0
19380   Select Case ERR.Number
        Case Else
19390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19400   End Select
19410   Resume EXITP

End Sub

Private Sub cmdConvLog_Click()

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdConvLog_Click"

        Dim strDocName As String

19510   With Me
19520     If ConvLogChk = True Then  ' ** Function: Below.

19530       gblnSetFocus = True
19540       strDocName = "frmVersion_Log"
19550       If strLogPathFile = vbNullString Then strLogPathFile = "{Not Found}"
19560       DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & strLogPathFile

19570       If gblnGoToReport = True And blnGoingToReport = True Then
19580         .TimerInterval = 0&
19590         Forms(strDocName).TimerInterval = 50&
19600         blnGoingToReport = False
19610         blnGTR_Emblem = False
19620         .GoToReport_arw_rpt_img.Visible = False
19630         GTREmblem_Off  ' ** Procedure: Below.
19640       End If

19650     Else
19660       Beep
19670     End If
19680   End With

EXITP:
19690   Exit Sub

ERRH:
19700   Select Case ERR.Number
        Case Else
19710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19720   End Select
19730   Resume EXITP

End Sub

Private Sub cmdConvLog_KeyDown(KeyCode As Integer, Shift As Integer)

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdConvLog_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19820   intShiftDown = (Shift And acShiftMask) > 0
19830   intAltDown = (Shift And acAltMask) > 0
19840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19860     Select Case intRetVal
          Case vbKeyTab
19870       With Me
19880         intRetVal = 0
19890         .cmdClose.SetFocus
19900       End With
19910     End Select
19920   End If

        ' ** Shift keys.
19930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19940     Select Case intRetVal
          Case vbKeyTab
19950       With Me
19960         intRetVal = 0
19970         If .cmdLocationGo.Enabled = True Then
19980           .cmdLocationGo.SetFocus
19990         Else
20000           .lbxReports.SetFocus
20010         End If
20020       End With
20030     End Select
20040   End If

EXITP:
20050   KeyCode = intRetVal
20060   Exit Sub

ERRH:
20070   intRetVal = 0
20080   Select Case ERR.Number
        Case Else
20090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20100   End Select
20110   Resume EXITP

End Sub

Private Function RptArray_Load() As Boolean

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "RptArray_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim blnLoadStage As Boolean
        Dim lngThisDbsID As Long
        Dim varTmp00 As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

20210   blnRetVal = True
20220   blnLoadStage = False

20230   DoCmd.Hourglass True
20240   DoEvents

20250   lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

20260   If lngRpts = 0& Or IsEmpty(arr_varRpt) = True Then

20270     varTmp00 = DCount("*", "tblReport_List_Staging")
20280     Select Case IsNull(varTmp00)
          Case True
20290       blnLoadStage = True
20300     Case False
20310       If varTmp00 = 0 Then
20320         blnLoadStage = True
20330       End If
20340     End Select

20350     If blnLoadStage = False Then
20360       Set dbs = CurrentDb
20370       With dbs
              ' ** tblReport_List_Report, sorted by rpt_caption, with rptlistrpt_datemodified.
20380         Set qdf = .QueryDefs("qryReport_List_50a")
20390         Set rst = qdf.OpenRecordset
20400         With rst
20410           If .BOF = True And .EOF = True Then
20420             blnRetVal = False
20430           Else
20440             .MoveLast
20450             lngRecs = .RecordCount
20460             .MoveFirst
20470             If ![rptlistrpt_order] <> 1& Then
20480               blnLoadStage = True
20490             Else
20500               .MoveLast
20510               If ![rptlistrpt_order] <> lngRecs Then
20520                 blnLoadStage = True
20530               End If
20540             End If
20550           End If
20560         End With  ' ** rst.
20570         Set rst = Nothing
20580         Set qdf = Nothing
20590         .Close
20600       End With  ' ** dbs.

20610       If blnLoadStage = True Then
20620         ReportList_Order  ' ** Module Function: modReportFunctions.
20630         Set dbs = CurrentDb
20640         With dbs
                ' ** Empty tblReport_List_Staging, by specified [dbid].
20650           Set qdf = .QueryDefs("qryReport_List_10a")
20660           With qdf.Parameters
20670             ![dbid] = lngThisDbsID
20680           End With
20690           qdf.Execute
20700           Set qdf = Nothing
20710           .Close
20720         End With  ' ** dbs.
20730         DoCmd.Hourglass True
20740         DoEvents
20750         ChangeSeed_Ext "tblReport_List_Staging"  ' ** Module Function: modAutonumberFieldFuncs.
20760         DoCmd.Hourglass True
20770         DoEvents
20780       End If
20790     End If  ' ** blnLoadStage.

20800     Set dbs = CurrentDb
20810     With dbs

20820       If blnLoadStage = True Then
              ' ** Append qryReport_List_88 (Union of qryReport_List_82 (xx), qryReport_List_83 (xx),
              ' ** qryReport_List_84 (xx), qryReport_List_85 (xx),
              ' ** qryReport_List_86 (xx), qryReport_List_87 (xx))
              ' ** to tblReport_List_Staging.
20830         Set qdf = .QueryDefs("qryReport_List_88a")
20840         qdf.Execute
20850         Set qdf = Nothing
20860       End If

20870       If blnRetVal = True Then
              ' ** qryReport_List_89c, just rptlistsort_id = Null.
20880         Set qdf = .QueryDefs("qryReport_List_89ca")
20890         Set rst = qdf.OpenRecordset
20900         If rst.BOF = True And rst.EOF = True Then
                ' ** All's well.
20910           rst.Close
20920           Set rst = Nothing
20930           Set qdf = Nothing
20940         Else
20950           rst.Close
20960           Set rst = Nothing
20970           Set qdf = Nothing
                ' ** Empty tblReport_List_Sort.
20980           Set qdf = .QueryDefs("qryReport_List_89cb")
20990           qdf.Execute
21000           Set qdf = Nothing
                ' ** Append qryReport_List_89ci (xx) to tblReport_List_Sort.
21010           Set qdf = .QueryDefs("qryReport_List_89cj")
21020           qdf.Execute
21030           Set qdf = Nothing
                ' **
21040           Set qdf = .QueryDefs("qryReport_List_89ck_04")
21050           qdf.Execute
21060           Set qdf = Nothing
21070         End If
21080       End If

21090       If blnRetVal = True Then

              ' ** qryReport_List_89 (tblReport_List_Staging, linked to tblReport_List_Report,
              ' ** with rpt_id, rpt_name, rpt_caption, rptlistrpt_order), linked to qryReport_List_89b
              ' ** (qryReport_List_89 (tblReport_List_Staging, linked to tblReport_List_Report,
              ' ** with rpt_id, rpt_name, rpt_caption, rptlistrpt_order), grouped by rptlistrpt_id, with cnt).
21100         Set qdf = .QueryDefs("qryReport_List_89c")
21110         Set rst = qdf.OpenRecordset
21120         With rst
21130           If .BOF = True And .EOF = True Then
21140             blnRetVal = False
21150           Else
21160             .MoveLast
21170             lngRpts = .RecordCount
21180             .MoveFirst
21190             arr_varRpt = .GetRows(lngRpts)
                  ' ********************************************************
                  ' ** Array: arr_varRpt()
                  ' **
                  ' **   Field  Element  Name                   Constant
                  ' **   =====  =======  =====================  ==========
                  ' **     1       0     rptlistrpt_id          R_RLRID
                  ' **     2       1     rptlistrptalt_id       R_RLRAID
                  ' **     3       2     rptlistrptaltctl_id    R_RLRACID
                  ' **     4       3     rptlist_id             R_RLID
                  ' **     5       4     dbs_id                 R_DID
                  ' **     6       5     rpt_id                 R_RID
                  ' **     7       6     rpt_name               R_RNAM
                  ' **     8       7     rpt_caption            R_RCAP
                  ' **     9       8     rptlistrpt_order       R_RORD
                  ' **    10       9     rptlistctl_id1         R_RLCID1
                  ' **    11      10     rptlistctltype_type1   R_TYP1
                  ' **    12      11     rptlistctl_level1      R_LVL1
                  ' **    13      12     rptlistctl_alt1        R_ALT1
                  ' **    14      13     frm_id1                R_FID1
                  ' **    15      14     frm_name1              R_FNAM1
                  ' **    16      15     frm_caption1           R_FCAP1
                  ' **    17      16     ctl_id1                R_CID1
                  ' **    18      17     ctl_name1              R_CNAM1
                  ' **    19      18     ctl_caption1           R_CCAP1
                  ' **    20      19     ctl_invoke1            R_INV1
                  ' **    21      20     ctl_focus1             R_FOC1
                  ' **    22      21     rptlistctl_id2         R_RLCID2
                  ' **    23      22     rptlistctltype_type2   R_TYP2
                  ' **    24      23     rptlistctl_level2      R_LVL2
                  ' **    25      24     rptlistctl_alt2        R_ALT2
                  ' **    26      25     frm_id2                R_FID2
                  ' **    27      26     frm_name2              R_FNAM2
                  ' **    28      27     frm_caption2           R_FCAP2
                  ' **    29      28     ctl_id2                R_CID2
                  ' **    30      29     ctl_name2              R_CNAM2
                  ' **    31      30     ctl_caption2           R_CCAP2
                  ' **    32      31     ctl_invoke2            R_INV2
                  ' **    33      32     ctl_focus2             R_FOC2
                  ' **    34      33     rptlistctl_id3         R_RLCID3
                  ' **    35      34     rptlistctltype_type3   R_TYP3
                  ' **    36      35     rptlistctl_level3      R_LVL3
                  ' **    37      36     rptlistctl_alt3        R_ALT3
                  ' **    38      37     frm_id3                R_FID3
                  ' **    39      38     frm_name3              R_FNAM3
                  ' **    40      39     frm_caption3           R_FCAP3
                  ' **    41      40     ctl_id3                R_CID3
                  ' **    42      41     ctl_name3              R_CNAM3
                  ' **    43      42     ctl_caption3           R_CCAP3
                  ' **    44      43     ctl_invoke3            R_INV3
                  ' **    45      44     ctl_focus3             R_FOC3
                  ' **    46      45     rptlistctl_id4         R_RLCID4
                  ' **    47      46     rptlistctltype_type4   R_TYP4
                  ' **    48      47     rptlistctl_level4      R_LVL4
                  ' **    49      48     rptlistctl_alt4        R_ALT4
                  ' **    50      49     frm_id4                R_FID4
                  ' **    51      50     frm_name4              R_FNAM4
                  ' **    52      51     frm_caption4           R_FCAP4
                  ' **    53      52     ctl_id4                R_CID4
                  ' **    54      53     ctl_name4              R_CNAM4
                  ' **    55      54     ctl_caption4           R_CCAP4
                  ' **    56      55     ctl_invoke4            R_INV4
                  ' **    57      56     ctl_focus4             R_FOC4
                  ' **    58      57     rptlistctl_id5         R_RLCID5
                  ' **    59      58     rptlistctltype_type5   R_TYP5
                  ' **    60      59     rptlistctl_level5      R_LVL5
                  ' **    61      60     rptlistctl_alt5        R_ALT5
                  ' **    62      61     frm_id5                R_FID5
                  ' **    63      62     frm_name5              R_FNAM5
                  ' **    64      63     frm_caption5           R_FCAP5
                  ' **    65      64     ctl_id5                R_CID5
                  ' **    66      65     ctl_name5              R_CNAM5
                  ' **    67      66     ctl_caption5           R_CCAP5
                  ' **    68      67     ctl_invoke5            R_INV5
                  ' **    69      68     ctl_focus5             R_FOC5
                  ' **    70      69     rptlistctl_id6         R_RLCID6
                  ' **    71      70     rptlistctltype_type6   R_TYP6
                  ' **    72      71     rptlistctl_level6      R_LVL6
                  ' **    73      72     rptlistctl_alt6        R_ALT6
                  ' **    74      73     frm_id6                R_FID6
                  ' **    75      74     frm_name6              R_FNAM6
                  ' **    76      75     frm_caption6           R_FCAP6
                  ' **    77      76     ctl_id6                R_CID6
                  ' **    78      77     ctl_name6              R_CNAM6
                  ' **    79      78     ctl_caption6           R_CCAP6
                  ' **    80      79     ctl_invoke6            R_INV6
                  ' **    81      80     ctl_focus6             R_FOC6
                  ' **    82      81     cnt                    R_CNT
                  ' **    83      82     rptlistsort_id         R_SID
                  ' **
                  ' ********************************************************
21200           End If
21210           .Close
21220         End With
21230         Set rst = Nothing
21240         Set qdf = Nothing

              'For lngX = 0& To (lngRpts - 1&)
              '  Debug.Print "'" & Left(CStr(lngX + 1&) & "    ", 4) & arr_varRpt(R_RNAM, lngX) & "  " & arr_varRpt(R_FNAM1, lngX)
              'Next

'1   rptAccountBalance  frmMenu_Report
'2   rptAccountContacts  frmMenu_Utility
'3   rptAccountProfile_SepChkAcct  frmMenu_Report
'4   rptAccountProfile_SepChkAcct  frmMenu_Report
'5   rptAccountProfile_SepChkAcct  frmMenu_Account
'6   rptAccountProfile_SepChkAcct  frmMenu_Account
'7   rptAccountReviews  frmMenu_Report
'8   rptAccountSummary  frmMenu_Report
'9   rptArchivedTransactions  frmMenu_Utility
'10  rptArchivedTransactions  frmMenu_Utility
'11  rptAssetHistory  frmMenu_Asset
'12  rptAssetHistory_All  frmMenu_Asset
'13  rptAssetList  frmMenu_Report
'14  rptAssetPricing_History  frmMenu_Asset
'15  rptAssetPricing  frmMenu_Asset
'16  rptPricingUpdateCusips_prn  frmMenu_Asset
'17  rptPurchasedSold_Purchased  frmMenu_Report
'18  rptPurchasedSold_Sold  frmMenu_Report
'19  rptTransaction_Audit_01  frmMenu_Utility
'20  rptCapitalGainAndLoss  frmMenu_Report
'21  rptCashControl  frmMenu_Report
'22  rptCheckList_All  frmMenu_Post
'23  rptCheckReconcile  frmMenu_Post
'24  rptChecks_Blank  frmMenu_Post
'25  rptChecks_Blank  frmMenu_Post
'26  rptChecks_Preprinted  frmMenu_Post
'27  rptNewClosedAccounts_Closed  frmMenu_Report
'28  rptVersion_Convert  frmMenu_Title
'29  rptVersion_Convert  frmMenu_Report
'30  rptCourtRptCA_01  frmMenu_Report
'31  rptCourtRptFL_00B  frmMenu_Report
'32  rptCourtRptFL_07  frmMenu_Report
'33  rptCourtRptFL_03  frmMenu_Report
'34  rptCourtRptCA_09  frmMenu_Report
'35  rptCourtRptNS_06  frmMenu_Report
'36  rptCourtRptFL_04  frmMenu_Report
'37  rptCourtRptCA_04  frmMenu_Report
'38  rptCourtRptFL_04B  frmMenu_Report
'39  rptCourtRptNS_08  frmMenu_Report
'40  rptCourtRptNS_03  frmMenu_Report
'41  rptCourtRptNS_04  frmMenu_Report
'42  rptCourtRptFL_06  frmMenu_Report
'43  rptCourtRptCA_06  frmMenu_Report
'44  rptCourtRptNS_09  frmMenu_Report
'45  rptCourtRptNS_02  frmMenu_Report
'46  rptCourtRptCA_03  frmMenu_Report
'47  rptCourtRptCA_08  frmMenu_Report
'48  rptCourtRptNS_05  frmMenu_Report
'49  rptCourtRptCA_05  frmMenu_Report
'50  rptCourtRptCA_10  frmMenu_Report
'51  rptCourtRptCA_11  frmMenu_Report
'52  rptCourtRptCA_00B  frmMenu_Report
'53  rptCourtRptNS_00B  frmMenu_Report
'54  rptCourtRptCA_07  frmMenu_Report
'55  rptCourtRptNS_00D  frmMenu_Report
'56  rptCourtRptCA_02  frmMenu_Report
'57  rptCourtRptFL_02  frmMenu_Report
'58  rptCourtRptNS_07  frmMenu_Report
'59  rptCourtRptNS_01  frmMenu_Report
'60  rptCourtRptCA_00  frmMenu_Report
'61  rptCourtRptFL_00  frmMenu_Report
'62  rptCourtRptNS_00  frmMenu_Report
'63  rptTax_DisbursementsDeductions  frmMenu_Report
'64  rptErrorLog  frmMenu_Utility
'65  rptIncExp_Expenses  frmMenu_Report
'66  rptIncExp_Expenses_Detailed  frmMenu_Report
'67  rptIncExp_Expenses_Summary  frmMenu_Report
'68  rptFeeCalculations  frmMenu_Post
'69  rptTransaction_Hidden  frmMenu_Account
'70  rptTransaction_Hidden  frmMenu_Account
'71  rptHoldings  frmMenu_Asset
'72  rptHoldings_Detailed  frmMenu_Asset
'73  rptIncExp_Income  frmMenu_Report
'74  rptIncExp_Income_Detailed  frmMenu_Report
'75  rptIncExp_Income_Summary  frmMenu_Report
'76  rptIncomeStatement  frmMenu_Report
'77  rptTax_ReceiptsIncome  frmMenu_Report
'78  rptListOfAccountTypeGroups  frmMenu_Utility
'79  rptListOfAccountTypeGroups_Detailed  frmMenu_Utility
'80  rptListOfAccountTypes  frmMenu_Utility
'81  rptListOfAdminOfficers  frmMenu_Utility
'82  rptListOfAdminOfficers_Detailed  frmMenu_Utility
'83  rptListOfTaxLots  frmMenu_Post
'84  rptListOfTaxLots  frmMenu_Post
'85  rptListOfAssetTypeGroups  frmMenu_Utility
'86  rptListOfAssetTypeGroups_Detailed  frmMenu_Utility
'87  rptListOfAssetTypes  frmMenu_Utility
'88  rptListOfAssets  frmMenu_Asset
'89  rptListOfFeeSchedules  frmMenu_Utility
'90  rptListOfIncExpCodes  frmMenu_Utility
'91  rptListOfInvestmentObjectives  frmMenu_Utility
'92  rptListOfInvestmentObjectives_Detailed  frmMenu_Utility
'93  rptListOfJournalTypes  frmMenu_Utility
'94  rptListOfLocations  frmMenu_Utility
'95  rptListOfRecurringItems_01  frmMenu_Utility
'96  rptListOfRelatedAccounts  frmMenu_Account
'97  rptListOfStateCodes  frmMenu_Utility
'98  rptListOfTaxCodes  frmMenu_Utility
'99  rptListOfUsers  frmMenu_Utility
'100 rptLocations  frmMenu_Asset
'101 rptLocations_Detailed  frmMenu_Asset
'102 rptMap_Dividend  frmMenu_Post
'103 rptMap_Dividend  frmMenu_Post
'104 rptMap_Reinvest_Div  frmMenu_Post
'105 rptMap_Reinvest_Div  frmMenu_Post
'106 rptMap_Interest  frmMenu_Post
'107 rptMap_Interest  frmMenu_Post
'108 rptMap_Reinvest_Int  frmMenu_Post
'109 rptMap_Reinvest_Int  frmMenu_Post
'110 rptMap_Received  frmMenu_Post
'111 rptMap_Received  frmMenu_Post
'112 rptMap_Reinvest_Rec  frmMenu_Post
'113 rptMap_Reinvest_Rec  frmMenu_Post
'114 rptMap_Split  frmMenu_Post
'115 rptMap_Split  frmMenu_Post
'116 rptMasterBalance  frmMenu_Utility
'117 rptNewClosedAccounts_New  frmMenu_Report
'118 rptPortfolioModeling  frmMenu_Asset
'119 rptPostingJournal_Classic  frmMenu_Post
'120 rptPostingJournal_Column  frmMenu_Post
'121 rptMaturity  frmMenu_Report
'122 rptStatement_Qualified  frmMenu_Report
'123 rptStatement_Balance  frmMenu_Report
'124 rptStatementOfCondition  frmMenu_Report
'125 rptTaxLot_All_SortDesc  frmMenu_Report
'126 rptTransaction_Statement_SortType  frmMenu_Report
'127 rptTransactions_ByType  frmMenu_Report
'128 rptUnrealizedGainAndLoss  frmMenu_Report

21250         Set rst = .OpenRecordset("tblReportListControlType", dbOpenDynaset, dbReadOnly)
21260         With rst
21270           .MoveLast
21280           lngRecs = .RecordCount
21290           .MoveFirst
21300           For lngX = 1& To lngRecs
21310             Select Case ![rptlistctltype_constant]
                  Case "taReport"
21320               taReport = ![rptlistctltype_type]
21330             Case "taMenu1"
21340               taMenu1 = ![rptlistctltype_type]
21350             Case "taMenu2"
21360               taMenu2 = ![rptlistctltype_type]
21370             Case "taMenu3"
21380               taMenu3 = ![rptlistctltype_type]
21390             Case "taVia1"
21400               taVia1 = ![rptlistctltype_type]
21410             Case "taVia2"
21420               taVia2 = ![rptlistctltype_type]
21430             Case "taVia3"
21440               taVia3 = ![rptlistctltype_type]
21450             Case "taVia4"
21460               taVia4 = ![rptlistctltype_type]
21470             End Select
21480             If lngX < lngRecs Then .MoveNext
21490           Next  ' ** lngX.
21500           .Close
21510         End With  ' ** rst.

21520       End If  ' ** blnRetVal.

21530       .Close
21540     End With  ' ** dbs.
21550     Set dbs = Nothing
21560   End If

21570   DoCmd.Hourglass False

EXITP:
21580   Set rst = Nothing
21590   Set qdf = Nothing
21600   Set dbs = Nothing
21610   RptArray_Load = blnRetVal
21620   Exit Function

ERRH:
21630   blnRetVal = False
21640   DoCmd.Hourglass False
21650   Select Case ERR.Number
        Case Else
21660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21670   End Select
21680   Resume EXITP

End Function

Private Function IsMenu(varType As Variant) As Boolean

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "IsMenu"

        Dim blnRetVal As Boolean

21710   blnRetVal = False

21720   If varType = taMenu1 Or varType = taMenu2 Or varType = taMenu3 Then
21730     blnRetVal = True
21740   End If

EXITP:
21750   IsMenu = blnRetVal
21760   Exit Function

ERRH:
21770   blnRetVal = False
21780   Select Case ERR.Number
        Case Else
21790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21800   End Select
21810   Resume EXITP

End Function

Private Function IsVia(varType As Variant) As Boolean

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "IsVia"

        Dim blnRetVal As Boolean

21910   blnRetVal = False

21920   If varType = taVia1 Or varType = taVia2 Or varType = taVia3 Or varType = taVia4 Then
21930     blnRetVal = True
21940   End If

EXITP:
21950   IsVia = blnRetVal
21960   Exit Function

ERRH:
21970   blnRetVal = False
21980   Select Case ERR.Number
        Case Else
21990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22000   End Select
22010   Resume EXITP

End Function

Private Function IsRpt(varType As Variant) As Boolean

22100 On Error GoTo ERRH

        Const THIS_PROC As String = "IsRpt"

        Dim blnRetVal As Boolean

22110   blnRetVal = False

22120   If varType = taReport Then
22130     blnRetVal = True
22140   End If

EXITP:
22150   IsRpt = blnRetVal
22160   Exit Function

ERRH:
22170   blnRetVal = False
22180   Select Case ERR.Number
        Case Else
22190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22200   End Select
22210   Resume EXITP

End Function

Private Function ConvLogChk() As Boolean

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "ConvLogChk"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strPath As String, strFile As String, strPathFile As String
        Dim lngRecs As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

22310   blnRetVal = False

22320   strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
22330   If strPath <> gstrDir_Dev Then
22340     strPath = CurrentBackendPath  ' ** Module Function: modFileUtilities.
22350   End If
22360   strPath = strPath & LNK_SEP & gstrDir_Convert
22370   strFile = gstrFile_ConvertLog  ' ** ConvLog.txt.
22380   strPathFile = strPath & LNK_SEP & strFile
22390   blnRetVal = FileExists(strPathFile)  ' ** Module Function: modFileUtilities.

22400   Select Case blnRetVal
        Case True
22410     strLogPathFile = strPathFile
22420   Case False
22430     Set dbs = CurrentDb
22440     With dbs
22450       Set rst = .OpenRecordset("tblVersion_Conversion", dbOpenDynaset, dbReadOnly)
22460       With rst
22470         If .BOF = True And .EOF = True Then
                ' ** Get the file.
22480         Else
22490           .MoveLast
22500           lngRecs = .RecordCount
22510           .MoveFirst
22520           For lngX = 1& To lngRecs  ' ** This may only ever have 1 record!
22530             If IsNull(![vercnv_name]) = False Then
22540               If IsNull(![vercnv_note]) = False Then
22550                 If Left(![vercnv_note], 27) = "Trust Accountant Conversion" Then
22560                   blnRetVal = True
22570                   strLogPathFile = "{From Table}"
22580                   Exit For
22590                 End If
22600               End If
22610             End If
22620             If lngX < lngRecs Then .MoveNext
22630           Next
22640         End If
22650         .Close
22660       End With
22670       Set rst = Nothing
22680       .Close
22690     End With
22700     Set dbs = Nothing
22710   End Select

EXITP:
22720   Set rst = Nothing
22730   Set dbs = Nothing
22740   ConvLogChk = blnRetVal
22750   Exit Function

ERRH:
22760   blnRetVal = False
22770   Select Case ERR.Number
        Case Else
22780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22790   End Select
22800   Resume EXITP

End Function

Public Sub GTREmblem_Off()

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

22910   With Me
22920     blnGTR_Emblem = False
22930     For lngX = 1& To 24&
22940       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
22950     Next
22960   End With

EXITP:
22970   Exit Sub

ERRH:
22980   Select Case ERR.Number
        Case Else
22990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23000   End Select
23010   Resume EXITP

End Sub
