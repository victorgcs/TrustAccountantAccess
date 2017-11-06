VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountExport"

'VGC 10/19/2017: CHANGES!

' #########################
' ## Use VBA_RenumErrh().  36620
' #########################

' ** Shortcut Alt keys responsive from this form:
' **   All Data:         A {opgExport_optAll}
' **   Browse 1:         B {cmdBrowse1}
' **   Cancel:           C {cmdCancel}
' **   Select None T3:   E {cmdTier3_SelectNone}
' **   File:             F {TargetFile}
' **   Select All T2:    L {cmdTier2_SelectAll}
' **   Select None T1:   N {cmdTier1_SelectNone}
' **   Select None T2:   O {cmdTier2_SelectNone}
' **   Path:             P {TargetPath}
' **   Export Now:       R {cmdExport}
' **   Select All T1:    S {cmdTier1_SelectAll}
' **   Select All T3:    T {cmdTier3_SelectAll}
' **   Acct Data:        U {opgExport_optAcctProf}
' **   Browse 2:         W {cmdBrowse2}
' **   Tier 1:           1 {ckgTier_chkTier1}
' **   Tier 2:           2 {ckgTier_chkTier2}
' **   Tier 3:           3 {ckgTier_chkTier3}

' ***********************************************
' ** Easter Egg (Secret Key):
' ** opgExport_lbl_cmd_MouseUp
' ***********************************************

' ######################################################################
' ## IMPORT/EXPORT SPEC UPDATED: 04/02/2017!
' ######################################################################

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Array: arr_varCtl().
Private lngCtls As Long, arr_varCtl As Variant
Private Const C_DID  As Integer = 0
Private Const C_FID  As Integer = 1
Private Const C_CID  As Integer = 2
Private Const C_TIER As Integer = 3
Private Const C_ITEM As Integer = 4
Private Const C_CNAM As Integer = 5
Private Const C_VAL  As Integer = 6

' ** Array: arr_varRec().
Private lngRecs As Long, arr_varRec As Variant
Private Const R_DID  As Integer = 0
Private Const R_SID  As Integer = 1
Private Const R_SNAM As Integer = 2
Private Const R_TIER As Integer = 3
Private Const R_ITEM As Integer = 4
Private Const R_TNAM As Integer = 5
Private Const R_QNAM As Integer = 6
Private Const R_CNT  As Integer = 7
Private Const R_DEF  As Integer = 8
Private Const R_DCNT As Integer = 9
Private Const R_EXP  As Integer = 10

' ** Array: arr_varItem().
Private lngItems As Long, arr_varItem() As Variant
Private Const I_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const I_CNAM As Integer = 0
Private Const I_TOP  As Integer = 1
Private Const I_VIS  As Integer = 2

Private blnBrowse1_Focus As Boolean, blnBrowse1_MouseDown As Boolean
Private blnBrowse2_Focus As Boolean, blnBrowse2_MouseDown As Boolean
Private blnTier1_SelectAll_Focus As Boolean, blnTier1_SelectAll_MouseDown As Boolean
Private blnTier1_SelectNone_Focus As Boolean, blnTier1_SelectNone_MouseDown As Boolean
Private blnTier2_SelectAll_Focus As Boolean, blnTier2_SelectAll_MouseDown As Boolean
Private blnTier2_SelectNone_Focus As Boolean, blnTier2_SelectNone_MouseDown As Boolean
Private blnTier3_SelectAll_Focus As Boolean, blnTier3_SelectAll_MouseDown As Boolean
Private blnTier3_SelectNone_Focus As Boolean, blnTier3_SelectNone_MouseDown As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngExport_Height As Long, lngExport_Diff As Long, lngExport_Top As Long, lngDetailHLine01_Top As Long

Private strTargetFile As String, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private blnRecalcTots As Boolean, blnEmpty As Boolean, blnNoData As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           blnBrowse1_Focus = False: blnBrowse1_MouseDown = False
170           blnBrowse2_Focus = False: blnBrowse2_MouseDown = False
180           blnTier1_SelectAll_Focus = False: blnTier1_SelectAll_MouseDown = False
190           blnTier1_SelectNone_Focus = False: blnTier1_SelectNone_MouseDown = False
200           blnTier2_SelectAll_Focus = False: blnTier2_SelectAll_MouseDown = False
210           blnTier2_SelectNone_Focus = False: blnTier2_SelectNone_MouseDown = False
220           blnTier3_SelectAll_Focus = False: blnTier3_SelectAll_MouseDown = False
230           blnTier3_SelectNone_Focus = False: blnTier3_SelectNone_MouseDown = False
240           blnRecalcTots = False: blnEmpty = False: blnNoData = False

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
250           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

260           CLR_DISABLED_FG = CLR_DKBLU2
270           CLR_DISABLED_BG = CLR_LTBLU

280           .DefMsg_lbl1.Visible = False
290           .DefMsg_lbl2.Visible = False
300           .detail_height_line.Top = 0&  ' ** Just get it out of the way.

310           lngExport_Height = .opgExport.Height
320           lngExport_Diff = (.Detail_hline01.Top - .AllData_alt_line2.Top)
330           lngDetailHLine01_Top = .Detail_hline01.Top
340           lngExport_Top = .cmdExport.Top

              ' ** Variables are fed empty, then populated ByRef.
350           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

360           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

370           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
380           If lngMonitorCnt > 1& Then
390             EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
400             lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
410             If lngMonitorNum > 1& Then
420               LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
430             End If
440           End If

              ' ** Load arr_varCtl() array.
450           GetCounts 1  ' ** Procedure: Below.

460           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
470           DoEvents

480           If CurrentUser <> "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
                ' ** May be overridden by special key: Ctrl-Alt-RightClick on 'Export Trust Accountant' tab label.
490             If CurrentUser = "TAAdmin" Then
500               Select Case .opgExport_Open
                  Case True
                    ' ** Leave the page open.
510                 blnNoData = False
520               Case False
530                 ShowAllData False  ' ** Procedure: Below.
540                 blnNoData = True
550               End Select
560             Else
570               ShowAllData False  ' ** Procedure: Below.
580               blnNoData = True
590               .opgExport_Open = False
600             End If
610           Else
620             Select Case .opgExport_Open
                Case True
                  ' ** Leave the page open.
630               blnNoData = False
640             Case False
650               ShowAllData False  ' ** Procedure: Below.
660               blnNoData = True
670             End Select
680           End If

690           If IsNull(.TargetPath) = True Then
700             .TargetPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
710           Else
720             If Trim(.TargetPath) = vbNullString Then
730               .TargetPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
740             End If
750           End If

760           TargetPath_AfterUpdate  ' ** Procedure: Below.

770           ckgTier_chkTier1_AfterUpdate  ' ** Procedure: Below.

780           ckgTier1_chk01_Account_AfterUpdate  ' ** Procedure: Below.
790           ckgTier1_chk02_MasterAsset_AfterUpdate  ' ** Procedure: Below.
800           ckgTier1_chk03_ActiveAssets_AfterUpdate  ' ** Procedure: Below.
810           ckgTier1_chk04_LedgerArchive_AfterUpdate  ' ** Procedure: Below.
820           ckgTier1_chk05_Ledger_AfterUpdate  ' ** Procedure: Below.
830           ckgTier1_chk06_Balance_AfterUpdate  ' ** Procedure: Below.

840           ckgTier_chkTier2_AfterUpdate  ' ** Procedure: Below.

850           ckgTier2_chk06_Location_AfterUpdate  ' ** Procedure: Below.
860           ckgTier2_chk05_LedgerHidden_AfterUpdate  ' ** Procedure: Below.
870           ckgTier2_chk04_Currency_History_AfterUpdate  ' ** Procedure: Below.
880           ckgTier2_chk03_Pricing_MasterAsset_History_AfterUpdate  ' ** Procedure: Below.
890           ckgTier2_chk02_RecurringItems_AfterUpdate  ' ** Procedure: Below.
900           ckgTier2_chk01_m_REVCODE_AfterUpdate  ' ** Procedure: Below.

910           ckgTier_chkTier3_AfterUpdate  ' ** Procedure: Below.

920           ckgTier3_chk01_AdminOfficer_AfterUpdate  ' ** Procedure: Below.
930           ckgTier3_chk02_PortfolioModel_AfterUpdate  ' ** Procedure: Below.
940           ckgTier3_chk03_Schedule_AfterUpdate  ' ** Procedure: Below.
950           ckgTier3_chk04_Schedule_Detail_AfterUpdate  ' ** Procedure: Below.
960           ckgTier3_chk05_CheckPOSPay_AfterUpdate  ' ** Procedure: Below.
970           ckgTier3_chk06_CheckPOSPay_Detail_AfterUpdate  ' ** Procedure: Below.
980           ckgTier3_chk07_CheckReconcile_Amount_AfterUpdate  ' ** Procedure: Below.
990           ckgTier3_chk08_CheckReconcile_Item_AfterUpdate  ' ** Procedure: Below.
1000          ckgTier3_chk09_CheckMemo_AfterUpdate  ' ** Procedure: Below.

1010          strTargetFile = "Accounts_" & year(Date) & Format(month(Date), "00") & Format(day(Date), "00") & ".csv"
1020          strTmp01 = strTargetFile
1030          If Right(strTmp01, 4) = ".csv" Then
1040            strTmp01 = Left(strTmp01, (Len(strTmp01) - 4))
1050          End If
1060          .TargetFile = strTmp01

1070          opgExport_AfterUpdate  ' ** Procedure: Below.

1080          DoCmd.Hourglass False

1090          blnIsOpen = False

1100        Else
1110          Cancel = -1
1120        End If
1130      Else
1140        Cancel = -1
1150      End If

1160    End With

1170    If Cancel = -1 Then
1180      Beep
1190      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1200    End If

EXITP:
1210    Exit Sub

ERRH:
1220    DoCmd.Hourglass False
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub Form_Load()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1310    With Me
1320      .AllData_Tbls.ForeColor = CLR_DISABLED_FG
1330      .AllData_Tbls.BackColor = CLR_DISABLED_BG
1340      .AllData_File.ForeColor = CLR_DISABLED_FG
1350      .AllData_File.BackColor = CLR_DISABLED_BG
1360    End With

EXITP:
1370    Exit Sub

ERRH:
1220    DoCmd.Hourglass False
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Tier 1:           1 {ckgTier_chkTier1}
        ' **   Tier 2:           2 {ckgTier_chkTier2}
        ' **   Tier 3:           3 {ckgTier_chkTier3}

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyEscape
1470        intRetVal = 0
1480        cmdCancel_Click  ' ** Procedure: Below.
1490      End Select
1500    End If

        ' ** Alt keys.
1510    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1520      Select Case intRetVal
          Case vbKey1
1530        With Me
1540          intRetVal = 0
1550          If .ckgTier_chkTier1.Enabled = True Then
1560            .ckgTier_chkTier1.SetFocus
1570            .ckgTier_chkTier1 = (Not .ckgTier_chkTier1)
1580            ckgTier_chkTier1_AfterUpdate  ' ** Procedure Below.
1590          Else
1600            Beep
1610          End If
1620        End With
1630      Case vbKey2
1640        With Me
1650          intRetVal = 0
1660          If .ckgTier_chkTier2.Enabled = True Then
1670            .ckgTier_chkTier2.SetFocus
1680            .ckgTier_chkTier2 = (Not .ckgTier_chkTier2)
1690            ckgTier_chkTier2_AfterUpdate  ' ** Procedure Below.
1700          Else
1710            Beep
1720          End If
1730        End With
1740      Case vbKey3
1750        With Me
1760          intRetVal = 0
1770          If .ckgTier_chkTier3.Enabled = True Then
1780            .ckgTier_chkTier3.SetFocus
1790            .ckgTier_chkTier3 = (Not .ckgTier_chkTier3)
1800            ckgTier_chkTier3_AfterUpdate  ' ** Procedure Below.
1810          Else
1820            Beep
1830          End If
1840        End With
1850      End Select
1860    End If

        ' ** Ctrl-Shift keys.
1870    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1880      Select Case intRetVal
          Case vbKeyF
1890        With Me
1900          intRetVal = 0
1910          .FocusHolder.SetFocus
1920        End With
1930      End Select
1940    End If

EXITP:
1950    KeyCode = intRetVal
1960    Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub Form_Timer()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2010    With Me
2020      .TimerInterval = 0&
2030      If gblnSetFocus = True Then
2040        gblnSetFocus = False
2050        strFormName = vbNullString
2060  On Error Resume Next
2070        Do Until strFormName = THIS_NAME
2080          DoCmd.SelectObject acForm, THIS_NAME, False
2090          strFormName = Screen.ActiveForm.Name
2100        Loop
2110  On Error GoTo ERRH
2120      End If
2130    End With

EXITP:
2140    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2210    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2220    gstrReturningForm = THIS_NAME
2230    gblnSetFocus = True
2240    Select Case strCallingForm
        Case "frmMenu_Utility"
2250      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2260    Case Else
          ' ** Nothing at the moment.
2270    End Select

EXITP:
2280    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2310    With Me
2320      If .cmdBrowse1_raised_focus_dots_img.Visible = True Or .cmdBrowse1_raised_focus_img.Visible = True Then
2330        Select Case blnBrowse1_Focus
            Case True
2340          .cmdBrowse1_raised_semifocus_dots_img.Visible = True
2350          .cmdBrowse1_raised_img.Visible = False
2360        Case False
2370          .cmdBrowse1_raised_img.Visible = True
2380          .cmdBrowse1_raised_semifocus_dots_img.Visible = False
2390        End Select
2400        .cmdBrowse1_raised_focus_img.Visible = False
2410        .cmdBrowse1_raised_focus_dots_img.Visible = False
2420        .cmdBrowse1_sunken_focus_dots_img.Visible = False
2430        .cmdBrowse1_raised_img_dis.Visible = False
2440      End If
2450      If .cmdBrowse2_raised_focus_dots_img.Visible = True Or .cmdBrowse2_raised_focus_img.Visible = True Then
2460        Select Case blnBrowse2_Focus
            Case True
2470          .cmdBrowse2_raised_semifocus_dots_img.Visible = True
2480          .cmdBrowse2_raised_img.Visible = False
2490        Case False
2500          .cmdBrowse2_raised_img.Visible = True
2510          .cmdBrowse2_raised_semifocus_dots_img.Visible = False
2520        End Select
2530        .cmdBrowse2_raised_focus_img.Visible = False
2540        .cmdBrowse2_raised_focus_dots_img.Visible = False
2550        .cmdBrowse2_sunken_focus_dots_img.Visible = False
2560        .cmdBrowse2_raised_img_dis.Visible = False
2570      End If
2580      If .cmdTier1_SelectAll_raised_focus_dots_img.Visible = True Or .cmdTier1_SelectAll_raised_focus_img.Visible = True Then
2590        Select Case blnTier1_SelectAll_Focus
            Case True
2600          .cmdTier1_SelectAll_raised_semifocus_dots_img.Visible = True
2610          .cmdTier1_SelectAll_raised_img.Visible = False
2620        Case False
2630          .cmdTier1_SelectAll_raised_img.Visible = True
2640          .cmdTier1_SelectAll_raised_semifocus_dots_img.Visible = False
2650        End Select
2660        .cmdTier1_SelectAll_raised_focus_img.Visible = False
2670        .cmdTier1_SelectAll_raised_focus_dots_img.Visible = False
2680        .cmdTier1_SelectAll_sunken_focus_dots_img.Visible = False
2690        .cmdTier1_SelectAll_raised_img_dis.Visible = False
2700      End If
2710      If .cmdTier1_SelectNone_raised_focus_dots_img.Visible = True Or .cmdTier1_SelectNone_raised_focus_img.Visible = True Then
2720        Select Case blnTier1_SelectNone_Focus
            Case True
2730          .cmdTier1_SelectNone_raised_semifocus_dots_img.Visible = True
2740          .cmdTier1_SelectNone_raised_img.Visible = False
2750        Case False
2760          .cmdTier1_SelectNone_raised_img.Visible = True
2770          .cmdTier1_SelectNone_raised_semifocus_dots_img.Visible = False
2780        End Select
2790        .cmdTier1_SelectNone_raised_focus_dots_img.Visible = False
2800        .cmdTier1_SelectNone_raised_focus_img.Visible = False
2810        .cmdTier1_SelectNone_sunken_focus_dots_img.Visible = False
2820        .cmdTier1_SelectNone_raised_img_dis.Visible = False
2830      End If
2840      If .cmdTier2_SelectAll_raised_focus_dots_img.Visible = True Or .cmdTier2_SelectAll_raised_focus_img.Visible = True Then
2850        Select Case blnTier2_SelectAll_Focus
            Case True
2860          .cmdTier2_SelectAll_raised_semifocus_dots_img.Visible = True
2870          .cmdTier2_SelectAll_raised_img.Visible = False
2880        Case False
2890          .cmdTier2_SelectAll_raised_img.Visible = True
2900          .cmdTier2_SelectAll_raised_semifocus_dots_img.Visible = False
2910        End Select
2920        .cmdTier2_SelectAll_raised_focus_img.Visible = False
2930        .cmdTier2_SelectAll_raised_focus_dots_img.Visible = False
2940        .cmdTier2_SelectAll_sunken_focus_dots_img.Visible = False
2950        .cmdTier2_SelectAll_raised_img_dis.Visible = False
2960      End If
2970      If .cmdTier2_SelectNone_raised_focus_dots_img.Visible = True Or .cmdTier2_SelectNone_raised_focus_img.Visible = True Then
2980        Select Case blnTier2_SelectNone_Focus
            Case True
2990          .cmdTier2_SelectNone_raised_semifocus_dots_img.Visible = True
3000          .cmdTier2_SelectNone_raised_img.Visible = False
3010        Case False
3020          .cmdTier2_SelectNone_raised_img.Visible = True
3030          .cmdTier2_SelectNone_raised_semifocus_dots_img.Visible = False
3040        End Select
3050        .cmdTier2_SelectNone_raised_focus_dots_img.Visible = False
3060        .cmdTier2_SelectNone_raised_focus_img.Visible = False
3070        .cmdTier2_SelectNone_sunken_focus_dots_img.Visible = False
3080        .cmdTier2_SelectNone_raised_img_dis.Visible = False
3090      End If
3100      If .cmdTier3_SelectAll_raised_focus_dots_img.Visible = True Or .cmdTier3_SelectAll_raised_focus_img.Visible = True Then
3110        Select Case blnTier3_SelectAll_Focus
            Case True
3120          .cmdTier3_SelectAll_raised_semifocus_dots_img.Visible = True
3130          .cmdTier3_SelectAll_raised_img.Visible = False
3140        Case False
3150          .cmdTier3_SelectAll_raised_img.Visible = True
3160          .cmdTier3_SelectAll_raised_semifocus_dots_img.Visible = False
3170        End Select
3180        .cmdTier3_SelectAll_raised_focus_img.Visible = False
3190        .cmdTier3_SelectAll_raised_focus_dots_img.Visible = False
3200        .cmdTier3_SelectAll_sunken_focus_dots_img.Visible = False
3210        .cmdTier3_SelectAll_raised_img_dis.Visible = False
3220      End If
3230      If .cmdTier3_SelectNone_raised_focus_dots_img.Visible = True Or .cmdTier3_SelectNone_raised_focus_img.Visible = True Then
3240        Select Case blnTier3_SelectNone_Focus
            Case True
3250          .cmdTier3_SelectNone_raised_semifocus_dots_img.Visible = True
3260          .cmdTier3_SelectNone_raised_img.Visible = False
3270        Case False
3280          .cmdTier3_SelectNone_raised_img.Visible = True
3290          .cmdTier3_SelectNone_raised_semifocus_dots_img.Visible = False
3300        End Select
3310        .cmdTier3_SelectNone_raised_focus_dots_img.Visible = False
3320        .cmdTier3_SelectNone_raised_focus_img.Visible = False
3330        .cmdTier3_SelectNone_sunken_focus_dots_img.Visible = False
3340        .cmdTier3_SelectNone_raised_img_dis.Visible = False
3350      End If
3360    End With

EXITP:
3370    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

3410    DoCmd.Close acForm, THIS_NAME

EXITP:
3420    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3520    intShiftDown = (Shift And acShiftMask) > 0
3530    intAltDown = (Shift And acAltMask) > 0
3540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3560      Select Case intRetVal
          Case vbKeyTab
3570        With Me
3580          intRetVal = 0
3590          Select Case .opgExport
              Case .opgExport_optAcctProf.OptionValue
3600            .TargetPath.SetFocus
3610          Case .opgExport_optAll.OptionValue
3620            .ckgTier_chkTier1.SetFocus
3630          End Select
3640        End With
3650      End Select
3660    End If

        ' ** Shift keys.
3670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3680      Select Case intRetVal
          Case vbKeyTab
3690        With Me
3700          intRetVal = 0
3710          If .cmdExport.Enabled = True Then
3720            .cmdExport.SetFocus
3730          Else
3740            Select Case .opgExport
                Case .opgExport_optAcctProf.OptionValue
3750              .TargetFile.SetFocus
3760            Case .opgExport_optAll.OptionValue
3770              If .cmdTier3_SelectNone.Enabled = True Then
3780                .cmdTier3_SelectNone.SetFocus
3790              ElseIf .cmdTier3_SelectAll.Enabled = True Then
3800                .cmdTier3_SelectAll.SetFocus
3810              Else
3820                .ckgTier_chkTier3.SetFocus
3830              End If
3840            End Select
3850          End If
3860        End With
3870      End Select
3880    End If

EXITP:
3890    KeyCode = intRetVal
3900    Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub opgExport_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgExport_box_MouseMove"

4010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4020    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub opgExport_lbl_cmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ** Easter Egg (Secret Key) to reveal/hide extended export options.

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgExport_lbl_cmd_MouseUp"

4110    With Me
4120      If Button = acRightButton And Shift = acCtrlMask + acAltMask Then
4130        If CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then
4140          ShowDataSwitch  ' ** Procedure: Below
4150        End If
4160      End If
4170    End With

EXITP:
4180    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub opgExport_AfterUpdate()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgExport_AfterUpdate"

        Dim blnFound As Boolean
        Dim lngX As Long

4210    With Me

4220      GetCounts 5  ' ** Procedure: Below.

4230      Select Case .opgExport
          Case .opgExport_optAcctProf.OptionValue

4240        .FocusHolder.SetFocus
4250        DoEvents

4260        .opgExport_optAcctProf_lbl.FontBold = True
4270        .opgExport_optAll_lbl.FontBold = False
4280        .opgExport_optAcctProf_lbl_box.Visible = True
4290        .opgExport_optAll_lbl_box.Visible = False
4300        .Target_box.BackStyle = acBackStyleNormal
4310        .Target_box2.BackColor = MY_CLR_VLTBGE
4320        .Target_hline03.BorderColor = MY_CLR_VLTBGE

4330        .Target_lbl.ForeColor = CLR_VDKGRY
4340        .Target_lbl_dim_hi.Visible = False
4350        .TargetPath.Enabled = True
4360        .TargetPath.BorderColor = CLR_LTBLU2
4370        .TargetPath.BackStyle = acBackStyleNormal
4380        .TargetFile.Enabled = True
4390        .TargetFile.BorderColor = CLR_LTBLU2
4400        .TargetFile.BackStyle = acBackStyleNormal
4410        .TargetFile_lbl2.ForeColor = CLR_BLK
4420        .TargetFile_lbl2_dim_hi.Visible = False

4430        .cmdBrowse1.Enabled = True
4440        .cmdBrowse1_raised_img.Visible = True
4450        .cmdBrowse1_raised_semifocus_dots_img.Visible = False
4460        .cmdBrowse1_raised_focus_img.Visible = False
4470        .cmdBrowse1_raised_focus_dots_img.Visible = False
4480        .cmdBrowse1_sunken_focus_dots_img.Visible = False
4490        .cmdBrowse1_raised_img_dis.Visible = False

4500        Tier1_Enable_AE False, blnNoData, Me  ' ** Module Procedure: modAcctExportFuncs.

4510        Tier2_Enable_AE False, blnNoData, Me  ' ** Module Procedure: modAcctExportFuncs.

4520        Tier3_Enable_AE False, blnNoData, Me  ' ** Module Procedure: modAcctExportFuncs.

4530        .AllData_box.BackStyle = acBackStyleTransparent
4540        .AllData_box2.BackStyle = acBackStyleTransparent
4550        .AllData_hline03.BorderColor = MY_CLR_LTBGE
4560        .AllData_lbl.ForeColor = WIN_CLR_DISF
4570        If blnNoData = False Then
4580          .AllData_lbl_dim_hi.Visible = True
4590        End If
4600        .AllData_Path.Enabled = False
4610        .AllData_Path.BackStyle = acBackStyleTransparent
4620        .AllData_Path.BorderColor = WIN_CLR_DISR
4630        .AllData_Tbls = 0&
4640        .AllData_Tbls.Locked = False
4650        .AllData_Tbls.BackStyle = acBackStyleTransparent
4660        .AllData_Tbls.BorderColor = WIN_CLR_DISR
4670        .AllData_File.Locked = False
4680        .AllData_File.BackStyle = acBackStyleTransparent
4690        .AllData_File.BorderColor = WIN_CLR_DISR
4700        .AllData_File_lbl2.ForeColor = WIN_CLR_DISF
4710        If blnNoData = False Then
4720          .AllData_File_lbl2_dim_hi.Visible = True
4730        End If

4740        If blnNoData = False Then
4750          .cmdBrowse2.Enabled = False
4760          .cmdBrowse2_raised_img_dis.Visible = True
4770          .cmdBrowse2_raised_img.Visible = False
4780          .cmdBrowse2_raised_semifocus_dots_img.Visible = False
4790          .cmdBrowse2_raised_focus_img.Visible = False
4800          .cmdBrowse2_raised_focus_dots_img.Visible = False
4810          .cmdBrowse2_sunken_focus_dots_img.Visible = False
4820        End If

4830        If blnRecalcTots = True Then
4840          GetCounts 4  ' ** Procedure: Below.
4850        End If

4860        .DefMsg_lbl1.ForeColor = WIN_CLR_DISF
4870        .DefMsg_lbl2.ForeColor = WIN_CLR_DISF
4880        If blnNoData = False Then
4890          .DefMsg_lbl1_dim_hi.Visible = True
4900          .DefMsg_lbl2_dim_hi.Visible = True
4910        End If

4920        .ckgTier1_chk01_Account_lbl_def.ForeColor = WIN_CLR_DISF
4930        .ckgTier1_chk02_MasterAsset_lbl_def.ForeColor = WIN_CLR_DISF
4940        .ckgTier1_chk06_Balance_lbl_def.ForeColor = WIN_CLR_DISF
4950        .ckgTier2_chk01_m_REVCODE_lbl_def.ForeColor = WIN_CLR_DISF
4960        .ckgTier2_chk02_RecurringItems_lbl_def.ForeColor = WIN_CLR_DISF
4970        .ckgTier2_chk04_Currency_History_lbl_def.ForeColor = WIN_CLR_DISF
4980        .ckgTier2_chk06_Location_lbl_def.ForeColor = WIN_CLR_DISF
4990        .ckgTier3_chk01_AdminOfficer_lbl_def.ForeColor = WIN_CLR_DISF
5000        .ckgTier3_chk09_CheckMemo_lbl_def.ForeColor = WIN_CLR_DISF
5010        If blnNoData = False Then
5020          If .ckgTier1_chk01_Account_lbl_def.Visible = True Then
5030            .ckgTier1_chk01_Account_lbl_def_dim_hi.Visible = True
5040          End If
5050          If .ckgTier1_chk02_MasterAsset_lbl_def.Visible = True Then
5060            .ckgTier1_chk02_MasterAsset_lbl_def_dim_hi.Visible = True
5070          End If
5080          If .ckgTier1_chk06_Balance_lbl_def.Visible = True Then
5090            .ckgTier1_chk06_Balance_lbl_def_dim_hi.Visible = True
5100          End If
5110          If .ckgTier2_chk01_m_REVCODE_lbl_def.Visible = True Then
5120            .ckgTier2_chk01_m_REVCODE_lbl_def_dim_hi.Visible = True
5130          End If
5140          If .ckgTier2_chk02_RecurringItems_lbl_def.Visible = True Then
5150            .ckgTier2_chk02_RecurringItems_lbl_def_dim_hi.Visible = True
5160          End If
5170          If .ckgTier2_chk04_Currency_History_lbl_def.Visible = True Then
5180            .ckgTier2_chk04_Currency_History_lbl_def_dim_hi.Visible = True
5190          End If
5200          If .ckgTier2_chk06_Location_lbl_def.Visible = True Then
5210            .ckgTier2_chk06_Location_lbl_def_dim_hi.Visible = True
5220          End If
5230          If .ckgTier3_chk01_AdminOfficer_lbl_def.Visible = True Then
5240            .ckgTier3_chk01_AdminOfficer_lbl_def_dim_hi.Visible = True
5250          End If
5260          If .ckgTier3_chk09_CheckMemo_lbl_def.Visible = True Then
5270            .ckgTier3_chk09_CheckMemo_lbl_def_dim_hi.Visible = True
5280          End If
5290        End If

5300        .cmdExport.Enabled = True

5310      Case .opgExport_optAll.OptionValue

5320        .FocusHolder.SetFocus
5330        DoEvents

5340        .opgExport_optAcctProf_lbl.FontBold = False
5350        .opgExport_optAll_lbl.FontBold = True
5360        .opgExport_optAcctProf_lbl_box.Visible = False
5370        .opgExport_optAll_lbl_box.Visible = True
5380        .Target_box.BackStyle = acBackStyleTransparent
5390        .Target_box2.BackColor = MY_CLR_LTBGE
5400        .Target_hline03.BorderColor = MY_CLR_LTBGE

5410        .Target_lbl.ForeColor = WIN_CLR_DISF
5420        .Target_lbl_dim_hi.Visible = True
5430        .TargetPath.Enabled = False
5440        .TargetPath.BorderColor = WIN_CLR_DISR
5450        .TargetPath.BackStyle = acBackStyleTransparent
5460        .TargetFile.Enabled = False
5470        .TargetFile.BorderColor = WIN_CLR_DISR
5480        .TargetFile.BackStyle = acBackStyleTransparent
5490        .TargetFile_lbl2.ForeColor = WIN_CLR_DISF
5500        .TargetFile_lbl2_dim_hi.Visible = True

5510        .cmdBrowse1.Enabled = False
5520        .cmdBrowse1_raised_img_dis.Visible = True
5530        .cmdBrowse1_raised_img.Visible = False
5540        .cmdBrowse1_raised_semifocus_dots_img.Visible = False
5550        .cmdBrowse1_raised_focus_img.Visible = False
5560        .cmdBrowse1_raised_focus_dots_img.Visible = False
5570        .cmdBrowse1_sunken_focus_dots_img.Visible = False

5580        Tier1_Enable_AE True, blnNoData, Me  ' ** Module Procedure: modAcctExportFuncs.

5590        Tier2_Enable_AE True, blnNoData, Me  ' ** Module Procedure: modAcctExportFuncs.

5600        Tier3_Enable_AE True, blnNoData, Me  ' ** Module Procedure: modAcctExportFuncs.

5610        .AllData_box.BackStyle = acBackStyleNormal
5620        .AllData_box2.BackStyle = acBackStyleNormal
5630        .AllData_hline03.BorderColor = MY_CLR_VLTBGE
5640        .AllData_lbl.ForeColor = CLR_VDKGRY
5650        .AllData_lbl_dim_hi.Visible = False
5660        .AllData_Path.Enabled = True
5670        .AllData_Path.BackStyle = acBackStyleNormal
5680        .AllData_Path.BorderColor = CLR_LTBLU2
5690        .AllData_Tbls.Locked = True
5700        .AllData_Tbls.BackStyle = acBackStyleNormal
5710        .AllData_Tbls.BorderColor = CLR_LTBLU2
5720        .AllData_File.Locked = True
5730        .AllData_File.BackStyle = acBackStyleNormal
5740        .AllData_File.BorderColor = CLR_LTBLU2
5750        .AllData_File_lbl2.ForeColor = CLR_BLK
5760        .AllData_File_lbl2_dim_hi.Visible = False

5770        .cmdBrowse2.Enabled = True
5780        .cmdBrowse2_raised_img.Visible = True
5790        .cmdBrowse2_raised_semifocus_dots_img.Visible = False
5800        .cmdBrowse2_raised_focus_img.Visible = False
5810        .cmdBrowse2_raised_focus_dots_img.Visible = False
5820        .cmdBrowse2_sunken_focus_dots_img.Visible = False
5830        .cmdBrowse2_raised_img_dis.Visible = False

5840        GetCounts 2  ' ** Procedure: Below.

5850        GetCounts 3  ' ** Procedure: Below.

5860        GetCounts 4  ' ** Procedure: Below.

5870        .DefMsg_lbl1.ForeColor = CLR_ROSE
5880        .DefMsg_lbl1_dim_hi.Visible = False
5890        .DefMsg_lbl2.ForeColor = CLR_TEAL
5900        .DefMsg_lbl2_dim_hi.Visible = False

5910        .ckgTier1_chk01_Account_lbl_def.ForeColor = CLR_ROSE
5920        .ckgTier1_chk02_MasterAsset_lbl_def.ForeColor = CLR_ROSE
5930        .ckgTier1_chk06_Balance_lbl_def.ForeColor = CLR_ROSE
5940        .ckgTier2_chk01_m_REVCODE_lbl_def.ForeColor = CLR_ROSE
5950        .ckgTier2_chk02_RecurringItems_lbl_def.ForeColor = CLR_ROSE
5960        .ckgTier2_chk04_Currency_History_lbl_def.ForeColor = CLR_ROSE
5970        .ckgTier2_chk06_Location_lbl_def.ForeColor = CLR_ROSE
5980        .ckgTier3_chk01_AdminOfficer_lbl_def.ForeColor = CLR_ROSE
5990        .ckgTier3_chk09_CheckMemo_lbl_def.ForeColor = CLR_ROSE
6000        .ckgTier1_chk01_Account_lbl_def_dim_hi.Visible = False
6010        .ckgTier1_chk02_MasterAsset_lbl_def_dim_hi.Visible = False
6020        .ckgTier1_chk06_Balance_lbl_def_dim_hi.Visible = False
6030        .ckgTier2_chk01_m_REVCODE_lbl_def_dim_hi.Visible = False
6040        .ckgTier2_chk02_RecurringItems_lbl_def_dim_hi.Visible = False
6050        .ckgTier2_chk04_Currency_History_lbl_def_dim_hi.Visible = False
6060        .ckgTier2_chk06_Location_lbl_def_dim_hi.Visible = False
6070        .ckgTier3_chk01_AdminOfficer_lbl_def_dim_hi.Visible = False
6080        .ckgTier3_chk09_CheckMemo_lbl_def_dim_hi.Visible = False

            ' ** If they're all Zero, except for the defaults,
            ' ** disable cmdExport and all the Select buttons.
6090        blnFound = False: blnEmpty = False
6100        For lngX = 0& To (lngRecs - 1&)
6110          Select Case arr_varRec(R_DEF, lngX)
              Case True
6120            If arr_varRec(R_CNT, lngX) > arr_varRec(R_DCNT, lngX) Then
6130              blnFound = True
6140              Exit For
6150            End If
6160          Case False
6170            If arr_varRec(R_CNT, lngX) <> 0& Then
6180              blnFound = True
6190              Exit For
6200            End If
6210          End Select
6220        Next
6230        Select Case blnFound
            Case True
              ' ** Tables aren't empty, so leave Select buttons active.
6240          .cmdExport.Enabled = True
6250        Case False
              ' ** All tables are either Zero, or default records only.
6260          blnEmpty = True
6270          .cmdExport.Enabled = False
6280          .cmdTier1_SelectAll.Enabled = False
6290          .cmdTier1_SelectAll_raised_img_dis.Visible = True
6300          .cmdTier1_SelectAll_raised_img.Visible = False
6310          .cmdTier1_SelectAll_raised_semifocus_dots_img.Visible = False
6320          .cmdTier1_SelectAll_raised_focus_img.Visible = False
6330          .cmdTier1_SelectAll_raised_focus_dots_img.Visible = False
6340          .cmdTier1_SelectAll_sunken_focus_dots_img.Visible = False
6350          .cmdTier1_SelectNone.Enabled = False
6360          .cmdTier1_SelectNone_raised_img_dis.Visible = True
6370          .cmdTier1_SelectNone_raised_img.Visible = False
6380          .cmdTier1_SelectNone_raised_semifocus_dots_img.Visible = False
6390          .cmdTier1_SelectNone_raised_focus_img.Visible = False
6400          .cmdTier1_SelectNone_raised_focus_dots_img.Visible = False
6410          .cmdTier1_SelectNone_sunken_focus_dots_img.Visible = False
6420          .cmdTier2_SelectAll.Enabled = False
6430          .cmdTier2_SelectAll_raised_img_dis.Visible = True
6440          .cmdTier2_SelectAll_raised_img.Visible = False
6450          .cmdTier2_SelectAll_raised_semifocus_dots_img.Visible = False
6460          .cmdTier2_SelectAll_raised_focus_img.Visible = False
6470          .cmdTier2_SelectAll_raised_focus_dots_img.Visible = False
6480          .cmdTier2_SelectAll_sunken_focus_dots_img.Visible = False
6490          .cmdTier2_SelectNone.Enabled = False
6500          .cmdTier2_SelectNone_raised_img_dis.Visible = True
6510          .cmdTier2_SelectNone_raised_img.Visible = False
6520          .cmdTier2_SelectNone_raised_semifocus_dots_img.Visible = False
6530          .cmdTier2_SelectNone_raised_focus_img.Visible = False
6540          .cmdTier2_SelectNone_raised_focus_dots_img.Visible = False
6550          .cmdTier2_SelectNone_sunken_focus_dots_img.Visible = False
6560          .cmdTier3_SelectAll.Enabled = False
6570          .cmdTier3_SelectAll_raised_img_dis.Visible = True
6580          .cmdTier3_SelectAll_raised_img.Visible = False
6590          .cmdTier3_SelectAll_raised_semifocus_dots_img.Visible = False
6600          .cmdTier3_SelectAll_raised_focus_img.Visible = False
6610          .cmdTier3_SelectAll_raised_focus_dots_img.Visible = False
6620          .cmdTier3_SelectAll_sunken_focus_dots_img.Visible = False
6630          .cmdTier3_SelectNone.Enabled = False
6640          .cmdTier3_SelectNone_raised_img_dis.Visible = True
6650          .cmdTier3_SelectNone_raised_img.Visible = False
6660          .cmdTier3_SelectNone_raised_semifocus_dots_img.Visible = False
6670          .cmdTier3_SelectNone_raised_focus_img.Visible = False
6680          .cmdTier3_SelectNone_raised_focus_dots_img.Visible = False
6690          .cmdTier3_SelectNone_sunken_focus_dots_img.Visible = False
6700        End Select

6710      End Select
6720    End With

EXITP:
6730    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub opgExport_optAcctProf_KeyDown(KeyCode As Integer, Shift As Integer)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgExport_optAcctProf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6820    intShiftDown = (Shift And acShiftMask) > 0
6830    intAltDown = (Shift And acAltMask) > 0
6840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6870        With Me
6880          intRetVal = 0
6890          Select Case .opgExport
              Case .opgExport_optAcctProf.OptionValue
6900            .TargetPath.SetFocus
6910          Case .opgExport_optAll.OptionValue
6920            .ckgTier_chkTier1.SetFocus
6930          End Select
6940        End With
6950      End Select
6960    End If

        ' ** Shift keys.
6970    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6980      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6990        With Me
7000          intRetVal = 0
7010          .cmdCancel.SetFocus
7020        End With
7030      End Select
7040    End If

EXITP:
7050    KeyCode = intRetVal
7060    Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub opgExport_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgExport_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7120    intShiftDown = (Shift And acShiftMask) > 0
7130    intAltDown = (Shift And acAltMask) > 0
7140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7170        With Me
7180          intRetVal = 0
7190          Select Case .opgExport
              Case .opgExport_optAcctProf.OptionValue
7200            .TargetPath.SetFocus
7210          Case .opgExport_optAll.OptionValue
7220            .ckgTier_chkTier1.SetFocus
7230          End Select
7240        End With
7250      End Select
7260    End If

        ' ** Shift keys.
7270    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7280      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7290        With Me
7300          intRetVal = 0
7310          .cmdCancel.SetFocus
7320        End With
7330      End Select
7340    End If

EXITP:
7350    KeyCode = intRetVal
7360    Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub opgExport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgExport_MouseMove"

7410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7420    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub Target_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "Target_box_MouseMove"

7510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7520    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub TargetPath_AfterUpdate()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "TargetPath_AfterUpdate"

        Dim strPath As String

7610    With Me
7620      If IsNull(.TargetPath) = False Then
7630        strPath = Trim(.TargetPath)
7640        If strPath <> vbNullString Then
7650          If .TargetPath.Enabled = True Then
7660            .TargetPath.SetFocus
7670            .TargetPath.SelLength = 0
7680            .TargetPath.SelStart = 99
7690          End If
7700        End If
7710      End If
7720    End With

EXITP:
7730    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub TargetPath_KeyDown(KeyCode As Integer, Shift As Integer)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "TargetPath_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7820    intShiftDown = (Shift And acShiftMask) > 0
7830    intAltDown = (Shift And acAltMask) > 0
7840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7870        With Me
7880          intRetVal = 0
7890          .TargetFile.SetFocus
7900        End With
7910      End Select
7920    End If

        ' ** Shift keys.
7930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7950        With Me
7960          intRetVal = 0
7970          .opgExport.SetFocus
7980        End With
7990      End Select
8000    End If

        ' ** Ctrl keys.
8010    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8020      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8030        With Me
8040          intRetVal = 0
8050          .cmdBrowse1.SetFocus
8060        End With
8070      End Select
8080    End If

EXITP:
8090    KeyCode = intRetVal
8100    Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub TargetPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "TargetPath_MouseMove"

8210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8220    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdBrowse1_GotFocus()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse1_GotFocus"

8310    With Me
8320      blnBrowse1_Focus = True
8330      .cmdBrowse1_raised_semifocus_dots_img.Visible = True
8340      .cmdBrowse1_raised_img.Visible = False
8350      .cmdBrowse1_raised_focus_img.Visible = False
8360      .cmdBrowse1_raised_focus_dots_img.Visible = False
8370      .cmdBrowse1_sunken_focus_dots_img.Visible = False
8380      .cmdBrowse1_raised_img_dis.Visible = False
8390    End With

EXITP:
8400    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdBrowse1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse1_MouseDown"

8510    With Me
8520      blnBrowse1_MouseDown = True
8530      .cmdBrowse1_sunken_focus_dots_img.Visible = True
8540      .cmdBrowse1_raised_img.Visible = False
8550      .cmdBrowse1_raised_semifocus_dots_img.Visible = False
8560      .cmdBrowse1_raised_focus_img.Visible = False
8570      .cmdBrowse1_raised_focus_dots_img.Visible = False
8580      .cmdBrowse1_raised_img_dis.Visible = False
8590    End With

EXITP:
8600    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdBrowse1_Click()
' ** Use Windows API for returning a path but no file name.

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse1_Click"

        Dim strTitle1 As String
        Dim strPath As String

8710    With Me
8720      strTitle1 = "Select a folder"
8730      strPath = GetFolderPathSIS(strTitle1)  ' ** Module Function: modBrowseFilesAndFolders.
8740      If strPath <> vbNullString Then
8750        .TargetPath = strPath
8760        DoEvents
8770        TargetPath_AfterUpdate  ' ** Procedure: Below.
8780      End If
8790      gblnSetFocus = True
8800      .TimerInterval = 100&
8810    End With

EXITP:
8820    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdBrowse1_KeyDown(KeyCode As Integer, Shift As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8920    intShiftDown = (Shift And acShiftMask) > 0
8930    intAltDown = (Shift And acAltMask) > 0
8940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8960      Select Case intRetVal
          Case vbKeyTab
8970        With Me
8980          intRetVal = 0
8990          .TargetFile.SetFocus
9000        End With
9010      End Select
9020    End If

        ' ** Shift keys.
9030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9040      Select Case intRetVal
          Case vbKeyTab
9050        With Me
9060          intRetVal = 0
9070          .TargetPath.SetFocus
9080        End With
9090      End Select
9100    End If

EXITP:
9110    KeyCode = intRetVal
9120    Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdBrowse1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse1_MouseMove"

9210    With Me
9220      If blnBrowse1_MouseDown = False Then
9230        Select Case blnBrowse1_Focus
            Case True
9240          .cmdBrowse1_raised_focus_dots_img.Visible = True
9250          .cmdBrowse1_raised_focus_img.Visible = False
9260        Case False
9270          .cmdBrowse1_raised_focus_img.Visible = True
9280          .cmdBrowse1_raised_focus_dots_img.Visible = False
9290        End Select
9300        .cmdBrowse1_raised_img.Visible = False
9310        .cmdBrowse1_raised_semifocus_dots_img.Visible = False
9320        .cmdBrowse1_sunken_focus_dots_img.Visible = False
9330        .cmdBrowse1_raised_img_dis.Visible = False
9340      End If
9350    End With

EXITP:
9360    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdBrowse1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse1_MouseUp"

9410    With Me
9420      .cmdBrowse1_raised_focus_dots_img.Visible = True
9430      .cmdBrowse1_raised_img.Visible = False
9440      .cmdBrowse1_raised_semifocus_dots_img.Visible = False
9450      .cmdBrowse1_raised_focus_img.Visible = False
9460      .cmdBrowse1_sunken_focus_dots_img.Visible = False
9470      .cmdBrowse1_raised_img_dis.Visible = False
9480      blnBrowse1_MouseDown = False
9490    End With

EXITP:
9500    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdBrowse1_LostFocus()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse1_LostFocus"

9610    With Me
9620      .cmdBrowse1_raised_img.Visible = True
9630      .cmdBrowse1_raised_semifocus_dots_img.Visible = False
9640      .cmdBrowse1_raised_focus_img.Visible = False
9650      .cmdBrowse1_raised_focus_dots_img.Visible = False
9660      .cmdBrowse1_sunken_focus_dots_img.Visible = False
9670      .cmdBrowse1_raised_img_dis.Visible = False
9680      blnBrowse1_Focus = False
9690    End With

EXITP:
9700    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub TargetFile_AfterUpdate()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "TargetFile_AfterUpdate"

        Dim intPos01 As Integer

9810    With Me
9820      If IsNull(.TargetFile) = False Then
9830        strTargetFile = Trim(.TargetFile)
9840        If strTargetFile <> vbNullString Then
9850          intPos01 = InStr(strTargetFile, ".")
9860          If intPos01 > 0 Then
9870            strTargetFile = Left(strTargetFile, (intPos01 - 1))
9880            .TargetFile = strTargetFile  ' ** No extension in field.
9890            strTargetFile = strTargetFile & ".csv"
9900          Else
9910            strTargetFile = strTargetFile & ".csv"
9920          End If
9930        End If
9940      Else
9950        strTargetFile = vbNullString
9960      End If
9970    End With

EXITP:
9980    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub TargetFile_KeyDown(KeyCode As Integer, Shift As Integer)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "TargetFile_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10020   intShiftDown = (Shift And acShiftMask) > 0
10030   intAltDown = (Shift And acAltMask) > 0
10040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10070       With Me
10080         intRetVal = 0
10090         If .cmdExport.Enabled = True Then
10100           .cmdExport.SetFocus
10110         Else
10120           .cmdCancel.SetFocus
10130         End If
10140       End With
10150     End Select
10160   End If

        ' ** Shift keys.
10170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10180     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10190       With Me
10200         intRetVal = 0
10210         .TargetPath.SetFocus
10220       End With
10230     End Select
10240   End If

        ' ** Ctrl-Shift keys.
10250   If intCtrlDown And (Not intAltDown) And intShiftDown Then
10260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10270       With Me
10280         intRetVal = 0
10290         .cmdBrowse1.SetFocus
10300       End With
10310     End Select
10320   End If

EXITP:
10330   KeyCode = intRetVal
10340   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdExport_Click()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExport_Click"

        Dim strTargetPath As String, strTargetPathFile As String

10410   With Me

10420     Select Case .opgExport
          Case .opgExport_optAcctProf.OptionValue

            ' ** Target path must be filled.
10430       If IsNull(.TargetPath) = True Then
10440         strTargetPath = CurDir()
10450         .TargetPath = strTargetPath
10460         DoEvents
10470       Else
10480         If Trim(.TargetPath) = vbNullString Then
10490           strTargetPath = CurDir()
10500           .TargetPath = strTargetPath
10510           DoEvents
10520         Else
10530           strTargetPath = IIf(Right(Trim(.TargetPath), 1) <> LNK_SEP, Trim(.TargetPath), _
                  Left(Trim(.TargetPath), Len(Trim(.TargetPath)) - 1))
10540         End If
10550       End If

            ' ** Target file must be filled.
10560       If IsNull(.TargetFile) Then
10570         MsgBox "Please specify a target file name.", vbInformation + vbOKOnly, "Entry Required"
10580         .TargetFile.SetFocus
10590       Else
              ' ** strTargetFile should already be filled.

10600         If DirExists(strTargetPath) = True Then  ' ** Module Function: modFileUtilities.

10610           strTargetPathFile = strTargetPath & LNK_SEP & strTargetFile

                ' ** Updated: 04/02/2017.
10620           DoCmd.TransferText acExportDelim, "TACloud_Account_Export_Specification", "qryAccountExport_02", strTargetPathFile, True
10630           DoEvents
10640           MsgBox "Export Complete." & vbCrLf & vbCrLf & _
                  "The file '" & strTargetFile & "' is in the '" & strTargetPath & "' folder.", _
                  vbInformation + vbOKOnly, "Export Successful"

10650         Else
10660           Beep
10670           MsgBox "The Folder '" & strTargetPath & "' cannot be found.", vbInformation + vbOKOnly, "Folder Not Found"
10680         End If
10690       End If

10700     Case .opgExport_optAll.OptionValue

10710       ExportToCloud  ' ** Procedure: Below.

10720     End Select

10730   End With

EXITP:
10740   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdExport_KeyDown(KeyCode As Integer, Shift As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10820   intShiftDown = (Shift And acShiftMask) > 0
10830   intAltDown = (Shift And acAltMask) > 0
10840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10860     Select Case intRetVal
          Case vbKeyTab
10870       With Me
10880         intRetVal = 0
10890         .cmdCancel.SetFocus
10900       End With
10910     End Select
10920   End If

        ' ** Shift keys.
10930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10940     Select Case intRetVal
          Case vbKeyTab
10950       With Me
10960         intRetVal = 0
10970         Select Case .opgExport
              Case .opgExport_optAcctProf.OptionValue
10980           .TargetFile.SetFocus
10990         Case .opgExport_optAll.OptionValue
11000           If .cmdTier3_SelectNone.Enabled = True Then
11010             .cmdTier3_SelectNone.SetFocus
11020           ElseIf .cmdTier3_SelectAll.Enabled = True Then
11030             .cmdTier3_SelectAll.SetFocus
11040           Else
11050             .ckgTier_chkTier3.SetFocus
11060           End If
11070         End Select
11080       End With
11090     End Select
11100   End If

EXITP:
11110   KeyCode = intRetVal
11120   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub ckgTier_chkTier1_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier_chkTier1_box_MouseMove"

11210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11220   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier_chkTier1_AfterUpdate()

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier_chkTier1_AfterUpdate"

11310   With Me
11320     Select Case .ckgTier_chkTier1
          Case True
11330       .ckgTier_chkTier1_lbl2.FontBold = True
11340       .ckgTier_chkTier1_lbl2_dim_hi.FontBold = True
11350       .ckgTier_chkTier1_box.BackStyle = acBackStyleNormal
11360       .ckgTier_chkTier1_box2.BackStyle = acBackStyleNormal
11370       .ckgTier_chkTier1_hline03.BorderColor = MY_CLR_VLTBGE
11380       .ckgTier1_chk01_Account.Enabled = True
11390       .ckgTier1_chk02_MasterAsset.Enabled = True
11400       .ckgTier1_chk03_ActiveAssets.Enabled = True
11410       .ckgTier1_chk04_LedgerArchive.Enabled = True
11420       .ckgTier1_chk05_Ledger.Enabled = True
11430       .ckgTier1_chk06_Balance.Enabled = True
11440       .ckgTier1_chk01_Account_lbl_recs.ForeColor = CLR_DKGRY
11450       .ckgTier1_chk01_Account_lbl_recs.BorderColor = CLR_VLTBLU2
11460       .ckgTier1_chk02_MasterAsset_lbl_recs.ForeColor = CLR_DKGRY
11470       .ckgTier1_chk02_MasterAsset_lbl_recs.BorderColor = CLR_VLTBLU2
11480       .ckgTier1_chk03_ActiveAssets_lbl_recs.ForeColor = CLR_DKGRY
11490       .ckgTier1_chk03_ActiveAssets_lbl_recs.BorderColor = CLR_VLTBLU2
11500       .ckgTier1_chk04_LedgerArchive_lbl_recs.ForeColor = CLR_DKGRY
11510       .ckgTier1_chk04_LedgerArchive_lbl_recs.BorderColor = CLR_VLTBLU2
11520       .ckgTier1_chk05_Ledger_lbl_recs.ForeColor = CLR_DKGRY
11530       .ckgTier1_chk05_Ledger_lbl_recs.BorderColor = CLR_VLTBLU2
11540       .ckgTier1_chk06_Balance_lbl_recs.ForeColor = CLR_DKGRY
11550       .ckgTier1_chk06_Balance_lbl_recs.BorderColor = CLR_VLTBLU2
11560       .cmdTier1_Select_box.BackStyle = acBackStyleNormal
11570       .cmdTier1_Select_hline03.BorderColor = MY_CLR_VLTBGE
11580       .cmdTier1_Select_hline04.BorderColor = MY_CLR_VLTBGE
11590       If blnEmpty = False And blnNoData = False Then
              ' ** If it's empty, they should already be disabled.
11600         .cmdTier1_SelectAll.Enabled = True
11610         .cmdTier1_SelectAll_raised_img.Visible = True
11620         .cmdTier1_SelectAll_raised_semifocus_dots_img.Visible = False
11630         .cmdTier1_SelectAll_raised_focus_img.Visible = False
11640         .cmdTier1_SelectAll_raised_focus_dots_img.Visible = False
11650         .cmdTier1_SelectAll_sunken_focus_dots_img.Visible = False
11660         .cmdTier1_SelectAll_raised_img_dis.Visible = False
11670         .cmdTier1_SelectNone.Enabled = True
11680         .cmdTier1_SelectNone_raised_img.Visible = True
11690         .cmdTier1_SelectNone_raised_semifocus_dots_img.Visible = False
11700         .cmdTier1_SelectNone_raised_focus_img.Visible = False
11710         .cmdTier1_SelectNone_raised_focus_dots_img.Visible = False
11720         .cmdTier1_SelectNone_sunken_focus_dots_img.Visible = False
11730         .cmdTier1_SelectNone_raised_img_dis.Visible = False
11740       End If
11750     Case False
11760       .ckgTier_chkTier1_lbl2.FontBold = False
11770       .ckgTier_chkTier1_lbl2_dim_hi.FontBold = False
11780       .ckgTier_chkTier1_box.BackStyle = acBackStyleTransparent
11790       .ckgTier_chkTier1_box2.BackStyle = acBackStyleTransparent
11800       .ckgTier_chkTier1_hline03.BorderColor = MY_CLR_LTBGE
11810       .ckgTier1_chk01_Account.Enabled = False
11820       .ckgTier1_chk02_MasterAsset.Enabled = False
11830       .ckgTier1_chk03_ActiveAssets.Enabled = False
11840       .ckgTier1_chk04_LedgerArchive.Enabled = False
11850       .ckgTier1_chk05_Ledger.Enabled = False
11860       .ckgTier1_chk06_Balance.Enabled = False
11870       .ckgTier1_chk01_Account_lbl_recs.ForeColor = WIN_CLR_DISF
11880       .ckgTier1_chk01_Account_lbl_recs.BorderColor = WIN_CLR_DISR
11890       .ckgTier1_chk02_MasterAsset_lbl_recs.ForeColor = WIN_CLR_DISF
11900       .ckgTier1_chk02_MasterAsset_lbl_recs.BorderColor = WIN_CLR_DISR
11910       .ckgTier1_chk03_ActiveAssets_lbl_recs.ForeColor = WIN_CLR_DISF
11920       .ckgTier1_chk03_ActiveAssets_lbl_recs.BorderColor = WIN_CLR_DISR
11930       .ckgTier1_chk04_LedgerArchive_lbl_recs.ForeColor = WIN_CLR_DISF
11940       .ckgTier1_chk04_LedgerArchive_lbl_recs.BorderColor = WIN_CLR_DISR
11950       .ckgTier1_chk05_Ledger_lbl_recs.ForeColor = WIN_CLR_DISF
11960       .ckgTier1_chk05_Ledger_lbl_recs.BorderColor = WIN_CLR_DISR
11970       .ckgTier1_chk06_Balance_lbl_recs.ForeColor = WIN_CLR_DISF
11980       .ckgTier1_chk06_Balance_lbl_recs.BorderColor = WIN_CLR_DISR
11990       .cmdTier1_Select_box.BackStyle = acBackStyleTransparent
12000       .cmdTier1_Select_hline03.BorderColor = MY_CLR_LTBGE
12010       .cmdTier1_Select_hline04.BorderColor = MY_CLR_LTBGE
12020       If blnNoData = False Then
12030         .cmdTier1_SelectAll.Enabled = False
12040         .cmdTier1_SelectAll_raised_img_dis.Visible = True
12050         .cmdTier1_SelectAll_raised_img.Visible = False
12060         .cmdTier1_SelectAll_raised_semifocus_dots_img.Visible = False
12070         .cmdTier1_SelectAll_raised_focus_img.Visible = False
12080         .cmdTier1_SelectAll_raised_focus_dots_img.Visible = False
12090         .cmdTier1_SelectAll_sunken_focus_dots_img.Visible = False
12100         .cmdTier1_SelectNone.Enabled = False
12110         .cmdTier1_SelectNone_raised_img_dis.Visible = True
12120         .cmdTier1_SelectNone_raised_img.Visible = False
12130         .cmdTier1_SelectNone_raised_semifocus_dots_img.Visible = False
12140         .cmdTier1_SelectNone_raised_focus_img.Visible = False
12150         .cmdTier1_SelectNone_raised_focus_dots_img.Visible = False
12160         .cmdTier1_SelectNone_sunken_focus_dots_img.Visible = False
12170       End If
12180     End Select
12190     GetCounts 2  ' ** Procedure: Below.
12200     GetCounts 3  ' ** Procedure: Below.
12210     GetCounts 4  ' ** Procedure: Below.
12220   End With

EXITP:
12230   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier_chkTier1_KeyDown(KeyCode As Integer, Shift As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier_chkTier1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12320   intShiftDown = (Shift And acShiftMask) > 0
12330   intAltDown = (Shift And acAltMask) > 0
12340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12370       With Me
12380         intRetVal = 0
12390         Select Case .ckgTier_chkTier1
              Case True
12400           .ckgTier1_chk01_Account.SetFocus
12410         Case False
12420           .ckgTier_chkTier2.SetFocus
12430         End Select
12440       End With
12450     End Select
12460   End If

        ' ** Shift keys.
12470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12490       With Me
12500         intRetVal = 0
12510         .opgExport.SetFocus
12520       End With
12530     End Select
12540   End If

EXITP:
12550   KeyCode = intRetVal
12560   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub ckgTier1_chk01_Account_AfterUpdate()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier1_chk01_Account_AfterUpdate"

12610   With Me
12620     Select Case .ckgTier1_chk01_Account
          Case True
12630       .ckgTier1_chk01_Account_lbl.FontBold = True
12640     Case False
12650       .ckgTier1_chk01_Account_lbl.FontBold = False
12660     End Select
12670     GetCounts 2  ' ** Procedure: Below.
12680   End With

EXITP:
12690   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier1_chk01_Account_KeyDown(KeyCode As Integer, Shift As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier1_chk01_Account_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12720   intShiftDown = (Shift And acShiftMask) > 0
12730   intAltDown = (Shift And acAltMask) > 0
12740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12770       With Me
12780         intRetVal = 0
12790         .ckgTier1_chk02_MasterAsset.SetFocus
12800       End With
12810     End Select
12820   End If

        ' ** Shift keys.
12830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12850       With Me
12860         intRetVal = 0
12870         .ckgTier_chkTier1.SetFocus
12880       End With
12890     End Select
12900   End If

EXITP:
12910   KeyCode = intRetVal
12920   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub ckgTier1_chk02_MasterAsset_AfterUpdate()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier1_chk02_MasterAsset_AfterUpdate"

13010   With Me
13020     Select Case .ckgTier1_chk02_MasterAsset
          Case True
13030       .ckgTier1_chk02_MasterAsset_lbl.FontBold = True
13040     Case False
13050       .ckgTier1_chk02_MasterAsset_lbl.FontBold = False
13060     End Select
13070     GetCounts 2  ' ** Procedure: Below.
13080   End With

EXITP:
13090   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier1_chk03_ActiveAssets_AfterUpdate()

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier1_chk03_ActiveAssets_AfterUpdate"

13110   With Me
13120     Select Case .ckgTier1_chk03_ActiveAssets
          Case True
13130       .ckgTier1_chk03_ActiveAssets_lbl.FontBold = True
13140     Case False
13150       .ckgTier1_chk03_ActiveAssets_lbl.FontBold = False
13160     End Select
13170     GetCounts 2  ' ** Procedure: Below.
13180   End With

EXITP:
13190   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier1_chk04_LedgerArchive_AfterUpdate()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier1_chk04_LedgerArchive_AfterUpdate"

13210   With Me
13220     Select Case .ckgTier1_chk04_LedgerArchive
          Case True
13230       .ckgTier1_chk04_LedgerArchive_lbl.FontBold = True
13240     Case False
13250       .ckgTier1_chk04_LedgerArchive_lbl.FontBold = False
13260     End Select
13270     GetCounts 2  ' ** Procedure: Below.
13280   End With

EXITP:
13290   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier1_chk05_Ledger_AfterUpdate()

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier1_chk05_Ledger_AfterUpdate"

13310   With Me
13320     Select Case .ckgTier1_chk05_Ledger
          Case True
13330       .ckgTier1_chk05_Ledger_lbl.FontBold = True
13340     Case False
13350       .ckgTier1_chk05_Ledger_lbl.FontBold = False
13360     End Select
13370     GetCounts 2  ' ** Procedure: Below.
13380   End With

EXITP:
13390   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier1_chk06_Balance_AfterUpdate()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier1_chk06_Balance_AfterUpdate"

13410   With Me
13420     Select Case .ckgTier1_chk06_Balance
          Case True
13430       .ckgTier1_chk06_Balance_lbl.FontBold = True
13440     Case False
13450       .ckgTier1_chk06_Balance_lbl.FontBold = False
13460     End Select
13470     GetCounts 2  ' ** Procedure: Below.
13480   End With

EXITP:
13490   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier1_chk06_Balance_KeyDown(KeyCode As Integer, Shift As Integer)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier1_chk06_Balance_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13520   intShiftDown = (Shift And acShiftMask) > 0
13530   intAltDown = (Shift And acAltMask) > 0
13540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13570       With Me
13580         intRetVal = 0
13590         If .cmdTier1_SelectAll.Enabled = True Then
13600           .cmdTier1_SelectAll.SetFocus
13610         Else
13620           .cmdTier1_SelectNone.SetFocus
13630         End If
13640       End With
13650     End Select
13660   End If

        ' ** Shift keys.
13670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13680     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13690       With Me
13700         intRetVal = 0
13710         .ckgTier1_chk05_Ledger.SetFocus
13720       End With
13730     End Select
13740   End If

EXITP:
13750   KeyCode = intRetVal
13760   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub ckgTier1_chk06_Balance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier1_chk06_Balance_MouseMove"

13810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
13820   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier1_SelectAll_GotFocus()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectAll_GotFocus"

13910   With Me
13920     blnTier1_SelectAll_Focus = True
13930     .cmdTier1_SelectAll_raised_semifocus_dots_img.Visible = True
13940     .cmdTier1_SelectAll_raised_img.Visible = False
13950     .cmdTier1_SelectAll_raised_focus_img.Visible = False
13960     .cmdTier1_SelectAll_raised_focus_dots_img.Visible = False
13970     .cmdTier1_SelectAll_sunken_focus_dots_img.Visible = False
13980     .cmdTier1_SelectAll_raised_img_dis.Visible = False
13990   End With

EXITP:
14000   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier1_SelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectAll_MouseDown"

14110   With Me
14120     blnTier1_SelectAll_MouseDown = True
14130     .cmdTier1_SelectAll_sunken_focus_dots_img.Visible = True
14140     .cmdTier1_SelectAll_raised_img.Visible = False
14150     .cmdTier1_SelectAll_raised_semifocus_dots_img.Visible = False
14160     .cmdTier1_SelectAll_raised_focus_img.Visible = False
14170     .cmdTier1_SelectAll_raised_focus_dots_img.Visible = False
14180     .cmdTier1_SelectAll_raised_img_dis.Visible = False
14190   End With

EXITP:
14200   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier1_SelectAll_Click()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectAll_Click"

14310   With Me
14320     DoCmd.Hourglass True
14330     DoEvents
14340     If .ckgTier1_chk01_Account.Enabled = True Then
14350       .ckgTier1_chk01_Account = True
14360       ckgTier1_chk01_Account_AfterUpdate  ' ** Procedure: Above.
14370       DoEvents
14380     End If
14390     If .ckgTier1_chk02_MasterAsset.Enabled = True Then
14400       .ckgTier1_chk02_MasterAsset = True
14410       ckgTier1_chk02_MasterAsset_AfterUpdate  ' ** Procedure: Above.
14420       DoEvents
14430     End If
14440     If .ckgTier1_chk03_ActiveAssets.Enabled = True Then
14450       .ckgTier1_chk03_ActiveAssets = True
14460       ckgTier1_chk03_ActiveAssets_AfterUpdate  ' ** Procedure: Above.
14470       DoEvents
14480     End If
14490     If .ckgTier1_chk04_LedgerArchive.Enabled = True Then
14500       .ckgTier1_chk04_LedgerArchive = True
14510       ckgTier1_chk04_LedgerArchive_AfterUpdate  ' ** Procedure: Above.
14520       DoEvents
14530     End If
14540     If .ckgTier1_chk05_Ledger.Enabled = True Then
14550       .ckgTier1_chk05_Ledger = True
14560       ckgTier1_chk05_Ledger_AfterUpdate  ' ** Procedure: Above.
14570       DoEvents
14580     End If
14590     If .ckgTier1_chk06_Balance.Enabled = True Then
14600       .ckgTier1_chk06_Balance = True
14610       ckgTier1_chk06_Balance_AfterUpdate  ' ** Procedure: Above.
14620       DoEvents
14630     End If
14640     GetCounts 2  ' ** Procedure: Below.
14650     DoCmd.Hourglass False
14660   End With

EXITP:
14670   Exit Sub

ERRH:
1220    DoCmd.Hourglass False
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdTier1_SelectAll_KeyDown(KeyCode As Integer, Shift As Integer)

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14720   intShiftDown = (Shift And acShiftMask) > 0
14730   intAltDown = (Shift And acAltMask) > 0
14740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14760     Select Case intRetVal
          Case vbKeyTab
14770       With Me
14780         intRetVal = 0
14790         If .cmdTier1_SelectNone.Enabled = True Then
14800           .cmdTier1_SelectNone.SetFocus
14810         Else
14820           .ckgTier_chkTier2.SetFocus
14830         End If
14840       End With
14850     End Select
14860   End If

        ' ** Shift keys.
14870   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14880     Select Case intRetVal
          Case vbKeyTab
14890       With Me
14900         intRetVal = 0
14910         .ckgTier1_chk06_Balance.SetFocus
14920       End With
14930     End Select
14940   End If

EXITP:
14950   KeyCode = intRetVal
14960   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdTier1_SelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectAll_MouseMove"

15010   With Me
15020     If blnTier1_SelectAll_MouseDown = False Then
15030       Select Case blnTier1_SelectAll_Focus
            Case True
15040         .cmdTier1_SelectAll_raised_focus_dots_img.Visible = True
15050         .cmdTier1_SelectAll_raised_focus_img.Visible = False
15060       Case False
15070         .cmdTier1_SelectAll_raised_focus_img.Visible = True
15080         .cmdTier1_SelectAll_raised_focus_dots_img.Visible = False
15090       End Select
15100       .cmdTier1_SelectAll_raised_img.Visible = False
15110       .cmdTier1_SelectAll_raised_semifocus_dots_img.Visible = False
15120       .cmdTier1_SelectAll_sunken_focus_dots_img.Visible = False
15130       .cmdTier1_SelectAll_raised_img_dis.Visible = False
15140     End If
15150   End With

EXITP:
15160   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier1_SelectAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectAll_MouseUp"

15210   With Me
15220     .cmdTier1_SelectAll_raised_focus_dots_img.Visible = True
15230     .cmdTier1_SelectAll_raised_img.Visible = False
15240     .cmdTier1_SelectAll_raised_semifocus_dots_img.Visible = False
15250     .cmdTier1_SelectAll_raised_focus_img.Visible = False
15260     .cmdTier1_SelectAll_sunken_focus_dots_img.Visible = False
15270     .cmdTier1_SelectAll_raised_img_dis.Visible = False
15280     blnTier1_SelectAll_MouseDown = False
15290   End With

EXITP:
15300   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier1_SelectAll_LostFocus()

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectAll_LostFocus"

15410   With Me
15420     .cmdTier1_SelectAll_raised_img.Visible = True
15430     .cmdTier1_SelectAll_raised_semifocus_dots_img.Visible = False
15440     .cmdTier1_SelectAll_raised_focus_img.Visible = False
15450     .cmdTier1_SelectAll_raised_focus_dots_img.Visible = False
15460     .cmdTier1_SelectAll_sunken_focus_dots_img.Visible = False
15470     .cmdTier1_SelectAll_raised_img_dis.Visible = False
15480     blnTier1_SelectAll_Focus = False
15490   End With

EXITP:
15500   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier1_SelectNone_GotFocus()

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectNone_GotFocus"

15610   With Me
15620     blnTier1_SelectNone_Focus = True
15630     .cmdTier1_SelectNone_raised_semifocus_dots_img.Visible = True
15640     .cmdTier1_SelectNone_raised_img.Visible = False
15650     .cmdTier1_SelectNone_raised_focus_img.Visible = False
15660     .cmdTier1_SelectNone_raised_focus_dots_img.Visible = False
15670     .cmdTier1_SelectNone_sunken_focus_dots_img.Visible = False
15680     .cmdTier1_SelectNone_raised_img_dis.Visible = False
15690   End With

EXITP:
15700   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier1_SelectNone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectNone_MouseDown"

15810   With Me
15820     blnTier1_SelectNone_MouseDown = True
15830     .cmdTier1_SelectNone_sunken_focus_dots_img.Visible = True
15840     .cmdTier1_SelectNone_raised_img.Visible = False
15850     .cmdTier1_SelectNone_raised_semifocus_dots_img.Visible = False
15860     .cmdTier1_SelectNone_raised_focus_img.Visible = False
15870     .cmdTier1_SelectNone_raised_focus_dots_img.Visible = False
15880     .cmdTier1_SelectNone_raised_img_dis.Visible = False
15890   End With

EXITP:
15900   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier1_SelectNone_Click()

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectNone_Click"

16010   With Me
16020     DoCmd.Hourglass True
16030     DoEvents
16040     If .ckgTier1_chk01_Account.Enabled = True Then
16050       .ckgTier1_chk01_Account = False
16060       ckgTier1_chk01_Account_AfterUpdate  ' ** Procedure: Above.
16070       DoEvents
16080     End If
16090     If .ckgTier1_chk02_MasterAsset.Enabled = True Then
16100       .ckgTier1_chk02_MasterAsset = False
16110       ckgTier1_chk02_MasterAsset_AfterUpdate  ' ** Procedure: Above.
16120       DoEvents
16130     End If
16140     If .ckgTier1_chk03_ActiveAssets.Enabled = True Then
16150       .ckgTier1_chk03_ActiveAssets = False
16160       ckgTier1_chk03_ActiveAssets_AfterUpdate  ' ** Procedure: Above.
16170       DoEvents
16180     End If
16190     If .ckgTier1_chk04_LedgerArchive.Enabled = True Then
16200       .ckgTier1_chk04_LedgerArchive = False
16210       ckgTier1_chk04_LedgerArchive_AfterUpdate  ' ** Procedure: Above.
16220       DoEvents
16230     End If
16240     If .ckgTier1_chk05_Ledger.Enabled = True Then
16250       .ckgTier1_chk05_Ledger = False
16260       ckgTier1_chk05_Ledger_AfterUpdate  ' ** Procedure: Above.
16270       DoEvents
16280     End If
16290     If .ckgTier1_chk06_Balance.Enabled = True Then
16300       .ckgTier1_chk06_Balance = False
16310       ckgTier1_chk06_Balance_AfterUpdate  ' ** Procedure: Above.
16320       DoEvents
16330     End If
16340     GetCounts 2  ' ** Procedure: Below.
16350     DoCmd.Hourglass False
16360   End With

EXITP:
16370   Exit Sub

ERRH:
1220    DoCmd.Hourglass False
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdTier1_SelectNone_KeyDown(KeyCode As Integer, Shift As Integer)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectNone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16420   intShiftDown = (Shift And acShiftMask) > 0
16430   intAltDown = (Shift And acAltMask) > 0
16440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16460     Select Case intRetVal
          Case vbKeyTab
16470       With Me
16480         intRetVal = 0
16490         .ckgTier_chkTier2.SetFocus
16500       End With
16510     End Select
16520   End If

        ' ** Shift keys.
16530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16540     Select Case intRetVal
          Case vbKeyTab
16550       With Me
16560         intRetVal = 0
16570         If .cmdTier1_SelectAll.Enabled = True Then
16580           .cmdTier1_SelectAll.SetFocus
16590         Else
16600           .ckgTier1_chk06_Balance.SetFocus
16610         End If
16620       End With
16630     End Select
16640   End If

EXITP:
16650   KeyCode = intRetVal
16660   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdTier1_SelectNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectNone_MouseMove"

16710   With Me
16720     If blnTier1_SelectNone_MouseDown = False Then
16730       Select Case blnTier1_SelectNone_Focus
            Case True
16740         .cmdTier1_SelectNone_raised_focus_dots_img.Visible = True
16750         .cmdTier1_SelectNone_raised_focus_img.Visible = False
16760       Case False
16770         .cmdTier1_SelectNone_raised_focus_img.Visible = True
16780         .cmdTier1_SelectNone_raised_focus_dots_img.Visible = False
16790       End Select
16800       .cmdTier1_SelectNone_raised_img.Visible = False
16810       .cmdTier1_SelectNone_raised_semifocus_dots_img.Visible = False
16820       .cmdTier1_SelectNone_sunken_focus_dots_img.Visible = False
16830       .cmdTier1_SelectNone_raised_img_dis.Visible = False
16840     End If
16850   End With

EXITP:
16860   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier1_SelectNone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectNone_MouseUp"

16910   With Me
16920     .cmdTier1_SelectNone_raised_focus_dots_img.Visible = True
16930     .cmdTier1_SelectNone_raised_img.Visible = False
16940     .cmdTier1_SelectNone_raised_semifocus_dots_img.Visible = False
16950     .cmdTier1_SelectNone_raised_focus_img.Visible = False
16960     .cmdTier1_SelectNone_sunken_focus_dots_img.Visible = False
16970     .cmdTier1_SelectNone_raised_img_dis.Visible = False
16980     blnTier1_SelectNone_MouseDown = False
16990   End With

EXITP:
17000   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier1_SelectNone_LostFocus()

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier1_SelectNone_LostFocus"

17110   With Me
17120     .cmdTier1_SelectNone_raised_img.Visible = True
17130     .cmdTier1_SelectNone_raised_semifocus_dots_img.Visible = False
17140     .cmdTier1_SelectNone_raised_focus_img.Visible = False
17150     .cmdTier1_SelectNone_raised_focus_dots_img.Visible = False
17160     .cmdTier1_SelectNone_sunken_focus_dots_img.Visible = False
17170     .cmdTier1_SelectNone_raised_img_dis.Visible = False
17180     blnTier1_SelectNone_Focus = False
17190   End With

EXITP:
17200   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier_chkTier2_AfterUpdate()

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier_chkTier2_AfterUpdate"

17310   With Me
17320     Select Case .ckgTier_chkTier2
          Case True
17330       .ckgTier_chkTier2_lbl2.FontBold = True
17340       .ckgTier_chkTier2_lbl2_dim_hi.FontBold = True
17350       .ckgTier_chkTier2_box.BackStyle = acBackStyleNormal
17360       .ckgTier_chkTier2_box2.BackStyle = acBackStyleNormal
17370       .ckgTier_chkTier2_hline03.BorderColor = MY_CLR_VLTBGE
17380       .ckgTier2_chk01_m_REVCODE.Enabled = True
17390       .ckgTier2_chk02_RecurringItems.Enabled = True
17400       .ckgTier2_chk03_Pricing_MasterAsset_History.Enabled = True
17410       .ckgTier2_chk04_Currency_History.Enabled = True
17420       .ckgTier2_chk05_LedgerHidden.Enabled = True
17430       .ckgTier2_chk06_Location.Enabled = True
17440       .ckgTier2_chk01_m_REVCODE_lbl_recs.ForeColor = CLR_DKGRY
17450       .ckgTier2_chk01_m_REVCODE_lbl_recs.BorderColor = CLR_VLTBLU2
17460       .ckgTier2_chk02_RecurringItems_lbl_recs.ForeColor = CLR_DKGRY
17470       .ckgTier2_chk02_RecurringItems_lbl_recs.BorderColor = CLR_VLTBLU2
17480       .ckgTier2_chk03_Pricing_MasterAsset_History_lbl_recs.ForeColor = CLR_DKGRY
17490       .ckgTier2_chk03_Pricing_MasterAsset_History_lbl_recs.BorderColor = CLR_VLTBLU2
17500       .ckgTier2_chk04_Currency_History_lbl_recs.ForeColor = CLR_DKGRY
17510       .ckgTier2_chk04_Currency_History_lbl_recs.BorderColor = CLR_VLTBLU2
17520       .ckgTier2_chk05_LedgerHidden_lbl_recs.ForeColor = CLR_DKGRY
17530       .ckgTier2_chk05_LedgerHidden_lbl_recs.BorderColor = CLR_VLTBLU2
17540       .ckgTier2_chk06_Location_lbl_recs.ForeColor = CLR_DKGRY
17550       .ckgTier2_chk06_Location_lbl_recs.BorderColor = CLR_VLTBLU2
17560       .cmdTier2_Select_box.BackStyle = acBackStyleNormal
17570       .cmdTier2_Select_hline03.BorderColor = MY_CLR_VLTBGE
17580       .cmdTier2_Select_hline04.BorderColor = MY_CLR_VLTBGE
17590       If blnEmpty = False And blnNoData = False Then
              ' ** If it's empty, they should already be disabled.
17600         .cmdTier2_SelectAll.Enabled = True
17610         .cmdTier2_SelectAll_raised_img.Visible = True
17620         .cmdTier2_SelectAll_raised_semifocus_dots_img.Visible = False
17630         .cmdTier2_SelectAll_raised_focus_img.Visible = False
17640         .cmdTier2_SelectAll_raised_focus_dots_img.Visible = False
17650         .cmdTier2_SelectAll_sunken_focus_dots_img.Visible = False
17660         .cmdTier2_SelectAll_raised_img_dis.Visible = False
17670         .cmdTier2_SelectNone.Enabled = True
17680         .cmdTier2_SelectNone_raised_img.Visible = True
17690         .cmdTier2_SelectNone_raised_semifocus_dots_img.Visible = False
17700         .cmdTier2_SelectNone_raised_focus_img.Visible = False
17710         .cmdTier2_SelectNone_raised_focus_dots_img.Visible = False
17720         .cmdTier2_SelectNone_sunken_focus_dots_img.Visible = False
17730         .cmdTier2_SelectNone_raised_img_dis.Visible = False
17740       End If
17750     Case False
17760       .ckgTier_chkTier2_lbl2.FontBold = False
17770       .ckgTier_chkTier2_lbl2_dim_hi.FontBold = False
17780       .ckgTier_chkTier2_box.BackStyle = acBackStyleTransparent
17790       .ckgTier_chkTier2_box2.BackStyle = acBackStyleTransparent
17800       .ckgTier_chkTier2_hline03.BorderColor = MY_CLR_LTBGE
17810       .ckgTier2_chk01_m_REVCODE.Enabled = False
17820       .ckgTier2_chk02_RecurringItems.Enabled = False
17830       .ckgTier2_chk03_Pricing_MasterAsset_History.Enabled = False
17840       .ckgTier2_chk04_Currency_History.Enabled = False
17850       .ckgTier2_chk05_LedgerHidden.Enabled = False
17860       .ckgTier2_chk06_Location.Enabled = False
17870       .ckgTier2_chk01_m_REVCODE_lbl_recs.ForeColor = WIN_CLR_DISF
17880       .ckgTier2_chk01_m_REVCODE_lbl_recs.BorderColor = WIN_CLR_DISR
17890       .ckgTier2_chk02_RecurringItems_lbl_recs.ForeColor = WIN_CLR_DISF
17900       .ckgTier2_chk02_RecurringItems_lbl_recs.BorderColor = WIN_CLR_DISR
17910       .ckgTier2_chk03_Pricing_MasterAsset_History_lbl_recs.ForeColor = WIN_CLR_DISF
17920       .ckgTier2_chk03_Pricing_MasterAsset_History_lbl_recs.BorderColor = WIN_CLR_DISR
17930       .ckgTier2_chk04_Currency_History_lbl_recs.ForeColor = WIN_CLR_DISF
17940       .ckgTier2_chk04_Currency_History_lbl_recs.BorderColor = WIN_CLR_DISR
17950       .ckgTier2_chk05_LedgerHidden_lbl_recs.ForeColor = WIN_CLR_DISF
17960       .ckgTier2_chk05_LedgerHidden_lbl_recs.BorderColor = WIN_CLR_DISR
17970       .ckgTier2_chk06_Location_lbl_recs.ForeColor = WIN_CLR_DISF
17980       .ckgTier2_chk06_Location_lbl_recs.BorderColor = WIN_CLR_DISR
17990       .cmdTier2_Select_box.BackStyle = acBackStyleTransparent
18000       .cmdTier2_Select_hline03.BorderColor = MY_CLR_LTBGE
18010       .cmdTier2_Select_hline04.BorderColor = MY_CLR_LTBGE
18020       If blnNoData = False Then
18030         .cmdTier2_SelectAll.Enabled = False
18040         .cmdTier2_SelectAll_raised_img_dis.Visible = True
18050         .cmdTier2_SelectAll_raised_img.Visible = False
18060         .cmdTier2_SelectAll_raised_semifocus_dots_img.Visible = False
18070         .cmdTier2_SelectAll_raised_focus_img.Visible = False
18080         .cmdTier2_SelectAll_raised_focus_dots_img.Visible = False
18090         .cmdTier2_SelectAll_sunken_focus_dots_img.Visible = False
18100         .cmdTier2_SelectNone.Enabled = False
18110         .cmdTier2_SelectNone_raised_img_dis.Visible = True
18120         .cmdTier2_SelectNone_raised_img.Visible = False
18130         .cmdTier2_SelectNone_raised_semifocus_dots_img.Visible = False
18140         .cmdTier2_SelectNone_raised_focus_img.Visible = False
18150         .cmdTier2_SelectNone_raised_focus_dots_img.Visible = False
18160         .cmdTier2_SelectNone_sunken_focus_dots_img.Visible = False
18170       End If
18180     End Select
18190     GetCounts 2  ' ** Procedure: Below.
18200     GetCounts 3  ' ** Procedure: Below.
18210     GetCounts 4  ' ** Procedure: Below.
18220   End With

EXITP:
18230   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier_chkTier2_KeyDown(KeyCode As Integer, Shift As Integer)

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier_chkTier2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18320   intShiftDown = (Shift And acShiftMask) > 0
18330   intAltDown = (Shift And acAltMask) > 0
18340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18370       With Me
18380         intRetVal = 0
18390         Select Case .ckgTier_chkTier2
              Case True
18400           .ckgTier2_chk01_m_REVCODE.SetFocus
18410         Case False
18420           .ckgTier_chkTier3.SetFocus
18430         End Select
18440       End With
18450     End Select
18460   End If

        ' ** Shift keys.
18470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18490       With Me
18500         intRetVal = 0
18510         If .cmdTier1_SelectNone.Enabled = True Then
18520           .cmdTier1_SelectNone.SetFocus
18530         ElseIf .cmdTier1_SelectAll.Enabled = True Then
18540           .cmdTier1_SelectAll.SetFocus
18550         Else
18560           .ckgTier_chkTier1.SetFocus
18570         End If
18580       End With
18590     End Select
18600   End If

EXITP:
18610   KeyCode = intRetVal
18620   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub ckgTier2_chk01_m_REVCODE_AfterUpdate()

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier2_chk01_m_REVCODE_AfterUpdate"

18710   With Me
18720     Select Case .ckgTier2_chk01_m_REVCODE
          Case True
18730       .ckgTier2_chk01_m_REVCODE_lbl.FontBold = True
18740     Case False
18750       .ckgTier2_chk01_m_REVCODE_lbl.FontBold = False
18760     End Select
18770     GetCounts 2  ' ** Procedure: Below.
18780   End With

EXITP:
18790   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier2_chk01_m_REVCODE_KeyDown(KeyCode As Integer, Shift As Integer)

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier2_chk01_m_REVCODE_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18820   intShiftDown = (Shift And acShiftMask) > 0
18830   intAltDown = (Shift And acAltMask) > 0
18840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18870       With Me
18880         intRetVal = 0
18890         .ckgTier2_chk02_RecurringItems.SetFocus
18900       End With
18910     End Select
18920   End If

        ' ** Shift keys.
18930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18950       With Me
18960         intRetVal = 0
18970         .ckgTier_chkTier2.SetFocus
18980       End With
18990     End Select
19000   End If

EXITP:
19010   KeyCode = intRetVal
19020   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub ckgTier2_chk02_RecurringItems_AfterUpdate()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier2_chk02_RecurringItems_AfterUpdate"

19110   With Me
19120     Select Case .ckgTier2_chk02_RecurringItems
          Case True
19130       .ckgTier2_chk02_RecurringItems_lbl.FontBold = True
19140     Case False
19150       .ckgTier2_chk02_RecurringItems_lbl.FontBold = False
19160     End Select
19170     GetCounts 2  ' ** Procedure: Below.
19180   End With

EXITP:
19190   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier2_chk03_Pricing_MasterAsset_History_AfterUpdate()

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier2_chk03_Pricing_MasterAsset_History_AfterUpdate"

19210   With Me
19220     Select Case .ckgTier2_chk03_Pricing_MasterAsset_History
          Case True
19230       .ckgTier2_chk03_Pricing_MasterAsset_History_lbl.FontBold = True
19240     Case False
19250       .ckgTier2_chk03_Pricing_MasterAsset_History_lbl.FontBold = False
19260     End Select
19270     GetCounts 2  ' ** Procedure: Below.
19280   End With

EXITP:
19290   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier2_chk04_Currency_History_AfterUpdate()

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier2_chk04_Currency_History_AfterUpdate"

19310   With Me
19320     Select Case .ckgTier2_chk04_Currency_History
          Case True
19330       .ckgTier2_chk04_Currency_History_lbl.FontBold = True
19340     Case False
19350       .ckgTier2_chk04_Currency_History_lbl.FontBold = False
19360     End Select
19370     GetCounts 2  ' ** Procedure: Below.
19380   End With

EXITP:
19390   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier2_chk05_LedgerHidden_AfterUpdate()

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier2_chk05_LedgerHidden_AfterUpdate"

19410   With Me
19420     Select Case .ckgTier2_chk05_LedgerHidden
          Case True
19430       .ckgTier2_chk05_LedgerHidden_lbl.FontBold = True
19440     Case False
19450       .ckgTier2_chk05_LedgerHidden_lbl.FontBold = False
19460     End Select
19470     GetCounts 2  ' ** Procedure: Below.
19480   End With

EXITP:
19490   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier2_chk06_Location_AfterUpdate()

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier2_chk06_Location_AfterUpdate"

19510   With Me
19520     Select Case .ckgTier2_chk06_Location
          Case True
19530       .ckgTier2_chk06_Location_lbl.FontBold = True
19540     Case False
19550       .ckgTier2_chk06_Location_lbl.FontBold = False
19560     End Select
19570     GetCounts 2  ' ** Procedure: Below.
19580   End With

EXITP:
19590   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier2_chk06_Location_KeyDown(KeyCode As Integer, Shift As Integer)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier2_chk06_Location_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19620   intShiftDown = (Shift And acShiftMask) > 0
19630   intAltDown = (Shift And acAltMask) > 0
19640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19670       With Me
19680         intRetVal = 0
19690         If .cmdTier2_SelectAll.Enabled = True Then
19700           .cmdTier2_SelectAll.SetFocus
19710         Else
19720           .cmdTier2_SelectNone.SetFocus
19730         End If
19740       End With
19750     End Select
19760   End If

        ' ** Shift keys.
19770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19790       With Me
19800         intRetVal = 0
19810         .ckgTier2_chk05_LedgerHidden.SetFocus
19820       End With
19830     End Select
19840   End If

EXITP:
19850   KeyCode = intRetVal
19860   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub ckgTier2_chk06_Location_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier2_chk06_Location_MouseMove"

19910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
19920   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier2_SelectAll_GotFocus()

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectAll_GotFocus"

20010   With Me
20020     blnTier2_SelectAll_Focus = True
20030     .cmdTier2_SelectAll_raised_semifocus_dots_img.Visible = True
20040     .cmdTier2_SelectAll_raised_img.Visible = False
20050     .cmdTier2_SelectAll_raised_focus_img.Visible = False
20060     .cmdTier2_SelectAll_raised_focus_dots_img.Visible = False
20070     .cmdTier2_SelectAll_sunken_focus_dots_img.Visible = False
20080     .cmdTier2_SelectAll_raised_img_dis.Visible = False
20090   End With

EXITP:
20100   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier2_SelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectAll_MouseDown"

20210   With Me
20220     blnTier2_SelectAll_MouseDown = True
20230     .cmdTier2_SelectAll_sunken_focus_dots_img.Visible = True
20240     .cmdTier2_SelectAll_raised_img.Visible = False
20250     .cmdTier2_SelectAll_raised_semifocus_dots_img.Visible = False
20260     .cmdTier2_SelectAll_raised_focus_img.Visible = False
20270     .cmdTier2_SelectAll_raised_focus_dots_img.Visible = False
20280     .cmdTier2_SelectAll_raised_img_dis.Visible = False
20290   End With

EXITP:
20300   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier2_SelectAll_Click()

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectAll_Click"

20410   With Me
20420     DoCmd.Hourglass True
20430     DoEvents
20440     If .ckgTier2_chk01_m_REVCODE.Enabled = True Then
20450       .ckgTier2_chk01_m_REVCODE = True
20460       ckgTier2_chk01_m_REVCODE_AfterUpdate  ' ** Procedure: Above.
20470       DoEvents
20480     End If
20490     If .ckgTier2_chk02_RecurringItems.Enabled = True Then
20500       .ckgTier2_chk02_RecurringItems = True
20510       ckgTier2_chk02_RecurringItems_AfterUpdate  ' ** Procedure: Above.
20520       DoEvents
20530     End If
20540     If .ckgTier2_chk03_Pricing_MasterAsset_History.Enabled = True Then
20550       .ckgTier2_chk03_Pricing_MasterAsset_History = True
20560       ckgTier2_chk03_Pricing_MasterAsset_History_AfterUpdate  ' ** Procedure: Above.
20570       DoEvents
20580     End If
20590     If .ckgTier2_chk04_Currency_History.Enabled = True Then
20600       .ckgTier2_chk04_Currency_History = True
20610       ckgTier2_chk04_Currency_History_AfterUpdate  ' ** Procedure: Above.
20620       DoEvents
20630     End If
20640     If .ckgTier2_chk05_LedgerHidden.Enabled = True Then
20650       .ckgTier2_chk05_LedgerHidden = True
20660       ckgTier2_chk05_LedgerHidden_AfterUpdate  ' ** Procedure: Above.
20670       DoEvents
20680     End If
20690     If .ckgTier2_chk06_Location.Enabled = True Then
20700       .ckgTier2_chk06_Location = True
20710       ckgTier2_chk06_Location_AfterUpdate  ' ** Procedure: Above.
20720       DoEvents
20730     End If
20740     GetCounts 2  ' ** Procedure: Below.
20750     DoCmd.Hourglass False
20760   End With

EXITP:
20770   Exit Sub

ERRH:
1220    DoCmd.Hourglass False
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdTier2_SelectAll_KeyDown(KeyCode As Integer, Shift As Integer)

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20820   intShiftDown = (Shift And acShiftMask) > 0
20830   intAltDown = (Shift And acAltMask) > 0
20840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20860     Select Case intRetVal
          Case vbKeyTab
20870       With Me
20880         intRetVal = 0
20890         If .cmdTier2_SelectNone.Enabled = True Then
20900           .cmdTier2_SelectNone.SetFocus
20910         Else
20920           .ckgTier_chkTier3.SetFocus
20930         End If
20940       End With
20950     End Select
20960   End If

        ' ** Shift keys.
20970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20980     Select Case intRetVal
          Case vbKeyTab
20990       With Me
21000         intRetVal = 0
21010         .ckgTier2_chk06_Location.SetFocus
21020       End With
21030     End Select
21040   End If

EXITP:
21050   KeyCode = intRetVal
21060   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdTier2_SelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectAll_MouseMove"

21110   With Me
21120     If blnTier2_SelectAll_MouseDown = False Then
21130       Select Case blnTier2_SelectAll_Focus
            Case True
21140         .cmdTier2_SelectAll_raised_focus_dots_img.Visible = True
21150         .cmdTier2_SelectAll_raised_focus_img.Visible = False
21160       Case False
21170         .cmdTier2_SelectAll_raised_focus_img.Visible = True
21180         .cmdTier2_SelectAll_raised_focus_dots_img.Visible = False
21190       End Select
21200       .cmdTier2_SelectAll_raised_img.Visible = False
21210       .cmdTier2_SelectAll_raised_semifocus_dots_img.Visible = False
21220       .cmdTier2_SelectAll_sunken_focus_dots_img.Visible = False
21230       .cmdTier2_SelectAll_raised_img_dis.Visible = False
21240     End If
21250   End With

EXITP:
21260   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier2_SelectAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectAll_MouseUp"

21310   With Me
21320     .cmdTier2_SelectAll_raised_focus_dots_img.Visible = True
21330     .cmdTier2_SelectAll_raised_img.Visible = False
21340     .cmdTier2_SelectAll_raised_semifocus_dots_img.Visible = False
21350     .cmdTier2_SelectAll_raised_focus_img.Visible = False
21360     .cmdTier2_SelectAll_sunken_focus_dots_img.Visible = False
21370     .cmdTier2_SelectAll_raised_img_dis.Visible = False
21380     blnTier2_SelectAll_MouseDown = False
21390   End With

EXITP:
21400   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier2_SelectAll_LostFocus()

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectAll_LostFocus"

21510   With Me
21520     .cmdTier2_SelectAll_raised_img.Visible = True
21530     .cmdTier2_SelectAll_raised_semifocus_dots_img.Visible = False
21540     .cmdTier2_SelectAll_raised_focus_img.Visible = False
21550     .cmdTier2_SelectAll_raised_focus_dots_img.Visible = False
21560     .cmdTier2_SelectAll_sunken_focus_dots_img.Visible = False
21570     .cmdTier2_SelectAll_raised_img_dis.Visible = False
21580     blnTier2_SelectAll_Focus = False
21590   End With

EXITP:
21600   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier2_SelectNone_GotFocus()

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectNone_GotFocus"

21710   With Me
21720     blnTier2_SelectNone_Focus = True
21730     .cmdTier2_SelectNone_raised_semifocus_dots_img.Visible = True
21740     .cmdTier2_SelectNone_raised_img.Visible = False
21750     .cmdTier2_SelectNone_raised_focus_img.Visible = False
21760     .cmdTier2_SelectNone_raised_focus_dots_img.Visible = False
21770     .cmdTier2_SelectNone_sunken_focus_dots_img.Visible = False
21780     .cmdTier2_SelectNone_raised_img_dis.Visible = False
21790   End With

EXITP:
21800   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier2_SelectNone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectNone_MouseDown"

21910   With Me
21920     blnTier2_SelectNone_MouseDown = True
21930     .cmdTier2_SelectNone_sunken_focus_dots_img.Visible = True
21940     .cmdTier2_SelectNone_raised_img.Visible = False
21950     .cmdTier2_SelectNone_raised_semifocus_dots_img.Visible = False
21960     .cmdTier2_SelectNone_raised_focus_img.Visible = False
21970     .cmdTier2_SelectNone_raised_focus_dots_img.Visible = False
21980     .cmdTier2_SelectNone_raised_img_dis.Visible = False
21990   End With

EXITP:
22000   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier2_SelectNone_Click()

22100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectNone_Click"

22110   With Me
22120     DoCmd.Hourglass True
22130     DoEvents
22140     If .ckgTier2_chk01_m_REVCODE.Enabled = True Then
22150       .ckgTier2_chk01_m_REVCODE = False
22160       ckgTier2_chk01_m_REVCODE_AfterUpdate  ' ** Procedure: Above.
22170       DoEvents
22180     End If
22190     If .ckgTier2_chk02_RecurringItems.Enabled = True Then
22200       .ckgTier2_chk02_RecurringItems = False
22210       ckgTier2_chk02_RecurringItems_AfterUpdate  ' ** Procedure: Above.
22220       DoEvents
22230     End If
22240     If .ckgTier2_chk03_Pricing_MasterAsset_History.Enabled = True Then
22250       .ckgTier2_chk03_Pricing_MasterAsset_History = False
22260       ckgTier2_chk03_Pricing_MasterAsset_History_AfterUpdate  ' ** Procedure: Above.
22270       DoEvents
22280     End If
22290     If .ckgTier2_chk04_Currency_History.Enabled = True Then
22300       .ckgTier2_chk04_Currency_History = False
22310       ckgTier2_chk04_Currency_History_AfterUpdate  ' ** Procedure: Above.
22320       DoEvents
22330     End If
22340     If .ckgTier2_chk05_LedgerHidden.Enabled = True Then
22350       .ckgTier2_chk05_LedgerHidden = False
22360       ckgTier2_chk05_LedgerHidden_AfterUpdate  ' ** Procedure: Above.
22370       DoEvents
22380     End If
22390     If .ckgTier2_chk06_Location.Enabled = True Then
22400       .ckgTier2_chk06_Location = False
22410       ckgTier2_chk06_Location_AfterUpdate  ' ** Procedure: Above.
22420       DoEvents
22430     End If
22440     GetCounts 2  ' ** Procedure: Below.
22450     DoCmd.Hourglass False
22460   End With

EXITP:
22470   Exit Sub

ERRH:
1220    DoCmd.Hourglass False
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdTier2_SelectNone_KeyDown(KeyCode As Integer, Shift As Integer)

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectNone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22520   intShiftDown = (Shift And acShiftMask) > 0
22530   intAltDown = (Shift And acAltMask) > 0
22540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22560     Select Case intRetVal
          Case vbKeyTab
22570       With Me
22580         intRetVal = 0
22590         .ckgTier_chkTier3.SetFocus
22600       End With
22610     End Select
22620   End If

        ' ** Shift keys.
22630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22640     Select Case intRetVal
          Case vbKeyTab
22650       With Me
22660         intRetVal = 0
22670         If .cmdTier2_SelectAll.Enabled = True Then
22680           .cmdTier2_SelectAll.SetFocus
22690         Else
22700           .ckgTier2_chk06_Location.SetFocus
22710         End If
22720       End With
22730     End Select
22740   End If

EXITP:
22750   KeyCode = intRetVal
22760   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdTier2_SelectNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectNone_MouseMove"

22810   With Me
22820     If blnTier2_SelectNone_MouseDown = False Then
22830       Select Case blnTier2_SelectNone_Focus
            Case True
22840         .cmdTier2_SelectNone_raised_focus_dots_img.Visible = True
22850         .cmdTier2_SelectNone_raised_focus_img.Visible = False
22860       Case False
22870         .cmdTier2_SelectNone_raised_focus_img.Visible = True
22880         .cmdTier2_SelectNone_raised_focus_dots_img.Visible = False
22890       End Select
22900       .cmdTier2_SelectNone_raised_img.Visible = False
22910       .cmdTier2_SelectNone_raised_semifocus_dots_img.Visible = False
22920       .cmdTier2_SelectNone_sunken_focus_dots_img.Visible = False
22930       .cmdTier2_SelectNone_raised_img_dis.Visible = False
22940     End If
22950   End With

EXITP:
22960   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier2_SelectNone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

23000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectNone_MouseUp"

23010   With Me
23020     .cmdTier2_SelectNone_raised_focus_dots_img.Visible = True
23030     .cmdTier2_SelectNone_raised_img.Visible = False
23040     .cmdTier2_SelectNone_raised_semifocus_dots_img.Visible = False
23050     .cmdTier2_SelectNone_raised_focus_img.Visible = False
23060     .cmdTier2_SelectNone_sunken_focus_dots_img.Visible = False
23070     .cmdTier2_SelectNone_raised_img_dis.Visible = False
23080     blnTier2_SelectNone_MouseDown = False
23090   End With

EXITP:
23100   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier2_SelectNone_LostFocus()

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier2_SelectNone_LostFocus"

23210   With Me
23220     .cmdTier2_SelectNone_raised_img.Visible = True
23230     .cmdTier2_SelectNone_raised_semifocus_dots_img.Visible = False
23240     .cmdTier2_SelectNone_raised_focus_img.Visible = False
23250     .cmdTier2_SelectNone_raised_focus_dots_img.Visible = False
23260     .cmdTier2_SelectNone_sunken_focus_dots_img.Visible = False
23270     .cmdTier2_SelectNone_raised_img_dis.Visible = False
23280     blnTier2_SelectNone_Focus = False
23290   End With

EXITP:
23300   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier_chkTier3_AfterUpdate()

23400 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier_chkTier3_AfterUpdate"

23410   With Me
23420     Select Case .ckgTier_chkTier3
          Case True
23430       .ckgTier_chkTier3_lbl2.FontBold = True
23440       .ckgTier_chkTier3_lbl2_dim_hi.FontBold = True
23450       .ckgTier_chkTier3_box.BackStyle = acBackStyleNormal
23460       .ckgTier_chkTier3_box2.BackStyle = acBackStyleNormal
23470       .ckgTier_chkTier3_hline03.BorderColor = MY_CLR_VLTBGE
23480       .ckgTier3_chk01_AdminOfficer.Enabled = True
23490       .ckgTier3_chk02_PortfolioModel.Enabled = True
23500       .ckgTier3_chk03_Schedule.Enabled = True
23510       .ckgTier3_chk04_Schedule_Detail.Enabled = True
23520       .ckgTier3_chk05_CheckPOSPay.Enabled = True
23530       .ckgTier3_chk06_CheckPOSPay_Detail.Enabled = True
23540       .ckgTier3_chk07_CheckReconcile_Amount.Enabled = True
23550       .ckgTier3_chk08_CheckReconcile_Item.Enabled = True
23560       .ckgTier3_chk09_CheckMemo.Enabled = True
23570       .ckgTier3_chk01_AdminOfficer_lbl_recs.ForeColor = CLR_DKGRY
23580       .ckgTier3_chk01_AdminOfficer_lbl_recs.BorderColor = CLR_VLTBLU2
23590       .ckgTier3_chk02_PortfolioModel_lbl_recs.ForeColor = CLR_DKGRY
23600       .ckgTier3_chk02_PortfolioModel_lbl_recs.BorderColor = CLR_VLTBLU2
23610       .ckgTier3_chk03_Schedule_lbl_recs.ForeColor = CLR_DKGRY
23620       .ckgTier3_chk03_Schedule_lbl_recs.BorderColor = CLR_VLTBLU2
23630       .ckgTier3_chk04_Schedule_Detail_lbl_recs.ForeColor = CLR_DKGRY
23640       .ckgTier3_chk04_Schedule_Detail_lbl_recs.BorderColor = CLR_VLTBLU2
23650       .ckgTier3_chk05_CheckPOSPay_lbl_recs.ForeColor = CLR_DKGRY
23660       .ckgTier3_chk05_CheckPOSPay_lbl_recs.BorderColor = CLR_VLTBLU2
23670       .ckgTier3_chk06_CheckPOSPay_Detail_lbl_recs.ForeColor = CLR_DKGRY
23680       .ckgTier3_chk06_CheckPOSPay_Detail_lbl_recs.BorderColor = CLR_VLTBLU2
23690       .ckgTier3_chk07_CheckReconcile_Amount_lbl_recs.ForeColor = CLR_DKGRY
23700       .ckgTier3_chk07_CheckReconcile_Amount_lbl_recs.BorderColor = CLR_VLTBLU2
23710       .ckgTier3_chk08_CheckReconcile_Item_lbl_recs.ForeColor = CLR_DKGRY
23720       .ckgTier3_chk08_CheckReconcile_Item_lbl_recs.BorderColor = CLR_VLTBLU2
23730       .ckgTier3_chk09_CheckMemo_lbl_recs.ForeColor = CLR_DKGRY
23740       .ckgTier3_chk09_CheckMemo_lbl_recs.BorderColor = CLR_VLTBLU2
23750       .cmdTier3_Select_box.BackStyle = acBackStyleNormal
23760       .cmdTier3_Select_hline03.BorderColor = MY_CLR_VLTBGE
23770       .cmdTier3_Select_hline04.BorderColor = MY_CLR_VLTBGE
23780       If blnEmpty = False And blnNoData = False Then
              ' ** If it's empty, they should already be disabled.
23790         .cmdTier3_SelectAll.Enabled = True
23800         .cmdTier3_SelectAll_raised_img.Visible = True
23810         .cmdTier3_SelectAll_raised_semifocus_dots_img.Visible = False
23820         .cmdTier3_SelectAll_raised_focus_img.Visible = False
23830         .cmdTier3_SelectAll_raised_focus_dots_img.Visible = False
23840         .cmdTier3_SelectAll_sunken_focus_dots_img.Visible = False
23850         .cmdTier3_SelectAll_raised_img_dis.Visible = False
23860         .cmdTier3_SelectNone.Enabled = True
23870         .cmdTier3_SelectNone_raised_img.Visible = True
23880         .cmdTier3_SelectNone_raised_semifocus_dots_img.Visible = False
23890         .cmdTier3_SelectNone_raised_focus_img.Visible = False
23900         .cmdTier3_SelectNone_raised_focus_dots_img.Visible = False
23910         .cmdTier3_SelectNone_sunken_focus_dots_img.Visible = False
23920         .cmdTier3_SelectNone_raised_img_dis.Visible = False
23930       End If
23940     Case False
23950       .ckgTier_chkTier3_lbl2.FontBold = False
23960       .ckgTier_chkTier3_lbl2_dim_hi.FontBold = False
23970       .ckgTier_chkTier3_box.BackStyle = acBackStyleTransparent
23980       .ckgTier_chkTier3_box2.BackStyle = acBackStyleTransparent
23990       .ckgTier_chkTier3_hline03.BorderColor = MY_CLR_LTBGE
24000       .ckgTier3_chk01_AdminOfficer.Enabled = False
24010       .ckgTier3_chk02_PortfolioModel.Enabled = False
24020       .ckgTier3_chk03_Schedule.Enabled = False
24030       .ckgTier3_chk04_Schedule_Detail.Enabled = False
24040       .ckgTier3_chk05_CheckPOSPay.Enabled = False
24050       .ckgTier3_chk06_CheckPOSPay_Detail.Enabled = False
24060       .ckgTier3_chk07_CheckReconcile_Amount.Enabled = False
24070       .ckgTier3_chk08_CheckReconcile_Item.Enabled = False
24080       .ckgTier3_chk09_CheckMemo.Enabled = False
24090       .ckgTier3_chk01_AdminOfficer_lbl_recs.ForeColor = WIN_CLR_DISF
24100       .ckgTier3_chk01_AdminOfficer_lbl_recs.BorderColor = WIN_CLR_DISR
24110       .ckgTier3_chk02_PortfolioModel_lbl_recs.ForeColor = WIN_CLR_DISF
24120       .ckgTier3_chk02_PortfolioModel_lbl_recs.BorderColor = WIN_CLR_DISR
24130       .ckgTier3_chk03_Schedule_lbl_recs.ForeColor = WIN_CLR_DISF
24140       .ckgTier3_chk03_Schedule_lbl_recs.BorderColor = WIN_CLR_DISR
24150       .ckgTier3_chk04_Schedule_Detail_lbl_recs.ForeColor = WIN_CLR_DISF
24160       .ckgTier3_chk04_Schedule_Detail_lbl_recs.BorderColor = WIN_CLR_DISR
24170       .ckgTier3_chk05_CheckPOSPay_lbl_recs.ForeColor = WIN_CLR_DISF
24180       .ckgTier3_chk05_CheckPOSPay_lbl_recs.BorderColor = WIN_CLR_DISR
24190       .ckgTier3_chk06_CheckPOSPay_Detail_lbl_recs.ForeColor = WIN_CLR_DISF
24200       .ckgTier3_chk06_CheckPOSPay_Detail_lbl_recs.BorderColor = WIN_CLR_DISR
24210       .ckgTier3_chk07_CheckReconcile_Amount_lbl_recs.ForeColor = WIN_CLR_DISF
24220       .ckgTier3_chk07_CheckReconcile_Amount_lbl_recs.BorderColor = WIN_CLR_DISR
24230       .ckgTier3_chk08_CheckReconcile_Item_lbl_recs.ForeColor = WIN_CLR_DISF
24240       .ckgTier3_chk08_CheckReconcile_Item_lbl_recs.BorderColor = WIN_CLR_DISR
24250       .ckgTier3_chk09_CheckMemo_lbl_recs.ForeColor = WIN_CLR_DISF
24260       .ckgTier3_chk09_CheckMemo_lbl_recs.BorderColor = WIN_CLR_DISR
24270       .cmdTier3_Select_box.BackStyle = acBackStyleTransparent
24280       .cmdTier3_Select_hline03.BorderColor = MY_CLR_LTBGE
24290       .cmdTier3_Select_hline04.BorderColor = MY_CLR_LTBGE
24300       If blnNoData = False Then
24310         .cmdTier3_SelectAll.Enabled = False
24320         .cmdTier3_SelectAll_raised_img_dis.Visible = True
24330         .cmdTier3_SelectAll_raised_img.Visible = False
24340         .cmdTier3_SelectAll_raised_semifocus_dots_img.Visible = False
24350         .cmdTier3_SelectAll_raised_focus_img.Visible = False
24360         .cmdTier3_SelectAll_raised_focus_dots_img.Visible = False
24370         .cmdTier3_SelectAll_sunken_focus_dots_img.Visible = False
24380         .cmdTier3_SelectNone.Enabled = False
24390         .cmdTier3_SelectNone_raised_img_dis.Visible = True
24400         .cmdTier3_SelectNone_raised_img.Visible = False
24410         .cmdTier3_SelectNone_raised_semifocus_dots_img.Visible = False
24420         .cmdTier3_SelectNone_raised_focus_img.Visible = False
24430         .cmdTier3_SelectNone_raised_focus_dots_img.Visible = False
24440         .cmdTier3_SelectNone_sunken_focus_dots_img.Visible = False
24450       End If
24460     End Select
24470     GetCounts 2  ' ** Procedure: Below.
24480     GetCounts 3  ' ** Procedure: Below.
24490     GetCounts 4  ' ** Procedure: Below.
24500   End With

EXITP:
24510   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier_chkTier3_KeyDown(KeyCode As Integer, Shift As Integer)

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier_chkTier3_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24620   intShiftDown = (Shift And acShiftMask) > 0
24630   intAltDown = (Shift And acAltMask) > 0
24640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24670       With Me
24680         intRetVal = 0
24690         Select Case .ckgTier_chkTier3
              Case True
24700           .ckgTier3_chk01_AdminOfficer.SetFocus
24710         Case False
24720           If .cmdExport.Enabled = True Then
24730             .cmdExport.SetFocus
24740           Else
24750             .cmdCancel.SetFocus
24760           End If
24770         End Select
24780       End With
24790     End Select
24800   End If

        ' ** Shift keys.
24810   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24820     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24830       With Me
24840         intRetVal = 0
24850         If .cmdTier2_SelectNone.Enabled = True Then
24860           .cmdTier2_SelectNone.SetFocus
24870         ElseIf .cmdTier2_SelectAll.Enabled = True Then
24880           .cmdTier2_SelectAll.SetFocus
24890         Else
24900           .ckgTier_chkTier2.SetFocus
24910         End If
24920       End With
24930     End Select
24940   End If

EXITP:
24950   KeyCode = intRetVal
24960   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub ckgTier3_chk01_AdminOfficer_AfterUpdate()

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier3_chk01_AdminOfficer_AfterUpdate"

25010   With Me
25020     Select Case .ckgTier3_chk01_AdminOfficer
          Case True
25030       .ckgTier3_chk01_AdminOfficer_lbl.FontBold = True
25040     Case False
25050       .ckgTier3_chk01_AdminOfficer_lbl.FontBold = False
25060     End Select
25070     GetCounts 2  ' ** Procedure: Below.
25080   End With

EXITP:
25090   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier3_chk01_AdminOfficer_KeyDown(KeyCode As Integer, Shift As Integer)

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier3_chk01_AdminOfficer_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25120   intShiftDown = (Shift And acShiftMask) > 0
25130   intAltDown = (Shift And acAltMask) > 0
25140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25170       With Me
25180         intRetVal = 0
25190         .ckgTier3_chk02_PortfolioModel.SetFocus
25200       End With
25210     End Select
25220   End If

        ' ** Shift keys.
25230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25250       With Me
25260         intRetVal = 0
25270         .ckgTier_chkTier3.SetFocus
25280       End With
25290     End Select
25300   End If

EXITP:
25310   KeyCode = intRetVal
25320   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub ckgTier3_chk02_PortfolioModel_AfterUpdate()

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier3_chk02_PortfolioModel_AfterUpdate"

25410   With Me
25420     Select Case .ckgTier3_chk02_PortfolioModel
          Case True
25430       .ckgTier3_chk02_PortfolioModel_lbl.FontBold = True
25440     Case False
25450       .ckgTier3_chk02_PortfolioModel_lbl.FontBold = False
25460     End Select
25470     GetCounts 2  ' ** Procedure: Below.
25480   End With

EXITP:
25490   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier3_chk03_Schedule_AfterUpdate()

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier3_chk03_Schedule_AfterUpdate"

25510   With Me
25520     Select Case .ckgTier3_chk03_Schedule
          Case True
25530       .ckgTier3_chk03_Schedule_lbl.FontBold = True
25540     Case False
25550       .ckgTier3_chk03_Schedule_lbl.FontBold = False
25560     End Select
25570     GetCounts 2  ' ** Procedure: Below.
25580   End With

EXITP:
25590   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier3_chk04_Schedule_Detail_AfterUpdate()

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier3_chk04_Schedule_Detail_AfterUpdate"

25610   With Me
25620     Select Case .ckgTier3_chk04_Schedule_Detail
          Case True
25630       .ckgTier3_chk04_Schedule_Detail_lbl.FontBold = True
25640     Case False
25650       .ckgTier3_chk04_Schedule_Detail_lbl.FontBold = False
25660     End Select
25670     GetCounts 2  ' ** Procedure: Below.
25680   End With

EXITP:
25690   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier3_chk05_CheckPOSPay_AfterUpdate()

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier3_chk05_CheckPOSPay_AfterUpdate"

25710   With Me
25720     Select Case .ckgTier3_chk05_CheckPOSPay
          Case True
25730       .ckgTier3_chk05_CheckPOSPay_lbl.FontBold = True
25740     Case False
25750       .ckgTier3_chk05_CheckPOSPay_lbl.FontBold = False
25760     End Select
25770     GetCounts 2  ' ** Procedure: Below.
25780   End With

EXITP:
25790   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier3_chk06_CheckPOSPay_Detail_AfterUpdate()

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier3_chk06_CheckPOSPay_Detail_AfterUpdate"

25810   With Me
25820     Select Case .ckgTier3_chk06_CheckPOSPay_Detail
          Case True
25830       .ckgTier3_chk06_CheckPOSPay_Detail_lbl.FontBold = True
25840     Case False
25850       .ckgTier3_chk06_CheckPOSPay_Detail_lbl.FontBold = False
25860     End Select
25870     GetCounts 2  ' ** Procedure: Below.
25880   End With

EXITP:
25890   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier3_chk07_CheckReconcile_Amount_AfterUpdate()

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier3_chk07_CheckReconcile_Amount_AfterUpdate"

25910   With Me
25920     Select Case .ckgTier3_chk07_CheckReconcile_Amount
          Case True
25930       .ckgTier3_chk07_CheckReconcile_Amount_lbl.FontBold = True
25940     Case False
25950       .ckgTier3_chk07_CheckReconcile_Amount_lbl.FontBold = False
25960     End Select
25970     GetCounts 2  ' ** Procedure: Below.
25980   End With

EXITP:
25990   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier3_chk08_CheckReconcile_Item_AfterUpdate()

26000 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier3_chk08_CheckReconcile_Item_AfterUpdate"

26010   With Me
26020     Select Case .ckgTier3_chk08_CheckReconcile_Item
          Case True
26030       .ckgTier3_chk08_CheckReconcile_Item_lbl.FontBold = True
26040     Case False
26050       .ckgTier3_chk08_CheckReconcile_Item_lbl.FontBold = False
26060     End Select
26070     GetCounts 2  ' ** Procedure: Below.
26080   End With

EXITP:
26090   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier3_chk09_CheckMemo_AfterUpdate()

26100 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier3_chk09_CheckMemo_AfterUpdate"

26110   With Me
26120     Select Case .ckgTier3_chk09_CheckMemo
          Case True
26130       .ckgTier3_chk09_CheckMemo_lbl.FontBold = True
26140     Case False
26150       .ckgTier3_chk09_CheckMemo_lbl.FontBold = False
26160     End Select
26170     GetCounts 2  ' ** Procedure: Below.
26180   End With

EXITP:
26190   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ckgTier3_chk09_CheckMemo_KeyDown(KeyCode As Integer, Shift As Integer)

26200 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier3_chk09_CheckMemo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26220   intShiftDown = (Shift And acShiftMask) > 0
26230   intAltDown = (Shift And acAltMask) > 0
26240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26270       With Me
26280         intRetVal = 0
26290         If .cmdTier3_SelectAll.Enabled = True Then
26300           .cmdTier3_SelectAll.SetFocus
26310         Else
26320           .cmdTier3_SelectNone.SetFocus
26330         End If
26340       End With
26350     End Select
26360   End If

        ' ** Shift keys.
26370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26390       With Me
26400         intRetVal = 0
26410         .ckgTier3_chk08_CheckReconcile_Item.SetFocus
26420       End With
26430     End Select
26440   End If

EXITP:
26450   KeyCode = intRetVal
26460   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub ckgTier3_chk09_CheckMemo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgTier3_chk09_CheckMemo_MouseMove"

26510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
26520   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier3_SelectAll_GotFocus()

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectAll_GotFocus"

26610   With Me
26620     blnTier3_SelectAll_Focus = True
26630     .cmdTier3_SelectAll_raised_semifocus_dots_img.Visible = True
26640     .cmdTier3_SelectAll_raised_img.Visible = False
26650     .cmdTier3_SelectAll_raised_focus_img.Visible = False
26660     .cmdTier3_SelectAll_raised_focus_dots_img.Visible = False
26670     .cmdTier3_SelectAll_sunken_focus_dots_img.Visible = False
26680     .cmdTier3_SelectAll_raised_img_dis.Visible = False
26690   End With

EXITP:
26700   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier3_SelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectAll_MouseDown"

26810   With Me
26820     blnTier3_SelectAll_MouseDown = True
26830     .cmdTier3_SelectAll_sunken_focus_dots_img.Visible = True
26840     .cmdTier3_SelectAll_raised_img.Visible = False
26850     .cmdTier3_SelectAll_raised_semifocus_dots_img.Visible = False
26860     .cmdTier3_SelectAll_raised_focus_img.Visible = False
26870     .cmdTier3_SelectAll_raised_focus_dots_img.Visible = False
26880     .cmdTier3_SelectAll_raised_img_dis.Visible = False
26890   End With

EXITP:
26900   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier3_SelectAll_Click()

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectAll_Click"

27010   With Me
27020     DoCmd.Hourglass True
27030     DoEvents
27040     If .ckgTier3_chk01_AdminOfficer.Enabled = True Then
27050       .ckgTier3_chk01_AdminOfficer = True
27060       ckgTier3_chk01_AdminOfficer_AfterUpdate  ' ** Procedure: Above.
27070       DoEvents
27080     End If
27090     If .ckgTier3_chk02_PortfolioModel.Enabled = True Then
27100       .ckgTier3_chk02_PortfolioModel = True
27110       ckgTier3_chk02_PortfolioModel_AfterUpdate  ' ** Procedure: Above.
27120       DoEvents
27130     End If
27140     If .ckgTier3_chk03_Schedule.Enabled = True Then
27150       .ckgTier3_chk03_Schedule = True
27160       ckgTier3_chk03_Schedule_AfterUpdate  ' ** Procedure: Above.
27170       DoEvents
27180     End If
27190     If .ckgTier3_chk04_Schedule_Detail.Enabled = True Then
27200       .ckgTier3_chk04_Schedule_Detail = True
27210       ckgTier3_chk04_Schedule_Detail_AfterUpdate  ' ** Procedure: Above.
27220       DoEvents
27230     End If
27240     If .ckgTier3_chk05_CheckPOSPay.Enabled = True Then
27250       .ckgTier3_chk05_CheckPOSPay = True
27260       ckgTier3_chk05_CheckPOSPay_AfterUpdate  ' ** Procedure: Above.
27270       DoEvents
27280     End If
27290     If .ckgTier3_chk06_CheckPOSPay_Detail.Enabled = True Then
27300       .ckgTier3_chk06_CheckPOSPay_Detail = True
27310       ckgTier3_chk06_CheckPOSPay_Detail_AfterUpdate  ' ** Procedure: Above.
27320       DoEvents
27330     End If
27340     If .ckgTier3_chk07_CheckReconcile_Amount.Enabled = True Then
27350       .ckgTier3_chk07_CheckReconcile_Amount = True
27360       ckgTier3_chk07_CheckReconcile_Amount_AfterUpdate  ' ** Procedure: Above.
27370       DoEvents
27380     End If
27390     If .ckgTier3_chk08_CheckReconcile_Item.Enabled = True Then
27400       .ckgTier3_chk08_CheckReconcile_Item = True
27410       ckgTier3_chk08_CheckReconcile_Item_AfterUpdate  ' ** Procedure: Above.
27420       DoEvents
27430     End If
27440     If .ckgTier3_chk09_CheckMemo.Enabled = True Then
27450       .ckgTier3_chk09_CheckMemo = True
27460       ckgTier3_chk09_CheckMemo_AfterUpdate  ' ** Procedure: Above.
27470       DoEvents
27480     End If
27490     GetCounts 2  ' ** Procedure: Below.
27500     DoCmd.Hourglass False
27510   End With

EXITP:
27520   Exit Sub

ERRH:
1220    DoCmd.Hourglass False
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdTier3_SelectAll_KeyDown(KeyCode As Integer, Shift As Integer)

27600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27620   intShiftDown = (Shift And acShiftMask) > 0
27630   intAltDown = (Shift And acAltMask) > 0
27640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27660     Select Case intRetVal
          Case vbKeyTab
27670       With Me
27680         intRetVal = 0
27690         If .cmdTier3_SelectNone.Enabled = True Then
27700           .cmdTier3_SelectNone.SetFocus
27710         ElseIf .cmdExport.Enabled = True Then
27720           .cmdExport.SetFocus
27730         Else
27740           .cmdCancel.SetFocus
27750         End If
27760       End With
27770     End Select
27780   End If

        ' ** Shift keys.
27790   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27800     Select Case intRetVal
          Case vbKeyTab
27810       With Me
27820         intRetVal = 0
27830         .ckgTier3_chk09_CheckMemo.SetFocus
27840       End With
27850     End Select
27860   End If

EXITP:
27870   KeyCode = intRetVal
27880   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdTier3_SelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectAll_MouseMove"

27910   With Me
27920     If blnTier3_SelectAll_MouseDown = False Then
27930       Select Case blnTier3_SelectAll_Focus
            Case True
27940         .cmdTier3_SelectAll_raised_focus_dots_img.Visible = True
27950         .cmdTier3_SelectAll_raised_focus_img.Visible = False
27960       Case False
27970         .cmdTier3_SelectAll_raised_focus_img.Visible = True
27980         .cmdTier3_SelectAll_raised_focus_dots_img.Visible = False
27990       End Select
28000       .cmdTier3_SelectAll_raised_img.Visible = False
28010       .cmdTier3_SelectAll_raised_semifocus_dots_img.Visible = False
28020       .cmdTier3_SelectAll_sunken_focus_dots_img.Visible = False
28030       .cmdTier3_SelectAll_raised_img_dis.Visible = False
28040     End If
28050   End With

EXITP:
28060   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier3_SelectAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

28100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectAll_MouseUp"

28110   With Me
28120     .cmdTier3_SelectAll_raised_focus_dots_img.Visible = True
28130     .cmdTier3_SelectAll_raised_img.Visible = False
28140     .cmdTier3_SelectAll_raised_semifocus_dots_img.Visible = False
28150     .cmdTier3_SelectAll_raised_focus_img.Visible = False
28160     .cmdTier3_SelectAll_sunken_focus_dots_img.Visible = False
28170     .cmdTier3_SelectAll_raised_img_dis.Visible = False
28180     blnTier3_SelectAll_MouseDown = False
28190   End With

EXITP:
28200   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier3_SelectAll_LostFocus()

28300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectAll_LostFocus"

28310   With Me
28320     .cmdTier3_SelectAll_raised_img.Visible = True
28330     .cmdTier3_SelectAll_raised_semifocus_dots_img.Visible = False
28340     .cmdTier3_SelectAll_raised_focus_img.Visible = False
28350     .cmdTier3_SelectAll_raised_focus_dots_img.Visible = False
28360     .cmdTier3_SelectAll_sunken_focus_dots_img.Visible = False
28370     .cmdTier3_SelectAll_raised_img_dis.Visible = False
28380     blnTier3_SelectAll_Focus = False
28390   End With

EXITP:
28400   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier3_SelectNone_GotFocus()

28500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectNone_GotFocus"

28510   With Me
28520     blnTier3_SelectNone_Focus = True
28530     .cmdTier3_SelectNone_raised_semifocus_dots_img.Visible = True
28540     .cmdTier3_SelectNone_raised_img.Visible = False
28550     .cmdTier3_SelectNone_raised_focus_img.Visible = False
28560     .cmdTier3_SelectNone_raised_focus_dots_img.Visible = False
28570     .cmdTier3_SelectNone_sunken_focus_dots_img.Visible = False
28580     .cmdTier3_SelectNone_raised_img_dis.Visible = False
28590   End With

EXITP:
28600   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier3_SelectNone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectNone_MouseDown"

28710   With Me
28720     blnTier3_SelectNone_MouseDown = True
28730     .cmdTier3_SelectNone_sunken_focus_dots_img.Visible = True
28740     .cmdTier3_SelectNone_raised_img.Visible = False
28750     .cmdTier3_SelectNone_raised_semifocus_dots_img.Visible = False
28760     .cmdTier3_SelectNone_raised_focus_img.Visible = False
28770     .cmdTier3_SelectNone_raised_focus_dots_img.Visible = False
28780     .cmdTier3_SelectNone_raised_img_dis.Visible = False
28790   End With

EXITP:
28800   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier3_SelectNone_Click()

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectNone_Click"

28910   With Me
28920     DoCmd.Hourglass True
28930     DoEvents
28940     If .ckgTier3_chk01_AdminOfficer.Enabled = True Then
28950       .ckgTier3_chk01_AdminOfficer = False
28960       ckgTier3_chk01_AdminOfficer_AfterUpdate  ' ** Procedure: Above.
28970       DoEvents
28980     End If
28990     If .ckgTier3_chk02_PortfolioModel.Enabled = True Then
29000       .ckgTier3_chk02_PortfolioModel = False
29010       ckgTier3_chk02_PortfolioModel_AfterUpdate  ' ** Procedure: Above.
29020       DoEvents
29030     End If
29040     If .ckgTier3_chk03_Schedule.Enabled = True Then
29050       .ckgTier3_chk03_Schedule = False
29060       ckgTier3_chk03_Schedule_AfterUpdate  ' ** Procedure: Above.
29070       DoEvents
29080     End If
29090     If .ckgTier3_chk04_Schedule_Detail.Enabled = True Then
29100       .ckgTier3_chk04_Schedule_Detail = False
29110       ckgTier3_chk04_Schedule_Detail_AfterUpdate  ' ** Procedure: Above.
29120       DoEvents
29130     End If
29140     If .ckgTier3_chk05_CheckPOSPay.Enabled = True Then
29150       .ckgTier3_chk05_CheckPOSPay = False
29160       ckgTier3_chk05_CheckPOSPay_AfterUpdate  ' ** Procedure: Above.
29170       DoEvents
29180     End If
29190     If .ckgTier3_chk06_CheckPOSPay_Detail.Enabled = True Then
29200       .ckgTier3_chk06_CheckPOSPay_Detail = False
29210       ckgTier3_chk06_CheckPOSPay_Detail_AfterUpdate  ' ** Procedure: Above.
29220       DoEvents
29230     End If
29240     If .ckgTier3_chk07_CheckReconcile_Amount.Enabled = True Then
29250       .ckgTier3_chk07_CheckReconcile_Amount = False
29260       ckgTier3_chk07_CheckReconcile_Amount_AfterUpdate  ' ** Procedure: Above.
29270       DoEvents
29280     End If
29290     If .ckgTier3_chk08_CheckReconcile_Item.Enabled = True Then
29300       .ckgTier3_chk08_CheckReconcile_Item = False
29310       ckgTier3_chk08_CheckReconcile_Item_AfterUpdate  ' ** Procedure: Above.
29320       DoEvents
29330     End If
29340     If .ckgTier3_chk09_CheckMemo.Enabled = True Then
29350       .ckgTier3_chk09_CheckMemo = False
29360       ckgTier3_chk09_CheckMemo_AfterUpdate  ' ** Procedure: Above.
29370       DoEvents
29380     End If
29390     GetCounts 2  ' ** Procedure: Below.
29400     DoCmd.Hourglass False
29410   End With

EXITP:
29420   Exit Sub

ERRH:
1220    DoCmd.Hourglass False
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdTier3_SelectNone_KeyDown(KeyCode As Integer, Shift As Integer)

29500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectNone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29520   intShiftDown = (Shift And acShiftMask) > 0
29530   intAltDown = (Shift And acAltMask) > 0
29540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29560     Select Case intRetVal
          Case vbKeyTab
29570       With Me
29580         intRetVal = 0
29590         If .cmdExport.Enabled = True Then
29600           .cmdExport.SetFocus
29610         Else
29620           .cmdCancel.SetFocus
29630         End If
29640       End With
29650     End Select
29660   End If

        ' ** Shift keys.
29670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29680     Select Case intRetVal
          Case vbKeyTab
29690       With Me
29700         intRetVal = 0
29710         If .cmdTier3_SelectAll.Enabled = True Then
29720           .cmdTier3_SelectAll.SetFocus
29730         Else
29740           .ckgTier3_chk09_CheckMemo.SetFocus
29750         End If
29760       End With
29770     End Select
29780   End If

EXITP:
29790   KeyCode = intRetVal
29800   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdTier3_SelectNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectNone_MouseMove"

29910   With Me
29920     If blnTier3_SelectNone_MouseDown = False Then
29930       Select Case blnTier3_SelectNone_Focus
            Case True
29940         .cmdTier3_SelectNone_raised_focus_dots_img.Visible = True
29950         .cmdTier3_SelectNone_raised_focus_img.Visible = False
29960       Case False
29970         .cmdTier3_SelectNone_raised_focus_img.Visible = True
29980         .cmdTier3_SelectNone_raised_focus_dots_img.Visible = False
29990       End Select
30000       .cmdTier3_SelectNone_raised_img.Visible = False
30010       .cmdTier3_SelectNone_raised_semifocus_dots_img.Visible = False
30020       .cmdTier3_SelectNone_sunken_focus_dots_img.Visible = False
30030       .cmdTier3_SelectNone_raised_img_dis.Visible = False
30040     End If
30050   End With

EXITP:
30060   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier3_SelectNone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

30100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectNone_MouseUp"

30110   With Me
30120     .cmdTier3_SelectNone_raised_focus_dots_img.Visible = True
30130     .cmdTier3_SelectNone_raised_img.Visible = False
30140     .cmdTier3_SelectNone_raised_semifocus_dots_img.Visible = False
30150     .cmdTier3_SelectNone_raised_focus_img.Visible = False
30160     .cmdTier3_SelectNone_sunken_focus_dots_img.Visible = False
30170     .cmdTier3_SelectNone_raised_img_dis.Visible = False
30180     blnTier3_SelectNone_MouseDown = False
30190   End With

EXITP:
30200   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdTier3_SelectNone_LostFocus()

30300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTier3_SelectNone_LostFocus"

30310   With Me
30320     .cmdTier3_SelectNone_raised_img.Visible = True
30330     .cmdTier3_SelectNone_raised_semifocus_dots_img.Visible = False
30340     .cmdTier3_SelectNone_raised_focus_img.Visible = False
30350     .cmdTier3_SelectNone_raised_focus_dots_img.Visible = False
30360     .cmdTier3_SelectNone_sunken_focus_dots_img.Visible = False
30370     .cmdTier3_SelectNone_raised_img_dis.Visible = False
30380     blnTier3_SelectNone_Focus = False
30390   End With

EXITP:
30400   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub AllData_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30500 On Error GoTo ERRH

        Const THIS_PROC As String = "AllData_box_MouseMove"

30510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
30520   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub AllData_Path_AfterUpdate()

30600 On Error GoTo ERRH

        Const THIS_PROC As String = "AllData_Path_AfterUpdate"

        Dim strPath As String

30610   With Me
30620     If IsNull(.AllData_Path) = False Then
30630       strPath = Trim(.AllData_Path)
30640       If strPath <> vbNullString Then
30650         If .AllData_Path.Enabled = True Then
30660           .AllData_Path.SetFocus
30670           .AllData_Path.SelLength = 0
30680           .AllData_Path.SelStart = 99
30690         End If
30700       End If
30710     End If
30720   End With

EXITP:
30730   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub AllData_Path_KeyDown(KeyCode As Integer, Shift As Integer)

30800 On Error GoTo ERRH

        Const THIS_PROC As String = "AllData_Path_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
30820   intShiftDown = (Shift And acShiftMask) > 0
30830   intAltDown = (Shift And acAltMask) > 0
30840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30860     Select Case intRetVal
          Case vbKeyTab
30870       With Me
30880         intRetVal = 0

30890       End With
30900     End Select
30910   End If

        ' ** Shift keys.
30920   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
30930     Select Case intRetVal
          Case vbKeyTab
30940       With Me
30950         intRetVal = 0

30960       End With
30970     End Select
30980   End If

EXITP:
30990   KeyCode = intRetVal
31000   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub AllData_Path_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

31100 On Error GoTo ERRH

        Const THIS_PROC As String = "AllData_Path_MouseMove"

31110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
31120   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdBrowse2_GotFocus()

31200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse2_GotFocus"

31210   With Me
31220     blnBrowse2_Focus = True
31230     .cmdBrowse2_raised_semifocus_dots_img.Visible = True
31240     .cmdBrowse2_raised_img.Visible = False
31250     .cmdBrowse2_raised_focus_img.Visible = False
31260     .cmdBrowse2_raised_focus_dots_img.Visible = False
31270     .cmdBrowse2_sunken_focus_dots_img.Visible = False
31280     .cmdBrowse2_raised_img_dis.Visible = False
31290   End With

EXITP:
31300   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdBrowse2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

31400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse2_MouseDown"

31410   With Me
31420     blnBrowse2_MouseDown = True
31430     .cmdBrowse2_sunken_focus_dots_img.Visible = True
31440     .cmdBrowse2_raised_img.Visible = False
31450     .cmdBrowse2_raised_semifocus_dots_img.Visible = False
31460     .cmdBrowse2_raised_focus_img.Visible = False
31470     .cmdBrowse2_raised_focus_dots_img.Visible = False
31480     .cmdBrowse2_raised_img_dis.Visible = False
31490   End With

EXITP:
31500   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdBrowse2_Click()
' ** Use Windows API for returning a path but no file name.

31600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse2_Click"

        Dim strTitle1 As String
        Dim strPath As String

31610   With Me
31620     strTitle1 = "Select a folder"
31630     strPath = GetFolderPathSIS(strTitle1)  ' ** Module Function: modBrowseFilesAndFolders.
31640     If strPath <> vbNullString Then
31650       .AllData_Path = strPath
31660       DoEvents
31670       AllData_Path_AfterUpdate  ' ** Procedure: Above.
31680     End If
31690     gblnSetFocus = True
31700     .TimerInterval = 100&
31710   End With

EXITP:
31720   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdBrowse2_KeyDown(KeyCode As Integer, Shift As Integer)

31800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

31810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
31820   intShiftDown = (Shift And acShiftMask) > 0
31830   intAltDown = (Shift And acAltMask) > 0
31840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
31850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31860     Select Case intRetVal
          Case vbKeyTab
31870       With Me
31880         intRetVal = 0
31890         If .cmdExport.Enabled = True Then
31900           .cmdExport.SetFocus
31910         Else
31920           .cmdCancel.SetFocus
31930         End If
31940       End With
31950     End Select
31960   End If

        ' ** Shift keys.
31970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31980     Select Case intRetVal
          Case vbKeyTab
31990       With Me
32000         intRetVal = 0
32010         .AllData_Path.SetFocus
32020       End With
32030     End Select
32040   End If

EXITP:
32050   KeyCode = intRetVal
32060   Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub cmdBrowse2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

32100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse2_MouseMove"

32110   With Me
32120     If blnBrowse2_MouseDown = False Then
32130       Select Case blnBrowse2_Focus
            Case True
32140         .cmdBrowse2_raised_focus_dots_img.Visible = True
32150         .cmdBrowse2_raised_focus_img.Visible = False
32160       Case False
32170         .cmdBrowse2_raised_focus_img.Visible = True
32180         .cmdBrowse2_raised_focus_dots_img.Visible = False
32190       End Select
32200       .cmdBrowse2_raised_img.Visible = False
32210       .cmdBrowse2_raised_semifocus_dots_img.Visible = False
32220       .cmdBrowse2_sunken_focus_dots_img.Visible = False
32230       .cmdBrowse2_raised_img_dis.Visible = False
32240     End If
32250   End With

EXITP:
32260   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdBrowse2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

32300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse2_MouseUp"

32310   With Me
32320     .cmdBrowse2_raised_focus_dots_img.Visible = True
32330     .cmdBrowse2_raised_img.Visible = False
32340     .cmdBrowse2_raised_semifocus_dots_img.Visible = False
32350     .cmdBrowse2_raised_focus_img.Visible = False
32360     .cmdBrowse2_sunken_focus_dots_img.Visible = False
32370     .cmdBrowse2_raised_img_dis.Visible = False
32380     blnBrowse2_MouseDown = False
32390   End With

EXITP:
32400   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub cmdBrowse2_LostFocus()

32500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse2_LostFocus"

32510   With Me
32520     .cmdBrowse2_raised_img.Visible = True
32530     .cmdBrowse2_raised_semifocus_dots_img.Visible = False
32540     .cmdBrowse2_raised_focus_img.Visible = False
32550     .cmdBrowse2_raised_focus_dots_img.Visible = False
32560     .cmdBrowse2_sunken_focus_dots_img.Visible = False
32570     .cmdBrowse2_raised_img_dis.Visible = False
32580     blnBrowse2_Focus = False
32590   End With

EXITP:
32600   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub GetCounts(intMode As Integer)

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "GetCounts"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, ctl As Access.Control
        Dim blnFound As Boolean, blnDefault As Boolean
        Dim lngCnt As Long
        Dim strTmp01 As String, strTmp02 As String
        Dim lngX As Long, lngY As Long

32710   With Me
32720     Set dbs = CurrentDb
32730     Select Case intMode
          Case 1  ' ** Load check box and table arrays.

            ' ** tblForm_Control, just 'frmAccountExport', table acCheckBox.
32740       Set qdf = dbs.QueryDefs("qryAccountExport_04_01")
32750       Set rst = qdf.OpenRecordset
32760       With rst
32770         .MoveLast
32780         lngCtls = .RecordCount
32790         .MoveFirst
32800         arr_varCtl = .GetRows(lngCtls)
              ' ************************************************
              ' ** Array: arr_varCtl()
              ' **
              ' **   Field  Element  Name           Constant
              ' **   =====  =======  =============  ==========
              ' **     1       0     dbs_id         C_DID
              ' **     2       1     frm_id         C_FID
              ' **     3       2     ctl_id         C_CID
              ' **     4       3     export_tier    C_TIER
              ' **     5       4     export_item    C_ITEM
              ' **     6       5     ctl_name       C_CNAM
              ' **     7       6     ctl_value      C_VAL
              ' **
              ' ************************************************
32810         .Close
32820       End With
32830       Set rst = Nothing
32840       Set qdf = Nothing

            ' ** qryAccountExport_05_04 (qryAccountExport_05_02 (qryAccountExport_05_01 (tblDatabase_Table,
            ' ** just tables to export), linked to qryAccountExport_03_01 (qryImportExport_Specifications_01
            ' ** (MSysIMEXSpecs, just SpecID, SpecName, FileType), just 'TACloud_.._Export_Specification',
            ' ** with len_name, tbl_name)), linked to qryAccountExport_05_03 (tblQuery, just 'Export'
            ' ** queries, with SpecName)), just needed fields.
32850       Set qdf = dbs.QueryDefs("qryAccountExport_05_05")
32860       Set rst = qdf.OpenRecordset
32870       With rst
32880         .MoveLast
32890         lngRecs = .RecordCount
32900         .MoveFirst
32910         arr_varRec = .GetRows(lngRecs)
              ' *************************************************
              ' ** Array: arr_varRec()
              ' **
              ' **   Field  Element  Name            Constant
              ' **   =====  =======  ==============  ==========
              ' **     1       0     dbs_id_asof     R_DID
              ' **     2       1     SpecID          R_SID
              ' **     3       2     SpecName        R_SNAM
              ' **     4       3     export_tier     R_TIER
              ' **     5       4     export_item     R_ITEM
              ' **     6       5     tbllnk_name     R_TNAM
              ' **     7       6     qry_name        R_QNAM
              ' **     8       7     reccnt_count    R_CNT
              ' **     9       8     def_tbl         R_DEF
              ' **    10       9     def_cnt         R_DCNT
              ' **    11      10     DoExport        R_EXP
              ' **
              ' *************************************************
32920         .Close
32930       End With
32940       Set rst = Nothing
32950       Set qdf = Nothing

32960     Case 2  ' ** Count of tables to export.

32970       For lngX = 0& To (lngCtls - 1&)
32980         arr_varCtl(C_VAL, lngX) = CBool(False)
32990       Next  ' ** lngX.

33000       For lngX = 0& To (lngRecs - 1&)
33010         arr_varRec(R_EXP, lngX) = CBool(False)
33020       Next  ' ** lngX.

33030       For lngX = 0& To (lngCtls - 1&)
33040         Select Case arr_varCtl(C_TIER, lngX)
              Case 1&
33050           If .ckgTier_chkTier1 = True Then
33060             arr_varCtl(C_VAL, lngX) = .Controls(arr_varCtl(C_CNAM, lngX))
33070           End If
33080         Case 2&
33090           If .ckgTier_chkTier2 = True Then
33100             arr_varCtl(C_VAL, lngX) = .Controls(arr_varCtl(C_CNAM, lngX))
33110           End If
33120         Case 3&
33130           If .ckgTier_chkTier3 = True Then
33140             arr_varCtl(C_VAL, lngX) = .Controls(arr_varCtl(C_CNAM, lngX))
33150           End If
33160         End Select
33170       Next  ' ** lngX.

33180       lngCnt = 0&
33190       For lngX = 0& To (lngCtls - 1&)
33200         If arr_varCtl(C_VAL, lngX) = True Then
33210           lngCnt = lngCnt + 1&
33220           For lngY = 0& To (lngRecs - 1&)
33230             If arr_varRec(R_TIER, lngY) = arr_varCtl(C_TIER, lngX) And arr_varRec(R_ITEM, lngY) = arr_varCtl(C_ITEM, lngX) Then
33240               arr_varRec(R_EXP, lngY) = CBool(True)
33250               Exit For
33260             End If
33270           Next  ' ** lngY.
33280         End If
33290       Next  ' ** lngX.
33300       .AllData_Tbls = lngCnt

33310     Case 3  ' ** Count of records in each table.

33320       blnFound = False
33330       For lngX = 0& To (lngRecs - 1&)
33340         If arr_varRec(R_CNT, lngX) > 0& Then
33350           blnFound = True
33360           Exit For
33370         End If
33380       Next

33390       If blnFound = False Then
33400         For lngX = 0& To (lngRecs - 1&)
33410           Set qdf = dbs.QueryDefs(arr_varRec(R_QNAM, lngX))
33420           Set rst = qdf.OpenRecordset
33430           With rst
33440             If .BOF = True And .EOF = True Then
33450               arr_varRec(R_CNT, lngX) = 0&
33460             Else
33470               .MoveLast
33480               arr_varRec(R_CNT, lngX) = .RecordCount
33490             End If
33500             .Close
33510           End With
33520           Set rst = Nothing
33530           Set qdf = Nothing
33540         Next
33550         blnRecalcTots = True
33560       End If  ' ** blnFound.

33570     Case 4&   ' ** Display counts.

33580       For lngX = 0& To (lngRecs - 1&)
33590         For lngY = 0& To (lngCtls - 1&)
33600           If arr_varCtl(C_TIER, lngY) = arr_varRec(R_TIER, lngX) And arr_varCtl(C_ITEM, lngY) = arr_varRec(R_ITEM, lngX) Then
33610             strTmp01 = arr_varCtl(C_CNAM, lngY) & "_lbl_recs"
33620             strTmp02 = Format(arr_varRec(R_CNT, lngX), "#,##0")
33630             .Controls(strTmp01).Caption = strTmp02
33640           End If
33650         Next
33660       Next

            ' ** Disable zeroes.  (I think I got a little carried away!)
33670       For lngX = 0& To (lngRecs - 1&)

33680         blnFound = True: blnDefault = False
33690         If arr_varRec(R_CNT, lngX) = 0 Then
33700           blnFound = False
33710         Else
                ' ** Disable if only defaults present.
33720           If (arr_varRec(R_TNAM, lngX) = "masterasset" Or arr_varRec(R_TNAM, lngX) = "Location" Or _
                    arr_varRec(R_TNAM, lngX) = "adminofficer" Or arr_varRec(R_TNAM, lngX) = "tblCheckMemo") And _
                    arr_varRec(R_CNT, lngX) = 1& Then
                  ' ** MasterAsset, Location, AdminOfficer, tblCheckMemo.
33730             blnFound = False
33740             blnDefault = True
33750           ElseIf (arr_varRec(R_TNAM, lngX) = "account" Or arr_varRec(R_TNAM, lngX) = "Balance" Or _
                    arr_varRec(R_TNAM, lngX) = "RecurringItems") And arr_varRec(R_CNT, lngX) = 2& Then
                  ' ** Account, Balance, RecurringItems.
33760             blnFound = False
33770             blnDefault = True
33780           ElseIf arr_varRec(R_TNAM, lngX) = "m_REVCODE" And arr_varRec(R_CNT, lngX) = 6& Then
                  ' ** m_REVCODE.
33790             blnFound = False
33800             blnDefault = True
33810           ElseIf arr_varRec(R_TNAM, lngX) = "tblCurrency_History" And arr_varRec(R_CNT, lngX) = 1388& Then
                  ' ** tblCurrency_History.
33820             blnFound = False
33830             blnDefault = True
33840           End If
33850         End If
              ' ** "masterasset","Location","adminofficer","tblCheckMemo","account","Balance","RecurringItems","m_REVCODE","tblCurrency_History"

33860         If blnFound = False Then
33870           For lngY = 0& To (lngCtls - 1&)
33880             If arr_varCtl(C_TIER, lngY) = arr_varRec(R_TIER, lngX) And arr_varCtl(C_ITEM, lngY) = arr_varRec(R_ITEM, lngX) Then
33890               If .Controls(arr_varCtl(C_CNAM, lngX)) = True Then
33900                 .Controls(arr_varCtl(C_CNAM, lngX)) = False
33910               End If
33920               If .Controls(arr_varCtl(C_CNAM, lngX)).Enabled = True Then
33930                 .Controls(arr_varCtl(C_CNAM, lngX)).Enabled = False
33940               End If
33950               Select Case .opgExport
                    Case .opgExport_optAcctProf.OptionValue
                      ' ** Whole section is disabled.
33960                 .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").ForeColor = 13749193  ' ** Lighter disabled text.
33970                 .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").BorderColor = 14407123  ' ** Lighter disabled border.
33980               Case .opgExport_optAll.OptionValue
                      ' ** Whole section is enabled.
33990                 Select Case arr_varRec(R_TIER, lngX)
                      Case 1&
34000                   Select Case .ckgTier_chkTier1
                        Case True
                          ' ** Would be enabled.
34010                     .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").ForeColor = CLR_VLTROSE
34020                     .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").BorderColor = CLR_VLTBLU3
34030                   Case False
                          ' ** Would be disabled.
34040                     .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").ForeColor = 13749193  ' ** Lighter disabled text.
34050                     .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").BorderColor = 14407123  ' ** Lighter disabled border.
34060                   End Select
34070                 Case 2&
34080                   Select Case .ckgTier_chkTier2
                        Case True
                          ' ** Would be enabled.
34090                     .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").ForeColor = CLR_VLTROSE
34100                     .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").BorderColor = CLR_VLTBLU3
34110                   Case False
                          ' ** Would be disabled.
34120                     .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").ForeColor = 13749193  ' ** Lighter disabled text.
34130                     .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").BorderColor = 14407123  ' ** Lighter disabled border.
34140                   End Select
34150                 Case 3&
34160                   Select Case .ckgTier_chkTier3
                        Case True
                          ' ** Would be enabled.
34170                     .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").ForeColor = CLR_VLTROSE
34180                     .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").BorderColor = CLR_VLTBLU3
34190                   Case False
                          ' ** Would be disabled.
34200                     .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").ForeColor = 13749193  ' ** Lighter disabled text.
34210                     .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_recs").BorderColor = 14407123  ' ** Lighter disabled border.
34220                   End Select
34230                 End Select
34240               End Select
34250               If blnDefault = True And blnNoData = False Then
34260                 .Controls(arr_varCtl(C_CNAM, lngX) & "_lbl_def").Visible = True
34270                 .DefMsg_lbl1.Visible = True
34280                 .DefMsg_lbl2.Visible = True
34290               End If
34300               Exit For
34310             End If
34320           Next  ' **  lngY.
34330         End If  ' ** blnFound.

34340       Next  ' ** lngX.

34350     Case 5&  ' ** Clear asterisks.

34360       For Each ctl In .Detail.Controls
34370         With ctl
34380           If Right(.Name, 8) = "_lbl_def" Then
34390             .Visible = False
34400           End If
34410         End With
34420         .DefMsg_lbl1.Visible = False
34430         .DefMsg_lbl2.Visible = False
34440       Next

34450     End Select
34460     dbs.Close
34470   End With

        ' ** CLR_VLTBLU3  '15986665  ' ** Very, very light blue border.
        ' ** CLR_VLTROSE  '9342639   ' ** Very light rose zero.
        ' ** 13749193  ' ** Lighter disabled text.
        ' ** 14407123  ' ** Lighter disabled border.

EXITP:
34480   Set ctl = Nothing
34490   Set rst = Nothing
34500   Set qdf = Nothing
34510   Set dbs = Nothing
34520   Exit Sub

ERRH:
1220    DoCmd.Hourglass False
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub ExportToCloud()

34600 On Error GoTo ERRH

        Const THIS_PROC As String = "ExportToCloud"

        Dim strTargetPath As String, strTargetFile As String, strTargetPathFile As String
        Dim blnContinue As Boolean
        Dim lngCnt As Long
        Dim strTmp01 As String
        Dim lngX As Long, lngY As Long

34610   With Me
34620     If blnNoData = False Then

34630       DoCmd.Hourglass True
34640       DoEvents

34650       blnContinue = True

            ' ** Target path must be filled.
34660       If IsNull(.AllData_Path) = True Then
34670         blnContinue = False
34680         DoCmd.Hourglass False
34690         MsgBox "A path is required to save the exported files.", vbInformation + vbOKOnly, "Entry Required"
34700       Else
34710         If Trim(.AllData_Path) = vbNullString Then
34720           blnContinue = False
34730           DoCmd.Hourglass False
34740           MsgBox "A path is required to save the exported files.", vbInformation + vbOKOnly, "Entry Required"
34750         Else
34760           strTargetPath = IIf(Right(Trim(.AllData_Path), 1) <> LNK_SEP, Trim(.AllData_Path), _
                  Left(Trim(.AllData_Path), Len(Trim(.AllData_Path)) - 1))
34770           If DirExists(strTargetPath) = False Then  ' ** Module Function: modFileUtilities.
34780             blnContinue = False
34790             DoCmd.Hourglass False
34800             MsgBox "The folder cannot be found." & vbCrLf & vbCrLf & strTargetPath, vbInformation + vbOKOnly, "Invalid Entry"
34810           End If
34820         End If
34830       End If

            ' *************************************************
            ' ** Array: arr_varRec()
            ' **
            ' **   Field  Element  Name            Constant
            ' **   =====  =======  ==============  ==========
            ' **     1       0     dbs_id_asof     R_DID
            ' **     2       1     SpecID          R_SID
            ' **     3       2     SpecName        R_SNAM
            ' **     4       3     export_tier     R_TIER
            ' **     5       4     export_item     R_ITEM
            ' **     6       5     tbllnk_name     R_TNAM
            ' **     7       6     qry_name        R_QNAM
            ' **     8       7     reccnt_count    R_CNT
            ' **     9       8     DoExport        R_EXP
            ' **
            ' *************************************************

34840       If blnContinue = True Then

              ' ** Recheck check boxes.
34850         GetCounts 2  ' ** Procedure: Above.

34860         lngCnt = 0&
34870         For lngX = 0& To (lngRecs - 1&)
34880           strTargetFile = vbNullString
34890           If arr_varRec(R_EXP, lngX) = True Then

34900             For lngY = 0& To (lngCtls - 1&)
34910               If arr_varCtl(C_TIER, lngY) = arr_varRec(R_TIER, lngX) And arr_varCtl(C_ITEM, lngY) = arr_varRec(R_ITEM, lngX) Then
34920                 strTmp01 = arr_varCtl(C_CNAM, lngY) & "_lbl"
34930                 strTargetFile = .Controls(strTmp01).Caption
34940                 Exit For
34950               End If
34960             Next

34970             If strTargetFile <> vbNullString Then
34980               strTmp01 = Rem_Under(Right(arr_varRec(R_QNAM, lngX), 5))  ' ** Module Function: modStringFuncs.
34990               strTargetFile = "TA_" & CStr(year(Date)) & Format(month(Date), "00") & Format(day(Date), "00") & "_" & _
                      strTmp01 & "_" & strTargetFile
35000               .AllData_File = strTargetFile
35010               DoEvents
35020               strTargetFile = strTargetFile & ".csv"
35030               strTargetPathFile = strTargetPath & LNK_SEP & strTargetFile
                    ' ** Access creates a new text file when you export data from Access.If the file name is
                    ' ** the same as the name of an existing text file, Access replaces the existing text file.
35040               DoCmd.TransferText acExportDelim, arr_varRec(R_SNAM, lngX), arr_varRec(R_QNAM, lngX), strTargetPathFile, True
35050               lngCnt = lngCnt + 1&
35060               DoEvents
35070               ForcePause 1  ' ** Module Procedure: modCodeUtilities.
35080             End If
35090           End If

35100         Next

35110         Beep
35120         DoCmd.Hourglass False
35130         If lngCnt > 0& Then
35140           MsgBox CStr(lngCnt) & " table" & IIf(lngCnt = 1, " has ", "s have ") & "been exported to the folder:" & vbCrLf & vbCrLf & _
                  strTargetPath, vbInformation + vbOKOnly, "Export Successful"
35150         Else
35160           MsgBox "No tables were exported.", vbInformation + vbOKOnly, "Export Failed"
35170         End If

35180         .cmdCancel.SetFocus

35190       End If  ' ** blnContinue.

            ' ** These tables have end-user data.
            ' ** Tier 1:
            'account
            '  TACloud_Account_Export_Specification
            '  qryAccountExport_02_01_01
            'masterasset
            '  TACloud_MasterAsset_Export_Specification
            '  qryAccountExport_02_01_02
            'ActiveAssets
            '  TACloud_ActiveAssets_Export_Specification
            '  qryAccountExport_02_01_03
            'LedgerArchive
            '  TACloud_LedgerArchive_Export_Specification
            '  qryAccountExport_02_01_04
            'ledger
            '  TACloud_Ledger_Export_Specification
            '  qryAccountExport_02_01_05
            'Balance
            '  TACloud_Balance_Export_Specification
            '  qryAccountExport_02_01_06

            ' ** Tier 2:
            'm_REVCODE
            '  TACloud_m_REVCODE_Export_Specification
            '  qryAccountExport_02_02_01
            'RecurringItems
            '  TACloud_RecurringItems_Export_Specification
            '  qryAccountExport_02_02_02
            'tblPricing_MasterAsset_History
            '  TACloud_tblPricing_MasterAsset_History_Export_Specification
            '  qryAccountExport_02_02_03
            'tblCurrency_History
            '  TACloud_tblCurrency_History_Export_Specification
            '  qryAccountExport_02_02_04
            'tblLedgerHidden
            '  TACloud_tblLedgerHidden_Export_Specification
            '  qryAccountExport_02_02_05
            'Location
            '  TACloud_Location_Export_Specification
            '  qryAccountExport_02_02_06

            ' ** Tier 3:
            'adminofficer
            '  TACloud_AdminOfficer_Export_Specification
            '  qryAccountExport_02_03_01
            'PortfolioModel
            '  TACloud_PortfolioModel_Export_Specification
            '  qryAccountExport_02_03_02
            'Schedule
            '  TACloud_Schedule_Export_Specification
            '  qryAccountExport_02_03_03
            'ScheduleDetail
            '  TACloud_ScheduleDetail_Export_Specification
            '  qryAccountExport_02_03_04
            'tblCheckPOSPay
            '  TACloud_tblCheckPOSPay_Export_Specification
            '  qryAccountExport_02_03_05
            'tblCheckPOSPay_Detail
            '  TACloud_tblCheckPOSPay_Detail_Export_Specification
            '  qryAccountExport_02_03_06ckgTier3_chk06_CheckPOSPay_Detail
            'tblCheckReconcile_Account
            '  TACloud_tblCheckReconcile_Account_Export_Specification
            '  qryAccountExport_02_03_07
            'tblCheckReconcile_Item
            '  TACloud_tblCheckReconcile_Item_Export_Specification
            '  qryAccountExport_02_03_08
            'tblCheckMemo
            '  TACloud_tblCheckMemo_Export_Specification
            '  qryAccountExport_02_03_09

35200     End If  ' ** blnNoData.
35210   End With

EXITP:
35220   Exit Sub

ERRH:
1220    DoCmd.Hourglass False
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub ShowAllData(blnShow As Boolean)

35300 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAllData"

        Dim ctl As Access.Control
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long
        Dim lngX As Long, lngE As Long

35310   With Me

35320     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
35330       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
35340     End If

35350     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
35360     lngMonitorNum = 1&: lngTmp04 = 0&
35370     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
35380     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

35390     If .opgExport.Height <> lngExport_Height Then
35400       blnShow = True
35410     End If

35420     Select Case blnShow
          Case True

35430       DoCmd.SelectObject acForm, THIS_NAME, False
35440       If lngMonitorNum = 1& Then lngTmp04 = lngFrm_Top
35450       DoCmd.MoveSize lngFrm_Left, lngTmp04, lngFrm_Width, lngFrm_Height  'lngFrm_Top
35460       If lngMonitorNum > 1& Then
35470         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
35480       End If

35490       .opgExport.Height = lngExport_Height
35500       .opgExport_box.Height = lngExport_Height
35510       For lngX = 0& To (lngItems - 1&)
35520         .Controls(arr_varItem(I_CNAM, lngX)).Top = arr_varItem(I_TOP, lngX)
35530         .Controls(arr_varItem(I_CNAM, lngX)).Visible = arr_varItem(I_VIS, lngX)
35540       Next
35550       .Detail_hline01.Top = lngDetailHLine01_Top
35560       .Detail_hline02.Top = (.Detail_hline01.Top + lngTpp)
35570       .Detail_vline01.Top = .Detail_hline01.Top
35580       .Detail_vline02.Top = .Detail_hline01.Top
35590       .cmdExport.Top = lngExport_Top
35600       .opgExport_optAll.Enabled = True

35610     Case False

35620       lngItems = 0&
35630       ReDim arr_varItem(I_ELEMS, 0)
35640       If .opgExport = .opgExport_optAll.OptionValue Then
35650         .opgExport = .opgExport_optAcctProf.OptionValue
35660         opgExport_AfterUpdate  ' ** Procedure: Above.
35670         DoEvents
35680       End If
35690       .opgExport_optAll.Enabled = False

35700       For Each ctl In .Detail.Controls
35710         With ctl
35720           If Left(.Name, 7) = "ckgTier" Then
35730             lngItems = lngItems + 1&
35740             lngE = (lngItems - 1&)
35750             ReDim Preserve arr_varItem(I_ELEMS, lngE)
35760             arr_varItem(I_CNAM, lngE) = .Name
35770             arr_varItem(I_TOP, lngE) = .Top
35780             arr_varItem(I_VIS, lngE) = .Visible
35790             .Visible = False
35800           ElseIf Left(.Name, 7) = "cmdTier" Then
35810             lngItems = lngItems + 1&
35820             lngE = (lngItems - 1&)
35830             ReDim Preserve arr_varItem(I_ELEMS, lngE)
35840             arr_varItem(I_CNAM, lngE) = .Name
35850             arr_varItem(I_TOP, lngE) = .Top
35860             arr_varItem(I_VIS, lngE) = .Visible
35870             .Visible = False
35880           ElseIf Left(.Name, 8) = "AllData_" Then
35890             If .Name <> "AllData_alt_line1" And .Name <> "AllData_alt_line2" Then
35900               lngItems = lngItems + 1&
35910               lngE = (lngItems - 1&)
35920               ReDim Preserve arr_varItem(I_ELEMS, lngE)
35930               arr_varItem(I_CNAM, lngE) = .Name
35940               arr_varItem(I_TOP, lngE) = .Top
35950               arr_varItem(I_VIS, lngE) = .Visible
35960               .Visible = False
35970             End If
35980           ElseIf Left(.Name, 10) = "cmdBrowse2" Then
35990             lngItems = lngItems + 1&
36000             lngE = (lngItems - 1&)
36010             ReDim Preserve arr_varItem(I_ELEMS, lngE)
36020             arr_varItem(I_CNAM, lngE) = .Name
36030             arr_varItem(I_TOP, lngE) = .Top
36040             arr_varItem(I_VIS, lngE) = .Visible
36050             .Visible = False
36060           ElseIf Left(.Name, 7) = "DefMsg_" Then
36070             lngItems = lngItems + 1&
36080             lngE = (lngItems - 1&)
36090             ReDim Preserve arr_varItem(I_ELEMS, lngE)
36100             arr_varItem(I_CNAM, lngE) = .Name
36110             arr_varItem(I_TOP, lngE) = .Top
36120             arr_varItem(I_VIS, lngE) = .Visible
36130             .Visible = False
36140           End If
36150         End With
36160       Next

36170       lngTmp01 = ((.opgExport.Top + .opgExport.Height) - .AllData_alt_line1.Top)
36180       lngTmp02 = ((.AllData_box3.Top + .AllData_box3.Height) - (.opgExport.Top + .opgExport.Height))
36190       .opgExport.Height = (.opgExport.Height - lngTmp01)
36200       .opgExport_box.Height = .opgExport.Height

36210       For lngX = 0& To (lngItems - 1&)
36220         .Controls(arr_varItem(I_CNAM, lngX)).Top = 0&
36230       Next

36240       lngTmp03 = (.cmdExport.Top - .Detail_hline01.Top)
36250       .Detail_hline01.Top = ((.opgExport_box.Top + .opgExport_box.Height) + (6& * lngTpp))
36260       .Detail_hline02.Top = .Detail_hline01.Top + lngTpp
36270       .Detail_vline01.Top = .Detail_hline01.Top
36280       .Detail_vline02.Top = .Detail_hline01.Top
36290       .cmdExport.Top = (.Detail_hline01.Top + lngTmp03)

36300       lngTmp03 = ((lngFrm_Height - (lngFrm_Height - (lngTmp01 + lngTmp02))) / 2)

36310       DoCmd.SelectObject acForm, THIS_NAME, False
36320       If lngMonitorNum = 1& Then lngTmp04 = lngFrm_Top
36330       DoCmd.MoveSize lngFrm_Left, (lngTmp04 + lngTmp03), lngFrm_Width, (lngFrm_Height - (lngTmp01 + lngTmp02))
36340       If lngMonitorNum > 1& Then
36350         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
36360       End If

36370     End Select

36380   End With

EXITP:
36390   Set ctl = Nothing
36400   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub ShowDataSwitch()

36500 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowDataSwitch"

36510   With Me
36520     Select Case blnNoData
          Case True
            ' ** Was off, turn on.
36530       blnNoData = False
36540       ShowAllData True  ' ** Procedure: Above.
36550       .opgExport_Open = True
36560     Case False
            ' ** Was on, turn off.
36570       blnNoData = True
36580       ShowAllData False  ' ** Procedure: Above.
36590       .opgExport_Open = False
36600     End Select
36610   End With

EXITP:
36620   Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub
