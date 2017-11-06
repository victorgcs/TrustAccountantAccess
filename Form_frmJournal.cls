VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmJournal"

'VGC 10/01/2017: CHANGES!
' ** TimerInterval: 100 (100 set by GoToReport)  NO! NEEDED FIRST TIME!!

' #######################################
' ## Monitor Funcs:
' ##   Form_Load()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Map:              A {cmdDividendMap, cmdInterestMap, cmdPurchaseMap, cmdMiscMap}
' **   Close:            C {cmdDividendClose, cmdInterestClose, cmdPurchaseClose, cmdSaleClose, cmdMiscClose}
' **   Dividend:         D {opgJournal_optDividend}
' **   Cancel:           E {cmdDividendCancel, cmdInterestCancel, cmdPurchaseCancel, cmdSaleCancel, cmdMiscCancel}
' **   STCGL Map:        G {cmdMiscMap_STCGL on frmJournal_Sub5_Misc}
' **   Print Check:      H {chkChecks on frmJournal_Sub5_Misc}
' **   Interest:         I {opgJournal_optInterest}
' **   New Loc:          L {cmdPurchaseLocNew on frmJournal_Sub3_Purchase}
' **   Lot Info:         L {cmdSaleLotInfo on frmJournal_Sub4_Sold}
' **   LTCL Map:         L {cmdMiscMap_LTCL on frmJournal_Sub5_Misc}
' **   Miscellaneous:    M {opgJournal_optMisc}
' **   Inc/Exp Codes:    N {cmbRevenueCodes on each subform}
' **   OK:               O {cmdDividendOK, cmdInterestOK, cmdPurchaseOK, cmdSaleOK, cmdMiscOK}
' **   Purchase:         P {opgJournal_optPurchase}
' **   Sale:             S {opgJournal_optSale}
' **   Tax Codes:        T {cmbTaxCodes on each subform}
' **   New Recur:        U {cmdMiscRecurNew on frmJournal_Sub5_Misc}
' **   New Asset:        W {cmdPurchaseAssetNew on frmJournal_Sub3_Purchase}
' **   Exit:             X {cmdDividendClose, cmdInterestClose, cmdPurchaseClose, cmdSaleClose, cmdMiscClose}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1 on frmJournal_Sub1_Dividend, frmJournal_Sub2_Interest, frmJournal_Sub3_Purchase, frmJournal_Sub4_Sold, frmJournal_Sub5_Misc}
' **   Date Picker:      F8 {cmdCalendar2 on frmJournal_Sub3_Purchase, frmJournal_Sub4_Sold}

' ** Shortcut Ctrl keys responsive from this form:
' **   Next Record:      N {MoveRec on each subform}
' **   Previous Record:  P {MoveRec on each subform}
' **   Switch:           T {cmdSwitch}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Reinvest:         R {tglDividendReinvest, tglInterestReinvest}
' **   Mark Sale:        S {tglMiscSale on frmJournal_Sub5_Misc}
' **   Close Timer:      T {Form_KeyDown}

' ** Additional matching forms:
' **   frmReinvest_Dividend
' **     Gets it's transdate from frmJournal_Sub1_Dividend.
' **   frmReinvest_Interest
' **     Gets it's transdate from frmJournal_Sub2_Interest.

' ** NOTE: THES FORMS ARE LAID OUT TO LOOK LIKE THE JOURNAL FORMS:
' **   frmReinvest_Dividend
' **   frmReinvest_Interest
' ** AND THESE CAN ALSO PUT IN JOURNAL DATA:
' **   frmMap_Div
' **   frmMap_Int
' **   frmMap_Rec
' **   frmMap_Split

' ** Proposed journalSubType List:
' **   Reinvest
' **   LTCG          (VGC 02/15/09)
' **   StockSplit    (VGC 02/15/09)
' **   MutiLot       ?  Sale of multiple tax lots, where shown as 1 transaction.
' **   Liability (+) (VGC 08/19/09)
' **   Liability (-) (VGC 08/19/09)

' ** Array: arr_varOpt().
Private lngOpts As Long, arr_varOpt() As Variant
Private Const OPT_ELEMS As Integer = 1  ' ** Array's first-element UBound().
Private Const OPT_NAM As Integer = 0
Private Const OPT_VAL As Integer = 1

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long

Private lngNoChars As Long, arr_varNoChar As Variant
Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private blnShowingDividend As Boolean
Private lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long ', lngRecsCur As Long
Private THAT_PROC As String, That_Erl As Long
Private strCallingForm As String, strSubform As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim grp As DAO.Group, usr As DAO.User
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)
150           blnIsOpen = True: blnShowingDividend = False
160           blnGTR_Emblem = False: lngGTR_Stat = 0&: blnGoingToReport2 = False

              ' ** See modQueryFunctions1.FormRef() for how this is used.
170           gstrFormQuerySpec = THIS_NAME

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190           gblnMessage = False  ' ** frmJournal_Sub1_Dividend uses this.
200           gblnSwitchTo = False

              ' ** Variables are fed empty, then populated ByRef.
210           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

220           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

230           Journal_OptLoad  ' ** Procedure: Below.
240           JRN_DIV = .opgJournal_optDividend.OptionValue
250           JRN_INT = .opgJournal_optInterest.OptionValue
260           JRN_PUR = .opgJournal_optPurchase.OptionValue
270           JRN_SAL = .opgJournal_optSale.OptionValue
280           JRN_MIS = .opgJournal_optMisc.OptionValue

              ' ** Loop through all the security groups, checking if our current user is in the admins group.
290           gblnAdmin = False
300           For Each grp In DBEngine.Workspaces(0).Groups
310             If grp.Name = "Admins" Then
320               For Each usr In grp.Users
330                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
340                   gblnAdmin = True
350                   Exit For
360                 End If
370               Next
380             End If
390           Next

              ' ** Turn the MouseWheel Off.
              ' ** Call our MouseHook function in the MouseHook dll.
              ' ** Please not the Optional GlobalHook BOOLEAN parameter
              ' ** Several developers asked for the MouseHook to be able to work with
              ' ** multiple instances of Access. In order to accomodate this request I
              ' ** have modified the function to allow the caller to
              ' ** specify a thread specific(this current instance of Access only) or
              ' ** a global(all applications) MouseWheel Hook.
              ' ** Only use the GlobalHook if you will be running multiple instances of Access!
400           blnRetVal = MouseWheelOFF(False)  ' ** Module Function: modMouseWheel.
              ' ** Turned back on 12/25/2008. {Remarked out by VGC for 2.1.46 and beyond.}

              'NOT SURE HOW USEFUL THIS IS ANYMORE!
              'SetNav_Access2007 Me  ' ** Module Function: modXAccess_07_10_Funcs.

410           Select Case glngJournalForm
              Case JRN_DIV
420             ShowDividend  ' ** Procedure: Below.
430           Case JRN_INT
440             ShowInterest  ' ** Procedure: Below.
450           Case JRN_PUR
460             ShowPurchase  ' ** Procedure: Below.
470           Case JRN_SAL
480             ShowSale  ' ** Procedure: Below.
490           Case JRN_MIS
500             ShowMisc  ' ** Procedure: Below.
510           Case Else
520             ShowDividend  ' ** Procedure: Below.
530           End Select

540           Select Case glngJournalForm
              Case JRN_DIV
550             .frmJournal_Sub1_Dividend.SetFocus
560           Case JRN_INT
570             .frmJournal_Sub2_Interest.SetFocus
580           Case JRN_PUR
590             .frmJournal_Sub3_Purchase.SetFocus
600           Case JRN_SAL
610             .frmJournal_Sub4_Sold.SetFocus
620           Case JRN_MIS
630             .frmJournal_Sub5_Misc.SetFocus
640           End Select

              ' ** Advance to new record on subform.
650           DoCmd.GoToRecord , , acNewRec

660         Else
670           Cancel = -1
680         End If
690       Else
700         Cancel = -1
710       End If
720     End With

730     If Cancel = -1 Then
740       Beep
750       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
760     End If

EXITP:
770     Set usr = Nothing
780     Set grp = Nothing
790     Exit Sub

ERRH:
800     DoCmd.Hourglass False
810     Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Ignore.
820     Case Else
830       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
840     End Select
850     Resume EXITP

End Sub

Private Sub Form_Load()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngTmp01 As Long

910     With Me

920       Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

930       IncludeCurrency  ' ** Procedure: Below.

940       If gblnGoToReport = True Then

950         lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
960         lngMonitorNum = 1&: lngTmp01 = 0&
970         EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
980         If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

990         If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
1000        DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, (lngFrm_Height + .FormFooter.Height)  'lngFrm_Top
1010        If lngMonitorNum > 1& Then
1020          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
1030        End If

1040        .FormFooter.Visible = True
1050        DoEvents

1060      End If

1070      DoCmd.SelectObject acForm, THIS_NAME, False

1080      DoCmd.Hourglass False

1090    End With

EXITP:
1100    Set rst = Nothing
1110    Set qdf = Nothing
1120    Set dbs = Nothing
1130    Exit Sub

ERRH:
1140    DoCmd.Hourglass False
1150    Select Case ERR.Number
        Case Else
1160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1170    End Select
1180    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1220    intShiftDown = (Shift And acShiftMask) > 0
1230    intAltDown = (Shift And acAltMask) > 0
1240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Next Record:      N {MoveRec on each subform}
        ' **   Previous Record:  P {MoveRec on each subform}
        ' **   Switch:           T {cmdSwitch}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Reinvest:         R {tglDividendReinvest, tglInterestReinvest}
        ' **   Mark Sale:        S {tglMiscSale on frmJournal_Sub5_Misc}
        ' **   Close Timer:      T {Form_KeyDown}

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyEscape
1270        intRetVal = 0
1280        cmdClose_Click  ' ** Procedure: Below.
1290      Case vbKeyF7
1300        With Me
1310          intRetVal = 0
1320          Select Case .opgJournal
              Case .opgJournal_optDividend.OptionValue
1330            If .frmJournal_Sub1_Dividend.Visible = True Then
1340              .frmJournal_Sub1_Dividend.SetFocus
1350              .frmJournal_Sub1_Dividend.Form.cmdCalendar1.SetFocus
1360              .frmJournal_Sub1_Dividend.Form.cmdCalendar1_Click  ' ** Form Procedure: frmJournal_Sub1_Dividend.
1370            End If
1380          Case .opgJournal_optInterest.OptionValue
1390            If .frmJournal_Sub2_Interest.Visible = True Then
1400              .frmJournal_Sub2_Interest.SetFocus
1410              .frmJournal_Sub2_Interest.Form.cmdCalendar1.SetFocus
1420              .frmJournal_Sub2_Interest.Form.cmdCalendar1_Click  ' ** Form Procedure: frmJournal_Sub2_Interest.
1430            End If
1440          Case .opgJournal_optPurchase.OptionValue
1450            If .frmJournal_Sub3_Purchase.Visible = True Then
1460              .frmJournal_Sub3_Purchase.SetFocus
1470              .frmJournal_Sub3_Purchase.Form.cmdCalendar1.SetFocus
1480              .frmJournal_Sub3_Purchase.Form.cmdCalendar1_Click  ' ** Form Procedure: frmJournal_Sub3_Purchase.
1490            End If
1500          Case .opgJournal_optSale.OptionValue
1510            If .frmJournal_Sub4_Sold.Visible = True Then
1520              .frmJournal_Sub4_Sold.SetFocus
1530              .frmJournal_Sub4_Sold.Form.cmdCalendar1.SetFocus
1540              .frmJournal_Sub4_Sold.Form.cmdCalendar1_Click  ' ** Form Procedure: frmJournal_Sub4_Sold.
1550            End If
1560          Case .opgJournal_optMisc.OptionValue
1570            If .opgJournal_optMisc.Visible = True Then
1580              .frmJournal_Sub5_Misc.SetFocus
1590              .frmJournal_Sub5_Misc.Form.cmdCalendar1.SetFocus
1600              .frmJournal_Sub5_Misc.Form.cmdCalendar1_Click  ' ** Form Procedure: frmJournal_Sub5_Misc.
1610            End If
1620          End Select
1630        End With
1640      Case vbKeyF8
1650        With Me
1660          Select Case .opgJournal
              Case .opgJournal_optDividend.OptionValue
1670            Beep
1680          Case .opgJournal_optInterest.OptionValue
1690            Beep
1700          Case .opgJournal_optPurchase.OptionValue
1710            If .frmJournal_Sub3_Purchase.Visible = True Then
1720              .frmJournal_Sub3_Purchase.SetFocus
1730              .frmJournal_Sub3_Purchase.Form.cmdCalendar2.SetFocus
1740              .frmJournal_Sub3_Purchase.Form.cmdCalendar2_Click  ' ** Form Procedure: frmJournal_Sub3_Purchase.
1750            End If
1760          Case .opgJournal_optSale.OptionValue
1770            If .frmJournal_Sub4_Sold.Visible = True Then
1780              .frmJournal_Sub4_Sold.SetFocus
1790              .frmJournal_Sub4_Sold.Form.cmdCalendar2.SetFocus
1800              .frmJournal_Sub4_Sold.Form.cmdCalendar2_Click  ' ** Form Procedure: frmJournal_Sub4_Sold.
1810            End If
1820          Case .opgJournal_optMisc.OptionValue
1830            Beep
1840          End Select
1850        End With
1860      End Select
1870    End If

        ' ** Ctrl keys.
1880    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1890      Select Case intRetVal
          Case vbKeyN
1900        With Me
1910          intRetVal = 0
1920          Select Case .opgJournal
              Case .opgJournal_optDividend.OptionValue
1930            .frmJournal_Sub1_Dividend.SetFocus
1940            .frmJournal_Sub1_Dividend.Form.MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Sub1_Dividend.
1950          Case .opgJournal_optInterest.OptionValue
1960            .frmJournal_Sub2_Interest.SetFocus
1970            .frmJournal_Sub2_Interest.Form.MoveRec acCmdRecordsGoToNext   ' ** Form Procedure: frmJournal_Sub2_Interest.
1980          Case .opgJournal_optPurchase.OptionValue
1990            .frmJournal_Sub3_Purchase.SetFocus
2000            .frmJournal_Sub3_Purchase.Form.MoveRec acCmdRecordsGoToNext   ' ** Form Procedure: frmJournal_Sub3_Purchase.
2010          Case .opgJournal_optSale.OptionValue
2020            .frmJournal_Sub4_Sold.SetFocus
2030            .frmJournal_Sub4_Sold.Form.MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Sub4_Sold.
2040          Case .opgJournal_optMisc.OptionValue
2050            .frmJournal_Sub5_Misc.SetFocus
2060            .frmJournal_Sub5_Misc.Form.MoveRec acCmdRecordsGoToNext   ' ** Form Procedure: frmJournal_Sub5_Misc.
2070          End Select
2080        End With
2090      Case vbKeyP
2100        With Me
2110          intRetVal = 0
2120          Select Case .opgJournal
              Case .opgJournal_optDividend.OptionValue
2130            .frmJournal_Sub1_Dividend.SetFocus
2140            .frmJournal_Sub1_Dividend.Form.MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmJournal_Sub1_Dividend.
2150          Case .opgJournal_optInterest.OptionValue
2160            .frmJournal_Sub2_Interest.SetFocus
2170            .frmJournal_Sub2_Interest.Form.MoveRec acCmdRecordsGoToPrevious   ' ** Form Procedure: frmJournal_Sub2_Interest.
2180          Case .opgJournal_optPurchase.OptionValue
2190            .frmJournal_Sub3_Purchase.SetFocus
2200            .frmJournal_Sub3_Purchase.Form.MoveRec acCmdRecordsGoToPrevious   ' ** Form Procedure: frmJournal_Sub3_Purchase.
2210          Case .opgJournal_optSale.OptionValue
2220            .frmJournal_Sub4_Sold.SetFocus
2230            .frmJournal_Sub4_Sold.Form.MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmJournal_Sub4_Sold.
2240          Case .opgJournal_optMisc.OptionValue
2250            .frmJournal_Sub5_Misc.SetFocus
2260            .frmJournal_Sub5_Misc.Form.MoveRec acCmdRecordsGoToPrevious   ' ** Form Procedure: frmJournal_Sub5_Misc.
2270          End Select
2280        End With
2290      Case vbKeyT
2300        With Me
2310          intRetVal = 0
2320          .cmdSwitch.SetFocus
2330          DoEvents
2340          cmdSwitch_Click  ' ** Procedure: Below.
2350        End With
2360      End Select
2370    End If

        ' ** Ctrl-Shift keys.
2380    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2390      Select Case intRetVal
          Case vbKeyF
2400        With Me
2410          intRetVal = 0
2420          .FocusHolder.SetFocus
2430        End With
2440      Case vbKeyR
2450        With Me
2460          intRetVal = 0
2470          If .frmJournal_Sub1_Dividend.Visible = True Then
2480            .frmJournal_Sub1_Dividend.SetFocus
2490            .frmJournal_Sub1_Dividend.Form.tglDividendReinvest.SetFocus
2500            .frmJournal_Sub1_Dividend.Form.tglDividendReinvest_Click  ' ** Form Procedure: frmJournal_Sub1_Dividend.
2510          ElseIf .frmJournal_Sub2_Interest.Visible = True Then
2520            .frmJournal_Sub2_Interest.SetFocus
2530            .frmJournal_Sub2_Interest.Form.tglInterestReinvest.SetFocus
2540            .frmJournal_Sub2_Interest.Form.tglInterestReinvest_Click  ' ** Form Procedure: frmJournal_Sub2_Interest.
2550          Else
2560            Beep
2570          End If
2580        End With
2590      Case vbKeyT
2600        intRetVal = 0
2610        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
2620      End Select
2630    End If

EXITP:
2640    KeyCode = intRetVal
2650    Exit Sub

ERRH:
2660    intRetVal = 0
2670    THAT_PROC = THIS_PROC
2680    That_Erl = Erl
2690    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2700    Resume EXITP

End Sub

Private Sub Form_Timer()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String
        Dim lngTmp03 As Long

        Static lngTmp01 As Long, lngTmp02 As Long
        Static frm As Access.Form

2810    With Me
2820      .TimerInterval = 0&
2830      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
2840        lngGTR_Stat = lngGTR_Stat + 1&
2850        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
2860        lngTmp01 = lngTmp01 + 1&
2870        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
2880        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
2890        If lngTmp01 > 1& Then
2900          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
2910        End If
2920        If lngTmp01 = 24& Then
2930          lngTmp01 = 0&
2940          lngTmp02 = lngTmp02 + 1&
2950          If lngTmp02 < 5& Or (blnGoingToReport2 = True And lngTmp02 < 20&) Then
2960            .TimerInterval = 50&
2970          Else
2980            blnGTR_Emblem = False
2990            Set frm = Nothing
3000          End If
3010        Else
3020          .TimerInterval = 50&
3030        End If
3040      End If  ' ** blnGTR_Emblem.
3050      If gblnSetFocus = True Then
3060        gblnSetFocus = False
3070        lngTmp03 = 0&
3080        strFormName = vbNullString
3090  On Error Resume Next
3100        Do Until strFormName = THIS_NAME
3110          lngTmp03 = lngTmp03 + 1&
3120          DoCmd.SelectObject acForm, THIS_NAME, False
3130          strFormName = Screen.ActiveForm.Name
3140          If lngTmp03 > 100& Then Exit Do
3150        Loop
3160  On Error GoTo ERRH
3170        If blnIsOpen = True Then
3180          .opgJournal = .opgJournal_optInterest.OptionValue
3190          opgJournal_AfterUpdate  ' ** Procedure: Below.
3200          DoEvents
3210          .Repaint
3220          .opgJournal = .opgJournal_optDividend.OptionValue
3230          opgJournal_AfterUpdate  ' ** Procedure: Below.
3240          DoEvents
3250          .Repaint
3260          blnIsOpen = False
3270          DoCmd.Hourglass False
3280          DoEvents
3290          gblnSetFocus = True
3300          .TimerInterval = 100&
3310        End If
3320      End If
3330      If gblnGoToReport = True Then
3340        DoCmd.Hourglass True  ' ** Make sure it's still running.
3350        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
3360        Else
3370          DoCmd.Hourglass True  ' ** Make sure it's still running.
3380          If blnGTR_Emblem = False Then
3390            blnGTR_Emblem = True
3400            lngTmp01 = 0&: lngTmp02 = 0&
3410          End If
3420          DoEvents
3430          Select Case blnGoingToReport2
              Case True
                ' ** Just saleCost.
3440            If blnGTR_Emblem = False Then
3450              DoCmd.Hourglass False
3460              blnGoingToReport = False
3470              blnGoingToReport2 = False
3480              gblnGoToReport = False
3490              GTREmblem_Off  ' ** Procedure: Below.
3500              Beep
3510            End If
3520          Case False
3530            Select Case blnGoingToReport
                Case True
3540              Select Case garr_varGoToReport(GTR_CTL2)
                  Case "cmdDividendMap"
3550                If garr_varGoToReport(GTR_INV2) = True Then
3560                  .TimerInterval = 50&
3570                  .Controls(strSubform).SetFocus
3580                  frm.cmdDividendMap_Click  ' ** Form Procedure: frmJournal_Sub1_Dividend.
3590                End If
3600              Case "cmdInterestMap"
3610                If garr_varGoToReport(GTR_INV2) = True Then
3620                  .TimerInterval = 50&
3630                  .Controls(strSubform).SetFocus
3640                  frm.cmdInterestMap_Click  ' ** Form Procedure: frmJournal_Sub2_Interest.
3650                End If
3660              Case "cmdPurchaseMap"
3670                If garr_varGoToReport(GTR_INV2) = True Then
3680                  .TimerInterval = 50&
3690                  .Controls(strSubform).SetFocus
3700                  frm.cmdPurchaseMap_Click  ' ** Form Procedure: frmJournal_Sub3_Purchase.
3710                End If
3720              Case "cmdMiscMap"
3730                If garr_varGoToReport(GTR_INV2) = True Then
3740                  .TimerInterval = 50&
3750                  .Controls(strSubform).SetFocus
3760                  frm.cmdMiscMap_Click  ' ** Form Procedure: frmJournal_Sub5_Misc.
3770                End If
3780              Case "cmdMiscMap_LTCL"
3790                If garr_varGoToReport(GTR_INV2) = True Then
3800                  .TimerInterval = 50&
3810                  .Controls(strSubform).SetFocus
3820                  frm.cmdMiscMap_LTCL_Click  ' ** Form Procedure: frmJournal_Sub5_Misc.
3830                End If
3840              Case "cmdMiscMap_STCGL"
3850                If garr_varGoToReport(GTR_INV2) = True Then
3860                  .TimerInterval = 50&
3870                  .Controls(strSubform).SetFocus
3880                  frm.cmdMiscMap_STCGL_Click  ' ** Form Procedure: frmJournal_Sub5_Misc.
3890                End If
3900              Case "saleCost"
                    ' ** List of Holdings by Tax Lot.
3910                .Controls(strSubform).SetFocus
3920                frm.saleAccountNo.SetFocus
3930                If garr_varGoToReport(GTR_INV2) = True Or garr_varGoToReport(GTR_FOC2) = True Then
                      ' ** Here we turn it over to frmJournal_Sub4_Sold.
                      ' ** See GTREmblem_Set() and GTREmblem_Get(), below.
3940                  blnGoingToReport2 = True
3950                  frm.TimerInterval = 50&
3960                  .TimerInterval = 50&
3970                End If
3980              End Select
3990              DoEvents
4000              blnGoingToReport = False
4010            Case False
4020              blnGoingToReport = True
4030              DoCmd.Hourglass True  ' ** Make sure it's still running.
4040              DoEvents
4050              Select Case garr_varGoToReport(GTR_CTL2)
                  Case "cmdDividendMap"
4060                .opgJournal = .opgJournal_optDividend.OptionValue
4070                opgJournal_AfterUpdate  ' ** Procedure: Below.
4080                DoEvents
4090                strSubform = "frmJournal_Sub1_Dividend"
4100                Set frm = .Controls(strSubform).Form
4110                With frm
4120                  .GoToReport_arw_mapdiv_img.Visible = True
4130                  .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
4140                End With
4150                DoEvents
4160              Case "cmdInterestMap"
4170                .opgJournal = .opgJournal_optInterest.OptionValue
4180                opgJournal_AfterUpdate  ' ** Procedure: Below.
4190                DoEvents
4200                strSubform = "frmJournal_Sub2_Interest"
4210                Set frm = .Controls(strSubform).Form
4220                With frm
4230                  .GoToReport_arw_mapint_img.Visible = True
4240                  .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
4250                End With
4260                DoEvents
4270              Case "cmdPurchaseMap"
4280                .opgJournal = .opgJournal_optPurchase.OptionValue
4290                opgJournal_AfterUpdate  ' ** Procedure: Below.
4300                DoEvents
4310                strSubform = "frmJournal_Sub3_Purchase"
4320                Set frm = .Controls(strSubform).Form
4330                With frm
4340                  .GoToReport_arw_mapsplit_img.Visible = True
4350                  .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
4360                End With
4370                DoEvents
4380              Case "cmdMiscMap"
4390                .opgJournal = .opgJournal_optMisc.OptionValue
4400                opgJournal_AfterUpdate  ' ** Procedure: Below.
4410                DoEvents
4420                strSubform = "frmJournal_Sub5_Misc"
4430                Set frm = .Controls(strSubform).Form
4440                With frm
4450                  If gblnSpecialCapGainLoss = True And ((gblnAdmin = True) Or (gblnAdmin = False And gintSpecialCapGainLossOpt <> 2)) Then
4460                    .cmdMiscMap_LTCL.Enabled = False
4470                    .cmdMiscMap_LTCL.Visible = False
4480                    .cmdMiscMap_LTCL_raised_img.Visible = False
4490                    .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
4500                    .cmdMiscMap_LTCL_raised_focus_img.Visible = False
4510                    .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
4520                    .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
4530                    .cmdMiscMap_LTCL_raised_img_dis.Visible = False
4540                    DoEvents
4550                  End If
4560                  .GoToReport_arw_maprec_img.Visible = True
4570                  .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
4580                End With
4590                DoEvents
4600              Case "cmdMiscMap_LTCL"
4610                .opgJournal = .opgJournal_optMisc.OptionValue
4620                opgJournal_AfterUpdate  ' ** Procedure: Below.
4630                DoEvents
4640                strSubform = "frmJournal_Sub5_Misc"
4650                Set frm = .Controls(strSubform).Form
4660                With frm
4670                  .cmdMiscMap.Visible = False
4680                  .GoToReport_arw_mapltcl_img.Visible = True
4690                  If .Controls(garr_varGoToReport(GTR_CTL2)).Visible = True And .Controls(garr_varGoToReport(GTR_CTL2)).Enabled = True Then
4700                    .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
4710                  Else
4720                    Beep
                        'MsgBox "CAN'T GET THERE FROM HERE!"
4730                  End If
4740                End With
4750                DoEvents
4760              Case "cmdMiscMap_STCGL"
4770                .opgJournal = .opgJournal_optMisc.OptionValue
4780                opgJournal_AfterUpdate  ' ** Procedure: Below.
4790                DoEvents
4800                strSubform = "frmJournal_Sub5_Misc"
4810                Set frm = .Controls(strSubform).Form
4820                With frm
4830                  .cmdMiscMap_LTCL.Visible = False
4840                  .cmdMiscMap_LTCL_raised_img.Visible = False
4850                  .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
4860                  .cmdMiscMap_LTCL_raised_focus_img.Visible = False
4870                  .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
4880                  .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
4890                  .cmdMiscMap_LTCL_raised_img_dis.Visible = False
4900                  .GoToReport_arw_mapstcgl_img.Visible = True
4910                  If .Controls(garr_varGoToReport(GTR_CTL2)).Visible = True And .Controls(garr_varGoToReport(GTR_CTL2)).Enabled = True Then
4920                    .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
4930                  Else
4940                    Beep
                        'MsgBox "CAN'T GET THERE FROM HERE!"
4950                  End If
4960                End With
4970                DoEvents
4980              Case "saleCost"
4990                .opgJournal = .opgJournal_optSale.OptionValue
5000                opgJournal_AfterUpdate  ' ** Procedure: Below.
5010                DoEvents
5020                strSubform = "frmJournal_Sub4_Sold"
5030                Set frm = .Controls(strSubform).Form
5040                With frm
5050                  .cmdLock_open_raised_img_dis.Visible = False
5060                  .GoToReport_box_shortname.Visible = True
5070                  .saleShortName.Visible = False
5080                  .GoToReport_arw_tl_acct_img.Top = (.saleAccountNo.Top + (2& * lngTpp))
5090                  .GoToReport_arw_tl_acct_img.Left = (.saleAccountNo.Left + .saleAccountNo.Width)
5100                  .GoToReport_arw_tl_acct_img.Visible = True
5110                  .Controls("saleAccountNo").SetFocus
5120                End With
5130              End Select
5140              DoEvents
5150              .TimerInterval = 50&
5160            End Select
5170          End Select  ' ** blnGoingToReport2.
5180        End If
5190      End If  ' ** gblnGoToReport.
5200      Select Case gstrReturningForm
          Case "frmMap_Div", "frmMap_Div_Detail", "frmMap_Int", "frmMap_Int_Detail", "frmMap_Rec", "frmMap_Rec_Detail", _
              "frmMap_Split", "frmMap_Split_Detail", "frmMap_Misc_LTCL", "frmMap_Misc_LTCL_Detail", _
              "frmMap_Misc_STCGL", "frmMap_Misc_STCGL_Detail", "frmMap_Reinvest_DivInt_Price", "frmMap_Reinvest_DivInt_Detail", _
              "frmMap_Reinvest_Rec_Price", "frmMap_Reinvest_Rec_Detail", _
              "frmReinvest_Dividend", "frmReinvest_Interest", "frmReinvest_Received", "frmReinvest_Sold"
            ' ** From cmdCancel's.
5210        gstrReturningForm = vbNullString
5220        DoCmd.Hourglass False
5230      Case Else
            ' ** Nothing at the moment.
5240      End Select

5250      If gblnTimer = True Then
5260        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
5270          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
5280          gblnTimer = False
5290          DoCmd.Hourglass False
5300        End If
5310      End If

5320    End With

        'HERE! 1  Form_Open()  frmJournal_Sub5_Misc
        'HERE! 2  Form_Load()  frmJournal_Sub5_Misc
        'HERE! 3  Form_Current()  frmJournal_Sub5_Misc
        'HERE! 1  Form_Open()  frmJournal_Sub4_Sold
        'HERE! 2  Form_Load()  frmJournal_Sub4_Sold
        'HERE! 3  Form_Current()  frmJournal_Sub4_Sold
        'HERE! 1  Form_Open()  frmJournal_Sub3_Purchase
        'HERE! 2  Form_Load()  frmJournal_Sub3_Purchase
        'HERE! 3  Form_Current()  frmJournal_Sub3_Purchase
        'HERE! 1  Form_Open()  frmJournal_Sub2_Interest
        'HERE! 2  Form_Load()  frmJournal_Sub2_Interest
        'HERE! 3  Form_Current()  frmJournal_Sub2_Interest
        'HERE! 1  Form_Open()  frmJournal_Sub1_Dividend
        'HERE! 2  Form_Load()  frmJournal_Sub1_Dividend
        'HERE! 3  Form_Current()  frmJournal_Sub1_Dividend
        'HERE! 2  Form_Load()  frmJournal
        'HERE! 3  Form_Timer()  frmJournal

EXITP:
5330    Exit Sub

ERRH:
5340    DoCmd.Hourglass False
5350    blnGoingToReport = False
5360    gblnGoToReport = False
5370    Select Case ERR.Number
        Case Else
5380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5390    End Select
5400    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

5510    If DataErr <> 0 Then
5520      Select Case DataErr
          Case 3315, 2237, 3101, 8519, 2108, 2116, 3020, 2501, 2169, 7753, 3314, 2110
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
5530      Case 2046  ' ** The command or action isn't available now (first or last record).
5540        DoCmd.CancelEvent
5550      Case 3162
5560        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
5570        DoCmd.CancelEvent
5580      Case Else
5590        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
5600        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
5610      End Select
5620    End If

EXITP:
5630    Exit Sub

ERRH:
5640    Select Case ERR.Number
        Case Else
5650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5660    End Select
5670    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim blnRetVal As Boolean
        Dim lngX As Long

5710    With Me

5720      gstrDividendAccountNumber = vbNullString
5730      gstrInterestAccountNumber = vbNullString
5740      gstrPurchaseAccountNumber = vbNullString
5750      gstrSaleAccountNumber = vbNullString
5760      gstrMiscAccountNumber = vbNullString
5770      gstrAccountNo = vbNullString
5780      gstrAccountName = vbNullString

          ' ** Save opgJournal choice (not currently Active), LastAcctNo.
5790      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
          ' ** Though miscCurr_ID has a preference listed, it's not currently being used.
          ' ** I'm not sure how it can be used anyway.

5800      blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

5810      FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
5820      DoEvents

5830      Set dbs = CurrentDb
5840      With dbs
            ' ** Empty tmpAccount.
5850        Set qdf = dbs.QueryDefs("qryMap_09")
5860        qdf.Execute
5870        Set qdf = Nothing
5880        DoEvents
5890        .Close
5900      End With
5910      DoEvents

5920      gblnMessage = False  ' ** frmJournal_Sub1_Dividend may have used this.
5930      blnGoingToReport = False: gblnGoToReport = False

          ' ** Make sure any secondary windows are closed.
5940      For lngX = (Forms.Count - 1) To 0& Step -1&
5950        strDocName = Forms(lngX).Name
5960        If strDocName <> "frmJournal" And strDocName <> "frmMenu_Background" Then
5970          If gblnSwitchTo = False And strDocName = "frmMenu_Post" Then
                ' ** Leave this open if returning to it.
5980          Else
5990            DoCmd.Close acForm, strDocName
6000          End If
6010        End If
6020      Next

          ' ** Make sure all reports are closed.
6030      Do While Reports.Count > 0
6040        DoCmd.Close acReport, Reports(0).Name
6050      Loop

          ' ** Since we're going right to another form, we can't delay these items.
6060      Select Case gblnSwitchTo
          Case True
            ' ** Since they're switching, delete their preference.
6070        Set dbs = CurrentDb
6080        With dbs
              ' ** Delete qryPreferences_16_01 (tblPreference_User, by specified [frmnam], [ctlnam], [usr]).
6090          Set qdf = .QueryDefs("qryPreferences_16_02")
6100          With qdf.Parameters
6110            ![frmnam] = "frmPostingDate"
6120            ![ctlnam] = "opgInput"
6130            ![usr] = CurrentUser ' ** Internal Access Function: Trust Accountant login.
6140          End With
6150          qdf.Execute
6160          .Close
6170        End With  ' ** dbs.
6180        blnRetVal = JrnlCol_Set  ' ** Module Function: modUtilities.
6190      Case False
            ' ** Nothing.
6200      End Select
6210    End With

EXITP:
6220    Set qdf = Nothing
6230    Set dbs = Nothing
6240    Exit Sub

ERRH:
6250    Select Case ERR.Number
        Case Else
6260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6270    End Select
6280    Resume EXITP

End Sub

Private Sub Form_Close()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Close"

        Dim strDocName As String

6310    With Me
6320      Select Case gblnSwitchTo
          Case True
6330        gblnSetFocus = True
6340        strDocName = "frmJournal_Columns"
6350        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
6360      Case False
6370        strDocName = "frmMenu_Background"
6380        gstrReturningForm = THIS_NAME
6390        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
6400      End Select
6410    End With

EXITP:
6420    Exit Sub

ERRH:
6430    Select Case ERR.Number
        Case Else
6440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6450    End Select
6460    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

6510    With Me
6520      If .cmdSwitch_raised_focus_img.Visible = True Then
6530        .cmdSwitch_raised_img.Visible = True
6540        .cmdSwitch_raised_focus_img.Visible = False
6550      ElseIf .cmdSwitch_raised_focus_dots_img.Visible = True Then
6560        .cmdSwitch_raised_semifocus_dots_img.Visible = True
6570        .cmdSwitch_raised_focus_dots_img.Visible = False
6580      End If
6590    End With

EXITP:
6600    Exit Sub

ERRH:
6610    Select Case ERR.Number
        Case Else
6620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6630    End Select
6640    Resume EXITP

End Sub

Private Sub cmdClose_Click()
' ** This is a fake click event.
' ** Called by:
' **   frmJournal (this):
' **     Form_KeyDown()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim strDocName As String

6710    DoCmd.Hourglass True
6720    DoEvents

6730    glngJournalForm = 0&
6740    gstrFormQuerySpec = vbNullString
6750    gstrReturningForm = THIS_NAME
6760    strDocName = "frmMenu_Post"
6770    If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
          ' ** I don't think it'll ever be loaded at this point.
6780      Forms("frmMenu_Post").UpdateTransCnt  ' ** Form Function: frmMenu_Post.
6790    Else
6800      DoCmd.OpenForm strDocName, , , , , , "frmMenu_Main"
6810    End If
6820    DoCmd.Close acForm, THIS_NAME

EXITP:
6830    Exit Sub

ERRH:
6840    DoCmd.Hourglass False
6850    Select Case ERR.Number
        Case Else
6860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6870    End Select
6880    Resume EXITP

End Sub

Private Sub opgJournal_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgJournal_box_MouseMove"

6910    With Me
6920      If .cmdSwitch_raised_focus_img.Visible = True Then
6930        .cmdSwitch_raised_img.Visible = True
6940        .cmdSwitch_raised_focus_img.Visible = False
6950      ElseIf .cmdSwitch_raised_focus_dots_img.Visible = True Then
6960        .cmdSwitch_raised_semifocus_dots_img.Visible = True
6970        .cmdSwitch_raised_focus_dots_img.Visible = False
6980      End If
6990    End With

EXITP:
7000    Exit Sub

ERRH:
7010    Select Case ERR.Number
        Case Else
7020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7030    End Select
7040    Resume EXITP

End Sub

Private Sub opgJournal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgJournal_MouseMove"

7110    With Me
7120      If .cmdSwitch_raised_focus_img.Visible = True Then
7130        .cmdSwitch_raised_img.Visible = True
7140        .cmdSwitch_raised_focus_img.Visible = False
7150      ElseIf .cmdSwitch_raised_focus_dots_img.Visible = True Then
7160        .cmdSwitch_raised_semifocus_dots_img.Visible = True
7170        .cmdSwitch_raised_focus_dots_img.Visible = False
7180      End If
7190    End With

EXITP:
7200    Exit Sub

ERRH:
7210    Select Case ERR.Number
        Case Else
7220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7230    End Select
7240    Resume EXITP

End Sub

Private Sub opgJournal_AfterUpdate()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgJournal_AfterUpdate"

        Dim strDocName As String
        Dim lngX As Long

7310    If blnShowingDividend = False Then

7320      With Me
7330        Application.Echo False
7340        .opgJournal_optMisc_lbl_box.Visible = False
7350        .opgJournal_optSale_lbl_box.Visible = False
7360        .opgJournal_optPurchase_lbl_box.Visible = False
7370        .opgJournal_optInterest_lbl_box.Visible = False
7380        .opgJournal_optDividend_lbl_box.Visible = False
7390        Select Case .opgJournal
            Case .opgJournal_optDividend.OptionValue
              ' ** Dividend was selected.
7400          ShowDividend  ' ** Procedure: Below.
7410        Case .opgJournal_optInterest.OptionValue
              ' ** Interest was selected.
7420          If blnIsOpen = False Then
7430            ShowInterest  ' ** Procedure: Below.
7440          End If
7450        Case .opgJournal_optPurchase.OptionValue
              ' ** Purchase / Deposit was selected..
7460          ShowPurchase  ' ** Procedure: Below.
7470        Case .opgJournal_optSale.OptionValue
              ' ** Sale / Withdrawal was selected.
7480          ShowSale  ' ** Procedure: Below.
7490        Case .opgJournal_optMisc.OptionValue
              ' ** Miscellaneous was selected.
7500          ShowMisc  ' ** Procedure: Below.
7510        End Select
7520      End With

          ' ** Make sure any secondary windows are closed.
7530      For lngX = (Forms.Count - 1) To 0& Step -1&
7540        strDocName = Forms(lngX).Name
7550        If strDocName <> "frmJournal" And strDocName <> "frmMenu_Background" And strDocName <> "frmLoadTimer" Then
7560          DoCmd.Close acForm, strDocName
7570        End If
7580      Next

          ' ** Advance to new record on subform.
7590  On Error Resume Next
7600      DoCmd.GoToRecord , , acNewRec
7610  On Error GoTo ERRH

7620    End If  ' ** blnShowingDividend.

EXITP:
7630    Application.Echo True
7640    Exit Sub

ERRH:
7650    Application.Echo True
7660    blnShowingDividend = False
7670    THAT_PROC = THIS_PROC
7680    That_Erl = Erl
7690    Application.Echo True
7700    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7710    Resume EXITP

End Sub

Private Sub opgJournal_optMisc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgJournal_optMisc_MouseMove"

7810    With Me
7820      If .cmdSwitch_raised_focus_img.Visible = True Then
7830        .cmdSwitch_raised_img.Visible = True
7840        .cmdSwitch_raised_focus_img.Visible = False
7850      ElseIf .cmdSwitch_raised_focus_dots_img.Visible = True Then
7860        .cmdSwitch_raised_semifocus_dots_img.Visible = True
7870        .cmdSwitch_raised_focus_dots_img.Visible = False
7880      End If
7890    End With

EXITP:
7900    Exit Sub

ERRH:
7910    Select Case ERR.Number
        Case Else
7920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7930    End Select
7940    Resume EXITP

End Sub

Public Sub cmdSwitch_Click()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSwitch_Click"

8010    With Me
8020      DoEvents
8030      gblnSwitchTo = True
8040      Select Case .opgJournal
          Case .opgJournal_optDividend.OptionValue
8050        .frmJournal_Sub1_Dividend.Form.cmdDividendClose_Click  ' ** Form Procedure: frmJournal_Sub1_Dividend.
8060      Case .opgJournal_optInterest.OptionValue
8070        .frmJournal_Sub2_Interest.Form.cmdInterestClose_Click  ' ** Form Procedure: frmJournal_Sub2_Interest.
8080      Case .opgJournal_optPurchase.OptionValue
8090        .frmJournal_Sub3_Purchase.Form.cmdPurchaseClose_Click  ' ** Form Procedure: frmJournal_Sub3_Purchase.
8100      Case .opgJournal_optSale.OptionValue
8110        .frmJournal_Sub4_Sold.Form.cmdSaleClose_Click  ' ** Form Procedure: frmJournal_Sub4_Sold.
8120      Case .opgJournal_optMisc.OptionValue
8130        .frmJournal_Sub5_Misc.Form.cmdMiscClose_Click  ' ** Form Procedure: frmJournal_Sub5_Misc.
8140      End Select
8150    End With

EXITP:
8160    Exit Sub

ERRH:
8170    Select Case ERR.Number
        Case Else
8180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8190    End Select
8200    Resume EXITP

End Sub

Private Sub cmdSwitch_GotFocus()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSwitch_GotFocus"

8310    With Me
8320      .cmdSwitch_raised_semifocus_dots_img.Visible = True
8330      .cmdSwitch_raised_img.Visible = False
8340      .cmdSwitch_raised_focus_img.Visible = False
8350    End With

EXITP:
8360    Exit Sub

ERRH:
8370    Select Case ERR.Number
        Case Else
8380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8390    End Select
8400    Resume EXITP

End Sub

Private Sub cmdSwitch_LostFocus()

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSwitch_LostFocus"

8510    With Me
8520      .cmdSwitch_raised_img.Visible = True
8530      .cmdSwitch_raised_semifocus_dots_img.Visible = False
8540      .cmdSwitch_raised_focus_dots_img.Visible = False
8550    End With

EXITP:
8560    Exit Sub

ERRH:
8570    Select Case ERR.Number
        Case Else
8580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8590    End Select
8600    Resume EXITP

End Sub

Private Sub cmdSwitch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSwitch_MouseDown"

8710    With Me
8720      .cmdSwitch_sunken_focus_dots_img.Visible = True
8730      .cmdSwitch_raised_img.Visible = False
8740      .cmdSwitch_raised_semifocus_dots_img.Visible = False
8750      .cmdSwitch_raised_focus_img.Visible = False
8760      .cmdSwitch_raised_focus_dots_img.Visible = False
8770    End With

EXITP:
8780    Exit Sub

ERRH:
8790    Select Case ERR.Number
        Case Else
8800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8810    End Select
8820    Resume EXITP

End Sub

Private Sub cmdSwitch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSwitch_MouseMove"

8910    With Me
8920      If .cmdSwitch_raised_img.Visible = True Then
8930        .cmdSwitch_raised_focus_img.Visible = True
8940        .cmdSwitch_raised_img.Visible = False
8950      ElseIf .cmdSwitch_raised_semifocus_dots_img.Visible = True Then
8960        .cmdSwitch_raised_focus_dots_img.Visible = True
8970        .cmdSwitch_raised_semifocus_dots_img.Visible = False
8980      ElseIf .cmdSwitch_raised_img_dis.Visible = True Then
            ' ** No reaction.
8990      Else
            ' ** Nothing.
9000      End If

9010    End With

EXITP:
9020    Exit Sub

ERRH:
9030    Select Case ERR.Number
        Case Else
9040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9050    End Select
9060    Resume EXITP

End Sub

Private Sub cmdSwitch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSwitch_MouseUp"

9110    With Me
9120      .cmdSwitch_raised_focus_dots_img.Visible = True
9130      .cmdSwitch_raised_img.Visible = False
9140      .cmdSwitch_raised_semifocus_dots_img.Visible = False
9150      .cmdSwitch_raised_focus_img.Visible = False
9160      .cmdSwitch_sunken_focus_dots_img.Visible = False
9170    End With

EXITP:
9180    Exit Sub

ERRH:
9190    Select Case ERR.Number
        Case Else
9200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9210    End Select
9220    Resume EXITP

End Sub

Private Sub frmJournal_Sub_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "frmJournal_Sub_box_MouseMove"

9310    With Me
9320      If .cmdSwitch_raised_focus_img.Visible = True Then
9330        .cmdSwitch_raised_img.Visible = True
9340        .cmdSwitch_raised_focus_img.Visible = False
9350      ElseIf .cmdSwitch_raised_focus_dots_img.Visible = True Then
9360        .cmdSwitch_raised_semifocus_dots_img.Visible = True
9370        .cmdSwitch_raised_focus_dots_img.Visible = False
9380      End If
9390    End With

EXITP:
9400    Exit Sub

ERRH:
9410    Select Case ERR.Number
        Case Else
9420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9430    End Select
9440    Resume EXITP

End Sub

Public Sub JournalVerification()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalVerification"

        Dim dbs As DAO.Database, rst As DAO.Recordset, frm As Access.Form
        Dim strMsg As String, intStyle As Integer, strTitle1 As String, strTitle2 As String
        Dim strAccountNo As String

9510    With Me

9520      strMsg = "Please enter a valid account number to continue."
9530      intStyle = vbInformation + vbOKOnly
9540      strTitle1 = "Entry Required": strTitle2 = "Invalid Entry"

9550      If .frmJournal_Sub1_Dividend.Visible = True Then
9560        Set frm = .frmJournal_Sub1_Dividend.Form
9570        If gblnDividendChanged = True Then
9580          If gblnDividendValidated = False Then
9590            frm.dividendAccountNo.SetFocus
9600            Select Case IsNull(frm.dividendAccountNo)
                Case True
9610              MsgBox strMsg, intStyle, strTitle1
9620              DoCmd.CancelEvent
9630            Case False
                  ' ** Check to see if it is a valid account.
9640              strAccountNo = Trim(frm.dividendAccountNo.text)
9650              Set dbs = CurrentDb
9660              Set rst = dbs.OpenRecordset("SELECT account.shortname FROM account WHERE account.accountno = '" & strAccountNo & "'")
9670              If rst.BOF = True And rst.EOF = True Then
9680                MsgBox strMsg, intStyle, strTitle2
9690              End If
9700              rst.Close
9710              Set rst = Nothing
9720              dbs.Close
9730              Set dbs = Nothing
9740            End Select
9750          End If
9760        End If
9770      ElseIf .frmJournal_Sub2_Interest.Visible = True Then
9780        Set frm = .frmJournal_Sub2_Interest.Form
9790        If gblnInterestChanged = True Then
9800          If gblnInterestValidated = False Then
9810            frm.interestAccountNo.SetFocus
9820            Select Case IsNull(frm.interestAccountNo)
                Case True
9830              MsgBox strMsg, intStyle, strTitle1
9840              DoCmd.CancelEvent
9850            Case False
                  ' ** Check to see if it is a valid account.
9860              strAccountNo = Trim(frm.interestAccountNo.text)
9870              Set dbs = CurrentDb
9880              Set rst = dbs.OpenRecordset("SELECT account.shortname FROM account WHERE account.accountno = '" & strAccountNo & "'")
9890              If rst.BOF = True And rst.EOF = True Then
9900                MsgBox strMsg, intStyle, strTitle2
9910              End If
9920              rst.Close
9930              Set rst = Nothing
9940              dbs.Close
9950              Set dbs = Nothing
9960            End Select
9970          End If
9980        End If
9990      ElseIf .frmJournal_Sub3_Purchase.Visible = True Then
10000       Set frm = .frmJournal_Sub3_Purchase.Form
10010       If gblnPurchaseChanged = True Then
10020         If gblnPurchaseValidated = False Then
10030           frm.purchaseAccountNo.SetFocus
10040           Select Case IsNull(frm.purchaseAccountNo)
                Case True
10050             MsgBox strMsg, intStyle, strTitle1
10060             DoCmd.CancelEvent
10070           Case False
                  ' ** Check to see if it is a valid account.
10080             strAccountNo = Trim(frm.purchaseAccountNo.text)
10090             Set dbs = CurrentDb
10100             Set rst = dbs.OpenRecordset("SELECT account.shortname FROM account WHERE account.accountno = '" & strAccountNo & "'")
10110             If rst.BOF = True And rst.EOF = True Then
10120               MsgBox strMsg, intStyle, strTitle2
10130             End If
10140             rst.Close
10150             Set rst = Nothing
10160             dbs.Close
10170             Set dbs = Nothing
10180           End Select
10190         End If
10200       End If
10210     ElseIf .frmJournal_Sub4_Sold.Visible = True Then
10220       Set frm = .frmJournal_Sub4_Sold.Form
10230       If gblnSaleChanged = True Then
10240         If gblnSaleValidated = False Then
10250           frm.saleAccountNo.SetFocus
10260           Select Case IsNull(frm.saleAccountNo)
                Case True
10270             MsgBox strMsg, intStyle, strTitle1
10280             DoCmd.CancelEvent
10290           Case False
                  ' ** Check to see if it is a valid account.
10300             strAccountNo = Trim(frm.saleAccountNo.text)
10310             Set dbs = CurrentDb
10320             Set rst = dbs.OpenRecordset("SELECT account.shortname FROM account WHERE account.accountno = '" & strAccountNo & "'")
10330             If rst.BOF = True And rst.EOF = True Then
10340               MsgBox strMsg, intStyle, strTitle2
10350             End If
10360             rst.Close
10370             Set rst = Nothing
10380             dbs.Close
10390             Set dbs = Nothing
10400           End Select
10410         End If
10420       End If
10430     ElseIf .frmJournal_Sub5_Misc.Visible = True Then
10440       Set frm = .frmJournal_Sub5_Misc.Form
10450       If gblnMiscChanged = True Then
10460         If gblnMiscValidated = False Then
10470           frm.miscAccountNo.SetFocus
10480           Select Case IsNull(frm.miscAccountNo)
                Case True
10490             MsgBox strMsg, intStyle, strTitle1
10500             DoCmd.CancelEvent
10510           Case False
                  ' ** Check to see if it is a valid account.
10520             strAccountNo = Trim(frm.miscAccountNo.text)
10530             Set dbs = CurrentDb
10540             Set rst = dbs.OpenRecordset("SELECT account.shortname FROM account WHERE account.accountno = '" & strAccountNo & "'")
10550             If rst.BOF = True And rst.EOF = True Then
10560               MsgBox strMsg, intStyle, strTitle2
10570             End If
10580             rst.Close
10590             Set rst = Nothing
10600             dbs.Close
10610             Set dbs = Nothing
10620           End Select
10630         End If
10640       End If
10650     End If
10660   End With

EXITP:
10670   Set frm = Nothing
10680   Set rst = Nothing
10690   Set dbs = Nothing
10700   Exit Sub

ERRH:
10710   THAT_PROC = THIS_PROC
10720   That_Erl = Erl
10730   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10740   Resume EXITP

End Sub

Private Sub Journal_OptLoad()

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "Journal_OptLoad"

        Dim ctl As Access.Control
        Dim lngE As Long

10810   lngOpts = 0&
10820   ReDim arr_varOpt(OPT_ELEMS, 0)

10830   With Me
10840     For Each ctl In .Detail.Controls
10850       With ctl
10860         If .ControlType = acOptionButton Then
10870           If Left(.Name, 14) = "opgJournal_opt" Then
10880             lngOpts = lngOpts + 1&
10890             lngE = lngOpts - 1&
10900             ReDim Preserve arr_varOpt(OPT_ELEMS, lngE)
10910             arr_varOpt(OPT_NAM, lngE) = .Name
10920             arr_varOpt(OPT_VAL, lngE) = .OptionValue
10930           End If
10940         End If
10950       End With
10960     Next
10970   End With

EXITP:
10980   Set ctl = Nothing
10990   Exit Sub

ERRH:
11000   Select Case ERR.Number
        Case Else
11010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11020   End Select
11030   Resume EXITP

End Sub

Public Sub ShowDividend()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowDividend"

        Dim frm As Access.Form
        Dim strFilter As String
        Dim lngX As Long

11110   With Me

11120     blnShowingDividend = True

11130     For lngX = 0& To (lngOpts - 1&)
11140       .Controls(arr_varOpt(OPT_NAM, lngX) & "_lbl").FontBold = False
11150     Next

11160     Set frm = .frmJournal_Sub1_Dividend.Form

11170     frm.frmgfx_id1 = .frmgfx_id1

11180     frm.cmdLock_open_raised_img = .cmdLock_open_raised_img
11190     frm.cmdLock_open_raised_img_dis = .cmdLock_open_raised_img_dis
11200     frm.cmdLock_closed_raised_img = .cmdLock_closed_raised_img

11210     frm.cmdCalendar1_raised_img = .cmdCalendar1_raised_img
11220     frm.cmdCalendar1_raised_semifocus_dots_img = .cmdCalendar1_raised_semifocus_dots_img
11230     frm.cmdCalendar1_raised_focus_img = .cmdCalendar1_raised_focus_img
11240     frm.cmdCalendar1_raised_focus_dots_img = .cmdCalendar1_raised_focus_dots_img
11250     frm.cmdCalendar1_sunken_focus_dots_img = .cmdCalendar1_sunken_focus_dots_img
11260     frm.cmdCalendar1_raised_img_dis = .cmdCalendar1_raised_img_dis

11270     frm.GoToReport_arw_mapdiv_img = .GoToReport_arw_mapdiv_img

11280     frm.tglDividendReinvest_false_raised_img = .tglMiscReinvest_false_raised_img
11290     frm.tglDividendReinvest_false_raised_semifocus_dots_img = .tglMiscReinvest_false_raised_semifocus_dots_img
11300     frm.tglDividendReinvest_false_raised_focus_img = .tglMiscReinvest_false_raised_focus_img
11310     frm.tglDividendReinvest_false_raised_focus_dots_img = .tglMiscReinvest_false_raised_focus_dots_img
11320     frm.tglDividendReinvest_false_sunken_focus_dots_img = .tglMiscReinvest_false_sunken_focus_dots_img
11330     frm.tglDividendReinvest_false_raised_img_dis = .tglMiscReinvest_false_raised_img_dis
11340     frm.tglDividendReinvest_true_raised_img = .tglMiscReinvest_true_raised_img
11350     frm.tglDividendReinvest_true_raised_focus_img = .tglMiscReinvest_true_raised_focus_img
11360     frm.tglDividendReinvest_true_raised_focus_dots_img = .tglMiscReinvest_true_raised_focus_dots_img
11370     frm.tglDividendReinvest_true_sunken_focus_dots_img = .tglMiscReinvest_true_sunken_focus_dots_img
11380     frm.tglDividendReinvest_true_raised_img_dis = .tglMiscReinvest_true_raised_img_dis

11390     .opgJournal = .opgJournal_optDividend.OptionValue
11400     .opgJournal_optDividend_lbl.FontBold = True
11410     glngJournalForm = JRN_DIV
11420     .frmJournal_Sub1_Dividend.Visible = True
11430     DoEvents
11440     DoCmd.SelectObject acForm, "frmJournal", False
11450 On Error Resume Next
11460     .opgJournal.SetFocus
11470     DoEvents
11480     .frmJournal_Sub2_Interest.Visible = False
11490     .frmJournal_Sub3_Purchase.Visible = False
11500     .frmJournal_Sub4_Sold.Visible = False
11510     .frmJournal_Sub5_Misc.Visible = False
11520     .frmJournal_Sub1_Dividend.SetFocus
11530 On Error GoTo ERRH

11540     strFilter = vbNullString
11550     Select Case gblnAdmin
          Case True
11560       If frm.Filter <> vbNullString Then
11570         frm.Filter = strFilter
11580       End If
11590       If frm.FilterOn = True Then frm.FilterOn = False
11600     Case False
11610       strFilter = "[Journal_USER] = '" & CurrentUser & "'"
11620       If frm.Filter <> strFilter Then
11630         frm.Filter = strFilter
11640       End If
11650       If frm.FilterOn = False Then frm.FilterOn = True
11660     End Select

11670     blnShowingDividend = False

11680   End With

EXITP:
11690   Set frm = Nothing
11700   Exit Sub

ERRH:
11710   blnShowingDividend = False
11720   Select Case ERR.Number
        Case Else
11730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11740   End Select
11750   Resume EXITP

End Sub

Public Sub ShowInterest()

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowInterest"

        Dim frm As Access.Form
        Dim strFilter As String
        Dim lngX As Long

11810   With Me

11820     For lngX = 0& To (lngOpts - 1&)
11830       .Controls(arr_varOpt(OPT_NAM, lngX) & "_lbl").FontBold = False
11840     Next

11850     Set frm = .frmJournal_Sub2_Interest.Form

11860     frm.frmgfx_id2 = .frmgfx_id2

11870     frm.cmdLock_open_raised_img = .cmdLock_open_raised_img
11880     frm.cmdLock_open_raised_img_dis = .cmdLock_open_raised_img_dis
11890     frm.cmdLock_closed_raised_img = .cmdLock_closed_raised_img

11900     frm.cmdCalendar1_raised_img = .cmdCalendar1_raised_img
11910     frm.cmdCalendar1_raised_semifocus_dots_img = .cmdCalendar1_raised_semifocus_dots_img
11920     frm.cmdCalendar1_raised_focus_img = .cmdCalendar1_raised_focus_img
11930     frm.cmdCalendar1_raised_focus_dots_img = .cmdCalendar1_raised_focus_dots_img
11940     frm.cmdCalendar1_sunken_focus_dots_img = .cmdCalendar1_sunken_focus_dots_img
11950     frm.cmdCalendar1_raised_img_dis = .cmdCalendar1_raised_img_dis

11960     frm.GoToReport_arw_mapint_img = .GoToReport_arw_mapint_img

11970     frm.tglInterestReinvest_false_raised_img = .tglMiscReinvest_false_raised_img
11980     frm.tglInterestReinvest_false_raised_semifocus_dots_img = .tglMiscReinvest_false_raised_semifocus_dots_img
11990     frm.tglInterestReinvest_false_raised_focus_img = .tglMiscReinvest_false_raised_focus_img
12000     frm.tglInterestReinvest_false_raised_focus_dots_img = .tglMiscReinvest_false_raised_focus_dots_img
12010     frm.tglInterestReinvest_false_sunken_focus_dots_img = .tglMiscReinvest_false_sunken_focus_dots_img
12020     frm.tglInterestReinvest_false_raised_img_dis = .tglMiscReinvest_false_raised_img_dis
12030     frm.tglInterestReinvest_true_raised_img = .tglMiscReinvest_true_raised_img
12040     frm.tglInterestReinvest_true_raised_focus_img = .tglMiscReinvest_true_raised_focus_img
12050     frm.tglInterestReinvest_true_raised_focus_dots_img = .tglMiscReinvest_true_raised_focus_dots_img
12060     frm.tglInterestReinvest_true_sunken_focus_dots_img = .tglMiscReinvest_true_sunken_focus_dots_img
12070     frm.tglInterestReinvest_true_raised_img_dis = .tglMiscReinvest_true_raised_img_dis

12080     .opgJournal = .opgJournal_optInterest.OptionValue
12090     .opgJournal_optInterest_lbl.FontBold = True
12100     glngJournalForm = JRN_INT
12110     .frmJournal_Sub2_Interest.Visible = True
12120     DoEvents
12130     DoCmd.SelectObject acForm, "frmJournal", False
12140 On Error Resume Next
12150     .opgJournal.SetFocus
12160     DoEvents
12170     .frmJournal_Sub1_Dividend.Visible = False
12180     .frmJournal_Sub3_Purchase.Visible = False
12190     .frmJournal_Sub4_Sold.Visible = False
12200     .frmJournal_Sub5_Misc.Visible = False
12210     .frmJournal_Sub2_Interest.SetFocus
12220 On Error GoTo ERRH

12230     strFilter = vbNullString
12240     Select Case gblnAdmin
          Case True
12250       If .frmJournal_Sub2_Interest.Form.Filter <> vbNullString Then
12260         .frmJournal_Sub2_Interest.Form.Filter = strFilter
12270       End If
12280       If .frmJournal_Sub2_Interest.Form.FilterOn = True Then .frmJournal_Sub2_Interest.Form.FilterOn = False
12290     Case False
12300       strFilter = "[Journal_USER] = '" & CurrentUser & "'"
12310       If .frmJournal_Sub2_Interest.Form.Filter <> strFilter Then
12320         .frmJournal_Sub2_Interest.Form.Filter = strFilter
12330       End If
12340       If .frmJournal_Sub2_Interest.Form.FilterOn = False Then .frmJournal_Sub2_Interest.Form.FilterOn = True
12350     End Select

12360   End With

EXITP:
12370   Set frm = Nothing
12380   Exit Sub

ERRH:
12390   Select Case ERR.Number
        Case Else
12400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12410   End Select
12420   Resume EXITP

End Sub

Public Sub ShowPurchase()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowPurchase"

        Dim frm As Access.Form
        Dim strFilter As String
        Dim lngX As Long

12510   With Me

12520     For lngX = 0& To (lngOpts - 1&)
12530       .Controls(arr_varOpt(OPT_NAM, lngX) & "_lbl").FontBold = False
12540     Next

12550     Set frm = .frmJournal_Sub3_Purchase.Form

12560     frm.frmgfx_id3 = .frmgfx_id3

12570     frm.cmdLock_open_raised_img = .cmdLock_open_raised_img
12580     frm.cmdLock_open_raised_img_dis = .cmdLock_open_raised_img_dis
12590     frm.cmdLock_closed_raised_img = .cmdLock_closed_raised_img

12600     frm.cmdCalendar1_raised_img = .cmdCalendar1_raised_img
12610     frm.cmdCalendar1_raised_semifocus_dots_img = .cmdCalendar1_raised_semifocus_dots_img
12620     frm.cmdCalendar1_raised_focus_img = .cmdCalendar1_raised_focus_img
12630     frm.cmdCalendar1_raised_focus_dots_img = .cmdCalendar1_raised_focus_dots_img
12640     frm.cmdCalendar1_sunken_focus_dots_img = .cmdCalendar1_sunken_focus_dots_img
12650     frm.cmdCalendar1_raised_img_dis = .cmdCalendar1_raised_img_dis

12660     frm.cmdCalendar2_raised_img = .cmdCalendar2_raised_img
12670     frm.cmdCalendar2_raised_semifocus_dots_img = .cmdCalendar2_raised_semifocus_dots_img
12680     frm.cmdCalendar2_raised_focus_img = .cmdCalendar2_raised_focus_img
12690     frm.cmdCalendar2_raised_focus_dots_img = .cmdCalendar2_raised_focus_dots_img
12700     frm.cmdCalendar2_sunken_focus_dots_img = .cmdCalendar2_sunken_focus_dots_img
12710     frm.cmdCalendar2_raised_img_dis = .cmdCalendar2_raised_img_dis

12720     frm.GoToReport_arw_mapsplit_img = .GoToReport_arw_mapsplit_img

12730     frm.tglPurchaseSale_false_raised_img = .tglMiscSale_false_raised_img
12740     frm.tglPurchaseSale_false_raised_semifocus_dots_img = .tglMiscSale_false_raised_semifocus_dots_img
12750     frm.tglPurchaseSale_false_raised_focus_img = .tglMiscSale_false_raised_focus_img
12760     frm.tglPurchaseSale_false_raised_focus_dots_img = .tglMiscSale_false_raised_focus_dots_img
12770     frm.tglPurchaseSale_false_sunken_focus_dots_img = .tglMiscSale_false_sunken_focus_dots_img
12780     frm.tglPurchaseSale_false_raised_img_dis = .tglMiscSale_false_raised_img_dis
12790     frm.tglPurchaseSale_true_raised_img = .tglMiscSale_true_raised_img
12800     frm.tglPurchaseSale_true_raised_focus_img = .tglMiscSale_true_raised_focus_img
12810     frm.tglPurchaseSale_true_raised_focus_dots_img = .tglMiscSale_true_raised_focus_dots_img
12820     frm.tglPurchaseSale_true_sunken_focus_dots_img = .tglMiscSale_true_sunken_focus_dots_img
12830     frm.tglPurchaseSale_true_raised_img_dis = .tglMiscSale_true_raised_img_dis

12840     .opgJournal = .opgJournal_optPurchase.OptionValue
12850     .opgJournal_optPurchase_lbl.FontBold = True
12860     glngJournalForm = JRN_PUR
12870     .frmJournal_Sub3_Purchase.Visible = True
12880     DoEvents
12890     DoCmd.SelectObject acForm, "frmJournal", False
12900 On Error Resume Next
12910     .opgJournal.SetFocus
12920     DoEvents
12930     .frmJournal_Sub1_Dividend.Visible = False
12940     .frmJournal_Sub2_Interest.Visible = False
12950     .frmJournal_Sub4_Sold.Visible = False
12960     .frmJournal_Sub5_Misc.Visible = False
12970     .frmJournal_Sub3_Purchase.SetFocus
12980     .frmJournal_Sub3_Purchase.Form.Requery
12990 On Error GoTo ERRH

13000     strFilter = vbNullString
13010     Select Case gblnAdmin
          Case True
13020       If .frmJournal_Sub3_Purchase.Form.Filter <> vbNullString Then
13030         .frmJournal_Sub3_Purchase.Form.Filter = strFilter
13040       End If
13050       If .frmJournal_Sub3_Purchase.Form.FilterOn = True Then .frmJournal_Sub3_Purchase.Form.FilterOn = False
13060     Case False
13070       strFilter = "[Journal_USER] = '" & CurrentUser & "'"
13080       If .frmJournal_Sub3_Purchase.Form.Filter <> strFilter Then
13090         .frmJournal_Sub3_Purchase.Form.Filter = strFilter
13100       End If
13110       If .frmJournal_Sub3_Purchase.Form.FilterOn = False Then .frmJournal_Sub3_Purchase.Form.FilterOn = True
13120     End Select

13130   End With

EXITP:
13140   Set frm = Nothing
13150   Exit Sub

ERRH:
13160   Select Case ERR.Number
        Case Else
13170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13180   End Select
13190   Resume EXITP

End Sub

Public Sub ShowSale()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowSale"

        Dim frm As Access.Form
        Dim strFilter As String
        Dim lngX As Long

13210   With Me

13220     For lngX = 0& To (lngOpts - 1&)
13230       .Controls(arr_varOpt(OPT_NAM, lngX) & "_lbl").FontBold = False
13240     Next

13250     Set frm = .frmJournal_Sub4_Sold.Form

13260     frm.frmgfx_id4 = .frmgfx_id4

13270     frm.cmdLock_open_raised_img = .cmdLock_open_raised_img
13280     frm.cmdLock_open_raised_img_dis = .cmdLock_open_raised_img_dis
13290     frm.cmdLock_closed_raised_img = .cmdLock_closed_raised_img

13300     frm.cmdCalendar1_raised_img = .cmdCalendar1_raised_img
13310     frm.cmdCalendar1_raised_semifocus_dots_img = .cmdCalendar1_raised_semifocus_dots_img
13320     frm.cmdCalendar1_raised_focus_img = .cmdCalendar1_raised_focus_img
13330     frm.cmdCalendar1_raised_focus_dots_img = .cmdCalendar1_raised_focus_dots_img
13340     frm.cmdCalendar1_sunken_focus_dots_img = .cmdCalendar1_sunken_focus_dots_img
13350     frm.cmdCalendar1_raised_img_dis = .cmdCalendar1_raised_img_dis

13360     frm.cmdCalendar2_raised_img = .cmdCalendar2_raised_img
13370     frm.cmdCalendar2_raised_semifocus_dots_img = .cmdCalendar2_raised_semifocus_dots_img
13380     frm.cmdCalendar2_raised_focus_img = .cmdCalendar2_raised_focus_img
13390     frm.cmdCalendar2_raised_focus_dots_img = .cmdCalendar2_raised_focus_dots_img
13400     frm.cmdCalendar2_sunken_focus_dots_img = .cmdCalendar2_sunken_focus_dots_img
13410     frm.cmdCalendar2_raised_img_dis = .cmdCalendar2_raised_img_dis

13420     frm.GoToReport_arw_tl_acct_img = .GoToReport_arw_tl_acct_img
13430     frm.GoToReport_arw_tl_tdate_img = .GoToReport_arw_tl_tdate_img
13440     frm.GoToReport_arw_tl_asset_img = .GoToReport_arw_tl_asset_img
13450     frm.GoToReport_arw_tl_jtype_img = .GoToReport_arw_tl_jtype_img
13460     frm.GoToReport_arw_tl_share_img = .GoToReport_arw_tl_share_img
13470     frm.GoToReport_arw_tl_adate_img = .GoToReport_arw_tl_adate_img
13480     frm.GoToReport_arw_tl_taxlot_img = .GoToReport_arw_tl_taxlot_img

13490     frm.tglSaleReinvest_false_raised_img = .tglMiscReinvest_false_raised_img
13500     frm.tglSaleReinvest_false_raised_semifocus_dots_img = .tglMiscReinvest_false_raised_semifocus_dots_img
13510     frm.tglSaleReinvest_false_raised_focus_img = .tglMiscReinvest_false_raised_focus_img
13520     frm.tglSaleReinvest_false_raised_focus_dots_img = .tglMiscReinvest_false_raised_focus_dots_img
13530     frm.tglSaleReinvest_false_sunken_focus_dots_img = .tglMiscReinvest_false_sunken_focus_dots_img
13540     frm.tglSaleReinvest_false_raised_img_dis = .tglMiscReinvest_false_raised_img_dis
13550     frm.tglSaleReinvest_true_raised_img = .tglMiscReinvest_true_raised_img
13560     frm.tglSaleReinvest_true_raised_focus_img = .tglMiscReinvest_true_raised_focus_img
13570     frm.tglSaleReinvest_true_raised_focus_dots_img = .tglMiscReinvest_true_raised_focus_dots_img
13580     frm.tglSaleReinvest_true_sunken_focus_dots_img = .tglMiscReinvest_true_sunken_focus_dots_img
13590     frm.tglSaleReinvest_true_raised_img_dis = .tglMiscReinvest_true_raised_img_dis

13600     frm.cmdSaleLotInfo.Enabled = False

13610     .opgJournal = .opgJournal_optSale.OptionValue
13620     .opgJournal_optSale_lbl.FontBold = True
13630     glngJournalForm = JRN_SAL
13640     .frmJournal_Sub4_Sold.Visible = True
13650     DoEvents
13660     DoCmd.SelectObject acForm, "frmJournal", False
13670 On Error Resume Next
13680     .opgJournal.SetFocus
13690     DoEvents
13700     .frmJournal_Sub1_Dividend.Visible = False
13710     .frmJournal_Sub2_Interest.Visible = False
13720     .frmJournal_Sub3_Purchase.Visible = False
13730     .frmJournal_Sub5_Misc.Visible = False
13740     .frmJournal_Sub4_Sold.SetFocus
13750 On Error GoTo ERRH

13760     strFilter = vbNullString
13770     Select Case gblnAdmin
          Case True
13780       If .frmJournal_Sub4_Sold.Form.Filter <> vbNullString Then
13790         .frmJournal_Sub4_Sold.Form.Filter = strFilter
13800       End If
13810       If .frmJournal_Sub4_Sold.Form.FilterOn = True Then .frmJournal_Sub4_Sold.Form.FilterOn = False
13820     Case False
13830       strFilter = "[Journal_USER] = '" & CurrentUser & "'"
13840       If .frmJournal_Sub4_Sold.Form.Filter <> strFilter Then
13850         .frmJournal_Sub4_Sold.Form.Filter = strFilter
13860       End If
13870       If .frmJournal_Sub4_Sold.Form.FilterOn = False Then .frmJournal_Sub4_Sold.Form.FilterOn = True
13880     End Select

13890   End With

EXITP:
13900   Set frm = Nothing
13910   Exit Sub

ERRH:
13920   Select Case ERR.Number
        Case Else
13930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13940   End Select
13950   Resume EXITP

End Sub

Public Sub ShowMisc()

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowMisc"

        Dim frm As Access.Form
        Dim strFilter As String
        Dim lngX As Long

14010   With Me

14020     For lngX = 0& To (lngOpts - 1&)
14030       .Controls(arr_varOpt(OPT_NAM, lngX) & "_lbl").FontBold = False
14040     Next

14050     Set frm = .frmJournal_Sub5_Misc.Form

14060     frm.frmgfx_id5 = .frmgfx_id5

14070     frm.cmdLock_open_raised_img = .cmdLock_open_raised_img
14080     frm.cmdLock_open_raised_img_dis = .cmdLock_open_raised_img_dis
14090     frm.cmdLock_closed_raised_img = .cmdLock_closed_raised_img

14100     frm.cmdCalendar1_raised_img = .cmdCalendar1_raised_img
14110     frm.cmdCalendar1_raised_semifocus_dots_img = .cmdCalendar1_raised_semifocus_dots_img
14120     frm.cmdCalendar1_raised_focus_img = .cmdCalendar1_raised_focus_img
14130     frm.cmdCalendar1_raised_focus_dots_img = .cmdCalendar1_raised_focus_dots_img
14140     frm.cmdCalendar1_sunken_focus_dots_img = .cmdCalendar1_sunken_focus_dots_img
14150     frm.cmdCalendar1_raised_img_dis = .cmdCalendar1_raised_img_dis

14160     frm.cmdMiscMap_LTCL_raised_img = .cmdMiscMap_LTCL_raised_img
14170     frm.cmdMiscMap_LTCL_raised_semifocus_dots_img = .cmdMiscMap_LTCL_raised_semifocus_dots_img
14180     frm.cmdMiscMap_LTCL_raised_focus_img = .cmdMiscMap_LTCL_raised_focus_img
14190     frm.cmdMiscMap_LTCL_raised_focus_dots_img = .cmdMiscMap_LTCL_raised_focus_dots_img
14200     frm.cmdMiscMap_LTCL_sunken_focus_dots_img = .cmdMiscMap_LTCL_sunken_focus_dots_img
14210     frm.cmdMiscMap_LTCL_raised_img_dis = .cmdMiscMap_LTCL_raised_img_dis

14220     frm.cmdMiscMap_STCGL_raised_img = .cmdMiscMap_STCGL_raised_img
14230     frm.cmdMiscMap_STCGL_raised_semifocus_dots_img = .cmdMiscMap_STCGL_raised_semifocus_dots_img
14240     frm.cmdMiscMap_STCGL_raised_focus_img = .cmdMiscMap_STCGL_raised_focus_img
14250     frm.cmdMiscMap_STCGL_raised_focus_dots_img = .cmdMiscMap_STCGL_raised_focus_dots_img
14260     frm.cmdMiscMap_STCGL_sunken_focus_dots_img = .cmdMiscMap_STCGL_sunken_focus_dots_img
14270     frm.cmdMiscMap_STCGL_raised_img_dis = .cmdMiscMap_STCGL_raised_img_dis

14280     frm.GoToReport_arw_maprec_img = .GoToReport_arw_maprec_img
14290     frm.GoToReport_arw_mapltcl_img = .GoToReport_arw_mapltcl_img
14300     frm.GoToReport_arw_mapstcgl_img = .GoToReport_arw_mapstcgl_img

14310     frm.tglMiscSale_false_raised_img = .tglMiscSale_false_raised_img
14320     frm.tglMiscSale_false_raised_semifocus_dots_img = .tglMiscSale_false_raised_semifocus_dots_img
14330     frm.tglMiscSale_false_raised_focus_img = .tglMiscSale_false_raised_focus_img
14340     frm.tglMiscSale_false_raised_focus_dots_img = .tglMiscSale_false_raised_focus_dots_img
14350     frm.tglMiscSale_false_sunken_focus_dots_img = .tglMiscSale_false_sunken_focus_dots_img
14360     frm.tglMiscSale_false_raised_img_dis = .tglMiscSale_false_raised_img_dis
14370     frm.tglMiscSale_true_raised_img = .tglMiscSale_true_raised_img
14380     frm.tglMiscSale_true_raised_focus_img = .tglMiscSale_true_raised_focus_img
14390     frm.tglMiscSale_true_raised_focus_dots_img = .tglMiscSale_true_raised_focus_dots_img
14400     frm.tglMiscSale_true_sunken_focus_dots_img = .tglMiscSale_true_sunken_focus_dots_img
14410     frm.tglMiscSale_true_raised_img_dis = .tglMiscSale_true_raised_img_dis
14420     frm.tglMiscReinvest_false_raised_img = .tglMiscReinvest_false_raised_img
14430     frm.tglMiscReinvest_false_raised_semifocus_dots_img = .tglMiscReinvest_false_raised_semifocus_dots_img
14440     frm.tglMiscReinvest_false_raised_focus_img = .tglMiscReinvest_false_raised_focus_img
14450     frm.tglMiscReinvest_false_raised_focus_dots_img = .tglMiscReinvest_false_raised_focus_dots_img
14460     frm.tglMiscReinvest_false_sunken_focus_dots_img = .tglMiscReinvest_false_sunken_focus_dots_img
14470     frm.tglMiscReinvest_false_raised_img_dis = .tglMiscReinvest_false_raised_img_dis
14480     frm.tglMiscReinvest_true_raised_img = .tglMiscReinvest_true_raised_img
14490     frm.tglMiscReinvest_true_raised_focus_img = .tglMiscReinvest_true_raised_focus_img
14500     frm.tglMiscReinvest_true_raised_focus_dots_img = .tglMiscReinvest_true_raised_focus_dots_img
14510     frm.tglMiscReinvest_true_sunken_focus_dots_img = .tglMiscReinvest_true_sunken_focus_dots_img
14520     frm.tglMiscReinvest_true_raised_img_dis = .tglMiscReinvest_true_raised_img_dis

14530     .opgJournal = .opgJournal_optMisc.OptionValue
14540     .opgJournal_optMisc_lbl.FontBold = True
14550     glngJournalForm = JRN_MIS
14560     .frmJournal_Sub5_Misc.Visible = True
14570     DoEvents
14580     DoCmd.SelectObject acForm, "frmJournal", False
14590 On Error Resume Next
14600     .opgJournal.SetFocus
14610     DoEvents
14620     .frmJournal_Sub1_Dividend.Visible = False
14630     .frmJournal_Sub2_Interest.Visible = False
14640     .frmJournal_Sub3_Purchase.Visible = False
14650     .frmJournal_Sub4_Sold.Visible = False
14660     .frmJournal_Sub5_Misc.SetFocus
14670 On Error GoTo ERRH

14680     strFilter = vbNullString
14690     Select Case gblnAdmin
          Case True
14700       If .frmJournal_Sub5_Misc.Form.Filter <> vbNullString Then
14710         .frmJournal_Sub5_Misc.Form.Filter = strFilter
14720       End If
14730       If .frmJournal_Sub5_Misc.Form.FilterOn = True Then .frmJournal_Sub5_Misc.Form.FilterOn = False
14740     Case False
14750       strFilter = "[Journal_USER] = '" & CurrentUser & "'"
14760       If .frmJournal_Sub5_Misc.Form.Filter <> strFilter Then
14770         .frmJournal_Sub5_Misc.Form.Filter = strFilter
14780       End If
14790       If .frmJournal_Sub5_Misc.Form.FilterOn = False Then .frmJournal_Sub5_Misc.Form.FilterOn = True
14800     End Select

14810   End With

EXITP:
14820   Set frm = Nothing
14830   Exit Sub

ERRH:
14840   Select Case ERR.Number
        Case Else
14850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14860   End Select
14870   Resume EXITP

End Sub

Public Sub NavVis(blnVisible As Boolean)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "NavVis"

14910   With Me
14920     Select Case blnVisible
          Case True
14930       .Nav_box01.Visible = True
14940       .Nav_hline01.Visible = True
14950       .Nav_hline02.Visible = True
            'If IsAccess2007 = True Then  ' ** Module Function: modXAccess_07_10_Funcs.
            '  .Nav_hline03.Visible = True
            'End If
14960       .Nav_vline01.Visible = True
14970       .Nav_vline02.Visible = True
14980       .Nav_vline03.Visible = True
14990       .Nav_vline04.Visible = True
15000     Case False
15010       .Nav_box01.Visible = False
15020       .Nav_hline01.Visible = False
15030       .Nav_hline02.Visible = False
15040       .Nav_hline03.Visible = False
15050       .Nav_vline01.Visible = False
15060       .Nav_vline02.Visible = False
15070       .Nav_vline03.Visible = False
15080       .Nav_vline04.Visible = False
15090     End Select
15100   End With

EXITP:
15110   Exit Sub

ERRH:
15120   Select Case ERR.Number
        Case Else
15130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15140   End Select
15150   Resume EXITP

End Sub

Public Sub GTREmblem_Set(blnSetEmblem As Boolean)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Set"

15210   With Me
15220     blnGTR_Emblem = blnSetEmblem
15230     If blnGTR_Emblem = False Then
15240       .TimerInterval = 0&
15250       blnGoingToReport = False
15260       GTREmblem_Off  ' ** Procedure: Below.
15270     End If
15280   End With

EXITP:
15290   Exit Sub

ERRH:
15300   DoCmd.Hourglass False
15310   Select Case ERR.Number
        Case Else
15320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15330   End Select
15340   Resume EXITP

End Sub

Public Function GTREmblem_Get() As Boolean

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Get"

        Dim blnRetVal As Boolean

15410   blnRetVal = blnGTR_Emblem

EXITP:
15420   GTREmblem_Get = blnRetVal
15430   Exit Function

ERRH:
15440   Select Case ERR.Number
        Case Else
15450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15460   End Select
15470   Resume EXITP

End Function

Public Sub GTREmblem_Off()

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

15510   With Me
15520     blnGTR_Emblem = False
15530     For lngX = 1& To 24&
15540       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
15550     Next
15560   End With

EXITP:
15570   Exit Sub

ERRH:
15580   Select Case ERR.Number
        Case Else
15590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15600   End Select
15610   Resume EXITP

End Sub

Public Sub IncludeCurrency()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form
        Dim blnShow As Boolean

15710   With Me

15720     blnShow = Pref_CurrID  ' ** Module Function: modPreferenceFuncs.

15730 On Error Resume Next
15740     Set frm = .frmJournal_Sub1_Dividend.Form
15750     If ERR.Number <> 0 Then
15760 On Error GoTo ERRH
15770     Else
15780 On Error GoTo ERRH
15790       With frm
15800         Select Case blnShow
              Case True
15810           .dividendCurr_ID.Visible = True
15820           .dividendCurr_ID_cmd.Visible = True
15830           .dividendCurr_ID_cmd.Enabled = True
15840         Case False
15850           .dividendCurr_ID.Visible = False
15860           .dividendCurr_ID_cmd.Visible = False
15870           .dividendCurr_ID_cmd.Enabled = False
15880         End Select
15890       End With
15900     End If
15910     Set frm = Nothing

15920 On Error Resume Next
15930     Set frm = .frmJournal_Sub2_Interest.Form
15940     If ERR.Number <> 0 Then
15950 On Error GoTo ERRH
15960     Else
15970 On Error GoTo ERRH
15980       With frm
15990         Select Case blnShow
              Case True
16000           .interestCurr_ID.Visible = True
16010           .interestCurr_ID_cmd.Visible = True
16020           .interestCurr_ID_cmd.Enabled = True
16030         Case False
16040           .interestCurr_ID.Visible = False
16050           .interestCurr_ID_cmd.Visible = False
16060           .interestCurr_ID_cmd.Enabled = False
16070         End Select
16080       End With
16090     End If
16100     Set frm = Nothing

16110 On Error Resume Next
16120     Set frm = .frmJournal_Sub3_Purchase.Form
16130     If ERR.Number <> 0 Then
16140 On Error GoTo ERRH
16150     Else
16160 On Error GoTo ERRH
16170       With frm
16180         Select Case blnShow
              Case True
                ' ** PCash and Cost are where they belong. (At least on opening.)
16190           .purchaseCurr_ID.Visible = True
16200           .purchaseCurr_ID_cmd.Visible = True
16210           .purchaseCurr_ID_cmd.Enabled = True
16220           If .purchasePCash.Left <> .purchasePCash_forex_box.Left Then
16230             .purchasePCash.Left = .purchasePCash_forex_box.Left
16240             .purchasePCash_lbl.Left = .purchasePCash.Left
16250             .purchasePCash_lbl_box.Left = .purchasePCash_lbl.Left
16260             .purchaseCost.Left = .purchaseCost_forext_box.Left
16270             .purchaseCost_lbl.Left = .purchaseCost.Left
16280           End If
16290         Case False
                ' ** Move PCash and Cost to their original locations.
16300           .purchaseCurr_ID.Visible = False
16310           .purchaseCurr_ID_cmd.Visible = False
16320           .purchaseCurr_ID_cmd.Enabled = False
16330           .purchasePCash.Left = .purchasePCash_alt_box.Left
16340           .purchasePCash_lbl.Left = .purchasePCash.Left
16350           .purchasePCash_lbl_box.Left = .purchasePCash_lbl.Left
16360           .purchaseCost.Left = .purchaseCost_alt_box.Left
16370           .purchaseCost_lbl.Left = .purchaseCost.Left
16380         End Select
16390       End With
16400     End If
16410     Set frm = Nothing

16420 On Error Resume Next
16430     Set frm = .frmJournal_Sub4_Sold.Form
16440     If ERR.Number <> 0 Then
16450 On Error GoTo ERRH
16460     Else
16470 On Error GoTo ERRH
16480       With frm
16490         Select Case blnShow
              Case True
                ' ** PCash and Cost are where they belong.
16500           .saleCurr_ID.Visible = True
16510           .saleCurr_ID_cmd.Visible = True
16520           .saleCurr_ID_cmd.Enabled = True
16530         Case False
                ' ** Move PCash and Cost to their original locations.
16540           .saleCurr_ID.Visible = False
16550           .saleCurr_ID_cmd.Visible = False
16560           .saleCurr_ID_cmd.Enabled = False
16570           .salePCash.Left = .salePCash_alt_box.Left
16580           .salePCash_lbl.Left = .salePCash.Left
16590           .salePCash_lbl_box.Left = .salePCash_lbl.Left
16600           .saleCost.Left = .saleCost_alt_box.Left
16610           .saleCost_lbl.Left = .saleCost.Left
16620         End Select
16630       End With
16640     End If
16650     Set frm = Nothing

16660 On Error Resume Next
16670     Set frm = .frmJournal_Sub5_Misc.Form
16680     If ERR.Number <> 0 Then
16690 On Error GoTo ERRH
16700     Else
16710 On Error GoTo ERRH
16720       With frm
16730         Select Case blnShow
              Case True
16740           .miscCurr_ID.Visible = True
16750           .IncludeCurrency True  ' ** Form Procedure: frmJournal_Sub5_Misc.
16760           .miscCurr_ID_cmd.Visible = True
16770           .miscCurr_ID_cmd.Enabled = True
16780         Case False
16790           .miscCurr_ID.Visible = False
16800           .IncludeCurrency False  ' ** Form Procedure: frmJournal_Sub5_Misc.
16810           .miscCurr_ID_cmd.Visible = False
16820           .miscCurr_ID_cmd.Enabled = False
16830         End Select
16840       End With
16850     End If
16860     Set frm = Nothing

16870     If blnShow = True Then
16880       Set dbs = CurrentDb
16890       With dbs
              ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
              ' ** just 'Fixedsys', currfont3_nochar = True.
16900         Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
16910         Set rst = qdf.OpenRecordset
16920         With rst
16930           .MoveLast
16940           lngNoChars = .RecordCount
16950           .MoveFirst
16960           arr_varNoChar = .GetRows(lngNoChars)
                ' *****************************************************
                ' ** Array: arr_varNoChar()
                ' **
                ' **   Field  Element  Name                Constant
                ' **   =====  =======  ==================  ==========
                ' **      1      0     curr_id             N_CID
                ' **      2      1     currsym_name        N_CNAM
                ' **      3      2     font_id             N_FID
                ' **      4      3     font_name           N_FNAM
                ' **      5      4     currfont3_nochar    N_NOCHR
                ' **
                ' *****************************************************
16970           .Close
16980         End With
16990         .Close
17000       End With
17010     End If

17020   End With

EXITP:
17030   Set frm = Nothing
17040   Set rst = Nothing
17050   Set qdf = Nothing
17060   Set dbs = Nothing
17070   Exit Sub

ERRH:
17080   Select Case ERR.Number
        Case Else
17090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17100   End Select
17110   Resume EXITP

End Sub

Public Function NoChar_Get() As Variant
' ** Fixed Width Fonts:
' **   Fixedsys
' **   Miriam Fixed
' **   Courier New

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Get"

        Dim varRetVal As Variant

17210   If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
17220     IncludeCurrency  ' ** Procedure: Above.
17230   End If

17240   varRetVal = arr_varNoChar

EXITP:
17250   NoChar_Get = varRetVal
17260   Exit Function

ERRH:
17270   varRetVal = Null
17280   Select Case ERR.Number
        Case Else
17290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17300   End Select
17310   Resume EXITP

End Function

Public Sub LoadMapDropdowns()

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "LoadMapDropdowns"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnContinue As Boolean

17410   blnContinue = True

17420   Set dbs = CurrentDb
17430   Set rst = dbs.OpenRecordset("tmpAccount", dbOpenDynaset, dbReadOnly)
17440   If rst.BOF = True And rst.EOF = True Then
          ' ** Continue.
17450   Else
17460     blnContinue = False
17470   End If
17480   rst.Close
17490   Set rst = Nothing
17500   DoEvents

17510   If blnContinue = True Then
          ' ** Empty tmpAccount.
17520     Set qdf = dbs.QueryDefs("qryMap_09")
17530     qdf.Execute
17540     Set qdf = Nothing
17550     DoEvents
          ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
          ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
          ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
          ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
          ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
          ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
          ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
          ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
17560     Set qdf = dbs.QueryDefs("qryMap_10")
17570     qdf.Execute
17580     Set qdf = Nothing
17590     DoEvents
17600   End If

17610   dbs.Close
17620   Set dbs = Nothing
17630   DoEvents

EXITP:
17640   Set rst = Nothing
17650   Set qdf = Nothing
17660   Set dbs = Nothing
17670   Exit Sub

ERRH:
17680   Select Case ERR.Number
        Case Else
17690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17700   End Select
17710   Resume EXITP

End Sub
