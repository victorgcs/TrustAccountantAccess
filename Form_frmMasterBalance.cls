VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMasterBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMasterBalance"

'VGC 10/24/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' #######################################

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** Shortcut Alt keys responsive from this form:
' **   Show Acct Master: A {chkShowAcctMast}
' **   Close:            C {cmdClose}
' **   Include Closed:   D {chkIncludeClosed}
' **   Excel:            E {cmdExcel}
' **   Disc Only:        H {chkDiscrepancies}
' **   Fix:              I {fix on frmMasterBalance_Sub}
' **   Name:             M {shortname on frmMasterBalance_Sub}
' **   Select None:      N {cmdSelectNone}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Show Acct Sort:   R {chkShowAcctSort}
' **   Select All:       S {cmdSelectAll}
' **   Update:           U {cmdUpdate}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}
' **   Save Size/Pos:    Z {chkSaveSizePos}

' ** Shortcut F-Keys responsive from this form:
' **   Recalc:           F5 {RecalcTots}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}
' **   Save:             S {cmdSave on frmMasterBalance_Sub}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnSelectAll_Focus As Boolean, blnSelectAll_MouseDown As Boolean
Private blnSelectNone_Focus As Boolean, blnSelectNone_MouseDown As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngSub_Top As Long, lngSub_Left As Long, lngSub_Width As Long, lngSub_Height As Long
Private lngSubBox_Top As Long, lngSubBox_Left As Long, lngSubBox_Width As Long, lngSubBox_Height As Long
Private lngForm_Width As Long, lngClose_Left As Long, lngDetail_Height As Long, lngSizable_Offset As Long
Private lngNavOffset_Top As Long, lngShortMenuOffset_Top As Long, lngChkBoxLbl_Offset As Long
Private lngUpdate_Left As Long, lngUpdateLbl_left As Long, lngSelectAll_Left As Long, lngSelectNone_Left As Long, lngAcctSort_Left As Long
Private lngDiscrepancies_Left As Long, lngDiscrepanciesBox_Left As Long, lngDiscrepanciesBox_Width As Long
Private lngDiscrepanciesBox_Offset As Long, lngDiscrepanciesVLine_Offset As Long, lngPageOf_Left As Long
Private lngRecsSumBox_Left As Long, lngRecsSumBox_Width As Long, lngRecsSumDisc_Left As Long, lngRecsSumFix_Left As Long
Private lngAcctSort_Width As Long, lngAcctMast_Width As Long, lngShortName_Width As Long

Private blnDateEndAfterUpdate As Boolean, blnBackFromKeyDown As Boolean, blnRecalcTots As Boolean, blnHasForeign As Boolean
Private blnIsMaximized As Boolean, blnCmdBars As Boolean
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String, strAccountNo As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim lngMyAdjust_Width As Long, lngMyAdjust_Left As Long
        Dim strFormName As String
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           blnCmdBars = False

170           DoCmd.Hourglass True  ' ** Make sure it's still there.
180           DoEvents

              ' ** Just get these out of the way.
190           .form_width_line.Left = 0&
200           .detail_height_line.Top = 0&
210           .cmdSelectAll_lbl.Top = 0&
220           .cmdSelectAll_lbl.Left = 0&
230           .cmdSelectNone_lbl.Top = 0&
240           .cmdSelectNone_lbl.Left = 0&

250           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
260           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

              ' ** Variables are fed empty, then populated ByRef.
270           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

280           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

290           DoEvents

300           lngForm_Width = .Width
310           lngClose_Left = .cmdClose.Left
320           lngUpdate_Left = .cmdUpdate.Left
330           lngUpdateLbl_left = .cmdUpdate_lbl.Left
340           lngSelectAll_Left = .cmdSelectAll.Left
350           lngSelectNone_Left = .cmdSelectNone.Left
360           lngSub_Top = .frmMasterBalance_Sub.Top
370           lngSub_Left = .frmMasterBalance_Sub.Left
380           lngSub_Width = .frmMasterBalance_Sub.Width
390           lngSub_Height = .frmMasterBalance_Sub.Height
400           lngSubBox_Top = .frmMasterBalance_Sub_box.Top
410           lngSubBox_Left = .frmMasterBalance_Sub_box.Left
420           lngSubBox_Width = .frmMasterBalance_Sub_box.Width
430           lngSubBox_Height = .frmMasterBalance_Sub_box.Height
440           lngDetail_Height = .Detail.Height
450           lngSizable_Offset = .Sizable_lbl1.Left - .cmdClose.Left
460           lngNavOffset_Top = (.frmMasterBalance_Sub_box.Top + .frmMasterBalance_Sub_box.Height) - .Nav_hline01.Top
470           lngShortMenuOffset_Top = (.ShortcutMenu_lbl.Top - (.frmMasterBalance_Sub_box.Top + .frmMasterBalance_Sub_box.Height))
480           lngDiscrepancies_Left = .chkDiscrepancies.Left
490           lngDiscrepanciesBox_Left = .chkDiscrepancies_box.Left
500           lngDiscrepanciesBox_Width = .chkDiscrepancies_box.Width
510           lngDiscrepanciesBox_Offset = (.Detail.Height - .chkDiscrepancies_box.Top)
520           lngDiscrepanciesVLine_Offset = (.chkDiscrepancies_vline03.Left - .chkDiscrepancies_vline01.Left)
530           lngAcctSort_Left = .chkShowAcctSort.Left
540           lngPageOf_Left = .chkPageOf.Left
550           lngChkBoxLbl_Offset = (.chkDiscrepancies_lbl.Left - .chkDiscrepancies.Left)
560           lngRecsSumBox_Left = .RecsSum_box.Left
570           lngRecsSumBox_Width = .RecsSum_box.Width
580           lngRecsSumDisc_Left = .RecsSum_Disc.Left
590           lngRecsSumFix_Left = .RecsSum_Fix.Left
600           lngShortName_Width = .frmMasterBalance_Sub.Form.shortname.Width
610           DoEvents

620           SetNav_Access2007 Me  ' ** Module Procedure: modXAccess_07_10_Funcs.
              'SystemColor_Move Me  ' ** Module Procedure: modSystemColorFuncs.
630           DoEvents

640           gstrFormQuerySpec = THIS_NAME
650           blnDateEndAfterUpdate = False: blnBackFromKeyDown = False: blnRecalcTots = False
660           blnSelectAll_Focus = False: blnSelectAll_MouseDown = False
670           blnSelectNone_Focus = False: blnSelectNone_MouseDown = False

680           CLR_DISABLED_FG = CLR_DKGRY
690           CLR_DISABLED_BG = CLR_LTTEAL

700           If strCallingForm <> vbNullString Then
710             Select Case strCallingForm
                Case "frmAccountProfile"
720               Forms(strCallingForm).Visible = False
730               .chkDiscrepancies.Visible = False
740               .chkShowAcctMast.Visible = False
750               .chkShowAcctSort = False
760               .chkShowAcctSort.Visible = False
770               .cmdUpdate_lbl.Visible = True
780               .RecsSum_Fix.Visible = False
790               .RecsSum_Disc.Visible = False
800               .cmdSelectAll.Enabled = False
810               .cmdSelectAll_raised_img_dis.Visible = True
820               .cmdSelectAll_raised_img.Visible = False
830               .cmdSelectAll_raised_semifocus_dots_img.Visible = False
840               .cmdSelectAll_raised_focus_img.Visible = False
850               .cmdSelectAll_raised_focus_dots_img.Visible = False
860               .cmdSelectAll_sunken_focus_dots_img.Visible = False
870               .cmdSelectNone.Enabled = False
880               .cmdSelectNone_raised_img_dis.Visible = True
890               .cmdSelectNone_raised_img.Visible = False
900               .cmdSelectNone_raised_semifocus_dots_img.Visible = False
910               .cmdSelectNone_raised_focus_img.Visible = False
920               .cmdSelectNone_raised_focus_dots_img.Visible = False
930               .cmdSelectNone_sunken_focus_dots_img.Visible = False
940               .frmMasterBalance_Sub.Form.SetAcctSortCols False, .chkDiscrepancies, strCallingForm  ' ** Form Procedure: frmMasterBalance_Sub.
950             Case "frmMenu_Maintenance"
960               strAccountNo = vbNullString
970               .chkDiscrepancies.Visible = True
980               .chkShowAcctMast.Visible = True
990               .chkShowAcctSort.Visible = True
1000              .cmdUpdate_lbl.Visible = True
1010              .RecsSum_Fix.Visible = True
1020              .RecsSum_Disc.Visible = True
1030              .frmMasterBalance_Sub.Form.SetAcctSortCols True, .chkDiscrepancies, strCallingForm  ' ** Form Procedure: frmMasterBalance_Sub.
1040            Case Else
1050              strAccountNo = vbNullString
1060            End Select
1070          End If  ' ** strCallingForm.
1080          DoEvents

              ' ** Either just above or just below it moves to Form_Load().

1090          .chkDiscrepancies = False  ' ** AfterUpdate() called in Form_Timer().

      #If NoExcel Then
1100          .cmdExcel.Enabled = False
      #Else
1110          .cmdExcel.Enabled = True
      #End If

              ' ** Make sure it's on this form before resizing.
1120          strFormName = vbNullString
1130  On Error Resume Next
1140          Do Until strFormName = THIS_NAME
1150            DoCmd.SelectObject acForm, THIS_NAME, False
1160            strFormName = Screen.ActiveForm.Name
1170          Loop
1180  On Error GoTo ERRH

1190          lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
1200          lngMonitorNum = 1&
1210          EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
1220          If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
1230          If lngMonitorNum > 1& Then
1240            LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
1250          End If

1260          blnIsMaximized = IsMaximized(Me)  ' ** Module Function: modWindowFunctions.
1270          DoEvents

1280          blnIsOpen = False

              ' ** If none of the Boolean's are TRUE yet, this should start the Form_Timer() sequence.
1290          .TimerInterval = 100&

1300        Else
1310          Cancel = -1
1320        End If
1330      Else
1340        Cancel = -1
1350      End If
1360    End With

1370    If Cancel = -1 Then
1380      Beep
1390      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1400    End If

EXITP:
1410    Exit Sub

ERRH:
1420    If strCallingForm = "frmAccountProfile" Then
1430      If Forms(strCallingForm).Visible = False Then
1440        Forms(strCallingForm).Visible = True
1450      End If
1460    End If
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub Form_Load()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strRptPath As String
        Dim lngRecs As Long
        Dim varArchiveCnt As Variant, varClosedCnt As Variant
        Dim strCtlToClear As String

1610    With Me

1620      DoCmd.Hourglass True  ' ** Make sure it's still there.
1630      DoEvents

1640      strCtlToClear = vbNullString
1650      blnHasForeign = False

1660      .RecsSum_Fix.ForeColor = CLR_DISABLED_FG
1670      .RecsSum_Fix.BackColor = CLR_DISABLED_BG
1680      .RecsSum_Disc.ForeColor = CLR_DISABLED_FG
1690      .RecsSum_Disc.BackColor = CLR_DISABLED_BG
1700      DoEvents

          ' ** Check whether there are records in LedgerArchive.
1710      .LedgerArchive_lbl.Visible = False
1720      .LedgerArchive_lbl_box.Visible = False
1730      varArchiveCnt = DCount("*", "LedgerArchive")
1740      If IsNull(varArchiveCnt) = False Then
1750        If varArchiveCnt > 0 Then
1760          .LedgerArchive_lbl.Visible = True
1770          .LedgerArchive_lbl_box.Visible = True
1780        End If
1790      End If
1800      DoEvents

          ' ** Check whether there are closed accounts.
1810      .IsClosed_lbl.Visible = False
1820      .IsClosed_lbl_box.Visible = False
1830      varClosedCnt = DCount("[dateclosed]", "account", "IsNull([dateclosed])=False")
1840      If IsNull(varClosedCnt) = False Then
1850        If varClosedCnt > 0 Then
1860          .chkIncludeClosed.Enabled = True
1870        Else
1880          .chkIncludeClosed.Enabled = False
1890          .chkIncludeClosed = False
1900          .chkIncludeClosed_lbl.Visible = False
1910          .chkIncludeClosed_lbl2.Visible = True
1920          .chkIncludeClosed_lbl2_dim_hi.Visible = True
1930        End If
1940      Else
1950        .chkIncludeClosed.Enabled = False
1960        .chkIncludeClosed_lbl.Visible = False
1970        .chkIncludeClosed_lbl2.Visible = True
1980        .chkIncludeClosed_lbl2_dim_hi.Visible = True
1990      End If
2000      DoEvents

2010      Set dbs = CurrentDb

          ' ** Check whether there are any foreign currencies.
          ' ** qryMasterBalance_09_04 (xx), grouped and summed, by curr_id, with cnt_curr.
2020      Set qdf = dbs.QueryDefs("qryMasterBalance_09_05")
2030      Set rst = qdf.OpenRecordset
2040      With rst
2050        If .BOF = True And .EOF = True Then
              ' ** No transactions!
2060          lngRecs = 0&
2070        Else
2080          .MoveLast
2090          lngRecs = .RecordCount
2100          .MoveFirst
2110          If lngRecs = 1& Then
2120            If ![curr_id] <> 150& Then
                  ' ** Unlikely!
2130              blnHasForeign = True
2140            End If
2150          Else
2160            blnHasForeign = True
2170          End If
2180        End If
2190        .Close
2200      End With
2210      Set rst = Nothing
2220      Set qdf = Nothing
2230      DoEvents

          ' ** Empty tblMasterBalance.
2240      Set qdf = dbs.QueryDefs("qryMasterBalance_12")
2250      qdf.Execute
2260      Set qdf = Nothing
2270      DoEvents

2280      Select Case blnHasForeign
          Case True
            ' ** Append qryMasterBalance_07b_00 (xx) to tblAccountBalance.
2290        Set qdf = dbs.QueryDefs("qryMasterBalance_13b_00")
            ' ** This query links all of these, below!
            'qryMasterBalance_13b_00
            '  qryMasterBalance_07b_00
            '    qryMasterBalance_06_00
            '      Account
            '    qryMasterBalance_05b_00
            '      qryMasterBalance_04b_00
            '        qryMasterBalance_03c_00
            '          qryMasterBalance_02c_00
            '            qryMasterBalance_01c_00
            '              qryMasterBalance_01c_07
            '                qryMasterBalance_01c_07_08
            '                  qryMasterBalance_01c_07_02
            '                    qryMasterBalance_01c_07_01
            '                      qryMasterBalance_01c_01
            '                        Ledger
            '                      Ledger
            '                  qryMasterBalance_01c_07_06                    icash_usd, pcash_usd, cost_usd
            '                    qryMasterBalance_01c_07_05
            '                      qryMasterBalance_01c_07_04
            '                        qryMasterBalance_01c_07_03
            '                          qryMasterBalance_01c_07_01
            '                            qryMasterBalance_01c_01
            '                              Ledger
            '                            Ledger
            '                        tblCurrency_History                     ' ** Make sure these don't get hit twice!
            '                    tblCurrency_History
            '                  qryMasterBalance_01c_07_07                    icash_usd, pcash_usd, cost_usd FROM tblCurrency, NOT HISTORY! ####
            '                    qryMasterBalance_01c_07_03
            '                      qryMasterBalance_01c_07_01
            '                        qryMasterBalance_01c_01
            '                          Ledger
            '                        Ledger
            '                    tblCurrency                                 ' ** Make sure these don't get hit twice!
            '                    qryMasterBalance_01c_07_06  {not in}
            '                      {above}
            '              qryMasterBalance_01c_08
            '                qryMasterBalance_01c_12
            '                  qryMasterBalance_01c_11
            '                    qryMasterBalance_01c_09
            '                      qryMasterBalance_01c_09_08
            '                        qryMasterBalance_01c_09_02
            '                          qryMasterBalance_01c_09_01
            '                            qryMasterBalance_01c_03
            '                              qryMasterBalance_01a_08
            '                                Account
            '                              LedgerArchive
            '                            LedgerArchive
            '                        qryMasterBalance_01c_09_06              icash_usd, pcash_usd, cost_usd
            '                          qryMasterBalance_01c_09_05
            '                            qryMasterBalance_01c_09_04
            '                              qryMasterBalance_01c_09_03
            '                                qryMasterBalance_01c_09_01
            '                                  qryMasterBalance_01c_03
            '                                    qryMasterBalance_01a_08
            '                                      Account
            '                                    LedgerArchive
            '                                  LedgerArchive
            '                              tblCurrency_History               ' ** Make sure these don't get hit twice!
            '                          tblCurrency_History
            '                        qryMasterBalance_01c_09_07              icash_usd, pcash_usd, cost_usd FROM tblCurrency, NOT HISTORY! ####
            '                          qryMasterBalance_01c_09_03
            '                            qryMasterBalance_01c_09_01
            '                              qryMasterBalance_01c_03
            '                                qryMasterBalance_01a_08
            '                                  Account
            '                                LedgerArchive
            '                              LedgerArchive
            '                          tblCurrency                           ' ** Make sure these don't get hit twice!
            '                          qryMasterBalance_01c_09_06  {not in}
            '                            {above}
            '                    qryMasterBalance_01c_10
            '                      qryMasterBalance_01c_10_08
            '                        qryMasterBalance_01c_10_02
            '                          qryMasterBalance_01c_10_01
            '                            qryMasterBalance_01c_04
            '                              qryMasterBalance_01a_08
            '                                Account
            '                              LedgerArchive
            '                            LedgerArchive
            '                        qryMasterBalance_01c_10_06              icash_usd, pcash_usd, cost_usd
            '                          qryMasterBalance_01c_10_05
            '                            qryMasterBalance_01c_10_04
            '                              qryMasterBalance_01c_10_03
            '                                qryMasterBalance_01c_10_01
            '                                  qryMasterBalance_01c_04
            '                                    qryMasterBalance_01a_08
            '                                      Account
            '                                    LedgerArchive
            '                                  LedgerArchive
            '                              tblCurrency_History               ' ** Make sure these don't get hit twice!
            '                          tblCurrency_History
            '                        qryMasterBalance_01c_10_07              icash_usd, pcash_usd, cost_usd FROM tblCurrency, NOT HISTORY! ####
            '                          qryMasterBalance_01c_10_03
            '                            qryMasterBalance_01c_10_01
            '                              qryMasterBalance_01c_04
            '                                qryMasterBalance_01a_08
            '                                  Account
            '                                LedgerArchive
            '                              LedgerArchive
            '                          tblCurrency                           ' ** Make sure these don't get hit twice!
            '                          qryMasterBalance_01c_10_06  {not in}
            '                            {above}
            '        qryMasterBalance_03d_00
            '          qryMasterBalance_02d_00
            '            qryMasterBalance_01d_00
            '              qryMasterBalance_01d_07
            '                qryMasterBalance_01d_07_06
            '                  qryMasterBalance_01d_07_03
            '                    qryMasterBalance_01d_07_02                  icash_usd, pcash_usd, cost_usd  IS THIS ONE OK?
            '                      qryMasterBalance_01d_07_01
            '                        qryMasterBalance_01d_01
            '                          Ledger
            '                        Ledger
            '                      tblCurrency_History                       ' ** Make sure these don't get hit twice!
            '                  qryMasterBalance_01d_07_05                    icash_usd, pcash_usd, cost_usd  NOT CALC'D! ONLY 150?
            '                    qryMasterBalance_01d_07_04
            '                      qryMasterBalance_01d_07_01
            '                        qryMasterBalance_01d_01
            '                          Ledger
            '                        Ledger
            '                      qryMasterBalance_01d_07_03  {not in}
            '                        {above}
            '                    tblCurrency_History                         ' ** Make sure these don't get hit twice!
            '                tblCurrency_History
            '              qryMasterBalance_01d_08
            '                qryMasterBalance_01d_12
            '                  qryMasterBalance_01d_11
            '                    qryMasterBalance_01d_09                     icash_usd, pcash_usd, cost_usd
            '                      qryMasterBalance_01d_09_01
            '                        qryMasterBalance_01d_03
            '                          qryMasterBalance_01a_08
            '                            Account
            '                          LedgerArchive
            '                        LedgerArchive
            '                      tblCurrency_History                       ' ** Make sure these don't get hit twice!
            '                    qryMasterBalance_01d_10                     icash_usd, pcash_usd, cost_usd
            '                      qryMasterBalance_01d_10_01
            '                        qryMasterBalance_01d_04
            '                          qryMasterBalance_01a_08
            '                            Account
            '                          LedgerArchive
            '                        LedgerArchive
            '                      tblCurrencyHistory                        ' ** Make sure these don't get hit twice!
            ' **
            'qryMasterBalance_01c_07_06   icash_usd, pcash_usd, cost_usd
            'qryMasterBalance_01c_07_07   icash_usd, pcash_usd, cost_usd FROM tblCurrency, NOT HISTORY! ####
            'qryMasterBalance_01c_09_06   icash_usd, pcash_usd, cost_usd
            'qryMasterBalance_01c_09_07   icash_usd, pcash_usd, cost_usd FROM tblCurrency, NOT HISTORY! ####
            'qryMasterBalance_01c_10_06   icash_usd, pcash_usd, cost_usd
            'qryMasterBalance_01c_10_07   icash_usd, pcash_usd, cost_usd FROM tblCurrency, NOT HISTORY! ####
            'qryMasterBalance_01d_07_02   icash_usd, pcash_usd, cost_usd  IS THIS ONE OK?
            'qryMasterBalance_01d_07_05   icash_usd, pcash_usd, cost_usd  NOT CALC'D! ONLY 150?
            'qryMasterBalance_01d_09      icash_usd, pcash_usd, cost_usd
            'qryMasterBalance_01d_10      icash_usd, pcash_usd, cost_usd

2300        .ForeignCurrencies_lbl.Visible = True
2310      Case False
            ' ** Append qryMasterBalance_07b (xx) to tblMasterBalance.
2320        Set qdf = dbs.QueryDefs("qryMasterBalance_13b")
2330        .ForeignCurrencies_lbl.Visible = False
2340      End Select
2350      qdf.Execute
2360      Set qdf = Nothing
2370      DoEvents

          ' ******************************************************************************
          ' ** This query is the end of a very long thread, involving all these queries!
          ' ******************************************************************************
          ' ** .._06, linked to .._05b, with, Ix, Px, Cx.
          ' ** qryMasterBalance_07b
          ' ** ======================
          ' ** |  Account, with account_sort, icash, pcash, cost.
          ' ** |> qryMasterBalance_06  ' ** End of line.
          ' ** |  ========================================
          ' ** |
          ' ** |  .._04b, grouped and summed, by accountno.
          ' ** |> qryMasterBalance_05b
          ' **    ======================
          ' **    |  Union of .._03c, .._03d.
          ' **    |> qryMasterBalance_04b
          ' **       ====================
          ' **       |  .._02c, grouped and summed, by accountno.
          ' **       |> qryMasterBalance_03c
          ' **       |  ======================
          ' **       |  |  .._01c, grouped and summed, by accountno, journaltype.
          ' **       |  |> qryMasterBalance_02c
          ' **       |     ======================
          ' **       |     |  Ledger, all entries, Union of .._01c_01, .._01c_02.
          ' **       |     |> qryMasterBalance_01c
          ' **       |        ======================
          ' **       |        |  Ledger (.._01a_01), by specified FormRef('datbeg'), FormRef('datend'), with .._11; Cartesian.
          ' **       |        |> qryMasterBalance_01c_01  ' ** End of line.
          ' **       |        |  ============================================
          ' **       |        |
          ' **       |        |  LedgerArchive (.._01a_02), by specified FormRef('datbeg'), FormRef('datend').
          ' **       |        |> qryMasterBalance_01c_02
          ' **       |           =========================
          ' **       |           |  .._01c_05, with correct account_sort, accountno, by specified FormRef().
          ' **       |           |> qryMasterBalance_01c_06
          ' **       |              =========================
          ' **       |              |  Union of .._01c_03, .._01c_04.
          ' **       |              |> qryMasterBalance_01c_05
          ' **       |                 =========================
          ' **       |                 |  LedgerArchive, linked to .._01a_08, by accountno_type, with .._11; Cartesian.
          ' **       |                 |> qryMasterBalance_01c_03
          ' **       |                 |  =========================
          ' **       |                 |  |  Account, with account_type, account_notype.
          ' **       |                 |  |> qryMasterBalance_01a_08  ' ** End of line.
          ' **       |                 |     ============================================
          ' **       |                 |
          ' **       |                 |  LedgerArchive, linked to .._01a_08, by accountno_notype, with .._11; Cartesian.
          ' **       |                 |> qryMasterBalance_01c_04
          ' **       |                    =========================
          ' **       |                    |  Account, with account_type, account_notype.
          ' **       |                    |> qryMasterBalance_01a_08  ' ** End of line.
          ' **       |                       ============================================
          ' **       |
          ' **       |  .._02d, grouped and summed, by accountno.
          ' **       |> qryMasterBalance_03d
          ' **          ======================
          ' **          |  .._01d, grouped and summed, by accountno, journaltype.
          ' **          |> qryMasterBalance_02d
          ' **             ======================
          ' **             |  Ledger, all entries, Union of .._01d_01, .._01d_02.
          ' **             |> qryMasterBalance_01d
          ' **                ======================
          ' **                |  Ledger (.._01b_01), by specified FormRef('datbeg'), FormRef('datend'), with .._11; Cartesian.
          ' **                |> qryMasterBalance_01d_01  ' ** End of line.
          ' **                |  ============================================
          ' **                |
          ' **                |  LedgerArchive (.._01b_02), by specified FormRef('datbeg'), FormRef('datend').
          ' **                |> qryMasterBalance_01d_02
          ' **                   =========================
          ' **                   |  .._01d_05, with correct account_sort, accountno, by specified FormRef().
          ' **                   |> qryMasterBalance_01d_06
          ' **                      =========================
          ' **                      |  Union of .._01d_03, .._01d_04.
          ' **                      |> qryMasterBalance_01d_05
          ' **                         =========================
          ' **                         |  LedgerArchive, linked to .._01a_08, by accountno_type, with .._11; Cartesian.
          ' **                         |> qryMasterBalance_01d_03
          ' **                         |  =========================
          ' **                         |  |  Account, with account_type, account_notype.
          ' **                         |  |> qryMasterBalance_01a_08  ' ** End of line.
          ' **                         |     ============================================
          ' **                         |
          ' **                         |  LedgerArchive, linked to .._01a_08, by accountno_notype, with .._11; Cartesian.
          ' **                         |> qryMasterBalance_01d_04
          ' **                            =========================
          ' **                            |  Account, with account_type, account_notype.
          ' **                            |> qryMasterBalance_01a_08  ' ** End of line.
          ' **                               ============================================
          ' **
          ' **   .._10, grouped, with Max(acct_len).
          ' **   qryMasterBalance_11
          ' **   =====================
          ' **   |  Account, with Len(accountno).
          ' **   |> qryMasterBalance_10  ' ** End of line.
          ' **      ========================================
          ' **

2380      dbs.Close
2390      Set dbs = Nothing

2400      .frmMasterBalance_Sub.Form.Requery
2410      DoEvents
2420      lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
2430      If lngRecsCur > 0& Then
2440        .RecsSum_Fix = .frmMasterBalance_Sub.Form.RecsSum_Fix
2450        .RecsSum_Disc = .frmMasterBalance_Sub.Form.RecsSum_Disc
2460  On Error Resume Next
            ' ** Got an error on customer's machine, but didn't take all the details!
2470        .frmMasterBalance_Sub.SetFocus
2480        .frmMasterBalance_Sub.Form.shortname.SetFocus
2490  On Error GoTo ERRH
2500      Else
2510        .RecsSum_Fix = CDbl(0)
2520        .RecsSum_Disc = CDbl(0)
2530        DisableSelectBtns  ' ** Procedure: Below.
2540      End If
2550      DoEvents

2560      If strCtlToClear <> vbNullString Then
2570        .Controls(strCtlToClear) = Null
2580      End If

2590      If strCallingForm <> "frmAccountProfile" Then
2600        Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
2610        DoEvents
2620      End If

2630      If IsNull(.UserReportPath) = True Then
            ' ** if the user has a UserReportPath preference from another screen, use that.
2640        Set dbs = CurrentDb
2650        With dbs
              ' ** qryReport_Path_02 (qryReport_Path_01 (tblPreference_Control, just 'UserReportPath'),
              ' ** linked to tblPreference_User, by specified [usr]), grouped by prefuser_text.
2660          Set qdf = .QueryDefs("qryReport_Path_03")  '##dbs_id
2670          With qdf.Parameters
2680            ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2690          End With
2700          Set rst = qdf.OpenRecordset
2710          With rst
2720            If .BOF = True And .EOF = True Then
                  ' ** Nope.
2730            Else
2740              .MoveLast
2750              lngRecs = .RecordCount
2760              .MoveFirst
2770              If lngRecs = 1& Then
                    ' ** If they only have one, or all of them are the same, use it!
2780                strRptPath = ![prefuser_text]
2790              End If
2800            End If
2810            .Close
2820          End With
2830          .Close
2840        End With
2850        If strRptPath <> vbNullString Then
2860          .UserReportPath = strRptPath
2870        End If
2880      End If
2890      DoEvents

2900      If .chkSaveSizePos = True Then
2910        If .frm_top <> 0& And .frm_left <> 0& And .frm_width <> 0& And .frm_height <> 0& Then
2920          DoCmd.MoveSize .frm_left, .frm_top, .frm_width, .frm_height
2930        End If
2940      End If

2950      chkShowAcctSort_AfterUpdate  ' ** Procedure: Below.
2960      DoEvents
2970      chkShowAcctMast_AfterUpdate  ' ** Procedure: Below.
2980      DoEvents
2990      chkPageOf_AfterUpdate  ' ** Procedure: Below.
3000      DoEvents
3010      chkSaveSizePos_AfterUpdate  ' ** Procedure: Below.
3020      DoEvents
3030      chkIncludeClosed_AfterUpdate  ' ** Procedure: Below.

3040      DoCmd.Hourglass False
3050      DoEvents

3060    End With  ' ** Me.

EXITP:
3070    Set rst = Nothing
3080    Set qdf = Nothing
3090    Set dbs = Nothing
3100    Exit Sub

ERRH:
3110    DoCmd.Hourglass False
3120    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Ignore.
3130    Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3220    intShiftDown = (Shift And acShiftMask) > 0
3230    intAltDown = (Shift And acAltMask) > 0
3240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Recalc:           F5 {RecalcTots}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Fix:              I {fix on frmMasterBalance_Sub}
        ' **   Name:             M {shortname on frmMasterBalance_Sub}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}
        ' **   Save:             S {cmdSave on frmMasterBalance_Sub}

        ' ** Plain keys.
3250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3260      Select Case intRetVal
          Case vbKeyF5
3270        intRetVal = 0
3280        RecalcTots  ' ** Procedure: Below.
3290      Case vbKeyEscape
3300        If strCallingForm <> vbNullString Then
3310          intRetVal = 0
3320          cmdClose_Click  ' ** Procedure: Below.
3330        End If
3340      Case vbKeyDown
3350        With Me
3360          intRetVal = 0
3370          lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
3380          If lngRecsCur > 0& Then
3390            .frmMasterBalance_Sub.SetFocus
3400            .frmMasterBalance_Sub.Form.MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmMasterBalance_Sub.
3410          End If
3420        End With
3430      Case vbKeyUp
3440        With Me
3450          intRetVal = 0
3460          lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
3470          If lngRecsCur > 0& Then
3480            .frmMasterBalance_Sub.SetFocus
3490            .frmMasterBalance_Sub.Form.MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmMasterBalance_Sub.
3500          End If
3510        End With
3520      End Select
3530    End If

        ' ** Alt keys.
3540    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
3550      Select Case intRetVal
          Case vbKeyI
3560        With Me
3570          intRetVal = 0
3580          lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
3590          If lngRecsCur > 0& Then
3600            .frmMasterBalance_Sub.SetFocus
3610            .frmMasterBalance_Sub.Form.fix.SetFocus
3620            .frmMasterBalance_Sub.Form.fix = Not (.frmMasterBalance_Sub.Form.fix)
3630          End If
3640        End With
3650      Case vbKeyM
3660        With Me
3670          intRetVal = 0
3680          lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
3690          If lngRecsCur > 0& Then
3700            .frmMasterBalance_Sub.SetFocus
3710            .frmMasterBalance_Sub.Form.shortname.SetFocus
3720          End If
3730        End With
3740      Case vbKeyX
3750        intRetVal = 0
3760        cmdClose_Click  ' ** Procedure: Below.
3770      End Select
3780    End If

        ' ** Ctrl keys.
3790    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3800      Select Case intRetVal
          Case vbKeyG
3810        With Me
3820          intRetVal = 0
3830          If .chkPageOf.Enabled = True Then
3840            .chkPageOf.SetFocus
3850            .chkPageOf = (Not .chkPageOf)
3860            chkPageOf_AfterUpdate  ' ** Procedure: Below.
3870          End If
3880        End With
3890      Case vbKeyS
3900        With Me
3910          intRetVal = 0
3920          .frmMasterBalance_Sub.SetFocus
3930          .frmMasterBalance_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmMasterBalance_Sub.
3940        End With
3950      End Select
3960    End If

        ' ** Ctrl-Shift keys.
3970    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3980      Select Case intRetVal
          Case vbKeyF
3990        With Me
4000          intRetVal = 0
4010          .FocusHolder.SetFocus
4020        End With
4030      End Select
4040    End If

EXITP:
4050    KeyCode = intRetVal
4060    Exit Sub

ERRH:
4070    intRetVal = 0
4080    Select Case ERR.Number
        Case Else
4090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4100    End Select
4110    Resume EXITP

End Sub

Private Sub Form_Resize()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Resize"

        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngSubHeight As Long, lngNavTop As Long, lngShortMenuTop As Long
        Dim lngDiff As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long, lngTmp05 As Long, lngTmp06 As Long, lngTmp07 As Long, lngTmp08 As Long

4210    With Me

4220      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
4230        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
4240      End If

          ' ** Variables are fed empty, then populated ByRef.
4250      GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Procedure: modWindowFunctions.

4260      If lngAcctSort_Width = 0& Or lngAcctMast_Width = 0& Then
4270        lngAcctSort_Width = (.frmMasterBalance_Sub.Form.account_sort.Width + (4& * lngTpp))
4280        lngTmp01 = ((.frmMasterBalance_Sub.Form.ICash.Width + (4& * lngTpp)) + (.frmMasterBalance_Sub.Form.Ix.Width + (4& * lngTpp)))
4290        lngTmp02 = ((.frmMasterBalance_Sub.Form.PCash.Width + (4& * lngTpp)) + (.frmMasterBalance_Sub.Form.Px.Width + (4& * lngTpp)))
4300        lngTmp03 = ((.frmMasterBalance_Sub.Form.Cost.Width + (4& * lngTpp)) + (.frmMasterBalance_Sub.Form.Cx.Width + (4& * lngTpp)))
4310        lngAcctMast_Width = (lngTmp01 + lngTmp02 + lngTmp03)
4320        lngTmp01 = 0&: lngTmp02 = 0&: lngTmp03 = 0&
4330      End If

4340      If ((lngWidth <> lngFrm_Width) Or (lngHeight <> lngFrm_Height)) Then

            ' ** Standard widths.
4350        lngTmp01 = lngFrm_Width
4360        lngTmp02 = (lngFrm_Width - lngAcctMast_Width)
4370        lngTmp03 = (lngFrm_Width - lngAcctSort_Width)
4380        lngTmp04 = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))

            ' ** These are WINDOW widths.
4390        If .chkShowAcctSort = True And .chkShowAcctMast = True Then
              ' ** Original full size.
4400          lngTmp06 = lngTmp01
4410        ElseIf .chkShowAcctSort = True And .chkShowAcctMast = False Then
              ' ** Without master, with sort.
4420          lngTmp06 = lngTmp02
4430        ElseIf .chkShowAcctSort = False And .chkShowAcctMast = True Then
              ' ** With master, without sort.
4440          lngTmp06 = lngTmp03
4450        ElseIf .chkShowAcctSort = False And .chkShowAcctMast = False Then
              ' ** Narrowest, without master, without sort.
4460          lngTmp06 = lngTmp04
4470        End If

4480        If lngWidth < lngTmp04 Then
              ' ** Don't let it shrink narrower than without both options.
4490        Else

4500          lngTmp05 = (lngClose_Left + (lngWidth - lngFrm_Width))  ' ** New cmdClose.Left.
4510          lngTmp05 = (lngTmp05 - (lngClose_Left - lngUpdate_Left))  ' ** New cmdUpdate.Left.

4520          .cmdClose.Left = (lngClose_Left + (lngWidth - lngFrm_Width))
4530          .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
4540          .Sizable_lbl2.Left = (.cmdClose.Left + lngSizable_Offset)
4550          .cmdUpdate.Left = (lngUpdate_Left + (lngWidth - lngFrm_Width))
4560          .frmMasterBalance_Sub.Width = (lngSub_Width + (lngWidth - lngFrm_Width))
4570          .frmMasterBalance_Sub_box.Width = ((lngSub_Width + (lngWidth - lngFrm_Width)) + (2& * lngTpp))
4580          .Nav_box01.Width = .frmMasterBalance_Sub.Width
4590          .Nav_hline03.Width = .frmMasterBalance_Sub.Width
4600          lngTmp05 = ((.cmdClose.Left + .cmdClose.Width) + (9& * lngTpp))
4610          .Header_hline01.Width = lngTmp05
4620          .Header_hline02.Width = lngTmp05
4630          .Header_vline01.Left = lngTmp05
4640          .Header_vline02.Left = lngTmp05
4650          .Detail_hline01.Width = lngTmp05
4660          .Detail_hline02.Width = lngTmp05
4670          .Detail_vline01.Left = lngTmp05
4680          .Detail_vline02.Left = lngTmp05
4690          .Footer_hline01.Width = lngTmp05
4700          .Footer_hline02.Width = lngTmp05
4710          .Footer_vline01.Left = lngTmp05
4720          .Footer_vline02.Left = lngTmp05  ' ** This is FORM width.

              ' ** Shortname starts at 2", let it go down to 1.5", and no limit to wider.
4730          lngTmp07 = (lngTmp06 - (lngFrm_Width - lngForm_Width))  ' ** lngTmp06 is a WINDOW width, lngTmp07 is a FORM width.
4740          If lngTmp05 < lngTmp07 Then
                ' ** The new form width is less than the standard form width.
4750            lngTmp08 = (lngTmp07 - lngTmp05)  ' ** Amount narrower than standard.
4760            If lngTmp08 <= (1440& / 2&) Then  ' ** 1/2 inch.
                  ' ** OK.
4770            Else
4780              lngTmp08 = (1440& / 2&)
4790            End If
4800            lngTmp08 = (lngShortName_Width - lngTmp08)  ' ** New shortname width.
4810            .frmMasterBalance_Sub.Form.SetColWidths lngTmp08  ' ** Form Procedure: frmMasterBalance_Sub.
4820          Else
                ' ** The new form width is greater than the standard form width.
4830            lngTmp08 = (lngTmp05 - lngTmp07)  ' ** Amount wider than standard.
4840            lngTmp08 = (lngShortName_Width + lngTmp08)  ' ** New shortname width.
4850            .frmMasterBalance_Sub.Form.SetColWidths lngTmp08  ' ** Form Procedure: frmMasterBalance_Sub.
4860          End If

4870        End If

4880        lngSubHeight = .frmMasterBalance_Sub.Height
4890        If lngHeight > lngFrm_Height Then

4900          lngDiff = (lngHeight - lngFrm_Height)
4910          .Detail.Height = (lngDetail_Height + lngDiff)
4920          .frmMasterBalance_Sub.Height = (lngSub_Height + lngDiff)
4930          .frmMasterBalance_Sub_box.Height = ((lngSub_Height + lngDiff) + (6& * lngTpp))
4940          lngNavTop = ((.frmMasterBalance_Sub_box.Top + .frmMasterBalance_Sub_box.Height) - lngNavOffset_Top)
4950          .Nav_hline01.Top = lngNavTop
4960          .Nav_hline02.Top = (lngNavTop + lngTpp)
4970          .Nav_hline03.Top = lngNavTop
4980          .Nav_vline01.Top = lngNavTop
4990          .Nav_vline02.Top = lngNavTop
5000          .Nav_vline03.Top = lngNavTop
5010          .Nav_vline04.Top = lngNavTop
5020          .Nav_box01.Top = lngNavTop
5030          lngShortMenuTop = ((.frmMasterBalance_Sub_box.Top + .frmMasterBalance_Sub_box.Height) + lngShortMenuOffset_Top)
5040          .ShortcutMenu_lbl.Top = lngShortMenuTop
5050          .ShortcutMenu_up_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
5060          .ShortcutMenu_down_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
5070          lngTmp05 = (.ShortcutMenu_lbl.Top + .ShortcutMenu_lbl.Height)
5080          .Detail_hline01.Top = lngTmp05
5090          .Detail_hline02.Top = (.Detail_hline01.Top + lngTpp)
5100          .Detail_vline01.Top = .Detail_hline01.Top
5110          .Detail_vline02.Top = .Detail_hline01.Top
5120          lngTmp05 = ((lngDetail_Height + lngDiff) - lngDiscrepanciesBox_Offset)
5130          .RecsSum_box.Top = lngTmp05
5140          .chkDiscrepancies_box.Top = lngTmp05
5150          .RecsSum_Disc.Top = (lngTmp05 + (4& * lngTpp))
5160          .RecsSum_Disc_lbl.Top = .RecsSum_Disc.Top
5170          .RecsSum_Disc_lbl2.Top = (.RecsSum_Disc_lbl.Top - (2& * lngTpp))
5180          .RecsSum_Fix.Top = (.RecsSum_Disc.Top + (18& * lngTpp))
5190          .RecsSum_Fix_lbl.Top = .RecsSum_Fix.Top
5200          .cmdSelectAll.Top = .RecsSum_Disc.Top
5210          .cmdSelectAll_raised_img.Top = .RecsSum_Disc.Top
5220          .cmdSelectAll_raised_semifocus_dots_img.Top = .RecsSum_Disc.Top
5230          .cmdSelectAll_raised_focus_img.Top = .RecsSum_Disc.Top
5240          .cmdSelectAll_raised_focus_dots_img.Top = .RecsSum_Disc.Top
5250          .cmdSelectAll_sunken_focus_dots_img.Top = .RecsSum_Disc.Top
5260          .cmdSelectAll_raised_img_dis.Top = .RecsSum_Disc.Top
5270          .cmdSelectNone.Top = .RecsSum_Fix.Top
5280          .cmdSelectNone_raised_img.Top = .cmdSelectNone.Top
5290          .cmdSelectNone_raised_semifocus_dots_img.Top = .cmdSelectNone.Top
5300          .cmdSelectNone_raised_focus_img.Top = .cmdSelectNone.Top
5310          .cmdSelectNone_raised_focus_dots_img.Top = .cmdSelectNone.Top
5320          .cmdSelectNone_sunken_focus_dots_img.Top = .cmdSelectNone.Top
5330          .cmdSelectNone_raised_img_dis.Top = .cmdSelectNone.Top
5340          .chkDiscrepancies_lbl.Top = .cmdSelectAll.Top
5350          .chkDiscrepancies_lbl2.Top = (.chkDiscrepancies_lbl.Top - lngTpp)  ' ** Not sure why this is.
5360          .chkDiscrepancies_lbl2_dim_hi.Top = .chkDiscrepancies_lbl.Top      ' ** Yes, that's right.
5370          .chkDiscrepancies.Top = (.chkDiscrepancies_lbl.Top + (2& * lngTpp))
5380          .chkDiscrepancies_vline01.Top = (.chkDiscrepancies_box.Top + lngTpp)
5390          .chkDiscrepancies_vline02.Top = .chkDiscrepancies_vline01.Top
5400          .chkDiscrepancies_vline03.Top = .chkDiscrepancies_vline01.Top
5410          .chkDiscrepancies_vline04.Top = .chkDiscrepancies_vline01.Top
5420          .chkIncludeClosed.Top = (.cmdSelectNone.Top + lngTpp)
5430          .chkIncludeClosed_lbl.Top = (.cmdSelectNone.Top - lngTpp)
5440          .chkIncludeClosed_lbl2.Top = (.chkIncludeClosed_lbl.Top - lngTpp)
5450          .chkIncludeClosed_lbl2_dim_hi.Top = (.chkIncludeClosed_lbl2.Top + lngTpp)
5460          .chkShowAcctSort.Top = .chkDiscrepancies.Top
5470          .chkShowAcctSort_lbl.Top = .chkDiscrepancies_lbl.Top
5480          .chkShowAcctMast.Top = .chkIncludeClosed.Top
5490          .chkShowAcctMast_lbl.Top = .chkIncludeClosed_lbl.Top
5500          .chkPageOf.Top = .chkShowAcctSort.Top
5510          .chkPageOf_lbl.Top = .chkShowAcctSort_lbl.Top
5520          .chkSaveSizePos.Top = .chkShowAcctMast.Top
5530          .chkSaveSizePos_lbl.Top = .chkShowAcctMast_lbl.Top
5540          lngTmp05 = (.chkShowAcctSort2.Top - .chkShowAcctSort2_box.Top)
5550          lngTmp06 = (.chkShowAcctMast2.Top - .chkShowAcctSort2.Top)
5560          .chkShowAcctSort2_box.Top = (.Detail_hline02.Top + (2& * lngTpp))
5570          .chkShowAcctSort2.Top = (.chkShowAcctSort2_box.Top + lngTmp05)
5580          .chkShowAcctSort2_lbl.Top = .chkShowAcctSort2.Top
5590          .chkShowAcctMast2.Top = (.chkShowAcctSort2.Top + lngTmp06)
5600          .chkShowAcctMast2_lbl.Top = .chkShowAcctMast2.Top
5610          .chkPageOf2.Top = (.chkShowAcctMast2.Top + lngTmp06)
5620          .chkPageOf2_lbl.Top = .chkPageOf2.Top
5630          .chkSaveSizePos2.Top = (.chkPageOf2.Top + lngTmp06)
5640          .chkSaveSizePos2_lbl.Top = .chkSaveSizePos2.Top

5650        Else
              ' ** .Detail.Height = 5610  '2 recs
              ' ** .Detail.Height = 3630  '1 rec
              ' ** .frmMasterBalance_Sub.Height = 5055  '2 recs
              ' ** .frmMasterBalance_Sub.Height = 2685  '1 rec
5660          If (lngSub_Height - (lngFrm_Height - lngHeight)) > (.frmMasterBalance_Sub.Form.Detail.Height + (19& * lngTpp)) Then

5670            lngDiff = (lngFrm_Height - lngHeight)
5680            .frmMasterBalance_Sub.Height = (lngSub_Height - lngDiff)
5690            .frmMasterBalance_Sub_box.Height = ((lngSub_Height - lngDiff) + (6& * lngTpp))
5700            lngNavTop = ((.frmMasterBalance_Sub_box.Top + .frmMasterBalance_Sub_box.Height) - lngNavOffset_Top)
5710            .Nav_hline01.Top = lngNavTop
5720            .Nav_hline02.Top = (lngNavTop + lngTpp)
5730            .Nav_hline03.Top = lngNavTop
5740            .Nav_vline01.Top = lngNavTop
5750            .Nav_vline02.Top = lngNavTop
5760            .Nav_vline03.Top = lngNavTop
5770            .Nav_vline04.Top = lngNavTop
5780            .Nav_box01.Top = lngNavTop
5790            lngShortMenuTop = ((.frmMasterBalance_Sub_box.Top + .frmMasterBalance_Sub_box.Height) + lngShortMenuOffset_Top)
5800            .ShortcutMenu_lbl.Top = lngShortMenuTop
5810            .ShortcutMenu_up_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
5820            .ShortcutMenu_down_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
5830            lngTmp05 = (.ShortcutMenu_lbl.Top + .ShortcutMenu_lbl.Height)
5840            .Detail_hline01.Top = lngTmp05
5850            .Detail_hline02.Top = (.Detail_hline01.Top + lngTpp)
5860            .Detail_vline01.Top = .Detail_hline01.Top
5870            .Detail_vline02.Top = .Detail_hline01.Top
5880            lngTmp05 = ((.ShortcutMenu_lbl.Top + .ShortcutMenu_lbl.Height) + (10& * lngTpp))  '((lngDetail_Height + lngDiff) - lngDiscrepanciesBox_Offset)
5890            .RecsSum_box.Top = lngTmp05
5900            .chkDiscrepancies_box.Top = lngTmp05
5910            .RecsSum_Disc.Top = (lngTmp05 + (4& * lngTpp))
5920            .RecsSum_Disc_lbl.Top = .RecsSum_Disc.Top
5930            .RecsSum_Disc_lbl2.Top = (.RecsSum_Disc_lbl.Top - (2& * lngTpp))
5940            .RecsSum_Fix.Top = (.RecsSum_Disc.Top + (18& * lngTpp))
5950            .RecsSum_Fix_lbl.Top = .RecsSum_Fix.Top
5960            .cmdSelectAll.Top = .RecsSum_Disc.Top
5970            .cmdSelectAll_raised_img.Top = .RecsSum_Disc.Top
5980            .cmdSelectAll_raised_semifocus_dots_img.Top = .RecsSum_Disc.Top
5990            .cmdSelectAll_raised_focus_img.Top = .RecsSum_Disc.Top
6000            .cmdSelectAll_raised_focus_dots_img.Top = .RecsSum_Disc.Top
6010            .cmdSelectAll_sunken_focus_dots_img.Top = .RecsSum_Disc.Top
6020            .cmdSelectAll_raised_img_dis.Top = .RecsSum_Disc.Top
6030            .cmdSelectNone.Top = .RecsSum_Fix.Top
6040            .cmdSelectNone_raised_img.Top = .cmdSelectNone.Top
6050            .cmdSelectNone_raised_semifocus_dots_img.Top = .cmdSelectNone.Top
6060            .cmdSelectNone_raised_focus_img.Top = .cmdSelectNone.Top
6070            .cmdSelectNone_raised_focus_dots_img.Top = .cmdSelectNone.Top
6080            .cmdSelectNone_sunken_focus_dots_img.Top = .cmdSelectNone.Top
6090            .cmdSelectNone_raised_img_dis.Top = .cmdSelectNone.Top
6100            .chkDiscrepancies_lbl.Top = .cmdSelectAll.Top
6110            .chkDiscrepancies_lbl2.Top = (.chkDiscrepancies_lbl.Top - lngTpp)  ' ** Not sure why this is.
6120            .chkDiscrepancies_lbl2_dim_hi.Top = .chkDiscrepancies_lbl.Top      ' ** Yes, that's right.
6130            .chkDiscrepancies.Top = (.chkDiscrepancies_lbl.Top + (2& * lngTpp))
6140            .chkDiscrepancies_vline01.Top = (.chkDiscrepancies_box.Top + lngTpp)
6150            .chkDiscrepancies_vline02.Top = .chkDiscrepancies_vline01.Top
6160            .chkDiscrepancies_vline03.Top = .chkDiscrepancies_vline01.Top
6170            .chkDiscrepancies_vline04.Top = .chkDiscrepancies_vline01.Top
6180            .chkIncludeClosed.Top = (.cmdSelectNone.Top + lngTpp)
6190            .chkIncludeClosed_lbl.Top = (.cmdSelectNone.Top - lngTpp)
6200            .chkIncludeClosed_lbl2.Top = (.chkIncludeClosed_lbl.Top - lngTpp)
6210            .chkIncludeClosed_lbl2_dim_hi.Top = (.chkIncludeClosed_lbl2.Top + lngTpp)
6220            .chkShowAcctSort.Top = .chkDiscrepancies.Top
6230            .chkShowAcctSort_lbl.Top = .chkDiscrepancies_lbl.Top
6240            .chkShowAcctMast.Top = .chkIncludeClosed.Top
6250            .chkShowAcctMast_lbl.Top = .chkIncludeClosed_lbl.Top
6260            .chkPageOf.Top = .chkShowAcctSort.Top
6270            .chkPageOf_lbl.Top = .chkShowAcctSort_lbl.Top
6280            .chkSaveSizePos.Top = .chkShowAcctMast.Top
6290            .chkSaveSizePos_lbl.Top = .chkShowAcctMast_lbl.Top
6300            lngTmp05 = (.chkShowAcctSort2.Top - .chkShowAcctSort2_box.Top)
6310            lngTmp06 = (.chkShowAcctMast2.Top - .chkShowAcctSort2.Top)
6320            .chkShowAcctSort2_box.Top = (.Detail_hline02.Top + (2& * lngTpp))
6330            .chkShowAcctSort2.Top = (.chkShowAcctSort2_box.Top + lngTmp05)
6340            .chkShowAcctSort2_lbl.Top = .chkShowAcctSort2.Top
6350            .chkShowAcctMast2.Top = (.chkShowAcctSort2.Top + lngTmp06)
6360            .chkShowAcctMast2_lbl.Top = .chkShowAcctMast2.Top
6370            .chkPageOf2.Top = (.chkShowAcctMast2.Top + lngTmp06)
6380            .chkPageOf2_lbl.Top = .chkPageOf2.Top
6390            .chkSaveSizePos2.Top = (.chkPageOf2.Top + lngTmp06)
6400            .chkSaveSizePos2_lbl.Top = .chkSaveSizePos2.Top
6410            If (.Detail.Height - lngDiff) <> (.chkDiscrepancies_box.Top + lngDiscrepanciesBox_Offset) Then
6420              .Detail.Height = (.chkDiscrepancies_box.Top + lngDiscrepanciesBox_Offset)
6430            Else
6440              .Detail.Height = (.Detail.Height - lngDiff)
6450            End If

6460          End If
6470        End If

6480      End If

6490    End With

EXITP:
6500    Exit Sub

ERRH:
6510    Select Case ERR.Number
        Case 2100  ' ** The control or subform control is too large for this location.
          ' ** Ignore.
6520    Case Else
6530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6540    End Select
6550    Resume EXITP

End Sub

Private Sub Form_Timer()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String
        Dim dblRecsSum_Disc As Double
        Dim blnIsErr As Boolean

6610    With Me
6620      .TimerInterval = 0&
6630      If gblnSetFocus = True Then  ' ** This is NOT to be set on the calling menu! It isn't set till later, below.
6640        gblnSetFocus = False
6650        strFormName = vbNullString
6660  On Error Resume Next
6670        Do Until strFormName = THIS_NAME
6680          DoCmd.SelectObject acForm, THIS_NAME, False
6690          strFormName = Screen.ActiveForm.Name
6700        Loop
6710  On Error GoTo ERRH
6720        If blnIsMaximized = False Then
6730          DoCmd.Restore
6740        End If
6750  On Error Resume Next
            ' ** Got an error on customer's machine, but didn't take all the details!
6760        .frmMasterBalance_Sub.SetFocus
6770        .frmMasterBalance_Sub.Form.accountno.SetFocus
6780  On Error GoTo ERRH
            ' ** This sets the TimerInterval to come back here for blnRecalcTots.
6790        chkDiscrepancies_AfterUpdate  ' ** Procedure: Below.
6800        blnCmdBars = True
6810      ElseIf blnRecalcTots = True Then
6820        blnRecalcTots = False
6830        RecalcTots  ' ** Procedure: Below.
6840        If blnCmdBars = True Then
6850          .TimerInterval = 1000&
6860        End If
6870      ElseIf blnCmdBars = True Then
6880        blnCmdBars = False
6890        CmdBars_Hide True  ' ** Module Procedure: modWindowFunctions.
6900      Else
            ' ** This SHOULD be the one hit first on opening.
6910        If strCallingForm <> vbNullString Then
6920          Select Case strCallingForm
              Case "frmAccountProfile"
6930            If .RecsSum_Disc > 0 Then
6940              .cmdUpdate.Enabled = True
6950              .chkDiscrepancies.Enabled = True
6960              .chkDiscrepancies_lbl.Visible = True
6970              .chkDiscrepancies_lbl2.Visible = False
6980              .chkDiscrepancies_lbl2_dim_hi.Visible = False
6990              .cmdUpdate_lbl.Caption = "If you wish to correct the discrepancies found, click the 'Update' button."
7000              .cmdUpdate_lbl2.Visible = False
7010              .frmMasterBalance_Sub.Form.fix = True
7020            Else
7030              .cmdUpdate.Enabled = False
7040              .chkDiscrepancies.Enabled = False
7050              .chkDiscrepancies_lbl.Visible = False
7060              .chkDiscrepancies_lbl2.Visible = True
7070              .chkDiscrepancies_lbl2_dim_hi.Visible = True
7080              .cmdUpdate_lbl.Caption = "No discrepancies found."
7090              .cmdUpdate_lbl2.Visible = False
7100            End If
7110          Case "frmMenu_Maintenance"
7120            lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
7130            If lngRecsCur > 0& Then
7140              .RecsSum_Fix = .frmMasterBalance_Sub.Form.RecsSum_Fix
7150              .RecsSum_Disc = .frmMasterBalance_Sub.Form.RecsSum_Disc
7160  On Error Resume Next
7170              dblRecsSum_Disc = .RecsSum_Disc
7180              If ERR.Number <> 0 Then
7190  On Error GoTo ERRH
7200                blnIsErr = True
7210                Do While blnIsErr = True
7220                  ForcePause 0.25  ' ** Module Function: modCodeUtilities.
7230  On Error Resume Next
7240                  dblRecsSum_Disc = .frmMasterBalance_Sub.Form.RecsSum_Disc
7250                  If ERR.Number <> 0 Then
7260  On Error GoTo ERRH
7270                  Else
7280  On Error GoTo ERRH
7290                    blnIsErr = False
7300                    Exit Do
7310                  End If
7320                Loop
7330              Else
7340  On Error GoTo ERRH
7350              End If
7360            Else
7370              .RecsSum_Fix = CDbl(0)
7380              .RecsSum_Disc = CDbl(0)
7390              dblRecsSum_Disc = CDbl(0)
7400            End If
7410            gblnSetFocus = True
7420            .TimerInterval = 100&
7430          End Select
7440        End If
7450      End If
7460      If gblnGoToReport = True Then
7470        gblnGoToReport = False
7480        .cmdPrint.SetFocus
7490        Beep
7500        DoCmd.Hourglass False
7510        DoEvents
7520      End If
7530    End With  ' ** Me.

        'SO, HOW DO I WANT TO ORCHESTRATE ALL THESE FORM_TIMER() EVENTS?
        'IF THE TIMER IS SET TO GO OFF FIRST, WITH NO TRUE'S, IT'LL HIT THE LAST SECTION, ABOVE.
        'THAT SETS THE TIMER TO 100, WITH gblnSetFocus = TRUE.
        'THAT FINISHES WITH chkDiscrepancies_AfterUpdate, WHICH SETS THE TIMER TO 500, WITH blnRecalcTots = TRUE.
        'THAT FINISHES BY SETTING THE TIMER TO 1000, WITH blnCmdBars ALREADY SET TO TRUE AT THE END OF THE FOCUS SECTION.
        'THE COMMAND BARS ENDS THE SEQUENCE, HAVING HIT ALL SECTIONS!

        'WE NEED TO STOP chkIncludeClosed_AfterUpdate() FROM STARTING THE TIMER WHEN CALLED AT THE END OF FORM_LOAD()!

EXITP:
7540    Exit Sub

ERRH:
7550    Select Case ERR.Number
        Case Else
7560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7570    End Select
7580    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String
        Dim lngTop As Long, lngLeft As Long, lngWidth As Long, lngHeight As Long
        Dim blnRetVal As Boolean

7610    With Me

7620      gstrFormQuerySpec = vbNullString

7630      Select Case .chkSaveSizePos
          Case True
            ' ** Variables are fed empty, then populated ByRef.
7640        GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Function: modWindowFunctions.
7650        .frm_top = lngTop
7660        .frm_left = lngLeft
7670        .frm_width = lngWidth
7680        .frm_height = lngHeight
7690      Case False
7700        .frm_top = 0&
7710        .frm_left = 0&
7720        .frm_width = 0&
7730        .frm_height = 0&
7740      End Select

7750      If strCallingForm <> "frmAccountProfile" Then
7760        Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
7770      End If

7780      blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel

7790      If gblnCBarVis = True Then
7800        CmdBars_Hide False  ' ** Module Procedure: modWindowFunctions.
7810      End If

7820      Select Case strCallingForm
          Case "frmAccountProfile"
7830        If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
7840          gblnMessage = True
7850          Forms(strCallingForm).Visible = True
7860          DoCmd.SelectObject acForm, strCallingForm, False
7870  On Error Resume Next
7880          Forms(strCallingForm).RestoreForm  ' ** Form Procedure: frmAccountProfile.
7890  On Error GoTo ERRH
7900          Forms(strCallingForm).TimerInterval = 250&
7910        Else
7920          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Account"
7930        End If
7940      Case "frmMenu_Maintenance"
7950        gblnSetFocus = True
7960        gstrReturningForm = THIS_NAME
7970        If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
7980          DoCmd.SelectObject acForm, strCallingForm, False
7990  On Error Resume Next
8000          Forms(strCallingForm).RestoreForm
8010  On Error GoTo ERRH
8020          Forms(strCallingForm).TimerInterval = 250&
8030        Else
8040          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Utility"
8050        End If
8060      Case vbNullString
8070        gblnSetFocus = True
8080        gstrReturningForm = THIS_NAME
8090        strDocName = "frmMenu_Maintenance"
8100        If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
8110          DoCmd.SelectObject acForm, strDocName, False
8120          Forms(strDocName).TimerInterval = 250&
8130        Else
8140          DoCmd.OpenForm strDocName, , , , , , "frmMenu_Utility"
8150        End If
8160      Case Else
8170        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
8180      End Select

8190    End With

EXITP:
8200    Exit Sub

ERRH:
8210    Select Case ERR.Number
        Case Else
8220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8230    End Select
8240    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

8310    With Me
8320      If .cmdSelectAll_raised_focus_dots_img.Visible = True Or .cmdSelectAll_raised_focus_img.Visible = True Then
8330        Select Case blnSelectAll_Focus
            Case True
8340          .cmdSelectAll_raised_semifocus_dots_img.Visible = True
8350          .cmdSelectAll_raised_img.Visible = False
8360        Case False
8370          .cmdSelectAll_raised_img.Visible = True
8380          .cmdSelectAll_raised_semifocus_dots_img.Visible = False
8390        End Select
8400        .cmdSelectAll_raised_focus_img.Visible = False
8410        .cmdSelectAll_raised_focus_dots_img.Visible = False
8420        .cmdSelectAll_sunken_focus_dots_img.Visible = False
8430        .cmdSelectAll_raised_img_dis.Visible = False
8440      End If
8450      If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
8460        Select Case blnSelectNone_Focus
            Case True
8470          .cmdSelectNone_raised_semifocus_dots_img.Visible = True
8480          .cmdSelectNone_raised_img.Visible = False
8490        Case False
8500          .cmdSelectNone_raised_img.Visible = True
8510          .cmdSelectNone_raised_semifocus_dots_img.Visible = False
8520        End Select
8530        .cmdSelectNone_raised_focus_img.Visible = False
8540        .cmdSelectNone_raised_focus_dots_img.Visible = False
8550        .cmdSelectNone_sunken_focus_dots_img.Visible = False
8560        .cmdSelectNone_raised_img_dis.Visible = False
8570      End If
8580    End With

EXITP:
8590    Exit Sub

ERRH:
8600    Select Case ERR.Number
        Case Else
8610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8620    End Select
8630    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

8710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8720    Exit Sub

ERRH:
8730    Select Case ERR.Number
        Case Else
8740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8750    End Select
8760    Resume EXITP

End Sub

Public Sub cmdClose_Click()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

8810    DoCmd.Close acForm, THIS_NAME

EXITP:
8820    Exit Sub

ERRH:
8830    Select Case ERR.Number
        Case Else
8840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8850    End Select
8860    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
8990          lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
9000          If lngRecsCur > 0& Then
9010            .frmMasterBalance_Sub.SetFocus
9020            .frmMasterBalance_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmMasterBalance_Sub.
9030            .frmMasterBalance_Sub.Form.account_sort.SetFocus
9040          Else
9050            If .cmdSelectAll.Enabled = True Then
9060              .cmdSelectAll.SetFocus
9070            ElseIf .chkDiscrepancies.Enabled = True And .chkDiscrepancies.Visible = True Then
9080              .chkDiscrepancies.SetFocus
9090            ElseIf .chkIncludeClosed.Enabled = True Then
9100              .chkIncludeClosed.SetFocus
9110            ElseIf .chkShowAcctSort.Enabled = True And .chkShowAcctSort.Visible = True Then
9120              .chkShowAcctSort.SetFocus
9130            ElseIf .chkShowAcctMast.Enabled = True And .chkShowAcctMast.Visible = True Then
9140              .chkShowAcctMast.SetFocus
9150            Else
9160              .chkPageOf.SetFocus
9170            End If
9180          End If
9190        End With
9200      End Select
9210    End If

        ' ** Shift keys.
9220    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9230      Select Case intRetVal
          Case vbKeyTab
9240        With Me
9250          intRetVal = 0
9260          If .cmdUpdate.Enabled = True Then
9270            .cmdUpdate.SetFocus
9280          ElseIf .cmdExcel.Enabled = True Then
9290            .cmdExcel.SetFocus
9300          ElseIf .cmdWord.Enabled = True Then
9310            .cmdWord.SetFocus
9320          Else
9330            .chkSaveSizePos.SetFocus
9340          End If
9350        End With
9360      End Select
9370    End If

EXITP:
9380    KeyCode = intRetVal
9390    Exit Sub

ERRH:
9400    intRetVal = 0
9410    Select Case ERR.Number
        Case Else
9420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9430    End Select
9440    Resume EXITP

End Sub

Private Sub cmdPreview_Click()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

9510    blnIsMaximized = IsMaximized(Me)  ' ** Module Function: modWindowFunctions.

9520    DoCmd.OpenReport "rptMasterBalance", acViewPreview
9530    DoCmd.Maximize
9540    DoCmd.RunCommand acCmdFitToWindow

EXITP:
9550    Exit Sub

ERRH:
9560    Select Case ERR.Number
        Case Else
9570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9580    End Select
9590    Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9620    intShiftDown = (Shift And acShiftMask) > 0
9630    intAltDown = (Shift And acAltMask) > 0
9640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9660      Select Case intRetVal
          Case vbKeyTab
9670        With Me
9680          intRetVal = 0
9690          .cmdPrint.SetFocus
9700        End With
9710      End Select
9720    End If

        ' ** Shift keys.
9730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9740      Select Case intRetVal
          Case vbKeyTab
9750        With Me
9760          intRetVal = 0
9770          lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
9780          If lngRecsCur > 0& Then
9790            .frmMasterBalance_Sub.SetFocus
9800            .frmMasterBalance_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMasterBalance_Sub.
9810            .frmMasterBalance_Sub.Form.Cx.SetFocus
9820          Else
9830            .cmdClose.SetFocus
9840          End If
9850        End With
9860      End Select
9870    End If

EXITP:
9880    KeyCode = intRetVal
9890    Exit Sub

ERRH:
9900    intRetVal = 0
9910    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
9920    Case Else
9930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9940    End Select
9950    Resume EXITP

End Sub

Private Sub cmdPrint_Click()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

10010   strDocName = "rptMasterBalance"
10020   If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
10030     DoCmd.OpenReport strDocName, acViewPreview
10040     DoCmd.Maximize
10050     DoCmd.RunCommand acCmdFitToWindow
10060   Else
          '##GTR_Ref: rptMasterBalance
10070     DoCmd.OpenReport strDocName, acViewNormal
10080   End If

EXITP:
10090   Exit Sub

ERRH:
10100   Select Case ERR.Number
        Case Else
10110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10120   End Select
10130   Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10220   intShiftDown = (Shift And acShiftMask) > 0
10230   intAltDown = (Shift And acAltMask) > 0
10240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10260     Select Case intRetVal
          Case vbKeyTab
10270       With Me
10280         intRetVal = 0
10290         .cmdWord.SetFocus
10300       End With
10310     End Select
10320   End If

        ' ** Shift keys.
10330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10340     Select Case intRetVal
          Case vbKeyTab
10350       With Me
10360         intRetVal = 0
10370         .cmdPreview.SetFocus
10380       End With
10390     End Select
10400   End If

EXITP:
10410   KeyCode = intRetVal
10420   Exit Sub

ERRH:
10430   intRetVal = 0
10440   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
10450   Case Else
10460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10470   End Select
10480   Resume EXITP

End Sub

Private Sub cmdWord_Click()

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

10510   With Me

10520     If IsNull(.UserReportPath) = True Then
10530       strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
10540     Else
10550       strRptPath = .UserReportPath
10560     End If
10570     strRptCap = "rptMasterBalance_" & Format(Date, "yyyymmdd")

10580     strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

10590     If strRptPathFile <> vbNullString Then
10600       strRpt = "rptMasterBalance"
10610       DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
10620       .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
10630     End If

10640   End With

EXITP:
10650   Exit Sub

ERRH:
10660   Select Case ERR.Number
        Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
10670     Beep
10680     MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
10690   Case Else
10700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10710   End Select
10720   Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

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
10890         If .cmdExcel.Enabled = True Then
10900           .cmdExcel.SetFocus
10910         ElseIf .cmdUpdate.Enabled = True Then
10920           .cmdUpdate.SetFocus
10930         Else
10940           .cmdClose.SetFocus
10950         End If
10960       End With
10970     End Select
10980   End If

        ' ** Shift keys.
10990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11000     Select Case intRetVal
          Case vbKeyTab
11010       With Me
11020         intRetVal = 0
11030         .cmdPrint.SetFocus
11040       End With
11050     End Select
11060   End If

EXITP:
11070   KeyCode = intRetVal
11080   Exit Sub

ERRH:
11090   intRetVal = 0
11100   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
11110   Case Else
11120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11130   End Select
11140   Resume EXITP

End Sub

Private Sub cmdExcel_Click()

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

11210   With Me

11220     If IsNull(.UserReportPath) = True Then
11230       strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
11240     Else
11250       strRptPath = .UserReportPath
11260     End If
11270     strRptCap = "rptMasterBalance_" & Format(Date, "yyyymmdd")

11280     strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

11290     If strRptPathFile <> vbNullString Then
            ' ** qryMasterBalance_14 (tblMasterBalance), For Export.
11300       strQry = "qryMasterBalance_18"
11310       DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, True
11320       .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
11330     End If

11340   End With

EXITP:
11350   Exit Sub

ERRH:
11360   Select Case ERR.Number
        Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
11370     Beep
11380     MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
11390   Case Else
11400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11410   End Select
11420   Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11520   intShiftDown = (Shift And acShiftMask) > 0
11530   intAltDown = (Shift And acAltMask) > 0
11540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11560     Select Case intRetVal
          Case vbKeyTab
11570       With Me
11580         intRetVal = 0
11590         If .cmdUpdate.Enabled = True Then
11600           .cmdUpdate.SetFocus
11610         Else
11620           .cmdClose.SetFocus
11630         End If
11640       End With
11650     End Select
11660   End If

        ' ** Shift keys.
11670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11680     Select Case intRetVal
          Case vbKeyTab
11690       With Me
11700         intRetVal = 0
11710         .cmdWord.SetFocus
11720       End With
11730     End Select
11740   End If

EXITP:
11750   KeyCode = intRetVal
11760   Exit Sub

ERRH:
11770   intRetVal = 0
11780   Select Case ERR.Number
        Case Else
11790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11800   End Select
11810   Resume EXITP

End Sub

Private Sub cmdSave_Click()

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

11910   With Me
11920     lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
11930     If lngRecsCur > 0& Then
11940       .frmMasterBalance_Sub.SetFocus
11950       .frmMasterBalance_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmMasterBalance_Sub.
11960     End If
11970   End With

EXITP:
11980   Exit Sub

ERRH:
11990   Select Case ERR.Number
        Case Else
12000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12010   End Select
12020   Resume EXITP

End Sub

Private Sub RecsSum_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "RecsSum_box_MouseMove"

12110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12120   Exit Sub

ERRH:
12130   Select Case ERR.Number
        Case Else
12140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12150   End Select
12160   Resume EXITP

End Sub

Private Sub RecsSum_Disc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "RecsSum_Disc_MouseMove"

12210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12220   Exit Sub

ERRH:
12230   Select Case ERR.Number
        Case Else
12240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12250   End Select
12260   Resume EXITP

End Sub

Private Sub RecsSum_Fix_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "RecsSum_Fix_MouseMove"

12310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12320   Exit Sub

ERRH:
12330   Select Case ERR.Number
        Case Else
12340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12350   End Select
12360   Resume EXITP

End Sub

Private Sub cmdSelectAll_GotFocus()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_GotFocus"

12410   With Me
12420     blnSelectAll_Focus = True
12430     .cmdSelectAll_raised_semifocus_dots_img.Visible = True
12440     .cmdSelectAll_raised_img.Visible = False
12450     .cmdSelectAll_raised_focus_img.Visible = False
12460     .cmdSelectAll_raised_focus_dots_img.Visible = False
12470     .cmdSelectAll_sunken_focus_dots_img.Visible = False
12480     .cmdSelectAll_raised_img_dis.Visible = False
12490   End With

EXITP:
12500   Exit Sub

ERRH:
12510   Select Case ERR.Number
        Case Else
12520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12530   End Select
12540   Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseDown"

12610   With Me
12620     blnSelectAll_MouseDown = True
12630     .cmdSelectAll_sunken_focus_dots_img.Visible = True
12640     .cmdSelectAll_raised_img.Visible = False
12650     .cmdSelectAll_raised_semifocus_dots_img.Visible = False
12660     .cmdSelectAll_raised_focus_img.Visible = False
12670     .cmdSelectAll_raised_focus_dots_img.Visible = False
12680     .cmdSelectAll_raised_img_dis.Visible = False
12690   End With

EXITP:
12700   Exit Sub

ERRH:
12710   Select Case ERR.Number
        Case Else
12720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12730   End Select
12740   Resume EXITP

End Sub

Private Sub cmdSelectAll_Click()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_Click"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long

12810   With Me
12820     Set rst = .frmMasterBalance_Sub.Form.RecordsetClone
12830     With rst
12840       If .BOF = True And .EOF = True Then
              ' ** Shouldn't be here.
12850       Else
12860         .MoveLast
12870         lngRecs = .RecordCount
12880         .MoveFirst
12890         For lngX = 1& To lngRecs
12900           If ![DiscCnt] <> 0& Then
12910             .Edit
12920             ![fix] = True
12930             .Update
12940           End If
12950           If lngX < lngRecs Then .MoveNext
12960         Next
12970       End If
12980       .Close
12990     End With
13000     .frmMasterBalance_Sub.Form.Requery
13010     DoEvents
13020     blnRecalcTots = True
13030     .TimerInterval = 500&
13040   End With

EXITP:
13050   Set rst = Nothing
13060   Exit Sub

ERRH:
13070   Select Case ERR.Number
        Case Else
13080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13090   End Select
13100   Resume EXITP

End Sub

Private Sub cmdSelectAll_KeyDown(KeyCode As Integer, Shift As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13220   intShiftDown = (Shift And acShiftMask) > 0
13230   intAltDown = (Shift And acAltMask) > 0
13240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13260     Select Case intRetVal
          Case vbKeyTab
13270       With Me
13280         intRetVal = 0
13290         .cmdSelectNone.SetFocus
13300       End With
13310     End Select
13320   End If

        ' ** Shift keys.
13330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13340     Select Case intRetVal
          Case vbKeyTab
13350       With Me
13360         intRetVal = 0
13370         lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
13380         If lngRecsCur > 0& Then
13390           .frmMasterBalance_Sub.SetFocus
13400         Else
13410           .cmdClose.SetFocus
13420         End If
13430       End With
13440     End Select
13450   End If

EXITP:
13460   KeyCode = intRetVal
13470   Exit Sub

ERRH:
13480   intRetVal = 0
13490   Select Case ERR.Number
        Case Else
13500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13510   End Select
13520   Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseMove"

13610   With Me
13620     If blnSelectAll_MouseDown = False Then
13630       Select Case blnSelectAll_Focus
            Case True
13640         .cmdSelectAll_raised_focus_dots_img.Visible = True
13650         .cmdSelectAll_raised_focus_img.Visible = False
13660       Case False
13670         .cmdSelectAll_raised_focus_img.Visible = True
13680         .cmdSelectAll_raised_focus_dots_img.Visible = False
13690       End Select
13700       .cmdSelectAll_raised_img.Visible = False
13710       .cmdSelectAll_raised_semifocus_dots_img.Visible = False
13720       .cmdSelectAll_sunken_focus_dots_img.Visible = False
13730       .cmdSelectAll_raised_img_dis.Visible = False
13740     End If
13750     If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
13760       Select Case blnSelectNone_Focus
            Case True
13770         .cmdSelectNone_raised_semifocus_dots_img.Visible = True
13780         .cmdSelectNone_raised_img.Visible = False
13790       Case False
13800         .cmdSelectNone_raised_img.Visible = True
13810         .cmdSelectNone_raised_semifocus_dots_img.Visible = False
13820       End Select
13830       .cmdSelectNone_raised_focus_img.Visible = False
13840       .cmdSelectNone_raised_focus_dots_img.Visible = False
13850       .cmdSelectNone_sunken_focus_dots_img.Visible = False
13860       .cmdSelectNone_raised_img_dis.Visible = False
13870     End If
13880   End With

EXITP:
13890   Exit Sub

ERRH:
13900   Select Case ERR.Number
        Case Else
13910     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13920   End Select
13930   Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseUp"

14010   With Me
14020     .cmdSelectAll_raised_focus_dots_img.Visible = True
14030     .cmdSelectAll_raised_img.Visible = False
14040     .cmdSelectAll_raised_semifocus_dots_img.Visible = False
14050     .cmdSelectAll_raised_focus_img.Visible = False
14060     .cmdSelectAll_sunken_focus_dots_img.Visible = False
14070     .cmdSelectAll_raised_img_dis.Visible = False
14080     blnSelectAll_MouseDown = False
14090   End With

EXITP:
14100   Exit Sub

ERRH:
14110   Select Case ERR.Number
        Case Else
14120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14130   End Select
14140   Resume EXITP

End Sub

Private Sub cmdSelectAll_LostFocus()

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_LostFocus"

14210   With Me
14220     .cmdSelectAll_raised_img.Visible = True
14230     .cmdSelectAll_raised_semifocus_dots_img.Visible = False
14240     .cmdSelectAll_raised_focus_img.Visible = False
14250     .cmdSelectAll_raised_focus_dots_img.Visible = False
14260     .cmdSelectAll_sunken_focus_dots_img.Visible = False
14270     .cmdSelectAll_raised_img_dis.Visible = False
14280     blnSelectAll_Focus = False
14290   End With

EXITP:
14300   Exit Sub

ERRH:
14310   Select Case ERR.Number
        Case Else
14320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14330   End Select
14340   Resume EXITP

End Sub

Private Sub cmdSelectNone_GotFocus()

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_GotFocus"

14410   With Me
14420     blnSelectNone_Focus = True
14430     .cmdSelectNone_raised_semifocus_dots_img.Visible = True
14440     .cmdSelectNone_raised_img.Visible = False
14450     .cmdSelectNone_raised_focus_img.Visible = False
14460     .cmdSelectNone_raised_focus_dots_img.Visible = False
14470     .cmdSelectNone_sunken_focus_dots_img.Visible = False
14480     .cmdSelectNone_raised_img_dis.Visible = False
14490   End With

EXITP:
14500   Exit Sub

ERRH:
14510   Select Case ERR.Number
        Case Else
14520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14530   End Select
14540   Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseDown"

14610   With Me
14620     blnSelectNone_MouseDown = True
14630     .cmdSelectNone_sunken_focus_dots_img.Visible = True
14640     .cmdSelectNone_raised_img.Visible = False
14650     .cmdSelectNone_raised_semifocus_dots_img.Visible = False
14660     .cmdSelectNone_raised_focus_img.Visible = False
14670     .cmdSelectNone_raised_focus_dots_img.Visible = False
14680     .cmdSelectNone_raised_img_dis.Visible = False
14690   End With

EXITP:
14700   Exit Sub

ERRH:
14710   Select Case ERR.Number
        Case Else
14720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14730   End Select
14740   Resume EXITP

End Sub

Private Sub cmdSelectNone_Click()

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_Click"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long

14810   With Me
14820     Set rst = .frmMasterBalance_Sub.Form.RecordsetClone
14830     With rst
14840       If .BOF = True And .EOF = True Then
              ' ** Shouldn't be here.
14850       Else
14860         .MoveLast
14870         lngRecs = .RecordCount
14880         .MoveFirst
14890         For lngX = 1& To lngRecs
14900           .Edit
14910           ![fix] = False
14920           .Update
14930           If lngX < lngRecs Then .MoveNext
14940         Next
14950       End If
14960       .Close
14970     End With
14980     .frmMasterBalance_Sub.Form.Requery
14990     DoEvents
15000     blnRecalcTots = True
15010     .TimerInterval = 500&
15020   End With

EXITP:
15030   Set rst = Nothing
15040   Exit Sub

ERRH:
15050   Select Case ERR.Number
        Case Else
15060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15070   End Select
15080   Resume EXITP

End Sub

Private Sub cmdSelectNone_KeyDown(KeyCode As Integer, Shift As Integer)

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15120   intShiftDown = (Shift And acShiftMask) > 0
15130   intAltDown = (Shift And acAltMask) > 0
15140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15160     Select Case intRetVal
          Case vbKeyTab
15170       With Me
15180         intRetVal = 0
15190         If .chkDiscrepancies.Enabled = True And .chkDiscrepancies.Visible = True Then
15200           .chkDiscrepancies.SetFocus
15210         ElseIf .chkIncludeClosed.Enabled = True Then
15220           .chkIncludeClosed.SetFocus
15230         ElseIf .chkShowAcctSort.Enabled = True And .chkShowAcctSort.Visible = True Then
15240           .chkShowAcctSort.SetFocus
15250         ElseIf .chkShowAcctMast.Enabled = True And .chkShowAcctMast.Visible = True Then
15260           .chkShowAcctMast.SetFocus
15270         Else
15280           .chkPageOf.SetFocus
15290         End If
15300       End With
15310     End Select
15320   End If

        ' ** Shift keys.
15330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15340     Select Case intRetVal
          Case vbKeyTab
15350       With Me
15360         intRetVal = 0
15370         .cmdSelectAll.SetFocus
15380       End With
15390     End Select
15400   End If

EXITP:
15410   KeyCode = intRetVal
15420   Exit Sub

ERRH:
15430   intRetVal = 0
15440   Select Case ERR.Number
        Case Else
15450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15460   End Select
15470   Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseMove"

15510   With Me
15520     If blnSelectNone_MouseDown = False Then
15530       Select Case blnSelectNone_Focus
            Case True
15540         .cmdSelectNone_raised_focus_dots_img.Visible = True
15550         .cmdSelectNone_raised_focus_img.Visible = False
15560       Case False
15570         .cmdSelectNone_raised_focus_img.Visible = True
15580         .cmdSelectNone_raised_focus_dots_img.Visible = False
15590       End Select
15600       .cmdSelectNone_raised_img.Visible = False
15610       .cmdSelectNone_raised_semifocus_dots_img.Visible = False
15620       .cmdSelectNone_sunken_focus_dots_img.Visible = False
15630       .cmdSelectNone_raised_img_dis.Visible = False
15640     End If
15650     If .cmdSelectAll_raised_focus_dots_img.Visible = True Or .cmdSelectAll_raised_focus_img.Visible = True Then
15660       Select Case blnSelectAll_Focus
            Case True
15670         .cmdSelectAll_raised_semifocus_dots_img.Visible = True
15680         .cmdSelectAll_raised_img.Visible = False
15690       Case False
15700         .cmdSelectAll_raised_img.Visible = True
15710         .cmdSelectAll_raised_semifocus_dots_img.Visible = False
15720       End Select
15730       .cmdSelectAll_raised_focus_img.Visible = False
15740       .cmdSelectAll_raised_focus_dots_img.Visible = False
15750       .cmdSelectAll_sunken_focus_dots_img.Visible = False
15760       .cmdSelectAll_raised_img_dis.Visible = False
15770     End If
15780   End With

EXITP:
15790   Exit Sub

ERRH:
15800   Select Case ERR.Number
        Case Else
15810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15820   End Select
15830   Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseUp"

15910   With Me
15920     .cmdSelectNone_raised_focus_dots_img.Visible = True
15930     .cmdSelectNone_raised_img.Visible = False
15940     .cmdSelectNone_raised_semifocus_dots_img.Visible = False
15950     .cmdSelectNone_raised_focus_img.Visible = False
15960     .cmdSelectNone_sunken_focus_dots_img.Visible = False
15970     .cmdSelectNone_raised_img_dis.Visible = False
15980     blnSelectNone_MouseDown = False
15990   End With

EXITP:
16000   Exit Sub

ERRH:
16010   Select Case ERR.Number
        Case Else
16020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16030   End Select
16040   Resume EXITP

End Sub

Private Sub cmdSelectNone_LostFocus()

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_LostFocus"

16110   With Me
16120     .cmdSelectNone_raised_img.Visible = True
16130     .cmdSelectNone_raised_semifocus_dots_img.Visible = False
16140     .cmdSelectNone_raised_focus_img.Visible = False
16150     .cmdSelectNone_raised_focus_dots_img.Visible = False
16160     .cmdSelectNone_sunken_focus_dots_img.Visible = False
16170     .cmdSelectNone_raised_img_dis.Visible = False
16180     blnSelectNone_Focus = False
16190   End With

EXITP:
16200   Exit Sub

ERRH:
16210   Select Case ERR.Number
        Case Else
16220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16230   End Select
16240   Resume EXITP

End Sub

Private Sub chkDiscrepancies_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkDiscrepancies_box_MouseMove"

16310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
16320   Exit Sub

ERRH:
16330   Select Case ERR.Number
        Case Else
16340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16350   End Select
16360   Resume EXITP

End Sub

Private Sub chkDiscrepancies_AfterUpdate()

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkDiscrepancies_AfterUpdate"

        Dim strFilter As String

16410   With Me
16420     strFilter = vbNullString
16430     Select Case .chkDiscrepancies
          Case True
16440       .chkDiscrepancies_lbl.FontBold = True
16450       strFilter = "[Ix] = 'X' Or [Px] = 'X' Or [Cx] = 'X'"
16460       If .chkIncludeClosed = False Then
16470         strFilter = "(" & strFilter & ") And [IsClosed] = False"
16480       End If
16490       .frmMasterBalance_Sub.Form.Filter = strFilter
16500       .frmMasterBalance_Sub.Form.FilterOn = True
16510     Case False
16520       .chkDiscrepancies_lbl.FontBold = False
16530       If strCallingForm <> "frmAccountProfile" Then
16540         If .chkIncludeClosed = False Then
16550           strFilter = "[IsClosed] = False"
16560         End If
16570         .frmMasterBalance_Sub.Form.Filter = strFilter
16580         If strFilter = vbNullString Then
16590           .frmMasterBalance_Sub.Form.FilterOn = False
16600         Else
16610           .frmMasterBalance_Sub.Form.FilterOn = True
16620         End If
16630       End If
16640     End Select
16650     DoEvents
16660     .frmMasterBalance_Sub.Form.Requery
16670     DoEvents
16680     blnRecalcTots = True
16690     .TimerInterval = 500&
16700   End With

EXITP:
16710   Exit Sub

ERRH:
16720   Select Case ERR.Number
        Case Else
16730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16740   End Select
16750   Resume EXITP

End Sub

Private Sub chkDiscrepancies_KeyDown(KeyCode As Integer, Shift As Integer)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkDiscrepancies_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16820   intShiftDown = (Shift And acShiftMask) > 0
16830   intAltDown = (Shift And acAltMask) > 0
16840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16860     Select Case intRetVal
          Case vbKeyTab
16870       With Me
16880         intRetVal = 0
16890         If .chkIncludeClosed.Enabled = True Then
16900           .chkIncludeClosed.SetFocus
16910         ElseIf .chkShowAcctSort.Enabled = True And .chkShowAcctSort.Visible = True Then
16920           .chkShowAcctSort.SetFocus
16930         ElseIf .chkShowAcctMast.Enabled = True And .chkShowAcctMast.Visible = True Then
16940           .chkShowAcctMast.SetFocus
16950         Else
16960           .chkPageOf.SetFocus
16970         End If
16980       End With
16990     End Select
17000   End If

        ' ** Shift keys.
17010   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17020     Select Case intRetVal
          Case vbKeyTab
17030       With Me
17040         intRetVal = 0
17050         If .cmdSelectNone.Enabled = True Then
17060           .cmdSelectNone.SetFocus
17070         Else
17080           lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
17090           If lngRecsCur > 0& Then
17100             .frmMasterBalance_Sub.SetFocus
17110           Else
17120             .cmdClose.SetFocus
17130           End If
17140         End If
17150       End With
17160     End Select
17170   End If

EXITP:
17180   KeyCode = intRetVal
17190   Exit Sub

ERRH:
17200   intRetVal = 0
17210   Select Case ERR.Number
        Case Else
17220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17230   End Select
17240   Resume EXITP

End Sub

Private Sub chkDiscrepancies_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkDiscrepancies_MouseMove"

17310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17320   Exit Sub

ERRH:
17330   Select Case ERR.Number
        Case Else
17340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17350   End Select
17360   Resume EXITP

End Sub

Private Sub chkIncludeClosed_AfterUpdate()

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeClosed_AfterUpdate"

        Dim strFilter As String

17410   With Me
17420     strFilter = vbNullString
17430     Select Case .chkIncludeClosed
          Case True
17440       .chkIncludeClosed_lbl.FontBold = True
17450       If .chkDiscrepancies = True Then
17460         strFilter = "[Ix] = 'X' Or [Px] = 'X' Or [Cx] = 'X'"
17470       End If
17480       .IsClosed_lbl.Visible = True
17490       .IsClosed_lbl_box.Visible = True
17500     Case False
17510       .chkIncludeClosed_lbl.FontBold = False
17520       strFilter = "[IsClosed] = False"
17530       If .chkDiscrepancies = True Then
17540         strFilter = "([Ix] = 'X' Or [Px] = 'X' Or [Cx] = 'X') And " & strFilter
17550       End If
17560       .IsClosed_lbl.Visible = False
17570       .IsClosed_lbl_box.Visible = False
17580     End Select
17590     .frmMasterBalance_Sub.Form.Filter = strFilter
17600     If strFilter = vbNullString Then
17610       .frmMasterBalance_Sub.Form.FilterOn = False
17620     Else
17630       .frmMasterBalance_Sub.Form.FilterOn = True
17640     End If
17650     DoEvents
17660     .frmMasterBalance_Sub.Form.Requery
17670     DoEvents
17680     If blnIsOpen = False Then
17690       blnRecalcTots = True
17700       .TimerInterval = 500&
17710     End If
17720   End With

EXITP:
17730   Exit Sub

ERRH:
17740   Select Case ERR.Number
        Case Else
17750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17760   End Select
17770   Resume EXITP

End Sub

Private Sub chkIncludeClosed_KeyDown(KeyCode As Integer, Shift As Integer)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeClosed_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17820   intShiftDown = (Shift And acShiftMask) > 0
17830   intAltDown = (Shift And acAltMask) > 0
17840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17860     Select Case intRetVal
          Case vbKeyTab
17870       With Me
17880         intRetVal = 0
17890         If .chkShowAcctSort.Enabled = True And .chkShowAcctSort.Visible = True Then
17900           .chkShowAcctSort.SetFocus
17910         ElseIf .chkShowAcctMast.Enabled = True And .chkShowAcctMast.Visible = True Then
17920           .chkShowAcctMast.SetFocus
17930         Else
17940           .chkPageOf.SetFocus
17950         End If
17960       End With
17970     End Select
17980   End If

        ' ** Shift keys.
17990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18000     Select Case intRetVal
          Case vbKeyTab
18010       With Me
18020         intRetVal = 0
18030         If .chkDiscrepancies.Enabled = True Then
18040           .chkDiscrepancies.SetFocus
18050         ElseIf .cmdSelectNone.Enabled = True Then
18060           .cmdSelectNone.SetFocus
18070         Else
18080           lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
18090           If lngRecsCur > 0& Then
18100             .frmMasterBalance_Sub.SetFocus
18110           Else
18120             .cmdClose.SetFocus
18130           End If
18140         End If
18150       End With
18160     End Select
18170   End If

EXITP:
18180   KeyCode = intRetVal
18190   Exit Sub

ERRH:
18200   intRetVal = 0
18210   Select Case ERR.Number
        Case Else
18220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18230   End Select
18240   Resume EXITP

End Sub

Private Sub chkIncludeClosed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeClosed_MouseMove"

18310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
18320   Exit Sub

ERRH:
18330   Select Case ERR.Number
        Case Else
18340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18350   End Select
18360   Resume EXITP

End Sub

Private Sub chkShowAcctSort_AfterUpdate()

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAcctSort_AfterUpdate"

        Dim strFilter As String

18410   With Me
18420     strFilter = .frmMasterBalance_Sub.Form.Filter
18430     If .chkDiscrepancies = False And .chkIncludeClosed = True And strCallingForm = "frmMenu_Maintenance" Then
            ' ** I don't know why this filter keeps reappearing!
18440       strFilter = vbNullString
18450     End If
18460     .chkShowAcctSort2 = .chkShowAcctSort
18470     Select Case .chkShowAcctSort
          Case True
18480       .chkShowAcctSort_lbl.FontBold = True
18490       .chkShowAcctSort2_lbl.FontBold = True
18500       .frmMasterBalance_Sub.Form.SetAcctSortCols True, .chkDiscrepancies, strCallingForm  ' ** Form Procedure: frmMasterBalance_Sub.
18510       ShowAcctSort True  ' ** Procedure: Below
18520       ShowAcctSort_Win True  ' ** Procedure: Below
18530     Case False
18540       .chkShowAcctSort_lbl.FontBold = False
18550       .chkShowAcctSort2_lbl.FontBold = False
18560       .frmMasterBalance_Sub.Form.SetAcctSortCols False, .chkDiscrepancies, strCallingForm  ' ** Form Procedure: frmMasterBalance_Sub.
18570       ShowAcctSort False  ' ** Procedure: Below
18580       ShowAcctSort_Win False  ' ** Procedure: Below
18590     End Select
18600     DoEvents
18610     .frmMasterBalance_Sub.Form.Filter = strFilter
18620     If strFilter = vbNullString Then
18630       .frmMasterBalance_Sub.Form.FilterOn = False
18640     Else
18650       .frmMasterBalance_Sub.Form.FilterOn = True
18660     End If
18670     DoEvents
18680     .frmMasterBalance_Sub.SetFocus
18690   End With

EXITP:
18700   Exit Sub

ERRH:
18710   Select Case ERR.Number
        Case Else
18720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18730   End Select
18740   Resume EXITP

End Sub

Private Sub chkShowAcctSort_KeyDown(KeyCode As Integer, Shift As Integer)

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAcctSort_KeyDown"

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
          Case vbKeyTab
18870       With Me
18880         intRetVal = 0
18890         If .chkShowAcctMast.Enabled = True And .chkShowAcctMast.Visible = True Then
18900           .chkShowAcctMast.SetFocus
18910         Else
18920           .chkPageOf.SetFocus
18930         End If
18940       End With
18950     End Select
18960   End If

        ' ** Shift keys.
18970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18980     Select Case intRetVal
          Case vbKeyTab
18990       With Me
19000         intRetVal = 0
19010         If .chkIncludeClosed.Enabled = True Then
19020           .chkIncludeClosed.SetFocus
19030         ElseIf .chkDiscrepancies.Visible = True And .chkDiscrepancies.Enabled = True Then
19040           .chkDiscrepancies.SetFocus
19050         ElseIf .cmdSelectNone.Enabled = True Then
19060           .cmdSelectNone.SetFocus
19070         Else
19080           lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
19090           If lngRecsCur > 0& Then
19100             .frmMasterBalance_Sub.SetFocus
19110           Else
19120             .cmdClose.SetFocus
19130           End If
19140         End If
19150       End With
19160     End Select
19170   End If

EXITP:
19180   KeyCode = intRetVal
19190   Exit Sub

ERRH:
19200   intRetVal = 0
19210   Select Case ERR.Number
        Case Else
19220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19230   End Select
19240   Resume EXITP

End Sub

Private Sub chkShowAcctSort2_AfterUpdate()

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAcctSort2_AfterUpdate"

        Dim strFilter As String

19310   With Me
19320     strFilter = .frmMasterBalance_Sub.Form.Filter
19330     If .chkDiscrepancies = False And .chkIncludeClosed = True And strCallingForm = "frmMenu_Maintenance" Then
            ' ** I don't know why this filter keeps reappearing!
19340       strFilter = vbNullString
19350     End If
19360     .chkShowAcctSort = .chkShowAcctSort2
19370     Select Case .chkShowAcctSort2
          Case True
19380       .chkShowAcctSort2_lbl.FontBold = True
19390       .chkShowAcctSort_lbl.FontBold = True
19400       .frmMasterBalance_Sub.Form.SetAcctSortCols True, .chkDiscrepancies, strCallingForm  ' ** Form Procedure: frmMasterBalance_Sub.
19410       ShowAcctSort True  ' ** Procedure: Below
19420       ShowAcctSort_Win True  ' ** Procedure: Below
19430     Case False
19440       .chkShowAcctSort2_lbl.FontBold = False
19450       .chkShowAcctSort_lbl.FontBold = False
19460       .frmMasterBalance_Sub.Form.SetAcctSortCols False, .chkDiscrepancies, strCallingForm  ' ** Form Procedure: frmMasterBalance_Sub.
19470       ShowAcctSort False  ' ** Procedure: Below
19480       ShowAcctSort_Win False  ' ** Procedure: Below
19490     End Select
19500     DoEvents
19510     .frmMasterBalance_Sub.Form.Filter = strFilter
19520     If strFilter = vbNullString Then
19530       .frmMasterBalance_Sub.Form.FilterOn = False
19540     Else
19550       .frmMasterBalance_Sub.Form.FilterOn = True
19560     End If
19570     DoEvents
19580     .frmMasterBalance_Sub.SetFocus
19590   End With

EXITP:
19600   Exit Sub

ERRH:
19610   Select Case ERR.Number
        Case Else
19620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19630   End Select
19640   Resume EXITP

End Sub

Private Sub chkShowAcctMast_AfterUpdate()

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAcctMast_AfterUpdate"

        Dim strFilter As String

19710   With Me
19720     strFilter = .frmMasterBalance_Sub.Form.Filter
19730     .chkShowAcctMast2 = .chkShowAcctMast
19740     Select Case .chkShowAcctMast
          Case True
19750       .chkShowAcctMast_lbl.FontBold = True
19760       .chkShowAcctMast2_lbl.FontBold = True
19770       .frmMasterBalance_Sub.Form.SetAcctMastCols True  ' ** Form Procedure: frmMasterBalance_Sub.
19780       ShowAcctMast_Win True  ' ** Procedure: Below
19790     Case False
19800       .chkShowAcctMast_lbl.FontBold = False
19810       .chkShowAcctMast2_lbl.FontBold = False
19820       .frmMasterBalance_Sub.Form.SetAcctMastCols False  ' ** Form Procedure: frmMasterBalance_Sub.
19830       ShowAcctMast_Win False  ' ** Procedure: Below
19840     End Select
19850     DoEvents
19860     .frmMasterBalance_Sub.Form.Filter = strFilter
19870     If strFilter = vbNullString Then
19880       .frmMasterBalance_Sub.Form.FilterOn = False
19890     Else
19900       .frmMasterBalance_Sub.Form.FilterOn = True
19910     End If
19920     DoEvents
19930     .frmMasterBalance_Sub.SetFocus
19940   End With

EXITP:
19950   Exit Sub

ERRH:
19960   Select Case ERR.Number
        Case Else
19970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19980   End Select
19990   Resume EXITP

End Sub

Private Sub chkShowAcctMast_KeyDown(KeyCode As Integer, Shift As Integer)

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAcctMast_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20020   intShiftDown = (Shift And acShiftMask) > 0
20030   intAltDown = (Shift And acAltMask) > 0
20040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20060     Select Case intRetVal
          Case vbKeyTab
20070       With Me
20080         intRetVal = 0
20090         .chkPageOf.SetFocus
20100       End With
20110     End Select
20120   End If

        ' ** Shift keys.
20130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20140     Select Case intRetVal
          Case vbKeyTab
20150       With Me
20160         intRetVal = 0
20170         If .chkShowAcctSort.Visible = True And .chkShowAcctSort.Enabled = True Then
20180           .chkShowAcctSort.SetFocus
20190         ElseIf .chkIncludeClosed.Enabled = True Then
20200           .chkIncludeClosed.SetFocus
20210         ElseIf .chkDiscrepancies.Visible = True And .chkDiscrepancies.Enabled = True Then
20220           .chkDiscrepancies.SetFocus
20230         ElseIf .cmdSelectNone.Enabled = True Then
20240           .cmdSelectNone.SetFocus
20250         Else
20260           lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
20270           If lngRecsCur > 0& Then
20280             .frmMasterBalance_Sub.SetFocus
20290           Else
20300             .cmdClose.SetFocus
20310           End If
20320         End If
20330       End With
20340     End Select
20350   End If

EXITP:
20360   KeyCode = intRetVal
20370   Exit Sub

ERRH:
20380   intRetVal = 0
20390   Select Case ERR.Number
        Case Else
20400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20410   End Select
20420   Resume EXITP

End Sub

Private Sub chkShowAcctMast2_AfterUpdate()

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAcctMast2_AfterUpdate"

        Dim strFilter As String

20510   With Me
20520     strFilter = .frmMasterBalance_Sub.Form.Filter
20530     .chkShowAcctMast = .chkShowAcctMast2
20540     Select Case .chkShowAcctMast2
          Case True
20550       .chkShowAcctMast2_lbl.FontBold = True
20560       .chkShowAcctMast_lbl.FontBold = True
20570       .frmMasterBalance_Sub.Form.SetAcctMastCols True  ' ** Form Procedure: frmMasterBalance_Sub.
20580       ShowAcctMast_Win True  ' ** Procedure: Below
20590     Case False
20600       .chkShowAcctMast2_lbl.FontBold = False
20610       .chkShowAcctMast_lbl.FontBold = False
20620       .frmMasterBalance_Sub.Form.SetAcctMastCols False  ' ** Form Procedure: frmMasterBalance_Sub.
20630       ShowAcctMast_Win False  ' ** Procedure: Below
20640     End Select
20650     DoEvents
20660     .frmMasterBalance_Sub.Form.Filter = strFilter
20670     If strFilter = vbNullString Then
20680       .frmMasterBalance_Sub.Form.FilterOn = False
20690     Else
20700       .frmMasterBalance_Sub.Form.FilterOn = True
20710     End If
20720     DoEvents
20730     .frmMasterBalance_Sub.SetFocus
20740   End With

EXITP:
20750   Exit Sub

ERRH:
20760   Select Case ERR.Number
        Case Else
20770     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20780   End Select
20790   Resume EXITP

End Sub

Public Sub chkPageOf_AfterUpdate()

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

20810   With Me
20820     .chkPageOf2 = .chkPageOf
20830     Select Case .chkPageOf
          Case True
20840       .chkPageOf_lbl.FontBold = True
20850       .chkPageOf2_lbl.FontBold = True
20860     Case False
20870       .chkPageOf_lbl.FontBold = False
20880       .chkPageOf2_lbl.FontBold = False
20890     End Select
20900   End With

EXITP:
20910   Exit Sub

ERRH:
20920   Select Case ERR.Number
        Case Else
20930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20940   End Select
20950   Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21020   intShiftDown = (Shift And acShiftMask) > 0
21030   intAltDown = (Shift And acAltMask) > 0
21040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21060     Select Case intRetVal
          Case vbKeyTab
21070       With Me
21080         intRetVal = 0
21090         .chkSaveSizePos.SetFocus
21100       End With
21110     End Select
21120   End If

        ' ** Shift keys.
21130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21140     Select Case intRetVal
          Case vbKeyTab
21150       With Me
21160         intRetVal = 0
21170         If .chkShowAcctMast.Visible = True And .chkShowAcctMast.Enabled = True Then
21180           .chkShowAcctMast.SetFocus
21190         ElseIf .chkShowAcctSort.Visible = True And .chkShowAcctSort.Enabled = True Then
21200           .chkShowAcctSort.SetFocus
21210         ElseIf .chkIncludeClosed.Enabled = True Then
21220           .chkIncludeClosed.SetFocus
21230         ElseIf .chkDiscrepancies.Visible = True And .chkDiscrepancies.Enabled = True Then
21240           .chkDiscrepancies.SetFocus
21250         ElseIf .cmdSelectNone.Enabled = True Then
21260           .cmdSelectNone.SetFocus
21270         Else
21280           lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
21290           If lngRecsCur > 0& Then
21300             .frmMasterBalance_Sub.SetFocus
21310           Else
21320             .cmdClose.SetFocus
21330           End If
21340         End If
21350       End With
21360     End Select
21370   End If

EXITP:
21380   KeyCode = intRetVal
21390   Exit Sub

ERRH:
21400   intRetVal = 0
21410   Select Case ERR.Number
        Case Else
21420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21430   End Select
21440   Resume EXITP

End Sub

Private Sub chkPageOf_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_MouseMove"

21510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
21520   Exit Sub

ERRH:
21530   Select Case ERR.Number
        Case Else
21540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21550   End Select
21560   Resume EXITP

End Sub

Private Sub chkPageOf2_AfterUpdate()

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf2_AfterUpdate"

21610   With Me
21620     .chkPageOf = .chkPageOf2
21630     Select Case .chkPageOf2
          Case True
21640       .chkPageOf2_lbl.FontBold = True
21650       .chkPageOf_lbl.FontBold = True
21660     Case False
21670       .chkPageOf2_lbl.FontBold = False
21680       .chkPageOf_lbl.FontBold = False
21690     End Select
21700   End With

EXITP:
21710   Exit Sub

ERRH:
21720   Select Case ERR.Number
        Case Else
21730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21740   End Select
21750   Resume EXITP

End Sub

Private Sub chkSaveSizePos_AfterUpdate()

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkSaveSizePos_AfterUpdate"

21810   With Me
21820     .chkSaveSizePos2 = .chkSaveSizePos
21830     Select Case .chkSaveSizePos
          Case True
21840       .chkSaveSizePos_lbl.FontBold = True
21850       .chkSaveSizePos2_lbl.FontBold = True
21860     Case False
21870       .chkSaveSizePos_lbl.FontBold = False
21880       .chkSaveSizePos2_lbl.FontBold = False
21890     End Select
21900   End With

EXITP:
21910   Exit Sub

ERRH:
21920   Select Case ERR.Number
        Case Else
21930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21940   End Select
21950   Resume EXITP

End Sub

Private Sub chkSaveSizePos_KeyDown(KeyCode As Integer, Shift As Integer)

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkSaveSizePos_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22020   intShiftDown = (Shift And acShiftMask) > 0
22030   intAltDown = (Shift And acAltMask) > 0
22040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22060     Select Case intRetVal
          Case vbKeyTab
22070       With Me
22080         intRetVal = 0
22090         If .cmdPreview.Enabled = True Then
22100           .cmdPreview.SetFocus
22110         ElseIf .cmdUpdate.Enabled = True Then
22120           .cmdUpdate.SetFocus
22130         Else
22140           .cmdClose.SetFocus
22150         End If
22160       End With
22170     End Select
22180   End If

        ' ** Shift keys.
22190   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22200     Select Case intRetVal
          Case vbKeyTab
22210       With Me
22220         intRetVal = 0
22230         .chkPageOf.SetFocus
22240       End With
22250     End Select
22260   End If

EXITP:
22270   KeyCode = intRetVal
22280   Exit Sub

ERRH:
22290   intRetVal = 0
22300   Select Case ERR.Number
        Case Else
22310     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22320   End Select
22330   Resume EXITP

End Sub

Private Sub chkSaveSizePos2_AfterUpdate()

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkSaveSizePos2_AfterUpdate"

22410   With Me
22420     .chkSaveSizePos = .chkSaveSizePos2
22430     Select Case .chkSaveSizePos2
          Case True
22440       .chkSaveSizePos2_lbl.FontBold = True
22450       .chkSaveSizePos_lbl.FontBold = True
22460     Case False
22470       .chkSaveSizePos2_lbl.FontBold = False
22480       .chkSaveSizePos_lbl.FontBold = False
22490     End Select
22500   End With

EXITP:
22510   Exit Sub

ERRH:
22520   Select Case ERR.Number
        Case Else
22530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22540   End Select
22550   Resume EXITP

End Sub

Private Sub cmdUpdate_Click()

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim dblRecsSum_Disc As Double, lngRecsSumDisc_Save As Long, lngRecsSumFix_Save As Long
        Dim blnIsErr As Boolean

22610   With Me
22620     DoCmd.Hourglass True
22630     DoEvents
22640     If .RecsSum_Fix = 0 Or .RecsSum_Disc = 0 Then
22650       If .frmMasterBalance_Sub.Form.RecsSum_Fix = 0 Then
22660         .frmMasterBalance_Sub.Form.RecsSum_Fix.Requery
22670         DoEvents
22680       End If
22690       If .frmMasterBalance_Sub.Form.RecsSum_Fix <> 0 Then
22700         If .RecsSum_Fix = 0 Then
22710           .RecsSum_Fix = .frmMasterBalance_Sub.Form.RecsSum_Fix
22720           DoEvents
22730         End If
22740       End If
22750       If .frmMasterBalance_Sub.Form.RecsSum_Disc = 0 Then
22760         .frmMasterBalance_Sub.Form.RecsSum_Disc.Requery
22770         DoEvents
22780       End If
22790       If .frmMasterBalance_Sub.Form.RecsSum_Disc <> 0 Then
22800         If .RecsSum_Disc = 0 Then
22810           .RecsSum_Disc = .frmMasterBalance_Sub.Form.RecsSum_Disc
22820           DoEvents
22830         End If
22840       End If
22850     End If
22860     DoEvents
22870     If .RecsSum_Fix > 0 Then
22880       If .RecsSum_Disc > 0 Then
22890         lngRecsSumDisc_Save = .RecsSum_Disc
22900         lngRecsSumFix_Save = .RecsSum_Fix
22910         DoCmd.Hourglass False
22920         If MsgBox("Are you sure you want to update the account table Numbers?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, ("Update Account Table" & Space(40))) = vbYes Then
                ' ** Get the accountno of the 1st one checked.
22930           DoCmd.Hourglass True
22940           DoEvents
22950           strAccountNo = DLookup("[accountno]", "tblMasterBalance", "[fix] = True")
22960           Set dbs = CurrentDb
22970           With dbs
                  ' ** Update Account, from tblMasterBalance.
22980             Set qdf = .QueryDefs("qryMasterBalance_15")
22990             qdf.Execute
23000             Set qdf = Nothing
23010             .Close
23020           End With
23030           Set dbs = Nothing
23040           If .chkDiscrepancies = True Then .chkDiscrepancies = False
23050           chkDiscrepancies_AfterUpdate  ' ** Procedure: Below.
23060           Form_Load  ' ** Procedure: Above.
23070           DoEvents
23080           Set rst = .frmMasterBalance_Sub.Form.RecordsetClone
23090           rst.FindFirst "[accountno] = '" & strAccountNo & "'"
23100           If rst.NoMatch = False Then
23110             .frmMasterBalance_Sub.Form.MoveRec 0, strAccountNo  ' ** Form Function: frmMasterBalance_Sub.
23120           Else
23130             .frmMasterBalance_Sub.Form.Requery
23140           End If
23150           rst.Close
23160           DoEvents
23170           lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
23180           If lngRecsCur > 0& Then
23190 On Error Resume Next
23200             dblRecsSum_Disc = .frmMasterBalance_Sub.Form.RecsSum_Disc
23210             If ERR.Number <> 0 Then
23220 On Error GoTo ERRH
23230               blnIsErr = True
23240               Do While blnIsErr = True
23250                 ForcePause 0.25  ' ** Module Function: modCodeUtilities.
23260 On Error Resume Next
23270                 dblRecsSum_Disc = .frmMasterBalance_Sub.Form.RecsSum_Disc
23280                 If ERR.Number <> 0 Then
23290 On Error GoTo ERRH
23300                 Else
23310 On Error GoTo ERRH
23320                   blnIsErr = False
23330                   Exit Do
23340                 End If
23350               Loop
23360             Else
23370 On Error GoTo ERRH
23380             End If
23390             .RecsSum_Fix = .frmMasterBalance_Sub.Form.RecsSum_Fix
23400             .RecsSum_Disc = .frmMasterBalance_Sub.Form.RecsSum_Disc
23410             DoEvents
                  ' ** It always seems to think there are no more discrepancies,
                  ' ** even when only a few were fixed!
23420             If lngRecsSumFix_Save < lngRecsSumDisc_Save Then
23430               If .RecsSum_Disc = 0 Then
23440                 If .frmMasterBalance_Sub.Form.RecsSum_Disc = 0 Then
23450                   .frmMasterBalance_Sub.Form.RecsSum_Disc.Requery
23460                   DoEvents
23470                 End If
23480                 If .frmMasterBalance_Sub.Form.RecsSum_Disc = 0 Then
                        ' ** I'll try it some other way.
23490                 Else
23500                   .RecsSum_Disc = .frmMasterBalance_Sub.Form.RecsSum_Disc
23510                 End If
23520               End If
23530             End If
23540           Else
23550             .RecsSum_Fix = CDbl(0)
23560             .RecsSum_Disc = CDbl(0)
23570             dblRecsSum_Disc = 0#
23580           End If
23590           If ((dblRecsSum_Disc > 0#) Or (lngRecsSumFix_Save < lngRecsSumDisc_Save)) Then
23600             .cmdUpdate.Enabled = True
23610             .chkDiscrepancies.Enabled = True
23620             .chkDiscrepancies_lbl.Visible = True
23630             .chkDiscrepancies_lbl2.Visible = False
23640             .chkDiscrepancies_lbl2_dim_hi.Visible = False
23650             .cmdUpdate_lbl.Caption = "To update an item with a red 'X', check" & vbCrLf & "the Fix box, then click 'Update'."
23660             .cmdUpdate_lbl2.Visible = True
23670           Else
23680             .cmdUpdate.Enabled = False
23690             .chkDiscrepancies.Enabled = False
23700             .chkDiscrepancies_lbl.Visible = False
23710             .chkDiscrepancies_lbl2.Visible = True
23720             .chkDiscrepancies_lbl2_dim_hi.Visible = True
23730             .cmdUpdate_lbl.Caption = "No discrepancies found."
23740             .cmdUpdate_lbl2.Visible = False
23750           End If
23760           .frmMasterBalance_Sub.SetFocus
23770           .frmMasterBalance_Sub.Form.accountno.SetFocus
23780           DoEvents
23790           Beep
23800           DoCmd.Hourglass False
23810           MsgBox "Master totals updated.", vbInformation + vbOKOnly, ("Finished" & Space(40))
23820           If lngRecsSumFix_Save < lngRecsSumDisc_Save Then
23830             ResetTots True  ' ** Procedure: Below.
23840           Else
23850             blnRecalcTots = True
23860             .TimerInterval = 500&
23870           End If
23880         End If
23890       Else
23900         DoCmd.Hourglass False
23910         Beep
23920         MsgBox "Amounts are already in balance!", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
23930       End If
23940     Else
23950       DoCmd.Hourglass False
23960       Beep
23970       MsgBox "Nothing checked!", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
23980     End If
23990   End With

EXITP:
24000   Set rst = Nothing
24010   Set qdf = Nothing
24020   Set dbs = Nothing
24030   Exit Sub

ERRH:
24040   DoCmd.Hourglass False
24050   Select Case ERR.Number
        Case Else
24060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24070   End Select
24080   Resume EXITP

End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24120   intShiftDown = (Shift And acShiftMask) > 0
24130   intAltDown = (Shift And acAltMask) > 0
24140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24160     Select Case intRetVal
          Case vbKeyTab
24170       With Me
24180         intRetVal = 0
24190         .cmdClose.SetFocus
24200       End With
24210     End Select
24220   End If

        ' ** Shift keys.
24230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24240     Select Case intRetVal
          Case vbKeyTab
24250       With Me
24260         intRetVal = 0
24270         If .cmdExcel.Enabled = True Then
24280           .cmdExcel.SetFocus
24290         ElseIf .cmdWord.Enabled = True Then
24300           .cmdWord.SetFocus
24310         Else
24320           .chkSaveSizePos.SetFocus
24330         End If
24340       End With
24350     End Select
24360   End If

EXITP:
24370   KeyCode = intRetVal
24380   Exit Sub

ERRH:
24390   intRetVal = 0
24400   Select Case ERR.Number
        Case Else
24410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24420   End Select
24430   Resume EXITP

End Sub

Private Sub DisableSelectBtns(Optional varAble As Variant)

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "DisableSelectBtns"

        Dim blnAble As Boolean

24510   With Me
24520     .FocusHolder.SetFocus
24530     DoEvents
24540     Select Case IsMissing(varAble)
          Case True
24550       blnAble = False
24560     Case False
24570       blnAble = varAble
24580     End Select
24590     Select Case blnAble
          Case True
24600       .cmdSelectAll.Enabled = True
24610       .cmdSelectAll_raised_img.Visible = True
24620       .cmdSelectAll_raised_semifocus_dots_img.Visible = False
24630       .cmdSelectAll_raised_focus_img.Visible = False
24640       .cmdSelectAll_raised_focus_dots_img.Visible = False
24650       .cmdSelectAll_sunken_focus_dots_img.Visible = False
24660       .cmdSelectAll_raised_img_dis.Visible = False
24670       .cmdSelectNone.Enabled = True
24680       .cmdSelectNone_raised_img.Visible = True
24690       .cmdSelectNone_raised_semifocus_dots_img.Visible = False
24700       .cmdSelectNone_raised_focus_img.Visible = False
24710       .cmdSelectNone_raised_focus_dots_img.Visible = False
24720       .cmdSelectNone_sunken_focus_dots_img.Visible = False
24730       .cmdSelectNone_raised_img_dis.Visible = False
24740     Case False
24750       .cmdSelectAll.Enabled = False
24760       .cmdSelectAll_raised_img_dis.Visible = True
24770       .cmdSelectAll_raised_img.Visible = False
24780       .cmdSelectAll_raised_semifocus_dots_img.Visible = False
24790       .cmdSelectAll_raised_focus_img.Visible = False
24800       .cmdSelectAll_raised_focus_dots_img.Visible = False
24810       .cmdSelectAll_sunken_focus_dots_img.Visible = False
24820       .cmdSelectNone.Enabled = False
24830       .cmdSelectNone_raised_img_dis.Visible = True
24840       .cmdSelectNone_raised_img.Visible = False
24850       .cmdSelectNone_raised_semifocus_dots_img.Visible = False
24860       .cmdSelectNone_raised_focus_img.Visible = False
24870       .cmdSelectNone_raised_focus_dots_img.Visible = False
24880       .cmdSelectNone_sunken_focus_dots_img.Visible = False
24890     End Select
24900   End With

EXITP:
24910   Exit Sub

ERRH:
24920   Select Case ERR.Number
        Case Else
24930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24940   End Select
24950   Resume EXITP

End Sub

Public Function FrmIsMax() As Boolean

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "FrmIsMax"

        Dim blnRetVal As Boolean

25010   blnRetVal = blnIsMaximized

EXITP:
25020   FrmIsMax = blnRetVal
25030   Exit Function

ERRH:
25040   Select Case ERR.Number
        Case Else
25050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25060   End Select
25070   Resume EXITP

End Function

Public Sub ResetTots(blnReset As Boolean)

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "ResetTots"

25110   With Me
25120     blnRecalcTots = blnReset
25130     If blnRecalcTots = True Then
25140       .TimerInterval = 1000&
25150     End If
25160   End With

EXITP:
25170   Exit Sub

ERRH:
25180   Select Case ERR.Number
        Case Else
25190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25200   End Select
25210   Resume EXITP

End Sub

Public Sub RecalcTots()
' ** It's possible that discrepancies could be present in the closed
' ** accounts, so presense or absense of discrepancies may change.

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

        Dim dblRecsSum_Disc As Double
        Dim strFilter As String

25310   With Me
25320     DoCmd.Hourglass True
25330     DoEvents
25340     .FocusHolder.SetFocus
25350     DoEvents
25360     lngRecsCur = .frmMasterBalance_Sub.Form.RecCnt  ' ** Form Function: frmMasterBalance_Sub.
25370     If lngRecsCur > 0& Then
25380       .RecsSum_Fix = .frmMasterBalance_Sub.Form.RecsSum_Fix
25390       DoEvents
25400       .RecsSum_Disc = .frmMasterBalance_Sub.Form.RecsSum_Disc
25410       DoEvents
25420     Else
25430       .RecsSum_Fix = CDbl(0)
25440       .RecsSum_Disc = CDbl(0)
25450     End If
25460     dblRecsSum_Disc = .RecsSum_Disc
25470     If dblRecsSum_Disc = 0# Then
25480       .cmdUpdate.Enabled = False
25490       .chkDiscrepancies = False
25500       .chkDiscrepancies_lbl.FontBold = False
25510       .chkDiscrepancies.Enabled = False
25520       .chkDiscrepancies_lbl.Visible = False
25530       .chkDiscrepancies_lbl2.Visible = True
25540       .chkDiscrepancies_lbl2_dim_hi.Visible = True
25550       .cmdUpdate_lbl.Caption = "No discrepancies found."
25560       .cmdUpdate_lbl2.Visible = False
25570     Else
25580       .cmdUpdate.Enabled = True
25590       .chkDiscrepancies.Enabled = True
25600       .chkDiscrepancies_lbl.Visible = True
25610       .chkDiscrepancies_lbl2.Visible = False
25620       .chkDiscrepancies_lbl2_dim_hi.Visible = False
25630       .cmdUpdate_lbl.Caption = "To update an item with a red 'X', check" & vbCrLf & "the Fix box, then click 'Update'."
25640       .cmdUpdate_lbl2.Visible = True
25650     End If
          ' ** I hope this filter stuff doesn't change the totals, above!
25660     strFilter = vbNullString
25670     Select Case .chkDiscrepancies
          Case True
25680       If dblRecsSum_Disc = 0# Then
25690         Select Case .chkIncludeClosed
              Case True
25700           .frmMasterBalance_Sub.Form.Filter = strFilter
25710           .frmMasterBalance_Sub.Form.FilterOn = False
25720           .frmMasterBalance_Sub.Form.Requery
25730         Case False
25740           strFilter = "[IsClosed] = False"
25750           .frmMasterBalance_Sub.Form.Filter = strFilter
25760           .frmMasterBalance_Sub.Form.FilterOn = True
25770           .frmMasterBalance_Sub.Form.Requery
25780         End Select
25790       Else
25800         strFilter = "([Ix] = 'X' Or [Px] = 'X' Or [Cx] = 'X')"
25810         If .chkIncludeClosed = False Then
25820           strFilter = strFilter & " And [IsClosed] = False"
25830         End If
25840         .frmMasterBalance_Sub.Form.Filter = strFilter
25850         .frmMasterBalance_Sub.Form.FilterOn = True
25860         .frmMasterBalance_Sub.Form.Requery
25870       End If
25880     Case False
25890       Select Case .chkIncludeClosed
            Case True
25900         .frmMasterBalance_Sub.Form.Filter = strFilter
25910         .frmMasterBalance_Sub.Form.FilterOn = False
25920         .frmMasterBalance_Sub.Form.Requery
25930       Case False
25940         strFilter = "[IsClosed] = False"
25950         .frmMasterBalance_Sub.Form.Filter = strFilter
25960         .frmMasterBalance_Sub.Form.FilterOn = True
25970         .frmMasterBalance_Sub.Form.Requery
25980       End Select
25990     End Select
26000     If dblRecsSum_Disc = 0# Then
26010       If .cmdSelectAll.Enabled = True Then
26020         DisableSelectBtns False  ' ** Procedure: Below.
26030       End If
26040     Else
26050       If .cmdSelectAll.Enabled = False Then
26060         DisableSelectBtns True  ' ** Procedure: Below.
26070       End If
26080     End If
26090     DoEvents
26100     DoCmd.Hourglass False
26110   End With

EXITP:
26120   Exit Sub

ERRH:
26130   DoCmd.Hourglass False
26140   Select Case ERR.Number
        Case Else
26150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26160   End Select
26170   Resume EXITP

End Sub

Private Sub ShowAcctSort(blnShow As Boolean)

26200 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAcctSort"

        Dim lngAcctSort_Width As Long
        Dim lngTmp01 As Long

26210   With Me

26220     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
26230       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
26240     End If

26250     If lngAcctSort_Width = 0& Then
26260       lngAcctSort_Width = (.frmMasterBalance_Sub.Form.account_sort.Width + (4& * lngTpp))
26270     End If

26280     Select Case blnShow
          Case True
26290       If .chkDiscrepancies_box.Left <> lngDiscrepanciesBox_Left Then
26300         .chkDiscrepancies_box.Left = lngDiscrepanciesBox_Left
26310         .chkShowAcctSort_lbl.Left = (lngAcctSort_Left + lngChkBoxLbl_Offset)
26320         .chkShowAcctSort.Left = lngAcctSort_Left
26330         .chkPageOf_lbl.Left = (lngPageOf_Left + lngChkBoxLbl_Offset)
26340         .chkPageOf.Left = lngPageOf_Left
26350         .chkSaveSizePos_lbl.Left = (lngPageOf_Left + lngChkBoxLbl_Offset)
26360         .chkSaveSizePos.Left = lngPageOf_Left
26370         .chkDiscrepancies_vline02.Left = (.chkShowAcctSort.Left - (4& * lngTpp))
26380         .chkDiscrepancies_vline01.Left = (.chkDiscrepancies_vline02.Left - lngTpp)
26390         .chkDiscrepancies_vline04.Left = (.chkPageOf.Left - (4& * lngTpp))
26400         .chkDiscrepancies_vline03.Left = (.chkDiscrepancies_vline04.Left - lngTpp)
26410         .chkDiscrepancies_lbl.Left = (lngDiscrepancies_Left + lngChkBoxLbl_Offset)
26420         .chkDiscrepancies_lbl2.Left = .chkDiscrepancies_lbl.Left
26430         .chkDiscrepancies_lbl2_dim_hi.Left = (.chkDiscrepancies_lbl2.Left + lngTpp)
26440         .chkDiscrepancies.Left = lngDiscrepancies_Left
26450         .chkShowAcctMast_lbl.Left = (lngAcctSort_Left + lngChkBoxLbl_Offset)
26460         .chkShowAcctMast.Left = lngAcctSort_Left
26470         .chkIncludeClosed.Left = .chkDiscrepancies.Left
26480         .chkIncludeClosed_lbl.Left = .chkDiscrepancies_lbl.Left
26490         .chkIncludeClosed_lbl2.Left = .chkDiscrepancies_lbl2.Left
26500         .chkIncludeClosed_lbl2_dim_hi.Left = .chkDiscrepancies_lbl2_dim_hi.Left
              ' ** Box is narrow only when both are off, otherwise box is wide.
26510         .RecsSum_box.Left = lngRecsSumBox_Left
26520         .RecsSum_box.Width = lngRecsSumBox_Width
26530         .cmdSelectAll.Left = lngSelectAll_Left
26540         .cmdSelectAll_raised_img.Left = lngSelectAll_Left
26550         .cmdSelectAll_raised_semifocus_dots_img.Left = lngSelectAll_Left
26560         .cmdSelectAll_raised_focus_img.Left = lngSelectAll_Left
26570         .cmdSelectAll_raised_focus_dots_img.Left = lngSelectAll_Left
26580         .cmdSelectAll_sunken_focus_dots_img.Left = lngSelectAll_Left
26590         .cmdSelectAll_raised_img_dis.Left = lngSelectAll_Left
26600         .cmdSelectNone.Left = lngSelectNone_Left
26610         .cmdSelectNone_raised_img.Left = lngSelectNone_Left
26620         .cmdSelectNone_raised_semifocus_dots_img.Left = lngSelectNone_Left
26630         .cmdSelectNone_raised_focus_img.Left = lngSelectNone_Left
26640         .cmdSelectNone_raised_focus_dots_img.Left = lngSelectNone_Left
26650         .cmdSelectNone_sunken_focus_dots_img.Left = lngSelectNone_Left
26660         .cmdSelectNone_raised_img_dis.Left = lngSelectNone_Left
26670         .RecsSum_Disc.Left = lngRecsSumDisc_Left
26680         .RecsSum_Disc_lbl.Left = lngRecsSumBox_Left
26690         .RecsSum_Disc_lbl2.Left = .RecsSum_Disc_lbl.Left
26700         .RecsSum_Disc_lbl.Visible = True
26710         .RecsSum_Disc_lbl2.Visible = False
26720         .RecsSum_Fix.Left = lngRecsSumFix_Left
26730         .RecsSum_Fix_lbl.Left = lngRecsSumBox_Left
26740       End If
26750     Case False
26760       If .chkDiscrepancies_box.Left <> (lngDiscrepanciesBox_Left - lngAcctSort_Width) Then
26770         .RecsSum_box.Left = (lngRecsSumBox_Left - lngAcctSort_Width)
26780         .RecsSum_Disc_lbl.Left = .RecsSum_box.Left
26790         .RecsSum_Disc_lbl2.Left = .RecsSum_Disc_lbl.Left
26800         .RecsSum_Disc.Left = (lngRecsSumDisc_Left - lngAcctSort_Width)
26810         .RecsSum_Fix_lbl.Left = .RecsSum_box.Left
26820         .RecsSum_Fix.Left = (lngRecsSumFix_Left - lngAcctSort_Width)
26830         .cmdSelectAll.Left = (lngSelectAll_Left - lngAcctSort_Width)
26840         .cmdSelectAll_raised_img.Left = .cmdSelectAll.Left
26850         .cmdSelectAll_raised_semifocus_dots_img.Left = .cmdSelectAll.Left
26860         .cmdSelectAll_raised_focus_img.Left = .cmdSelectAll.Left
26870         .cmdSelectAll_raised_focus_dots_img.Left = .cmdSelectAll.Left
26880         .cmdSelectAll_sunken_focus_dots_img.Left = .cmdSelectAll.Left
26890         .cmdSelectAll_raised_img_dis.Left = .cmdSelectAll.Left
26900         .cmdSelectNone.Left = (lngSelectNone_Left - lngAcctSort_Width)
26910         .cmdSelectNone_raised_img.Left = .cmdSelectNone.Left
26920         .cmdSelectNone_raised_semifocus_dots_img.Left = .cmdSelectNone.Left
26930         .cmdSelectNone_raised_focus_img.Left = .cmdSelectNone.Left
26940         .cmdSelectNone_raised_focus_dots_img.Left = .cmdSelectNone.Left
26950         .cmdSelectNone_sunken_focus_dots_img.Left = .cmdSelectNone.Left
26960         .cmdSelectNone_raised_img_dis.Left = .cmdSelectNone.Left
26970         .chkDiscrepancies_box.Left = (lngDiscrepanciesBox_Left - lngAcctSort_Width)
26980         .chkDiscrepancies.Left = (lngDiscrepancies_Left - lngAcctSort_Width)
26990         .chkDiscrepancies_lbl.Left = (.chkDiscrepancies.Left + lngChkBoxLbl_Offset)
27000         .chkDiscrepancies_lbl2.Left = .chkDiscrepancies_lbl.Left
27010         .chkDiscrepancies_lbl2_dim_hi.Left = (.chkDiscrepancies_lbl2.Left + lngTpp)
27020         .chkShowAcctSort.Left = (lngAcctSort_Left - lngAcctSort_Width)
27030         .chkShowAcctSort_lbl.Left = (.chkShowAcctSort.Left + lngChkBoxLbl_Offset)
27040         .chkShowAcctMast.Left = .chkShowAcctSort.Left
27050         .chkShowAcctMast_lbl.Left = (.chkShowAcctMast.Left + lngChkBoxLbl_Offset)
27060         .chkIncludeClosed.Left = .chkDiscrepancies.Left
27070         .chkIncludeClosed_lbl.Left = .chkDiscrepancies_lbl.Left
27080         .chkIncludeClosed_lbl2.Left = .chkDiscrepancies_lbl2.Left
27090         .chkIncludeClosed_lbl2_dim_hi.Left = .chkDiscrepancies_lbl2_dim_hi.Left
27100         .chkPageOf.Left = (lngPageOf_Left - lngAcctSort_Width)
27110         .chkPageOf_lbl.Left = (.chkPageOf.Left + lngChkBoxLbl_Offset)
27120         .chkSaveSizePos.Left = .chkPageOf.Left
27130         .chkSaveSizePos_lbl.Left = (.chkSaveSizePos.Left + lngChkBoxLbl_Offset)
27140         .chkDiscrepancies_vline02.Left = (.chkShowAcctSort.Left - (4& * lngTpp))
27150         .chkDiscrepancies_vline01.Left = (.chkDiscrepancies_vline02.Left - lngTpp)
27160         .chkDiscrepancies_vline04.Left = (.chkPageOf.Left - (4& * lngTpp))
27170         .chkDiscrepancies_vline03.Left = (.chkDiscrepancies_vline04.Left - lngTpp)
27180         Select Case .chkShowAcctMast
              Case True
27190           .RecsSum_Disc_lbl.Visible = True
27200           .RecsSum_Disc_lbl2.Visible = False
                ' ** Box is narrow only when both are off, otherwise box is wide.
27210           .RecsSum_box.Width = lngRecsSumBox_Width
27220         Case False
27230           .RecsSum_Disc_lbl.Visible = False
27240           .RecsSum_Disc_lbl2.Visible = True
                ' ** If they're overlapping, this will calculate how much.
27250           lngTmp01 = ((.chkShowAcctSort2_box.Left + .chkShowAcctSort2_box.Width) - .RecsSum_box.Left)
27260           lngTmp01 = (lngTmp01 + (2& * lngTpp))
                ' ** Box is narrow only when both are off, otherwise box is wide.
27270           .RecsSum_box.Width = (lngRecsSumBox_Width - lngTmp01)
27280           .RecsSum_box.Left = (.RecsSum_box.Left + lngTmp01)
27290         End Select
27300       End If

27310     End Select

27320   End With

EXITP:
27330   Exit Sub

ERRH:
27340   Select Case ERR.Number
        Case Else
27350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27360   End Select
27370   Resume EXITP

End Sub

Private Sub ShowAcctSort_Win(blnShow As Boolean)

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAcctSort_Win"

        Dim lngTop As Long, lngLeft As Long, lngWidth As Long, lngHeight As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long

27410   With Me

27420     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
27430       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
27440     End If

27450     If lngAcctSort_Width = 0& Or lngAcctMast_Width = 0& Then
27460       lngAcctSort_Width = (.frmMasterBalance_Sub.Form.account_sort.Width + (4& * lngTpp))
27470       lngTmp01 = ((.frmMasterBalance_Sub.Form.ICash.Width + (4& * lngTpp)) + (.frmMasterBalance_Sub.Form.Ix.Width + (4& * lngTpp)))
27480       lngTmp02 = ((.frmMasterBalance_Sub.Form.PCash.Width + (4& * lngTpp)) + (.frmMasterBalance_Sub.Form.Px.Width + (4& * lngTpp)))
27490       lngTmp03 = ((.frmMasterBalance_Sub.Form.Cost.Width + (4& * lngTpp)) + (.frmMasterBalance_Sub.Form.Cx.Width + (4& * lngTpp)))
27500       lngAcctMast_Width = (lngTmp01 + lngTmp02 + lngTmp03)
27510       lngTmp01 = 0&: lngTmp02 = 0&: lngTmp03 = 0&
27520     End If

27530     lngTmp01 = 0&: lngTmp02 = 0&
27540     Select Case blnShow
          Case True
27550       Select Case .chkShowAcctMast
            Case True
              ' ** As Opening, with both showing.
27560         lngTmp01 = lngFrm_Width
27570         lngTmp02 = lngFrm_Left
27580         .cmdClose.Left = lngClose_Left
27590         .cmdUpdate.Left = lngUpdate_Left
27600         .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
27610         .Sizable_lbl2.Left = .Sizable_lbl1.Left
27620         .frmMasterBalance_Sub.Width = lngSub_Width
27630         .frmMasterBalance_Sub_box.Width = lngSubBox_Width
27640         .Nav_box01.Width = .frmMasterBalance_Sub.Width
27650         .Nav_hline03.Width = .frmMasterBalance_Sub.Width
27660         .Header_vline01.Left = lngFrm_Width
27670         .Header_vline02.Left = lngFrm_Width
27680         .Detail_vline01.Left = lngFrm_Width
27690         .Detail_vline02.Left = lngFrm_Width
27700         .Footer_vline01.Left = lngFrm_Width
27710         .Footer_vline02.Left = lngFrm_Width
27720         .Header_hline01.Width = lngFrm_Width
27730         .Header_hline02.Width = lngFrm_Width
27740         .Detail_hline01.Width = lngFrm_Width
27750         .Detail_hline02.Width = lngFrm_Width
27760         .Footer_hline01.Width = lngFrm_Width
27770         .Footer_hline02.Width = lngFrm_Width
27780         .Width = lngFrm_Width
27790         .frmMasterBalance_Sub.Form.Width = .frmMasterBalance_Sub.Form.IsClosed_bg1.Width
27800         lngTmp03 = (.cmdUpdate_lbl2.Left - .cmdUpdate_lbl.Left)
27810         .cmdUpdate_lbl.Left = lngUpdateLbl_left
27820         .cmdUpdate_lbl2.Left = (.cmdUpdate_lbl.Left + lngTmp03)
27830       Case False
27840         lngTmp01 = (lngFrm_Width - lngAcctMast_Width)
27850         lngTmp02 = (lngFrm_Left + (lngAcctMast_Width / 2&))
27860         .cmdClose.Left = (lngClose_Left - lngAcctMast_Width)
27870         .cmdUpdate.Left = (lngUpdate_Left - lngAcctMast_Width)
27880         .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
27890         .Sizable_lbl2.Left = .Sizable_lbl1.Left
27900         .frmMasterBalance_Sub.Width = (lngSub_Width - lngAcctMast_Width)
27910         .frmMasterBalance_Sub_box.Width = (lngSubBox_Width - lngAcctMast_Width)
27920         .Nav_box01.Width = .frmMasterBalance_Sub.Width
27930         .Nav_hline03.Width = .frmMasterBalance_Sub.Width
27940         .Header_vline01.Left = (lngFrm_Width - lngAcctMast_Width)
27950         .Header_vline02.Left = (lngFrm_Width - lngAcctMast_Width)
27960         .Detail_vline01.Left = (lngFrm_Width - lngAcctMast_Width)
27970         .Detail_vline02.Left = (lngFrm_Width - lngAcctMast_Width)
27980         .Footer_vline01.Left = (lngFrm_Width - lngAcctMast_Width)
27990         .Footer_vline02.Left = (lngFrm_Width - lngAcctMast_Width)
28000         .Header_hline01.Width = (lngFrm_Width - lngAcctMast_Width)
28010         .Header_hline02.Width = (lngFrm_Width - lngAcctMast_Width)
28020         .Detail_hline01.Width = (lngFrm_Width - lngAcctMast_Width)
28030         .Detail_hline02.Width = (lngFrm_Width - lngAcctMast_Width)
28040         .Footer_hline01.Width = (lngFrm_Width - lngAcctMast_Width)
28050         .Footer_hline02.Width = (lngFrm_Width - lngAcctMast_Width)
28060         .Width = (lngFrm_Width - lngAcctMast_Width)
28070         .frmMasterBalance_Sub.Form.Width = .frmMasterBalance_Sub.Form.IsClosed_bg1.Width
28080         lngTmp03 = (.cmdUpdate_lbl2.Left - .cmdUpdate_lbl.Left)
28090         .cmdUpdate_lbl.Left = .cmdUpdate_alt_left3.Left
28100         .cmdUpdate_lbl2.Left = (.cmdUpdate_lbl.Left + lngTmp03)
28110       End Select
28120     Case False
28130       Select Case .chkShowAcctMast
            Case True
28140         lngTmp01 = (lngFrm_Width - lngAcctSort_Width)
28150         lngTmp02 = (lngFrm_Left + (lngAcctSort_Width / 2&))
28160         .cmdClose.Left = (lngClose_Left - lngAcctSort_Width)
28170         .cmdUpdate.Left = (lngUpdate_Left - lngAcctSort_Width)
28180         .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
28190         .Sizable_lbl2.Left = .Sizable_lbl1.Left
28200         .frmMasterBalance_Sub.Width = (lngSub_Width - lngAcctSort_Width)
28210         .frmMasterBalance_Sub_box.Width = (lngSubBox_Width - lngAcctSort_Width)
28220         .Nav_box01.Width = .frmMasterBalance_Sub.Width
28230         .Nav_hline03.Width = .frmMasterBalance_Sub.Width
28240         .Header_vline01.Left = (lngFrm_Width - lngAcctSort_Width)
28250         .Header_vline02.Left = (lngFrm_Width - lngAcctSort_Width)
28260         .Detail_vline01.Left = (lngFrm_Width - lngAcctSort_Width)
28270         .Detail_vline02.Left = (lngFrm_Width - lngAcctSort_Width)
28280         .Footer_vline01.Left = (lngFrm_Width - lngAcctSort_Width)
28290         .Footer_vline02.Left = (lngFrm_Width - lngAcctSort_Width)
28300         .Header_hline01.Width = (lngFrm_Width - lngAcctSort_Width)
28310         .Header_hline02.Width = (lngFrm_Width - lngAcctSort_Width)
28320         .Detail_hline01.Width = (lngFrm_Width - lngAcctSort_Width)
28330         .Detail_hline02.Width = (lngFrm_Width - lngAcctSort_Width)
28340         .Footer_hline01.Width = (lngFrm_Width - lngAcctSort_Width)
28350         .Footer_hline02.Width = (lngFrm_Width - lngAcctSort_Width)
28360         .Width = (lngFrm_Width - lngAcctSort_Width)
28370         .frmMasterBalance_Sub.Form.Width = .frmMasterBalance_Sub.Form.IsClosed_bg1.Width
28380         lngTmp03 = (.cmdUpdate_lbl2.Left - .cmdUpdate_lbl.Left)
28390         .cmdUpdate_lbl.Left = .cmdUpdate_alt_left2.Left
28400         .cmdUpdate_lbl2.Left = (.cmdUpdate_lbl.Left + lngTmp03)
28410       Case False
28420         lngTmp01 = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28430         lngTmp02 = (lngFrm_Left + ((lngAcctSort_Width + lngAcctMast_Width) / 2&))
28440         .cmdClose.Left = (lngClose_Left - (lngAcctSort_Width + lngAcctMast_Width))
28450         .cmdUpdate.Left = (lngUpdate_Left - (lngAcctSort_Width + lngAcctMast_Width))
28460         .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
28470         .Sizable_lbl2.Left = .Sizable_lbl1.Left
28480         .frmMasterBalance_Sub.Width = (lngSub_Width - (lngAcctSort_Width + lngAcctMast_Width))
28490         .frmMasterBalance_Sub_box.Width = (lngSubBox_Width - (lngAcctSort_Width + lngAcctMast_Width))
28500         .Nav_box01.Width = .frmMasterBalance_Sub.Width
28510         .Nav_hline03.Width = .frmMasterBalance_Sub.Width
28520         .Header_vline01.Left = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28530         .Header_vline02.Left = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28540         .Detail_vline01.Left = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28550         .Detail_vline02.Left = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28560         .Footer_vline01.Left = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28570         .Footer_vline02.Left = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28580         .Header_hline01.Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28590         .Header_hline02.Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28600         .Detail_hline01.Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28610         .Detail_hline02.Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28620         .Footer_hline01.Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28630         .Footer_hline02.Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28640         .Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
28650         .frmMasterBalance_Sub.Form.Width = .frmMasterBalance_Sub.Form.IsClosed_bg1.Width
28660         lngTmp03 = (.cmdUpdate_lbl2.Left - .cmdUpdate_lbl.Left)
28670         .cmdUpdate_lbl.Left = .cmdUpdate_alt_left4.Left
28680         .cmdUpdate_lbl2.Left = (.cmdUpdate_lbl.Left + lngTmp03)
28690       End Select
28700     End Select

          ' ** Variables are fed empty, then populated ByRef.
28710     GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Function: modWindowFunctions.

28720     DoCmd.MoveSize (lngTmp02 - 120&), (lngTop - 120&), lngTmp01, lngHeight  ' ** Values are always 8 pixels off!
28730     DoEvents

28740   End With

EXITP:
28750   Exit Sub

ERRH:
28760   Select Case ERR.Number
        Case Else
28770     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28780   End Select
28790   Resume EXITP

End Sub

Private Sub ShowAcctMast_Win(blnShow As Boolean)

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAcctMast_Win"

        Dim lngTop As Long, lngLeft As Long, lngWidth As Long, lngHeight As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long

28810   With Me

28820     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
28830       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
28840     End If

28850     If lngAcctSort_Width = 0& Or lngAcctMast_Width = 0& Then
28860       lngAcctSort_Width = (.frmMasterBalance_Sub.Form.account_sort.Width + (4& * lngTpp))
28870       lngTmp01 = ((.frmMasterBalance_Sub.Form.ICash.Width + (4& * lngTpp)) + (.frmMasterBalance_Sub.Form.Ix.Width + (4& * lngTpp)))
28880       lngTmp02 = ((.frmMasterBalance_Sub.Form.PCash.Width + (4& * lngTpp)) + (.frmMasterBalance_Sub.Form.Px.Width + (4& * lngTpp)))
28890       lngTmp03 = ((.frmMasterBalance_Sub.Form.Cost.Width + (4& * lngTpp)) + (.frmMasterBalance_Sub.Form.Cx.Width + (4& * lngTpp)))
28900       lngAcctMast_Width = (lngTmp01 + lngTmp02 + lngTmp03)
28910       lngTmp01 = 0&: lngTmp02 = 0&: lngTmp03 = 0&
28920     End If

28930     lngTmp01 = 0&: lngTmp02 = 0&
28940     .FocusHolder.SetFocus
28950     Select Case blnShow
          Case True
28960       .chkDiscrepancies_box.Width = lngDiscrepanciesBox_Width
28970       .chkShowAcctSort2_box.Visible = False
28980       .chkShowAcctSort2.Enabled = False
28990       .chkShowAcctMast2.Enabled = False
29000       .chkPageOf2.Enabled = False
29010       .chkSaveSizePos2.Enabled = False
29020       .chkShowAcctSort2.Visible = False
29030       .chkShowAcctMast2.Visible = False
29040       .chkPageOf2.Visible = False
29050       .chkSaveSizePos2.Visible = False
29060       .chkDiscrepancies_vline01.Visible = True
29070       .chkDiscrepancies_vline02.Visible = True
29080       .chkDiscrepancies_vline03.Left = (.chkDiscrepancies_vline01.Left + lngDiscrepanciesVLine_Offset)
29090       .chkDiscrepancies_vline04.Left = (.chkDiscrepancies_vline03.Left + lngTpp)
29100       .chkDiscrepancies_vline03.Visible = True
29110       .chkDiscrepancies_vline04.Visible = True
29120       .RecsSum_Disc_lbl.Visible = True
29130       .RecsSum_Disc_lbl2.Visible = False
29140       lngTmp01 = (lngAcctSort_Left - lngDiscrepancies_Left)
29150       .chkShowAcctSort.Left = (.chkDiscrepancies.Left + lngTmp01)
29160       .chkShowAcctSort_lbl.Left = (.chkShowAcctSort.Left + lngChkBoxLbl_Offset)
29170       .chkShowAcctSort.Enabled = True
29180       .chkShowAcctMast.Left = .chkShowAcctSort.Left
29190       .chkShowAcctMast_lbl.Left = (.chkShowAcctMast.Left + lngChkBoxLbl_Offset)
29200       .chkShowAcctMast.Enabled = True
29210       lngTmp01 = (lngPageOf_Left - lngDiscrepancies_Left)
29220       .chkPageOf.Left = (.chkDiscrepancies.Left + lngTmp01)
29230       .chkPageOf_lbl.Left = (.chkPageOf.Left + lngChkBoxLbl_Offset)
29240       .chkPageOf.Enabled = True
29250       .chkSaveSizePos.Left = .chkPageOf.Left
29260       .chkSaveSizePos_lbl.Left = (.chkSaveSizePos.Left + lngChkBoxLbl_Offset)
29270       .chkSaveSizePos.Enabled = True
29280       .chkShowAcctSort.Visible = True
29290       .chkShowAcctMast.Visible = True
29300       .chkPageOf.Visible = True
29310       .chkSaveSizePos.Visible = True
29320       Select Case .chkShowAcctSort
            Case True
              ' ** As Opening, with both showing.
29330         lngTmp01 = lngFrm_Width
29340         lngTmp02 = lngFrm_Left
29350         .cmdClose.Left = lngClose_Left
29360         .cmdUpdate.Left = lngUpdate_Left
29370         .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
29380         .Sizable_lbl2.Left = .Sizable_lbl1.Left
29390         .frmMasterBalance_Sub.Width = lngSub_Width
29400         .frmMasterBalance_Sub_box.Width = lngSubBox_Width
29410         .Nav_box01.Width = .frmMasterBalance_Sub.Width
29420         .Nav_hline03.Width = .frmMasterBalance_Sub.Width
29430         .Header_vline01.Left = lngFrm_Width
29440         .Header_vline02.Left = lngFrm_Width
29450         .Detail_vline01.Left = lngFrm_Width
29460         .Detail_vline02.Left = lngFrm_Width
29470         .Footer_vline01.Left = lngFrm_Width
29480         .Footer_vline02.Left = lngFrm_Width
29490         .Header_hline01.Width = lngFrm_Width
29500         .Header_hline02.Width = lngFrm_Width
29510         .Detail_hline01.Width = lngFrm_Width
29520         .Detail_hline02.Width = lngFrm_Width
29530         .Footer_hline01.Width = lngFrm_Width
29540         .Footer_hline02.Width = lngFrm_Width
29550         .Width = lngFrm_Width
29560         .frmMasterBalance_Sub.Form.Width = .frmMasterBalance_Sub.Form.IsClosed_bg1.Width
29570         lngTmp03 = .cmdUpdate_lbl2.Left - .cmdUpdate_lbl.Left
29580         .cmdUpdate_lbl.Left = lngUpdateLbl_left
29590         .cmdUpdate_lbl2.Left = .cmdUpdate_lbl.Left + lngTmp03
29600       Case False
29610         lngTmp01 = (lngFrm_Width - lngAcctSort_Width)
29620         lngTmp02 = (lngFrm_Left + (lngAcctSort_Width / 2&))
29630         .cmdClose.Left = (lngClose_Left - lngAcctSort_Width)
29640         .cmdUpdate.Left = (lngUpdate_Left - lngAcctSort_Width)
29650         .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
29660         .Sizable_lbl2.Left = .Sizable_lbl1.Left
29670         .frmMasterBalance_Sub.Width = (lngSub_Width - lngAcctSort_Width)
29680         .frmMasterBalance_Sub_box.Width = (lngSubBox_Width - lngAcctSort_Width)
29690         .Nav_box01.Width = .frmMasterBalance_Sub.Width
29700         .Nav_hline03.Width = .frmMasterBalance_Sub.Width
29710         .Header_vline01.Left = (lngFrm_Width - lngAcctSort_Width)
29720         .Header_vline02.Left = (lngFrm_Width - lngAcctSort_Width)
29730         .Detail_vline01.Left = (lngFrm_Width - lngAcctSort_Width)
29740         .Detail_vline02.Left = (lngFrm_Width - lngAcctSort_Width)
29750         .Footer_vline01.Left = (lngFrm_Width - lngAcctSort_Width)
29760         .Footer_vline02.Left = (lngFrm_Width - lngAcctSort_Width)
29770         .Header_hline01.Width = (lngFrm_Width - lngAcctSort_Width)
29780         .Header_hline02.Width = (lngFrm_Width - lngAcctSort_Width)
29790         .Detail_hline01.Width = (lngFrm_Width - lngAcctSort_Width)
29800         .Detail_hline02.Width = (lngFrm_Width - lngAcctSort_Width)
29810         .Footer_hline01.Width = (lngFrm_Width - lngAcctSort_Width)
29820         .Footer_hline02.Width = (lngFrm_Width - lngAcctSort_Width)
29830         .Width = (lngFrm_Width - lngAcctSort_Width)
29840         .frmMasterBalance_Sub.Form.Width = .frmMasterBalance_Sub.Form.IsClosed_bg1.Width
29850         lngTmp03 = .cmdUpdate_lbl2.Left - .cmdUpdate_lbl.Left
29860         .cmdUpdate_lbl.Left = .cmdUpdate_alt_left2.Left
29870         .cmdUpdate_lbl2.Left = .cmdUpdate_lbl.Left + lngTmp03
29880       End Select
            ' ** Box is narrow only when both are off, otherwise box is wide.
29890       If .RecsSum_box.Width <> lngRecsSumBox_Width Then
29900         lngTmp04 = lngRecsSumBox_Width - .RecsSum_box.Width
29910         .RecsSum_box.Left = (.RecsSum_box.Left - lngTmp04)
29920         .RecsSum_box.Width = lngRecsSumBox_Width
29930       End If
29940     Case False
29950       .chkShowAcctSort2_box.Visible = True
29960       .chkShowAcctSort2.Enabled = True
29970       .chkShowAcctMast2.Enabled = True
29980       .chkPageOf2.Enabled = True
29990       .chkSaveSizePos2.Enabled = True
30000       .chkShowAcctSort2.Visible = True
30010       .chkShowAcctMast2.Visible = True
30020       .chkPageOf2.Visible = True
30030       .chkSaveSizePos2.Visible = True
30040       .chkShowAcctSort.Visible = False
30050       .chkShowAcctMast.Visible = False
30060       .chkPageOf.Visible = False
30070       .chkSaveSizePos.Visible = False
30080       .chkShowAcctSort.Enabled = False
30090       .chkShowAcctMast.Enabled = False
30100       .chkPageOf.Enabled = False
30110       .chkSaveSizePos.Enabled = False
30120       .chkShowAcctSort.Left = 0&
30130       .chkShowAcctSort_lbl.Left = 0&
30140       .chkShowAcctMast.Left = 0&
30150       .chkShowAcctMast_lbl.Left = 0&
30160       .chkPageOf.Left = 0&
30170       .chkPageOf_lbl.Left = 0&
30180       .chkSaveSizePos.Left = 0&
30190       .chkSaveSizePos_lbl.Left = 0&
30200       .chkDiscrepancies_vline01.Visible = False
30210       .chkDiscrepancies_vline02.Visible = False
30220       .chkDiscrepancies_vline03.Visible = False
30230       .chkDiscrepancies_vline04.Visible = False
30240       .chkDiscrepancies_vline03.Left = 0&
30250       .chkDiscrepancies_vline04.Left = 0&
30260       .chkDiscrepancies_box.Width = (.chkDiscrepancies_vline02.Left - .chkDiscrepancies_box.Left)
30270       Select Case .chkShowAcctSort
            Case True
30280         lngTmp01 = (lngFrm_Width - lngAcctMast_Width)
30290         lngTmp02 = (lngFrm_Left + (lngAcctMast_Width / 2&))
30300         .cmdClose.Left = (lngClose_Left - lngAcctMast_Width)
30310         .cmdUpdate.Left = (lngUpdate_Left - lngAcctMast_Width)
30320         .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
30330         .Sizable_lbl2.Left = .Sizable_lbl1.Left
30340         .RecsSum_Disc_lbl.Visible = True
30350         .RecsSum_Disc_lbl2.Visible = False
              ' ** Box is narrow only when both are off, otherwise box is wide.
30360         If .RecsSum_box.Width <> lngRecsSumBox_Width Then
30370           lngTmp04 = lngRecsSumBox_Width - .RecsSum_box.Width
30380           .RecsSum_box.Left = (.RecsSum_box.Left - lngTmp04)
30390           .RecsSum_box.Width = lngRecsSumBox_Width
30400         End If
30410         .frmMasterBalance_Sub.Width = (lngSub_Width - lngAcctMast_Width)
30420         .frmMasterBalance_Sub_box.Width = (lngSubBox_Width - lngAcctMast_Width)
30430         .Nav_box01.Width = .frmMasterBalance_Sub.Width
30440         .Nav_hline03.Width = .frmMasterBalance_Sub.Width
30450         .Header_vline01.Left = (lngFrm_Width - lngAcctMast_Width)
30460         .Header_vline02.Left = (lngFrm_Width - lngAcctMast_Width)
30470         .Detail_vline01.Left = (lngFrm_Width - lngAcctMast_Width)
30480         .Detail_vline02.Left = (lngFrm_Width - lngAcctMast_Width)
30490         .Footer_vline01.Left = (lngFrm_Width - lngAcctMast_Width)
30500         .Footer_vline02.Left = (lngFrm_Width - lngAcctMast_Width)
30510         .Header_hline01.Width = (lngFrm_Width - lngAcctMast_Width)
30520         .Header_hline02.Width = (lngFrm_Width - lngAcctMast_Width)
30530         .Detail_hline01.Width = (lngFrm_Width - lngAcctMast_Width)
30540         .Detail_hline02.Width = (lngFrm_Width - lngAcctMast_Width)
30550         .Footer_hline01.Width = (lngFrm_Width - lngAcctMast_Width)
30560         .Footer_hline02.Width = (lngFrm_Width - lngAcctMast_Width)
30570         .Width = (lngFrm_Width - lngAcctMast_Width)
30580         .frmMasterBalance_Sub.Form.Width = .frmMasterBalance_Sub.Form.IsClosed_bg1.Width
30590         lngTmp03 = .cmdUpdate_lbl2.Left - .cmdUpdate_lbl.Left
30600         .cmdUpdate_lbl.Left = .cmdUpdate_alt_left3.Left
30610         .cmdUpdate_lbl2.Left = .cmdUpdate_lbl.Left + lngTmp03
30620       Case False
30630         lngTmp01 = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30640         lngTmp02 = (lngFrm_Left + ((lngAcctSort_Width + lngAcctMast_Width) / 2&))
30650         .cmdClose.Left = (lngClose_Left - (lngAcctSort_Width + lngAcctMast_Width))
30660         .cmdUpdate.Left = (lngUpdate_Left - (lngAcctSort_Width + lngAcctMast_Width))
30670         .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
30680         .Sizable_lbl2.Left = .Sizable_lbl1.Left
30690         .RecsSum_Disc_lbl.Visible = False
30700         .RecsSum_Disc_lbl2.Visible = True
30710         .frmMasterBalance_Sub.Width = (lngSub_Width - (lngAcctSort_Width + lngAcctMast_Width))
30720         .frmMasterBalance_Sub_box.Width = (lngSubBox_Width - (lngAcctSort_Width + lngAcctMast_Width))
30730         .Nav_box01.Width = .frmMasterBalance_Sub.Width
30740         .Nav_hline03.Width = .frmMasterBalance_Sub.Width
30750         .Header_vline01.Left = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30760         .Header_vline02.Left = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30770         .Detail_vline01.Left = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30780         .Detail_vline02.Left = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30790         .Footer_vline01.Left = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30800         .Footer_vline02.Left = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30810         .Header_hline01.Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30820         .Header_hline02.Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30830         .Detail_hline01.Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30840         .Detail_hline02.Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30850         .Footer_hline01.Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30860         .Footer_hline02.Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30870         .Width = (lngFrm_Width - (lngAcctSort_Width + lngAcctMast_Width))
30880         .frmMasterBalance_Sub.Form.Width = .frmMasterBalance_Sub.Form.IsClosed_bg1.Width
30890         lngTmp03 = .cmdUpdate_lbl2.Left - .cmdUpdate_lbl.Left
30900         .cmdUpdate_lbl.Left = .cmdUpdate_alt_left4.Left
30910         .cmdUpdate_lbl2.Left = .cmdUpdate_lbl.Left + lngTmp03
30920         If .RecsSum_box.Width = lngRecsSumBox_Width Then
30930           lngTmp04 = ((.chkShowAcctSort2_box.Left + .chkShowAcctSort2_box.Width) - .RecsSum_box.Left)
30940           lngTmp04 = (lngTmp04 + (2& * lngTpp))
30950           .RecsSum_box.Width = (.RecsSum_box.Width - lngTmp04)
30960           .RecsSum_box.Left = (.RecsSum_box.Left + lngTmp04)
30970         End If
30980       End Select
30990     End Select

          ' ** Variables are fed empty, then populated ByRef.
31000     GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Function: modWindowFunctions.

31010     DoCmd.MoveSize (lngTmp02 - (8& * lngTpp)), (lngTop - (8& * lngTpp)), lngTmp01, lngHeight  ' ** Values are always 8 pixels off!
31020     DoEvents

31030   End With

EXITP:
31040   Exit Sub

ERRH:
31050   Select Case ERR.Number
        Case Else
31060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31070   End Select
31080   Resume EXITP

End Sub
