VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJournal_Columns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmJournal_Columns"

'VGC 09/23/2017: CHANGES!

' #########################
' ## Use VBA_RenumErrh().  39190
' #########################

'tblJOURNAL_FIELD DOESN'T GET AUTOMATICALLY UPDATED!!!    'SEE zz_qry_Journal_Column_01+

' ** Shortcut Alt keys responsive from this form:
' **   Show All Assets:  A {chkShowAllAssets on frmJournal_Columns_Sub}
' **   Close:            C {cmdClose}
' **   Date:             D {transdate on frmJournal_Columns_Sub}
' **   Enter Right:      G {opgEnterKey_optRight}
' **   Income Cash:      H {icash on frmJournal_Columns_Sub}
' **   Journal Type:     J {journaltype on frmJournal_Columns_Sub}
' **   Print Check:      K {PrintCheck on frmJournal_Columns_Sub}
' **   New Location:     L {cmdLocNew}

' **   Memo:             M {JrnlMemo_Memo}
' **   Account Number:   N {accountno on frmJournal_Columns_Sub}
' **   Commit:           O {posted on frmJournal_Columns_Sub}
' **   Print Report:     P {cmdPrintReport}
' **   Recurring Item:   R {Recur_Name on frmJournal_Columns_Sub}
' **   Asset:            S {assetno on frmJournal_Columns_Sub}
' **   Transactions:     T {frmJournal_Columns_Sub}
' **   New Recur Item:   U {cmdRecurNew}
' **   Preview Report:   V {cmdPreviewReport}
' **   New Asset:        W {cmdAssetNew}
' **   Exit:             X {cmdClose}
' **   Enter Down:       Z {opgEnterKey_optDown}
'Available Letters:
'B E F I Q Y

' ** Shortcut F-keys responsive from this form:
' **   Recalc:           F5 {RecalcTots on frmJournal_Columns_Sub}
' **   Date Picker:      F7 {cmdCalendar1 on frmJournal_Columns_Sub}
' **   Date Picker:      F8 {cmdCalendar2 on frmJournal_Columns_Sub}

' ** Shortcut Ctrl keys responsive from this form:
' **   Add:              A {cmdAdd}
' **   Delete:           D {cmdDelete}
' **   Edit:             E {cmdEdit} (also Alt+E)
' **   Refresh:          H {cmdRefresh}
' **   Scroll Left:      L {cmdScrollLeft}
' **   Check Memo:       M {cmdMemoReveal}
' **   Scroll Right:     R {cmdScrollRight}
' **   Save:             S {cmdSave on frmJournal_Columns_Sub}
' **   Switch:           T {cmdSwitch}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Commit All:       A {cmdUncomComAll}
' **   Delete All:       D {cmdUncomDelAll}

' ** Shortcut Alt-Shift keys responsive from this form:
' **   Show All:         A {opgFilter_optAll}
' **   Show Committed:   C {opgFilter_optCommitted}
' **   Map Div:          D {cmdSpecPurp_Div_Map}
' **   Map STCG/L:       G {cmdSpecPurp_Misc_MapSTCGL}
' **   Paid Total:       I {cmdSpecPurp_Sold_PaidTotal}
' **   Map LTCG:         L {cmdSpecPurp_Misc_MapLTCG}
' **   Map Int:          P {cmdSpecPurp_Int_Map}
' **   Map Split:        S {cmdSpecPurp_Purch_MapSplit}
' **   Map LTCL:         T {cmdSpecPurp_Misc_MapLTCL}
' **   Show Uncommitted: U {opgFilter_optUncommitted}

' ** Double-Click: Open Full Screen Editor  NOT YET IMPLEMENTED!

' ** Fully-constructed buttons ZOrder, bottom-to-top:
' **   .._box         : Bottom
' **   .._dots
' **   .._lbl_dim_hi
' **   .._lbl
' **   .._border
' **   .._cmd         : Top

' ** Bound Object Frame buttons ZOrder, bottom-to-top:
' **   .._lbl                 : Bottom (visible only during design view)
' **   .._raised_img_dis
' **   .._sunken_focus_img/.._sunken_focus_dots_img
' **   .._raised_focus_img/.._raised_focus_dots_img
' **   .._raised_img
' **   .._cmd                 : Top

' ** Subform events fire first!
' **   SUB Form_Open()
' **   SUB Form_Load()
' **   SUB Form_Current()
' **   SUB Form_Current()
' **   Form_Open()
' **   Form_Load()
' ** This means that JC_Msc_Cur_Set(), from the subform's OnCurrent
' ** event, fires before we even get to this OnOpen!

' ** Memo field maximum characters.
Private Const MEMO_MAX As Integer = 40

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngTpp As Long, lngHLine_Width As Long, lngWideBtn_Width As Long, lngSizable_Offset As Long
Private lngLastHeight As Long
Private lngDetail_Height As Long, lngDetail_Height_New As Long, lngDetail_Height_Diff As Long, lngDetail_Height_Min As Long
Private lngVLine5_Left As Long, lngVLine5_Left_New As Long, lngVLine5_Left_Diff As Long, lngVLine5_Left_Min As Long
Private lngVLine3_Left As Long, lngVLine7_Left As Long, lngSizable_Left As Long
Private lngSub_Width As Long, lngSub_Height As Long, lngSub_Bottom As Long
Private lngClose_Left As Long, lngAdd_Left As Long, lngEdit_Left As Long, lngDelete_Left As Long
Private lngAssets_Left As Long, lngLoc_Left As Long, lngRecur_Left As Long, lngRefresh_Left As Long
Private lngOpGrp_Height As Long, lngOpt_Offset As Long, lngOptLbl_Offset As Long

' ** Dummy variable for the subform's JC_Key_Sub_Next() function.
Private blnD1 As Boolean, blnD2 As Boolean

Private THIS_NAME_SUB As String

Private blnAssetNew_Focus As Boolean, blnLocNew_Focus As Boolean, blnRecurNew_Focus As Boolean, blnRefresh_Focus As Boolean

' ** Array: arr_varStat().
Private lngStats As Long, arr_varStat As Variant
'Private Const S_SEC As Integer = 0
Private Const S_NAM As Integer = 1
Private Const S_TXT As Integer = 2

Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGoneToReport As Boolean
Private blnGTR_Emblem As Boolean, lngGTR_Stat As Long, lngGTR_ID As Long
Private lngJrnlColID_Max As Long
Private blnIsOpen As Boolean, blnIsMaximized As Boolean, blnNotPopup As Boolean, blnCommitNoClose As Boolean
Private lngRecsCur As Long, blnSetFocus3 As Boolean, blnFromSwitch As Boolean
Private strCallingForm As String, strControlName As String
Private THAT_PROC As String, That_Erl As Long, That_Desc As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim grp As DAO.Group, usr As DAO.User
        Dim blnRetVal As Boolean

110     blnIsOpen = True

120     With Me
130       If IsNull(.OpenArgs) = False Then
140         If Trim(.OpenArgs) <> vbNullString Then

150           strCallingForm = Trim(.OpenArgs)

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170           blnGTR_Emblem = False: lngGTR_Stat = 0&: lngGTR_ID = 0&: blnGoingToReport2 = False: blnGoneToReport = False: blnSetFocus3 = False

180           If gblnGoToReport = True Then
190             If IsNull(garr_varGoToReport(GTR_CTL2)) = False Then
200               .GoToReport_Emblem_01_img.Visible = True
210               .cmdClose.Visible = False  ' ** Just temporarily.
220               DoEvents
230             End If
240           End If

250           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel
260           SetNav_Access2007 Me  ' ** Module Procedure: modXAccess_07_10_Funcs.
              'SystemColor_Move Me  ' ** Module Procedure: modSystemColorFuncs.

270           gstrFormQuerySpec = THIS_NAME
280           blnFromSwitch = gblnSwitchTo: gblnSwitchTo = False
290           blnD1 = False: blnD2 = False
300           blnCommitNoClose = False

310           .LastAcctNo = vbNullString

320           JC_Btn_BoolReset  ' ** Module Procedure: modJrnlCol_Buttons.
330           blnAssetNew_Focus = False: blnLocNew_Focus = False: blnRecurNew_Focus = False: blnRefresh_Focus = False

              ' ** CoOptions_Read() is in the subform's Form_Open().
340           If gstrTrustDataLocation = vbNullString Then
350             IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
360           End If

370           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

              ' ** THIS_NAME_SUB.
380           ThisNameSub_Set  ' ** Function: Below.

              ' ** gblnSingleUser set on frmMenu_Post.
              ' ** Determine whether the data is on this computer.
390           gblnLocalData = IsLocalData  ' ** Module Function: modStartupFuncs.

              ' ** Determine whether cmdRefresh is applicable.
400           If gblnAdmin = False Or (gblnAdmin = True And gblnSingleUser = True And gblnLocalData = True) Then
410             .cmdRefresh.Enabled = False
420           ElseIf GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
430             .cmdRefresh.Enabled = False
440           Else
450             .cmdRefresh.Enabled = True
460           End If

              ' ** Variables are fed empty, then populated ByRef.
470           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

480           lngSub_Width = .frmJournal_Columns_Sub.Width
490           lngSub_Height = .frmJournal_Columns_Sub.Height
500           lngDetail_Height = .Detail.Height

              ' ** Minimum Left for Footer_vline05:
510           lngVLine5_Left = .Footer_vline05.Left
520           lngVLine5_Left_Min = ((.cmdLocNew_raised_img.Left + .cmdLocNew_raised_img.Width) + (16& * lngTpp))
530           lngDetail_Height_Min = ((2160& / 15&) * lngTpp)
540           lngVLine3_Left = .Footer_vline03.Left
550           lngVLine7_Left = .Footer_vline07.Left

560           lngAssets_Left = .cmdAssetNew.Left
570           lngLoc_Left = .cmdLocNew.Left
580           lngRecur_Left = .cmdRecurNew.Left
590           lngRefresh_Left = .cmdRefresh.Left
600           lngAdd_Left = .cmdAdd.Left
610           lngEdit_Left = .cmdEdit.Left
620           lngDelete_Left = .cmdDelete.Left
630           lngClose_Left = .cmdClose.Left
640           lngSizable_Left = .Sizable_lbl1.Left
650           lngSizable_Offset = ((lngSizable_Left + .Sizable_lbl1.Width) - _
                (.frmJournal_Columns_Sub_box.Left + .frmJournal_Columns_Sub_box.Width))  ' ** 30&

660           lngOpGrp_Height = .opgEnterKey.Height
670           lngOpt_Offset = (.opgEnterKey_optRight.Top - .opgEnterKey.Top)
680           lngOptLbl_Offset = (.opgEnterKey_optRight_lbl.Top - .opgEnterKey.Top)

690           lngWideBtn_Width = .cmdLocNew.Width
700           lngHLine_Width = .Header_hline01.Width  ' ** Arbitrary choice, but representative.

710           lngVLine5_Left_Diff = 0&: lngDetail_Height_Diff = 0&
720           lngLastHeight = 0&
730           lngVLine5_Left_New = 0&: lngDetail_Height_New = 0&
740           lngSub_Bottom = 0&

              ' ** Pass the initial positions to modJrnlCol_Buttons.
750           JC_Btn_Resize_Set lngWideBtn_Width, lngVLine5_Left, lngAssets_Left, lngLoc_Left, lngRecur_Left, _
                lngRefresh_Left, lngAdd_Left, lngEdit_Left, lngDelete_Left, lngClose_Left, lngOpGrp_Height, _
                lngOpt_Offset, lngOptLbl_Offset, lngSizable_Offset  ' ** Module Procedure: modJrnlCol_Buttons.

760           If IsNull(.currow_hilite_forecolor) = True Then
770             .currow_hilite_forecolor = CLR_AC07  ' ** Default.
780           End If

790           opgEnterKey_AfterUpdate  ' ** Procedure: Below.
800           opgFilter_AfterUpdate  ' ** Procedure: Below.

810           blnIsMaximized = IsMaximized(Me)  ' ** Module Function: modWindowFunctions.

820           Form_Resize  ' ** Procedure: Below.

830         Else
840           Cancel = -1
850         End If
860       Else
870         Cancel = -1
880       End If

890     End With

900     If Cancel = -1 Then
910       DoCmd.Hourglass False
920       Beep
930       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
940     End If

EXITP:
950     Set usr = Nothing
960     Set grp = Nothing
970     Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     Application.Echo True
1000    THAT_PROC = THIS_PROC
1010    That_Erl = Erl: That_Desc = ERR.description
1020    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1030    Resume EXITP

End Sub

Private Sub Form_Load()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

1010    With Me

1020      StatusBar_Load  ' ** Procedure: Below.

1030      Set dbs = CurrentDb
          ' ** Empty tmpAccount.
1040      Set qdf = dbs.QueryDefs("qryJournal_Columns_10_Account_04")
1050      qdf.Execute
1060      Set qdf = Nothing
1070      DoEvents
          ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
          ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
          ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
          ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
          ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
          ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
          ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
          ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
1080      Set qdf = dbs.QueryDefs("qryJournal_Columns_10_Account_05")
1090      qdf.Execute
1100      Set qdf = Nothing
1110      DoEvents
1120      dbs.Close
1130      Set dbs = Nothing
1140      DoEvents

1150      If IsNull(.currow_hilite_forecolor) = False Then
1160        If .currow_hilite_forecolor = 0& Then .currow_hilite_forecolor = CLR_AC07
1170        CLR_HILITE = .currow_hilite_forecolor
1180      Else
1190        If CLR_HILITE = 0& Then
1200          CLR_HILITE = CLR_AC07  ' ** Default.
1210        End If
1220        .currow_hilite_forecolor = CLR_HILITE
1230      End If

1240      .frmJournal_Columns_Sub_lbl_raised_img_dis.Visible = False  ' ** Never used.
1250      .frmJournal_Columns_Sub_lbl_raised_img.Visible = True

1260      .frmJournal_Columns_Sub.Form.ctlBack1.ForeColor = CLR_HILITE
1270      .frmJournal_Columns_Sub.Form.ctlBack2.ForeColor = CLR_HILITE
1280      DoEvents

1290      If gblnGoToReport = True Then
1300        .frmJournal_Columns_Sub.Form.GoToReport_arw_jcol_bge_img = .GoToReport_arw_jcol_bge_img
1310        .frmJournal_Columns_Sub.Form.GoToReport_arw_jcol_blu1_img = .GoToReport_arw_jcol_blu1_img
1320        .frmJournal_Columns_Sub.Form.GoToReport_arw_jcol_blu2_img = .GoToReport_arw_jcol_blu2_img
1330      End If

          ' ** If this came from frmJournal, the Post menu may still be open.
1340      strDocName = "frmMenu_Post"
1350      If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
1360        DoCmd.Close acForm, strDocName
1370      End If

1380      DoCmd.Hourglass False

1390    End With

EXITP:
1400    Set qdf = Nothing
1410    Set dbs = Nothing
1420    Exit Sub

ERRH:
100     DoCmd.Hourglass False
110     Application.Echo True
120     THAT_PROC = THIS_PROC
130     That_Erl = Erl: That_Desc = ERR.description
140     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
150     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer, lngRetVal As Long

1510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1520    intShiftDown = (Shift And acShiftMask) > 0
1530    intAltDown = (Shift And acAltMask) > 0
1540    intCtrlDown = (Shift And acCtrlMask) > 0

1550    If lngRecsCur = 0& Then lngRecsCur = Me.frmJournal_Columns_Sub.Form.RecCnt  ' ** Form Function: frmJournal_Columns_Sub.

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Recalc:           F5 {RecalcTots on frmJournal_Columns_Sub}
        ' **   Date Picker:      F7 {cmdCalendar1 on frmJournal_Columns_Sub}
        ' **   Date Picker:      F8 {cmdCalendar2 on frmJournal_Columns_Sub}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   TransDate:        D {Date on frmJournal_Columns_Sub}
        ' **   Income Cash:      H {icash on frmJournal_Columns_Sub}
        ' **   Journal Type:     J {journaltype on frmJournal_Columns_Sub}
        ' **   Print Check:      K {PrintCheck on frmJournal_Columns_Sub}
        ' **   Account Number:   N {accountno on frmJournal_Columns_Sub}
        ' **   Commit:           O {posted on frmJournal_Columns_Sub}
        ' **   Recurring Item:   R {Recur_Name on frmJournal_Columns_Sub}
        ' **   Asset:            S {assetno on frmJournal_Columns_Sub}
        ' **   Transactions:     T {frmJournal_Columns_Sub}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Add:              A {cmdAdd}
        ' **   Delete:           D {cmdDelete}
        ' **   Edit:             E {cmdEdit}
        ' **   Refresh:          H {cmdRefresh}
        ' **   Scroll Left:      L {cmdScrollLeft}
        ' **   Check Memo:       M {cmdMemoReveal}
        ' **   Scroll Right:     R {cmdScrollRight}
        ' **   Save:             S {cmdSave on frmJournal_Columns_Sub}
        ' **   Switch:           T {cmdSwitch}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Commit All:       A {cmdUncomComAll}
        ' **   Delete All:       D {cmdUncomDelAll}

        ' ** Shortcut Alt-Shift keys to other forms and functionality:
        ' **   Show All:         A {opgFilter_optAll}
        ' **   Show Committed:   C {opgFilter_optCommitted}
        ' **   Map Div:          D {cmdSpecPurp_Div_Map}
        ' **   Map STCG/L:       G {cmdSpecPurp_Misc_MapSTCGL}
        ' **   Paid Total:       I {cmdSpecPurp_Sold_PaidTotal}
        ' **   Map LTCG:         L {cmdSpecPurp_Misc_MapLTCG}
        ' **   Map Int:          P {cmdSpecPurp_Int_Map}
        ' **   Map Split:        S {cmdSpecPurp_Purch_MapSplit}
        ' **   Map LTCL:         T {cmdSpecPurp_Misc_MapLTCL}
        ' **   Show Uncommitted: U {opgFilter_optUncommitted}

        ' ** Plain keys.
1560    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1570      intRetVal = JC_Key_Frm(KeyCode, 1, THAT_PROC, That_Erl, That_Desc, Me)  ' ** Module Function: modJrnlCol_Keys.
1580    End If

        ' ** Alt keys.
1590    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1600      intRetVal = JC_Key_Frm(KeyCode, 2, THAT_PROC, That_Erl, That_Desc, Me)  ' ** Module Function: modJrnlCol_Keys.
1610    End If

        ' ** Ctrl keys.
1620    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1630      intRetVal = JC_Key_Frm(KeyCode, 3, THAT_PROC, That_Erl, That_Desc, Me)  ' ** Module Function: modJrnlCol_Keys.
1640    End If

        ' ** Ctrl-Shift keys.
1650    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1660      intRetVal = JC_Key_Frm(KeyCode, 4, THAT_PROC, That_Erl, That_Desc, Me)  ' ** Module Function: modJrnlCol_Keys.
1670    End If

        ' ** Alt-Shift keys.
1680    If (Not intCtrlDown) And intAltDown And intShiftDown Then
1690      intRetVal = JC_Key_Frm(KeyCode, 5, THAT_PROC, That_Erl, That_Desc, Me)  ' ** Module Function: modJrnlCol_Keys.
1700    End If

EXITP:
1710    KeyCode = intRetVal
1720    Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2046  ' ** The command or action isn't available now (first or last record).
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub Form_Resize()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Resize"

        Dim ctl As Access.Control
        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngClose_Width As Long, lngClose_Diff As Long
        Dim blnShrinking As Boolean

1810    blnShrinking = False

        ' ** Control ZOrder must always keep cmdAssetNew, cmdRefresh, cmdLocNew, cmdRecurNew on top!
        ' ** All 4 Tops are normally standard Top
        ' ** All 4 Heights are normally standard Height.
        ' ** All 4 Widths are normally standard Width.

1820    With Me

1830      GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Procedure: modWindowFunctions.

          ' ** On open, hits once True, once False.
          ' ** Both cases require no changes on my screen.
          ' ** On my screen, cmdClose gutter is 135& in Design, 300& in Run (from lngWidth).
          ' ** So, 300& is taken up by form borders, etc.

1840      Select Case blnIsOpen
          Case True

1850        If lngWidth <> lngFrm_Width Then

              ' ** Normal Footer_vline05 design-time position:
              ' **   Forms(0).Width - Forms(0).Footer_vline05.Left = 4575 @ 15 TPP
              ' ** Normal cmdClose design-time gutter:
              ' **   Forms(0).Width - (Forms(0).cmdClose.Left + Forms(0).cmdClose.Width) = 135
              ' ** On my machine, lngWidth = Me.Width + 180&.
              ' ** So, cmdClose adjustment should be +180&, rather than +285&
              ' ** With 180& adjustment:
              ' **   Form_Resize() 1
              ' **   Form_Resize() 2
              ' **   lngVLine5_Left_Diff = 0
              ' **   Form_Resize() 1
              ' **   Form_Resize() 3

              ' ** GoToReport_Emblem_00_img design-time:
              ' **   Top:    45    Twips
              ' **   Left:   17925
              ' **   Width:  720
              ' **   Height: 720
              ' **   Gutter: 0 ((Width - (Left + Width)) = 0)

1860          lngVLine5_Left_New = (lngWidth - (((4575& / 15&) * lngTpp) + ((180& / 15&) * lngTpp)))  ' ** My Twips, divided by My TPP, times Their TPP.
1870          lngDetail_Height_Diff = (lngHeight - lngFrm_Height)
1880          If lngDetail_Height_New = 0& Then
1890            lngDetail_Height_New = (lngDetail_Height + lngDetail_Height_Diff)
1900            If lngLastHeight = 0& Then
1910              lngLastHeight = lngDetail_Height_New
1920              If lngVLine5_Left_Diff = 0& Then
                    ' ** Since I'm no longer maximizing on open, no adjustments need be made.
1930                lngVLine5_Left_Diff = (lngVLine5_Left_New - lngVLine5_Left)  ' ** May be positive or negative.
1940              End If  ' ** lngVLine5_Left_Diff.
1950            End If  ' ** lngLastHeight.
1960          End If  ' ** lngDetail_Height_New.

1970        Else
              ' ** No change in window width.

1980          lngClose_Width = .cmdClose.Width

              'In Design mode, the subform gutter is 45&, its box, 30&.
              'lngWidth - subwidth = 420&.
              'If we use the same 300& as above, that leaves 120&, or 60& per side.
              'I've never been able to figure out why this arithmetic never matches!

1990          If (lngClose_Left > lngWidth) Or (lngClose_Left > ((lngWidth - (lngClose_Width + (20& * lngTpp))) - (8& * lngTpp))) Then
                ' ** Way off the right side!    Not far enough from the right-hand edge.          300&                120&
                ' ** Last 120& is additional adjustment made on 09/23/2011.

                'These changes aren't really accurate unless
                'each one is related to the current lngWidth!
                'Positioning them by percentage may be necessary.
                'For example,
                '  The subform is 0.993563958165728% of form width.
                '  cmdClose is at 0.944489139179405%
                '  cmdAssetNew is at 0.49798873692679%
2000            lngClose_Diff = (lngClose_Left - (((lngWidth - lngClose_Width) - (20& * lngTpp)) + (8& * lngTpp)))
                ' ** The additional adjustment keeps the Sizable labels correct,
                ' ** but the cmdClose gutter is still too small!
2010            lngVLine5_Left_Diff = (lngVLine5_Left - lngClose_Diff)
2020            .cmdAssetNew.Left = (lngAssets_Left - lngClose_Diff)
2030            .cmdLocNew.Left = (lngLoc_Left - lngClose_Diff)
2040            .cmdRecurNew.Left = (lngRecur_Left - lngClose_Diff)
2050            .Footer_vline03.Left = (lngVLine3_Left - lngVLine5_Left_Diff)
2060            .Footer_vline04.Left = ((lngVLine3_Left - lngVLine5_Left_Diff) + lngTpp)
2070            .cmdRefresh.Left = (lngRefresh_Left - lngClose_Diff)
2080            .Footer_vline05.Left = (lngVLine5_Left - lngVLine5_Left_Diff)
2090            .Footer_vline06.Left = ((lngVLine5_Left - lngVLine5_Left_Diff) + lngTpp)
2100            .cmdAdd_box.Left = (((lngVLine5_Left - lngVLine5_Left_Diff) + lngTpp) + lngTpp)
2110            .cmdAdd.Left = (lngAdd_Left - lngClose_Diff)
2120            .cmdEdit.Left = (lngEdit_Left - lngClose_Diff)
2130            .cmdDelete.Left = (lngDelete_Left - lngClose_Diff)
2140            .Footer_vline07.Left = (lngVLine7_Left - lngVLine5_Left_Diff)
2150            .Footer_vline08.Left = ((lngVLine7_Left - lngVLine5_Left_Diff) + lngTpp)
2160            .cmdClose.Left = (lngClose_Left - lngClose_Diff)
2170            .frmJournal_Columns_Sub.Width = (lngSub_Width - lngClose_Diff)
2180            .frmJournal_Columns_Sub_box.Width = (.frmJournal_Columns_Sub.Width + (2& * lngTpp))
2190            .Sizable_lbl1.Left = (((.frmJournal_Columns_Sub_box.Left + .frmJournal_Columns_Sub_box.Width) + _
                  lngSizable_Offset) - .Sizable_lbl1.Width)
2200            .Sizable_lbl2.Left = .Sizable_lbl1.Left
                ' ** Sizable, like Emblem, has Zero gutter.
2210            GTRStuff 3, False  ' ** Procedure: Below.

2220          End If

2230        End If  ' ** lngWidth, lngFrm_Width.

2240        blnIsOpen = False

2250      Case False  ' ** Not on open.

            ' ** If it's wider, all controls remain at normal position. They do not move right!
            ' ** If it's narrower, the smaller buttons take over, and the others move left to the minimum.

2260        lngVLine5_Left_New = (lngWidth - (((4575& / 15&) * lngTpp) + ((180& / 15&) * lngTpp)))  ' ** My Twips, divided by My TPP, times Their TPP.
2270        If lngHeight > lngFrm_Height Then
2280          lngDetail_Height_Diff = (lngHeight - lngFrm_Height)
2290          lngDetail_Height_New = (lngDetail_Height + lngDetail_Height_Diff)
2300        ElseIf lngHeight < lngFrm_Height Then
2310          lngDetail_Height_Diff = (lngFrm_Height - lngHeight)
2320          lngDetail_Height_New = (lngDetail_Height - lngDetail_Height_Diff)
2330        Else
2340          lngDetail_Height_Diff = 0&
2350          lngDetail_Height_New = lngDetail_Height
2360        End If

2370        If lngDetail_Height_New < lngLastHeight Then  ' ** Won't shrink if lngLastHeight = 0.
2380          blnShrinking = True
2390        Else
2400          blnShrinking = False
2410        End If
2420        lngLastHeight = lngDetail_Height_New

2430        If lngWidth <> lngFrm_Width Then
2440          If lngVLine5_Left_New >= lngVLine5_Left_Min Then
                ' ** Adjustments may be necessary.

2450            If lngVLine5_Left_New > lngVLine5_Left Then
                  ' ** The adjusted window width is bigger than the non-maximized (restored) form width,
                  ' ** so everything remains at opening position, except the subform.

2460              lngVLine5_Left_Diff = (lngVLine5_Left_New - lngVLine5_Left)

                  ' ** Buttons back to normal.
2470              JC_Btn_Resize_Btn Me, True, lngVLine5_Left_New, lngVLine5_Left_Diff  ' ** Module Procedure: modJrnlCol_Buttons.

2480              .Header_hline01.Width = (lngHLine_Width + lngVLine5_Left_Diff)
2490              .frmJournal_Columns_Sub.Width = ((lngSub_Width + lngVLine5_Left_Diff) - (8& * lngTpp))
2500              .Nav_box01.Width = .frmJournal_Columns_Sub.Width
2510              .Nav_hline03.Width = .frmJournal_Columns_Sub.Width

2520            ElseIf lngVLine5_Left_New < lngVLine5_Left Then
                  ' ** The adjusted window width is smaller than the non-maximized (restored) form width.

2530              lngVLine5_Left_Diff = (lngVLine5_Left - lngVLine5_Left_New)

                  ' ** Smaller buttons take over.
2540              JC_Btn_Resize_Btn Me, False, lngVLine5_Left_New, lngVLine5_Left_Diff  ' ** Module Procedure: modJrnlCol_Buttons.

2550              .Header_hline01.Width = (lngHLine_Width - lngVLine5_Left_Diff)
2560              .frmJournal_Columns_Sub.Width = ((lngSub_Width - lngVLine5_Left_Diff) - (8& * lngTpp))
2570              .Nav_box01.Width = .frmJournal_Columns_Sub.Width
2580              .Nav_hline03.Width = .frmJournal_Columns_Sub.Width

2590            Else
                  ' ** It's been restored to the non-maximized form width, though buttons and such may still be elsewhere!

2600              lngVLine5_Left_Diff = 0&

                  ' ** Buttons back to normal.
2610              JC_Btn_Resize_Btn Me, True, lngVLine5_Left_New, lngVLine5_Left_Diff  ' ** Module Procedure: modJrnlCol_Buttons.

2620              .Header_hline01.Width = lngHLine_Width
2630              .frmJournal_Columns_Sub.Width = lngSub_Width
2640              .Nav_box01.Width = .frmJournal_Columns_Sub.Width
2650              .Nav_hline03.Width = .frmJournal_Columns_Sub.Width

2660            End If

2670            .Header_hline02.Width = .Header_hline01.Width
2680            .Header_vline01.Left = .Header_hline01.Width
2690            .Header_vline02.Left = .Header_vline01.Left
2700            .Footer_hline01.Width = .Header_hline01.Width
2710            .Footer_hline02.Width = .Header_hline01.Width
2720            .Footer_vline01.Left = .Header_vline01.Left
2730            .Footer_vline02.Left = .Header_vline01.Left
2740            .frmJournal_Columns_Sub_box.Width = .frmJournal_Columns_Sub.Width + (2& * lngTpp)
2750            .Sizable_lbl1.Left = (((.frmJournal_Columns_Sub_box.Left + .frmJournal_Columns_Sub_box.Width) + _
                  lngSizable_Offset) - .Sizable_lbl1.Width)
2760            .Sizable_lbl2.Left = .Sizable_lbl1.Left
                ' ** Sizable, like Emblem, has Zero gutter.
2770            GTRStuff 3, False  ' ** Procedure: Below.

2780          Else
                ' ** Hit minimum Left, so stop moving left and make sure things are at their minimum position.

2790            lngVLine5_Left_New = lngVLine5_Left_Min
2800            lngVLine5_Left_Diff = (lngVLine5_Left - lngVLine5_Left_New)

                ' ** Smaller buttons take over.
2810            JC_Btn_Resize_Btn Me, False, lngVLine5_Left_New, lngVLine5_Left_Diff  ' ** Module Procedure: modJrnlCol_Buttons.

2820            .Header_hline01.Width = (lngHLine_Width - lngVLine5_Left_Diff)
2830            .frmJournal_Columns_Sub.Width = ((lngSub_Width - lngVLine5_Left_Diff) - (8& * lngTpp))
2840            .Nav_hline03.Width = .frmJournal_Columns_Sub.Width

2850            .Header_hline02.Width = .Header_hline01.Width
2860            .Header_vline01.Left = .Header_hline01.Width
2870            .Header_vline02.Left = .Header_vline01.Left
2880            .Footer_hline01.Width = .Header_hline01.Width
2890            .Footer_hline02.Width = .Header_hline01.Width
2900            .Footer_vline01.Left = .Header_vline01.Left
2910            .Footer_vline02.Left = .Header_vline01.Left
2920            .frmJournal_Columns_Sub_box.Width = .frmJournal_Columns_Sub.Width + (2& * lngTpp)
2930            .Sizable_lbl1.Left = (((.frmJournal_Columns_Sub_box.Left + .frmJournal_Columns_Sub_box.Width) + _
                  lngSizable_Offset) - .Sizable_lbl1.Width)
2940            .Sizable_lbl2.Left = .Sizable_lbl1.Left
                ' ** Sizable, like Emblem, has Zero gutter.
2950            GTRStuff 3, False  ' ** Procedure: Below.

2960          End If  ' ** lngVLine5_Left_Min.
2970        End If  ' ** lngWidth, lngFrm_Width.

2980        If lngHeight <> lngFrm_Height Then
2990          If lngDetail_Height_New >= lngDetail_Height_Min Then

3000            If lngHeight > lngFrm_Height Then
                  ' ** The adjusted window height is bigger than the non-maximized (restored) form height.
3010              If blnShrinking = False Then
3020                .Detail.Height = lngDetail_Height_New
3030              End If
3040              .frmJournal_Columns_Sub.Height = (lngSub_Height + lngDetail_Height_Diff)
3050            ElseIf lngHeight < lngFrm_Height Then
                  ' ** The adjusted window height is smaller than the non-maximized (restored) form height.
3060              If blnShrinking = False Then
3070                .Detail.Height = lngDetail_Height_New  'NOT SURE HOW THIS WORKS IF IT'S BEFORE THE LEGEND IS MOVED!!
3080              End If
3090              .frmJournal_Columns_Sub.Height = (lngSub_Height - lngDetail_Height_Diff)
3100            Else
                  ' ** It's been restored to the non-maximized form height, though labels and such may still be elsewhere!
3110              If blnShrinking = False Then
3120                .Detail.Height = lngDetail_Height_New
3130              End If
3140              .frmJournal_Columns_Sub.Height = lngSub_Height
3150            End If

3160            .detail_height_line.Top = (lngDetail_Height_New - .detail_height_line.Height)
3170            .frmJournal_Columns_Sub_box.Height = (.frmJournal_Columns_Sub.Height + (6& * lngTpp))  ' ** 90 Twips.
3180            lngSub_Bottom = (.frmJournal_Columns_Sub.Top + .frmJournal_Columns_Sub.Height)
3190            For Each ctl In .Detail.Controls
3200              With ctl
3210                If IsNull(.Tag) = False Then
3220                  If Trim(.Tag) <> vbNullString Then
3230                    If .Tag = "Subform Footer" Then
3240                      Select Case .ControlType
                          Case acLine
3250                        If Left(.Name, 4) = "Nav_" Then
3260                          .Top = lngSub_Bottom
3270                          If .Name = "Nav_hline02" Then  ' ** The only one with a different Top.
3280                            .Top = .Top + lngTpp
3290                          End If
3300                        Else
                              ' ** opgEnterKey_vline01 - opgEnterKey_vline04, opgFilter_vline01 - opgFilter_vline04.
3310                          .Top = (lngSub_Bottom + (6& * lngTpp))
3320                        End If
3330                      Case acLabel
                            ' ** ShortcutMenu_lbl, opgFilter_lbl, opgEnterKey_lbl.
3340                        If .Name = "ShortcutMenu_up_arrow_lbl" Or .Name = "ShortcutMenu_down_arrow_lbl" Then
3350                          .Top = (lngSub_Bottom + (5& * lngTpp))
3360                        Else
3370                          .Top = (lngSub_Bottom + (6& * lngTpp))
3380                        End If
3390                      Case acRectangle
3400                        Select Case .Name
                            Case "opgEnterKey_box", "opgFilter_box"
3410                          .Top = (lngSub_Bottom + (6& * lngTpp))
3420                        Case "Nav_box01"
3430                          .Top = lngSub_Bottom
3440                        End Select
3450                      Case acCommandButton  'cmdMemoReveal
3460                        .Top = lngSub_Bottom + (8& * lngTpp)
3470                      Case acBoundObjectFrame
3480                        .Top = lngSub_Bottom + (8& * lngTpp)
3490                      Case acImage
3500                        .Top = lngSub_Bottom + (8& * lngTpp)
3510                      End Select
3520                    End If  ' ** Subform Footer.
3530                  End If  ' ** Tag.
3540                End If  ' ** Tag.
3550              End With  ' ** ctl.
3560            Next  ' ** ctl.

                ' ** Now move the Option Group (which doesn't have a Tag).
3570            JC_Btn_Resize_Opg Me, blnShrinking, lngSub_Bottom  ' ** Module Procedure: modJrnlCol_Buttons.

3580            If blnShrinking = True Then
3590              .Detail.Height = lngDetail_Height_New
3600            End If
3610            If (.detail_height_line.Top + .detail_height_line.Height) < .Detail.Height Then
3620              .Detail.Height = (.detail_height_line.Top + .detail_height_line.Height)
3630            End If

                '.frmJournal_Columns_Sub.Form.TimerInterval = 500&

3640          Else
                ' ** Hit minimum Height, so stop moving up and make sure things are at their minimum position.

3650          End If  ' ** lngDetail_Height_Min.
3660        End If  ' ** lngHeight, lngFrm_Height.

3670      End Select  ' ** blnIsOpen.

          '.TwipsWidth_Form = .Width
          '.TwipsWidth_Window = lngWidth
          '.TwipsClose_Left = .cmdClose.Left
          '.TwipsClose_Width = .cmdClose.Width
          '.TwipsSub_Height = .frmJournal_Columns_Sub.Height
          '.TwipsSub_Hdr_Height = .frmJournal_Columns_Sub.Form.FormHeader.Height
          '.TwipsSub_HdrLine_Top = .frmJournal_Columns_Sub.Form.form_header_height_line.Top
          '.TwipsSub_HdrLine_Height = .frmJournal_Columns_Sub.Form.form_header_height_line.Height

3680    End With  ' ** Me.

EXITP:
3690    Set ctl = Nothing
3700    Exit Sub

ERRH:
980     Application.Echo True
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Select Case ERR.Number
        Case 2100  ' ** The control or subform control is too large for this location.
          ' ** Ignore.
1020    Case Else
1030      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1040    End Select
1050    Resume EXITP

End Sub

Private Sub Form_Timer()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim frmSub As Access.Form
        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim strNextSub As String, blnSetFocus2 As Boolean
        Dim strFormName As String
        Dim lngJrnlColID As Long
        Dim blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean, lngNewJrnlColID As Long
        Dim blnToTaxLot As Boolean, datPostingDate As Date, strSaveMoveCtl As String, blnGTR_NoAdd As Boolean
        Dim lngTmp03 As Long

        Static lngTmp01 As Long, lngTmp02 As Long, lngTmp04 As Long

3810    With Me
3820      .TimerInterval = 0&
3830      Application.Echo True

3840      Set frmSub = .frmJournal_Columns_Sub.Form

3850      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
3860        lngGTR_Stat = lngGTR_Stat + 1&
3870        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
3880        lngTmp01 = lngTmp01 + 1&
3890        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
3900        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
3910        If lngTmp01 > 1& Then
3920          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
3930        End If
3940        If lngTmp01 = 24& Then
3950          lngTmp01 = 0&
3960          lngTmp02 = lngTmp02 + 1&
3970          If lngTmp02 < 5& Or (blnGoingToReport2 = True And lngTmp02 < 26&) Then
3980            .TimerInterval = 50&
3990          Else
4000            blnGTR_Emblem = False
4010          End If
4020        Else
4030          .TimerInterval = 50&
4040        End If
4050      End If  ' ** blnGTR_Emblem.

4060      If gblnSetFocus = True Then
4070        gblnSetFocus = False
4080        lngTmp03 = 0&
4090        strFormName = vbNullString
4100  On Error Resume Next
4110        Do Until strFormName = THIS_NAME
4120          lngTmp03 = lngTmp03 + 1&
4130          DoCmd.SelectObject acForm, THIS_NAME, False
4140          strFormName = Screen.ActiveForm.Name
4150          If lngTmp03 > 100& Then Exit Do  ' ** This got stuck once, because of a mapping modal business!
4160        Loop
4170  On Error GoTo ERRH

4180        If IsMaximized(Me) = True And blnIsMaximized = False Then  ' ** Module Function: modWindowFunctions.
4190          DoCmd.Restore
4200        End If

4210        .cmdAssetNew.Enabled = True
4220        .cmdLocNew.Enabled = True
4230        .cmdRecurNew.Enabled = True

4240        blnSetFocus2 = True

4250      End If  ' ** gblnSetFocus.

4260      If blnSetFocus2 = True Then
4270        blnSetFocus2 = False

4280        If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
4290          CmdBars_Hide True  ' ** Module Procedure: modWindowFunctions
4300        End If

4310        If .ToTaxLot <> 0& Then

4320          DoCmd.Hourglass False
4330          If .JrnlCol_ID > 0& Then
4340            frmSub.MoveRec 0, .JrnlCol_ID  ' ** Form Procedure: frmJournal_Columns_Sub.
4350            DoEvents
4360          End If

4370          Select Case .ToTaxLot
              Case 1&
                ' ** Was sent, but nothing happened! (Or a problem within the KeyDown event!)
4380            .frmJournal_Columns_Sub.SetFocus
4390            frmSub.Cost.Undo  ' = 0@
4400            frmSub.PCash.SetFocus
4410            lngTmp03 = fSetScrollBarPosHZ(frmSub, 999&)  ' ** Module Function: modScrollBarFuncs.
4420          Case 2&
                ' ** cmdOK Return, Single-Lot Sale.
4430            .frmJournal_Columns_Sub.SetFocus
4440            frmSub.description.SetFocus
4450            lngTmp03 = fSetScrollBarPosHZ(frmSub, 999&)  ' ** Module Function: modScrollBarFuncs.
4460          Case 3&
                ' ** cmdOK Return, Multi-Lot Sale.
4470            .frmJournal_Columns_Sub.SetFocus

4480            blnToTaxLot = frmSub.TaxLot_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.
4490            blnNextRec = frmSub.NextRec_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.
4500            blnFromZero = frmSub.FromZero_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.
4510            datPostingDate = frmSub.PostDate_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.
4520            strSaveMoveCtl = frmSub.SaveMoveCtl_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.
4530            blnNoMove = frmSub.NoMove_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.
4540            lngNewJrnlColID = frmSub.NewJColID_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.
4550            blnGoingToReport = frmSub.GoingToReport_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.
4560            blnGoneToReport = frmSub.GTRGone_GetSet(True, blnGoneToReport, lngGTR_ID)  ' ** Form Procedure: frmJournal_Columns_Sub.
4570            blnGTR_Emblem = frmSub.GTR_Emblem_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.
4580            blnGTR_NoAdd = frmSub.GTR_NoAdd_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.

                '.frmJournal_Columns_Sub.Form.TaxLot_Form False  ' ** Form Procedure: frmJournal_Columns_Sub.
4590            JC_Frm_TaxLot_Form Me, False, blnToTaxLot, blnGoingToReport, blnGoneToReport, _
                  blnNextRec, blnFromZero, datPostingDate, strSaveMoveCtl, blnNoMove, lngNewJrnlColID, _
                  lngGTR_ID, blnGTR_Emblem, blnGTR_NoAdd  ' ** Module Procedure: modJrnlCol_Forms.

4600            frmSub.TaxLot_GetSet False, blnToTaxLot  ' ** Form Procedure: frmJournal_Columns_Sub.
4610            frmSub.NextRec_GetSet False, blnNextRec  ' ** Form Procedure: frmJournal_Columns_Sub.
4620            frmSub.FromZero_GetSet False, blnFromZero  ' ** Form Procedure: frmJournal_Columns_Sub.
4630            frmSub.PostDate_GetSet False, datPostingDate  ' ** Form Procedure: frmJournal_Columns_Sub.
4640            frmSub.SaveMoveCtl_GetSet False, strSaveMoveCtl  ' ** Form Procedure: frmJournal_Columns_Sub.
4650            frmSub.NoMove_GetSet False, blnNoMove  ' ** Form Procedure: frmJournal_Columns_Sub.
4660            frmSub.NewJColID_GetSet False, lngNewJrnlColID  ' ** Form Procedure: frmJournal_Columns_Sub.
4670            frmSub.GoingToReport_GetSet False, blnGoingToReport  ' ** Form Procedure: frmJournal_Columns_Sub.
4680            frmSub.GTRGone_GetSet True, blnGoneToReport, lngGTR_ID  ' ** Form Procedure: frmJournal_Columns_Sub.
4690            frmSub.GTR_Emblem_GetSet False, blnGTR_Emblem  ' ** Form Procedure: frmJournal_Columns_Sub.
4700            frmSub.GTR_NoAdd_GetSet False, blnGTR_NoAdd  ' ** Form Procedure: frmJournal_Columns_Sub.

4710          Case -1&
                ' **   -1 Input missing.
4720            .frmJournal_Columns_Sub.SetFocus
4730            If .TaxLotFrom <> vbNullString And .TaxLotFrom <> "cmdAssetNew" Then
4740              frmSub.Controls(.TaxLotFrom).Undo
4750            End If
4760            frmSub.assetno.SetFocus
4770            DoEvents
4780            lngTmp03 = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
4790          Case -2&
                ' **   -2 No holdings.
4800            .frmJournal_Columns_Sub.SetFocus
4810            If .TaxLotFrom <> vbNullString And .TaxLotFrom <> "cmdAssetNew" Then
4820              frmSub.Controls(.TaxLotFrom).Undo
4830            End If
4840            frmSub.assetno.SetFocus
4850            DoEvents
4860            lngTmp03 = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
4870          Case -3&
                ' **   -3 Insufficient holdings.
4880            .frmJournal_Columns_Sub.SetFocus
4890            frmSub.shareface.SetFocus  ' ** Field already updated.
4900  On Error Resume Next
4910            frmSub.shareface = 0#
4920  On Error GoTo ERRH
4930            DoEvents
4940            lngTmp03 = fSetScrollBarPosHZ(frmSub, 999&)  ' ** Module Function: modScrollBarFuncs.
4950          Case -4&
                ' ** cmdCancel Return.
4960            If blnGoneToReport = False Then
4970              .frmJournal_Columns_Sub.SetFocus
4980              frmSub.Cost.Undo  ' = 0@
4990              frmSub.PCash.SetFocus
5000              DoEvents
5010              frmSub.cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
5020              lngTmp03 = fSetScrollBarPosHZ(frmSub, 999&)  ' ** Module Function: modScrollBarFuncs.
5030            Else
5040              blnGoneToReport = False
5050              .cmdAdd.SetFocus
                  'frmSub.AddRec  ' ** Form Procedure: frmJournal_Columns_Sub.
5060              frmSub.AddRec_Send  ' ** Form Procedure: frmJournal_Columns_Sub.
5070              DoEvents
5080              frmSub.transdate.SetFocus
5090              frmSub.transdate_GotFocus  ' ** Form Procedure: frmJournal_Columns_Sub.
5100            End If
5110          Case -9
                ' **   -9 Data problem.
5120            .frmJournal_Columns_Sub.SetFocus
5130            If .TaxLotFrom <> vbNullString And .TaxLotFrom <> "cmdAssetNew" Then
5140              frmSub.Controls(.TaxLotFrom).Undo
5150            End If
5160            frmSub.assetno.SetFocus
5170            DoEvents
5180            lngTmp03 = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
5190          End Select
5200          .ToTaxLot = 0&
5210          .TaxLotFrom = vbNullString
5220          .JrnlCol_ID = 0&
              ' **    0 OK.
              ' **   -1 Input missing.
              ' **   -2 No holdings.
              ' **   -3 Insufficient holdings.
              ' **   -9 Data problem.

5230        ElseIf gstrReturningForm = "frmMap_Div" Or gstrReturningForm = "frmMap_Div_Detail" Or _
                gstrReturningForm = "frmMap_Int" Or gstrReturningForm = "frmMap_Int_Detail" Or _
                gstrReturningForm = "frmMap_Rec" Or gstrReturningForm = "frmMap_Rec_Detail" Or _
                gstrReturningForm = "frmMap_Split" Or gstrReturningForm = "frmMap_Split_Detail" Or _
                gstrReturningForm = "frmMap_Misc_LTCL" Or gstrReturningForm = "frmMap_Misc_LTCL_Detail" Or _
                gstrReturningForm = "frmMap_Misc_STCGL" Or gstrReturningForm = "frmMap_Misc_STCGL_Detail" Or _
                gstrReturningForm = "frmMap_Reinvest_DivInt_Detail" Or gstrReturningForm = "frmMap_Reinvest_Rec_Detail" Then
              ' ** Should only get here if they were canceled.
5240          .FocusHolder2.SetFocus
5250          DoEvents
5260          strControlName = vbNullString
5270          Select Case gstrReturningForm
              Case "frmMap_Div", "frmMap_Div_Detail"
5280            .cmdSpecPurp_Div_Map.SetFocus
5290            strControlName = "cmdSpecPurp_Div_Map"
5300          Case "frmMap_Int", "frmMap_Int_Detail"
5310            .cmdSpecPurp_Int_Map.SetFocus
5320            strControlName = "cmdSpecPurp_Int_Map"
5330          Case "frmMap_Rec", "frmMap_Rec_Detail", "frmMap_Reinvest_Rec_Detail"
5340            .cmdSpecPurp_Misc_MapLTCG.SetFocus
5350            strControlName = "cmdSpecPurp_Misc_MapLTCG"
5360          Case "frmMap_Split", "frmMap_Split_Detail"
5370            .cmdSpecPurp_Purch_MapSplit.SetFocus
5380            strControlName = "cmdSpecPurp_Purch_MapSplit"
5390          Case "frmMap_Misc_LTCL", "frmMap_Misc_LTCL_Detail"
5400            .cmdSpecPurp_Misc_MapLTCL.SetFocus
5410            strControlName = "cmdSpecPurp_Misc_MapLTCL"
5420          Case "frmMap_Misc_STCGL", "frmMap_Misc_STCGL_Detail"
5430            .cmdSpecPurp_Misc_MapSTCGL.SetFocus
5440            strControlName = "cmdSpecPurp_Misc_MapSTCGL"
5450          Case "frmMap_Reinvest_DivInt_Detail"
5460            Select Case gstrCrtRpt_NetLoss
                Case "Dividend"
5470              .cmdSpecPurp_Div_Map.SetFocus
5480              strControlName = "cmdSpecPurp_Div_Map"
5490            Case "Interest"
5500              .cmdSpecPurp_Int_Map.SetFocus
5510              strControlName = "cmdSpecPurp_Int_Map"
5520            End Select
5530          End Select
5540          gstrCrtRpt_NetLoss = vbNullString
5550          gstrReturningForm = vbNullString
5560          DoCmd.Hourglass False
5570          DoEvents

5580        ElseIf gstrReturningForm <> "frmMap_Return" And gstrReturningForm <> "frmAssets" Then
              ' ** This should hit OnOpen, and any time other than
              ' ** a return from the Tax Lot form or a Map cmdOK.

5590          .frmJournal_Columns_Sub.SetFocus
5600          .Controls(THIS_NAME_SUB).SetFocus
5610          .Controls(THIS_NAME_SUB).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
              ' ** OnOpen, this is the first call to JC_Key_Sub_Next(), and blnFromZero (blnD2) is False.
              ' ** blnNextRec (blnD1) and blnFromZero (blnD2) are frmJournal_Columns_Sub variables.
5620          strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
              ' ** Form_Timer() -> JC_Key_Sub_Next() -> AddRec() -> JC_Key_Sub_Next() -> AddRec()
              ' ** So, blnFromZero needs to get to AddRec(), so it doesn't call JC_Key_Sub_Next() again!
5630          .Controls(THIS_NAME_SUB).Form.Controls(strNextSub).SetFocus
              ' ** On a completely empty record (save transdate),
              ' ** I want to keep the top buttons active!

5640          frmSub.RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.

5650        ElseIf gstrReturningForm = "frmMap_Return" Then
5660          gstrReturningForm = vbNullString
5670          blnNextRec = frmSub.NextRec_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
5680          blnFromZero = frmSub.FromZero_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
5690          JC_Frm_Map_Return lngJrnlColID_Max, blnNextRec, blnFromZero  ' ** Module Procedure: modJrnlCol_Forms.
5700          frmSub.RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.
5710        End If  ' ** ToTaxLot.

5720      ElseIf gstrReturningForm = "frmMap_Return" Then
5730        gstrReturningForm = vbNullString
5740        blnNextRec = frmSub.NextRec_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
5750        blnFromZero = frmSub.FromZero_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
5760        JC_Frm_Map_Return lngJrnlColID_Max, blnNextRec, blnFromZero  ' ** Module Procedure: modJrnlCol_Forms.
5770        frmSub.RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.
5780      End If  ' ** blnSetFocus2.
5790      DoEvents

5800      If gblnGoToReport = True Then
5810        DoCmd.Hourglass True  ' ** Make sure it's still running.
5820        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48& And lngGTR_Stat < 96&) Or (lngGTR_Stat > 96&)) Then
              ' ** Just let the emblem go.
5830        Else
5840          DoCmd.Hourglass True  ' ** Make sure it's still running.
5850          If blnGTR_Emblem = False And IsNull(garr_varGoToReport(GTR_CTL2)) = False Then
5860            blnGTR_Emblem = True
5870            If IsNull(garr_varGoToReport(GTR_CTL2)) = False Then
5880              lngTmp01 = 0&: lngTmp02 = 0&
5890              lngTmp03 = (.GoToReport_Emblem_01_img.Width + (8& * lngTpp))
5900              .cmdRefresh.Visible = False
5910              .Footer_vline03.Visible = False
5920              .Footer_vline04.Visible = False
5930              .Footer_vline05.Left = (.Footer_vline05.Left - lngTmp03)
5940              .Footer_vline06.Left = (.Footer_vline06.Left - lngTmp03)
5950              .cmdAdd_box.Left = (.cmdAdd_box.Left - lngTmp03)
5960              .cmdAdd.Left = (.cmdAdd.Left - lngTmp03)
5970              .cmdEdit.Left = (.cmdEdit.Left - lngTmp03)
5980              .cmdDelete.Left = (.cmdDelete.Left - lngTmp03)
5990              .Footer_vline07.Left = (.Footer_vline07.Left - lngTmp03)
6000              .Footer_vline08.Left = (.Footer_vline08.Left - lngTmp03)
6010              .cmdClose.Left = (.cmdClose.Left - lngTmp03)
6020              .cmdClose.Visible = True  ' ** Turned it off temporarily in Form_Open(), above.
6030              .Sizable_lbl1.Visible = False
6040              .Sizable_lbl2.Visible = False
6050            End If
6060          End If
6070          DoEvents
6080          Select Case blnGoingToReport2
              Case True
                ' ** Just Cost.
6090            If blnGTR_Emblem = False Then
6100              DoCmd.Hourglass False
6110              blnGoingToReport = False
6120              blnGoingToReport2 = False
6130              gblnGoToReport = False
6140              GTRStuff 4, False  ' ** Procedure: Below.
6150              Beep
6160            End If
6170          Case False
6180            strControlName = vbNullString
6190            Select Case blnGoingToReport
                Case True
6200              Select Case IsNull(garr_varGoToReport(GTR_CTL2))
                  Case True
6210                gblnGoToReport = False
6220                blnGoingToReport = False
6230                .GoToReport_arw_pr_img.Visible = False
6240                .cmdClose.SetFocus
6250                .RecsTot_Committed.Locked = True
6260                .RecsTot_Committed.BorderColor = CLR_LTBLU2
6270                .RecsTot_Uncommitted.Locked = True
6280                .RecsTot_Uncommitted.BorderColor = CLR_LTBLU2
6290                DoCmd.Hourglass False
6300                DoEvents
6310              Case False
6320                Select Case garr_varGoToReport(GTR_CTL2)
                    Case "cmdSpecPurp_Div_Map"
6330                  If garr_varGoToReport(GTR_INV2) = True Then
                        ' ** Borrowing this from Court Reports.
6340                    gstrCrtRpt_NetLoss = "Dividend"
6350                    .TimerInterval = 50&
6360                    cmdSpecPurp_Div_Map_Click  ' ** Procedure: Below.
6370                  End If
6380                Case "cmdSpecPurp_Int_Map"
6390                  If garr_varGoToReport(GTR_INV2) = True Then
6400                    gstrCrtRpt_NetLoss = "Interest"
6410                    .TimerInterval = 50&
6420                    cmdSpecPurp_Int_Map_Click  ' ** Procedure: Below.
6430                  End If
6440                Case "cmdSpecPurp_Misc_MapLTCG"
6450                  If garr_varGoToReport(GTR_INV2) = True Then
6460                    gstrCrtRpt_NetLoss = "Received"
6470                    .TimerInterval = 50&
6480                    cmdSpecPurp_Misc_MapLTCG_Click  ' ** Procedure: Below.
6490                  End If
6500                Case "cmdSpecPurp_Purch_MapSplit"
6510                  If garr_varGoToReport(GTR_INV2) = True Then
6520                    gstrCrtRpt_NetLoss = "Split"
6530                    .TimerInterval = 50&
6540                    cmdSpecPurp_Purch_MapSplit_Click  ' ** Procedure: Below.
6550                  End If
6560                Case "cmdSpecPurp_Misc_MapLTCL"
6570                  If garr_varGoToReport(GTR_INV2) = True Then
6580                    gstrCrtRpt_NetLoss = "LTCL"
6590                    .TimerInterval = 50&
6600                    cmdSpecPurp_Misc_MapLTCL_Click  ' ** Procedure: Below.
6610                  End If
6620                Case "cmdSpecPurp_Misc_MapSTCGL"
6630                  If garr_varGoToReport(GTR_INV2) = True Then
6640                    If InStr(garr_varGoToReport(GTR_RNAM), "STCG") > 0 Then
6650                      gstrCrtRpt_NetLoss = "STCG"
6660                    ElseIf InStr(garr_varGoToReport(GTR_RNAM), "STCL") > 0 Then
6670                      gstrCrtRpt_NetLoss = "STCL"
6680                    End If
6690                    .TimerInterval = 50&
6700                    cmdSpecPurp_Misc_MapSTCGL_Click  ' ** Procedure: Below.
6710                  End If
6720                Case "cost"
                      ' ** List of Holdings by Tax Lot.
6730                  .frmJournal_Columns_Sub.SetFocus
6740                  frmSub.accountno.SetFocus
6750                  If garr_varGoToReport(GTR_INV2) = True Or garr_varGoToReport(GTR_FOC2) = True Then
                        ' ** Here we turn it over to frmJournal_Columns_Sub.
                        ' ** See GTREmblem_Set() and GTREmblem_Get(), below.
6760                    blnGoingToReport2 = True
6770                    frmSub.TimerInterval = 50&
6780                    .TimerInterval = 50&
6790                  End If
6800                End Select
6810                DoEvents
6820              End Select
6830            Case False
6840              blnGoingToReport = True
6850              DoCmd.Hourglass True  ' ** Make sure it's still running.
6860              DoEvents
6870              Select Case IsNull(garr_varGoToReport(GTR_CTL2))
                  Case True
                    ' ** Posting Journal report.
6880                .frmJournal_Columns_Sub.SetFocus
6890                frmSub.transdate.SetFocus
6900                frmSub.transdate.SelLength = 0
6910                frmSub.transdate.SelStart = 11
6920                DoCmd.SelectObject acForm, THIS_NAME, False
6930                .RecsTot_Committed.Locked = False  ' ** To Make it look disabled.
6940                .RecsTot_Committed.BorderColor = WIN_CLR_DISR
6950                .RecsTot_Uncommitted.Locked = False
6960                .RecsTot_Uncommitted.BorderColor = WIN_CLR_DISR
6970                .GoToReport_arw_pr_img.Top = (.cmdPrintReport.Top + (3& * lngTpp))
6980                .GoToReport_arw_pr_img.Left = ((.cmdPrintReport.Left + .cmdPrintReport.Width) - lngTpp)
6990                .GoToReport_arw_pr_img.Visible = True
7000                If .cmdPrintReport.Visible = False Then
                      ' ** May be on a Paid with Check Memo.
                      '.frmJournal_Columns_Sub.Form.AddRec  ' ** Form Procedure: frmJournal_Columns_Sub.
7010                  frmSub.AddRec_Send  ' ** Form Procedure: frmJournal_Columns_Sub.
7020                  DoEvents
7030                End If
7040                If .cmdPrintReport.Enabled = True Then
7050                  .cmdPrintReport.SetFocus
7060                End If
7070                Beep
7080                DoEvents
7090                .TimerInterval = (GTR_WAIT * 2&)
7100              Case False
                    ' ** Make sure we're dealing with the default configuration.
7110                If .opgFilter <> .opgFilter_optAll.OptionValue Then
7120                  .opgFilter = .opgFilter_optAll.OptionValue
7130                  opgFilter_AfterUpdate  ' ** Procedure: Below.
7140                  DoEvents
7150                End If
7160                If .opgEnterKey <> .opgEnterKey_optRight.OptionValue Then
7170                  .opgEnterKey = .opgEnterKey_optRight.OptionValue
7180                  opgEnterKey_AfterUpdate  ' ** Procedure: Below.
7190                  DoEvents
7200                End If
7210                Select Case garr_varGoToReport(GTR_CTL2)
                    Case "cmdSpecPurp_Div_Map"
7220                  lngBtn_Top = .Header_lbl.Top
7230                  lngBtn_Left = ((.Controls(garr_varGoToReport(GTR_CTL2)).Left - .GoToReport_arw_mapdiv_img.Width) + (1& * lngTpp))
7240                  .GoToReport_arw_mapdiv_img.Top = lngBtn_Top
7250                  .GoToReport_arw_mapdiv_img.Left = lngBtn_Left
7260                  .GoToReport_arw_mapdiv_img.Visible = True
7270                  .GoToReport_lbl_mapdiv_img.Visible = True
7280                  .Header_lbl.Visible = False
7290                  .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
7300                  DoEvents
7310                Case "cmdSpecPurp_Int_Map"
7320                  lngBtn_Top = (.Controls(garr_varGoToReport(GTR_CTL2)).Top - (3& * lngTpp))
7330                  lngBtn_Left = ((.Controls(garr_varGoToReport(GTR_CTL2)).Left - .GoToReport_arw_mapint_img.Width) + (1& * lngTpp))
7340                  .GoToReport_arw_mapint_img.Top = lngBtn_Top
7350                  .GoToReport_arw_mapint_img.Left = lngBtn_Left
7360                  .GoToReport_arw_mapint_img.Visible = True
7370                  .cmdSpecPurp_Div_Map.Visible = False
7380                  .cmdSpecPurp_Div_Map_raised_img.Visible = False
7390                  .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
7400                  DoEvents
7410                Case "cmdSpecPurp_Misc_MapLTCG"
7420                  lngBtn_Top = (.Controls(garr_varGoToReport(GTR_CTL2)).Top - (3& * lngTpp))
7430                  lngBtn_Left = ((.Controls(garr_varGoToReport(GTR_CTL2)).Left - .GoToReport_arw_maprec_img.Width) + (1& * lngTpp))
7440                  .GoToReport_arw_maprec_img.Top = lngBtn_Top
7450                  .GoToReport_arw_maprec_img.Left = lngBtn_Left
7460                  .GoToReport_arw_maprec_img.Visible = True
7470                  .cmdSpecPurp_Int_Map.Visible = False
7480                  .cmdSpecPurp_Int_Map_raised_img.Visible = False
7490                  .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
7500                  DoEvents
7510                Case "cmdSpecPurp_Purch_MapSplit"
7520                  lngBtn_Top = (.Controls(garr_varGoToReport(GTR_CTL2)).Top - (3& * lngTpp))
7530                  lngBtn_Left = ((.Controls(garr_varGoToReport(GTR_CTL2)).Left - .GoToReport_arw_mapsplit_img.Width) + (1& * lngTpp))
7540                  Select Case gblnSpecialCapGainLoss
                      Case True
7550                    .GoToReport_arw_mapsplit_img.Top = lngBtn_Top
7560                    .GoToReport_arw_mapsplit_img.Left = lngBtn_Left
7570                    .GoToReport_arw_mapsplit_img.Visible = True
7580                    .cmdSpecPurp_Misc_MapSTCGL.Visible = False
7590                    .cmdSpecPurp_Misc_MapSTCGL_raised_img.Visible = False
7600                    .cmdSpecPurp_Misc_MapSTCGL_raised_semifocus_dots_img.Visible = False
7610                    .cmdSpecPurp_Misc_MapSTCGL_raised_focus_img.Visible = False
7620                    .cmdSpecPurp_Misc_MapSTCGL_raised_focus_dots_img.Visible = False
7630                    .cmdSpecPurp_Misc_MapSTCGL_sunken_focus_dots_img.Visible = False
7640                    .cmdSpecPurp_Misc_MapSTCGL_raised_img_dis.Visible = False
7650                  Case False
7660                    .GoToReport_arw_mapsplit_nosc_img.Top = lngBtn_Top
7670                    .GoToReport_arw_mapsplit_nosc_img.Left = lngBtn_Left
7680                    .GoToReport_arw_mapsplit_nosc_img.Visible = True
7690                    .cmdSpecPurp_Misc_MapLTCG.Visible = False
7700                    .cmdSpecPurp_Misc_MapLTCG_raised_img.Visible = False
7710                  End Select
7720                  .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
7730                  DoEvents
7740                Case "cmdSpecPurp_Misc_MapLTCL"
7750                  lngBtn_Top = (.Controls(garr_varGoToReport(GTR_CTL2)).Top - (3& * lngTpp))
7760                  lngBtn_Left = ((.Controls(garr_varGoToReport(GTR_CTL2)).Left - .GoToReport_arw_mapltcl_img.Width) + (1& * lngTpp))
7770                  .GoToReport_arw_mapltcl_img.Top = lngBtn_Top
7780                  .GoToReport_arw_mapltcl_img.Left = lngBtn_Left
7790                  .GoToReport_arw_mapltcl_img.Visible = True
7800                  .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
7810                  DoEvents
7820                Case "cmdSpecPurp_Misc_MapSTCGL"
7830                  lngBtn_Top = (.Controls(garr_varGoToReport(GTR_CTL2)).Top - (3& * lngTpp))
7840                  lngBtn_Left = ((.Controls(garr_varGoToReport(GTR_CTL2)).Left - .GoToReport_arw_mapstcgl_img.Width) + (1& * lngTpp))
7850                  .cmdSpecPurp_Misc_MapLTCL.Visible = False
7860                  .cmdSpecPurp_Misc_MapLTCL_raised_img.Visible = False
7870                  .cmdSpecPurp_Misc_MapLTCL_raised_semifocus_dots_img.Visible = False
7880                  .cmdSpecPurp_Misc_MapLTCL_raised_focus_img.Visible = False
7890                  .cmdSpecPurp_Misc_MapLTCL_raised_focus_dots_img.Visible = False
7900                  .cmdSpecPurp_Misc_MapLTCL_sunken_focus_dots_img.Visible = False
7910                  .cmdSpecPurp_Misc_MapLTCL_raised_img_dis.Visible = False
7920                  .GoToReport_arw_mapstcgl_img.Top = lngBtn_Top
7930                  .GoToReport_arw_mapstcgl_img.Left = lngBtn_Left
7940                  .GoToReport_arw_mapstcgl_img.Visible = True
7950                  .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
7960                  DoEvents
7970                Case "cost"
7980                  lngJrnlColID = GetJrnlColID(frmSub.Filter)  ' ** Module Function: modGoToReportFuncs.
7990                  If lngJrnlColID = 0& Then
                        '.frmJournal_Columns_Sub.Form.AddRec  ' ** Form Procedure: frmJournal_Columns_Sub.
8000                    frmSub.AddRec_Send  ' ** Form Procedure: frmJournal_Columns_Sub.
8010                    DoEvents
8020                    lngJrnlColID = GetJrnlColID(frmSub.Filter)  ' ** Module Function: modGoToReportFuncs.
8030                  End If
8040                  If lngJrnlColID > 0& Then
8050                    If frmSub.JrnlCol_ID <> lngJrnlColID Then
8060                      frmSub.MoveRec 0, lngJrnlColID  ' ** Form Procedure: frmJournal_Columns_Sub.
8070                      DoEvents
8080                    End If
8090                    .frmJournal_Columns_Sub.SetFocus
8100                    frmSub.transdate.SetFocus
8110                    frmSub.transdate.SelLength = 0
8120                    frmSub.transdate.SelStart = 19
8130                    frmSub.GoToReport_arw_jcol_bge_img.Left = frmSub.accountno.Left
8140                    frmSub.GoToReport_arw_jcol_blu1_img.Left = frmSub.GoToReport_arw_jcol_bge_img.Left
8150                    frmSub.GoToReport_arw_jcol_blu2_img.Left = frmSub.GoToReport_arw_jcol_bge_img.Left
8160                    frmSub.GoToReport_arw_jcol_bge_img.Visible = True
8170                    DoEvents
8180                  Else
8190                    .TimerInterval = 0&
8200                    blnGoingToReport = False
8210                    gblnGoToReport = False
8220                    blnGTR_Emblem = False
8230                    GTRStuff 4, False  ' ** Procedure: Below.
8240                    Beep
8250                    DoCmd.Hourglass False
8260                    MsgBox "A new, empty record was not found.", vbCritical + vbOKOnly, "Report Location Unavailable"
8270                  End If
8280                End Select
8290                DoEvents
8300                .TimerInterval = 50&
8310              End Select
8320            End Select
8330          End Select  ' ** gblnGoToReport2.
8340        End If
8350      End If  ' ** gblnGoToReport

          ' ** I'm trying to get the focus back up here on the button it left with,
          ' ** but it seems to be competing with the subform, or something!
8360      Select Case blnSetFocus3
          Case True
8370        lngTmp04 = lngTmp04 + 1&
8380        If strControlName <> vbNullString Then
8390          DoCmd.SelectObject acForm, THIS_NAME, False
8400          .Controls(strControlName).SetFocus
8410          Select Case strControlName
              Case "cmdSpecPurp_Div_Map"
8420            cmdSpecPurp_Div_Map_GotFocus  ' ** Procedure: Below.
8430          Case "cmdSpecPurp_Int_Map"
8440            cmdSpecPurp_Int_Map_GotFocus  ' ** Procedure: Below.
8450          Case "cmdSpecPurp_Misc_MapLTCG"
8460            cmdSpecPurp_Misc_MapLTCG_GotFocus  ' ** Procedure: Below.
8470          Case "cmdSpecPurp_Purch_MapSplit"
8480            cmdSpecPurp_Purch_MapSplit_GotFocus  ' ** Procedure: Below.
8490          Case "cmdSpecPurp_Misc_MapLTCL"
8500            cmdSpecPurp_Misc_MapLTCL_GotFocus  ' ** Procedure: Below.
8510          Case "cmdSpecPurp_Misc_MapSTCGL"
8520            cmdSpecPurp_Misc_MapSTCGL_GotFocus  ' ** Procedure: Below.
8530          End Select
8540          .TimerInterval = 250&
8550        End If
8560        If lngTmp04 >= 2& Then
8570          If strControlName <> vbNullString Then
8580            blnSetFocus3 = False
8590            DoCmd.SelectObject acForm, THIS_NAME, False
8600            .Controls(strControlName).SetFocus
8610            Select Case strControlName
                Case "cmdSpecPurp_Div_Map"
8620              cmdSpecPurp_Div_Map_GotFocus  ' ** Procedure: Below.
8630            Case "cmdSpecPurp_Int_Map"
8640              cmdSpecPurp_Int_Map_GotFocus  ' ** Procedure: Below.
8650            Case "cmdSpecPurp_Misc_MapLTCG"
8660              cmdSpecPurp_Misc_MapLTCG_GotFocus  ' ** Procedure: Below.
8670            Case "cmdSpecPurp_Purch_MapSplit"
8680              cmdSpecPurp_Purch_MapSplit_GotFocus  ' ** Procedure: Below.
8690            Case "cmdSpecPurp_Misc_MapLTCL"
8700              cmdSpecPurp_Misc_MapLTCL_GotFocus  ' ** Procedure: Below.
8710            Case "cmdSpecPurp_Misc_MapSTCGL"
8720              cmdSpecPurp_Misc_MapSTCGL_GotFocus  ' ** Procedure: Below.
8730            End Select
8740            .TimerInterval = 0&
8750          End If
8760        End If
8770      Case False
8780        If strControlName <> vbNullString Then
8790          .Controls(strControlName).SetFocus
8800          Select Case strControlName
              Case "cmdSpecPurp_Div_Map"
8810            cmdSpecPurp_Div_Map_GotFocus  ' ** Procedure: Below.
8820          Case "cmdSpecPurp_Int_Map"
8830            cmdSpecPurp_Int_Map_GotFocus  ' ** Procedure: Below.
8840          Case "cmdSpecPurp_Misc_MapLTCG"
8850            cmdSpecPurp_Misc_MapLTCG_GotFocus  ' ** Procedure: Below.
8860          Case "cmdSpecPurp_Purch_MapSplit"
8870            cmdSpecPurp_Purch_MapSplit_GotFocus  ' ** Procedure: Below.
8880          Case "cmdSpecPurp_Misc_MapLTCL"
8890            cmdSpecPurp_Misc_MapLTCL_GotFocus  ' ** Procedure: Below.
8900          Case "cmdSpecPurp_Misc_MapSTCGL"
8910            cmdSpecPurp_Misc_MapSTCGL_GotFocus  ' ** Procedure: Below.
8920          End Select
8930          blnSetFocus3 = True
8940          lngTmp04 = 0&
8950          .TimerInterval = 250&
8960        End If
8970      End Select

8980    End With

EXITP:
8990    Set frmSub = Nothing
9000    Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     Application.Echo True
1000    THAT_PROC = THIS_PROC
1010    That_Erl = Erl: That_Desc = ERR.description
1020    If ERR.Number > 0 Then
1030      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1040    Else  ' ** Those huge negative error numbers are too big for the Integer in the Form_Error() parameters.
1050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1060    End If
1070    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim strTmp01 As String, dblTmp02 As Double, lngTmp03 As Long

9110    Application.Echo True
9120    If DataErr <> 0 Then
9130      dblTmp02 = DataErr
9140      lngTmp03 = Erl
9150      Select Case DataErr
          Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
            ' ** Do nothing.
9160      Case 2237  ' ** The text you entered isn't an item in the list.
9170        MsgBox "You must select a valid item from the list.", vbInformation + vbOKOnly, "Invalid Entry"
9180        Response = acDataErrContinue
9190      Case 2427  ' ** You entered an expression that has no value.
            ' ** A reference to an empty subform.
9200        Response = acDataErrContinue
9210      Case 3020  ' ** Update or CancelUpdate without AddNew or Edit.
            ' ** Always totally bizarre!
9220        Me.Undo
9230        Response = acDataErrContinue
9240      Case Else
9250        strTmp01 = AccessError(DataErr)
9260        zErrorHandler THIS_NAME, THIS_PROC, CVar(DataErr), 0&, strTmp01
9270      End Select
9280    End If

EXITP:
9290    Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String
        Dim lngX As Long

9310    With Me

9320      .TimerInterval = 0&  ' ** Make sure!

9330      DoCmd.Hourglass True
9340      DoEvents

          ' ** Save LastAcctNo, and others.
9350      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

9360      gblnGoToReport = False  ' ** Make sure this is off.

          ' ** Empty all the arrays.
9370      lngStats = 0&
9380      arr_varStat = Empty
9390      .frmJournal_Columns_Sub.Form.JrnlCol_Sub_Clear  ' ** Form Procedure: frmJournal_Columns_Sub.
9400      JC_Btn_Clear  ' ** Module Procedure: modJrnlCol_Buttons.
9410      JC_Key_Clear  ' ** Module Procedure: modJrnlCol_Keys.
9420      JC_Msc_Clear  ' ** Module Procedure: modJrnlCol_Misc.
9430      JC_Frm_Clear  ' ** Module Procedure: modJrnlCol_Forms.

9440      If .cmdSpecPurp_Sold_PaidTotal.Enabled = False And .cmdSpecPurp_Div_Map.Enabled = True Then
9450        .cmdSpecPurp_Sold_lbl.ForeColor = CLR_BLUGRY  ' ** Just my obsessiveness
9460        .cmdSpecPurp_Sold_lbl_dim_hi.Visible = False
9470        .cmdSpecPurp_Sold_PaidTotal_raised_img.Visible = True
9480        DoEvents
9490      End If

          ' ** At least once, this caused a crash.
9500      If .opgFilter <> .opgFilter_optAll.OptionValue Then
9510        .opgFilter = .opgFilter_optAll.OptionValue
9520        DoEvents
9530      End If

          ' ** Make sure any secondary windows are closed.
9540      For lngX = (Forms.Count - 1) To 0& Step -1&
9550        strDocName = Forms(lngX).Name
9560        If strDocName <> "frmJournal_Columns" Then
9570          If gblnSwitchTo = False And strDocName = "frmMenu_Post" Then
                ' ** Leave this open if returning to it.
9580          ElseIf strDocName = "frmMenu_Background" Then
                ' ** Leave this open.
9590          Else
9600            DoCmd.Close acForm, strDocName
9610          End If
9620        End If
9630      Next

          ' ** Make sure all reports are closed.
9640      Do While Reports.Count > 0
9650        DoCmd.Close acReport, Reports(0).Name
9660      Loop

9670      Set dbs = CurrentDb
9680      With dbs
              ' ** Delete qryJournal_Columns_02j (tblJournal_Column, just Journal_ID > 0 and Posted = False) from Journal.
9690          Set qdf = dbs.QueryDefs("qryJournal_Columns_02k")
9700          qdf.Execute
9710          Set qdf = Nothing
9720          DoEvents
              ' ** Delete tblJournalColumn, for posted = True.
9730          Set qdf = dbs.QueryDefs("qryJournal_Columns_02b")
9740          qdf.Execute
9750          Set qdf = Nothing
9760          DoEvents
              ' ** Delete qryJournal_Columns_02f (tblJournal_Column, just totally empty records).
9770          Set qdf = dbs.QueryDefs("qryJournal_Columns_02g")
9780          qdf.Execute
9790          Set qdf = Nothing
9800          DoEvents
              ' ** Empty tmpAccount.
9810          Set qdf = dbs.QueryDefs("qryJournal_Columns_10_Account_04")
9820          qdf.Execute
9830          DoEvents
9840        Set qdf = Nothing
9850      End With

9860      Select Case gblnSwitchTo
          Case True
            ' ** Since they're switching, delete their preference.
9870        With dbs
              ' ** Delete qryPreferences_16_01 (tblPreference_User, by specified [frmnam], [ctlnam], [usr]).
9880          Set qdf = .QueryDefs("qryPreferences_16_02")  '##dbs_id
9890          With qdf.Parameters
9900            ![frmnam] = "frmPostingDate"
9910            ![ctlnam] = "opgInput"
9920            ![usr] = CurrentUser ' ** Internal Access Function: Trust Accountant login.
9930          End With
9940          qdf.Execute
9950          Set qdf = Nothing
9960        End With
9970        glngJournalForm = JRN_DIV
9980      Case False
9990        If blnFromSwitch = True Then
10000         With dbs
                ' ** tblPreference_user, linked to tblPreference_Control,
                ' ** just 'opgInput' on 'frmPostingDate', by specified [usr].
10010           Set qdf = .QueryDefs("qryPreferences_09_01")
10020           With qdf.Parameters
10030             ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10040           End With
10050           Set rst = qdf.OpenRecordset
10060           If rst.BOF = True And rst.EOF = True Then
                  ' ** No preference (which may not ever happen).
10070             rst.Close
10080             Set rst = Nothing
10090             Set qdf = Nothing
                    ' ** Append qryPreferences_09_06 (qryPreferences_09_05 (tblPreference_Control,
                    ' ** just 'opgInput' on 'frmPostingDate', with ctl_name_opt, ctl_name_short),
                    ' ** as new tblPreference_User record, by specified [usr], [opt]) to tblPreference_User.
10100               Set qdf = .QueryDefs("qryPreferences_09_07")
10110               With qdf.Parameters
10120                 ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10130                 ![opt] = "Columnar"
10140               End With
10150               qdf.Execute
10160             Set qdf = Nothing
10170           Else
10180             rst.Close
10190             Set rst = Nothing
10200             Set qdf = Nothing
                    ' ** Update qryPreferences_09_03 (qryPreferences_09_01 (tblPreference_user,
                    ' ** linked to tblPreference_Control, just 'opgInput' on 'frmPostingDate',
                    ' ** by specified [usr]), with DLookups() to qryPreferences_09_02 (tblForm_Control,
                    ' ** just 'opgInput_optClassic', 'opgInput_optColumnar', with ctl_name_short),
                    ' ** with prefuser_integer_new, by specified [opt]).
10210               Set qdf = .QueryDefs("qryPreferences_09_04")
10220               With qdf.Parameters
10230                 ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10240                 ![opt] = "Columnar"
10250               End With
10260               qdf.Execute
10270             Set qdf = Nothing
10280           End If
10290         End With
10300       End If  ' ** blnFromSwitch.
10310     End Select

10320     dbs.Close

10330   End With

EXITP:
10340   Set rst = Nothing
10350   Set qdf = Nothing
10360   Set dbs = Nothing
10370   Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2501  ' ** The Close action was canceled.
1020    Resume EXITP

End Sub

Private Sub Form_Close()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Close"

        Dim strDocName As String

10410   With Me
10420     .TimerInterval = 0&  ' ** Make sure!

10430     If IsMaximized(Me) = True Then
10440       DoCmd.Restore
10450     End If

10460     gblnClosing = False: gblnDeleting = False
10470     gstrFormQuerySpec = vbNullString
10480     gblnSetFocus = True
10490     gstrReturningForm = THIS_NAME

10500     DoCmd.Hourglass False

10510     Select Case gblnSwitchTo
          Case True
10520       strDocName = "frmJournal"
10530       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
10540     Case False
10550       strDocName = "frmMenu_Background"
10560       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
10570     End Select

10580   End With

EXITP:
10590   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2501  ' ** The Close action was canceled.
1010    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

10610   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
10620   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

10710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10720   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

10810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10820   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub Header_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl_MouseMove"

10910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10920   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdClose_GotFocus()

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_GotFocus"

11010   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
11020   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdClose_Click()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

11110   If blnCommitNoClose = False Then

11120     gblnClosing = True
11130     DoEvents

11140     DoCmd.Close acForm, THIS_NAME

11150   Else
11160     Beep
11170   End If

EXITP:
11180   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2501  ' ** The Close action was canceled.
1010    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

11210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11220   intShiftDown = (Shift And acShiftMask) > 0
11230   intAltDown = (Shift And acAltMask) > 0
11240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11260     Select Case intRetVal
          Case vbKeyTab
11270       With Me
11280         intRetVal = 0
11290         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
11300         If strNext = (THIS_NAME_SUB) Then
11310           .Controls(strNext).SetFocus
11320           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
11330           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
11340           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
11350           .Controls(strNext).Form.Controls(strNextSub).SetFocus
11360         Else
11370           .Controls(strNext).SetFocus
11380         End If
11390       End With
11400     End Select
11410   End If

        ' ** Shift keys.
11420   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11430     Select Case intRetVal
          Case vbKeyTab
11440       With Me
11450         intRetVal = 0
11460         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
11470         If strNext = (THIS_NAME_SUB) Then
11480           .Controls(strNext).SetFocus
11490           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
11500           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
11510           .Controls(strNext).Form.Controls(strNextSub).SetFocus
11520           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
11530         Else
11540           .Controls(strNext).SetFocus
11550         End If
11560       End With
11570     End Select
11580   End If

EXITP:
11590   KeyCode = intRetVal
11600   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdSwitch_GotFocus()

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSwitch_GotFocus"

11710   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
11720   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
11730   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSwitch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSwitch_MouseDown"

11810   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
11820   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Public Sub cmdSwitch_Click()

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSwitch_Click"

11910   With Me
11920     gblnSwitchTo = True
11930     DoEvents
11940     cmdClose_Click  ' ** Procedure: Above.
11950   End With

EXITP:
11960   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSwitch_KeyDown(KeyCode As Integer, Shift As Integer)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSwitch_KeyDown"

        Dim intRetVal As Integer

12010   intRetVal = JC_Key_SwitchScroll(KeyCode, Shift, THIS_PROC, THIS_NAME_SUB, Me)  ' ** Module Function: modJrnlCol_Keys.

EXITP:
12020   KeyCode = intRetVal
12030   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdSwitch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSwitch_MouseMove"

12110   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
12120   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdSwitch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSwitch_MouseUp"

12210   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
12220   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdSwitch_LostFocus()

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSwitch_LostFocus"

12310   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
12320   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdScrollLeft_GotFocus()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdScrollLeft_GotFocus"

12410   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
12420   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
12430   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdScrollLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdScrollLeft_MouseDown"

12510   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
12520   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdScrollLeft_Click()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdScrollLeft_Click"

        Dim lngRetVal As Long

12610   With Me
12620     lngRetVal = fSetScrollBarPosHZ(.frmJournal_Columns_Sub.Form, 2&)  ' ** Module Function: modScrollBarFuncs.
12630   End With

EXITP:
12640   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdScrollLeft_KeyDown(KeyCode As Integer, Shift As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdScrollLeft_KeyDown"

        Dim intRetVal As Integer

12710   intRetVal = JC_Key_SwitchScroll(KeyCode, Shift, THIS_PROC, THIS_NAME_SUB, Me)  ' ** Module Function: modJrnlCol_Keys.

EXITP:
12720   KeyCode = intRetVal
12730   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1020    Resume EXITP

End Sub

Private Sub cmdScrollLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdScrollLeft_MouseUp"

12810   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
12820   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdScrollLeft_LostFocus()

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdScrollLeft_LostFocus"

12910   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
12920   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdScrollRight_GotFocus()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdScrollRight_GotFocus"

13010   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
13020   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
13030   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdScrollRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdScrollRight_MouseDown"

13110   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
13120   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdScrollRight_Click()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdScrollRight_Click"

        Dim lngRetVal As Long

13210   With Me
13220     lngRetVal = fSetScrollBarPosHZ(.frmJournal_Columns_Sub.Form, 999&)  ' ** Module Function: modScrollBarFuncs.
13230   End With

EXITP:
13240   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdScrollRight_KeyDown(KeyCode As Integer, Shift As Integer)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdScrollRight_KeyDown"

        Dim intRetVal As Integer

13310   intRetVal = JC_Key_SwitchScroll(KeyCode, Shift, THIS_PROC, THIS_NAME_SUB, Me)  ' ** Module Function: modJrnlCol_Keys.

EXITP:
13320   KeyCode = intRetVal
13330   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdScrollRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdScrollRight_MouseUp"

13410   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
13420   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdScrollRight_LostFocus()

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdScrollRight_LostFocus"

13510   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
13520   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub frmJournal_Columns_Sub_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "frmJournal_Columns_Sub_box_MouseMove"

13610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
13620   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub frmJournal_Columns_Sub_Enter()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "frmJournal_Columns_Sub_Enter"

        Dim Button As Integer, Shift As Integer, X As Single, Y As Single

13710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

13720   If blnNotPopup = True Then
13730     SysCmd acSysCmdClearStatus
13740   End If

EXITP:
13750   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

13810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
13820   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Public Sub opgFilter_AfterUpdate()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFilter_AfterUpdate"

13910   JCol_OpgFilter_AfterUpdate THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.

EXITP:
13920   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub opgFilter_optAll_GotFocus()

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFilter_optAll_GotFocus"

14010   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
14020   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub opgFilter_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFilter_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

14110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14120   intShiftDown = (Shift And acShiftMask) > 0
14130   intAltDown = (Shift And acAltMask) > 0
14140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14170       With Me
14180         intRetVal = 0
14190         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
14200         If strNext = (THIS_NAME_SUB) Then
14210           .Controls(strNext).SetFocus
14220           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
14230           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
14240           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
14250           .Controls(strNext).Form.Controls(strNextSub).SetFocus
14260         Else
14270           .Controls(strNext).SetFocus
14280         End If
14290       End With
14300     Case vbKeyLeft
14310       With Me
14320         intRetVal = 0
14330         If .cmdPrintReport.Enabled = True Then
14340           .cmdPrintReport.SetFocus
14350         Else
14360           .cmdClose.SetFocus
14370         End If
14380       End With
14390     Case vbKeyRight
14400       With Me
14410         intRetVal = 0
14420         .opgFilter_optCommitted.SetFocus
14430       End With
14440     End Select
14450   End If

        ' ** Shift keys.
14460   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14470     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14480       With Me
14490         intRetVal = 0
14500         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
14510         If strNext = (THIS_NAME_SUB) Then
14520           .Controls(strNext).SetFocus
14530           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
14540           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
14550           .Controls(strNext).Form.Controls(strNextSub).SetFocus
14560           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
14570         Else
14580           .Controls(strNext).SetFocus
14590         End If
14600       End With
14610     End Select
14620   End If

EXITP:
14630   KeyCode = intRetVal
14640   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub opgFilter_optCommitted_GotFocus()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFilter_optCommitted_GotFocus"

14710   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
14720   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub opgFilter_optCommitted_KeyDown(KeyCode As Integer, Shift As Integer)

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFilter_optCommitted_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

14810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14820   intShiftDown = (Shift And acShiftMask) > 0
14830   intAltDown = (Shift And acAltMask) > 0
14840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14870       With Me
14880         intRetVal = 0
14890         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
14900         If strNext = (THIS_NAME_SUB) Then
14910           .Controls(strNext).SetFocus
14920           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
14930           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
14940           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
14950           .Controls(strNext).Form.Controls(strNextSub).SetFocus
14960         Else
14970           .Controls(strNext).SetFocus
14980         End If
14990       End With
15000     Case vbKeyLeft
15010       With Me
15020         intRetVal = 0
15030         .opgFilter_optAll.SetFocus
15040       End With
15050     Case vbKeyRight
15060       With Me
15070         intRetVal = 0
15080         .opgFilter_optUncommitted.SetFocus
15090       End With
15100     End Select
15110   End If

        ' ** Shift keys.
15120   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15130     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15140       With Me
15150         intRetVal = 0
15160         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
15170         If strNext = (THIS_NAME_SUB) Then
15180           .Controls(strNext).SetFocus
15190           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
15200           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
15210           .Controls(strNext).Form.Controls(strNextSub).SetFocus
15220           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
15230         Else
15240           .Controls(strNext).SetFocus
15250         End If
15260       End With
15270     End Select
15280   End If

EXITP:
15290   KeyCode = intRetVal
15300   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub opgFilter_optUncommitted_GotFocus()

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFilter_optUncommitted_GotFocus"

15410   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
15420   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub opgFilter_optUncommitted_KeyDown(KeyCode As Integer, Shift As Integer)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFilter_optUncommitted_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

15510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15520   intShiftDown = (Shift And acShiftMask) > 0
15530   intAltDown = (Shift And acAltMask) > 0
15540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15570       With Me
15580         intRetVal = 0
15590         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
15600         If strNext = (THIS_NAME_SUB) Then
15610           .Controls(strNext).SetFocus
15620           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
15630           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
15640           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
15650           .Controls(strNext).Form.Controls(strNextSub).SetFocus
15660         Else
15670           .Controls(strNext).SetFocus
15680         End If
15690       End With
15700     Case vbKeyLeft
15710       With Me
15720         intRetVal = 0
15730         .opgFilter_optCommitted.SetFocus
15740       End With
15750     Case vbKeyRight
15760       With Me
15770         intRetVal = 0
15780         .opgEnterKey_optRight.SetFocus
15790       End With
15800     End Select
15810   End If

        ' ** Shift keys.
15820   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15830     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15840       With Me
15850         intRetVal = 0
15860         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
15870         If strNext = (THIS_NAME_SUB) Then
15880           .Controls(strNext).SetFocus
15890           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
15900           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
15910           .Controls(strNext).Form.Controls(strNextSub).SetFocus
15920           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
15930         Else
15940           .Controls(strNext).SetFocus
15950         End If
15960       End With
15970     End Select
15980   End If

EXITP:
15990   KeyCode = intRetVal
16000   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1020    Resume EXITP

End Sub

Private Sub opgEnterKey_AfterUpdate()

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgEnterKey_AfterUpdate"

16110   With Me
16120     Select Case .opgEnterKey
          Case .opgEnterKey_optRight.OptionValue
16130       .opgEnterKey_optRight_lbl.FontBold = True
16140       .opgEnterKey_optDown_lbl.FontBold = False
16150     Case .opgEnterKey_optDown.OptionValue
16160       .opgEnterKey_optRight_lbl.FontBold = False
16170       .opgEnterKey_optDown_lbl.FontBold = True
16180     End Select
16190   End With

EXITP:
16200   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub opgEnterKey_optRight_GotFocus()

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgEnterKey_optRight_GotFocus"

16310   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
16320   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub opgEnterKey_optRight_KeyDown(KeyCode As Integer, Shift As Integer)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgEnterKey_optRight_KeyDown"

        Dim intRetVal As Integer

16410   intRetVal = JC_Key_EnterMemo(KeyCode, Shift, THIS_PROC, THIS_NAME_SUB, Me)  ' ** Module Function: modJrnlCol_Keys.

EXITP:
16420   KeyCode = intRetVal
16430   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub opgEnterKey_optDown_GotFocus()

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgEnterKey_optDown_GotFocus"

16510   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
16520   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub opgEnterKey_optDown_KeyDown(KeyCode As Integer, Shift As Integer)

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgEnterKey_optDown_KeyDown"

        Dim intRetVal As Integer

16610   intRetVal = JC_Key_EnterMemo(KeyCode, Shift, THIS_PROC, THIS_NAME_SUB, Me)  ' ** Module Function: modJrnlCol_Keys.

EXITP:
16620   KeyCode = intRetVal
16630   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdMemoReveal_GotFocus()

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMemoReveal_GotFocus"

16710   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
16720   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdMemoReveal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMemoReveal_MouseDown"

16810   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
16820   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Public Sub cmdMemoReveal_Click()

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMemoReveal_Click"

16910   JC_Ctl_MemoReveal Me  ' ** Module Procedure: modJrnlCol_Controls.

EXITP:
16920   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdMemoReveal_KeyDown(KeyCode As Integer, Shift As Integer)

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMemoReveal_KeyDown"

        Dim intRetVal As Integer

17010   intRetVal = JC_Key_EnterMemo(KeyCode, Shift, THIS_PROC, THIS_NAME_SUB, Me)  ' ** Module Function: modJrnlCol_Keys.

EXITP:
17020   KeyCode = intRetVal
17030   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1020    Resume EXITP

End Sub

Private Sub cmdMemoReveal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMemoReveal_MouseMove"

17110   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
17120   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdMemoReveal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMemoReveal_MouseUp"

17210   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
17220   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdMemoReveal_LostFocus()

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMemoReveal_LostFocus"

17310   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
17320   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdPreviewReport_GotFocus()

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_GotFocus"

17410   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
17420   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
17430   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdPreviewReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_MouseDown"

17510   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
17520   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdPreviewReport_Click()

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_Click"

17610   With Me
17620     blnIsMaximized = IsMaximized(Me)  ' ** Module Function: modWindowFunctions.
17630     JC_Frm_Report .frmJournal_Columns_Sub.Form.RecCnt, acViewPreview  ' ** Module Procedure: modJrnlCol_Forms
17640   End With                                                            ' ** Form Function: frmJournal_Columns_Sub.

EXITP:
17650   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case 2202  ' ** You must set a default printer before you design, print, or preview.
990       MsgBox "There does not appear to be a default printer defined." & vbCrLf & _
            "Please define a default printer before trying to print this report.", _
            vbCritical + vbOKOnly, "Default Printer Required"
1000    Case Else
1010      THAT_PROC = THIS_PROC
1020      That_Erl = Erl: That_Desc = ERR.description
1030      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1040    End Select
1050    Resume EXITP

End Sub

Private Sub cmdPreviewReport_KeyDown(KeyCode As Integer, Shift As Integer)

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

17710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17720   intShiftDown = (Shift And acShiftMask) > 0
17730   intAltDown = (Shift And acAltMask) > 0
17740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17760     Select Case intRetVal
          Case vbKeyTab
17770       With Me
17780         intRetVal = 0
17790         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
17800         If strNext = (THIS_NAME_SUB) Then
17810           .Controls(strNext).SetFocus
17820           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
17830           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
17840           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
17850           .Controls(strNext).Form.Controls(strNextSub).SetFocus
17860         Else
17870           .Controls(strNext).SetFocus
17880         End If
17890       End With
17900     End Select
17910   End If

        ' ** Shift keys.
17920   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17930     Select Case intRetVal
          Case vbKeyTab
17940       With Me
17950         intRetVal = 0
17960         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
17970         If strNext = (THIS_NAME_SUB) Then
17980           .Controls(strNext).SetFocus
17990           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
18000           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
18010           .Controls(strNext).Form.Controls(strNextSub).SetFocus
18020           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
18030         Else
18040           .Controls(strNext).SetFocus
18050         End If
18060       End With
18070     End Select
18080   End If

EXITP:
18090   KeyCode = intRetVal
18100   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdPreviewReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_MouseMove"

18210   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
18220   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdPreviewReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_MouseUp"

18310   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
18320   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdPreviewReport_LostFocus()

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewReport_LostFocus"

18410   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
18420   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

18510   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
18520   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
18530   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

18610   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
18620   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdPrintReport_Click()

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

18710   With Me
18720     blnIsMaximized = IsMaximized(Me)  ' ** Module Function: modWindowFunctions.
          '##GTR_Ref: rptPostingJournal_Column
18730     JC_Frm_Report .frmJournal_Columns_Sub.Form.RecCnt, acViewNormal  ' ** Module Procedure: modJrnlCol_Forms.
18740   End With                                                           ' ** Form Function: frmJournal_Columns_Sub.

EXITP:
18750   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case 2202  ' ** You must set a default printer before you design, print, or preview.
990       MsgBox "There does not appear to be a default printer defined." & vbCrLf & _
            "Please define a default printer before trying to print this report.", _
            vbCritical + vbOKOnly, "Default Printer Required"
1000    Case Else
1010      THAT_PROC = THIS_PROC
1020      That_Erl = Erl: That_Desc = ERR.description
1030      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1040    End Select
1050    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

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
18890         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
18900         If strNext = (THIS_NAME_SUB) Then
18910           .Controls(strNext).SetFocus
18920           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
18930           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
18940           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
18950           .Controls(strNext).Form.Controls(strNextSub).SetFocus
18960         Else
18970           .Controls(strNext).SetFocus
18980         End If
18990       End With
19000     End Select
19010   End If

        ' ** Shift keys.
19020   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19030     Select Case intRetVal
          Case vbKeyTab
19040       With Me
19050         intRetVal = 0
19060         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
19070         If strNext = (THIS_NAME_SUB) Then
19080           .Controls(strNext).SetFocus
19090           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
19100           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
19110           .Controls(strNext).Form.Controls(strNextSub).SetFocus
19120           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
19130         Else
19140           .Controls(strNext).SetFocus
19150         End If
19160       End With
19170     End Select
19180   End If

EXITP:
19190   KeyCode = intRetVal
19200   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

19310   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
19320   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

19410   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
19420   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

19510   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
19520   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUncomComAll_GotFocus()

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomComAll_GotFocus"

19610   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
19620   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
19630   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUncomComAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomComAll_MouseDown"

19710   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
19720   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdUncomComAll_Click()

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomComAll_Click"

        Dim blnRetVal As Boolean

19810   With Me
19820     blnRetVal = CommitAll(False, .frmJournal_Columns_Sub.Form)   ' ** Module Function: modJrnlCol_Recs.
19830   End With

EXITP:
19840   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUncomComAll_KeyDown(KeyCode As Integer, Shift As Integer)

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomComAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

19910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19920   intShiftDown = (Shift And acShiftMask) > 0
19930   intAltDown = (Shift And acAltMask) > 0
19940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19960     Select Case intRetVal
          Case vbKeyTab
19970       With Me
19980         intRetVal = 0
19990         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
20000         If strNext = (THIS_NAME_SUB) Then
20010           .Controls(strNext).SetFocus
20020           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
20030           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
20040           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
20050           .Controls(strNext).Form.Controls(strNextSub).SetFocus
20060         Else
20070           .Controls(strNext).SetFocus
20080         End If
20090       End With
20100     End Select
20110   End If

        ' ** Shift keys.
20120   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20130     Select Case intRetVal
          Case vbKeyTab
20140       With Me
20150         intRetVal = 0
20160         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
20170         If strNext = (THIS_NAME_SUB) Then
20180           .Controls(strNext).SetFocus
20190           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
20200           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
20210           .Controls(strNext).Form.Controls(strNextSub).SetFocus
20220           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
20230         Else
20240           .Controls(strNext).SetFocus
20250         End If
20260       End With
20270     End Select
20280   End If

EXITP:
20290   KeyCode = intRetVal
20300   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdUncomComAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomComAll_MouseMove"

20410   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
20420   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUncomComAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomComAll_MouseUp"

20510   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
20520   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUncomComAll_LostFocus()

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomComAll_LostFocus"

20610   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
20620   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUncomDelAll_GotFocus()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomDelAll_GotFocus"

20710   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
20720   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
20730   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUncomDelAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomDelAll_MouseDown"

20810   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
20820   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdUncomDelAll_Click()

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomDelAll_Click"

20910   JC_Ctl_UncomDelAll Me  ' ** Module Procedure: modJrnlCol_Controls.

EXITP:
20920   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUncomDelAll_KeyDown(KeyCode As Integer, Shift As Integer)

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomDelAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

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
21090         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
21100         If strNext = (THIS_NAME_SUB) Then
21110           .Controls(strNext).SetFocus
21120           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
21130           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
21140           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
21150           .Controls(strNext).Form.Controls(strNextSub).SetFocus
21160         Else
21170           .Controls(strNext).SetFocus
21180         End If
21190       End With
21200     End Select
21210   End If

        ' ** Shift keys.
21220   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21230     Select Case intRetVal
          Case vbKeyTab
21240       With Me
21250         intRetVal = 0
21260         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
21270         If strNext = (THIS_NAME_SUB) Then
21280           .Controls(strNext).SetFocus
21290           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
21300           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
21310           .Controls(strNext).Form.Controls(strNextSub).SetFocus
21320           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
21330         Else
21340           .Controls(strNext).SetFocus
21350         End If
21360       End With
21370     End Select
21380   End If

EXITP:
21390   KeyCode = intRetVal
21400   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdUncomDelAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomDelAll_MouseMove"

21510   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
21520   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUncomDelAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomDelAll_MouseUp"

21610   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
21620   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUncomDelAll_LostFocus()

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUncomDelAll_LostFocus"

21710   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
21720   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUnCommitOne_GotFocus()

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUnCommitOne_GotFocus"

21810   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
21820   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
21830   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUnCommitOne_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUnCommitOne_MouseDown"

21910   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
21920   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUnCommitOne_Click()

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUnCommitOne_Click"

22010   JC_Ctl_UnCommitOne Me  ' ** Module Procedure: modJrnlCol_Controls.

EXITP:
22020   Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1020    Resume EXITP

End Sub

Private Sub cmdUnCommitOne_KeyDown(KeyCode As Integer, Shift As Integer)

22100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUnCommitOne_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String
        Dim intRetVal As Integer

22110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22120   intShiftDown = (Shift And acShiftMask) > 0
22130   intAltDown = (Shift And acAltMask) > 0
22140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22160     Select Case intRetVal
          Case vbKeyTab
22170       With Me
22180         intRetVal = 0
22190         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
22200       End With
22210     End Select
22220   End If

        ' ** Shift keys.
22230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22240     Select Case intRetVal
          Case vbKeyTab
22250       With Me
22260         intRetVal = 0
22270         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
22280       End With
22290     End Select
22300   End If

EXITP:
22310   KeyCode = intRetVal
22320   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1020    Resume EXITP

End Sub

Private Sub cmdUnCommitOne_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUnCommitOne_MouseMove"

22410   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
22420   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUnCommitOne_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUnCommitOne_MouseUp"

22510   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
22520   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdUnCommitOne_LostFocus()

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUnCommitOne_LostFocus"

22610   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
22620   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub JrnlMemo_Memo_GotFocus()

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlMemo_Memo_GotFocus"

22710   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
22720   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub JrnlMemo_Memo_AfterUpdate()

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlMemo_Memo_AfterUpdate"

22810   JCol_Memo_AfterUpdate THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.

EXITP:
22820   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub JrnlMemo_Memo_KeyDown(KeyCode As Integer, Shift As Integer)

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlMemo_Memo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intLen As Integer
        Dim intRetVal As Integer

22910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22920   intShiftDown = (Shift And acShiftMask) > 0
22930   intAltDown = (Shift And acAltMask) > 0
22940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22970       With Me
22980         intRetVal = 0
22990         .frmJournal_Columns_Sub.SetFocus
23000         .frmJournal_Columns_Sub.Form.description.SetFocus
23010       End With
23020     Case Else
23030       With Me
23040         If IsNull(.JrnlMemo_Memo.text) = False Then
23050           If Trim(.JrnlMemo_Memo.text) <> vbNullString Then
23060             intLen = Len(Trim(.JrnlMemo_Memo.text))
23070             If intLen > MEMO_MAX Then
23080               intRetVal = 0
23090               Beep
23100               .JrnlMemo_Memo.text = Left(Trim(.JrnlMemo_Memo.text), MEMO_MAX)
23110               .JrnlMemo_Memo.SelStart = 99
23120             End If
23130           End If
23140         End If
23150       End With
23160     End Select
23170   End If

        ' ** Shift keys.
23180   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23190     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23200       With Me
23210         intRetVal = 0
23220         .frmJournal_Columns_Sub.SetFocus
23230         .frmJournal_Columns_Sub.Form.PrintCheck.SetFocus
23240       End With
23250     Case Else
23260       With Me
23270         If IsNull(.JrnlMemo_Memo.text) = False Then
23280           If Trim(.JrnlMemo_Memo.text) <> vbNullString Then
23290             intLen = Len(Trim(.JrnlMemo_Memo.text))
23300             If intLen > MEMO_MAX Then
23310               intRetVal = 0
23320               Beep
23330               .JrnlMemo_Memo.text = Left(Trim(.JrnlMemo_Memo.text), MEMO_MAX)
23340               .JrnlMemo_Memo.SelStart = 99
23350             End If
23360           End If
23370         End If
23380       End With
23390     End Select
23400   End If

EXITP:
23410   KeyCode = intRetVal
23420   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub JrnlMemo_Memo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlMemo_Memo_MouseMove"

23510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
23520   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdAssetNew_GotFocus()

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetNew_GotFocus"

23610   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
23620     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
23630   Else
23640     blnAssetNew_Focus = True
23650     JC_Btn_FocusSet THIS_PROC, blnAssetNew_Focus  ' ** Module Procedure: modJrnlCol_Buttons.
23660   End If
23670   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
23680   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdAssetNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetNew_MouseDown"

23710   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
23720     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
23730   End If

EXITP:
23740   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdAssetNew_Click()

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetNew_Click"

        Dim blnToTaxLot As Boolean

23810   With Me
23820     blnToTaxLot = .frmJournal_Columns_Sub.Form.TaxLot_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
23830     JC_Frm_Assets Me, .frmJournal_Columns_Sub.Form, blnToTaxLot  ' ** Module Procedure: modJrnlCol_Forms.
23840     .frmJournal_Columns_Sub.Form.TaxLot_GetSet False, blnToTaxLot  ' ** Form Procedure: frmJournal_Columns_Sub.
23850   End With

EXITP:
23860   Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     Me.cmdLocNew.Enabled = True
1000    Me.cmdRecurNew.Enabled = True
1010    THAT_PROC = THIS_PROC
1020    That_Erl = Erl: That_Desc = ERR.description
1030    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1040    Resume EXITP

End Sub

Private Sub cmdAssetNew_KeyDown(KeyCode As Integer, Shift As Integer)

23900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetNew_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

23910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23920   intShiftDown = (Shift And acShiftMask) > 0
23930   intAltDown = (Shift And acAltMask) > 0
23940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23960     Select Case intRetVal
          Case vbKeyTab
23970       With Me
23980         intRetVal = 0
23990         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
24000         If strNext = (THIS_NAME_SUB) Then
24010           .Controls(strNext).SetFocus
24020           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
24030           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
24040           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
24050           .Controls(strNext).Form.Controls(strNextSub).SetFocus
24060         Else
24070           .Controls(strNext).SetFocus
24080         End If
24090       End With
24100     End Select
24110   End If

        ' ** Shift keys.
24120   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24130     Select Case intRetVal
          Case vbKeyTab
24140       With Me
24150         intRetVal = 0
24160         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
24170         If strNext = (THIS_NAME_SUB) Then
24180           .Controls(strNext).SetFocus
24190           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
24200           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
24210           .Controls(strNext).Form.Controls(strNextSub).SetFocus
24220           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
24230         Else
24240           .Controls(strNext).SetFocus
24250         End If
24260       End With
24270     End Select
24280   End If

EXITP:
24290   KeyCode = intRetVal
24300   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdAssetNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetNew_MouseMove"

24410   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
24420     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
24430   End If

EXITP:
24440   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdAssetNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetNew_MouseUp"

24510   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
24520     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
24530   End If

EXITP:
24540   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdAssetNew_LostFocus()

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetNew_LostFocus"

24610   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
24620     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
24630   Else
24640     blnAssetNew_Focus = False
24650     JC_Btn_FocusSet THIS_PROC, blnAssetNew_Focus  ' ** Module Procedure: modJrnlCol_Buttons.
24660   End If

EXITP:
24670   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdLocNew_GotFocus()

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLocNew_GotFocus"

24710   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
24720     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
24730   Else
24740     blnLocNew_Focus = True
24750     JC_Btn_FocusSet THIS_PROC, blnLocNew_Focus  ' ** Module Procedure: modJrnlCol_Buttons.
24760   End If
24770   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
24780   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdLocNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLocNew_MouseDown"

24810   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
24820     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
24830   End If

EXITP:
24840   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdLocNew_Click()

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLocNew_Click"

24910   With Me
24920     JC_Frm_Locations Me, .frmJournal_Columns_Sub.Form  ' ** Module Procedure: modJrnlCol_Forms.
24930   End With

EXITP:
24940   Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     Me.cmdAssetNew.Enabled = True
1000    Me.cmdRecurNew.Enabled = True
1010    THAT_PROC = THIS_PROC
1020    That_Erl = Erl: That_Desc = ERR.description
1030    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1040    Resume EXITP

End Sub

Private Sub cmdLocNew_KeyDown(KeyCode As Integer, Shift As Integer)

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLocNew_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

25010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25020   intShiftDown = (Shift And acShiftMask) > 0
25030   intAltDown = (Shift And acAltMask) > 0
25040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25060     Select Case intRetVal
          Case vbKeyTab
25070       With Me
25080         intRetVal = 0
25090         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
25100         If strNext = (THIS_NAME_SUB) Then
25110           .Controls(strNext).SetFocus
25120           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
25130           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
25140           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
25150           .Controls(strNext).Form.Controls(strNextSub).SetFocus
25160         Else
25170           .Controls(strNext).SetFocus
25180         End If
25190       End With
25200     End Select
25210   End If

        ' ** Shift keys.
25220   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25230     Select Case intRetVal
          Case vbKeyTab
25240       With Me
25250         intRetVal = 0
25260         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
25270         If strNext = (THIS_NAME_SUB) Then
25280           .Controls(strNext).SetFocus
25290           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
25300           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
25310           .Controls(strNext).Form.Controls(strNextSub).SetFocus
25320           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
25330         Else
25340           .Controls(strNext).SetFocus
25350         End If
25360       End With
25370     End Select
25380   End If

EXITP:
25390   KeyCode = intRetVal
25400   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdLocNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLocNew_MouseMove"

25510   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
25520     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
25530   End If

EXITP:
25540   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdLocNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLocNew_MouseUp"

25610   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
25620     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
25630   End If

EXITP:
25640   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdLocNew_LostFocus()

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLocNew_LostFocus"

25710   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
25720     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
25730   Else
25740     blnLocNew_Focus = False
25750     JC_Btn_FocusSet THIS_PROC, blnLocNew_Focus  ' ** Module Procedure: modJrnlCol_Buttons.
25760   End If

EXITP:
25770   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdRecurNew_GotFocus()

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRecurNew_GotFocus"

25810   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
25820     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
25830   Else
25840     blnRecurNew_Focus = True
25850     JC_Btn_FocusSet THIS_PROC, blnRecurNew_Focus  ' ** Module Procedure: modJrnlCol_Buttons.
25860   End If
25870   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
25880   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdRecurNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRecurNew_MouseDown"

25910   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
25920     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
25930   End If

EXITP:
25940   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdRecurNew_Click()

26000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRecurNew_Click"

26010   With Me
26020     JC_Frm_RecurringItems Me, .frmJournal_Columns_Sub.Form  ' ** Module Procedure: modJrnlCol_Forms.
26030   End With

EXITP:
26040   Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     Me.cmdAssetNew.Enabled = True
1000    Me.cmdLocNew.Enabled = True
1010    THAT_PROC = THIS_PROC
1020    That_Erl = Erl: That_Desc = ERR.description
1030    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1040    Resume EXITP

End Sub

Private Sub cmdRecurNew_KeyDown(KeyCode As Integer, Shift As Integer)

26100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRecurNew_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

26110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26120   intShiftDown = (Shift And acShiftMask) > 0
26130   intAltDown = (Shift And acAltMask) > 0
26140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26160     Select Case intRetVal
          Case vbKeyTab
26170       With Me
26180         intRetVal = 0
26190         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
26200         If strNext = (THIS_NAME_SUB) Then
26210           .Controls(strNext).SetFocus
26220           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
26230           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
26240           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
26250           .Controls(strNext).Form.Controls(strNextSub).SetFocus
26260         Else
26270           .Controls(strNext).SetFocus
26280         End If
26290       End With
26300     End Select
26310   End If

        ' ** Shift keys.
26320   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26330     Select Case intRetVal
          Case vbKeyTab
26340       With Me
26350         intRetVal = 0
26360         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
26370         If strNext = (THIS_NAME_SUB) Then
26380           .Controls(strNext).SetFocus
26390           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
26400           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
26410           .Controls(strNext).Form.Controls(strNextSub).SetFocus
26420           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
26430         Else
26440           .Controls(strNext).SetFocus
26450         End If
26460       End With
26470     End Select
26480   End If

EXITP:
26490   KeyCode = intRetVal
26500   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdRecurNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRecurNew_MouseMove"

26610   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
26620     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
26630   End If

EXITP:
26640   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdRecurNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRecurNew_MouseUp"

26710   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
26720     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
26730   End If

EXITP:
26740   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdRecurNew_LostFocus()

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRecurNew_LostFocus"

26810   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
26820     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
26830   Else
26840     blnRecurNew_Focus = False
26850     JC_Btn_FocusSet THIS_PROC, blnRecurNew_Focus  ' ** Module Procedure: modJrnlCol_Buttons.
26860   End If

EXITP:
26870   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdRefresh_GotFocus()

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRefresh_GotFocus"

26910   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
26920     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
26930   Else
26940     blnRefresh_Focus = True
26950     JC_Btn_FocusSet THIS_PROC, blnRefresh_Focus  ' ** Module Procedure: modJrnlCol_Buttons.
26960   End If
26970   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
26980   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdRefresh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRefresh_MouseDown"

27010   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
27020     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
27030   End If

EXITP:
27040   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdRefresh_Click()

27100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRefresh_Click"

27110   With Me
27120     JC_Frm_Refresh .frmJournal_Columns_Sub.Form  ' ** Module Procedure: modJrnlCol_Forms.
27130   End With

EXITP:
27140   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRefresh_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

27210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27220   intShiftDown = (Shift And acShiftMask) > 0
27230   intAltDown = (Shift And acAltMask) > 0
27240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27260     Select Case intRetVal
          Case vbKeyTab
27270       With Me
27280         intRetVal = 0
27290         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
27300         If strNext = (THIS_NAME_SUB) Then
27310           .Controls(strNext).SetFocus
27320           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
27330           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
27340           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
27350           .Controls(strNext).Form.Controls(strNextSub).SetFocus
27360         Else
27370           .Controls(strNext).SetFocus
27380         End If
27390       End With
27400     End Select
27410   End If

        ' ** Shift keys.
27420   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27430     Select Case intRetVal
          Case vbKeyTab
27440       With Me
27450         intRetVal = 0
27460         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
27470         If strNext = (THIS_NAME_SUB) Then
27480           .Controls(strNext).SetFocus
27490           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
27500           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
27510           .Controls(strNext).Form.Controls(strNextSub).SetFocus
27520           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
27530         Else
27540           .Controls(strNext).SetFocus
27550         End If
27560       End With
27570     End Select
27580   End If

EXITP:
27590   KeyCode = intRetVal
27600   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRefresh_MouseMove"

27710   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
27720     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
27730   End If

EXITP:
27740   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdRefresh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRefresh_MouseUp"

27810   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
27820     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
27830   End If

EXITP:
27840   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdRefresh_LostFocus()

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRefresh_LostFocus"

27910   If JC_Btn_Resize_Norm(Me) = False Then  ' ** Module Procedure: modJrnlCol_Buttons.
27920     JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
27930   Else
27940     blnRefresh_Focus = False
27950     JC_Btn_FocusSet THIS_PROC, blnRefresh_Focus  ' ** Module Procedure: modJrnlCol_Buttons.
27960   End If

EXITP:
27970   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdAdd_GotFocus()

28000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_GotFocus"

28010   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
28020   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdAdd_Click()

28100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_Click"

28110   With Me
28120     .frmJournal_Columns_Sub.SetFocus
          '.frmJournal_Columns_Sub.Form.AddRec  ' ** Form Procedure: frmJournal_Columns_Sub.
28130     .frmJournal_Columns_Sub.Form.AddRec_Send  ' ** Form Procedure: frmJournal_Columns_Sub.
28140   End With

EXITP:
28150   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)

28200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

28210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28220   intShiftDown = (Shift And acShiftMask) > 0
28230   intAltDown = (Shift And acAltMask) > 0
28240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28260     Select Case intRetVal
          Case vbKeyTab
28270       With Me
28280         intRetVal = 0
28290         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
28300         If strNext = (THIS_NAME_SUB) Then
28310           .Controls(strNext).SetFocus
28320           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
28330           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
28340           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
28350           .Controls(strNext).Form.Controls(strNextSub).SetFocus
28360         Else
28370           .Controls(strNext).SetFocus
28380         End If
28390       End With
28400     End Select
28410   End If

        ' ** Shift keys.
28420   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28430     Select Case intRetVal
          Case vbKeyTab
28440       With Me
28450         intRetVal = 0
28460         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
28470         If strNext = (THIS_NAME_SUB) Then
28480           .Controls(strNext).SetFocus
28490           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
28500           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
28510           .Controls(strNext).Form.Controls(strNextSub).SetFocus
28520           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
28530         Else
28540           .Controls(strNext).SetFocus
28550         End If
28560       End With
28570     End Select
28580   End If

EXITP:
28590   KeyCode = intRetVal
28600   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdEdit_GotFocus()

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEdit_GotFocus"

28710   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
28720   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdEdit_Click()

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEdit_Click"

28810   MsgBox "Under Construction", vbInformation + vbOKOnly, "Full Screen Edit Unavailable"
        'With Me
        '  Cancel = 0
        '  .frmJournal_Columns_Sub.Form.Detail_DblClick Cancel  ' ** Form Procedure: frmJournal_Columns_Sub.
        'End With

EXITP:
28820   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdEdit_KeyDown(KeyCode As Integer, Shift As Integer)

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEdit_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

28910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28920   intShiftDown = (Shift And acShiftMask) > 0
28930   intAltDown = (Shift And acAltMask) > 0
28940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28960     Select Case intRetVal
          Case vbKeyTab
28970       With Me
28980         intRetVal = 0
28990         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
29000         If strNext = (THIS_NAME_SUB) Then
29010           .Controls(strNext).SetFocus
29020           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
29030           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
29040           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
29050           .Controls(strNext).Form.Controls(strNextSub).SetFocus
29060         Else
29070           .Controls(strNext).SetFocus
29080         End If
29090       End With
29100     End Select
29110   End If

        ' ** Shift keys.
29120   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29130     Select Case intRetVal
          Case vbKeyTab
29140       With Me
29150         intRetVal = 0
29160         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
29170         If strNext = (THIS_NAME_SUB) Then
29180           .Controls(strNext).SetFocus
29190           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
29200           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
29210           .Controls(strNext).Form.Controls(strNextSub).SetFocus
29220           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
29230         Else
29240           .Controls(strNext).SetFocus
29250         End If
29260       End With
29270     End Select
29280   End If

EXITP:
29290   KeyCode = intRetVal
29300   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdDelete_GotFocus()

29400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_GotFocus"

29410   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
29420   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdDelete_Click()

29500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_Click"

29510   JC_Ctl_Delete blnGoneToReport, Me  ' ** Module Procedure: modJrnlCol_Controls.

EXITP:
29520   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)

29600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

29610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29620   intShiftDown = (Shift And acShiftMask) > 0
29630   intAltDown = (Shift And acAltMask) > 0
29640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29660     Select Case intRetVal
          Case vbKeyTab
29670       With Me
29680         intRetVal = 0
29690         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
29700         If strNext = (THIS_NAME_SUB) Then
29710           .Controls(strNext).SetFocus
29720           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
29730           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
29740           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
29750           .Controls(strNext).Form.Controls(strNextSub).SetFocus
29760         Else
29770           .Controls(strNext).SetFocus
29780         End If
29790       End With
29800     End Select
29810   End If

        ' ** Shift keys.
29820   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29830     Select Case intRetVal
          Case vbKeyTab
29840       With Me
29850         intRetVal = 0
29860         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
29870         If strNext = (THIS_NAME_SUB) Then
29880           .Controls(strNext).SetFocus
29890           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
29900           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
29910           .Controls(strNext).Form.Controls(strNextSub).SetFocus
29920           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
29930         Else
29940           .Controls(strNext).SetFocus
29950         End If
29960       End With
29970     End Select
29980   End If

EXITP:
29990   KeyCode = intRetVal
30000   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_box_MouseMove"

30110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
30120   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Div_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Div_lbl_MouseMove"

30210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
30220   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Div_Map_GotFocus()

30300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Div_Map_GotFocus"

30310   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
30320   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
30330   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Div_Map_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

30400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Div_Map_MouseDown"

30410   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
30420   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdSpecPurp_Div_Map_Click()

30500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Div_Map_Click"

30510   With Me
30520     JC_Frm_Dividend .frmJournal_Columns_Sub.Form.RecCnt, lngJrnlColID_Max  ' ** Module Procedure: modJrnlCol_Forms.
30530   End With                                                                 ' ** Form Function: frmJournal_Columns_Sub.

EXITP:
30540   Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Div_Map_KeyDown(KeyCode As Integer, Shift As Integer)

30600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Div_Map_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

30610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
30620   intShiftDown = (Shift And acShiftMask) > 0
30630   intAltDown = (Shift And acAltMask) > 0
30640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30660     Select Case intRetVal
          Case vbKeyTab
30670       With Me
30680         intRetVal = 0
30690         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
30700         If strNext = (THIS_NAME_SUB) Then
30710           .Controls(strNext).SetFocus
30720           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
30730           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
30740           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
30750           .Controls(strNext).Form.Controls(strNextSub).SetFocus
30760         Else
30770           .Controls(strNext).SetFocus
30780         End If
30790       End With
30800     End Select
30810   End If

        ' ** Shift keys.
30820   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
30830     Select Case intRetVal
          Case vbKeyTab
30840       With Me
30850         intRetVal = 0
30860         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
30870         If strNext = (THIS_NAME_SUB) Then
30880           .Controls(strNext).SetFocus
30890           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
30900           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
30910           .Controls(strNext).Form.Controls(strNextSub).SetFocus
30920           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
30930         Else
30940           .Controls(strNext).SetFocus
30950         End If
30960       End With
30970     End Select
30980   End If

EXITP:
30990   KeyCode = intRetVal
31000   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Div_Map_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

31100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Div_Map_MouseMove"

31110   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
31120   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Div_Map_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

31200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Div_Map_MouseUp"

31210   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
31220   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Div_Map_LostFocus()

31300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Div_Map_LostFocus"

31310   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
31320   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Int_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

31400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Int_lbl_MouseMove"

31410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
31420   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Int_Map_GotFocus()

31500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Int_Map_GotFocus"

31510   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
31520   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
31530   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Int_Map_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

31600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Int_Map_MouseDown"

31610   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
31620   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdSpecPurp_Int_Map_Click()

31700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Int_Map_Click"

31710   With Me
31720     JC_Frm_Interest .frmJournal_Columns_Sub.Form.RecCnt, lngJrnlColID_Max  ' ** Module Procedure: modJrnlCol_Forms.
31730   End With                                                                 ' ** Form Function: frmJournal_Columns_Sub.

EXITP:
31740   Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Int_Map_KeyDown(KeyCode As Integer, Shift As Integer)

31800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Int_Map_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

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
31890         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
31900         If strNext = (THIS_NAME_SUB) Then
31910           .Controls(strNext).SetFocus
31920           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
31930           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
31940           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
31950           .Controls(strNext).Form.Controls(strNextSub).SetFocus
31960         Else
31970           .Controls(strNext).SetFocus
31980         End If
31990       End With
32000     End Select
32010   End If

        ' ** Shift keys.
32020   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
32030     Select Case intRetVal
          Case vbKeyTab
32040       With Me
32050         intRetVal = 0
32060         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
32070         If strNext = (THIS_NAME_SUB) Then
32080           .Controls(strNext).SetFocus
32090           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
32100           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
32110           .Controls(strNext).Form.Controls(strNextSub).SetFocus
32120           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
32130         Else
32140           .Controls(strNext).SetFocus
32150         End If
32160       End With
32170     End Select
32180   End If

EXITP:
32190   KeyCode = intRetVal
32200   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Int_Map_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

32300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Int_Map_MouseMove"

32310   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
32320   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Int_Map_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

32400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Int_Map_MouseUp"

32410   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
32420   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Int_Map_LostFocus()

32500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Int_Map_LostFocus"

32510   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
32520   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Sold_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

32600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Sold_lbl_MouseMove"

32610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
32620   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Sold_PaidTotal_GotFocus()

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Sold_PaidTotal_GotFocus"

32710   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
32720   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
32730   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Sold_PaidTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

32800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Sold_PaidTotal_MouseDown"

32810   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
32820   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdSpecPurp_Sold_PaidTotal_Click()

32900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Sold_PaidTotal_Click"

32910   With Me
32920     JC_Frm_Paid .frmJournal_Columns_Sub.Form  ' ** Module Procedure: modJrnlCol_Forms.
32930     .frmJournal_Columns_Sub.SetFocus
32940   End With

EXITP:
32950   Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Sold_PaidTotal_KeyDown(KeyCode As Integer, Shift As Integer)

33000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Sold_PaidTotal_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

33010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
33020   intShiftDown = (Shift And acShiftMask) > 0
33030   intAltDown = (Shift And acAltMask) > 0
33040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
33050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
33060     Select Case intRetVal
          Case vbKeyTab
33070       With Me
33080         intRetVal = 0
33090         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
33100         If strNext = (THIS_NAME_SUB) Then
33110           .Controls(strNext).SetFocus
33120           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
33130           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
33140           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
33150           .Controls(strNext).Form.Controls(strNextSub).SetFocus
33160         Else
33170           .Controls(strNext).SetFocus
33180         End If
33190       End With
33200     End Select
33210   End If

        ' ** Shift keys.
33220   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
33230     Select Case intRetVal
          Case vbKeyTab
33240       With Me
33250         intRetVal = 0
33260         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
33270         If strNext = (THIS_NAME_SUB) Then
33280           .Controls(strNext).SetFocus
33290           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
33300           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
33310           .Controls(strNext).Form.Controls(strNextSub).SetFocus
33320           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
33330         Else
33340           .Controls(strNext).SetFocus
33350         End If
33360       End With
33370     End Select
33380   End If

EXITP:
33390   KeyCode = intRetVal
33400   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Sold_PaidTotal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

33500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Sold_PaidTotal_MouseMove"

33510   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
33520   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Sold_PaidTotal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

33600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Sold_PaidTotal_MouseUp"

33610   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
33620   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Sold_PaidTotal_LostFocus()

33700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Sold_PaidTotal_LostFocus"

33710   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
33720   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Purch_MapSplit_GotFocus()

33800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Purch_MapSplit_GotFocus"

33810   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
33820   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
33830   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Purch_MapSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

33900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Purch_MapSplit_MouseDown"

33910   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
33920   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdSpecPurp_Purch_MapSplit_Click()

34000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Purch_MapSplit_Click"

34010   With Me
34020     JC_Frm_Split .frmJournal_Columns_Sub.Form.RecCnt, lngJrnlColID_Max  ' ** Module Procedure: modJrnlCol_Forms.
34030   End With                                                              ' ** Form Function: frmJournal_Columns_Sub.

EXITP:
34040   Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Purch_MapSplit_KeyDown(KeyCode As Integer, Shift As Integer)

34100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Purch_MapSplit_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

34110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
34120   intShiftDown = (Shift And acShiftMask) > 0
34130   intAltDown = (Shift And acAltMask) > 0
34140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
34150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
34160     Select Case intRetVal
          Case vbKeyTab
34170       With Me
34180         intRetVal = 0
34190         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
34200         If strNext = (THIS_NAME_SUB) Then
34210           .Controls(strNext).SetFocus
34220           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
34230           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
34240           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
34250           .Controls(strNext).Form.Controls(strNextSub).SetFocus
34260         Else
34270           .Controls(strNext).SetFocus
34280         End If
34290       End With
34300     End Select
34310   End If

        ' ** Shift keys.
34320   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
34330     Select Case intRetVal
          Case vbKeyTab
34340       With Me
34350         intRetVal = 0
34360         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
34370         If strNext = (THIS_NAME_SUB) Then
34380           .Controls(strNext).SetFocus
34390           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
34400           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
34410           .Controls(strNext).Form.Controls(strNextSub).SetFocus
34420           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
34430         Else
34440           .Controls(strNext).SetFocus
34450         End If
34460       End With
34470     End Select
34480   End If

EXITP:
34490   KeyCode = intRetVal
34500   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Purch_MapSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

34600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Purch_MapSplit_MouseMove"

34610   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
34620   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Purch_MapSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

34700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Purch_MapSplit_MouseUp"

34710   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
34720   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdSpecPurp_Purch_MapSplit_LostFocus()
' ** Made Public for extra GTR check.

34800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Purch_MapSplit_LostFocus"

34810   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
34820   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdSpecPurp_Misc_MapLTCG_GotFocus()
' ** Made Public for extra GTR check.

34900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCG_GotFocus"

34910   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
34920   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
34930   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapLTCG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

35000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCG_MouseDown"

35010   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
35020   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdSpecPurp_Misc_MapLTCG_Click()

35100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCG_Click"

35110   With Me
35120     JC_Frm_LTCG .frmJournal_Columns_Sub.Form.RecCnt, lngJrnlColID_Max  ' ** Module Procedure: modJrnlCol_Forms.
35130   End With                                                             ' ** Form Function: frmJournal_Columns_Sub.

EXITP:
35140   Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapLTCG_KeyDown(KeyCode As Integer, Shift As Integer)

35200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCG_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

35210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
35220   intShiftDown = (Shift And acShiftMask) > 0
35230   intAltDown = (Shift And acAltMask) > 0
35240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
35250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
35260     Select Case intRetVal
          Case vbKeyTab
35270       With Me
35280         intRetVal = 0
35290         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
35300         If strNext = (THIS_NAME_SUB) Then
35310           .Controls(strNext).SetFocus
35320           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
35330           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
35340           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
35350           .Controls(strNext).Form.Controls(strNextSub).SetFocus
35360         Else
35370           .Controls(strNext).SetFocus
35380         End If
35390       End With
35400     End Select
35410   End If

        ' ** Shift keys.
35420   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
35430     Select Case intRetVal
          Case vbKeyTab
35440       With Me
35450         intRetVal = 0
35460         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
35470         If strNext = (THIS_NAME_SUB) Then
35480           .Controls(strNext).SetFocus
35490           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
35500           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
35510           .Controls(strNext).Form.Controls(strNextSub).SetFocus
35520           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
35530         Else
35540           .Controls(strNext).SetFocus
35550         End If
35560       End With
35570     End Select
35580   End If

EXITP:
35590   KeyCode = intRetVal
35600   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapLTCG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCG_MouseMove"

35710   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
35720   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapLTCG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

35800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCG_MouseUp"

35810   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
35820   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapLTCG_LostFocus()

35900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCG_LostFocus"

35910   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
35920   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapLTCL_GotFocus()

36000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCL_GotFocus"

36010   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
36020   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
36030   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapLTCL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

36100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCL_MouseDown"

36110   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
36120   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdSpecPurp_Misc_MapLTCL_Click()

36200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCL_Click"

36210   With Me
36220     JC_Frm_LTCL .frmJournal_Columns_Sub.Form.RecCnt, lngJrnlColID_Max  ' ** Module Procedure: modJrnlCol_Forms.
36230   End With                                                             ' ** Form Function: frmJournal_Columns_Sub.

EXITP:
36240   Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapLTCL_KeyDown(KeyCode As Integer, Shift As Integer)

36300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCL_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

36310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
36320   intShiftDown = (Shift And acShiftMask) > 0
36330   intAltDown = (Shift And acAltMask) > 0
36340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
36350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
36360     Select Case intRetVal
          Case vbKeyTab
36370       With Me
36380         intRetVal = 0
36390         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
36400         If strNext = (THIS_NAME_SUB) Then
36410           .Controls(strNext).SetFocus
36420           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
36430           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
36440           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
36450           .Controls(strNext).Form.Controls(strNextSub).SetFocus
36460         Else
36470           .Controls(strNext).SetFocus
36480         End If
36490       End With
36500     End Select
36510   End If

        ' ** Shift keys.
36520   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
36530     Select Case intRetVal
          Case vbKeyTab
36540       With Me
36550         intRetVal = 0
36560         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
36570         If strNext = (THIS_NAME_SUB) Then
36580           .Controls(strNext).SetFocus
36590           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
36600           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
36610           .Controls(strNext).Form.Controls(strNextSub).SetFocus
36620           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
36630         Else
36640           .Controls(strNext).SetFocus
36650         End If
36660       End With
36670     End Select
36680   End If

EXITP:
36690   KeyCode = intRetVal
36700   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapLTCL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

36800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCL_MouseMove"

36810   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
36820   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapLTCL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

36900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCL_MouseUp"

36910   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
36920   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapLTCL_LostFocus()

37000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapLTCL_LostFocus"

37010   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
37020   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapSTCGL_GotFocus()

37100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapSTCGL_GotFocus"

37110   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.
37120   StatusBar_Set THIS_PROC  ' ** Procedure: Below.

EXITP:
37130   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapSTCGL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

37200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapSTCGL_MouseDown"

37210   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
37220   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub cmdSpecPurp_Misc_MapSTCGL_Click()

37300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapSTCGL_Click"

37310   With Me
37320     JC_Frm_STCGL .frmJournal_Columns_Sub.Form.RecCnt, lngJrnlColID_Max  ' ** Module Procedure: modJrnlCol_Forms.
37330   End With                                                             ' ** Form Function: frmJournal_Columns_Sub.

EXITP:
37340   Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapSTCGL_KeyDown(KeyCode As Integer, Shift As Integer)

37400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapSTCGL_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strNext As String, strNextSub As String
        Dim intRetVal As Integer, lngRetVal As Long

37410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
37420   intShiftDown = (Shift And acShiftMask) > 0
37430   intAltDown = (Shift And acAltMask) > 0
37440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
37450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
37460     Select Case intRetVal
          Case vbKeyTab
37470       With Me
37480         intRetVal = 0
37490         strNext = JC_Key_Par_Next(THIS_PROC, Me)  ' ** Module Function: modJrnlCol_Keys.
37500         If strNext = (THIS_NAME_SUB) Then
37510           .Controls(strNext).SetFocus
37520           .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
37530           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
37540           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Module Function: modJrnlCol_Keys.
37550           .Controls(strNext).Form.Controls(strNextSub).SetFocus
37560         Else
37570           .Controls(strNext).SetFocus
37580         End If
37590       End With
37600     End Select
37610   End If

        ' ** Shift keys.
37620   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
37630     Select Case intRetVal
          Case vbKeyTab
37640       With Me
37650         intRetVal = 0
37660         strNext = JC_Key_Par_Next(THIS_PROC, Me, False)  ' ** Module Function: modJrnlCol_Keys.
37670         If strNext = (THIS_NAME_SUB) Then
37680           .Controls(strNext).SetFocus
37690           .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
37700           strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Module Function: modJrnlCol_Keys.
37710           .Controls(strNext).Form.Controls(strNextSub).SetFocus
37720           lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
37730         Else
37740           .Controls(strNext).SetFocus
37750         End If
37760       End With
37770     End Select
37780   End If

EXITP:
37790   KeyCode = intRetVal
37800   Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
1020    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapSTCGL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

37900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapSTCGL_MouseMove"

37910   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
37920   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapSTCGL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

38000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapSTCGL_MouseUp"

38010   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
38020   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub cmdSpecPurp_Misc_MapSTCGL_LostFocus()

38100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSpecPurp_Misc_MapSTCGL_LostFocus"

38110   JC_Btn_Image THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Buttons.

EXITP:
38120   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub ThisNameSub_Set()

38200 On Error GoTo ERRH

        Const THIS_PROC As String = "ThisNameSub_Set"

        Dim ctl As Access.Control

38210   With Me
38220     For Each ctl In .Controls
38230       With ctl
38240         If .ControlType = acSubform Then
38250           THIS_NAME_SUB = .Name
38260           Exit For
38270         End If
38280       End With
38290     Next
38300   End With

EXITP:
38310   Set ctl = Nothing
38320   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub StatusBar_Set(strProc As String)

38400 On Error GoTo ERRH

        Const THIS_PROC As String = "StatusBar_Set"

        Dim strControl As String
        Dim intPos01 As Integer
        Dim lngX As Long

38410   StatusBar_Load  ' ** Procedure: Below.

        ' ********************************************
        ' ** Array: arr_varStat()
        ' **
        ' **   Element  Name              Constant
        ' **   =======  ================  ==========
        ' **      0     Section Number    S_SEC
        ' **      1     Control Name      S_NAM
        ' **      2     StatusBarText     S_TXT
        ' **
        ' ********************************************

38420   If blnNotPopup = True And strProc <> vbNullString Then
38430     intPos01 = InStr(strProc, "_GotFocus")
38440     If intPos01 > 0 Then
38450       strControl = Left(strProc, (intPos01 - 1))
38460       For lngX = 0& To (lngStats - 1&)
38470         If arr_varStat(S_NAM, lngX) = strControl Then
38480           If arr_varStat(S_TXT, lngX) <> vbNullString Then
38490             SysCmd acSysCmdSetStatus, arr_varStat(S_TXT, lngX)
38500           Else
38510             SysCmd acSysCmdClearStatus
38520           End If
38530           Exit For
38540         End If
38550       Next
38560     End If
38570   End If

EXITP:
38580   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Private Sub StatusBar_Load()

38600 On Error GoTo ERRH

        Const THIS_PROC As String = "StatusBar_Load"

        'Dim intMaxLen As Integer
        'Dim lngX As Long
        Dim arr_varRetVal As Variant

        Const RV_ERR   As Integer = 0
        Const RV_SBARS As Integer = 1
        Const RV_POPUP As Integer = 2

38610   If lngStats = 0& Or IsEmpty(arr_varStat) Then
38620     arr_varRetVal = JC_Msc_SBar_Load(Me)  ' ** Module Function: modJrnlCol_Misc.
38630     If arr_varRetVal(RV_ERR, 0) = vbNullString Then
38640       arr_varStat = arr_varRetVal(RV_SBARS, 0)
38650       lngStats = (UBound(arr_varStat, 2) + 1&)
38660       blnNotPopup = arr_varRetVal(RV_POPUP, 0)
38670     End If
38680   End If

        'intMaxLen = 0&
        'For lngX = 0& To (lngStats - 1)
        '  If arr_varStat(S_TXT, lngX) <> vbNullString Then
        '    If Len(arr_varStat(S_NAM, lngX)) > intMaxLen Then intMaxLen = Len(arr_varStat(S_NAM, lngX))
        '  End If
        'Next
        'For lngX = 0& To (lngStats - 1)
        '  If arr_varStat(S_TXT, lngX) <> vbNullString Then
        '    Debug.Print "'CTL: " & Left(arr_varStat(S_NAM, lngX) & Space(intMaxLen), intMaxLen) & "  TXT: " & arr_varStat(S_TXT, lngX)
        '  End If
        'Next
        'For lngX = 0& To (lngStats - 1)
        '  If arr_varStat(S_TXT, lngX) = vbNullString Then
        '    Debug.Print "'NO TXT: " & arr_varStat(S_NAM, lngX)
        '  End If
        'Next

        ' ** Won't ever get hit:
        ' **   NO TXT: RecsTot_Committed
        ' **   NO TXT: RecsTot_Uncommitted

EXITP:
38690   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Function IsMaximized_Get() As Boolean

38700 On Error GoTo ERRH

        Const THIS_PROC As String = "IsMaximized_Get"

        Dim blnRetVal As Boolean

38710   blnRetVal = IsMaximized(Me)  ' ** Module Function: modWindowFunctions.

EXITP:
38720   IsMaximized_Get = blnRetVal
38730   Exit Function

ERRH:
980     blnRetVal = False
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl: That_Desc = ERR.description
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1020    Resume EXITP

End Function

Public Sub IsMaximized_Set(blnMax As Boolean)

38800 On Error GoTo ERRH

        Const THIS_PROC As String = "IsMaximized_Set"

38810   Select Case blnMax
        Case True
38820     DoCmd.SelectObject acForm, THIS_NAME, False
38830     DoCmd.Maximize
38840   Case False
38850     DoCmd.SelectObject acForm, THIS_NAME, False
38860     DoCmd.Restore
38870   End Select
38880   DoEvents

EXITP:
38890   Exit Sub

ERRH:
980     THAT_PROC = THIS_PROC
990     That_Erl = Erl: That_Desc = ERR.description
1000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1010    Resume EXITP

End Sub

Public Sub JrnlCol_FocusSet(blnGotFocus_Asset As Boolean, blnGotFocus_Loc As Boolean, blnGotFocus_Recur As Boolean, blnGotFocus_Refresh As Boolean)

38900 On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlCol_FocusSet"

38910   blnAssetNew_Focus = blnGotFocus_Asset
38920   blnLocNew_Focus = blnGotFocus_Loc
38930   blnRecurNew_Focus = blnGotFocus_Recur
38940   blnRefresh_Focus = blnGotFocus_Refresh

EXITP:
38950   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Public Sub CommitNoClose_Set(blnCommit As Boolean)

39000 On Error GoTo ERRH

        Const THIS_PROC As String = "CommitNoClose_Set"

39010   blnCommitNoClose = blnCommit

EXITP:
39020   Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Public Function GTRStuff(intMode As Integer, blnGTR As Boolean, Optional varJColID As Variant) As Boolean

39100 On Error GoTo ERRH

        Const THIS_PROC As String = "GTRStuff"

        Dim lngJColID As Long
        Dim blnRetVal As Boolean

39110   blnRetVal = False

        ' ** GTRStuff
        ' **   1 : GTREmblem_Set blnSetEmblem
        ' **   2 : GTREmblem_Get {none}
        ' **   3 : GTREmblem_Move {none}
        ' **   4 : GTREmblem_Off {none}
        ' **   5 : GTRGone_Set blnWentToReport As Boolean, lngJColID As Long

39120   Select Case IsMissing(varJColID)
        Case True
39130     lngJColID = 0&
39140   Case False
39150     lngJColID = varJColID
39160   End Select

39170   blnRetVal = GTRStuff_Mod(Me, intMode, blnGTR, blnGTR_Emblem, blnGoingToReport, _
          blnGoneToReport, lngGTR_Stat, lngGTR_ID, lngJColID)  ' ** Module Function: modGoToReportFuncs.

EXITP:
39180   GTRStuff = blnRetVal
39190   Exit Function

ERRH:
980     blnRetVal = False
990     Select Case ERR.Number
        Case Else
1000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1010    End Select
1020    Resume EXITP

End Function
