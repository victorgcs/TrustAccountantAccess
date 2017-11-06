VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJournal_Columns_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmJournal_Columns_Sub"

'VGC 09/09/2017: CHANGES!

' #########################
' ## Use VBA_RenumErrh().  39750
' #########################

'tblJOURNAL_FIELD DOESN'T GET AUTOMATICALLY UPDATED!!!

' ** Shortcut Alt keys responsive from this form:
' **   Show All Assets:  A {chkShowAllAssets}
' **   Close:            C {cmdClose on frmJournal_Columns}
' **   Date:             D {transdate}
' **   Enter Right:      G {opgEnterKey_optRight on frmJournal_Columns}
' **   Income Cash:      H {icash}
' **   Journal Type:     J {journaltype}
' **   Print Check:      K {PrintCheck}
' **   New Location:     L {cmdLocNew on frmJournal_Columns}
' **   Memo:             M {JrnlMemo_Memo on frmJournal_Columns}
' **   Account Number:   N {accountno}
' **   Commit:           O {posted}
' **   Print Report:     P {cmdPrintReport on frmJournal_Columns}
' **   Recurring Item:   R {Recur_Name}
' **   Asset:            S {assetno}
' **   Transactions:     T {frmJournal_Columns_Sub on frmJournal_Columns}
' **   New Recur Item:   U {cmdRecurNew on frmJournal_Columns}
' **   Preview Report:   V {cmdPreviewReport on frmJournal_Columns}
' **   New Asset:        W {cmdAssetNew on frmJournal_Columns}
' **   Exit:             X {cmdClose on frmJournal_Columns}
' **   Enter Down:         {opgEnterKey_optDown on frmJournal_Columns}
'Available Letters:
'B E F I Q Y Z

' ** Shortcut F-keys responsive from this form:
' **   Dropdown Signal:  F4 {Dropdown}
' **   Recalc:           F5 {RecalcTots}
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Add:              A {cmdAdd on frmJournal_Columns}
' **   Delete:           D {cmdDelete on frmJournal_Columns}
' **   Edit:             E {cmdEdit on frmJournal_Columns} (also Alt+E)
' **   Refresh:          H {cmdRefresh on frmJournal_Columns}
' **   Scroll Left:      L {cmdScrollLeft on frmJournal_Columns}
' **   Check Memo:       M {cmdMemoReveal on frmJournal_Columns}
' **   Scroll Right:     R {cmdScrollRight on frmJournal_Columns}
' **   Save:             S {cmdSave}
' **   Switch:           T {cmdSwitch on frmJournal_Columns}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Commit All:       A {cmdUncomComAll on frmJournal_Columns}
' **   Delete All:       D {cmdUncomDelAll on frmJournal_Columns}

' ** Shortcut Alt-Shift keys responsive from this form:
' **   Show All:         A {opgFilter_optAll on frmJournal_Columns}
' **   Show Committed:   C {opgFilter_optCommitted on frmJournal_Columns}
' **   Map Div:          D {cmdSpecPurp_Div_Map on frmJournal_Columns}
' **   Map STCG/L:       G {cmdSpecPurp_Misc_MapSTCGL on frmJournal_Columns}
' **   Paid Total:       I {cmdSpecPurp_Sold_PaidTotal on frmJournal_Columns}
' **   Map LTCG:         L {cmdSpecPurp_Misc_MapLTCG on frmJournal_Columns}
' **   Map Int:          P {cmdSpecPurp_Int_Map on frmJournal_Columns}
' **   Map Split:        S {cmdSpecPurp_Purch_MapSplit on frmJournal_Columns}
' **   Map LTCL:         T {cmdSpecPurp_Misc_MapLTCL on frmJournal_Columns}
' **   Show Uncommitted: U {opgFilter_optUncommitted on frmJournal_Columns}

' ** Current line highlight:
' **   Name: ctlCurrentLine    ControlSource: =GetLineNumber("frmJournal_Columns_Sub","JrnlCol_ID")
' **   Name: ctlCurrentRecord  ControlSource: {unbound}
' **   Name: ctlBack1          ControlSource: =IIf([SelTop]=[ctlCurrentLine],String(255,"Û"),Null)   FontName: Terminal
' **   Name: ctlBack2          ControlSource: =IIf([SelTop]=[ctlCurrentLine],String(255,"Û"),Null)   FontName: Terminal
' **   See GetLineNumber() in modUtilities.

' ** Date checking functions:
' **   DateCheck_Post(varDate As Variant) As Boolean
' **     modUtilities
' **   DateCheck_Trade(varDate As Variant) As Boolean
' **     modUtilities
' **   Compare_DateA_DateB(datA As Date, strOperator As String, datB As Date) As Boolean
' **     modStringFuncs

' ** Subform events fire first!
' **   SUB Form_Open()
' **   SUB Form_Load()
' **   SUB Form_Current()
' **   SUB Form_Current()
' **   Form_Open()
' **   Form_Load()

' ** JrnlCol_ID_lbl.ForeColor  : CLR_GRY
' ** Other_Labels.ForeColor    : CLR_DKGRY2
' ** Label_Lines.BorderColor   : CLR_DKGRY
' ** Sort_Lines.BorderColor    : CLR_LTCYAN
' ** Text_Boxes.ForeColor      : CLR_BLK
' ** Text_Boxes.BackColor      : CLR_WHT
' ** Text_Boxes.BorderColor    : CLR_GRY
' ** SpecPurp_.._lbl.ForeColor : CLR_BLUGRY

' ** Combo Box ListWidths are total columns + 17 pixels for the scrollbar.

' ** Combo box column constants: assetno.
'Private Const CBX_AST_ASTNO As Integer = 0  'assetno
Private Const CBX_AST_DESC  As Integer = 1  'totdesc
'Private Const CBX_AST_CUSIP As Integer = 2  'cusip
'Private Const CBX_AST_TYPE  As Integer = 3  'assettype
'Private Const CBX_AST_TAX   As Integer = 4  'taxcode
'Private Const CBX_AST_RATE  As Integer = 5  'rate
'Private Const CBX_AST_DUE   As Integer = 6  'due
'Private Const CBX_AST_D4D   As Integer = 7  'AssetType_D4D

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Memo field maximum characters.
'Private Const MEMO_MAX As Integer = 40

' ** Array: arr_varStat().
Private lngStats As Long, arr_varStat As Variant
'Private Const S_SEC As Integer = 0
'Private Const S_NAM As Integer = 1
'Private Const S_TXT As Integer = 2

' ** Array: arr_varNewRec().
Private lngNewRecs As Long, arr_varNewRec() As Variant
Private Const N_ELEMS As Integer = 1  ' ** Array's first-element UBound().
Private Const N_ID   As Integer = 0
Private Const N_CMTD As Integer = 1

Private blnJTypeSet As Boolean, blnNextRec As Boolean, blnFromZero As Boolean, blnNotPopup As Boolean, blnWarnZeroCost As Boolean
Private blnToTaxLot As Boolean, blnNoMove As Boolean, blnRecsTotUpdate As Boolean, blnReinvestment As Boolean, blnActNoTab As Boolean
Private blnDontCommitTwice As Boolean

Private blnGTR_NoAdd As Boolean
Private blnGoingToReport As Boolean, blnGoneToReport As Boolean, blnGoneToReport2 As Boolean
Private blnGTR_Emblem As Boolean, lngGTR_Stat As Long, lngGTR_ID As Long
Private lngNewJrnlColID As Long, lngTpp As Long
Private datPostingDate As Date
Private strAccountNo_OldValue As String, strShortName_OldValue As String, strJournalType_OldValue As String
Private lngAssetNo_OldValue As Long, strAssetDesc_OldValue As String
Private dblShareFace_OldValue As Double, dblICash_OldValue As Double, dblPCash_OldValue As Double

Private blnF4Invoked As Boolean, blnF4InvokedMouse As Boolean, strF4LastControl As String
Private strPageMoveCtl As String, strSaveMoveCtl As String
Private lngRecsCur As Long
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long, That_Desc As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim grp As DAO.Group, usr As DAO.User
        Dim blnRetVal As Boolean

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmJournal_Columns" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

180         gblnClosing = False: gblnDeleting = False
190         blnF4Invoked = False: blnF4InvokedMouse = False: strF4LastControl = vbNullString: blnToTaxLot = False: blnNoMove = False
200         blnJTypeSet = False: blnNextRec = False: blnFromZero = False: blnWarnZeroCost = False
210         strPageMoveCtl = vbNullString: strSaveMoveCtl = vbNullString
220         blnRecsTotUpdate = False: blnReinvestment = False
230         .Parent.LastAcctNo = vbNullString: blnActNoTab = False  ' ** This one's 'cause I haven't figured out why it keeps going!
240         .Parent.ToTaxLot = 0&
250         .Parent.TaxLotFrom = vbNullString
260         lngNewJrnlColID = 0&
270         blnGTR_Emblem = False: lngGTR_Stat = 0&: lngGTR_ID = 0&
280         blnGoneToReport = False: blnGoneToReport2 = False: blnGTR_NoAdd = False

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
290         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

            ' ** NEW OPTION?
300         gblnLocSuggest = True

310         .AllowAdditions = False
320         .AllowDeletions = False

            ' ** Loop through all the security groups, checking if our current user is in the admins group.
330         gblnAdmin = False
340         For Each grp In DBEngine.Workspaces(0).Groups
350           If grp.Name = "Admins" Then
360             For Each usr In grp.Users
370               If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
380                 gblnAdmin = True
390                 Exit For
400               End If
410             Next
420           End If
430         Next

440         blnRetVal = CoOptions_Read  ' ** Module Function: modStartupFuncs.

450       Else
460         Cancel = -1
470       End If
480     End With

490     If Cancel = -1 Then
500       Beep
510       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
520     End If

EXITP:
530     Set usr = Nothing
540     Set grp = Nothing
550     Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database

        ' ** Copyright Lebans Holdings 1999 Ltd.
        ' ** Create an instance of the Calendar class
610     Set clsMonthClass = New clsMonthCal
        ' ** You MUST SET the class hWndForm prop!!!
620     clsMonthClass.hWndForm = Me.hwnd
        ' ** Let's default to PositionAtCursor.
630     clsMonthClass.PositionAtCursor = True

640     Set dbs = CurrentDb
650     With dbs
660       Set grstPostingDate = .OpenRecordset("PostingDate")
670       With grstPostingDate
680         .FindFirst "[Username] = '" & CurrentUser & "'"  ' ** Internal Access Function: Trust Accountant login.
690         If .NoMatch = False Then
700           datPostingDate = ![Posting_Date]
710         Else
              ' ** There should always be one by this time.
720           .AddNew
730           ![Posting_Date] = Date
740           ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
750           .Update
760           datPostingDate = Date
770         End If
780         .Close
790       End With  ' ** grstPostingDate.
800       .Close
810     End With  ' ** dbs.

820     With Me

830       StatusBar_Load  ' ** Procedure: Below.

840       Select Case gblnRevenueExpenseTracking
          Case True
850         .revcode_ID.Enabled = True
860         .revcode_ID.Locked = False
870         .revcode_DESC_display.Enabled = True
880         .revcode_DESC_display.Visible = True
890       Case False
900         .revcode_ID.Enabled = False
910         .revcode_ID.Locked = False  ' ** This we do want to look disabled.
920         .revcode_DESC_display.Visible = False
930         .revcode_DESC_display.Enabled = False
940         .revcode_ID_lbl.ForeColor = WIN_CLR_DISF
950         .revcode_ID_lbl_dim_hi.Visible = True
960         .revcode_ID_lbl_line.BorderColor = WIN_CLR_DISR
970         .revcode_ID_lbl_line_dim_hi.Visible = True
980       End Select

990       Select Case gblnIncomeTaxCoding
          Case True
1000        .taxcode.Enabled = True
1010        .taxcode.Locked = False
1020        .taxcode_description_display.Enabled = True
1030        .taxcode_description_display.Visible = True
1040      Case False
1050        .taxcode.Enabled = False
1060        .taxcode.Locked = False  ' ** This we do want to look disabled.
1070        .taxcode_description_display.Visible = False
1080        .taxcode_description_display.Enabled = False
1090        .taxcode_lbl.ForeColor = WIN_CLR_DISF
1100        .taxcode_lbl_dim_hi.Visible = True
1110        .taxcode_lbl_line.BorderColor = WIN_CLR_DISR
1120        .taxcode_lbl_line_dim_hi.Visible = True
1130      End Select

1140      .journal_USER.ForeColor = CLR_DISABLED_FG
1150      .journal_USER.BackColor = CLR_DISABLED_BG
1160      .journal_USER.Locked = True
1170      .journal_USER.Enabled = False

1180      If CLR_HILITE = 0& Then CLR_HILITE = CLR_AC07
1190      .ctlBack1.ForeColor = CLR_HILITE
1200      .ctlBack2.ForeColor = CLR_HILITE

1210      .JrnlCol_ID.ForeColor = CLR_DISABLED_FG
1220      .JrnlCol_ID.BackColor = CLR_DISABLED_BG

1230      JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

1240      blnRecsTotUpdate = True
1250      .TimerInterval = 3000&

1260    End With

EXITP:
1270    Set dbs = Nothing
1280    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Form_Current()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim ctl As Access.Control

1310    With Me

1320      .ctlCurrentRecord = .SelTop

          ' ** If this is a new record, locks are set by the thread set off by journaltype_AfterUpdate().
          ' ** Up until that time, all fields to the right of JournalType should be locked.
          ' ** Should we let them partially enter several records? Then let them complete them one-at-a-time?
          ' ** That could be convenient. But we'd have to have a mechanism to then commit an entry [NOW WE DO!].
          ' ** Non-committed entries don't go into the real Journal,
          ' ** but could they remain here until they're either committed or deleted?
          ' ** Again, I think that might be convenient for large batches of entries!

1330      .Parent.JrnlMemo_Memo = Null
1340      blnF4Invoked = False: blnF4InvokedMouse = False: strF4LastControl = vbNullString: blnActNoTab = False

1350      JC_Msc_Cur_Set Me, .Parent  ' ** Module Procedure: modJrnlCol_Misc.
          ' ** On a completely empty record (save transdate),
          ' ** I want to keep the top buttons active!

1360      strAccountNo_OldValue = vbNullString: strShortName_OldValue = vbNullString
1370      If IsNull(.accountno) = False Then
1380        If Trim(.accountno) <> vbNullString Then
1390          strAccountNo_OldValue = .accountno
1400          strShortName_OldValue = Nz(.accountno.Column(1), vbNullString)
1410        End If
1420      End If

1430      strJournalType_OldValue = vbNullString
1440      If IsNull(.journaltype) = False Then
1450        If Trim(.journaltype) <> vbNullString Then
1460          strJournalType_OldValue = .journaltype
1470          JC_Key_JType_Set strJournalType_OldValue, Me  ' ** Module Function: modJrnlCol_Keys.
1480        End If
1490      End If

1500      lngAssetNo_OldValue = -1&: strAssetDesc_OldValue = vbNullString
1510      If IsNull(.assetno) = False Then
1520        lngAssetNo_OldValue = .assetno
1530        If .assetno > 0& Then
1540          If IsNull(.assetno_description) = False Then
1550            strAssetDesc_OldValue = .assetno_description
1560          Else
                ' ** Might happen.
1570            strAssetDesc_OldValue = Nz(.assetno.Column(CBX_AST_DESC), vbNullString)
1580          End If
1590        End If
1600      End If

1610      dblShareFace_OldValue = 0#: dblICash_OldValue = 0@: dblPCash_OldValue = 0@

1620      If CLR_HILITE = 0& Then CLR_HILITE = CLR_AC07
1630      If .Parent.currow_hilite_forecolor <> CLR_HILITE Then .Parent.currow_hilite_forecolor = CLR_HILITE

1640      If .ctlBack1.ForeColor <> CLR_HILITE Then
1650        .ctlBack1.ForeColor = CLR_HILITE
1660      End If
1670      If .ctlBack2.ForeColor <> CLR_HILITE Then
1680        .ctlBack2.ForeColor = CLR_HILITE
1690      End If

1700      If strPageMoveCtl <> vbNullString Then
1710        If strPageMoveCtl = "assetno_description" Then strPageMoveCtl = "assetno"
1720        If strPageMoveCtl = "Loc_Name_display" Then strPageMoveCtl = "Location_ID"
1730        If strPageMoveCtl = "revcode_DESC_display" Then strPageMoveCtl = "revcode_ID"
1740        If strPageMoveCtl = "taxcode_description_display" Then strPageMoveCtl = "taxcode"
1750        If Left(strPageMoveCtl, 3) <> "cmd" Then
1760          .Controls(strPageMoveCtl).SetFocus
1770        Else
1780          DoCmd.SelectObject acForm, .Parent.Name, False
1790          .Parent.Controls(strPageMoveCtl).SetFocus
1800        End If
1810        strPageMoveCtl = vbNullString
1820      End If

1830      DoCmd.Hourglass False
1840      RecalcTots  ' ** Procedure: Below.

1850    End With

EXITP:
1860    Set ctl = Nothing
1870    Exit Sub

ERRH:
560     DoCmd.Hourglass False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intRetVal As Integer

1910    intRetVal = JC_Key_SubFrm(KeyCode, Shift, strSaveMoveCtl, blnF4Invoked, _
          blnF4InvokedMouse, strF4LastControl, blnNextRec, blnFromZero, Me)  ' ** Module Function: modJrnlCol_Keys.

EXITP:
1920    KeyCode = intRetVal
1930    Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2046  ' ** The command or action isn't available now (first or last record).
        'Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
600     Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

2010    With Me
2020  On Error Resume Next
2030      .JrnlCol_DateModified = Now()
2040  On Error GoTo ERRH
2050    End With

EXITP:
2060    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Form_Timer()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strThisJType As String
        Dim varTmp00 As Variant, strTmp01 As String, lngTmp02 As Long, lngTmp03 As Long

2110    With Me
2120      .TimerInterval = 0&

2130      Timer1_JC blnGoingToReport, blnGTR_NoAdd, blnGTR_Emblem, blnGoneToReport, blnGoneToReport2, _
            lngGTR_Stat, Me  ' ** Module Procedure: modJrnlCol_Timer.

2140      If blnRecsTotUpdate = True Then
2150        blnRecsTotUpdate = False
2160        gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2170  On Error Resume Next
2180        strThisJType = Nz(.journaltype, vbNullString)
2190        If ERR.Number = 0 Then
2200  On Error GoTo ERRH
2210          If .Filter = vbNullString Then
2220            lngTmp02 = 0&
2230            lngRecsCur = RecCnt  ' ** Function: Below.
2240            If lngRecsCur > 0& Then
2250              lngTmp02 = .RecsTot_Committed
2260            End If
2270            .Parent.RecsTot_Committed = Abs(lngTmp02)
2280            .Parent.RecsTot_Uncommitted = lngRecsCur - Abs(lngTmp02)
2290            If strThisJType = vbNullString Then
2300              JC_Btn_Set strThisJType, .posted, Me.Parent, True  ' ** Module Procedure: modJrnlCol_Buttons.
2310            Else
2320              JC_Btn_Set strThisJType, .posted, Me.Parent  ' ** Module Procedure: modJrnlCol_Buttons.
2330            End If
2340          Else
                ' ** Give full results, even if not seeing all of them.
2350            Select Case gblnAdmin
                Case True
2360              varTmp00 = DCount("*", "tblJournal_Column")
2370              If IsNull(varTmp00) = False Then
2380                lngTmp02 = varTmp00
2390              Else
2400                lngTmp02 = 0&
2410              End If
2420              varTmp00 = DCount("*", "tblJournal_Column", "[posted] = True")
2430              If IsNull(varTmp00) = False Then
2440                lngTmp03 = varTmp00
2450              Else
2460                lngTmp03 = 0&
2470              End If
2480            Case False
2490              varTmp00 = DCount("*", "tblJournal_Column", "[journal_USER] = '" & gstrJournalUser & "'")
2500              If IsNull(varTmp00) = False Then
2510                lngTmp02 = varTmp00
2520              Else
2530                lngTmp02 = 0&
2540              End If
2550              varTmp00 = DCount("*", "tblJournal_Column", "[posted] = True And [journal_USER] = '" & gstrJournalUser & "'")
2560              If IsNull(varTmp00) = False Then
2570                lngTmp03 = varTmp00
2580              Else
2590                lngTmp03 = 0&
2600              End If
2610            End Select
2620            .Parent.RecsTot_Committed = lngTmp03
2630            .Parent.RecsTot_Uncommitted = lngTmp02 - lngTmp03
2640            If strThisJType = vbNullString Then
2650              JC_Btn_Set strThisJType, .posted, Me.Parent, True  ' ** Module Procedure: modJrnlCol_Buttons.
2660            Else
2670              JC_Btn_Set strThisJType, .posted, Me.Parent  ' ** Module Procedure: modJrnlCol_Buttons.
2680            End If
2690          End If
2700        Else
2710  On Error GoTo ERRH
2720        End If
2730      ElseIf strPageMoveCtl <> vbNullString Then
2740        .Controls(strPageMoveCtl).SetFocus
2750        strPageMoveCtl = vbNullString
2760      ElseIf strSaveMoveCtl <> vbNullString Then

2770        strTmp01 = strSaveMoveCtl
2780        strSaveMoveCtl = vbNullString

2790        Timer2_JC blnJTypeSet, blnReinvestment, strSaveMoveCtl, Me  ' ** Module Procedure: modJrnlCol_Timer.

2800      End If  ' ** blnRecsTotUpdate, strPageMoveCtl, strSaveMoveCtl.
2810    End With  ' ** Me.

EXITP:
2820    Exit Sub

ERRH:
560     DoCmd.Hourglass False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2424  ' ** The expression you entered has a field, control, or property name that Microsoft Access can't find.
600     Resume EXITP

End Sub

Private Sub Form_Click()
' ** GetLineNumber() in modMiscFuncs.
' ** ctlBack1.ControlSource: =IIf([SelTop]=[ctlCurrentLine],String(255,"Û"),Null)
' ** ctlBack2.ControlSource:

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Click"

2910    With Me
2920      If IsNull(.JrnlCol_ID) = False Or .NewRecord = True Then
2930        .ctlCurrentRecord = .SelTop
2940      End If
2950    End With

EXITP:
2960    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2427  ' ** You entered an expression that has no value.
590     Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3010    If DataErr <> 0 Then
          'Debug.Print "'" & THIS_PROC & " 1  ERR: " & CStr(DataErr)
          'JC_Tst_ErrLog DataErr, IIf(That_Desc = vbNullString, ERR.description, That_Desc), THIS_NAME, _
          '  IIf(THAT_PROC = vbNullString, THIS_PROC, THAT_PROC), That_Erl ' ** Module Procedure: modJrnlCol_Test.
3020      Select Case DataErr
          Case 13  ' ** Type mismatch.
            ' ** A typo in a numeric field.
3030        Response = acDataErrContinue
3040      Case 2108  ' ** You must save the field before you execute the GoToControl action, the GoToControl method, or the SetFocus method.
            ' ** Do Nothing
3050        Response = acDataErrContinue
3060      Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
            ' ** Do nothing.
3070        Response = acDataErrContinue
3080      Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.
            ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
3090        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3100        Response = acDataErrContinue
3110      Case 2115  ' ** The macro or function set to the BeforeUpdate or ValidationRule property for
3120        Me.Undo  ' ** this field is preventing Microsoft Access from saving the data in the field.
3130        Response = acDataErrContinue
3140      Case 2116  ' ** The value in the field or record violates the validation rule for the record or field.
3150        Me.Undo
3160        Response = acDataErrDisplay
3170      Case 2237  ' ** The text you entered isn't an item in the list.
3180        MsgBox "You must select a valid item from the list.", vbInformation + vbOKOnly, "Invalid Entry"
3190        Response = acDataErrContinue
3200      Case 2473  ' ** The expression On Exit you entered as the event property setting produced the following error:
3210        Response = acDataErrContinue  ' ** There was a problem referencing a property or method of the object.
3220      Case 3020  ' ** Update or CancelUpdate without AddNew or Edit.
            ' ** And, incidentally, because DataErr is an integer, the large negative error
            ' ** number (-2147352567) causes an Overflow error (6) here if compared to DataErr.
3230        Response = acDataErrContinue
3240      Case 7787  ' ** Can't update, currently locked by another user.
            ' ** The mysterious 7787 error, {unknown}! A web search provided several descriptions,
            ' ** all similar to the one above, though none of them 'Official'.
3250        Me.Undo
3260        Response = acDataErrContinue
3270      Case 7878  ' ** The data has been changed. Another user edited this record and saved the changes before
3280        Response = acDataErrContinue  ' ** you attempted to save your changes. Re-edit the record.
3290      Case Else
3300        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3310        zErrorHandler THIS_NAME, THAT_PROC, CLng(DataErr), That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3320      End Select
3330    End If

3340    THAT_PROC = vbNullString
3350    That_Erl = 0&
3360    That_Desc = vbNullString

        ' **************************************************************************************************************************
        ' ** Response enumeration:
        ' **   Value   Constant             Description
        ' **   =====   ===================  ======================================================================================
        ' **     0     acDataErrContinue    Doesn't display the default message to the user. You can use this when you want to
        ' **                                display a custom message to the user. For example, the event procedure could display
        ' **                                a custom dialog box asking if the user wanted to save the new entry. If the response
        ' **                                is Yes, the event procedure would add the new entry to the list and set the Response
        ' **                                argument to acDataErrAdded. If the response is No, the event procedure would set the
        ' **                                Response argument to acDataErrContinue.
        ' **     1     acDataErrDisplay     Displays the default message to the user. You can use this when you don't want to
        ' **                                allow the user to add a new value to the combo box list. (Default)
        ' **     2     acDataErrAdded       Doesn't display a message to the user but enables you to add the entry to the combo
        ' **                                box list in the NotInList event procedure. After the entry is added, Microsoft
        ' **                                Access updates the list by requerying the combo box. Microsoft Access then rechecks
        ' **                                the string against the combo box list, and saves the value in the NewData argument
        ' **                                in the field the combo box is bound to. If the string is not in the list, then
        ' **                                Microsoft Access displays an error message.
        ' **************************************************************************************************************************

EXITP:
3370    Exit Sub

ERRH:
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3410    If Not clsMonthClass Is Nothing Then
3420      If clsMonthClass.IsCalendar Then
3430        Cancel = -1
3440      Else
3450        Set clsMonthClass = Nothing
3460      End If
3470    End If

3480    With Me
3490      .TimerInterval = 0&  ' ** Make sure!
3500      .OrderBy = vbNullString
3510      .OrderByOn = False
3520    End With

EXITP:
3530    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub cmdSave_Click()
' ** This is a fake click event!

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

        Dim strControl As String
        Dim lngRetVal As Long

3610    With Me
3620      .Refresh
3630  On Error Resume Next
3640      strControl = Screen.ActiveControl.Name
3650  On Error GoTo ERRH
3660      If strControl = "assetno_description" Then strControl = "assetno"
3670      If strControl = "Loc_Name_display" Then strControl = "Location_ID"
3680      If strControl = "revcode_DESC_display" Then strControl = "revcode_ID"
3690      If strControl = "taxcode_description_display" Then strControl = "taxcode"
3700      If IsNull(.accountno) = True Then
3710        If strAccountNo_OldValue = vbNullString Then
3720          .Undo
3730        End If
3740      End If
          ' ** Don't do an acCmdSelectRecord, because is slides the form!
3750      DoCmd.RunCommand acCmdSaveRecord
3760      If blnFromZero = False Then
3770        If strSaveMoveCtl <> vbNullString Then
3780          .TimerInterval = 100&
3790        ElseIf strControl <> vbNullString Then
              ' ** strControl is THIS field.
3800          If strControl <> "JrnlMemo_Memo" Then
3810            If blnReinvestment = True Then
                  'blnReinvestment = False  'WAIT, SINCE THEY MIGHT CHANGE THE ASSET, AND WE DON'T WANT TO LOSE THE AMOUNT!
3820              lngRetVal = fSetScrollBarPosHZ(Me, 1&)  ' ** Module Function: modScrollBarFuncs.
3830            End If
3840            Select Case blnNoMove
                Case True
3850              blnNoMove = False
3860            Case False
3870              If Left(strControl, 3) <> "cmd" Then
3880                .Controls(strControl).SetFocus
3890              Else
3900                DoCmd.SelectObject acForm, .Parent.Name, False
3910                .Parent.Controls(strControl).SetFocus
3920              End If
3930            End Select
3940          Else
3950            DoCmd.SelectObject acForm, .Parent.Name, False
3960            .Parent.JrnlMemo_Memo.SetFocus
3970          End If
3980        End If
3990      End If
4000    End With

EXITP:
4010    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2427  ' ** You entered an expression that has no value.
        'Case 2455, 3709, 7885  ' ** You entered an expression that has an invalid reference to the property MaxRecButton.
        'Case 2501  ' ** The RunCommand action was canceled.
        'Case 3020, -2147352567  ' ** Update or CancelUpdate without AddNew or Edit.
590     Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intRetVal As Integer

4110    intRetVal = JC_Key_FocusDown(KeyCode, Shift, THIS_PROC, blnNextRec, blnFromZero, Me)  ' ** Module Function: modJrnlCol_Keys.

EXITP:
4120    KeyCode = intRetVal
4130    Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
600     Resume EXITP

End Sub

Private Sub FocusHolder2_GotFocus()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder2_GotFocus"

4210    blnNoMove = False

EXITP:
4220    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub FocusHolder2_KeyDown(KeyCode As Integer, Shift As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder2_KeyDown"

        Dim intRetVal As Integer

4310    intRetVal = JC_Key_FocusDown(KeyCode, Shift, THIS_PROC, blnNextRec, blnFromZero, Me)  ' ** Module Function: modJrnlCol_Keys.

EXITP:
4320    KeyCode = intRetVal
4330    Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
600     Resume EXITP

End Sub

Private Sub posted_lbl_DblClick(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "posted_lbl_DblClick"

4410    JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
4420    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub posted_GotFocus()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "posted_GotFocus"

4510    JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
4520    blnNoMove = False

EXITP:
4530    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub posted_AfterUpdate()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "posted_AfterUpdate"

4610    JC_Ctl_Posted_Update Me, blnWarnZeroCost, strSaveMoveCtl, blnNoMove, blnNextRec, blnFromZero  ' ** Module Procedure: modJrnlCol_Controls.

EXITP:
4620    Exit Sub

ERRH:
560     Me.posted = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub posted_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "posted_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4710    intRetVal = KeyCode
4720    strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
4730    intShiftDown = (Shift And acShiftMask) > 0
4740    intAltDown = (Shift And acAltMask) > 0
4750    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4760    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4770      Select Case intRetVal
          Case vbKeyTab
4780        intRetVal = 0
4790        JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
4800      Case vbKeyReturn
4810        intRetVal = 0
4820        strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
4830      Case vbKeyPageUp, vbKeyPageDown
4840        strPageMoveCtl = "posted"
4850      End Select
4860    End If

        ' ** Shift keys.
4870    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4880      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4890        intRetVal = 0
4900        JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
4910      End Select
4920    End If

EXITP:
4930    KeyCode = intRetVal
4940    Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub JrnlCol_ID_lbl_DblClick(Cancel As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlCol_ID_lbl_DblClick"

5010    JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
5020    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub transdate_lbl_DblClick(Cancel As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

5110    JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
5120    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub transdate_GotFocus()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_GotFocus"

5210    JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
5220    blnNoMove = False

EXITP:
5230    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub transdate_BeforeUpdate(Cancel As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_BeforeUpdate"

5310    With Me
5320      If IsNull(.transdate) = True Then
            ' ** They must have opportunity to use calendar!
5330      Else
5340        If DateCheck_Post(.transdate) = False Then  ' ** Module Function: modUtilities.
5350          blnNoMove = True
5360          Cancel = -1
5370        End If
5380      End If
5390    End With

EXITP:
5400    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub transdate_AfterUpdate()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_AfterUpdate"

5510    With Me

5520      If .posted = True Then
5530        .posted = False
5540        .posted.Locked = False
5550        cmdSave_Click  ' ** Procedure: Above.
5560        DoEvents
5570      End If

5580      strSaveMoveCtl = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
5590      cmdSave_Click  ' ** Procedure: Above.

5600    End With

EXITP:
5610    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub transdate_KeyDown(KeyCode As Integer, Shift As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5710    intRetVal = KeyCode
5720    strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
5730    intShiftDown = (Shift And acShiftMask) > 0
5740    intAltDown = (Shift And acAltMask) > 0
5750    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5760    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5770      Select Case intRetVal
          Case vbKeyTab
            ' ** They must have opportunity to use calendar!
5780        Select Case blnNoMove
            Case True
5790          blnNoMove = False
5800        Case False
5810          With Me
5820            intRetVal = 0
5830            JC_Key_Sub Me, THIS_PROC, KeyCode, 2, blnNextRec, blnFromZero, blnToTaxLot, .Parent.LastAcctNo, "accountno"  ' ** Module Function: modJrnlCol_Keys.
5840            If IsNull(.accountno) = False Then
5850              strAccountNo_OldValue = .accountno
5860              strShortName_OldValue = .accountno.Column(1)
5870            End If
5880            If IsNull(.Parent.LastAcctNo) = False Then
5890              If .Parent.LastAcctNo <> vbNullString Then
5900                blnActNoTab = True
5910                blnNoMove = True  ' ** JC_Key_Sub() sets focus to journaltype.
5920              End If
5930            End If
5940          End With
5950        End Select
5960      Case vbKeyReturn
            ' ** They must have opportunity to use calendar!
5970        Select Case blnNoMove
            Case True
5980          blnNoMove = False
5990        Case False
6000          intRetVal = 0
6010          strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
6020        End Select
6030      Case vbKeyPageUp, vbKeyPageDown
6040        Select Case blnNoMove
            Case True
6050          blnNoMove = False
6060        Case False
6070          strPageMoveCtl = "transdate"
6080        End Select
6090      End Select
6100    End If

        ' ** Shift keys.
6110    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6120      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6130        Select Case blnNoMove
            Case True
6140          blnNoMove = False
6150        Case False
6160          intRetVal = 0
6170          JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
6180        End Select
6190      End Select
6200    End If

        ' ** Available form records info:
        ' **   Me.CurrentRecord  The value specified by this property corresponds to the value shown
        ' **                     in the record number box found in the lower-left corner of the form.
        ' **   Me.NewRecord      When a user has moved to a new record, the NewRecord property setting
        ' **                     will be True whether the user has started to edit the record or not.

EXITP:
6210    KeyCode = intRetVal
6220    Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub transdate_Exit(Cancel As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_Exit"

6310    With Me
          ' ** They must have opportunity to use calendar!
6320      If .transdate.Locked = False Then
6330        If .Dirty = True And blnReinvestment = False Then
6340          Select Case blnNoMove
              Case True
6350            blnNoMove = False
6360            strSaveMoveCtl = vbNullString
6370          Case False
6380            strSaveMoveCtl = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
6390          End Select
6400          cmdSave_Click  ' ** Procedure: Above.
6410        End If
6420      End If
6430    End With

EXITP:
6440    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

6510    JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
6520    blnNoMove = False

EXITP:
6530    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

6610    JC_Key_CalendarClick THIS_PROC, blnNoMove, blnNextRec, blnFromZero, _
          strSaveMoveCtl, clsMonthClass, Me  ' ** Module Procedure: modJrnlCol_Keys.

EXITP:
6620    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Access can't move the focus to the control '|'.
590     Resume EXITP

End Sub

Private Sub cmdCalendar1_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6710    intRetVal = KeyCode
6720    strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
6730    intShiftDown = (Shift And acShiftMask) > 0
6740    intAltDown = (Shift And acAltMask) > 0
6750    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6760    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6770      Select Case intRetVal
          Case vbKeyTab
6780        With Me
6790          Select Case blnNoMove
              Case True
6800            blnNoMove = False
6810          Case False
6820            intRetVal = 0
6830            If IsDate(.transdate) = True Then
6840              JC_Key_Sub Me, THIS_PROC, KeyCode, 2, blnNextRec, blnFromZero, blnToTaxLot, .Parent.LastAcctNo, "accountno"  ' ** Module Function: modJrnlCol_Keys.
6850            Else
6860              Beep
6870              MsgBox "A Posting Date is required to continue.", vbInformation + vbOKOnly, "Entry Required  1"
6880              .transdate.SetFocus
6890            End If
6900          End Select
6910        End With
6920      Case vbKeyPageUp, vbKeyPageDown
6930        Select Case blnNoMove
            Case True
6940          blnNoMove = False
6950        Case False
6960          strPageMoveCtl = "transdate"
6970        End Select
6980      End Select
6990    End If

        ' ** Shift keys.
7000    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7010      Select Case intRetVal
          Case vbKeyTab
7020        With Me
7030          Select Case blnNoMove
              Case True
7040            blnNoMove = False
7050          Case False
7060            intRetVal = 0
7070            .transdate.SetFocus
7080          End Select
7090        End With
7100      End Select
7110    End If

EXITP:
7120    KeyCode = intRetVal
7130    Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub cmdCalendar1_Exit(Cancel As Integer)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Exit"

7210    With Me
7220      If IsNull(.transdate) = True Then
7230        If gblnClosing = False And gblnDeleting = False Then
7240          Beep
7250          MsgBox "A Posting Date is required to continue.", vbInformation + vbOKOnly, "Entry Required  3"
7260          Cancel = -1
7270          .transdate.SetFocus
7280        End If
7290      End If
7300    End With

EXITP:
7310    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

7410    JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
7420    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub accountno_Enter()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_Enter"

7510    With Me
7520      If IsNull(.transdate) = True Then
7530        Beep
7540        MsgBox "A Posting Date is required to continue.", vbInformation + vbOKOnly, "Entry Required  4"
7550        blnNoMove = True
7560        .transdate.SetFocus
7570      ElseIf IsNull(.accountno) = False Then
7580        If Trim(.accountno) <> vbNullString Then
7590          strAccountNo_OldValue = .accountno
7600          strShortName_OldValue = .accountno.Column(1)
7610        Else
7620          strAccountNo_OldValue = vbNullString
7630          strShortName_OldValue = vbNullString
7640        End If
7650      Else
7660        strAccountNo_OldValue = vbNullString
7670        strShortName_OldValue = vbNullString
7680        If gblnGoToReport = False Then
7690          blnF4Invoked = True: strF4LastControl = "accountno"
7700          ExpandCombo .accountno  ' ** Module Procedure: modUtilities.
7710        End If
7720      End If
7730    End With

EXITP:
7740    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub accountno_GotFocus()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_GotFocus"

7810    JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
7820    blnNoMove = False

EXITP:
7830    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub accountno_BeforeUpdate(Cancel As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_BeforeUpdate"

        Dim blnContinue As Boolean

7910    With Me

7920      blnContinue = True

7930      JCol_AcctNo_BeforeUpdate Cancel, blnContinue, strAccountNo_OldValue, strShortName_OldValue, _
            THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

7940    End With

EXITP:
7950    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub accountno_AfterUpdate()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_AfterUpdate"

        Dim blnDelete As Boolean

8010    With Me

8020      blnDelete = False

8030      JCol_AcctNo_AfterUpdate blnDelete, strAccountNo_OldValue, strShortName_OldValue, strSaveMoveCtl, _
            blnNextRec, blnFromZero, blnNoMove, THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

8040    End With

EXITP:
8050    Exit Sub

ERRH:
560     DoCmd.Hourglass False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub accountno_NotInList(NewData As String, Response As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_NotInList"

8110    With Me
8120      If gblnClosing = False And gblnDeleting = False Then
8130        .accountno = Null
8140        blnNoMove = True
8150        MsgBox "You must select a valid Account from the list.", vbInformation + vbOKOnly, "Invalid Entry"
8160        Response = acDataErrContinue
8170      End If
8180    End With

EXITP:
8190    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     If ERR.Number > 0 Then
590       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Else  ' ** Those huge negative error numbers are too big for the Integer in the Form_Error() parameters.
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End If
630     Resume EXITP

End Sub

Private Sub accountno_KeyDown(KeyCode As Integer, Shift As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8210    intRetVal = KeyCode
8220    strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
8230    intShiftDown = (Shift And acShiftMask) > 0
8240    intAltDown = (Shift And acAltMask) > 0
8250    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8260    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8270      Select Case intRetVal
          Case vbKeyTab
8280        Select Case blnNoMove
            Case True
8290          blnNoMove = False
8300        Case True
8310          With Me
8320            intRetVal = 0
8330            JC_Key_Sub Me, THIS_PROC, KeyCode, 2, blnNextRec, blnFromZero, blnToTaxLot, .Parent.LastAcctNo, "accountno2"  ' ** Module Function: modJrnlCol_Keys.
8340            If IsNull(.accountno) = False Then
8350              strAccountNo_OldValue = .accountno
8360              strShortName_OldValue = .accountno.Column(1)
8370            End If
8380          End With
8390        End Select
8400      Case vbKeyReturn
8410        Select Case blnNoMove
            Case True
8420          blnNoMove = False
8430        Case False
8440          intRetVal = 0
8450          strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
8460        End Select
8470      Case vbKeyPageUp, vbKeyPageDown
8480        Select Case blnNoMove
            Case True
8490          blnNoMove = False
8500        Case False
8510          strPageMoveCtl = "accountno"
8520        End Select
8530      End Select
8540    End If

        ' ** Shift keys.
8550    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8570        Select Case blnNoMove
            Case True
8580          blnNoMove = False
8590        Case False
8600          intRetVal = 0
8610          JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
8620        End Select
8630      End Select
8640    End If

EXITP:
8650    KeyCode = intRetVal
8660    Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub accountno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_MouseDown"

8710    blnF4InvokedMouse = True  ' ** Maybe, maybe not.

EXITP:
8720    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub accountno_Exit(Cancel As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_Exit"

        ' ** They must have opportunity to use accountno2!
8810    blnF4InvokedMouse = False

EXITP:
8820    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub accountno2_lbl_DblClick(Cancel As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno2_lbl_DblClick"

8910    JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
8920    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub accountno2_Enter()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno2_Enter"

9010    With Me
9020      If IsNull(.transdate) = True Then
9030        Beep
9040        MsgBox "A Posting Date is required to continue.", vbInformation + vbOKOnly, "Entry Required  5"
9050        blnNoMove = True
9060        .transdate.SetFocus
9070      ElseIf IsNull(.accountno2) = False Then
9080        If Trim(.accountno2) <> vbNullString Then
9090          strAccountNo_OldValue = .accountno2
9100          strShortName_OldValue = .accountno2.Column(1)
9110        Else
9120          strAccountNo_OldValue = vbNullString
9130          strShortName_OldValue = vbNullString
9140        End If
9150      Else
9160        strAccountNo_OldValue = vbNullString
9170        strShortName_OldValue = vbNullString
9180        If gblnGoToReport = False Then
9190          blnF4Invoked = True: strF4LastControl = "accountno2"
9200          ExpandCombo .accountno2  ' ** Module Procedure: modUtilities.
9210        End If
9220      End If
9230    End With

EXITP:
9240    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub accountno2_GotFocus()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno2_GotFocus"

9310    JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
9320    blnNoMove = False

EXITP:
9330    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub accountno2_BeforeUpdate(Cancel As Integer)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno2_BeforeUpdate"

        Dim blnContinue As Boolean

9410    With Me

9420      blnContinue = True

9430      JCol_AcctNo2_BeforeUpdate Cancel, blnContinue, strAccountNo_OldValue, strShortName_OldValue, _
            blnNoMove, THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

9440    End With

EXITP:
9450    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub accountno2_AfterUpdate()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno2_AfterUpdate"

        Dim blnDelete As Boolean

9510    With Me

9520      blnDelete = False

9530      JCol_AcctNo2_AfterUpdate blnDelete, strAccountNo_OldValue, strShortName_OldValue, strSaveMoveCtl, _
            blnNextRec, blnFromZero, THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

9540    End With

EXITP:
9550    Exit Sub

ERRH:
560     DoCmd.Hourglass False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub accountno2_NotInList(NewData As String, Response As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno2_NotInList"

9610    With Me
9620      If gblnClosing = False And gblnDeleting = False Then
9630        .accountno2 = Null
9640        blnNoMove = True
9650        MsgBox "You must select a valid Account from the list.", vbInformation + vbOKOnly, "Invalid Entry"
9660        Response = acDataErrContinue
9670      End If
9680    End With

EXITP:
9690    Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     If ERR.Number > 0 Then
590       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Else  ' ** Those huge negative error numbers are too big for the Integer in the Form_Error() parameters.
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End If
630     Resume EXITP

End Sub

Private Sub accountno2_KeyDown(KeyCode As Integer, Shift As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9710    intRetVal = KeyCode
9720    strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
9730    intShiftDown = (Shift And acShiftMask) > 0
9740    intAltDown = (Shift And acAltMask) > 0
9750    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9760    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9770      Select Case intRetVal
          Case vbKeyTab
9780        With Me
9790          intRetVal = 0
9800          If IsNull(.accountno2) = False Then
9810            JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
9820          Else
9830            If IsNull(.accountno2.text) = True Then
9840              Beep
9850              MsgBox "An Account is required to continue.", vbInformation + vbOKOnly, "Entry Required  2"
9860            Else
9870              JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
9880            End If
9890          End If
9900        End With
9910      Case vbKeyReturn
9920        With Me
9930          intRetVal = 0
9940          If IsNull(.accountno2) = False Then
9950            strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
9960          Else
9970            If IsNull(.accountno2.text) = True Then
9980              Beep
9990              MsgBox "An Account is required to continue.", vbInformation + vbOKOnly, "Entry Required  3"
10000           Else
10010             strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
10020           End If
10030         End If
10040       End With
10050     Case vbKeyPageUp, vbKeyPageDown
10060       strPageMoveCtl = "accountno2"
10070     End Select
10080   End If

        ' ** Shift keys.
10090   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10100     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10110       intRetVal = 0
10120       JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
10130     End Select
10140   End If

EXITP:
10150   KeyCode = intRetVal
10160   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub accountno2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "accountno2_MouseDown"

10210   blnF4InvokedMouse = True  ' ** Maybe, maybe not.

EXITP:
10220   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub accountno2_Exit(Cancel As Integer)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "accountno2_Exit"

10310   With Me
10320     blnF4InvokedMouse = False
10330     If IsNull(.accountno2) = True Then
10340       If IsNull(.accountno) = True Then
10350         If gblnClosing = False And gblnDeleting = False Then
10360           Beep
10370           MsgBox "An Account is required to continue.", vbInformation + vbOKOnly, "Entry Required  4"
10380           Cancel = -1
10390           .accountno.SetFocus
10400         End If
10410       Else
10420         .accountno2 = .accountno
10430       End If
10440     Else
10450       If IsNull(.accountno) = True Then
10460         .accountno = .accountno2
10470       End If
10480     End If
10490   End With

EXITP:
10500   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub journaltype_lbl_DblClick(Cancel As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_lbl_DblClick"

10610   JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
10620   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub journaltype_Enter()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_Enter"

10710   With Me
10720     If IsNull(.transdate) = True Then
10730       Beep
10740       MsgBox "A Posting Date is required to continue.", vbInformation + vbOKOnly, "Entry Required  6"
10750       blnNoMove = True
10760       .transdate.SetFocus
10770     ElseIf IsNull(.accountno) = True Then
10780       Beep
10790       MsgBox "An Account is required to continue.", vbInformation + vbOKOnly, "Entry Required  5"
10800       blnNoMove = True
10810       .accountno.SetFocus
10820     Else
10830       If gblnGoToReport = False Then
10840         If IsNull(.journaltype) = True Then
10850           blnF4Invoked = True: strF4LastControl = "journaltype"
10860           ExpandCombo .journaltype  ' ** Module Procedure: modUtilities.
10870         Else
10880           strJournalType_OldValue = .journaltype
10890         End If
10900       End If
10910     End If
10920   End With

EXITP:
10930   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub journaltype_GotFocus()

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_GotFocus"

11010   JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
11020   blnNoMove = False

EXITP:
11030   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub journaltype_BeforeUpdate(Cancel As Integer)

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_BeforeUpdate"

        Dim blnContinue As Boolean

11110   With Me

11120     blnContinue = True

11130     JCol_JType_BeforeUpdate Cancel, blnContinue, strJournalType_OldValue, blnNoMove, _
            THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

11140   End With

EXITP:
11150   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub journaltype_AfterUpdate()

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_AfterUpdate"

11210   With Me

11220     JCol_JType_AfterUpdate blnJTypeSet, strSaveMoveCtl, blnWarnZeroCost, blnReinvestment, _
            blnNextRec, blnFromZero, blnNoMove, THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

11230   End With

        'WHY ALL LOCKED? IT SHOULDN'T BE LOCKED!
        'HERE! 4  journaltype_KeyDown()  True
        'HERE! 1  journaltype_AfterUpdate()
        'HERE! 2  journaltype_AfterUpdate()  True
        'HERE! 1  journaltype_Exit()  True
        'HERE! 1  assetno_description_GotFocus()  True
        'HERE! 2  assetno_description_GotFocus()  True
        'HERE! 3  assetno_description_GotFocus()  True
        'HERE! 1  assetno_Enter()  True
        '** BECAUSE IT WASN'T PASSING THIS FIELD NAME, THIS_PROC!

EXITP:
11240   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub journaltype_NotInList(NewData As String, Response As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_NotInList"

11310   With Me
11320     If gblnClosing = False And gblnDeleting = False Then
11330       .journaltype = Null
11340       blnNoMove = True
11350       MsgBox "You must select a valid Journal Type from the list.", vbInformation + vbOKOnly, "Invalid Entry"
11360       Response = acDataErrContinue
11370     End If
11380   End With

EXITP:
11390   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     If ERR.Number > 0 Then
590       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Else  ' ** Those huge negative error numbers are too big for the Integer in the Form_Error() parameters.
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End If
630     Resume EXITP

End Sub

Private Sub JournalType_KeyDown(KeyCode As Integer, Shift As Integer)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11410   intRetVal = KeyCode
11420   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
11430   intShiftDown = (Shift And acShiftMask) > 0
11440   intAltDown = (Shift And acAltMask) > 0
11450   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11460   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11470     Select Case intRetVal
          Case vbKeyTab
11480       With Me
11490         Select Case blnNoMove
              Case True
11500           blnNoMove = False
11510         Case False
11520           intRetVal = 0
11530           If IsNull(.journaltype) = False Then
11540             JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
11550           Else
11560             If IsNull(.journaltype.text) = True Then
11570               Beep
11580               MsgBox "A Journal Type is required to continue.", vbInformation + vbOKOnly, "Entry Required  3"
11590             Else
11600               JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
11610             End If
11620           End If
11630         End Select
11640       End With
11650     Case vbKeyReturn
11660       With Me
11670         Select Case blnNoMove
              Case True
11680           blnNoMove = False
11690         Case False
11700           intRetVal = 0
11710           If IsNull(.journaltype) = False Then
11720             strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
11730           Else
11740             If IsNull(.journaltype.text) = True Then
11750               Beep
11760               MsgBox "A Journal Type is required to continue.", vbInformation + vbOKOnly, "Entry Required  4"
11770             Else
11780               strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
11790             End If
11800           End If
11810         End Select
11820       End With
11830     Case vbKeyPageUp, vbKeyPageDown
11840       Select Case blnNoMove
            Case True
11850         blnNoMove = False
11860       Case False
11870         strPageMoveCtl = "journaltype"
11880       End Select
11890     End Select
11900   End If

        ' ** Shift keys.
11910   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11920     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11930       Select Case blnNoMove
            Case True
11940         blnNoMove = False
11950       Case False
11960         intRetVal = 0
11970         JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
11980       End Select
11990     End Select
12000   End If

EXITP:
12010   KeyCode = intRetVal
12020   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub journaltype_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_MouseDown"

12110   blnF4InvokedMouse = True  ' ** Maybe, maybe not.

EXITP:
12120   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub journaltype_Exit(Cancel As Integer)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_Exit"

12210   blnF4InvokedMouse = False

12220   With Me
12230     If IsNull(.journaltype) = True Then
12240       If gblnClosing = False And gblnDeleting = False Then
12250         If blnActNoTab = True Then
                ' ** Stay here without the message!
                ' ** I haven't figured out why it wants to keep going!
12260           blnActNoTab = False
12270           Cancel = -1
12280         Else
12290           If blnGoingToReport = False Then
12300             Beep
12310             MsgBox "A Journal Type is required to continue.", vbInformation + vbOKOnly, "Entry Required  5"
12320           End If
12330           Cancel = -1
12340         End If
12350       End If
12360     Else
            ' ** I can't figure out why assetno is locked, and/or where it's supposed to get unlocked!
12370       Select Case .journaltype
            Case "Deposit", "Purchase", "Withdrawn", "Sold"
12380         .assetno.Locked = False
12390       End Select
12400     End If
12410   End With

EXITP:
12420   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub chkShowAllAssets_GotFocus()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAllAssets_GotFocus"

12510   With Me
12520     JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
12530     blnNoMove = False
12540     .chkShowAllAssets_lbl_dots.Visible = True
12550   End With

EXITP:
12560   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub chkShowAllAssets_AfterUpdate()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAllAssets_AfterUpdate"

12610   With Me
12620     Select Case .chkShowAllAssets
          Case True
12630       .chkShowAllAssets_lbl.FontBold = True
12640     Case False
12650       .chkShowAllAssets_lbl.FontBold = False
12660     End Select
12670     JC_Msc_Asset_Set Me  ' ** Module Procedure: modJrnlCol_Misc.
12680     DoEvents
12690     .assetno.SetFocus
12700   End With

EXITP:
12710   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub chkShowAllAssets_KeyDown(KeyCode As Integer, Shift As Integer)

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAllAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12810   intRetVal = KeyCode
12820   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
12830   intShiftDown = (Shift And acShiftMask) > 0
12840   intAltDown = (Shift And acAltMask) > 0
12850   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12860   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12870     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12880       With Me
12890         intRetVal = 0
12900         .assetno.SetFocus
12910       End With
12920     End Select
12930   End If

        ' ** Shift keys.
12940   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12950     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12960       With Me
12970         intRetVal = 0
12980         .assetno.SetFocus  ' ** Wherever they wanna go, it's to assetno!
12990       End With
13000     End Select
13010   End If

EXITP:
13020   KeyCode = intRetVal
13030   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
        'Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
600     Resume EXITP

End Sub

Private Sub chkShowAllAssets_LostFocus()

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAllAssets_LostFocus"

13110   With Me
13120     .chkShowAllAssets_lbl_dots.Visible = False
13130   End With

EXITP:
13140   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub assetno_description_GotFocus()
' ** This control sits on top of the assetno combo box.
' ** It's used to display the MasterAsset description.
' ** CONTROL IS ENABLED, BUT LOCKED.

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_description_GotFocus"

13210   With Me
13220     JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
13230     blnNoMove = False
13240     If .assetno.ListCount <= 1 Then
13250       JC_Msc_Asset_Set Me  ' ** Module Procedure: modJrnlCol_Misc.
13260       DoEvents
13270     End If
13280     .assetno.SetFocus
13290   End With

EXITP:
13300   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub assetno_description_KeyDown(KeyCode As Integer, Shift As Integer)
' ** Ordinarily, they shouldn't have time to press a key while on this field!

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_description_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13410   intRetVal = KeyCode
13420   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
13430   intShiftDown = (Shift And acShiftMask) > 0
13440   intAltDown = (Shift And acAltMask) > 0
13450   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13460   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13470     Select Case intRetVal
          Case vbKeyTab
13480       With Me
13490         intRetVal = 0
13500         If .assetno.Locked = False Then
13510           .assetno.SetFocus
13520         Else
13530           JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
13540         End If
13550       End With
13560     Case vbKeyReturn
13570       With Me
13580         intRetVal = 0
13590         If .assetno.Locked = False Then
13600           .assetno.SetFocus
13610         Else
13620           strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
13630         End If
13640       End With
13650     Case vbKeyPageUp, vbKeyPageDown
13660       strPageMoveCtl = "assetno"
13670     End Select
13680   End If

        ' ** Shift keys.
13690   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13700     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13710       intRetVal = 0
13720       JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
13730     End Select
13740   End If

EXITP:
13750   KeyCode = intRetVal
13760   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub assetno_lbl_DblClick(Cancel As Integer)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_lbl_DblClick"

13810   JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
13820   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub assetno_Enter()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_Enter"

13910   With Me
13920     If IsNull(.transdate) = True Then
13930       Beep
13940       MsgBox "A Posting Date is required to continue.", vbInformation + vbOKOnly, "Entry Required  7"
13950       blnNoMove = True
13960       .transdate.SetFocus
13970     ElseIf IsNull(.accountno) = True Then
13980       Beep
13990       MsgBox "An Account is required to continue.", vbInformation + vbOKOnly, "Entry Required  6"
14000       blnNoMove = True
14010       .accountno.SetFocus
14020     ElseIf IsNull(.journaltype) = True Then
14030       Beep
14040       MsgBox "A Journal Type is required to continue.", vbInformation + vbOKOnly, "Entry Required  6"
14050       blnNoMove = True
14060       .journaltype.SetFocus
14070     ElseIf IsNull(.assetno) = False Then
14080       lngAssetNo_OldValue = .assetno
14090       If .assetno > 0& Then
14100         If IsNull(.assetno_description) = False Then
14110           strAssetDesc_OldValue = .assetno_description
14120         Else
                ' ** I don't know. It might happen.
14130           strAssetDesc_OldValue = .assetno.Column(CBX_AST_DESC)
14140         End If
14150       Else
14160         strAssetDesc_OldValue = vbNullString
14170       End If
14180     Else
14190       lngAssetNo_OldValue = -1&
14200       strAssetDesc_OldValue = vbNullString
14210       If .assetno.ListCount <= 1 Then
14220         JC_Msc_Asset_Set Me  ' ** Module Procedure: modJrnlCol_Misc.
14230         DoEvents
14240       End If
14250       If gblnGoToReport = False Then
14260         blnF4Invoked = True: strF4LastControl = "assetno"
14270         ExpandCombo .assetno  ' ** Module Procedure: modUtilities.
14280       End If
14290     End If
14300   End With

EXITP:
14310   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub assetno_GotFocus()

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_GotFocus"

14410   With Me
14420     JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
14430     blnNoMove = False
14440     If .chkShowAllAssets.Visible = True Then
14450       .chkShowAllAssets.Locked = False
14460     End If
14470   End With

EXITP:
14480   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub assetno_AfterUpdate()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_AfterUpdate"

        Dim blnDelete As Boolean

14510   With Me

14520     blnDelete = False

14530     JCol_Astno_AfterUpdate blnDelete, blnReinvestment, lngAssetNo_OldValue, strAssetDesc_OldValue, _
            strSaveMoveCtl, blnNextRec, blnFromZero, THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

14540   End With

EXITP:
14550   Exit Sub

ERRH:
560     DoCmd.Hourglass False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub assetno_NotInList(NewData As String, Response As Integer)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_NotInList"

14610   With Me
14620     If gblnClosing = False And gblnDeleting = False Then
14630       .assetno = Null
14640       blnNoMove = True
14650       MsgBox "You must select a valid Asset from the list.", vbInformation + vbOKOnly, "Invalid Entry"
14660       Response = acDataErrContinue
14670     End If
14680   End With

EXITP:
14690   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     If ERR.Number > 0 Then
590       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Else  ' ** Those huge negative error numbers are too big for the Integer in the Form_Error() parameters.
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End If
630     Resume EXITP

End Sub

Private Sub assetno_KeyDown(KeyCode As Integer, Shift As Integer)

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14710   intRetVal = KeyCode
14720   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
14730   intShiftDown = (Shift And acShiftMask) > 0
14740   intAltDown = (Shift And acAltMask) > 0
14750   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14760   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14770     Select Case intRetVal
          Case vbKeyTab
14780       intRetVal = 0
14790       JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
14800     Case vbKeyReturn
14810       intRetVal = 0
14820       strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
14830     Case vbKeyPageUp, vbKeyPageDown
14840       strPageMoveCtl = "assetno"
14850     End Select
14860   End If

        ' ** Shift keys.
14870   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14880     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14890       intRetVal = 0
14900       JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
14910     End Select
14920   End If

EXITP:
14930   KeyCode = intRetVal
14940   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub assetno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_MouseDown"

15010   blnF4InvokedMouse = True  ' ** Maybe, maybe not.

EXITP:
15020   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub assetno_Exit(Cancel As Integer)

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_Exit"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccountNo As String, lngAssetNo As Long, dblShareface As Double

15110   With Me
15120     If blnReinvestment = True Then
15130       blnReinvestment = False  'NOW WE SHOULD BE SAFE IN TURNING THIS OFF!
15140       blnJTypeSet = False  ' ** Still haven't figured out why this is True!
15150     End If
15160     blnF4InvokedMouse = False
15170     If IsNull(.accountno) = False And IsNull(.assetno) = False And IsNull(.journaltype) = False Then
15180       If .assetno > 0& And (.journaltype = "Dividend" Or .journaltype = "Interest") And .posted = False Then

15190         .assetdate = Now()
15200         .assetdate_display = CDate(Format(.assetdate, "mm/dd/yyyy"))
15210         strAccountNo = .accountno
15220         lngAssetNo = .assetno
15230         dblShareface = 0#

15240         Set dbs = CurrentDb
15250         With dbs
                ' ** ActiveAssets, grouped and summed, by accountno, assetno, by specified [actno], [astno].
15260           Set qdf = .QueryDefs("qryJournal_Columns_29")
15270           With qdf.Parameters
15280             ![actno] = strAccountNo
15290             ![astno] = lngAssetNo
15300           End With
15310           Set rst = qdf.OpenRecordset
15320           With rst
15330             If .BOF = True And .EOF = True Then
                    ' ** Shouldn't happen!
15340             Else
15350               .MoveFirst
15360               dblShareface = ![sumsf]
15370             End If
15380             .Close
15390           End With
15400           .Close
15410         End With

15420         .shareface = dblShareface

15430         strSaveMoveCtl = vbNullString 'JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
15440         cmdSave_Click  ' ** Procedure: Below.

15450       ElseIf blnReinvestment = True Then
15460         strSaveMoveCtl = vbNullString 'JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
15470         cmdSave_Click  ' ** Procedure: Below.
15480       End If
15490     End If
15500   End With

EXITP:
15510   Set rst = Nothing
15520   Set qdf = Nothing
15530   Set dbs = Nothing
15540   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Recur_Name_lbl_DblClick(Cancel As Integer)

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_lbl_DblClick"

15610   JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
15620   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Recur_Name_Enter()
' ** Unlike Trust Import, in Trust Accountant we only show table items in the
' ** dropdown, rather than including other manually entered descriptions.

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_Enter"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strThisJType As String
        Dim blnNoRowSource As Boolean

15710   blnNoRowSource = False

15720   With Me
15730     strThisJType = Nz(.journaltype, vbNullString)
15740     Select Case strThisJType
          Case "Misc.", "Paid", "Received"
15750       If IsNull(.Recur_Name.RowSource) = True Then
15760         blnNoRowSource = True
15770       Else
15780         If .Recur_Name.RowSource = vbNullString Then blnNoRowSource = True
15790       End If
15800       If blnNoRowSource = True Then  ' ** Unlikely.
              ' ** RecurringItems; Recur_Name, Recur_Type, RecurringItem_ID; By specified FormRef().
15810         .Recur_Name.RowSource = "qryJournal_Columns_10_RecurringItems_05"
15820         .Recur_Name.Requery
15830       Else
15840         If Recur_Name.ListCount <= 1 Then
                ' ** If this is a 'Paid' or 'Received', it's possible there are no entries yet!
15850           Set dbs = CurrentDb
15860           With dbs
                  ' ** Union of qryJournal_Columns_10_RecurringItems_06 (RecurringItems, by specified FormRef('Recur_Type')),
                  ' ** qryJournal_Columns_10_RecurringItems_07 (tblJournal_Column, grouped by Recur_Name, by specified FormRef('Recur_Type')),
                  ' ** qryJournal_Columns_10_RecurringItems_08 (Journal, grouped by Recur_Name, by specified FormRef('Recur_Type')),
                  ' ** qryJournal_Columns_10_RecurringItems_09 (Ledger, grouped by Recur_Name, by specified FormRef('Recur_Type')).
15870             Set qdf = .QueryDefs("qryJournal_Columns_10_RecurringItems_10")
15880             Set rst = qdf.OpenRecordset
15890             With rst
15900               If .BOF = True And .EOF = True Then
15910                 If strThisJType = "Paid" Or strThisJType = "Received" Then
15920                   MsgBox "There are no existing 'Payee' or 'Payor' Recurring Items." & vbCrLf & vbCrLf & _
                          "You may type in a description, or click the Recurring Items" & vbCrLf & _
                          "button at the bottom of the window.", vbInformation + vbOKOnly, "Dropdown List Is Empty"
15930                 Else
15940                   Beep
15950                   MsgBox "Default Miscellaneous Recurring Items were not found!" & vbCrLf & vbCrLf & _
                          "Close the Import File window and reenter." & vbCrLf & _
                          "If the problem persists, contact Delta Data, Inc., for assistance.", _
                          vbCritical + vbOKOnly, "Default Entries Not Found"
15960                 End If
15970               Else
                      ' ** RecurringItems; Recur_Name, Recur_Type, RecurringItem_ID; By specified FormRef('Recur_Type').
15980                 Me.Recur_Name.RowSource = "qryJournal_Columns_10_RecurringItems_05"
15990                 Me.Recur_Name.Requery
16000               End If
16010               .Close
16020             End With  ' ** rst.
16030             .Close
16040           End With  ' ** dbs.
                ' ** From modQueryFunctions1:
                ' **   Case "RecurType"
                ' **     Select Case Forms("frmJournal_Columns").frmJournal_Columns_Sub.Form.journaltype
                ' **     Case "Paid"
                ' **       varRetVal = "Payee"
                ' **     Case "Received"
                ' **       varRetVal = "Payor"
                ' **     Case "Misc."
                ' **       varRetVal = "Misc"
                ' **     Case Else
                ' **       varRetVal = RET_ERR
                ' **     End Select
16050         End If  ' ** ListCount.
16060         If IsNull(.Recur_Name) = True Then
16070           blnF4Invoked = True: strF4LastControl = "Recur_Name"
16080           ExpandCombo .Recur_Name  ' ** Module Procedure: modUtilities.
16090         End If
16100       End If  ' ** blnNoRowSource.
16110     Case Else
            ' ** They shouldn't be here!
16120     End Select
16130   End With

EXITP:
16140   Set rst = Nothing
16150   Set qdf = Nothing
16160   Set dbs = Nothing
16170   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Recur_Name_GotFocus()

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_GotFocus"

16210   JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
16220   blnNoMove = False

EXITP:
16230   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub Recur_Name_AfterUpdate()

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_AfterUpdate"

        Dim strThisJType As String
        Dim varTmp00 As Variant

16310   With Me

16320     If .posted = True Then
16330       .posted = False
16340       .posted.Locked = False
16350     End If

16360     .MiscXFer_ItoP_lbl.Visible = False
16370     .MiscXFer_PtoI_lbl.Visible = False

16380     If IsNull(.Recur_Name) = False Then
16390       If Trim(.Recur_Name) <> vbNullString Then
16400         varTmp00 = .Recur_Name.Column(2)  ' ** Recur_Name is Column(0).
16410         If IsNull(varTmp00) = False Then
16420           If IsNumeric(varTmp00) = True Then
16430             If CLng(varTmp00) > 0& Then
16440               .RecurringItem_ID = CLng(varTmp00)
16450               Select Case .RecurringItem_ID
                    Case RECUR_I_TO_P_ID  ' ** RECUR_I_TO_P is text.
16460                 .MiscXFer_ItoP_lbl.Visible = True
16470                 .MiscXFer_PtoI_lbl.Visible = False
16480               Case RECUR_P_TO_I_ID  ' ** RECUR_P_TO_I is text.
16490                 .MiscXFer_ItoP_lbl.Visible = False
16500                 .MiscXFer_PtoI_lbl.Visible = True
16510               Case Else
16520                 .MiscXFer_ItoP_lbl.Visible = False
16530                 .MiscXFer_PtoI_lbl.Visible = False
16540               End Select
16550             Else
16560               .RecurringItem_ID = Null
16570             End If
16580           Else
16590             .RecurringItem_ID = Null
16600           End If
16610         Else
16620           .RecurringItem_ID = Null
16630         End If
16640       Else
16650         .Recur_Name = Null
16660         .RecurringItem_ID = Null
16670       End If
16680     Else
16690       .RecurringItem_ID = Null
16700     End If

16710     If IsNull(.Recur_Type) = True Then
            ' ** Make sure there's one there.
16720       strThisJType = Nz(.journaltype, vbNullString)
16730       Select Case strThisJType
            Case "Misc."
16740         .Recur_Type = "Misc"
16750       Case "Paid"
16760         .Recur_Type = "Payee"
16770       Case "Received"
16780         .Recur_Type = "Payor"
16790       End Select
16800     End If

16810     strSaveMoveCtl = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
16820     cmdSave_Click  ' ** Procedure: Above.

16830   End With

EXITP:
16840   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Recur_Name_NotInList(NewData As String, Response As Integer)

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_NotInList"

16910   With Me
16920     Response = acDataErrContinue
16930     cmdSave_Click  ' ** Procedure: Above.
16940     Response = acDataErrAdded
          '################
          'CHECK WHETHER 'ADDED', WHEN IN FACT HERE WE DON'T, CAUSES A PROBLEM!
          '################
16950   End With

        ' **************************************************************************************************************************
        ' ** Response enumeration:
        ' **   Value   Constant             Description
        ' **   =====   ===================  ======================================================================================
        ' **     0     acDataErrContinue    Doesn't display the default message to the user. You can use this when you want to
        ' **                                display a custom message to the user. For example, the event procedure could display
        ' **                                a custom dialog box asking if the user wanted to save the new entry. If the response
        ' **                                is Yes, the event procedure would add the new entry to the list and set the Response
        ' **                                argument to acDataErrAdded. If the response is No, the event procedure would set the
        ' **                                Response argument to acDataErrContinue.
        ' **     1     acDataErrDisplay     Displays the default message to the user. You can use this when you don't want to
        ' **                                allow the user to add a new value to the combo box list. (Default)
        ' **     2     acDataErrAdded       Doesn't display a message to the user but enables you to add the entry to the combo
        ' **                                box list in the NotInList event procedure. After the entry is added, Microsoft
        ' **                                Access updates the list by requerying the combo box. Microsoft Access then rechecks
        ' **                                the string against the combo box list, and saves the value in the NewData argument
        ' **                                in the field the combo box is bound to. If the string is not in the list, then
        ' **                                Microsoft Access displays an error message.
        ' **************************************************************************************************************************

EXITP:
16960   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     If ERR.Number > 0 Then
590       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Else  ' ** Those huge negative error numbers are too big for the Integer in the Form_Error() parameters.
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End If
630     Resume EXITP

End Sub

Private Sub Recur_Name_KeyDown(KeyCode As Integer, Shift As Integer)

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17010   intRetVal = KeyCode
17020   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
17030   intShiftDown = (Shift And acShiftMask) > 0
17040   intAltDown = (Shift And acAltMask) > 0
17050   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17060   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17070     Select Case intRetVal
          Case vbKeyTab
17080       intRetVal = 0
17090       JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
17100     Case vbKeyReturn
17110       intRetVal = 0
17120       strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
17130     Case vbKeyPageUp, vbKeyPageDown
17140       strPageMoveCtl = "Recur_Name"
17150     End Select
17160   End If

        ' ** Shift keys.
17170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17180     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17190       intRetVal = 0
17200       JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
17210     End Select
17220   End If

EXITP:
17230   KeyCode = intRetVal
17240   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub Recur_Name_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_MouseDown"

17310   blnF4InvokedMouse = True  ' ** Maybe, maybe not.

EXITP:
17320   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Recur_Name_Exit(Cancel As Integer)

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_Exit"

17410   blnF4InvokedMouse = False

EXITP:
17420   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub assetdate_display_lbl_DblClick(Cancel As Integer)

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_lbl_DblClick"

17510   JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
17520   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub assetdate_display_GotFocus()

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_GotFocus"

17610   JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
17620   blnNoMove = False

EXITP:
17630   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub assetdate_display_BeforeUpdate(Cancel As Integer)

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_BeforeUpdate"

        Dim blnContinue As Boolean

17710   blnContinue = True

17720   With Me
17730     Select Case .journaltype
          Case "Deposit", "Purchase", "Liability (+)", "Liability (-)", "Sold", "Withdrawn", "Cost Adj.", "Received", "Dividend", "Interest"
17740       If .journaltype = "Received" And .assetno = 0& Then
17750         blnContinue = False
17760         If gblnClosing = False And gblnDeleting = False Then
17770           Beep
17780           MsgBox "A Trade Date on a Received is not allowed without a valid Asset.", _
                  vbInformation + vbOKOnly, "Invalid Entry"
17790           Cancel = -1
17800           blnNoMove = True
17810           .assetdate_display.Undo
17820         End If
17830       End If
17840       If blnContinue = True Then
17850         If IsNull(.assetdate_display) = True Then
17860           If gblnClosing = False And gblnDeleting = False Then
17870             Beep
17880             MsgBox "A Trade Date is required.", vbInformation + vbOKOnly, "Entry Required"
17890             Cancel = -1
17900             blnNoMove = True
17910             .assetdate_display.Undo
17920           End If
17930         Else
17940           If DateCheck_Trade(.assetdate_display) = False Then  ' ** Module Function: modUtilities.
17950             Cancel = -1
17960             blnNoMove = True
17970           End If
17980         End If
17990       End If
18000     Case Else
            ' ** No assetdate.
18010       If IsNull(.assetdate_display) = False Then
18020         If gblnClosing = False And gblnDeleting = False Then
18030           Beep
18040           MsgBox "A Trade Date on a " & .journaltype & " is not allowed.", vbInformation + vbOKOnly, "Invalid Entry"
18050           Cancel = -1
18060           blnNoMove = True
18070           .assetdate_display.Undo
18080         End If
18090       End If
18100     End Select
18110   End With

EXITP:
18120   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub assetdate_display_AfterUpdate()

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_AfterUpdate"

        Dim dblTmp01 As Double, strTmp02 As String, dblTmp03 As Double
        Dim strTmp04 As String, lngTmp05 As Long, strTmp06 As String

18210   With Me

18220     If .posted = True Then
18230       .posted = False
18240       .posted.Locked = False
18250     End If

18260     If IsNull(.assetdate_display) = False Then
18270       If IsDate(.assetdate_display) = True Then
18280         dblTmp01 = CDbl(.assetdate_display)  ' ** Date as Double number.
18290         strTmp02 = CStr(dblTmp01)            ' ** String of Double number.
18300         If InStr(strTmp02, ".") <> 0 Then    ' ** Entered date has time component (unlikely).
18310           strTmp02 = Left(strTmp02, (InStr(strTmp02, ".") - 1))  ' ** String without decimal component (time).
18320           .assetdate_display = CDate(CLng(strTmp02))
18330         End If
18340         If IsNull(.assetdate) = True Then  ' ** No pre-existing assetdate.
18350           .assetdate = .assetdate_display + time()  ' ** Add the time of day.
18360         Else
18370           dblTmp03 = CDbl(.assetdate)  ' ** Date as Double number.
18380           strTmp04 = CStr(dblTmp03)    ' ** String of Double number.
18390           If InStr(strTmp04, ".") > 0 Then  ' ** Pre-existing date has time component.
18400             lngTmp05 = CLng(Left(strTmp04, (InStr(strTmp04, ".") - 1)))  ' ** Day component only of assetdate.
18410             strTmp06 = Mid(strTmp04, (InStr(strTmp04, ".") + 1))  ' ** String of decimal component (time).
18420           Else
18430             lngTmp05 = CLng(strTmp04)  ' ** Day component only of assetdate.
18440             strTmp06 = "0"  ' ** String of decimal component (time).
18450           End If
18460           If CLng(strTmp02) = lngTmp05 Then  ' ** Day of entered date same as pre-existing date.
18470             If strTmp06 <> "0" Then  ' ** Pre-existing date has time component.
                    ' ** Leave pre-existing assetdate as-is.
18480             Else
18490               .assetdate = .assetdate_display + time()  ' ** Add the time of day.
18500             End If
18510           Else  ' ** Entered date different from pre-existing date.
18520             .assetdate = .assetdate_display + time()  ' ** Add the time of day.
18530           End If
18540         End If
18550       Else
18560         blnNoMove = True
18570       End If
18580     Else
18590       .assetdate = Null
18600     End If

18610     If blnNoMove = False Then
18620       strSaveMoveCtl = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
18630       cmdSave_Click  ' ** Procedure: Above.
18640     End If

18650   End With

EXITP:
18660   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub assetdate_display_KeyDown(KeyCode As Integer, Shift As Integer)

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18710   intRetVal = KeyCode
18720   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
18730   intShiftDown = (Shift And acShiftMask) > 0
18740   intAltDown = (Shift And acAltMask) > 0
18750   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18760   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18770     Select Case intRetVal
          Case vbKeyTab
18780       Select Case blnNoMove
            Case True
18790         blnNoMove = False
18800       Case False
18810         intRetVal = 0
18820         JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
18830       End Select
18840     Case vbKeyReturn
18850       Select Case blnNoMove
            Case True
18860         blnNoMove = False
18870       Case False
18880         intRetVal = 0
18890         strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
18900       End Select
18910     Case vbKeyPageUp, vbKeyPageDown
18920       Select Case blnNoMove
            Case True
18930         blnNoMove = False
18940       Case False
18950         strPageMoveCtl = "assetdate_display"
18960       End Select
18970     End Select
18980   End If

        ' ** Shift keys.
18990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19000     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19010       Select Case blnNoMove
            Case True
19020         blnNoMove = False
19030       Case False
19040         intRetVal = 0
19050         JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
19060       End Select
19070     End Select
19080   End If

EXITP:
19090   KeyCode = intRetVal
19100   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub assetdate_display_Exit(Cancel As Integer)

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_Exit"

        Dim datTmp01 As Date
        Dim blnContinue As Boolean

19210   blnContinue = False

19220   With Me
19230     If .assetdate_display.Locked = False Then
19240       If IsNull(.assetdate_display) = True Then
19250         Select Case .journaltype
              Case "Dividend", "Interest", "Purchase", "Deposit", "Sold", "Withdrawn", "Liability (+)", "Liability (-)", "Cost Adj."
19260           datTmp01 = Now()
19270           .assetdate = datTmp01
19280           .assetdate_display = CDate(Format(datTmp01, "mm/dd/yyyy"))  ' ** Strip the timestamp.
                'blnContinue = True  ' ** May cause problems!
19290         Case "Misc.", "Paid", "Received"
                ' ** Shouldn't be here.
19300         End Select
19310       End If
19320       If blnContinue = True Then
19330         strSaveMoveCtl = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
19340         cmdSave_Click  ' ** Procedure: Above.
19350       End If
19360     End If
19370   End With

EXITP:
19380   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

19410   JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
19420   blnNoMove = False

EXITP:
19430   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

19510   JC_Key_CalendarClick THIS_PROC, blnNoMove, blnNextRec, blnFromZero, _
          strSaveMoveCtl, clsMonthClass, Me  ' ** Module Procedure: modJrnlCol_Keys.

EXITP:
19520   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19610   intRetVal = KeyCode
19620   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
19630   intShiftDown = (Shift And acShiftMask) > 0
19640   intAltDown = (Shift And acAltMask) > 0
19650   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19660   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19670     Select Case intRetVal
          Case vbKeyTab
19680       Select Case blnNoMove
            Case True
19690         blnNoMove = False
19700       Case False
19710         intRetVal = 0
19720         JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
19730       End Select
19740     Case vbKeyPageUp, vbKeyPageDown
19750       Select Case blnNoMove
            Case True
19760         blnNoMove = False
19770       Case False
19780         strPageMoveCtl = "assetdate_display"
19790       End Select
19800     End Select
19810   End If

        ' ** Shift keys.
19820   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19830     Select Case intRetVal
          Case vbKeyTab
19840       With Me
19850         Select Case blnNoMove
              Case True
19860           blnNoMove = False
19870         Case False
19880           intRetVal = 0
19890           .assetdate_display.SetFocus
19900         End Select
19910       End With
19920     End Select
19930   End If

EXITP:
19940   KeyCode = intRetVal
19950   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub shareface_lbl_DblClick(Cancel As Integer)

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_lbl_DblClick"

20010   JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
20020   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub shareface_Enter()

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_Enter"

20110   With Me
20120     If IsNull(.shareface) = False Then
20130       dblShareFace_OldValue = .shareface
20140     Else
20150       dblShareFace_OldValue = 0#
20160     End If
20170     .Parent.ToTaxLot = 0&
20180     .Parent.TaxLotFrom = vbNullString
20190   End With

EXITP:
20200   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub shareface_GotFocus()

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_GotFocus"

20310   JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
20320   blnNoMove = False

EXITP:
20330   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub shareface_BeforeUpdate(Cancel As Integer)

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_BeforeUpdate"

        Dim dblShareFace_NewValue As Double
        Dim msgResponse As VbMsgBoxResult

20410   With Me
20420     If IsNull(.shareface) = True Then
20430       dblShareFace_NewValue = 0#
20440     Else
20450       dblShareFace_NewValue = .shareface
20460     End If
20470     If dblShareFace_NewValue <> dblShareFace_OldValue And dblShareFace_OldValue <> 0# Then
20480       If .posted = True Then
20490         msgResponse = MsgBox("You've changed the Share/Face value." & vbCrLf & _
                "Do you really want to continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Share/Face Changed")
20500         If msgResponse <> vbYes Then
20510           Cancel = -1
20520           blnNoMove = True
20530           .shareface.Undo
20540         End If
20550       End If
20560     End If
20570   End With

EXITP:
20580   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub shareface_change()

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_Change"

20610   JCol_Shareface_Change THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.

EXITP:
20620   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub shareface_AfterUpdate()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_AfterUpdate"

        Dim strThisJType As String

20710   With Me

20720     If .posted = True Then
20730       .posted = False
20740       .posted.Locked = False
20750     End If

20760     If IsNull(.shareface) = True Then
20770       .shareface = 0#
20780     End If

20790     strThisJType = Nz(.journaltype, vbNullString)
20800     Select Case strThisJType
          Case "Withdrawn", "Sold"
20810       If IsNull(.Cost) = False Then
20820         If .Cost <> 0@ Then
20830           .Cost = 0@
20840         End If
20850       End If
20860     End Select

20870     Select Case strThisJType
          Case "Purchase"
20880       If .shareface <> 0# And Nz(.ICash, 0) < 0@ And Nz(.Cost, 0) = Abs(Nz(.ICash, 0)) Then
              ' ** Reinvestment from Dividend or Interest.
20890         strSaveMoveCtl = "Loc_Name_display"
20900       Else
              ' ** Proceed normally.
20910         strSaveMoveCtl = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
20920       End If
20930     Case Else
            ' ** Proceed normally.
20940       strSaveMoveCtl = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
20950     End Select
20960     blnNoMove = True
20970     cmdSave_Click  ' ** Procedure: Above.

20980   End With

EXITP:
20990   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub shareface_KeyDown(KeyCode As Integer, Shift As Integer)

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21010   intRetVal = KeyCode
21020   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
21030   intShiftDown = (Shift And acShiftMask) > 0
21040   intAltDown = (Shift And acAltMask) > 0
21050   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21060   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21070     Select Case intRetVal
          Case vbKeyTab
21080       Select Case blnNoMove
            Case True
21090         blnNoMove = False
21100       Case False
21110         intRetVal = 0
21120         JC_Key_Sub Me, THIS_PROC, KeyCode, 3, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
21130       End Select
21140     Case vbKeyReturn
21150       Select Case blnNoMove
            Case True
21160         blnNoMove = False
21170       Case False
21180         With Me
21190           intRetVal = 0
21200           If .shareface = 0# And IIf(.shareface.text = vbNullString, 0#, Val(.shareface.text)) Then
21210             If .Dirty = True Then
21220               DoCmd.RunCommand acCmdSaveRecord
21230             End If
21240           End If
21250           strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 3, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
21260         End With
21270       End Select
21280     Case vbKeyPageUp, vbKeyPageDown
21290       Select Case blnNoMove
            Case True
21300         blnNoMove = False
21310       Case False
21320         strPageMoveCtl = "shareface"
21330       End Select
21340     End Select
21350   End If

        ' ** Shift keys.
21360   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21370     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21380       Select Case blnNoMove
            Case True
21390         blnNoMove = False
21400       Case False
21410         intRetVal = 0
21420         JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
21430       End Select
21440     End Select
21450   End If

EXITP:
21460   KeyCode = intRetVal
21470   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub shareface_Exit(Cancel As Integer)

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_Exit"

21510   With Me

21520     JCol_ShareFace_Exit Cancel, blnToTaxLot, blnGoingToReport, blnGoneToReport, blnNextRec, blnFromZero, _
            datPostingDate, strSaveMoveCtl, blnNoMove, lngNewJrnlColID, lngGTR_ID, blnGTR_Emblem, _
            blnGTR_NoAdd, THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

21530   End With

EXITP:
21540   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub icash_lbl_DblClick(Cancel As Integer)

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_lbl_DblClick"

21610   JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
21620   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub icash_Enter()

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_Enter"

        Dim lngRetVal As Long

21710   With Me
21720     If IsNull(.ICash) = False Then
21730       dblICash_OldValue = .ICash
21740     Else
21750       dblICash_OldValue = 0@
21760     End If
          ' ** Doesn't seem to want to scroll all
          ' ** the way right on Rich's machine!
21770     lngRetVal = fSetScrollBarPosHZ(Me, 999&)  ' ** Module Function: modScrollBarFuncs.
21780     gstrSaleICash = "0"
21790   End With

EXITP:
21800   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub icash_GotFocus()

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_GotFocus"

21910   JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
21920   blnNoMove = False

EXITP:
21930   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub icash_BeforeUpdate(Cancel As Integer)

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_BeforeUpdate"

        Dim dblICash_NewValue As Double
        Dim msgResponse As VbMsgBoxResult

22010   With Me
22020     If IsNull(.ICash) = True Then
22030       dblICash_NewValue = 0#
22040     Else
22050       dblICash_NewValue = .ICash
22060     End If
22070     If dblICash_NewValue <> dblICash_OldValue And dblICash_OldValue <> 0# Then
22080       If .posted = True Then
22090         msgResponse = MsgBox("You've changed the Income Cash value." & vbCrLf & _
                "Do you really want to continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Income Cash Changed")
22100         If msgResponse <> vbYes Then
22110           Cancel = -1
22120           blnNoMove = True
22130           .ICash.Undo
22140         End If
22150       End If
22160     End If
22170   End With

EXITP:
22180   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub icash_Change()

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_Change"

22210   JCol_ICash_Change THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.

EXITP:
22220   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub icash_AfterUpdate()

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_AfterUpdate"

22310   With Me

22320     JCol_ICash_AfterUpdate strSaveMoveCtl, blnNextRec, blnFromZero, blnNoMove, _
            THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

22330   End With

EXITP:
22340   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub icash_KeyDown(KeyCode As Integer, Shift As Integer)

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22410   intRetVal = KeyCode
22420   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
22430   intShiftDown = (Shift And acShiftMask) > 0
22440   intAltDown = (Shift And acAltMask) > 0
22450   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22460   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22470     Select Case intRetVal
          Case vbKeyTab
22480       Select Case blnNoMove
            Case True
22490         blnNoMove = False
22500       Case False
22510         intRetVal = 0
22520         With Me
22530           gstrSaleICash = CStr(CDbl(Val(Nz(.ICash.text, 0))))
22540           gstrSaleICash = Rem_Dollar(gstrSaleICash)  ' ** Module Function: modStringFuncs.
22550         End With
22560         JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
22570       End Select
22580     Case vbKeyReturn
22590       Select Case blnNoMove
            Case True
22600         blnNoMove = False
22610       Case False
22620         intRetVal = 0
22630         With Me
22640           gstrSaleICash = CStr(CDbl(Val(Nz(.ICash.text, 0))))
22650           gstrSaleICash = Rem_Dollar(gstrSaleICash)  ' ** Module Function: modStringFuncs.
22660         End With
22670         strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
22680       End Select
22690     Case vbKeyPageUp, vbKeyPageDown
22700       Select Case blnNoMove
            Case True
22710         blnNoMove = False
22720       Case False
22730         strPageMoveCtl = "icash"
22740       End Select
22750     End Select
22760   End If

        ' ** Shift keys.
22770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22790       Select Case blnNoMove
            Case True
22800         blnNoMove = False
22810       Case False
22820         intRetVal = 0
22830         JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
22840       End Select
22850     End Select
22860   End If

EXITP:
22870   KeyCode = intRetVal
22880   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub icash_Exit(Cancel As Integer)

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_Exit"

        Dim strThisJType As String
        Dim blnContinue As Boolean

22910   blnContinue = False

22920   With Me
22930     If .ICash.Locked = False Then

22940       If IsNull(.ICash) = True Then
22950         .ICash = 0@
22960         blnContinue = True
22970       Else
22980         If .ICash.text = vbNullString Then
22990           .ICash = 0@
23000           blnContinue = True
23010         End If
23020       End If

23030       strThisJType = Nz(.journaltype, vbNullString)
23040       Select Case strThisJType
            Case "Dividend", "Interest"
23050         blnContinue = False
23060         If .ICash = 0& Then
23070           If gblnClosing = False And gblnDeleting = False Then
23080             Beep
23090             MsgBox "An Income Cash value is required for a " & strThisJType & " transaction.", _
                    vbInformation + vbOKOnly, "Entry Required"
23100             Cancel = -1
23110           End If
23120         End If
23130       Case Else
              ' ** Nothing at the moment.
23140       End Select

23150       If blnContinue = True Then
23160         strSaveMoveCtl = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
23170         cmdSave_Click  ' ** Procedure: Above.
23180       End If

23190     End If
23200   End With

EXITP:
23210   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub pcash_lbl_DblClick(Cancel As Integer)

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_lbl_DblClick"

23310   JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
23320   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub pcash_Enter()

23400 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_Enter"

23410   With Me
23420     If IsNull(.PCash) = False Then
23430       dblPCash_OldValue = .PCash
23440     Else
23450       dblPCash_OldValue = 0@
23460     End If
23470     gstrSalePCash = "0"
23480   End With

EXITP:
23490   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub pcash_GotFocus()

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_GotFocus"

23510   JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
23520   blnNoMove = False

EXITP:
23530   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub pcash_BeforeUpdate(Cancel As Integer)

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_BeforeUpdate"

        Dim dblPcash_NewValue As Double
        Dim msgResponse As VbMsgBoxResult

23610   With Me
23620     If IsNull(.PCash) = True Then
23630       dblPcash_NewValue = 0#
23640     Else
23650       dblPcash_NewValue = .PCash
23660     End If
23670     If dblPcash_NewValue <> dblPCash_OldValue And dblPCash_OldValue <> 0# Then
23680       If .posted = True Then
23690         msgResponse = MsgBox("You've changed the Principal Cash value." & vbCrLf & _
                "Do you really want to continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Principal Cash Changed")
23700         If msgResponse <> vbYes Then
23710           Cancel = -1
23720           blnNoMove = True
23730           .PCash.Undo
23740         End If
23750       End If
23760     End If
23770   End With

EXITP:
23780   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub pcash_Change()

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_Change"

23810   JCol_PCash_Change THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.

EXITP:
23820   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub pcash_AfterUpdate()
' ** Additional checks in JC_Key_Sub().

23900 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_AfterUpdate"

23910   With Me

23920     JCol_PCash_AfterUpdate strSaveMoveCtl, blnNextRec, blnFromZero, blnNoMove, _
            THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

23930   End With

EXITP:
23940   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub pcash_KeyDown(KeyCode As Integer, Shift As Integer)
' ** Additional checks in JC_Key_Sub().

24000 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24010   intRetVal = KeyCode
24020   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
24030   intShiftDown = (Shift And acShiftMask) > 0
24040   intAltDown = (Shift And acAltMask) > 0
24050   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24060   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24070     Select Case intRetVal
          Case vbKeyTab
24080       Select Case blnNoMove
            Case True
24090         blnNoMove = False
24100       Case False
24110         intRetVal = 0
24120         With Me
24130           If Val(gstrSaleICash) = 0 Then
24140             gstrSaleICash = CStr(CDbl(Val(Nz(.ICash, 0))))
24150             gstrSaleICash = Rem_Dollar(gstrSaleICash)  ' ** Module Function: modStringFuncs.
24160           End If
24170           If Val(gstrSalePCash) = 0 Then
24180             gstrSalePCash = CStr(CDbl(Val(Nz(.PCash.text, 0))))
24190             gstrSalePCash = Rem_Dollar(gstrSalePCash)  ' ** Module Function: modStringFuncs.
24200           End If
24210         End With
24220         JC_Key_Sub Me, THIS_PROC, KeyCode, 3, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
24230       End Select
24240     Case vbKeyReturn
24250       Select Case blnNoMove
            Case True
24260         blnNoMove = False
24270       Case False
24280         intRetVal = 0
24290         With Me
24300           If Val(gstrSaleICash) = 0 Then
24310             gstrSaleICash = CStr(CDbl(Val(Nz(.ICash, 0))))
24320             gstrSaleICash = Rem_Dollar(gstrSaleICash)  ' ** Module Function: modStringFuncs.
24330           End If
24340           If Val(gstrSalePCash) = 0 Then
24350             gstrSalePCash = CStr(CDbl(Val(Nz(.PCash.text, 0))))
24360             gstrSalePCash = Rem_Dollar(gstrSalePCash)  ' ** Module Function: modStringFuncs.
24370           End If
24380         End With
24390         strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 3, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
24400       End Select
24410     Case vbKeyPageUp, vbKeyPageDown
24420       Select Case blnNoMove
            Case True
24430         blnNoMove = False
24440       Case False
24450         strPageMoveCtl = "pcash"
24460       End Select
24470     End Select
24480   End If

        ' ** Shift keys.
24490   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24500     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24510       Select Case blnNoMove
            Case True
24520         blnNoMove = False
24530       Case False
24540         intRetVal = 0
24550         JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
24560       End Select
24570     End Select
24580   End If

EXITP:
24590   KeyCode = intRetVal
24600   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub pcash_Exit(Cancel As Integer)
' ** Additional checks in JC_Key_Sub().

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_Exit"

24710   With Me

24720     JCol_PCash_Exit Cancel, blnToTaxLot, blnWarnZeroCost, blnGoingToReport, blnGoneToReport, _
            blnNextRec, blnFromZero, datPostingDate, strSaveMoveCtl, blnNoMove, lngNewJrnlColID, _
            lngGTR_ID, blnGTR_Emblem, blnGTR_NoAdd, THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

24730   End With

EXITP:
24740   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub cost_lbl_DblClick(Cancel As Integer)

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_lbl_DblClick"

24810   JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
24820   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub cost_GotFocus()

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_GotFocus"

24910   JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
24920   blnNoMove = False

EXITP:
24930   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub cost_Change()

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_Change"

25010   JCol_Cost_Change THAT_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.

EXITP:
25020   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub cost_AfterUpdate()

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_AfterUpdate"

25110   With Me

25120     JCol_Cost_AfterUpdate strSaveMoveCtl, blnNextRec, blnFromZero, blnNoMove, _
            THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

25130   End With

EXITP:
25140   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub cost_KeyDown(KeyCode As Integer, Shift As Integer)

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25210   intRetVal = KeyCode
25220   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
25230   intShiftDown = (Shift And acShiftMask) > 0
25240   intAltDown = (Shift And acAltMask) > 0
25250   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25260   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25270     Select Case intRetVal
          Case vbKeyTab
25280       Select Case blnNoMove
            Case True
25290         blnNoMove = False
25300       Case False
25310         intRetVal = 0
25320         JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
25330       End Select
25340     Case vbKeyReturn
25350       Select Case blnNoMove
            Case True
25360         blnNoMove = False
25370       Case False
25380         intRetVal = 0
25390         strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
25400       End Select
25410     Case vbKeyPageUp, vbKeyPageDown
25420       Select Case blnNoMove
            Case True
25430         blnNoMove = False
25440       Case False
25450         strPageMoveCtl = "cost"
25460       End Select
25470     End Select
25480   End If

        ' ** Shift keys.
25490   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25500     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25510       Select Case blnNoMove
            Case True
25520         blnNoMove = False
25530       Case False
25540         intRetVal = 0
25550         JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
25560       End Select
25570     End Select
25580   End If

EXITP:
25590   KeyCode = intRetVal
25600   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub cost_Exit(Cancel As Integer)

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_Exit"

        Dim blnContinue As Boolean

25710   With Me

25720     blnContinue = False

25730     JCol_Cost_Exit Cancel, blnContinue, blnWarnZeroCost, strSaveMoveCtl, blnNextRec, _
            blnFromZero, blnNoMove, THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

25740   End With

EXITP:
25750   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Loc_Name_display_GotFocus()
' ** This control sits on top of the Location_ID combo box.
' ** It's used to display the Loc_Name.
' ** CONTROL IS ENABLED, BUT LOCKED.

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Name_display_GotFocus"

25810   JC_Msc_LocNameDisp blnNoMove, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.

EXITP:
25820   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Loc_Name_display_KeyDown(KeyCode As Integer, Shift As Integer)
' ** Ordinarily, they shouldn't have time to press a key while on this field!

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Name_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25910   intRetVal = KeyCode
25920   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
25930   intShiftDown = (Shift And acShiftMask) > 0
25940   intAltDown = (Shift And acAltMask) > 0
25950   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25960   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25970     Select Case intRetVal
          Case vbKeyTab
25980       With Me
25990         intRetVal = 0
26000         If .Location_ID.Locked = False Then
26010           .Location_ID.SetFocus
26020         Else
26030           JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
26040         End If
26050       End With
26060     Case vbKeyReturn
26070       With Me
26080         intRetVal = 0
26090         If .Location_ID.Locked = False Then
26100           .Location_ID.SetFocus
26110         Else
26120           strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
26130         End If
26140       End With
26150     Case vbKeyPageUp, vbKeyPageDown
26160       strPageMoveCtl = "Location_ID"
26170     End Select
26180   End If

        ' ** Shift keys.
26190   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26200     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26210       intRetVal = 0
26220       JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
26230     End Select
26240   End If

EXITP:
26250   KeyCode = intRetVal
26260   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub Location_ID_lbl_DblClick(Cancel As Integer)

26300 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_lbl_DblClick"

26310   JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
26320   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Location_ID_Enter()

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_Enter"

26410   With Me
26420     If gblnGoToReport = False Then
26430       If IsNull(.Location_ID) = True Then
26440         blnF4Invoked = True: strF4LastControl = "Location_ID"
26450         ExpandCombo .Location_ID  ' ** Module Procedure: modUtilities.
26460       Else
26470         If .Location_ID = 1& Then
26480           blnF4Invoked = True: strF4LastControl = "Location_ID"
26490           ExpandCombo .Location_ID  ' ** Module Procedure: modUtilities.
26500         End If
26510       End If
26520     End If
26530   End With

EXITP:
26540   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Location_ID_GotFocus()

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_GotFocus"

26610   JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
26620   blnNoMove = False

EXITP:
26630   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Location_ID_AfterUpdate()

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnPosted As Boolean

26710   With Me

          ' ** Don't automatically uncommit for a change in Location_ID.
26720     blnPosted = .posted

26730     If IsNull(.Location_ID) = False Then
26740       .Loc_Name = .Location_ID.Column(1)
26750       If .Location_ID = 1& Then
26760         .Loc_Name_display = vbNullString
26770       Else
26780         .Loc_Name_display = .Loc_Name
26790       End If
26800     Else
26810       .Location_ID = 1&
26820       .Loc_Name = "{no entry}"
26830       .Loc_Name_display = vbNullString
26840     End If

26850     If blnPosted = True Then
26860       Set dbs = CurrentDb
26870       With dbs
              ' ** Update Journal, for Location_ID, by specified [JrnlID], [locid].
26880         Set qdf = .QueryDefs("qryJournal_Columns_11c")
26890         With qdf.Parameters
26900           ![jrnlid] = Me.Journal_ID
26910           ![locid] = Me.Location_ID
26920         End With
26930         qdf.Execute
26940         .Close
26950       End With
26960     End If

26970     strSaveMoveCtl = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
26980     cmdSave_Click  ' ** Procedure: Above.

26990   End With

EXITP:
27000   Set qdf = Nothing
27010   Set dbs = Nothing
27020   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Location_ID_NotInList(NewData As String, Response As Integer)

27100 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_NotInList"

27110   With Me
27120     If gblnClosing = False And gblnDeleting = False Then
27130       .Location_ID = 1&
27140       blnNoMove = True
27150       MsgBox "You must select a valid Location from the list," & vbCrLf & _
              "or choose {no entry}.", vbInformation + vbOKOnly, "Invalid Entry"
27160       Response = acDataErrContinue
27170     End If
27180   End With

EXITP:
27190   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     If ERR.Number > 0 Then
590       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Else  ' ** Those huge negative error numbers are too big for the Integer in the Form_Error() parameters.
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End If
630     Resume EXITP

End Sub

Private Sub Location_ID_KeyDown(KeyCode As Integer, Shift As Integer)

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27210   intRetVal = KeyCode
27220   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
27230   intShiftDown = (Shift And acShiftMask) > 0
27240   intAltDown = (Shift And acAltMask) > 0
27250   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27260   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27270     Select Case intRetVal
          Case vbKeyTab
27280       With Me
27290         intRetVal = 0
27300         If IsNull(.Location_ID) = True Then
27310           .Location_ID = 1&
27320         Else
27330           If IsNull(.Location_ID.text) = True Then
27340             .Location_ID.text = "{no entry}"
27350           Else
27360             If .Location_ID.text = vbNullString Then
27370               .Location_ID.text = "{no entry}"
27380             End If
27390           End If
27400         End If
27410         JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
27420       End With
27430     Case vbKeyReturn
27440       With Me
27450         intRetVal = 0
27460         If IsNull(.Location_ID) = True Then
27470           .Location_ID = 1&
27480         Else
27490           If IsNull(.Location_ID.text) = True Then
27500             .Location_ID.text = "{no entry}"
27510           Else
27520             If .Location_ID.text = vbNullString Then
27530               .Location_ID.text = "{no entry}"
27540             End If
27550           End If
27560         End If
27570         strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
27580       End With
27590     Case vbKeyPageUp, vbKeyPageDown
27600       strPageMoveCtl = "Location_ID"
27610     End Select
27620   End If

        ' ** Shift keys.
27630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27650       intRetVal = 0
27660       JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
27670     End Select
27680   End If

EXITP:
27690   KeyCode = intRetVal
27700   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub Location_ID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_MouseDown"

27810   blnF4InvokedMouse = True  ' ** Maybe, maybe not.

EXITP:
27820   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Location_ID_Exit(Cancel As Integer)

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_Exit"

27910   blnF4InvokedMouse = False

EXITP:
27920   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub PrintCheck_lbl_DblClick(Cancel As Integer)

28000 On Error GoTo ERRH

        Const THIS_PROC As String = "PrintCheck_lbl_DblClick"

28010   JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
28020   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub PrintCheck_GotFocus()

28100 On Error GoTo ERRH

        Const THIS_PROC As String = "PrintCheck_GotFocus"

28110   JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
28120   blnNoMove = False

28130   With Me
          ' ** Because this is the same color as the Detail BackColor,
          ' ** it should only show when the highlight bar is beneath it.
28140     .PrintCheck_box.Visible = True
28150   End With

EXITP:
28160   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub PrintCheck_BeforeUpdate(Cancel As Integer)

28200 On Error GoTo ERRH

        Const THIS_PROC As String = "PrintCheck_BeforeUpdate"

28210   With Me
28220     Select Case .PrintCheck
          Case True
28230       If .journaltype <> "Paid" Then
28240         If gblnClosing = False And gblnDeleting = False Then
28250           Beep
28260           MsgBox "A check is only allowed for a 'Paid' Journal Type.", vbInformation + vbOKOnly, "Invalid Entry"
28270           Cancel = -1
28280         End If
28290       End If
28300     Case False
            ' ** Nothing.
28310     End Select
28320   End With

EXITP:
28330   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub PrintCheck_AfterUpdate()

28400 On Error GoTo ERRH

        Const THIS_PROC As String = "PrintCheck_AfterUpdate"

28410   With Me

28420     If .posted = True Then
28430       .posted = False
28440       .posted.Locked = False
28450     End If

28460     Select Case .PrintCheck
          Case True
28470       JC_Msc_Memo_Set .PrintCheck, .Parent  ' ** Module Procedure: modJrnlCol_Misc.
28480       If IsNull(.JrnlMemo_Memo) = False Then
28490         .Parent.JrnlMemo_Memo = .JrnlMemo_Memo
28500       End If
28510       DoEvents
28520       strSaveMoveCtl = vbNullString
28530       cmdSave_Click  ' ** Procedure: Above.
28540       DoCmd.SelectObject acForm, .Parent.Name, False
28550       .Parent.JrnlMemo_Memo.SetFocus
28560     Case False
28570       .JrnlMemo_Memo = Null
28580       .Parent.JrnlMemo_Memo = Null
28590       JC_Msc_Memo_Set .PrintCheck, .Parent  ' ** Module Procedure: modJrnlCol_Misc.
28600       strSaveMoveCtl = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
28610       cmdSave_Click  ' ** Procedure: Above.
28620     End Select

28630   End With

EXITP:
28640   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub PrintCheck_KeyDown(KeyCode As Integer, Shift As Integer)

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "PrintCheck_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28710   intRetVal = KeyCode
28720   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
28730   intShiftDown = (Shift And acShiftMask) > 0
28740   intAltDown = (Shift And acAltMask) > 0
28750   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28760   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28770     Select Case intRetVal
          Case vbKeyTab
28780       With Me
28790         intRetVal = 0
28800         Select Case .PrintCheck
              Case True
28810           DoCmd.SelectObject acForm, .Parent.Name, False
28820           .Parent.JrnlMemo_Memo.SetFocus
28830         Case False
28840           JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
28850         End Select
28860       End With
28870     Case vbKeyReturn
28880       With Me
28890         intRetVal = 0
28900         Select Case .PrintCheck
              Case True
28910           DoCmd.SelectObject acForm, .Parent.Name, False
28920           .Parent.JrnlMemo_Memo.SetFocus
28930         Case False
28940           strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
28950         End Select
28960       End With
28970     Case vbKeyPageUp, vbKeyPageDown
28980       strPageMoveCtl = "PrintCheck"
28990     End Select
29000   End If

        ' ** Shift keys.
29010   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29020     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29030       intRetVal = 0
29040       JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
29050     End Select
29060   End If

EXITP:
29070   KeyCode = intRetVal
29080   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub PrintCheck_LostFocus()

29100 On Error GoTo ERRH

        Const THIS_PROC As String = "PrintCheck_LostFocus"

29110   With Me
29120     .PrintCheck_box.Visible = False
29130   End With

EXITP:
29140   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub description_lbl_DblClick(Cancel As Integer)

29200 On Error GoTo ERRH

        Const THIS_PROC As String = "description_lbl_DblClick"

29210   JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
29220   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub description_GotFocus()

29300 On Error GoTo ERRH

        Const THIS_PROC As String = "description_GotFocus"

29310   JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
29320   blnNoMove = False

EXITP:
29330   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub description_AfterUpdate()
' ** Let them edit a committed entry without having to re-commit.

29400 On Error GoTo ERRH

        Const THIS_PROC As String = "description_AfterUpdate"

        Dim blnIsPosted As Boolean

29410   With Me

29420     blnIsPosted = False

29430     JCol_Desc_AfterUpdate blnIsPosted, strSaveMoveCtl, blnNextRec, blnFromZero, blnNoMove, _
            THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

29440   End With

EXITP:
29450   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub description_KeyDown(KeyCode As Integer, Shift As Integer)

29500 On Error GoTo ERRH

        Const THIS_PROC As String = "description_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29510   intRetVal = KeyCode
29520   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
29530   intShiftDown = (Shift And acShiftMask) > 0
29540   intAltDown = (Shift And acAltMask) > 0
29550   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29560   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29570     Select Case intRetVal
          Case vbKeyTab
29580       Select Case blnNoMove
            Case True
29590         blnNoMove = False
29600       Case False
29610         intRetVal = 0
29620         JC_Key_Sub Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
29630       End Select
29640     Case vbKeyReturn
29650       Select Case blnNoMove
            Case True
29660         blnNoMove = False
29670       Case False
29680         intRetVal = 0
29690         strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
29700       End Select
29710     Case vbKeyPageUp, vbKeyPageDown
29720       Select Case blnNoMove
            Case True
29730         blnNoMove = False
29740       Case False
29750         With Me
29760           If .posted = False And .journaltype = "Cost Adj." And .Cost <> 0@ And .assetno > 0& Then
29770             intRetVal = 0
29780           End If
29790         End With
29800         strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
29810       End Select
29820     End Select
29830   End If

        ' ** Shift keys.
29840   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29850     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29860       Select Case blnNoMove
            Case True
29870         blnNoMove = False
29880       Case False
              ' ** If they're moving backwards, don't bother checking for Commit.
29890         intRetVal = 0
29900         JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
29910       End Select
29920     End Select
29930   End If

EXITP:
29940   KeyCode = intRetVal
29950   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub revcode_DESC_display_GotFocus()
' ** This control sits on top of the revcode_ID combo box.
' ** It's used to display the revcode_DESC when the
' ** RowSource is showing the other set.
' ** CONTROL IS ENABLED, BUT LOCKED.

30000 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_display_GotFocus"

30010   JC_Msc_RevCodeDescDisp blnNoMove, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.

EXITP:
30020   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub revcode_DESC_display_KeyDown(KeyCode As Integer, Shift As Integer)
' ** Ordinarily, they shouldn't have time to press a key while on this field!

30100 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30110   intRetVal = KeyCode
30120   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
30130   intShiftDown = (Shift And acShiftMask) > 0
30140   intAltDown = (Shift And acAltMask) > 0
30150   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30160   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30170     Select Case intRetVal
          Case vbKeyTab
30180       With Me
30190         intRetVal = 0
30200         If .revcode_ID.Locked = False Then
30210           .revcode_ID.SetFocus
30220         Else
30230           JC_Key_Sub Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
30240         End If
30250       End With
30260     Case vbKeyReturn
30270       With Me
30280         intRetVal = 0
30290         If .revcode_ID.Locked = False Then
30300           .revcode_ID.SetFocus
30310         Else
30320           strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
30330         End If
30340       End With
30350     Case vbKeyPageUp, vbKeyPageDown
30360       Select Case blnNoMove
            Case True
30370         blnNoMove = False
30380       Case False
30390         With Me
30400           If .posted = False And .journaltype = "Cost Adj." And .Cost <> 0@ And .assetno > 0& Then
30410             intRetVal = 0
30420           End If
30430         End With
30440         strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
30450       End Select
30460     End Select
30470   End If

        ' ** Shift keys.
30480   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
30490     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30500       intRetVal = 0
30510       JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
30520     End Select
30530   End If

EXITP:
30540   KeyCode = intRetVal
30550   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub revcode_ID_lbl_DblClick(Cancel As Integer)

30600 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ID_lbl_DblClick"

30610   Select Case gblnRevenueExpenseTracking
        Case True
30620     JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.
30630   Case False
          ' ** Disabled.
30640   End Select

EXITP:
30650   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub revcode_ID_Enter()

30700 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ID_Enter"

30710   With Me
30720     If IsNull(.revcode_ID) = True Then
30730       blnF4Invoked = True: strF4LastControl = "revcode_ID"
30740       ExpandCombo .revcode_ID  ' ** Module Procedure: modUtilities.
30750     Else
30760       If .revcode_ID.text = vbNullString Then
30770         blnF4Invoked = True: strF4LastControl = "revcode_ID"
30780         ExpandCombo .revcode_ID  ' ** Module Procedure: modUtilities.
30790       Else
30800         If .revcode_ID = REVID_INC Or .revcode_ID = REVID_EXP Then
30810           blnF4Invoked = True: strF4LastControl = "revcode_ID"
30820           ExpandCombo .revcode_ID  ' ** Module Procedure: modUtilities.
30830         End If
30840       End If
30850     End If
30860   End With

EXITP:
30870   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub revcode_ID_GotFocus()

30900 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ID_GotFocus"

30910   JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
30920   blnNoMove = False

EXITP:
30930   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub revcode_ID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

31000 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ID_MouseDown"

31010   blnF4InvokedMouse = True  ' ** Maybe, maybe not.

EXITP:
31020   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub revcode_ID_AfterUpdate()

31100 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ID_AfterUpdate"

        Dim blnPosted As Boolean

31110   With Me

          ' ** Don't automatically uncommit for a change in revcode_ID or taxcode.
31120     blnPosted = .posted

31130     JCol_RevID_AfterUpdate blnPosted, strSaveMoveCtl, blnNextRec, blnFromZero, _
            blnNoMove, THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

31140   End With

EXITP:
31150   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub revcode_ID_NotInList(NewData As String, Response As Integer)

31200 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ID_NotInList"

31210   With Me
31220     If gblnClosing = False And gblnDeleting = False Then
31230       .revcode_ID = REVID_INC
31240       blnNoMove = True
31250       MsgBox "You must select a valid Code from the list," & vbCrLf & _
              "or choose {no entry}.", vbInformation + vbOKOnly, "Invalid Entry"
31260       Response = acDataErrContinue
31270     End If
31280   End With

EXITP:
31290   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     If ERR.Number > 0 Then
590       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Else  ' ** Those huge negative error numbers are too big for the Integer in the Form_Error() parameters.
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End If
630     Resume EXITP

End Sub

Private Sub revcode_ID_KeyDown(KeyCode As Integer, Shift As Integer)

31300 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ID_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

31310   intRetVal = KeyCode
31320   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
31330   intShiftDown = (Shift And acShiftMask) > 0
31340   intAltDown = (Shift And acAltMask) > 0
31350   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
31360   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31370     Select Case intRetVal
          Case vbKeyTab
31380       Select Case blnNoMove
            Case True
31390         blnNoMove = False
31400       Case False
31410         With Me
31420           intRetVal = 0
31430           If IsNull(.revcode_ID) = True Then
31440             Select Case .revcode_TYPE
                  Case REVTYP_INC
31450               .revcode_ID = REVID_INC
31460             Case REVTYP_EXP
31470               .revcode_ID = REVID_EXP
31480             End Select
31490           Else
31500             If IsNull(.revcode_ID.text) = True Then
31510               .revcode_ID.text = "{no entry}"
31520             Else
31530               If .revcode_ID.text = vbNullString Then
31540                 .revcode_ID.text = "{no entry}"
31550               End If
31560             End If
31570           End If
31580           JC_Key_Sub Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
31590         End With
31600       End Select
31610     Case vbKeyReturn
31620       Select Case blnNoMove
            Case True
31630         blnNoMove = False
31640       Case False
31650         With Me
31660           intRetVal = 0
31670           If IsNull(.revcode_ID) = True Then
31680             Select Case .revcode_TYPE
                  Case REVTYP_INC
31690               .revcode_ID = REVID_INC
31700             Case REVTYP_EXP
31710               .revcode_ID = REVID_EXP
31720             End Select
31730           Else
31740             If IsNull(.revcode_ID.text) = True Then
31750               .revcode_ID.text = "{no entry}"
31760             Else
31770               If .revcode_ID.text = vbNullString Then
31780                 .revcode_ID.text = "{no entry}"
31790               End If
31800             End If
31810           End If
31820           strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
31830         End With
31840       End Select
31850     Case vbKeyPageUp, vbKeyPageDown
31860       Select Case blnNoMove
            Case True
31870         blnNoMove = False
31880       Case False
31890         With Me
31900           If .posted = False And .journaltype = "Cost Adj." And .Cost <> 0@ And .assetno > 0& Then
31910             intRetVal = 0
31920           End If
31930         End With
31940         strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
31950       End Select
31960     End Select
31970   End If

        ' ** Shift keys.
31980   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31990     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
32000       intRetVal = 0
32010       JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
32020     End Select
32030   End If

EXITP:
32040   KeyCode = intRetVal
32050   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub revcode_ID_Exit(Cancel As Integer)

32100 On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ID_Exit"

32110   blnF4InvokedMouse = False

EXITP:
32120   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub taxcode_description_display_GotFocus()
' ** This control sits on top of the taxcode combo box.
' ** It's used to display the taxcode 'taxcode_description'
' ** when the combo box RowSource is showing the other set.
' ** CONTROL IS ENABLED, BUT LOCKED.

32200 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_display_GotFocus"

32210   JC_Msc_TaxCodeDescDisp blnNoMove, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.

EXITP:
32220   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub taxcode_description_display_KeyDown(KeyCode As Integer, Shift As Integer)
' ** Ordinarily, they shouldn't have time to press a key while on this field!

32300 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

32310   intRetVal = KeyCode
32320   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
32330   intShiftDown = (Shift And acShiftMask) > 0
32340   intAltDown = (Shift And acAltMask) > 0
32350   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
32360   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
32370     Select Case intRetVal
          Case vbKeyTab
32380       With Me
32390         intRetVal = 0
32400         If .taxcode.Locked = False Then
32410           .taxcode.SetFocus
32420         Else
32430           JC_Key_Sub Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
32440         End If
32450       End With
32460     Case vbKeyReturn
32470       With Me
32480         intRetVal = 0
32490         If .taxcode.Locked = False Then
32500           .taxcode.SetFocus
32510         Else
32520           strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
32530         End If
32540       End With
32550     Case vbKeyPageUp, vbKeyPageDown
32560       Select Case blnNoMove
            Case True
32570         blnNoMove = False
32580       Case False
32590         With Me
32600           If .posted = False And .journaltype = "Cost Adj." And .Cost <> 0@ And .assetno > 0& Then
32610             intRetVal = 0
32620           End If
32630         End With
32640         strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
32650       End Select
32660     End Select
32670   End If

        ' ** Shift keys.
32680   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
32690     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
32700       intRetVal = 0
32710       JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
32720     End Select
32730   End If

EXITP:
32740   KeyCode = intRetVal
32750   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub taxcode_lbl_DblClick(Cancel As Integer)

32800 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_lbl_DblClick"

32810   Select Case gblnIncomeTaxCoding
        Case True
32820     JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.
32830   Case False
          ' ** Disabled.
32840   End Select

EXITP:
32850   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub taxcode_Enter()

32900 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_Enter"

32910   With Me
32920     If IsNull(.taxcode) = True Then
32930       blnF4Invoked = True: strF4LastControl = "taxcode"
32940       ExpandCombo .taxcode  ' ** Module Procedure: modUtilities.
32950     Else
32960       If .taxcode.text = vbNullString Or .taxcode = 0& Then
32970         blnF4Invoked = True: strF4LastControl = "taxcode"
32980         ExpandCombo .taxcode  ' ** Module Procedure: modUtilities.
32990       ElseIf .taxcode = TAXID_INC Or .taxcode = TAXID_DED Then
33000         blnF4Invoked = True: strF4LastControl = "taxcode"
33010         ExpandCombo .taxcode  ' ** Module Procedure: modUtilities.
33020       End If
33030     End If
33040   End With

EXITP:
33050   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub taxcode_GotFocus()

33100 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_GotFocus"

33110   JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
33120   blnNoMove = False
33130   blnDontCommitTwice = False

EXITP:
33140   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub taxcode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

33200 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_MouseDown"

33210   blnF4InvokedMouse = True  ' ** Maybe, maybe not.

EXITP:
33220   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub taxcode_AfterUpdate()

33300 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_AfterUpdate"

        Dim blnPosted As Boolean

33310   With Me

          ' ** Don't automatically uncommit for a change in revcode_ID or taxcode.
33320     blnPosted = .posted

33330     JCol_TaxCode_AfterUpdate blnPosted, blnDontCommitTwice, strSaveMoveCtl, blnNextRec, _
            blnFromZero, blnNoMove, THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

33340   End With

EXITP:
33350   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub taxcode_NotInList(NewData As String, Response As Integer)

33400 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_NotInList"

33410   With Me
33420     If gblnClosing = False And gblnDeleting = False Then
33430       .taxcode = TAXID_INC
33440       blnNoMove = True
33450       MsgBox "Must select a valid Code from the list," & vbCrLf & _
              "or choose {no entry}.", vbInformation + vbOKOnly, "Invalid Entry"
33460       Response = acDataErrContinue
33470     End If
33480   End With

EXITP:
33490   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     If ERR.Number > 0 Then
590       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Else  ' ** Those huge negative error numbers are too big for the Integer in the Form_Error() parameters.
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End If
630     Resume EXITP

End Sub

Private Sub taxcode_KeyDown(KeyCode As Integer, Shift As Integer)

33500 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

33510   intRetVal = KeyCode
33520   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
33530   intShiftDown = (Shift And acShiftMask) > 0
33540   intAltDown = (Shift And acAltMask) > 0
33550   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
33560   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
33570     Select Case intRetVal
          Case vbKeyTab
33580       Select Case blnNoMove
            Case True
33590         blnNoMove = False
33600       Case False
33610         With Me
33620           intRetVal = 0
33630           If IsNull(.taxcode) = True Then
33640             Select Case .taxcode_type
                  Case TAXTYP_INC
33650               .taxcode = TAXID_INC
33660             Case TAXTYP_DED
33670               .taxcode = TAXID_DED
33680             End Select
33690           Else
33700             If IsNull(.taxcode.text) = True Then
33710               .taxcode.text = "{no entry}"
33720             Else
33730               If .taxcode.text = vbNullString Then
33740                 .taxcode.text = "{no entry}"
33750               End If
33760             End If
33770           End If
33780           blnDontCommitTwice = True
33790           JC_Key_Sub Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
33800         End With
33810       End Select
33820     Case vbKeyReturn
33830       Select Case blnNoMove
            Case True
33840         blnNoMove = False
33850       Case False
33860         With Me
33870           intRetVal = 0
33880           If IsNull(.taxcode) = True Then
33890             Select Case .taxcode_type
                  Case TAXTYP_INC
33900               .taxcode = TAXID_INC
33910             Case TAXTYP_DED
33920               .taxcode = TAXID_DED
33930             End Select
33940           Else
33950             If IsNull(.taxcode.text) = True Then
33960               .taxcode.text = "{no entry}"
33970             Else
33980               If .taxcode.text = vbNullString Then
33990                 .taxcode.text = "{no entry}"
34000               End If
34010             End If
34020           End If
34030           blnDontCommitTwice = True
34040           strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot)   ' ** Module Function: modJrnlCol_Keys.
34050         End With
34060       End Select
34070     Case vbKeyPageUp, vbKeyPageDown
34080       Select Case blnNoMove
            Case True
34090         blnNoMove = False
34100       Case False
34110         With Me
34120           If .posted = False And .journaltype = "Cost Adj." And .Cost <> 0@ And .assetno > 0& Then
34130             intRetVal = 0
34140           End If
34150         End With
34160         strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 4, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
34170       End Select
34180     End Select
34190   End If

        ' ** Shift keys.
34200   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
34210     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
34220       intRetVal = 0
34230       JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
34240     End Select
34250   End If

EXITP:
34260   KeyCode = intRetVal
34270   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub taxcode_Exit(Cancel As Integer)

34300 On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_Exit"

34310   blnF4InvokedMouse = False

EXITP:
34320   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Reinvested_lbl_DblClick(Cancel As Integer)

34400 On Error GoTo ERRH

        Const THIS_PROC As String = "Reinvested_lbl_DblClick"

34410   JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
34420   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Reinvested_GotFocus()

34500 On Error GoTo ERRH

        Const THIS_PROC As String = "Reinvested_GotFocus"

34510   With Me

34520     JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, Me  ' ** Module Procedure: modJrnlCol_Misc.
34530     blnNoMove = False

          ' ** Because this is the same color as the Detail BackColor,
          ' ** it should only show when the highlight bar is beneath it.
34540     .Reinvested_box.Visible = True

34550   End With

EXITP:
34560   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Reinvested_AfterUpdate()

34600 On Error GoTo ERRH

        Const THIS_PROC As String = "Reinvested_AfterUpdate"

34610   With Me

34620     blnReinvestment = False

34630     JCol_Reinvest_AfterUpdate blnReinvestment, strSaveMoveCtl, blnNextRec, _
            blnFromZero, THIS_PROC, That_Erl, That_Desc, Me  ' ** Module Procedure: modJrnlCol_Procs.
          '03/24/2017: CHANGED THAT_PROC TO THIS_PROC!

34640   End With

EXITP:
34650   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub Reinvested_KeyDown(KeyCode As Integer, Shift As Integer)

34700 On Error GoTo ERRH

        Const THIS_PROC As String = "Reinvested_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim blnContinue As Boolean
        Dim intRetVal As Integer

34710   intRetVal = KeyCode
34720   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
34730   intShiftDown = (Shift And acShiftMask) > 0
34740   intAltDown = (Shift And acAltMask) > 0
34750   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
34760   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
34770     Select Case intRetVal
          Case vbKeyTab
34780       With Me
34790         intRetVal = 0
34800         blnContinue = True
34810         Select Case .posted
              Case True
                ' ** Proceed normally.
34820         Case False
34830           Select Case blnReinvestment
                Case True
34840             blnContinue = False
34850           Case False
34860             strSaveMoveCtl = vbNullString
34870             cmdSave_Click  ' ** Procedure: Above.
34880             CommitRec Me, blnNextRec, blnFromZero  ' ** Module Function: modJrnlCol_Recs.
34890           End Select
34900         End Select
34910         If blnContinue = True Then
34920           JC_Key_Sub Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
                ' ** Keep us in the subform!
34930           DoCmd.SelectObject acForm, .Parent.Name, False
34940           .Parent.frmJournal_Columns_Sub.SetFocus
34950         End If  ' ** blnContinue.
34960       End With
34970     Case vbKeyReturn
34980       With Me
34990         intRetVal = 0
35000         Select Case .posted
              Case True
                ' ** Proceed normally.
35010         Case False
35020           Select Case blnReinvestment
                Case True
35030             blnContinue = False
35040           Case False
35050             strSaveMoveCtl = vbNullString
35060             cmdSave_Click  ' ** Procedure: Above.
35070             CommitRec Me, blnNextRec, blnFromZero  ' ** Module Function: modJrnlCol_Recs.
35080           End Select
35090         End Select
35100         If blnContinue = True Then
35110           strPageMoveCtl = JC_Key_Sub(Me, THIS_PROC, KeyCode, 0, blnNextRec, blnFromZero, blnToTaxLot)  ' ** Module Function: modJrnlCol_Keys.
35120         End If
35130       End With
35140     Case vbKeyPageUp, vbKeyPageDown
35150       strPageMoveCtl = "Reinvested"
35160     End Select
35170   End If

        ' ** Shift keys.
35180   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
35190     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
35200       intRetVal = 0
35210       JC_Key_Sub Me, THIS_PROC, KeyCode, 1, blnNextRec, blnFromZero, blnToTaxLot  ' ** Module Function: modJrnlCol_Keys.
35220     End Select
35230   End If

EXITP:
35240   KeyCode = intRetVal
35250   Exit Sub

ERRH:
560     intRetVal = 0
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Sub

Private Sub Reinvested_LostFocus()

35300 On Error GoTo ERRH

        Const THIS_PROC As String = "Reinvested_LostFocus"

35310   With Me
35320     .Reinvested_box.Visible = False
35330   End With

EXITP:
35340   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Private Sub journal_USER_lbl_DblClick(Cancel As Integer)

35400 On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_lbl_DblClick"

35410   JC_Sort_Now THIS_PROC, Me  ' ** Module Procedure: modJrnlCol_Sort.

EXITP:
35420   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Function ADateFromTDate() As Boolean

35500 On Error GoTo ERRH

        Const THIS_PROC As String = "ADateFromTDate"

        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String, lngTmp02 As Long, datTmp03 As Date
        Dim blnRetVal As Boolean

35510   blnRetVal = True

35520   With Me
35530     datTmp03 = CDate(.transdate)
35540     strTmp01 = CStr(CDbl(datTmp03))
35550     If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
35560     lngTmp02 = CLng(strTmp01)
35570     If lngTmp02 > CLng(Date) Then
35580       msgResponse = MsgBox("A Trade Date is required for this Journal transaction." & vbCrLf & vbCrLf & _
              "Click OK to use today's date." & vbCrLf & _
              "Click Cancel to enter your own", vbQuestion + vbOKCancel + vbDefaultButton2, "Trade Date Today")
35590       Select Case msgResponse
            Case vbOK
35600         datTmp03 = Now()
35610       Case Else
35620         blnRetVal = False
35630       End Select
35640     Else
35650       datTmp03 = (CDate(lngTmp02) + time)
35660     End If
35670     If blnRetVal = True Then
35680       .assetdate = datTmp03
35690       .assetdate_display = CDate(Format(datTmp03, "mm/dd/yyyy"))
35700     End If
35710   End With

EXITP:
35720   ADateFromTDate = blnRetVal
35730   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function ADateFromADateDisp() As Boolean

35800 On Error GoTo ERRH

        Const THIS_PROC As String = "ADateFromADateDisp"

        Dim strTmp01 As String, lngTmp02 As Long, datTmp03 As Date
        Dim blnRetVal As Boolean

35810   blnRetVal = True

35820   With Me
35830     datTmp03 = CDate(.assetdate_display)
35840     strTmp01 = CStr(CDbl(datTmp03))
35850     If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
35860     lngTmp02 = CLng(strTmp01)
35870     datTmp03 = (CDate(lngTmp02) + time)
35880     .assetdate = datTmp03
35890   End With

EXITP:
35900   ADateFromADateDisp = blnRetVal
35910   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function ADateDispFromADate() As Boolean

36000 On Error GoTo ERRH

        Const THIS_PROC As String = "ADateDispFromADate"

        Dim strTmp01 As String, lngTmp02 As Long, datTmp03 As Date
        Dim blnRetVal As Boolean

36010   blnRetVal = True

36020   With Me
36030     datTmp03 = .assetdate
36040     strTmp01 = CStr(CDbl(datTmp03))
36050     If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
36060     lngTmp02 = CLng(strTmp01)
36070     datTmp03 = CDate(lngTmp02)
36080     .assetdate_display = datTmp03
36090   End With

EXITP:
36100   ADateDispFromADate = blnRetVal
36110   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function ADateNull() As Boolean

36200 On Error GoTo ERRH

        Const THIS_PROC As String = "ADateNull"

        Dim blnRetVal As Boolean

36210   blnRetVal = True

36220   With Me
36230     If IsNull(.assetdate) = False Then
36240       .assetdate = Null
36250       .assetdate_display = Null
36260     End If
36270   End With

EXITP:
36280   ADateNull = blnRetVal
36290   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function PDateNull() As Boolean

36300 On Error GoTo ERRH

        Const THIS_PROC As String = "PDateNull"

        Dim blnRetVal As Boolean

36310   blnRetVal = True

36320   With Me
36330     If IsNull(.PurchaseDate) = False Then
36340       .PurchaseDate = Null
36350     End If
36360   End With

EXITP:
36370   PDateNull = blnRetVal
36380   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Sub StatusBar_Load()

36400 On Error GoTo ERRH

        Const THIS_PROC As String = "StatusBar_Load"

        Dim arr_varRetVal As Variant

        Const RV_ERR   As Integer = 0
        Const RV_SBARS As Integer = 1
        Const RV_POPUP As Integer = 2

36410 On Error Resume Next
36420   If lngStats = 0& Or IsEmpty(arr_varStat) Then
36430     arr_varRetVal = JC_Msc_SBar_Load(Me, True)  ' ** Module Function: modJrnlCol_Misc.
36440     If arr_varRetVal(RV_ERR, 0) = vbNullString Then
36450       Select Case IsEmpty(arr_varRetVal(RV_SBARS, 0))
            Case True
36460         arr_varStat = Empty
36470         lngStats = 0&
36480         blnNotPopup = False
36490       Case False
36500         Select Case IsNull(arr_varRetVal(RV_SBARS, 0))
              Case True
36510           arr_varStat = Empty
36520           lngStats = 0&
36530           blnNotPopup = False
36540         Case False
36550           arr_varStat = arr_varRetVal(RV_SBARS, 0)
36560           lngStats = (UBound(arr_varStat, 2) + 1&)
36570           blnNotPopup = arr_varRetVal(RV_POPUP, 0)
36580         End Select
36590       End Select
36600     End If
36610   End If
36620 On Error GoTo ERRH

EXITP:
36630   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Function StatusBar_Get() As Long
' ** Called by:
' **   modJrnlCol_Misc:
' **     JC_Msc_Pub_Reset()

36700 On Error GoTo ERRH

        Const THIS_PROC As String = "StatusBar_Get"

        Dim lngRetVal As Long

36710   lngRetVal = lngStats

EXITP:
36720   StatusBar_Get = lngRetVal
36730   Exit Function

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Function

Public Sub AddRec_Send(Optional varFromZero As Variant)
' ** Forward the request with the necessary subform variables.

36800 On Error GoTo ERRH

        Const THIS_PROC As String = "AddRec_Send"

36810   Select Case IsMissing(varFromZero)
        Case True
36820     AddRec Me, blnGTR_NoAdd, datPostingDate, strSaveMoveCtl, blnNoMove, lngNewJrnlColID  ' ** Module Procedure: modJrnlCol_Recs.
36830   Case False
36840     AddRec Me, blnGTR_NoAdd, datPostingDate, strSaveMoveCtl, blnNoMove, lngNewJrnlColID, varFromZero  ' ** Module Procedure: modJrnlCol_Recs.
36850   End Select

EXITP:
36860   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub DelRec_Send()
' ** Forward the request with the necessary subform variables.

36900 On Error GoTo ERRH

        Const THIS_PROC As String = "DelRec_Send"

36910   DelRec Me, blnGoneToReport, blnGoneToReport2, blnRecsTotUpdate  ' ** Module Procedure: modJrnlCol_Recs.

EXITP:
36920   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub NewRecAdd(varInput As Variant)

37000 On Error GoTo ERRH

        Const THIS_PROC As String = "NewRecAdd"

        Dim lngE As Long

37010   If IsNull(varInput) = False Then
37020     lngNewRecs = lngNewRecs + 1&
37030     lngE = lngNewRecs - 1&
37040     ReDim Preserve arr_varNewRec(N_ELEMS, lngE)
37050     arr_varNewRec(N_ID, lngE) = CLng(varInput)
37060     arr_varNewRec(N_CMTD, lngE) = CBool(False)
37070   End If

EXITP:
37080   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub DoTheDeed()
' ** Called by:
' **   modJrnlCol_Recs:
' **     JC_Rec_DelRec()

37100 On Error GoTo ERRH

        Const THIS_PROC As String = "DoTheDeed"

37110   With Me
37120     DoCmd.SelectObject acForm, .Parent.Name, False
37130     .Parent.frmJournal_Columns_Sub.SetFocus
37140     DoEvents
37150     DoCmd.RunCommand acCmdDeleteRecord
37160   End With

EXITP:
37170   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Function RecCnt() As Long

37200 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

37210   lngRetVal = 0&

37220   With Me
37230 On Error Resume Next
37240     Set rst = .RecordsetClone
37250     If ERR.Number = 0 Then
37260 On Error GoTo ERRH
37270       With rst
37280         If .BOF = True And .EOF = True Then
                ' ** Shouldn't really be here if that's the case.
37290         Else
37300           .MoveLast
37310           lngRetVal = .RecordCount
37320         End If
37330         .Close
37340       End With
37350     Else
37360 On Error GoTo ERRH
37370     End If
37380   End With

EXITP:
37390   Set rst = Nothing
37400   RecCnt = lngRetVal
37410   Exit Function

ERRH:
560     lngRetVal = 0&
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Sub MoveRec(lngRecsGoTo As AcCommand, Optional varRecID As Variant)

37500 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

37510   If IsMissing(varRecID) = True Then
37520     Select Case lngRecsGoTo
          Case acCmdRecordsGoToNew, acCmdRecordsGoToNext, acCmdRecordsGoToPrevious, acCmdRecordsGoToFirst, acCmdRecordsGoToLast
37530 On Error Resume Next
37540       DoCmd.RunCommand lngRecsGoTo
37550 On Error GoTo ERRH
37560     End Select
37570   Else
37580     If IsNull(varRecID) = False Then
37590       If varRecID > 0& Then
37600 On Error Resume Next
37610         Set rst = Me.RecordsetClone
37620         If ERR.Number = 0 Then
37630 On Error GoTo ERRH
37640           With rst
37650 On Error Resume Next
37660             .FindFirst "[JrnlCol_ID] = " & CStr(varRecID)  ' ** Some combination of the events produced the error:
37670             If ERR.Number = 0 Then                         ' **   3420 - Object invalid or no longer set.
37680 On Error GoTo ERRH                                         ' ** on this line. The added On Error's should cover it.
37690               If .NoMatch = False Then
37700                 Me.Bookmark = .Bookmark
37710               End If
37720               .Close
37730             Else
37740 On Error GoTo ERRH
37750             End If
37760           End With
37770         Else
37780 On Error GoTo ERRH
37790         End If
37800       Else
37810         Select Case lngRecsGoTo
              Case acCmdRecordsGoToNew, acCmdRecordsGoToNext, acCmdRecordsGoToPrevious, acCmdRecordsGoToFirst, acCmdRecordsGoToLast
37820           DoCmd.RunCommand lngRecsGoTo
37830         End Select
37840       End If
37850     Else
37860       Select Case lngRecsGoTo
            Case acCmdRecordsGoToNew, acCmdRecordsGoToNext, acCmdRecordsGoToPrevious, acCmdRecordsGoToFirst, acCmdRecordsGoToLast
37870         DoCmd.RunCommand lngRecsGoTo
37880       End Select
37890     End If
37900   End If

EXITP:
37910   Set rst = Nothing
37920   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub RecalcTots()

38000 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

38010   With Me
38020     blnRecsTotUpdate = True
38030     .TimerInterval = 2000&
38040   End With

EXITP:
38050   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Function FromZero_GetSet(blnGet As Boolean, Optional varZero As Variant) As Boolean
' ** Called by:
' **   modJrnlCol_Forms:
' **     JC_Frm_TaxLot()
' **     JC_Frm_Map_Return()
' **   modJrnlCol_Recs:
' **     JC_Rec_AddRec()
' **     JC_Rec_CommitAll()
' **     JC_Rec_CostAdj()

38100 On Error GoTo ERRH

        Const THIS_PROC As String = "FromZero_GetSet"

        Dim blnRetVal As Boolean

38110   Select Case blnGet
        Case True
38120     blnRetVal = blnFromZero
38130   Case False
38140     blnFromZero = CBool(varZero)
38150     blnRetVal = True
38160   End Select

EXITP:
38170   FromZero_GetSet = blnRetVal
38180   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function GoneToReport2_GetSet(blnGet As Boolean, Optional varGone2 As Variant) As Boolean
' ** Called by:
' **   frmJournal_Columns_TaxLot:
' **     cmdCancel_Click()

38200 On Error GoTo ERRH

        Const THIS_PROC As String = "GoneToReport2_GetSet"

        Dim blnRetVal As Boolean

38210   Select Case blnGet
        Case True
38220     blnRetVal = blnGoneToReport2
38230   Case False
38240     blnGoneToReport2 = CBool(varGone2)
38250     blnRetVal = True
38260   End Select

EXITP:
38270   GoneToReport2_GetSet = blnRetVal
38280   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function GTR_NoAdd_GetSet(blnGet As Boolean, Optional varNoAdd As Variant) As Boolean
' ** Called by:
' **   frmJournal_Columns:
' **     Form_Timer()
' **   modJrnlCol_Recs:
' **     JC_Rec_AddRec()

38300 On Error GoTo ERRH

        Const THIS_PROC As String = "GTR_NoAdd_GetSet"

        Dim blnRetVal As Boolean

38310   Select Case blnGet
        Case True
38320     blnRetVal = blnGTR_NoAdd
38330   Case False
38340     blnGTR_NoAdd = CBool(varNoAdd)
38350     blnRetVal = True
38360   End Select

EXITP:
38370   GTR_NoAdd_GetSet = blnRetVal
38380   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function NewJColID_GetSet(blnGet As Boolean, Optional varNewID As Variant) As Long
' ** Called by:
' **   frmJournal_Columns:
' **     Form_Timer()
' **   modJrnlCol_Recs:
' **     JC_Rec_AddRec()

38400 On Error GoTo ERRH

        Const THIS_PROC As String = "NewJColID_GetSet"

        Dim lngRetVal As Long

38410   Select Case blnGet
        Case True
38420     lngRetVal = lngNewJrnlColID
38430   Case False
38440     lngNewJrnlColID = CLng(varNewID)
38450     lngRetVal = 0&
38460   End Select

EXITP:
38470   NewJColID_GetSet = lngRetVal
38480   Exit Function

ERRH:
560     lngRetVal = -1&
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function NextRec_GetSet(blnGet As Boolean, Optional varNext As Variant) As Boolean
' ** Called by:
' **   modJrnlCol_Forms:
' **     JC_Frm_TaxLot()
' **     JC_Frm_Map_Return()
' **   modJrnlCol_Recs:
' **     JC_Rec_AddRec()
' **     JC_Rec_CommitAll()
' **     JC_Rec_CostAdj()

38500 On Error GoTo ERRH

        Const THIS_PROC As String = "NextRec_GetSet"

        Dim blnRetVal As Boolean

38510   Select Case blnGet
        Case True
38520     blnRetVal = blnNextRec
38530   Case False
38540     blnNextRec = CBool(varNext)
38550     blnRetVal = True
38560   End Select

EXITP:
38570   NextRec_GetSet = blnRetVal
38580   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function NoMove_GetSet(blnGet As Boolean, Optional varNoMove As Variant) As Boolean
' ** Called by:
' **   frmJournal_Columns:
' **     Form_Timer()
' **   modJrnlCol_Recs:
' **     JC_Rec_AddRec()

38600 On Error GoTo ERRH

        Const THIS_PROC As String = "NoMove_GetSet"

        Dim blnRetVal As Boolean

38610   Select Case blnGet
        Case True
38620     blnRetVal = blnNoMove
38630   Case False
38640     blnNoMove = CBool(varNoMove)
38650     blnRetVal = True
38660   End Select

EXITP:
38670   NoMove_GetSet = blnRetVal
38680   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function PostDate_GetSet(blnGet As Boolean, Optional varPostDate As Variant) As Date
' ** Called by:
' **   frmJournal_Columns:
' **     Form_Timer()
' **   modJrnlCol_Misc:
' **     JC_Msc_Pub_Reset()
' **   modJrnlCol_Recs:
' **     JC_Rec_AddRec()

38700 On Error GoTo ERRH

        Const THIS_PROC As String = "PostDate_GetSet"

        Dim datRetVal As Date

38710   Select Case blnGet
        Case True
38720     datRetVal = datPostingDate
38730   Case False
38740     datPostingDate = CDate(varPostDate)
38750     datRetVal = 0
38760   End Select

EXITP:
38770   PostDate_GetSet = datRetVal
38780   Exit Function

ERRH:
560     datRetVal = -1
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function RecsTotUpdate_GetSet(blnGet As Boolean, Optional varRecTot As Variant) As Boolean
' ** Called by:
' **   frmJournal_Columns_TaxLot:
' **     cmdCancel_Click()

38800 On Error GoTo ERRH

        Const THIS_PROC As String = "RecsTotUpdate_GetSet"

        Dim blnRetVal As Boolean

38810   Select Case blnGet
        Case True
38820     blnRetVal = blnRecsTotUpdate
38830   Case False
38840     blnRecsTotUpdate = CBool(varRecTot)
38850     blnRetVal = True
38860   End Select

EXITP:
38870   RecsTotUpdate_GetSet = blnRetVal
38880   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function SaveMoveCtl_GetSet(blnGet As Boolean, Optional varMoveCtl As Variant) As String
' ** Called by:
' **   frmJournal_Columns:
' **     Form_Timer()
' **   modJrnlCol_Recs:
' **     JC_Rec_AddRec()

38900 On Error GoTo ERRH

        Const THIS_PROC As String = "SaveMoveCtl_GetSet"

        Dim strRetVal As String

38910   Select Case blnGet
        Case True
38920     strRetVal = strSaveMoveCtl
38930   Case False
38940     strSaveMoveCtl = CStr(varMoveCtl)
38950     strRetVal = vbNullString
38960   End Select

EXITP:
38970   SaveMoveCtl_GetSet = strRetVal
38980   Exit Function

ERRH:
560     strRetVal = RET_ERR
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function TaxLot_GetSet(blnGet As Boolean, Optional varLot As Variant) As Boolean
' ** Called by:
' **   frmJournal_Columns:
' **     Form_Timer()
' **     cmdAssetNew_Click()

39000 On Error GoTo ERRH

        Const THIS_PROC As String = "TaxLot_GetSet"

        Dim blnRetVal As Boolean

39010   Select Case blnGet
        Case True
39020     blnRetVal = blnToTaxLot
39030   Case False
39040     blnToTaxLot = CBool(varLot)
39050     blnRetVal = True
39060   End Select

EXITP:
39070   TaxLot_GetSet = blnRetVal
39080   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function WarnZeroCash_GetSet(blnGet As Boolean, Optional varWarn As Variant) As Boolean
' ** Called by:
' **   modJrnlCol_Forms:
' **     JC_Frm_Map_Return()
' **   modJrnlCol_Keys:
' **     JC_Key_Sub()

39100 On Error GoTo ERRH

        Const THIS_PROC As String = "WarnZeroCash_GetSet"

        Dim blnRetVal As Boolean

39110   Select Case blnGet
        Case True
39120     blnRetVal = blnWarnZeroCost
39130   Case False
39140     blnWarnZeroCost = CBool(varWarn)
39150     blnRetVal = True
39160   End Select

EXITP:
39170   WarnZeroCash_GetSet = blnRetVal
39180   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function GTR_Emblem_GetSet(blnGet As Boolean, Optional varGTR As Variant) As Boolean
' ** Called by:
' **   frmJournal_Columns:
' **     Form_Timer()

39200 On Error GoTo ERRH

        Const THIS_PROC As String = "GTR_Emblem_GetSet"

        Dim blnRetVal As Boolean

39210   Select Case blnGet
        Case True
39220     blnRetVal = blnGTR_Emblem
39230   Case False
39240     blnGTR_Emblem = CBool(varGTR)
39250     blnRetVal = True
39260   End Select

EXITP:
39270   GTR_Emblem_GetSet = blnRetVal
39280   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Function GTRGone_GetSet(blnGet As Boolean, blnWentToReport As Boolean, lngJColID As Long) As Boolean
' ** Called by:
' **   frmJournal_Columns:
' **     Form_Timer()

39300 On Error GoTo ERRH

        Const THIS_PROC As String = "GTRGone_GetSet"

        Dim blnRetVal As Boolean

39310   Select Case blnGet
        Case True
39320     blnRetVal = blnGoneToReport
39330     lngJColID = lngGTR_ID
39340   Case False
39350     blnGoneToReport = blnWentToReport
39360     lngGTR_ID = lngJColID
39370   End Select

EXITP:
39380   GTRGone_GetSet = blnRetVal
39390   Exit Function

ERRH:
560     blnRetVal = False
570     Select Case ERR.Number
        Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Function

Public Function GoingToReport_GetSet(blnGet As Boolean, Optional varGoing As Variant) As Boolean
' ** Called by:
' **   frmJournal_Columns:
' **     Form_Timer()

39400 On Error GoTo ERRH

        Const THIS_PROC As String = "GoingToReport_GetSet"

        Dim blnRetVal As Boolean

39410   Select Case blnGet
        Case True
39420     blnRetVal = blnGoingToReport
39430   Case False
39440     blnGoingToReport = CBool(varGoing)
39450     blnRetVal = True
39460   End Select

EXITP:
39470   GoingToReport_GetSet = blnRetVal
39480   Exit Function

ERRH:
560     blnRetVal = False
570     THAT_PROC = THIS_PROC
580     That_Erl = Erl: That_Desc = ERR.description
590     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
600     Resume EXITP

End Function

Public Sub CalendarCheck()
' ** Called by:
' **   modJrnlCol_Misc:
' **     JC_Msc_Pub_Reset()

39500 On Error GoTo ERRH

        Const THIS_PROC As String = "CalendarCheck"

39510   If clsMonthClass Is Nothing Then
          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
39520     Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
39530     clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
39540     clsMonthClass.PositionAtCursor = True
39550   End If

EXITP:
39560   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub ColorCheck()
' ** Called by:
' **   modJrnlCol_Misc:
' **     JC_Msc_Pub_Reset()

39600 On Error GoTo ERRH

        Const THIS_PROC As String = "ColorCheck"

39610   With Me
39620     If CLR_HILITE = 0& Then CLR_HILITE = CLR_AC07
39630     .ctlBack1.ForeColor = CLR_HILITE
39640     .ctlBack2.ForeColor = CLR_HILITE
39650     CLR_DISABLED_FG = CLR_DKGRY
39660     CLR_DISABLED_BG = CLR_LTTEAL
39670   End With

EXITP:
39680   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub

Public Sub JrnlCol_Sub_Clear()

39700 On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlCol_Sub_Clear"

39710   lngStats = 0&
39720   arr_varStat = Empty
39730   lngNewRecs = 0&
39740   ReDim arr_varNewRec(N_ELEMS, 0)

EXITP:
39750   Exit Sub

ERRH:
560     THAT_PROC = THIS_PROC
570     That_Erl = Erl: That_Desc = ERR.description
580     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
590     Resume EXITP

End Sub
