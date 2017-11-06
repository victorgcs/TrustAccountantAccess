VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Rec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Rec"

'VGC 09/09/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Select Account:   A {cmbAccounts}
' **   By Number:        B {opgAccountSource_optNumber}
' **   Cancel:           C {cmdCancel}
' **   Date:             D {transdate}
' **   Trade Date:       E {assetdate_display}
' **   Asset Cusip:      I {opgAssetSource_optCusip}
' **   All:              L {opgAccountNumber_optAll}
' **   By Name:          M {opgAccountSource_optName}
' **   Asset Name:       N {opgAssetSource_optName}
' **   OK:               O {cmdOK}
' **   Per Share:        R {pershare}
' **   Select Asset:     S {cmbAssets}
' **   Amount:           T {Amount}
' **   Spec. Account:    U {opgAccountNumber_optSpecified}
' **   Asset Type:       Y {opgAssetSource_optType}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** MapType = "Received"

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of Calendar Class.
Public clsMonthClass As clsMonthCal

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Combo box column constants: cmbAssets.
'Private Const CBX_A_ASTNO  As Integer = 0  'assetno
Private Const CBX_A_TDESC  As Integer = 1  'totdesc
Private Const CBX_A_CUSIP  As Integer = 2  'cusip
Private Const CBX_A_DESC   As Integer = 3  'description
'Private Const CBX_A_TYPE   As Integer = 4  'assettype
Private Const CBX_A_CURRID As Integer = 5  'curr_id

' ** Combo box column constants: curr_id.
Private Const CBX_C_CURRID As Integer = 0  'curr_id
'Private Const CBX_C_CODE   As Integer = 1  'curr_code
'Private Const CBX_C_NAME   As Integer = 2  'curr_name
Private Const CBX_C_SYM    As Integer = 3  'currsym_symbol
Private Const CBX_C_DEC    As Integer = 4  'curr_decimal
'Private Const CBX_C_RATE1  As Integer = 5  'curr_rate1
Private Const CBX_C_RATE2  As Integer = 6  'curr_rate2
Private Const CBX_C_DATE   As Integer = 7  'curr_date

' ** Array: arr_varNoChar().
Private lngNoChars As Long, arr_varNoChar As Variant
Private Const N_CID   As Integer = 0
'Private Const N_CNAM  As Integer = 1
'Private Const N_FID   As Integer = 2
'Private Const N_FNAM  As Integer = 3
'Private Const N_NOCHR As Integer = 4

' ** Trying to force case on icash/ICash and pcash/PCash by use of a constant doesn't work; neither Private nor Public!

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean

Private strFrmt1 As String, strFrmt2 As String, intDec1 As Integer, intDec2 As Integer, lngBkClr As Long
Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private dblPerShare_Calc As Double, lngTpp As Long, blnCurrID As Boolean, lngCurrID As Long
Private datPostingDate As Date, intInputType As Integer
Private blnAssetUpdated As Boolean, blnAssetChecked As Boolean, blnDrivingMeBackwards As Boolean
Private strAccountNo_OU As String, datAssetDate_OldValue As Date, lngAssetDate_OldValue As Long, blnAssetDateChecked As Boolean
Private blnExitingAssetNo As Boolean, blnCancel As Boolean
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnCurrID = Pref_CurrID  ' ** Module Function: modPreferenceFuncs.

160           blnExitingAssetNo = False: blnCancel = False
170           blnAssetUpdated = False: blnAssetChecked = False: blnDrivingMeBackwards = False
180           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
190           blnCalendar2_Focus = False: blnCalendar2_MouseDown = False
200           blnGTR_Emblem = False: lngGTR_Stat = 0&
210           intInputType = 0
220           dblPerShare_Calc = 0#
230           gstrAccountNo = vbNullString: gstrAccountName = vbNullString

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
240           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

250           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

              ' ** Borrowing these variables from the Court Reports.
260           gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString

270           CLR_DISABLED_FG = CLR_DKGRY
280           CLR_DISABLED_BG = CLR_LTTEAL

290           Select Case gblnAccountNoWithType
              Case True
300             strAccountNo_OU = "99-INCOME O/U"
310           Case False
320             strAccountNo_OU = "INCOME O/U"
330           End Select

              ' ** EVENT CHECK: chkRememberMe!
340           cmbAccounts_AfterUpdate  ' ** Procedure: Below.

350           opgAccountNumber_AfterUpdate  ' ** Procedure: Below.
360           opgAccountSource_AfterUpdate  ' ** Procedure: Below.
370           chkRememberMe_AfterUpdate  ' ** Procedure: Below.
380           opgAssetSource_AfterUpdate  ' ** Procedure: Below.

390         Else
400           Cancel = -1
410         End If
420       Else
430         Cancel = -1
440       End If
450     End With

460     If Cancel = -1 Then
470       Beep
480       DoCmd.Hourglass False
490       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
500     End If

EXITP:
510     Exit Sub

ERRH:
520     DoCmd.Hourglass False
530     Select Case ERR.Number
        Case Else
540       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
550     End Select
560     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strTmp01 As String, dblTmp02 As Double

610     With Me

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
620       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
630       clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
640       clsMonthClass.PositionAtCursor = True

650       gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
660       datPostingDate = Date

670       Set dbs = CurrentDb
680       With dbs
690         Set grstPostingDate = .OpenRecordset("PostingDate", dbOpenDynaset, dbConsistent)
700         With grstPostingDate
710           .MoveFirst
720           .FindFirst "[Username] = '" & CurrentUser & "'"  ' ** Internal Access Function: Trust Accountant login.
730           If .NoMatch = False Then
740             If IsNull(![Posting_Date]) = False Then
750               datPostingDate = ![Posting_Date]
760             Else
770               .Edit
780               ![Posting_Date] = datPostingDate
790               .Update
800             End If
810           Else
820             .AddNew
830             ![Posting_Date] = datPostingDate
840             ![Username] = gstrJournalUser
850             .Update
860           End If
870           .Close
880         End With
            ' ** Empty tmpAccount.
890         Set qdf = .QueryDefs("qryMap_Rec_08_01")
900         qdf.Execute
910         Set qdf = Nothing
920         DoEvents
            ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
            ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
            ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
            ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
            ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
            ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
            ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
            ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
930         Set qdf = .QueryDefs("qryMap_Rec_08_02")
940         qdf.Execute
950         Set qdf = Nothing
960         DoEvents
970         .Close
980       End With
990       Set dbs = Nothing
1000      DoEvents

1010      blnAssetDateChecked = False
1020      If IsNull(.assetdate) = True Then
1030        datAssetDate_OldValue = 0
1040        lngAssetDate_OldValue = 0&
1050      Else
1060        If IsDate(.assetdate) = False Then
1070          datAssetDate_OldValue = 0
1080          lngAssetDate_OldValue = 0&
1090          .assetdate = Null
1100          .assetdate_display = Null
1110        Else
1120          datAssetDate_OldValue = .assetdate
1130          dblTmp02 = CDbl(datAssetDate_OldValue)
1140          strTmp01 = CStr(dblTmp02)
1150          If InStr(strTmp01, ".") > 0 Then
1160            strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
1170          End If
1180          lngAssetDate_OldValue = CLng(strTmp01)
1190          .assetdate_display = CDate(lngAssetDate_OldValue)  ' ** Display without timestamp.
1200        End If
1210      End If

1220      .transdate.DefaultValue = "#" & Format(datPostingDate, "mm/dd/yyyy") & "#"
1230      .shareface.ForeColor = CLR_DISABLED_FG
1240      .shareface.BackColor = CLR_DISABLED_BG

1250      IncludeCurrency  ' ** Procedure: Below.

1260    End With

1270    DoCmd.Hourglass False

EXITP:
1280    Set qdf = Nothing
1290    Set dbs = Nothing
1300    Exit Sub

ERRH:
1310    DoCmd.Hourglass False
1320    Select Case ERR.Number
        Case Else
1330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1340    End Select
1350    Resume EXITP

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

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   By Number:        B {opgAccountSource_optNumber}
        ' **   By Name:          M {opgAccountSource_optName}

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyEscape
1470        intRetVal = 0
1480        cmdCancel_Click  ' ** Procedure: Below.
1490      Case vbKeyF7
1500        intRetVal = 0
1510        cmdCalendar1_Click  ' ** Procedure: Below.
1520      Case vbKeyF8
1530        intRetVal = 0
1540        cmdCalendar2_Click  ' ** Procedure: Below.
1550      End Select
1560    End If

        ' ** Alt keys.
1570    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1580      Select Case intRetVal
          Case vbKeyB  ' ** Because the '_lbl' is empty.
1590        With Me
1600          intRetVal = 0
1610          If .opgAccountSource.Enabled = True Then
1620            .opgAccountSource.SetFocus
1630            .opgAccountSource = .opgAccountSource_optNumber.OptionValue
1640            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1650          Else
1660            Beep
1670          End If
1680        End With
1690      Case vbKeyM  ' ** Because the '_lbl' is empty.
1700        With Me
1710          intRetVal = 0
1720          If .opgAccountSource.Enabled = True Then
1730            .opgAccountSource.SetFocus
1740            .opgAccountSource = .opgAccountSource_optName.OptionValue
1750            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1760          Else
1770            Beep
1780          End If
1790        End With
1800      End Select
1810    End If

        ' ** Ctrl-Shift keys.
1820    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1830      Select Case intRetVal
          Case vbKeyF
1840        With Me
1850          intRetVal = 0
1860          .FocusHolder.SetFocus
1870        End With
1880      End Select
1890    End If

EXITP:
1900    KeyCode = intRetVal
1910    Exit Sub

ERRH:
1920    intRetVal = 0
1930    THAT_PROC = THIS_PROC
1940    That_Erl = Erl
1950    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1960    Resume EXITP

End Sub

Private Sub Form_Timer()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngAssetNo As Long, strAccountNo As String
        Dim Cancel As Integer
        Dim lngTmp03 As Long

        Static lngTmp01 As Long, lngTmp02 As Long

2010    With Me
2020      .TimerInterval = 0
2030      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
2040        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
2050      End If
2060      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
2070        lngGTR_Stat = lngGTR_Stat + 1&
2080        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
2090        lngTmp01 = lngTmp01 + 1&
2100        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
2110        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
2120        If lngTmp01 > 1& Then
2130          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
2140        End If
2150        If lngTmp01 = 24& Then
2160          lngTmp01 = 0&
2170          lngTmp02 = lngTmp02 + 1&
2180          If lngTmp02 < 8& Then
2190            .TimerInterval = 50&
2200          Else
2210            blnGTR_Emblem = False
2220          End If
2230        Else
2240          .TimerInterval = 50&
2250        End If
2260      End If  ' ** blnGTR_Emblem.
2270      If blnExitingAssetNo = True Then
            ' ** This was set up this way so that the user doesn't have to hit Cancel twice
            ' ** if they happened to have entered an asset, but then changed their mind.
2280        blnExitingAssetNo = False
2290        If blnCancel = False Then
2300          lngAssetNo = .cmbAssets
2310          strAccountNo = NullStrIfNull(.cmbAccounts)  ' ** Module Function: modStringFuncs.
2320          Set dbs = CurrentDb
2330          Select Case strCallingForm
              Case "frmJournal"
2340            Select Case .opgAccountNumber
                Case .opgAccountNumber_optAll.OptionValue
                  ' ** ActiveAssets, grouped and summed, by assetno, by specified [astno].
2350              Set qdf = dbs.QueryDefs("qryMap_Rec_07_01")
2360            Case .opgAccountNumber_optSpecified.OptionValue
                  ' ** ActiveAssets, grouped and summed, by assetno, by specified [astno], [actno].
2370              Set qdf = dbs.QueryDefs("qryMap_Rec_07_02")
2380              With qdf.Parameters
2390                ![actno] = strAccountNo
2400              End With
2410            End Select
2420          Case "frmJournal_Columns"
2430            Select Case .opgAccountNumber
                Case .opgAccountNumber_optAll.OptionValue
                  ' ** ActiveAssets, grouped and summed, by assetno, by specified [astno].
2440              Set qdf = dbs.QueryDefs("qryJournal_Columns_35_03")
2450            Case .opgAccountNumber_optSpecified.OptionValue
                  ' ** ActiveAssets, grouped and summed, by assetno, by specified [astno], [actno].
2460              Set qdf = dbs.QueryDefs("qryJournal_Columns_35_03_01")
2470              With qdf.Parameters
2480                ![actno] = strAccountNo
2490              End With
2500            End Select
2510          End Select
2520          With qdf.Parameters
2530            ![astno] = lngAssetNo
2540          End With
2550          Set rst = qdf.OpenRecordset
2560          If rst.BOF = True And rst.EOF = True Then
                ' ** Shouldn't happen.
2570          Else
2580            blnAssetUpdated = False
2590            blnAssetChecked = True
2600            .TotalShares = rst![sumsf]
2610            .shareface = Round(rst![sumsf], 4)
2620            If .amount.Enabled = True Then
2630              dblPerShare_Calc = (Nz(.amount, 0) / IIf(Nz(.shareface, 0) = 0, 1, .shareface))
2640              .pershare = dblPerShare_Calc  '5 dec.
2650  On Error Resume Next
2660              .amount.SetFocus
2670  On Error GoTo ERRH
2680            Else
2690              .amount = (.pershare * IIf(Nz(.shareface, 0) = 0, 1, .shareface))
2700  On Error Resume Next
2710              .pershare.SetFocus
2720  On Error GoTo ERRH
2730            End If
2740          End If
2750          rst.Close
2760          dbs.Close
2770        End If
2780      End If
2790      If gblnGoToReport = True Then
2800        DoCmd.Hourglass True  ' ** Make sure it's still running.
2810        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48& And lngGTR_Stat < 96&) Or _
                (lngGTR_Stat > 96& And lngGTR_Stat < 120&) Or (lngGTR_Stat > 120&)) Then
              ' ** Just let the emblem go.
2820        Else
2830          DoCmd.Hourglass True  ' ** Make sure it's still running.
2840          If blnGTR_Emblem = False Then
2850            blnGTR_Emblem = True
2860            lngTmp01 = 0&: lngTmp02 = 0&
2870            lngTmp03 = (.GoToReport_Emblem_01_img.Width + (8& * lngTpp))
2880            .GoToReport_arw_ok_img.Left = (.GoToReport_arw_ok_img.Left - lngTmp03)
2890            .cmdOK.Left = (.cmdOK.Left - lngTmp03)
2900            .cmdCancel.Left = (.cmdCancel.Left - lngTmp03)
2910          End If
2920          DoEvents
2930          Select Case blnGoingToReport
              Case True
2940            Select Case blnGoingToReport2
                Case True
2950              Select Case .GoToReport_arw_ok_img.Visible
                  Case True
2960                blnGoingToReport2 = False
2970                .TimerInterval = 50&
2980                cmdOK_Click  ' ** Procedure: Below.
2990              Case False
3000                .amount.SetFocus
3010                .amount = 100
3020                Cancel = 0
3030                amount_Exit Cancel  ' ** Procedure: Below.
3040                .assetdate = Now()
3050                .assetdate_display = Date
3060                DoEvents
3070                .GoToReport_arw_ok_img.Visible = True
3080                .GoToReport_arw_map_amt_img.Visible = False
3090                .assetdate_display_lbl.Visible = True
3100                .assetdate_display_lbl2.Visible = True
                    '.assetdate_display_lbl_box.Visible = True
3110                .GoToReport_arw_map_per_img.Visible = False
3120                .GoToReport_lbl_amtper.Visible = False
3130                .GoToReport_lin_hline01.Visible = False
3140                .GoToReport_lin_hline02.Visible = False
3150                .GoToReport_lbl_arrow01.Visible = False
3160                .GoToReport_lbl_arrow02.Visible = False
3170                If strCallingForm = "frmJournalColums" Then
                      ' ** Don't know why this continues to light up.
3180                  Forms(strCallingForm).cmdSpecPurp_Purch_MapSplit_LostFocus  ' ** Form Procedure: frmJournal_Columns.
3190                End If
3200                DoEvents
3210                .TimerInterval = 50&
3220              End Select
3230            Case False
3240              blnGoingToReport2 = True
3250              lngAssetNo = GetDivAsset  ' ** Module Function: modGoToReportFuncs.
3260              If lngAssetNo > 0& Then
3270                .cmbAssets = lngAssetNo
3280                cmbAssets_AfterUpdate  ' ** Procedure: Below.
3290                DoEvents
3300                .GoToReport_arw_map_amt_img.Visible = True
3310                .GoToReport_arw_map_per_img.Visible = True
3320                .assetdate_display_lbl.Visible = False
3330                .assetdate_display_lbl2.Visible = False
                    '.assetdate_display_lbl_box.Visible = False
3340                .GoToReport_lbl_amtper.Visible = True
3350                .GoToReport_lin_hline01.Visible = True
3360                .GoToReport_lin_hline02.Visible = True
3370                .GoToReport_lbl_arrow01.Visible = True
3380                .GoToReport_lbl_arrow02.Visible = True
3390                .GoToReport_arw_map_asset_img.Visible = False
3400                .opgAssetSource.Visible = True
3410                .opgAssetSource_box.Visible = True
3420                .TotalShares = 4000
3430                .shareface = 4000
3440                .FocusHolder.SetFocus  ' ** So they don't see a flashing cursor.
3450                DoEvents
3460                .TimerInterval = 50&
3470              Else
3480                .TimerInterval = 0&
3490                Select Case garr_varGoToReport(GTR_FRM2)
                    Case "frmJournal"
3500                  Forms(garr_varGoToReport(GTR_FRM2)).GTREmblem_Set False  ' ** Form Procedure: frmJournal.
3510                Case "frmJournal_Columns"
3520                  Forms(garr_varGoToReport(GTR_FRM2)).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
3530                End Select
3540                gblnGoToReport = False
3550                blnGoingToReport = False
3560                blnGoingToReport2 = False
3570                GTREmblem_Off  ' ** Procedure: Below.
3580                .GoToReport_arw_map_asset_img.Visible = False
3590                .opgAssetSource.Visible = True
3600                .opgAssetSource_box.Visible = True
3610                Beep
3620                DoCmd.Hourglass False
3630                MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                      "There are insufficient asset holdings to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
3640              End If
3650            End Select
3660          Case False
3670            blnGoingToReport = True
3680            .opgAssetSource.Visible = False
3690            .opgAssetSource_box.Visible = False
3700            .GoToReport_arw_map_asset_img.Visible = True
3710            .FocusHolder.SetFocus  ' ** So they don't see a flashing cursor.
3720            DoEvents
3730            Select Case garr_varGoToReport(GTR_FRM2)
                Case "frmJournal"
3740              Forms(garr_varGoToReport(GTR_FRM2)).GTREmblem_Set False  ' ** Form Procedure: frmJournal.
3750            Case "frmJournal_Columns"
3760              Forms(garr_varGoToReport(GTR_FRM2)).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
3770            End Select
3780            .TimerInterval = 50&
3790          End Select
3800        End If
3810      End If

3820    End With  ' ** Me.

EXITP:
3830    Set rst = Nothing
3840    Set qdf = Nothing
3850    Set dbs = Nothing
3860    Exit Sub

ERRH:
3870    Select Case ERR.Number
        Case Else
3880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3890    End Select
3900    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim strControlName As String

4010    strControlName = vbNullString

4020    If DataErr <> 0 Then
4030      Select Case DataErr
          Case 13, 94, 2108, 2110, 2116, 2135, 2169, 2474, 2501, 3020, 3163, 3314, 3315, 7753, 8519
            ' ** 13  : Type mismatch.
            ' ** 94  : Invalid use of Null.
            ' ** 2108: You must save the field before you execute the GoToControl action, the GoToControl method, or the SetFocus method.
            ' ** 2110: Microsoft Access can't move the focus to the control |.
            ' ** 2116: The value in the field or record violates the validation rule for the record or field.
            ' ** 2135: This property is read-only and can't be set.
            ' ** 2169: You can't save this record at this time.
            ' ** 2474: The expression you entered requires the control to be in the active window.
            ' ** 2501: The '|' action was Canceled.
            ' ** 3020: Update or CancelUpdate without AddNew or Edit.
            ' ** 3163: The field is too small to accept the amount of data you attempted to add.
            ' ** 3314: The field '|' cannot contain a Null value because the Required property for this field is set to True.
            ' ** 3315: Field '|' cannot be a zero-length string.
            ' ** 7753: The value you entered doesn't meet the validation rule defined for the field or control.
            ' ** 8519: You are about to delete | record(s).
            ' ** Do nothing.
4040      Case 2046  ' ** The command or action isn't available now (first or last record).
4050        DoCmd.CancelEvent
4060      Case 2113  ' ** The value you entered isn't valid for this field.
4070        Select Case Me.ActiveControl.Name
            Case "shareface"
4080          strControlName = "Shares"
4090        Case "pershare"
4100          strControlName = "Rate"
4110        Case "amount"
4120          strControlName = "Amount"
4130        Case "transdate"
4140          strControlName = vbNullString
4150          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
4160          Response = acDataErrContinue
4170          DoCmd.CancelEvent
4180          Exit Sub
4190        Case Else
4200          strControlName = vbNullString
4210        End Select
4220        If strControlName <> vbNullString Then
4230          MsgBox "You must enter a numeric value in the " & strControlName & " field.", vbInformation + vbOKOnly, "Invalid Entry"
4240          Response = acDataErrContinue
4250          DoCmd.CancelEvent
4260        End If
4270      Case 2237  ' ** The text you entered isn't an item in the list.
4280        DoCmd.CancelEvent
4290      Case 2279  ' ** The value you entered isn't appropriate for the input mask '|' specified for this field.
4300        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
4310        Response = acDataErrContinue
4320        DoCmd.CancelEvent
4330      Case 2580  ' ** The record source '|' specified on this form or report does not exist.
            ' ** Ignore! This form has no RecordSource.
4340        Response = acDataErrContinue
4350      Case 3162  ' ** You tried to assign the Null value to a variable that is not a Variant data type.
4360        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
            ' ** Do nothing.
4370      Case Else
4380        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
4390        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
4400      End Select
4410    End If

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

4510    With Me

4520      If .chkRememberMe = False Then
4530        .cmbAccounts = Null
4540      End If

4550      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

4560      If Not clsMonthClass Is Nothing Then
4570        If clsMonthClass.IsCalendar Then
4580          Cancel = -1
4590        Else
4600          Set clsMonthClass = Nothing
4610        End If
4620      End If

4630    End With

EXITP:
4640    Set qdf = Nothing
4650    Set dbs = Nothing
4660    Exit Sub

ERRH:
4670    DoCmd.Hourglass False
4680    Select Case ERR.Number
        Case Else
4690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4700    End Select
4710    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4810    With Me
4820      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
4830        Select Case blnCalendar1_Focus
            Case True
4840          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
4850          .cmdCalendar1_raised_img.Visible = False
4860        Case False
4870          .cmdCalendar1_raised_img.Visible = True
4880          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
4890        End Select
4900        .cmdCalendar1_raised_focus_dots_img.Visible = False
4910        .cmdCalendar1_raised_focus_img.Visible = False
4920        .cmdCalendar1_sunken_focus_dots_img.Visible = False
4930        .cmdCalendar1_raised_img_dis.Visible = False
4940      End If
4950      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
4960        Select Case blnCalendar2_Focus
            Case True
4970          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
4980          .cmdCalendar2_raised_img.Visible = False
4990        Case False
5000          .cmdCalendar2_raised_img.Visible = True
5010          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
5020        End Select
5030        .cmdCalendar2_raised_focus_dots_img.Visible = False
5040        .cmdCalendar2_raised_focus_img.Visible = False
5050        .cmdCalendar2_sunken_focus_dots_img.Visible = False
5060        .cmdCalendar2_raised_img_dis.Visible = False
5070      End If
5080    End With

EXITP:
5090    Exit Sub

ERRH:
5100    Select Case ERR.Number
        Case Else
5110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5120    End Select
5130    Resume EXITP

End Sub

Private Sub cmdOK_Click()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        'Dim cnxn As ADODB.Connection, rsx1 As ADODB.Recordset  ' ** Early binding.
        Dim cnxn As Object, rsx1 As Object                      ' ** Late binding.
        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim dblShareface As Double, dblAmount As Double, dblPerShare As Double
        Dim dblSumPCash As Double, dblSumShareface As Double, dblOverUnder As Double
        Dim datTransDate As Date, datAssetDate As Date, strMapType As String
        Dim lngAssetNo As Long, lngTaxcode As Long
        Dim strDocName As String
        Dim lngMaps As Long, arr_varMap As Variant
        Dim lngAdjs As Long, arr_varAdj As Variant
        Dim lngRecs As Long
        Dim blnContinue As Boolean, blnFound As Boolean, blnSkip As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String, lngTmp02 As Long, lngTmp03 As Long, dblTmp04 As Double, dblTmp05 As Double
        Dim lngX As Long, lngY As Long

        ' ** Array: arr_varMap().
        Const M_ACTNO   As Integer = 0
        Const M_ASTNO   As Integer = 1
        'Const M_SHARES  As Integer = 2
        'Const M_SHARESR As Integer = 3
        Const M_PER     As Integer = 4
        'Const M_ICASH   As Integer = 5
        'Const M_ICASHR  As Integer = 6

        ' ** Array: arr_varAdj().
        Const A_ID    As Integer = 0
        'Const A_ACTNO As Integer = 1
        'Const A_ASTNO As Integer = 2
        'Const A_PCASH As Integer = 3
        Const A_PCADJ As Integer = 4

5210    blnContinue = True

5220    DoCmd.Hourglass True
5230    DoEvents

5240    dblShareface = 0#
5250    dblAmount = 0#
5260    dblSumPCash = 0#
5270    dblSumShareface = 0#
5280    dblPerShare = 0#
5290    dblOverUnder = 0#
5300    lngAssetNo = 0&
5310    datTransDate = 0
5320    datAssetDate = 0
5330    lngTaxcode = 0&

5340    With Me

          ' ** Step 1: Validation.
5350      .Map_Status_lbl.Caption = " 1 of 22: Validation . . ."
          '.Map_Status_lbl.Visible = True
5360      DoEvents

5370      Select Case .opgAccountNumber
          Case .opgAccountNumber_optAll.OptionValue
            ' ** Let it pass.
5380        gstrAccountNo = vbNullString: gstrAccountName = vbNullString
5390      Case .opgAccountNumber_optSpecified.OptionValue
5400        Select Case IsNull(.cmbAccounts)
            Case True
5410          gstrAccountNo = vbNullString: gstrAccountName = vbNullString
5420          blnContinue = False
5430          DoCmd.Hourglass False
5440          MsgBox "Please select a valid account.", vbInformation + vbOKOnly, "Entry Required"
5450          .cmbAccounts.SetFocus
5460        Case False
              ' ** Let it pass.
5470          gstrAccountNo = .cmbAccounts.Column(0)
5480          gstrAccountName = .cmbAccounts.Column(2)
5490        End Select
5500      End Select

5510      If blnContinue = True Then
5520        If IsNull(.cmbAssets) = True Then
5530          blnContinue = False
5540          DoCmd.Hourglass False
5550          MsgBox "Please select a valid asset.", vbInformation + vbOKOnly, "Entry Required"
5560          .cmbAssets.SetFocus
5570        Else
5580          If Trim(.cmbAssets) = vbNullString Then
5590            blnContinue = False
5600            DoCmd.Hourglass False
5610            MsgBox "Please select a valid asset.", vbInformation + vbOKOnly, "Entry Required"
5620            .cmbAssets.SetFocus
5630          Else
5640            lngAssetNo = .cmbAssets
5650            If Nz(.shareface, 0) = 0 Then
5660              blnContinue = False
5670              DoCmd.Hourglass False
5680              MsgBox "Share/Face must be a number greater than zero.", vbInformation + vbOKOnly, "Entry Required"
5690              .cmbAssets.SetFocus
5700            Else
5710              If IsNull(.amount) = True Then
5720                blnContinue = False
5730                DoCmd.Hourglass False
5740                MsgBox "Amount must be a number greater than zero.", vbInformation + vbOKOnly, "Entry Required"
5750                .amount.SetFocus
5760              Else
5770                If IsNull(.pershare) = True Then
5780                  blnContinue = False
5790                  DoCmd.Hourglass False
5800                  MsgBox "Per Share must be greater than zero to continue.", vbInformation + vbOKOnly, "Entry Required"
5810                  .pershare.SetFocus
5820                Else
5830                  If Not IsDate(.assetdate) Then
5840                    blnContinue = False
5850                    DoCmd.Hourglass False
5860                    MsgBox "Please enter a valid trade date.", vbInformation + vbOKOnly, "Entry Required"
5870                    .assetdate.SetFocus
5880                  Else
5890                    If Compare_DateA_DateB(CDate(.assetdate), ">", Date) = True Then  ' ** Module Function: modStringFuncs.
5900                      blnContinue = False
5910                      DoCmd.Hourglass False
5920                      MsgBox "Future trade dates are not allowed.", vbInformation + vbOKOnly, "Invalid Date"
5930                    Else
5940                      If Not IsDate(.transdate) Then
5950                        blnContinue = False
5960                        DoCmd.Hourglass False
5970                        MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Date"
5980                        .transdate.SetFocus
5990                      Else
                            ' ** Allow posting up to 1 month into the future.
6000                        If .transdate > DateAdd("m", 1, Date) Then
6010                          blnContinue = False
6020                          DoCmd.Hourglass False
6030                          MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
6040                        Else
6050                          If .amount = 0 And .pershare = 0 Then
6060                            blnContinue = False
6070                            DoCmd.Hourglass False
6080                            MsgBox "Either Amount or Per Share must be greater than zero to continue.", vbInformation + vbOKOnly, "Entry Required"
6090                            If .amount.Enabled = False Or .pershare.Enabled = False Then
6100                              .amount = Null
6110                              .amount.Enabled = True
6120                              .amount.Locked = False
6130                              .amount.ForeColor = CLR_BLK
6140                              .pershare = Null
6150                              .pershare.Enabled = True
6160                              .pershare.Locked = False
6170                              .pershare.ForeColor = CLR_BLK
6180                              If lngCurrID = 150& Then
6190                                .amount.BackColor = CLR_WHT
6200                                .pershare.BackColor = CLR_WHT
6210                              Else
6220                                .amount.BackColor = lngBkClr
6230                                .pershare.BackColor = lngBkClr
6240                              End If
6250                            End If
6260                            .amount.SetFocus
6270                          End If
6280                        End If
6290                      End If
6300                    End If
6310                  End If
6320                End If
6330              End If
6340            End If
6350          End If
6360        End If
6370      End If  ' ** blnContinue.
6380      DoEvents

6390      If blnContinue = True Then

6400        DoCmd.Hourglass True
6410        DoEvents

6420        Set dbs = CurrentDb

            ' ** Step 2: Empty Journal Map.
6430        .Map_Status_lbl.Caption = " 2 of 22: Empty temporary table 1 . . ."
6440        DoEvents

6450        Select Case strCallingForm
            Case "frmJournal"
              ' ** Empty Journal Map.
6460          Set qdf = dbs.QueryDefs("qryMap_01")
6470        Case "frmJournal_Columns"
              ' ** Empty Journal Map.
6480          Set qdf = dbs.QueryDefs("qryJournal_Columns_35_04")
6490        End Select
6500        qdf.Execute
6510        DoEvents

            ' ** Step 3: Set Tax Code.
6520        .Map_Status_lbl.Caption = " 3 of 22: Set Tax Code . . ."
6530        DoEvents

6540        Select Case strCallingForm
            Case "frmJournal"
              ' ** MasterAsset, linked to AssetType, by specified [astno].
6550          Set qdf = dbs.QueryDefs("qryMap_Rec_06")
6560        Case "frmJournal_Columns"
              ' ** MasterAsset, linked to AssetType, by specified [astno].
6570          Set qdf = dbs.QueryDefs("qryJournal_Columns_35_01")
6580        End Select
6590        With qdf.Parameters
6600          ![astno] = lngAssetNo
6610        End With
6620        Set rst = qdf.OpenRecordset
6630        If rst.BOF = True And rst.EOF = True Then
              ' ** Shouldn't happen.
6640        Else
6650          rst.MoveFirst
6660          lngTaxcode = rst![taxcode]
6670        End If
6680        .taxcode = lngTaxcode
6690        DoEvents

            ' ** Step 4: Calculate PerShare/Amount.
6700        .Map_Status_lbl.Caption = " 4 of 22: Calculate PerShare/Amount . . ."
6710        DoEvents

6720        dblShareface = .TotalShares
6730        Select Case intInputType
            Case 1  ' ** User entered an Amount.
6740          dblAmount = .amount
6750          dblPerShare = (dblAmount / dblShareface)
6760        Case 2  ' ** User entered a PerShare.
6770          dblPerShare = .pershare
6780          dblAmount = (dblShareface * dblPerShare)
6790        Case Else  ' ** Shouldn't be an else!
6800          dblAmount = .amount
6810          dblPerShare = (dblAmount / dblShareface)
6820        End Select
6830        datTransDate = .transdate
6840        datAssetDate = .assetdate
6850        strMapType = .maptype  ' ** JournalType.
6860        DoEvents

            ' ** Step 5: Check for closed accounts.
6870        .Map_Status_lbl.Caption = " 5 of 22: Check for closed accounts . . ."
6880        DoEvents

6890        Select Case .opgAccountNumber
            Case .opgAccountNumber_optAll.OptionValue
              ' ** Check for closed accounts holding assets.
              ' ** qryMap_07 (qryMap_06 (ActiveAssets, linked to Account, just closed accounts, by specified [astno]),
              ' ** grouped by accountno, with cnt_astno), grouped and summed, with cnt_actno, cnt_astno.
6900          Set qdf = dbs.QueryDefs("qryMap_08")
6910          With qdf.Parameters
6920            ![astno] = Me.cmbAssets
6930          End With
6940          Set rst = qdf.OpenRecordset
6950          If rst.BOF = True And rst.EOF = True Then
                ' ** No closed accounts holding assets.
6960            msgResponse = vbYes
6970            rst.Close
6980          Else
6990            rst.MoveFirst
7000            If IsNull(rst![cnt_actno]) = True And IsNull(rst![cnt_astno]) = True Then
7010              msgResponse = vbYes
7020              rst.Close
7030            Else
7040              If rst![cnt_actno] = 0 And IsNull(rst![cnt_astno]) = True Then
7050                msgResponse = vbYes
7060                rst.Close
7070              ElseIf rst![cnt_actno] = 0 And rst![cnt_astno] = 0 Then
7080                msgResponse = vbYes
7090                rst.Close
7100              Else
7110                lngTmp02 = rst![cnt_actno]
7120                lngTmp03 = rst![cnt_astno]
7130                rst.Close
                    ' ** qryMap_06 (ActiveAssets, linked to Account, just closed accounts,
                    ' ** by specified [astno]), grouped by accountno, with cnt_astno.
7140                Set qdf = dbs.QueryDefs("qryMap_07")
7150                With qdf.Parameters
7160                  ![astno] = lngAssetNo
7170                End With
7180                Set rst = qdf.OpenRecordset
7190                rst.MoveFirst
7200                If lngTmp02 = 1& Then
7210                  strTmp01 = "There is 1 closed account currently holding this asset:" & vbCrLf & _
                        "    " & rst![accountno] & "  " & rst![shortname] & vbCrLf
7220                Else
7230                  strTmp01 = "There are " & CStr(lngTmp02) & " closed accounts currently holding this asset:" & vbCrLf
7240                  For lngX = 1& To lngTmp02
7250                    strTmp01 = strTmp01 & "    " & rst![accountno] & "  " & rst![shortname] & vbCrLf
7260                    If lngX < lngTmp02 Then rst.MoveNext
7270                  Next
7280                End If
7290                strTmp01 = strTmp01 & "If you choose to continue with this mapping," & vbCrLf & _
                      "the results may be inaccurate." & vbCrLf & vbCrLf & "Do you wish to continue?"
7300                DoCmd.Hourglass False
7310                Beep
7320                msgResponse = MsgBox(strTmp01, vbQuestion + vbYesNo + vbDefaultButton2, "Closed Accounts Holding Assets")
7330              End If
7340            End If
7350          End If
7360        Case .opgAccountNumber_optSpecified.OptionValue
7370          msgResponse = vbYes
7380        End Select
7390        DoEvents

7400        If msgResponse = vbYes Then

7410          DoCmd.Hourglass True
7420          DoEvents

              ' ** Step 6: Collect Active Assets.
7430          .Map_Status_lbl.Caption = " 6 of 22: Collect Active Assets . . ."
7440          DoEvents

7450          Select Case strCallingForm
              Case "frmJournal"
7460            Select Case .opgAccountNumber
                Case .opgAccountNumber_optAll.OptionValue
                  ' ** Append ActiveAssets to Journal Map, by specified [astno], [astdat], [jrntyp], [persh], [transdat].  #curr_id
7470              Set qdf = dbs.QueryDefs("qryMap_Rec_01")
7480            Case .opgAccountNumber_optSpecified.OptionValue
                  ' ** Append ActiveAssets to Journal Map, by specified [astno], [astdat], [jrntyp], [persh], [transdat], [actno].  #curr_id
7490              Set qdf = dbs.QueryDefs("qryMap_Rec_01a")
7500              With qdf.Parameters
7510                ![actno] = gstrAccountNo
7520              End With
7530            End Select
7540          Case "frmJournal_Columns"
7550            Select Case .opgAccountNumber
                Case .opgAccountNumber_optAll.OptionValue
                  ' ** Append ActiveAssets to Journal Map, by specified [astno], [astdat], [jrntyp], [persh], [transdat].  #curr_id
7560              Set qdf = dbs.QueryDefs("qryJournal_Columns_35_02")
7570            Case .opgAccountNumber_optSpecified.OptionValue
                  ' ** Append ActiveAssets to Journal Map, by specified [astno], [astdat], [jrntyp], [persh], [transdat], [actno].  #curr_id
7580              Set qdf = dbs.QueryDefs("qryJournal_Columns_35_02a")
7590              With qdf.Parameters
7600                ![actno] = gstrAccountNo
7610              End With
7620            End Select
7630          End Select
7640          With qdf.Parameters
7650            ![astno] = lngAssetNo
7660            ![astdat] = datAssetDate
7670            ![jrntyp] = strMapType
7680            ![persh] = dblPerShare  ' ** Used for both pcash and pershare.
7690            ![transdat] = datTransDate
7700          End With
7710          qdf.Execute
7720          dbs.Close
7730          Set qdf = Nothing
7740          Set dbs = Nothing
7750          DoEvents

7760          blnSkip = True
              ' ################################################################################
7770          If blnSkip = False Then
                'Set rsx1 = New ADODB.Recordset             ' ** Early binding.
7780            Set rsx1 = CreateObject("ADODB.Recordset")  ' ** Late binding.
7790            Set cnxn = CurrentProject.Connection
7800            rsx1.Open "journal map", cnxn, adOpenKeyset, adLockOptimistic, adCmdTableDirect
7810            With rsx1
7820              .MoveLast
7830              lngRecs = .RecordCount
7840              .MoveFirst
7850              For lngX = 1& To lngRecs
7860                dblTmp04 = Nz(.Fields("pcash"), 0)
7870                If ((dblTmp04 > 0.009) Or (dblTmp04 < -0.009)) Then
7880                  dblTmp04 = dbl_Round(dblTmp04, 3)  ' ** Module Function: modUtilities.
7890                Else
7900                  dblTmp04 = 0#
7910                End If
7920                If dblTmp04 <> 0# Then
7930                  dblTmp05 = Nz(.Fields("shareface"), 0)
7940                  If ((dblTmp05 > 0.00009) Or (dblTmp05 < -0.00009)) Then
7950                    dblTmp05 = dbl_Round(dblTmp05, 5)  ' ** Module Function: modUtilities.
7960                  Else
7970                    dblTmp05 = 0#
7980                    dblTmp04 = 0#  ' ** Also zero-out the pcash.
7990                  End If
8000                End If
8010                .Fields("pcash") = dblTmp04
8020                .Fields("shareface") = dblTmp05
8030                .Update
8040                dblSumPCash = dblSumPCash + dblTmp04
8050                dblSumShareface = dblSumShareface + dblTmp05
8060                dblTmp04 = 0#: dblTmp05 = 0#
8070                If lngX < lngRecs Then .MoveNext
8080              Next
8090            End With
8100          End If  ' ** blnSkip.
              ' ################################################################################

8110          DoEvents
8120          Set dbs = CurrentDb

              ' ** Step 7: Sum holdings.
8130          .Map_Status_lbl.Caption = " 7 of 22: Sum holdings . . ."
8140          DoEvents

8150          Select Case intInputType
              Case 1  ' ** User entered an Amount.
8160            Select Case .opgAccountNumber
                Case .opgAccountNumber_optAll.OptionValue
                  ' ** qryMap_Rec_13a (Journal Map, with qryMap_Rec_07_01 (ActiveAssets,
                  ' ** grouped and summed, by assetno, by specified [astno]), with calc'd
                  ' ** fields, by specified [astno], [amt], [prec]), grouped and summed by account.
8170              Set qdf = dbs.QueryDefs("qryMap_Rec_14_05")
8180              With qdf.Parameters
8190                ![astno] = lngAssetNo
8200                ![amt] = dblAmount
8210                ![prec] = gintShareFaceDecimals
8220              End With
8230            Case .opgAccountNumber_optSpecified.OptionValue
                  ' ** qryMap_Rec_14cc (qryMap_Rec_13c (Journal Map, with qryMap_Rec_07_02 (ActiveAssets,
                  ' ** grouped and summed, by assetno, by specified [astno], [actno]), with calc'd fields,
                  ' ** by specified [astno], [amt], [prec], [actno]), grouped and summed by account),
                  ' ** with .._rnd rounded again.
8240              Set qdf = dbs.QueryDefs("qryMap_Rec_14_07")
8250              With qdf.Parameters
8260                ![astno] = lngAssetNo
8270                ![actno] = gstrAccountNo
8280                ![amt] = dblAmount
8290                ![prec] = gintShareFaceDecimals
8300              End With
8310            End Select
8320          Case 2  ' ** User entered a PerShare.
8330            Select Case .opgAccountNumber
                Case .opgAccountNumber_optAll.OptionValue
                  ' ** qryMap_Rec_13b (Journal Map, with qryMap_Rec_07_01 (ActiveAssets,
                  ' ** grouped and summed, by assetno, by specified [astno]), with calc'd
                  ' ** fields, by specified [astno], [persh], [prec]), grouped and summed by account.
8340              Set qdf = dbs.QueryDefs("qryMap_Rec_14_06")
8350              With qdf.Parameters
8360                ![astno] = lngAssetNo
8370                ![persh] = dblPerShare
8380                ![prec] = gintShareFaceDecimals
8390              End With
8400            Case .opgAccountNumber_optSpecified.OptionValue
                  ' ** qryMap_Rec_14dd (qryMap_Rec_13d (Journal Map, with qryMap_Rec_07_02 (ActiveAssets,
                  ' ** grouped and summed, by assetno, by specified [astno], [actno]), with calc'd fields,
                  ' ** by specified [astno], [persh], [prec], [actno]), grouped and summed by account),
                  ' ** with .._rnd rounded again.
8410              Set qdf = dbs.QueryDefs("qryMap_Rec_14_08")
8420              With qdf.Parameters
8430                ![astno] = lngAssetNo
8440                ![actno] = gstrAccountNo
8450                ![persh] = dblPerShare
8460                ![prec] = gintShareFaceDecimals
8470              End With
8480            End Select
8490          End Select
8500          Set rst = qdf.OpenRecordset
8510          With rst
8520            If .BOF = True And .EOF = True Then
                  ' ** Something's seriously wrong!
8530              blnContinue = False
8540            Else
8550              .MoveLast
8560              lngMaps = .RecordCount
8570              .MoveFirst
8580              arr_varMap = .GetRows(lngMaps)
                  ' ***************************************************
                  ' ** Array: arr_varMap()
                  ' **
                  ' **   Field  Element  Name             Constant
                  ' **   =====  =======  ===============  ===========
                  ' **     1       0     accountno        M_ACTNO
                  ' **     2       1     assetno          M_ASTNO
                  ' **     3       2     shareface        M_SHARES       ' ** Total shareface by account.
                  ' **     4       3     shareface_rnd    M_SHARESR      ' ** Total shareface by account, rounded to gintShareFaceDecimals.
                  ' **     5       4     pershare         M_PER
                  ' **     6       5     pcash            M_ICASH        ' ** Total pcash by account: ([shareface]*([amt]/[sumsf])).
                  ' **     7       6     pcash_rnd        M_ICASHR       ' ** Total pcash by account, rounded to 2 decimal places.
                  ' **
                  ' ***************************************************
8590            End If
8600            .Close
8610          End With
8620          Set rst = Nothing
8630          Set qdf = Nothing
8640          DoEvents

              ' ** Step 8: Rounding Stage 1.
8650          .Map_Status_lbl.Caption = " 8 of 22: Rounding Stage 1 . . ."
8660          DoEvents

8670          Set rst = dbs.OpenRecordset("journal map", dbOpenDynaset, dbConsistent)
8680          With rst
8690            If .BOF = True And .EOF = True Then
                  ' ** Might be coming from GoToReport.
8700              blnContinue = False
8710            Else
8720              .MoveLast
8730              lngRecs = .RecordCount
8740              .MoveFirst
8750              For lngX = 1& To lngRecs
8760                If InStr(![accountno], strAccountNo_OU) = 0 Then  ' ** This is excluded from the arr_varMap() query.
8770                  blnFound = False
8780                  For lngY = 0& To (lngMaps - 1&)
8790                    If arr_varMap(M_ACTNO, lngY) = ![accountno] And arr_varMap(M_ASTNO, lngY) = ![assetno] Then
8800                      blnFound = True
8810                      .Edit
8820                      dblTmp04 = Round(![shareface], gintShareFaceDecimals)
8830                      ![shareface] = dblTmp04
8840                      ![ICash] = CCur(Round((dblTmp04 * arr_varMap(M_PER, lngY)), 2))
8850                      ![pershare] = arr_varMap(M_PER, lngY)
8860                      .Update
8870                      Exit For
8880                    End If
8890                  Next
8900                  If blnFound = False Then
8910                    blnContinue = False
8920                    Exit For
8930                  End If
8940                End If
8950                DoEvents
8960                If lngX < lngRecs Then .MoveNext
8970              Next
8980            End If  ' ** BOF, EOF.
8990            .Close
9000          End With
9010          Set rst = Nothing
9020          DoEvents

9030          If blnContinue = True Then

                ' ** Step 9: Delete under-threshold entries.
9040            .Map_Status_lbl.Caption = " 9 of 22: Delete under-threshold entries . . ."
9050            DoEvents

                ' ** Delete qryMap_Rec_11 (Journal Map, for shareface < 0.00006, pcash < 0.006, by specified [astno]).
9060            Set qdf = dbs.QueryDefs("qryMap_Rec_12")
9070            With qdf.Parameters
9080              ![astno] = lngAssetNo
9090            End With
9100            qdf.Execute
9110            Set qdf = Nothing
9120            DoEvents

                ' ** Step 10: Sum distributions.
9130            .Map_Status_lbl.Caption = "10 of 22: Sum distributions . . ."
9140            DoEvents

9150            Select Case strCallingForm
                Case "frmJournal"
                  ' ** qryMap_Rec_02_06 (Journal Map, linked to Account, qryMap_Rec_02_02 (qryMap_Rec_02_01
                  ' ** (Journal Map, grouped by accountno, assetno, Location_ID), grouped by accountno,
                  ' ** assetno, with cnt, Location_ID_min, Location_ID_max), grouped and summed,
                  ' ** by accountno), grouped and summed, by assetno, by specified [astno].
9160              Set qdf = dbs.QueryDefs("qryMap_Rec_02_07")
9170            Case "frmJournal_Columns"
                  ' ** qryJournal_Columns_35_05_06 (xx), grouped and summed, by assetno, by specified [astno].
9180              Set qdf = dbs.QueryDefs("qryJournal_Columns_35_05_07")
9190            End Select
9200            With qdf.Parameters
9210              ![astno] = lngAssetNo
9220            End With
9230            Set rst = qdf.OpenRecordset
9240            With rst
9250              .MoveFirst
9260              dblSumPCash = ![PCash]
9270              dblSumShareface = ![shareface]
9280              dblPerShare = ![pershare]  ' ** Max([pershare]), though they should all be identical.
9290              .Close
9300            End With
9310            Set rst = Nothing
9320            Set qdf = Nothing
9330            DoEvents

9340            dblTmp04 = (Round(dblAmount, 2) - dblSumPCash)
9350            If ((dblTmp04 > 0.009) Or (dblTmp04 < -0.009)) And (dblTmp04 > 1#) Then
                  ' ** There's an Over/Under of more than $1.00, so let's see if we can bring it down.

                  ' ** Step 11: Rounding Stage 2.
9360              .Map_Status_lbl.Caption = "11 of 22: Rounding Stage 2 . . ."
9370              DoEvents

                  '.opgAccountNumber_optAll.OptionValue:
                  '  .._07_01:
                  '    [astno]
                  '  .._13_01
                  '    [astno]
                  '    [amt]
                  '    [prec]
                  '  .._13_02:
                  '    [astno]
                  '    [persh]
                  '    [prec]
                  '  .._16_01, .._16_02:
                  '    [thrsh]
                  '  .._17_01, .._17_02:
                  '    [astno]
                  '    [pcashou]
                  '=====================
                  '    [astno]
                  '    [amt]
                  '    [prec]
                  '    [thrsh]
                  '    [pcashou]
                  '--------------
                  '    [astno]
                  '    [persh]
                  '    [prec]
                  '    [thrsh]
                  '    [pcashou]

                  '.opgAccountNumber_optSpecified.OptionValue:
                  '  .._07_02:
                  '    [astno]
                  '    [actno]
                  '  .._13_03:
                  '    [astno]
                  '    [amt]
                  '    [prec]
                  '  .._13_04:
                  '    [astno]
                  '    [persh]
                  '    [prec]
                  '  .._16_03, .._16_04:
                  '    [thrsh]
                  '  .._17_03, .._17_04:
                  '    [astno]
                  '    [pcashou]
                  '=====================
                  '    [astno]
                  '    [actno]
                  '    [amt]
                  '    [prec]
                  '    [thrsh]
                  '    [pcashou]
                  '--------------
                  '    [astno]
                  '    [actno]
                  '    [persh]
                  '    [prec]
                  '    [thrsh]
                  '    [pcashou]

9380              Select Case intInputType
                  Case 1  ' ** User entered an Amount.
9390                Select Case .opgAccountNumber
                    Case .opgAccountNumber_optAll.OptionValue
                      ' ** qryMap_Rec_18a (Journal Map, linked to qryMap_Rec_17a (Journal Map, linked to
                      ' ** qryMap_Rec_16a (qryMap_Rec_14a (qryMap_Rec_13a (Journal Map, with qryMap_Rec_07_01
                      ' ** (ActiveAssets, grouped and summed, by assetno, by specified [astno]), with calc'd
                      ' ** fields, by specified [astno], [amt], [prec]), grouped and summed by account),
                      ' ** just needed fields, by specified [thrsh]), with pcash_adj, by specified [astno],
                      ' ** [thrsh], [pcashou]; Cartesian), with pcash_adjx), just entries with adjustment.
9400                  Set qdf = dbs.QueryDefs("qryMap_Rec_19_01")
9410                  With qdf.Parameters
9420                    ![astno] = lngAssetNo
9430                    ![amt] = dblAmount
9440                    ![prec] = gintShareFaceDecimals
9450                    dblTmp05 = (0.01 / (dblTmp04 / dblShareface))  ' ** Threshhold needed for a penny adjustment.
9460                    ![thrsh] = dblTmp05
9470                    ![pcashou] = dblTmp04
9480                  End With
                      '    [astno]
                      '    [amt]
                      '    [prec]
                      '    [thrsh]
                      '    [pcashou]
9490                Case .opgAccountNumber_optSpecified.OptionValue
                      ' ** qryMap_Rec_18c (Journal Map, linked to qryMap_Rec_17c (Journal Map, linked to
                      ' ** qryMap_Rec_16c (qryMap_Rec_14c (qryMap_Rec_14cc (qryMap_Rec_13c (Journal Map,
                      ' ** with qryMap_Rec_07_02 (ActiveAssets, grouped and summed, by assetno, by specified
                      ' ** [astno], [actno]), with calc'd fields, by specified [astno], [amt], [prec], [actno]),
                      ' ** grouped and summed by account), with .._rnd rounded again), just needed fields, by
                      ' ** specified [thrsh]), with qryMap_Rec_07_02 (ActiveAssets, grouped and summed, by assetno,
                      ' ** by specified [astno], [actno]), with icash_adj, by specified [astno], [thrsh], [pcashou],
                      ' ** [actno]; Cartesian), with pcash_adjx), just entries with adjustment.
9500                  Set qdf = dbs.QueryDefs("qryMap_Rec_19_03")
9510                  With qdf.Parameters
9520                    ![astno] = lngAssetNo
9530                    ![actno] = gstrAccountNo
9540                    ![amt] = dblAmount
9550                    ![prec] = gintShareFaceDecimals
9560                    dblTmp05 = (0.01 / (dblTmp04 / dblShareface))  ' ** Threshhold needed for a penny adjustment.
9570                    ![thrsh] = dblTmp05
9580                    ![pcashou] = dblTmp04
9590                  End With
                      '    [astno]
                      '    [actno]
                      '    [amt]
                      '    [prec]
                      '    [thrsh]
                      '    [pcashou]
9600                End Select
9610              Case 2  ' ** User entered a PerShare.
9620                Select Case .opgAccountNumber
                    Case .opgAccountNumber_optAll.OptionValue
                      ' ** qryMap_Rec_18b (Journal Map, linked to qryMap_Rec_17b (Journal Map, linked to
                      ' ** qryMap_Rec_16b (qryMap_Rec_14b (qryMap_Rec_13b (Journal Map, with qryMap_Rec_07_01
                      ' ** (ActiveAssets, grouped and summed, by assetno, by specified [astno]), with calc'd
                      ' ** fields, by specified [astno], [persh], [prec]), grouped and summed by account),
                      ' ** just needed fields, by specified [thrsh]), with pcash_adj, by specified [astno],
                      ' ** [thrsh], [pcashou]; Cartesian), with pcash_adjx), just entries with adjustment.
9630                  Set qdf = dbs.QueryDefs("qryMap_Rec_19_02")
9640                  With qdf.Parameters
9650                    ![astno] = lngAssetNo
9660                    ![persh] = dblPerShare
9670                    ![prec] = gintShareFaceDecimals
9680                    dblTmp05 = (0.01 / (dblTmp04 / dblShareface))  ' ** Threshhold needed for a penny adjustment.
9690                    ![thrsh] = dblTmp05
9700                    ![pcashou] = dblTmp04
9710                  End With
                      '    [astno]
                      '    [persh]
                      '    [prec]
                      '    [thrsh]
                      '    [pcashou]
9720                Case .opgAccountNumber_optSpecified.OptionValue
                      ' ** qryMap_Rec_18d (Journal Map, linked to qryMap_Rec_17d (Journal Map, linked to
                      ' ** qryMap_Rec_16d (qryMap_Rec_14d (qryMap_Rec_14dd (qryMap_Rec_13d (Journal Map,
                      ' ** with qryMap_Rec_07_02 (ActiveAssets, grouped and summed, by assetno, by specified
                      ' ** [astno], [actno]), with calc'd fields, by specified [astno], [persh], [prec], [actno]),
                      ' ** grouped and summed by account), with .._rnd rounded again), just needed fields, by
                      ' ** specified [thrsh]), with qryMap_Rec_07_02 (ActiveAssets, grouped and summed, by assetno,
                      ' ** by specified [astno], [actno]), with icash_adj, by specified [astno], [thrsh], [pcashou],
                      ' ** [actno]; Cartesian), with pcash_adjx), just entries with adjustment.
9730                  Set qdf = dbs.QueryDefs("qryMap_Rec_19_04")
9740                  With qdf.Parameters
9750                    ![astno] = lngAssetNo
9760                    ![actno] = gstrAccountNo
9770                    ![persh] = dblPerShare
9780                    ![prec] = gintShareFaceDecimals
9790                    dblTmp05 = (0.01 / (dblTmp04 / dblShareface))  ' ** Threshhold needed for a penny adjustment.
9800                    ![thrsh] = dblTmp05
9810                    ![pcashou] = dblTmp04
9820                  End With
                      '    [astno]
                      '    [actno]
                      '    [persh]
                      '    [prec]
                      '    [thrsh]
                      '    [pcashou]
9830                End Select
9840              End Select
9850              DoEvents

                  ' ** Step 12: Collect adjustments.
9860              .Map_Status_lbl.Caption = "12 of 22: Collect adjustments . . ."
9870              DoEvents

9880              lngAdjs = 0&
9890              Set rst = qdf.OpenRecordset
9900              With rst
9910                If .BOF = True And .EOF = True Then
                      ' ** Not enough to adjust.
9920                Else
9930                  .MoveLast
9940                  lngAdjs = .RecordCount
9950                  .MoveFirst
9960                  arr_varAdj = .GetRows(lngAdjs)
                      ' **********************************************
                      ' ** Array: arr_varAdj()
                      ' **
                      ' **   Field  Element  Name         Constant
                      ' **   =====  =======  ===========  ==========
                      ' **     1       0     ID           A_ID
                      ' **     2       1     accountno    A_ACTNO
                      ' **     3       2     assetno      A_ASTNO
                      ' **     4       3     pcash        A_PCASH
                      ' **     5       4     pcash_adj    A_PCADJ
                      ' **
                      ' **********************************************
9970                End If
9980                .Close
9990              End With
10000             Set rst = Nothing
10010             Set qdf = Nothing
10020             DoEvents  ' ** All these DoEvents are to keep the window from showing a 'Not Responding'!

                  ' ** Step 13: Update Journal Map.
10030             .Map_Status_lbl.Caption = "13 of 22: Update temporary table . . ."
10040             DoEvents

10050             Set rst = dbs.OpenRecordset("Journal Map", dbOpenDynaset, dbConsistent)
10060             With rst
10070               .MoveFirst
10080               For lngX = 0& To (lngAdjs - 1&)
10090                 .FindFirst "[ID] = " & CStr(arr_varAdj(A_ID, lngX))
10100                 If .NoMatch = False Then
10110                   .Edit
10120                   ![PCash] = arr_varAdj(A_PCADJ, lngX)
10130                   .Update
10140                 Else
10150                   blnContinue = False
10160                   Exit For
10170                 End If
10180               Next
10190             End With
10200             Set rst = Nothing
10210             Set qdf = Nothing
10220             DoEvents

                  ' ** Step 14: Confirm Rounding Stage 2.
10230             .Map_Status_lbl.Caption = "14 of 22: Confirm Rounding Stage 2 . . ."
10240             DoEvents

10250             Select Case strCallingForm
                  Case "frmJournal"
                    ' ** Now check the Over/Under again!
                    ' ** qryMap_Rec_02_06 (Journal Map, linked to Account, qryMap_Rec_02_02 (qryMap_Rec_02_01
                    ' ** (Journal Map, grouped by accountno, assetno, Location_ID), grouped by accountno,
                    ' ** assetno, with cnt, Location_ID_min, Location_ID_max), grouped and summed,
                    ' ** by accountno), grouped and summed, by assetno, by specified [astno].
10260               Set qdf = dbs.QueryDefs("qryMap_Rec_02_07")
10270             Case "frmJournal_Columns"
                    ' ** qryJournal_Columns_35_05_06 (Journal Map, linked to Account, qryJournal_Columns_35_05_02
                    ' ** (qryJournal_Columns_35_05_01 (Journal Map, grouped by accountno, assetno, Location_ID),
                    ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max), grouped
                    ' ** and summed, by accountno), grouped and summed, by assetno, by specified [astno].
10280               Set qdf = dbs.QueryDefs("qryJournal_Columns_35_05_07")
10290             End Select
10300             With qdf.Parameters
10310               ![astno] = lngAssetNo
10320             End With
10330             Set rst = qdf.OpenRecordset
10340             With rst
10350               .MoveFirst
10360               dblSumPCash = ![PCash]
10370               dblSumShareface = ![shareface]
10380               dblPerShare = ![pershare]  ' ** Max([pershare]), though they should all be identical.
10390               .Close
10400             End With
10410             Set rst = Nothing
10420             Set qdf = Nothing
10430             DoEvents

10440             dblTmp04 = (Round(dblAmount, 2) - dblSumPCash)

10450           End If

                ' ** Step 15: Create Over/Under entry.
10460           .Map_Status_lbl.Caption = "15 of 22: Create Over/Under entry . . ."
10470           DoEvents

10480           If ((dblTmp04 > 0.009) Or (dblTmp04 < -0.009)) Then
                  ' ** Create a new record in Journal Map.
10490             Set rst = dbs.OpenRecordset("journal map", dbOpenDynaset, dbAppendOnly)
10500             With rst
10510               .AddNew
10520               ![assetno] = lngAssetNo
10530               ![accountno] = strAccountNo_OU
10540               ![shareface] = Round(Abs(dblShareface - dblSumShareface), gintShareFaceDecimals)
10550               ![pershare] = dblPerShare
10560               ![transdate] = datTransDate
10570               ![journaltype] = strMapType
10580               dblOverUnder = Round((Round(dblAmount, 2) - dblSumPCash), 2)  ' ** Module Function: modUtilities.
10590               ![PCash] = CCur(Round(dblOverUnder, 2))
10600               ![description] = "Over/Short of " & strMapType & " Mapping for " & Me.cmbAssets.Column(CBX_A_DESC)
10610               ![curr_id] = lngCurrID
10620               .Update
10630               .Close
10640             End With
10650             Set rst = Nothing
10660             .overshort = dblOverUnder
10670           Else
10680             .overshort = 0#
10690           End If

10700           dbs.Close
10710           Set dbs = Nothing
10720           DoEvents

                ' ** Step 16: Delete zero entries.
10730           .Map_Status_lbl.Caption = "16 of 22: Delete zero entries . . ."
10740           DoEvents

                ' ################################################################################
10750           If blnSkip = False Then
10760             Set dbs = CurrentDb
                  ' ** Journal Map, just zero shareface, icash, pcash.
10770             Set qdf = dbs.QueryDefs("qryMap_Div_01a")
10780             Set rst = qdf.OpenRecordset
10790             If rst.BOF = True And rst.EOF = True Then
                    ' ** All's well.
10800               rst.Close
10810             Else
10820               rst.Close
                    ' ** Delete qryMap_Div_01a (Journal Map, just zero shareface, icash, pcash).
10830               Set qdf = dbs.QueryDefs("qryMap_Div_01b")
10840               qdf.Execute
10850             End If
10860             dbs.Close
10870           End If  ' ** blnSkip.
                ' ################################################################################

10880           dblOverUnder = AdvancedRounding(lngAssetNo)  ' ** Procedure: Below.

                ' ** Step 22: Display detail.
10890           .Map_Status_lbl.Caption = "22 of 22: Display detail . . ."
10900           DoEvents

                ' ** Borrowing these variables from the Court Reports.
10910           gstrCrtRpt_Ordinal = .cmbAssets.Column(CBX_A_CUSIP)  ' ** cusip.
10920           gstrCrtRpt_Version = .cmbAssets.Column(CBX_A_TDESC)  ' ** totdesc.

10930           strDocName = "frmMap_Rec_Detail"
10940           DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & CStr(lngAssetNo) & "~" & _
                  CStr(dblOverUnder) & "~" & CStr(dblPerShare) & "~" & Format(datTransDate, "mm/dd/yyyy") & "~" & _
                  Format(datAssetDate, "mm/dd/yyyy hh:nn:ss") & "~" & CStr(lngCurrID)

10950           If gblnGoToReport = True Then
10960             DoCmd.Hourglass True  ' ** Make sure it's still running.
10970             DoEvents
10980             .GoToReport_arw_ok_img.Visible = False
10990             Select Case IsNull(garr_varGoToReport(GTR_FRM4))
                  Case True
                    ' ** Finished at frmMap_Rec_Detail.
11000               Forms(strDocName).cmdCancel.SetFocus
11010               Forms(strDocName).cmdOK.Enabled = False  ' ** Only allow Cancel on this fake dividend.
11020               Forms(strDocName).cmdReinvest.Enabled = False
11030               DoEvents
11040               Forms(strDocName).TimerInterval = 50&
11050               blnGoingToReport = False
11060             Case False
                    ' ** Continuing on to frmMap_Reinvest_Rec_Price.
11070               Forms(strDocName).cmdReinvest.SetFocus
11080               Forms(strDocName).cmdOK.Enabled = False  ' ** Only allow Reinvest on this fake dividend.
11090               Forms(strDocName).cmdCancel.Enabled = False
11100               DoEvents
11110               Forms(strDocName).TimerInterval = 100&
11120               blnGoingToReport = False
11130             End Select
11140           End If

11150           DoCmd.Close acForm, THIS_NAME

11160         Else
11170           Select Case gblnGoToReport
                Case True
11180             Beep
11190             DoCmd.Hourglass False
11200             MsgBox "The report site cannot be shown because there are no records.", _
                    vbInformation + vbOKOnly, "Record Not Found"
11210             blnGoingToReport = False
11220             .GoToReport_arw_ok_img.Visible = False
11230             gblnGoToReport = False
11240           Case False
11250             Beep
11260             DoCmd.Hourglass False
11270             MsgBox "This procedure cannot continue." & vbCrLf & "Expected records were not found.", _
                    vbInformation + vbOKOnly, "Record Not Found"
11280             blnGoingToReport = False
11290             .GoToReport_arw_ok_img.Visible = False
11300             gblnGoToReport = False
11310           End Select
11320           .cmdCancel.SetFocus
11330         End If  ' ** blnContinue.

11340       Else
11350         .cmdCancel.SetFocus
11360       End If  ' ** msgResponse.

11370     End If

11380   End With

EXITP:
11390   Set rsx1 = Nothing
11400   Set cnxn = Nothing
11410   Set rst = Nothing
11420   Set qdf = Nothing
11430   Set dbs = Nothing
11440   Exit Sub

ERRH:
11450   DoCmd.Hourglass False
11460   THAT_PROC = THIS_PROC
11470   That_Erl = Erl
11480   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11490   Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
11590         .cmdCancel.SetFocus
11600       End With
11610     End Select
11620   End If

        ' ** Shift keys.
11630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11640     Select Case intRetVal
          Case vbKeyTab
11650       With Me
11660         intRetVal = 0
11670         .transdate.SetFocus
11680       End With
11690     End Select
11700   End If

        ' ** Ctrl-Shift keys.
11710   If intCtrlDown And (Not intAltDown) And intShiftDown Then
11720     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11730       With Me
11740         intRetVal = 0
11750         .cmdCalendar1.SetFocus
11760       End With
11770     End Select
11780   End If

EXITP:
11790   KeyCode = intRetVal
11800   Exit Sub

ERRH:
11810   intRetVal = 0
11820   THAT_PROC = THIS_PROC
11830   That_Erl = Erl
11840   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11850   Resume EXITP

End Sub

Private Sub cmdCancel_Click()

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

11910   DoCmd.Hourglass True
11920   DoEvents

11930   blnCancel = True

11940   Set dbs = CurrentDb
11950   With dbs
11960     Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
11970       Set qdf = .QueryDefs("qryMap_01")
11980     Case "frmJournal_Columns"
            ' ** Empty Journal Map.
11990       Set qdf = .QueryDefs("qryJournal_Columns_35_04")
12000     End Select
12010     qdf.Execute
12020     Set qdf = Nothing
12030     DoEvents
          ' ** Empty tmpAccount.
12040     Set qdf = dbs.QueryDefs("qryMap_Rec_08_01")
12050     qdf.Execute
12060     Set qdf = Nothing
12070     DoEvents
12080     .Close
12090   End With

12100   gblnSetFocus = True
12110   Select Case strCallingForm
        Case "frmJournal"
12120     If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
12130       DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
12140     Else
12150       gstrReturningForm = THIS_NAME
12160       DoCmd.SelectObject acForm, strCallingForm, False
12170       Forms(strCallingForm).TimerInterval = 100&
12180     End If
12190   Case "frmJournal_Columns"
12200     If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
12210       DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
12220     Else
12230       gstrReturningForm = THIS_NAME
12240       DoCmd.SelectObject acForm, strCallingForm, False
12250       Forms(strCallingForm).TimerInterval = 250&
12260     End If
12270   End Select

12280   DoCmd.Close acForm, THIS_NAME

EXITP:
12290   Set qdf = Nothing
12300   Set dbs = Nothing
12310   Exit Sub

ERRH:
12320   DoCmd.Hourglass False
12330   THAT_PROC = THIS_PROC
12340   That_Erl = Erl
12350   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12360   Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12420   intShiftDown = (Shift And acShiftMask) > 0
12430   intAltDown = (Shift And acAltMask) > 0
12440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12460     Select Case intRetVal
          Case vbKeyTab
12470       With Me
12480         intRetVal = 0
12490         .opgAccountNumber.SetFocus
12500       End With
12510     End Select
12520   End If

        ' ** Shift keys.
12530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12540     Select Case intRetVal
          Case vbKeyTab
12550       With Me
12560         intRetVal = 0
12570         .cmdOK.SetFocus
12580       End With
12590     End Select
12600   End If

EXITP:
12610   KeyCode = intRetVal
12620   Exit Sub

ERRH:
12630   intRetVal = 0
12640   THAT_PROC = THIS_PROC
12650   That_Erl = Erl
12660   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12670   Resume EXITP

End Sub

Private Sub opgAccountNumber_AfterUpdate()

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_AfterUpdate"

12710   With Me
12720     .cmbAssets = Null
12730     .cmbAccounts = Null
12740     .shareface = 0#
12750     .amount = Null
12760     .pershare = Null
12770     .assetdate = Null
12780     .assetdate_display = Null
12790     .Amount_usd = Null
12800     .pershare_usd = Null
12810     Select Case .opgAccountNumber
          Case .opgAccountNumber_optAll.OptionValue
12820       .opgAccountNumber_optAll_lbl.FontBold = True
12830       .opgAccountNumber_optSpecified_lbl.FontBold = False
12840       .opgAccountNumber_optAll_lbl_box.Visible = True
12850       .opgAccountNumber_optSpecified_lbl_box.Visible = False
12860       .cmbAccounts.Enabled = False
12870       .cmbAccounts.BorderColor = WIN_CLR_DISR
12880       .cmbAccounts.BackStyle = acBackStyleTransparent
12890       .cmbAccounts_lbl_box.Visible = True
12900       .opgAccountSource.Enabled = False
12910       .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
12920       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
12930       .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
12940       .opgAccountSource_optName_lbl2_dim_hi.Visible = True
12950       .chkRememberMe.Enabled = False
12960       .chkRememberMe_lbl.Visible = False
12970       .chkRememberMe_lbl2_dim.Visible = True
12980       .chkRememberMe_lbl2_dim_hi.Visible = True
12990       DoEvents
            ' ** cmbAssets RowSource's:
            ' **   qryMap_Rec_05a  .opgAssetSource_optType.OptionValue
            ' **   qryMap_Rec_05c    GlobalVarGet('gstrAccountNo')
            ' **   qryMap_Rec_05d  .opgAssetSource_optName.OptionValue
            ' **   qryMap_Rec_05f    GlobalVarGet('gstrAccountNo')
            ' **   qryMap_Rec_05g  .opgAssetSource_optCusip.OptionValue
            ' **   qryMap_Rec_05i    GlobalVarGet('gstrAccountNo')
13000       Select Case .opgAssetSource
            Case .opgAssetSource_optType.OptionValue
13010         If .cmbAssets.RowSource <> "qryMap_Rec_05a" Then
13020           .cmbAssets.RowSource = "qryMap_Rec_05a"
13030         End If
13040       Case .opgAssetSource_optName.OptionValue
13050         If .cmbAssets.RowSource <> "qryMap_Rec_05d" Then
13060           .cmbAssets.RowSource = "qryMap_Rec_05d"
13070         End If
13080       Case .opgAssetSource_optCusip.OptionValue
13090         If .cmbAssets.RowSource <> "qryMap_Rec_05g" Then
13100           .cmbAssets.RowSource = "qryMap_Rec_05g"
13110         End If
13120       End Select
13130       .cmbAssets.Requery
13140       DoEvents
13150       .cmbAssets.SetFocus
13160     Case .opgAccountNumber_optSpecified.OptionValue
            ' ** cmbAssets can remain as-is.
13170       .opgAccountNumber_optAll_lbl.FontBold = False
13180       .opgAccountNumber_optSpecified_lbl.FontBold = True
13190       .opgAccountNumber_optAll_lbl_box.Visible = False
13200       .opgAccountNumber_optSpecified_lbl_box.Visible = True
13210       .cmbAccounts.Enabled = True
13220       .cmbAccounts.BorderColor = CLR_LTBLU2
13230       .cmbAccounts.BackStyle = acBackStyleNormal
13240       .cmbAccounts_lbl_box.Visible = False
13250       .opgAccountSource.Enabled = True
13260       .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
13270       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
13280       .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
13290       .opgAccountSource_optName_lbl2_dim_hi.Visible = False
13300       .chkRememberMe.Enabled = True
13310       .chkRememberMe_lbl.Visible = True
13320       .chkRememberMe_lbl2_dim.Visible = False
13330       .chkRememberMe_lbl2_dim_hi.Visible = False
13340       .cmbAccounts.Requery
13350       DoEvents
13360       .cmbAccounts.SetFocus
13370     End Select
13380   End With

EXITP:
13390   Exit Sub

ERRH:
13400   Select Case ERR.Number
        Case Else
13410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13420   End Select
13430   Resume EXITP

End Sub

Private Sub opgAccountNumber_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optAll_KeyDown"

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
13590         Select Case .opgAccountNumber
              Case .opgAccountNumber_optAll.OptionValue
13600           .cmbAssets.SetFocus
13610         Case .opgAccountNumber_optSpecified.OptionValue
13620           .cmbAccounts.SetFocus
13630         End Select
13640       End With
13650     End Select
13660   End If

        ' ** Shift keys.
13670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13680     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13690       With Me
13700         intRetVal = 0
13710         .cmdCancel.SetFocus
13720       End With
13730     End Select
13740   End If

EXITP:
13750   KeyCode = intRetVal
13760   Exit Sub

ERRH:
13770   intRetVal = 0
13780   THAT_PROC = THIS_PROC
13790   That_Erl = Erl
13800   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13810   Resume EXITP

End Sub

Private Sub opgAccountNumber_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optSpecified_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13920   intShiftDown = (Shift And acShiftMask) > 0
13930   intAltDown = (Shift And acAltMask) > 0
13940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13970       With Me
13980         intRetVal = 0
13990         Select Case .opgAccountNumber
              Case .opgAccountNumber_optAll.OptionValue
14000           .cmbAssets.SetFocus
14010         Case .opgAccountNumber_optSpecified.OptionValue
14020           .cmbAccounts.SetFocus
14030         End Select
14040       End With
14050     End Select
14060   End If

        ' ** Shift keys.
14070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14080     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14090       With Me
14100         intRetVal = 0
14110         .cmdCancel.SetFocus
14120       End With
14130     End Select
14140   End If

EXITP:
14150   KeyCode = intRetVal
14160   Exit Sub

ERRH:
14170   intRetVal = 0
14180   THAT_PROC = THIS_PROC
14190   That_Erl = Erl
14200   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14210   Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

14310   With Me
14320     Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
14330       .opgAccountSource_optNumber_lbl2.FontBold = True
14340       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = True
14350       .opgAccountSource_optName_lbl2.FontBold = False
14360       .opgAccountSource_optName_lbl2_dim_hi.FontBold = False
14370       If .cmbAccounts.RowSource <> "qryMap_Rec_08a" Then
14380         .cmbAccounts.RowSource = "qryMap_Rec_08a"
14390       End If
14400     Case .opgAccountSource_optName.OptionValue
14410       .opgAccountSource_optNumber_lbl2.FontBold = False
14420       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = False
14430       .opgAccountSource_optName_lbl2.FontBold = True
14440       .opgAccountSource_optName_lbl2_dim_hi.FontBold = True
14450       If .cmbAccounts.RowSource <> "qryMap_Rec_08b" Then
14460         .cmbAccounts.RowSource = "qryMap_Rec_08b"
14470       End If
14480     End Select
14490     .cmbAccounts.Requery
14500   End With

EXITP:
14510   Exit Sub

ERRH:
14520   Select Case ERR.Number
        Case Else
14530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14540   End Select
14550   Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14620   intShiftDown = (Shift And acShiftMask) > 0
14630   intAltDown = (Shift And acAltMask) > 0
14640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14670       With Me
14680         intRetVal = 0
14690         .chkRememberMe.SetFocus
14700       End With
14710     End Select
14720   End If

        ' ** Shift keys.
14730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14750       With Me
14760         intRetVal = 0
14770         .opgAccountNumber.SetFocus
14780       End With
14790     End Select
14800   End If

EXITP:
14810   KeyCode = intRetVal
14820   Exit Sub

ERRH:
14830   intRetVal = 0
14840   THAT_PROC = THIS_PROC
14850   That_Erl = Erl
14860   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14870   Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14920   intShiftDown = (Shift And acShiftMask) > 0
14930   intAltDown = (Shift And acAltMask) > 0
14940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14970       With Me
14980         intRetVal = 0
14990         .chkRememberMe.SetFocus
15000       End With
15010     End Select
15020   End If

        ' ** Shift keys.
15030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15050       With Me
15060         intRetVal = 0
15070         .opgAccountNumber.SetFocus
15080       End With
15090     End Select
15100   End If

EXITP:
15110   KeyCode = intRetVal
15120   Exit Sub

ERRH:
15130   intRetVal = 0
15140   THAT_PROC = THIS_PROC
15150   That_Erl = Erl
15160   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15170   Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

15210   With Me
15220     Select Case .chkRememberMe
          Case True
15230       .chkRememberMe_lbl.FontBold = True
15240       .chkRememberMe_lbl2_dim.FontBold = True
15250       .chkRememberMe_lbl2_dim_hi.FontBold = True
15260     Case False
15270       .chkRememberMe_lbl.FontBold = False
15280       .chkRememberMe_lbl2_dim.FontBold = False
15290       .chkRememberMe_lbl2_dim_hi.FontBold = False
15300     End Select
15310   End With

EXITP:
15320   Exit Sub

ERRH:
15330   Select Case ERR.Number
        Case Else
15340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15350   End Select
15360   Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15420   intShiftDown = (Shift And acShiftMask) > 0
15430   intAltDown = (Shift And acAltMask) > 0
15440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15470       With Me
15480         intRetVal = 0
15490         .cmbAccounts.SetFocus
15500       End With
15510     End Select
15520   End If

        ' ** Shift keys.
15530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15550       With Me
15560         intRetVal = 0
15570         .opgAccountSource.SetFocus
15580       End With
15590     End Select
15600   End If

EXITP:
15610   KeyCode = intRetVal
15620   Exit Sub

ERRH:
15630   intRetVal = 0
15640   THAT_PROC = THIS_PROC
15650   That_Erl = Erl
15660   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15670   Resume EXITP

End Sub

Private Sub cmbAccounts_AfterUpdate()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

15710   With Me
15720     If IsNull(.cmbAccounts) = False Then
15730       gstrAccountNo = .cmbAccounts.Column(0)
15740       gstrAccountName = .cmbAccounts.Column(2)
15750       Select Case .opgAssetSource
            Case .opgAssetSource_optType.OptionValue
              ' ** qryMap_Rec_05a (qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** for Dividend = True, shareface <> 0; by assettype), linked to qryMap_Rec_05b (ActiveAssets,
              ' ** grouped by accountno, assetno), by specified GlobalVarGet('gstrAccountNo'); by assettype.
15760         .cmbAssets.RowSource = "qryMap_Rec_05c"
15770         .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
15780         .cmbAssets.ListWidth = 7890
15790       Case .opgAssetSource_optName.OptionValue
              ' ** qryMap_Rec_05d (qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** for Dividend = True, shareface <> 0; by totdesc), linked to qryMap_Rec_05e (ActiveAssets,
              ' ** grouped by accountno, assetno), by specified GlobalVarGet('gstrAccountNo'); by totdesc.
15800         .cmbAssets.RowSource = "qryMap_Rec_05f"
15810         .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
15820         .cmbAssets.ListWidth = 7890
15830       Case .opgAssetSource_optCusip.OptionValue
              ' ** qryMap_Rec_05g (qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** with qryMap_Rec_05k (qryMap_Rec_05j (MasterAsset, with cusip_len), grouped, with Max(cusip_len)),
              ' ** for Dividend = True, shareface <> 0; by cusip; Cartesian), linked to qryMap_Rec_05h (ActiveAssets,
              ' ** grouped by accountno, assetno), by specified GlobalVarGet('gstrAccountNo'); by cusip.
15840         .cmbAssets.RowSource = "qryMap_Rec_05i"
15850         .cmbAssets.ColumnWidths = "0;7275;0;0;360;0"
15860         .cmbAssets.ListWidth = 7890
15870       End Select
15880     Else
15890       gstrAccountNo = vbNullString: gstrAccountName = vbNullString
15900       Select Case .opgAssetSource
            Case .opgAssetSource_optType.OptionValue
              ' ** qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** for Dividend = True, shareface <> 0; by assettype.
15910         .cmbAssets.RowSource = "qryMap_Rec_05a"
15920         .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
15930         .cmbAssets.ListWidth = 7890
15940       Case .opgAssetSource_optName.OptionValue
              ' ** qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** for Dividend = True, shareface <> 0; by totdesc.
15950         .cmbAssets.RowSource = "qryMap_Rec_05d"
15960         .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
15970         .cmbAssets.ListWidth = 7890
15980       Case .opgAssetSource_optCusip.OptionValue
              ' ** qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** with qryMap_Rec_05k (qryMap_Rec_05j (MasterAsset, with cusip_len), grouped,
              ' ** with Max(cusip_len)), for Dividend = True, shareface <> 0; by cusip; Cartesian.
15990         .cmbAssets.RowSource = "qryMap_Rec_05g"
16000         .cmbAssets.ColumnWidths = "0;7275;0;0;360;0"
16010         .cmbAssets.ListWidth = 7890
16020       End Select
16030     End If
16040     .cmbAssets.Requery
16050   End With

EXITP:
16060   Exit Sub

ERRH:
16070   Select Case ERR.Number
        Case Else
16080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16090   End Select
16100   Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16220   intShiftDown = (Shift And acShiftMask) > 0
16230   intAltDown = (Shift And acAltMask) > 0
16240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16270       With Me
16280         intRetVal = 0
16290         .cmbAssets.SetFocus
16300       End With
16310     End Select
16320   End If

        ' ** Shift keys.
16330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16350       With Me
16360         intRetVal = 0
16370         .opgAccountNumber.SetFocus
16380       End With
16390     End Select
16400   End If

        ' ** Ctrl keys.
16410   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
16420     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16430       With Me
16440         intRetVal = 0
16450         .opgAccountSource.SetFocus
16460       End With
16470     End Select
16480   End If

EXITP:
16490   KeyCode = intRetVal
16500   Exit Sub

ERRH:
16510   intRetVal = 0
16520   THAT_PROC = THIS_PROC
16530   That_Erl = Erl
16540   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16550   Resume EXITP

End Sub

Private Sub opgAssetSource_AfterUpdate()

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_AfterUpdate"

16610   With Me
16620     Select Case .opgAssetSource
          Case .opgAssetSource_optType.OptionValue
16630       .opgAssetSource_optType_lbl.FontBold = True
16640       .opgAssetSource_optName_lbl.FontBold = False
16650       .opgAssetSource_optCusip_lbl.FontBold = False
16660       Select Case .opgAccountNumber
            Case .opgAccountNumber_optAll.OptionValue
              ' ** qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to
              ' ** MasterAsset, for Dividend = True, shareface <> 0; by assettype.
16670         If .cmbAssets.RowSource <> "qryMap_Rec_05a" Then
16680           .cmbAssets.RowSource = "qryMap_Rec_05a"
16690         End If
16700       Case .opgAccountNumber_optSpecified.OptionValue
              ' ** qryMap_Rec_05a (qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** for Dividend = True, shareface <> 0; by assettype), linked to qryMap_Rec_05b (ActiveAssets,
              ' ** grouped by accountno, assetno), by specified GlobalVarGet('gstrAccountNo'); by assettype.
16710         If .cmbAssets.RowSource <> "qryMap_Rec_05c" Then
16720           .cmbAssets.RowSource = "qryMap_Rec_05c"
16730         End If
16740       End Select
16750       .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
16760       .cmbAssets.ListWidth = 7890
16770     Case .opgAssetSource_optName.OptionValue
16780       .opgAssetSource_optType_lbl.FontBold = False
16790       .opgAssetSource_optName_lbl.FontBold = True
16800       .opgAssetSource_optCusip_lbl.FontBold = False
16810       Select Case .opgAccountNumber
            Case .opgAccountNumber_optAll.OptionValue
              ' ** qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to
              ' ** MasterAsset, for Dividend = True, shareface <> 0; by totdesc.
16820         If .cmbAssets.RowSource <> "qryMap_Rec_05d" Then
16830           .cmbAssets.RowSource = "qryMap_Rec_05d"
16840         End If
16850       Case .opgAccountNumber_optSpecified.OptionValue
              ' ** qryMap_Rec_05d (qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** for Dividend = True, shareface <> 0; by totdesc), linked to qryMap_Rec_05e (ActiveAssets,
              ' ** grouped by accountno, assetno), by specified GlobalVarGet('gstrAccountNo'); by totdesc.
16860         If .cmbAssets.RowSource <> "qryMap_Rec_05f" Then
16870           .cmbAssets.RowSource = "qryMap_Rec_05f"
16880         End If
16890       End Select
16900       .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
16910       .cmbAssets.ListWidth = 7890
16920     Case .opgAssetSource_optCusip.OptionValue
16930       .opgAssetSource_optType_lbl.FontBold = False
16940       .opgAssetSource_optName_lbl.FontBold = False
16950       .opgAssetSource_optCusip_lbl.FontBold = True
16960       Select Case .opgAccountNumber
            Case .opgAccountNumber_optAll.OptionValue
              ' ** qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to
              ' ** MasterAsset, with qryMap_Rec_05k (qryMap_Rec_05j (MasterAsset,
              ' ** with cusip_len), grouped, with Max(cusip_len)), for
              ' ** Dividend = True, shareface <> 0; by cusip; Cartesian.
16970         If .cmbAssets.RowSource <> "qryMap_Rec_05g" Then
16980           .cmbAssets.RowSource = "qryMap_Rec_05g"
16990         End If
17000       Case .opgAccountNumber_optSpecified.OptionValue
              ' ** qryMap_Rec_05g (qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** with qryMap_Rec_05k (qryMap_Rec_05j (MasterAsset, with cusip_len), grouped, with Max(cusip_len)),
              ' ** for Dividend = True, shareface <> 0; by cusip; Cartesian), linked to qryMap_Rec_05h (ActiveAssets,
              ' ** grouped by accountno, assetno), by specified GlobalVarGet('gstrAccountNo'); by cusip.
17010         If .cmbAssets.RowSource <> "qryMap_Rec_05i" Then
17020           .cmbAssets.RowSource = "qryMap_Rec_05i"
17030         End If
17040       End Select
17050       .cmbAssets.ColumnWidths = "0;7275;0;0;360;0"
17060       .cmbAssets.ListWidth = 7890
17070     End Select
17080     .cmbAssets.Requery
17090   End With

EXITP:
17100   Exit Sub

ERRH:
17110   Select Case ERR.Number
        Case Else
17120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17130   End Select
17140   Resume EXITP

End Sub

Private Sub opgAssetSource_optType_KeyDown(KeyCode As Integer, Shift As Integer)

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17220   intShiftDown = (Shift And acShiftMask) > 0
17230   intAltDown = (Shift And acAltMask) > 0
17240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17270       With Me
17280         intRetVal = 0
17290         .cmbAssets.SetFocus
17300       End With
17310     End Select
17320   End If

        ' ** Shift keys.
17330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17350       With Me
17360         intRetVal = 0
17370         .opgAccountNumber.SetFocus
17380       End With
17390     End Select
17400   End If

EXITP:
17410   KeyCode = intRetVal
17420   Exit Sub

ERRH:
17430   intRetVal = 0
17440   Select Case ERR.Number
        Case Else
17450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17460   End Select
17470   Resume EXITP

End Sub

Private Sub opgAssetSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17520   intShiftDown = (Shift And acShiftMask) > 0
17530   intAltDown = (Shift And acAltMask) > 0
17540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17570       With Me
17580         intRetVal = 0
17590         .cmbAssets.SetFocus
17600       End With
17610     End Select
17620   End If

        ' ** Shift keys.
17630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17650       With Me
17660         intRetVal = 0
17670         .opgAccountNumber.SetFocus
17680       End With
17690     End Select
17700   End If

EXITP:
17710   KeyCode = intRetVal
17720   Exit Sub

ERRH:
17730   intRetVal = 0
17740   Select Case ERR.Number
        Case Else
17750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17760   End Select
17770   Resume EXITP

End Sub

Private Sub opgAssetSource_optCusip_KeyDown(KeyCode As Integer, Shift As Integer)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optCusip_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
17870       With Me
17880         intRetVal = 0
17890         .cmbAssets.SetFocus
17900       End With
17910     End Select
17920   End If

        ' ** Shift keys.
17930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17950       With Me
17960         intRetVal = 0
17970         .opgAccountNumber.SetFocus
17980       End With
17990     End Select
18000   End If

EXITP:
18010   KeyCode = intRetVal
18020   Exit Sub

ERRH:
18030   intRetVal = 0
18040   Select Case ERR.Number
        Case Else
18050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18060   End Select
18070   Resume EXITP

End Sub

Private Sub cmbAssets_AfterUpdate()

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_AfterUpdate"

18110   With Me

18120     If blnCurrID = True Then
18130       lngCurrID = 0&
18140       If IsNull(.cmbAssets) = False Then
18150         If IsNull(.cmbAssets.Column(CBX_A_CURRID)) = False Then
18160           lngCurrID = .cmbAssets.Column(CBX_A_CURRID)
18170           .curr_id = lngCurrID
18180         End If
18190       End If
18200     End If

18210     .amount = Null
18220     .amount.Enabled = True
18230     .amount.Locked = False
18240     .amount.ForeColor = CLR_BLK
18250     .amount.BackColor = CLR_WHT
18260     .pershare = Null
18270     .pershare.Enabled = True
18280     .pershare.Locked = False
18290     .pershare.ForeColor = CLR_BLK
18300     .pershare.BackColor = CLR_WHT
18310     blnAssetUpdated = True: blnAssetChecked = False
18320     .amount.SetFocus

18330     curr_id_AfterUpdate  ' ** Procedure: Below.

18340   End With

EXITP:
18350   Exit Sub

ERRH:
18360   THAT_PROC = THIS_PROC
18370   That_Erl = Erl
18380   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18390   Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18420   intShiftDown = (Shift And acShiftMask) > 0
18430   intAltDown = (Shift And acAltMask) > 0
18440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18470       With Me
18480         intRetVal = 0
              ' ** This enabling and disabling is causing all sorts of ...
18490         If .amount.Enabled = True Then
18500 On Error Resume Next
18510           .amount.SetFocus
18520 On Error GoTo ERRH
18530         Else
18540 On Error Resume Next
18550           .pershare.SetFocus
18560 On Error GoTo ERRH
18570         End If
18580       End With
18590     End Select
18600   End If

        ' ** Shift keys.
18610   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18620     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18630       With Me
18640         intRetVal = 0
18650         If .cmbAccounts.Enabled = True Then
18660           .cmbAccounts.SetFocus
18670         Else
18680           .opgAccountNumber.SetFocus
18690         End If
18700       End With
18710     End Select
18720   End If

        ' ** Ctrl keys.
18730   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
18740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18750       With Me
18760         intRetVal = 0
18770         .opgAssetSource.SetFocus
18780       End With
18790     End Select
18800   End If

EXITP:
18810   KeyCode = intRetVal
18820   Exit Sub

ERRH:
18830   intRetVal = 0
18840   THAT_PROC = THIS_PROC
18850   That_Erl = Erl
18860   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18870   Resume EXITP

End Sub

Private Sub cmbAssets_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_MouseMove"

18910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
18920   Exit Sub

ERRH:
18930   Select Case ERR.Number
        Case Else
18940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18950   End Select
18960   Resume EXITP

End Sub

Private Sub cmbAssets_LostFocus()

19000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_LostFocus"

19010   With Me
19020     If IsNull(.cmbAssets) = False Then
19030       If .cmbAssets > 0 And blnAssetUpdated = True And blnAssetChecked = False Then
19040         blnExitingAssetNo = True
19050         .TimerInterval = 100&
19060       End If
19070     End If
19080   End With  ' ** Me.

EXITP:
19090   Exit Sub

ERRH:
19100   THAT_PROC = THIS_PROC
19110   That_Erl = Erl
19120   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19130   Resume EXITP

End Sub

Private Sub curr_id_cmd_DblClick(Cancel As Integer)

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_cmd_DblClick"

19210   With Me
19220     If IsNull(.cmbAssets) = False Then
19230       Select Case .curr_date.Visible
            Case True
19240         .curr_date.Visible = False
19250       Case False
19260         .curr_date = .curr_id.Column(CBX_C_DATE)
19270         .curr_date.Visible = True
19280       End Select
19290     End If
19300   End With

EXITP:
19310   Exit Sub

ERRH:
19320   Select Case ERR.Number
        Case Else
19330     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
19340   End Select
19350   Resume EXITP

End Sub

Private Sub curr_id_AfterUpdate()

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_AfterUpdate"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String
        Dim lngX As Long

19410   With Me

19420     If IsNull(.curr_id) = False Then

19430       lngCurrID = .curr_id.Column(CBX_C_CURRID)

19440       blnFound = False
19450       If lngCurrID <> 150& Then
19460         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
19470           NoChar_Load  ' ** Procedure: Below.
19480         End If
19490         For lngX = 0& To (lngNoChars - 1&)
19500           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
19510             blnFound = True
19520           End If
19530         Next
19540       End If

19550       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
19560         .amount.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
19570         .amount.FontSize = 10
19580         .amount.FontBold = True
19590         .pershare.FontName = "Courier New"
19600         .pershare.FontSize = 10
19610         .pershare.FontBold = True
19620       Case False
19630         If .amount.FontName <> "Arial" Then
19640           .amount.FontName = "Arial"
19650           .amount.FontSize = 10
19660           .amount.FontBold = False
19670           .pershare.FontName = "Arial"
19680           .pershare.FontSize = 10
19690           .pershare.FontBold = False
19700         End If
19710       End Select

19720       If lngCurrID = 150& Then  ' ** USD.
19730         strFrmt1 = "Currency"
19740         intDec1 = 2
19750         strFrmt2 = "Currency"
19760         intDec2 = 5
19770         lngBkClr = CLR_WHT
19780       Else
19790         strTmp01 = .curr_id.Column(CBX_C_SYM)  ' ** Symbol.
19800         If .curr_id.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
19810           intDec1 = 4
19820         Else
19830           intDec1 = .curr_id.Column(CBX_C_DEC)  ' ** Decimal places.
19840         End If
19850         intDec2 = 5
19860         intLen = Len(strTmp01)
19870         Select Case intLen
              Case 1
                'strTmp01 = strTmp01
19880         Case 2
19890           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2)
19900         Case 3
19910           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3)
19920         Case 4
19930           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & "\" & Mid(strTmp01, 4)
19940         Case 5
19950           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & _
                  "\" & Mid(strTmp01, 4, 1) & "\" & Mid(strTmp01, 5)
19960         End Select
19970         Select Case intDec1
              Case 0
19980           strFrmt1 = strTmp01 & "#,##0;(" & strTmp01 & "#,##0);" & strTmp01 & "0;" & strTmp01 & "0"
19990           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
20000         Case 1  ' ** None have this currently.
20010           strFrmt1 = strTmp01 & "#,##0.0;(" & strTmp01 & "#,##0.0);" & strTmp01 & "0.0;" & strTmp01 & "0.0"
20020           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
20030         Case 2
20040           strFrmt1 = strTmp01 & "#,##0.00;(" & strTmp01 & "#,##0.00);" & strTmp01 & "0.00;" & strTmp01 & "0.00"
20050           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
20060         Case 3
20070           strFrmt1 = strTmp01 & "#,##0.000;(" & strTmp01 & "#,##0.000);" & strTmp01 & "0.000;" & strTmp01 & "0.000"
20080           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
20090         End Select
20100         lngBkClr = CLR_VLTGRN
20110       End If
20120       .amount.Format = strFrmt1
20130       .amount.DecimalPlaces = intDec1
20140       .amount.BackColor = lngBkClr
20150       .pershare.Format = strFrmt2
20160       .pershare.DecimalPlaces = intDec2
20170       .pershare.BackColor = lngBkClr

20180       If lngCurrID = 150& Then
20190         .Amount_usd.Visible = False
20200         .pershare_usd.Visible = False
20210       Else
20220         .Amount_usd = 0
20230         .Amount_usd.Visible = True
20240         .pershare_usd = 0
20250         .pershare_usd.Visible = True
20260       End If

20270     End If

20280   End With

EXITP:
20290   Exit Sub

ERRH:
20300   THAT_PROC = THIS_PROC
20310   That_Erl = Erl
20320   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20330   Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

20410   With Me
20420     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
20430     .curr_id.SetFocus
20440     Response = acDataErrContinue
20450   End With

EXITP:
20460   Exit Sub

ERRH:
20470   THAT_PROC = THIS_PROC
20480   That_Erl = Erl
20490   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20500   Resume EXITP

End Sub

Private Sub Amount_KeyDown(KeyCode As Integer, Shift As Integer)

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "Amount_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20620   intShiftDown = (Shift And acShiftMask) > 0
20630   intAltDown = (Shift And acAltMask) > 0
20640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20670       With Me
20680         intRetVal = 0
20690         If .pershare.Enabled = True Then
20700 On Error Resume Next
20710           .pershare.SetFocus
20720           If ERR.Number <> 0 Then
20730 On Error GoTo ERRH
20740             .assetdate.SetFocus
20750           Else
20760 On Error GoTo ERRH
20770           End If
20780         Else
20790           .assetdate.SetFocus
20800         End If
20810       End With
20820     End Select
20830   End If

        ' ** Shift keys.
20840   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20850     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20860       With Me
20870         intRetVal = 0
20880         blnDrivingMeBackwards = True
20890         .cmbAssets.SetFocus
20900       End With
20910     End Select
20920   End If

EXITP:
20930   KeyCode = intRetVal
20940   Exit Sub

ERRH:
20950   intRetVal = 0
20960   THAT_PROC = THIS_PROC
20970   That_Erl = Erl
20980   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20990   Resume EXITP

End Sub

Private Sub amount_Exit(Cancel As Integer)

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "amount_Exit"

        Dim dblRate2 As Double
        Dim intPos01 As Integer
        Dim blnContinue As Boolean
        Dim dblTmp01 As Double

21010   With Me

21020     blnContinue = True

21030     lngCurrID = .curr_id.Column(CBX_C_CURRID)

21040     If IsNumeric(.amount) = True Then
21050       If .amount < 0 Then
21060         blnContinue = False
21070         MsgBox "Amount must be positive.", vbInformation + vbOKOnly, "Invalid Entry"
21080         Cancel = -1
21090         .amount.SetFocus
21100       Else
21110         If CDbl(Rem_Dollar(.amount.text, lngCurrID)) > 999999999.99 Then  ' ** Module Function: modStringFuncs.
21120           blnContinue = False
21130           MsgBox "Amount is too large.", vbInformation + vbOKOnly, "Invalid Entry"
21140           .amount = vbNullString
21150           Cancel = -1
21160           .amount.SetFocus
21170         Else
21180           intPos01 = InStr(1, .amount, ".", vbTextCompare)
21190           If intPos01 > 0 Then
21200             If Len(.amount) - intPos01 > 2 Then
21210               blnContinue = False
21220               MsgBox "You can only enter Amount with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
21230               .amount = vbNullString
21240               Cancel = -1
21250             End If
21260           End If
21270           If blnContinue = True Then
21280             If Nz(.shareface, 0) = 0 Then
21290               blnContinue = False
21300               MsgBox "Shares must have a value", vbInformation + vbOKOnly, "Invalid Entry"
21310               .shareface = 1
21320               .cmbAssets.SetFocus
21330               If lngCurrID = 150& Then
21340                 .pershare = (.amount / IIf(Nz(.shareface, 0) = 0, 1, .shareface))  '5 dec.
21350                 .pershare.BackColor = CLR_DISABLED_BG
21360               Else
21370                 .pershare = (.amount / IIf(Nz(.shareface, 0) = 0, 1, .shareface))  '5 dec.
21380                 .pershare.BackColor = CLR_LTGRN  'lngBkClr
21390               End If
21400               .pershare.Locked = True
21410               .pershare.Enabled = False
21420               .pershare.ForeColor = CLR_DISABLED_FG
21430             Else
21440               If blnDrivingMeBackwards = False Then
21450                 If gblnGoToReport = False Then
21460                   .assetdate_display.SetFocus
21470                 Else
21480 On Error Resume Next
21490                   .cmdOK.SetFocus
21500 On Error GoTo ERRH
21510                 End If
21520               Else
21530                 blnDrivingMeBackwards = False
21540               End If
21550               intInputType = 1
21560               dblPerShare_Calc = (.amount / IIf(Nz(.shareface, 0) = 0, 1, .shareface))
21570               If lngCurrID = 150& Then
21580                 .pershare = dblPerShare_Calc  '5 dec.
21590                 .pershare.BackColor = CLR_DISABLED_BG
21600               Else
21610                 .pershare = dblPerShare_Calc  '5 dec.
21620                 .pershare.BackColor = CLR_LTGRN  'lngBkClr
21630                 dblRate2 = .curr_id.Column(CBX_C_RATE2)
21640                 dblTmp01 = (.amount * dblRate2)
21650                 .Amount_usd = Round(dblTmp01, 2)
21660                 .Amount_usd.Visible = True
21670                 dblTmp01 = (dblPerShare_Calc * dblRate2)
21680                 dblTmp01 = Round(dblTmp01, 5)
21690                 .pershare_usd = dblTmp01
21700                 .pershare_usd.Visible = True
21710               End If
21720               .pershare.Locked = True
21730               .pershare.Enabled = False
21740               .pershare.ForeColor = CLR_DISABLED_FG
21750             End If
21760           End If
21770         End If
21780       End If
21790     Else
21800       .pershare.Enabled = True
21810       .pershare.Locked = False
21820       .pershare.ForeColor = CLR_BLK
21830       If lngCurrID = 150& Then
21840         .pershare.BackColor = CLR_WHT
21850       Else
21860         .pershare.BackColor = CLR_VLTGRN
21870       End If
21880     End If

21890   End With

EXITP:
21900   Exit Sub

ERRH:
21910   Select Case ERR.Number
        Case Else
21920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21930   End Select
21940   Resume EXITP

End Sub

Private Sub pershare_Exit(Cancel As Integer)

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "pershare_Exit"

        Dim dblRate2 As Double
        Dim intPos01 As Integer
        Dim blnContinue As Boolean
        Dim dblTmp01 As Double

22010   With Me

22020     blnContinue = True

22030     lngCurrID = .curr_id.Column(CBX_C_CURRID)

22040     If IsNumeric(.pershare) = True Then
22050       If .pershare < 0 Then
22060         blnContinue = False
22070         MsgBox "Per Share must be positive.", vbInformation + vbOKOnly, "Invalid Entry"
22080         Cancel = -1
22090         .pershare.SetFocus
22100       Else
22110         If CDbl(Rem_Dollar(.pershare.text, lngCurrID)) > 999999999.9999 Then  ' ** Module Function: modStringFuncs.
22120           blnContinue = False
22130           MsgBox "Per Share is too large.", vbInformation + vbOKOnly, "Invalid Entry"
22140           .pershare = vbNullString
22150           Cancel = -1
22160           .pershare.SetFocus
22170         Else
22180           intPos01 = InStr(1, .pershare, ".", vbTextCompare)
22190           If intPos01 > 0 Then
22200             If Len(.pershare) - intPos01 > 5 Then
22210               blnContinue = False
22220               MsgBox "You can only enter Per Share with up to 5 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
22230               .pershare = vbNullString
22240               Cancel = -1
22250               .pershare.SetFocus
22260             End If
22270           End If
22280           If blnContinue = True Then
22290             intInputType = 2
22300             If lngCurrID = 150& Then
22310               .amount = (.pershare * Nz(.shareface, 1))
22320               .amount.BackColor = CLR_DISABLED_BG
22330             Else
22340               .amount = (.pershare * Nz(.shareface, 1))
22350               .amount.BackColor = CLR_LTGRN  'lngBkClr
22360               dblRate2 = .curr_id.Column(CBX_C_RATE2)
22370               dblTmp01 = (.pershare * dblRate2)
22380               dblTmp01 = Round(dblTmp01, 5)
22390               .pershare_usd = dblTmp01
22400               .pershare_usd.Visible = True
22410               dblTmp01 = (.amount * dblRate2)
22420               dblTmp01 = Round(dblTmp01, 2)
22430               .Amount_usd = dblTmp01
22440               .Amount_usd.Visible = True
22450             End If
22460             .amount.Locked = True
22470             .amount.Enabled = False
22480             .amount.ForeColor = CLR_DISABLED_FG
22490           End If
22500         End If
22510       End If
22520     Else
22530       .amount.Enabled = True
22540       .amount.Locked = False
22550       .amount.ForeColor = CLR_BLK
22560       If lngCurrID = 150& Then
22570         .amount.BackColor = CLR_WHT
22580       Else
22590         .amount.BackColor = CLR_VLTGRN
22600       End If
22610     End If

22620   End With

EXITP:
22630   Exit Sub

ERRH:
22640   Select Case ERR.Number
        Case Else
22650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22660   End Select
22670   Resume EXITP

End Sub

Private Sub assetdate_display_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_lbl2_MouseMove"

22710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
22720   Exit Sub

ERRH:
22730   Select Case ERR.Number
        Case Else
22740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22750   End Select
22760   Resume EXITP

End Sub

Private Sub assetdate_display_Enter()

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_Enter"

        Dim strTmp01 As String, dblTmp02 As Double

22810   With Me
22820     If IsNull(.assetdate) = True Then
22830       datAssetDate_OldValue = 0
22840       lngAssetDate_OldValue = 0&
22850       .assetdate_display = Null
22860     Else
22870       If IsDate(.assetdate) = False Then
22880         datAssetDate_OldValue = 0
22890         lngAssetDate_OldValue = 0&
22900         .assetdate = Null
22910         .assetdate_display = Null
22920       Else
22930         datAssetDate_OldValue = .assetdate
22940         dblTmp02 = CDbl(datAssetDate_OldValue)
22950         strTmp01 = CStr(dblTmp02)
22960         If InStr(strTmp01, ".") > 0 Then
22970           strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
22980         Else
22990           datAssetDate_OldValue = datAssetDate_OldValue + time()
23000         End If
23010         lngAssetDate_OldValue = CLng(strTmp01)
23020         .assetdate_display.text = Format(CDate(lngAssetDate_OldValue), "mm/dd/yyyy")  ' ** Display without timestamp
23030       End If
23040     End If
23050     blnAssetDateChecked = False
23060   End With

EXITP:
23070   Exit Sub

ERRH:
23080   THAT_PROC = THIS_PROC
23090   That_Erl = Erl
23100   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23110   Resume EXITP

End Sub

Private Sub assetdate_display_BeforeUpdate(Cancel As Integer)

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_BeforeUpdate"

23210   With Me
23220     If .assetdate_display.text <> "__/__/____" And .assetdate_display.text <> vbNullString Then
23230       If Left(.assetdate_display.text, 2) > 12 Then
23240         MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Date"
23250         Cancel = -1
23260         .assetdate_display.Undo
23270       End If
23280     End If
23290   End With

EXITP:
23300   Exit Sub

ERRH:
23310   THAT_PROC = THIS_PROC
23320   That_Erl = Erl
23330   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23340   Resume EXITP

End Sub

Private Sub assetdate_display_AfterUpdate()

23400 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_AfterUpdate"

        Dim strTmp01 As String, lngTmp02 As Long, dblTmp03 As Double, datTmp04 As Date

23410   With Me
23420     If IsNull(.assetdate_display) = True Then
23430       datAssetDate_OldValue = Now()
23440       lngAssetDate_OldValue = CLng(Date)
23450       .assetdate = datAssetDate_OldValue  ' ** Save with timestamp.
23460       .assetdate_display = CDate(lngAssetDate_OldValue)  ' ** Display without timestamp.
23470     Else
23480       If Trim(.assetdate_display) = vbNullString Then
23490         datAssetDate_OldValue = Now()
23500         lngAssetDate_OldValue = CLng(Date)
23510         .assetdate = datAssetDate_OldValue  ' ** Save with timestamp.
23520         .assetdate_display = CDate(lngAssetDate_OldValue)  ' ** Display without timestamp.
23530       Else
23540         If DateCheck_Trade(.assetdate_display.text) = True Then  ' ** Module Function: modUtilities.
23550           datTmp04 = CDate(.assetdate_display.text)
23560           dblTmp03 = CDbl(datTmp04)
23570           strTmp01 = CStr(dblTmp03)
23580           If InStr(strTmp01, ".") > 0 Then  ' ** There shouldn't be.
23590             strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
23600           End If
23610           lngTmp02 = CLng(strTmp01)
23620           If CDbl(datAssetDate_OldValue) = 0# Then
23630             datAssetDate_OldValue = datTmp04 + time()
23640             lngAssetDate_OldValue = lngTmp02
23650             .assetdate = datAssetDate_OldValue  ' ** Save with timestamp.
23660           Else
23670             If lngTmp02 <> lngAssetDate_OldValue Then
23680               datAssetDate_OldValue = datTmp04 + time()
23690               lngAssetDate_OldValue = lngTmp02
23700               .assetdate = datAssetDate_OldValue  ' ** Save with timestamp.
23710             Else
                    ' ** Let assetdate stand.
23720             End If
23730           End If
23740         Else
23750           datAssetDate_OldValue = Now()
23760           lngAssetDate_OldValue = CLng(Date)
23770           .assetdate = datAssetDate_OldValue  ' ** Save with timestamp.
23780           .assetdate_display = CDate(lngAssetDate_OldValue)  ' ** Display without timestamp.
23790         End If
23800       End If
23810     End If
23820     blnAssetDateChecked = True
23830   End With

EXITP:
23840   Exit Sub

ERRH:
23850   THAT_PROC = THIS_PROC
23860   That_Erl = Erl
23870   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23880   Resume EXITP

End Sub

Private Sub assetdate_display_KeyDown(KeyCode As Integer, Shift As Integer)

23900 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23920   intShiftDown = (Shift And acShiftMask) > 0
23930   intAltDown = (Shift And acAltMask) > 0
23940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23970       With Me
23980         intRetVal = 0
23990         .transdate.SetFocus
24000       End With
24010     End Select
24020   End If

        ' ** Shift keys.
24030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24050       With Me
24060         intRetVal = 0
24070         If .pershare.Enabled = True Then
24080           .pershare.SetFocus
24090         Else
24100           .amount.SetFocus
24110         End If
24120       End With
24130     End Select
24140   End If

        ' ** Ctrl keys.
24150   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
24160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24170       With Me
24180         intRetVal = 0
24190         .cmdCalendar2.SetFocus
24200       End With
24210     End Select
24220   End If

EXITP:
24230   KeyCode = intRetVal
24240   Exit Sub

ERRH:
24250   intRetVal = 0
24260   THAT_PROC = THIS_PROC
24270   That_Erl = Erl
24280   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
24290   Resume EXITP

End Sub

Private Sub assetdate_display_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_MouseMove"

24310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
24320   Exit Sub

ERRH:
24330   Select Case ERR.Number
        Case Else
24340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24350   End Select
24360   Resume EXITP

End Sub

Private Sub assetdate_display_Exit(Cancel As Integer)

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_Exit"

        Dim varTmp00 As Variant

24410   With Me
24420 On Error Resume Next
24430     varTmp00 = .assetdate_display.text  ' ** When called from cmdCalendar2_Click(), below, it's not always on focus.
24440     If ERR.Number <> 0 Then
24450       varTmp00 = .assetdate_display
24460     End If
24470 On Error GoTo ERRH
24480     If IsNull(varTmp00) = True Then
24490       If IsNull(.assetdate_display) = True Then
24500         .assetdate = Now()
24510         .assetdate_display = CDate(Format(.assetdate, "mm/dd/yyyy"))
24520       Else
24530         If IsDate(.assetdate_display) = True Then
24540           varTmp00 = .assetdate_display
24550         Else
24560           .assetdate = Now()
24570           .assetdate_display = CDate(Format(.assetdate, "mm/dd/yyyy"))
24580         End If
24590       End If
24600     Else
24610       If IsDate(varTmp00) = False Then
24620         .assetdate = Now()
24630         .assetdate_display = CDate(Format(.assetdate, "mm/dd/yyyy"))
24640         varTmp00 = Null
24650       End If
24660     End If
24670     If IsNull(varTmp00) = False Then
24680       If blnAssetDateChecked = False Then
24690         If DateCheck_Trade(varTmp00) = True Then  ' ** Module Function: modUtilities.
                ' ** '.Text' needed to reliably check date.
24700           .assetdate = .assetdate_display + time()
24710         Else
24720           If IsNull(.assetdate_display) = True Then
24730             .assetdate_display = Date
24740             .assetdate = Now()
24750           Else
24760             If Trim(.assetdate_display) = vbNullString Then
24770               .assetdate_display = Date
24780               .assetdate = Now()
24790             Else
24800               Cancel = -1
24810               .assetdate_display.SetFocus
24820             End If
24830           End If
24840         End If
24850       Else
24860         blnAssetDateChecked = False
24870       End If
24880     End If
24890   End With

EXITP:
24900   Exit Sub

ERRH:
24910   THAT_PROC = THIS_PROC
24920   That_Erl = Erl
24930   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** Ignore.
24940   Case Else
24950     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
24960   End Select
24970   Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

25010   With Me
25020     blnCalendar2_Focus = True
25030     .cmdCalendar2_raised_semifocus_dots_img.Visible = True
25040     .cmdCalendar2_raised_img.Visible = False
25050     .cmdCalendar2_raised_focus_img.Visible = False
25060     .cmdCalendar2_raised_focus_dots_img.Visible = False
25070     .cmdCalendar2_sunken_focus_dots_img.Visible = False
25080     .cmdCalendar2_raised_img_dis.Visible = False
25090   End With

EXITP:
25100   Exit Sub

ERRH:
25110   Select Case ERR.Number
        Case Else
25120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25130   End Select
25140   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

25210   With Me
25220     blnCalendar2_MouseDown = True
25230     .cmdCalendar2_sunken_focus_dots_img.Visible = True
25240     .cmdCalendar2_raised_img.Visible = False
25250     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
25260     .cmdCalendar2_raised_focus_img.Visible = False
25270     .cmdCalendar2_raised_focus_dots_img.Visible = False
25280     .cmdCalendar2_raised_img_dis.Visible = False
25290   End With

EXITP:
25300   Exit Sub

ERRH:
25310   Select Case ERR.Number
        Case Else
25320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25330   End Select
25340   Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

25410   With Me
25420     datStartDate = Date
25430     datEndDate = 0
25440     blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
25450     If blnRetVal = True Then
25460       If Compare_DateA_DateB(datStartDate, ">", Date) = True Then  ' ** Module Function: modStringFuncs.
25470         MsgBox "Future trade dates are not allowed.", vbInformation + vbOKOnly, "Invalid Date"
25480         .assetdate = Now()
25490         .assetdate_display = CDate(Format(.assetdate, "mm/dd/yyyy"))
25500       Else
25510         .assetdate = datStartDate + time
25520         .assetdate_display = CDate(Format(.assetdate, "mm/dd/yyyy"))
25530       End If
25540     Else
25550       .assetdate = Now()
25560       .assetdate_display = CDate(Format(.assetdate, "mm/dd/yyyy"))
25570     End If
25580     .transdate.SetFocus
25590   End With

EXITP:
25600   Exit Sub

ERRH:
25610   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
25620   Case Else
25630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25640   End Select
25650   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

25710   With Me
25720     If blnCalendar2_MouseDown = False Then
25730       Select Case blnCalendar2_Focus
            Case True
25740         .cmdCalendar2_raised_focus_dots_img.Visible = True
25750         .cmdCalendar2_raised_focus_img.Visible = False
25760       Case False
25770         .cmdCalendar2_raised_focus_img.Visible = True
25780         .cmdCalendar2_raised_focus_dots_img.Visible = False
25790       End Select
25800       .cmdCalendar2_raised_img.Visible = False
25810       .cmdCalendar2_raised_semifocus_dots_img.Visible = False
25820       .cmdCalendar2_sunken_focus_dots_img.Visible = False
25830       .cmdCalendar2_raised_img_dis.Visible = False
25840     End If
25850   End With

EXITP:
25860   Exit Sub

ERRH:
25870   Select Case ERR.Number
        Case Else
25880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25890   End Select
25900   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

26000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

26010   With Me
26020     .cmdCalendar2_raised_focus_dots_img.Visible = True
26030     .cmdCalendar2_raised_img.Visible = False
26040     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
26050     .cmdCalendar2_raised_focus_img.Visible = False
26060     .cmdCalendar2_sunken_focus_dots_img.Visible = False
26070     .cmdCalendar2_raised_img_dis.Visible = False
26080     blnCalendar2_MouseDown = False
26090   End With

EXITP:
26100   Exit Sub

ERRH:
26110   Select Case ERR.Number
        Case Else
26120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26130   End Select
26140   Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

26200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

26210   With Me
26220     .cmdCalendar2_raised_img.Visible = True
26230     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
26240     .cmdCalendar2_raised_focus_img.Visible = False
26250     .cmdCalendar2_raised_focus_dots_img.Visible = False
26260     .cmdCalendar2_sunken_focus_dots_img.Visible = False
26270     .cmdCalendar2_raised_img_dis.Visible = False
26280     blnCalendar2_Focus = False
26290   End With

EXITP:
26300   Exit Sub

ERRH:
26310   Select Case ERR.Number
        Case Else
26320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26330   End Select
26340   Resume EXITP

End Sub

Private Sub transdate_KeyDown(KeyCode As Integer, Shift As Integer)

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26420   intShiftDown = (Shift And acShiftMask) > 0
26430   intAltDown = (Shift And acAltMask) > 0
26440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26470       With Me
26480         intRetVal = 0
26490         .cmdOK.SetFocus
26500       End With
26510     End Select
26520   End If

        ' ** Shift keys.
26530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26550       With Me
26560         intRetVal = 0
26570         If .assetdate_display.Enabled = True Then
26580           .assetdate_display.SetFocus
26590         ElseIf .pershare.Enabled = True Then
26600           .pershare.SetFocus
26610         Else
26620           .amount.SetFocus
26630         End If
26640       End With
26650     End Select
26660   End If

        ' ** Ctrl keys.
26670   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
26680     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26690       With Me
26700         intRetVal = 0
26710         .cmdCalendar1.SetFocus
26720       End With
26730     End Select
26740   End If

        ' ** Ctrl-Shift keys.
26750   If intCtrlDown And (Not intAltDown) And intShiftDown Then
26760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26770       With Me
26780         intRetVal = 0
26790         .cmdCalendar2.SetFocus
26800       End With
26810     End Select
26820   End If

EXITP:
26830   KeyCode = intRetVal
26840   Exit Sub

ERRH:
26850   intRetVal = 0
26860   THAT_PROC = THIS_PROC
26870   That_Erl = Erl
26880   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
26890   Resume EXITP

End Sub

Private Sub transdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_MouseMove"

26910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
26920   Exit Sub

ERRH:
26930   Select Case ERR.Number
        Case Else
26940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26950   End Select
26960   Resume EXITP

End Sub

Private Sub transdate_Exit(Cancel As Integer)

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_Exit"

        Dim varTmp00 As Variant

27010   With Me
27020 On Error Resume Next
27030     varTmp00 = .transdate.text  ' ** When called from cmdCalendar1_Click(), below, it's not always on focus.
27040     If ERR.Number <> 0 Then
27050       varTmp00 = .transdate
27060     End If
27070 On Error GoTo ERRH
27080     If DateCheck_Post(varTmp00) = True Then  ' ** Module Function: modUtilities.
            ' ** '.Text' needed to reliably check date.
27090     Else
27100       If IsNull(.transdate) = True Then
27110         .transdate = Date
27120       Else
27130         If Trim(.transdate) = vbNullString Then
27140           .transdate = Date
27150         Else
27160           Cancel = -1
27170           .transdate.SetFocus
27180         End If
27190       End If
27200     End If
27210   End With

EXITP:
27220   Exit Sub

ERRH:
27230   Select Case ERR.Number
        Case Else
27240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27250   End Select
27260   Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

27300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

27310   With Me
27320     blnCalendar1_Focus = True
27330     .cmdCalendar1_raised_semifocus_dots_img.Visible = True
27340     .cmdCalendar1_raised_img.Visible = False
27350     .cmdCalendar1_raised_focus_img.Visible = False
27360     .cmdCalendar1_raised_focus_dots_img.Visible = False
27370     .cmdCalendar1_sunken_focus_dots_img.Visible = False
27380     .cmdCalendar1_raised_img_dis.Visible = False
27390   End With

EXITP:
27400   Exit Sub

ERRH:
27410   Select Case ERR.Number
        Case Else
27420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27430   End Select
27440   Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

27500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

27510   With Me
27520     blnCalendar1_MouseDown = True
27530     .cmdCalendar1_sunken_focus_dots_img.Visible = True
27540     .cmdCalendar1_raised_img.Visible = False
27550     .cmdCalendar1_raised_semifocus_dots_img.Visible = False
27560     .cmdCalendar1_raised_focus_img.Visible = False
27570     .cmdCalendar1_raised_focus_dots_img.Visible = False
27580     .cmdCalendar1_raised_img_dis.Visible = False
27590   End With

EXITP:
27600   Exit Sub

ERRH:
27610   Select Case ERR.Number
        Case Else
27620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27630   End Select
27640   Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

27700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

27710   With Me
27720     datStartDate = Date
27730     datEndDate = 0
27740     blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
27750     If blnRetVal = True Then
            ' ** Allow posting up to 1 month into the future.
27760       If datStartDate > DateAdd("m", 1, Date) Then
27770         MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
27780         .transdate = CDate(Format(Date, "mm/dd/yyyy"))
27790       Else
27800         .transdate = datStartDate
27810       End If
27820     Else
27830       .transdate = CDate(Format(Date, "mm/dd/yyyy"))
27840     End If
          ' ** Make the date 'is edited' as if were entered by hand.
27850     .transdate.SetFocus
27860     Cancel = 0
27870     transdate_Exit Cancel  ' ** Procedure: Above.
27880     If Cancel = 0 Then
27890       .cmdOK.SetFocus
27900     End If
27910   End With

EXITP:
27920   Exit Sub

ERRH:
27930   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
27940   Case Else
27950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27960   End Select
27970   Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

28010   With Me
28020     If blnCalendar1_MouseDown = False Then
28030       Select Case blnCalendar1_Focus
            Case True
28040         .cmdCalendar1_raised_focus_dots_img.Visible = True
28050         .cmdCalendar1_raised_focus_img.Visible = False
28060       Case False
28070         .cmdCalendar1_raised_focus_img.Visible = True
28080         .cmdCalendar1_raised_focus_dots_img.Visible = False
28090       End Select
28100       .cmdCalendar1_raised_img.Visible = False
28110       .cmdCalendar1_raised_semifocus_dots_img.Visible = False
28120       .cmdCalendar1_sunken_focus_dots_img.Visible = False
28130       .cmdCalendar1_raised_img_dis.Visible = False
28140     End If
28150   End With

EXITP:
28160   Exit Sub

ERRH:
28170   Select Case ERR.Number
        Case Else
28180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28190   End Select
28200   Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

28300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

28310   With Me
28320     .cmdCalendar1_raised_focus_dots_img.Visible = True
28330     .cmdCalendar1_raised_img.Visible = False
28340     .cmdCalendar1_raised_semifocus_dots_img.Visible = False
28350     .cmdCalendar1_raised_focus_img.Visible = False
28360     .cmdCalendar1_sunken_focus_dots_img.Visible = False
28370     .cmdCalendar1_raised_img_dis.Visible = False
28380     blnCalendar1_MouseDown = False
28390   End With

EXITP:
28400   Exit Sub

ERRH:
28410   Select Case ERR.Number
        Case Else
28420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28430   End Select
28440   Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

28500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

28510   With Me
28520     .cmdCalendar1_raised_img.Visible = True
28530     .cmdCalendar1_raised_semifocus_dots_img.Visible = False
28540     .cmdCalendar1_raised_focus_img.Visible = False
28550     .cmdCalendar1_raised_focus_dots_img.Visible = False
28560     .cmdCalendar1_sunken_focus_dots_img.Visible = False
28570     .cmdCalendar1_raised_img_dis.Visible = False
28580     blnCalendar1_Focus = False
28590   End With

EXITP:
28600   Exit Sub

ERRH:
28610   Select Case ERR.Number
        Case Else
28620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28630   End Select
28640   Resume EXITP

End Sub

Private Function AdvancedRounding(lngAssetNo As Long) As Double

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "AdvancedRounding"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim lngRecs As Long
        Dim lngLocationID As Long
        Dim blnNoDistribution As Boolean
        Dim lngX As Long
        Dim dblRetVal As Double

28710   blnNoDistribution = False
28720   dblRetVal = 0#

28730   Set dbs = CurrentDb
28740   With dbs

          ' ** Step 17: Rounding Stage 3.
28750     Me.Map_Status_lbl.Caption = "17 of 22: Rounding Stage 3 . . ."
28760     DoEvents

28770     Select Case strCallingForm
          Case "frmJournal"
            ' ** qryMap_Rec_02_06 (xx), just 'INCOME O/U', linked to qryMap_Rec_02_09 (xx), with pcash_per_acct, by specified [actno].
28780       Set qdf = .QueryDefs("qryMap_Rec_02_10")
28790     Case "frmJournal_Columns"
            ' ** qryJournal_Columns_35_05_06 (xx), just 'INCOME O/U', linked to qryJournal_Columns_35_05_09 (xx), with pcash_per_acct, by specified [actno].
28800       Set qdf = .QueryDefs("qryJournal_Columns_35_05_10")
28810     End Select
28820     With qdf.Parameters
28830       ![actno] = strAccountNo_OU
28840     End With
28850     Set rst1 = qdf.OpenRecordset
28860     With rst1
28870       If .BOF = True And .EOF = True Then
              ' ** No INCOME O/U to distribute!
28880         blnNoDistribution = True
28890       Else
28900         .MoveFirst
28910         If ![total_pcash_ou] < 1@ Then
                ' ** Less than a buck, leave it alone.
28920           blnNoDistribution = True
28930         Else
28940           If ![pennies_per_acct] < 0.01 Then
                  ' ** Too small to divvy up!
28950             blnNoDistribution = True
28960           Else
28970             If ![pcash_per_acct] < 0.01 Then
                    ' ** Rounded value too small to divvy up!
28980               blnNoDistribution = True
28990             Else
                    ' ** Continue with the distribution.
29000             End If
29010           End If
29020         End If
29030       End If
29040       .Close
29050     End With  ' ** rst1.
29060     Set rst1 = Nothing
29070     Set qdf = Nothing
29080     DoEvents

          ' ** Step 18: Empty tblJournal_Map_Staging2.
29090     Me.Map_Status_lbl.Caption = "18 of 22: Empty temporary table 2 . . ."
29100     DoEvents

29110     Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty tblJournal_Map_Staging2.
29120       Set qdf = .QueryDefs("qryMap_Rec_02_16")
29130     Case "frmJournal_Columns"
            ' ** Empty tblJournal_Map_Staging2.
29140       Set qdf = .QueryDefs("qryJournal_Columns_35_05_16")
29150     End Select
29160     qdf.Execute
29170     Set qdf = Nothing
29180     DoEvents

29190     Select Case blnNoDistribution
          Case True

            ' ** Step 19: Collect unchanged distributions.
29200       Me.Map_Status_lbl.Caption = "19 of 22: Collect unchanged distributions . . ."
29210       DoEvents

29220       Select Case strCallingForm
            Case "frmJournal"
              ' ** Append qryMap_Rec_02_06 (xx) to tblJournal_Map_Staging2.
29230         Set qdf = .QueryDefs("qryMap_Rec_02_15")
29240       Case "frmJournal_Columns"
              ' ** Append qryJournal_Columns_35_05_06 (xx) to tblJournal_Map_Staging2.
29250         Set qdf = .QueryDefs("qryJournal_Columns_35_05_15")
29260       End Select
29270       qdf.Execute
29280       Set qdf = Nothing

29290     Case False

29300       Set rst2 = .OpenRecordset("tblJournal_Map_Staging2", dbOpenDynaset, dbConsistent)

            ' ** Step 19: Create new distributions.
29310       Me.Map_Status_lbl.Caption = "19 of 22: Create new distributions . . ."
29320       DoEvents

29330       Select Case strCallingForm
            Case "frmJournal"
              ' ** qryMap_Rec_02_08 (xx), linked to qryMap_Rec_02_11 (xx), with total_shareface_new, total_pcash_new.
29340         Set qdf = .QueryDefs("qryMap_Rec_02_12")
29350       Case "frmJournal_Columns"
              ' ** qryJournal_Columns_35_05_08 (xx), linked to qryJournal_Columns_35_05_11 (xx), with total_shareface_new, total_pcash_new.
29360         Set qdf = .QueryDefs("qryJournal_Columns_35_05_12")
29370       End Select
29380       With qdf.Parameters
29390         ![actno] = strAccountNo_OU
29400       End With
29410       Set rst1 = qdf.OpenRecordset
29420       With rst1
29430         .MoveLast
29440         lngRecs = .RecordCount
29450         .MoveFirst
29460         For lngX = 1& To lngRecs
29470           With rst2
29480             .AddNew
                  ' ** ![jmap_id] : AutoNumber.
29490             ![accountno] = rst1![accountno]
29500             ![assetno] = rst1![assetno]
29510             ![transdate] = rst1![transdate]
29520             ![assetdate] = rst1![assetdate]
29530             ![shareface] = rst1![total_shareface_new]
29540             ![pershare] = rst1![pershare]
29550             ![journaltype] = "Received"
29560             ![ICash] = 0@
29570             ![PCash] = rst1![total_pcash_new]
29580             ![Cost] = 0@
29590             ![description] = "Long Term Capital Gain"
29600             ![rate] = 0#
29610             ![due] = Null
29620             ![Location_ID] = rst1![Location_ID]
29630             .Update
29640           End With  ' ** rst2.
29650           DoEvents
29660           If lngX < lngRecs Then .MoveNext
29670         Next  ' ** lngX.
29680         .Close
29690       End With  ' ** rst1.
29700       Set rst1 = Nothing
29710       Set qdf = Nothing
29720       DoEvents

            ' ** Step 20: Determine Over/Under location.
29730       Me.Map_Status_lbl.Caption = "20 of 22: Determine Over/Under location . . ."
29740       DoEvents

29750       Select Case strCallingForm
            Case "frmJournal"
              ' ** Ledger, just 'INCOME O/U', 'Received', grouped by accountno, with Location_ID_min, Location_ID_max, cnt, by specified [actno], [astno].
29760         Set qdf = .QueryDefs("qryMap_Rec_02_05")
29770       Case "frmJournal_Columns"
              ' ** Ledger, just 'INCOME O/U', 'Received', grouped by accountno, with Location_ID_min, Location_ID_max, cnt, by specified [actno], [astno].
29780         Set qdf = .QueryDefs("qryJournal_Columns_35_05_05")
29790       End Select
29800       With qdf.Parameters
29810         ![actno] = strAccountNo_OU
29820         ![astno] = lngAssetNo
29830       End With
29840       Set rst1 = qdf.OpenRecordset
29850       With rst1
29860         If .BOF = True And .EOF = True Then
29870           lngLocationID = 1&  ' ** Default.
29880         Else
29890           .MoveFirst
                ' ** CLng(IIf([cnt]>2,1, IIf([cnt]=1,[Location_ID_max],IIf([cnt]=2 And [Location_ID_min]=1,[Location_ID_max],1))))
29900           If ![cnt] > 2 Then
29910             lngLocationID = 1&
29920           Else
29930             If ![cnt] = 1 Then
29940               lngLocationID = ![Location_ID_max]
29950             Else
29960               If ![cnt] = 2 And [Location_ID_min] = 1 Then
29970                 lngLocationID = ![Location_ID_max]
29980               Else
29990                 lngLocationID = 1&
30000               End If
30010             End If
30020           End If
30030         End If
30040         .Close
30050       End With  ' ** rst1.
30060       Set rst1 = Nothing
30070       Set qdf = Nothing
30080       DoEvents

            ' ** Step 21: Create new Over/Under entry.
30090       Me.Map_Status_lbl.Caption = "22 of 22: Create new Over/Under entry . . ."
30100       DoEvents

30110       Select Case strCallingForm
            Case "frmJournal"
              ' ** qryMap_Rec_02_13 (xx), linked to qryMap_Rec_02_07 (xx), with shareface_ou_new, pcash_ou_new, by specified [actno], [locid].
30120         Set qdf = .QueryDefs("qryMap_Rec_02_14")
30130       Case "frmJournal_Columns"
              ' ** qryJournal_Columns_35_05_13 (xx), linked to qryJournal_Columns_35_05_07 (xx), with shareface_ou_new, pcash_ou_new, by specified [actno], [locid].
30140         Set qdf = .QueryDefs("qryJournal_Columns_35_05_14")
30150       End Select
30160       With qdf.Parameters
30170         ![astno] = lngAssetNo
30180         ![actno] = strAccountNo_OU
30190         ![locid] = lngLocationID
30200       End With
30210       Set rst1 = qdf.OpenRecordset
30220       With rst1
30230         .MoveFirst
30240         With rst2
30250           .AddNew
                ' ** ![jmap_id] : AutoNumber.
30260           ![accountno] = rst1![accountno]
30270           ![assetno] = rst1![assetno]
30280           ![transdate] = rst1![transdate]
30290           ![assetdate] = rst1![assetdate]
30300           ![shareface] = rst1![shareface_ou_new]
30310           ![pershare] = rst1![pershare]
30320           ![journaltype] = "Received"
30330           ![ICash] = 0@
30340           dblRetVal = rst1![pcash_ou_new]
30350           ![PCash] = rst1![pcash_ou_new]
30360           ![Cost] = 0@
30370           ![description] = "Long Term Capital Gain"
30380           ![rate] = 0#
30390           ![due] = Null
30400           ![Location_ID] = rst1![Location_ID]
30410           .Update
30420           .Close
30430         End With  ' ** rst2.
30440         .Close
30450       End With  ' ** rst1.
30460       Set rst1 = Nothing
30470       Set rst2 = Nothing
30480       Set qdf = Nothing

30490     End Select  ' ** blnNoDistribution.

30500     .Close
30510   End With  ' ** dbs.
30520   DoEvents

EXITP:
30530   Set rst1 = Nothing
30540   Set rst1 = Nothing
30550   Set qdf = Nothing
30560   Set dbs = Nothing
30570   AdvancedRounding = dblRetVal
30580   Exit Function

ERRH:
30590   DoCmd.Hourglass False
30600   dblRetVal = 0#
30610   Select Case ERR.Number
        Case Else
30620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30630   End Select
30640   Resume EXITP

End Function

Private Sub IncludeCurrency()

30700 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngTmp01 As Long

30710   With Me
30720     .curr_id.Enabled = False
30730     .curr_id.Locked = True
30740     .curr_id.ForeColor = CLR_DISABLED_FG
30750     .curr_id.BackColor = CLR_VLTGRN
30760     Select Case blnCurrID
          Case True
30770       .curr_id.Visible = True
30780       .curr_id_cmd.Visible = True
30790       .curr_id_cmd.Enabled = True
30800       .cmbAssets_vline01.Visible = True
30810       .cmbAssets_vline02.Visible = True
30820     Case False
30830       .curr_id.Visible = False
30840       .curr_id_cmd.Visible = False
30850       .curr_id_cmd.Enabled = False
30860       .cmbAssets_vline01.Visible = False
30870       .cmbAssets_vline02.Visible = False
30880       lngTmp01 = ((.curr_id.Left + .curr_id.Width) - (.cmbAssets.Left + .cmbAssets.Width))
30890       .cmbAssets.Width = (.cmbAssets.Width + lngTmp01)
30900     End Select
30910     .Amount_usd.Visible = False
30920     .pershare_usd.Visible = False
30930   End With

        ' ** cmbAssets RowSource's:
        ' **   qryMap_Rec_05a
        ' **   qryMap_Rec_05c
        ' **   qryMap_Rec_05d
        ' **   qryMap_Rec_05f
        ' **   qryMap_Rec_05g
        ' **   qryMap_Rec_05i

EXITP:
30940   Exit Sub

ERRH:
30950   DoCmd.Hourglass False
30960   Select Case ERR.Number
        Case Else
30970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30980   End Select
30990   Resume EXITP

End Sub

Private Sub NoChar_Load()

31000 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

31010   Set dbs = CurrentDb
31020   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
31030     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
31040     Set rst = qdf.OpenRecordset
31050     With rst
31060       .MoveLast
31070       lngNoChars = .RecordCount
31080       .MoveFirst
31090       arr_varNoChar = .GetRows(lngNoChars)
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
31100       .Close
31110     End With
31120     .Close
31130   End With

EXITP:
31140   Set rst = Nothing
31150   Set qdf = Nothing
31160   Set dbs = Nothing
31170   Exit Sub

ERRH:
31180   DoCmd.Hourglass False
31190   Select Case ERR.Number
        Case Else
31200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31210   End Select
31220   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

31300 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

31310   With Me
31320     blnGTR_Emblem = False
31330     For lngX = 1& To 24&
31340       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
31350     Next
31360   End With

EXITP:
31370   Exit Sub

ERRH:
31380   Select Case ERR.Number
        Case Else
31390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31400   End Select
31410   Resume EXITP

End Sub
