VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Split"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Split"

'VGC 09/09/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Select Account:   A {cmbAccounts}
' **   By Number:        B {opgAccountSource_optNumber}
' **   Cancel:           C {cmdCancel}
' **   Date:             D {transdate}
' **   Payable Date:     E {assetdate_display}
' **   Asset Cusip:      I {opgAssetSource_optCusip}
' **   All:              L {opgAccountNumber_optAll}
' **   By Name:          M {opgAccountSource_optName}
' **   Asset Name:       N {opgAssetSource_optName}
' **   OK:               O {cmdOK}
' **   Rate:             R {pershare}
' **   Select Asset:     S {cmbAssets}
' **   Spec. Account:    U {opgAccountNumber_optSpecified}
' **   Asset Type:       Y {opgAssetSource_optType}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** MapType = "Deposit", except...
' ** MapType = "Withdrawn" for Reverse Split

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of the Calendar Class.
Public clsMonthClass As clsMonthCal

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Combo box column constants: cmbAssets.
'Private Const CBX_A_ASTNO  As Integer = 0  'assetno
Private Const CBX_A_TDESC  As Integer = 1  'totdesc
Private Const CBX_A_CUSIP  As Integer = 2  'cusip
'Private Const CBX_A_DESC   As Integer = 3  'description
'Private Const CBX_A_TYPE   As Integer = 4  'assettype
Private Const CBX_A_CURRID As Integer = 5  'curr_id

' ** Combo box column constants: curr_id.
'Private Const CBX_C_CURRID As Integer = 0  'curr_id
'Private Const CBX_C_CODE   As Integer = 1  'curr_code
'Private Const CBX_C_NAME   As Integer = 2  'curr_name
'Private Const CBX_C_SYM    As Integer = 3  'currsym_symbol
'Private Const CBX_C_DEC    As Integer = 4  'curr_decimal
'Private Const CBX_C_RATE1  As Integer = 5  'curr_rate1
'Private Const CBX_C_RATE2  As Integer = 6  'curr_rate2
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

Private lngBkClr As Long
Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private lngTpp As Long, blnCurrID As Boolean, lngCurrID As Long
Private blnAssetUpdated As Boolean, blnAssetChecked As Boolean, blnDrivingMeBackwards As Boolean
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
210           gstrAccountNo = vbNullString

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
220           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

230           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

              ' ** Borrowing these variables from the Court Reports.
240           gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString

250           CLR_DISABLED_FG = CLR_DKGRY
260           CLR_DISABLED_BG = CLR_LTTEAL

270           .pershare_lbl3.Visible = False

              ' ** EVENT CHECK: chkRememberMe!
280           cmbAccounts_AfterUpdate  ' ** Procedure: Below.

290           opgAccountNumber_AfterUpdate  ' ** Procedure: Below.
300           opgAccountSource_AfterUpdate  ' ** Procedure: Below.
310           chkRememberMe_AfterUpdate  ' ** Procedure: Below.
320           opgAssetSource_AfterUpdate  ' ** Procedure: Below.

330         Else
340           Cancel = -1
350         End If
360       Else
370         Cancel = -1
380       End If
390     End With

400     If Cancel = -1 Then
410       DoCmd.Hourglass False
420       Beep
430       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
440     End If

EXITP:
450     Exit Sub

ERRH:
460     DoCmd.Hourglass False
470     THAT_PROC = THIS_PROC
480     That_Erl = Erl
490     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
500     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

610     With Me

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class.
620       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
630       clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
640       clsMonthClass.PositionAtCursor = True

650       Set dbs = CurrentDb
660       With dbs
670         Set grstPostingDate = .OpenRecordset("PostingDate", dbOpenDynaset, dbConsistent)
680         With grstPostingDate
690           .FindFirst "[Username] = '" & CurrentUser & "'"  ' ** Internal Access Function: Trust Accountant login.
700           If .NoMatch = False Then
710             Me.transdate.DefaultValue = "#" & CDate(Format(![Posting_Date], "mm/dd/yyyy")) & "#"
720           Else
730             .AddNew
740             ![Posting_Date] = Date
750             ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
760             .Update
770             Me.transdate.DefaultValue = Date
780           End If
790           .Close
800         End With
            ' ** Empty tmpAccount.
810         Set qdf = .QueryDefs("qryMap_Split_08_01")
820         qdf.Execute
830         Set qdf = Nothing
840         DoEvents
            ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
            ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
            ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
            ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
            ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
            ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
            ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
            ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
850         Set qdf = .QueryDefs("qryMap_Split_08_02")
860         qdf.Execute
870         Set qdf = Nothing
880         DoEvents
890         .Close
900       End With

910       .shareface.ForeColor = CLR_DISABLED_FG
920       .shareface.BackColor = CLR_DISABLED_BG

930       IncludeCurrency  ' ** Procedure: Below.

940     End With

950     DoCmd.Hourglass False

EXITP:
960     Set qdf = Nothing
970     Set dbs = Nothing
980     Exit Sub

ERRH:
990     DoCmd.Hourglass False
1000    Select Case ERR.Number
        Case Else
1010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1020    End Select
1030    Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   By Number:        B {opgAccountSource_optNumber}
        ' **   By Name:          M {opgAccountSource_optName}

        ' ** Plain keys.
1150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyEscape
1170        intRetVal = 0
1180        cmdCancel_Click  ' ** Procedure: Below.
1190      Case vbKeyF7
1200        intRetVal = 0
1210        cmdCalendar1_Click  ' ** Procedure: Below.
1220      Case vbKeyF8
1230        intRetVal = 0
1240        cmdCalendar2_Click  ' ** Procedure: Below.
1250      End Select
1260    End If

        ' ** Alt keys.
1270    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1280      Select Case intRetVal
          Case vbKeyB  ' ** Because the '_lbl' is empty.
1290        With Me
1300          intRetVal = 0
1310          If .opgAccountSource.Enabled = True Then
1320            .opgAccountSource.SetFocus
1330            .opgAccountSource = .opgAccountSource_optNumber.OptionValue
1340            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1350          Else
1360            Beep
1370          End If
1380        End With
1390      Case vbKeyM  ' ** Because the '_lbl' is empty.
1400        With Me
1410          intRetVal = 0
1420          If .opgAccountSource.Enabled = True Then
1430            .opgAccountSource.SetFocus
1440            .opgAccountSource = .opgAccountSource_optName.OptionValue
1450            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1460          Else
1470            Beep
1480          End If
1490        End With
1500      End Select
1510    End If

        ' ** Ctrl-Shift keys.
1520    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1530      Select Case intRetVal
          Case vbKeyF
1540        With Me
1550          intRetVal = 0
1560          .FocusHolder.SetFocus
1570        End With
1580      End Select
1590    End If

EXITP:
1600    KeyCode = intRetVal
1610    Exit Sub

ERRH:
1620    intRetVal = 0
1630    THAT_PROC = THIS_PROC
1640    That_Erl = Erl
1650    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1660    Resume EXITP

End Sub

Private Sub Form_Timer()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngAssetNo As Long, strAccountNo As String
        Dim lngTmp03 As Long

        Static lngTmp01 As Long, lngTmp02 As Long

1710    With Me
1720      .TimerInterval = 0
1730      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
1740        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
1750      End If
1760      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
1770        lngGTR_Stat = lngGTR_Stat + 1&
1780        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1790        lngTmp01 = lngTmp01 + 1&
1800        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1810        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1820        If lngTmp01 > 1& Then
1830          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1840        End If
1850        If lngTmp01 = 24& Then
1860          lngTmp01 = 0&
1870          lngTmp02 = lngTmp02 + 1&
1880          If lngTmp02 < 8& Then
1890            .TimerInterval = 50&
1900          Else
1910            blnGTR_Emblem = False
1920          End If
1930        Else
1940          .TimerInterval = 50&
1950        End If
1960      End If  ' ** blnGTR_Emblem.
1970      If blnExitingAssetNo = True Then
            ' ** This was set up this way so that the user doesn't have to hit Cancel twice
            ' ** if they happened to have entered an asset, but then changed their mind.
1980        blnExitingAssetNo = False
1990        If blnCancel = False Then
2000          lngAssetNo = .cmbAssets
2010          strAccountNo = NullStrIfNull(.cmbAccounts)  ' ** Module Function: modStringFuncs.
2020          Set dbs = CurrentDb
2030          Select Case strCallingForm
              Case "frmJournal"
2040            Select Case .opgAccountNumber
                Case .opgAccountNumber_optAll.OptionValue
                  ' ** ActiveAssets, grouped and summed, by assetno, by specified [astno].
2050              Set qdf = dbs.QueryDefs("qryMap_Split_07")
2060            Case .opgAccountNumber_optSpecified.OptionValue
                  ' ** ActiveAssets, grouped and summed, by assetno, by specified [astno], [actno].
2070              Set qdf = dbs.QueryDefs("qryMap_Split_07_01")
2080              With qdf.Parameters
2090                ![actno] = strAccountNo
2100              End With
2110            End Select
2120          Case "frmJournal_Columns"
2130            Select Case .opgAccountNumber
                Case .opgAccountNumber_optAll.OptionValue
                  ' ** ActiveAssets, grouped and summed, by assetno, by specified [astno].
2140              Set qdf = dbs.QueryDefs("qryJournal_Columns_34_03")
2150            Case .opgAccountNumber_optSpecified.OptionValue
                  ' ** ActiveAssets, grouped and summed, by assetno, by specified [astno], [actno].
2160              Set qdf = dbs.QueryDefs("qryJournal_Columns_34_03_01")
2170              With qdf.Parameters
2180                ![actno] = strAccountNo
2190              End With
2200            End Select
2210          End Select
2220          With qdf.Parameters
2230            ![astno] = lngAssetNo
2240          End With
2250          Set rst = qdf.OpenRecordset
              'Set rst = dbs.OpenRecordset("SELECT SUM(shareface) AS sumsf FROM activeAssets WHERE assetno = " & CStr(.cmbAssets) & ";")
2260          If rst.BOF = True And rst.EOF = True Then
                ' ** Shouldn't happen.
2270          Else
2280            blnAssetUpdated = False
2290            blnAssetChecked = True
2300            .TotalShares = rst![sumsf]
2310            .shareface = Round(rst![sumsf], 4)
2320  On Error Resume Next
2330            .pershare.SetFocus
2340  On Error GoTo ERRH
2350          End If
2360          rst.Close
2370          dbs.Close
2380        End If
2390      End If
2400      If gblnGoToReport = True Then
2410        DoCmd.Hourglass True  ' ** Make sure it's still running.
2420        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48& And lngGTR_Stat < 96&) Or _
                (lngGTR_Stat > 96& And lngGTR_Stat < 120&) Or (lngGTR_Stat > 120&)) Then
              ' ** Just let the emblem go.
2430        Else
2440          DoCmd.Hourglass True  ' ** Make sure it's still running.
2450          If blnGTR_Emblem = False Then
2460            blnGTR_Emblem = True
2470            lngTmp01 = 0&: lngTmp02 = 0&
2480            lngTmp03 = (.GoToReport_Emblem_01_img.Width + (8& * lngTpp))
2490            .GoToReport_arw_ok_img.Left = (.GoToReport_arw_ok_img.Left - lngTmp03)
2500            .cmdOK.Left = (.cmdOK.Left - lngTmp03)
2510            .cmdCancel.Left = (.cmdCancel.Left - lngTmp03)
2520          End If
2530          DoEvents
2540          Select Case blnGoingToReport
              Case True
2550            Select Case blnGoingToReport2
                Case True
2560              Select Case .GoToReport_arw_ok_img.Visible
                  Case True
2570                blnGoingToReport2 = False
2580                .TimerInterval = 50&
2590                cmdOK_Click  ' ** Procedure: Below.
2600              Case False
2610                .pershare.SetFocus
2620                .pershare = 66.666
2630                DoEvents
2640                .GoToReport_arw_ok_img.Visible = True
2650                .GoToReport_arw_map_per_img.Visible = False
2660                .assetdate = Now()
2670                .assetdate_display = Date
2680                .cmdOK.SetFocus
2690                DoEvents
2700                .TimerInterval = 50&
2710              End Select
2720            Case False
2730              blnGoingToReport2 = True
2740              lngAssetNo = GetDivAsset  ' ** Module Function: modGoToReportFuncs.
2750              If lngAssetNo > 0& Then
2760                .cmbAssets = lngAssetNo
2770                cmbAssets_AfterUpdate  ' ** Procedure: Below.
2780                DoEvents
2790                .GoToReport_arw_map_per_img.Visible = True
2800                .GoToReport_arw_map_asset_img.Visible = False
2810                .opgAssetSource.Visible = True
2820                .opgAssetSource_box.Visible = True
2830                .TotalShares = 4000
2840                .shareface = 4000
2850                .FocusHolder.SetFocus  ' ** So they don't see a flashing cursor.
2860                DoEvents
2870                .TimerInterval = 50&
2880              Else
2890                .TimerInterval = 0&
2900                Select Case garr_varGoToReport(GTR_FRM2)
                    Case "frmJournal"
2910                  Forms(garr_varGoToReport(GTR_FRM2)).GTREmblem_Set False  ' ** Form Procedure: frmJournal.
2920                Case "frmJournal_Columns"
2930                  Forms(garr_varGoToReport(GTR_FRM2)).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
2940                End Select
2950                gblnGoToReport = False
2960                blnGoingToReport = False
2970                blnGoingToReport2 = False
2980                GTREmblem_Off  ' ** Procedure: Below.
2990                .GoToReport_arw_map_asset_img.Visible = False
3000                .opgAssetSource.Visible = True
3010                .opgAssetSource_box.Visible = True
3020                Beep
3030                DoCmd.Hourglass False
3040                MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                      "There are insufficient asset holdings to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
3050              End If
3060            End Select
3070          Case False
3080            blnGoingToReport = True
3090            .GoToReport_arw_map_asset_img.Visible = True
3100            .opgAssetSource.Visible = False
3110            .opgAssetSource_box.Visible = False
3120            .FocusHolder.SetFocus  ' ** So they don't see a flashing cursor.
3130            DoEvents
3140            Select Case garr_varGoToReport(GTR_FRM2)
                Case "frmJournal"
3150              Forms(garr_varGoToReport(GTR_FRM2)).GTREmblem_Set False  ' ** Form Procedure: frmJournal.
3160            Case "frmJournal_Columns"
3170              Forms(garr_varGoToReport(GTR_FRM2)).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
3180            End Select
3190            .TimerInterval = 50&
3200          End Select
3210        End If
3220      End If
3230    End With  ' ** Me.

EXITP:
3240    Set rst = Nothing
3250    Set qdf = Nothing
3260    Set dbs = Nothing
3270    Exit Sub

ERRH:
3280    Select Case ERR.Number
        Case Else
3290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3300    End Select
3310    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim strControlName As String

3410    strControlName = vbNullString

3420    If DataErr <> 0 Then
3430      Select Case DataErr
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
3440      Case 2046  ' ** The command or action isn't available now (first or last record).
3450        DoCmd.CancelEvent
3460      Case 2113  ' ** 2113: The value you entered isn't valid for this field.
3470        Select Case Me.ActiveControl.Name
            Case "shareface"
3480          strControlName = "Shares"
3490        Case "pershare"
3500          strControlName = "Rate"
3510        Case "amount"
3520          strControlName = "Amount"
3530        Case "transdate"
3540          strControlName = vbNullString
3550          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3560          Response = acDataErrContinue
3570          DoCmd.CancelEvent
3580        Case Else
3590          strControlName = vbNullString
3600        End Select
3610        If strControlName <> vbNullString Then
3620          MsgBox "You must enter a numeric value in the " & strControlName & " field.", vbInformation + vbOKOnly, "Invalid Entry"
3630          Response = acDataErrContinue
3640          DoCmd.CancelEvent
3650        End If
3660      Case 2237  ' ** The text you entered isn't an item in the list.
3670        DoCmd.CancelEvent
3680      Case 2279  ' ** The value you entered isn't appropriate for the input mask '|' specified for this field.
3690        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3700        Response = acDataErrContinue
3710        DoCmd.CancelEvent
3720      Case 2580  ' ** The record source '|' specified on this form or report does not exist.
            ' ** Ignore! This form has no RecordSource.
3730        Response = acDataErrContinue
3740      Case 3162  ' ** You tried to assign the Null value to a variable that is not a Variant data type.
3750        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
            ' ** Do nothing.
3760      Case Else
3770        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3780        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3790      End Select
3800    End If

EXITP:
3810    Exit Sub

ERRH:
3820    Select Case ERR.Number
        Case Else
3830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3840    End Select
3850    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3910    With Me

3920      If .chkRememberMe = False Then
3930        .cmbAccounts = Null
3940      End If

3950      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3960      If Not clsMonthClass Is Nothing Then
3970        If clsMonthClass.IsCalendar Then
3980          Cancel = -1
3990        Else
4000          Set clsMonthClass = Nothing
4010        End If
4020      End If

4030    End With

EXITP:
4040    Exit Sub

ERRH:
4050    DoCmd.Hourglass False
4060    Select Case ERR.Number
        Case Else
4070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4080    End Select
4090    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4110    With Me
4120      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
4130        Select Case blnCalendar1_Focus
            Case True
4140          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
4150          .cmdCalendar1_raised_img.Visible = False
4160        Case False
4170          .cmdCalendar1_raised_img.Visible = True
4180          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
4190        End Select
4200        .cmdCalendar1_raised_focus_dots_img.Visible = False
4210        .cmdCalendar1_raised_focus_img.Visible = False
4220        .cmdCalendar1_sunken_focus_dots_img.Visible = False
4230        .cmdCalendar1_raised_img_dis.Visible = False
4240      End If
4250      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
4260        Select Case blnCalendar2_Focus
            Case True
4270          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
4280          .cmdCalendar2_raised_img.Visible = False
4290        Case False
4300          .cmdCalendar2_raised_img.Visible = True
4310          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
4320        End Select
4330        .cmdCalendar2_raised_focus_dots_img.Visible = False
4340        .cmdCalendar2_raised_focus_img.Visible = False
4350        .cmdCalendar2_sunken_focus_dots_img.Visible = False
4360        .cmdCalendar2_raised_img_dis.Visible = False
4370      End If
4380    End With

EXITP:
4390    Exit Sub

ERRH:
4400    Select Case ERR.Number
        Case Else
4410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4420    End Select
4430    Resume EXITP

End Sub

Private Sub cmdOK_Click()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim dblShareface As Double, dblPerShare As Double
        Dim datTransDate As Date, datAssetDate As Date, strMapType As String
        Dim lngAssetNo As Long, lngTaxcode As Long
        Dim strDocName As String
        Dim blnContinue As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, lngTmp02 As Long, lngTmp03 As Long, strTmp04 As String
        Dim lngX As Long

4510    blnContinue = True

4520    dblShareface = 0#
4530    dblPerShare = 0#
4540    lngAssetNo = 0&
4550    datAssetDate = 0
4560    datTransDate = 0
4570    lngTaxcode = 0&

4580    With Me

4590      Select Case .opgAccountNumber
          Case .opgAccountNumber_optAll.OptionValue
            ' ** Let it pass.
4600        gstrAccountNo = vbNullString
4610      Case .opgAccountNumber_optSpecified.OptionValue
4620        Select Case IsNull(.cmbAccounts)
            Case True
4630          gstrAccountNo = vbNullString
4640          blnContinue = False
4650          DoCmd.Hourglass False
4660          MsgBox "Please select a valid account.", vbInformation + vbOKOnly, "Entry Required"
4670          .cmbAccounts.SetFocus
4680        Case False
              ' ** Let it pass.
4690          gstrAccountNo = .cmbAccounts
4700        End Select
4710      End Select

4720      If blnContinue = True Then
4730        If IsNull(.cmbAssets) = True Then
4740          blnContinue = False
4750          MsgBox "Please select a valid asset.", vbInformation + vbOKOnly, "Entry Required"
4760          .cmbAssets.SetFocus
4770        Else
4780          If Trim(.cmbAssets) = vbNullString Then
4790            blnContinue = False
4800            MsgBox "Please select a valid asset.", vbInformation + vbOKOnly, "Entry Required"
4810            .cmbAssets.SetFocus
4820          Else
4830            lngAssetNo = .cmbAssets
4840            If IsNull(.pershare) Then
4850              blnContinue = False
4860              MsgBox "Rate must be greater than zero to continue.", vbInformation + vbOKOnly, "Entry Required"
4870              .pershare.SetFocus
4880            Else
4890              If Compare_DateA_DateB(CDate(.assetdate_display), ">", Date) = True Then  ' ** Module Function: modStringFuncs.
4900                blnContinue = False
4910                MsgBox "Future payable dates are not allowed.", vbInformation + vbOKOnly, "Invalid Date"
4920              Else
                    ' ** Allow posting up to 1 month into the future.
4930                If .transdate > DateAdd("m", 1, Date) Then
4940                  blnContinue = False
4950                  MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
4960                Else
4970                  If .pershare = 0 Then
4980                    blnContinue = False
4990                    MsgBox "Per Share must be greater than zero to continue.", vbInformation + vbOKOnly, "Entry Required"
5000                    If .pershare.Enabled = False Then
5010                      .pershare = Null
5020                      .pershare.Enabled = True
5030                      .pershare.Locked = False
5040                      .pershare.ForeColor = CLR_BLK
5050                      If lngCurrID = 150& Then
5060                        .pershare.BackColor = CLR_WHT
5070                      Else
5080                        .pershare.BackColor = lngBkClr  ' ** I know it's not currency.
5090                      End If
5100                    End If
5110                    .pershare.SetFocus
5120                  End If
5130                End If
5140              End If
5150            End If
5160          End If
5170        End If
5180      End If  ' ** blnContinue.

5190      If blnContinue = True Then

5200        DoCmd.Hourglass False
5210        DoEvents

5220        Set dbs = CurrentDb

            ' ** Empty Journal Map.
5230        Set qdf = dbs.QueryDefs("qryMap_01")
5240        qdf.Execute

5250        Select Case strCallingForm
            Case "frmJournal"
              ' ** MasterAsset, linked to AssetType, by specified [astno].
5260          Set qdf = dbs.QueryDefs("qryMap_Split_06")
5270        Case "frmJournal_Columns"
              ' ** MasterAsset, linked to AssetType, by specified [astno].
5280          Set qdf = dbs.QueryDefs("qryJournal_Columns_34_01")
5290        End Select
5300        With qdf.Parameters
5310          ![astno] = lngAssetNo
5320        End With
5330        Set rst = qdf.OpenRecordset
5340        If rst.BOF = True And rst.EOF = True Then
              ' ** Shouldn't happen.
5350        Else
5360          rst.MoveFirst
5370          lngTaxcode = rst![taxcode]
5380        End If
5390        .taxcode = lngTaxcode
5400        DoEvents

5410        dblShareface = .TotalShares
5420        dblPerShare = .pershare
5430        datAssetDate = .assetdate
5440        datTransDate = .transdate
5450        strMapType = .maptype  ' ** JournalType.
5460        DoEvents

5470        Select Case .opgAccountNumber
            Case .opgAccountNumber_optAll.OptionValue
              ' ** Check for closed accounts holding assets.
              ' ** qryMap_07 (qryMap_06 (ActiveAssets, linked to Account, just closed accounts, by specified [astno]),
              ' ** grouped by accountno, with cnt_astno), grouped and summed, with cnt_actno, cnt_astno.
5480          Set qdf = dbs.QueryDefs("qryMap_08")
5490          With qdf.Parameters
5500            ![astno] = Me.cmbAssets
5510          End With
5520          Set rst = qdf.OpenRecordset
5530          If rst.BOF = True And rst.EOF = True Then
                ' ** No closed accounts holding assets.
5540            msgResponse = vbYes
5550            rst.Close
5560          Else
5570            rst.MoveFirst
5580            If IsNull(rst![cnt_actno]) = True And IsNull(rst![cnt_astno]) = True Then
5590              msgResponse = vbYes
5600              rst.Close
5610            Else
5620              If rst![cnt_actno] = 0 And IsNull(rst![cnt_astno]) = True Then
5630                msgResponse = vbYes
5640                rst.Close
5650              ElseIf rst![cnt_actno] = 0 And rst![cnt_astno] = 0 Then
5660                msgResponse = vbYes
5670                rst.Close
5680              Else
5690                lngTmp02 = rst![cnt_actno]
5700                lngTmp03 = rst![cnt_astno]
5710                rst.Close
                    ' ** qryMap_06 (ActiveAssets, linked to Account, just closed accounts,
                    ' ** by specified [astno]), grouped by accountno, with cnt_astno.
5720                Set qdf = dbs.QueryDefs("qryMap_07")
5730                With qdf.Parameters
5740                  ![astno] = Me.cmbAssets
5750                End With
5760                Set rst = qdf.OpenRecordset
5770                rst.MoveFirst
5780                If lngTmp02 = 1& Then
5790                  strTmp04 = "There is 1 closed account currently holding this asset:" & vbCrLf & _
                        "    " & rst![accountno] & "  " & rst![shortname] & vbCrLf
5800                Else
5810                  strTmp04 = "There are " & CStr(lngTmp02) & " closed accounts currently holding this asset:" & vbCrLf
5820                  For lngX = 1& To lngTmp02
5830                    strTmp04 = strTmp04 & "    " & rst![accountno] & "  " & rst![shortname] & vbCrLf
5840                    If lngX < lngTmp02 Then rst.MoveNext
5850                  Next
5860                End If
5870                strTmp04 = strTmp04 & "If you choose to continue with this mapping," & vbCrLf & _
                      "the results may be inaccurate." & vbCrLf & vbCrLf & "Do you wish to continue?"
5880                DoCmd.Hourglass False
5890                Beep
5900                msgResponse = MsgBox(strTmp04, vbQuestion + vbYesNo + vbDefaultButton2, "Closed Accounts Holding Assets")
5910              End If
5920            End If
5930          End If
5940        Case .opgAccountNumber_optSpecified.OptionValue
5950          msgResponse = vbYes
5960        End Select
5970        DoEvents

5980        If msgResponse = vbYes Then

5990          DoCmd.Hourglass True
6000          DoEvents

6010          Select Case strCallingForm
              Case "frmJournal"
6020            Select Case .opgAccountNumber
                Case .opgAccountNumber_optAll.OptionValue
                  ' ** Append ActiveAssets to Journal Map, by specified [astno], [jrntyp], [persh], [transdat].
6030              Set qdf = dbs.QueryDefs("qryMap_Split_01")
6040            Case .opgAccountNumber_optSpecified.OptionValue
                  ' ** Append ActiveAssets to Journal Map, by specified [astno], [jrntyp], [persh], [transdat], [actno].
6050              Set qdf = dbs.QueryDefs("qryMap_Split_01a")
6060              With qdf.Parameters
6070                ![actno] = gstrAccountNo
6080              End With
6090            End Select
6100          Case "frmJournal_Columns"
6110            Select Case .opgAccountNumber
                Case .opgAccountNumber_optAll.OptionValue
                  ' ** Append ActiveAssets to Journal Map, by specified [astno], [jrntyp], [persh], [transdat].
6120              Set qdf = dbs.QueryDefs("qryJournal_Columns_34_02")
6130            Case .opgAccountNumber_optSpecified.OptionValue
                  ' ** Append ActiveAssets to Journal Map, by specified [astno], [jrntyp], [persh], [transdat], [actno].
6140              Set qdf = dbs.QueryDefs("qryJournal_Columns_34_02a")
6150              With qdf.Parameters
6160                ![actno] = gstrAccountNo
6170              End With
6180            End Select
6190          End Select
6200          With qdf.Parameters
6210            ![astno] = lngAssetNo
6220            ![jrntyp] = strMapType
6230            ![persh] = dblPerShare  'CCur(dblPerShare)  ' ** Used for icash.
6240            ![transdat] = datTransDate
6250          End With
6260          qdf.Execute
6270          Set qdf = Nothing
6280          DoEvents

              ' ** Journal Map, just zero shareface, icash, pcash.
6290          Set qdf = dbs.QueryDefs("qryMap_Div_01a")
6300          Set rst = qdf.OpenRecordset
6310          If rst.BOF = True And rst.EOF = True Then
                ' ** All's well.
6320            rst.Close
6330          Else
6340            rst.Close
                ' ** Delete qryMap_Div_01a (Journal Map, just zero shareface, icash, pcash).
6350            Set qdf = dbs.QueryDefs("qryMap_Div_01b")
6360            qdf.Execute
6370          End If
6380          dbs.Close
6390          Set rst = Nothing
6400          Set qdf = Nothing
6410          Set dbs = Nothing

6420          varTmp00 = .cmbAssets.Column(CBX_A_CUSIP)
6430          Select Case IsNull(varTmp00)
              Case True
                ' ** Might be coming from GoToReport.
6440            blnContinue = False
6450          Case False
                ' ** Proceed.
6460          End Select

6470          If blnContinue = True Then

                ' ** Borrowing these variables from the Court Reports.
6480            gstrCrtRpt_Ordinal = .cmbAssets.Column(CBX_A_CUSIP)  ' ** cusip.
6490            gstrCrtRpt_Version = .cmbAssets.Column(CBX_A_TDESC)  ' ** totdesc.

6500            strDocName = "frmMap_Split_Detail"
6510            DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & CStr(dblPerShare) & "~" & _
                  Format(datTransDate, "mm/dd/yyyy") & "~" & Format(datAssetDate, "mm/dd/yyyy hh:nn:ss") & "~" & _
                  CStr(lngCurrID) & "~" & .pershare_lbl3.Caption

6520            If gblnGoToReport = True Then
6530              DoCmd.Hourglass True  ' ** Make sure it's still running.
6540              DoEvents
6550              .GoToReport_arw_ok_img.Visible = False
6560              Forms(strDocName).cmdCancel.SetFocus
6570              Forms(strDocName).cmdOK.Enabled = False  ' ** Only allow Cancel on this fake split.
6580              DoEvents
6590              Forms(strDocName).TimerInterval = 50&
6600              blnGoingToReport = False
6610            End If

6620            DoCmd.Close acForm, THIS_NAME

6630          Else
6640            Select Case gblnGoToReport
                Case True
6650              Beep
6660              DoCmd.Hourglass False
6670              MsgBox "The report site cannot be shown because there are no records.", _
                    vbInformation + vbOKOnly, "Record Not Found"
6680              blnGoingToReport = False
6690              .GoToReport_arw_ok_img.Visible = False
6700              gblnGoToReport = False
6710            Case False
6720              Beep
6730              DoCmd.Hourglass False
6740              MsgBox "This procedure cannot continue." & vbCrLf & "Expected records were not found.", _
                    vbInformation + vbOKOnly, "Record Not Found"
6750              blnGoingToReport = False
6760              .GoToReport_arw_ok_img.Visible = False
6770              gblnGoToReport = False
6780            End Select
6790            .cmdCancel.SetFocus
6800          End If  ' ** blnContinue.

6810        Else
6820          .cmdCancel.SetFocus
6830        End If  ' ** msgResponse.

6840      End If

6850    End With

EXITP:
6860    Set rst = Nothing
6870    Set qdf = Nothing
6880    Set dbs = Nothing
6890    Exit Sub

ERRH:
6900    DoCmd.Hourglass False
6910    THAT_PROC = THIS_PROC
6920    That_Erl = Erl
6930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6940    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7020    intShiftDown = (Shift And acShiftMask) > 0
7030    intAltDown = (Shift And acAltMask) > 0
7040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7060      Select Case intRetVal
          Case vbKeyTab
7070        With Me
7080          intRetVal = 0
7090          .cmdCancel.SetFocus
7100        End With
7110      End Select
7120    End If

        ' ** Shift keys.
7130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7140      Select Case intRetVal
          Case vbKeyTab
7150        With Me
7160          intRetVal = 0
7170          .transdate.SetFocus
7180        End With
7190      End Select
7200    End If

        ' ** Ctrl-Shift keys.
7210    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7220      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7230        With Me
7240          intRetVal = 0
7250          .cmdCalendar1.SetFocus
7260        End With
7270      End Select
7280    End If

EXITP:
7290    KeyCode = intRetVal
7300    Exit Sub

ERRH:
7310    intRetVal = 0
7320    THAT_PROC = THIS_PROC
7330    That_Erl = Erl
7340    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7350    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

7410    DoCmd.Hourglass True
7420    DoEvents

7430    blnCancel = True

7440    Set dbs = CurrentDb
7450    With dbs
7460      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
7470        Set qdf = .QueryDefs("qryMap_01")
7480      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
7490        Set qdf = .QueryDefs("qryJournal_Columns_34_04")
7500      End Select
7510      qdf.Execute
7520      Set qdf = Nothing
7530      DoEvents
          ' ** Empty tmpAccount.
7540      Set qdf = .QueryDefs("qryMap_Split_08_01")
7550      qdf.Execute
7560      Set qdf = Nothing
7570      DoEvents
7580      .Close
7590    End With

7600    gblnSetFocus = True
7610    Select Case strCallingForm
        Case "frmJournal"
7620      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
7630        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
7640      Else
7650        gstrReturningForm = THIS_NAME
7660        DoCmd.SelectObject acForm, strCallingForm, False
7670        Forms(strCallingForm).TimerInterval = 100&
7680      End If
7690    Case "frmJournal_Columns"
7700      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
7710        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
7720      Else
7730        gstrReturningForm = THIS_NAME
7740        DoCmd.SelectObject acForm, strCallingForm, False
7750        Forms(strCallingForm).TimerInterval = 250&
7760      End If
7770    End Select

7780    DoCmd.Close acForm, THIS_NAME

EXITP:
7790    Set qdf = Nothing
7800    Set dbs = Nothing
7810    Exit Sub

ERRH:
7820    DoCmd.Hourglass False
7830    THAT_PROC = THIS_PROC
7840    That_Erl = Erl
7850    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7860    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7920    intShiftDown = (Shift And acShiftMask) > 0
7930    intAltDown = (Shift And acAltMask) > 0
7940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7960      Select Case intRetVal
          Case vbKeyTab
7970        With Me
7980          intRetVal = 0
7990          .opgAccountNumber.SetFocus
8000        End With
8010      End Select
8020    End If

        ' ** Shift keys.
8030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8040      Select Case intRetVal
          Case vbKeyTab
8050        With Me
8060          intRetVal = 0
8070          .cmdOK.SetFocus
8080        End With
8090      End Select
8100    End If

EXITP:
8110    KeyCode = intRetVal
8120    Exit Sub

ERRH:
8130    intRetVal = 0
8140    THAT_PROC = THIS_PROC
8150    That_Erl = Erl
8160    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8170    Resume EXITP

End Sub

Private Sub opgAccountNumber_AfterUpdate()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_AfterUpdate"

8210    With Me
8220      .cmbAssets = Null
8230      .cmbAccounts = Null
8240      .shareface = 0#
8250      .pershare = Null
8260      .assetdate = Null
8270      .assetdate_display = Null
8280      .pershare_usd = Null
8290      Select Case .opgAccountNumber
          Case .opgAccountNumber_optAll.OptionValue
8300        .opgAccountNumber_optAll_lbl.FontBold = True
8310        .opgAccountNumber_optSpecified_lbl.FontBold = False
8320        .opgAccountNumber_optAll_lbl_box.Visible = True
8330        .opgAccountNumber_optSpecified_lbl_box.Visible = False
8340        .cmbAccounts.Enabled = False
8350        .cmbAccounts.BorderColor = WIN_CLR_DISR
8360        .cmbAccounts.BackStyle = acBackStyleTransparent
8370        .cmbAccounts_lbl_box.Visible = True
8380        .opgAccountSource.Enabled = False
8390        .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
8400        .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
8410        .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
8420        .opgAccountSource_optName_lbl2_dim_hi.Visible = True
8430        .chkRememberMe.Enabled = False
8440        .chkRememberMe_lbl.Visible = False
8450        .chkRememberMe_lbl2_dim.Visible = True
8460        .chkRememberMe_lbl2_dim_hi.Visible = True
8470        DoEvents
            ' ** cmbAssets RowSource's:
            ' **   qryMap_Rec_05a  .opgAssetSource_optType.OptionValue
            ' **   qryMap_Rec_05c    GlobalVarGet('gstrAccountNo')
            ' **   qryMap_Rec_05d  .opgAssetSource_optName.OptionValue
            ' **   qryMap_Rec_05f    GlobalVarGet('gstrAccountNo')
            ' **   qryMap_Rec_05g  .opgAssetSource_optCusip.OptionValue
            ' **   qryMap_Rec_05i    GlobalVarGet('gstrAccountNo')
8480        Select Case .opgAssetSource
            Case .opgAssetSource_optType.OptionValue
8490          If .cmbAssets.RowSource <> "qryMap_Rec_05a" Then
8500            .cmbAssets.RowSource = "qryMap_Rec_05a"
8510          End If
8520        Case .opgAssetSource_optName.OptionValue
8530          If .cmbAssets.RowSource <> "qryMap_Rec_05d" Then
8540            .cmbAssets.RowSource = "qryMap_Rec_05d"
8550          End If
8560        Case .opgAssetSource_optCusip.OptionValue
8570          If .cmbAssets.RowSource <> "qryMap_Rec_05g" Then
8580            .cmbAssets.RowSource = "qryMap_Rec_05g"
8590          End If
8600        End Select
8610        .cmbAssets.Requery
8620        DoEvents
8630        .cmbAssets.SetFocus
8640      Case .opgAccountNumber_optSpecified.OptionValue
            ' ** cmbAssets can remain as-is.
8650        .opgAccountNumber_optAll_lbl.FontBold = False
8660        .opgAccountNumber_optSpecified_lbl.FontBold = True
8670        .opgAccountNumber_optAll_lbl_box.Visible = False
8680        .opgAccountNumber_optSpecified_lbl_box.Visible = True
8690        .cmbAccounts.Enabled = True
8700        .cmbAccounts.BorderColor = CLR_LTBLU2
8710        .cmbAccounts.BackStyle = acBackStyleNormal
8720        .cmbAccounts_lbl_box.Visible = False
8730        .opgAccountSource.Enabled = True
8740        .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
8750        .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
8760        .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
8770        .opgAccountSource_optName_lbl2_dim_hi.Visible = False
8780        .chkRememberMe.Enabled = True
8790        .chkRememberMe_lbl.Visible = True
8800        .chkRememberMe_lbl2_dim.Visible = False
8810        .chkRememberMe_lbl2_dim_hi.Visible = False
8820        .cmbAccounts.Requery
8830        DoEvents
8840        .cmbAccounts.SetFocus
8850      End Select
8860    End With

EXITP:
8870    Exit Sub

ERRH:
8880    Select Case ERR.Number
        Case Else
8890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8900    End Select
8910    Resume EXITP

End Sub

Private Sub opgAccountNumber_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9020    intShiftDown = (Shift And acShiftMask) > 0
9030    intAltDown = (Shift And acAltMask) > 0
9040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9070        With Me
9080          intRetVal = 0
9090          Select Case .opgAccountNumber
              Case .opgAccountNumber_optAll.OptionValue
9100            .cmbAssets.SetFocus
9110          Case .opgAccountNumber_optSpecified.OptionValue
9120            .cmbAccounts.SetFocus
9130          End Select
9140        End With
9150      End Select
9160    End If

        ' ** Shift keys.
9170    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9180      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9190        With Me
9200          intRetVal = 0
9210          .cmdCancel.SetFocus
9220        End With
9230      End Select
9240    End If

EXITP:
9250    KeyCode = intRetVal
9260    Exit Sub

ERRH:
9270    intRetVal = 0
9280    THAT_PROC = THIS_PROC
9290    That_Erl = Erl
9300    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9310    Resume EXITP

End Sub

Private Sub opgAccountNumber_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optSpecified_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9420    intShiftDown = (Shift And acShiftMask) > 0
9430    intAltDown = (Shift And acAltMask) > 0
9440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9470        With Me
9480          intRetVal = 0
9490          Select Case .opgAccountNumber
              Case .opgAccountNumber_optAll.OptionValue
9500            .cmbAssets.SetFocus
9510          Case .opgAccountNumber_optSpecified.OptionValue
9520            .cmbAccounts.SetFocus
9530          End Select
9540        End With
9550      End Select
9560    End If

        ' ** Shift keys.
9570    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9580      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9590        With Me
9600          intRetVal = 0
9610          .cmdCancel.SetFocus
9620        End With
9630      End Select
9640    End If

EXITP:
9650    KeyCode = intRetVal
9660    Exit Sub

ERRH:
9670    intRetVal = 0
9680    THAT_PROC = THIS_PROC
9690    That_Erl = Erl
9700    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9710    Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

9810    With Me
9820      Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
9830        .opgAccountSource_optNumber_lbl2.FontBold = True
9840        .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = True
9850        .opgAccountSource_optName_lbl2.FontBold = False
9860        .opgAccountSource_optName_lbl2_dim_hi.FontBold = False
9870        If .cmbAccounts.RowSource <> "qryMap_Split_08a" Then
9880          .cmbAccounts.RowSource = "qryMap_Split_08a"
9890        End If
9900      Case .opgAccountSource_optName.OptionValue
9910        .opgAccountSource_optNumber_lbl2.FontBold = False
9920        .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = False
9930        .opgAccountSource_optName_lbl2.FontBold = True
9940        .opgAccountSource_optName_lbl2_dim_hi.FontBold = True
9950        If .cmbAccounts.RowSource <> "qryMap_Split_08b" Then
9960          .cmbAccounts.RowSource = "qryMap_Split_08b"
9970        End If
9980      End Select
9990      .cmbAccounts.Requery
10000   End With

EXITP:
10010   Exit Sub

ERRH:
10020   Select Case ERR.Number
        Case Else
10030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10040   End Select
10050   Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10120   intShiftDown = (Shift And acShiftMask) > 0
10130   intAltDown = (Shift And acAltMask) > 0
10140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10170       With Me
10180         intRetVal = 0
10190         .chkRememberMe.SetFocus
10200       End With
10210     End Select
10220   End If

        ' ** Shift keys.
10230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10250       With Me
10260         intRetVal = 0
10270         .opgAccountNumber.SetFocus
10280       End With
10290     End Select
10300   End If

EXITP:
10310   KeyCode = intRetVal
10320   Exit Sub

ERRH:
10330   intRetVal = 0
10340   THAT_PROC = THIS_PROC
10350   That_Erl = Erl
10360   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10370   Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10420   intShiftDown = (Shift And acShiftMask) > 0
10430   intAltDown = (Shift And acAltMask) > 0
10440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10470       With Me
10480         intRetVal = 0
10490         .chkRememberMe.SetFocus
10500       End With
10510     End Select
10520   End If

        ' ** Shift keys.
10530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10550       With Me
10560         intRetVal = 0
10570         .opgAccountNumber.SetFocus
10580       End With
10590     End Select
10600   End If

EXITP:
10610   KeyCode = intRetVal
10620   Exit Sub

ERRH:
10630   intRetVal = 0
10640   THAT_PROC = THIS_PROC
10650   That_Erl = Erl
10660   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10670   Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

10710   With Me
10720     Select Case .chkRememberMe
          Case True
10730       .chkRememberMe_lbl.FontBold = True
10740       .chkRememberMe_lbl2_dim.FontBold = True
10750       .chkRememberMe_lbl2_dim_hi.FontBold = True
10760     Case False
10770       .chkRememberMe_lbl.FontBold = False
10780       .chkRememberMe_lbl2_dim.FontBold = False
10790       .chkRememberMe_lbl2_dim_hi.FontBold = False
10800     End Select
10810   End With

EXITP:
10820   Exit Sub

ERRH:
10830   Select Case ERR.Number
        Case Else
10840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10850   End Select
10860   Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10920   intShiftDown = (Shift And acShiftMask) > 0
10930   intAltDown = (Shift And acAltMask) > 0
10940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10970       With Me
10980         intRetVal = 0
10990         .cmbAccounts.SetFocus
11000       End With
11010     End Select
11020   End If

        ' ** Shift keys.
11030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11050       With Me
11060         intRetVal = 0
11070         .opgAccountSource.SetFocus
11080       End With
11090     End Select
11100   End If

EXITP:
11110   KeyCode = intRetVal
11120   Exit Sub

ERRH:
11130   intRetVal = 0
11140   THAT_PROC = THIS_PROC
11150   That_Erl = Erl
11160   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11170   Resume EXITP

End Sub

Private Sub cmbAccounts_AfterUpdate()

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

11210   With Me
11220     If IsNull(.cmbAccounts) = False Then
11230       gstrAccountNo = .cmbAccounts
11240       Select Case .opgAssetSource
            Case .opgAssetSource_optType.OptionValue
              ' ** qryMap_Split_05a (qryMap_Split_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** for Dividend = True, shareface <> 0; by assettype), linked to qryMap_Split_05b (ActiveAssets,
              ' ** grouped by accountno, assetno), by specified GlobalVarGet('gstrAccountNo'); by assettype.
11250         .cmbAssets.RowSource = "qryMap_Split_05c"
11260         .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
11270         .cmbAssets.ListWidth = 7890
11280       Case .opgAssetSource_optName.OptionValue
              ' ** qryMap_Split_05d (qryMap_Split_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** for Dividend = True, shareface <> 0; by totdesc), linked to qryMap_Split_05e (ActiveAssets,
              ' ** grouped by accountno, assetno), by specified GlobalVarGet('gstrAccountNo'); by totdesc.
11290         .cmbAssets.RowSource = "qryMap_Split_05f"
11300         .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
11310         .cmbAssets.ListWidth = 7890
11320       Case .opgAssetSource_optCusip.OptionValue
              ' ** qryMap_Split_05g (qryMap_Split_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** with qryMap_Split_05k (qryMap_Split_05j (MasterAsset, with cusip_len), grouped, with Max(cusip_len)),
              ' ** for Dividend = True, shareface <> 0; by cusip; Cartesian), linked to qryMap_Split_05h (ActiveAssets,
              ' ** grouped by accountno, assetno), by specified GlobalVarGet('gstrAccountNo'); by cusip.
11330         .cmbAssets.RowSource = "qryMap_Split_05i"
11340         .cmbAssets.ColumnWidths = "0;7275;0;0;360;0"
11350         .cmbAssets.ListWidth = 7890
11360       End Select
11370     Else
11380       gstrAccountNo = vbNullString
11390       Select Case .opgAssetSource
            Case .opgAssetSource_optType.OptionValue
              ' ** qryMap_Split_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** for Dividend = True, shareface <> 0; by assettype.
11400         .cmbAssets.RowSource = "qryMap_Split_05a"
11410         .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
11420         .cmbAssets.ListWidth = 7890
11430       Case .opgAssetSource_optName.OptionValue
              ' ** qryMap_Split_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** for Dividend = True, shareface <> 0; by totdesc.
11440         .cmbAssets.RowSource = "qryMap_Split_05d"
11450         .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
11460         .cmbAssets.ListWidth = 7890
11470       Case .opgAssetSource_optCusip.OptionValue
              ' ** qryMap_Split_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** with qryMap_Split_05k (qryMap_Split_05j (MasterAsset, with cusip_len), grouped,
              ' ** with Max(cusip_len)), for Dividend = True, shareface <> 0; by cusip; Cartesian.
11480         .cmbAssets.RowSource = "qryMap_Split_05g"
11490         .cmbAssets.ColumnWidths = "0;7275;0;0;360;0"
11500         .cmbAssets.ListWidth = 7890
11510       End Select
11520     End If
11530     .cmbAssets.Requery
11540   End With

EXITP:
11550   Exit Sub

ERRH:
11560   Select Case ERR.Number
        Case Else
11570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11580   End Select
11590   Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11620   intShiftDown = (Shift And acShiftMask) > 0
11630   intAltDown = (Shift And acAltMask) > 0
11640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11670       With Me
11680         intRetVal = 0
11690         .cmbAssets.SetFocus
11700       End With
11710     End Select
11720   End If

        ' ** Shift keys.
11730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11750       With Me
11760         intRetVal = 0
11770         .opgAccountNumber.SetFocus
11780       End With
11790     End Select
11800   End If

        ' ** Ctrl keys.
11810   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
11820     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11830       With Me
11840         intRetVal = 0
11850         .opgAccountSource.SetFocus
11860       End With
11870     End Select
11880   End If

EXITP:
11890   KeyCode = intRetVal
11900   Exit Sub

ERRH:
11910   intRetVal = 0
11920   THAT_PROC = THIS_PROC
11930   That_Erl = Erl
11940   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11950   Resume EXITP

End Sub

Private Sub opgAssetSource_AfterUpdate()

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_AfterUpdate"

12010   With Me
12020     Select Case .opgAssetSource
          Case .opgAssetSource_optType.OptionValue
12030       .opgAssetSource_optType_lbl.FontBold = True
12040       .opgAssetSource_optName_lbl.FontBold = False
12050       .opgAssetSource_optCusip_lbl.FontBold = False
12060       Select Case .opgAccountNumber
            Case .opgAccountNumber_optAll.OptionValue
              ' ** qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to
              ' ** MasterAsset, for Dividend = True, shareface <> 0; by assettype.
12070         If .cmbAssets.RowSource <> "qryMap_Rec_05a" Then
12080           .cmbAssets.RowSource = "qryMap_Rec_05a"
12090         End If
12100       Case .opgAccountNumber_optSpecified.OptionValue
              ' ** qryMap_Rec_05a (qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** for Dividend = True, shareface <> 0; by assettype), linked to qryMap_Rec_05b (ActiveAssets,
              ' ** grouped by accountno, assetno), by specified GlobalVarGet('gstrAccountNo'); by assettype.
12110         If .cmbAssets.RowSource <> "qryMap_Rec_05c" Then
12120           .cmbAssets.RowSource = "qryMap_Rec_05c"
12130         End If
12140       End Select
12150       .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
12160       .cmbAssets.ListWidth = 7890
12170     Case .opgAssetSource_optName.OptionValue
12180       .opgAssetSource_optType_lbl.FontBold = False
12190       .opgAssetSource_optName_lbl.FontBold = True
12200       .opgAssetSource_optCusip_lbl.FontBold = False
12210       Select Case .opgAccountNumber
            Case .opgAccountNumber_optAll.OptionValue
              ' ** qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to
              ' ** MasterAsset, for Dividend = True, shareface <> 0; by totdesc.
12220         If .cmbAssets.RowSource <> "qryMap_Rec_05d" Then
12230           .cmbAssets.RowSource = "qryMap_Rec_05d"
12240         End If
12250       Case .opgAccountNumber_optSpecified.OptionValue
              ' ** qryMap_Rec_05d (qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** for Dividend = True, shareface <> 0; by totdesc), linked to qryMap_Rec_05e (ActiveAssets,
              ' ** grouped by accountno, assetno), by specified GlobalVarGet('gstrAccountNo'); by totdesc.
12260         If .cmbAssets.RowSource <> "qryMap_Rec_05f" Then
12270           .cmbAssets.RowSource = "qryMap_Rec_05f"
12280         End If
12290       End Select
12300       .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
12310       .cmbAssets.ListWidth = 7890
12320     Case .opgAssetSource_optCusip.OptionValue
12330       .opgAssetSource_optType_lbl.FontBold = False
12340       .opgAssetSource_optName_lbl.FontBold = False
12350       .opgAssetSource_optCusip_lbl.FontBold = True
12360       Select Case .opgAccountNumber
            Case .opgAccountNumber_optAll.OptionValue
              ' ** qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to
              ' ** MasterAsset, with qryMap_Rec_05k (qryMap_Rec_05j (MasterAsset,
              ' ** with cusip_len), grouped, with Max(cusip_len)), for
              ' ** Dividend = True, shareface <> 0; by cusip; Cartesian.
12370         If .cmbAssets.RowSource <> "qryMap_Rec_05g" Then
12380           .cmbAssets.RowSource = "qryMap_Rec_05g"
12390         End If
12400       Case .opgAccountNumber_optSpecified.OptionValue
              ' ** qryMap_Rec_05g (qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to MasterAsset,
              ' ** with qryMap_Rec_05k (qryMap_Rec_05j (MasterAsset, with cusip_len), grouped, with Max(cusip_len)),
              ' ** for Dividend = True, shareface <> 0; by cusip; Cartesian), linked to qryMap_Rec_05h (ActiveAssets,
              ' ** grouped by accountno, assetno), by specified GlobalVarGet('gstrAccountNo'); by cusip.
12410         If .cmbAssets.RowSource <> "qryMap_Rec_05i" Then
12420           .cmbAssets.RowSource = "qryMap_Rec_05i"
12430         End If
12440       End Select
12450       .cmbAssets.ColumnWidths = "0;7275;0;0;360;0"
12460       .cmbAssets.ListWidth = 7890
12470     End Select
12480     .cmbAssets.Requery
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

Private Sub opgAssetSource_optType_KeyDown(KeyCode As Integer, Shift As Integer)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12620   intShiftDown = (Shift And acShiftMask) > 0
12630   intAltDown = (Shift And acAltMask) > 0
12640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12670       With Me
12680         intRetVal = 0
12690         .cmbAssets.SetFocus
12700       End With
12710     End Select
12720   End If

        ' ** Shift keys.
12730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12750       With Me
12760         intRetVal = 0
12770         .opgAccountNumber.SetFocus
12780       End With
12790     End Select
12800   End If

EXITP:
12810   KeyCode = intRetVal
12820   Exit Sub

ERRH:
12830   intRetVal = 0
12840   Select Case ERR.Number
        Case Else
12850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12860   End Select
12870   Resume EXITP

End Sub

Private Sub opgAssetSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12920   intShiftDown = (Shift And acShiftMask) > 0
12930   intAltDown = (Shift And acAltMask) > 0
12940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12970       With Me
12980         intRetVal = 0
12990         .cmbAssets.SetFocus
13000       End With
13010     End Select
13020   End If

        ' ** Shift keys.
13030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13050       With Me
13060         intRetVal = 0
13070         .opgAccountNumber.SetFocus
13080       End With
13090     End Select
13100   End If

EXITP:
13110   KeyCode = intRetVal
13120   Exit Sub

ERRH:
13130   intRetVal = 0
13140   Select Case ERR.Number
        Case Else
13150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13160   End Select
13170   Resume EXITP

End Sub

Private Sub opgAssetSource_optCusip_KeyDown(KeyCode As Integer, Shift As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optCusip_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
13270       With Me
13280         intRetVal = 0
13290         .cmbAssets.SetFocus
13300       End With
13310     End Select
13320   End If

        ' ** Shift keys.
13330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13350       With Me
13360         intRetVal = 0
13370         .opgAccountNumber.SetFocus
13380       End With
13390     End Select
13400   End If

EXITP:
13410   KeyCode = intRetVal
13420   Exit Sub

ERRH:
13430   intRetVal = 0
13440   Select Case ERR.Number
        Case Else
13450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13460   End Select
13470   Resume EXITP

End Sub

Private Sub cmbAssets_AfterUpdate()

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_AfterUpdate"

13510   With Me

13520     If blnCurrID = True Then
13530       lngCurrID = 0&
13540       If IsNull(.cmbAssets) = False Then
13550         If IsNull(.cmbAssets.Column(CBX_A_CURRID)) = False Then
13560           lngCurrID = .cmbAssets.Column(CBX_A_CURRID)
13570           .curr_id = lngCurrID
13580         End If
13590       End If
13600     End If

13610     blnAssetUpdated = True: blnAssetChecked = False
13620     .pershare.SetFocus

13630     curr_id_AfterUpdate  ' ** Procedure: Below.

13640   End With

EXITP:
13650   Exit Sub

ERRH:
13660   THAT_PROC = THIS_PROC
13670   That_Erl = Erl
13680   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13690   Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13720   intShiftDown = (Shift And acShiftMask) > 0
13730   intAltDown = (Shift And acAltMask) > 0
13740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13770       With Me
13780         intRetVal = 0
13790         .pershare.SetFocus
13800       End With
13810     End Select
13820   End If

        ' ** Shift keys.
13830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13850       With Me
13860         intRetVal = 0
13870         If .cmbAccounts.Enabled = True Then
13880           .cmbAccounts.SetFocus
13890         Else
13900           .opgAccountNumber.SetFocus
13910         End If
13920       End With
13930     End Select
13940   End If

        ' ** Ctrl keys.
13950   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
13960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13970       With Me
13980         intRetVal = 0
13990         .opgAssetSource.SetFocus
14000       End With
14010     End Select
14020   End If

EXITP:
14030   KeyCode = intRetVal
14040   Exit Sub

ERRH:
14050   intRetVal = 0
14060   THAT_PROC = THIS_PROC
14070   That_Erl = Erl
14080   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14090   Resume EXITP

End Sub

Private Sub cmbAssets_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_MouseMove"

14110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
14120   Exit Sub

ERRH:
14130   Select Case ERR.Number
        Case Else
14140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14150   End Select
14160   Resume EXITP

End Sub

Private Sub cmbAssets_LostFocus()

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_LostFocus"

14210   With Me
14220     If IsNull(.cmbAssets) = False Then
14230       If .cmbAssets > 0 And blnAssetUpdated = True And blnAssetChecked = False Then
14240         blnExitingAssetNo = True
14250         .TimerInterval = 100&
14260       End If
14270     End If
14280   End With  ' ** Me.

EXITP:
14290   Exit Sub

ERRH:
14300   THAT_PROC = THIS_PROC
14310   That_Erl = Erl
14320   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14330   Resume EXITP

End Sub

Private Sub curr_id_cmd_DblClick(Cancel As Integer)

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_cmd_DblClick"

14410   With Me
14420     If IsNull(.cmbAssets) = False Then
14430       Select Case .curr_date.Visible
            Case True
14440         .curr_date.Visible = False
14450       Case False
14460         .curr_date = .curr_id.Column(CBX_C_DATE)
14470         .curr_date.Visible = True
14480       End Select
14490     End If
14500   End With

EXITP:
14510   Exit Sub

ERRH:
14520   Select Case ERR.Number
        Case Else
14530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
14540   End Select
14550   Resume EXITP

End Sub

Private Sub curr_id_AfterUpdate()

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_AfterUpdate"

        Dim blnFound As Boolean
        Dim lngX As Long

14610   With Me

14620     If IsNull(.curr_id) = False Then

14630       lngCurrID = .curr_id

14640       blnFound = False
14650       If lngCurrID <> 150& Then
14660         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
14670           NoChar_Load  ' ** Procedure: Below.
14680         End If
14690         For lngX = 0& To (lngNoChars - 1&)
14700           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
14710             blnFound = True
14720           End If
14730         Next
14740       End If

14750       Select Case blnFound  ' ** I know this isn't currency.
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
14760         .pershare.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
14770         .pershare.FontSize = 10
14780         .pershare.FontBold = True
14790       Case False
14800         If .pershare.FontName <> "Arial" Then
14810           .pershare.FontName = "Arial"
14820           .pershare.FontSize = 10
14830           .pershare.FontBold = False
14840         End If
14850       End Select

14860       If lngCurrID = 150& Then  ' ** USD.
14870         lngBkClr = CLR_WHT
14880       Else
14890         lngBkClr = CLR_VLTGRN
14900       End If
14910       .pershare.BackColor = lngBkClr

14920     End If

14930   End With

EXITP:
14940   Exit Sub

ERRH:
14950   THAT_PROC = THIS_PROC
14960   That_Erl = Erl
14970   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14980   Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

15010   With Me
15020     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
15030     .curr_id.SetFocus
15040     Response = acDataErrContinue
15050   End With

EXITP:
15060   Exit Sub

ERRH:
15070   THAT_PROC = THIS_PROC
15080   That_Erl = Erl
15090   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15100   Resume EXITP

End Sub

Private Sub pershare_Enter()

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "pershare_Enter"

15210   With Me
15220     .pershare.BackColor = CLR_WHT
15230   End With

EXITP:
15240   Exit Sub

ERRH:
15250   THAT_PROC = THIS_PROC
15260   That_Erl = Erl
15270   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15280   Resume EXITP

End Sub

Private Sub pershare_Exit(Cancel As Integer)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "pershare_Exit"

        Dim msgResponse As VbMsgBoxResult
        Dim blnReverse As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String, dblTmp02 As Double
        Dim blnContinue As Boolean

15310   With Me
15320     blnContinue = True
15330     blnReverse = False
15340     If IsNumeric(.pershare) = True Then
15350       If .pershare < 0 Then
15360         Beep
15370         msgResponse = MsgBox("Are you sure you want a Reverse Stock Split?", vbQuestion + vbYesNo, "Negative Rate")
15380         If msgResponse <> vbYes Then
15390           blnContinue = False
15400           DoCmd.CancelEvent
15410           .pershare.SetFocus
15420         Else
15430           blnReverse = True
15440           .maptype.DefaultValue = (Chr(34) & "Withdrawn" & Chr(34))
15450           .pershare.BackColor = CLR_LTRED2
15460         End If
15470       Else
15480         If .maptype.DefaultValue <> (Chr(34) & "Deposit" & Chr(34)) Then
15490           .maptype.DefaultValue = (Chr(34) & "Deposit" & Chr(34))
15500         End If
15510       End If
15520       If blnContinue = True Then
15530         If CDbl(.pershare.Value) > 999999999.9999 Then
15540           blnContinue = False
15550           MsgBox "Rate is too large.", vbInformation + vbOKOnly, "Invalid Entry"
15560           .pershare = vbNullString
15570           DoCmd.CancelEvent
15580           .pershare.SetFocus
15590         Else
15600           intPos01 = InStr(1, .pershare, ".", vbTextCompare)
15610           If intPos01 > 0 Then
15620             If Len(.pershare) - intPos01 > 3 Then
15630               blnContinue = False
15640               MsgBox "You can only enter a Rate with up to 3 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
15650               .pershare = vbNullString
15660               DoCmd.CancelEvent
15670               .pershare.SetFocus
15680             End If
15690           End If
15700         End If
15710       End If
15720       If blnContinue = True And .pershare > 0 Then
              ' ** 100% = 2-for-1    100% increase in shares.
              ' ** 150% = 2.5-for-1  150% increase in shares.
15730         dblTmp02 = ((.pershare + 100#) / 100#)
15740         dblTmp02 = Round(dblTmp02, 2)
15750         If InStr(dblTmp02, ".") > 0 Then
                ' ** Try doubling or tripling.
15760           If InStr(CStr(dblTmp02 * 2#), ".") = 0 Then
15770             strTmp01 = CStr(dblTmp02 * 2#) & "-for-2"
15780           ElseIf InStr(CStr(dblTmp02 * 3#), ".") = 0 Then
15790             strTmp01 = CStr(dblTmp02 * 3#) & "-for-3"
15800           Else
15810             strTmp01 = CStr(dblTmp02) & "-for-1"
15820           End If
15830         Else
15840           strTmp01 = CStr(dblTmp02) & "-for-1"
15850         End If
15860         .pershare_lbl3.Caption = strTmp01
15870         .pershare_lbl3.Visible = True
15880         If blnReverse = True Then
15890           .pershare.BackColor = CLR_VLTRED
15900         End If
15910       Else
15920         .pershare_lbl3.Visible = False
15930         .pershare_lbl3.Caption = vbNullString
15940       End If
15950     Else
15960       .pershare_lbl3.Visible = False
15970       .pershare_lbl3.Caption = vbNullString
15980     End If
15990   End With  ' ** Me.

EXITP:
16000   Exit Sub

ERRH:
16010   Select Case ERR.Number
        Case Else
16020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16030   End Select
16040   Resume EXITP

End Sub

Private Sub assetdate_display_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_lbl2_MouseMove"

16110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
16120   Exit Sub

ERRH:
16130   Select Case ERR.Number
        Case Else
16140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16150   End Select
16160   Resume EXITP

End Sub

Private Sub assetdate_display_Enter()

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_Enter"

16210   With Me
16220     If IsNull(.assetdate) = False Then
16230       .assetdate_display = Format(.assetdate, "mm/dd/yyyy")
16240     End If
16250   End With

EXITP:
16260   Exit Sub

ERRH:
16270   Select Case ERR.Number
        Case Else
16280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16290   End Select
16300   Resume EXITP

End Sub

Private Sub assetdate_display_KeyDown(KeyCode As Integer, Shift As Integer)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
16470       With Me
16480         intRetVal = 0
16490         .transdate.SetFocus
16500       End With
16510     End Select
16520   End If

        ' ** Shift keys.
16530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16550       With Me
16560         intRetVal = 0
16570         .pershare.SetFocus
16580       End With
16590     End Select
16600   End If

        ' ** Ctrl keys.
16610   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
16620     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16630       With Me
16640         intRetVal = 0
16650         .cmdCalendar2.SetFocus
16660       End With
16670     End Select
16680   End If

EXITP:
16690   KeyCode = intRetVal
16700   Exit Sub

ERRH:
16710   intRetVal = 0
16720   THAT_PROC = THIS_PROC
16730   That_Erl = Erl
16740   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16750   Resume EXITP

End Sub

Private Sub assetdate_display_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_MouseMove"

16810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
16820   Exit Sub

ERRH:
16830   Select Case ERR.Number
        Case Else
16840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16850   End Select
16860   Resume EXITP

End Sub

Private Sub assetdate_display_Exit(Cancel As Integer)

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_Exit"

        Dim varTmp00 As Variant

16910   With Me
16920 On Error Resume Next
16930     varTmp00 = .assetdate_display.text  ' ** When called from cmdCalendar2_Click(), below, it's not always on focus.
16940     If ERR.Number <> 0 Then
16950       varTmp00 = .assetdate_display
16960     End If
16970 On Error GoTo ERRH
16980     If DateCheck_Trade(varTmp00) = True Then  ' ** Module Function: modUtilities.
            ' ** '.Text' needed to reliably check date.
16990       .assetdate = .assetdate_display + time()
17000     Else
17010       If IsNull(.assetdate_display) = True Then
17020         .assetdate_display = Date
17030         .assetdate = Now()
17040       Else
17050         If Trim(.assetdate_display) = vbNullString Then
17060           .assetdate_display = Date
17070           .assetdate = Now()
17080         Else
17090           Cancel = -1
17100           .assetdate_display.SetFocus
17110         End If
17120       End If
17130     End If
17140   End With

EXITP:
17150   Exit Sub

ERRH:
17160   THAT_PROC = THIS_PROC
17170   That_Erl = Erl
17180   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17190   Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

17210   With Me
17220     blnCalendar2_Focus = True
17230     .cmdCalendar2_raised_semifocus_dots_img.Visible = True
17240     .cmdCalendar2_raised_img.Visible = False
17250     .cmdCalendar2_raised_focus_img.Visible = False
17260     .cmdCalendar2_raised_focus_dots_img.Visible = False
17270     .cmdCalendar2_sunken_focus_dots_img.Visible = False
17280     .cmdCalendar2_raised_img_dis.Visible = False
17290   End With

EXITP:
17300   Exit Sub

ERRH:
17310   Select Case ERR.Number
        Case Else
17320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17330   End Select
17340   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

17410   With Me
17420     blnCalendar2_MouseDown = True
17430     .cmdCalendar2_sunken_focus_dots_img.Visible = True
17440     .cmdCalendar2_raised_img.Visible = False
17450     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
17460     .cmdCalendar2_raised_focus_img.Visible = False
17470     .cmdCalendar2_raised_focus_dots_img.Visible = False
17480     .cmdCalendar2_raised_img_dis.Visible = False
17490   End With

EXITP:
17500   Exit Sub

ERRH:
17510   Select Case ERR.Number
        Case Else
17520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17530   End Select
17540   Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

17610   With Me
17620     datStartDate = Date
17630     datEndDate = 0
17640     blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
17650     If blnRetVal = True Then
17660       If Compare_DateA_DateB(datStartDate, ">", Date) = True Then  ' ** Module Function: modStringFuncs.
17670         MsgBox "Future payable dates are not allowed.", vbInformation + vbOKOnly, "Invalid Date"
17680         .assetdate = Now()
17690         .assetdate_display = CDate(Format(.assetdate, "mm/dd/yyyy"))
17700       Else
17710         .assetdate = datStartDate + time()
17720         .assetdate_display = CDate(Format(.assetdate, "mm/dd/yyyy"))
17730       End If
17740     Else
17750       .assetdate = Now()
17760       .assetdate_display = CDate(Format(.assetdate, "mm/dd/yyyy"))
17770     End If
17780     .transdate.SetFocus
17790   End With

EXITP:
17800   Exit Sub

ERRH:
17810   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
17820   Case Else
17830     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17840   End Select
17850   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

17910   With Me
17920     If blnCalendar2_MouseDown = False Then
17930       Select Case blnCalendar2_Focus
            Case True
17940         .cmdCalendar2_raised_focus_dots_img.Visible = True
17950         .cmdCalendar2_raised_focus_img.Visible = False
17960       Case False
17970         .cmdCalendar2_raised_focus_img.Visible = True
17980         .cmdCalendar2_raised_focus_dots_img.Visible = False
17990       End Select
18000       .cmdCalendar2_raised_img.Visible = False
18010       .cmdCalendar2_raised_semifocus_dots_img.Visible = False
18020       .cmdCalendar2_sunken_focus_dots_img.Visible = False
18030       .cmdCalendar2_raised_img_dis.Visible = False
18040     End If
18050   End With

EXITP:
18060   Exit Sub

ERRH:
18070   Select Case ERR.Number
        Case Else
18080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18090   End Select
18100   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

18210   With Me
18220     .cmdCalendar2_raised_focus_dots_img.Visible = True
18230     .cmdCalendar2_raised_img.Visible = False
18240     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
18250     .cmdCalendar2_raised_focus_img.Visible = False
18260     .cmdCalendar2_sunken_focus_dots_img.Visible = False
18270     .cmdCalendar2_raised_img_dis.Visible = False
18280     blnCalendar2_MouseDown = False
18290   End With

EXITP:
18300   Exit Sub

ERRH:
18310   Select Case ERR.Number
        Case Else
18320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18330   End Select
18340   Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

18410   With Me
18420     .cmdCalendar2_raised_img.Visible = True
18430     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
18440     .cmdCalendar2_raised_focus_img.Visible = False
18450     .cmdCalendar2_raised_focus_dots_img.Visible = False
18460     .cmdCalendar2_sunken_focus_dots_img.Visible = False
18470     .cmdCalendar2_raised_img_dis.Visible = False
18480     blnCalendar2_Focus = False
18490   End With

EXITP:
18500   Exit Sub

ERRH:
18510   Select Case ERR.Number
        Case Else
18520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18530   End Select
18540   Resume EXITP

End Sub

Private Sub transdate_KeyDown(KeyCode As Integer, Shift As Integer)

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18620   intShiftDown = (Shift And acShiftMask) > 0
18630   intAltDown = (Shift And acAltMask) > 0
18640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18670       With Me
18680         intRetVal = 0
18690         .cmdOK.SetFocus
18700       End With
18710     End Select
18720   End If

        ' ** Shift keys.
18730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18750       With Me
18760         intRetVal = 0
18770         .assetdate_display.SetFocus
18780       End With
18790     End Select
18800   End If

        ' ** Ctrl keys.
18810   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
18820     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18830       With Me
18840         intRetVal = 0
18850         .cmdCalendar1.SetFocus
18860       End With
18870     End Select
18880   End If

        ' ** Ctrl-Shift keys.
18890   If intCtrlDown And (Not intAltDown) And intShiftDown Then
18900     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18910       With Me
18920         intRetVal = 0
18930         .cmdCalendar2.SetFocus
18940       End With
18950     End Select
18960   End If

EXITP:
18970   KeyCode = intRetVal
18980   Exit Sub

ERRH:
18990   intRetVal = 0
19000   THAT_PROC = THIS_PROC
19010   That_Erl = Erl
19020   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19030   Resume EXITP

End Sub

Private Sub transdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_MouseMove"

19110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
19120   Exit Sub

ERRH:
19130   Select Case ERR.Number
        Case Else
19140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19150   End Select
19160   Resume EXITP

End Sub

Private Sub transdate_Exit(Cancel As Integer)

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_Exit"

        Dim varTmp00 As Variant

19210   With Me
19220 On Error Resume Next
19230     varTmp00 = .transdate.text  ' ** When called from cmdCalendar1_Click(), below, it's not always on focus.
19240     If ERR.Number <> 0 Then
19250       varTmp00 = .transdate
19260     End If
19270 On Error GoTo ERRH
19280     If DateCheck_Post(varTmp00) = True Then  ' ** Module Function: modUtilities.
            ' ** '.Text' needed to reliably check date.
19290     Else
19300       If IsNull(.transdate) = True Then
19310         .transdate = Date
19320       Else
19330         If Trim(.transdate) = vbNullString Then
19340           .transdate = Date
19350         Else
19360           Cancel = -1
19370           .transdate.SetFocus
19380         End If
19390       End If
19400     End If
19410   End With

EXITP:
19420   Exit Sub

ERRH:
19430   Select Case ERR.Number
        Case Else
19440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19450   End Select
19460   Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

19510   With Me
19520     blnCalendar1_Focus = True
19530     .cmdCalendar1_raised_semifocus_dots_img.Visible = True
19540     .cmdCalendar1_raised_img.Visible = False
19550     .cmdCalendar1_raised_focus_img.Visible = False
19560     .cmdCalendar1_raised_focus_dots_img.Visible = False
19570     .cmdCalendar1_sunken_focus_dots_img.Visible = False
19580     .cmdCalendar1_raised_img_dis.Visible = False
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

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

19710   With Me
19720     blnCalendar1_MouseDown = True
19730     .cmdCalendar1_sunken_focus_dots_img.Visible = True
19740     .cmdCalendar1_raised_img.Visible = False
19750     .cmdCalendar1_raised_semifocus_dots_img.Visible = False
19760     .cmdCalendar1_raised_focus_img.Visible = False
19770     .cmdCalendar1_raised_focus_dots_img.Visible = False
19780     .cmdCalendar1_raised_img_dis.Visible = False
19790   End With

EXITP:
19800   Exit Sub

ERRH:
19810   Select Case ERR.Number
        Case Else
19820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19830   End Select
19840   Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

19910   With Me
19920     datStartDate = Date
19930     datEndDate = 0
19940     blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
19950     If blnRetVal = True Then
            ' ** Allow posting up to 1 month into the future.
19960       If .transdate > DateAdd("m", 1, Date) Then
19970         MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
19980         .transdate = CDate(Format(Date, "mm/dd/yyyy"))
19990       Else
20000         .transdate = datStartDate
20010       End If
20020     Else
20030       .transdate = CDate(Format(Date, "mm/dd/yyyy"))
20040     End If
          ' ** Make the date 'is edited' as if were entered by hand.
20050     .transdate.SetFocus
20060     Cancel = 0
20070     transdate_Exit Cancel  ' ** Procedure: Above.
20080     If Cancel = 0 Then
20090       .cmdOK.SetFocus
20100     End If
20110   End With

EXITP:
20120   Exit Sub

ERRH:
20130   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
20140   Case Else
20150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20160   End Select
20170   Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

20210   With Me
20220     If blnCalendar1_MouseDown = False Then
20230       Select Case blnCalendar1_Focus
            Case True
20240         .cmdCalendar1_raised_focus_dots_img.Visible = True
20250         .cmdCalendar1_raised_focus_img.Visible = False
20260       Case False
20270         .cmdCalendar1_raised_focus_img.Visible = True
20280         .cmdCalendar1_raised_focus_dots_img.Visible = False
20290       End Select
20300       .cmdCalendar1_raised_img.Visible = False
20310       .cmdCalendar1_raised_semifocus_dots_img.Visible = False
20320       .cmdCalendar1_sunken_focus_dots_img.Visible = False
20330       .cmdCalendar1_raised_img_dis.Visible = False
20340     End If
20350   End With

EXITP:
20360   Exit Sub

ERRH:
20370   Select Case ERR.Number
        Case Else
20380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20390   End Select
20400   Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

20510   With Me
20520     .cmdCalendar1_raised_focus_dots_img.Visible = True
20530     .cmdCalendar1_raised_img.Visible = False
20540     .cmdCalendar1_raised_semifocus_dots_img.Visible = False
20550     .cmdCalendar1_raised_focus_img.Visible = False
20560     .cmdCalendar1_sunken_focus_dots_img.Visible = False
20570     .cmdCalendar1_raised_img_dis.Visible = False
20580     blnCalendar1_MouseDown = False
20590   End With

EXITP:
20600   Exit Sub

ERRH:
20610   Select Case ERR.Number
        Case Else
20620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20630   End Select
20640   Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

20710   With Me
20720     .cmdCalendar1_raised_img.Visible = True
20730     .cmdCalendar1_raised_semifocus_dots_img.Visible = False
20740     .cmdCalendar1_raised_focus_img.Visible = False
20750     .cmdCalendar1_raised_focus_dots_img.Visible = False
20760     .cmdCalendar1_sunken_focus_dots_img.Visible = False
20770     .cmdCalendar1_raised_img_dis.Visible = False
20780     blnCalendar1_Focus = False
20790   End With

EXITP:
20800   Exit Sub

ERRH:
20810   Select Case ERR.Number
        Case Else
20820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20830   End Select
20840   Resume EXITP

End Sub

Private Sub IncludeCurrency()

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngTmp01 As Long

20910   With Me
20920     .curr_id.Enabled = False
20930     .curr_id.Locked = True
20940     .curr_id.ForeColor = CLR_DISABLED_FG
20950     .curr_id.BackColor = CLR_VLTGRN
20960     Select Case blnCurrID
          Case True
20970       .curr_id.Visible = True
20980       .curr_id_cmd.Visible = True
20990       .curr_id_cmd.Enabled = True
21000       .cmbAssets_vline01.Visible = True
21010       .cmbAssets_vline02.Visible = True
21020     Case False
21030       .curr_id.Visible = False
21040       .curr_id_cmd.Visible = False
21050       .curr_id_cmd.Enabled = False
21060       .cmbAssets_vline01.Visible = False
21070       .cmbAssets_vline02.Visible = False
21080       lngTmp01 = ((.curr_id.Left + .curr_id.Width) - (.cmbAssets.Left + .cmbAssets.Width))
21090       .cmbAssets.Width = (.cmbAssets.Width + lngTmp01)
21100     End Select
21110     .pershare_usd.Visible = False  ' ** It'll always be False!
21120   End With

        ' ** cmbAssets RowSource's:
        ' **   qryMap_Split_05a
        ' **   qryMap_Split_05c
        ' **   qryMap_Split_05d
        ' **   qryMap_Split_05f
        ' **   qryMap_Split_05g
        ' **   qryMap_Split_05i

EXITP:
21130   Exit Sub

ERRH:
21140   DoCmd.Hourglass False
21150   Select Case ERR.Number
        Case Else
21160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21170   End Select
21180   Resume EXITP

End Sub

Private Sub NoChar_Load()

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

21210   Set dbs = CurrentDb
21220   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
21230     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
21240     Set rst = qdf.OpenRecordset
21250     With rst
21260       .MoveLast
21270       lngNoChars = .RecordCount
21280       .MoveFirst
21290       arr_varNoChar = .GetRows(lngNoChars)
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
21300       .Close
21310     End With
21320     .Close
21330   End With

EXITP:
21340   Set rst = Nothing
21350   Set qdf = Nothing
21360   Set dbs = Nothing
21370   Exit Sub

ERRH:
21380   DoCmd.Hourglass False
21390   Select Case ERR.Number
        Case Else
21400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21410   End Select
21420   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

21510   With Me
21520     blnGTR_Emblem = False
21530     For lngX = 1& To 24&
21540       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
21550     Next
21560   End With

EXITP:
21570   Exit Sub

ERRH:
21580   Select Case ERR.Number
        Case Else
21590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21600   End Select
21610   Resume EXITP

End Sub
