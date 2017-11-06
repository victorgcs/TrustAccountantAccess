VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Misc_STCGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Misc_STCGL"

'VGC 09/09/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Date:             D {transdate}
' **   Income Cash:      H {Amount}
' **   Asset Cusip:      I {opgAssetSource_optCusip}
' **   Asset Name:       N {opgAssetSource_optName}
' **   OK:               O {cmdOK}
' **   Per Share:        R {pershare}
' **   Select Asset:     S {cmbAssets}
' **   Asset Type:       Y {opgAssetSource_optType}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar}

' ** MapType = "Misc."

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of Calendar Class.
Public clsMonthClass As clsMonthCal

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long
Private CLR_DISABLED_BG_R As Long
Private CLR_DISABLED_BG_VLR As Long

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

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean

Private strFrmt1 As String, strFrmt2 As String, intDec1 As Integer, intDec2 As Integer, lngBkClr1 As Long, lngBkClr2 As Long
Private strAccountNo_OU As String
Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private dblPerShare_Calc As Double, lngTpp As Long, blnCurrID As Boolean, lngCurrID As Long
Private datPostingDate As Date, intInputType As Integer
Private blnAssetUpdated As Boolean, blnAssetChecked As Boolean, blnDrivingMeBackwards As Boolean
Private strCallingForm As String
Private blnExitingAssetNo As Boolean, blnCancel As Boolean
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
190           blnGTR_Emblem = False: lngGTR_Stat = 0&
200           intInputType = 0
210           dblPerShare_Calc = 0#

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
220           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

230           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

240           Select Case gblnAccountNoWithType
              Case True
250             strAccountNo_OU = "99-INCOME O/U"
260           Case False
270             strAccountNo_OU = "INCOME O/U"
280           End Select

              ' ** Borrowing these variables from the Court Reports.
290           gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString

300           CLR_DISABLED_FG = CLR_DKGRY
310           CLR_DISABLED_BG = CLR_LTTEAL
320           CLR_DISABLED_BG_R = CLR_LTRED
330           CLR_DISABLED_BG_VLR = CLR_VLTRED

340           opgAssetSource_AfterUpdate  ' ** Procedure: Below.

350         Else
360           Cancel = -1
370         End If
380       Else
390         Cancel = -1
400       End If
410     End With

420     If Cancel = -1 Then
430       Beep
440       DoCmd.Hourglass False
450       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
460     End If

EXITP:
470     Exit Sub

ERRH:
480     DoCmd.Hourglass False
490     Select Case ERR.Number
        Case Else
500       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
510     End Select
520     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

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
890         DoEvents
            ' ** Empty tmpAccount.
900         Set qdf = .QueryDefs("qryMap_Misc_STCGL_08_01")
910         qdf.Execute
920         Set qdf = Nothing
930         DoEvents
            ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
            ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
            ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
            ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
            ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
            ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
            ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
            ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
940         Set qdf = .QueryDefs("qryMap_Misc_STCGL_08_02")
950         qdf.Execute
960         Set qdf = Nothing
970         DoEvents
980         .Close
990       End With
1000      Set dbs = Nothing
1010      DoEvents

1020      .transdate.DefaultValue = "#" & Format(datPostingDate, "mm/dd/yyyy") & "#"
1030      .shareface.ForeColor = CLR_DISABLED_FG
1040      .shareface.BackColor = CLR_DISABLED_BG

1050      IncludeCurrency  ' ** Procedure: Below.

1060    End With

1070    DoCmd.Hourglass False

EXITP:
1080    Set qdf = Nothing
1090    Set dbs = Nothing
1100    Exit Sub

ERRH:
1110    DoCmd.Hourglass False
1120    Select Case ERR.Number
        Case Else
1130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1140    End Select
1150    Resume EXITP

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

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyEscape
1270        intRetVal = 0
1280        cmdCancel_Click  ' ** Procedure: Below.
1290      Case vbKeyF7
1300        intRetVal = 0
1310        cmdCalendar_Click  ' ** Procedure: Below.
1320      End Select
1330    End If

        ' ** Ctrl-Shift keys.
1340    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1350      Select Case intRetVal
          Case vbKeyF
1360        With Me
1370          intRetVal = 0
1380          .FocusHolder.SetFocus
1390        End With
1400      End Select
1410    End If

EXITP:
1420    KeyCode = intRetVal
1430    Exit Sub

ERRH:
1440    intRetVal = 0
1450    THAT_PROC = THIS_PROC
1460    That_Erl = Erl
1470    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1480    Resume EXITP

End Sub

Private Sub Form_Timer()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngAssetNo As Long
        Dim Cancel As Integer
        Dim lngTmp03 As Long

        Static lngTmp01 As Long, lngTmp02 As Long

1510    With Me
1520      .TimerInterval = 0
1530      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
1540        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
1550      End If
1560      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
1570        lngGTR_Stat = lngGTR_Stat + 1&
1580        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1590        lngTmp01 = lngTmp01 + 1&
1600        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1610        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1620        If lngTmp01 > 1& Then
1630          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1640        End If
1650        If lngTmp01 = 24& Then
1660          lngTmp01 = 0&
1670          lngTmp02 = lngTmp02 + 1&
1680          If lngTmp02 < 8& Then
1690            .TimerInterval = 50&
1700          Else
1710            blnGTR_Emblem = False
1720          End If
1730        Else
1740          .TimerInterval = 50&
1750        End If
1760      End If  ' ** blnGTR_Emblem.
1770      If blnExitingAssetNo = True Then
            ' ** This was set up this way so that the user doesn't have to hit Cancel twice
            ' ** if they happened to have entered an asset, but then changed their mind.
1780        blnExitingAssetNo = False
1790        If blnCancel = False Then
1800          Set dbs = CurrentDb
1810          Select Case strCallingForm
              Case "frmJournal"
                ' ** ActiveAssets, grouped and summed, by assetno, by specified [astno].
1820            Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_07")
1830          Case "frmJournal_Columns"
                ' ** ActiveAssets, grouped and summed, by assetno, by specified [astno].
1840            Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_57")
1850          End Select
1860          With qdf.Parameters
1870            ![astno] = Me.cmbAssets
1880          End With
1890          Set rst = qdf.OpenRecordset
1900          If rst.BOF = True And rst.EOF = True Then
                ' ** Shouldn't happen.
1910          Else
1920            blnAssetUpdated = False
1930            blnAssetChecked = True
1940            .TotalShares = rst![sumsf]
1950            .shareface = rst![sumsf]
1960            If .amount.Enabled = True Then
1970              dblPerShare_Calc = (Nz(.amount, 0) / IIf(Nz(.shareface, 0) = 0, 1, .shareface))
1980              .pershare = dblPerShare_Calc  '5 dec.
1990  On Error Resume Next
2000              .amount.SetFocus
2010  On Error GoTo ERRH
2020            Else
2030              .amount = (.pershare * IIf(Nz(.shareface, 0) = 0, 1, .shareface))
2040  On Error Resume Next
2050              .pershare.SetFocus
2060  On Error GoTo ERRH
2070            End If
2080          End If
2090          rst.Close
2100          dbs.Close
2110        End If
2120      End If
2130      If gblnGoToReport = True Then
2140        DoCmd.Hourglass True  ' ** Make sure it's still running.
2150        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48& And lngGTR_Stat < 96&) Or _
                (lngGTR_Stat > 96& And lngGTR_Stat < 120&) Or (lngGTR_Stat > 120&)) Then
              ' ** Just let the emblem go.
2160        Else
2170          DoCmd.Hourglass True  ' ** Make sure it's still running.
2180          If lngTpp = 0& Then
                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
2190            lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
2200          End If
2210          If blnGTR_Emblem = False Then
2220            blnGTR_Emblem = True
2230            lngTmp01 = 0&: lngTmp02 = 0&
2240            lngTmp03 = (.GoToReport_Emblem_01_img.Width + (8& * lngTpp))
2250            .GoToReport_arw_ok_img.Left = (.GoToReport_arw_ok_img.Left - lngTmp03)
2260            .cmdOK.Left = (.cmdOK.Left - lngTmp03)
2270            .cmdCancel.Left = (.cmdCancel.Left - lngTmp03)
2280          End If
2290          DoEvents
2300          Select Case blnGoingToReport
              Case True
2310            Select Case blnGoingToReport2
                Case True
2320              Select Case .GoToReport_arw_ok_img.Visible
                  Case True
2330                blnGoingToReport2 = False
2340                .TimerInterval = 50&
2350                cmdOK_Click  ' ** Procedure: Below.
2360              Case False
2370                .amount.SetFocus
2380                If garr_varGoToReport(GTR_RNAM) = "rptMap_Misc_STCG" Then
2390                  .amount = 100
2400                ElseIf garr_varGoToReport(GTR_RNAM) = "rptMap_Misc_STCL" Then
2410                  .amount = -100
2420                End If
2430                Cancel = 0
2440                amount_Exit Cancel  ' ** Procedure: Below.
2450                DoEvents
2460                .GoToReport_arw_ok_img.Visible = True
2470                .GoToReport_arw_map_amt_img.Visible = False
2480                .GoToReport_arw_map_per_img.Visible = False
2490                .GoToReport_lbl_amtper.Visible = False
2500                .GoToReport_lin_hline01.Visible = False
2510                .GoToReport_lin_hline02.Visible = False
2520                .GoToReport_lbl_arrow01.Visible = False
2530                .GoToReport_lbl_arrow02.Visible = False
2540                DoEvents
2550                .TimerInterval = 50&
2560              End Select
2570            Case False
2580              blnGoingToReport2 = True
2590              lngAssetNo = GetDivAsset  ' ** Module Function: modGoToReportFuncs.
2600              If lngAssetNo > 0& Then
2610                .cmbAssets = lngAssetNo
2620                cmbAssets_AfterUpdate  ' ** Procedure: Below.
2630                DoEvents
2640                .GoToReport_arw_map_amt_img.Visible = True
2650                .GoToReport_arw_map_per_img.Visible = True
2660                .GoToReport_lbl_amtper.Visible = True
2670                If strCallingForm = "frmJournal_Columns" Then
                      ' ** Don't know why this continues to light up.
2680                  Forms(strCallingForm).cmdSpecPurp_Purch_MapSplit_raised_img.Visible = True
2690                  Forms(strCallingForm).cmdSpecPurp_Purch_MapSplit_raised_semifocus_dots_img.Visible = False
2700                End If
2710                .GoToReport_lin_hline01.Visible = True
2720                .GoToReport_lin_hline02.Visible = True
2730                .GoToReport_lbl_arrow01.Visible = True
2740                .GoToReport_lbl_arrow02.Visible = True
2750                .GoToReport_arw_map_asset_img.Visible = False
2760                .opgAssetSource.Visible = True
2770                .opgAssetSource_box.Visible = True
2780                .TotalShares = 4000
2790                .shareface = 4000
2800                .FocusHolder.SetFocus  ' ** So they don't see a flashing cursor.
2810                DoEvents
2820                .TimerInterval = 50&
2830              Else
2840                .TimerInterval = 0&
2850                Select Case garr_varGoToReport(GTR_FRM2)
                    Case "frmJournal"
2860                  Forms(garr_varGoToReport(GTR_FRM2)).GTREmblem_Set False  ' ** Form Procedure: frmJournal.
2870                Case "frmJournal_Columns"
2880                  Forms(garr_varGoToReport(GTR_FRM2)).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
2890                End Select
2900                gblnGoToReport = False
2910                blnGoingToReport = False
2920                blnGoingToReport2 = False
2930                GTREmblem_Off  ' ** Procedure: Below.
2940                .GoToReport_arw_map_asset_img.Visible = False
2950                .opgAssetSource.Visible = True
2960                .opgAssetSource_box.Visible = True
2970                Beep
2980                DoCmd.Hourglass False
2990                MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                      "There are insufficient asset holdings to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
3000              End If
3010            End Select
3020          Case False
3030            blnGoingToReport = True
3040            .GoToReport_arw_map_asset_img.Visible = True
3050            If strCallingForm = "frmJournal_Columns" Then
                  ' ** Don't know why this continues to light up.
3060              Forms(strCallingForm).cmdSpecPurp_Purch_MapSplit_raised_img.Visible = True
3070              Forms(strCallingForm).cmdSpecPurp_Purch_MapSplit_raised_semifocus_dots_img.Visible = False
3080            End If
3090            .opgAssetSource.Visible = False
3100            .opgAssetSource_box.Visible = False
3110            .FocusHolder.SetFocus  ' ** So they don't see a flashing cursor.
3120            DoEvents
3130            Select Case garr_varGoToReport(GTR_FRM2)
                Case "frmJournal"
3140              Forms(garr_varGoToReport(GTR_FRM2)).GTREmblem_Set False  ' ** Form Procedure: frmJournal.
3150            Case "frmJournal_Columns"
3160              Forms(garr_varGoToReport(GTR_FRM2)).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
3170            End Select
3180            .TimerInterval = 50&
3190          End Select
3200        End If
3210      End If

3220    End With  ' ** Me.

EXITP:
3230    Set rst = Nothing
3240    Set qdf = Nothing
3250    Set dbs = Nothing
3260    Exit Sub

ERRH:
3270    Select Case ERR.Number
        Case Else
3280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3290    End Select
3300    Resume EXITP

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
3460      Case 2113  ' ** The value you entered isn't valid for this field.
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
3580          Exit Sub
3590        Case Else
3600          strControlName = vbNullString
3610        End Select
3620        If strControlName <> vbNullString Then
3630          MsgBox "You must enter a numeric value in the " & strControlName & " field.", vbInformation + vbOKOnly, "Invalid Entry"
3640          Response = acDataErrContinue
3650          DoCmd.CancelEvent
3660        End If
3670      Case 2237  ' ** The text you entered isn't an item in the list.
3680        DoCmd.CancelEvent
3690      Case 2279  ' ** The value you entered isn't appropriate for the input mask '|' specified for this field.
3700        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3710        Response = acDataErrContinue
3720        DoCmd.CancelEvent
3730      Case 2580  ' ** The record source '|' specified on this form or report does not exist.
            ' ** Ignore! This form has no RecordSource.
3740        Response = acDataErrContinue
3750      Case 3162  ' ** You tried to assign the Null value to a variable that is not a Variant data type.
3760        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
            ' ** Do nothing.
3770      Case Else
3780        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3790        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3800      End Select
3810    End If

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3910    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3920    If Not clsMonthClass Is Nothing Then
3930      If clsMonthClass.IsCalendar Then
3940        Cancel = -1
3950      Else
3960        Set clsMonthClass = Nothing
3970      End If
3980    End If

3990    DoCmd.Hourglass False

EXITP:
4000    Exit Sub

ERRH:
4010    DoCmd.Hourglass False
4020    Select Case ERR.Number
        Case Else
4030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4040    End Select
4050    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4110    With Me
4120      If .cmdCalendar_raised_focus_dots_img.Visible = True Or .cmdCalendar_raised_focus_img.Visible = True Then
4130        Select Case blnCalendar1_Focus
            Case True
4140          .cmdCalendar_raised_semifocus_dots_img.Visible = True
4150          .cmdCalendar_raised_img.Visible = False
4160        Case False
4170          .cmdCalendar_raised_img.Visible = True
4180          .cmdCalendar_raised_semifocus_dots_img.Visible = False
4190        End Select
4200        .cmdCalendar_raised_focus_dots_img.Visible = False
4210        .cmdCalendar_raised_focus_img.Visible = False
4220        .cmdCalendar_sunken_focus_dots_img.Visible = False
4230        .cmdCalendar_raised_img_dis.Visible = False
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

Private Sub cmdOK_Click()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        'Dim cnxn As ADODB.Connection, rsx1 As ADODB.Recordset  ' ** Early binding.
        Dim cnxn As Object, rsx1 As Object                      ' ** Late binding.
        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim dblShareface As Double, dblAmount As Double, dblPerShare As Double
        Dim dblSumICash As Double, dblSumShareface As Double, dblOverUnder As Double
        Dim datTransDate As Date, strMapType As String
        Dim lngAssetNo As Long, lngTaxcode As Long
        Dim strDocName As String
        Dim lngMaps As Long, arr_varMap As Variant
        Dim lngAdjs As Long, arr_varAdj As Variant
        Dim lngRecs As Long
        Dim blnContinue As Boolean, blnFound As Boolean, blnSkip As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim intPos01 As Integer
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
        'Const A_ICASH As Integer = 3
        Const A_PCADJ As Integer = 4

4410    blnContinue = True

4420    DoCmd.Hourglass True
4430    DoEvents

4440    dblShareface = 0#
4450    dblAmount = 0#
4460    dblSumICash = 0#
4470    dblSumShareface = 0#
4480    dblPerShare = 0#
4490    dblOverUnder = 0#
4500    lngAssetNo = 0&
4510    datTransDate = 0
4520    lngTaxcode = 0&

4530    With Me

          ' ** Step 1: Validation.
4540      .Map_Status_lbl.Caption = " 1 of 22: Validation . . ."
          '.Map_Status_lbl.Visible = True
4550      DoEvents

4560      If IsNull(.cmbAssets) = True Then
4570        blnContinue = False
4580        DoCmd.Hourglass False
4590        MsgBox "Please select a valid asset.", vbInformation + vbOKOnly, "Entry Required"
4600        .cmbAssets.SetFocus
4610      Else
4620        If Trim(.cmbAssets) = vbNullString Then
4630          blnContinue = False
4640          DoCmd.Hourglass False
4650          MsgBox "Please select a valid asset.", vbInformation + vbOKOnly, "Entry Required"
4660          .cmbAssets.SetFocus
4670        Else
4680          lngAssetNo = .cmbAssets
4690          If Nz(.shareface, 0) = 0 Then
4700            blnContinue = False
4710            DoCmd.Hourglass False
4720            MsgBox "Share/Face must be a number greater than zero.", vbInformation + vbOKOnly, "Entry Required"
4730            .cmbAssets.SetFocus
4740          Else
4750            If IsNull(.amount) = True Then
4760              blnContinue = False
4770              DoCmd.Hourglass False
4780              MsgBox "Amount cannot be zero.", vbInformation + vbOKOnly, "Entry Required"
4790              .amount.SetFocus
4800            Else
4810              If IsNull(.pershare) = True Then
4820                blnContinue = False
4830                DoCmd.Hourglass False
4840                MsgBox "Per Share cannot be zero.", vbInformation + vbOKOnly, "Entry Required"
4850                .pershare.SetFocus
4860              Else
4870                If Not IsDate(.transdate) Then
4880                  blnContinue = False
4890                  DoCmd.Hourglass False
4900                  MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Date"
4910                  .transdate.SetFocus
4920                Else
                      ' ** Allow posting up to 1 month into the future.
4930                  If .transdate > DateAdd("m", 1, Date) Then
4940                    blnContinue = False
4950                    DoCmd.Hourglass False
4960                    MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
4970                  Else
4980                    If .amount = 0 And .pershare = 0 Then
4990                      blnContinue = False
5000                      DoCmd.Hourglass False
5010                      MsgBox "Either Amount or Per Share must be greater than zero to continue.", vbInformation + vbOKOnly, "Entry Required"
5020                      If .amount.Enabled = False Or .pershare.Enabled = False Then
5030                        .amount = Null
5040                        .amount.Enabled = True
5050                        .amount.Locked = False
5060                        .amount.ForeColor = CLR_BLK
5070                        .pershare = Null
5080                        .pershare.Enabled = True
5090                        .pershare.Locked = False
5100                        .pershare.ForeColor = CLR_BLK
5110                        If lngCurrID = 150& Then
5120                          .amount.BackColor = CLR_WHT
5130                          .pershare.BackColor = CLR_WHT
5140                        Else
5150                          .amount.BackColor = CLR_VLTGRN
5160                          .pershare.BackColor = CLR_VLTGRN
5170                        End If
5180                      End If
5190                      .amount.SetFocus
5200                    End If
5210                  End If
5220                End If
5230              End If
5240            End If
5250          End If
5260        End If
5270      End If
5280      DoEvents

5290      If blnContinue = True Then

5300        DoCmd.Hourglass True
5310        DoEvents

5320        Set dbs = CurrentDb

            ' ** Step 2: Empty Journal Map.
5330        .Map_Status_lbl.Caption = " 2 of 22: Empty temporary table 1 . . ."
5340        DoEvents

5350        Select Case strCallingForm
            Case "frmJournal"
              ' ** Empty Journal Map.
5360          Set qdf = dbs.QueryDefs("qryMap_01")
5370        Case "frmJournal_Columns"
              ' ** Empty Journal Map.
5380          Set qdf = dbs.QueryDefs("qryJournal_Columns_35_04")
5390        End Select
5400        qdf.Execute
5410        DoEvents
5420        Set qdf = Nothing

            ' ** Step 3: Set Tax Code.
5430        .Map_Status_lbl.Caption = " 3 of 22: Set Tax Code . . ."
5440        DoEvents

5450        Select Case strCallingForm
            Case "frmJournal"
              ' ** MasterAsset, linked to AssetType, by specified [astno].
5460          Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_02")
5470        Case "frmJournal_Columns"
              ' ** MasterAsset, linked to AssetType, by specified [astno].
5480          Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_52")
5490        End Select
5500        With qdf.Parameters
5510          ![astno] = lngAssetNo
5520        End With
5530        Set rst = qdf.OpenRecordset
5540        If rst.BOF = True And rst.EOF = True Then
              ' ** Shouldn't happen.
5550        Else
5560          rst.MoveFirst
5570          lngTaxcode = rst![taxcode]
5580        End If
5590        rst.Close
5600        Set rst = Nothing
5610        Set qdf = Nothing
5620        .taxcode = lngTaxcode
5630        DoEvents

            ' ** Step 4: Calculate PerShare/Amount.
5640        .Map_Status_lbl.Caption = " 4 of 22: Calculate PerShare/Amount . . ."
5650        DoEvents

5660        dblShareface = .TotalShares
5670        Select Case intInputType
            Case 1  ' ** User entered an Amount.
5680          dblAmount = .amount
5690          dblPerShare = (dblAmount / dblShareface)
5700        Case 2  ' ** User entered a PerShare.
5710          dblPerShare = .pershare
5720          dblAmount = (dblShareface * dblPerShare)
5730        Case Else  ' ** Shouldn't be an else!
5740          dblAmount = .amount
5750          dblPerShare = (dblAmount / dblShareface)
5760        End Select
5770        datTransDate = .transdate

5780        strMapType = .maptype  ' ** JournalType.
5790        DoEvents

5800        If dblPerShare > 0# Then
              ' ** Append qryMap_Misc_STCGL_Detail_03a (Query Source; RecurringType, just 'Misc',
              ' ** as new RecurringItems entry, for 'Short Term Capital Gain') to RecurringItems.
5810          Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_Detail_05a")
5820        Else
              ' ** Append qryMap_Misc_STCGL_Detail_03b (Query Source; RecurringType, just 'Misc',
              ' ** as new RecurringItems entry, for 'Short Term Capital Loss') to RecurringItems.
5830          Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_Detail_05b")
5840        End If
5850        qdf.Execute
5860        Set qdf = Nothing

            ' ** Step 5: Check for closed accounts.
5870        .Map_Status_lbl.Caption = " 5 of 22: Check for closed accounts . . ."
5880        DoEvents

            ' ** Check for closed accounts holding assets.
            ' ** qryMap_Misc_STCGL_04 (qryMap_Misc_STCGL_03 (ActiveAssets, linked to
            ' ** Account, just closed accounts, by specified [astno]), grouped by
            ' ** accountno, with cnt_astno), grouped and summed, with cnt_actno, cnt_astno.
5890        Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_05")
5900        With qdf.Parameters
5910          ![astno] = Me.cmbAssets
5920        End With
5930        Set rst = qdf.OpenRecordset
5940        If rst.BOF = True And rst.EOF = True Then
              ' ** No closed accounts holding assets.
5950          msgResponse = vbYes
5960          rst.Close
5970          Set rst = Nothing
5980          Set qdf = Nothing
5990        Else
6000          rst.MoveFirst
6010          If IsNull(rst![cnt_actno]) = True And IsNull(rst![cnt_astno]) = True Then
6020            msgResponse = vbYes
6030            rst.Close
6040            Set rst = Nothing
6050            Set qdf = Nothing
6060          Else
6070            If rst![cnt_actno] = 0 And IsNull(rst![cnt_astno]) = True Then
6080              msgResponse = vbYes
6090              rst.Close
6100              Set rst = Nothing
6110              Set qdf = Nothing
6120            ElseIf rst![cnt_actno] = 0 And rst![cnt_astno] = 0 Then
6130              msgResponse = vbYes
6140              rst.Close
6150              Set rst = Nothing
6160              Set qdf = Nothing
6170            Else
6180              lngTmp02 = rst![cnt_actno]
6190              lngTmp03 = rst![cnt_astno]
6200              rst.Close
6210              Set rst = Nothing
6220              Set qdf = Nothing
                  ' ** qryMap_Misc_STCGL_03 (ActiveAssets, linked to Account, just closed
                  ' ** accounts, by specified [astno]), grouped by accountno, with cnt_astno.
6230              Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_04")
6240              With qdf.Parameters
6250                ![astno] = lngAssetNo
6260              End With
6270              Set rst = qdf.OpenRecordset
6280              rst.MoveFirst
6290              If lngTmp02 = 1& Then
6300                strTmp01 = "There is 1 closed account currently holding this asset:" & vbCrLf & _
                      "    " & rst![accountno] & "  " & rst![shortname] & vbCrLf
6310              Else
6320                strTmp01 = "There are " & CStr(lngTmp02) & " closed accounts currently holding this asset:" & vbCrLf
6330                For lngX = 1& To lngTmp02
6340                  strTmp01 = strTmp01 & "    " & rst![accountno] & "  " & rst![shortname] & vbCrLf
6350                  If lngX < lngTmp02 Then rst.MoveNext
6360                Next
6370              End If
6380              strTmp01 = strTmp01 & "If you choose to continue with this mapping," & vbCrLf & _
                    "the results may be inaccurate." & vbCrLf & vbCrLf & "Do you wish to continue?"
6390              rst.Close
6400              Set rst = Nothing
6410              Set qdf = Nothing
6420              DoCmd.Hourglass False
6430              Beep
6440              msgResponse = MsgBox(strTmp01, vbQuestion + vbYesNo + vbDefaultButton2, "Closed Accounts Holding Assets")
6450            End If
6460          End If
6470        End If
6480        DoEvents

6490        If msgResponse = vbYes Then

6500          DoCmd.Hourglass True
6510          DoEvents

              ' ** Step 6: Collect Active Assets.
6520          .Map_Status_lbl.Caption = " 6 of 22: Collect Active Assets . . ."
6530          DoEvents

6540          Select Case strCallingForm
              Case "frmJournal"
                ' ** Append ActiveAssets to Journal Map, by specified [astno], [jrntyp], [persh], [transdat].  #curr_id
6550            Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_06")
6560          Case "frmJournal_Columns"
                ' ** Append ActiveAssets to Journal Map, by specified [astno], [jrntyp], [persh], [transdat].  #curr_id
6570            Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_56")
6580          End Select
6590          With qdf.Parameters
6600            ![astno] = lngAssetNo
6610            ![jrntyp] = strMapType
6620            ![persh] = dblPerShare  ' ** Used for both ICash and pershare, negative and positive.
6630            ![transdat] = datTransDate
6640          End With
6650          qdf.Execute
6660          dbs.Close
6670          Set qdf = Nothing
6680          Set dbs = Nothing
6690          DoEvents

6700          blnSkip = True
              ' ################################################################################
6710          If blnSkip = False Then
                'Set rsx1 = New ADODB.Recordset             ' ** Early binding.
6720            Set rsx1 = CreateObject("ADODB.Recordset")  ' ** Late binding.
6730            Set cnxn = CurrentProject.Connection
6740            rsx1.Open "journal map", cnxn, adOpenKeyset, adLockOptimistic, adCmdTableDirect
6750            With rsx1
6760              .MoveLast
6770              lngRecs = .RecordCount
6780              .MoveFirst
6790              For lngX = 1& To lngRecs
6800                dblTmp04 = Nz(.Fields("ICash"), 0)
6810                If ((dblTmp04 > 0.009) Or (dblTmp04 < -0.009)) Then
6820                  dblTmp04 = dbl_Round(dblTmp04, 3)  ' ** Module Function: modUtilities.
6830                Else
6840                  dblTmp04 = 0#
6850                End If
6860                If dblTmp04 <> 0# Then
6870                  dblTmp05 = Nz(.Fields("shareface"), 0)
6880                  If ((dblTmp05 > 0.00009) Or (dblTmp05 < -0.00009)) Then
6890                    dblTmp05 = dbl_Round(dblTmp05, 5)  ' ** Module Function: modUtilities.
6900                  Else
6910                    dblTmp05 = 0#
6920                    dblTmp04 = 0#  ' ** Also zero-out the ICash.
6930                  End If
6940                End If
6950                .Fields("ICash") = dblTmp04
6960                .Fields("shareface") = dblTmp05
6970                .Update
6980                dblSumICash = dblSumICash + dblTmp04
6990                dblSumShareface = dblSumShareface + dblTmp05
7000                dblTmp04 = 0#: dblTmp05 = 0#
7010                If lngX < lngRecs Then .MoveNext
7020              Next
7030            End With
7040          End If  ' ** blnSkip.
              ' ################################################################################

7050          DoEvents
7060          Set dbs = CurrentDb

              ' ** Step 7: Sum holdings.
7070          .Map_Status_lbl.Caption = " 7 of 22: Sum holdings . . ."
7080          DoEvents

7090          Select Case intInputType
              Case 1  ' ** User entered an Amount.
                ' ** qryMap_Misc_STCGL_09a (qryMap_Misc_STCGL_08a (Journal Map, with qryMap_Misc_STCGL_07
                ' ** (ActiveAssets, grouped and summed, by assetno, by specified [astno]), with calc'd fields,
                ' ** by specified [astno], [amt], [prec]), grouped and summed by account), with .._rnd rounded again.
7100            Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_10a")
7110            With qdf.Parameters
7120              ![astno] = lngAssetNo
7130              ![amt] = dblAmount
7140              ![prec] = gintShareFaceDecimals
7150            End With
7160          Case 2  ' ** User entered a PerShare.
                ' ** qryMap_Misc_STCGL_09b (qryMap_Misc_STCGL_08b (Journal Map, with qryMap_Misc_STCGL_07
                ' ** (ActiveAssets, grouped and summed, by assetno, by specified [astno]), with calc'd fields,
                ' ** by specified [astno], [persh], [prec]), grouped and summed by account), with .._rnd rounded again.
7170            Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_10b")
7180            With qdf.Parameters
7190              ![astno] = lngAssetNo
7200              ![persh] = dblPerShare
7210              ![prec] = gintShareFaceDecimals
7220            End With
7230          End Select
7240          Set rst = qdf.OpenRecordset
7250          With rst
7260            If .BOF = True And .EOF = True Then
                  ' ** Something's seriously wrong!
7270              blnContinue = False
7280            Else
7290              .MoveLast
7300              lngMaps = .RecordCount
7310              .MoveFirst
7320              arr_varMap = .GetRows(lngMaps)
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
                  ' **     6       5     ICash            M_ICASH        ' ** Total ICash by account: ([shareface]*([amt]/[sumsf])).
                  ' **     7       6     ICash_rnd        M_ICASHR       ' ** Total ICash by account, rounded to 2 decimal places.
                  ' **
                  ' ***************************************************
7330            End If
7340            .Close
7350          End With
7360          Set rst = Nothing
7370          Set qdf = Nothing
7380          DoEvents

              ' ** Step 8: Rounding Stage 1.
7390          .Map_Status_lbl.Caption = " 8 of 22: Rounding Stage 1 . . ."
7400          DoEvents

7410          Set rst = dbs.OpenRecordset("journal map", dbOpenDynaset, dbConsistent)
7420          With rst
7430            If .BOF = True And .EOF = True Then
                  ' ** Might be coming from GoToReport.
7440              blnContinue = False
7450            Else
7460              .MoveLast
7470              lngRecs = .RecordCount
7480              .MoveFirst
7490              For lngX = 1& To lngRecs
7500                If InStr(![accountno], strAccountNo_OU) = 0 Then  ' ** This is excluded from the arr_varMap() query.
7510                  blnFound = False
7520                  For lngY = 0& To (lngMaps - 1&)
7530                    If arr_varMap(M_ACTNO, lngY) = ![accountno] And arr_varMap(M_ASTNO, lngY) = ![assetno] Then
7540                      blnFound = True
7550                      .Edit
7560                      dblTmp04 = Round(![shareface], gintShareFaceDecimals)
7570                      ![shareface] = dblTmp04
7580                      ![ICash] = CCur(Round((dblTmp04 * arr_varMap(M_PER, lngY)), 2))
7590                      ![pershare] = arr_varMap(M_PER, lngY)
7600                      .Update
7610                      Exit For
7620                    End If
7630                  Next
7640                  If blnFound = False Then
7650                    blnContinue = False
7660                    Exit For
7670                  End If
7680                End If
7690                DoEvents
7700                If lngX < lngRecs Then .MoveNext
7710              Next
7720            End If  ' ** BOF, EOF.
7730            .Close
7740          End With
7750          Set rst = Nothing
7760          DoEvents

7770          If blnContinue = True Then

                ' ** Step 9: Delete under-threshold entries.
7780            .Map_Status_lbl.Caption = " 9 of 22: Delete under-threshold entries . . ."
7790            DoEvents

                ' ** Delete qryMap_Misc_STCGL_11a (Journal Map, for shareface < 0.00006
                ' ** and > -0.00006, icash < 0.006 and > -0.006, by specified [astno]).
7800            Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_11b")
7810            With qdf.Parameters
7820              ![astno] = lngAssetNo
7830            End With
7840            qdf.Execute
7850            Set qdf = Nothing
7860            DoEvents

                ' ** Step 10: Sum distributions.
7870            .Map_Status_lbl.Caption = "10 of 22: Sum distributions . . ."
7880            DoEvents

7890            Select Case strCallingForm
                Case "frmJournal"
                  ' ** qryMap_Misc_STCGL_12c (Journal Map, linked to Account, qryMap_Misc_STCGL_12b
                  ' ** (qryMap_Misc_STCGL_12a (Journal Map, grouped by accountno, assetno, Location_ID),
                  ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max),
                  ' ** grouped and summed, by accountno), grouped and summed, by assetno, by specified [astno].
7900              Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_12d")
7910            Case "frmJournal_Columns"
                  ' ** qryMap_Misc_STCGL_62c (Journal Map, linked to Account, qryMap_Misc_STCGL_62b
                  ' ** (qryMap_Misc_STCGL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
                  ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max)),
                  ' ** grouped and summed, by accountno), grouped and summed, by assetno, by specified [astno].
7920              Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_62d")
7930            End Select
7940            With qdf.Parameters
7950              ![astno] = lngAssetNo
7960            End With
7970            Set rst = qdf.OpenRecordset
7980            With rst
7990              .MoveFirst
8000              dblSumICash = ![ICash]
8010              dblSumShareface = ![shareface]
8020              dblPerShare = ![pershare]  ' ** Max([pershare]), though they should all be identical.
8030              .Close
8040            End With
8050            Set rst = Nothing
8060            Set qdf = Nothing
8070            DoEvents

8080            dblTmp04 = (Round(dblAmount, 2) - dblSumICash)
8090            If ((dblTmp04 > 0.009) Or (dblTmp04 < -0.009)) And (dblTmp04 > 1#) Then
                  ' ** There's an Over/Under of more than $1.00, so let's see if we can bring it down.

                  ' ** Step 11: Rounding Stage 2.
8100              .Map_Status_lbl.Caption = "11 of 22: Rounding Stage 2 . . ."
8110              DoEvents

8120              Select Case intInputType
                  Case 1  ' ** User entered an Amount.
                    ' ** qryMap_Misc_STCGL_15a (Journal Map, linked to qryMap_Misc_STCGL_14a (Journal Map,
                    ' ** linked to qryMap_Misc_STCGL_13a (qryMap_Misc_STCGL_10a (qryMap_Misc_STCGL_09a
                    ' ** (qryMap_Misc_STCGL_08a (Journal Map, with qryMap_Misc_STCGL_07 (ActiveAssets,
                    ' ** grouped and summed, by assetno, by specified [astno]), with calc'd fields, by
                    ' ** specified [astno], [amt], [prec]), grouped and summed by account), with .._rnd
                    ' ** rounded again), just needed fields, by specified [thrsh]), with qryMap_Misc_STCGL_07
                    ' ** (ActiveAssets, grouped and summed, by assetno, by specified [astno]), with
                    ' ** icash_adj, specified [astno], [thrsh], [pcashou]; Cartesian), with pcash_adjx),
                    ' ** just entries with adjustment.
8130                Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_16a")
8140                With qdf.Parameters
8150                  ![astno] = lngAssetNo
8160                  ![amt] = dblAmount
8170                  ![prec] = gintShareFaceDecimals
8180                  dblTmp05 = (0.01 / (dblTmp04 / dblShareface))  ' ** Threshhold needed for a penny adjustment.
8190                  ![thrsh] = dblTmp05
8200                  ![ICashou] = dblTmp04
8210                End With
8220              Case 2  ' ** User entered a PerShare.
                    ' ** qryMap_Misc_STCGL_15b (Journal Map, linked to qryMap_Misc_STCGL_14b (Journal Map,
                    ' ** linked to qryMap_Misc_STCGL_13b (qryMap_Misc_STCGL_10b (qryMap_Misc_STCGL_09b
                    ' ** (qryMap_Misc_STCGL_08b (Journal Map, with qryMap_Misc_STCGL_07 (ActiveAssets,
                    ' ** grouped and summed, by assetno, by specified [astno]), with calc'd fields, by
                    ' ** specified [astno], [persh], [prec]), grouped and summed by account), with .._rnd
                    ' ** rounded again), just needed fields, by specified [thrsh]), with qryMap_Misc_STCGL_07
                    ' ** (ActiveAssets, grouped and summed, by assetno, by specified [astno]), with
                    ' ** icash_adj, by specified [astno], [thrsh], [pcashou]; Cartesian), with pcash_adjx),
                    ' ** just entries with adjustment.
8230                Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_16b")
8240                With qdf.Parameters
8250                  ![astno] = lngAssetNo
8260                  ![persh] = dblPerShare
8270                  ![prec] = gintShareFaceDecimals
8280                  dblTmp05 = (0.01 / (dblTmp04 / dblShareface))  ' ** Threshhold needed for a penny adjustment.
8290                  ![thrsh] = dblTmp05
8300                  ![ICashou] = dblTmp04
8310                End With
8320              End Select
8330              DoEvents

                  ' ** Step 12: Collect adjustments.
8340              .Map_Status_lbl.Caption = "12 of 22: Collect adjustments . . ."
8350              DoEvents

8360              lngAdjs = 0&
8370              Set rst = qdf.OpenRecordset
8380              With rst
8390                If .BOF = True And .EOF = True Then
                      ' ** Not enough to adjust.
8400                Else
8410                  .MoveLast
8420                  lngAdjs = .RecordCount
8430                  .MoveFirst
8440                  arr_varAdj = .GetRows(lngAdjs)
                      ' **********************************************
                      ' ** Array: arr_varAdj()
                      ' **
                      ' **   Field  Element  Name         Constant
                      ' **   =====  =======  ===========  ==========
                      ' **     1       0     ID           A_ID
                      ' **     2       1     accountno    A_ACTNO
                      ' **     3       2     assetno      A_ASTNO
                      ' **     4       3     ICash        A_ICash
                      ' **     5       4     ICash_adj    A_PCADJ
                      ' **
                      ' **********************************************
8450                End If
8460                .Close
8470              End With
8480              Set rst = Nothing
8490              Set qdf = Nothing
8500              DoEvents  ' ** All these DoEvents are to keep the window from showing a 'Not Responding'!

                  ' ** Step 13: Update Journal Map.
8510              .Map_Status_lbl.Caption = "13 of 22: Update temporary table . . ."
8520              DoEvents

8530              Set rst = dbs.OpenRecordset("Journal Map", dbOpenDynaset, dbConsistent)
8540              With rst
8550                .MoveFirst
8560                For lngX = 0& To (lngAdjs - 1&)
8570                  .FindFirst "[ID] = " & CStr(arr_varAdj(A_ID, lngX))
8580                  If .NoMatch = False Then
8590                    .Edit
8600                    ![ICash] = arr_varAdj(A_PCADJ, lngX)
8610                    .Update
8620                  Else
8630                    blnContinue = False
8640                    Exit For
8650                  End If
8660                Next
8670              End With
8680              Set rst = Nothing
8690              Set qdf = Nothing
8700              DoEvents

                  ' ** Step 14: Confirm Rounding Stage 2.
8710              .Map_Status_lbl.Caption = "14 of 22: Confirm Rounding Stage 2 . . ."
8720              DoEvents

8730              Select Case strCallingForm
                  Case "frmJournal"
                    ' ** Now check the Over/Under again!
                    ' ** qryMap_Misc_STCGL_12c (Journal Map, linked to Account, qryMap_Misc_STCGL_12b
                    ' ** (qryMap_Misc_STCGL_12a (Journal Map, grouped by accountno, assetno, Location_ID),
                    ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max),
                    ' ** grouped and summed, by accountno), grouped and summed, by assetno, by specified [astno].
8740                Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_12d")
8750              Case "frmJournal_Columns"
                    ' ** qryMap_Misc_STCGL_62c (Journal Map, linked to Account, qryMap_Misc_STCGL_62b
                    ' ** (qryMap_Misc_STCGL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
                    ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max)),
                    ' ** grouped and summed, by accountno), grouped and summed, by assetno, by specified [astno]
8760                Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_62d")
8770              End Select
8780              With qdf.Parameters
8790                ![astno] = lngAssetNo
8800              End With
8810              Set rst = qdf.OpenRecordset
8820              With rst
8830                .MoveFirst
8840                dblSumICash = ![ICash]
8850                dblSumShareface = ![shareface]
8860                dblPerShare = ![pershare]  ' ** Max([pershare]), though they should all be identical.
8870                .Close
8880              End With
8890              Set rst = Nothing
8900              Set qdf = Nothing
8910              DoEvents

8920              dblTmp04 = (Round(dblAmount, 2) - dblSumICash)

8930            End If

                ' ** Step 15: Create Over/Under entry.
8940            .Map_Status_lbl.Caption = "15 of 22: Create Over/Under entry . . ."
8950            DoEvents

8960            If ((dblTmp04 > 0.009) Or (dblTmp04 < -0.009)) Then
                  ' ** Create a new record in Journal Map.
8970              Set rst = dbs.OpenRecordset("journal map", dbOpenDynaset, dbAppendOnly)
8980              With rst
8990                .AddNew
9000                ![assetno] = lngAssetNo
9010                ![accountno] = strAccountNo_OU
9020                ![shareface] = Round(Abs(dblShareface - dblSumShareface), gintShareFaceDecimals)
9030                ![pershare] = dblPerShare
9040                ![transdate] = datTransDate
9050                ![journaltype] = strMapType
9060                dblOverUnder = Round((Round(dblAmount, 2) - dblSumICash), 2)  ' ** Module Function: modUtilities.
9070                ![ICash] = CCur(Round(dblOverUnder, 2))
9080                ![description] = "Over/Short of " & strMapType & " Mapping for " & Me.cmbAssets.Column(CBX_A_DESC)
9090                ![curr_id] = lngCurrID
9100                .Update
9110                .Close
9120              End With
9130              Set rst = Nothing
9140              .overshort = dblOverUnder
9150            Else
9160              .overshort = 0#
9170            End If

9180            dbs.Close
9190            Set dbs = Nothing
9200            DoEvents

                ' ** Step 16: Delete zero entries.
9210            .Map_Status_lbl.Caption = "16 of 22: Delete zero entries . . ."
9220            DoEvents

                ' ################################################################################
9230            If blnSkip = False Then
9240              Set dbs = CurrentDb
                  ' ** Journal Map, just zero shareface, icash, pCash.
9250              Set qdf = dbs.QueryDefs("qryMap_Div_01a")
9260              Set rst = qdf.OpenRecordset
9270              If rst.BOF = True And rst.EOF = True Then
                    ' ** All's well.
9280                rst.Close
9290              Else
9300                rst.Close
                    ' ** Delete qryMap_Div_01a (Journal Map, just zero shareface, icash, ICash).
9310                Set qdf = dbs.QueryDefs("qryMap_Div_01b")
9320                qdf.Execute
9330              End If
9340              dbs.Close
9350            End If  ' ** blnSkip.
                ' ################################################################################

9360            dblOverUnder = AdvancedRounding(lngAssetNo, dblPerShare)  ' ** Procedure: Below.
9370            If dblOverUnder = 0# And .overshort <> 0 Then
                  ' ** Not sure why this disagrees.
9380              dblOverUnder = .overshort
9390            End If

                ' ** Step 22: Display detail.
9400            .Map_Status_lbl.Caption = "22 of 22: Display detail . . ."
9410            DoEvents

                ' ** Borrowing these variables from the Court Reports.
9420            gstrCrtRpt_Ordinal = .cmbAssets.Column(CBX_A_CUSIP)  ' ** cusip.
9430            strTmp01 = .cmbAssets.Column(CBX_A_TDESC)            ' ** totdesc.
9440            intPos01 = InStr(strTmp01, Chr(160))  ' ** Hard space.
9450            If intPos01 > 0 Then
9460              strTmp01 = Mid(strTmp01, (intPos01 + 1))
9470              Do While Left(strTmp01, 1) = Chr(160)
9480                strTmp01 = Mid(strTmp01, 2)
9490              Loop
9500            End If
9510            gstrCrtRpt_Version = strTmp01

9520            strDocName = "frmMap_Misc_STCGL_Detail"
9530            DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & CStr(lngAssetNo) & "~" & _
                  CStr(dblOverUnder) & "~" & CStr(dblPerShare) & "~" & Format(datTransDate, "mm/dd/yyyy") & "~" & _
                  CStr(dblAmount) & "~" & CStr(lngCurrID)

9540            DoCmd.Close acForm, THIS_NAME

9550          Else
9560            Select Case gblnGoToReport
                Case True
9570              Beep
9580              DoCmd.Hourglass False
9590              MsgBox "The report site cannot be shown because there are no records.", _
                    vbInformation + vbOKOnly, "Record Not Found"
9600              blnGoingToReport = False
9610              .GoToReport_arw_ok_img.Visible = False
9620              gblnGoToReport = False
9630            Case False
9640              Beep
9650              DoCmd.Hourglass False
9660              MsgBox "This procedure cannot continue." & vbCrLf & "Expected records were not found.", _
                    vbInformation + vbOKOnly, "Record Not Found"
9670              blnGoingToReport = False
9680              .GoToReport_arw_ok_img.Visible = False
9690              gblnGoToReport = False
9700            End Select
9710            .cmdCancel.SetFocus
9720          End If  ' ** blnContinue.

9730        Else
9740          .cmdCancel.SetFocus
9750        End If  ' ** msgResponse.

9760      End If

9770    End With

EXITP:
9780    Set rsx1 = Nothing
9790    Set cnxn = Nothing
9800    Set rst = Nothing
9810    Set qdf = Nothing
9820    Set dbs = Nothing
9830    Exit Sub

ERRH:
9840    DoCmd.Hourglass False
9850    THAT_PROC = THIS_PROC
9860    That_Erl = Erl
9870    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9880    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9920    intShiftDown = (Shift And acShiftMask) > 0
9930    intAltDown = (Shift And acAltMask) > 0
9940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9960      Select Case intRetVal
          Case vbKeyTab
9970        With Me
9980          intRetVal = 0
9990          .cmdCancel.SetFocus
10000       End With
10010     End Select
10020   End If

        ' ** Shift keys.
10030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10040     Select Case intRetVal
          Case vbKeyTab
10050       With Me
10060         intRetVal = 0
10070         .transdate.SetFocus
10080       End With
10090     End Select
10100   End If

        ' ** Ctrl-Shift keys.
10110   If intCtrlDown And (Not intAltDown) And intShiftDown Then
10120     Select Case intRetVal
          Case vbKeyTab
10130       With Me
10140         intRetVal = 0
10150         .cmdCalendar.SetFocus
10160       End With
10170     End Select
10180   End If

EXITP:
10190   KeyCode = intRetVal
10200   Exit Sub

ERRH:
10210   intRetVal = 0
10220   THAT_PROC = THIS_PROC
10230   That_Erl = Erl
10240   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10250   Resume EXITP

End Sub

Private Sub cmdCancel_Click()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

10310   DoCmd.Hourglass True
10320   DoEvents

10330   blnCancel = True

10340   Set dbs = CurrentDb
10350   With dbs
10360     Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
10370       Set qdf = .QueryDefs("qryMap_01")
10380     Case "frmJournal_Columns"
            ' ** Empty Journal Map.
10390       Set qdf = .QueryDefs("qryJournal_Columns_35_04")
10400     End Select
10410     qdf.Execute
10420     Set qdf = Nothing
10430     DoEvents
          ' ** Empty tmpAccount.
10440     Set qdf = .QueryDefs("qryMap_Misc_STCGL_08_01")
10450     qdf.Execute
10460     Set qdf = Nothing
10470     DoEvents
10480     .Close
10490   End With
10500   Set dbs = Nothing
10510   DoEvents

10520   gblnSetFocus = True
10530   Select Case strCallingForm
        Case "frmJournal"
10540     If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
10550       DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
10560     Else
10570       gstrReturningForm = THIS_NAME
10580       DoCmd.SelectObject acForm, strCallingForm, False
10590       Forms(strCallingForm).TimerInterval = 250&
10600     End If
10610   Case "frmJournal_Columns"
10620     If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
10630       DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
10640     Else
10650       gstrReturningForm = "frmMap_Return"
10660       DoCmd.SelectObject acForm, strCallingForm, False
10670       Forms(strCallingForm).TimerInterval = 250&
10680     End If
10690   End Select

10700   DoCmd.Close acForm, THIS_NAME

EXITP:
10710   Set qdf = Nothing
10720   Set dbs = Nothing
10730   Exit Sub

ERRH:
10740   DoCmd.Hourglass False
10750   THAT_PROC = THIS_PROC
10760   That_Erl = Erl
10770   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10780   Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
10890         .cmbAssets.SetFocus
10900       End With
10910     End Select
10920   End If

        ' ** Shift keys.
10930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10940     Select Case intRetVal
          Case vbKeyTab
10950       With Me
10960         intRetVal = 0
10970         .cmdOK.SetFocus
10980       End With
10990     End Select
11000   End If

EXITP:
11010   KeyCode = intRetVal
11020   Exit Sub

ERRH:
11030   intRetVal = 0
11040   THAT_PROC = THIS_PROC
11050   That_Erl = Erl
11060   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11070   Resume EXITP

End Sub

Private Sub opgAssetSource_AfterUpdate()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_AfterUpdate"

11110   With Me
11120     Select Case .opgAssetSource
          Case .opgAssetSource_optType.OptionValue
11130       .opgAssetSource_optType_lbl.FontBold = True
11140       .opgAssetSource_optName_lbl.FontBold = False
11150       .opgAssetSource_optCusip_lbl.FontBold = False
            ' ** qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to
            ' ** MasterAsset, for Dividend = True, shareface <> 0; by assettype.
11160       If .cmbAssets.RowSource <> "qryMap_Rec_05a" Then
11170         .cmbAssets.RowSource = "qryMap_Rec_05a"
11180       End If
11190       .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
11200       .cmbAssets.ListWidth = 7890
11210     Case .opgAssetSource_optName.OptionValue
11220       .opgAssetSource_optType_lbl.FontBold = False
11230       .opgAssetSource_optName_lbl.FontBold = True
11240       .opgAssetSource_optCusip_lbl.FontBold = False
            ' ** qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to
            ' ** MasterAsset, for Dividend = True, shareface <> 0; by totdesc.
11250       If .cmbAssets.RowSource <> "qryMap_Rec_05d" Then
11260         .cmbAssets.RowSource = "qryMap_Rec_05d"
11270       End If
11280       .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
11290       .cmbAssets.ListWidth = 7890
11300     Case .opgAssetSource_optCusip.OptionValue
11310       .opgAssetSource_optType_lbl.FontBold = False
11320       .opgAssetSource_optName_lbl.FontBold = False
11330       .opgAssetSource_optCusip_lbl.FontBold = True
            ' ** qryMap_Rec_04 (MasterAsset, linked to AssetType), linked to
            ' ** MasterAsset, with qryMap_Rec_05k (qryMap_Rec_05j (MasterAsset,
            ' ** with cusip_len), grouped, with Max(cusip_len)), for
            ' ** Dividend = True, shareface <> 0; by cusip; Cartesian.
11340       If .cmbAssets.RowSource <> "qryMap_Rec_05g" Then
11350         .cmbAssets.RowSource = "qryMap_Rec_05g"
11360       End If
11370       .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
11380       .cmbAssets.ListWidth = 7890
11390     End Select
11400     .cmbAssets.Requery
11410   End With

EXITP:
11420   Exit Sub

ERRH:
11430   Select Case ERR.Number
        Case Else
11440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11450   End Select
11460   Resume EXITP

End Sub

Private Sub opgAssetSource_optType_KeyDown(KeyCode As Integer, Shift As Integer)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optType_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
11570       With Me
11580         intRetVal = 0
11590         .cmbAssets.SetFocus
11600       End With
11610     End Select
11620   End If

        ' ** Shift keys.
11630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11650       With Me
11660         intRetVal = 0
11670         .cmdCancel.SetFocus
11680       End With
11690     End Select
11700   End If

EXITP:
11710   KeyCode = intRetVal
11720   Exit Sub

ERRH:
11730   intRetVal = 0
11740   Select Case ERR.Number
        Case Else
11750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11760   End Select
11770   Resume EXITP

End Sub

Private Sub opgAssetSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11820   intShiftDown = (Shift And acShiftMask) > 0
11830   intAltDown = (Shift And acAltMask) > 0
11840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11870       With Me
11880         intRetVal = 0
11890         .cmbAssets.SetFocus
11900       End With
11910     End Select
11920   End If

        ' ** Shift keys.
11930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11950       With Me
11960         intRetVal = 0
11970         .cmdCancel.SetFocus
11980       End With
11990     End Select
12000   End If

EXITP:
12010   KeyCode = intRetVal
12020   Exit Sub

ERRH:
12030   intRetVal = 0
12040   Select Case ERR.Number
        Case Else
12050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12060   End Select
12070   Resume EXITP

End Sub

Private Sub opgAssetSource_optCusip_KeyDown(KeyCode As Integer, Shift As Integer)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optCusip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12120   intShiftDown = (Shift And acShiftMask) > 0
12130   intAltDown = (Shift And acAltMask) > 0
12140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12170       With Me
12180         intRetVal = 0
12190         .cmbAssets.SetFocus
12200       End With
12210     End Select
12220   End If

        ' ** Shift keys.
12230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12250       With Me
12260         intRetVal = 0
12270         .cmdCancel.SetFocus
12280       End With
12290     End Select
12300   End If

EXITP:
12310   KeyCode = intRetVal
12320   Exit Sub

ERRH:
12330   intRetVal = 0
12340   Select Case ERR.Number
        Case Else
12350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12360   End Select
12370   Resume EXITP

End Sub

Private Sub cmbAssets_AfterUpdate()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_AfterUpdate"

12410   With Me

12420     If blnCurrID = True Then
12430       lngCurrID = 0&
12440       If IsNull(.cmbAssets) = False Then
12450         If IsNull(.cmbAssets.Column(CBX_A_CURRID)) = False Then
12460           lngCurrID = .cmbAssets.Column(CBX_A_CURRID)
12470           .curr_id = lngCurrID
12480         End If
12490       End If
12500     End If

12510     .amount = Null
12520     .amount.Enabled = True
12530     .amount.Locked = False
12540     .amount.ForeColor = CLR_BLK
12550     .pershare = Null
12560     .pershare.Enabled = True
12570     .pershare.Locked = False
12580     .pershare.ForeColor = CLR_BLK
12590     If lngCurrID = 150& Then
12600       .amount.BackColor = CLR_WHT
12610       .pershare.BackColor = CLR_WHT
12620     Else
12630       .amount.BackColor = CLR_VLTGRN
12640       .pershare.BackColor = CLR_VLTGRN
12650     End If
12660     blnAssetUpdated = True: blnAssetChecked = False
12670     .amount.SetFocus

12680     curr_id_AfterUpdate  ' ** Procedure: Below.

12690   End With

EXITP:
12700   Exit Sub

ERRH:
12710   THAT_PROC = THIS_PROC
12720   That_Erl = Erl
12730   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12740   Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12820   intShiftDown = (Shift And acShiftMask) > 0
12830   intAltDown = (Shift And acAltMask) > 0
12840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12870       With Me
12880         intRetVal = 0
              ' ** This enabling and disabling is causing all sorts of ...
12890         If .amount.Enabled = True Then
12900 On Error Resume Next
12910           .amount.SetFocus
12920 On Error GoTo ERRH
12930         Else
12940 On Error Resume Next
12950           .pershare.SetFocus
12960 On Error GoTo ERRH
12970         End If
12980       End With
12990     End Select
13000   End If

        ' ** Shift keys.
13010   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13020     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13030       With Me
13040         intRetVal = 0
13050         .cmdCancel.SetFocus
13060       End With
13070     End Select
13080   End If

        ' ** Ctrl keys.
13090   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
13100     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13110       With Me
13120         intRetVal = 0
13130         .opgAssetSource.SetFocus
13140       End With
13150     End Select
13160   End If

EXITP:
13170   KeyCode = intRetVal
13180   Exit Sub

ERRH:
13190   intRetVal = 0
13200   THAT_PROC = THIS_PROC
13210   That_Erl = Erl
13220   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13230   Resume EXITP

End Sub

Private Sub cmbAssets_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_MouseMove"

13310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
13320   Exit Sub

ERRH:
13330   Select Case ERR.Number
        Case Else
13340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13350   End Select
13360   Resume EXITP

End Sub

Private Sub cmbAssets_LostFocus()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_LostFocus"

13410   With Me
13420     If IsNull(.cmbAssets) = False Then
13430       If .cmbAssets > 0 And blnAssetUpdated = True And blnAssetChecked = False Then
13440         blnExitingAssetNo = True
13450         .TimerInterval = 100&
13460       End If
13470     End If
13480   End With  ' ** Me.

EXITP:
13490   Exit Sub

ERRH:
13500   THAT_PROC = THIS_PROC
13510   That_Erl = Erl
13520   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13530   Resume EXITP

End Sub

Private Sub curr_id_cmd_DblClick(Cancel As Integer)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_cmd_DblClick"

13610   With Me
13620     If IsNull(.cmbAssets) = False Then
13630       Select Case .curr_date.Visible
            Case True
13640         .curr_date.Visible = False
13650       Case False
13660         .curr_date = .curr_id.Column(CBX_C_DATE)
13670         .curr_date.Visible = True
13680       End Select
13690     End If
13700   End With

EXITP:
13710   Exit Sub

ERRH:
13720   Select Case ERR.Number
        Case Else
13730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
13740   End Select
13750   Resume EXITP

End Sub

Private Sub curr_id_AfterUpdate()

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_AfterUpdate"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String, lngTmp02 As Long
        Dim lngX As Long

13810   With Me

13820     If IsNull(.curr_id) = False Then

13830       lngCurrID = .curr_id.Column(CBX_C_CURRID)

13840       blnFound = False
13850       If lngCurrID <> 150& Then
13860         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
13870           NoChar_Load  ' ** Procedure: Below.
13880         End If
13890         For lngX = 0& To (lngNoChars - 1&)
13900           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
13910             blnFound = True
13920           End If
13930         Next
13940       End If

13950       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
13960         .amount.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
13970         .amount.FontSize = 10
13980         .amount.FontBold = True
13990         .pershare.FontName = "Courier New"
14000         .pershare.FontSize = 10
14010         .pershare.FontBold = True
14020       Case False
14030         If .amount.FontName <> "Arial" Then
14040           .amount.FontName = "Arial"
14050           .amount.FontSize = 10
14060           .amount.FontBold = False
14070           .pershare.FontName = "Arial"
14080           .pershare.FontSize = 10
14090           .pershare.FontBold = False
14100         End If
14110       End Select

14120       If lngCurrID = 150& Then  ' ** USD.
14130         strFrmt1 = "Currency"
14140         intDec1 = 2
14150         strFrmt2 = "Currency"
14160         intDec2 = 5
14170         lngBkClr1 = CLR_VLTRED
14180         lngBkClr2 = CLR_LTRED
14190         lngTmp02 = CLR_WHT
14200       Else
14210         strTmp01 = .curr_id.Column(CBX_C_SYM)  ' ** Symbol.
14220         If .curr_id.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
14230           intDec1 = 4
14240         Else
14250           intDec1 = .curr_id.Column(CBX_C_DEC)  ' ** Decimal places.
14260         End If
14270         intDec2 = 5
14280         intLen = Len(strTmp01)
14290         Select Case intLen
              Case 1
                'strTmp01 = strTmp01
14300         Case 2
14310           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2)
14320         Case 3
14330           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3)
14340         Case 4
14350           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & "\" & Mid(strTmp01, 4)
14360         Case 5
14370           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & _
                  "\" & Mid(strTmp01, 4, 1) & "\" & Mid(strTmp01, 5)
14380         End Select
14390         Select Case intDec1
              Case 0
14400           strFrmt1 = strTmp01 & "#,##0;(" & strTmp01 & "#,##0);" & strTmp01 & "0;" & strTmp01 & "0"
14410           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
14420         Case 1  ' ** None have this currently.
14430           strFrmt1 = strTmp01 & "#,##0.0;(" & strTmp01 & "#,##0.0);" & strTmp01 & "0.0;" & strTmp01 & "0.0"
14440           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
14450         Case 2
14460           strFrmt1 = strTmp01 & "#,##0.00;(" & strTmp01 & "#,##0.00);" & strTmp01 & "0.00;" & strTmp01 & "0.00"
14470           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
14480         Case 3
14490           strFrmt1 = strTmp01 & "#,##0.000;(" & strTmp01 & "#,##0.000);" & strTmp01 & "0.000;" & strTmp01 & "0.000"
14500           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
14510         End Select
14520         lngBkClr1 = CLR_VLTPRP
14530         lngBkClr2 = CLR_LTPRP
14540         lngTmp02 = CLR_VLTGRN
              'CLR_DISABLED_BG_R = CLR_LTRED     '14013951
              'CLR_DISABLED_BG_VLR = CLR_VLTRED  '15921919
              'CLR_LTGRN   '15138802
              'CLR_VLTGRN  '16056314
              'CLR_LTPRP   '16771829
              'CLR_VLTPRP  '16774907
14550       End If
14560       .amount.Format = strFrmt1
14570       .amount.DecimalPlaces = intDec1
14580       Select Case IsNull(.amount)
            Case True
14590         .amount.BackColor = CLR_WHT
14600       Case False
14610         If .amount >= 0 Then
14620           .amount.BackColor = CLR_WHT
14630         Else
14640           .amount.BackColor = lngBkClr1
14650         End If
14660       End Select
14670       .pershare.Format = strFrmt2
14680       .pershare.DecimalPlaces = intDec2
14690       Select Case IsNull(.pershare)
            Case True
14700         .pershare.BackColor = CLR_WHT
14710       Case False
14720         If .pershare >= 0 Then
14730           .pershare.BackColor = CLR_WHT
14740         Else
14750           .pershare.BackColor = lngTmp02
14760         End If
14770       End Select

14780       If lngCurrID = 150& Then
14790         .Amount_usd.Visible = False
14800         .pershare_usd.Visible = False
14810       Else
14820         .Amount_usd = 0
14830         .Amount_usd.Visible = True
14840         .pershare_usd = 0
14850         .pershare_usd.Visible = True
14860       End If

14870     End If

14880   End With

EXITP:
14890   Exit Sub

ERRH:
14900   THAT_PROC = THIS_PROC
14910   That_Erl = Erl
14920   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14930   Resume EXITP

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

Private Sub Amount_KeyDown(KeyCode As Integer, Shift As Integer)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "Amount_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15220   intShiftDown = (Shift And acShiftMask) > 0
15230   intAltDown = (Shift And acAltMask) > 0
15240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15270       With Me
15280         intRetVal = 0
15290         If .pershare.Enabled = True Then
15300 On Error Resume Next
15310           .pershare.SetFocus
15320           If ERR.Number <> 0 Then
15330 On Error GoTo ERRH
15340             .transdate.SetFocus
15350           Else
15360 On Error GoTo ERRH
15370           End If
15380         Else
15390           .transdate.SetFocus
15400         End If
15410       End With
15420     End Select
15430   End If

        ' ** Shift keys.
15440   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15450     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15460       With Me
15470         intRetVal = 0
15480         blnDrivingMeBackwards = True
15490         .cmbAssets.SetFocus
15500       End With
15510     End Select
15520   End If

EXITP:
15530   KeyCode = intRetVal
15540   Exit Sub

ERRH:
15550   intRetVal = 0
15560   THAT_PROC = THIS_PROC
15570   That_Erl = Erl
15580   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15590   Resume EXITP

End Sub

Private Sub amount_Exit(Cancel As Integer)

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "amount_Exit"

        Dim dblRate2 As Double
        Dim intPos01 As Integer
        Dim blnContinue As Boolean
        Dim dblTmp01 As Double

15610   With Me

15620     blnContinue = True

15630     lngCurrID = .curr_id.Column(CBX_C_CURRID)

15640     If IsNumeric(.amount) = True Then
15650       If CDbl(Rem_Dollar(.amount.text, lngCurrID)) > 999999999.99 Then  ' ** Module Function: modStringFuncs.
15660         blnContinue = False
15670         MsgBox "Amount is too large.", vbInformation + vbOKOnly, "Invalid Entry"
15680         .amount = vbNullString
15690         Cancel = -1
15700         .amount.SetFocus
15710       Else
15720         intPos01 = InStr(1, .amount, ".", vbTextCompare)
15730         If intPos01 > 0 Then
15740           If Len(.amount) - intPos01 > 2 Then
15750             blnContinue = False
15760             MsgBox "You can only enter Amount with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
15770             .amount = vbNullString
15780             Cancel = -1
15790           End If
15800         End If
15810         If blnContinue = True Then
15820           If Nz(.shareface, 0) = 0 Then
15830             blnContinue = False
15840             MsgBox "Shares must have a value", vbInformation + vbOKOnly, "Invalid Entry"
15850             .shareface = 1
15860             .cmbAssets.SetFocus
15870             If lngCurrID = 150& Then
15880               .pershare = (.amount / IIf(Nz(.shareface, 0) = 0, 1, .shareface))  '5 dec.
15890               .pershare.BackColor = CLR_DISABLED_BG
15900             Else
15910               .pershare = (.amount / IIf(Nz(.shareface, 0) = 0, 1, .shareface))  '5 dec.
15920               .pershare.BackColor = lngBkClr2
15930             End If
15940             .pershare.Locked = True
15950             .pershare.Enabled = False
15960             .pershare.ForeColor = CLR_DISABLED_FG
15970           Else
15980             If blnDrivingMeBackwards = False Then
15990 On Error Resume Next
16000               .cmdOK.SetFocus
16010 On Error GoTo ERRH
16020             Else
16030               blnDrivingMeBackwards = False
16040             End If
16050             intInputType = 1
16060             dblPerShare_Calc = (.amount / IIf(Nz(.shareface, 0) = 0, 1, .shareface))
16070             If lngCurrID = 150& Then
16080               .pershare = dblPerShare_Calc  '5 dec.
16090               If .amount >= 0 Then
16100                 .pershare.BackColor = CLR_DISABLED_BG
16110                 .amount.BackColor = CLR_WHT
16120               Else
16130                 .pershare.BackColor = CLR_DISABLED_BG_R
16140                 .amount.BackColor = CLR_DISABLED_BG_VLR
16150               End If
16160             Else
16170               .pershare = dblPerShare_Calc  '5 dec.
16180               If .amount >= 0 Then
16190                 .pershare.BackColor = CLR_LTGRN
16200                 .amount.BackColor = CLR_VLTGRN
16210               Else
16220                 .pershare.BackColor = lngBkClr2
16230                 .amount.BackColor = lngBkClr1
16240               End If
16250               dblRate2 = .curr_id.Column(CBX_C_RATE2)
16260               dblTmp01 = (.amount * dblRate2)
16270               .Amount_usd = Round(dblTmp01, 2)
16280               .Amount_usd.Visible = True
16290               dblTmp01 = (dblPerShare_Calc * dblRate2)
16300               dblTmp01 = Round(dblTmp01, 5)
16310               .pershare_usd = dblTmp01
16320               .pershare_usd.Visible = True
16330             End If
16340             .pershare.Locked = True
16350             .pershare.Enabled = False
16360             .pershare.ForeColor = CLR_DISABLED_FG
16370           End If
16380         End If
16390       End If
16400     Else
16410       .pershare.Enabled = True
16420       .pershare.Locked = False
16430       .pershare.ForeColor = CLR_BLK
16440       If lngCurrID = 150& Then
16450         .pershare.BackColor = CLR_WHT
16460       Else
16470         .pershare.BackColor = CLR_VLTGRN
16480       End If
16490     End If

16500   End With

EXITP:
16510   Exit Sub

ERRH:
16520   Select Case ERR.Number
        Case Else
16530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16540   End Select
16550   Resume EXITP

End Sub

Private Sub pershare_Exit(Cancel As Integer)

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "pershare_Exit"

        Dim dblRate2 As Double
        Dim intPos01 As Integer
        Dim blnContinue As Boolean
        Dim dblTmp01 As Double

16610   With Me

16620     blnContinue = True

16630     lngCurrID = .curr_id.Column(CBX_C_CURRID)

16640     If IsNumeric(.pershare) = True Then
16650       If CDbl(Rem_Dollar(.pershare.text, lngCurrID)) > 999999999.9999 Then  ' ** Module Function: modStringFuncs.
16660         blnContinue = False
16670         MsgBox "Per Share is too large.", vbInformation + vbOKOnly, "Invalid Entry"
16680         .pershare = vbNullString
16690         Cancel = -1
16700         .pershare.SetFocus
16710       Else
16720         intPos01 = InStr(1, .pershare, ".", vbTextCompare)
16730         If intPos01 > 0 Then
16740           If Len(.pershare) - intPos01 > 5 Then
16750             blnContinue = False
16760             MsgBox "You can only enter Per Share with up to 5 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
16770             .pershare = vbNullString
16780             Cancel = -1
16790             .pershare.SetFocus
16800           End If
16810         End If
16820         If blnContinue = True Then
16830           intInputType = 2
16840           If lngCurrID = 150& Then
16850             .amount = (.pershare * Nz(.shareface, 1))
16860             If .pershare >= 0 Then
16870               .pershare.BackColor = CLR_WHT
16880               .amount.BackColor = CLR_DISABLED_BG
16890             Else
16900               .pershare.BackColor = CLR_DISABLED_BG_VLR
16910               .amount.BackColor = CLR_DISABLED_BG_R
16920             End If
16930           Else
16940             .amount = (.pershare * Nz(.shareface, 1))
16950             If .pershare >= 0 Then
16960               .pershare.BackColor = CLR_VLTGRN
16970               .amount.BackColor = CLR_LTGRN
16980             Else
16990               .pershare.BackColor = lngBkClr1
17000               .amount.BackColor = lngBkClr2
17010             End If
17020             dblRate2 = .curr_id.Column(CBX_C_RATE2)
17030             dblTmp01 = (.pershare * dblRate2)
17040             dblTmp01 = Round(dblTmp01, 5)
17050             .pershare_usd = dblTmp01
17060             .pershare_usd.Visible = True
17070             dblTmp01 = (.amount * dblRate2)
17080             dblTmp01 = Round(dblTmp01, 2)
17090             .Amount_usd = dblTmp01
17100             .Amount_usd.Visible = True
17110           End If
17120           .amount.Locked = True
17130           .amount.Enabled = False
17140           .amount.ForeColor = CLR_DISABLED_FG
17150         End If
17160       End If
17170     Else
17180       .amount.Enabled = True
17190       .amount.Locked = False
17200       .amount.ForeColor = CLR_BLK
17210       If lngCurrID = 150& Then
17220         .amount.BackColor = CLR_WHT
17230       Else
17240         .amount.BackColor = CLR_VLTGRN
17250       End If
17260     End If

17270   End With

EXITP:
17280   Exit Sub

ERRH:
17290   Select Case ERR.Number
        Case Else
17300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17310   End Select
17320   Resume EXITP

End Sub

Private Sub transdate_KeyDown(KeyCode As Integer, Shift As Integer)

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17420   intShiftDown = (Shift And acShiftMask) > 0
17430   intAltDown = (Shift And acAltMask) > 0
17440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17470       With Me
17480         intRetVal = 0
17490         .cmdOK.SetFocus
17500       End With
17510     End Select
17520   End If

        ' ** Shift keys.
17530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17550       With Me
17560         intRetVal = 0
17570         If .pershare.Enabled = True Then
17580           .pershare.SetFocus
17590         Else
17600           .amount.SetFocus
17610         End If
17620       End With
17630     End Select
17640   End If

        ' ** Ctrl keys.
17650   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
17660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17670       With Me
17680         intRetVal = 0
17690         .cmdCalendar.SetFocus
17700       End With
17710     End Select
17720   End If

EXITP:
17730   KeyCode = intRetVal
17740   Exit Sub

ERRH:
17750   intRetVal = 0
17760   THAT_PROC = THIS_PROC
17770   That_Erl = Erl
17780   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17790   Resume EXITP

End Sub

Private Sub transdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_MouseMove"

17810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17820   Exit Sub

ERRH:
17830   Select Case ERR.Number
        Case Else
17840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17850   End Select
17860   Resume EXITP

End Sub

Private Sub transdate_Exit(Cancel As Integer)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_Exit"

        Dim varTmp00 As Variant

17910   With Me
17920 On Error Resume Next
17930     varTmp00 = .transdate.text  ' ** When called from cmdCalendar_Click(), below, it's not always on focus.
17940     If ERR.Number <> 0 Then
17950       varTmp00 = .transdate
17960     End If
17970 On Error GoTo ERRH
17980     If DateCheck_Post(varTmp00) = True Then  ' ** Module Function: modUtilities.
            ' ** '.Text' needed to reliably check date.
17990     Else
18000       If IsNull(.transdate) = True Then
18010         .transdate = Date
18020       Else
18030         If Trim(.transdate) = vbNullString Then
18040           .transdate = Date
18050         Else
18060           Cancel = -1
18070           .transdate.SetFocus
18080         End If
18090       End If
18100     End If
18110   End With

EXITP:
18120   Exit Sub

ERRH:
18130   Select Case ERR.Number
        Case Else
18140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18150   End Select
18160   Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

18210   With Me
18220     blnCalendar1_Focus = True
18230     .cmdCalendar_raised_semifocus_dots_img.Visible = True
18240     .cmdCalendar_raised_img.Visible = False
18250     .cmdCalendar_raised_focus_img.Visible = False
18260     .cmdCalendar_raised_focus_dots_img.Visible = False
18270     .cmdCalendar_sunken_focus_dots_img.Visible = False
18280     .cmdCalendar_raised_img_dis.Visible = False
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

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

18410   With Me
18420     blnCalendar1_MouseDown = True
18430     .cmdCalendar_sunken_focus_dots_img.Visible = True
18440     .cmdCalendar_raised_img.Visible = False
18450     .cmdCalendar_raised_semifocus_dots_img.Visible = False
18460     .cmdCalendar_raised_focus_img.Visible = False
18470     .cmdCalendar_raised_focus_dots_img.Visible = False
18480     .cmdCalendar_raised_img_dis.Visible = False
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

Private Sub cmdCalendar_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

18610   With Me
18620     datStartDate = Date
18630     datEndDate = 0
18640     blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
18650     If blnRetVal = True Then
            ' ** Allow posting up to 1 month into the future.
18660       If datStartDate > DateAdd("m", 1, Date) Then
18670         MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
18680         .transdate = CDate(Format(Date, "mm/dd/yyyy"))
18690       Else
18700         .transdate = datStartDate
18710       End If
18720     Else
18730       .transdate = CDate(Format(Date, "mm/dd/yyyy"))
18740     End If
          ' ** Make the date 'is edited' as if were entered by hand.
18750     .transdate.SetFocus
18760     Cancel = 0
18770     transdate_Exit Cancel  ' ** Procedure: Above.
18780     If Cancel = 0 Then
18790       .cmdOK.SetFocus
18800     End If
18810   End With

EXITP:
18820   Exit Sub

ERRH:
18830   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
18840   Case Else
18850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18860   End Select
18870   Resume EXITP

End Sub

Private Sub cmdCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseMove"

18910   With Me
18920     If blnCalendar1_MouseDown = False Then
18930       Select Case blnCalendar1_Focus
            Case True
18940         .cmdCalendar_raised_focus_dots_img.Visible = True
18950         .cmdCalendar_raised_focus_img.Visible = False
18960       Case False
18970         .cmdCalendar_raised_focus_img.Visible = True
18980         .cmdCalendar_raised_focus_dots_img.Visible = False
18990       End Select
19000       .cmdCalendar_raised_img.Visible = False
19010       .cmdCalendar_raised_semifocus_dots_img.Visible = False
19020       .cmdCalendar_sunken_focus_dots_img.Visible = False
19030       .cmdCalendar_raised_img_dis.Visible = False
19040     End If
19050   End With

EXITP:
19060   Exit Sub

ERRH:
19070   Select Case ERR.Number
        Case Else
19080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19090   End Select
19100   Resume EXITP

End Sub

Private Sub cmdCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseUp"

19210   With Me
19220     .cmdCalendar_raised_focus_dots_img.Visible = True
19230     .cmdCalendar_raised_img.Visible = False
19240     .cmdCalendar_raised_semifocus_dots_img.Visible = False
19250     .cmdCalendar_raised_focus_img.Visible = False
19260     .cmdCalendar_sunken_focus_dots_img.Visible = False
19270     .cmdCalendar_raised_img_dis.Visible = False
19280     blnCalendar1_MouseDown = False
19290   End With

EXITP:
19300   Exit Sub

ERRH:
19310   Select Case ERR.Number
        Case Else
19320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19330   End Select
19340   Resume EXITP

End Sub

Private Sub cmdCalendar_LostFocus()

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_LostFocus"

19410   With Me
19420     .cmdCalendar_raised_img.Visible = True
19430     .cmdCalendar_raised_semifocus_dots_img.Visible = False
19440     .cmdCalendar_raised_focus_img.Visible = False
19450     .cmdCalendar_raised_focus_dots_img.Visible = False
19460     .cmdCalendar_sunken_focus_dots_img.Visible = False
19470     .cmdCalendar_raised_img_dis.Visible = False
19480     blnCalendar1_Focus = False
19490   End With

EXITP:
19500   Exit Sub

ERRH:
19510   Select Case ERR.Number
        Case Else
19520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19530   End Select
19540   Resume EXITP

End Sub

Private Function AdvancedRounding(lngAssetNo As Long, dblPerShare As Double) As Double

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "AdvancedRounding"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim lngRecs As Long
        Dim lngLocationID As Long
        Dim blnNoDistribution As Boolean
        Dim lngX As Long
        Dim dblRetVal As Double

19610   blnNoDistribution = False
19620   dblRetVal = 0#

19630   Set dbs = CurrentDb
19640   With dbs

          ' ** Step 17: Rounding Stage 3.
19650     Me.Map_Status_lbl.Caption = "17 of 22: Rounding Stage 3 . . ."
19660     DoEvents

19670     Select Case strCallingForm
          Case "frmJournal"
            ' ** qryMap_Misc_STCGL_12c (Journal Map, linked to Account, qryMap_Misc_STCGL_12b
            ' ** (qryMap_Misc_STCGL_12a (Journal Map, grouped by accountno, assetno, Location_ID),
            ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max),
            ' ** grouped and summed, by accountno), just 'INCOME O/U', linked to qryMap_Misc_STCGL_17b
            ' ** (qryMap_Misc_STCGL_17a (qryMap_Misc_STCGL_12c (Journal Map, linked to Account,
            ' ** qryMap_Misc_STCGL_12b (qryMap_Misc_STCGL_12a (Journal Map, grouped by accountno,
            ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
            ' ** Location_ID_max), grouped and summed, by accountno), without 'INCOME O/U', with
            ' ** total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), grouped and
            ' ** summed, by assetno, with cnt), with pcash_per_acct, by specified [actno].
19680       Set qdf = .QueryDefs("qryMap_Misc_STCGL_17c")
19690     Case "frmJournal_Columns"
            ' ** qryMap_Misc_STCGL_62c (Journal Map, linked to Account, qryMap_Misc_STCGL_62b
            ' ** (qryMap_Misc_STCGL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
            ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max)),
            ' ** grouped and summed, by accountno), just 'INCOME O/U', linked to qryMap_Misc_STCGL_67b
            ' ** (qryMap_Misc_STCGL_67a (qryMap_Misc_STCGL_62c (Journal Map, linked to Account,
            ' ** qryMap_Misc_STCGL_62b (qryMap_Misc_STCGL_62a (Journal Map, grouped by accountno,
            ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
            ' ** Location_ID_max)), grouped and summed, by accountno), without 'INCOME O/U', with
            ' ** total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), grouped and
            ' ** summed, by assetno, with cnt), with pcash_per_acct, by specified [actno].
19700       Set qdf = .QueryDefs("qryMap_Misc_STCGL_67c")
19710     End Select
19720     With qdf.Parameters
19730       ![actno] = strAccountNo_OU
19740     End With
19750     Set rst1 = qdf.OpenRecordset
19760     With rst1
19770       If .BOF = True And .EOF = True Then
              ' ** No INCOME O/U to distribute!
19780         blnNoDistribution = True
19790       Else
19800         .MoveFirst
19810         If ![total_ICash_ou] < 1@ Then
                ' ** Less than a buck, leave it alone.
19820           blnNoDistribution = True
19830         Else
19840           If ![pennies_per_acct] < 0.01 Then
                  ' ** Too small to divvy up!
19850             blnNoDistribution = True
19860           Else
19870             If ![ICash_per_acct] < 0.01 Then
                    ' ** Rounded value too small to divvy up!
19880               blnNoDistribution = True
19890             Else
                    ' ** Continue with the distribution.
19900             End If
19910           End If
19920         End If
19930       End If
19940       .Close
19950     End With  ' ** rst1.
19960     Set rst1 = Nothing
19970     Set qdf = Nothing
19980     DoEvents

          ' ** Step 18: Empty tblJournal_Map_Staging2.
19990     Me.Map_Status_lbl.Caption = "18 of 22: Empty temporary table 2 . . ."
20000     DoEvents

20010     Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty tblJournal_Map_Staging2.
20020       Set qdf = .QueryDefs("qryMap_Misc_STCGL_18")
20030     Case "frmJournal_Columns"
            ' ** Empty tblJournal_Map_Staging2.
20040       Set qdf = .QueryDefs("qryMap_Misc_STCGL_68")
20050     End Select
20060     qdf.Execute
20070     Set qdf = Nothing
20080     DoEvents

20090     Select Case blnNoDistribution
          Case True

            ' ** Step 19: Collect unchanged distributions.
20100       Me.Map_Status_lbl.Caption = "19 of 22: Collect unchanged distributions . . ."
20110       DoEvents

20120       Select Case strCallingForm
            Case "frmJournal"
20130         If dblPerShare > 0# Then
                ' ** Append qryMap_Misc_STCGL_12c (Journal Map, linked to Account, qryMap_Misc_STCGL_12b
                ' ** (qryMap_Misc_STCGL_12a (Journal Map, grouped by accountno, assetno, Location_ID),
                ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max), grouped
                ' ** and summed, by accountno) to tblJournal_Map_Staging2, with journaltype = 'Misc.', Gain.
20140           Set qdf = .QueryDefs("qryMap_Misc_STCGL_19a")
20150         Else
                ' ** Append qryMap_Misc_STCGL_12c (Journal Map, linked to Account, qryMap_Misc_STCGL_12b
                ' ** (qryMap_Misc_STCGL_12a (Journal Map, grouped by accountno, assetno, Location_ID),
                ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max), grouped
                ' ** and summed, by accountno) to tblJournal_Map_Staging2, with journaltype = 'Misc.', Loss.
20160           Set qdf = .QueryDefs("qryMap_Misc_STCGL_19b")
20170         End If
20180       Case "frmJournal_Columns"
20190         If dblPerShare > 0# Then
                ' ** Append qryMap_Misc_STCGL_62c (Journal Map, linked to Account, qryMap_Misc_STCGL_62b
                ' ** (qryMap_Misc_STCGL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
                ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max)), grouped
                ' ** and summed, by accountno) to tblJournal_Map_Staging2, with journaltype = 'Misc.', Gain.
20200           Set qdf = .QueryDefs("qryMap_Misc_STCGL_69a")
20210         Else
                ' ** Append qryMap_Misc_STCGL_62c (Journal Map, linked to Account, qryMap_Misc_STCGL_62b
                ' ** (qryMap_Misc_STCGL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
                ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max)), grouped
                ' ** and summed, by accountno) to tblJournal_Map_Staging2, with journaltype = 'Misc.', Loss.
20220           Set qdf = .QueryDefs("qryMap_Misc_STCGL_69b")
20230         End If
20240       End Select
20250       qdf.Execute
20260       Set qdf = Nothing

20270     Case False

20280       Set rst2 = .OpenRecordset("tblJournal_Map_Staging2", dbOpenDynaset, dbConsistent)

            ' ** Step 19: Create new distributions.
20290       Me.Map_Status_lbl.Caption = "19 of 22: Create new distributions . . ."
20300       DoEvents

20310       Select Case strCallingForm
            Case "frmJournal"
              ' ** qryMap_Misc_STCGL_17a (qryMap_Misc_STCGL_12c (Journal Map, linked to Account,
              ' ** qryMap_Misc_STCGL_12b (qryMap_Misc_STCGL_12a (Journal Map, grouped by accountno,
              ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
              ' ** Location_ID_max), grouped and summed, by accountno), without 'INCOME O/U', with
              ' ** total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), linked to
              ' ** qryMap_Misc_STCGL_20a (qryMap_Misc_STCGL_17c (qryMap_Misc_STCGL_12c (Journal Map,
              ' ** linked to Account, qryMap_Misc_STCGL_12b (qryMap_Misc_STCGL_12a (Journal Map,
              ' ** grouped by accountno, assetno, Location_ID), grouped by accountno, assetno, with
              ' ** cnt, Location_ID_min, Location_ID_max), grouped and summed, by accountno), just
              ' ** 'INCOME O/U', linked to qryMap_Misc_STCGL_17b (qryMap_Misc_STCGL_17a
              ' ** (qryMap_Misc_STCGL_12c (Journal Map, linked to Account, qryMap_Misc_STCGL_12b
              ' ** (qryMap_Misc_STCGL_12a (Journal Map, grouped by accountno, assetno, Location_ID),
              ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max),
              ' ** grouped and summed, by accountno), without 'INCOME O/U', with total_pcash_rnd,
              ' ** total_pcash_dbl, gain_loss, by specified [actno]), grouped and summed, by
              ' ** assetno, with cnt), with pcash_per_acct, by specified [actno]), with
              ' ** shareface_per_acct), with total_shareface_new, total_pcash_new.
20320         Set qdf = .QueryDefs("qryMap_Misc_STCGL_20b")
20330       Case "frmJournal_Columns"
              ' ** qryMap_Misc_STCGL_67a (qryMap_Misc_STCGL_62c (Journal Map, linked to Account,
              ' ** qryMap_Misc_STCGL_62b (qryMap_Misc_STCGL_62a (Journal Map, grouped by accountno,
              ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
              ' ** Location_ID_max)), grouped and summed, by accountno), without 'INCOME O/U', with
              ' ** total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), linked to
              ' ** qryMap_Misc_STCGL_70a (qryMap_Misc_STCGL_67c (qryMap_Misc_STCGL_62c (Journal Map,
              ' ** linked to Account, qryMap_Misc_STCGL_62b (qryMap_Misc_STCGL_62a (Journal Map,
              ' ** grouped by accountno, assetno, Location_ID), grouped by accountno, assetno, with
              ' ** cnt, Location_ID_min, Location_ID_max)), grouped and summed, by accountno), just
              ' ** 'INCOME O/U', linked to qryMap_Misc_STCGL_67b (qryMap_Misc_STCGL_67a
              ' ** (qryMap_Misc_STCGL_62c (Journal Map, linked to Account, qryMap_Misc_STCGL_62b
              ' ** (qryMap_Misc_STCGL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
              ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max)),
              ' ** grouped and summed, by accountno), without 'INCOME O/U', with total_pcash_rnd,
              ' ** total_pcash_dbl, gain_loss, by specified [actno]), grouped and summed, by
              ' ** assetno, with cnt), with pcash_per_acct, by specified [actno]), with
              ' ** shareface_per_acct), with total_shareface_new, total_pcash_new.
20340         Set qdf = .QueryDefs("qryMap_Misc_STCGL_70b")
20350       End Select
20360       With qdf.Parameters
20370         ![actno] = strAccountNo_OU
20380       End With
20390       Set rst1 = qdf.OpenRecordset
20400       With rst1
20410         .MoveLast
20420         lngRecs = .RecordCount
20430         .MoveFirst
20440         For lngX = 1& To lngRecs
20450           With rst2
20460             .AddNew
                  ' ** ![jmap_id] : AutoNumber.
20470             ![accountno] = rst1![accountno]
20480             ![assetno] = rst1![assetno]
20490             ![transdate] = rst1![transdate]
20500             ![shareface] = rst1![total_shareface_new]
20510             ![pershare] = rst1![pershare]
20520             ![journaltype] = "Misc."
20530             ![ICash] = 0@
20540             ![ICash] = rst1![total_ICash_new]
20550             ![Cost] = 0@
20560             If dblPerShare > 0# Then
20570               ![description] = "STCG"
20580             Else
20590               ![description] = "STCL"
20600             End If
20610             ![rate] = 0#
20620             ![due] = Null
20630             ![Location_ID] = rst1![Location_ID]
20640             If dblPerShare > 0# Then
20650               ![RecurringItem] = "Short Term Capital Gain"
20660             Else
20670               ![RecurringItem] = "Short Term Capital Loss"
20680             End If
20690             ![RecurringType] = "Misc"
20700             .Update
20710           End With  ' ** rst2.
20720           DoEvents
20730           If lngX < lngRecs Then .MoveNext
20740         Next  ' ** lngX.
20750         .Close
20760       End With  ' ** rst1.
20770       Set rst1 = Nothing
20780       Set qdf = Nothing
20790       DoEvents

            ' ** Step 20: Determine Over/Under location.
20800       Me.Map_Status_lbl.Caption = "20 of 22: Determine Over/Under location . . ."
20810       DoEvents

20820       Select Case strCallingForm
            Case "frmJournal"
20830         If dblPerShare > 0# Then
                ' ** Ledger, just 'INCOME O/U', 'Received', grouped by accountno, with
                ' ** Location_ID_min, Location_ID_max, cnt, by specified [actno], [astno].
20840           Set qdf = .QueryDefs("qryMap_Misc_STCGL_21a")
20850         Else
                ' ** Ledger, just 'INCOME O/U', 'Paid', grouped by accountno, with
                ' ** Location_ID_min, Location_ID_max, cnt, by specified [actno], [astno]
20860           Set qdf = .QueryDefs("qryMap_Misc_STCGL_21b")
20870         End If
20880       Case "frmJournal_Columns"
20890         If dblPerShare > 0# Then
                ' ** Ledger, just 'INCOME O/U', 'Received', grouped by accountno, with
                ' ** Location_ID_min, Location_ID_max, cnt, by specified [actno], [astno].
20900           Set qdf = .QueryDefs("qryMap_Misc_STCGL_71a")
20910         Else
                ' ** Ledger, just 'INCOME O/U', 'Paid', grouped by accountno, with
                ' ** Location_ID_min, Location_ID_max, cnt, by specified [actno], [astno]
20920           Set qdf = .QueryDefs("qryMap_Misc_STCGL_71b")
20930         End If
20940       End Select
20950       With qdf.Parameters
20960         ![actno] = strAccountNo_OU
20970         ![astno] = lngAssetNo
20980       End With
20990       Set rst1 = qdf.OpenRecordset
21000       With rst1
21010         If .BOF = True And .EOF = True Then
21020           lngLocationID = 1&  ' ** Default.
21030         Else
21040           .MoveFirst
                ' ** CLng(IIf([cnt]>2,1, IIf([cnt]=1,[Location_ID_max],IIf([cnt]=2 And [Location_ID_min]=1,[Location_ID_max],1))))
21050           If ![cnt] > 2 Then
21060             lngLocationID = 1&
21070           Else
21080             If ![cnt] = 1 Then
21090               lngLocationID = ![Location_ID_max]
21100             Else
21110               If ![cnt] = 2 And [Location_ID_min] = 1 Then
21120                 lngLocationID = ![Location_ID_max]
21130               Else
21140                 lngLocationID = 1&
21150               End If
21160             End If
21170           End If
21180         End If
21190         .Close
21200       End With  ' ** rst1.
21210       Set rst1 = Nothing
21220       Set qdf = Nothing
21230       DoEvents

            ' ** Step 21: Create new Over/Under entry.
21240       Me.Map_Status_lbl.Caption = "22 of 22: Create new Over/Under entry . . ."
21250       DoEvents

21260       Select Case strCallingForm
            Case "frmJournal"
              ' ** qryMap_Misc_STCGL_22a (qryMap_Misc_STCGL_20b (qryMap_Misc_STCGL_17a (qryMap_Misc_STCGL_12c
              ' ** (Journal Map, linked to Account, qryMap_Misc_STCGL_12b (qryMap_Misc_STCGL_12a (Journal Map,
              ' ** grouped by accountno, assetno, Location_ID), grouped by accountno, assetno, with cnt,
              ' ** Location_ID_min, Location_ID_max), grouped and summed, by accountno), without 'INCOME O/U',
              ' ** with total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), linked to
              ' ** qryMap_Misc_STCGL_20a (qryMap_Misc_STCGL_17c (qryMap_Misc_STCGL_12c (Journal Map, linked to
              ' ** Account, qryMap_Misc_STCGL_12b (qryMap_Misc_STCGL_12a (Journal Map, grouped by accountno,
              ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
              ' ** Location_ID_max), grouped and summed, by accountno), just 'INCOME O/U', linked to
              ' ** qryMap_Misc_STCGL_17b (qryMap_Misc_STCGL_17a (qryMap_Misc_STCGL_12c (Journal Map, linked to
              ' ** Account, qryMap_Misc_STCGL_12b (qryMap_Misc_STCGL_12a (Journal Map, grouped by accountno,
              ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
              ' ** Location_ID_max), grouped and summed, by accountno), without 'INCOME O/U', with
              ' ** total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), grouped and summed, by
              ' ** assetno, with cnt), with pcash_per_acct, by specified [actno]), with shareface_per_acct),
              ' ** with total_shareface_new, total_pcash_new), grouped and summed, by assetno), linked to
              ' ** qryMap_Misc_STCGL_12d (qryMap_Misc_STCGL_12c (Journal Map, linked to Account, qryMap_Misc_STCGL_12b
              ' ** (qryMap_Misc_STCGL_12a (Journal Map, grouped by accountno, assetno, Location_ID), grouped by
              ' ** accountno, assetno, with cnt, Location_ID_min, Location_ID_max), grouped and summed, by
              ' ** accountno), grouped and summed, by assetno, by specified [astno]), with shareface_ou_new,
              ' ** pcash_ou_new, by specified [actno], [locid].
21270         Set qdf = .QueryDefs("qryMap_Misc_STCGL_22b")
21280       Case "frmJournal_Columns"
              ' ** qryMap_Misc_STCGL_72a (qryMap_Misc_STCGL_70b (qryMap_Misc_STCGL_67a (qryMap_Misc_STCGL_62c
              ' ** (Journal Map, linked to Account, qryMap_Misc_STCGL_62b (qryMap_Misc_STCGL_62a (Journal Map,
              ' ** grouped by accountno, assetno, Location_ID), grouped by accountno, assetno, with cnt,
              ' ** Location_ID_min, Location_ID_max)), grouped and summed, by accountno), without 'INCOME O/U',
              ' ** with total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), linked to
              ' ** qryMap_Misc_STCGL_70a (qryMap_Misc_STCGL_67c (qryMap_Misc_STCGL_62c (Journal Map, linked to
              ' ** Account, qryMap_Misc_STCGL_62b (qryMap_Misc_STCGL_62a (Journal Map, grouped by accountno,
              ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
              ' ** Location_ID_max)), grouped and summed, by accountno), just 'INCOME O/U', linked to
              ' ** qryMap_Misc_STCGL_67b (qryMap_Misc_STCGL_67a (qryMap_Misc_STCGL_62c (Journal Map, linked to
              ' ** Account, qryMap_Misc_STCGL_62b (qryMap_Misc_STCGL_62a (Journal Map, grouped by accountno,
              ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
              ' ** Location_ID_max)), grouped and summed, by accountno), without 'INCOME O/U', with
              ' ** total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), grouped and summed, by
              ' ** assetno, with cnt), with pcash_per_acct, by specified [actno]), with shareface_per_acct),
              ' ** with total_shareface_new, total_pcash_new), grouped and summed, by assetno), linked to
              ' ** qryMap_Misc_STCGL_62d (qryMap_Misc_STCGL_62c (Journal Map, linked to Account, qryMap_Misc_STCGL_62b
              ' ** (qryMap_Misc_STCGL_62a (Journal Map, grouped by accountno, assetno, Location_ID), grouped by
              ' ** accountno, assetno, with cnt, Location_ID_min, Location_ID_max)), grouped and summed, by
              ' ** accountno), grouped and summed, by assetno, by specified [astno]), with shareface_ou_new,
              ' ** pcash_ou_new, by specified [actno], [locid].
21290         Set qdf = .QueryDefs("qryMap_Misc_STCGL_72b")
21300       End Select
21310       With qdf.Parameters
21320         ![astno] = lngAssetNo
21330         ![actno] = strAccountNo_OU
21340         ![locid] = lngLocationID
21350       End With
21360       Set rst1 = qdf.OpenRecordset
21370       With rst1
21380         .MoveFirst
21390         With rst2
21400           .AddNew
                ' ** ![jmap_id] : AutoNumber.
21410           ![accountno] = rst1![accountno]
21420           ![assetno] = rst1![assetno]
21430           ![transdate] = rst1![transdate]
21440           ![shareface] = rst1![shareface_ou_new]
21450           ![pershare] = rst1![pershare]
21460           ![journaltype] = "Misc."
21470           ![ICash] = 0@
21480           dblRetVal = rst1![ICash_ou_new]
21490           ![ICash] = rst1![ICash_ou_new]
21500           ![Cost] = 0@
21510           If dblPerShare > 0# Then
21520             ![description] = "STCG"
21530           Else
21540             ![description] = "STCL"
21550           End If
21560           ![rate] = 0#
21570           ![due] = Null
21580           ![Location_ID] = rst1![Location_ID]
21590           If dblPerShare > 0# Then
21600             ![RecurringItem] = "Short Term Capital Gain"
21610           Else
21620             ![RecurringItem] = "Short Term Capital Loss"
21630           End If
21640           ![RecurringType] = "Misc"
21650           .Update
21660           .Close
21670         End With  ' ** rst2.
21680         .Close
21690       End With  ' ** rst1.
21700       Set rst1 = Nothing
21710       Set rst2 = Nothing
21720       Set qdf = Nothing

21730     End Select  ' ** blnNoDistribution.

21740     .Close
21750   End With  ' ** dbs.
21760   DoEvents

EXITP:
21770   Set rst1 = Nothing
21780   Set rst1 = Nothing
21790   Set qdf = Nothing
21800   Set dbs = Nothing
21810   AdvancedRounding = dblRetVal
21820   Exit Function

ERRH:
21830   DoCmd.Hourglass False
21840   dblRetVal = 0#
21850   Select Case ERR.Number
        Case Else
21860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21870   End Select
21880   Resume EXITP

End Function

Private Sub IncludeCurrency()

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngTmp01 As Long

21910   With Me
21920     .curr_id.Enabled = False
21930     .curr_id.Locked = True
21940     .curr_id.ForeColor = CLR_DISABLED_FG
21950     .curr_id.BackColor = CLR_VLTGRN
21960     Select Case blnCurrID
          Case True
21970       .curr_id.Visible = True
21980       .curr_id_cmd.Visible = True
21990       .curr_id_cmd.Enabled = True
22000       .cmbAssets_vline01.Visible = True
22010       .cmbAssets_vline02.Visible = True
22020     Case False
22030       .curr_id.Visible = False
22040       .curr_id_cmd.Visible = False
22050       .curr_id_cmd.Enabled = False
22060       .cmbAssets_vline01.Visible = False
22070       .cmbAssets_vline02.Visible = False
22080       lngTmp01 = ((.curr_id.Left + .curr_id.Width) - (.cmbAssets.Left + .cmbAssets.Width))
22090       .cmbAssets.Width = (.cmbAssets.Width + lngTmp01)
22100     End Select
22110     .Amount_usd.Visible = False
22120     .pershare_usd.Visible = False
22130   End With

        ' ** cmbAssets RowSource's:
        ' **   qryMap_Rec_05a
        ' **   qryMap_Rec_05c
        ' **   qryMap_Rec_05d
        ' **   qryMap_Rec_05f
        ' **   qryMap_Rec_05g
        ' **   qryMap_Rec_05i

EXITP:
22140   Exit Sub

ERRH:
22150   DoCmd.Hourglass False
22160   Select Case ERR.Number
        Case Else
22170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22180   End Select
22190   Resume EXITP

End Sub

Private Sub NoChar_Load()

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

22210   Set dbs = CurrentDb
22220   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
22230     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
22240     Set rst = qdf.OpenRecordset
22250     With rst
22260       .MoveLast
22270       lngNoChars = .RecordCount
22280       .MoveFirst
22290       arr_varNoChar = .GetRows(lngNoChars)
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
22300       .Close
22310     End With
22320     .Close
22330   End With

EXITP:
22340   Set rst = Nothing
22350   Set qdf = Nothing
22360   Set dbs = Nothing
22370   Exit Sub

ERRH:
22380   DoCmd.Hourglass False
22390   Select Case ERR.Number
        Case Else
22400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22410   End Select
22420   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

22510   With Me
22520     blnGTR_Emblem = False
22530     For lngX = 1& To 24&
22540       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
22550     Next
22560   End With

EXITP:
22570   Exit Sub

ERRH:
22580   Select Case ERR.Number
        Case Else
22590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22600   End Select
22610   Resume EXITP

End Sub
