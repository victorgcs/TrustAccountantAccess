VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Misc_LTCL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Misc_LTCL"

'VGC 09/09/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Date:             D {transdate}
' **   Principal Cash:   H {Amount}
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

' ** Trying to force case on icash/ICash and pcash/PCash by use of a constant doesn't work; neither Private nor Public!

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean

Private strFrmt1 As String, strFrmt2 As String, intDec1 As Integer, intDec2 As Integer, lngBkClr1 As Long, lngBkClr2 As Long
Private strAccountNo_OU As String
Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private dblPerShare_Calc As Double, lngTpp As Long, blnCurrID As Boolean, lngCurrID As Long
Private datPostingDate As Date, intInputType As Integer
Private blnAssetUpdated As Boolean, blnAssetChecked As Boolean, blnDrivingMeBackwards As Boolean
Private blnExitingAssetNo As Boolean, blnCancel As Boolean
Private THAT_PROC As String, That_Erl As Long
Private strCallingForm As String
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
320           CLR_DISABLED_BG_R = CLR_LTRED     '14013951
330           CLR_DISABLED_BG_VLR = CLR_VLTRED  '15921919

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
890         Set grstPostingDate = Nothing
            ' ** Append qryMap_Misc_LTCL_Detail_03 (Query Source; RecurringType, just 'Misc',
            ' ** as new RecurringItems entry, for 'Long Term Capital Loss') to RecurringItems.
900         Set qdf = .QueryDefs("qryMap_Misc_LTCL_Detail_05")
910         qdf.Execute
920         Set qdf = Nothing
930         DoEvents
            ' ** Empty tmpAccount.
940         Set qdf = .QueryDefs("qryMap_Misc_LTCL_08_01")
950         qdf.Execute
960         Set qdf = Nothing
970         DoEvents
            ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
            ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
            ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
            ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
            ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
            ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
            ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
            ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
980         Set qdf = .QueryDefs("qryMap_Misc_LTCL_08_02")
990         qdf.Execute
1000        Set qdf = Nothing
1010        DoEvents
1020        .Close
1030      End With
1040      Set dbs = Nothing
1050      DoEvents

1060      .transdate.DefaultValue = "#" & Format(datPostingDate, "mm/dd/yyyy") & "#"
1070      .shareface.ForeColor = CLR_DISABLED_FG
1080      .shareface.BackColor = CLR_DISABLED_BG

1090      IncludeCurrency  ' ** Procedure: Below.

1100    End With

1110    DoCmd.Hourglass False

EXITP:
1120    Set qdf = Nothing
1130    Set dbs = Nothing
1140    Exit Sub

ERRH:
1150    DoCmd.Hourglass False
1160    Select Case ERR.Number
        Case Else
1170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1180    End Select
1190    Resume EXITP

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
1820            Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_07")
1830          Case "frmJournal_Columns"
                ' ** ActiveAssets, grouped and summed, by assetno, by specified [astno].
1840            Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_57")
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
2380                .amount = 100
2390                Cancel = 0
2400                amount_Exit Cancel  ' ** Procedure: Below.
2410                DoEvents
2420                .GoToReport_arw_ok_img.Visible = True
2430                .GoToReport_arw_map_amt_img.Visible = False
2440                .GoToReport_arw_map_per_img.Visible = False
2450                .GoToReport_lbl_amtper.Visible = False
2460                .GoToReport_lin_hline01.Visible = False
2470                .GoToReport_lin_hline02.Visible = False
2480                .GoToReport_lbl_arrow01.Visible = False
2490                .GoToReport_lbl_arrow02.Visible = False
2500                If strCallingForm = "frmJournal_Columns" Then
                      ' ** Don't know why this continues to light up.
2510                  Forms(strCallingForm).cmdSpecPurp_Purch_MapSplit_raised_img.Visible = True
2520                  Forms(strCallingForm).cmdSpecPurp_Purch_MapSplit_raised_semifocus_dots_img.Visible = False
2530                End If
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
2670                .GoToReport_lin_hline01.Visible = True
2680                .GoToReport_lin_hline02.Visible = True
2690                .GoToReport_lbl_arrow01.Visible = True
2700                .GoToReport_lbl_arrow02.Visible = True
2710                .GoToReport_arw_map_asset_img.Visible = False
2720                .opgAssetSource.Visible = True
2730                .opgAssetSource_box.Visible = True
2740                .TotalShares = 4000
2750                .shareface = 4000
2760                .FocusHolder.SetFocus  ' ** So they don't see a flashing cursor.
2770                DoEvents
2780                .TimerInterval = 50&
2790              Else
2800                .TimerInterval = 0&
2810                Select Case garr_varGoToReport(GTR_FRM2)
                    Case "frmJournal"
2820                  Forms(garr_varGoToReport(GTR_FRM2)).GTREmblem_Set False  ' ** Form Procedure: frmJournal.
2830                Case "frmJournal_Columns"
2840                  Forms(garr_varGoToReport(GTR_FRM2)).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
2850                End Select
2860                gblnGoToReport = False
2870                blnGoingToReport = False
2880                blnGoingToReport2 = False
2890                GTREmblem_Off  ' ** Procedure: Below.
2900                .GoToReport_arw_map_asset_img.Visible = False
2910                .opgAssetSource.Visible = True
2920                .opgAssetSource_box.Visible = True
2930                Beep
2940                DoCmd.Hourglass False
2950                MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                      "There are insufficient asset holdings to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
2960              End If
2970            End Select
2980          Case False
2990            blnGoingToReport = True
3000            .GoToReport_arw_map_asset_img.Visible = True
3010            If strCallingForm = "frmJournal_Columns" Then
                  ' ** Don't know why this continues to light up.
3020              Forms(strCallingForm).cmdSpecPurp_Purch_MapSplit_raised_img.Visible = True
3030              Forms(strCallingForm).cmdSpecPurp_Purch_MapSplit_raised_semifocus_dots_img.Visible = False
3040            End If
3050            .opgAssetSource.Visible = False
3060            .opgAssetSource_box.Visible = False
3070            .FocusHolder.SetFocus  ' ** So they don't see a flashing cursor.
3080            DoEvents
3090            Select Case garr_varGoToReport(GTR_FRM2)
                Case "frmJournal"
3100              Forms(garr_varGoToReport(GTR_FRM2)).GTREmblem_Set False  ' ** Form Procedure: frmJournal.
3110            Case "frmJournal_Columns"
3120              Forms(garr_varGoToReport(GTR_FRM2)).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
3130            End Select
3140            .TimerInterval = 50&
3150          End Select
3160        End If
3170      End If

3180    End With  ' ** Me.

EXITP:
3190    Set rst = Nothing
3200    Set qdf = Nothing
3210    Set dbs = Nothing
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim strControlName As String

3310    strControlName = vbNullString

3320    If DataErr <> 0 Then
3330      Select Case DataErr
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
3340      Case 2046  ' ** The command or action isn't available now (first or last record).
3350        DoCmd.CancelEvent
3360      Case 2113  ' ** The value you entered isn't valid for this field.
3370        Select Case Me.ActiveControl.Name
            Case "shareface"
3380          strControlName = "Shares"
3390        Case "pershare"
3400          strControlName = "Rate"
3410        Case "amount"
3420          strControlName = "Amount"
3430        Case "transdate"
3440          strControlName = vbNullString
3450          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3460          Response = acDataErrContinue
3470          DoCmd.CancelEvent
3480          Exit Sub
3490        Case Else
3500          strControlName = vbNullString
3510        End Select
3520        If strControlName <> vbNullString Then
3530          MsgBox "You must enter a numeric value in the " & strControlName & " field.", vbInformation + vbOKOnly, "Invalid Entry"
3540          Response = acDataErrContinue
3550          DoCmd.CancelEvent
3560        End If
3570      Case 2237  ' ** The text you entered isn't an item in the list.
3580        DoCmd.CancelEvent
3590      Case 2279  ' ** The value you entered isn't appropriate for the input mask '|' specified for this field.
3600        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3610        Response = acDataErrContinue
3620        DoCmd.CancelEvent
3630      Case 2580  ' ** The record source '|' specified on this form or report does not exist.
            ' ** Ignore! This form has no RecordSource.
3640        Response = acDataErrContinue
3650      Case 3162  ' ** You tried to assign the Null value to a variable that is not a Variant data type.
3660        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
            ' ** Do nothing.
3670      Case Else
3680        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3690        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3700      End Select
3710    End If

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3810    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3820    If Not clsMonthClass Is Nothing Then
3830      If clsMonthClass.IsCalendar Then
3840        Cancel = -1
3850      Else
3860        Set clsMonthClass = Nothing
3870      End If
3880    End If

3890    DoCmd.Hourglass False

EXITP:
3900    Exit Sub

ERRH:
3910    DoCmd.Hourglass False
3920    Select Case ERR.Number
        Case Else
3930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3940    End Select
3950    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4010    With Me
4020      If .cmdCalendar_raised_focus_dots_img.Visible = True Or .cmdCalendar_raised_focus_img.Visible = True Then
4030        Select Case blnCalendar1_Focus
            Case True
4040          .cmdCalendar_raised_semifocus_dots_img.Visible = True
4050          .cmdCalendar_raised_img.Visible = False
4060        Case False
4070          .cmdCalendar_raised_img.Visible = True
4080          .cmdCalendar_raised_semifocus_dots_img.Visible = False
4090        End Select
4100        .cmdCalendar_raised_focus_dots_img.Visible = False
4110        .cmdCalendar_raised_focus_img.Visible = False
4120        .cmdCalendar_sunken_focus_dots_img.Visible = False
4130        .cmdCalendar_raised_img_dis.Visible = False
4140      End If
4150    End With

EXITP:
4160    Exit Sub

ERRH:
4170    Select Case ERR.Number
        Case Else
4180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4190    End Select
4200    Resume EXITP

End Sub

Private Sub cmdOK_Click()

4300  On Error GoTo ERRH

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
        Dim intPos01 As Integer
        Dim strTmp01 As String, dblTmp02 As Double, dblTmp03 As Double, lngTmp04 As Long, lngTmp05 As Long, strTmp06 As String
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

4310    blnContinue = True

4320    DoCmd.Hourglass True
4330    DoEvents

4340    dblShareface = 0#
4350    dblAmount = 0#
4360    dblSumPCash = 0#
4370    dblSumShareface = 0#
4380    dblPerShare = 0#
4390    dblOverUnder = 0#
4400    lngAssetNo = 0&
4410    datTransDate = 0
4420    datAssetDate = 0
4430    lngTaxcode = 0&

4440    With Me

          ' ** Step 1: Validation.
4450      .Map_Status_lbl.Caption = " 1 of 22: Validation . . ."
4460      DoEvents

4470      If IsNull(.cmbAssets) = True Then
4480        blnContinue = False
4490        DoCmd.Hourglass False
4500        MsgBox "Please select a valid asset.", vbInformation + vbOKOnly, "Entry Required"
4510        .cmbAssets.SetFocus
4520      Else
4530        If Trim(.cmbAssets) = vbNullString Then
4540          blnContinue = False
4550          DoCmd.Hourglass False
4560          MsgBox "Please select a valid asset.", vbInformation + vbOKOnly, "Entry Required"
4570          .cmbAssets.SetFocus
4580        Else
4590          lngAssetNo = .cmbAssets
4600          If Nz(.shareface, 0) = 0 Then
4610            blnContinue = False
4620            DoCmd.Hourglass False
4630            MsgBox "Share/Face must be a number greater than zero.", vbInformation + vbOKOnly, "Entry Required"
4640            .cmbAssets.SetFocus
4650          Else
4660            If IsNull(.amount) = True Then
4670              blnContinue = False
4680              DoCmd.Hourglass False
4690              MsgBox "Amount must be a number greater than zero.", vbInformation + vbOKOnly, "Entry Required"
4700              .amount.SetFocus
4710            Else
4720              If IsNull(.pershare) = True Then
4730                blnContinue = False
4740                DoCmd.Hourglass False
4750                MsgBox "Per Share must be greater than zero to continue.", vbInformation + vbOKOnly, "Entry Required"
4760                .pershare.SetFocus
4770              Else
4780                If IsDate(.transdate) = False Then
4790                  blnContinue = False
4800                  DoCmd.Hourglass False
4810                  MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Date"
4820                  .transdate.SetFocus
4830                Else
                      ' ** Allow posting up to 1 month into the future.
4840                  If .transdate > DateAdd("m", 1, Date) Then
4850                    blnContinue = False
4860                    DoCmd.Hourglass False
4870                    MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
4880                  Else
4890                    If .amount = 0 And .pershare = 0 Then
4900                      blnContinue = False
4910                      DoCmd.Hourglass False
4920                      MsgBox "Either Amount or Per Share must be greater than zero to continue.", vbInformation + vbOKOnly, "Entry Required"
4930                      If .amount.Enabled = False Or .pershare.Enabled = False Then
4940                        .amount = Null
4950                        .amount.Enabled = True
4960                        .amount.Locked = False
4970                        .amount.ForeColor = CLR_BLK
4980                        .pershare = Null
4990                        .pershare.Enabled = True
5000                        .pershare.Locked = False
5010                        .pershare.ForeColor = CLR_BLK
5020                        If lngCurrID = 150& Then
5030                          .amount.BackColor = CLR_WHT
5040                          .pershare.BackColor = CLR_WHT
5050                        Else
5060                          .amount.BackColor = lngBkClr1
5070                          .pershare.BackColor = lngBkClr1
5080                        End If
5090                      End If
5100                      .amount.SetFocus
5110                    End If
5120                  End If
5130                End If
5140              End If
5150            End If
5160          End If
5170        End If
5180      End If
5190      DoEvents

5200      If blnContinue = True Then

5210        DoCmd.Hourglass True
5220        DoEvents

5230        Set dbs = CurrentDb

            ' ** Step 2: Empty Journal Map.
5240        .Map_Status_lbl.Caption = " 2 of 22: Empty temporary table 1 . . ."
5250        DoEvents

5260        Select Case strCallingForm
            Case "frmJournal"
              ' ** Empty Journal Map.
5270          Set qdf = dbs.QueryDefs("qryMap_01")
5280        Case "frmJournal_Columns"
              ' ** Empty Journal Map.
5290          Set qdf = dbs.QueryDefs("qryJournal_Columns_35_04")
5300        End Select
5310        qdf.Execute
5320        DoEvents

            ' ** Step 3: Set Tax Code.
5330  On Error Resume Next
5340        .Map_Status_lbl.Caption = " 3 of 22: Set Tax Code . . ."
            ' ** 2467  The expression you entered refers to an object that is closed or doesn't exist.
5350  On Error GoTo ERRH
5360        DoEvents

5370        Select Case strCallingForm
            Case "frmJournal"
              ' ** MasterAsset, linked to AssetType, by specified [astno].
5380          Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_02")
5390        Case "frmJournal_Columns"
              ' ** MasterAsset, linked to AssetType, by specified [astno].
5400          Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_52")
5410        End Select
5420        With qdf.Parameters
5430          ![astno] = lngAssetNo
5440        End With
5450        Set rst = qdf.OpenRecordset
5460        If rst.BOF = True And rst.EOF = True Then
              ' ** Shouldn't happen.
5470        Else
5480          rst.MoveFirst
5490          lngTaxcode = rst![taxcode]
5500        End If
5510        .taxcode = lngTaxcode
5520        DoEvents

            ' ** Step 4: Calculate PerShare/Amount.
5530  On Error Resume Next
5540        .Map_Status_lbl.Caption = " 4 of 22: Calculate PerShare/Amount . . ."
5550  On Error GoTo ERRH
5560        DoEvents

5570        dblShareface = .TotalShares
5580        Select Case intInputType
            Case 1  ' ** User entered an Amount.
5590          dblAmount = .amount
5600          dblPerShare = (dblAmount / dblShareface)
5610        Case 2  ' ** User entered a PerShare.
5620          dblPerShare = .pershare
5630          dblAmount = (dblShareface * dblPerShare)
5640        Case Else  ' ** Shouldn't be an else!
5650          dblAmount = .amount
5660          dblPerShare = (dblAmount / dblShareface)
5670        End Select
5680        datTransDate = .transdate
            ' TM
5690        strMapType = .maptype  ' ** JournalType.
5700        DoEvents

            ' ** Step 5: Check for closed accounts.
5710  On Error Resume Next
5720        .Map_Status_lbl.Caption = " 5 of 22: Check for closed accounts . . ."
5730  On Error GoTo ERRH
5740        DoEvents

            ' ** Check for closed accounts holding assets.
            ' ** qryMap_Misc_LTCL_04 (qryMap_Misc_LTCL_03 (ActiveAssets, linked to Account,
            ' ** just closed accounts, by specified [astno]), grouped by accountno,
            ' ** with cnt_astno), grouped and summed, with cnt_actno, cnt_astno.
5750        Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_05")
5760        With qdf.Parameters
5770          ![astno] = Me.cmbAssets
5780        End With
5790        Set rst = qdf.OpenRecordset
5800        If rst.BOF = True And rst.EOF = True Then
              ' ** No closed accounts holding assets.
5810          msgResponse = vbYes
5820          rst.Close
5830        Else
5840          rst.MoveFirst
5850          If IsNull(rst![cnt_actno]) = True And IsNull(rst![cnt_astno]) = True Then
5860            msgResponse = vbYes
5870            rst.Close
5880          Else
5890            If rst![cnt_actno] = 0 And IsNull(rst![cnt_astno]) = True Then
5900              msgResponse = vbYes
5910              rst.Close
5920            ElseIf rst![cnt_actno] = 0 And rst![cnt_astno] = 0 Then
5930              msgResponse = vbYes
5940              rst.Close
5950            Else
5960              lngTmp04 = rst![cnt_actno]
5970              lngTmp05 = rst![cnt_astno]
5980              rst.Close
                  ' ** qryMap_Misc_LTCL_03 (ActiveAssets, linked to Account, just closed accounts,
                  ' ** by specified [astno]), grouped by accountno, with cnt_astno.
5990              Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_04")
6000              With qdf.Parameters
6010                ![astno] = lngAssetNo
6020              End With
6030              Set rst = qdf.OpenRecordset
6040              rst.MoveFirst
6050              If lngTmp04 = 1& Then
6060                strTmp06 = "There is 1 closed account currently holding this asset:" & vbCrLf & _
                      "    " & rst![accountno] & "  " & rst![shortname] & vbCrLf
6070              Else
6080                strTmp06 = "There are " & CStr(lngTmp04) & " closed accounts currently holding this asset:" & vbCrLf
6090                For lngX = 1& To lngTmp04
6100                  strTmp06 = strTmp06 & "    " & rst![accountno] & "  " & rst![shortname] & vbCrLf
6110                  If lngX < lngTmp04 Then rst.MoveNext
6120                Next
6130              End If
6140              strTmp06 = strTmp06 & "If you choose to continue with this mapping," & vbCrLf & _
                    "the results may be inaccurate." & vbCrLf & vbCrLf & "Do you wish to continue?"
6150              DoCmd.Hourglass False
6160              Beep
6170              msgResponse = MsgBox(strTmp06, vbQuestion + vbYesNo + vbDefaultButton2, "Closed Accounts Holding Assets")
6180            End If
6190          End If
6200        End If
6210        DoEvents

6220        If msgResponse = vbYes Then

6230          DoCmd.Hourglass True
6240          DoEvents

              ' ** Step 6: Collect Active Assets.
6250  On Error Resume Next
6260          .Map_Status_lbl.Caption = " 6 of 22: Collect Active Assets . . ."
6270  On Error GoTo ERRH
6280          DoEvents

6290          Select Case strCallingForm
              Case "frmJournal"
                ' ** Append ActiveAssets to Journal Map, by specified [astno], [astdat], [jrntyp], [persh], [transdat].  #curr_id
6300            Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_06")
6310          Case "frmJournal_Columns"
                ' ** Append ActiveAssets to Journal Map, by specified [astno], [astdat], [jrntyp], [persh], [transdat].  #curr_id
6320            Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_56")
6330          End Select
6340          With qdf.Parameters
6350            ![astno] = lngAssetNo
6360            ![astdat] = datAssetDate
6370            ![jrntyp] = strMapType
6380            ![persh] = dblPerShare  ' ** Used for both pcash and pershare.
6390            ![transdat] = datTransDate
6400          End With
6410          qdf.Execute
6420          dbs.Close
6430          Set qdf = Nothing
6440          Set dbs = Nothing
6450          DoEvents

6460          blnSkip = True
              ' ################################################################################
6470          If blnSkip = False Then
                'Set rsx1 = New ADODB.Recordset             ' ** Early binding.
6480            Set rsx1 = CreateObject("ADODB.Recordset")  ' ** Late binding.
6490            Set cnxn = CurrentProject.Connection
6500            rsx1.Open "journal map", cnxn, adOpenKeyset, adLockOptimistic, adCmdTableDirect
6510            With rsx1
6520              .MoveLast
6530              lngRecs = .RecordCount
6540              .MoveFirst
6550              For lngX = 1& To lngRecs
6560                dblTmp02 = Nz(.Fields("pcash"), 0)
6570                If ((dblTmp02 > 0.009) Or (dblTmp02 < -0.009)) Then
6580                  dblTmp02 = dbl_Round(dblTmp02, 3)  ' ** Module Function: modUtilities.
6590                Else
6600                  dblTmp02 = 0#
6610                End If
6620                If dblTmp02 <> 0# Then
6630                  dblTmp03 = Nz(.Fields("shareface"), 0)
6640                  If ((dblTmp03 > 0.00009) Or (dblTmp03 < -0.00009)) Then
6650                    dblTmp03 = dbl_Round(dblTmp03, 5)  ' ** Module Function: modUtilities.
6660                  Else
6670                    dblTmp03 = 0#
6680                    dblTmp02 = 0#  ' ** Also zero-out the pcash.
6690                  End If
6700                End If
6710                .Fields("pcash") = dblTmp02
6720                .Fields("shareface") = dblTmp03
6730                .Update
6740                dblSumPCash = dblSumPCash + dblTmp02
6750                dblSumShareface = dblSumShareface + dblTmp03
6760                dblTmp02 = 0#: dblTmp03 = 0#
6770                If lngX < lngRecs Then .MoveNext
6780              Next
6790            End With
6800          End If  ' ** blnSkip.
              ' ################################################################################

6810          DoEvents
6820          Set dbs = CurrentDb

              ' ** Step 7: Sum holdings.
6830  On Error Resume Next
6840          .Map_Status_lbl.Caption = " 7 of 22: Sum holdings . . ."
6850  On Error GoTo ERRH
6860          DoEvents

6870          Select Case intInputType
              Case 1  ' ** User entered an Amount.
                ' ** qryMap_Misc_LTCL_09a (qryMap_Misc_LTCL_08a (Journal Map, with qryMap_Misc_LTCL_07
                ' ** (ActiveAssets, grouped and summed, by assetno, by specified [astno]), with calc'd fields,
                ' ** by specified [astno], [amt], [prec]), grouped and summed by account), with .._rnd rounded again.
6880            Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_10a")
6890            With qdf.Parameters
6900              ![astno] = lngAssetNo
6910              ![amt] = dblAmount
6920              ![prec] = gintShareFaceDecimals
6930            End With
6940          Case 2  ' ** User entered a PerShare.
                ' ** qryMap_Misc_LTCL_09b (qryMap_Misc_LTCL_08b (Journal Map, with qryMap_Misc_LTCL_07
                ' ** (ActiveAssets, grouped and summed, by assetno, by specified [astno]), with calc'd fields,
                ' ** by specified [astno], [persh], [prec]), grouped and summed by account), with .._rnd rounded again.
6950            Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_10b")
6960            With qdf.Parameters
6970              ![astno] = lngAssetNo
6980              ![persh] = dblPerShare
6990              ![prec] = gintShareFaceDecimals
7000            End With
7010          End Select
7020          Set rst = qdf.OpenRecordset
7030          With rst
7040            If .BOF = True And .EOF = True Then
                  ' ** Something's seriously wrong!
7050              blnContinue = False
7060            Else
7070              .MoveLast
7080              lngMaps = .RecordCount
7090              .MoveFirst
7100              arr_varMap = .GetRows(lngMaps)
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
7110            End If
7120            .Close
7130          End With
7140          Set rst = Nothing
7150          Set qdf = Nothing
7160          DoEvents

              ' ** Step 8: Rounding Stage 1.
7170  On Error Resume Next
7180          .Map_Status_lbl.Caption = " 8 of 22: Rounding Stage 1 . . ."
7190  On Error GoTo ERRH
7200          DoEvents

7210          Set rst = dbs.OpenRecordset("journal map", dbOpenDynaset, dbConsistent)
7220          With rst
7230            If .BOF = True And .EOF = True Then
                  ' ** Might be coming from GoToReport.
7240              blnContinue = False
7250            Else
7260              .MoveLast
7270              lngRecs = .RecordCount
7280              .MoveFirst
7290              For lngX = 1& To lngRecs
7300                If InStr(![accountno], strAccountNo_OU) = 0 Then  ' ** This is excluded from the arr_varMap() query.
7310                  blnFound = False
7320                  For lngY = 0& To (lngMaps - 1&)
7330                    If arr_varMap(M_ACTNO, lngY) = ![accountno] And arr_varMap(M_ASTNO, lngY) = ![assetno] Then
7340                      blnFound = True
7350                      .Edit
7360                      dblTmp02 = Round(![shareface], gintShareFaceDecimals)
7370                      ![shareface] = dblTmp02
7380                      ![ICash] = CCur(Round((dblTmp02 * arr_varMap(M_PER, lngY)), 2))
7390                      ![pershare] = arr_varMap(M_PER, lngY)
7400                      .Update
7410                      Exit For
7420                    End If
7430                  Next
7440                  If blnFound = False Then
7450                    blnContinue = False
7460                    Exit For
7470                  End If
7480                End If
7490                DoEvents
7500                If lngX < lngRecs Then .MoveNext
7510              Next
7520            End If  ' ** BOF, EOF.
7530            .Close
7540          End With
7550          Set rst = Nothing
7560          DoEvents

7570          If blnContinue = True Then

                ' ** Step 9: Delete under-threshold entries.
7580  On Error Resume Next
7590            .Map_Status_lbl.Caption = " 9 of 22: Delete under-threshold entries . . ."
7600  On Error GoTo ERRH
7610            DoEvents

                ' ** Delete qryMap_Misc_LTCL_11a (Journal Map, for shareface < 0.00006
                ' ** and > -0.00006, icash < 0.006 and > -0.006, by specified [astno]).
7620            Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_11b")
7630            With qdf.Parameters
7640              ![astno] = lngAssetNo
7650            End With
7660            qdf.Execute
7670            Set qdf = Nothing
7680            DoEvents

                ' ** Step 10: Sum distributions.
7690  On Error Resume Next
7700            .Map_Status_lbl.Caption = "10 of 22: Sum distributions . . ."
7710  On Error GoTo ERRH
7720            DoEvents

7730            Select Case strCallingForm
                Case "frmJournal"
                  ' ** qryMap_Misc_LTCL_12c (Journal Map, linked to Account, qryMap_Misc_LTCL_12b
                  ' ** (qryMap_Misc_LTCL_12a (Journal Map, grouped by accountno, assetno, Location_ID),
                  ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max),
                  ' ** grouped and summed, by accountno), grouped and summed, by assetno, by specified [astno].
7740              Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_12d")
7750            Case "frmJournal_Columns"
                  ' ** qryMap_Misc_LTCL_62c (Journal Map, linked to Account, qryMap_Misc_LTCL_62b
                  ' ** (qryMap_Misc_LTCL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
                  ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max)),
                  ' ** grouped and summed, by accountno), grouped and summed, by assetno, by specified [astno]
7760              Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_62d")
7770            End Select
7780            With qdf.Parameters
7790              ![astno] = lngAssetNo
7800            End With
7810            Set rst = qdf.OpenRecordset
7820            With rst
7830              .MoveFirst
7840              dblSumPCash = ![PCash]
7850              dblSumShareface = ![shareface]
7860              dblPerShare = ![pershare]  ' ** Max([pershare]), though they should all be identical.
7870              .Close
7880            End With
7890            Set rst = Nothing
7900            Set qdf = Nothing
7910            DoEvents

7920            dblTmp02 = (Round(dblAmount, 2) - dblSumPCash)
7930            If ((dblTmp02 > 0.009) Or (dblTmp02 < -0.009)) And (dblTmp02 > 1#) Then
                  ' ** There's an Over/Under of more than $1.00, so let's see if we can bring it down.

                  ' ** Step 11: Rounding Stage 2.
7940  On Error Resume Next
7950              .Map_Status_lbl.Caption = "11 of 22: Rounding Stage 2 . . ."
7960  On Error GoTo ERRH
7970              DoEvents

7980              Select Case intInputType
                  Case 1  ' ** User entered an Amount.
                    ' ** qryMap_Misc_LTCL_15a (Journal Map, linked to qryMap_Misc_LTCL_14a (Journal Map,
                    ' ** linked to qryMap_Misc_LTCL_13a (qryMap_Misc_LTCL_10a (qryMap_Misc_LTCL_09a
                    ' ** (qryMap_Misc_LTCL_08a (Journal Map, with qryMap_Misc_LTCL_07 (ActiveAssets,
                    ' ** grouped and summed, by assetno, by specified [astno]), with calc'd fields, by
                    ' ** specified [astno], [amt], [prec]), grouped and summed by account), with .._rnd
                    ' ** rounded again), just needed fields, by specified [thrsh]), with qryMap_Misc_LTCL_07
                    ' ** (ActiveAssets, grouped and summed, by assetno, by specified [astno]), with
                    ' ** icash_adj, specified [astno], [thrsh], [pcashou]; Cartesian), with pcash_adjx),
                    ' ** just entries with adjustment.
7990                Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_16a")
8000                With qdf.Parameters
8010                  ![astno] = lngAssetNo
8020                  ![amt] = dblAmount
8030                  ![prec] = gintShareFaceDecimals
8040                  dblTmp03 = (0.01 / (dblTmp02 / dblShareface))  ' ** Threshhold needed for a penny adjustment.
8050                  ![thrsh] = dblTmp03
8060                  ![pcashou] = dblTmp02
8070                End With
8080              Case 2  ' ** User entered a PerShare.
                    ' ** qryMap_Misc_LTCL_15b (Journal Map, linked to qryMap_Misc_LTCL_14b (Journal Map,
                    ' ** linked to qryMap_Misc_LTCL_13b (qryMap_Misc_LTCL_10b (qryMap_Misc_LTCL_09b
                    ' ** (qryMap_Misc_LTCL_08b (Journal Map, with qryMap_Misc_LTCL_07 (ActiveAssets,
                    ' ** grouped and summed, by assetno, by specified [astno]), with calc'd fields, by
                    ' ** specified [astno], [persh], [prec]), grouped and summed by account), with .._rnd
                    ' ** rounded again), just needed fields, by specified [thrsh]), with qryMap_Misc_LTCL_07
                    ' ** (ActiveAssets, grouped and summed, by assetno, by specified [astno]), with
                    ' ** icash_adj, by specified [astno], [thrsh], [pcashou]; Cartesian), with pcash_adjx),
                    ' ** just entries with adjustment.
8090                Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_16b")
8100                With qdf.Parameters
8110                  ![astno] = lngAssetNo
8120                  ![persh] = dblPerShare
8130                  ![prec] = gintShareFaceDecimals
8140                  dblTmp03 = (0.01 / (dblTmp02 / dblShareface))  ' ** Threshhold needed for a penny adjustment.
8150                  ![thrsh] = dblTmp03
8160                  ![pcashou] = dblTmp02
8170                End With
8180              End Select
8190              DoEvents

                  ' ** Step 12: Collect adjustments.
8200  On Error Resume Next
8210              .Map_Status_lbl.Caption = "12 of 22: Collect adjustments . . ."
8220  On Error GoTo ERRH
8230              DoEvents

8240              lngAdjs = 0&
8250              Set rst = qdf.OpenRecordset
8260              With rst
8270                If .BOF = True And .EOF = True Then
                      ' ** Not enough to adjust.
8280                Else
8290                  .MoveLast
8300                  lngAdjs = .RecordCount
8310                  .MoveFirst
8320                  arr_varAdj = .GetRows(lngAdjs)
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
8330                End If
8340                .Close
8350              End With
8360              Set rst = Nothing
8370              Set qdf = Nothing
8380              DoEvents  ' ** All these DoEvents are to keep the window from showing a 'Not Responding'!

                  ' ** Step 13: Update Journal Map.
8390  On Error Resume Next
8400              .Map_Status_lbl.Caption = "13 of 22: Update temporary table . . ."
8410  On Error GoTo ERRH
8420              DoEvents

8430              Set rst = dbs.OpenRecordset("Journal Map", dbOpenDynaset, dbConsistent)
8440              With rst
8450                .MoveFirst
8460                For lngX = 0& To (lngAdjs - 1&)
8470                  .FindFirst "[ID] = " & CStr(arr_varAdj(A_ID, lngX))
8480                  If .NoMatch = False Then
8490                    .Edit
8500                    ![PCash] = arr_varAdj(A_PCADJ, lngX)
8510                    .Update
8520                  Else
8530                    blnContinue = False
8540                    Exit For
8550                  End If
8560                Next
8570              End With
8580              Set rst = Nothing
8590              Set qdf = Nothing
8600              DoEvents

                  ' ** Step 14: Confirm Rounding Stage 2.
8610  On Error Resume Next
8620              .Map_Status_lbl.Caption = "14 of 22: Confirm Rounding Stage 2 . . ."
8630  On Error GoTo ERRH
8640              DoEvents

8650              Select Case strCallingForm
                  Case "frmJournal"
                    ' ** Now check the Over/Under again!
                    ' ** qryMap_Misc_LTCL_12c (Journal Map, linked to Account, qryMap_Misc_LTCL_12b
                    ' ** (qryMap_Misc_LTCL_12a (Journal Map, grouped by accountno, assetno, Location_ID),
                    ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max),
                    ' ** grouped and summed, by accountno), grouped and summed, by assetno, by specified [astno].
8660                Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_12d")
8670              Case "frmJournal_Columns"
                    ' ** qryMap_Misc_LTCL_62c (Journal Map, linked to Account, qryMap_Misc_LTCL_62b
                    ' ** (qryMap_Misc_LTCL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
                    ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max)),
                    ' ** grouped and summed, by accountno), grouped and summed, by assetno, by specified [astno].
8680                Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_62d")
8690              End Select
8700              With qdf.Parameters
8710                ![astno] = lngAssetNo
8720              End With
8730              Set rst = qdf.OpenRecordset
8740              With rst
8750                .MoveFirst
8760                dblSumPCash = ![PCash]
8770                dblSumShareface = ![shareface]
8780                dblPerShare = ![pershare]  ' ** Max([pershare]), though they should all be identical.
8790                .Close
8800              End With
8810              Set rst = Nothing
8820              Set qdf = Nothing
8830              DoEvents

8840              dblTmp02 = (Round(dblAmount, 2) - dblSumPCash)

8850            End If

                ' ** Step 15: Create Over/Under entry.
8860  On Error Resume Next
8870            .Map_Status_lbl.Caption = "15 of 22: Create Over/Under entry . . ."
8880  On Error GoTo ERRH
8890            DoEvents

8900            If ((dblTmp02 > 0.009) Or (dblTmp02 < -0.009)) Then
                  ' ** Create a new record in Journal Map.
8910              Set rst = dbs.OpenRecordset("journal map", dbOpenDynaset, dbAppendOnly)
8920              With rst
8930                .AddNew
8940                ![assetno] = lngAssetNo
8950                ![accountno] = strAccountNo_OU
8960                ![shareface] = Round(Abs(dblShareface - dblSumShareface), gintShareFaceDecimals)
8970                ![pershare] = dblPerShare
8980                ![transdate] = datTransDate
8990                ![journaltype] = strMapType
9000                dblOverUnder = Round((Round(dblAmount, 2) - dblSumPCash), 2)  ' ** Module Function: modUtilities.
9010                ![PCash] = CCur(Round(dblOverUnder, 2))
9020                ![description] = "Over/Short of " & strMapType & " Mapping for " & Me.cmbAssets.Column(CBX_A_DESC)
9030                ![curr_id] = lngCurrID
9040                .Update
9050                .Close
9060              End With
9070              Set rst = Nothing
9080              .overshort = dblOverUnder
9090            Else
9100              .overshort = 0#
9110            End If

9120            dbs.Close
9130            Set dbs = Nothing
9140            DoEvents

                ' ** Step 16: Delete zero entries.
9150  On Error Resume Next
9160            .Map_Status_lbl.Caption = "16 of 22: Delete zero entries . . ."
9170  On Error GoTo ERRH
9180            DoEvents

                ' ################################################################################
9190            If blnSkip = False Then
9200              Set dbs = CurrentDb
                  ' ** Journal Map, just zero shareface, icash, pcash.
9210              Set qdf = dbs.QueryDefs("qryMap_Div_01a")
9220              Set rst = qdf.OpenRecordset
9230              If rst.BOF = True And rst.EOF = True Then
                    ' ** All's well.
9240                rst.Close
9250              Else
9260                rst.Close
                    ' ** Delete qryMap_Div_01a (Journal Map, just zero shareface, icash, pcash).
9270                Set qdf = dbs.QueryDefs("qryMap_Div_01b")
9280                qdf.Execute
9290              End If
9300              dbs.Close
9310            End If  ' ** blnSkip.
                ' ################################################################################

9320            dblOverUnder = AdvancedRounding(lngAssetNo)  ' ** Procedure: Below.
9330            If dblOverUnder = 0# And .overshort <> 0 Then
                  ' ** Not sure why this disagrees.
9340              dblOverUnder = .overshort
9350            End If

                ' ** Step 22: Display detail.
9360  On Error Resume Next
9370            .Map_Status_lbl.Caption = "22 of 22: Display detail . . ."
9380  On Error GoTo ERRH
9390            DoEvents

                ' ** Borrowing these variables from the Court Reports.
9400            gstrCrtRpt_Ordinal = .cmbAssets.Column(CBX_A_CUSIP)  ' ** cusip.
9410            strTmp01 = .cmbAssets.Column(CBX_A_TDESC)            ' ** totdesc.
9420            intPos01 = InStr(strTmp01, Chr(160))  ' ** Hard space.
9430            If intPos01 > 0 Then
9440              strTmp01 = Mid(strTmp01, (intPos01 + 1))
9450              Do While Left(strTmp01, 1) = Chr(160)
9460                strTmp01 = Mid(strTmp01, 2)
9470              Loop
9480            End If
9490            gstrCrtRpt_Version = strTmp01

9500            strDocName = "frmMap_Misc_LTCL_Detail"
9510            DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & CStr(lngAssetNo) & "~" & _
                  CStr(dblOverUnder) & "~" & CStr(dblPerShare) & "~" & Format(datTransDate, "mm/dd/yyyy") & "~" & _
                  CStr(dblAmount) & "~" & CStr(lngCurrID)

9520            DoCmd.Close acForm, THIS_NAME

9530          Else
9540            Select Case gblnGoToReport
                Case True
9550              Beep
9560              DoCmd.Hourglass False
9570              MsgBox "The report site cannot be shown because there are no records.", _
                    vbInformation + vbOKOnly, "Record Not Found"
9580              blnGoingToReport = False
9590              .GoToReport_arw_ok_img.Visible = False
9600              gblnGoToReport = False
9610            Case False
9620              Beep
9630              DoCmd.Hourglass False
9640              MsgBox "This procedure cannot continue." & vbCrLf & "Expected records were not found.", _
                    vbInformation + vbOKOnly, "Record Not Found"
9650              blnGoingToReport = False
9660              .GoToReport_arw_ok_img.Visible = False
9670              gblnGoToReport = False
9680            End Select
9690            .cmdCancel.SetFocus
9700          End If  ' ** blnContinue.

9710        Else
9720          .cmdCancel.SetFocus
9730        End If  ' ** msgResponse.

9740      End If

9750    End With

EXITP:
9760    Set rsx1 = Nothing
9770    Set cnxn = Nothing
9780    Set rst = Nothing
9790    Set qdf = Nothing
9800    Set dbs = Nothing
9810    Exit Sub

ERRH:
9820    DoCmd.Hourglass False
9830    THAT_PROC = THIS_PROC
9840    That_Erl = Erl
9850    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9860    Resume EXITP

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
10440     Set qdf = .QueryDefs("qryMap_Misc_LTCL_08_01")
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
10650       gstrReturningForm = THIS_NAME
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
14580       .amount.BackColor = lngTmp02
14590       .pershare.Format = strFrmt2
14600       .pershare.DecimalPlaces = intDec2
14610       .pershare.BackColor = lngTmp02

14620       If lngCurrID = 150& Then
14630         .Amount_usd.Visible = False
14640         .pershare_usd.Visible = False
14650       Else
14660         .Amount_usd = 0
14670         .Amount_usd.Visible = True
14680         .pershare_usd = 0
14690         .pershare_usd.Visible = True
14700       End If

14710     End If

14720   End With

EXITP:
14730   Exit Sub

ERRH:
14740   THAT_PROC = THIS_PROC
14750   That_Erl = Erl
14760   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14770   Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

14810   With Me
14820     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
14830     .curr_id.SetFocus
14840     Response = acDataErrContinue
14850   End With

EXITP:
14860   Exit Sub

ERRH:
14870   THAT_PROC = THIS_PROC
14880   That_Erl = Erl
14890   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14900   Resume EXITP

End Sub

Private Sub Amount_KeyDown(KeyCode As Integer, Shift As Integer)

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "Amount_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15020   intShiftDown = (Shift And acShiftMask) > 0
15030   intAltDown = (Shift And acAltMask) > 0
15040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15070       With Me
15080         intRetVal = 0
15090         If .pershare.Enabled = True Then
15100 On Error Resume Next
15110           .pershare.SetFocus
15120           If ERR.Number <> 0 Then
15130 On Error GoTo ERRH
15140             .transdate.SetFocus
15150           Else
15160 On Error GoTo ERRH
15170           End If
15180         Else
15190           .transdate.SetFocus
15200         End If
15210       End With
15220     End Select
15230   End If

        ' ** Shift keys.
15240   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15250     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15260       With Me
15270         intRetVal = 0
15280         blnDrivingMeBackwards = True
15290         .cmbAssets.SetFocus
15300       End With
15310     End Select
15320   End If

EXITP:
15330   KeyCode = intRetVal
15340   Exit Sub

ERRH:
15350   intRetVal = 0
15360   THAT_PROC = THIS_PROC
15370   That_Erl = Erl
15380   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15390   Resume EXITP

End Sub

Private Sub amount_Exit(Cancel As Integer)

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "amount_Exit"

        Dim dblRate2 As Double
        Dim intPos01 As Integer
        Dim blnContinue As Boolean
        Dim dblTmp01 As Double

15410   With Me

15420     blnContinue = True

15430     lngCurrID = .curr_id.Column(CBX_C_CURRID)

15440     If IsNumeric(.amount) = True Then
15450       If .amount > 0 Then
15460         .amount = -1 * .amount
15470       End If
15480       If CDbl(Rem_Dollar(.amount.text, lngCurrID)) < -99999999.99 Then  ' ** Module Function: modStringFuncs.
15490         blnContinue = False
15500         MsgBox "Amount is too small.", vbInformation + vbOKOnly, "Invalid Entry"
15510         .amount = vbNullString
15520         Cancel = -1
15530         .amount.SetFocus
15540       Else
15550         intPos01 = InStr(1, .amount, ".", vbTextCompare)
15560         If intPos01 > 0 Then
15570           If Len(.amount) - intPos01 > 2 Then
15580             blnContinue = False
15590             MsgBox "You can only enter Amount with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
15600             .amount = vbNullString
15610             Cancel = -1
15620           End If
15630         End If
15640         If blnContinue = True Then
15650           If Nz(.shareface, 0) = 0 Then
15660             blnContinue = False
15670             MsgBox "Shares must have a value", vbInformation + vbOKOnly, "Invalid Entry"
15680             .shareface = 1
15690             .cmbAssets.SetFocus
15700             If lngCurrID = 150& Then
15710               .pershare = (.amount / IIf(Nz(.shareface, 0) = 0, 1, .shareface))  '5 dec.
15720               .pershare.BackColor = CLR_DISABLED_BG
15730             Else
15740               .pershare = (.amount / IIf(Nz(.shareface, 0) = 0, 1, .shareface))  '5 dec.
15750               .pershare.BackColor = lngBkClr2
15760             End If
15770             .pershare.Locked = True
15780             .pershare.Enabled = False
15790             .pershare.ForeColor = CLR_DISABLED_FG
15800           Else
15810             If blnDrivingMeBackwards = False Then
15820               If gblnGoToReport = False Then
15830                 .transdate.SetFocus
15840               Else
15850 On Error Resume Next
15860                 .cmdOK.SetFocus
15870 On Error GoTo ERRH
15880               End If
15890             Else
15900               blnDrivingMeBackwards = False
15910             End If
15920             intInputType = 1
15930             dblPerShare_Calc = (.amount / IIf(Nz(.shareface, 0) = 0, 1, .shareface))
15940             If lngCurrID = 150& Then
15950               .pershare = dblPerShare_Calc  '5 dec.
15960               .pershare.BackColor = CLR_DISABLED_BG_R
15970               .amount.BackColor = CLR_DISABLED_BG_VLR
15980             Else
15990               .pershare = dblPerShare_Calc  '5 dec.
16000               .pershare.BackColor = lngBkClr2
16010               .amount.BackColor = lngBkClr1
16020               dblRate2 = .curr_id.Column(CBX_C_RATE2)
16030               dblTmp01 = (.amount * dblRate2)
16040               .Amount_usd = Round(dblTmp01, 2)
16050               .Amount_usd.Visible = True
16060               dblTmp01 = (dblPerShare_Calc * dblRate2)
16070               dblTmp01 = Round(dblTmp01, 5)
16080               .pershare_usd = dblTmp01
16090               .pershare_usd.Visible = True
16100             End If
16110             .pershare.Locked = True
16120             .pershare.Enabled = False
16130             .pershare.ForeColor = CLR_DISABLED_FG
16140           End If
16150         End If
16160       End If
16170     Else
16180       .pershare.Enabled = True
16190       .pershare.Locked = False
16200       .pershare.ForeColor = CLR_BLK
16210       If lngCurrID = 150& Then
16220         .pershare.BackColor = CLR_WHT
16230       Else
16240         .pershare.BackColor = CLR_VLTGRN
16250       End If
16260     End If

16270   End With

EXITP:
16280   Exit Sub

ERRH:
16290   Select Case ERR.Number
        Case Else
16300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16310   End Select
16320   Resume EXITP

End Sub

Private Sub pershare_Exit(Cancel As Integer)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "pershare_Exit"

        Dim dblRate2 As Double
        Dim intPos01 As Integer
        Dim blnContinue As Boolean
        Dim dblTmp01 As Double

16410   With Me

16420     blnContinue = True

16430     lngCurrID = .curr_id.Column(CBX_C_CURRID)

16440     If IsNumeric(.pershare) = True Then
16450       If .pershare > 0 Then
16460         .pershare = -1 * .pershare
16470       End If
16480       If CDbl(Rem_Dollar(.pershare.text, lngCurrID)) > 999999999.9999 Then  ' ** Module Function: modStringFuncs.
16490         blnContinue = False
16500         MsgBox "Per Share is too large.", vbInformation + vbOKOnly, "Invalid Entry"
16510         .pershare = vbNullString
16520         Cancel = -1
16530         .pershare.SetFocus
16540       Else
16550         intPos01 = InStr(1, .pershare, ".", vbTextCompare)
16560         If intPos01 > 0 Then
16570           If Len(.pershare) - intPos01 > 5 Then
16580             blnContinue = False
16590             MsgBox "You can only enter Per Share with up to 5 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
16600             .pershare = vbNullString
16610             Cancel = -1
16620             .pershare.SetFocus
16630           End If
16640         End If
16650         If blnContinue = True Then
16660           intInputType = 2
16670           If lngCurrID = 150& Then
16680             .amount = (.pershare * Nz(.shareface, 1))
16690             .amount.BackColor = CLR_DISABLED_BG_R
16700             .pershare.BackColor = CLR_DISABLED_BG_VLR
16710           Else
16720             .amount = (.pershare * Nz(.shareface, 1))
16730             .amount.BackColor = lngBkClr2
16740             .pershare.BackColor = lngBkClr1
16750             dblRate2 = .curr_id.Column(CBX_C_RATE2)
16760             dblTmp01 = (.pershare * dblRate2)
16770             dblTmp01 = Round(dblTmp01, 5)
16780             .pershare_usd = dblTmp01
16790             .pershare_usd.Visible = True
16800             dblTmp01 = (.amount * dblRate2)
16810             dblTmp01 = Round(dblTmp01, 2)
16820             .Amount_usd = dblTmp01
16830             .Amount_usd.Visible = True
16840           End If
16850           .amount.Locked = True
16860           .amount.Enabled = False
16870           .amount.ForeColor = CLR_DISABLED_FG
16880         End If
16890       End If
16900     Else
16910       .amount.Enabled = True
16920       .amount.Locked = False
16930       .amount.ForeColor = CLR_BLK
16940       If lngCurrID = 150& Then
16950         .amount.BackColor = CLR_WHT
16960       Else
16970         .amount.BackColor = CLR_VLTGRN
16980       End If
16990     End If

17000   End With

EXITP:
17010   Exit Sub

ERRH:
17020   Select Case ERR.Number
        Case Else
17030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17040   End Select
17050   Resume EXITP

End Sub

Private Sub transdate_KeyDown(KeyCode As Integer, Shift As Integer)

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17120   intShiftDown = (Shift And acShiftMask) > 0
17130   intAltDown = (Shift And acAltMask) > 0
17140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17170       With Me
17180         intRetVal = 0
17190         .cmdOK.SetFocus
17200       End With
17210     End Select
17220   End If

        ' ** Shift keys.
17230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17250       With Me
17260         intRetVal = 0
17270         If .pershare.Enabled = True Then
17280           .pershare.SetFocus
17290         Else
17300           .amount.SetFocus
17310         End If
17320       End With
17330     End Select
17340   End If

        ' ** Ctrl keys.
17350   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
17360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17370       With Me
17380         intRetVal = 0
17390         .cmdCalendar.SetFocus
17400       End With
17410     End Select
17420   End If

EXITP:
17430   KeyCode = intRetVal
17440   Exit Sub

ERRH:
17450   intRetVal = 0
17460   THAT_PROC = THIS_PROC
17470   That_Erl = Erl
17480   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17490   Resume EXITP

End Sub

Private Sub transdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_MouseMove"

17510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17520   Exit Sub

ERRH:
17530   Select Case ERR.Number
        Case Else
17540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17550   End Select
17560   Resume EXITP

End Sub

Private Sub transdate_Exit(Cancel As Integer)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_Exit"

        Dim varTmp00 As Variant

17610   With Me
17620 On Error Resume Next
17630     varTmp00 = .transdate.text  ' ** When called from cmdCalendar_Click(), below, it's not always on focus.
17640     If ERR.Number <> 0 Then
17650       varTmp00 = .transdate
17660     End If
17670 On Error GoTo ERRH
17680     If DateCheck_Post(varTmp00) = True Then  ' ** Module Function: modUtilities.
            ' ** '.Text' needed to reliably check date.
17690     Else
17700       If IsNull(.transdate) = True Then
17710         .transdate = Date
17720       Else
17730         If Trim(.transdate) = vbNullString Then
17740           .transdate = Date
17750         Else
17760           Cancel = -1
17770           .transdate.SetFocus
17780         End If
17790       End If
17800     End If
17810   End With

EXITP:
17820   Exit Sub

ERRH:
17830   Select Case ERR.Number
        Case Else
17840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17850   End Select
17860   Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

17910   With Me
17920     blnCalendar1_Focus = True
17930     .cmdCalendar_raised_semifocus_dots_img.Visible = True
17940     .cmdCalendar_raised_img.Visible = False
17950     .cmdCalendar_raised_focus_img.Visible = False
17960     .cmdCalendar_raised_focus_dots_img.Visible = False
17970     .cmdCalendar_sunken_focus_dots_img.Visible = False
17980     .cmdCalendar_raised_img_dis.Visible = False
17990   End With

EXITP:
18000   Exit Sub

ERRH:
18010   Select Case ERR.Number
        Case Else
18020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18030   End Select
18040   Resume EXITP

End Sub

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

18110   With Me
18120     blnCalendar1_MouseDown = True
18130     .cmdCalendar_sunken_focus_dots_img.Visible = True
18140     .cmdCalendar_raised_img.Visible = False
18150     .cmdCalendar_raised_semifocus_dots_img.Visible = False
18160     .cmdCalendar_raised_focus_img.Visible = False
18170     .cmdCalendar_raised_focus_dots_img.Visible = False
18180     .cmdCalendar_raised_img_dis.Visible = False
18190   End With

EXITP:
18200   Exit Sub

ERRH:
18210   Select Case ERR.Number
        Case Else
18220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18230   End Select
18240   Resume EXITP

End Sub

Private Sub cmdCalendar_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

18310   With Me
18320     datStartDate = Date
18330     datEndDate = 0
18340     blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
18350     If blnRetVal = True Then
            ' ** Allow posting up to 1 month into the future.
18360       If datStartDate > DateAdd("m", 1, Date) Then
18370         MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
18380         .transdate = CDate(Format(Date, "mm/dd/yyyy"))
18390       Else
18400         .transdate = datStartDate
18410       End If
18420     Else
18430       .transdate = CDate(Format(Date, "mm/dd/yyyy"))
18440     End If
          ' ** Make the date 'is edited' as if were entered by hand.
18450     .transdate.SetFocus
18460     Cancel = 0
18470     transdate_Exit Cancel  ' ** Procedure: Above.
18480     If Cancel = 0 Then
18490       .cmdOK.SetFocus
18500     End If
18510   End With

EXITP:
18520   Exit Sub

ERRH:
18530   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
18540   Case Else
18550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18560   End Select
18570   Resume EXITP

End Sub

Private Sub cmdCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseMove"

18610   With Me
18620     If blnCalendar1_MouseDown = False Then
18630       Select Case blnCalendar1_Focus
            Case True
18640         .cmdCalendar_raised_focus_dots_img.Visible = True
18650         .cmdCalendar_raised_focus_img.Visible = False
18660       Case False
18670         .cmdCalendar_raised_focus_img.Visible = True
18680         .cmdCalendar_raised_focus_dots_img.Visible = False
18690       End Select
18700       .cmdCalendar_raised_img.Visible = False
18710       .cmdCalendar_raised_semifocus_dots_img.Visible = False
18720       .cmdCalendar_sunken_focus_dots_img.Visible = False
18730       .cmdCalendar_raised_img_dis.Visible = False
18740     End If
18750   End With

EXITP:
18760   Exit Sub

ERRH:
18770   Select Case ERR.Number
        Case Else
18780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18790   End Select
18800   Resume EXITP

End Sub

Private Sub cmdCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseUp"

18910   With Me
18920     .cmdCalendar_raised_focus_dots_img.Visible = True
18930     .cmdCalendar_raised_img.Visible = False
18940     .cmdCalendar_raised_semifocus_dots_img.Visible = False
18950     .cmdCalendar_raised_focus_img.Visible = False
18960     .cmdCalendar_sunken_focus_dots_img.Visible = False
18970     .cmdCalendar_raised_img_dis.Visible = False
18980     blnCalendar1_MouseDown = False
18990   End With

EXITP:
19000   Exit Sub

ERRH:
19010   Select Case ERR.Number
        Case Else
19020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19030   End Select
19040   Resume EXITP

End Sub

Private Sub cmdCalendar_LostFocus()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_LostFocus"

19110   With Me
19120     .cmdCalendar_raised_img.Visible = True
19130     .cmdCalendar_raised_semifocus_dots_img.Visible = False
19140     .cmdCalendar_raised_focus_img.Visible = False
19150     .cmdCalendar_raised_focus_dots_img.Visible = False
19160     .cmdCalendar_sunken_focus_dots_img.Visible = False
19170     .cmdCalendar_raised_img_dis.Visible = False
19180     blnCalendar1_Focus = False
19190   End With

EXITP:
19200   Exit Sub

ERRH:
19210   Select Case ERR.Number
        Case Else
19220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19230   End Select
19240   Resume EXITP

End Sub

Private Function AdvancedRounding(lngAssetNo As Long) As Double

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "AdvancedRounding"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim lngRecs As Long
        Dim lngLocationID As Long
        Dim blnNoDistribution As Boolean
        Dim lngX As Long
        Dim dblRetVal As Double

19310   blnNoDistribution = False
19320   dblRetVal = 0#

19330   Set dbs = CurrentDb
19340   With dbs

          ' ** Step 17: Rounding Stage 3.
19350     Me.Map_Status_lbl.Caption = "17 of 22: Rounding Stage 3 . . ."
19360     DoEvents

19370     Select Case strCallingForm
          Case "frmJournal"
            ' ** qryMap_Misc_LTCL_12c (Journal Map, linked to Account, qryMap_Misc_LTCL_12b
            ' ** (qryMap_Misc_LTCL_12a (Journal Map, grouped by accountno, assetno, Location_ID),
            ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max),
            ' ** grouped and summed, by accountno), just 'INCOME O/U', linked to qryMap_Misc_LTCL_17b
            ' ** (qryMap_Misc_LTCL_17a (qryMap_Misc_LTCL_12c (Journal Map, linked to Account,
            ' ** qryMap_Misc_LTCL_12b (qryMap_Misc_LTCL_12a (Journal Map, grouped by accountno,
            ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
            ' ** Location_ID_max), grouped and summed, by accountno), without 'INCOME O/U', with
            ' ** total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), grouped and
            ' ** summed, by assetno, with cnt), with pcash_per_acct, by specified [actno].
19380       Set qdf = .QueryDefs("qryMap_Misc_LTCL_17c")
19390     Case "frmJournal_Columns"
            ' ** qryMap_Misc_LTCL_62c (Journal Map, linked to Account, qryMap_Misc_LTCL_62b
            ' ** (qryMap_Misc_LTCL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
            ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max)),
            ' ** grouped and summed, by accountno), just 'INCOME O/U', linked to qryMap_Misc_LTCL_67b
            ' ** (qryMap_Misc_LTCL_67a (qryMap_Misc_LTCL_62c (Journal Map, linked to Account,
            ' ** qryMap_Misc_LTCL_62b (qryMap_Misc_LTCL_62a (Journal Map, grouped by accountno,
            ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
            ' ** Location_ID_max)), grouped and summed, by accountno), without 'INCOME O/U', with
            ' ** total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), grouped and
            ' ** summed, by assetno, with cnt), with pcash_per_acct, by specified [actno].
19400       Set qdf = .QueryDefs("qryMap_Misc_LTCL_67c")

19410     End Select
19420     With qdf.Parameters
19430       ![actno] = strAccountNo_OU
19440     End With
19450     Set rst1 = qdf.OpenRecordset
19460     With rst1
19470       If .BOF = True And .EOF = True Then
              ' ** No INCOME O/U to distribute!
19480         blnNoDistribution = True
19490       Else
19500         .MoveFirst
19510         If ![total_pcash_ou] < 1@ Then
                ' ** Less than a buck, leave it alone.
19520           blnNoDistribution = True
19530         Else
19540           If ![pennies_per_acct] < 0.01 Then
                  ' ** Too small to divvy up!
19550             blnNoDistribution = True
19560           Else
19570             If ![pcash_per_acct] < 0.01 Then
                    ' ** Rounded value too small to divvy up!
19580               blnNoDistribution = True
19590             Else
                    ' ** Continue with the distribution.
19600             End If
19610           End If
19620         End If
19630       End If
19640       .Close
19650     End With  ' ** rst1.
19660     Set rst1 = Nothing
19670     Set qdf = Nothing
19680     DoEvents

          ' ** Step 18: Empty tblJournal_Map_Staging2.
19690     Me.Map_Status_lbl.Caption = "18 of 22: Empty temporary table 2 . . ."
19700     DoEvents

19710     Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty tblJournal_Map_Staging2.
19720       Set qdf = .QueryDefs("qryMap_Misc_LTCL_18")
19730     Case "frmJournal_Columns"
            ' ** Empty tblJournal_Map_Staging2.
19740       Set qdf = .QueryDefs("qryMap_Misc_LTCL_68")
19750     End Select
19760     qdf.Execute
19770     Set qdf = Nothing
19780     DoEvents

19790     Select Case blnNoDistribution
          Case True

            ' ** Step 19: Collect unchanged distributions.
19800       Me.Map_Status_lbl.Caption = "19 of 22: Collect unchanged distributions . . ."
19810       DoEvents

19820       Select Case strCallingForm
            Case "frmJournal"
              ' ** Append qryMap_Misc_LTCL_12c (Journal Map, linked to Account, qryMap_Misc_LTCL_12b
              ' ** (qryMap_Misc_LTCL_12a (Journal Map, grouped by accountno, assetno, Location_ID),
              ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max),
              ' ** grouped and summed, by accountno) to tblJournal_Map_Staging2.
19830         Set qdf = .QueryDefs("qryMap_Misc_LTCL_19")
19840       Case "frmJournal_Columns"
              ' ** Append qryMap_Misc_LTCL_62c (Journal Map, linked to Account, qryMap_Misc_LTCL_62b
              ' ** (qryMap_Misc_LTCL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
              ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max)),
              ' ** grouped and summed, by accountno) to tblJournal_Map_Staging2.
19850         Set qdf = .QueryDefs("qryMap_Misc_LTCL_69")
19860       End Select
19870       qdf.Execute
19880       Set qdf = Nothing

19890     Case False

19900       Set rst2 = .OpenRecordset("tblJournal_Map_Staging2", dbOpenDynaset, dbConsistent)

            ' ** Step 19: Create new distributions.
19910       Me.Map_Status_lbl.Caption = "19 of 22: Create new distributions . . ."
19920       DoEvents

19930       Select Case strCallingForm
            Case "frmJournal"
              ' ** qryMap_Misc_LTCL_17a (qryMap_Misc_LTCL_12c (Journal Map, linked to Account,
              ' ** qryMap_Misc_LTCL_12b (qryMap_Misc_LTCL_12a (Journal Map, grouped by accountno,
              ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
              ' ** Location_ID_max), grouped and summed, by accountno), without 'INCOME O/U', with
              ' ** total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), linked to
              ' ** qryMap_Misc_LTCL_20a (qryMap_Misc_LTCL_17c (qryMap_Misc_LTCL_12c (Journal Map,
              ' ** linked to Account, qryMap_Misc_LTCL_12b (qryMap_Misc_LTCL_12a (Journal Map,
              ' ** grouped by accountno, assetno, Location_ID), grouped by accountno, assetno, with
              ' ** cnt, Location_ID_min, Location_ID_max), grouped and summed, by accountno), just
              ' ** 'INCOME O/U', linked to qryMap_Misc_LTCL_17b (qryMap_Misc_LTCL_17a
              ' ** (qryMap_Misc_LTCL_12c (Journal Map, linked to Account, qryMap_Misc_LTCL_12b
              ' ** (qryMap_Misc_LTCL_12a (Journal Map, grouped by accountno, assetno, Location_ID),
              ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max),
              ' ** grouped and summed, by accountno), without 'INCOME O/U', with total_pcash_rnd,
              ' ** total_pcash_dbl, gain_loss, by specified [actno]), grouped and summed, by
              ' ** assetno, with cnt), with pcash_per_acct, by specified [actno]), with
              ' ** shareface_per_acct), with total_shareface_new, total_pcash_new.
19940         Set qdf = .QueryDefs("qryMap_Misc_LTCL_20b")
19950       Case "frmJournal_Columns"
              ' ** qryMap_Misc_LTCL_67a (qryMap_Misc_LTCL_62c (Journal Map, linked to Account,
              ' ** qryMap_Misc_LTCL_62b (qryMap_Misc_LTCL_62a (Journal Map, grouped by accountno,
              ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
              ' ** Location_ID_max)), grouped and summed, by accountno), without 'INCOME O/U', with
              ' ** total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), linked to
              ' ** qryMap_Misc_LTCL_70a (qryMap_Misc_LTCL_67c (qryMap_Misc_LTCL_62c (Journal Map,
              ' ** linked to Account, qryMap_Misc_LTCL_62b (qryMap_Misc_LTCL_62a (Journal Map,
              ' ** grouped by accountno, assetno, Location_ID), grouped by accountno, assetno, with
              ' ** cnt, Location_ID_min, Location_ID_max)), grouped and summed, by accountno), just
              ' ** 'INCOME O/U', linked to qryMap_Misc_LTCL_67b (qryMap_Misc_LTCL_67a
              ' ** (qryMap_Misc_LTCL_62c (Journal Map, linked to Account, qryMap_Misc_LTCL_62b
              ' ** (qryMap_Misc_LTCL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
              ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max)),
              ' ** grouped and summed, by accountno), without 'INCOME O/U', with total_pcash_rnd,
              ' ** total_pcash_dbl, gain_loss, by specified [actno]), grouped and summed, by
              ' ** assetno, with cnt), with pcash_per_acct, by specified [actno]), with
              ' ** shareface_per_acct), with total_shareface_new, total_pcash_new
19960         Set qdf = .QueryDefs("qryMap_Misc_LTCL_70b")
19970       End Select
19980       With qdf.Parameters
19990         ![actno] = strAccountNo_OU
20000       End With
20010       Set rst1 = qdf.OpenRecordset
20020       With rst1
20030         .MoveLast
20040         lngRecs = .RecordCount
20050         .MoveFirst
20060         For lngX = 1& To lngRecs
20070           With rst2
20080             .AddNew
                  ' ** ![jmap_id] : AutoNumber.
20090             ![accountno] = rst1![accountno]
20100             ![assetno] = rst1![assetno]
20110             ![transdate] = rst1![transdate]
20120             ![shareface] = rst1![total_shareface_new]
20130             ![pershare] = rst1![pershare]
20140             ![journaltype] = "Misc."
20150             ![ICash] = 0@
20160             ![PCash] = rst1![total_pcash_new]
20170             ![Cost] = 0@
20180             ![description] = "Long Term Capital Loss"
20190             ![rate] = 0#
20200             ![due] = Null
20210             ![Location_ID] = rst1![Location_ID]
20220             .Update
20230           End With  ' ** rst2.
20240           DoEvents
20250           If lngX < lngRecs Then .MoveNext
20260         Next  ' ** lngX.
20270         .Close
20280       End With  ' ** rst1.
20290       Set rst1 = Nothing
20300       Set qdf = Nothing
20310       DoEvents

            ' ** Step 20: Determine Over/Under location.
20320       Me.Map_Status_lbl.Caption = "20 of 22: Determine Over/Under location . . ."
20330       DoEvents

20340       Select Case strCallingForm
            Case "frmJournal"
              ' ** Ledger, just 'INCOME O/U', 'Received', grouped by accountno, with
              ' ** Location_ID_min, Location_ID_max, cnt, by specified [actno], [astno].
20350         Set qdf = .QueryDefs("qryMap_Misc_LTCL_21")
20360       Case "frmJournal_Columns"
              ' ** Ledger, just 'INCOME O/U', 'Received', grouped by accountno, with
              ' ** Location_ID_min, Location_ID_max, cnt, by specified [actno], [astno].
20370         Set qdf = .QueryDefs("qryMap_Misc_LTCL_71")
20380       End Select
20390       With qdf.Parameters
20400         ![actno] = strAccountNo_OU
20410         ![astno] = lngAssetNo
20420       End With
20430       Set rst1 = qdf.OpenRecordset
20440       With rst1
20450         If .BOF = True And .EOF = True Then
20460           lngLocationID = 1&  ' ** Default.
20470         Else
20480           .MoveFirst
                ' ** CLng(IIf([cnt]>2,1, IIf([cnt]=1,[Location_ID_max],IIf([cnt]=2 And [Location_ID_min]=1,[Location_ID_max],1))))
20490           If ![cnt] > 2 Then
20500             lngLocationID = 1&
20510           Else
20520             If ![cnt] = 1 Then
20530               lngLocationID = ![Location_ID_max]
20540             Else
20550               If ![cnt] = 2 And [Location_ID_min] = 1 Then
20560                 lngLocationID = ![Location_ID_max]
20570               Else
20580                 lngLocationID = 1&
20590               End If
20600             End If
20610           End If
20620         End If
20630         .Close
20640       End With  ' ** rst1.
20650       Set rst1 = Nothing
20660       Set qdf = Nothing
20670       DoEvents

            ' ** Step 21: Create new Over/Under entry.
20680       Me.Map_Status_lbl.Caption = "22 of 22: Create new Over/Under entry . . ."
20690       DoEvents

20700       Select Case strCallingForm
            Case "frmJournal"
              ' ** qryMap_Misc_LTCL_22a (qryMap_Misc_LTCL_20b (qryMap_Misc_LTCL_17a (qryMap_Misc_LTCL_12c
              ' ** (Journal Map, linked to Account, qryMap_Misc_LTCL_12b (qryMap_Misc_LTCL_12a (Journal Map,
              ' ** grouped by accountno, assetno, Location_ID), grouped by accountno, assetno, with cnt,
              ' ** Location_ID_min, Location_ID_max), grouped and summed, by accountno), without 'INCOME O/U',
              ' ** with total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), linked to
              ' ** qryMap_Misc_LTCL_20a (qryMap_Misc_LTCL_17c (qryMap_Misc_LTCL_12c (Journal Map, linked to
              ' ** Account, qryMap_Misc_LTCL_12b (qryMap_Misc_LTCL_12a (Journal Map, grouped by accountno,
              ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
              ' ** Location_ID_max), grouped and summed, by accountno), just 'INCOME O/U', linked to
              ' ** qryMap_Misc_LTCL_17b (qryMap_Misc_LTCL_17a (qryMap_Misc_LTCL_12c (Journal Map, linked to
              ' ** Account, qryMap_Misc_LTCL_12b (qryMap_Misc_LTCL_12a (Journal Map, grouped by accountno,
              ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
              ' ** Location_ID_max), grouped and summed, by accountno), without 'INCOME O/U', with
              ' ** total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), grouped and summed, by
              ' ** assetno, with cnt), with pcash_per_acct, by specified [actno]), with shareface_per_acct),
              ' ** with total_shareface_new, total_pcash_new), grouped and summed, by assetno), linked to
              ' ** qryMap_Misc_LTCL_12d (qryMap_Misc_LTCL_12c (Journal Map, linked to Account, qryMap_Misc_LTCL_12b
              ' ** (qryMap_Misc_LTCL_12a (Journal Map, grouped by accountno, assetno, Location_ID), grouped by
              ' ** accountno, assetno, with cnt, Location_ID_min, Location_ID_max), grouped and summed, by
              ' ** accountno), grouped and summed, by assetno, by specified [astno]), with shareface_ou_new,
              ' ** pcash_ou_new, by specified [actno], [locid].
20710         Set qdf = .QueryDefs("qryMap_Misc_LTCL_22b")
20720       Case "frmJournal_Columns"
              ' ** qryMap_Misc_LTCL_72a (qryMap_Misc_LTCL_70b (qryMap_Misc_LTCL_67a (qryMap_Misc_LTCL_62c
              ' ** (Journal Map, linked to Account, qryMap_Misc_LTCL_62b (qryMap_Misc_LTCL_62a (Journal Map,
              ' ** grouped by accountno, assetno, Location_ID), grouped by accountno, assetno, with cnt,
              ' ** Location_ID_min, Location_ID_max)), grouped and summed, by accountno), without 'INCOME O/U',
              ' ** with total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), linked to
              ' ** qryMap_Misc_LTCL_70a (qryMap_Misc_LTCL_67c (qryMap_Misc_LTCL_62c (Journal Map, linked to
              ' ** Account, qryMap_Misc_LTCL_62b (qryMap_Misc_LTCL_62a (Journal Map, grouped by accountno,
              ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
              ' ** Location_ID_max)), grouped and summed, by accountno), just 'INCOME O/U', linked to
              ' ** qryMap_Misc_LTCL_67b (qryMap_Misc_LTCL_67a (qryMap_Misc_LTCL_62c (Journal Map, linked to
              ' ** Account, qryMap_Misc_LTCL_62b (qryMap_Misc_LTCL_62a (Journal Map, grouped by accountno,
              ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
              ' ** Location_ID_max)), grouped and summed, by accountno), without 'INCOME O/U', with
              ' ** total_pcash_rnd, total_pcash_dbl, gain_loss, by specified [actno]), grouped and summed, by
              ' ** assetno, with cnt), with pcash_per_acct, by specified [actno]), with shareface_per_acct),
              ' ** with total_shareface_new, total_pcash_new), grouped and summed, by assetno), linked to
              ' ** qryMap_Misc_LTCL_62d (qryMap_Misc_LTCL_62c (Journal Map, linked to Account, qryMap_Misc_LTCL_62b
              ' ** (qryMap_Misc_LTCL_62a (Journal Map, grouped by accountno, assetno, Location_ID), grouped by
              ' ** accountno, assetno, with cnt, Location_ID_min, Location_ID_max)), grouped and summed, by
              ' ** accountno), grouped and summed, by assetno, by specified [astno]), with shareface_ou_new,
              ' ** pcash_ou_new, by specified [actno], [locid].
20730         Set qdf = .QueryDefs("qryMap_Misc_LTCL_72b")
20740       End Select
20750       With qdf.Parameters
20760         ![astno] = lngAssetNo
20770         ![actno] = strAccountNo_OU
20780         ![locid] = lngLocationID
20790       End With
20800       Set rst1 = qdf.OpenRecordset
20810       With rst1
20820         .MoveFirst
20830         With rst2
20840           .AddNew
                ' ** ![jmap_id] : AutoNumber.
20850           ![accountno] = rst1![accountno]
20860           ![assetno] = rst1![assetno]
20870           ![transdate] = rst1![transdate]
20880           ![shareface] = rst1![shareface_ou_new]
20890           ![pershare] = rst1![pershare]
20900           ![journaltype] = "Paid"
20910           ![ICash] = 0@
20920           dblRetVal = rst1![pcash_ou_new]
20930           ![PCash] = rst1![pcash_ou_new]
20940           ![Cost] = 0@
20950           ![description] = "Long Term Capital Loss"
20960           ![rate] = 0#
20970           ![due] = Null
20980           ![Location_ID] = rst1![Location_ID]
20990           .Update
21000           .Close
21010         End With  ' ** rst2.
21020         .Close
21030       End With  ' ** rst1.
21040       Set rst1 = Nothing
21050       Set rst2 = Nothing
21060       Set qdf = Nothing

21070     End Select  ' ** blnNoDistribution.

21080     .Close
21090   End With  ' ** dbs.
21100   DoEvents

EXITP:
21110   Set rst1 = Nothing
21120   Set rst1 = Nothing
21130   Set qdf = Nothing
21140   Set dbs = Nothing
21150   AdvancedRounding = dblRetVal
21160   Exit Function

ERRH:
21170   DoCmd.Hourglass False
21180   dblRetVal = 0#
21190   Select Case ERR.Number
        Case Else
21200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21210   End Select
21220   Resume EXITP

End Function

Private Sub IncludeCurrency()

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngTmp01 As Long

21310   With Me
21320     .curr_id.Enabled = False
21330     .curr_id.Locked = True
21340     .curr_id.ForeColor = CLR_DISABLED_FG
21350     .curr_id.BackColor = CLR_VLTGRN
21360     Select Case blnCurrID
          Case True
21370       .curr_id.Visible = True
21380       .curr_id_cmd.Visible = True
21390       .curr_id_cmd.Enabled = True
21400       .cmbAssets_vline01.Visible = True
21410       .cmbAssets_vline02.Visible = True
21420     Case False
21430       .curr_id.Visible = False
21440       .curr_id_cmd.Visible = False
21450       .curr_id_cmd.Enabled = False
21460       .cmbAssets_vline01.Visible = False
21470       .cmbAssets_vline02.Visible = False
21480       lngTmp01 = ((.curr_id.Left + .curr_id.Width) - (.cmbAssets.Left + .cmbAssets.Width))
21490       .cmbAssets.Width = (.cmbAssets.Width + lngTmp01)
21500     End Select
21510     .Amount_usd.Visible = False
21520     .pershare_usd.Visible = False
21530   End With

        ' ** cmbAssets RowSource's:
        ' **   qryMap_Rec_05a
        ' **   qryMap_Rec_05c
        ' **   qryMap_Rec_05d
        ' **   qryMap_Rec_05f
        ' **   qryMap_Rec_05g
        ' **   qryMap_Rec_05i

EXITP:
21540   Exit Sub

ERRH:
21550   DoCmd.Hourglass False
21560   Select Case ERR.Number
        Case Else
21570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21580   End Select
21590   Resume EXITP

End Sub

Private Sub NoChar_Load()

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

21610   Set dbs = CurrentDb
21620   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
21630     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
21640     Set rst = qdf.OpenRecordset
21650     With rst
21660       .MoveLast
21670       lngNoChars = .RecordCount
21680       .MoveFirst
21690       arr_varNoChar = .GetRows(lngNoChars)
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
21700       .Close
21710     End With
21720     .Close
21730   End With

EXITP:
21740   Set rst = Nothing
21750   Set qdf = Nothing
21760   Set dbs = Nothing
21770   Exit Sub

ERRH:
21780   DoCmd.Hourglass False
21790   Select Case ERR.Number
        Case Else
21800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21810   End Select
21820   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

21910   With Me
21920     blnGTR_Emblem = False
21930     For lngX = 1& To 24&
21940       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
21950     Next
21960   End With

EXITP:
21970   Exit Sub

ERRH:
21980   Select Case ERR.Number
        Case Else
21990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22000   End Select
22010   Resume EXITP

End Sub
