VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Int"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Int"

'VGC 09/09/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Date:             D {transdate}
' **   Asset Cusip:      I {opgAssetSource_optCusip}
' **   Asset Name:       N {opgAssetSource_optName}
' **   OK:               O {cmdOK}
' **   Rate:             R {pershare}
' **   Select Asset:     S {cmbAssets}
' **   Amount:           T {Amount}
' **   Asset Type:       Y {opgAssetSource_optType}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar}

' ** MapType = "Interest"

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of the Calendar class.
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

Private strFrmt1 As String, strFrmt2 As String, intDec1 As Integer, intDec2 As Integer, lngBkClr As Long
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

              ' ** Borrowing these variables from the Court Reports.
240           gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString

250           CLR_DISABLED_FG = CLR_DKGRY
260           CLR_DISABLED_BG = CLR_LTTEAL

270           opgAssetSource_AfterUpdate  ' ** Procedure: Below.

280         Else
290           Cancel = -1
300         End If
310       Else
320         Cancel = -1
330       End If
340     End With

350     If Cancel = -1 Then
360       DoCmd.Hourglass False
370       Beep
380       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
390     End If

EXITP:
400     Exit Sub

ERRH:
410     DoCmd.Hourglass False
420     Select Case ERR.Number
        Case Else
430       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
440     End Select
450     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        'Dim cnxn As ADODB.Connection, rsx1 As ADODB.Recordset  ' ** Early binding.
        Dim cnxn As Object, rsx1 As Object                      ' ** Late binding.

510     With Me

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
520       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
530       clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
540       clsMonthClass.PositionAtCursor = True

550       gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
560       datPostingDate = Date

          'Set rsx1 = New ADODB.Recordset             ' ** Early binding.
570       Set rsx1 = CreateObject("ADODB.Recordset")  ' ** Late binding.
580       Set cnxn = CurrentProject.Connection
590       rsx1.Open "PostingDate", cnxn, adOpenKeyset, adLockOptimistic, adCmdTableDirect
600       With rsx1
610         .MoveFirst
620         .Find "[Username] = '" & gstrJournalUser & "'"
630         If .EOF = False Then
640           If IsNull(.Fields("Posting_Date")) = False Then
650             datPostingDate = .Fields("Posting_Date")
660           Else
670             .Fields("Posting_Date") = datPostingDate
680             .Update
690           End If
700         Else
710           .AddNew
720           .Fields("Posting_Date") = datPostingDate
730           .Fields("Username") = gstrJournalUser
740           .Update
750         End If
760         .Close
770       End With
780       cnxn.Close

790       Set dbs = CurrentDb
          ' ** Empty tmpAccount.
800       Set qdf = dbs.QueryDefs("qryMap_Int_08_01")
810       qdf.Execute
820       Set qdf = Nothing
830       DoEvents
          ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
          ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
          ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
          ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
          ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
          ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
          ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
          ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
840       Set qdf = dbs.QueryDefs("qryMap_Int_08_02")
850       qdf.Execute
860       Set qdf = Nothing
870       DoEvents
880       dbs.Close
890       Set dbs = Nothing
900       DoEvents

910       .transdate.DefaultValue = "#" & Format(datPostingDate, "mm/dd/yyyy") & "#"
920       .shareface.ForeColor = CLR_DISABLED_FG
930       .shareface.BackColor = CLR_DISABLED_BG

940       IncludeCurrency  ' ** Procedure: Below.

950     End With

960     DoCmd.Hourglass False

EXITP:
970     Set rsx1 = Nothing
980     Set cnxn = Nothing
990     Set qdf = Nothing
1000    Set dbs = Nothing
1010    Exit Sub

ERRH:
1020    DoCmd.Hourglass False
1030    Select Case ERR.Number
        Case Else
1040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1050    End Select
1060    Resume EXITP

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

        ' ** Plain keys.
1150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyEscape
1170        intRetVal = 0
1180        cmdCancel_Click  ' ** Procedure: Below.
1190      Case vbKeyF7
1200        intRetVal = 0
1210        cmdCalendar_Click  ' ** Procedure: Below.
1220      End Select
1230    End If

        ' ** Ctrl-Shift keys.
1240    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1250      Select Case intRetVal
          Case vbKeyF
1260        With Me
1270          intRetVal = 0
1280          .FocusHolder.SetFocus
1290        End With
1300      End Select
1310    End If

EXITP:
1320    KeyCode = intRetVal
1330    Exit Sub

ERRH:
1340    intRetVal = 0
1350    THAT_PROC = THIS_PROC
1360    That_Erl = Erl
1370    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1380    Resume EXITP

End Sub

Private Sub Form_Timer()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngAssetNo As Long
        Dim Cancel As Integer
        Dim lngTmp03 As Long

        Static lngTmp01 As Long, lngTmp02 As Long

1410    With Me
1420      .TimerInterval = 0
1430      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
1440        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
1450      End If
1460      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
1470        lngGTR_Stat = lngGTR_Stat + 1&
1480        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1490        lngTmp01 = lngTmp01 + 1&
1500        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1510        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1520        If lngTmp01 > 1& Then
1530          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1540        End If
1550        If lngTmp01 = 24& Then
1560          lngTmp01 = 0&
1570          lngTmp02 = lngTmp02 + 1&
1580          If lngTmp02 < 8& Then
1590            .TimerInterval = 50&
1600          Else
1610            blnGTR_Emblem = False
1620          End If
1630        Else
1640          .TimerInterval = 50&
1650        End If
1660      End If  ' ** blnGTR_Emblem.
1670      If blnExitingAssetNo = True Then
            ' ** This was set up this way so that the user doesn't have to hit Cancel twice
            ' ** if they happened to have entered an asset, but then changed their mind.
1680        blnExitingAssetNo = False
1690        If blnCancel = False Then
1700          Set dbs = CurrentDb
1710          Select Case strCallingForm
              Case "frmJournal"
                ' ** ActiveAssets, grouped and summed, by assetno, by specified [astno].
1720            Set qdf = dbs.QueryDefs("qryMap_Int_07")
1730          Case "frmJournal_Columns"
                ' ** ActiveAssets, grouped and summed, by assetno, by specified [astno].
1740            Set qdf = dbs.QueryDefs("qryJournal_Columns_31_03")
1750          End Select
1760          With qdf.Parameters
1770            ![astno] = Me.cmbAssets
1780          End With
1790          Set rst = qdf.OpenRecordset
1800          If rst.BOF = True And rst.EOF = True Then
                ' ** Shouldn't happen.
1810          Else
1820            blnAssetUpdated = False
1830            blnAssetChecked = True
1840            rst.MoveFirst
1850            .TotalShares = rst![sumsf]
1860            .shareface = Round(rst![sumsf], 4)
                ' ** This enabling and disabling is causing all sorts of ...
1870            If .amount.Enabled = True Then
1880              dblPerShare_Calc = (Nz(.amount, 0) / IIf(Nz(.shareface, 0) = 0, 1, .shareface) * 100)
1890              .pershare = dblPerShare_Calc  ' ** Rate %, 5 dec.
1900  On Error Resume Next
1910              .amount.SetFocus
1920  On Error GoTo ERRH
1930            Else
1940              .amount = (.pershare / 100#) * IIf(Nz(.shareface, 0) = 0, 1, .shareface)
1950  On Error Resume Next
1960              .pershare.SetFocus
1970  On Error GoTo ERRH
1980            End If
1990          End If
2000          rst.Close
2010          dbs.Close
2020        End If
2030      End If
2040      If gblnGoToReport = True Then
2050        DoCmd.Hourglass True  ' ** Make sure it's still running.
2060        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48& And lngGTR_Stat < 96&) Or _
                (lngGTR_Stat > 96& And lngGTR_Stat < 120&) Or (lngGTR_Stat > 120&)) Then
              ' ** Just let the emblem go.
2070        Else
2080          DoCmd.Hourglass True  ' ** Make sure it's still running.
2090          If blnGTR_Emblem = False Then
2100            blnGTR_Emblem = True
2110            lngTmp01 = 0&: lngTmp02 = 0&
2120            lngTmp03 = (.GoToReport_Emblem_01_img.Width + (8& * lngTpp))
2130            .GoToReport_arw_ok_img.Left = (.GoToReport_arw_ok_img.Left - lngTmp03)
2140            .cmdOK.Left = (.cmdOK.Left - lngTmp03)
2150            .cmdCancel.Left = (.cmdCancel.Left - lngTmp03)
2160          End If
2170          DoEvents
2180          Select Case blnGoingToReport
              Case True
2190            Select Case blnGoingToReport2
                Case True
2200              Select Case .GoToReport_arw_ok_img.Visible
                  Case True
2210                blnGoingToReport2 = False
2220                .TimerInterval = 50&
2230                cmdOK_Click  ' ** Procedure: Below.
2240              Case False
2250                .amount.SetFocus
2260                .amount = 100
2270                Cancel = 0
2280                amount_Exit Cancel  ' ** Procedure: Below.
2290                DoEvents
2300                .GoToReport_arw_ok_img.Visible = True
2310                .GoToReport_arw_map_amt_img.Visible = False
2320                .GoToReport_arw_map_per_img.Visible = False
2330                .GoToReport_lbl_amtper.Visible = False
2340                .GoToReport_lin_hline01.Visible = False
2350                .GoToReport_lin_hline02.Visible = False
2360                .GoToReport_lbl_arrow01.Visible = False
2370                .GoToReport_lbl_arrow02.Visible = False
2380                DoEvents
2390                .TimerInterval = 50&
2400              End Select
2410            Case False
2420              blnGoingToReport2 = True
2430              lngAssetNo = GetIntAsset  ' ** Module Function: modGoToReportFuncs.
2440              If lngAssetNo > 0& Then
2450                .cmbAssets = lngAssetNo
2460                cmbAssets_AfterUpdate  ' ** Procedure: Below.
2470                DoEvents
2480                .GoToReport_arw_map_amt_img.Visible = True
2490                .GoToReport_arw_map_per_img.Visible = True
2500                .GoToReport_lbl_amtper.Visible = True
2510                .GoToReport_lin_hline01.Visible = True
2520                .GoToReport_lin_hline02.Visible = True
2530                .GoToReport_lbl_arrow01.Visible = True
2540                .GoToReport_lbl_arrow02.Visible = True
2550                .GoToReport_arw_map_asset_img.Visible = False
2560                .opgAssetSource.Visible = True
2570                .opgAssetSource_box.Visible = True
2580                .TotalShares = 4000
2590                .shareface = 4000
2600                .FocusHolder.SetFocus  ' ** So they don't see a flashing cursor.
2610                DoEvents
2620                .TimerInterval = 50&
2630              Else
2640                .TimerInterval = 0&
2650                Select Case garr_varGoToReport(GTR_FRM2)
                    Case "frmJournal"
2660                  Forms(garr_varGoToReport(GTR_FRM2)).GTREmblem_Set False  ' ** Form Procedure: frmJournal.
2670                Case "frmJournal_Columns"
2680                  Forms(garr_varGoToReport(GTR_FRM2)).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
2690                End Select
2700                gblnGoToReport = False
2710                blnGoingToReport = False
2720                blnGoingToReport2 = False
2730                GTREmblem_Off  ' ** Procedure: Below.
2740                .GoToReport_arw_map_asset_img.Visible = False
2750                .opgAssetSource.Visible = True
2760                .opgAssetSource_box.Visible = True
2770                Beep
2780                DoCmd.Hourglass False
2790                MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                      "There are insufficient asset holdings to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
2800              End If
2810            End Select
2820          Case False
2830            blnGoingToReport = True
2840            .opgAssetSource.Visible = False
2850            .opgAssetSource_box.Visible = False
2860            .GoToReport_arw_map_asset_img.Visible = True
2870            .FocusHolder.SetFocus  ' ** So they don't see a flashing cursor.
2880            DoEvents
2890            Select Case garr_varGoToReport(GTR_FRM2)
                Case "frmJournal"
2900              Forms(garr_varGoToReport(GTR_FRM2)).GTREmblem_Set False  ' ** Form Procedure: frmJournal.
2910            Case "frmJournal_Columns"
2920              Forms(garr_varGoToReport(GTR_FRM2)).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
2930            End Select
2940            .TimerInterval = 50&
2950          End Select
2960        End If
2970      End If
2980    End With  ' ** Me.

EXITP:
2990    Set rst = Nothing
3000    Set qdf = Nothing
3010    Set dbs = Nothing
3020    Exit Sub

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim strControlName As String

3110    If DataErr <> 0 Then
3120      Select Case DataErr
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
3130      Case 2046  ' ** The command or action isn't available now (first or last record).
3140        DoCmd.CancelEvent
3150      Case 2113  ' ** The value you entered isn't valid for this field.
3160        Select Case Me.ActiveControl.Name
            Case "shareface"
3170          strControlName = "Shares"
3180        Case "pershare"
3190          strControlName = "Rate"
3200        Case "amount"
3210          strControlName = "Amount"
3220        Case "transdate"
3230          strControlName = "Date"
3240          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3250          Response = acDataErrContinue
3260          DoCmd.CancelEvent
3270          Exit Sub
3280        Case Else
3290          strControlName = ""
3300        End Select
3310        MsgBox "You must enter a numeric value in the " & strControlName & " field.", vbInformation + vbOKOnly, "Invalid Entry"
3320        Response = acDataErrContinue
3330        DoCmd.CancelEvent
3340      Case 2237  ' ** The text you entered isn't an item in the list.
3350        DoCmd.CancelEvent
3360      Case 2279  ' ** The value you entered isn't appropriate for the input mask '|' specified for this field.
3370        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3380        Response = acDataErrContinue
3390        DoCmd.CancelEvent
3400      Case 3162 ' ** You tried to assign the Null value to a variable that is not a Variant data type.
3410        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
            ' ** Do nothing.
3420      Case Else
3430        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3440        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450      End Select
3460    End If

EXITP:
3470    Exit Sub

ERRH:
3480    Select Case ERR.Number
        Case Else
3490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3500    End Select
3510    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3610    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3620    If Not clsMonthClass Is Nothing Then
3630      If clsMonthClass.IsCalendar Then
3640        Cancel = -1
3650      Else
3660        Set clsMonthClass = Nothing
3670      End If
3680    End If

EXITP:
3690    Exit Sub

ERRH:
3700    DoCmd.Hourglass False
3710    Select Case ERR.Number
        Case Else
3720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3730    End Select
3740    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3810    With Me
3820      If .cmdCalendar_raised_focus_dots_img.Visible = True Or .cmdCalendar_raised_focus_img.Visible = True Then
3830        Select Case blnCalendar1_Focus
            Case True
3840          .cmdCalendar_raised_semifocus_dots_img.Visible = True
3850          .cmdCalendar_raised_img.Visible = False
3860        Case False
3870          .cmdCalendar_raised_img.Visible = True
3880          .cmdCalendar_raised_semifocus_dots_img.Visible = False
3890        End Select
3900        .cmdCalendar_raised_focus_dots_img.Visible = False
3910        .cmdCalendar_raised_focus_img.Visible = False
3920        .cmdCalendar_sunken_focus_dots_img.Visible = False
3930        .cmdCalendar_raised_img_dis.Visible = False
3940      End If
3950    End With

EXITP:
3960    Exit Sub

ERRH:
3970    Select Case ERR.Number
        Case Else
3980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3990    End Select
4000    Resume EXITP

End Sub

Private Sub cmdOK_Click()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        'Dim cnxn As ADODB.Connection, rsx1 As ADODB.Recordset  ' ** Early binding.
        Dim cnxn As Object, rsx1 As Object                      ' ** Late binding.
        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim dblShareface As Double, dblAmount As Double, dblPerShare As Double
        Dim dblSumICash As Double, dblSumShareface As Double, dblOverUnder As Double
        Dim datTransDate As Date, datAssetDate As Date, strMapType As String
        Dim lngAssetNo As Long, intTaxCode As Integer
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
        'Const A_ICASH As Integer = 3
        Const A_ICADJ As Integer = 4

4110    blnContinue = True

4120    dblShareface = 0#
4130    dblAmount = 0#
4140    dblPerShare = 0#
4150    dblSumICash = 0#
4160    dblSumShareface = 0#
4170    dblOverUnder = 0#
4180    intTaxCode = 0
4190    datTransDate = 0
4200    datAssetDate = 0
4210    lngAssetNo = 0&

4220    With Me

4230      If IsNull(.cmbAssets) = True Then
4240        blnContinue = False
4250        MsgBox "Please select a valid asset.", vbInformation + vbOKOnly, "Entry Required"
4260        .cmbAssets.SetFocus
4270      Else
4280        If Trim(.cmbAssets) = vbNullString Then
4290          blnContinue = False
4300          MsgBox "Please select a valid asset.", vbInformation + vbOKOnly, "Entry Required"
4310          .cmbAssets.SetFocus
4320        Else
4330          lngAssetNo = .cmbAssets
4340          If Nz(.shareface, 0) = 0 Then
4350            blnContinue = False
4360            MsgBox "Face must be greater than zero to continue.", vbInformation + vbOKOnly, "Invalid Entry"
4370            .cmbAssets.SetFocus
4380          Else
4390            If IsNull(.amount) Then
4400              blnContinue = False
4410              MsgBox "Amount must be greater than zero to continue.", vbInformation + vbOKOnly, "Entry Required"
4420              .amount.SetFocus
4430            Else
4440              If IsNull(.pershare) Then
4450                blnContinue = False
4460                MsgBox "Rate must be greater than zero to continue.", vbInformation + vbOKOnly, "Entry Required"
4470                .pershare.SetFocus
4480              Else
4490                If .amount = 0 And .pershare = 0 Then
4500                  blnContinue = False
4510                  MsgBox "Either Amount or Rate must be greater than zero to continue.", vbInformation + vbOKOnly, "Entry Required"
4520                  If .amount.Enabled = False Or .pershare.Enabled = False Then
4530                    .amount = Null
4540                    .amount.Enabled = True
4550                    .amount.Locked = False
4560                    .amount.ForeColor = CLR_BLK
4570                    .pershare = Null
4580                    .pershare.Enabled = True
4590                    .pershare.Locked = False
4600                    .pershare.ForeColor = CLR_BLK
4610                    If lngCurrID = 150& Then
4620                      .amount.BackColor = CLR_WHT
4630                      .pershare.BackColor = CLR_WHT
4640                    Else
4650                      .amount.BackColor = lngBkClr
4660                      .pershare.BackColor = lngBkClr
4670                    End If
4680                  End If
4690                  .amount.SetFocus
4700                End If
4710              End If
4720            End If
4730          End If
4740        End If
4750      End If

4760      If blnContinue = True Then

4770        DoCmd.Hourglass True
4780        DoEvents

4790        Set dbs = CurrentDb

            ' ** Empty Journal Map.
4800        Set qdf = dbs.QueryDefs("qryMap_01")
4810        qdf.Execute

4820        Select Case strCallingForm
            Case "frmJournal"
              ' ** MasterAsset, linked to AssetType, by specified [astno].
4830          Set qdf = dbs.QueryDefs("qryMap_Int_06")
4840        Case "frmJournal_Columns"
              ' ** MasterAsset, linked to AssetType, by specified [astno].
4850          Set qdf = dbs.QueryDefs("qryJournal_Columns_31_01")
4860        End Select
4870        With qdf.Parameters
4880          ![astno] = lngAssetNo
4890        End With
4900        Set rst = qdf.OpenRecordset
4910        If rst.BOF = True And rst.EOF = True Then
              ' ** Shouldn't happen.
4920        Else
4930          rst.MoveFirst
4940          intTaxCode = rst![taxcode]
4950        End If
4960        .taxcode = intTaxCode
4970        DoEvents

4980        dblShareface = .TotalShares
4990        Select Case intInputType
            Case 1  ' ** User entered an Amount.
5000          dblAmount = .amount
5010          dblPerShare = ((dblAmount / dblShareface) * 100#)
5020        Case 2  ' ** User entered a PerShare.
5030          dblPerShare = .pershare
5040          dblAmount = (dblShareface * (dblPerShare / 100#))
5050        Case Else  ' ** Shouldn't be an else!
5060          dblAmount = .amount
5070          dblPerShare = ((dblAmount / dblShareface) * 100#)
5080        End Select
            ' ** dblPerShare is visible number, in percent, so real number is 1/100th of that.
5090        datTransDate = .transdate
5100        strMapType = .maptype  ' ** JournalType.
5110        DoEvents

            ' ** Check for closed accounts holding assets.
            ' ** qryMap_07 (qryMap_06 (ActiveAssets, linked to Account, just closed accounts, by specified [astno]),
            ' ** grouped by accountno, with cnt_astno), grouped and summed, with cnt_actno, cnt_astno.
5120        Set qdf = dbs.QueryDefs("qryMap_08")
5130        With qdf.Parameters
5140          ![astno] = Me.cmbAssets
5150        End With
5160        Set rst = qdf.OpenRecordset
5170        If rst.BOF = True And rst.EOF = True Then
              ' ** No closed accounts holding assets.
5180          msgResponse = vbYes
5190          rst.Close
5200        Else
5210          rst.MoveFirst
5220          If IsNull(rst![cnt_actno]) = True And IsNull(rst![cnt_astno]) = True Then
5230            msgResponse = vbYes
5240            rst.Close
5250          Else
5260            If rst![cnt_actno] = 0 And IsNull(rst![cnt_astno]) = True Then
5270              msgResponse = vbYes
5280              rst.Close
5290            ElseIf rst![cnt_actno] = 0 And rst![cnt_astno] = 0 Then
5300              msgResponse = vbYes
5310              rst.Close
5320            Else
5330              lngTmp02 = rst![cnt_actno]
5340              lngTmp03 = rst![cnt_astno]
5350              rst.Close
                  ' ** qryMap_06 (ActiveAssets, linked to Account, just closed accounts,
                  ' ** by specified [astno]), grouped by accountno, with cnt_astno.
5360              Set qdf = dbs.QueryDefs("qryMap_07")
5370              With qdf.Parameters
5380                ![astno] = Me.cmbAssets
5390              End With
5400              Set rst = qdf.OpenRecordset
5410              rst.MoveFirst
5420              If lngTmp02 = 1& Then
5430                strTmp01 = "There is 1 closed account currently holding this asset:" & vbCrLf & _
                      "    " & rst![accountno] & "  " & rst![shortname] & vbCrLf
5440              Else
5450                strTmp01 = "There are " & CStr(lngTmp02) & " closed accounts currently holding this asset:" & vbCrLf
5460                For lngX = 1& To lngTmp02
5470                  strTmp01 = strTmp01 & "    " & rst![accountno] & "  " & rst![shortname] & vbCrLf
5480                  If lngX < lngTmp02 Then rst.MoveNext
5490                Next
5500              End If
5510              strTmp01 = strTmp01 & "If you choose to continue with this mapping," & vbCrLf & _
                    "the results may be inaccurate." & vbCrLf & vbCrLf & "Do you wish to continue?"
5520              DoCmd.Hourglass False
5530              Beep
5540              msgResponse = MsgBox(strTmp01, vbQuestion + vbYesNo + vbDefaultButton2, "Closed Accounts Holding Assets")
5550            End If
5560          End If
5570        End If
5580        DoEvents

5590        If msgResponse = vbYes Then

5600          DoCmd.Hourglass True
5610          DoEvents

5620          Select Case strCallingForm
              Case "frmJournal"
                ' ** Append ActiveAssets to Journal Map, by specified [astno], [jrntyp], [persh], [transdat].  #curr_id
5630            Set qdf = dbs.QueryDefs("qryMap_Int_01")
5640          Case "frmJournal_Columns"
                ' ** Append ActiveAssets to Journal Map, by specified [astno], [jrntyp], [persh], [transdat].  #curr_id
5650            Set qdf = dbs.QueryDefs("qryJournal_Columns_31_02")
5660          End Select
5670          With qdf.Parameters
5680            ![astno] = lngAssetNo
5690            ![jrntyp] = strMapType
5700            ![persh] = dblPerShare  ' ** Used for both icash and pershare.
                ' ** Remember, dblPerShare is visible number, in percent.
                ' ** Query divides by 100.
5710            ![transdat] = datTransDate
5720          End With
5730          qdf.Execute  ' ** Calculates icash as (([persh]/100)*[ActiveAssets].[shareface])
5740          dbs.Close
5750          Set qdf = Nothing
5760          Set dbs = Nothing
5770          DoEvents

5780          blnSkip = True
              ' ################################################################################
5790          If blnSkip = False Then
                'Set rsx1 = New ADODB.Recordset             ' ** Early binding.
5800            Set rsx1 = CreateObject("ADODB.Recordset")  ' ** Late binding.
5810            Set cnxn = CurrentProject.Connection
5820            rsx1.Open "journal map", cnxn, adOpenKeyset, adLockOptimistic, adCmdTableDirect
5830            With rsx1
5840              .MoveLast
5850              lngRecs = .RecordCount
5860              .MoveFirst
5870              For lngX = 1& To lngRecs
5880                dblTmp04 = Nz(.Fields("icash"), 0)
5890                If ((dblTmp04 > 0.009) Or (dblTmp04 < -0.009)) Then
5900                  dblTmp04 = dbl_Round(dblTmp04, 3)  ' ** Module Function: modUtilities.
5910                Else
5920                  dblTmp04 = 0#
5930                End If
5940                If dblTmp04 <> 0# Then
5950                  dblTmp05 = Nz(.Fields("shareface"), 0)
5960                  If ((dblTmp05 > 0.00009) Or (dblTmp05 < -0.00009)) Then
5970                    dblTmp05 = dbl_Round(dblTmp05, 5)  ' ** Module Function: modUtilities.
5980                  Else
5990                    dblTmp05 = 0#
6000                    dblTmp04 = 0#  ' ** Also zero-out the icash.
6010                  End If
6020                End If
6030                .Fields("icash") = dblTmp04
6040                .Fields("shareface") = dblTmp05
6050                .Update
6060                dblSumICash = dblSumICash + dblTmp04
6070                dblSumShareface = dblSumShareface + dblTmp05
6080                dblTmp04 = 0#: dblTmp05 = 0#
6090                If lngX < lngRecs Then .MoveNext
6100              Next
6110            End With
6120          End If  ' ** blnSkip.
              ' ################################################################################

6130          DoEvents
6140          Set dbs = CurrentDb

6150          Select Case intInputType
              Case 1  ' ** User entered an Amount.
                ' ** qryMap_Int_13a (Journal Map, with qryMap_Int_07 (ActiveAssets,
                ' ** grouped and summed, by assetno, by specified [astno]), with calc'd
                ' ** fields, by specified [astno], [amt], [prec]), grouped and summed by account.
6160            Set qdf = dbs.QueryDefs("qryMap_Int_14a")
6170            With qdf.Parameters
6180              ![astno] = lngAssetNo
6190              ![amt] = dblAmount
6200              ![prec] = gintShareFaceDecimals
6210            End With
6220          Case 2  ' ** User entered a PerShare.
                ' ** qryMap_Int_13b (Journal Map, with qryMap_Int_07 (ActiveAssets,
                ' ** grouped and summed, by assetno, by specified [astno]), with calc'd
                ' ** fields, by specified [astno], [persh], [prec]), grouped and summed by account.
6230            Set qdf = dbs.QueryDefs("qryMap_Int_14b")
6240            With qdf.Parameters
6250              ![astno] = lngAssetNo
6260              ![persh] = dblPerShare
                  ' ** Remember, dblPerShare is visible number, in percent.
                  ' ** Query divides by 100.
6270              ![prec] = gintShareFaceDecimals
6280            End With
6290          End Select
6300          Set rst = qdf.OpenRecordset
6310          With rst
6320            If .BOF = True And .EOF = True Then
                  ' ** Something's seriously wrong!
6330              blnContinue = False
6340            Else
6350              .MoveLast
6360              lngMaps = .RecordCount
6370              .MoveFirst
6380              arr_varMap = .GetRows(lngMaps)
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
                  ' **     6       5     icash            M_ICASH        ' ** Total icash by account: ([shareface]*([amt]/[sumsf])).
                  ' **     7       6     icash_rnd        M_ICASHR       ' ** Total icash by account, rounded to 2 decimal places.
                  ' **
                  ' ***************************************************
6390            End If  ' ** THIS PERSHARE IS REAL!!!
6400            .Close
6410          End With
6420          Set rst = Nothing
6430          Set qdf = Nothing
6440          DoEvents

6450          Set rst = dbs.OpenRecordset("journal map", dbOpenDynaset, dbConsistent)
6460          With rst
6470            If .BOF = True And .EOF = True Then
                  ' ** Might be coming from GoToReport.
6480              blnContinue = False
6490            Else
6500              .MoveLast
6510              lngRecs = .RecordCount
6520              .MoveFirst
6530              For lngX = 1& To lngRecs
6540                If InStr(![accountno], "INCOME O/U") = 0 Then  ' ** This is excluded from the arr_varMap() query.
6550                  blnFound = False
6560                  For lngY = 0& To (lngMaps - 1&)
6570                    If arr_varMap(M_ACTNO, lngY) = ![accountno] And arr_varMap(M_ASTNO, lngY) = ![assetno] Then
6580                      blnFound = True
6590                      .Edit
6600                      dblTmp04 = Round(![shareface], gintShareFaceDecimals)
6610                      ![shareface] = dblTmp04
6620                      ![ICash] = CCur(Round((dblTmp04 * arr_varMap(M_PER, lngY)), 2))  ' ** THIS PERSHARE IS REAL!!!
6630                      ![pershare] = arr_varMap(M_PER, lngY)
6640                      .Update
6650                      Exit For
6660                    End If
6670                  Next
6680                  If blnFound = False Then
6690                    blnContinue = False
6700                    Exit For
6710                  End If
6720                End If
6730                If lngX < lngRecs Then .MoveNext
6740              Next
6750            End If  ' ** BOF, EOF.
6760            .Close
6770          End With
6780          Set rst = Nothing
6790          DoEvents

6800          If blnContinue = True Then

                ' ** Delete qryMap_Int_11 (Journal Map, for shareface < 0.00006, icash < 0.006, by specified [astno]).
6810            Set qdf = dbs.QueryDefs("qryMap_Int_12")
6820            With qdf.Parameters
6830              ![astno] = lngAssetNo
6840            End With
6850            qdf.Execute
6860            Set qdf = Nothing
6870            DoEvents

                ' ** qryMap_Int_02_04 (Journal Map, linked to Account, qryMap_Int_02_02 (qryMap_Int_02_01
                ' ** (Journal Map, grouped by accountno, assetno, Location_ID), grouped by accountno,
                ' ** assetno, with cnt, Location_ID_min, Location_ID_max), grouped and summed,
                ' ** by accountno), grouped and summed, by assetno, by specified [astno].
6880            Set qdf = dbs.QueryDefs("qryMap_Int_02_05")
6890            With qdf.Parameters
6900              ![astno] = lngAssetNo
6910            End With
6920            Set rst = qdf.OpenRecordset
6930            With rst
6940              .MoveFirst
6950              dblSumICash = ![ICash]
6960              dblSumShareface = ![shareface]
6970              dblPerShare = (![pershare] * 100#)  ' ** Max([pershare]), though they should all be identical.
                  ' ** ![pershare] should be real number, so
                  ' ** dblPerShare is visible number, in perecent.
6980              .Close
6990            End With
7000            Set rst = Nothing
7010            Set qdf = Nothing
7020            DoEvents

7030            dblTmp04 = (Round(dblAmount, 2) - dblSumICash)
7040            If ((dblTmp04 > 0.009) Or (dblTmp04 < -0.009)) And (dblTmp04 > 1#) Then
                  ' ** There's an Over/Under of more than $1.00, so let's see if we can bring it down.

7050              Select Case intInputType
                  Case 1  ' ** User entered an Amount.
                    ' ** qryMap_Int_18a (Journal Map, linked to qryMap_Int_17a (Journal Map, linked to
                    ' ** qryMap_Int_16a (qryMap_Int_14a (qryMap_Int_13a (Journal Map, with qryMap_Int_07
                    ' ** (ActiveAssets, grouped and summed, by assetno, by specified [astno]), with calc'd
                    ' ** fields, by specified [astno], [amt], [prec]), grouped and summed by account),
                    ' ** just needed fields, by specified [thrsh]), with icash_adj, by specified [astno],
                    ' ** [thrsh], [icashou]; Cartesian), with icash_adjx), just entries with adjustment.
7060                Set qdf = dbs.QueryDefs("qryMap_Int_19a")
7070                With qdf.Parameters
7080                  ![astno] = lngAssetNo
7090                  ![amt] = dblAmount
7100                  ![prec] = gintShareFaceDecimals
7110                  dblTmp05 = (0.01 / (dblTmp04 / dblShareface))  ' ** Threshhold needed for a penny adjustment.
7120                  ![thrsh] = dblTmp05
7130                  ![ICashou] = dblTmp04
7140                End With
7150              Case 2  ' ** User entered a PerShare.
                    ' ** qryMap_Int_18b (Journal Map, linked to qryMap_Int_17b (Journal Map, linked to
                    ' ** qryMap_Int_16b (qryMap_Int_14b (qryMap_Int_13b (Journal Map, with qryMap_Int_07
                    ' ** (ActiveAssets, grouped and summed, by assetno, by specified [astno]), with calc'd
                    ' ** fields, by specified [astno], [persh], [prec]), grouped and summed by account),
                    ' ** just needed fields, by specified [thrsh]), with icash_adj, by specified [astno],
                    ' ** [thrsh], [icashou]; Cartesian), with icash_adjx), just entries with adjustment.
7160                Set qdf = dbs.QueryDefs("qryMap_Int_19b")
7170                With qdf.Parameters
7180                  ![astno] = lngAssetNo
7190                  ![persh] = dblPerShare
                      ' ** Remember, dblPerShare is visible number, in percent.
                      ' ** Query divides by 100.
7200                  ![prec] = gintShareFaceDecimals
7210                  dblTmp05 = (0.01 / (dblTmp04 / dblShareface))  ' ** Threshhold needed for a penny adjustment.
7220                  ![thrsh] = dblTmp05
7230                  ![ICashou] = dblTmp04
7240                End With
7250              End Select

7260              lngAdjs = 0&
7270              Set rst = qdf.OpenRecordset
7280              With rst
7290                If .BOF = True And .EOF = True Then
                      ' ** Not enough to adjust.
7300                Else
7310                  .MoveLast
7320                  lngAdjs = .RecordCount
7330                  .MoveFirst
7340                  arr_varAdj = .GetRows(lngAdjs)
                      ' **********************************************
                      ' ** Array: arr_varAdj()
                      ' **
                      ' **   Field  Element  Name         Constant
                      ' **   =====  =======  ===========  ==========
                      ' **     1       0     ID           A_ID
                      ' **     2       1     accountno    A_ACTNO
                      ' **     3       2     assetno      A_ASTNO
                      ' **     4       3     icash        A_ICASH
                      ' **     5       4     icash_adj    A_ICADJ
                      ' **
                      ' **********************************************
7350                End If
7360                .Close
7370              End With
7380              Set rst = Nothing
7390              Set qdf = Nothing
7400              DoEvents  ' ** All these DoEvents are to keep the window from showing a 'Not Responding'!

7410              Set rst = dbs.OpenRecordset("Journal Map", dbOpenDynaset, dbConsistent)
7420              With rst
7430                .MoveFirst
7440                For lngX = 0& To (lngAdjs - 1&)
7450                  .FindFirst "[ID] = " & CStr(arr_varAdj(A_ID, lngX))
7460                  If .NoMatch = False Then
7470                    .Edit
7480                    ![ICash] = arr_varAdj(A_ICADJ, lngX)
7490                    .Update
7500                  Else
7510                    blnContinue = False
7520                    Exit For
7530                  End If
7540                Next
7550              End With
7560              Set rst = Nothing
7570              Set qdf = Nothing
7580              DoEvents

                  ' ** Now check the Over/Under again!
                  ' ** qryMap_Int_02_04 (Journal Map, linked to Account, qryMap_Int_02_02 (qryMap_Int_02_01
                  ' ** (Journal Map, grouped by accountno, assetno, Location_ID), grouped by accountno,
                  ' ** assetno, with cnt, Location_ID_min, Location_ID_max), grouped and summed,
                  ' ** by accountno), grouped and summed, by assetno, by specified [astno].
7590              Set qdf = dbs.QueryDefs("qryMap_Int_02_05")
7600              With qdf.Parameters
7610                ![astno] = lngAssetNo
7620              End With
7630              Set rst = qdf.OpenRecordset
7640              With rst
7650                .MoveFirst
7660                dblSumICash = ![ICash]
7670                dblSumShareface = ![shareface]
7680                dblPerShare = (![pershare] * 100#)  ' ** Max([pershare]), though they should all be identical.
                    ' ** ![pershare] should be the real number, so
                    ' ** dblPerShare is the visible number, in percent.
7690                .Close
7700              End With
7710              Set rst = Nothing
7720              Set qdf = Nothing
7730              DoEvents

7740              dblTmp04 = (Round(dblAmount, 2) - dblSumICash)

7750            End If

7760            If ((dblTmp04 > 0.009) Or (dblTmp04 < -0.009)) Then
                  ' ** Create a new record in Journal Map.
7770              Set rst = dbs.OpenRecordset("journal map", dbOpenDynaset, dbAppendOnly)
7780              With rst
7790                .AddNew
7800                ![assetno] = lngAssetNo
7810                ![accountno] = "INCOME O/U"
7820                ![shareface] = Round(Abs(dblShareface - dblSumShareface), gintShareFaceDecimals)
7830                ![pershare] = (dblPerShare / 100#)
                    ' ** dblPerShare is the visible number, in percent.
7840                ![transdate] = datTransDate
7850                ![journaltype] = strMapType
7860                dblOverUnder = Round((Round(dblAmount, 2) - dblSumICash), 2)
7870                ![ICash] = CCur(Round(dblOverUnder, 2))
7880                ![description] = "Over/Short of " & strMapType & " Mapping for " & Me.cmbAssets.Column(CBX_A_DESC)
7890                ![curr_id] = lngCurrID
7900                .Update
7910                .Close
7920              End With
7930              Set rst = Nothing
7940              .overshort = dblOverUnder
7950            Else
7960              .overshort = 0#
7970            End If

7980            dbs.Close
7990            Set dbs = Nothing
8000            DoEvents

                ' ################################################################################
8010            If blnSkip = False Then
8020              dblTmp04 = (dbl_Round(dblAmount, 2) - dblSumICash)  ' ** Module Function: modUtilities.
8030              If ((dblTmp04 > 0.009) Or (dblTmp04 < -0.009)) Then
                    ' ** Create a new record in the 'journal map' table.
8040                With rsx1
8050                  .AddNew
8060                  .Fields("assetno") = lngAssetNo
8070                  .Fields("accountno") = "INCOME O/U"
8080                  .Fields("shareface") = dblShareface - dblSumShareface
8090                  .Fields("pershare") = dblPerShare / 100
8100                  .Fields("transdate") = datTransDate
8110                  .Fields("journaltype") = "Interest"
8120                  dblOverUnder = Format(dbl_Round(dblAmount, 2) - dblSumICash, "FIXED")  ' ** Module Function: modUtilities.
8130                  .Fields("icash") = dblOverUnder
8140                  .Update
8150                End With
8160                .overshort = dblOverUnder
8170              Else
8180                .overshort = 0#
8190              End If

8200              rsx1.Close
8210              cnxn.Close

8220              Set dbs = CurrentDb
                  ' ** Journal Map, just zero shareface, icash, pcash.
8230              Set qdf = dbs.QueryDefs("qryMap_Div_01a")
8240              Set rst = qdf.OpenRecordset
8250              If rst.BOF = True And rst.EOF = True Then
                    ' ** All's well.
8260                rst.Close
8270              Else
8280                rst.Close
                    ' ** Delete qryMap_Div_01a (Journal Map, just zero shareface, icash, pcash).
8290                Set qdf = dbs.QueryDefs("qryMap_Div_01b")
8300                qdf.Execute
8310              End If
8320              dbs.Close
8330            End If  ' ** blnSkip.
                ' ################################################################################

                ' ** Borrowing these variables from the Court Reports.
8340            gstrCrtRpt_Ordinal = .cmbAssets.Column(CBX_A_CUSIP)  ' ** cusip.
8350            gstrCrtRpt_Version = .cmbAssets.Column(CBX_A_TDESC)  ' ** totdesc.
8360            gdatStartDate = datTransDate

8370            DoEvents

8380            dblAmount = .amount
8390            dblShareface = .shareface

8400            strDocName = "frmMap_Int_Detail"
8410            DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & CStr(dblOverUnder) & "~" & _
                  CStr(dblPerShare) & "~" & CStr(intTaxCode) & "~" & CStr(lngAssetNo) & "~" & _
                  CStr(dblShareface) & "~" & CStr(dblAmount) & "~" & CStr(lngCurrID)

8420            If gblnGoToReport = True Then
8430              DoCmd.Hourglass True  ' ** Make sure it's still running.
8440              DoEvents
8450              .GoToReport_arw_ok_img.Visible = False
8460              Select Case IsNull(garr_varGoToReport(GTR_FRM4))
                  Case True
                    ' ** Finished at frmMap_Int_Detail.
8470                Forms(strDocName).cmdCancel.SetFocus
8480                Forms(strDocName).cmdOK.Enabled = False  ' ** Only allow Cancel on this fake dividend.
8490                Forms(strDocName).cmdReinvest.Enabled = False
8500                DoEvents
8510                Forms(strDocName).TimerInterval = 50&
8520                blnGoingToReport = False
8530              Case False
                    ' ** Continuing on to frmMap_Reinvest_DivInt_Price.
8540                Forms(strDocName).cmdReinvest.SetFocus
8550                Forms(strDocName).cmdOK.Enabled = False  ' ** Only allow Reinvest on this fake dividend.
8560                Forms(strDocName).cmdCancel.Enabled = False
8570                DoEvents
8580                Forms(strDocName).TimerInterval = 100&
8590                blnGoingToReport = False
8600              End Select
8610            End If

8620            DoCmd.Close acForm, THIS_NAME

8630          Else
8640            Select Case gblnGoToReport
                Case True
8650              Beep
8660              DoCmd.Hourglass False
8670              MsgBox "The report site cannot be shown because there are no records.", _
                    vbInformation + vbOKOnly, "Record Not Found"
8680              blnGoingToReport = False
8690              .GoToReport_arw_ok_img.Visible = False
8700              gblnGoToReport = False
8710            Case False
8720              Beep
8730              DoCmd.Hourglass False
8740              MsgBox "This procedure cannot continue." & vbCrLf & "Expected records were not found.", _
                    vbInformation + vbOKOnly, "Record Not Found"
8750              blnGoingToReport = False
8760              .GoToReport_arw_ok_img.Visible = False
8770              gblnGoToReport = False
8780            End Select
8790            .cmdCancel.SetFocus
8800          End If  ' ** blnContinue.

8810        Else
8820          .cmdCancel.SetFocus
8830        End If  ' ** msgResponse.

8840      End If

8850    End With

EXITP:
8860    Set rsx1 = Nothing
8870    Set cnxn = Nothing
8880    Set rst = Nothing
8890    Set qdf = Nothing
8900    Set dbs = Nothing
8910    Exit Sub

ERRH:
8920    DoCmd.Hourglass True
8930    THAT_PROC = THIS_PROC
8940    That_Erl = Erl
8950    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8960    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
          Case vbKeyTab
9070        With Me
9080          intRetVal = 0
9090          .cmdCancel.SetFocus
9100        End With
9110      End Select
9120    End If

        ' ** Shift keys.
9130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9140      Select Case intRetVal
          Case vbKeyTab
9150        With Me
9160          intRetVal = 0
9170          .transdate.SetFocus
9180        End With
9190      End Select
9200    End If

        ' ** Ctrl-Shift keys.
9210    If intCtrlDown And (Not intAltDown) And intShiftDown Then
9220      Select Case intRetVal
          Case vbKeyTab
9230        With Me
9240          intRetVal = 0
9250          .cmdCalendar.SetFocus
9260        End With
9270      End Select
9280    End If

EXITP:
9290    KeyCode = intRetVal
9300    Exit Sub

ERRH:
9310    intRetVal = 0
9320    THAT_PROC = THIS_PROC
9330    That_Erl = Erl
9340    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9350    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

9410    DoCmd.Hourglass True
9420    DoEvents

9430    blnCancel = True

9440    Set dbs = CurrentDb
9450    With dbs
9460      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
9470        Set qdf = .QueryDefs("qryMap_01")
9480      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
9490        Set qdf = .QueryDefs("qryJournal_Columns_31_04")
9500      End Select
9510      qdf.Execute
9520      Set qdf = Nothing
9530      DoEvents
          ' ** Empty tmpAccount.
9540      Set qdf = .QueryDefs("qryMap_Int_08_01")
9550      qdf.Execute
9560      Set qdf = Nothing
9570      DoEvents
9580      .Close
9590    End With
9600    Set dbs = Nothing
9610    DoEvents

9620    gblnSetFocus = True
9630    Select Case strCallingForm
        Case "frmJournal"
9640      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
9650        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
9660      Else
9670        gstrReturningForm = THIS_NAME
9680        DoCmd.SelectObject acForm, strCallingForm, False
9690        Forms(strCallingForm).TimerInterval = 100&
9700      End If
9710    Case "frmJournal_Columns"
9720      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
9730        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
9740      Else
9750        gstrReturningForm = THIS_NAME
9760        DoCmd.SelectObject acForm, strCallingForm, False
9770        Forms(strCallingForm).TimerInterval = 250&
9780      End If
9790    End Select

9800    DoCmd.Close acForm, THIS_NAME

EXITP:
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

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
9990          .cmbAssets.SetFocus
10000       End With
10010     End Select
10020   End If

        ' ** Shift keys.
10030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10040     Select Case intRetVal
          Case vbKeyTab
10050       With Me
10060         intRetVal = 0
10070         .cmdOK.SetFocus
10080       End With
10090     End Select
10100   End If

EXITP:
10110   KeyCode = intRetVal
10120   Exit Sub

ERRH:
10130   intRetVal = 0
10140   THAT_PROC = THIS_PROC
10150   That_Erl = Erl
10160   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10170   Resume EXITP

End Sub

Private Sub opgAssetSource_AfterUpdate()

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_AfterUpdate"

10210   With Me
10220     Select Case .opgAssetSource
          Case .opgAssetSource_optType.OptionValue
10230       .opgAssetSource_optType_lbl.FontBold = True
10240       .opgAssetSource_optName_lbl.FontBold = False
10250       .opgAssetSource_optCusip_lbl.FontBold = False
            ' ** qryMap_Int_04 (MasterAsset, linked to AssetType), linked to
            ' ** MasterAsset, for Dividend = True, shareface <> 0; by assettype.
10260       If .cmbAssets.RowSource <> "qryMap_Int_05a" Then
10270         .cmbAssets.RowSource = "qryMap_Int_05a"
10280       End If
10290       .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
10300       .cmbAssets.ListWidth = 7890
10310     Case .opgAssetSource_optName.OptionValue
10320       .opgAssetSource_optType_lbl.FontBold = False
10330       .opgAssetSource_optName_lbl.FontBold = True
10340       .opgAssetSource_optCusip_lbl.FontBold = False
            ' ** qryMap_Int_04 (MasterAsset, linked to AssetType), linked to
            ' ** MasterAsset, for Dividend = True, shareface <> 0; by totdesc.
10350       If .cmbAssets.RowSource <> "qryMap_Int_05d" Then
10360         .cmbAssets.RowSource = "qryMap_Int_05d"
10370       End If
10380       .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
10390       .cmbAssets.ListWidth = 7890
10400     Case .opgAssetSource_optCusip.OptionValue
10410       .opgAssetSource_optType_lbl.FontBold = False
10420       .opgAssetSource_optName_lbl.FontBold = False
10430       .opgAssetSource_optCusip_lbl.FontBold = True
            ' ** qryMap_Int_04 (MasterAsset, linked to AssetType), linked to
            ' ** MasterAsset, with qryMap_Int_05k (qryMap_Int_05j (MasterAsset,
            ' ** with cusip_len), grouped, with Max(cusip_len)), for
            ' ** Dividend = True, shareface <> 0; by cusip; Cartesian.
10440       If .cmbAssets.RowSource <> "qryMap_Int_05g" Then
10450         .cmbAssets.RowSource = "qryMap_Int_05g"
10460       End If
10470       .cmbAssets.ColumnWidths = "0;6090;1185;0;360;0"
10480       .cmbAssets.ListWidth = 7890
10490     End Select
10500     .cmbAssets.Requery
10510   End With

EXITP:
10520   Exit Sub

ERRH:
10530   Select Case ERR.Number
        Case Else
10540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10550   End Select
10560   Resume EXITP

End Sub

Private Sub opgAssetSource_optType_KeyDown(KeyCode As Integer, Shift As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10620   intShiftDown = (Shift And acShiftMask) > 0
10630   intAltDown = (Shift And acAltMask) > 0
10640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10670       With Me
10680         intRetVal = 0
10690         .cmbAssets.SetFocus
10700       End With
10710     End Select
10720   End If

        ' ** Shift keys.
10730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10750       With Me
10760         intRetVal = 0
10770         .cmdCancel.SetFocus
10780       End With
10790     End Select
10800   End If

EXITP:
10810   KeyCode = intRetVal
10820   Exit Sub

ERRH:
10830   intRetVal = 0
10840   Select Case ERR.Number
        Case Else
10850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10860   End Select
10870   Resume EXITP

End Sub

Private Sub opgAssetSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optName_KeyDown"

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
10990         .cmbAssets.SetFocus
11000       End With
11010     End Select
11020   End If

        ' ** Shift keys.
11030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11050       With Me
11060         intRetVal = 0
11070         .cmdCancel.SetFocus
11080       End With
11090     End Select
11100   End If

EXITP:
11110   KeyCode = intRetVal
11120   Exit Sub

ERRH:
11130   intRetVal = 0
11140   Select Case ERR.Number
        Case Else
11150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11160   End Select
11170   Resume EXITP

End Sub

Private Sub opgAssetSource_optCusip_KeyDown(KeyCode As Integer, Shift As Integer)

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optCusip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11220   intShiftDown = (Shift And acShiftMask) > 0
11230   intAltDown = (Shift And acAltMask) > 0
11240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11270       With Me
11280         intRetVal = 0
11290         .cmbAssets.SetFocus
11300       End With
11310     End Select
11320   End If

        ' ** Shift keys.
11330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11350       With Me
11360         intRetVal = 0
11370         .cmdCancel.SetFocus
11380       End With
11390     End Select
11400   End If

EXITP:
11410   KeyCode = intRetVal
11420   Exit Sub

ERRH:
11430   intRetVal = 0
11440   Select Case ERR.Number
        Case Else
11450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11460   End Select
11470   Resume EXITP

End Sub

Private Sub cmbAssets_AfterUpdate()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_AfterUpdate"

11510   With Me

11520     If blnCurrID = True Then
11530       lngCurrID = 0&
11540       If IsNull(.cmbAssets) = False Then
11550         If IsNull(.cmbAssets.Column(CBX_A_CURRID)) = False Then
11560           lngCurrID = .cmbAssets.Column(CBX_A_CURRID)
11570           .curr_id = lngCurrID
11580         End If
11590       End If
11600     End If

11610     .amount = Null
11620     .amount.Enabled = True
11630     .amount.Locked = False
11640     .amount.ForeColor = CLR_BLK
11650     .amount.BackColor = CLR_WHT
11660     .pershare = Null
11670     .pershare.Enabled = True
11680     .pershare.Locked = False
11690     .pershare.ForeColor = CLR_BLK
11700     .pershare.BackColor = CLR_WHT
11710     blnAssetUpdated = True: blnAssetChecked = False
11720     .amount.SetFocus

11730     curr_id_AfterUpdate  ' ** Procedure: Below.

11740   End With

EXITP:
11750   Exit Sub

ERRH:
11760   THAT_PROC = THIS_PROC
11770   That_Erl = Erl
11780   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11790   Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

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
              ' ** This enabling and disabling is causing all sorts of ...
11890         If .amount.Enabled = True Then
11900 On Error Resume Next
11910           .amount.SetFocus
11920 On Error GoTo ERRH
11930         Else
11940 On Error Resume Next
11950           .pershare.SetFocus
11960 On Error GoTo ERRH
11970         End If
11980       End With
11990     End Select
12000   End If

        ' ** Shift keys.
12010   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12020     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12030       With Me
12040         intRetVal = 0
12050         .cmdCancel.SetFocus
12060       End With
12070     End Select
12080   End If

        ' ** Ctrl keys.
12090   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
12100     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12110       With Me
12120         intRetVal = 0
12130         .opgAssetSource.SetFocus
12140       End With
12150     End Select
12160   End If

EXITP:
12170   KeyCode = intRetVal
12180   Exit Sub

ERRH:
12190   intRetVal = 0
12200   THAT_PROC = THIS_PROC
12210   That_Erl = Erl
12220   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12230   Resume EXITP

End Sub

Private Sub cmbAssets_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_MouseMove"

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

Private Sub cmbAssets_LostFocus()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_LostFocus"

12410   With Me
12420     If IsNull(.cmbAssets) = False Then
12430       If .cmbAssets > 0 And blnAssetUpdated = True And blnAssetChecked = False Then
12440         blnExitingAssetNo = True
12450         .TimerInterval = 100&
12460       End If
12470     End If
12480   End With  ' ** Me.

EXITP:
12490   Exit Sub

ERRH:
12500   THAT_PROC = THIS_PROC
12510   That_Erl = Erl
12520   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12530   Resume EXITP

End Sub

Private Sub curr_id_cmd_DblClick(Cancel As Integer)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_cmd_DblClick"

12610   With Me
12620     If IsNull(.cmbAssets) = False Then
12630       Select Case .curr_date.Visible
            Case True
12640         .curr_date.Visible = False
12650       Case False
12660         .curr_date = .curr_id.Column(CBX_C_DATE)
12670         .curr_date.Visible = True
12680       End Select
12690     End If
12700   End With

EXITP:
12710   Exit Sub

ERRH:
12720   Select Case ERR.Number
        Case Else
12730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
12740   End Select
12750   Resume EXITP

End Sub

Private Sub curr_id_AfterUpdate()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_AfterUpdate"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String
        Dim lngX As Long

12810   With Me

12820     If IsNull(.curr_id) = False Then

12830       lngCurrID = .curr_id.Column(CBX_C_CURRID)

12840       blnFound = False
12850       If lngCurrID <> 150& Then
12860         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
12870           NoChar_Load  ' ** Procedure: Below.
12880         End If
12890         For lngX = 0& To (lngNoChars - 1&)
12900           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
12910             blnFound = True
12920           End If
12930         Next
12940       End If

12950       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
12960         .amount.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
12970         .amount.FontSize = 10
12980         .amount.FontBold = True
12990         .pershare.FontName = "Courier New"
13000         .pershare.FontSize = 10
13010         .pershare.FontBold = True
13020       Case False
13030         If .amount.FontName <> "Arial" Then
13040           .amount.FontName = "Arial"
13050           .amount.FontSize = 10
13060           .amount.FontBold = False
13070           .pershare.FontName = "Arial"
13080           .pershare.FontSize = 10
13090           .pershare.FontBold = False
13100         End If
13110       End Select

13120       If lngCurrID = 150& Then  ' ** USD.
13130         strFrmt1 = "Currency"
13140         intDec1 = 2
13150         strFrmt2 = "Fixed"
13160         intDec2 = 5
13170         lngBkClr = CLR_WHT
13180       Else
13190         strTmp01 = .curr_id.Column(CBX_C_SYM)  ' ** Symbol.
13200         If .curr_id.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
13210           intDec1 = 4
13220         Else
13230           intDec1 = .curr_id.Column(CBX_C_DEC)  ' ** Decimal places.
13240         End If
13250         intDec2 = 5
13260         intLen = Len(strTmp01)
13270         Select Case intLen
              Case 1
                'strTmp01 = strTmp01
13280         Case 2
13290           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2)
13300         Case 3
13310           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3)
13320         Case 4
13330           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & "\" & Mid(strTmp01, 4)
13340         Case 5
13350           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & _
                  "\" & Mid(strTmp01, 4, 1) & "\" & Mid(strTmp01, 5)
13360         End Select
13370         Select Case intDec1
              Case 0
13380           strFrmt1 = strTmp01 & "#,##0;(" & strTmp01 & "#,##0);" & strTmp01 & "0;" & strTmp01 & "0"
13390           strFrmt2 = "#,##0.00000;(" & "#,##0.00000);" & "0.00000;" & "0.00000"
13400         Case 1  ' ** None have this currently.
13410           strFrmt1 = strTmp01 & "#,##0.0;(" & strTmp01 & "#,##0.0);" & strTmp01 & "0.0;" & strTmp01 & "0.0"
13420           strFrmt2 = "#,##0.00000;(" & "#,##0.00000);" & "0.00000;" & "0.00000"
13430         Case 2
13440           strFrmt1 = strTmp01 & "#,##0.00;(" & strTmp01 & "#,##0.00);" & strTmp01 & "0.00;" & strTmp01 & "0.00"
13450           strFrmt2 = "#,##0.00000;(" & "#,##0.00000);" & "0.00000;" & "0.00000"
13460         Case 3
13470           strFrmt1 = strTmp01 & "#,##0.000;(" & strTmp01 & "#,##0.000);" & strTmp01 & "0.000;" & strTmp01 & "0.000"
13480           strFrmt2 = "#,##0.00000;(" & "#,##0.00000);" & "0.00000;" & "0.00000"
13490         End Select
13500         lngBkClr = CLR_VLTGRN
13510       End If
13520       .amount.Format = strFrmt1
13530       .amount.DecimalPlaces = intDec1
13540       .amount.BackColor = lngBkClr
13550       .pershare.Format = strFrmt2
13560       .pershare.RightMargin = 0&
13570       .pershare.DecimalPlaces = intDec2
13580       .pershare.BackColor = lngBkClr

13590       If lngCurrID = 150& Then
13600         .Amount_usd.Visible = False
13610         .pershare_usd.Visible = False
13620       Else
13630         .Amount_usd = 0
13640         .Amount_usd.Visible = True
13650         .pershare_usd = 0
13660         .pershare_usd.Visible = False  ' ** Not currency!
13670       End If

13680     End If

13690   End With

EXITP:
13700   Exit Sub

ERRH:
13710   THAT_PROC = THIS_PROC
13720   That_Erl = Erl
13730   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13740   Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

13810   With Me
13820     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
13830     .curr_id.SetFocus
13840     Response = acDataErrContinue
13850   End With

EXITP:
13860   Exit Sub

ERRH:
13870   THAT_PROC = THIS_PROC
13880   That_Erl = Erl
13890   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13900   Resume EXITP

End Sub

Private Sub Amount_KeyDown(KeyCode As Integer, Shift As Integer)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "Amount_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14020   intShiftDown = (Shift And acShiftMask) > 0
14030   intAltDown = (Shift And acAltMask) > 0
14040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14070       With Me
14080         intRetVal = 0
14090         If .pershare.Enabled = True Then
14100 On Error Resume Next
14110           .pershare.SetFocus
14120           If ERR.Number <> 0 Then
14130 On Error GoTo ERRH
14140             .transdate.SetFocus
14150           Else
14160 On Error GoTo ERRH
14170           End If
14180         Else
14190           .transdate.SetFocus
14200         End If
14210       End With
14220     End Select
14230   End If

        ' ** Shift keys.
14240   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14250     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14260       With Me
14270         intRetVal = 0
14280         blnDrivingMeBackwards = True
14290         .cmbAssets.SetFocus
14300       End With
14310     End Select
14320   End If

EXITP:
14330   KeyCode = intRetVal
14340   Exit Sub

ERRH:
14350   intRetVal = 0
14360   THAT_PROC = THIS_PROC
14370   That_Erl = Erl
14380   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14390   Resume EXITP

End Sub

Private Sub amount_Exit(Cancel As Integer)

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "amount_Exit"

        Dim dblRate2 As Double
        Dim intPos01 As Integer
        Dim blnContinue As Boolean
        Dim dblTmp01 As Double

14410   With Me

14420     blnContinue = True

14430     lngCurrID = .curr_id.Column(CBX_C_CURRID)

14440     If IsNumeric(.amount) = True Then
14450       If .amount < 0 Then
14460         blnContinue = False
14470         MsgBox "Amount must be positive.", vbInformation + vbOKOnly, "Invalid Entry"
14480         DoCmd.CancelEvent
14490         .amount.SetFocus
14500       Else
14510         If CDbl(Rem_Dollar(.amount.text, lngCurrID)) > 999999999.99 Then  ' ** Module Function: modStringFuncs.
14520           blnContinue = False
14530           MsgBox "Amount is too large.", vbInformation + vbOKOnly, "Invalid Entry"
14540           .amount = vbNullString
14550           DoCmd.CancelEvent
14560           .amount.SetFocus
14570         Else
14580           intPos01 = InStr(1, .amount, ".", vbTextCompare)
14590           If intPos01 > 0 Then
14600             If Len(.amount) - intPos01 > 2 Then
14610               blnContinue = False
14620               MsgBox "You can only enter Amount with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
14630               .amount = vbNullString
14640               DoCmd.CancelEvent
14650             End If
14660           End If
14670           If blnContinue = True Then
14680             If Nz(.shareface, 0) = 0 Then
14690               MsgBox "Face must have a value.", vbInformation + vbOKOnly, "Invalid Entry"
14700               .shareface = 1
14710               .cmbAssets.SetFocus
14720               If lngCurrID = 150& Then
14730                 .pershare = .amount / IIf(Nz(.shareface, 0) = 0, 1, .shareface)  '5 dec.
14740                 .pershare.BackColor = CLR_DISABLED_BG
14750               Else
14760                 .pershare = .amount / IIf(Nz(.shareface, 0) = 0, 1, .shareface)  '5 dec.
14770                 .pershare.BackColor = CLR_LTGRN 'lngBkClr
14780               End If
14790               .pershare.Locked = True
14800               .pershare.Enabled = False
14810               .pershare.ForeColor = CLR_DISABLED_FG
14820             Else
14830               If blnDrivingMeBackwards = False Then
14840                 If gblnGoToReport = False Then
14850                   .transdate.SetFocus
14860                 Else
14870 On Error Resume Next
14880                   .cmdOK.SetFocus
14890 On Error GoTo ERRH
14900                 End If
14910               Else
14920                 blnDrivingMeBackwards = False
14930               End If
14940               intInputType = 1
14950               dblPerShare_Calc = (.amount / IIf(Nz(.shareface, 0) = 0, 1, .shareface))
14960               If lngCurrID = 150& Then
14970                 .pershare = Round((dblPerShare_Calc * 100#), 5) '5 dec.
14980                 .pershare.BackColor = CLR_DISABLED_BG
14990               Else
15000                 .pershare = Round((dblPerShare_Calc * 100#), 5)  '5 dec.
15010                 .pershare.BackColor = CLR_LTGRN  'lngBkClr
15020                 dblRate2 = .curr_id.Column(CBX_C_RATE2)
15030                 dblTmp01 = (.amount * dblRate2)
15040                 .Amount_usd = Round(dblTmp01, 2)
15050                 .Amount_usd.Visible = True
                      'dblTmp01 = (dblPerShare_Calc * dblRate2)  ' ** Not currency.
                      'dblTmp01 = Round(dblTmp01, 5)
                      '.pershare_usd = dblTmp01
                      '.pershare_usd.Visible = True
15060               End If
15070               .pershare.Locked = True
15080               .pershare.Enabled = False
15090               .pershare.ForeColor = CLR_DISABLED_FG
15100             End If
15110           End If
15120         End If
15130       End If
15140     Else
15150       .pershare.Enabled = True
15160       .pershare.Locked = False
15170       .pershare.ForeColor = CLR_BLK
15180       If lngCurrID = 150& Then
15190         .pershare.BackColor = CLR_WHT
15200       Else
15210         .pershare.BackColor = CLR_VLTGRN
15220       End If
15230     End If

15240   End With

EXITP:
15250   Exit Sub

ERRH:
15260   Select Case ERR.Number
        Case Else
15270     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15280   End Select
15290   Resume EXITP

End Sub

Private Sub pershare_Exit(Cancel As Integer)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "pershare_Exit"

        Dim dblRate2 As Double
        Dim intPos01 As Integer
        Dim blnContinue As Boolean
        Dim dblTmp01 As Double

15310   With Me

15320     blnContinue = True

15330     lngCurrID = .curr_id.Column(CBX_C_CURRID)

15340     If IsNumeric(.pershare) = True Then
15350       If .pershare < 0 Then
15360         blnContinue = False
15370         MsgBox "Per Share must be positive.", vbInformation + vbOKOnly, "Invalid Entry"
15380         DoCmd.CancelEvent
15390         .pershare.SetFocus
15400       Else
15410         If CDbl(Rem_Dollar(.pershare.text, lngCurrID)) > 999999999.9999 Then  ' ** Module Function: modStringFuncs.
15420           blnContinue = False
15430           MsgBox "Per Share is too large.", vbInformation + vbOKOnly, "Invalid Entry"
15440           .pershare = vbNullString
15450           DoCmd.CancelEvent
15460           .pershare.SetFocus
15470         Else
15480           intPos01 = InStr(1, .pershare, ".", vbTextCompare)
15490           If intPos01 > 0 Then
15500             If Len(.pershare) - intPos01 > 5 Then
15510               blnContinue = False
15520               MsgBox "You can only enter a Per Share with up to 5 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
15530               .pershare = vbNullString
15540               DoCmd.CancelEvent
15550             End If
15560           End If
15570           If blnContinue = True Then
15580             intInputType = 2
15590             If lngCurrID = 150& Then
15600               .amount = Round(((.pershare / 100#) * Nz(.shareface, 1)), 2)
15610               .amount.BackColor = CLR_DISABLED_BG
15620             Else
15630               .amount = Round(((.pershare / 100#) * Nz(.shareface, 1)), 2)
15640               .amount.BackColor = CLR_LTGRN  'lngBkClr
15650               dblRate2 = .curr_id.Column(CBX_C_RATE2)
                    'blTmp01 = (.pershare * dblRate2)  ' ** Not currency.
                    'dblTmp01 = Round(dblTmp01, 5)
                    '.pershare_usd = dblTmp01
                    '.pershare_usd.Visible = True
15660               dblTmp01 = (.amount * dblRate2)
15670               dblTmp01 = Round(dblTmp01, 2)
15680               .Amount_usd = dblTmp01
15690               .Amount_usd.Visible = True
15700             End If
15710             .amount.Locked = True
15720             .amount.Enabled = False
15730             .amount.ForeColor = CLR_DISABLED_FG
15740           End If
15750         End If
15760       End If
15770     Else
15780       .amount.Enabled = True
15790       .amount.Locked = False
15800       .amount.ForeColor = CLR_BLK
15810       If lngCurrID = 150& Then
15820         .amount.BackColor = CLR_WHT
15830       Else
15840         .amount.BackColor = CLR_VLTGRN
15850       End If
15860     End If

15870   End With

EXITP:
15880   Exit Sub

ERRH:
15890   Select Case ERR.Number
        Case Else
15900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15910   End Select
15920   Resume EXITP

End Sub

Private Sub transdate_KeyDown(KeyCode As Integer, Shift As Integer)

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16020   intShiftDown = (Shift And acShiftMask) > 0
16030   intAltDown = (Shift And acAltMask) > 0
16040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16070       With Me
16080         intRetVal = 0
16090         .cmdOK.SetFocus
16100       End With
16110     End Select
16120   End If

        ' ** Shift keys.
16130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16150       With Me
16160         intRetVal = 0
16170         If .pershare.Enabled = True Then
16180           .pershare.SetFocus
16190         Else
16200           .amount.SetFocus
16210         End If
16220       End With
16230     End Select
16240   End If

        ' ** Ctrl keys.
16250   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
16260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16270       With Me
16280         intRetVal = 0
16290         .cmdCalendar.SetFocus
16300       End With
16310     End Select
16320   End If

EXITP:
16330   KeyCode = intRetVal
16340   Exit Sub

ERRH:
16350   intRetVal = 0
16360   THAT_PROC = THIS_PROC
16370   That_Erl = Erl
16380   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16390   Resume EXITP

End Sub

Private Sub transdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_MouseMove"

16410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
16420   Exit Sub

ERRH:
16430   Select Case ERR.Number
        Case Else
16440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16450   End Select
16460   Resume EXITP

End Sub

Private Sub transdate_Exit(Cancel As Integer)

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_Exit"

        Dim varTmp00 As Variant

16510   With Me
16520 On Error Resume Next
16530     varTmp00 = .transdate.text  ' ** When called from cmdCalendar_Click(), below, it's not always on focus.
16540     If ERR.Number <> 0 Then
16550       varTmp00 = .transdate
16560     End If
16570 On Error GoTo ERRH
16580     If DateCheck_Post(varTmp00) = True Then  ' ** Module Function: modUtilities.
            ' ** '.Text' needed to reliably check date.
16590     Else
16600       If IsNull(.transdate) = True Then
16610         .transdate = Date
16620       Else
16630         If Trim(.transdate) = vbNullString Then
16640           .transdate = Date
16650         Else
16660           Cancel = -1
16670           .transdate.SetFocus
16680         End If
16690       End If
16700     End If
16710   End With

EXITP:
16720   Exit Sub

ERRH:
16730   Select Case ERR.Number
        Case Else
16740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16750   End Select
16760   Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

16810   With Me
16820     blnCalendar1_Focus = True
16830     .cmdCalendar_raised_semifocus_dots_img.Visible = True
16840     .cmdCalendar_raised_img.Visible = False
16850     .cmdCalendar_raised_focus_img.Visible = False
16860     .cmdCalendar_raised_focus_dots_img.Visible = False
16870     .cmdCalendar_sunken_focus_dots_img.Visible = False
16880     .cmdCalendar_raised_img_dis.Visible = False
16890   End With

EXITP:
16900   Exit Sub

ERRH:
16910   Select Case ERR.Number
        Case Else
16920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16930   End Select
16940   Resume EXITP

End Sub

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

17010   With Me
17020     blnCalendar1_MouseDown = True
17030     .cmdCalendar_sunken_focus_dots_img.Visible = True
17040     .cmdCalendar_raised_img.Visible = False
17050     .cmdCalendar_raised_semifocus_dots_img.Visible = False
17060     .cmdCalendar_raised_focus_img.Visible = False
17070     .cmdCalendar_raised_focus_dots_img.Visible = False
17080     .cmdCalendar_raised_img_dis.Visible = False
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

Private Sub cmdCalendar_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

17210   With Me
17220     datStartDate = Date
17230     datEndDate = 0
17240     blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
17250     If blnRetVal = True Then
            ' ** Allow posting up to 1 month into the future.
17260       If datStartDate > DateAdd("m", 1, Date) Then
17270         MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
17280         .transdate = CDate(Format(Date, "mm/dd/yyyy"))
17290       Else
17300         .transdate = datStartDate
17310       End If
17320     Else
17330       .transdate = CDate(Format(Date, "mm/dd/yyyy"))
17340     End If
          ' ** Make the date 'is edited' as if were entered by hand.
17350     .transdate.SetFocus
17360     Cancel = 0
17370     transdate_Exit Cancel  ' ** Procedure: Above.
17380     If Cancel = 0 Then
17390       .cmdOK.SetFocus
17400     End If
17410   End With

EXITP:
17420   Exit Sub

ERRH:
17430   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
17440   Case Else
17450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17460   End Select
17470   Resume EXITP

End Sub

Private Sub cmdCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseMove"

17510   With Me
17520     If blnCalendar1_MouseDown = False Then
17530       Select Case blnCalendar1_Focus
            Case True
17540         .cmdCalendar_raised_focus_dots_img.Visible = True
17550         .cmdCalendar_raised_focus_img.Visible = False
17560       Case False
17570         .cmdCalendar_raised_focus_img.Visible = True
17580         .cmdCalendar_raised_focus_dots_img.Visible = False
17590       End Select
17600       .cmdCalendar_raised_img.Visible = False
17610       .cmdCalendar_raised_semifocus_dots_img.Visible = False
17620       .cmdCalendar_sunken_focus_dots_img.Visible = False
17630       .cmdCalendar_raised_img_dis.Visible = False
17640     End If
17650   End With

EXITP:
17660   Exit Sub

ERRH:
17670   Select Case ERR.Number
        Case Else
17680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17690   End Select
17700   Resume EXITP

End Sub

Private Sub cmdCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseUp"

17810   With Me
17820     .cmdCalendar_raised_focus_dots_img.Visible = True
17830     .cmdCalendar_raised_img.Visible = False
17840     .cmdCalendar_raised_semifocus_dots_img.Visible = False
17850     .cmdCalendar_raised_focus_img.Visible = False
17860     .cmdCalendar_sunken_focus_dots_img.Visible = False
17870     .cmdCalendar_raised_img_dis.Visible = False
17880     blnCalendar1_MouseDown = False
17890   End With

EXITP:
17900   Exit Sub

ERRH:
17910   Select Case ERR.Number
        Case Else
17920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17930   End Select
17940   Resume EXITP

End Sub

Private Sub cmdCalendar_LostFocus()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_LostFocus"

18010   With Me
18020     .cmdCalendar_raised_img.Visible = True
18030     .cmdCalendar_raised_semifocus_dots_img.Visible = False
18040     .cmdCalendar_raised_focus_img.Visible = False
18050     .cmdCalendar_raised_focus_dots_img.Visible = False
18060     .cmdCalendar_sunken_focus_dots_img.Visible = False
18070     .cmdCalendar_raised_img_dis.Visible = False
18080     blnCalendar1_Focus = False
18090   End With

EXITP:
18100   Exit Sub

ERRH:
18110   Select Case ERR.Number
        Case Else
18120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18130   End Select
18140   Resume EXITP

End Sub

Private Sub IncludeCurrency()

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngTmp01 As Long

18210   With Me
18220     .curr_id.Enabled = False
18230     .curr_id.Locked = True
18240     .curr_id.ForeColor = CLR_DISABLED_FG
18250     .curr_id.BackColor = CLR_VLTGRN
18260     Select Case blnCurrID
          Case True
18270       .curr_id.Visible = True
18280       .curr_id_cmd.Visible = True
18290       .curr_id_cmd.Enabled = True
18300       .cmbAssets_vline01.Visible = True
18310       .cmbAssets_vline02.Visible = True
18320     Case False
18330       .curr_id.Visible = False
18340       .curr_id_cmd.Visible = False
18350       .curr_id_cmd.Enabled = False
18360       .cmbAssets_vline01.Visible = False
18370       .cmbAssets_vline02.Visible = False
18380       lngTmp01 = ((.curr_id.Left + .curr_id.Width) - (.cmbAssets.Left + .cmbAssets.Width))
18390       .cmbAssets.Width = (.cmbAssets.Width + lngTmp01)
18400     End Select
18410     .Amount_usd.Visible = False
18420     .pershare_usd.Visible = False
18430   End With

        ' ** cmbAssets RowSource's:
        ' **   qryMap_Int_05a
        ' **   qryMap_Int_05c
        ' **   qryMap_Int_05d
        ' **   qryMap_Int_05f
        ' **   qryMap_Int_05g
        ' **   qryMap_Int_05i

EXITP:
18440   Exit Sub

ERRH:
18450   DoCmd.Hourglass False
18460   Select Case ERR.Number
        Case Else
18470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18480   End Select
18490   Resume EXITP

End Sub

Private Sub NoChar_Load()

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

18510   Set dbs = CurrentDb
18520   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
18530     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
18540     Set rst = qdf.OpenRecordset
18550     With rst
18560       .MoveLast
18570       lngNoChars = .RecordCount
18580       .MoveFirst
18590       arr_varNoChar = .GetRows(lngNoChars)
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
18600       .Close
18610     End With
18620     .Close
18630   End With

EXITP:
18640   Set rst = Nothing
18650   Set qdf = Nothing
18660   Set dbs = Nothing
18670   Exit Sub

ERRH:
18680   DoCmd.Hourglass False
18690   Select Case ERR.Number
        Case Else
18700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18710   End Select
18720   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

18810   With Me
18820     blnGTR_Emblem = False
18830     For lngX = 1& To 24&
18840       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
18850     Next
18860   End With

EXITP:
18870   Exit Sub

ERRH:
18880   Select Case ERR.Number
        Case Else
18890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18900   End Select
18910   Resume EXITP

End Sub
