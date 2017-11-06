VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Reinvest_Rec_Price"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Reinvest_Rec_Price"

'VGC 08/30/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   OK:               O {cmdOK}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

'QRY: qryJournal_Columns_35_08           REF: FormRef('frmMap_Reinvest_Rec_Price')
'QRY: qryMap_Received_02_04              REF: FormRef('frmMap_Reinvest_Rec_Price')

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of the Calendar Class.
Public clsMonthClass As clsMonthCal

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Combo box column constants: curr_id.
'Private Const CBX_C_CURRID As Integer = 0  'curr_id
'Private Const CBX_C_CODE   As Integer = 1  'curr_code
'Private Const CBX_C_NAME   As Integer = 2  'curr_name
Private Const CBX_C_SYM    As Integer = 3  'currsym_symbol
'Private Const CBX_C_DEC    As Integer = 4  'curr_decimal
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
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean

Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private lngTpp As Long, blnCurrID As Boolean, lngCurrID As Long
Private datPostingDate As Date, datAssetDate As Date, blnIsCancel As Boolean
Private lngAssetNo As Long, dblPerShare As Double
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strTmp01 = Trim(.OpenArgs)
150           If CharCnt(strTmp01, "~") = 5 Then  ' ** Module Function: modStringFuncs.

                ' ** strCallingForm & "~" & CStr(lngAssetNo) & "~" & CStr(dblPerShare) & "~" & _
                ' **   Format(datPostingDate, "mm/dd/yyyy") & "~" & Format(datAssetDate, "mm/dd/yyyy hh:nn:ss")
160             intPos01 = InStr(strTmp01, "~")
170             strCallingForm = Left(strTmp01, (intPos01 - 1))
180             strTmp01 = Mid(strTmp01, (intPos01 + 1))
190             intPos01 = InStr(strTmp01, "~")
200             lngAssetNo = Val(Left(strTmp01, (intPos01 - 1)))
210             strTmp01 = Mid(strTmp01, (intPos01 + 1))
220             intPos01 = InStr(strTmp01, "~")
230             dblPerShare = Val(Left(strTmp01, (intPos01 - 1)))
240             strTmp01 = Mid(strTmp01, (intPos01 + 1))
250             intPos01 = InStr(strTmp01, "~")
260             datPostingDate = CDate(Left(strTmp01, (intPos01 - 1)))
270             strTmp01 = Mid(strTmp01, (intPos01 + 1))
280             intPos01 = InStr(strTmp01, "~")
290             datAssetDate = CDate(Left(strTmp01, (intPos01 - 1)))
300             lngCurrID = Val(Mid(strTmp01, (intPos01 + 1)))
                ' ** gstrAccountNo has specified account, otherwise vbNullString.

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
310             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

320             CLR_DISABLED_FG = CLR_DKGRY
330             CLR_DISABLED_BG = CLR_LTTEAL

340             blnCurrID = Pref_CurrID  ' ** Module Function: modPreferenceFuncs.
350             .curr_id = lngCurrID
360             IncludeCurrency  ' ** Procedure: Below.

370             gstrFormQuerySpec = THIS_NAME
380             blnIsCancel = False
390             blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
400             blnCalendar2_Focus = False: blnCalendar2_MouseDown = False
410             blnGTR_Emblem = False: lngGTR_Stat = 0&

420             DoCmd.Hourglass False

430           Else
440             Cancel = -1
450           End If
460         Else
470           Cancel = -1
480         End If
490       Else
500         Cancel = -1
510       End If
520     End With

530     If Cancel = -1 Then
540       Beep
550       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
560     End If

EXITP:
570     Set rst = Nothing
580     Set dbs = Nothing
590     Exit Sub

ERRH:
600     DoCmd.Hourglass False
610     Select Case ERR.Number
        Case Else
620       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
630     End Select
640     Resume EXITP

End Sub

Private Sub Form_Load()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        ' ** Copyright Lebans Holdings 1999 Ltd.
        ' ** Create an instance of the Calendar class.
710     Set clsMonthClass = New clsMonthCal
        ' ** You MUST SET the class hWndForm prop!!!
720     clsMonthClass.hWndForm = Me.hwnd
        ' ** Let's default to PositionAtCursor.
730     clsMonthClass.PositionAtCursor = True

740     gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.

750     With Me

760       .txtPostDate = datPostingDate
770       .txtTradeDate = datAssetDate

780       If blnCurrID = True And lngCurrID <> 150& Then
790         .txtPrice.BackColor = CLR_VLTGRN
800       Else
810         .txtPrice.BackColor = CLR_WHT
820       End If

830       If .txtTradeDate.Enabled = False Then
840         .txtTradeDate.ForeColor = CLR_DISABLED_FG
850         .txtTradeDate.BackColor = CLR_DISABLED_BG
860         .cmdCalendar1.Enabled = False
870         .cmdCalendar1_raised_img_dis.Visible = True
880         .cmdCalendar1_raised_img.Visible = False
890         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
900         .cmdCalendar1_raised_focus_img.Visible = False
910         .cmdCalendar1_raised_focus_dots_img.Visible = False
920         .cmdCalendar1_sunken_focus_dots_img.Visible = False
930       End If

940       If .txtPostDate.Enabled = False Then
950         .txtPostDate.ForeColor = CLR_DISABLED_FG
960         .txtTradeDate.BackColor = CLR_DISABLED_BG
970         .cmdCalendar2.Enabled = False
980         .cmdCalendar2_raised_img_dis.Visible = True
990         .cmdCalendar2_raised_img.Visible = False
1000        .cmdCalendar2_raised_semifocus_dots_img.Visible = False
1010        .cmdCalendar2_raised_focus_img.Visible = False
1020        .cmdCalendar2_raised_focus_dots_img.Visible = False
1030        .cmdCalendar2_sunken_focus_dots_img.Visible = False
1040      End If

          ' ** Borrowing these variables from the Court Reports.
1050      If gstrCrtRpt_Version <> vbNullString Then
1060        .cusip = gstrCrtRpt_Ordinal
1070        .cusip.ForeColor = CLR_DISABLED_FG
1080        .cusip.BackColor = CLR_DISABLED_BG
1090        .cusip.Visible = True
1100        .totdesc = gstrCrtRpt_Version
1110        .totdesc.ForeColor = CLR_DISABLED_FG
1120        .totdesc.BackColor = CLR_DISABLED_BG
1130        .totdesc.Visible = True
1140      End If

1150      If gstrAccountNo <> vbNullString Then
            ' ** An account was specified.
1160        If lngTpp = 0& Then
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
1170          lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
1180        End If
1190        .cusip.Top = .totdesc.Top
1200        .totdesc.Width = .shortname.Width
1210        .totdesc.Left = ((.cusip.Left + .cusip.Width) + (4& * lngTpp))
1220        .accountno = gstrAccountNo
1230        .shortname = gstrAccountName
1240        .accountno.Visible = True
1250        .shortname.Visible = True
1260      End If

1270    End With

1280    DoCmd.Hourglass False

EXITP:
1290    Exit Sub

ERRH:
1300    DoCmd.Hourglass False
1310    Select Case ERR.Number
        Case Else
1320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1330    End Select
1340    Resume EXITP

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

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyEscape
1470        intRetVal = 0
1480        cmdCancel_Click  ' ** Procedure: Below.
1490      Case vbKeyF7
1500        With Me
1510          intRetVal = 0
1520          If .cmdCalendar1.Enabled = True Then
1530            cmdCalendar1_Click  ' ** Procedure: Below.
1540          End If
1550        End With
1560      Case vbKeyF8
1570        With Me
1580          intRetVal = 0
1590          If .cmdCalendar2.Enabled = True Then
1600            cmdCalendar2_Click  ' ** Procedure: Below.
1610          End If
1620        End With
1630      End Select
1640    End If

        ' ** Ctrl-Shift keys.
1650    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1660      Select Case intRetVal
          Case vbKeyF
1670        With Me
1680          intRetVal = 0
1690          .FocusHolder.SetFocus
1700        End With
1710      End Select
1720    End If

EXITP:
1730    KeyCode = intRetVal
1740    Exit Sub

ERRH:
1750    intRetVal = 0
1760    Select Case ERR.Number
        Case Else
1770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1780    End Select
1790    Resume EXITP

End Sub

Private Sub Form_Timer()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngTmp03 As Long

        Static lngTmp01 As Long, lngTmp02 As Long

1810    With Me
1820      .TimerInterval = 0&
1830      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
1840        lngGTR_Stat = lngGTR_Stat + 1&
1850        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1860        lngTmp01 = lngTmp01 + 1&
1870        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1880        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1890        If lngTmp01 > 1& Then
1900          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1910        End If
1920        If lngTmp01 = 24& Then
1930          lngTmp01 = 0&
1940          lngTmp02 = lngTmp02 + 1&
1950          If lngTmp02 < 8& Then
1960            .TimerInterval = 50&
1970          Else
1980            blnGTR_Emblem = False
1990          End If
2000        Else
2010          .TimerInterval = 50&
2020        End If
2030      End If  ' ** blnGTR_Emblem.
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
2210                .TimerInterval = 50&
2220                cmdOK_Click  ' ** Procedure: Below.
2230              Case False
2240                .GoToReport_arw_ok_img.Visible = True
2250                .cmdOK.SetFocus
2260                DoEvents
2270                .TimerInterval = 50&
2280              End Select
2290            Case False
2300              blnGoingToReport2 = True
2310              .txtPrice.SetFocus
2320              .txtPrice = 1
2330              .GoToReport_arw_map_per_img.Visible = False
2340              If strCallingForm = "frmJournal_Columns" Then
                    ' ** Don't know why this continues to light up.
2350                Forms(strCallingForm).cmdSpecPurp_Purch_MapSplit_raised_img.Visible = True
2360                Forms(strCallingForm).cmdSpecPurp_Purch_MapSplit_raised_semifocus_dots_img.Visible = False
2370              End If
2380              DoEvents
2390              .TimerInterval = 50&
2400            End Select
2410          Case False
2420            blnGoingToReport = True
2430            DoCmd.Hourglass True  ' ** Make sure it's still running.
2440            DoEvents
2450            .FocusHolder.SetFocus  ' ** So they don't see a flashing cursor.
2460            .GoToReport_arw_map_per_img.Visible = True
2470            If strCallingForm = "frmJournal_Columns" Then
                  ' ** Don't know why this continues to light up.
2480              Forms(strCallingForm).cmdSpecPurp_Purch_MapSplit_raised_img.Visible = True
2490              Forms(strCallingForm).cmdSpecPurp_Purch_MapSplit_raised_semifocus_dots_img.Visible = False
2500            End If
2510            .TimerInterval = 50&
2520          End Select
2530        End If
2540      End If
2550    End With

EXITP:
2560    Exit Sub

ERRH:
2570    DoCmd.Hourglass False
2580    Select Case ERR.Number
        Case Else
2590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2600    End Select
2610    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2710    If DataErr <> 0 Then
2720      Select Case DataErr
          Case 2501, 3101, 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2237
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
2730      Case 3162
2740        If blnIsCancel = False Then
2750          MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
2760          Response = acDataErrContinue
2770          DoCmd.CancelEvent
2780        Else
2790          Response = acDataErrContinue
2800        End If
2810      Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
2820        If blnIsCancel = False Then
2830          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
2840          Response = acDataErrContinue
2850          DoCmd.CancelEvent
2860        Else
2870          Response = acDataErrContinue
2880        End If
2890      Case Else
2900        zErrorHandler THIS_NAME, "Undefined Form Error", DataErr  ' ** Module Function: modErrorHandler
2910      End Select
2920    End If

EXITP:
2930    Exit Sub

ERRH:
2940    Select Case ERR.Number
        Case Else
2950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2960    End Select
2970    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3010    If Not clsMonthClass Is Nothing Then
3020      If clsMonthClass.IsCalendar Then
3030        Cancel = -1
3040      Else
3050        Set clsMonthClass = Nothing
3060      End If
3070    End If

EXITP:
3080    Exit Sub

ERRH:
3090    Select Case ERR.Number
        Case Else
3100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3110    End Select
3120    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3210    With Me
3220      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
3230        Select Case blnCalendar1_Focus
            Case True
3240          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
3250          .cmdCalendar1_raised_img.Visible = False
3260        Case False
3270          .cmdCalendar1_raised_img.Visible = True
3280          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
3290        End Select
3300        .cmdCalendar1_raised_focus_dots_img.Visible = False
3310        .cmdCalendar1_raised_focus_img.Visible = False
3320        .cmdCalendar1_sunken_focus_dots_img.Visible = False
3330        .cmdCalendar1_raised_img_dis.Visible = False
3340      End If
3350      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
3360        Select Case blnCalendar2_Focus
            Case True
3370          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
3380          .cmdCalendar2_raised_img.Visible = False
3390        Case False
3400          .cmdCalendar2_raised_img.Visible = True
3410          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
3420        End Select
3430        .cmdCalendar2_raised_focus_dots_img.Visible = False
3440        .cmdCalendar2_raised_focus_img.Visible = False
3450        .cmdCalendar2_sunken_focus_dots_img.Visible = False
3460        .cmdCalendar2_raised_img_dis.Visible = False
3470      End If
3480    End With

EXITP:
3490    Exit Sub

ERRH:
3500    Select Case ERR.Number
        Case Else
3510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3520    End Select
3530    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub cmdOK_Click()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim dblSum As Double, dblSumShareface As Double, dblOverUnder As Double, dblPrice As Double
        Dim strDocName As String
        Dim blnContinue As Boolean

3710    blnContinue = True

3720    dblSum = 0#
3730    dblSumShareface = 0#
3740    dblOverUnder = 0#

3750    With Me

3760      If IsDate(.txtTradeDate) = False Then
3770        blnContinue = False
3780        MsgBox "You must enter a valid trade date!", vbInformation + vbOKOnly, "Invalid Entry"
3790        .txtTradeDate.SetFocus
3800      Else
3810        If IsDate(.txtPostDate) = False Then
3820          blnContinue = False
3830          MsgBox "You must enter a valid date!", vbInformation + vbOKOnly, "Invalid Entry"
3840          .txtPostDate.SetFocus
3850        Else
3860          If IsNumeric(.txtPrice) = False Then
3870            blnContinue = False
3880            MsgBox "You must enter a valid price", vbInformation + vbOKOnly, "Invalid Entry"
3890            .txtPrice.SetFocus
3900          End If
3910        End If
3920      End If

3930      If blnContinue = True Then

3940        DoCmd.Hourglass True
3950        DoEvents

            ' ** Make sure these are up-to-date.
3960        datPostingDate = .txtPostDate
3970        datAssetDate = .txtTradeDate
3980        dblPrice = .txtPrice

3990        Set dbs = CurrentDb
4000        With dbs
4010          Select Case strCallingForm
              Case "frmJournal"
                '#########################################
                '## HERE'S WHERE IT LOST SPECIFICITY!
                '#########################################
                ' ** gstrAccountNo has specified account.
4020            If gstrAccountNo = vbNullString Then  ' ** Choice was All.
                  ' ** Append ActiveAssets to Journal Map, by specified [astno], [persh], [pric], [transdat].
4030              Set qdf = .QueryDefs("qryMap_Received_01")
4040            Else  ' ** Choice was Specified.
                  ' ** Append ActiveAssets to Journal Map, by specified [astno], [actno], [persh], [pric], [transdat].
4050              Set qdf = .QueryDefs("qryMap_Received_01_01")
4060              With qdf.Parameters
4070                ![actno] = gstrAccountNo
4080              End With
4090            End If
4100          Case "frmJournal_Columns"
4110            If gstrAccountNo = vbNullString Then  ' ** Choice was All.
                  ' ** Append ActiveAssets to Journal Map, by specified [astno], [persh], [pric], [transdat].
4120              Set qdf = .QueryDefs("qryJournal_Columns_35_07")
4130            Else  ' ** Choice was Specified.
                  ' ** Append ActiveAssets to Journal Map, by specified [astno], [actno], [persh], [pric], [transdat].
4140              Set qdf = .QueryDefs("qryJournal_Columns_35_07_01")
4150              With qdf.Parameters
4160                ![actno] = gstrAccountNo
4170              End With
4180            End If
4190          End Select
4200          With qdf.Parameters
4210            ![astno] = lngAssetNo
4220            ![persh] = dblPerShare  ' ** Used for pcash (formatted to Currency within query).
4230            ![pric] = CCur(dblPrice)  ' ** Used for pershare.
4240            ![transdat] = datPostingDate
4250          End With
4260          qdf.Execute
4270          Set qdf = Nothing

              ' ** Update qryMap_Reinvest_05_09 (tblJournal_Map_Staging3, with DLookups()
              ' ** to qryMap_Reinvest_05_08 (tblJournal_Map_Staging3, linked to Journal Map,
              ' ** with jmap_id2_new, journaltype2_new)).
4280          Set qdf = .QueryDefs("qryMap_Reinvest_05_10")
4290          qdf.Execute
4300          Set qdf = Nothing

4310          .Close
4320        End With
4330        Set dbs = Nothing
4340        DoEvents

            ' ** Borrowing this variable from the Court Reports.
4350        gdblCrtRpt_PrinTot = dblPrice

4360        strDocName = "frmMap_Reinvest_Rec_Detail"
4370        DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & CStr(dblPrice) & "~" & _
              Format(datPostingDate, "mm/dd/yyyy") & "~" & Format(datAssetDate, "mm/dd/yyyy hh:nn:ss") & "~" & _
              CStr(lngCurrID)

4380        If gblnGoToReport = True Then
4390          DoCmd.Hourglass True  ' ** Make sure it's still running.
4400          DoEvents
4410          .GoToReport_arw_ok_img.Visible = False
4420          Forms(strDocName).cmdCancel.SetFocus
4430          Forms(strDocName).cmdOK.Enabled = False  ' ** Only allow Cancel for this fake reinvestment.
4440          DoEvents
4450          Forms(strDocName).TimerInterval = 50&
4460          blnGoingToReport = False
4470          blnGoingToReport2 = False
4480        End If

4490        DoCmd.Close acForm, THIS_NAME

4500      End If

4510    End With

EXITP:
4520    Set qdf = Nothing
4530    Set dbs = Nothing
4540    Exit Sub

ERRH:
4550    DoCmd.Hourglass False
4560    Select Case ERR.Number
        Case Else
4570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4580    End Select
4590    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4620    intShiftDown = (Shift And acShiftMask) > 0
4630    intAltDown = (Shift And acAltMask) > 0
4640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4660      Select Case intRetVal
          Case vbKeyTab
4670        With Me
4680          intRetVal = 0
4690          .cmdCancel.SetFocus
4700        End With
4710      End Select
4720    End If

        ' ** Shift keys.
4730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4740      Select Case intRetVal
          Case vbKeyTab
4750        With Me
4760          intRetVal = 0
4770          If .txtPostDate.Enabled = True Then
4780            .txtPostDate.SetFocus
4790          ElseIf .txtTradeDate.Enabled = True Then
4800            .txtTradeDate.SetFocus
4810          Else
4820            .txtPrice.SetFocus
4830          End If
4840        End With
4850      End Select
4860    End If

        ' ** Ctrl-Shift keys.
4870    If intCtrlDown And (Not intAltDown) And intShiftDown Then
4880      Select Case intRetVal
          Case vbKeyTab
4890        With Me
4900          intRetVal = 0
4910          If .cmdCalendar2.Enabled = True Then
4920            .cmdCalendar2.SetFocus
4930          End If
4940        End With
4950      End Select
4960    End If

EXITP:
4970    KeyCode = intRetVal
4980    Exit Sub

ERRH:
4990    intRetVal = 0
5000    Select Case ERR.Number
        Case Else
5010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5020    End Select
5030    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intTmp01 As Integer
        Dim intX As Integer

5110    DoCmd.Hourglass True
5120    DoEvents

5130    blnIsCancel = True
5140    intTmp01 = 0
5150    Do While Forms.Count > 1
5160      intTmp01 = intTmp01 + 1
5170      For intX = (Forms.Count - 1) To 0 Step -1
5180        Select Case Forms(intX).Name
            Case THIS_NAME
              ' ** Skip.
5190        Case "frmJournal", "frmJournal_Columns"
              ' ** Skip.
5200        Case Else
5210          DoCmd.Close acForm, Forms(intX).Name
5220          Exit For  ' ** Do them one-at-a-time.
5230        End Select
5240      Next
5250      If intTmp01 >= 10 Then Exit Do  ' ** Don't get caught in a loop.
5260    Loop

5270    Set dbs = CurrentDb
5280    With dbs

          ' ** See that posted and reinvested are turned off,
          ' ** and Null-out CheckNum on parent transactions.
5290      Select Case strCallingForm
          Case "frmJournal"
            ' ** Update qryMap_Reinvest_06_01 (tblJournal_Map_Staging3, linked to
            ' ** Journal, with posted_new, Reinvested_new, Check_Num_new).
5300        Set qdf = .QueryDefs("qryMap_Reinvest_06_02")
5310        qdf.Execute
5320        Set qdf = Nothing
5330        DoEvents
            ' ** Empty Journal Map.
5340        Set qdf = .QueryDefs("qryMap_01")
5350        qdf.Execute
5360        Set qdf = Nothing
5370        DoEvents
            ' ** Empty tblJournal_Map_Staging3
5380        Set qdf = .QueryDefs("qryMap_Reinvest_05_07")
5390        qdf.Execute
5400        Set qdf = Nothing
5410      Case "frmJournal_Columns"
            ' ** Update qryJournal_Columns_32_01 (tblJournal_Map_Staging3, linked to
            ' ** tblJournal_Column, with posted_new, Reinvested_new, Check_Num_new).
5420        Set qdf = .QueryDefs("qryJournal_Columns_32_02")
5430        qdf.Execute
5440        Set qdf = Nothing
5450        DoEvents
            ' ** Empty Journal Map.
5460        Set qdf = .QueryDefs("qryJournal_Columns_30_04")
5470        qdf.Execute
5480        Set qdf = Nothing
5490        DoEvents
            ' ** Empty tblJournal_Map_Staging3.
5500        Set qdf = .QueryDefs("qryJournal_Columns_32_03")
5510        qdf.Execute
5520        Set qdf = Nothing
5530      End Select
5540      DoEvents
          ' ** Empty tmpAccount.
5550      Set qdf = dbs.QueryDefs("qryMap_Rec_08_01")
5560      qdf.Execute
5570      Set qdf = Nothing
5580      DoEvents

5590      .Close
5600    End With
5610    Set dbs = Nothing
5620    DoEvents

5630    gblnSetFocus = True
5640    Select Case strCallingForm
        Case "frmJournal"
5650      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5660        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
5670      Else
5680        gstrReturningForm = THIS_NAME
5690        gstrFormQuerySpec = strCallingForm
5700        Forms(strCallingForm).frmJournal_Sub5_Misc.Form.Requery
5710        DoCmd.SelectObject acForm, strCallingForm, False
5720        Forms(strCallingForm).TimerInterval = 250&
5730      End If
5740    Case "frmJournal_Columns"
5750      gstrReturningForm = "frmMap_Return"
5760      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5770        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
5780      Else
5790        Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
5800        DoCmd.SelectObject acForm, strCallingForm, False
5810        Forms(strCallingForm).TimerInterval = 250&
5820      End If
5830    End Select

5840    DoCmd.Close acForm, THIS_NAME

EXITP:
5850    Set qdf = Nothing
5860    Set dbs = Nothing
5870    Exit Sub

ERRH:
5880    Select Case ERR.Number
        Case Else
5890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5900    End Select
5910    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6020    intShiftDown = (Shift And acShiftMask) > 0
6030    intAltDown = (Shift And acAltMask) > 0
6040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6060      Select Case intRetVal
          Case vbKeyTab
6070        With Me
6080          intRetVal = 0
6090          .txtPrice.SetFocus
6100        End With
6110      End Select
6120    End If

        ' ** Shift keys.
6130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6140      Select Case intRetVal
          Case vbKeyTab
6150        With Me
6160          intRetVal = 0
6170          .cmdOK.SetFocus
6180        End With
6190      End Select
6200    End If

EXITP:
6210    KeyCode = intRetVal
6220    Exit Sub

ERRH:
6230    intRetVal = 0
6240    Select Case ERR.Number
        Case Else
6250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6260    End Select
6270    Resume EXITP

End Sub

Private Sub txtPrice_Change()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "txtPrice_Change"

6310    With Me
6320      If IsNumeric(.txtPrice.text) = True Then
6330        If CDbl(.txtPrice.text) >= 100000 Then
6340          MsgBox "Per Share is too large.", vbInformation + vbOKOnly, "Invalid Entry"
6350          .txtPrice = vbNullString
6360          DoCmd.CancelEvent
6370          .txtPrice.SetFocus
6380        End If
6390      Else
6400        If .txtPrice.text <> "." Then
6410          MsgBox "You must enter a numeric value.", vbInformation + vbOKOnly, "Invalid Entry"
6420          .txtPrice = vbNullString
6430          DoCmd.CancelEvent
6440          .txtPrice.SetFocus
6450        End If
6460      End If
6470    End With

EXITP:
6480    Exit Sub

ERRH:
6490    Select Case ERR.Number
        Case Else
6500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6510    End Select
6520    Resume EXITP

End Sub

Private Sub txtPrice_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "txtPrice_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6620    intShiftDown = (Shift And acShiftMask) > 0
6630    intAltDown = (Shift And acAltMask) > 0
6640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6670        With Me
6680          intRetVal = 0
6690          If .txtTradeDate.Enabled = True Then
6700            .txtTradeDate.SetFocus
6710          ElseIf .txtPostDate.Enabled = True Then
6720            .txtPostDate.SetFocus
6730          Else
6740            .cmdOK.SetFocus
6750          End If
6760        End With
6770      End Select
6780    End If

        ' ** Shift keys.
6790    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6800      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6810        With Me
6820          intRetVal = 0
6830          .cmdCancel.SetFocus
6840        End With
6850      End Select
6860    End If

EXITP:
6870    KeyCode = intRetVal
6880    Exit Sub

ERRH:
6890    intRetVal = 0
6900    Select Case ERR.Number
        Case Else
6910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6920    End Select
6930    Resume EXITP

End Sub

Private Sub txtPrice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "txtPrice_MouseMove"

7010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7020    Exit Sub

ERRH:
7030    Select Case ERR.Number
        Case Else
7040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7050    End Select
7060    Resume EXITP

End Sub

Private Sub txtPrice_Exit(Cancel As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "txtPrice_Exit"

        Dim dblFullValue As Double, dblTruncatedValue As Double, dblRate2 As Double
        Dim dblTmp01 As Double

7110    With Me
7120      If blnIsCancel = False Then
7130        If IsNull(.txtPrice) = False Then
7140          If IsNumeric(.txtPrice) = True Then
7150            dblFullValue = .txtPrice
7160            dblTruncatedValue = Val(Format(dblFullValue, "0.00000"))
7170            If dblFullValue <> dblTruncatedValue Then
7180              Cancel = -1
7190              MsgBox "You can only enter a price with up to 5 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
7200              .txtPrice.Value = vbNullString
7210            Else
7220              If blnCurrID = True And lngCurrID <> 150& Then
7230                dblRate2 = .curr_id.Column(CBX_C_RATE2)
7240                dblTmp01 = Round((dblFullValue * dblRate2), 5)
7250                .txtPrice_usd = dblTmp01
7260                .txtPrice_usd.Visible = True
7270              Else
7280                .txtPrice_usd.Visible = False
7290              End If
7300            End If
7310          End If
7320        Else
7330          .txtPrice_usd = 0
7340          .txtPrice_usd.Visible = False
7350        End If
7360      Else
7370        .txtPrice_usd = 0
7380        .txtPrice_usd.Visible = False
7390      End If
7400    End With

EXITP:
7410    Exit Sub

ERRH:
7420    Select Case ERR.Number
        Case Else
7430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7440    End Select
7450    Resume EXITP

End Sub

Private Sub curr_id_cmd_DblClick(Cancel As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_cmd_DblClick"

7510    With Me
7520      Select Case .curr_date.Visible
          Case True
7530        .curr_date.Visible = False
7540      Case False
7550        .curr_date = .curr_id.Column(CBX_C_DATE)
7560        .curr_date.Visible = True
7570      End Select
7580    End With

EXITP:
7590    Exit Sub

ERRH:
7600    Select Case ERR.Number
        Case Else
7610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
7620    End Select
7630    Resume EXITP

End Sub

Private Sub txtTradeDate_BeforeUpdate(Cancel As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "txtTradeDate_BeforeUpdate"

7710    With Me
7720      If IsNull(.txtTradeDate) = False Then
7730        If Trim(.txtTradeDate) = vbNullString Then
7740          .txtTradeDate = Now()
7750        Else
7760          If blnIsCancel = False Then
7770            If DateCheck_Trade(.txtTradeDate.text) = True Then  ' ** Module Function: modUtilities.
                  ' ** '.Text' needed to reliably check date.
7780              If InStr(CStr(CDbl(CDate(.txtTradeDate))), ".") = 0 Then
7790                .txtTradeDate = .txtTradeDate + time()
7800              End If
7810            Else
7820              Cancel = -1
7830            End If
7840          End If
7850        End If
7860      Else
7870        .txtTradeDate = Now()
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

Private Sub txtTradeDate_KeyDown(KeyCode As Integer, Shift As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "txtTradeDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8020    intShiftDown = (Shift And acShiftMask) > 0
8030    intAltDown = (Shift And acAltMask) > 0
8040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8070        With Me
8080          intRetVal = 0
8090          If .txtPostDate.Enabled = True Then
8100            .txtPostDate.SetFocus
8110          Else
8120            .cmdOK.SetFocus
8130          End If
8140        End With
8150      End Select
8160    End If

        ' ** Shift keys.
8170    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8180      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8190        With Me
8200          intRetVal = 0
8210          .txtPrice.SetFocus
8220        End With
8230      End Select
8240    End If

        ' ** Ctrl keys.
8250    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8270        With Me
8280          intRetVal = 0
8290          If .cmdCalendar1.Enabled = True Then
8300            .cmdCalendar1.SetFocus
8310          End If
8320        End With
8330      End Select
8340    End If

EXITP:
8350    KeyCode = intRetVal
8360    Exit Sub

ERRH:
8370    intRetVal = 0
8380    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
8390    Case Else
8400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8410    End Select
8420    Resume EXITP

End Sub

Private Sub txtTradeDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "txtTradeDate_MouseMove"

8510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8520    Exit Sub

ERRH:
8530    Select Case ERR.Number
        Case Else
8540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8550    End Select
8560    Resume EXITP

End Sub

Private Sub txtTradeDate_Exit(Cancel As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "txtTradeDate_Exit"

8610    With Me
8620      If blnIsCancel = False Then
8630        If DateCheck_Trade(.txtTradeDate.text) = True Then  ' ** Module Function: modUtilities.
              ' ** '.Text' needed to reliably check date.
8640          If InStr(CStr(CDbl(CDate(.txtTradeDate))), ".") = 0 Then
8650            .txtTradeDate = .txtTradeDate + time()
8660          End If
8670        Else
8680          If IsNull(.txtTradeDate) = True Then
8690            .txtTradeDate = Now()
8700          Else
8710            If Trim(.txtTradeDate) = vbNullString Then
8720              .txtTradeDate = Now()
8730            Else
8740              Cancel = -1
8750              .txtTradeDate.SetFocus
8760            End If
8770          End If
8780        End If
8790      End If
8800    End With

EXITP:
8810    Exit Sub

ERRH:
8820    Select Case ERR.Number
        Case Else
8830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8840    End Select
8850    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

8910    With Me
8920      blnCalendar1_Focus = True
8930      .cmdCalendar1_raised_semifocus_dots_img.Visible = True
8940      .cmdCalendar1_raised_img.Visible = False
8950      .cmdCalendar1_raised_focus_img.Visible = False
8960      .cmdCalendar1_raised_focus_dots_img.Visible = False
8970      .cmdCalendar1_sunken_focus_dots_img.Visible = False
8980      .cmdCalendar1_raised_img_dis.Visible = False
8990    End With

EXITP:
9000    Exit Sub

ERRH:
9010    Select Case ERR.Number
        Case Else
9020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9030    End Select
9040    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

9110    With Me
9120      blnCalendar1_MouseDown = True
9130      .cmdCalendar1_sunken_focus_dots_img.Visible = True
9140      .cmdCalendar1_raised_img.Visible = False
9150      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
9160      .cmdCalendar1_raised_focus_img.Visible = False
9170      .cmdCalendar1_raised_focus_dots_img.Visible = False
9180      .cmdCalendar1_raised_img_dis.Visible = False
9190    End With

EXITP:
9200    Exit Sub

ERRH:
9210    Select Case ERR.Number
        Case Else
9220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9230    End Select
9240    Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

9310    With Me
9320      datStartDate = Date
9330      datEndDate = 0
9340      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
9350      If blnRetVal = True Then
9360        .txtTradeDate = datStartDate + time()
9370      Else
9380        .txtTradeDate = Now()
9390      End If
9400      If .txtPostDate.Enabled = True Then
9410        .txtPostDate.SetFocus
9420      Else
9430        .cmdOK.SetFocus
9440      End If
9450    End With

EXITP:
9460    Exit Sub

ERRH:
9470    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
9480    Case Else
9490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9500    End Select
9510    Resume EXITP

End Sub

Private Sub cmdCalendar1_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_KeyDown"

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
9690          If .txtPostDate.Enabled = True Then
9700            .txtPostDate.SetFocus
9710          Else
9720            .cmdOK.SetFocus
9730          End If
9740        End With
9750      End Select
9760    End If

        ' ** Shift keys.
9770    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9780      Select Case intRetVal
          Case vbKeyTab
9790        With Me
9800          intRetVal = 0
9810          If .txtTradeDate.Enabled = True Then
9820            .txtTradeDate.SetFocus
9830          ElseIf .txtPrice.Enabled = True Then
9840            .txtPrice.SetFocus
9850          Else
9860            .cmdCancel.SetFocus
9870          End If
9880        End With
9890      End Select
9900    End If

EXITP:
9910    KeyCode = intRetVal
9920    Exit Sub

ERRH:
9930    intRetVal = 0
9940    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
9950    Case Else
9960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9970    End Select
9980    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

10010   With Me
10020     If blnCalendar1_MouseDown = False Then
10030       Select Case blnCalendar1_Focus
            Case True
10040         .cmdCalendar1_raised_focus_dots_img.Visible = True
10050         .cmdCalendar1_raised_focus_img.Visible = False
10060       Case False
10070         .cmdCalendar1_raised_focus_img.Visible = True
10080         .cmdCalendar1_raised_focus_dots_img.Visible = False
10090       End Select
10100       .cmdCalendar1_raised_img.Visible = False
10110       .cmdCalendar1_raised_semifocus_dots_img.Visible = False
10120       .cmdCalendar1_sunken_focus_dots_img.Visible = False
10130       .cmdCalendar1_raised_img_dis.Visible = False
10140     End If
10150   End With

EXITP:
10160   Exit Sub

ERRH:
10170   Select Case ERR.Number
        Case Else
10180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10190   End Select
10200   Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

10310   With Me
10320     .cmdCalendar1_raised_focus_dots_img.Visible = True
10330     .cmdCalendar1_raised_img.Visible = False
10340     .cmdCalendar1_raised_semifocus_dots_img.Visible = False
10350     .cmdCalendar1_raised_focus_img.Visible = False
10360     .cmdCalendar1_sunken_focus_dots_img.Visible = False
10370     .cmdCalendar1_raised_img_dis.Visible = False
10380     blnCalendar1_MouseDown = False
10390   End With

EXITP:
10400   Exit Sub

ERRH:
10410   Select Case ERR.Number
        Case Else
10420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10430   End Select
10440   Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

10510   With Me
10520     .cmdCalendar1_raised_img.Visible = True
10530     .cmdCalendar1_raised_semifocus_dots_img.Visible = False
10540     .cmdCalendar1_raised_focus_img.Visible = False
10550     .cmdCalendar1_raised_focus_dots_img.Visible = False
10560     .cmdCalendar1_sunken_focus_dots_img.Visible = False
10570     .cmdCalendar1_raised_img_dis.Visible = False
10580     blnCalendar1_Focus = False
10590   End With

EXITP:
10600   Exit Sub

ERRH:
10610   Select Case ERR.Number
        Case Else
10620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10630   End Select
10640   Resume EXITP

End Sub

Private Sub txtPostDate_KeyDown(KeyCode As Integer, Shift As Integer)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "txtPostDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10720   intShiftDown = (Shift And acShiftMask) > 0
10730   intAltDown = (Shift And acAltMask) > 0
10740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10770       With Me
10780         intRetVal = 0
10790         If IsDate(.txtPostDate) = True Then
10800           .cmdOK.SetFocus
10810         Else
10820           .txtPostDate = Date
10830           .cmdOK.SetFocus
10840         End If
10850       End With
10860     End Select
10870   End If

        ' ** Shift keys.
10880   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10890     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10900       With Me
10910         intRetVal = 0
10920         If IsDate(.txtPostDate) = True Then
10930           If .txtTradeDate.Enabled = True Then
10940             .txtTradeDate.SetFocus
10950           Else
10960             .txtPrice.SetFocus
10970           End If
10980         Else
10990           .txtPostDate = Date
11000           If .txtTradeDate.Enabled = True Then
11010             .txtTradeDate.SetFocus
11020           Else
11030             .txtPrice.SetFocus
11040           End If
11050         End If
11060       End With
11070     End Select
11080   End If

        ' ** Ctrl keys.
11090   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
11100     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11110       With Me
11120         intRetVal = 0
11130         If .cmdCalendar2.Enabled = True Then
11140           .cmdCalendar2.SetFocus
11150         End If
11160       End With
11170     End Select
11180   End If

        ' ** Ctrl-Shift keys.
11190   If intCtrlDown And (Not intAltDown) And intShiftDown Then
11200     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11210       With Me
11220         intRetVal = 0
11230         If .cmdCalendar1.Enabled = True Then
11240           .cmdCalendar1.SetFocus
11250         End If
11260       End With
11270     End Select
11280   End If

EXITP:
11290   KeyCode = intRetVal
11300   Exit Sub

ERRH:
11310   intRetVal = 0
11320   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
11330   Case Else
11340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11350   End Select
11360   Resume EXITP

End Sub

Private Sub txtPostDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "txtPostDate_MouseMove"

11410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11420   Exit Sub

ERRH:
11430   Select Case ERR.Number
        Case Else
11440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11450   End Select
11460   Resume EXITP

End Sub

Private Sub txtPostDate_Exit(Cancel As Integer)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "txtPostDate_Exit"

11510   With Me
11520     If blnIsCancel = False Then
11530       If DateCheck_Post(.txtPostDate.text) = True Then  ' ** Module Function: modUtilities.
              ' ** '.Text' needed to reliably check date.
11540       Else
11550         If IsNull(.txtPostDate) = True Then
11560           .txtPostDate = Date
11570         Else
11580           If Trim(.txtPostDate) = vbNullString Then
11590             .txtPostDate = Date
11600           Else
11610             Cancel = -1
11620             .txtPostDate.SetFocus
11630           End If
11640         End If
11650       End If
11660     End If
11670   End With

EXITP:
11680   Exit Sub

ERRH:
11690   Select Case ERR.Number
        Case Else
11700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11710   End Select
11720   Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

11810   With Me
11820     blnCalendar2_Focus = True
11830     .cmdCalendar2_raised_semifocus_dots_img.Visible = True
11840     .cmdCalendar2_raised_img.Visible = False
11850     .cmdCalendar2_raised_focus_img.Visible = False
11860     .cmdCalendar2_raised_focus_dots_img.Visible = False
11870     .cmdCalendar2_sunken_focus_dots_img.Visible = False
11880     .cmdCalendar2_raised_img_dis.Visible = False
11890   End With

EXITP:
11900   Exit Sub

ERRH:
11910   Select Case ERR.Number
        Case Else
11920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11930   End Select
11940   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

12010   With Me
12020     blnCalendar2_MouseDown = True
12030     .cmdCalendar2_sunken_focus_dots_img.Visible = True
12040     .cmdCalendar2_raised_img.Visible = False
12050     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
12060     .cmdCalendar2_raised_focus_img.Visible = False
12070     .cmdCalendar2_raised_focus_dots_img.Visible = False
12080     .cmdCalendar2_raised_img_dis.Visible = False
12090   End With

EXITP:
12100   Exit Sub

ERRH:
12110   Select Case ERR.Number
        Case Else
12120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12130   End Select
12140   Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

12210   With Me
12220     datStartDate = Date
12230     datEndDate = 0
12240     blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
12250     If blnRetVal = True Then
12260       .txtTradeDate = datStartDate + time()
12270     Else
12280       .txtTradeDate = Now()
12290     End If
12300     If .txtPostDate.Enabled = True Then
12310       .txtPostDate.SetFocus
12320     Else
12330       .cmdOK.SetFocus
12340     End If
12350   End With

EXITP:
12360   Exit Sub

ERRH:
12370   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
12380   Case Else
12390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12400   End Select
12410   Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12520   intShiftDown = (Shift And acShiftMask) > 0
12530   intAltDown = (Shift And acAltMask) > 0
12540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12560     Select Case intRetVal
          Case vbKeyTab
12570       With Me
12580         intRetVal = 0
12590         .cmdOK.SetFocus
12600       End With
12610     End Select
12620   End If

        ' ** Shift keys.
12630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12640     Select Case intRetVal
          Case vbKeyTab
12650       With Me
12660         intRetVal = 0
12670         If .txtPostDate.Enabled = True Then
12680           .txtPostDate.SetFocus
12690         ElseIf .txtTradeDate.Enabled = True Then
12700           .txtTradeDate.SetFocus
12710         ElseIf .txtPrice.Enabled = True Then
12720           .txtPrice.SetFocus
12730         Else
12740           .cmdCancel.SetFocus
12750         End If
12760       End With
12770     End Select
12780   End If

EXITP:
12790   KeyCode = intRetVal
12800   Exit Sub

ERRH:
12810   intRetVal = 0
12820   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
12830   Case Else
12840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12850   End Select
12860   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

12910   With Me
12920     If blnCalendar2_MouseDown = False Then
12930       Select Case blnCalendar2_Focus
            Case True
12940         .cmdCalendar2_raised_focus_dots_img.Visible = True
12950         .cmdCalendar2_raised_focus_img.Visible = False
12960       Case False
12970         .cmdCalendar2_raised_focus_img.Visible = True
12980         .cmdCalendar2_raised_focus_dots_img.Visible = False
12990       End Select
13000       .cmdCalendar2_raised_img.Visible = False
13010       .cmdCalendar2_raised_semifocus_dots_img.Visible = False
13020       .cmdCalendar2_sunken_focus_dots_img.Visible = False
13030       .cmdCalendar2_raised_img_dis.Visible = False
13040     End If
13050   End With

EXITP:
13060   Exit Sub

ERRH:
13070   Select Case ERR.Number
        Case Else
13080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13090   End Select
13100   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

13210   With Me
13220     .cmdCalendar2_raised_focus_dots_img.Visible = True
13230     .cmdCalendar2_raised_img.Visible = False
13240     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
13250     .cmdCalendar2_raised_focus_img.Visible = False
13260     .cmdCalendar2_sunken_focus_dots_img.Visible = False
13270     .cmdCalendar2_raised_img_dis.Visible = False
13280     blnCalendar2_MouseDown = False
13290   End With

EXITP:
13300   Exit Sub

ERRH:
13310   Select Case ERR.Number
        Case Else
13320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13330   End Select
13340   Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

13410   With Me
13420     .cmdCalendar2_raised_img.Visible = True
13430     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
13440     .cmdCalendar2_raised_focus_img.Visible = False
13450     .cmdCalendar2_raised_focus_dots_img.Visible = False
13460     .cmdCalendar2_sunken_focus_dots_img.Visible = False
13470     .cmdCalendar2_raised_img_dis.Visible = False
13480     blnCalendar2_Focus = False
13490   End With

EXITP:
13500   Exit Sub

ERRH:
13510   Select Case ERR.Number
        Case Else
13520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13530   End Select
13540   Resume EXITP

End Sub

Private Sub IncludeCurrency()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim blnFound As Boolean
        Dim intDec As Integer, intLen As Integer, strFrmt As String
        Dim strTmp01 As String
        Dim lngX As Long

13610   With Me

13620     .curr_id.Enabled = False
13630     .curr_id_cmd.Visible = True
13640     .curr_id_cmd.Enabled = True
13650     .curr_id.Locked = True
13660     .curr_id.ForeColor = CLR_DISABLED_FG
13670     .curr_id.BackColor = CLR_VLTGRN
13680     If blnCurrID = True And lngCurrID <> 150& Then

13690       .curr_id.Visible = True
13700       .txtPrice_usd.Visible = False

13710       If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
13720         NoChar_Load  ' ** Procedure: Below.
13730       End If
13740       blnFound = False
13750       For lngX = 0& To (lngNoChars - 1&)
13760         If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                ' ** This currency's symbol does not show in the Fixedsys font.
13770           blnFound = True
13780         End If
13790       Next

13800       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
13810         .txtPrice.FontName = "Courier New"
13820         .txtPrice.FontSize = 10
13830         .txtPrice.FontBold = True
13840       Case False
13850         If .txtPrice.FontName <> "Arial" Then
13860           .txtPrice.FontName = "Arial"
13870           .txtPrice.FontSize = 10
13880           .txtPrice.FontBold = False
13890         End If
13900       End Select

13910       strTmp01 = .curr_id.Column(CBX_C_SYM)  ' ** Symbol.
13920       intDec = 5
13930       intLen = Len(strTmp01)
13940       Select Case intLen
            Case 1
              'strTmp01 = strTmp01
13950       Case 2
13960         strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2)
13970       Case 3
13980         strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3)
13990       Case 4
14000         strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & "\" & Mid(strTmp01, 4)
14010       Case 5
14020         strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & _
                "\" & Mid(strTmp01, 4, 1) & "\" & Mid(strTmp01, 5)
14030       End Select
14040       strFrmt = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"

14050       .txtPrice.Format = strFrmt
14060       .txtPrice.DecimalPlaces = intDec
14070       .txtPrice.BackColor = CLR_LTGRN

14080     Else
14090       .curr_id.Visible = False
14100       .curr_id_cmd.Visible = False
14110       .curr_id_cmd.Enabled = False
'100         .curr_id_cmd.Visible = True
'110         .curr_id_cmd.Enabled = True
'120         .curr_id_cmd.Visible = False
'130         .curr_id_cmd.Enabled = False
14120       .txtPrice_usd.Visible = False
14130     End If
14140   End With

EXITP:
14150   Exit Sub

ERRH:
14160   DoCmd.Hourglass False
14170   Select Case ERR.Number
        Case Else
14180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14190   End Select
14200   Resume EXITP

End Sub

Private Sub NoChar_Load()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

14310   Set dbs = CurrentDb
14320   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
14330     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
14340     Set rst = qdf.OpenRecordset
14350     With rst
14360       .MoveLast
14370       lngNoChars = .RecordCount
14380       .MoveFirst
14390       arr_varNoChar = .GetRows(lngNoChars)
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
14400       .Close
14410     End With
14420     .Close
14430   End With

EXITP:
14440   Set rst = Nothing
14450   Set qdf = Nothing
14460   Set dbs = Nothing
14470   Exit Sub

ERRH:
14480   DoCmd.Hourglass False
14490   Select Case ERR.Number
        Case Else
14500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14510   End Select
14520   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

14610   With Me
14620     blnGTR_Emblem = False
14630     For lngX = 1& To 24&
14640       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
14650     Next
14660   End With

EXITP:
14670   Exit Sub

ERRH:
14680   Select Case ERR.Number
        Case Else
14690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14700   End Select
14710   Resume EXITP

End Sub
