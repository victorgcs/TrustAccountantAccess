VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Reinvest_DivInt_Price"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Reinvest_DivInt_Price"

'VGC 08/30/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   OK:               O {cmdOK}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

'QRY: qryJournal_Columns_30_08           REF: FormRef('frmMap_Reinvest_DivInt_Price')
'QRY: qryMap_Reinvest_02_04              REF: FormRef('frmMap_Reinvest_DivInt_Price')

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
Private strJType As String, datPostingDate As Date, blnIsCancel As Boolean
Private lngAssetNo As Long, dblPerShare As Double, curAmount As Currency, dblShareface As Double
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strTmp01 = Trim(.OpenArgs)
150           If CharCnt(strTmp01, "~") = 5 Then  ' ** Module Function: modStringFuncs.

                ' ** strCallingForm & "~" & "Interest" & "~" & lngAssetNo & "~" &
                ' ** curAmount & "~" & dblShareFace & "~" & lngCurrID
160             intPos01 = InStr(strTmp01, "~")
170             strCallingForm = Left(strTmp01, (intPos01 - 1))
180             strTmp01 = Mid(strTmp01, (intPos01 + 1))
190             intPos01 = InStr(strTmp01, "~")
200             strJType = Left(strTmp01, (intPos01 - 1))
210             strTmp01 = Mid(strTmp01, (intPos01 + 1))
220             intPos01 = InStr(strTmp01, "~")
230             lngAssetNo = Val(Left(strTmp01, (intPos01 - 1)))
240             strTmp01 = Mid(strTmp01, (intPos01 + 1))
250             intPos01 = InStr(strTmp01, "~")
260             Select Case strJType
                Case "Dividend"
                  ' ** Form_frmMap_Div.pershare
270               dblPerShare = Val(Left(strTmp01, (intPos01 - 1)))
280               strTmp01 = Mid(strTmp01, (intPos01 + 1))
290               intPos01 = InStr(strTmp01, "~")
                  ' ** Last slot not used for Dividend.
300             Case "Interest"
                  ' ** Form_frmMap_Int.Amount
310               curAmount = Val(Left(strTmp01, (intPos01 - 1)))
320               strTmp01 = Mid(strTmp01, (intPos01 + 1))
330               intPos01 = InStr(strTmp01, "~")
                  ' ** Form_frmMap_Int.shareface
340               dblShareface = Val(Left(strTmp01, (intPos01 - 1)))
350             End Select
360             lngCurrID = Val(Mid(strTmp01, (intPos01 + 1)))

370             .Caption = strJType & " Map Reinvest Price"

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
380             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

390             CLR_DISABLED_FG = CLR_DKGRY
400             CLR_DISABLED_BG = CLR_LTTEAL

410             blnCurrID = Pref_CurrID  ' ** Module Function: modPreferenceFuncs.
420             .curr_id = lngCurrID
430             IncludeCurrency  ' ** Procedure: Below.

440             blnIsCancel = False
450             blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
460             blnCalendar2_Focus = False: blnCalendar2_MouseDown = False
470             blnGTR_Emblem = False: lngGTR_Stat = 0&

480           Else
490             Cancel = -1
500           End If
510         Else
520           Cancel = -1
530         End If
540       Else
550         Cancel = -1
560       End If
570     End With

580     If Cancel = -1 Then
590       Beep
600       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
610     End If

EXITP:
620     Exit Sub

ERRH:
630     DoCmd.Hourglass False
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub Form_Load()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        'Dim cnxn As ADODB.Connection, rsx1 As ADODB.Recordset  ' ** Early binding.
        Dim cnxn As Object, rsx1 As Object                      ' ** Late binding.

        ' ** Copyright Lebans Holdings 1999 Ltd.
        ' ** Create an instance of the Calendar class.
710     Set clsMonthClass = New clsMonthCal
        ' ** You MUST SET the class hWndForm prop!!!
720     clsMonthClass.hWndForm = Me.hwnd
        ' ** Let's default to PositionAtCursor.
730     clsMonthClass.PositionAtCursor = True

740     gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
750     datPostingDate = Date

        'Set rsx1 = New ADODB.Recordset             ' ** Early binding.
760     Set rsx1 = CreateObject("ADODB.Recordset")  ' ** Late binding.
770     Set cnxn = CurrentProject.Connection
780     rsx1.Open "PostingDate", cnxn, adOpenKeyset, adLockOptimistic, adCmdTableDirect
790     With rsx1
800       .MoveFirst
810       .Find "[Username] = '" & gstrJournalUser & "'"
820       If .EOF = False Then
830         If IsNull(.Fields("Posting_Date")) = False Then
840           datPostingDate = .Fields("Posting_Date")
850         Else
860           .Fields("Posting_Date") = datPostingDate
870           .Update
880         End If
890       Else
900         .AddNew
910         .Fields("Posting_Date") = datPostingDate
920         .Fields("Username") = gstrJournalUser
930         .Update
940       End If
950       .Close
960     End With
970     Set rsx1 = Nothing
980     cnxn.Close
990     Set cnxn = Nothing

1000    With Me

1010      .txtPostDate = datPostingDate

1020      Select Case strJType
          Case "Interest"
1030        .txtPrice_lbl2.Visible = True
1040        .txtPrice_lbl.Visible = False
1050        .txtPrice.DefaultValue = 1@
1060        .txtPrice.Enabled = False
1070      Case "Dividend"
1080        .txtPrice_lbl.Visible = True
1090        .txtPrice_lbl2.Visible = False
1100        .txtPrice.DefaultValue = vbNullString
1110        .txtPrice.Enabled = True
1120      End Select

1130      If .txtPrice.Enabled = False Then
1140        If .txtPrice.Locked = False Then .txtPrice.Locked = True
1150        .txtPrice.ForeColor = CLR_DISABLED_FG
1160        If blnCurrID = True And lngCurrID <> 150& Then
1170          .txtPrice.BackColor = CLR_LTGRN
1180        Else
1190          .txtPrice.BackColor = CLR_DISABLED_BG
1200        End If
1210      End If

          ' ** Borrowing these variables from the Court Reports.
1220      If gstrCrtRpt_Version <> vbNullString Then
1230        .cusip = gstrCrtRpt_Ordinal
1240        .cusip.ForeColor = CLR_DISABLED_FG
1250        .cusip.BackColor = CLR_DISABLED_BG
1260        .cusip.Visible = True
1270        .totdesc = gstrCrtRpt_Version
1280        .totdesc.ForeColor = CLR_DISABLED_FG
1290        .totdesc.BackColor = CLR_DISABLED_BG
1300        .totdesc.Visible = True
1310      End If

1320    End With

1330    DoCmd.Hourglass False

EXITP:
1340    Set rsx1 = Nothing
1350    Set cnxn = Nothing
1360    Exit Sub

ERRH:
1370    DoCmd.Hourglass False
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1520    intShiftDown = (Shift And acShiftMask) > 0
1530    intAltDown = (Shift And acAltMask) > 0
1540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Plain keys.
1550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1560      Select Case intRetVal
          Case vbKeyEscape
1570        intRetVal = 0
1580        cmdCancel_Click  ' ** Procedure: Below.
1590      Case vbKeyF7
1600        With Me
1610          intRetVal = 0
1620          If .cmdCalendar1.Enabled = True Then
1630            cmdCalendar1_Click  ' ** Procedure: Below.
1640          End If
1650        End With
1660      Case vbKeyF8
1670        With Me
1680          intRetVal = 0
1690          If .cmdCalendar2.Enabled = True Then
1700            cmdCalendar2_Click  ' ** Procedure: Below.
1710          End If
1720        End With
1730      End Select
1740    End If

        ' ** Ctrl-Shift keys.
1750    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1760      Select Case intRetVal
          Case vbKeyF
1770        With Me
1780          intRetVal = 0
1790          .FocusHolder.SetFocus
1800        End With
1810      End Select
1820    End If

EXITP:
1830    KeyCode = intRetVal
1840    Exit Sub

ERRH:
1850    intRetVal = 0
1860    Select Case ERR.Number
        Case Else
1870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1880    End Select
1890    Resume EXITP

End Sub

Private Sub Form_Timer()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngTmp03 As Long

        Static lngTmp01 As Long, lngTmp02 As Long

1910    With Me
1920      .TimerInterval = 0&
1930      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
1940        lngGTR_Stat = lngGTR_Stat + 1&
1950        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1960        lngTmp01 = lngTmp01 + 1&
1970        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1980        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1990        If lngTmp01 > 1& Then
2000          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
2010        End If
2020        If lngTmp01 = 24& Then
2030          lngTmp01 = 0&
2040          lngTmp02 = lngTmp02 + 1&
2050          If lngTmp02 < 8& Then
2060            .TimerInterval = 50&
2070          Else
2080            blnGTR_Emblem = False
2090          End If
2100        Else
2110          .TimerInterval = 50&
2120        End If
2130      End If  ' ** blnGTR_Emblem.
2140      If gblnGoToReport = True Then
2150        DoCmd.Hourglass True  ' ** Make sure it's still running.
2160        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48& And lngGTR_Stat < 96&) Or _
                (lngGTR_Stat > 96& And lngGTR_Stat < 120&) Or (lngGTR_Stat > 120&)) Then
              ' ** Just let the emblem go.
2170        Else
2180          DoCmd.Hourglass True  ' ** Make sure it's still running.
2190          If blnGTR_Emblem = False Then
2200            blnGTR_Emblem = True
2210            lngTmp01 = 0&: lngTmp02 = 0&
2220            lngTmp03 = (.GoToReport_Emblem_01_img.Width + (8& * lngTpp))
2230            .GoToReport_arw_ok_img.Left = (.GoToReport_arw_ok_img.Left - lngTmp03)
2240            .cmdOK.Left = (.cmdOK.Left - lngTmp03)
2250            .cmdCancel.Left = (.cmdCancel.Left - lngTmp03)
2260          End If
2270          DoEvents
2280          Select Case blnGoingToReport
              Case True
2290            Select Case blnGoingToReport2
                Case True
2300              Select Case .GoToReport_arw_ok_img.Visible
                  Case True
2310                .TimerInterval = 50&
2320                cmdOK_Click  ' ** Procedure: Below.
2330              Case False
2340                .txtTradeDate = Now()
2350                .GoToReport_arw_ok_img.Visible = True
2360                .GoToReport_arw_map_adate_img.Visible = False
2370                .cmdCalendar1.Enabled = False
2380                .cmdCalendar1.Visible = True
2390                .cmdCalendar1_raised_img.Visible = True
2400                .cmdOK.SetFocus
2410                DoEvents
2420                .TimerInterval = 50&
2430              End Select
2440            Case False
2450              blnGoingToReport2 = True
2460              If strJType = "Dividend" Then
2470                .txtPrice.SetFocus
2480                .txtPrice = 1
2490              End If
2500              .GoToReport_arw_map_adate_img.Visible = True
2510              .cmdCalendar1.Enabled = False
2520              .cmdCalendar1.Visible = False
2530              .cmdCalendar1_raised_img.Visible = False
2540              .GoToReport_arw_map_per_img.Visible = False
2550              .FocusHolder.SetFocus  ' ** So they don't see a flashing cursor.
2560              DoEvents
2570              .TimerInterval = 50&
2580            End Select
2590          Case False
2600            blnGoingToReport = True
2610            DoCmd.Hourglass True  ' ** Make sure it's still running.
2620            DoEvents
2630            .FocusHolder.SetFocus  ' ** So they don't see a flashing cursor.
2640            Select Case strJType
                Case "Dividend"
2650              .GoToReport_arw_map_per_img.Visible = True
2660              .TimerInterval = 50&
2670            Case "Interest"
2680              .TimerInterval = 50&
2690            End Select
2700          End Select
2710        End If
2720      End If
2730    End With

EXITP:
2740    Exit Sub

ERRH:
2750    DoCmd.Hourglass False
2760    Select Case ERR.Number
        Case Else
2770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2780    End Select
2790    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2810    If DataErr <> 0 Then
2820      Select Case DataErr
          Case 2501, 3101, 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2237
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
2830      Case 3162
2840        If blnIsCancel = False Then
2850          MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
2860          Response = acDataErrContinue
2870          DoCmd.CancelEvent
2880        Else
2890          Response = acDataErrContinue
2900        End If
2910      Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
2920        If blnIsCancel = False Then
2930          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
2940          Response = acDataErrContinue
2950          DoCmd.CancelEvent
2960        Else
2970          Response = acDataErrContinue
2980        End If
2990      Case Else
3000        zErrorHandler THIS_NAME, "Undefined Form Error", DataErr  ' ** Module Function: modErrorHandler.
3010      End Select
3020    End If

EXITP:
3030    Exit Sub

ERRH:
3040    Select Case ERR.Number
        Case Else
3050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3060    End Select
3070    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3110    If Not clsMonthClass Is Nothing Then
3120      If clsMonthClass.IsCalendar Then
3130        Cancel = -1
3140      Else
3150        Set clsMonthClass = Nothing
3160      End If
3170    End If

EXITP:
3180    Exit Sub

ERRH:
3190    DoCmd.Hourglass False
3200    Select Case ERR.Number
        Case Else
3210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3220    End Select
3230    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3310    With Me
3320      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
3330        Select Case blnCalendar1_Focus
            Case True
3340          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
3350          .cmdCalendar1_raised_img.Visible = False
3360        Case False
3370          .cmdCalendar1_raised_img.Visible = True
3380          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
3390        End Select
3400        .cmdCalendar1_raised_focus_dots_img.Visible = False
3410        .cmdCalendar1_raised_focus_img.Visible = False
3420        .cmdCalendar1_sunken_focus_dots_img.Visible = False
3430        .cmdCalendar1_raised_img_dis.Visible = False
3440      End If
3450      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
3460        Select Case blnCalendar2_Focus
            Case True
3470          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
3480          .cmdCalendar2_raised_img.Visible = False
3490        Case False
3500          .cmdCalendar2_raised_img.Visible = True
3510          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
3520        End Select
3530        .cmdCalendar2_raised_focus_dots_img.Visible = False
3540        .cmdCalendar2_raised_focus_img.Visible = False
3550        .cmdCalendar2_sunken_focus_dots_img.Visible = False
3560        .cmdCalendar2_raised_img_dis.Visible = False
3570      End If
3580    End With

EXITP:
3590    Exit Sub

ERRH:
3600    Select Case ERR.Number
        Case Else
3610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3620    End Select
3630    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub cmdOK_Click()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim datPostingDate As Date, datTradeDate As Date, dblPrice As Double
        Dim strDocName As String
        Dim blnContinue As Boolean

3810    blnContinue = True

3820    With Me

3830      If IsDate(.txtTradeDate) = False Then
3840        blnContinue = False
3850        MsgBox "You must enter a valid trade date!", vbInformation + vbOKOnly, "Invalid Entry"
3860        .txtTradeDate.SetFocus
3870      Else
3880        If IsDate(.txtPostDate) = False Then
3890          blnContinue = False
3900          MsgBox "You must enter a valid date!", vbInformation + vbOKOnly, "Invalid Entry"
3910          .txtPostDate.SetFocus
3920        Else
3930          If IsNumeric(.txtPrice) = False Then
3940            blnContinue = False
3950            MsgBox "You must enter a valid price", vbInformation + vbOKOnly, "Invalid Entry"
3960            .txtPrice.SetFocus
3970          End If
3980        End If
3990      End If

4000      If blnContinue = True Then

4010        DoCmd.Hourglass True
4020        DoEvents

            ' ** Make sure these are up-to-date.
4030        datPostingDate = CDate(.txtPostDate)
4040        datTradeDate = CDate(.txtTradeDate)
4050        dblPrice = CDbl(.txtPrice)

4060        Set dbs = CurrentDb
4070        With dbs
4080          Select Case strCallingForm
              Case "frmJournal"
                ' ** Append ActiveAssets to Journal Map, by specified [astno], [persh], [pric], [transdat].  #curr_id
4090            Set qdf = .QueryDefs("qryMap_Reinvest_01")
4100          Case "frmJournal_Columns"
                ' ** Append ActiveAssets to Journal Map, by specified [astno], [persh], [pric], [transdat].  #curr_id
4110            Set qdf = .QueryDefs("qryJournal_Columns_30_07")
4120          End Select
4130          With qdf.Parameters
4140            Select Case strJType
                Case "Dividend"
4150              ![astno] = lngAssetNo
4160              ![persh] = dblPerShare  ' ** Used for icash (formatted to Currency within query).
4170            Case "Interest"
4180              ![astno] = lngAssetNo
4190              ![persh] = CDbl(curAmount / dblShareface)  ' ** Used for icash (formatted to Currency within query).
4200            End Select
4210            ![pric] = dblPrice  ' ** Used for pershare.
4220            ![transdat] = datPostingDate
4230          End With
4240          qdf.Execute
4250          Set qdf = Nothing
4260          DoEvents

              ' ** Delete qryMap_Reinvest_02_06 (qryMap_Reinvest_02_05 (Journal Map, linked to Account,
              ' ** qryMap_Reinvest_02_02 (qryMap_Reinvest_02_01 (Journal Map, grouped by accountno,
              ' ** assetno, Location_ID), grouped by accountno, assetno, with cnt, Location_ID_min,
              ' ** Location_ID_max), grouped and summed by accountno), just total_shareface < 0.0001,
              ' ** total_icash > -0.01).
4270          Set qdf = .QueryDefs("qryMap_Reinvest_02_07")
4280          qdf.Execute
4290          Set qdf = Nothing
4300          DoEvents

              ' ** Update qryMap_Reinvest_05_09 (tblJournal_Map_Staging3, with DLookups()
              ' ** to qryMap_Reinvest_05_08 (tblJournal_Map_Staging3, linked to Journal Map,
              ' ** with jmap_id2_new, journaltype2_new)).
4310          Set qdf = .QueryDefs("qryMap_Reinvest_05_10")
4320          qdf.Execute
4330          Set qdf = Nothing

4340          .Close
4350        End With
4360        Set dbs = Nothing
4370        DoEvents

            ' ** Borrowing this variable from the Court Reports.
4380        gdblCrtRpt_PrinTot = dblPrice

4390        strDocName = "frmMap_Reinvest_DivInt_Detail"
4400        DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & strJType & "~" & _
              CStr(dblPrice) & "~" & Format(datPostingDate, "mm/dd/yyyy") & "~" & Format(datTradeDate, "mm/dd/yyyy") & "~" & _
              CStr(lngCurrID)

4410        If gblnGoToReport = True Then
4420          DoCmd.Hourglass True  ' ** Make sure it's still running.
4430          DoEvents
4440          .GoToReport_arw_ok_img.Visible = False
4450          Forms(strDocName).cmdCancel.SetFocus
4460          Forms(strDocName).cmdOK.Enabled = False  ' ** Only allow Cancel for this fake reinvestment.
4470          DoEvents
4480          Forms(strDocName).TimerInterval = 50&
4490          blnGoingToReport = False
4500          blnGoingToReport2 = False
4510        End If

4520        DoCmd.Close acForm, THIS_NAME

4530      End If

4540    End With

EXITP:
4550    Set qdf = Nothing
4560    Set dbs = Nothing
4570    Exit Sub

ERRH:
4580    DoCmd.Hourglass False
4590    Select Case ERR.Number
        Case Else
4600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4610    End Select
4620    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4720    intShiftDown = (Shift And acShiftMask) > 0
4730    intAltDown = (Shift And acAltMask) > 0
4740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4760      Select Case intRetVal
          Case vbKeyTab
4770        With Me
4780          intRetVal = 0
4790          .cmdCancel.SetFocus
4800        End With
4810      End Select
4820    End If

        ' ** Shift keys.
4830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4840      Select Case intRetVal
          Case vbKeyTab
4850        With Me
4860          intRetVal = 0
4870          If .txtPostDate.Enabled = True Then
4880            .txtPostDate.SetFocus
4890          ElseIf .txtTradeDate.Enabled = True Then
4900            .txtTradeDate.SetFocus
4910          ElseIf .txtPrice.Enabled = True Then
4920            .txtPrice.SetFocus
4930          Else
4940            .cmdCancel.SetFocus
4950          End If
4960        End With
4970      End Select
4980    End If

        ' ** Ctrl-Shift keys.
4990    If intCtrlDown And (Not intAltDown) And intShiftDown Then
5000      Select Case intRetVal
          Case vbKeyTab
5010        With Me
5020          intRetVal = 0
5030          If .cmdCalendar2.Enabled = True Then
5040            .cmdCalendar2.SetFocus
5050          End If
5060        End With
5070      End Select
5080    End If

EXITP:
5090    KeyCode = intRetVal
5100    Exit Sub

ERRH:
5110    intRetVal = 0
5120    Select Case ERR.Number
        Case Else
5130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5140    End Select
5150    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intTmp01 As Integer
        Dim intX As Integer

5210    DoCmd.Hourglass True
5220    DoEvents

5230    blnIsCancel = True
5240    intTmp01 = 0
5250    Do While Forms.Count > 1
5260      intTmp01 = intTmp01 + 1
5270      For intX = (Forms.Count - 1) To 0 Step -1
5280        Select Case Forms(intX).Name
            Case THIS_NAME
              ' ** Skip.
5290        Case "frmJournal", "frmJournal_Columns"
              ' ** Skip.
5300        Case Else
5310          DoCmd.Close acForm, Forms(intX).Name
5320          Exit For  ' ** Do them one-at-a-time.
5330        End Select
5340      Next
5350      If intTmp01 >= 10 Then Exit Do  ' ** Don't get caught in a loop.
5360    Loop

5370    Set dbs = CurrentDb
5380    With dbs

          ' ** See that posted and reinvested are turned off,
          ' ** and Null-out CheckNum on parent transactions.
5390      Select Case strCallingForm
          Case "frmJournal"
            ' ** Update qryMap_Reinvest_06_01 (tblJournal_Map_Staging3, linked to
            ' ** Journal, with posted_new, Reinvested_new, Check_Num_new).
5400        Set qdf = .QueryDefs("qryMap_Reinvest_06_02")
5410        qdf.Execute
5420        Set qdf = Nothing
5430        DoEvents
            ' ** Empty Journal Map.
5440        Set qdf = .QueryDefs("qryMap_01")
5450        qdf.Execute
5460        Set qdf = Nothing
5470        DoEvents
            ' ** Empty tblJournal_Map_Staging3
5480        Set qdf = .QueryDefs("qryMap_Reinvest_05_07")
5490        qdf.Execute
5500        Set qdf = Nothing
5510      Case "frmJournal_Columns"
            ' ** Update qryJournal_Columns_32_01 (tblJournal_Map_Staging3, linked to
            ' ** tblJournal_Column, with posted_new, Reinvested_new, Check_Num_new).
5520        Set qdf = .QueryDefs("qryJournal_Columns_32_02")
5530        qdf.Execute
5540        Set qdf = Nothing
5550        DoEvents
            ' ** Empty Journal Map.
5560        Set qdf = .QueryDefs("qryJournal_Columns_30_04")
5570        qdf.Execute
5580        Set qdf = Nothing
5590        DoEvents
            ' ** Empty tblJournal_Map_Staging3.
5600        Set qdf = .QueryDefs("qryJournal_Columns_32_03")
5610        qdf.Execute
5620        Set qdf = Nothing
5630      End Select
5640      DoEvents
          ' ** Empty tmpAccount.
5650      Set qdf = .QueryDefs("qryMap_Div_08_01")
5660      qdf.Execute
5670      Set qdf = Nothing
5680      DoEvents
5690      .Close
5700    End With
5710    Set dbs = Nothing
5720    DoEvents

5730    gblnSetFocus = True
5740    Select Case strCallingForm
        Case "frmJournal"
5750      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5760        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
5770      Else
5780        gstrReturningForm = THIS_NAME
5790        Select Case strJType
            Case "Dividend"
5800          Forms(strCallingForm).frmJournal_Sub1_Dividend.Form.Requery
5810          DoEvents
5820          Forms(strCallingForm).frmJournal_Sub1_Dividend.Form.Map_NewRec  ' ** Form Procedure: frmJournal_Sub1_Dividend.
5830        Case "Interest"
5840          Forms(strCallingForm).frmJournal_Sub2_Interest.Form.Requery
5850          DoEvents
5860          Forms(strCallingForm).frmJournal_Sub2_Interest.Form.Map_NewRec  ' ** Form Procedure: frmJournal_Sub2_Interest.
5870        End Select
5880        DoCmd.SelectObject acForm, strCallingForm, False
5890        Forms(strCallingForm).TimerInterval = 250&
5900      End If
5910    Case "frmJournal_Columns"
5920      gstrReturningForm = "frmMap_Return"
5930      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5940        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
5950      Else
5960        Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
5970        DoCmd.SelectObject acForm, strCallingForm, False
5980        Forms(strCallingForm).TimerInterval = 250&
5990      End If
6000    End Select

6010    DoCmd.Close acForm, THIS_NAME

EXITP:
6020    Set qdf = Nothing
6030    Set dbs = Nothing
6040    Exit Sub

ERRH:
6050    DoCmd.Hourglass False
6060    Select Case ERR.Number
        Case Else
6070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6080    End Select
6090    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6120    intShiftDown = (Shift And acShiftMask) > 0
6130    intAltDown = (Shift And acAltMask) > 0
6140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6160      Select Case intRetVal
          Case vbKeyTab
6170        With Me
6180          intRetVal = 0
6190          If .txtPrice.Enabled = True Then
6200            .txtPrice.SetFocus
6210          ElseIf .txtTradeDate.Enabled = True Then
6220            .txtTradeDate.SetFocus
6230          ElseIf .txtPostDate.Enabled = True Then
6240            .txtPostDate.SetFocus
6250          Else
6260            .cmdOK.SetFocus
6270          End If
6280        End With
6290      End Select
6300    End If

        ' ** Shift keys.
6310    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6320      Select Case intRetVal
          Case vbKeyTab
6330        With Me
6340          intRetVal = 0
6350          .cmdOK.SetFocus
6360        End With
6370      End Select
6380    End If

EXITP:
6390    KeyCode = intRetVal
6400    Exit Sub

ERRH:
6410    intRetVal = 0
6420    Select Case ERR.Number
        Case Else
6430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6440    End Select
6450    Resume EXITP

End Sub

Private Sub txtPrice_Change()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "txtPrice_Change"

6510    With Me
6520      If IsNumeric(.txtPrice.text) = True Then
6530        If CDbl(Rem_Dollar(.txtPrice.text, lngCurrID)) >= 100000 Then  ' ** Module Function: modStringFuncs.
6540          MsgBox "Per Share is too large.", vbInformation + vbOKOnly, "Invalid Entry"
6550          .txtPrice = vbNullString
6560          DoCmd.CancelEvent
6570          .txtPrice.SetFocus
6580        End If
6590      Else
6600        If .txtPrice.text <> "." Then
6610          MsgBox "You must enter a numeric value.", vbInformation + vbOKOnly, "Invalid Entry"
6620          .txtPrice = vbNullString
6630          DoCmd.CancelEvent
6640          .txtPrice.SetFocus
6650        End If
6660      End If
6670    End With

EXITP:
6680    Exit Sub

ERRH:
6690    Select Case ERR.Number
        Case Else
6700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6710    End Select
6720    Resume EXITP

End Sub

Private Sub txtPrice_KeyDown(KeyCode As Integer, Shift As Integer)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "txtPrice_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6820    intShiftDown = (Shift And acShiftMask) > 0
6830    intAltDown = (Shift And acAltMask) > 0
6840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6870        With Me
6880          intRetVal = 0
6890          If .txtTradeDate.Enabled = True Then
6900            .txtTradeDate.SetFocus
6910          ElseIf .txtPostDate.Enabled = True Then
6920            .txtPostDate.SetFocus
6930          Else
6940            .cmdOK.SetFocus
6950          End If
6960        End With
6970      End Select
6980    End If

        ' ** Shift keys.
6990    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7000      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7010        With Me
7020          intRetVal = 0
7030          .cmdCancel.SetFocus
7040        End With
7050      End Select
7060    End If

EXITP:
7070    KeyCode = intRetVal
7080    Exit Sub

ERRH:
7090    intRetVal = 0
7100    Select Case ERR.Number
        Case Else
7110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7120    End Select
7130    Resume EXITP

End Sub

Private Sub txtPrice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "txtPrice_MouseMove"

7210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7220    Exit Sub

ERRH:
7230    Select Case ERR.Number
        Case Else
7240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7250    End Select
7260    Resume EXITP

End Sub

Private Sub txtPrice_Exit(Cancel As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "txtPrice_Exit"

        Dim dblFullValue As Double, dblTruncatedValue As Double, dblRate2 As Double
        Dim dblTmp01 As Double

7310    With Me
7320      If blnIsCancel = False Then
7330        If IsNull(.txtPrice) = False Then
7340          If IsNumeric(.txtPrice) = True Then
7350            dblFullValue = .txtPrice
7360            dblTruncatedValue = Val(Format(dblFullValue, "0.00000"))
7370            If dblFullValue <> dblTruncatedValue Then
7380              Cancel = -1
7390              MsgBox "You can only enter a price with up to 5 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
7400              .txtPrice.Value = vbNullString
7410            Else
7420              If blnCurrID = True And lngCurrID <> 150& Then
7430                dblRate2 = .curr_id.Column(CBX_C_RATE2)
7440                dblTmp01 = Round((dblFullValue * dblRate2), 5)
7450                .txtPrice_usd = dblTmp01
7460                .txtPrice_usd.Visible = True
7470              Else
7480                .txtPrice_usd.Visible = False
7490              End If
7500            End If
7510          Else
7520            .txtPrice_usd = 0
7530            .txtPrice_usd.Visible = False
7540          End If
7550        Else
7560          .txtPrice_usd = 0
7570          .txtPrice_usd.Visible = False
7580        End If
7590      End If
7600    End With

EXITP:
7610    Exit Sub

ERRH:
7620    Select Case ERR.Number
        Case Else
7630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7640    End Select
7650    Resume EXITP

End Sub

Private Sub curr_id_cmd_DblClick(Cancel As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_cmd_DblClick"

7710    With Me
7720      Select Case .curr_date.Visible
          Case True
7730        .curr_date.Visible = False
7740      Case False
7750        .curr_date = .curr_id.Column(CBX_C_DATE)
7760        .curr_date.Visible = True
7770      End Select
7780    End With

EXITP:
7790    Exit Sub

ERRH:
7800    Select Case ERR.Number
        Case Else
7810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
7820    End Select
7830    Resume EXITP

End Sub

Private Sub txtTradeDate_Enter()

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "txtTradeDate_Enter"

        Dim Cancel As Integer

7910    With Me
7920      txtPrice_Exit Cancel  ' ** Procedure: Above.
7930    End With

EXITP:
7940    Exit Sub

ERRH:
7950    Select Case ERR.Number
        Case Else
7960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7970    End Select
7980    Resume EXITP

End Sub

Private Sub txtTradeDate_AfterUpdate()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "txtTradeDate_AfterUpdate"

8010    With Me
8020      If IsNull(.txtTradeDate) = False Then
8030        If Trim(.txtTradeDate) = vbNullString Then
8040          .txtTradeDate = Now()
8050        Else
8060          If blnIsCancel = False Then
8070            If DateCheck_Trade(.txtTradeDate.text) = True Then  ' ** Module Function: modUtilities.
                  ' ** '.Text' needed to reliably check date.
8080              If InStr(CStr(CDbl(CDate(.txtTradeDate))), ".") = 0 Then
8090                .txtTradeDate = .txtTradeDate + time()
8100              End If
8110            Else
8120              .txtTradeDate.Undo
8130            End If
8140          End If
8150        End If
8160      Else
8170        .txtTradeDate = Now()
8180      End If
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

Private Sub txtTradeDate_KeyDown(KeyCode As Integer, Shift As Integer)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "txtTradeDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8320    intShiftDown = (Shift And acShiftMask) > 0
8330    intAltDown = (Shift And acAltMask) > 0
8340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8370        With Me
8380          intRetVal = 0
8390          If .txtPostDate.Enabled = True Then
8400            .txtPostDate.SetFocus
8410          Else
8420            .cmdOK.SetFocus
8430          End If
8440        End With
8450      End Select
8460    End If

        ' ** Shift keys.
8470    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8480      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8490        With Me
8500          intRetVal = 0
8510          If .txtPrice.Enabled = True Then
8520            .txtPrice.SetFocus
8530          Else
8540            .cmdCancel.SetFocus
8550          End If
8560        End With
8570      End Select
8580    End If

        ' ** Ctrl keys.
8590    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8600      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8610        With Me
8620          intRetVal = 0
8630          If .cmdCalendar1.Enabled = True Then
8640            .cmdCalendar1.SetFocus
8650          End If
8660        End With
8670      End Select
8680    End If

EXITP:
8690    KeyCode = intRetVal
8700    Exit Sub

ERRH:
8710    intRetVal = 0
8720    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
8730    Case Else
8740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8750    End Select
8760    Resume EXITP

End Sub

Private Sub txtTradeDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "txtTradeDate_MouseMove"

8810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8820    Exit Sub

ERRH:
8830    Select Case ERR.Number
        Case Else
8840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8850    End Select
8860    Resume EXITP

End Sub

Private Sub txtTradeDate_Exit(Cancel As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "txtTradeDate_Exit"

8910    With Me
8920      If blnIsCancel = False Then
8930        If DateCheck_Trade(.txtTradeDate.text) = True Then  ' ** Module Function: modUtilities.
              ' ** '.Text' needed to reliably check date.
8940          If InStr(CStr(CDbl(CDate(.txtTradeDate))), ".") = 0 Then
8950            .txtTradeDate = .txtTradeDate + time()
8960          End If
8970        Else
8980          If IsNull(.txtTradeDate) = True Then
8990            .txtTradeDate = Now()
9000          Else
9010            If Trim(.txtTradeDate) = vbNullString Then
9020              .txtTradeDate = Now()
9030            Else
9040              Cancel = -1
9050              .txtTradeDate.SetFocus
9060            End If
9070          End If
9080        End If
9090      End If
9100    End With

EXITP:
9110    Exit Sub

ERRH:
9120    Select Case ERR.Number
        Case Else
9130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9140    End Select
9150    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

9210    With Me
9220      blnCalendar1_Focus = True
9230      .cmdCalendar1_raised_semifocus_dots_img.Visible = True
9240      .cmdCalendar1_raised_img.Visible = False
9250      .cmdCalendar1_raised_focus_img.Visible = False
9260      .cmdCalendar1_raised_focus_dots_img.Visible = False
9270      .cmdCalendar1_sunken_focus_dots_img.Visible = False
9280      .cmdCalendar1_raised_img_dis.Visible = False
9290    End With

EXITP:
9300    Exit Sub

ERRH:
9310    Select Case ERR.Number
        Case Else
9320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9330    End Select
9340    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

9410    With Me
9420      blnCalendar1_MouseDown = True
9430      .cmdCalendar1_sunken_focus_dots_img.Visible = True
9440      .cmdCalendar1_raised_img.Visible = False
9450      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
9460      .cmdCalendar1_raised_focus_img.Visible = False
9470      .cmdCalendar1_raised_focus_dots_img.Visible = False
9480      .cmdCalendar1_raised_img_dis.Visible = False
9490    End With

EXITP:
9500    Exit Sub

ERRH:
9510    Select Case ERR.Number
        Case Else
9520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9530    End Select
9540    Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

9610    With Me
9620      datStartDate = Date
9630      datEndDate = 0
9640      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
9650      If blnRetVal = True Then
9660        .txtTradeDate = datStartDate + time()
9670      Else
9680        .txtTradeDate = Now()
9690      End If
9700      If .txtPostDate.Enabled = True Then
9710        .txtPostDate.SetFocus
9720      Else
9730        .cmdOK.SetFocus
9740      End If
9750    End With

EXITP:
9760    Exit Sub

ERRH:
9770    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
9780    Case Else
9790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9800    End Select
9810    Resume EXITP

End Sub

Private Sub cmdCalendar1_KeyDown(KeyCode As Integer, Shift As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_KeyDown"

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
9990          If .txtPostDate.Enabled = True Then
10000           .txtPostDate.SetFocus
10010         Else
10020           .cmdOK.SetFocus
10030         End If
10040       End With
10050     End Select
10060   End If

        ' ** Shift keys.
10070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10080     Select Case intRetVal
          Case vbKeyTab
10090       With Me
10100         intRetVal = 0
10110         If .txtTradeDate.Enabled = True Then
10120           .txtTradeDate.SetFocus
10130         ElseIf .txtPrice.Enabled = True Then
10140           .txtPrice.SetFocus
10150         Else
10160           .cmdCancel.SetFocus
10170         End If
10180       End With
10190     End Select
10200   End If

EXITP:
10210   KeyCode = intRetVal
10220   Exit Sub

ERRH:
10230   intRetVal = 0
10240   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
10250   Case Else
10260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10270   End Select
10280   Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

10310   With Me
10320     If blnCalendar1_MouseDown = False Then
10330       Select Case blnCalendar1_Focus
            Case True
10340         .cmdCalendar1_raised_focus_dots_img.Visible = True
10350         .cmdCalendar1_raised_focus_img.Visible = False
10360       Case False
10370         .cmdCalendar1_raised_focus_img.Visible = True
10380         .cmdCalendar1_raised_focus_dots_img.Visible = False
10390       End Select
10400       .cmdCalendar1_raised_img.Visible = False
10410       .cmdCalendar1_raised_semifocus_dots_img.Visible = False
10420       .cmdCalendar1_sunken_focus_dots_img.Visible = False
10430       .cmdCalendar1_raised_img_dis.Visible = False
10440     End If
10450   End With

EXITP:
10460   Exit Sub

ERRH:
10470   Select Case ERR.Number
        Case Else
10480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10490   End Select
10500   Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

10610   With Me
10620     .cmdCalendar1_raised_focus_dots_img.Visible = True
10630     .cmdCalendar1_raised_img.Visible = False
10640     .cmdCalendar1_raised_semifocus_dots_img.Visible = False
10650     .cmdCalendar1_raised_focus_img.Visible = False
10660     .cmdCalendar1_sunken_focus_dots_img.Visible = False
10670     .cmdCalendar1_raised_img_dis.Visible = False
10680     blnCalendar1_MouseDown = False
10690   End With

EXITP:
10700   Exit Sub

ERRH:
10710   Select Case ERR.Number
        Case Else
10720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10730   End Select
10740   Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

10810   With Me
10820     .cmdCalendar1_raised_img.Visible = True
10830     .cmdCalendar1_raised_semifocus_dots_img.Visible = False
10840     .cmdCalendar1_raised_focus_img.Visible = False
10850     .cmdCalendar1_raised_focus_dots_img.Visible = False
10860     .cmdCalendar1_sunken_focus_dots_img.Visible = False
10870     .cmdCalendar1_raised_img_dis.Visible = False
10880     blnCalendar1_Focus = False
10890   End With

EXITP:
10900   Exit Sub

ERRH:
10910   Select Case ERR.Number
        Case Else
10920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10930   End Select
10940   Resume EXITP

End Sub

Private Sub txtPostDate_KeyDown(KeyCode As Integer, Shift As Integer)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "txtPostDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11020   intShiftDown = (Shift And acShiftMask) > 0
11030   intAltDown = (Shift And acAltMask) > 0
11040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11070       With Me
11080         intRetVal = 0
11090         If IsDate(.txtPostDate) = True Then
11100           .cmdOK.SetFocus
11110         Else
11120           .txtPostDate = Date
11130           .cmdOK.SetFocus
11140         End If
11150       End With
11160     End Select
11170   End If

        ' ** Shift keys.
11180   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11190     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11200       With Me
11210         intRetVal = 0
11220         If IsDate(.txtPostDate) = True Then
11230           If .txtTradeDate.Enabled = True Then
11240             .txtTradeDate.SetFocus
11250           ElseIf .txtPrice.Enabled = True Then
11260             .txtPrice.SetFocus
11270           Else
11280             .cmdCancel.SetFocus
11290           End If
11300         Else
11310           .txtPostDate = Date
11320           If .txtTradeDate.Enabled = True Then
11330             .txtTradeDate.SetFocus
11340           ElseIf .txtPrice.Enabled = True Then
11350             .txtPrice.SetFocus
11360           Else
11370             .cmdCancel.SetFocus
11380           End If
11390         End If
11400       End With
11410     End Select
11420   End If

        ' ** Ctrl keys.
11430   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
11440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11450       With Me
11460         intRetVal = 0
11470         If .cmdCalendar2.Enabled = True Then
11480           .cmdCalendar2.SetFocus
11490         End If
11500       End With
11510     End Select
11520   End If

        ' ** Ctrl-Shift keys.
11530   If intCtrlDown And (Not intAltDown) And intShiftDown Then
11540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11550       With Me
11560         intRetVal = 0
11570         If .cmdCalendar1.Enabled = True Then
11580           .cmdCalendar1.SetFocus
11590         End If
11600       End With
11610     End Select
11620   End If

EXITP:
11630   KeyCode = intRetVal
11640   Exit Sub

ERRH:
11650   intRetVal = 0
11660   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
11670   Case Else
11680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11690   End Select
11700   Resume EXITP

End Sub

Private Sub txtPostDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "txtPostDate_MouseMove"

11810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11820   Exit Sub

ERRH:
11830   Select Case ERR.Number
        Case Else
11840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11850   End Select
11860   Resume EXITP

End Sub

Private Sub txtPostDate_Exit(Cancel As Integer)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "txtPostDate_Exit"

11910   With Me
11920     If blnIsCancel = False Then
11930       If DateCheck_Post(.txtPostDate.text) = True Then  ' ** Module Function: modUtilities.
              ' ** '.Text' needed to reliably check date.
11940       Else
11950         If IsNull(.txtPostDate) = True Then
11960           .txtPostDate = Date
11970         Else
11980           If Trim(.txtPostDate) = vbNullString Then
11990             .txtPostDate = Date
12000           Else
12010             Cancel = -1
12020             .txtPostDate.SetFocus
12030           End If
12040         End If
12050       End If
12060     End If
12070   End With

EXITP:
12080   Exit Sub

ERRH:
12090   Select Case ERR.Number
        Case Else
12100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12110   End Select
12120   Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

12210   With Me
12220     blnCalendar2_Focus = True
12230     .cmdCalendar2_raised_semifocus_dots_img.Visible = True
12240     .cmdCalendar2_raised_img.Visible = False
12250     .cmdCalendar2_raised_focus_img.Visible = False
12260     .cmdCalendar2_raised_focus_dots_img.Visible = False
12270     .cmdCalendar2_sunken_focus_dots_img.Visible = False
12280     .cmdCalendar2_raised_img_dis.Visible = False
12290   End With

EXITP:
12300   Exit Sub

ERRH:
12310   Select Case ERR.Number
        Case Else
12320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12330   End Select
12340   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

12410   With Me
12420     blnCalendar2_MouseDown = True
12430     .cmdCalendar2_sunken_focus_dots_img.Visible = True
12440     .cmdCalendar2_raised_img.Visible = False
12450     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
12460     .cmdCalendar2_raised_focus_img.Visible = False
12470     .cmdCalendar2_raised_focus_dots_img.Visible = False
12480     .cmdCalendar2_raised_img_dis.Visible = False
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

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

12610   With Me
12620     datStartDate = Date
12630     datEndDate = 0
12640     blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
12650     If blnRetVal = True Then
12660       .txtPostDate = datStartDate + time()
12670     Else
12680       .txtPostDate = Now()
12690     End If
12700     If .txtPostDate.Enabled = True Then
12710       .txtPostDate.SetFocus
12720     Else
12730       .cmdOK.SetFocus
12740     End If
12750   End With

EXITP:
12760   Exit Sub

ERRH:
12770   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
12780   Case Else
12790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12800   End Select
12810   Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

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
          Case vbKeyTab
12970       With Me
12980         intRetVal = 0
12990         .cmdOK.SetFocus
13000       End With
13010     End Select
13020   End If

        ' ** Shift keys.
13030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13040     Select Case intRetVal
          Case vbKeyTab
13050       With Me
13060         intRetVal = 0
13070         If .txtPostDate.Enabled = True Then
13080           .txtPostDate.SetFocus
13090         ElseIf .txtTradeDate.Enabled = True Then
13100           .txtTradeDate.SetFocus
13110         ElseIf .txtPrice.Enabled = True Then
13120           .txtPrice.SetFocus
13130         Else
13140           .cmdCancel.SetFocus
13150         End If
13160       End With
13170     End Select
13180   End If

EXITP:
13190   KeyCode = intRetVal
13200   Exit Sub

ERRH:
13210   intRetVal = 0
13220   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
13230   Case Else
13240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13250   End Select
13260   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

13310   With Me
13320     If blnCalendar2_MouseDown = False Then
13330       Select Case blnCalendar2_Focus
            Case True
13340         .cmdCalendar2_raised_focus_dots_img.Visible = True
13350         .cmdCalendar2_raised_focus_img.Visible = False
13360       Case False
13370         .cmdCalendar2_raised_focus_img.Visible = True
13380         .cmdCalendar2_raised_focus_dots_img.Visible = False
13390       End Select
13400       .cmdCalendar2_raised_img.Visible = False
13410       .cmdCalendar2_raised_semifocus_dots_img.Visible = False
13420       .cmdCalendar2_sunken_focus_dots_img.Visible = False
13430       .cmdCalendar2_raised_img_dis.Visible = False
13440     End If
13450   End With

EXITP:
13460   Exit Sub

ERRH:
13470   Select Case ERR.Number
        Case Else
13480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13490   End Select
13500   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

13610   With Me
13620     .cmdCalendar2_raised_focus_dots_img.Visible = True
13630     .cmdCalendar2_raised_img.Visible = False
13640     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
13650     .cmdCalendar2_raised_focus_img.Visible = False
13660     .cmdCalendar2_sunken_focus_dots_img.Visible = False
13670     .cmdCalendar2_raised_img_dis.Visible = False
13680     blnCalendar2_MouseDown = False
13690   End With

EXITP:
13700   Exit Sub

ERRH:
13710   Select Case ERR.Number
        Case Else
13720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13730   End Select
13740   Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

13810   With Me
13820     .cmdCalendar2_raised_img.Visible = True
13830     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
13840     .cmdCalendar2_raised_focus_img.Visible = False
13850     .cmdCalendar2_raised_focus_dots_img.Visible = False
13860     .cmdCalendar2_sunken_focus_dots_img.Visible = False
13870     .cmdCalendar2_raised_img_dis.Visible = False
13880     blnCalendar2_Focus = False
13890   End With

EXITP:
13900   Exit Sub

ERRH:
13910   Select Case ERR.Number
        Case Else
13920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13930   End Select
13940   Resume EXITP

End Sub

Private Sub IncludeCurrency()

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim blnFound As Boolean
        Dim intDec As Integer, intLen As Integer, strFrmt As String
        Dim strTmp01 As String
        Dim lngX As Long

14010   With Me

14020     .curr_id.Enabled = False
14030     .curr_id.Locked = True
14040     .curr_id.ForeColor = CLR_DISABLED_FG
14050     .curr_id.BackColor = CLR_VLTGRN
14060     If blnCurrID = True And lngCurrID <> 150& Then

14070       .curr_id.Visible = True
14080       .curr_id_cmd.Visible = True
14090       .curr_id_cmd.Enabled = True
14100       .txtPrice_usd.Visible = False

14110       If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
14120         NoChar_Load  ' ** Procedure: Below.
14130       End If
14140       blnFound = False
14150       For lngX = 0& To (lngNoChars - 1&)
14160         If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                ' ** This currency's symbol does not show in the Fixedsys font.
14170           blnFound = True
14180         End If
14190       Next

14200       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
14210         .txtPrice.FontName = "Courier New"
14220         .txtPrice.FontSize = 10
14230         .txtPrice.FontBold = True
14240       Case False
14250         If .txtPrice.FontName <> "Arial" Then
14260           .txtPrice.FontName = "Arial"
14270           .txtPrice.FontSize = 10
14280           .txtPrice.FontBold = False
14290         End If
14300       End Select

14310       strTmp01 = .curr_id.Column(CBX_C_SYM)  ' ** Symbol.
14320       intDec = 5
14330       intLen = Len(strTmp01)
14340       Select Case intLen
            Case 1
              'strTmp01 = strTmp01
14350       Case 2
14360         strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2)
14370       Case 3
14380         strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3)
14390       Case 4
14400         strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & "\" & Mid(strTmp01, 4)
14410       Case 5
14420         strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & _
                "\" & Mid(strTmp01, 4, 1) & "\" & Mid(strTmp01, 5)
14430       End Select
14440       strFrmt = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"

14450       .txtPrice.Format = strFrmt
14460       .txtPrice.DecimalPlaces = intDec
14470       .txtPrice.BackColor = CLR_LTGRN

14480     Else
14490       .curr_id.Visible = False
14500       .curr_id_cmd.Visible = False
14510       .curr_id_cmd.Enabled = False
14520       .txtPrice_usd.Visible = False
14530     End If
14540   End With

EXITP:
14550   Exit Sub

ERRH:
14560   DoCmd.Hourglass False
14570   Select Case ERR.Number
        Case Else
14580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14590   End Select
14600   Resume EXITP

End Sub

Private Sub NoChar_Load()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

14710   Set dbs = CurrentDb
14720   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
14730     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
14740     Set rst = qdf.OpenRecordset
14750     With rst
14760       .MoveLast
14770       lngNoChars = .RecordCount
14780       .MoveFirst
14790       arr_varNoChar = .GetRows(lngNoChars)
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
14800       .Close
14810     End With
14820     .Close
14830   End With

EXITP:
14840   Set rst = Nothing
14850   Set qdf = Nothing
14860   Set dbs = Nothing
14870   Exit Sub

ERRH:
14880   DoCmd.Hourglass False
14890   Select Case ERR.Number
        Case Else
14900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14910   End Select
14920   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

15010   With Me
15020     blnGTR_Emblem = False
15030     For lngX = 1& To 24&
15040       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
15050     Next
15060   End With

EXITP:
15070   Exit Sub

ERRH:
15080   Select Case ERR.Number
        Case Else
15090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15100   End Select
15110   Resume EXITP

End Sub
