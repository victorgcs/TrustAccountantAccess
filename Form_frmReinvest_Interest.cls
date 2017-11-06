VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReinvest_Interest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmReinvest_Interest"

'VGC 04/23/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           E {cmdCancel}
' **   New Location:     L {cmdReinvestIntLocNew}
' **   Inc/Exp Codes:    N {cmbRevenueCodes_Purchase}
' **   OK:               O {cmdOK}
' **   Tax Codes:        T {cmbTaxCodes_Purchase}
' **   New Asset:        W {cmdReinvestIntAssetNew}

' ** Shortcut Ctrl keys responsive from this form:
' **   Comments:         M {description}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' *************************
' ** Date defaults:
' **   transdate = Date()
' **   assetdate = Now()
' *************************

' ** Shift enumeration:
' **   1  acShiftMask  The bit mask for the SHIFT key.
' **   2  acCtrlMask   The bit mask for the CTRL key.
' **   4  acAltMask    The bit mask for the ALT key.

' ** Button enumeration:
' **   1  acLeftButton    The bit mask for the left mouse button.
' **   2  acRightButton   The bit mask for the right mouse button.
' **   4  acMiddleButton  The bit mask for the middle mouse button.

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Combo box column constants: cmbAssets.
'Private Const CBX_A_DESC   As Integer = 0  'totdesc
'Private Const CBX_A_CUSIP  As Integer = 1  'cusip
'Private Const CBX_A_ASTNO  As Integer = 2  'assetno
'Private Const CBX_A_TYPE   As Integer = 3  'assettype
'Private Const CBX_A_DUE    As Integer = 4  'due
'Private Const CBX_A_RATE   As Integer = 5  'rate
Private Const CBX_A_TAX    As Integer = 6  'taxcode
'Private Const CBX_A_CURRID As Integer = 7  'curr_id

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of the Calendar class.
Public clsMonthClass As clsMonthCal

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

' ** Used with frmCalendar subforms.
Private Const opgDateType_optDateOnly As Integer = 1
Private Const opgDateType_optDateHour As Integer = 2

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long

Private frm As Access.Form
Private strAccountNo As String, strShortName As String, datTransDate As Date, dblICash As Double
Private lngAssetNo As Long, lngJrnlID As Long, blnCurrID As Boolean, lngCurrID As Long
Private lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long ', lngRecsCur As Long
Private THAT_PROC As String, That_Erl As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer
        Dim varTmp00 As Variant, strTmp01 As String, lngTmp02 As Long
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         strTmp01 = Trim(.OpenArgs)
140         If strTmp01 <> vbNullString Then
150           If CharCnt(strTmp01, "~") = 7 Then
                ' ** strCallingForm & "~" & strAccountNo & "~" & datTransDate & "~" & lngCurrID & "~" & _
                ' ** dblICash & "~" & lngAssetNo & "~" & lngJrnlID & "~" & strShortName

160             blnIsOpen = True

170             intPos01 = InStr(strTmp01, "~")
180             strCallingForm = Left(strTmp01, (intPos01 - 1))
190             strTmp01 = Mid(strTmp01, (intPos01 + 1))
200             intPos01 = InStr(strTmp01, "~")
210             strAccountNo = Left(strTmp01, (intPos01 - 1))
220             strTmp01 = Mid(strTmp01, (intPos01 + 1))
230             intPos01 = InStr(strTmp01, "~")
240             datTransDate = CDate(Left(strTmp01, (intPos01 - 1)))
250             strTmp01 = Mid(strTmp01, (intPos01 + 1))
260             intPos01 = InStr(strTmp01, "~")
270             lngCurrID = Val(Left(strTmp01, (intPos01 - 1)))
280             strTmp01 = Mid(strTmp01, (intPos01 + 1))
290             intPos01 = InStr(strTmp01, "~")
300             dblICash = Val(Left(strTmp01, (intPos01 - 1)))
310             strTmp01 = Mid(strTmp01, (intPos01 + 1))
320             intPos01 = InStr(strTmp01, "~")
330             lngAssetNo = Val(Left(strTmp01, (intPos01 - 1)))
340             strTmp01 = Mid(strTmp01, (intPos01 + 1))
350             intPos01 = InStr(strTmp01, "~")
360             lngJrnlID = Val(Left(strTmp01, (intPos01 - 1)))
370             strShortName = Mid(strTmp01, (intPos01 + 1))

                ' ** Copyright Lebans Holdings 1999 Ltd.
                ' ** Create an instance of the Calendar class.
380             Set clsMonthClass = New clsMonthCal
                ' ** You MUST SET the class hWndForm prop!!!
390             clsMonthClass.hWndForm = Me.hwnd
                ' ** Let's default to PositionAtCursor.
400             clsMonthClass.PositionAtCursor = True

410             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

420             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
430             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

440             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
450             lngMonitorNum = 1&: lngTmp02 = 0&
460             EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
470             If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

                ' ** Variables are fed empty, then populated ByRef.
480             GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

490             FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

500             If lngMonitorNum = 1& Then lngTmp02 = (lngFrm_Top + (25& * lngTpp))
510             DoCmd.MoveSize lngFrm_Left - (8& * lngTpp), lngTmp02, lngFrm_Width, lngFrm_Height  'lngFrm_Top + (25& * lngTpp)
520             If lngMonitorNum > 1& Then
530               LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
540             End If

550             Set frm = Forms("frmJournal")
560             frm.Visible = False

570             .cmdReinvestIntAssetNew.Enabled = False
580             .cmdReinvestIntAssetNew.Visible = True
590             .cmdReinvestIntLocNew.Enabled = False
600             .cmdReinvestIntLocNew.Visible = True

610             .accountno.DefaultValue = Chr(34) & strAccountNo & Chr(34)
620             .shortname.DefaultValue = Chr(34) & strShortName & Chr(34)
630             .transdate.DefaultValue = "#" & Format(datTransDate, "mm/dd/yyyy") & "#"
640             .cmbAssets.DefaultValue = lngAssetNo
650             .description.DefaultValue = Chr(34) & "Interest Reinvestment" & Chr(34)

660             blnCurrID = frm.frmJournal_Sub2_Interest.Form.interestCurr_ID.Visible

670             If blnCurrID = True And lngCurrID <> 150& Then  ' ** USD.
680               glngCurrID = lngCurrID
690               .cmbAssets.RowSource = "qryJournal_Interest_07_03"
700               .cmbAssets.Requery
710             End If
720             If lngAssetNo <> 0& Then
730               varTmp00 = DLookup("[curr_id]", "masterasset", "[assetno] = " & CStr(lngAssetNo))
740               If varTmp00 <> lngCurrID Then
750                 lngAssetNo = 0&
760               End If
770             End If
780             .cmbAssets.DefaultValue = lngAssetNo  ' ** May be Zero.

790             If dblICash = 0 Then
800               .ICash.DefaultValue = 0
810             Else
820               .ICash.DefaultValue = (dblICash * -1)
830             End If
840             .PCash.DefaultValue = 0
850             .Cost.DefaultValue = dblICash
860             .CheckNum.DefaultValue = lngJrnlID  ' ** Borrowing this for the Reinvest/Received reference.

870             If blnCurrID = True And lngCurrID <> 150& Then
880               .reinvestintCurr_ID.Visible = True
890               .reinvestintCurr_ID.DefaultValue = lngCurrID
900               .reinvestintCurr_ID_cmd.Visible = True
910               .reinvestintCurr_ID_cmd.Enabled = True
920               .icash_usd.Visible = True
930               .pcash_usd.Visible = True
940               .cost_usd.Visible = True
950               NoChar_Load  ' ** Procedure: Below.
960             Else
970               .reinvestintCurr_ID.Visible = False
980               .reinvestintCurr_ID_cmd.Enabled = False
990               .reinvestintCurr_ID_cmd.Visible = False
1000              .icash_usd.Visible = False
1010              .PCash.Left = .pcash_alt_box.Left
1020              .pcash_lbl.Left = .PCash.Left
1030              .pcash_usd.Visible = False
1040              .Cost.Left = .cost_alt_box.Left
1050              .cost_lbl.Left = .Cost.Left
1060              .cost_usd.Visible = False
1070            End If

1080            DoCmd.Hourglass False

1090            .shareface.SetFocus

1100            blnIsOpen = False

1110          Else
1120            Cancel = -1
1130          End If
1140        Else
1150          Cancel = -1
1160        End If
1170      Else
1180        Cancel = -1
1190      End If
1200    End With

1210    If Cancel = -1 Then
1220      Beep
1230      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1240    End If

EXITP:
1250    Exit Sub

ERRH:
1260    DoCmd.Hourglass False
1270    Select Case ERR.Number
        Case Else
1280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1290    End Select
1300    Resume EXITP

End Sub

Private Sub Form_Load()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim varTmp00 As Variant

1410    With Me

          ' ** Do this whether or not Currency is visible.
1420      CLR_DISABLED_FG = CLR_DKGRY
1430      CLR_DISABLED_BG = CLR_LTTEAL

1440      .reinvestintCurr_ID.Enabled = False
1450      .reinvestintCurr_ID.Locked = True
1460      .reinvestintCurr_ID.ForeColor = CLR_DISABLED_FG
1470      .reinvestintCurr_ID.BackColor = CLR_VLTGRN
1480      reinvestintCurr_ID_AfterUpdate  ' ** Procedure: Below.

1490      Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
1500        .cmbRevenueCodes_Purchase = REVID_INC
1510        .cmbRevenueCodes_Purchase_display = Null
1520      Case False
1530        If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
1540          .cmbRevenueCodes_Purchase_display = Null
1550        Else
1560          .cmbRevenueCodes_Purchase_display = .cmbRevenueCodes_Purchase.Column(1)
1570        End If
1580      End Select

1590      Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
1600        .cmbTaxCodes_Purchase = TAXID_INC
1610        .cmbTaxCodes_Purchase_display = Null
1620      Case False
1630        If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
1640          .cmbTaxCodes_Purchase_display = Null
1650        Else
1660          .cmbTaxCodes_Purchase_display = .cmbTaxCodes_Purchase.Column(1)
1670        End If
1680      End Select

1690      Select Case gblnRevenueExpenseTracking
          Case True
1700        .cmbRevenueCodes_Purchase.Visible = True
1710        .cmbRevenueCodes_Purchase_display.Visible = True
1720        .cmbRevenueCodes_Purchase_lbl2.Visible = False  ' ** Option off msg.
1730      Case False
1740        .cmbRevenueCodes_Purchase.Visible = False
1750        .cmbRevenueCodes_Purchase_display.Visible = False
1760        .cmbRevenueCodes_Purchase_lbl2.Visible = True  ' ** Option off msg.
1770      End Select

1780      Select Case gblnIncomeTaxCoding
          Case True
1790        .cmbTaxCodes_Purchase.Visible = True
1800        .cmbTaxCodes_Purchase_display.Visible = True
1810        .cmbTaxCodes_Purchase_lbl2.Visible = False  ' ** Option off msg.
1820      Case False
1830        .cmbTaxCodes_Purchase.Visible = False
1840        .cmbTaxCodes_Purchase_display.Visible = False
1850        .cmbTaxCodes_Purchase_lbl2.Visible = True  ' ** Option off msg.
1860      End Select

          ' ** ActiveAssets, grouped, by accountno, assetno, Location_ID, with cnt,
          ' ** by specified GlobalVarGet("glngAssetNo","gstrAccountNo").
1870      varTmp00 = DLookup("[Location_ID]", "qryJournal_Interest_16")
1880      If varTmp00 > 1 Then
1890        .Location_ID = varTmp00
1900      End If

1910    End With

EXITP:
1920    Exit Sub

ERRH:
1930    DoCmd.Hourglass False
1940    Select Case ERR.Number
        Case Else
1950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1960    End Select
1970    Resume EXITP

End Sub

Private Sub Form_Current()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strSQL As String

2010    With Me

2020      strSQL = "SELECT masterasset.assetno, masterasset.assettype " & _
            "FROM masterasset INNER JOIN assettype ON masterasset.assettype = assettype.assettype " & _
            "WHERE (((masterasset.assetno) = " & Me.cmbAssets & "));"

2030      Set dbs = CurrentDb
2040      Set rst = dbs.OpenRecordset(strSQL)

2050      If rst("assettype") = "90" Then
2060        .journaltype = "Liability"
2070        If .Cost > 0 Then
2080          .Cost = .Cost * -1
2090        End If
2100      Else
2110        .journaltype = "Purchase"
2120      End If

2130      If .[Location_ID] = 1& Then
2140        .cmbLocations = Null
2150      Else
2160        .cmbLocations = .[Location_ID]
2170      End If

2180      rst.Close
2190      dbs.Close

2200      If .reinvestintCurr_ID.Visible = True Then
2210        If .reinvestintCurr_Date.Visible = True Then .reinvestintCurr_Date.Visible = False
2220        .icash_usd = Null
2230        .icash_usd.Visible = False
2240        .ICash.Format = "Currency"
2250        .ICash.DecimalPlaces = 2
2260        .ICash.BackColor = CLR_WHT
2270        .pcash_usd = Null
2280        .pcash_usd.Visible = False
2290        .PCash.Format = "Currency"
2300        .PCash.DecimalPlaces = 2
2310        .PCash.BackColor = CLR_WHT
2320        .cost_usd = Null
2330        .cost_usd.Visible = False
2340        .Cost.Format = "Currency"
2350        .Cost.DecimalPlaces = 2
2360        .Cost.BackColor = CLR_WHT
2370      End If

          ' ** Set the currency symbol.
2380      reinvestintCurr_ID_AfterUpdate  ' ** Procedure: Below.

2390      If .reinvestintCurr_ID.Visible = True Then
2400        icash_AfterUpdate  ' ** Procedure: Below.
2410        cost_AfterUpdate  ' ** Procedure: Below.
2420      End If

2430    End With

EXITP:
2440    Set rst = Nothing
2450    Set dbs = Nothing
2460    Exit Sub

ERRH:
2470    DoCmd.Hourglass False
2480    Select Case ERR.Number
        Case Else
2490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2500    End Select
2510    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2620    intShiftDown = (Shift And acShiftMask) > 0
2630    intAltDown = (Shift And acAltMask) > 0
2640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Comments:         M {description}

        ' ** Plain keys.
2650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2660      Select Case intRetVal
          Case vbKeyEscape
2670        intRetVal = 0
2680        cmdCancel_Click  ' ** Procedure: Below.
2690      Case vbKeyF7
2700        With Me
2710          intRetVal = 0
2720          .cmdCalendar1.SetFocus
2730          .cmdCalendar1.Form.cmdCalendar_Click  ' ** Form Procedure: frmCalendar (cmdCalendar1).
2740        End With
2750      Case vbKeyF8
2760        With Me
2770          intRetVal = 0
2780          .cmdCalendar2.SetFocus
2790          .cmdCalendar2.Form.cmdCalendar_Click  ' ** Form Procedure: frmCalendar (cmdCalendar2).
2800        End With
2810      End Select
2820    End If

        ' ** Ctrl keys.
2830    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2840      Select Case intRetVal
          Case vbKeyM
2850        With Me
2860          intRetVal = 0
2870          .description.SetFocus
2880        End With
2890      End Select
2900    End If

        ' ** Ctrl-Shift keys.
2910    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2920      Select Case intRetVal
          Case vbKeyF
2930        With Me
2940          intRetVal = 0
2950          .FocusHolder.SetFocus
2960        End With
2970      End Select
2980    End If

EXITP:
2990    KeyCode = intRetVal
3000    Exit Sub

ERRH:
3010    intRetVal = 0
3020    THAT_PROC = THIS_PROC
3030    That_Erl = Erl
3040    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
3050    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3110    If DataErr <> 0 Then
3120      Select Case DataErr
          Case 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2110, 13, 2474, 2135, 10503, 10508
            ' ** Do nothing.
3130      Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
3140        Select Case Me.ActiveControl.Name
            Case "assetdate", "assetdate_display"
3150          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Entry Required"
3160        Case "shareface"
3170          MsgBox "Please enter a valid number.", vbInformation + vbOKOnly, "Entry Required"
3180        End Select
3190        Response = acDataErrContinue
3200      Case 3101
            ' ** Not an item in the table.
3210        DoCmd.CancelEvent
3220      Case 2237
            ' ** Item not in list.
3230        DoCmd.CancelEvent
3240      Case 3162
3250        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
            ' ** Do nothing.
3260      Case 2501  ' ** The '|' action was Canceled.
            ' ** Do nothing.
3270      Case 3315
            ' ** Do nothing.
3280      Case Else
3290        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3300        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3310      End Select
3320    End If

EXITP:
3330    Exit Sub

ERRH:
3340    Select Case ERR.Number
        Case Else
3350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3360    End Select
3370    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3410    DoCmd.Hourglass True  ' ** Make sure it's still running.
3420    DoEvents

3430    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3440    If Not clsMonthClass Is Nothing Then
3450      If clsMonthClass.IsCalendar Then
3460        Cancel = -1
3470      Else
3480        Set clsMonthClass = Nothing
3490      End If
3500    End If

3510    gblnSetFocus = True
3520    gstrReturningForm = THIS_NAME
3530    frm.Visible = True
3540    DoCmd.SelectObject acForm, frm.Name, False
3550    frm.frmJournal_Sub2_Interest.SetFocus
3560    frm.frmJournal_Sub2_Interest.Form.interestAccountNo.SetFocus
3570    frm.TimerInterval = 250&

EXITP:
3580    Set frm = Nothing
3590    Exit Sub

ERRH:
3600    Select Case ERR.Number
        Case Else
3610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3620    End Select
3630    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3710    With Me
3720      .cmdCalendar1.Form.Detail_MouseMove Button, Shift, X, Y
3730      .cmdCalendar2.Form.Detail_MouseMove Button, Shift, X, Y
3740    End With

EXITP:
3750    Exit Sub

ERRH:
3760    Select Case ERR.Number
        Case Else
3770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3780    End Select
3790    Resume EXITP

End Sub

Private Sub cmdOK_Click()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim strRevCode As String, intTaxCode As Integer
        Dim lngJrnlID As Long
        Dim blnContinue As Boolean

3810    blnContinue = True

3820    DoCmd.Hourglass True
3830    DoEvents

3840    With Me

3850      If IsNumeric(.shareface) = False Then
3860        blnContinue = False
3870        DoCmd.Hourglass False
3880        MsgBox "The shareface must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
3890        .shareface.SetFocus
3900      Else
3910        If .shareface = 0 Then
3920          blnContinue = False
3930          DoCmd.Hourglass False
3940          MsgBox "The shareface must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
3950          .shareface.SetFocus
3960        Else
3970          If IsDate(.assetdate) = False Or IsDate(.assetdate_display) = False Then
3980            blnContinue = False
3990            DoCmd.Hourglass False
4000            MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
4010            .assetdate_display.SetFocus
4020          End If
4030        End If
4040      End If

4050      If blnContinue = True Then
4060        If IsNull(.cmbRevenueCodes_Purchase) = True Then
4070          .cmbRevenueCodes_Purchase = REVID_INC  ' ** Unspecified Income.
4080          strRevCode = "I"
4090        Else
4100          If .cmbRevenueCodes_Purchase = 0& Then
4110            .cmbRevenueCodes_Purchase = REVID_INC  ' ** Unspecified Income.
4120            strRevCode = "I"
4130          Else
4140            strRevCode = DLookup("[revcode_TYPE_Code]", "qryRevCodeComboIE", "[revcode_ID] = " & CStr(.cmbRevenueCodes_Purchase))
4150          End If
4160        End If
4170        intTaxCode = Nz(.cmbTaxCodes_Purchase, 0)
4180        Select Case .journaltype
            Case "Purchase"
              ' ** INCOME.
4190          If strRevCode <> "I" Then
4200            .cmbRevenueCodes_Purchase = REVID_INC  ' ** Unspecified Income.
4210          End If
4220          If IsNull(.cmbTaxCodes_Purchase) = True Then
4230            .cmbTaxCodes_Purchase = TAXID_INC
4240          Else
4250            If .taxcode_type = TAXTYP_DED Then
4260              .cmbTaxCodes_Purchase = .cmbAssets.Column(CBX_A_TAX)  ' ** All AssetType-based Tax Codes are INCOME.
4270            End If
4280          End If
4290        End Select
4300      End If

4310      If blnContinue = True Then

4320        DoCmd.Hourglass True  ' ** Make sure it's still running.
4330        DoEvents

            ' ** This was parent.
4340        lngJrnlID = .interestreinvestID
4350        gdblCrtRpt_CostTot = lngJrnlID  ' ** Borrowing this variable from the Court Reports.

4360        .priceperunit = Abs(.Cost / .shareface)

4370        If InStr(CStr(CDbl(CDate(.assetdate))), ".") = 0 Then  ' ** The hidden one.
4380          .assetdate = .assetdate + time()
4390        End If
4400        DoCmd.RunCommand acCmdSaveRecord

4410        DoCmd.Close acForm, THIS_NAME

4420      End If  ' ** blnContinue.

4430    End With  ' ** Me.

EXITP:
4440    Exit Sub

ERRH:
4450    DoCmd.Hourglass False
4460    Select Case ERR.Number
        Case Else
4470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4480    End Select
4490    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim lngJrnlID As Long

4510    With Me

4520      DoCmd.Hourglass True
4530      DoEvents

4540      lngJrnlID = .interestreinvestID

4550      gblnMessage = False
4560      gdblCrtRpt_CostTot = lngJrnlID  ' ** Borrowing this variable from the Court Reports.

4570      DoCmd.Close acForm, THIS_NAME

4580    End With

EXITP:
4590    Exit Sub

ERRH:
4600    DoCmd.Hourglass False
4610    THAT_PROC = THIS_PROC
4620    That_Erl = Erl
4630    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4640    Resume EXITP

End Sub

Private Sub transdate_BeforeUpdate(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_BeforeUpdate"

4710    With Me
4720      If .transdate.text <> "__/__/____" And .transdate.text <> vbNullString Then
4730        If Left(.transdate.text, 2) > 12 Then
4740          MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Entry"
4750          Cancel = -1
4760        End If
4770      End If
4780    End With

EXITP:
4790    Exit Sub

ERRH:
4800    THAT_PROC = THIS_PROC
4810    That_Erl = Erl
4820    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4830    Resume EXITP

End Sub

Private Sub transdate_KeyDown(KeyCode As Integer, Shift As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4920    intShiftDown = (Shift And acShiftMask) > 0
4930    intAltDown = (Shift And acAltMask) > 0
4940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4970        With Me
4980          intRetVal = 0
4990          .cmbAssets.SetFocus
5000        End With
5010      End Select
5020    End If

        ' ** Shift keys.
5030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5050        With Me
5060          intRetVal = 0
5070          If .journaltype.Enabled = True Then
5080            .journaltype.SetFocus
5090          ElseIf .accountno.Enabled = True Then
5100            .accountno.SetFocus
5110          ElseIf .cmbTaxCodes_Purchase_display.Visible = True And .cmbTaxCodes_Purchase_display.Enabled = True Then
5120            .cmbTaxCodes_Purchase_display.SetFocus
5130          ElseIf .cmbRevenueCodes_Purchase_display.Visible = True And .cmbRevenueCodes_Purchase_display.Enabled = True Then
5140            .cmbRevenueCodes_Purchase_display.SetFocus
5150          Else
5160            .description.SetFocus
5170          End If
5180        End With
5190      End Select
5200    End If

        ' ** Ctrl keys.
5210    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
5220      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5230        With Me
5240          intRetVal = 0
5250          .cmdCalendar1.SetFocus
5260        End With
5270      End Select
5280    End If

EXITP:
5290    KeyCode = intRetVal
5300    Exit Sub

ERRH:
5310    intRetVal = 0
5320    THAT_PROC = THIS_PROC
5330    That_Erl = Erl
5340    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5350    Resume EXITP

End Sub

Private Sub transdate_Exit(Cancel As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_Exit"

5410    With Me
5420      If DateCheck_Post(.transdate.text) = True Then  ' ** Module Function: modUtilities.
            ' ** '.Text' needed to reliably check date.
5430      Else
5440        If IsNull(.transdate) = True Then
5450          .transdate = Date
5460        Else
5470          If Trim(.transdate) = vbNullString Then
5480            .transdate = Date
5490          Else
5500            Cancel = -1
5510            .transdate.SetFocus
5520          End If
5530        End If
5540      End If
5550    End With

EXITP:
5560    Exit Sub

ERRH:
5570    Select Case ERR.Number
        Case Else
5580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5590    End Select
5600    Resume EXITP

End Sub

Private Sub cmdCalendar1_Enter()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Enter"

5710    With Me
5720      .cmdCalendar1.Form.DateDest = "transdate"
5730      .cmdCalendar1.Form.opgDateType = opgDateType_optDateOnly
5740      DoEvents
5750      .cmdCalendar1.Form.cmdCalendar.SetFocus
5760    End With

EXITP:
5770    Exit Sub

ERRH:
5780    Select Case ERR.Number
        Case Else
5790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5800    End Select
5810    Resume EXITP

End Sub

Private Sub cmdCalendar1_Exit(Cancel As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Exit"

        Dim Button As Integer, Shift As Integer, X As Single, Y As Single

5910    With Me
5920      .cmdCalendar1.Form.Detail_MouseMove Button, Shift, X, Y
5930    End With

EXITP:
5940    Exit Sub

ERRH:
5950    Select Case ERR.Number
        Case Else
5960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5970    End Select
5980    Resume EXITP

End Sub

Private Sub cmbAssets_Enter()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_Enter"

6010    With Me
6020      ExpandCombo .cmbAssets  ' ** Module Procedure: modUtilities.
6030    End With

EXITP:
6040    Exit Sub

ERRH:
6050    Select Case ERR.Number
        Case Else
6060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6070    End Select
6080    Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
6170        With Me
6180          intRetVal = 0
6190          .shareface.SetFocus
6200        End With
6210      End Select
6220    End If

        ' ** Shift keys.
6230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6250        With Me
6260          intRetVal = 0
6270          .transdate.SetFocus
6280        End With
6290      End Select
6300    End If

        ' ** Ctrl-Shift keys.
6310    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6320      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6330        With Me
6340          intRetVal = 0
6350          .cmdCalendar1.SetFocus
6360        End With
6370      End Select
6380    End If

EXITP:
6390    KeyCode = intRetVal
6400    Exit Sub

ERRH:
6410    intRetVal = 0
6420    THAT_PROC = THIS_PROC
6430    That_Erl = Erl
6440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6450    Resume EXITP

End Sub

Private Sub shareface_Exit(Cancel As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_Exit"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

6510    blnContinue = True

6520    With Me
6530      If IsNumeric(.shareface) = True Then
6540        If .shareface < 0 Then
6550          blnContinue = False
6560          MsgBox "Share/Face must be positive.", vbInformation + vbOKOnly, "Invalid Entry"
6570          Cancel = -1
6580          .shareface.SetFocus
6590        Else
6600          If CDbl(.shareface) > 999999999.9999 Then
6610            blnContinue = False
6620            MsgBox "Share/Face is too large.", vbInformation + vbOKOnly, "Invalid Entry"
6630            .shareface = vbNullString
6640            Cancel = -1
6650            .shareface.SetFocus
6660          End If
6670          If blnContinue = True Then
6680            intPos01 = InStr(1, .shareface, ".", vbTextCompare)
6690            If intPos01 > 0 Then
6700              If Len(.shareface) - intPos01 > gintShareFaceDecimals Then
6710                blnContinue = False
6720                MsgBox "You can only enter a Share/Face with up to " & Trim(str(gintShareFaceDecimals)) & " decimals.", _
                      vbInformation + vbOKOnly, "Invalid Entry"
6730                .shareface = vbNullString
6740                Cancel = -1
6750                .shareface.SetFocus
6760              Else
6770                .shareface.Format = "#,###." & String(gintShareFaceDecimals, "0")
6780              End If
6790            Else
6800              .shareface.Format = "#,###"
6810            End If
6820          End If
6830        End If
6840      Else
6850        blnContinue = False
6860        .shareface = 0
6870        MsgBox "Share/Face must have a value.", vbInformation + vbOKOnly, "Entry Required"
6880        .shareface = vbNullString
6890        DoCmd.RunCommand acCmdSaveRecord
6900        Cancel = -1
6910        .shareface.SetFocus
6920      End If
6930    End With  ' ** Me.

EXITP:
6940    Exit Sub

ERRH:
6950    Select Case ERR.Number
        Case Else
6960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6970    End Select
6980    Resume EXITP

End Sub

Private Sub assetdate_display_KeyDown(KeyCode As Integer, Shift As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
7070        With Me
7080          intRetVal = 0
7090          .cmbLocations.SetFocus
7100        End With
7110      End Select
7120    End If

        ' ** Shift keys.
7130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7150        With Me
7160          intRetVal = 0
7170          .shareface.SetFocus
7180        End With
7190      End Select
7200    End If

        ' ** Ctrl keys.
7210    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
7220      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7230        With Me
7240          intRetVal = 0
7250          .cmdCalendar2.SetFocus
7260        End With
7270      End Select
7280    End If

EXITP:
7290    KeyCode = intRetVal
7300    Exit Sub

ERRH:
7310    intRetVal = 0
7320    Select Case ERR.Number
        Case Else
7330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7340    End Select
7350    Resume EXITP

End Sub

Private Sub assetdate_display_Exit(Cancel As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_Exit"

7410    With Me
7420      If DateCheck_Trade(.assetdate_display.text) = True Then  ' ** Module Function: modUtilities.
            ' ** '.Text' needed to reliably check date.
7430        If InStr(CStr(CDbl(CDate(.assetdate_display))), ".") = 0 Then
7440          .assetdate = .assetdate_display + time()
7450        Else
7460          .assetdate = .assetdate_display
7470          .assetdate_display = CDate(Format(.assetdate_display, "mm/dd/yyyy"))
7480        End If
7490      Else
7500        If IsNull(.assetdate_display) = True Then
7510          .assetdate_display = Date
7520          .assetdate = Now()
7530        Else
7540          If Trim(.assetdate_display) = vbNullString Then
7550            .assetdate_display = Date
7560            .assetdate = Now()
7570          Else
7580            Cancel = -1
7590            .assetdate_display.SetFocus
7600          End If
7610        End If
7620      End If
7630    End With

EXITP:
7640    Exit Sub

ERRH:
7650    THAT_PROC = THIS_PROC
7660    That_Erl = Erl
7670    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7680    Resume EXITP

End Sub

Private Sub cmdCalendar2_Enter()

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Enter"

7710    With Me
7720      .cmdCalendar2.Form.DateDest = "assetdate"
7730      .cmdCalendar2.Form.opgDateType = opgDateType_optDateHour
7740      DoEvents
7750      .cmdCalendar2.Form.cmdCalendar.SetFocus
7760    End With

EXITP:
7770    Exit Sub

ERRH:
7780    Select Case ERR.Number
        Case Else
7790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7800    End Select
7810    Resume EXITP

End Sub

Private Sub cmdCalendar2_Exit(Cancel As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Exit"

        Dim Button As Integer, Shift As Integer, X As Single, Y As Single

7910    With Me
7920      .cmdCalendar2.Form.Detail_MouseMove Button, Shift, X, Y
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

Private Sub cmbLocations_Enter()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_Enter"

8010    With Me
8020      If IsNull(.cmbLocations) = True Then
8030        .[Location_ID] = 1&
8040      Else
8050        If .cmbLocations = 0& Then
8060          .[Location_ID] = 1&
8070        End If
8080      End If
8090    End With

EXITP:
8100    Exit Sub

ERRH:
8110    THAT_PROC = THIS_PROC
8120    That_Erl = Erl
8130    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8140    Resume EXITP

End Sub

Private Sub cmbLocations_GotFocus()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_GotFocus"

8210    With Me
8220      If .Location_ID = 1& Then
8230        ExpandCombo .cmbLocations  ' ** Module Procedure: modUtilities.
8240      End If
8250    End With

EXITP:
8260    Exit Sub

ERRH:
8270    THAT_PROC = THIS_PROC
8280    That_Erl = Erl
8290    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8300    Resume EXITP

End Sub

Private Sub cmbLocations_AfterUpdate()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_AfterUpdate"

8410    With Me
8420      .[Location_ID] = .cmbLocations
8430    End With

EXITP:
8440    Exit Sub

ERRH:
8450    THAT_PROC = THIS_PROC
8460    That_Erl = Erl
8470    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8480    Resume EXITP

End Sub

Private Sub cmbLocations_KeyDown(KeyCode As Integer, Shift As Integer)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8520    intShiftDown = (Shift And acShiftMask) > 0
8530    intAltDown = (Shift And acAltMask) > 0
8540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8570        With Me
8580          intRetVal = 0
8590          If .ICash.Enabled = True Then
8600            .ICash.SetFocus
8610          ElseIf .Cost.Enabled = True Then
8620            .Cost.SetFocus
8630          Else
8640            .description.SetFocus
8650          End If
8660        End With
8670      End Select
8680    End If

        ' ** Shift keys.
8690    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8700      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8710        With Me
8720          intRetVal = 0
8730          .assetdate_display.SetFocus
8740        End With
8750      End Select
8760    End If

        ' ** Ctrl-Shift keys.
8770    If intCtrlDown And (Not intAltDown) And intShiftDown Then
8780      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8790        With Me
8800          intRetVal = 0
8810          .cmdCalendar2.SetFocus
8820        End With
8830      End Select
8840    End If

EXITP:
8850    KeyCode = intRetVal
8860    Exit Sub

ERRH:
8870    intRetVal = 0
8880    Select Case ERR.Number
        Case Else
8890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8900    End Select
8910    Resume EXITP

End Sub

Private Sub icash_AfterUpdate()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "ICash_AfterUpdate"

        Dim dblICash As Double
        Dim varTmp00 As Variant

9010    With Me
9020      If blnCurrID = True And lngCurrID <> 150& Then
9030        .icash_usd = 0#
9040        If .icash_usd.Visible = False Then .icash_usd.Visible = True
9050        If .pcash_usd.Visible = False Then .pcash_usd.Visible = True
9060        If .cost_usd.Visible = False Then .cost_usd.Visible = True
9070        If IsNull(.ICash) = False Then
9080          If .ICash <> 0 Then
9090            If IsNull(.reinvestintCurr_ID.Column(CBX_C_CURRID)) = False Then
9100              lngCurrID = .reinvestintCurr_ID.Column(CBX_C_CURRID)
9110              If lngCurrID <> 150& Then  ' ** USD.
9120                varTmp00 = .reinvestintCurr_ID.Column(CBX_C_RATE2)
9130                If IsNull(varTmp00) = False Then
9140                  If varTmp00 > 0 Then
9150                    dblICash = .ICash
9160                    dblICash = (dblICash * varTmp00)
9170                    dblICash = Round(dblICash, 2)
9180                    .icash_usd = dblICash
9190                  End If
9200                End If
9210              End If
9220            End If
9230          End If
9240        End If
9250      End If
9260      DoEvents
9270    End With

EXITP:
9280    Exit Sub

ERRH:
9290    Select Case ERR.Number
        Case Else
9300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
9310    End Select
9320    Resume EXITP

End Sub

Private Sub icash_Exit(Cancel As Integer)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_Exit"

9410    With Me
9420      Select Case .journaltype
          Case "Purchase"
9430        If .ICash > 0 Then
9440          .ICash = .ICash * -1
9450          If .icash_usd.Visible = True Then
9460            .icash_usd = Abs(.icash_usd) * -1
9470          End If
9480        End If
9490      Case "Liability"
9500        If .ICash < 0 Then
9510          .ICash = Abs(.ICash)
9520          If .icash_usd.Visible = True Then
9530            .icash_usd = Abs(.icash_usd)
9540          End If
9550        End If
9560      End Select
9570    End With

EXITP:
9580    Exit Sub

ERRH:
9590    Select Case ERR.Number
        Case Else
9600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9610    End Select
9620    Resume EXITP

End Sub

Private Sub pcash_AfterUpdate()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_AfterUpdate"

        Dim dblPCash As Double
        Dim varTmp00 As Variant

9710    With Me
9720      If blnCurrID = True And lngCurrID <> 150& Then
9730        .pcash_usd = 0#
9740        If .icash_usd.Visible = False Then .icash_usd.Visible = True
9750        If .pcash_usd.Visible = False Then .pcash_usd.Visible = True
9760        If .cost_usd.Visible = False Then .cost_usd.Visible = True
9770        If IsNull(.PCash) = False Then
9780          If .PCash <> 0 Then
9790            If IsNull(.reinvestintCurr_ID.Column(CBX_C_CURRID)) = False Then
9800              lngCurrID = .reinvestintCurr_ID.Column(CBX_C_CURRID)
9810              If lngCurrID <> 150& Then  ' ** USD.
9820                varTmp00 = .reinvestintCurr_ID.Column(CBX_C_RATE2)
9830                If IsNull(varTmp00) = False Then
9840                  If varTmp00 > 0 Then
9850                    dblPCash = .PCash
9860                    dblPCash = (dblPCash * varTmp00)
9870                    dblPCash = Round(dblPCash, 2)
9880                    .pcash_usd = dblPCash
9890                  End If
9900                End If
9910              End If
9920            End If
9930          End If
9940        End If
9950      End If
9960      DoEvents
9970    End With

EXITP:
9980    Exit Sub

ERRH:
9990    Select Case ERR.Number
        Case Else
10000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10010   End Select
10020   Resume EXITP

End Sub

Private Sub pcash_Exit(Cancel As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_Exit"

10110   With Me
10120     Select Case .journaltype
          Case "Purchase"
10130       If .PCash > 0 Then
10140         .PCash = .PCash * -1
10150         If .pcash_usd.Visible = True Then
10160           .pcash_usd = Abs(.pcash_usd) * -1
10170         End If
10180       End If
10190     Case "Liability"
10200       If .PCash < 0 Then
10210         .PCash = Abs(.PCash)
10220         If .pcash_usd.Visible = True Then
10230           .pcash_usd = Abs(.pcash_usd)
10240         End If
10250       End If
10260     End Select
10270   End With

EXITP:
10280   Exit Sub

ERRH:
10290   Select Case ERR.Number
        Case Else
10300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10310   End Select
10320   Resume EXITP

End Sub

Private Sub cost_AfterUpdate()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "Cost_AfterUpdate"

        Dim dblCost As Double
        Dim varTmp00 As Variant

10410   With Me
10420     If .reinvestintCurr_ID.Visible = True Then
10430       .cost_usd = Null
10440       If .icash_usd.Visible = False Then .icash_usd.Visible = True
10450       If .pcash_usd.Visible = False Then .pcash_usd.Visible = True
10460       If .cost_usd.Visible = False Then .cost_usd.Visible = True
10470       If IsNull(.Cost) = False Then
10480         If .Cost <> 0 Then
10490           If IsNull(.reinvestintCurr_ID.Column(CBX_C_CURRID)) = False Then
10500             lngCurrID = .reinvestintCurr_ID.Column(CBX_C_CURRID)
10510             If lngCurrID <> 150& Then  ' ** USD.
10520               varTmp00 = .reinvestintCurr_ID.Column(CBX_C_RATE2)
10530               If IsNull(varTmp00) = False Then
10540                 If varTmp00 > 0 Then
10550                   dblCost = .Cost
10560                   dblCost = (dblCost * varTmp00)
10570                   dblCost = Round(dblCost, 2)
10580                   .cost_usd = dblCost
10590                 End If
10600               End If
10610             End If
10620           End If
10630         End If
10640       End If
10650     End If
10660     DoEvents
10670   End With

EXITP:
10680   Exit Sub

ERRH:
10690   Select Case ERR.Number
        Case Else
10700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
10710   End Select
10720   Resume EXITP

End Sub

Private Sub cost_Exit(Cancel As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_Exit"

10810   With Me
10820     Select Case .journaltype
          Case "Purchase"
10830       If .Cost < 0 Then
10840         .Cost = Abs(.Cost)
10850         If .cost_usd.Visible = True Then
10860           .cost_usd = Abs(.cost_usd)
10870         End If
10880       End If
10890     Case "Liability"
10900       If .Cost > 0 Then
10910         .Cost = Abs(.Cost) * -1
10920         If .cost_usd.Visible = True Then
10930           .cost_usd = Abs(.cost_usd) * -1
10940         End If
10950       End If
10960     End Select
10970   End With

EXITP:
10980   Exit Sub

ERRH:
10990   Select Case ERR.Number
        Case Else
11000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11010   End Select
11020   Resume EXITP

End Sub

Private Sub reinvestintCurr_ID_cmd_DblClick(Cancel As Integer)

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "reinvestintCurr_ID_cmd_DblClick"

11110   With Me
11120     If IsNull(.reinvestintCurr_Date) = True Then
11130       .reinvestintCurr_Date = .reinvestintCurr_ID.Column(CBX_C_DATE)
11140     End If
11150     DoEvents
11160     Select Case .reinvestintCurr_Date.Visible
          Case True
11170       .reinvestintCurr_Date.Visible = False
11180     Case False
11190       .reinvestintCurr_Date.Visible = True
11200     End Select
11210   End With

EXITP:
11220   Exit Sub

ERRH:
11230   Select Case ERR.Number
        Case Else
11240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11250   End Select
11260   Resume EXITP

End Sub

Private Sub reinvestintCurr_ID_AfterUpdate()

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "reinvestintCurr_ID_AfterUpdate"

        Dim lngCurrID As Long
        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String, strTmp03 As String, intTmp04 As Integer, lngTmp06 As Long
        Dim lngX As Long

11310   With Me

11320     If IsNull(.reinvestintCurr_ID) = False Then

11330       lngCurrID = .reinvestintCurr_ID.Column(CBX_C_CURRID)

11340       blnFound = False
11350       If lngCurrID <> 150& Then
11360         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
11370           NoChar_Load  ' ** Procedure: Below.
11380         End If
11390         For lngX = 0& To (lngNoChars - 1&)
11400           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
11410             blnFound = True
11420           End If
11430         Next
11440       End If

11450       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
11460         .ICash.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
11470         .ICash.FontSize = 10
11480         .ICash.FontBold = True
11490         .PCash.FontName = "Courier New"
11500         .PCash.FontSize = 10
11510         .PCash.FontBold = True
11520         .Cost.FontName = "Courier New"
11530         .Cost.FontSize = 10
11540         .Cost.FontBold = True
11550       Case False
11560         If .ICash.FontName <> "Fixedsys" Then
11570           .ICash.FontName = "Fixedsys"
11580           .ICash.FontSize = 9
11590           .ICash.FontBold = False
11600           .PCash.FontName = "Fixedsys"
11610           .PCash.FontSize = 9
11620           .PCash.FontBold = False
11630           .Cost.FontName = "Fixedsys"
11640           .Cost.FontSize = 9
11650           .Cost.FontBold = False
11660         End If
11670       End Select

11680       If lngCurrID = 150& Then  ' ** USD.
11690         strTmp01 = "Currency"
11700         intTmp04 = 2
11710         lngTmp06 = CLR_WHT
11720       Else
11730         strTmp03 = .reinvestintCurr_ID.Column(CBX_C_SYM)  ' ** Symbol.
11740         If .reinvestintCurr_ID.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
11750           intTmp04 = 4
11760         Else
11770           intTmp04 = .reinvestintCurr_ID.Column(CBX_C_DEC)  ' ** Decimal places.
11780         End If
11790         intLen = Len(strTmp03)
11800         Select Case intLen
              Case 1
                'strTmp03 = strTmp03
11810         Case 2
11820           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2)
11830         Case 3
11840           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3)
11850         Case 4
11860           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & "\" & Mid(strTmp03, 4)
11870         Case 5
11880           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & _
                  "\" & Mid(strTmp03, 4, 1) & "\" & Mid(strTmp03, 5)
11890         End Select
11900         Select Case intTmp04
              Case 0
11910           strTmp01 = strTmp03 & "#,##0;(" & strTmp03 & "#,##0);" & strTmp03 & "0;" & strTmp03 & "0"
11920         Case 1  ' ** None have this currently.
11930           strTmp01 = strTmp03 & "#,##0.0;(" & strTmp03 & "#,##0.0);" & strTmp03 & "0.0;" & strTmp03 & "0.0"
11940         Case 2
11950           strTmp01 = strTmp03 & "#,##0.00;(" & strTmp03 & "#,##0.00);" & strTmp03 & "0.00;" & strTmp03 & "0.00"
11960         Case 3
11970           strTmp01 = strTmp03 & "#,##0.000;(" & strTmp03 & "#,##0.000);" & strTmp03 & "0.000;" & strTmp03 & "0.000"
11980         End Select
11990         lngTmp06 = CLR_VLTGRN
12000       End If
12010       .ICash.Format = strTmp01
12020       .ICash.DecimalPlaces = intTmp04
12030       .ICash.BackColor = lngTmp06
12040       .PCash.Format = strTmp01
12050       .PCash.DecimalPlaces = intTmp04
12060       .PCash.BackColor = lngTmp06
12070       .Cost.Format = strTmp01
12080       .Cost.DecimalPlaces = intTmp04
12090       .Cost.BackColor = lngTmp06
12100     End If

12110   End With

EXITP:
12120   Exit Sub

ERRH:
12130   THAT_PROC = THIS_PROC
12140   That_Erl = Erl
12150   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12160   Resume EXITP

End Sub

Private Sub reinvestintCurr_ID_NotInList(NewData As String, Response As Integer)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "reinvestintCurr_ID_NotInList"

12210   With Me
12220     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
12230     .reinvestintCurr_ID.SetFocus
12240     Response = acDataErrContinue
12250   End With

EXITP:
12260   Exit Sub

ERRH:
12270   THAT_PROC = THIS_PROC
12280   That_Erl = Erl
12290   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12300   Resume EXITP

End Sub

Private Sub description_GotFocus()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "description_GotFocus"

12410   With Me
12420     .description.SelLength = 0
12430     .description.SelStart = 99
12440   End With

EXITP:
12450   Exit Sub

ERRH:
12460   Select Case ERR.Number
        Case Else
12470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12480   End Select
12490   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_display_GotFocus()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_GotFocus"

12510   With Me
12520     .cmbRevenueCodes_Purchase.SetFocus
12530   End With

EXITP:
12540   Exit Sub

ERRH:
12550   Select Case ERR.Number
        Case Else
12560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12570   End Select
12580   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_KeyDown"

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
12690         .cmbRevenueCodes_Purchase.SetFocus
12700       End With
12710     End Select
12720   End If

        ' ** Shift keys.
12730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12750       With Me
12760         intRetVal = 0
12770         .description.SetFocus
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

Private Sub cmbRevenueCodes_Purchase_Enter()

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_Enter"

12910   With Me
12920     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
12930       ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
12940     Case False
12950       If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
12960         ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
12970       End If
12980     End Select
12990   End With

EXITP:
13000   Exit Sub

ERRH:
13010   Select Case ERR.Number
        Case Else
13020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13030   End Select
13040   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_AfterUpdate()

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_AfterUpdate"

13110   With Me
13120     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
13130       .cmbRevenueCodes_Purchase = REVID_INC
13140     Case False
            ' ** Nothing at the moment.
13150     End Select
13160     If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
13170       .cmbRevenueCodes_Purchase_display = Null
13180     Else
13190       .cmbRevenueCodes_Purchase_display = .cmbRevenueCodes_Purchase.Column(1)
13200     End If
13210   End With

EXITP:
13220   Exit Sub

ERRH:
13230   Select Case ERR.Number
        Case Else
13240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13250   End Select
13260   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_KeyDown(KeyCode As Integer, Shift As Integer)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13320   intShiftDown = (Shift And acShiftMask) > 0
13330   intAltDown = (Shift And acAltMask) > 0
13340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13370       With Me
13380         intRetVal = 0
13390         If .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
13400           .cmbTaxCodes_Purchase.SetFocus
13410         ElseIf .cmdOK.Enabled = True Then
13420           .cmdOK.SetFocus
13430         Else
13440           .cmdCancel.SetFocus
13450         End If
13460       End With
13470     End Select
13480   End If

        ' ** Shift keys.
13490   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13500     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13510       With Me
13520         intRetVal = 0
13530         .description.SetFocus
13540       End With
13550     End Select
13560   End If

EXITP:
13570   KeyCode = intRetVal
13580   Exit Sub

ERRH:
13590   intRetVal = 0
13600   Select Case ERR.Number
        Case Else
13610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13620   End Select
13630   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_display_GotFocus()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_GotFocus"

13710   With Me
13720     .cmbTaxCodes_Purchase.SetFocus
13730   End With

EXITP:
13740   Exit Sub

ERRH:
13750   Select Case ERR.Number
        Case Else
13760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13770   End Select
13780   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13820   intShiftDown = (Shift And acShiftMask) > 0
13830   intAltDown = (Shift And acAltMask) > 0
13840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13870       With Me
13880         intRetVal = 0
13890         .cmbTaxCodes_Purchase.SetFocus
13900       End With
13910     End Select
13920   End If

        ' ** Shift keys.
13930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13950       With Me
13960         intRetVal = 0
13970         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
13980           .cmbRevenueCodes_Purchase.SetFocus
13990         Else
14000           .description.SetFocus
14010         End If
14020       End With
14030     End Select
14040   End If

EXITP:
14050   KeyCode = intRetVal
14060   Exit Sub

ERRH:
14070   intRetVal = 0
14080   Select Case ERR.Number
        Case Else
14090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14100   End Select
14110   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_Enter()

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_Enter"

14210   With Me
14220     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
14230       ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
14240     Case False
14250       If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
14260         ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
14270       End If
14280     End Select
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

Private Sub cmbTaxCodes_Purchase_AfterUpdate()

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_AfterUpdate"

14410   With Me
14420     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
14430       .cmbTaxCodes_Purchase = TAXID_INC
14440     Case False
            ' ** Nothing at the moment.
14450     End Select
14460     If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
14470       .cmbTaxCodes_Purchase_display = Null
14480     Else
14490       .cmbTaxCodes_Purchase_display = .cmbTaxCodes_Purchase.Column(1)
14500     End If
14510   End With

EXITP:
14520   Exit Sub

ERRH:
14530   Select Case ERR.Number
        Case Else
14540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14550   End Select
14560   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_KeyDown(KeyCode As Integer, Shift As Integer)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_KeyDown"

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
14690         If .cmdOK.Enabled = True Then
14700           .cmdOK.SetFocus
14710         Else
14720           .cmdCancel.SetFocus
14730         End If
14740       End With
14750     End Select
14760   End If

        ' ** Shift keys.
14770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14790       With Me
14800         intRetVal = 0
14810         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
14820           .cmbRevenueCodes_Purchase.SetFocus
14830         Else
14840           .description.SetFocus
14850         End If
14860       End With
14870     End Select
14880   End If

EXITP:
14890   KeyCode = intRetVal
14900   Exit Sub

ERRH:
14910   intRetVal = 0
14920   Select Case ERR.Number
        Case Else
14930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14940   End Select
14950   Resume EXITP

End Sub

Public Sub CheckCalendar1(blnGoodDate As Boolean, datStartDate As Date)
' ** Called by the Calendar subform.

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckCalendar1"

15010   With Me
15020     Select Case blnGoodDate
          Case True
            ' ** Allow posting up to 1 month into the future.
15030       If datStartDate > DateAdd("m", 1, Date) Then
15040         MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
15050         .transdate = CDate(Format(Date, "mm/dd/yyyy"))
15060       Else
15070         .transdate = datStartDate
15080       End If
15090     Case False
15100       .transdate = CDate(Format(Date, "mm/dd/yyyy"))
15110     End Select
15120     Select Case .cmbAssets.Enabled
          Case True
15130       .cmbAssets.SetFocus
15140     Case False
15150       .shareface.SetFocus
15160     End Select
15170   End With

EXITP:
15180   Exit Sub

ERRH:
15190   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
15200   Case Else
15210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15220   End Select
15230   Resume EXITP

End Sub

Public Sub CheckCalendar2(blnGoodDate As Boolean, datStartDate As Date)
' ** Called by the Calendar subform.

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckCalendar2"

15310   With Me
15320     Select Case blnGoodDate
          Case True
            ' ** Don't allow any future dates.
15330       If Compare_DateA_DateB(datStartDate, ">", Date) = True Then  ' ** Module Function: modStringFuncs.
15340         MsgBox "Future trade dates are not allowed.", vbInformation + vbOKOnly, "Invalid Date"
15350         .assetdate_display = Date
15360         .assetdate = Now()
15370       Else
15380         .assetdate_display = datStartDate
15390         .assetdate = datStartDate + time
15400       End If
15410     Case False
15420       .assetdate_display = Date
15430       .assetdate = Now()
15440     End Select
15450     Select Case .cmbLocations.Enabled
          Case True
15460       .cmbLocations.SetFocus
15470     Case False
15480       .ICash.SetFocus
15490     End Select
15500   End With

EXITP:
15510   Exit Sub

ERRH:
15520   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
15530   Case Else
15540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15550   End Select
15560   Resume EXITP

End Sub

Private Sub NoChar_Load()
' ** Load list of currencies that the Fixedsys font doesn't show.

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

15610   Set dbs = CurrentDb
15620   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
15630     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
15640     Set rst = qdf.OpenRecordset
15650     With rst
15660       .MoveLast
15670       lngNoChars = .RecordCount
15680       .MoveFirst
15690       arr_varNoChar = .GetRows(lngNoChars)
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
15700       .Close
15710     End With
15720     .Close
15730   End With

EXITP:
15740   Set rst = Nothing
15750   Set qdf = Nothing
15760   Set dbs = Nothing
15770   Exit Sub

ERRH:
15780   DoCmd.Hourglass False
15790   Select Case ERR.Number
        Case Else
15800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15810   End Select
15820   Resume EXITP

End Sub
