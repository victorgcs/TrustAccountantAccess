VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReinvest_Dividend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmReinvest_Dividend"

'VGC 04/23/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           E {cmdCancel}
' **   New Location:     L {cmdReinvestDivLocNew}
' **   Inc/Exp Codes:    N {cmbRevenueCodes_Purchase}
' **   OK:               O {cmdOK}
' **   Tax Codes:        T {cmbTaxCodes_Purchase}
' **   New Asset:        W {cmdReinvestDivAssetNew}

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
Private Const CBX_A_ASTNO  As Integer = 2  'assetno
Private Const CBX_A_TYPE   As Integer = 3  'assettype
Private Const CBX_A_TAX    As Integer = 4  'taxcode
Private Const CBX_A_CURRID As Integer = 5  'curr_id

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
                ' ** strCallingForm & "~" & strAccountNo & "~" & datTransDate & "~" & lngCurrID & "~" &
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

570             .cmdReinvestDivAssetNew.Enabled = False
580             .cmdReinvestDivAssetNew.Visible = True
590             .cmdReinvestDivLocNew.Enabled = False
600             .cmdReinvestDivLocNew.Visible = True

610             .accountno.DefaultValue = Chr(34) & strAccountNo & Chr(34)
620             .shortname.DefaultValue = Chr(34) & strShortName & Chr(34)
630             .transdate.DefaultValue = "#" & Format(datTransDate, "mm/dd/yyyy") & "#"
640             .cmbAssets.DefaultValue = lngAssetNo
650             .description.DefaultValue = Chr(34) & "Dividend Reinvestment" & Chr(34)

660             blnCurrID = frm.frmJournal_Sub1_Dividend.Form.dividendCurr_ID.Visible

670             If blnCurrID = True And lngCurrID <> 150& Then  ' ** USD.
680               glngCurrID = lngCurrID
                  ' ** qryJournal_Dividend_07_01 (MasterAsset, without 'HA-', with totdesc),
                  ' ** linked to assettype, by specified GlobalVarGet("glngCurrID").  #curr_id
690               .cmbAssets.RowSource = "qryJournal_Dividend_07_03"
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
880               .reinvestdivCurr_ID.Visible = True
890               .reinvestdivCurr_ID.DefaultValue = lngCurrID
900               .reinvestdivCurr_ID_cmd.Visible = True
910               .reinvestdivCurr_ID_cmd.Enabled = True
920               .icash_usd.Visible = True
930               .pcash_usd.Visible = True
940               .cost_usd.Visible = True
950               NoChar_Load  ' ** Procedure: Below.
960             Else
970               .reinvestdivCurr_ID.Visible = False
980               .reinvestdivCurr_ID_cmd.Enabled = False
990               .reinvestdivCurr_ID_cmd.Visible = False
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

1440      .reinvestdivCurr_ID.Enabled = False
1450      .reinvestdivCurr_ID.Locked = True
1460      .reinvestdivCurr_ID.ForeColor = CLR_DISABLED_FG
1470      .reinvestdivCurr_ID.BackColor = CLR_VLTGRN
1480      reinvestdivCurr_ID_AfterUpdate  ' ** Procedure: Below.

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
1870      varTmp00 = DLookup("[Location_ID]", "qryJournal_Dividend_13")
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
            "WHERE (((masterasset.assetno) = " & CStr(.cmbAssets.Column(CBX_A_ASTNO)) & "));"

2030      Set dbs = CurrentDb
2040      Set rst = dbs.OpenRecordset(strSQL)

2050      If rst![assettype] = "90" Then
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

2200      If .reinvestdivCurr_ID.Visible = True Then
2210        If .reinvestdivCurr_Date.Visible = True Then .reinvestdivCurr_Date.Visible = False
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
2380      reinvestdivCurr_ID_AfterUpdate  ' ** Procedure: Below.

2390      If .reinvestdivCurr_ID.Visible = True Then
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
3150          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3160        Case "shareface"
3170          MsgBox "Please enter a valid number.", vbInformation + vbOKOnly, "Invalid Entry"
3180        End Select
3190        Response = acDataErrContinue
3200      Case 3101
            ' ** Not an item in the table.
3210        DoCmd.CancelEvent
3220      Case 2237
            '** Item not in list.
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
3550    frm.frmJournal_Sub1_Dividend.SetFocus
3560    frm.frmJournal_Sub1_Dividend.Form.dividendAccountNo.SetFocus
3570    frm.TimerInterval = 250&

EXITP:
3580    Set frm = Nothing
3590    Exit Sub

ERRH:
3600    DoCmd.Hourglass False
3610    Select Case ERR.Number
        Case Else
3620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3630    End Select
3640    Resume EXITP

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
4250            If taxcode_type = TAXTYP_DED Then
4260              .cmbTaxCodes_Purchase = .cmbAssets.Column(CBX_A_TAX)  ' ** All AssetType-based Tax Codes are INCOME.
4270            End If
4280          End If
4290        End Select
4300      End If

4310      If blnContinue = True Then

4320        DoCmd.Hourglass True  ' ** Make sure it's still running.
4330        DoEvents

            ' ** This was parent.
4340        lngJrnlID = .dividendreinvestID
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
4460    THAT_PROC = THIS_PROC
4470    That_Erl = Erl
4480    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4490    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

4510    With Me

4520      DoCmd.Hourglass True
4530      DoEvents

          ' ** This was parent.
4540      lngJrnlID = .dividendreinvestID

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
6020      If IsNull(.cmbAssets) = True Then
6030        ExpandCombo .cmbAssets  ' ** Module Procedure: modUtilities.
6040      Else
6050        If IsNull(.cmbAssets.Column(CBX_A_ASTNO)) = True Then
6060          ExpandCombo .cmbAssets  ' ** Module Procedure: modUtilities.
6070        End If
6080      End If
6090    End With

EXITP:
6100    Exit Sub

ERRH:
6110    Select Case ERR.Number
        Case Else
6120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6130    End Select
6140    Resume EXITP

End Sub

Private Sub cmbAssets_AfterUpdate()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_AfterUpdate"

        Dim lngTmp01 As Long

6210    With Me
6220      If IsNull(.cmbAssets) = False Then
6230        If .cmbAssets.Column(CBX_A_ASTNO) > 0& Then
6240          If .cmbAssets.Column(CBX_A_TYPE) = "90" Then
6250            .journaltype = "Liability"
6260            If .Cost > 0 Then
6270              .Cost = .Cost * -1
6280            End If
6290          Else
6300            .journaltype = "Purchase"
6310          End If
6320          If blnCurrID = True Then
6330            lngTmp01 = .cmbAssets.Column(CBX_A_CURRID)
6340            If lngTmp01 <> lngCurrID Then
6350              lngCurrID = lngTmp01
6360              reinvestdivCurr_ID_AfterUpdate  ' ** Procedure: Below.
6370              DoEvents
6380              icash_AfterUpdate  ' ** Procedure: Below.
6390              DoEvents
6400              pcash_AfterUpdate  ' ** Procedure: Below.
6410              DoEvents
6420              cost_AfterUpdate  ' ** Procedure: Below.
6430              DoEvents
6440            End If
6450          End If
6460        End If
6470      End If
6480    End With

EXITP:
6490    Exit Sub

ERRH:
6500    THAT_PROC = THIS_PROC
6510    That_Erl = Erl
6520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6530    Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

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
6690          .shareface.SetFocus
6700        End With
6710      End Select
6720    End If

        ' ** Shift keys.
6730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6750        With Me
6760          intRetVal = 0
6770          .transdate.SetFocus
6780        End With
6790      End Select
6800    End If

        ' ** Ctrl-Shift keys.
6810    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6820      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6830        With Me
6840          intRetVal = 0
6850          .cmdCalendar1.SetFocus
6860        End With
6870      End Select
6880    End If

EXITP:
6890    KeyCode = intRetVal
6900    Exit Sub

ERRH:
6910    intRetVal = 0
6920    THAT_PROC = THIS_PROC
6930    That_Erl = Erl
6940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6950    Resume EXITP

End Sub

Private Sub shareface_Exit(Cancel As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_Exit"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

7010    blnContinue = True

7020    With Me
7030      If IsNumeric(.shareface) = True Then
7040        If .shareface < 0 Then
7050          blnContinue = False
7060          MsgBox "Share/Face must be positive.", vbInformation + vbOKOnly, "Invalid Entry"
7070          Cancel = -1
7080          .shareface.SetFocus
7090        Else
7100          If CDbl(.shareface) > 999999999.9999 Then
7110            blnContinue = False
7120            MsgBox "Share/Face is too large.", vbInformation + vbOKOnly, "Invalid Entry"
7130            .shareface = vbNullString
7140            Cancel = -1
7150            .shareface.SetFocus
7160          End If
7170          If blnContinue = True Then
7180            intPos01 = InStr(1, .shareface, ".", vbTextCompare)
7190            If intPos01 > 0 Then
7200              If Len(.shareface) - intPos01 > gintShareFaceDecimals Then
7210                blnContinue = False
7220                MsgBox "You can only enter a Share/Face with up to " & Trim(str(gintShareFaceDecimals)) & " decimals.", _
                      vbInformation + vbOKOnly, "Invalid Entry"
7230                .shareface = vbNullString
7240                Cancel = -1
7250                .shareface.SetFocus
7260              Else
7270                .shareface.Format = "#,###." & String(gintShareFaceDecimals, "0")
7280              End If
7290            Else
7300              .shareface.Format = "#,###"
7310            End If
7320          End If
7330        End If
7340      Else
7350        blnContinue = False
7360        .shareface = 0
7370        MsgBox "Share/Face must have a value.", vbInformation + vbOKOnly, "Entry Required"
7380        .shareface = vbNullString
7390        DoCmd.RunCommand acCmdSaveRecord
7400        Cancel = -1
7410        .shareface.SetFocus
7420      End If
7430    End With  ' ** Me.

EXITP:
7440    Exit Sub

ERRH:
7450    Select Case ERR.Number
        Case Else
7460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7470    End Select
7480    Resume EXITP

End Sub

Private Sub assetdate_display_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7520    intShiftDown = (Shift And acShiftMask) > 0
7530    intAltDown = (Shift And acAltMask) > 0
7540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7570        With Me
7580          intRetVal = 0
7590          .cmbLocations.SetFocus
7600        End With
7610      End Select
7620    End If

        ' ** Shift keys.
7630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7650        With Me
7660          intRetVal = 0
7670          .shareface.SetFocus
7680        End With
7690      End Select
7700    End If

        ' ** Ctrl keys.
7710    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
7720      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7730        With Me
7740          intRetVal = 0
7750          .cmdCalendar2.SetFocus
7760        End With
7770      End Select
7780    End If

EXITP:
7790    KeyCode = intRetVal
7800    Exit Sub

ERRH:
7810    intRetVal = 0
7820    Select Case ERR.Number
        Case Else
7830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7840    End Select
7850    Resume EXITP

End Sub

Private Sub assetdate_display_Exit(Cancel As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_Exit"

7910    With Me
7920      If DateCheck_Trade(.assetdate_display.text) = True Then  ' ** Module Function: modUtilities.
            ' ** '.Text' needed to reliably check date.
7930        If InStr(CStr(CDbl(CDate(.assetdate_display))), ".") = 0 Then
7940          .assetdate = .assetdate_display + time()
7950        Else
7960          .assetdate = .assetdate_display
7970          .assetdate_display = CDate(Format(.assetdate_display, "mm/dd/yyyy"))
7980        End If
7990      Else
8000        If IsNull(.assetdate_display) = True Then
8010          .assetdate_display = Date
8020          .assetdate = Now()
8030        Else
8040          If Trim(.assetdate_display) = vbNullString Then
8050            .assetdate_display = Date
8060            .assetdate = Now()
8070          Else
8080            Cancel = -1
8090            .assetdate_display.SetFocus
8100          End If
8110        End If
8120      End If
8130    End With

EXITP:
8140    Exit Sub

ERRH:
8150    THAT_PROC = THIS_PROC
8160    That_Erl = Erl
8170    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8180    Resume EXITP

End Sub

Private Sub cmdCalendar2_Enter()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Enter"

8210    With Me
8220      .cmdCalendar2.Form.DateDest = "assetdate_display"
8230      .cmdCalendar2.Form.opgDateType = opgDateType_optDateHour
8240      DoEvents
8250      .cmdCalendar2.Form.cmdCalendar.SetFocus
8260    End With

EXITP:
8270    Exit Sub

ERRH:
8280    Select Case ERR.Number
        Case Else
8290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8300    End Select
8310    Resume EXITP

End Sub

Private Sub cmdCalendar2_Exit(Cancel As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Exit"

        Dim Button As Integer, Shift As Integer, X As Single, Y As Single

8410    With Me
8420      .cmdCalendar2.Form.Detail_MouseMove Button, Shift, X, Y
8430    End With

EXITP:
8440    Exit Sub

ERRH:
8450    Select Case ERR.Number
        Case Else
8460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8470    End Select
8480    Resume EXITP

End Sub

Private Sub cmbLocations_Enter()

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_Enter"

8510    With Me
8520      If IsNull(.cmbLocations) = True Then
8530        .[Location_ID] = 1&
8540      Else
8550        If .cmbLocations = 0& Then
8560          .[Location_ID] = 1&
8570        End If
8580      End If
8590    End With

EXITP:
8600    Exit Sub

ERRH:
8610    THAT_PROC = THIS_PROC
8620    That_Erl = Erl
8630    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8640    Resume EXITP

End Sub

Private Sub cmbLocations_GotFocus()

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_GotFocus"

8710    With Me
8720      If .Location_ID = 1& Then
8730        ExpandCombo .cmbLocations  ' ** Module Procedure: modUtilities.
8740      End If
8750    End With

EXITP:
8760    Exit Sub

ERRH:
8770    THAT_PROC = THIS_PROC
8780    That_Erl = Erl
8790    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8800    Resume EXITP

End Sub

Private Sub cmbLocations_AfterUpdate()

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_AfterUpdate"

8910    With Me
8920      .[Location_ID] = .cmbLocations
8930    End With

EXITP:
8940    Exit Sub

ERRH:
8950    THAT_PROC = THIS_PROC
8960    That_Erl = Erl
8970    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8980    Resume EXITP

End Sub

Private Sub cmbLocations_KeyDown(KeyCode As Integer, Shift As Integer)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_KeyDown"

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
9090          If .ICash.Enabled = True Then
9100            .ICash.SetFocus
9110          ElseIf .Cost.Enabled = True Then
9120            .Cost.SetFocus
9130          Else
9140            .description.SetFocus
9150          End If
9160        End With
9170      End Select
9180    End If

        ' ** Shift keys.
9190    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9200      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9210        With Me
9220          intRetVal = 0
9230          .assetdate_display.SetFocus
9240        End With
9250      End Select
9260    End If

        ' ** Ctrl-Shift keys.
9270    If intCtrlDown And (Not intAltDown) And intShiftDown Then
9280      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9290        With Me
9300          intRetVal = 0
9310          .cmdCalendar2.SetFocus
9320        End With
9330      End Select
9340    End If

EXITP:
9350    KeyCode = intRetVal
9360    Exit Sub

ERRH:
9370    intRetVal = 0
9380    Select Case ERR.Number
        Case Else
9390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9400    End Select
9410    Resume EXITP

End Sub

Private Sub icash_AfterUpdate()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "ICash_AfterUpdate"

        Dim dblICash As Double
        Dim varTmp00 As Variant

9510    With Me
9520      If blnCurrID = True And lngCurrID <> 150& Then
9530        .icash_usd = 0#
9540        If .icash_usd.Visible = False Then .icash_usd.Visible = True
9550        If .pcash_usd.Visible = False Then .pcash_usd.Visible = True
9560        If .cost_usd.Visible = False Then .cost_usd.Visible = True
9570        If IsNull(.ICash) = False Then
9580          If .ICash <> 0 Then
9590            If IsNull(.reinvestdivCurr_ID.Column(CBX_C_CURRID)) = False Then
9600              lngCurrID = .reinvestdivCurr_ID.Column(CBX_C_CURRID)
9610              If lngCurrID <> 150& Then  ' ** USD.
9620                varTmp00 = .reinvestdivCurr_ID.Column(CBX_C_RATE2)
9630                If IsNull(varTmp00) = False Then
9640                  If varTmp00 > 0 Then
9650                    dblICash = .ICash
9660                    dblICash = (dblICash * varTmp00)
9670                    dblICash = Round(dblICash, 2)
9680                    .icash_usd = dblICash
9690                  End If
9700                End If
9710              End If
9720            End If
9730          End If
9740        End If
9750      End If
9760      DoEvents
9770    End With

EXITP:
9780    Exit Sub

ERRH:
9790    Select Case ERR.Number
        Case Else
9800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
9810    End Select
9820    Resume EXITP

End Sub

Private Sub icash_Exit(Cancel As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_Exit"

9910    With Me
9920      Select Case .journaltype
          Case "Purchase"
9930        If .ICash > 0 Then
9940          .ICash = .ICash * -1
9950          If .icash_usd.Visible = True Then
9960            .icash_usd = Abs(.icash_usd) * -1
9970          End If
9980        End If
9990      Case "Liability"
10000       If .ICash < 0 Then
10010         .ICash = Abs(.ICash)
10020         If .icash_usd.Visible = True Then
10030           .icash_usd = Abs(.icash_usd)
10040         End If
10050       End If
10060     End Select
10070   End With

EXITP:
10080   Exit Sub

ERRH:
10090   Select Case ERR.Number
        Case Else
10100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10110   End Select
10120   Resume EXITP

End Sub

Private Sub pcash_AfterUpdate()

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_AfterUpdate"

        Dim dblPCash As Double
        Dim varTmp00 As Variant

10210   With Me
10220     If blnCurrID = True And lngCurrID <> 150& Then
10230       .pcash_usd = 0#
10240       If .icash_usd.Visible = False Then .icash_usd.Visible = True
10250       If .pcash_usd.Visible = False Then .pcash_usd.Visible = True
10260       If .cost_usd.Visible = False Then .cost_usd.Visible = True
10270       If IsNull(.PCash) = False Then
10280         If .PCash <> 0 Then
10290           If IsNull(.reinvestdivCurr_ID.Column(CBX_C_CURRID)) = False Then
10300             lngCurrID = .reinvestdivCurr_ID.Column(CBX_C_CURRID)
10310             If lngCurrID <> 150& Then  ' ** USD.
10320               varTmp00 = .reinvestdivCurr_ID.Column(CBX_C_RATE2)
10330               If IsNull(varTmp00) = False Then
10340                 If varTmp00 > 0 Then
10350                   dblPCash = .PCash
10360                   dblPCash = (dblPCash * varTmp00)
10370                   dblPCash = Round(dblPCash, 2)
10380                   .pcash_usd = dblPCash
10390                 End If
10400               End If
10410             End If
10420           End If
10430         End If
10440       End If
10450     End If
10460     DoEvents
10470   End With

EXITP:
10480   Exit Sub

ERRH:
10490   Select Case ERR.Number
        Case Else
10500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10510   End Select
10520   Resume EXITP

End Sub

Private Sub pcash_Exit(Cancel As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_Exit"

10610   With Me
10620     Select Case .journaltype
          Case "Purchase"
10630       If .PCash > 0 Then
10640         .PCash = .PCash * -1
10650         If .pcash_usd.Visible = True Then
10660           .pcash_usd = Abs(.pcash_usd) * -1
10670         End If
10680       End If
10690     Case "Liability"
10700       If .PCash < 0 Then
10710         .PCash = Abs(.PCash)
10720         If .pcash_usd.Visible = True Then
10730           .pcash_usd = Abs(.pcash_usd)
10740         End If
10750       End If
10760     End Select
10770   End With

EXITP:
10780   Exit Sub

ERRH:
10790   Select Case ERR.Number
        Case Else
10800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10810   End Select
10820   Resume EXITP

End Sub

Private Sub cost_AfterUpdate()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "Cost_AfterUpdate"

        Dim dblCost As Double
        Dim varTmp00 As Variant

10910   With Me
10920     If blnCurrID = True And lngCurrID <> 150& Then
10930       .cost_usd = 0#
10940       If .icash_usd.Visible = False Then .icash_usd.Visible = True
10950       If .pcash_usd.Visible = False Then .pcash_usd.Visible = True
10960       If .cost_usd.Visible = False Then .cost_usd.Visible = True
10970       If IsNull(.Cost) = False Then
10980         If .Cost <> 0 Then
10990           If IsNull(.reinvestdivCurr_ID.Column(CBX_C_CURRID)) = False Then
11000             lngCurrID = .reinvestdivCurr_ID.Column(CBX_C_CURRID)
11010             If lngCurrID <> 150& Then  ' ** USD.
11020               varTmp00 = .reinvestdivCurr_ID.Column(CBX_C_RATE2)
11030               If IsNull(varTmp00) = False Then
11040                 If varTmp00 > 0 Then
11050                   dblCost = .Cost
11060                   dblCost = (dblCost * varTmp00)
11070                   dblCost = Round(dblCost, 2)
11080                   .cost_usd = dblCost
11090                 End If
11100               End If
11110             End If
11120           End If
11130         End If
11140       End If
11150     End If
11160     DoEvents
11170   End With

EXITP:
11180   Exit Sub

ERRH:
11190   Select Case ERR.Number
        Case Else
11200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
11210   End Select
11220   Resume EXITP

End Sub

Private Sub cost_Exit(Cancel As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_Exit"

11310   With Me
11320     Select Case .journaltype
          Case "Purchase"
11330       If .Cost < 0 Then
11340         .Cost = Abs(.Cost)
11350         If .cost_usd.Visible = True Then
11360           .cost_usd = Abs(.cost_usd)
11370         End If
11380       End If
11390     Case "Liability"
11400       If .Cost > 0 Then
11410         .Cost = Abs(.Cost) * -1
11420         If .cost_usd.Visible = True Then
11430           .cost_usd = Abs(.cost_usd) * -1
11440         End If
11450       End If
11460     End Select
11470   End With

EXITP:
11480   Exit Sub

ERRH:
11490   Select Case ERR.Number
        Case Else
11500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11510   End Select
11520   Resume EXITP

End Sub

Private Sub reinvestdivCurr_ID_cmd_DblClick(Cancel As Integer)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "reinvestdivCurr_ID_cmd_DblClick"

11610   With Me
11620     If IsNull(.reinvestdivCurr_Date) = True Then
11630       .reinvestdivCurr_Date = .reinvestdivCurr_ID.Column(CBX_C_DATE)
11640     End If
11650     DoEvents
11660     Select Case .reinvestdivCurr_Date.Visible
          Case True
11670       .reinvestdivCurr_Date.Visible = False
11680     Case False
11690       .reinvestdivCurr_Date.Visible = True
11700     End Select
11710   End With

EXITP:
11720   Exit Sub

ERRH:
11730   Select Case ERR.Number
        Case Else
11740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11750   End Select
11760   Resume EXITP

End Sub

Private Sub reinvestdivCurr_ID_AfterUpdate()

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "reinvestdivCurr_ID_AfterUpdate"

        Dim lngCurrID As Long
        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String, strTmp03 As String, intTmp04 As Integer, lngTmp06 As Long
        Dim lngX As Long

11810   With Me

11820     If IsNull(.reinvestdivCurr_ID) = False Then

11830       lngCurrID = .reinvestdivCurr_ID.Column(CBX_C_CURRID)

11840       blnFound = False
11850       If lngCurrID <> 150& Then
11860         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
11870           NoChar_Load  ' ** Procedure: Below.
11880         End If
11890         For lngX = 0& To (lngNoChars - 1&)
11900           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
11910             blnFound = True
11920           End If
11930         Next
11940       End If

11950       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
11960         .ICash.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
11970         .ICash.FontSize = 10
11980         .ICash.FontBold = True
11990         .PCash.FontName = "Courier New"
12000         .PCash.FontSize = 10
12010         .PCash.FontBold = True
12020         .Cost.FontName = "Courier New"
12030         .Cost.FontSize = 10
12040         .Cost.FontBold = True
12050       Case False
12060         If .ICash.FontName <> "Fixedsys" Then
12070           .ICash.FontName = "Fixedsys"
12080           .ICash.FontSize = 9
12090           .ICash.FontBold = False
12100           .PCash.FontName = "Fixedsys"
12110           .PCash.FontSize = 9
12120           .PCash.FontBold = False
12130           .Cost.FontName = "Fixedsys"
12140           .Cost.FontSize = 9
12150           .Cost.FontBold = False
12160         End If
12170       End Select

12180       If lngCurrID = 150& Then  ' ** USD.
12190         strTmp01 = "Currency"
12200         intTmp04 = 2
12210         lngTmp06 = CLR_WHT
12220       Else
12230         strTmp03 = .reinvestdivCurr_ID.Column(CBX_C_SYM)  ' ** Symbol.
12240         If .reinvestdivCurr_ID.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
12250           intTmp04 = 4
12260         Else
12270           intTmp04 = .reinvestdivCurr_ID.Column(CBX_C_DEC)  ' ** Decimal places.
12280         End If
12290         intLen = Len(strTmp03)
12300         Select Case intLen
              Case 1
                'strTmp03 = strTmp03
12310         Case 2
12320           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2)
12330         Case 3
12340           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3)
12350         Case 4
12360           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & "\" & Mid(strTmp03, 4)
12370         Case 5
12380           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & _
                  "\" & Mid(strTmp03, 4, 1) & "\" & Mid(strTmp03, 5)
12390         End Select
12400         Select Case intTmp04
              Case 0
12410           strTmp01 = strTmp03 & "#,##0;(" & strTmp03 & "#,##0);" & strTmp03 & "0;" & strTmp03 & "0"
12420         Case 1  ' ** None have this currently.
12430           strTmp01 = strTmp03 & "#,##0.0;(" & strTmp03 & "#,##0.0);" & strTmp03 & "0.0;" & strTmp03 & "0.0"
12440         Case 2
12450           strTmp01 = strTmp03 & "#,##0.00;(" & strTmp03 & "#,##0.00);" & strTmp03 & "0.00;" & strTmp03 & "0.00"
12460         Case 3
12470           strTmp01 = strTmp03 & "#,##0.000;(" & strTmp03 & "#,##0.000);" & strTmp03 & "0.000;" & strTmp03 & "0.000"
12480         End Select
12490         lngTmp06 = CLR_VLTGRN
12500       End If
12510       .ICash.Format = strTmp01
12520       .ICash.DecimalPlaces = intTmp04
12530       .ICash.BackColor = lngTmp06
12540       .PCash.Format = strTmp01
12550       .PCash.DecimalPlaces = intTmp04
12560       .PCash.BackColor = lngTmp06
12570       .Cost.Format = strTmp01
12580       .Cost.DecimalPlaces = intTmp04
12590       .Cost.BackColor = lngTmp06
12600     End If

12610   End With

EXITP:
12620   Exit Sub

ERRH:
12630   THAT_PROC = THIS_PROC
12640   That_Erl = Erl
12650   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12660   Resume EXITP

End Sub

Private Sub reinvestdivCurr_ID_NotInList(NewData As String, Response As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "reinvestdivCurr_ID_NotInList"

12710   With Me
12720     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
12730     .reinvestdivCurr_ID.SetFocus
12740     Response = acDataErrContinue
12750   End With

EXITP:
12760   Exit Sub

ERRH:
12770   THAT_PROC = THIS_PROC
12780   That_Erl = Erl
12790   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12800   Resume EXITP

End Sub

Private Sub description_GotFocus()

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "description_GotFocus"

12910   With Me
12920     .description.SelLength = 0
12930     .description.SelStart = 99
12940   End With

EXITP:
12950   Exit Sub

ERRH:
12960   Select Case ERR.Number
        Case Else
12970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12980   End Select
12990   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_display_GotFocus()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_GotFocus"

13010   With Me
13020     .cmbRevenueCodes_Purchase.SetFocus
13030   End With

EXITP:
13040   Exit Sub

ERRH:
13050   Select Case ERR.Number
        Case Else
13060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13070   End Select
13080   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13120   intShiftDown = (Shift And acShiftMask) > 0
13130   intAltDown = (Shift And acAltMask) > 0
13140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13170       With Me
13180         intRetVal = 0
13190         .cmbRevenueCodes_Purchase.SetFocus
13200       End With
13210     End Select
13220   End If

        ' ** Shift keys.
13230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13250       With Me
13260         intRetVal = 0
13270         .description.SetFocus
13280       End With
13290     End Select
13300   End If

EXITP:
13310   KeyCode = intRetVal
13320   Exit Sub

ERRH:
13330   intRetVal = 0
13340   Select Case ERR.Number
        Case Else
13350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13360   End Select
13370   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_Enter()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_Enter"

13410   With Me
13420     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
13430       ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
13440     Case False
13450       If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
13460         ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
13470       End If
13480     End Select
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

Private Sub cmbRevenueCodes_Purchase_AfterUpdate()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_AfterUpdate"

13610   With Me
13620     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
13630       .cmbRevenueCodes_Purchase = REVID_INC
13640     Case False
            ' ** Nothing at the moment.
13650     End Select
13660     If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
13670       .cmbRevenueCodes_Purchase_display = Null
13680     Else
13690       .cmbRevenueCodes_Purchase_display = .cmbRevenueCodes_Purchase.Column(1)
13700     End If
13710   End With

EXITP:
13720   Exit Sub

ERRH:
13730   Select Case ERR.Number
        Case Else
13740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13750   End Select
13760   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_KeyDown(KeyCode As Integer, Shift As Integer)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_KeyDown"

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
13890         If .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
13900           .cmbTaxCodes_Purchase.SetFocus
13910         ElseIf .cmdOK.Enabled = True Then
13920           .cmdOK.SetFocus
13930         Else
13940           .cmdCancel.SetFocus
13950         End If
13960       End With
13970     End Select
13980   End If

        ' ** Shift keys.
13990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14000     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14010       With Me
14020         intRetVal = 0
14030         .description.SetFocus
14040       End With
14050     End Select
14060   End If

EXITP:
14070   KeyCode = intRetVal
14080   Exit Sub

ERRH:
14090   intRetVal = 0
14100   Select Case ERR.Number
        Case Else
14110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14120   End Select
14130   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_display_GotFocus()

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_GotFocus"

14210   With Me
14220     .cmbTaxCodes_Purchase.SetFocus
14230   End With

EXITP:
14240   Exit Sub

ERRH:
14250   Select Case ERR.Number
        Case Else
14260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14270   End Select
14280   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14320   intShiftDown = (Shift And acShiftMask) > 0
14330   intAltDown = (Shift And acAltMask) > 0
14340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14370       With Me
14380         intRetVal = 0
14390         .cmbTaxCodes_Purchase.SetFocus
14400       End With
14410     End Select
14420   End If

        ' ** Shift keys.
14430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14450       With Me
14460         intRetVal = 0
14470         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
14480           .cmbRevenueCodes_Purchase.SetFocus
14490         Else
14500           .description.SetFocus
14510         End If
14520       End With
14530     End Select
14540   End If

EXITP:
14550   KeyCode = intRetVal
14560   Exit Sub

ERRH:
14570   intRetVal = 0
14580   Select Case ERR.Number
        Case Else
14590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14600   End Select
14610   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_Enter()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_Enter"

14710   With Me
14720     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
14730       ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
14740     Case False
14750       If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
14760         ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
14770       End If
14780     End Select
14790   End With

EXITP:
14800   Exit Sub

ERRH:
14810   Select Case ERR.Number
        Case Else
14820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14830   End Select
14840   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_AfterUpdate()

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_AfterUpdate"

14910   With Me
14920     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
14930       .cmbTaxCodes_Purchase = TAXID_INC
14940     Case False
            ' ** Nothing at the moment.
14950     End Select
14960     If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
14970       .cmbTaxCodes_Purchase_display = Null
14980     Else
14990       .cmbTaxCodes_Purchase_display = .cmbTaxCodes_Purchase.Column(1)
15000     End If
15010   End With

EXITP:
15020   Exit Sub

ERRH:
15030   Select Case ERR.Number
        Case Else
15040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15050   End Select
15060   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_KeyDown(KeyCode As Integer, Shift As Integer)

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15120   intShiftDown = (Shift And acShiftMask) > 0
15130   intAltDown = (Shift And acAltMask) > 0
15140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15170       With Me
15180         intRetVal = 0
15190         If .cmdOK.Enabled = True Then
15200           .cmdOK.SetFocus
15210         Else
15220           .cmdCancel.SetFocus
15230         End If
15240       End With
15250     End Select
15260   End If

        ' ** Shift keys.
15270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15280     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15290       With Me
15300         intRetVal = 0
15310         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
15320           .cmbRevenueCodes_Purchase.SetFocus
15330         Else
15340           .description.SetFocus
15350         End If
15360       End With
15370     End Select
15380   End If

EXITP:
15390   KeyCode = intRetVal
15400   Exit Sub

ERRH:
15410   intRetVal = 0
15420   Select Case ERR.Number
        Case Else
15430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15440   End Select
15450   Resume EXITP

End Sub

Public Sub CheckCalendar1(blnGoodDate As Boolean, datStartDate As Date)
' ** Called by the Calendar subform.

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckCalendar1"

15510   With Me
15520     Select Case blnGoodDate
          Case True
            ' ** Allow posting up to 1 month into the future.
15530       If datStartDate > DateAdd("m", 1, Date) Then
15540         MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
15550         .transdate = CDate(Format(Date, "mm/dd/yyyy"))
15560       Else
15570         .transdate = datStartDate
15580       End If
15590     Case False
15600       .transdate = CDate(Format(Date, "mm/dd/yyyy"))
15610     End Select
15620     Select Case .cmbAssets.Enabled
          Case True
15630       .cmbAssets.SetFocus
15640     Case False
15650       .shareface.SetFocus
15660     End Select
15670   End With

EXITP:
15680   Exit Sub

ERRH:
15690   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
15700   Case Else
15710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15720   End Select
15730   Resume EXITP

End Sub

Public Sub CheckCalendar2(blnGoodDate As Boolean, datStartDate As Date)
' ** Called by the Calendar subform.

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckCalendar2"

15810   With Me
15820     Select Case blnGoodDate
          Case True
            ' ** Don't allow any future dates.
15830       If Compare_DateA_DateB(datStartDate, ">", Date) = True Then  ' ** Module Function: modStringFuncs.
15840         MsgBox "Future trade dates are not allowed.", vbInformation + vbOKOnly, "Invalid Date"
15850         .assetdate_display = Date
15860         .assetdate = Now()
15870       Else
15880         .assetdate = datStartDate + time
15890         .assetdate_display = datStartDate
15900       End If
15910     Case False
15920       .assetdate_display = Date
15930       .assetdate = Now()
15940     End Select
15950     Select Case .cmbLocations.Enabled
          Case True
15960       .cmbLocations.SetFocus
15970     Case False
15980       If .ICash.Enabled = True Then
15990         .ICash.SetFocus
16000       ElseIf .Cost.Enabled = True Then
16010         .Cost.SetFocus
16020       Else
16030         .description.SetFocus
16040       End If
16050     End Select
16060   End With

EXITP:
16070   Exit Sub

ERRH:
16080   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
16090   Case Else
16100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16110   End Select
16120   Resume EXITP

End Sub

Private Sub NoChar_Load()
' ** Load list of currencies that the Fixedsys font doesn't show.

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

16210   Set dbs = CurrentDb
16220   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
16230     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
16240     Set rst = qdf.OpenRecordset
16250     With rst
16260       .MoveLast
16270       lngNoChars = .RecordCount
16280       .MoveFirst
16290       arr_varNoChar = .GetRows(lngNoChars)
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
16300       .Close
16310     End With
16320     .Close
16330   End With

EXITP:
16340   Set rst = Nothing
16350   Set qdf = Nothing
16360   Set dbs = Nothing
16370   Exit Sub

ERRH:
16380   DoCmd.Hourglass False
16390   Select Case ERR.Number
        Case Else
16400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16410   End Select
16420   Resume EXITP

End Sub
