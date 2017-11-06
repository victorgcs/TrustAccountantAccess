VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReinvest_Sold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmReinvest_Sold"

'VGC 04/23/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           E {cmdCancel}
' **   New Location:     L {cmdReinvestSaleLocNew}
' **   Inc/Exp Codes:    N {cmbRevenueCodes_Purchase}
' **   OK:               O {cmdOK}
' **   Tax Codes:        T {cmbTaxCodes_Purchase}
' **   New Asset:        W {cmdReinvestSaleAssetNew}

' ** Shortcut Ctrl keys responsive from this form:
' **   Toggle ID:        D {salereinvestID}  (Superuser only)
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
'Private Const CBX_A_DUE    As Integer = 4  'due
'Private Const CBX_A_RATE   As Integer = 5  'rate
'Private Const CBX_A_TAX    As Integer = 6  'taxcode
Private Const CBX_A_CURRID As Integer = 7  'curr_id

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
Private strAccountNo As String, strShortName As String, datTransDate As Date, dblICash As Double, dblPCash As Double
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
150           If CharCnt(strTmp01, "~") = 8 Then
                ' ** strCallingForm & "~" & strAccountNo & "~" & datTransDate & "~" & lngCurrID & "~" &
                ' ** dblICash & "~" & dblPCash & "~" & lngAssetNo & "~" & lngJrnlID & "~" & strShortName

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
330             dblPCash = Val(Left(strTmp01, (intPos01 - 1)))
340             strTmp01 = Mid(strTmp01, (intPos01 + 1))
350             intPos01 = InStr(strTmp01, "~")
360             lngAssetNo = Val(Left(strTmp01, (intPos01 - 1)))
370             strTmp01 = Mid(strTmp01, (intPos01 + 1))
380             intPos01 = InStr(strTmp01, "~")
390             lngJrnlID = Val(Left(strTmp01, (intPos01 - 1)))
400             strShortName = Mid(strTmp01, (intPos01 + 1))

                ' ** Copyright Lebans Holdings 1999 Ltd.
                ' ** Create an instance of the Calendar class.
410             Set clsMonthClass = New clsMonthCal
                ' ** You MUST SET the class hWndForm prop!!!
420             clsMonthClass.hWndForm = Me.hwnd
                ' ** Let's default to PositionAtCursor.
430             clsMonthClass.PositionAtCursor = True

440             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

450             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
460             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

470             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
480             lngMonitorNum = 1&: lngTmp02 = 0&
490             EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
500             If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

                ' ** Variables are fed empty, then populated ByRef.
510             GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFuncs.

520             FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

530             If lngMonitorNum = 1& Then lngTmp02 = (lngFrm_Top + (25& * lngTpp))
540             DoCmd.MoveSize lngFrm_Left - (8& * lngTpp), lngTmp02, lngFrm_Width, lngFrm_Height  'lngFrm_Top + (25& * lngTpp)
550             If lngMonitorNum > 1& Then
560               LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
570             End If

580             Set frm = Forms("frmJournal")
590             frm.Visible = False

600             .cmdReinvestSaleAssetNew.Enabled = False
610             .cmdReinvestSaleAssetNew.Visible = True
620             .cmdReinvestSaleLocNew.Enabled = False
630             .cmdReinvestSaleLocNew.Visible = True

640             .accountno.DefaultValue = Chr(34) & strAccountNo & Chr(34)
650             .shortname.DefaultValue = Chr(34) & strShortName & Chr(34)
660             .transdate.DefaultValue = "#" & Format$(datTransDate, "mm/dd/yyyy") & "#"
670             .description.DefaultValue = Chr(34) & "Sold Reinvestment" & Chr(34)

680             blnCurrID = frm.frmJournal_Sub4_Sold.Form.saleCurr_ID.Visible

690             If blnCurrID = True And lngCurrID <> 150& Then  ' ** USD.
700               glngCurrID = lngCurrID
                  ' ** qryJournal_Dividend_07_01 (MasterAsset, without 'HA-', with totdesc),
                  ' ** linked to assettype, by specified GlobalVarGet("glngCurrID").  #curr_id
710               .cmbAssets.RowSource = "qryJournal_Sale_16_03"
720               .cmbAssets.Requery
730             End If
740             If lngAssetNo <> 0& Then
750               varTmp00 = DLookup("[curr_id]", "masterasset", "[assetno] = " & CStr(lngAssetNo))
760               If varTmp00 <> lngCurrID Then
770                 lngAssetNo = 0&
780               End If
790             End If
800             .cmbAssets.DefaultValue = lngAssetNo  ' ** May be Zero.

810             If dblICash = 0 Then
820               .ICash.DefaultValue = 0
830               .ICash.DefaultValue = (dblICash * -1)
840             End If
850             If dblPCash = 0 Then
860               .PCash.DefaultValue = 0
870             Else
880               .PCash.DefaultValue = (dblPCash * -1)
890             End If
900             .Cost.DefaultValue = (dblICash + dblPCash)
910             .CheckNum.DefaultValue = lngJrnlID  ' ** Borrowing this for the Reinvest/Sold reference.

920             If blnCurrID = True And lngCurrID <> 150& Then  ' ** USD.
930               .reinvestsaleCurr_ID.Visible = True
940               .reinvestsaleCurr_ID.DefaultValue = lngCurrID
950               .reinvestsaleCurr_ID_cmd.Visible = True
960               .reinvestsaleCurr_ID_cmd.Enabled = True
970               .icash_usd.Visible = True
980               .pcash_usd.Visible = True
990               .cost_usd.Visible = True
1000              NoChar_Load  ' ** Procedure: Below.
1010            Else
1020              .reinvestsaleCurr_ID.Visible = False
1030              .reinvestsaleCurr_ID_cmd.Enabled = False
1040              .reinvestsaleCurr_ID_cmd.Visible = False
1050              .icash_usd.Visible = False
1060              .PCash.Left = .pcash_alt_box.Left
1070              .pcash_lbl.Left = .PCash.Left
1080              .pcash_usd.Visible = False
1090              .Cost.Left = .cost_alt_box.Left
1100              .cost_lbl.Left = .Cost.Left
1110              .cost_usd.Visible = False
1120            End If

1130            DoCmd.Hourglass False

1140            .shareface.SetFocus

1150            blnIsOpen = False

1160          Else
1170            Cancel = -1
1180          End If
1190        Else
1200          Cancel = -1
1210        End If
1220      Else
1230        Cancel = -1
1240      End If
1250    End With

1260    If Cancel = -1 Then
1270      Beep
1280      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1290    End If

EXITP:
1300    Exit Sub

ERRH:
1310    DoCmd.Hourglass False
1320    Select Case ERR.Number
        Case Else
1330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1340    End Select
1350    Resume EXITP

End Sub

Private Sub Form_Load()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim varTmp00 As Variant

1410    With Me

          ' ** Do this whether or not Currency is visible.
1420      CLR_DISABLED_FG = CLR_DKGRY
1430      CLR_DISABLED_BG = CLR_LTTEAL

1440      .reinvestsaleCurr_ID.Enabled = False
1450      .reinvestsaleCurr_ID.Locked = True
1460      .reinvestsaleCurr_ID.ForeColor = CLR_DISABLED_FG
1470      .reinvestsaleCurr_ID.BackColor = CLR_VLTGRN
1480      reinvestsaleCurr_ID_AfterUpdate  ' ** Procedure: Below.

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
1870      varTmp00 = DLookup("[Location_ID]", "qryJournal_Sale_21")
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
            "WHERE (((masterasset.assetno) = " & CStr(cmbAssets) & "));"

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
2190      Set rst = Nothing
2200      dbs.Close
2210      Set dbs = Nothing

2220      If .reinvestsaleCurr_ID.Visible = True Then 'If blnCurrID = True And lngCurrID <> 150& Then
2230        If .reinvestsaleCurr_Date.Visible = True Then .reinvestsaleCurr_Date.Visible = False
2240        .icash_usd = Null
2250        .icash_usd.Visible = False
2260        .ICash.Format = "Currency"
2270        .ICash.DecimalPlaces = 2
2280        .ICash.BackColor = CLR_WHT
2290        .pcash_usd = Null
2300        .pcash_usd.Visible = False
2310        .PCash.Format = "Currency"
2320        .PCash.DecimalPlaces = 2
2330        .PCash.BackColor = CLR_WHT
2340        .cost_usd = Null
2350        .cost_usd.Visible = False
2360        .Cost.Format = "Currency"
2370        .Cost.DecimalPlaces = 2
2380        .Cost.BackColor = CLR_WHT
2390      End If

          ' ** Set the currency symbol.
2400      reinvestsaleCurr_ID_AfterUpdate  ' ** Procedure: Below.

2410      If blnCurrID = True And lngCurrID <> 150& Then  ' ** USD.
2420        icash_AfterUpdate  ' ** Procedure: Below.
2430        pcash_AfterUpdate  ' ** Procedure: Below.
2440        cost_AfterUpdate  ' ** Procedure: Below.
2450      End If

2460    End With

EXITP:
2470    Set rst = Nothing
2480    Set dbs = Nothing
2490    Exit Sub

ERRH:
2500    Select Case ERR.Number
        Case Else
2510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2520    End Select
2530    Resume EXITP

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
        ' **   Toggle ID:        D {salereinvestID}  (Superuser only)
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
          Case vbKeyD
2850        With Me
2860          If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
2870            intRetVal = 0
2880            .salereinvestID.Visible = Not .salereinvestID.Visible
2890          End If
2900        End With
2910      Case vbKeyM
2920        With Me
2930          intRetVal = 0
2940          .description.SetFocus
2950        End With
2960      End Select
2970    End If

        ' ** Ctrl-Shift keys.
2980    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2990      Select Case intRetVal
          Case vbKeyF
3000        With Me
3010          intRetVal = 0
3020          .FocusHolder.SetFocus
3030        End With
3040      End Select
3050    End If

EXITP:
3060    KeyCode = intRetVal
3070    Exit Sub

ERRH:
3080    intRetVal = 0
3090    THAT_PROC = THIS_PROC
3100    That_Erl = Erl
3110    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
3120    Resume EXITP

End Sub

Private Sub Form_Timer()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

3210    With Me
3220      .TimerInterval = 0&
3230      If gblnSetFocus = True Then
3240        gblnSetFocus = False
3250        strFormName = vbNullString
3260  On Error Resume Next
3270        Do Until strFormName = THIS_NAME
3280          DoCmd.SelectObject acForm, THIS_NAME, False
3290          strFormName = Screen.ActiveForm.Name
3300        Loop
3310  On Error GoTo ERRH
3320      End If
3330    End With

EXITP:
3340    Exit Sub

ERRH:
3350    DoCmd.Hourglass False
3360    Select Case ERR.Number
        Case Else
3370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3380    End Select
3390    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3410    If DataErr <> 0 Then
3420      Select Case DataErr
          Case 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2110, 13, 2474, 2135, 10503, 10508
            ' ** Do nothing.
3430      Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
3440        Select Case Me.ActiveControl.Name
            Case "assetdate"
3450          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Entry Required"
3460        Case "shareface"
3470          MsgBox "Please enter a valid number.", vbInformation + vbOKOnly, "Entry Required"
3480        End Select
3490        Response = acDataErrContinue
3500      Case 3101
            ' ** Not an item in the table.
3510        DoCmd.CancelEvent
3520      Case 2237
            ' ** Item not in list.
3530        DoCmd.CancelEvent
3540      Case 3162
3550        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
            ' ** Do nothing.
3560      Case 2501  ' ** The '|' action was Canceled.
            ' ** Do nothing.
3570      Case 3315
            ' ** Do nothing.
3580      Case Else
3590        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3600        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3610      End Select
3620    End If

EXITP:
3630    Exit Sub

ERRH:
3640    Select Case ERR.Number
        Case Else
3650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3660    End Select
3670    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3710    DoCmd.Hourglass True  ' ** Make sure it's still running.
3720    DoEvents

3730    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3740    If Not clsMonthClass Is Nothing Then
3750      If clsMonthClass.IsCalendar Then
3760        Cancel = -1
3770      Else
3780        Set clsMonthClass = Nothing
3790      End If
3800    End If

3810    gblnSetFocus = True
3820    gstrReturningForm = THIS_NAME
3830    frm.Visible = True
3840    DoCmd.SelectObject acForm, frm.Name, False
3850    frm.frmJournal_Sub4_Sold.SetFocus
3860    frm.frmJournal_Sub4_Sold.Form.saleAccountNo.SetFocus
3870    frm.TimerInterval = 250&

EXITP:
3880    Exit Sub

ERRH:
3890    DoCmd.Hourglass False
3900    Select Case ERR.Number
        Case Else
3910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3920    End Select
3930    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4010    With Me
4020      .cmdCalendar1.Form.Detail_MouseMove Button, Shift, X, Y
4030      .cmdCalendar2.Form.Detail_MouseMove Button, Shift, X, Y
4040    End With

EXITP:
4050    Exit Sub

ERRH:
4060    Select Case ERR.Number
        Case Else
4070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4080    End Select
4090    Resume EXITP

End Sub

Private Sub cmdOK_Click()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim strRevCode As String, intTaxCode As Integer
        Dim lngJrnlID As Long
        Dim blnContinue As Boolean

4110    blnContinue = True

4120    DoCmd.Hourglass True
4130    DoEvents

4140    With Me

4150      If IsNumeric(.shareface) = False Then
4160        blnContinue = False
4170        DoCmd.Hourglass False
4180        MsgBox "The shareface must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
4190        .shareface.SetFocus
4200      Else
4210        If .shareface = 0 Then
4220          blnContinue = False
4230          DoCmd.Hourglass False
4240          MsgBox "The shareface must be greater than zero.", vbInformation + vbOKOnly, "Invalid Entry"
4250          .shareface.SetFocus
4260        Else
4270          If IsDate(.assetdate) = False Then
4280            blnContinue = False
4290            DoCmd.Hourglass False
4300            MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
4310            .assetdate.SetFocus
4320          Else
4330            If IsNull(.cmbAssets) = True Or .cmbAssets = 0 Or .cmbAssets = vbNullString Then
4340              blnContinue = False
4350              DoCmd.Hourglass False
4360              MsgBox "Please select an asset.", vbInformation + vbOKOnly, "Invalid Entry"
4370              .cmbAssets.SetFocus
4380            End If
4390          End If
4400        End If
4410      End If

4420      If blnContinue = True Then
4430        If IsNull(.cmbRevenueCodes_Purchase) = True Then
4440          .cmbRevenueCodes_Purchase = REVID_INC  ' ** Unspecified Income.
4450          strRevCode = "I"
4460        Else
4470          If .cmbRevenueCodes_Purchase = 0& Then
4480            .cmbRevenueCodes_Purchase = REVID_INC  ' ** Unspecified Income.
4490            strRevCode = "I"
4500          Else
4510            strRevCode = DLookup("[revcode_TYPE_Code]", "qryRevCodeComboIE", "[revcode_ID] = " & CStr(.cmbRevenueCodes_Purchase))
4520          End If
4530        End If
4540        intTaxCode = Nz(.cmbTaxCodes_Purchase, 0)
4550        Select Case .journaltype
            Case "Purchase"
              ' ** INCOME.
4560          If strRevCode <> "I" Then
4570            .cmbRevenueCodes_Purchase = REVID_INC  ' ** Unspecified Income.
4580          End If
4590          If IsNull(.cmbTaxCodes_Purchase) = True Then
4600            .cmbTaxCodes_Purchase = TAXID_INC
4610          Else
4620            If .taxcode_type = TAXTYP_DED Then
4630              .cmbTaxCodes_Purchase = .cmbAssets.Column(4)  ' ** All AssetType-based Tax Codes are INCOME.
4640            End If
4650          End If
4660        End Select
4670      End If

4680      If blnContinue = True Then

4690        DoCmd.Hourglass True  ' ** Make sure it's still running.
4700        DoEvents

            ' ** This was parent.
4710        lngJrnlID = .salereinvestID
4720        gdblCrtRpt_CostTot = lngJrnlID  ' ** Borrowing this variable from the Court Reports.

4730        .priceperunit = Abs(.Cost / .shareface)

4740        If InStr(CStr(CDbl(CDate(.assetdate))), ".") = 0 Then
4750          .assetdate = .assetdate + time()
4760        End If
4770        DoCmd.RunCommand acCmdSaveRecord

4780        Forms(strCallingForm).frmJournal_Sub4_Sold.Form.cmdSaleLotInfo.Enabled = False

4790        DoCmd.Close acForm, THIS_NAME

4800      End If  ' ** blnContinue.

4810    End With  ' ** Me.

EXITP:
4820    Exit Sub

ERRH:
4830    DoCmd.Hourglass False
4840    Select Case ERR.Number
        Case Else
4850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4860    End Select
4870    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim lngJrnlID As Long

4910    With Me

4920      DoCmd.Hourglass True
4930      DoEvents

4940      lngJrnlID = .salereinvestID

4950      gblnMessage = False
4960      gdblCrtRpt_CostTot = lngJrnlID  ' ** Borrowing this variable from the Court Reports.

4970      DoCmd.Close acForm, THIS_NAME

4980    End With

EXITP:
4990    Exit Sub

ERRH:
5000    DoCmd.Hourglass False
5010    THAT_PROC = THIS_PROC
5020    That_Erl = Erl
5030    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5040    Resume EXITP

End Sub

Private Sub transdate_BeforeUpdate(Cancel As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_BeforeUpdate"

5110    With Me
5120      If .transdate.text <> "__/__/____" And .transdate.text <> vbNullString Then
5130        If Left(.transdate.text, 2) > 12 Then
5140          MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Entry"
5150          DoCmd.CancelEvent
5160        End If
5170      End If
5180    End With

EXITP:
5190    Exit Sub

ERRH:
5200    THAT_PROC = THIS_PROC
5210    That_Erl = Erl
5220    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5230    Resume EXITP

End Sub

Private Sub transdate_KeyDown(KeyCode As Integer, Shift As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5320    intShiftDown = (Shift And acShiftMask) > 0
5330    intAltDown = (Shift And acAltMask) > 0
5340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5370        With Me
5380          intRetVal = 0
5390          .cmbAssets.SetFocus
5400        End With
5410      End Select
5420    End If

        ' ** Shift keys.
5430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5450        With Me
5460          intRetVal = 0
5470          If .journaltype.Enabled = True Then
5480            .journaltype.SetFocus
5490          ElseIf .accountno.Enabled = True Then
5500            .accountno.SetFocus
5510          ElseIf .cmbTaxCodes_Purchase_display.Visible = True And .cmbTaxCodes_Purchase_display.Enabled = True Then
5520            .cmbTaxCodes_Purchase_display.SetFocus
5530          ElseIf .cmbRevenueCodes_Purchase_display.Visible = True And .cmbRevenueCodes_Purchase_display.Enabled = True Then
5540            .cmbRevenueCodes_Purchase_display.SetFocus
5550          Else
5560            .description.SetFocus
5570          End If
5580        End With
5590      End Select
5600    End If

        ' ** Ctrl keys.
5610    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
5620      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5630        With Me
5640          intRetVal = 0
5650          .cmdCalendar1.SetFocus
5660        End With
5670      End Select
5680    End If

EXITP:
5690    KeyCode = intRetVal
5700    Exit Sub

ERRH:
5710    intRetVal = 0
5720    THAT_PROC = THIS_PROC
5730    That_Erl = Erl
5740    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5750    Resume EXITP

End Sub

Private Sub transdate_Exit(Cancel As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_Exit"

5810    With Me
5820      If DateCheck_Post(.transdate.text) = True Then  ' ** Module Function: modUtilities.
            ' ** '.Text' needed to reliably check date.
5830      Else
5840        If IsNull(.transdate) = True Then
5850          .transdate = Date
5860        Else
5870          If Trim(.transdate) = vbNullString Then
5880            .transdate = Date
5890          Else
5900            Cancel = -1
5910            .transdate.SetFocus
5920          End If
5930        End If
5940      End If
5950    End With

EXITP:
5960    Exit Sub

ERRH:
5970    Select Case ERR.Number
        Case Else
5980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5990    End Select
6000    Resume EXITP

End Sub

Private Sub cmdCalendar1_Enter()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Enter"

6110    With Me
6120      .cmdCalendar1.Form.DateDest = "transdate"
6130      .cmdCalendar1.Form.opgDateType = opgDateType_optDateOnly
6140      DoEvents
6150      .cmdCalendar1.Form.cmdCalendar.SetFocus
6160    End With

EXITP:
6170    Exit Sub

ERRH:
6180    Select Case ERR.Number
        Case Else
6190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6200    End Select
6210    Resume EXITP

End Sub

Private Sub cmdCalendar1_Exit(Cancel As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Exit"

        Dim Button As Integer, Shift As Integer, X As Single, Y As Single

6310    With Me
6320      .cmdCalendar1.Form.Detail_MouseMove Button, Shift, X, Y
6330    End With

EXITP:
6340    Exit Sub

ERRH:
6350    Select Case ERR.Number
        Case Else
6360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6370    End Select
6380    Resume EXITP

End Sub

Private Sub cmbAssets_Enter()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_Enter"

6410    With Me
6420      If IsNull(.cmbAssets) = True Then
6430        ExpandCombo .cmbAssets  ' ** Module Procedure: modUtilities.
6440      Else
6450        If IsNull(.cmbAssets.Column(CBX_A_ASTNO)) = True Then
6460          ExpandCombo .cmbAssets  ' ** Module Procedure: modUtilities.
6470        End If
6480      End If
6490    End With

EXITP:
6500    Exit Sub

ERRH:
6510    Select Case ERR.Number
        Case Else
6520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6530    End Select
6540    Resume EXITP

End Sub

Private Sub cmbAssets_AfterUpdate()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_AfterUpdate"

        Dim lngTmp01 As Long

6610    With Me
6620      If IsNull(.cmbAssets) = False Then
6630        If .cmbAssets.Column(CBX_A_ASTNO) > 0& Then
6640          If .cmbAssets.Column(CBX_A_TYPE) = "90" Then
6650            .journaltype = "Liability"
6660            If .Cost > 0 Then
6670              .Cost = .Cost * -1
6680            End If
6690          Else
6700            .journaltype = "Purchase"
6710          End If
6720          If blnCurrID = True Then
6730            lngTmp01 = .cmbAssets.Column(CBX_A_CURRID)
6740            If lngTmp01 <> lngCurrID Then
6750              lngCurrID = lngTmp01
6760              reinvestsaleCurr_ID_AfterUpdate  ' ** Procedure: Below.
6770              DoEvents
6780              icash_AfterUpdate  ' ** Procedure: Below.
6790              DoEvents
6800              pcash_AfterUpdate  ' ** Procedure: Below.
6810              DoEvents
6820              cost_AfterUpdate  ' ** Procedure: Below.
6830              DoEvents
6840            End If
6850          End If
6860        End If
6870      End If
6880    End With

EXITP:
6890    Exit Sub

ERRH:
6900    Select Case ERR.Number
        Case Else
6910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6920    End Select
6930    Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

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
7090          .shareface.SetFocus
7100        End With
7110      End Select
7120    End If

        ' ** Shift keys.
7130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
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

Private Sub shareface_Exit(Cancel As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_Exit"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

7410    blnContinue = True

7420    With Me
7430      If IsNumeric(.shareface) = True Then
7440        If .shareface < 0 Then
7450          blnContinue = False
7460          MsgBox "Share/Face must be positive.", vbInformation + vbOKOnly, "Invalid Entry"
7470          DoCmd.CancelEvent
7480          .shareface.SetFocus
7490        Else
7500          If CDbl(.shareface) > 999999999.9999 Then
7510            blnContinue = False
7520            MsgBox "Share/Face is too large.", vbInformation + vbOKOnly, "Invalid Entry"
7530            .shareface = vbNullString
7540            DoCmd.CancelEvent
7550            .shareface.SetFocus
7560          End If
7570          If blnContinue = True Then
7580            intPos01 = InStr(1, .shareface, ".", vbTextCompare)
7590            If intPos01 > 0 Then
7600              If Len(.shareface) - intPos01 > gintShareFaceDecimals Then
7610                blnContinue = False
7620                MsgBox "You can only enter a Share/Face with up to " & Trim(str(gintShareFaceDecimals)) & " decimals.", _
                      vbInformation + vbOKOnly, "Invalid Entry"
7630                .shareface = vbNullString
7640                DoCmd.CancelEvent
7650              Else
7660                .shareface.Format = "#,###." & String(gintShareFaceDecimals, "0")
7670              End If
7680            Else
7690              .shareface.Format = "#,###"
7700            End If
7710          End If
7720        End If
7730      Else
7740        blnContinue = False
7750        .shareface = 0
7760        MsgBox "Share/Face must have a value.", vbInformation + vbOKOnly, "Entry Required"
7770        .shareface = vbNullString
7780        DoCmd.RunCommand acCmdSaveRecord
7790        DoCmd.CancelEvent
7800        .shareface.SetFocus
7810      End If
7820    End With  ' ** Me.

EXITP:
7830    Exit Sub

ERRH:
7840    Select Case ERR.Number
        Case Else
7850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7860    End Select
7870    Resume EXITP

End Sub

Private Sub assetdate_display_KeyDown(KeyCode As Integer, Shift As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
7970        With Me
7980          intRetVal = 0
7990          .cmbLocations.SetFocus
8000        End With
8010      End Select
8020    End If

        ' ** Shift keys.
8030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8050        With Me
8060          intRetVal = 0
8070          .shareface.SetFocus
8080        End With
8090      End Select
8100    End If

        ' ** Ctrl keys.
8110    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8120      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8130        With Me
8140          intRetVal = 0
8150          .cmdCalendar2.SetFocus
8160        End With
8170      End Select
8180    End If

EXITP:
8190    KeyCode = intRetVal
8200    Exit Sub

ERRH:
8210    intRetVal = 0
8220    Select Case ERR.Number
        Case Else
8230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8240    End Select
8250    Resume EXITP

End Sub

Private Sub assetdate_display_Exit(Cancel As Integer)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_Exit"

8310    With Me
8320      If DateCheck_Trade(.assetdate_display.text) = True Then  ' ** Module Function: modUtilities.
            ' ** '.Text' needed to reliably check date.
8330        If InStr(CStr(CDbl(CDate(.assetdate_display))), ".") = 0 Then
8340          .assetdate = .assetdate_display + time()
8350        Else
8360          .assetdate = .assetdate_display
8370          .assetdate_display = CDate(Format(.assetdate_display, "mm/dd/yyyy"))
8380        End If
8390      Else
8400        If IsNull(.assetdate_display) = True Then
8410          .assetdate_display = Date
8420          .assetdate = Now()
8430        Else
8440          If Trim(.assetdate_display) = vbNullString Then
8450            .assetdate_display = Date
8460            .assetdate = Now()
8470          Else
8480            Cancel = -1
8490            .assetdate_display.SetFocus
8500          End If
8510        End If
8520      End If
8530    End With

EXITP:
8540    Exit Sub

ERRH:
8550    THAT_PROC = THIS_PROC
8560    That_Erl = Erl
8570    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8580    Resume EXITP

End Sub

Private Sub cmdCalendar2_Enter()

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Enter"

8610    With Me
8620      .cmdCalendar2.Form.DateDest = "assetdate_display"
8630      .cmdCalendar2.Form.opgDateType = opgDateType_optDateHour
8640      DoEvents
8650      .cmdCalendar2.Form.cmdCalendar.SetFocus
8660    End With

EXITP:
8670    Exit Sub

ERRH:
8680    Select Case ERR.Number
        Case Else
8690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8700    End Select
8710    Resume EXITP

End Sub

Private Sub cmdCalendar2_Exit(Cancel As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Exit"

        Dim Button As Integer, Shift As Integer, X As Single, Y As Single

8810    With Me
8820      .cmdCalendar2.Form.Detail_MouseMove Button, Shift, X, Y
8830    End With

EXITP:
8840    Exit Sub

ERRH:
8850    Select Case ERR.Number
        Case Else
8860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8870    End Select
8880    Resume EXITP

End Sub

Private Sub cmbLocations_Enter()

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_Enter"

8910    With Me
8920      If IsNull(.cmbLocations) = True Then
8930        .[Location_ID] = 1&
8940      Else
8950        If .cmbLocations = 0& Then
8960          .[Location_ID] = 1&
8970        End If
8980      End If
8990    End With

EXITP:
9000    Exit Sub

ERRH:
9010    THAT_PROC = THIS_PROC
9020    That_Erl = Erl
9030    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9040    Resume EXITP

End Sub

Private Sub cmbLocations_GotFocus()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_GotFocus"

9110    With Me
9120      If .[Location_ID] = 1& Then
9130        ExpandCombo .[Location_ID]  ' ** Module Procedure: modUtilities.
9140      End If
9150    End With

EXITP:
9160    Exit Sub

ERRH:
9170    THAT_PROC = THIS_PROC
9180    That_Erl = Erl
9190    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9200    Resume EXITP

End Sub

Private Sub cmbLocations_AfterUpdate()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_AfterUpdate"

9310    With Me
9320      .[Location_ID] = .cmbLocations
9330    End With

EXITP:
9340    Exit Sub

ERRH:
9350    THAT_PROC = THIS_PROC
9360    That_Erl = Erl
9370    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9380    Resume EXITP

End Sub

Private Sub cmbLocations_KeyDown(KeyCode As Integer, Shift As Integer)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_KeyDown"

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
9490          .ICash.SetFocus
9500        End With
9510      End Select
9520    End If

        ' ** Shift keys.
9530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9550        With Me
9560          intRetVal = 0
9570          .assetdate_display.SetFocus
9580        End With
9590      End Select
9600    End If

        ' ** Ctrl-Shift keys.
9610    If intCtrlDown And (Not intAltDown) And intShiftDown Then
9620      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9630        With Me
9640          intRetVal = 0
9650          .cmdCalendar2.SetFocus
9660        End With
9670      End Select
9680    End If

EXITP:
9690    KeyCode = intRetVal
9700    Exit Sub

ERRH:
9710    intRetVal = 0
9720    Select Case ERR.Number
        Case Else
9730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9740    End Select
9750    Resume EXITP

End Sub

Private Sub icash_AfterUpdate()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "ICash_AfterUpdate"

        Dim dblICash As Double
        Dim varTmp00 As Variant

9810    With Me
9820      If blnCurrID = True And lngCurrID <> 150& Then
9830        .icash_usd = 0#
9840        If .icash_usd.Visible = False Then .icash_usd.Visible = True
9850        If .pcash_usd.Visible = False Then .pcash_usd.Visible = True
9860        If .cost_usd.Visible = False Then .cost_usd.Visible = True
9870        If IsNull(.ICash) = False Then
9880          If .ICash <> 0 Then
9890            If IsNull(.reinvestsaleCurr_ID.Column(CBX_C_CURRID)) = False Then
9900              lngCurrID = .reinvestsaleCurr_ID.Column(CBX_C_CURRID)
9910              If lngCurrID <> 150& Then  ' ** USD.
9920                varTmp00 = .reinvestsaleCurr_ID.Column(CBX_C_RATE2)
9930                If IsNull(varTmp00) = False Then
9940                  If varTmp00 > 0 Then
9950                    dblICash = .ICash
9960                    dblICash = (dblICash * varTmp00)
9970                    dblICash = Round(dblICash, 2)
9980                    .icash_usd = dblICash
9990                  End If
10000               End If
10010             End If
10020           End If
10030         End If
10040       End If
10050     End If
10060     DoEvents
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

Private Sub icash_Exit(Cancel As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_Exit"

10210   With Me
10220     Select Case .journaltype
          Case "Purchase"
10230       If .ICash > 0 Then
10240         .ICash = .ICash * -1
10250         If .icash_usd.Visible = True Then
10260           .icash_usd = Abs(.icash_usd) * -1
10270         End If
10280       End If
10290     Case "Liability"
10300       If .ICash < 0 Then
10310         .ICash = Abs(.ICash)
10320         If .icash_usd.Visible = True Then
10330           .icash_usd = Abs(.icash_usd)
10340         End If
10350       End If
10360     End Select
10370   End With

EXITP:
10380   Exit Sub

ERRH:
10390   Select Case ERR.Number
        Case Else
10400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10410   End Select
10420   Resume EXITP

End Sub

Private Sub pcash_AfterUpdate()

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_AfterUpdate"

        Dim dblPCash As Double
        Dim varTmp00 As Variant

10510   With Me
10520     If blnCurrID = True And lngCurrID <> 150& Then
10530       .pcash_usd = 0#
10540       If .icash_usd.Visible = False Then .icash_usd.Visible = True
10550       If .pcash_usd.Visible = False Then .pcash_usd.Visible = True
10560       If .cost_usd.Visible = False Then .cost_usd.Visible = True
10570       If IsNull(.PCash) = False Then
10580         If .PCash <> 0 Then
10590           If IsNull(.reinvestsaleCurr_ID.Column(CBX_C_CURRID)) = False Then
10600             lngCurrID = .reinvestsaleCurr_ID.Column(CBX_C_CURRID)
10610             If lngCurrID <> 150& Then  ' ** USD.
10620               varTmp00 = .reinvestsaleCurr_ID.Column(CBX_C_RATE2)
10630               If IsNull(varTmp00) = False Then
10640                 If varTmp00 > 0 Then
10650                   dblPCash = .PCash
10660                   dblPCash = (dblPCash * varTmp00)
10670                   dblPCash = Round(dblPCash, 2)
10680                   .pcash_usd = dblPCash
10690                 End If
10700               End If
10710             End If
10720           End If
10730         End If
10740       End If
10750     End If
10760     DoEvents
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

Private Sub pcash_Exit(Cancel As Integer)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_Exit"

10910   With Me
10920     Select Case .journaltype
          Case "Purchase"
10930       If .PCash > 0 Then
10940         .PCash = .PCash * -1
10950         If .pcash_usd.Visible = True Then
10960           .pcash_usd = Abs(.pcash_usd) * -1
10970         End If
10980       End If
10990     Case "Liability"
11000       If .PCash < 0 Then
11010         .PCash = Abs(.PCash)
11020         If .pcash_usd.Visible = True Then
11030           .pcash_usd = Abs(.pcash_usd)
11040         End If
11050       End If
11060     End Select
11070   End With

EXITP:
11080   Exit Sub

ERRH:
11090   Select Case ERR.Number
        Case Else
11100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11110   End Select
11120   Resume EXITP

End Sub

Private Sub cost_AfterUpdate()

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "Cost_AfterUpdate"

        Dim dblCost As Double
        Dim varTmp00 As Variant

11210   With Me
11220     If blnCurrID = True And lngCurrID <> 150& Then
11230       .cost_usd = 0#
11240       If .icash_usd.Visible = False Then .icash_usd.Visible = True
11250       If .pcash_usd.Visible = False Then .pcash_usd.Visible = True
11260       If .cost_usd.Visible = False Then .cost_usd.Visible = True
11270       If IsNull(.Cost) = False Then
11280         If .Cost <> 0 Then
11290           If IsNull(.reinvestsaleCurr_ID.Column(CBX_C_CURRID)) = False Then
11300             lngCurrID = .reinvestsaleCurr_ID.Column(CBX_C_CURRID)
11310             If lngCurrID <> 150& Then  ' ** USD.
11320               varTmp00 = .reinvestsaleCurr_ID.Column(CBX_C_RATE2)
11330               If IsNull(varTmp00) = False Then
11340                 If varTmp00 > 0 Then
11350                   dblCost = .Cost
11360                   dblCost = (dblCost * varTmp00)
11370                   dblCost = Round(dblCost, 2)
11380                   .cost_usd = dblCost
11390                 End If
11400               End If
11410             End If
11420           End If
11430         End If
11440       End If
11450     End If
11460     DoEvents
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

Private Sub cost_Exit(Cancel As Integer)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_Exit"

11610   With Me
11620     Select Case .journaltype
          Case "Purchase"
11630       If .Cost < 0 Then
11640         .Cost = Abs(.Cost)
11650         If .cost_usd.Visible = True Then
11660           .cost_usd = Abs(.cost_usd)
11670         End If
11680       End If
11690     Case "Liability"
11700       If .Cost > 0 Then
11710         .Cost = Abs(.Cost) * -1
11720         If .cost_usd.Visible = True Then
11730           .cost_usd = Abs(.cost_usd) * -1
11740         End If
11750       End If
11760     End Select
11770   End With

EXITP:
11780   Exit Sub

ERRH:
11790   Select Case ERR.Number
        Case Else
11800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11810   End Select
11820   Resume EXITP

End Sub

Private Sub reinvestsaleCurr_ID_cmd_DblClick(Cancel As Integer)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "reinvestsaleCurr_ID_cmd_DblClick"

11910   With Me
11920     If IsNull(.reinvestsaleCurr_Date) = True Then
11930       .reinvestsaleCurr_Date = .reinvestsaleCurr_ID.Column(CBX_C_DATE)
11940     End If
11950     DoEvents
11960     Select Case .reinvestsaleCurr_Date.Visible
          Case True
11970       .reinvestsaleCurr_Date.Visible = False
11980     Case False
11990       .reinvestsaleCurr_Date.Visible = True
12000     End Select
12010   End With

EXITP:
12020   Exit Sub

ERRH:
12030   Select Case ERR.Number
        Case Else
12040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12050   End Select
12060   Resume EXITP

End Sub

Private Sub reinvestsaleCurr_ID_AfterUpdate()

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "reinvestsaleCurr_ID_AfterUpdate"

        Dim lngCurrID As Long
        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String, strTmp03 As String, intTmp04 As Integer, lngTmp06 As Long
        Dim lngX As Long

12110   With Me

12120     If IsNull(.reinvestsaleCurr_ID) = False Then

12130       lngCurrID = .reinvestsaleCurr_ID.Column(CBX_C_CURRID)

12140       blnFound = False
12150       If lngCurrID <> 150& Then
12160         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
12170           NoChar_Load  ' ** Procedure: Below.
12180         End If
12190         For lngX = 0& To (lngNoChars - 1&)
12200           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
12210             blnFound = True
12220           End If
12230         Next
12240       End If

12250       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
12260         .ICash.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
12270         .ICash.FontSize = 10
12280         .ICash.FontBold = True
12290         .PCash.FontName = "Courier New"
12300         .PCash.FontSize = 10
12310         .PCash.FontBold = True
12320         .Cost.FontName = "Courier New"
12330         .Cost.FontSize = 10
12340         .Cost.FontBold = True
12350       Case False
12360         If .ICash.FontName <> "Fixedsys" Then
12370           .ICash.FontName = "Fixedsys"
12380           .ICash.FontSize = 9
12390           .ICash.FontBold = False
12400           .PCash.FontName = "Fixedsys"
12410           .PCash.FontSize = 9
12420           .PCash.FontBold = False
12430           .Cost.FontName = "Fixedsys"
12440           .Cost.FontSize = 9
12450           .Cost.FontBold = False
12460         End If
12470       End Select

12480       If lngCurrID = 150& Then  ' ** USD.
12490         strTmp01 = "Currency"
12500         intTmp04 = 2
12510         lngTmp06 = CLR_WHT
12520       Else
12530         strTmp03 = .reinvestsaleCurr_ID.Column(CBX_C_SYM)  ' ** Symbol.
12540         If .reinvestsaleCurr_ID.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
12550           intTmp04 = 4
12560         Else
12570           intTmp04 = .reinvestsaleCurr_ID.Column(CBX_C_DEC)  ' ** Decimal places.
12580         End If
12590         intLen = Len(strTmp03)
12600         Select Case intLen
              Case 1
                'strTmp03 = strTmp03
12610         Case 2
12620           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2)
12630         Case 3
12640           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3)
12650         Case 4
12660           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & "\" & Mid(strTmp03, 4)
12670         Case 5
12680           strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & _
                  "\" & Mid(strTmp03, 4, 1) & "\" & Mid(strTmp03, 5)
12690         End Select
12700         Select Case intTmp04
              Case 0
12710           strTmp01 = strTmp03 & "#,##0;(" & strTmp03 & "#,##0);" & strTmp03 & "0;" & strTmp03 & "0"
12720         Case 1  ' ** None have this currently.
12730           strTmp01 = strTmp03 & "#,##0.0;(" & strTmp03 & "#,##0.0);" & strTmp03 & "0.0;" & strTmp03 & "0.0"
12740         Case 2
12750           strTmp01 = strTmp03 & "#,##0.00;(" & strTmp03 & "#,##0.00);" & strTmp03 & "0.00;" & strTmp03 & "0.00"
12760         Case 3
12770           strTmp01 = strTmp03 & "#,##0.000;(" & strTmp03 & "#,##0.000);" & strTmp03 & "0.000;" & strTmp03 & "0.000"
12780         End Select
12790         lngTmp06 = CLR_VLTGRN
12800       End If
12810       .ICash.Format = strTmp01
12820       .ICash.DecimalPlaces = intTmp04
12830       .ICash.BackColor = lngTmp06
12840       .PCash.Format = strTmp01
12850       .PCash.DecimalPlaces = intTmp04
12860       .PCash.BackColor = lngTmp06
12870       .Cost.Format = strTmp01
12880       .Cost.DecimalPlaces = intTmp04
12890       .Cost.BackColor = lngTmp06
12900     End If

12910   End With

EXITP:
12920   Exit Sub

ERRH:
12930   THAT_PROC = THIS_PROC
12940   That_Erl = Erl
12950   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12960   Resume EXITP

End Sub

Private Sub reinvestsaleCurr_ID_NotInList(NewData As String, Response As Integer)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "reinvestsaleCurr_ID_NotInList"

13010   With Me
13020     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
13030     .reinvestsaleCurr_ID.SetFocus
13040     Response = acDataErrContinue
13050   End With

EXITP:
13060   Exit Sub

ERRH:
13070   THAT_PROC = THIS_PROC
13080   That_Erl = Erl
13090   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13100   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_display_GotFocus()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_GotFocus"

13210   With Me
13220     .cmbRevenueCodes_Purchase.SetFocus
13230   End With

EXITP:
13240   Exit Sub

ERRH:
13250   Select Case ERR.Number
        Case Else
13260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13270   End Select
13280   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_KeyDown"

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
13390         .cmbRevenueCodes_Purchase.SetFocus
13400       End With
13410     End Select
13420   End If

        ' ** Shift keys.
13430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13450       With Me
13460         intRetVal = 0
13470         .description.SetFocus
13480       End With
13490     End Select
13500   End If

EXITP:
13510   KeyCode = intRetVal
13520   Exit Sub

ERRH:
13530   intRetVal = 0
13540   Select Case ERR.Number
        Case Else
13550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13560   End Select
13570   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_Enter()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_Enter"

13610   With Me
13620     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
13630       ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
13640     Case False
13650       If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
13660         ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
13670       End If
13680     End Select
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

Private Sub cmbRevenueCodes_Purchase_AfterUpdate()

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_AfterUpdate"

13810   With Me
13820     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
13830       .cmbRevenueCodes_Purchase = REVID_INC
13840     Case False
            ' ** Nothing at the moment.
13850     End Select
13860     If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
13870       .cmbRevenueCodes_Purchase_display = Null
13880     Else
13890       .cmbRevenueCodes_Purchase_display = .cmbRevenueCodes_Purchase.Column(1)
13900     End If
13910   End With

EXITP:
13920   Exit Sub

ERRH:
13930   Select Case ERR.Number
        Case Else
13940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13950   End Select
13960   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_display_GotFocus()

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_GotFocus"

14010   With Me
14020     .cmbTaxCodes_Purchase.SetFocus
14030   End With

EXITP:
14040   Exit Sub

ERRH:
14050   Select Case ERR.Number
        Case Else
14060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14070   End Select
14080   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14120   intShiftDown = (Shift And acShiftMask) > 0
14130   intAltDown = (Shift And acAltMask) > 0
14140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14170       With Me
14180         intRetVal = 0
14190         .cmbTaxCodes_Purchase.SetFocus
14200       End With
14210     End Select
14220   End If

        ' ** Shift keys.
14230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14250       With Me
14260         intRetVal = 0
14270         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
14280           .cmbRevenueCodes_Purchase.SetFocus
14290         Else
14300           .description.SetFocus
14310         End If
14320       End With
14330     End Select
14340   End If

EXITP:
14350   KeyCode = intRetVal
14360   Exit Sub

ERRH:
14370   intRetVal = 0
14380   Select Case ERR.Number
        Case Else
14390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14400   End Select
14410   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_Enter()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_Enter"

14510   With Me
14520     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
14530       ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
14540     Case False
14550       If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
14560         ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
14570       End If
14580     End Select
14590   End With

EXITP:
14600   Exit Sub

ERRH:
14610   Select Case ERR.Number
        Case Else
14620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14630   End Select
14640   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_AfterUpdate()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_AfterUpdate"

14710   With Me
14720     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
14730       .cmbTaxCodes_Purchase = TAXID_INC
14740     Case False
            ' ** Nothing at the moment.
14750     End Select
14760     If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
14770       .cmbTaxCodes_Purchase_display = Null
14780     Else
14790       .cmbTaxCodes_Purchase_display = .cmbTaxCodes_Purchase.Column(1)
14800     End If
14810   End With

EXITP:
14820   Exit Sub

ERRH:
14830   Select Case ERR.Number
        Case Else
14840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14850   End Select
14860   Resume EXITP

End Sub

Public Sub CheckCalendar1(blnGoodDate As Boolean, datStartDate As Date)
' ** Called by the Calendar subform.

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckCalendar1"

14910   With Me
14920     Select Case blnGoodDate
          Case True
            ' ** Allow posting up to 1 month into the future.
14930       If datStartDate > DateAdd("m", 1, Date) Then
14940         MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
14950         .transdate = CDate(Format(Date, "mm/dd/yyyy"))
14960       Else
14970         .transdate = datStartDate
14980       End If
14990     Case False
15000       .transdate = CDate(Format(Date, "mm/dd/yyyy"))
15010     End Select
15020     Select Case .cmbAssets.Enabled
          Case True
15030       .cmbAssets.SetFocus
15040     Case False
15050       .shareface.SetFocus
15060     End Select
15070   End With

EXITP:
15080   Exit Sub

ERRH:
15090   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
15100   Case Else
15110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15120   End Select
15130   Resume EXITP

End Sub

Public Sub CheckCalendar2(blnGoodDate As Boolean, datStartDate As Date)
' ** Called by the Calendar subform.

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckCalendar2"

15210   With Me
15220     Select Case blnGoodDate
          Case True
            ' ** Don't allow any future dates.
15230       If Compare_DateA_DateB(datStartDate, ">", Date) = True Then  ' ** Module Function: modStringFuncs.
15240         MsgBox "Future trade dates are not allowed.", vbInformation + vbOKOnly, "Invalid Date"
15250         .assetdate_display = Date
15260         .assetdate = Now()
15270       Else
15280         .assetdate_display = datStartDate
15290         .assetdate = datStartDate + time
15300       End If
15310     Case False
15320       .assetdate_display = Date
15330       .assetdate = Now()
15340     End Select
15350     Select Case .cmbLocations.Enabled
          Case True
15360       .cmbLocations.SetFocus
15370     Case False
15380       .ICash.SetFocus
15390     End Select
15400   End With

EXITP:
15410   Exit Sub

ERRH:
15420   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
15430   Case Else
15440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15450   End Select
15460   Resume EXITP

End Sub

Private Sub NoChar_Load()
' ** Load list of currencies that the Fixedsys font doesn't show.

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

15510   Set dbs = CurrentDb
15520   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
15530     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
15540     Set rst = qdf.OpenRecordset
15550     With rst
15560       .MoveLast
15570       lngNoChars = .RecordCount
15580       .MoveFirst
15590       arr_varNoChar = .GetRows(lngNoChars)
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
15600       .Close
15610     End With
15620     .Close
15630   End With

EXITP:
15640   Set rst = Nothing
15650   Set qdf = Nothing
15660   Set dbs = Nothing
15670   Exit Sub

ERRH:
15680   DoCmd.Hourglass False
15690   Select Case ERR.Number
        Case Else
15700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15710   End Select
15720   Resume EXITP

End Sub
