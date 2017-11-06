VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssets_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssets_Sub"

'VGC 04/16/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd on frmAssets}
' **   Close:            C {cmdClose on frmAssets}
' **   Edit:             E {cmdEdit on frmAssets}
' **   Show Hidden:      H {chkShowHidden on frmAssets}
' **   CUSIP:            I {cusip}
' **   Incl Currency:    L {chkIncludeCurrency on frmAssets}
' **   Show AssetNo:     N {chkShowAssetNo on frmAssets}
' **   Only Assets Held: O {opgShow_optHeldOnly on frmAssets}
' **   Print:            P {cmdPrintReport on frmAssets}
' **   Shorten:          S {cmdShorten}
' **   All Assets:       T {opgShow_optAll on frmAssets}
' **   Update:           U {cmdUpdate on frmAssets}
' **   Widen:            W {cmdWiden}
' **   Exit:             X {cmdClose on frmAssets}
' **   Asset Type:       Y {cmbAssetType}

' ** Shortcut F-Keys responsive from this form:
' **   Recalc New:      F5 {RecalcNewAssets, RecalcForEx on frmAssets}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ** MasterAsset table:
' **   Masterasset_TYPE field:
' **     RA = Regular Asset
' **     IA = Interest Asset
' **       CUSIP: 999999999
' **       This asset comes with the program, there can be only one, and it
' **       can't be deleted, though it can be renamed to whatever you want.
' **       Used with frmAccruedIncome.
' **   Description field prefixes:
' **     'HA-' = Hidden Asset
' **       Doesn't show up in drop-downs, and can only be hidden if 0 shares.
' **       Can have as many as you like.
' **     'SW-' = Sweep Asset
' **       Used with account_SWEEP field in Account table, and frmSweeper.
' **       Only 'SW-' assets show up in drop-down.
' **       Can have as many as you like.

' ** RecChanged = 0 : Unchanged.
' **              1 : Changed.
' **              3 : Marked for deletion (currently not allowed).

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[assettype], [description_masterasset_sort], [cusip]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

' ** Combo box column constants: curr_id.
Private Const CBX_C_CID   As Integer = 0  'curr_id
Private Const CBX_C_CODE  As Integer = 1  'curr_code
Private Const CBX_C_NAME  As Integer = 2  'curr_name

Private blnShorten_Focus As Boolean, blnShorten_MouseDown As Boolean
Private blnWiden_Focus As Boolean, blnWiden_MouseDown As Boolean

Private lngAssetNo_Width As Long, lngAssetDesc_left As Long, lngRateEntry_Left As Long
Private lngParForm_Width As Long, lngAssetDesc_Width As Long, lngAssetDescAstno_Width As Long
Private lngAssetNext_Left As Long, lngForm_Width As Long, lngRatePct_Offset As Long

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngCurrID_Old As Long, strCurrCode_Old As String, strCurrName_Old As String
Private lngCurrID_New As Long, strCurrCode_New As String, strCurrName_New As String
Private blnIAHasChanged As Boolean, blnSetToRate As Boolean, intRecChanged As Integer
Private strThisAssetType As String
Private lngRecsCur As Long, lngIsHid_Width As Long, blnIsHid As Boolean
Private strCallingForm As String, lngTpp As Long
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAssets" Then

160         blnSetToRate = False: blnIsHid = False

170         CLR_DISABLED_FG = CLR_DKGRY
180         CLR_DISABLED_BG = CLR_LTTEAL

190         blnShorten_Focus = False: blnShorten_MouseDown = False
200         blnWiden_Focus = False: blnWiden_MouseDown = False

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
210         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

220         .form_width_line.Left = 0&  ' ** Just get it out of the way.
230         lngForm_Width = .Width
240         lngSortLbl_Top = .Sort_lbl.Top
250         lngSortLbl_Width = .Sort_lbl.Width
260         lngAssetNo_Width = (.assetno.Width + (4& * lngTpp))
270         lngAssetDesc_Width = .description_masterasset.Width
280         lngAssetDescAstno_Width = (.description_masterasset.Width - lngAssetNo_Width)
290         lngAssetDesc_left = .description_masterasset.Left
300         lngAssetNext_Left = .curr_id.Left  ' ** Whether or not it's visible.
310         lngRateEntry_Left = .rate_entry.Left
320         lngRatePct_Offset = .rate_pct_lbl.Left - .rate_entry.Left

            ' ** Since there's nowhere to go, disable cmdShorten.
330         .cmdShorten.Enabled = False
340         .cmdShorten_raised_img_dis.Visible = True
350         .cmdShorten_raised_img.Visible = False
360         .cmdShorten_raised_semifocus_dots_img.Visible = False
370         .cmdShorten_raised_focus_img.Visible = False
380         .cmdShorten_raised_focus_dots_img.Visible = False
390         .cmdShorten_sunken_focus_dots_img.Visible = False

400       Else
410         Cancel = -1
420       End If
430     End With

440     If Cancel = -1 Then
450       Beep
460       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
470     End If

EXITP:
480     Exit Sub

ERRH:
490     Select Case ERR.Number
        Case Else
500       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
510     End Select
520     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     SortNow THIS_PROC  ' ** Procedure: Below

620     With Me

630       .assetno.ForeColor = CLR_DISABLED_FG
640       .assetno.BackColor = CLR_DISABLED_BG
650       .cusip.ForeColor = CLR_DISABLED_FG
660       .cusip.BackColor = CLR_DISABLED_BG
670       .description_masterasset.ForeColor = CLR_DISABLED_FG
680       .description_masterasset.BackColor = CLR_DISABLED_BG
690       .curr_id.ForeColor = CLR_DISABLED_FG
700       .curr_id.BackColor = CLR_DISABLED_BG
710       .curr_id_usd_box.ForeColor = CLR_DISABLED_BG
720       .curr_id_forex_box.ForeColor = CLR_VLTGRN
730       .rate_entry.ForeColor = CLR_DISABLED_FG
740       .rate_entry.BackColor = CLR_DISABLED_BG
750       .due.ForeColor = CLR_DISABLED_FG
760       .due.BackColor = CLR_DISABLED_BG
770       .shareface.ForeColor = CLR_DISABLED_FG
780       .shareface.BackColor = CLR_DISABLED_BG
790       .cmbAssetType.ForeColor = CLR_DISABLED_FG
800       .cmbAssetType.BackColor = CLR_DISABLED_BG

810       .FocusHolder.SetFocus

820     End With

EXITP:
830     Exit Sub

ERRH:
840     Select Case ERR.Number
        Case Else
850       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
860     End Select
870     Resume EXITP

End Sub

Private Sub Form_Current()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

910     With Me
920       intRecChanged = 0
930       If .Parent.chkReloading = False Then
940         If .cusip.Enabled = True Then
950           lngCurrID_Old = .curr_id.Column(CBX_C_CID)
960           strCurrCode_Old = .curr_id.Column(CBX_C_CODE)
970           strCurrName_Old = .curr_id.Column(CBX_C_NAME)
980           If .masterasset_TYPE = "IA" Then
990             .cusip.Locked = True
1000          ElseIf .RecChanged = 3 Then  ' ** Marked for deletion.
1010            .cusip.Locked = True
1020          Else
1030            .cusip.Locked = False
1040          End If
1050          If .RecChanged = 3 Then  ' ** Marked for deletion.
                ' ** Record is marked for deletion. (WHICH IS CURRENTLY NOT ALLOWED!)
1060            .description_masterasset.Locked = True
1070            .curr_id.Locked = True
1080            .rate_entry.Locked = True
1090            .due.Locked = True
1100            .cmbAssetType.Locked = True
1110          Else
1120            .description_masterasset.Locked = False
1130            .curr_id.Locked = False
1140            .rate_entry.Locked = False
1150            .due.Locked = False
1160            .cmbAssetType.Locked = False
1170          End If
1180          If IsNull(.cmbAssetType) = False Then
1190            strThisAssetType = .cmbAssetType
1200          Else
1210            strThisAssetType = vbNullString
1220          End If
1230        End If
1240      End If
1250    End With

EXITP:
1260    Exit Sub

ERRH:
1270    THAT_PROC = THIS_PROC
1280    That_Erl = Erl
1290    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1300    Resume EXITP

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

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Recalc New:      F5 {RecalcNewAssets on frmAssets}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   CUSIP:            I {cusip}
        ' **   Shorten:          S {cmdShorten}
        ' **   Widen:            W {cmdWiden}
        ' **   Exit:             X {cmdClose on frmAssets}
        ' **   Asset Type:       Y {cmbAssetType}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyF5
1470        With Me
1480          intRetVal = 0
1490          .Parent.RecalcNewAssets  ' ** Form Procedure: frmAssets.
1500          .Parent.RecalcForEx  ' ** Form Procedure: frmAssets.
1510        End With
1520      Case vbKeyDown
1530        intRetVal = 0
1540        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1550      Case vbKeyUp
1560        intRetVal = 0
1570        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1580      Case vbKeyEscape
1590        With Me
1600          intRetVal = 0
1610          .Parent.cmdClose_Click  ' ** Form Procedure: frmAssets.
1620        End With
1630      End Select
1640    End If

        ' ** Alt keys.
1650    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1660      Select Case intRetVal
          Case vbKeyI
1670        With Me
1680          intRetVal = 0
1690          If .cusip.Enabled = True Then
1700            .cusip.SetFocus
1710          End If
1720        End With
1730      Case vbKeyS
            ' ** Invoke the shorten-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
1740        With Me
1750          intRetVal = 0
1760          If .cmdShorten.Enabled = True Then
1770            cmdShorten_Click  ' ** Procedure: Below.
1780          Else
1790            Beep
1800          End If
1810        End With
1820      Case vbKeyW
            ' ** Invoke the widen-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
1830        With Me
1840          intRetVal = 0
1850          If .cmdWiden.Enabled = True Then
1860            cmdWiden_Click  ' ** Procedure: Below.
1870          Else
1880            Beep
1890          End If
1900        End With
1910      Case vbKeyX
1920        With Me
1930          intRetVal = 0
1940          .Parent.cmdClose_Click  ' ** Form Procedure: frmAssets.
1950        End With
1960      Case vbKeyY
1970        With Me
1980          intRetVal = 0
1990          If .cmbAssetType.Enabled = True Then
2000            .cmbAssetType.SetFocus
2010          End If
2020        End With
2030      End Select
2040    End If

        ' ** Ctrl keys.
2050    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2060      Select Case intRetVal
          Case vbKeyS
2070        intRetVal = 0
2080        cmdSave_Click  ' ** Procedure: Below.
2090      Case vbKeyTab
2100        With Me
2110          intRetVal = 0
2120          DoCmd.SelectObject acForm, .Parent.Name, False
2130          .Parent.opgShow.SetFocus
2140        End With
2150      Case vbKeyPageUp
2160        intRetVal = 0
2170        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2180      Case vbKeyPageDown
2190        intRetVal = 0
2200        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2210      End Select
2220    End If

        ' ** Ctrl-Shift keys.
2230    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2240      Select Case intRetVal
          Case vbKeyF
2250        With Me
2260          intRetVal = 0
2270          DoCmd.SelectObject acForm, .Parent.Name, False
2280          .Parent.FocusHolder.SetFocus
2290        End With
2300      Case vbKeyTab
2310        With Me
2320          intRetVal = 0
2330          DoCmd.SelectObject acForm, .Parent.Name, False
2340          .Parent.cmdClose.SetFocus
2350        End With
2360      End Select
2370    End If

EXITP:
2380    KeyCode = intRetVal
2390    Exit Sub

ERRH:
2400    intRetVal = 0
2410    THAT_PROC = THIS_PROC
2420    That_Erl = Erl
2430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2440    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

        Dim strMsg As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

2510    blnContinue = True

2520    With Me
2530      If IsNull(.cmbAssetType) Then
2540        MsgBox "You must select an asset type to continue.", vbInformation + vbOKOnly, "Entry Required"
2550        .cmbAssetType.SetFocus
2560      Else
2570        blnIAHasChanged = .Parent.IAHasChanged_Get  ' ** Form Procedure: frmAssets.
2580        If .masterasset_TYPE = "IA" And blnIAHasChanged = False Then
2590          strMsg = "Note:" & vbCrLf
2600          strMsg = strMsg & "This asset is a specially coded 'Accrued Interest Asset' for use" & vbCrLf & _
                "with the Accrued Interest feature found on the Post Menu." & vbCrLf
2610          strMsg = strMsg & "Though you may change its name, it will retain its unique Accrued Interest designation." & vbCrLf & vbCrLf
2620          strMsg = strMsg & "Do you wish to change the description of this asset?"
2630          msgResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Accrued Interest Asset")
2640          If msgResponse = vbNo Then
2650            blnContinue = False
2660            Cancel = -1: .Undo
2670          Else
                ' ** It doesn't matter whether they update this change or not, it
                ' ** just stops the message from appearing again during this session.
                ' ** And if they do update with the changes, the next time they come
                ' ** in here, LoadIAData() will see that, and suppress the message.
2680            blnIAHasChanged = True
2690            .Parent.IAHasChanged_Set blnIAHasChanged  ' ** Form Procedure: frmAssets.
2700          End If
2710        End If
2720        If blnContinue = True Then
2730          If .RecChanged = 0 Then  ' ** Unchanged.
2740            .RecChanged = 1        ' ** Changed.
2750          End If
2760          .Parent.chkUpdated = True
2770          If .Parent.cmdUpdate.Enabled <> .Parent.chkUpdated Then
2780            .Parent.cmdUpdate.Enabled = .Parent.chkUpdated
2790          End If
2800          strMsg = vbNullString
2810  On Error Resume Next
2820          strMsg = Screen.ActiveControl.Name
2830          If strMsg <> "cmdEdit" Then
2840            .Parent.cmdEdit.Enabled = False
2850          End If
2860  On Error GoTo ERRH
2870        End If
2880      End If
2890    End With

EXITP:
2900    Exit Sub

ERRH:
2910    THAT_PROC = THIS_PROC
2920    That_Erl = Erl
2930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2940    Resume EXITP

End Sub

Private Sub Form_Timer()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

3010    With Me
3020      .TimerInterval = 0&
3030      If blnSetToRate = True Then
3040        blnSetToRate = False
3050        .rate_entry.SetFocus
3060      End If
3070      If blnIsHid = True Then
3080        blnIsHid = False
3090        If .IsHid_Sum = 0 Then
3100          If .Parent.chkShowHidden.Enabled = True Then
3110            .Parent.chkShowHidden = False
3120            .Parent.chkShowHidden_lbl.FontBold = False
3130            .Parent.chkShowHidden_lbl2_dim.Visible = True
3140            .Parent.chkShowHidden_lbl2_dim_hi.Visible = True
3150            .Parent.chkShowHidden_lbl3_dim.Visible = True
3160            .Parent.chkShowHidden_lbl3_dim_hi.Visible = True
3170            .Parent.chkShowHidden_lbl.Visible = False
3180            .Parent.chkShowHidden.Enabled = False
3190            DoEvents
3200          End If
3210        Else
3220          If .Parent.chkShowHidden.Enabled = False Then
3230            .Parent.chkShowHidden.Enabled = True
3240            .Parent.chkShowHidden_lbl.Visible = True
3250            .Parent.chkShowHidden_lbl2_dim.Visible = False
3260            .Parent.chkShowHidden_lbl2_dim_hi.Visible = False
3270            .Parent.chkShowHidden_lbl3_dim.Visible = False
3280            .Parent.chkShowHidden_lbl3_dim_hi.Visible = False
3290            DoEvents
3300          End If
3310        End If
3320      End If
3330    End With

EXITP:
3340    Exit Sub

ERRH:
3350    Select Case ERR.Number
        Case Else
3360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3370    End Select
3380    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3410    If DataErr <> 0 Then
3420      Select Case DataErr
          Case 2046, 2108, 2116, 2169, 2237, 2501, 3020, 3101, 3314, 7753, 8519
            ' ** 2046: The command or action '|' isn't available now.
            ' ** 2108: You must save the field before you execute the GoToControl action, the GoToControl method, or the SetFocus method.
            ' ** 2116: The value in the field or record violates the validation rule for the record or field.
            ' ** 2169: You can't save this record at this time.
            ' ** 2237: The text you entered isn't an item in the list.
            ' ** 2501: The '|' action was Canceled.
            ' ** 3020: Update or CancelUpdate without AddNew or Edit.
            ' ** 3101: The Microsoft Jet database engine cannot find a record in the table '|2' with key matching field(s) '|1'.
            ' ** 3314: The field '|' cannot contain a Null value because the Required property for this field is set to True.
            ' ** 7753: The value you entered doesn't meet the validation rule defined for the field or control.
            ' ** 8519: You are about to delete | record(s).
3430        Response = acDataErrContinue
3440      Case 2105  ' ** You can't go to the specified record.
3450        Response = acDataErrContinue
3460      Case 2113  ' ** 2113: The value you entered isn't valid for this field.
3470        Select Case Me.ActiveControl.Name
            Case "due"
3480          MsgBox "The date you entered is invalid.", vbInformation + vbOKOnly, "Invalid Entry"
3490          Response = acDataErrContinue
3500        Case "shareface"
3510          MsgBox "The data you entered must be numeric.", vbInformation + vbOKOnly, "Invalid Entry"
3520          Response = acDataErrContinue
3530        Case "rate_entry", "rate"
3540          MsgBox "The data you entered must be numeric.", vbInformation + vbOKOnly, "Invalid Entry"
3550          Response = acDataErrContinue
3560        Case Else
3570          Response = acDataErrContinue
3580        End Select
3590      Case 2143  ' ** You didn't specify a search criteria with a FindRecord action.
            ' ** This only shows up when using Shift-Tab to go from cusip to the previous record. Totally BIZARRE!
3600        Response = acDataErrContinue
3610      Case 2279  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
3620        MsgBox "You must enter the data in the format MM/DD/YYYY.", vbInformation + vbOKOnly, "Invalid Date"
3630        Response = acDataErrContinue
3640        DoCmd.CancelEvent
3650      Case 2427  ' ** The expression you entered has no value.
3660        Response = acDataErrContinue
3670      Case 3022  ' ** The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.
3680        MsgBox "The changes you requested were not successful because they would create duplicate values." & vbCrLf & vbCrLf & _
              "Change the data in the fields that contain duplicate data!", vbInformation + vbOKOnly, "Invalid Entry"
3690        Response = acDataErrContinue
3700        DoCmd.CancelEvent
3710      Case 3162  ' ** You tried to assign the Null value to a variable that is not a Variant data type.
3720        If Me.ActiveControl.Name = "cmbAssetType" Then
3730          MsgBox "The Asset Type field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
3740        Else
3750          MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Invalid Entry"
3760        End If
3770        ERR.Clear
3780        DoCmd.CancelEvent
3790      Case Else
3800        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3810        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3820      End Select
3830    End If

        ' ** AcDataError enumeration:
        ' **   0  acDataErrContinue  Ignore the error and continue without displaying the default Microsoft Access
        ' **                         error message. A custom error message may be displayed in place of the default
        ' **                         error message.
        ' **   1  acDataErrDisplay   Display the default Microsoft Access error message. (Default)
        ' **   2  acDataErrAdded     Don't display the default Microsoft Access error message. The entry may be
        ' **                         added to the combo box list in the NotInList event procedure. After the entry
        ' **                         is added, Microsoft Access updates the list by requerying the combo box.
        ' **                         Microsoft Access then rechecks the string against the combo box list, and saves
        ' **                         the value in the NewData argument in the field the combo box is bound to. If
        ' **                         the string is not in the list, then Microsoft Access displays an error message.

EXITP:
3840    Exit Sub

ERRH:
3850    Select Case ERR.Number
        Case Else
3860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3870    End Select
3880    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3910    With Me
3920      If .cmdShorten_raised_focus_dots_img.Visible = True Or .cmdShorten_raised_focus_img.Visible = True Then
3930        Select Case blnShorten_Focus
            Case True
3940          .cmdShorten_raised_semifocus_dots_img.Visible = True
3950          .cmdShorten_raised_img.Visible = False
3960        Case False
3970          .cmdShorten_raised_img.Visible = True
3980          .cmdShorten_raised_semifocus_dots_img.Visible = False
3990        End Select
4000        .cmdShorten_raised_focus_img.Visible = False
4010        .cmdShorten_raised_focus_dots_img.Visible = False
4020        .cmdShorten_sunken_focus_dots_img.Visible = False
4030        .cmdShorten_raised_img_dis.Visible = False
4040      End If
4050      If .cmdWiden_raised_focus_dots_img.Visible = True Or .cmdWiden_raised_focus_img.Visible = True Then
4060        Select Case blnWiden_Focus
            Case True
4070          .cmdWiden_raised_semifocus_dots_img.Visible = True
4080          .cmdWiden_raised_img.Visible = False
4090        Case False
4100          .cmdWiden_raised_img.Visible = True
4110          .cmdWiden_raised_semifocus_dots_img.Visible = False
4120        End Select
4130        .cmdWiden_raised_focus_img.Visible = False
4140        .cmdWiden_raised_focus_dots_img.Visible = False
4150        .cmdWiden_sunken_focus_dots_img.Visible = False
4160        .cmdWiden_raised_img_dis.Visible = False
4170      End If
4180    End With

EXITP:
4190    Exit Sub

ERRH:
4200    Select Case ERR.Number
        Case Else
4210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4220    End Select
4230    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

4310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Public Sub cmdSave_Click()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

4410    DoCmd.RunCommand acCmdSaveRecord

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4520    intShiftDown = (Shift And acShiftMask) > 0
4530    intAltDown = (Shift And acAltMask) > 0
4540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4570        With Me
4580          intRetVal = 0
4590          lngRecsCur = RecCnt  ' ** Function: Below.
4600          If .CurrentRecord < lngRecsCur Then
4610            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4620            If .cusip.Enabled = True Then
4630              .cusip.SetFocus
4640            Else
4650              .FocusHolder.SetFocus
4660            End If
4670          Else
4680            If .cusip.Enabled = True Then
4690              .cusip.SetFocus
4700            Else
4710              .Parent.opgShow.SetFocus
4720            End If
4730          End If
4740        End With
4750      End Select
4760    End If

        ' ** Shift keys.
4770    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4780      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4790        With Me
4800          intRetVal = 0
4810          If .CurrentRecord > 1 Then
4820            MoveRec acCmdRecordsGoToPrevious   ' ** Procedure: Below.
4830            If .cusip.Enabled = True Then
4840              .cusip.SetFocus
4850            Else
4860              .FocusHolder.SetFocus
4870            End If
4880          Else
4890            .Parent.cmdClose.SetFocus
4900          End If
4910        End With
4920      End Select
4930    End If

EXITP:
4940    KeyCode = intRetVal
4950    Exit Sub

ERRH:
4960    intRetVal = 0
4970    THAT_PROC = THIS_PROC
4980    That_Erl = Erl
4990    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5000    Resume EXITP

End Sub

Private Sub Sort_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "Sort_lbl_MouseMove"

5110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
5120    Exit Sub

ERRH:
5130    Select Case ERR.Number
        Case Else
5140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5150    End Select
5160    Resume EXITP

End Sub

Private Sub assetno_lbl_DblClick(Cancel As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_lbl_DblClick"

5210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5220    Exit Sub

ERRH:
5230    Select Case ERR.Number
        Case Else
5240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5250    End Select
5260    Resume EXITP

End Sub

Private Sub cusip_lbl_DblClick(Cancel As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_lbl_DblClick"

5310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5320    Exit Sub

ERRH:
5330    Select Case ERR.Number
        Case Else
5340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5350    End Select
5360    Resume EXITP

End Sub

Private Sub cusip_BeforeUpdate(Cancel As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_BeforeUpdate"

        Dim strMsg As String

5410    With Me
5420      strMsg = vbNullString
5430      If IsNull(Me.cusip) = True Then
5440        strMsg = "CUSIP field must be filled in to continue."
5450      Else
5460        If Me.cusip = vbNullString Then
5470          strMsg = "CUSIP field must be filled in to continue."
5480        ElseIf Me.masterasset_TYPE = "IA" Then
5490          strMsg = "This asset is a specially coded 'Accrued Interest Asset' for use" & vbCrLf & _
                "with the Accrued Interest feature found on the Post Menu." & vbCrLf & _
                "Its CUSIP cannot be changed."
5500        End If
5510      End If
5520      If strMsg <> vbNullString Then
5530        If .masterasset_TYPE <> "IA" Then
5540          MsgBox strMsg, vbInformation + vbOKOnly, "Entry Required"
5550        Else
5560          MsgBox strMsg, vbInformation + vbOKOnly, "Invalid Entry"
5570          .Undo
5580        End If
5590        Cancel = -1
5600      End If
5610    End With

EXITP:
5620    Exit Sub

ERRH:
5630    Select Case ERR.Number
        Case Else
5640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5650    End Select
5660    Resume EXITP

End Sub

Private Sub cusip_KeyDown(KeyCode As Integer, Shift As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5720    intShiftDown = (Shift And acShiftMask) > 0
5730    intAltDown = (Shift And acAltMask) > 0
5740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5770        With Me
5780          intRetVal = 0
5790          .description_masterasset.SetFocus
5800        End With
5810      End Select
5820    End If

        ' ** Shift keys.
5830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5850        With Me
5860          intRetVal = 0
5870          If .CurrentRecord > 1 Then
5880            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
5890            .cmbAssetType.SetFocus  ' ** For some odd reason, an error shows up here.
5900          Else
5910            DoCmd.SelectObject acForm, .Parent.Name, False
5920            .Parent.cmdClose.SetFocus
5930          End If
5940        End With
5950      End Select
5960    End If

EXITP:
5970    KeyCode = intRetVal
5980    Exit Sub

ERRH:
5990    intRetVal = 0
6000    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
6010    Case Else
6020      THAT_PROC = THIS_PROC
6030      That_Erl = Erl
6040      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6050    End Select
6060    Resume EXITP

End Sub

Private Sub cusip_Exit(Cancel As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_Exit"

6110    cmdSave_Click  ' ** Procedure: Above.

EXITP:
6120    Exit Sub

ERRH:
6130    Select Case ERR.Number
        Case Else
6140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6150    End Select
6160    Resume EXITP

End Sub

Private Sub description_masterasset_lbl_DblClick(Cancel As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "description_masterasset_lbl_DblClick"

6210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6220    Exit Sub

ERRH:
6230    Select Case ERR.Number
        Case Else
6240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6250    End Select
6260    Resume EXITP

End Sub

Private Sub description_masterasset_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "description_masterasset_lbl_MouseMove"

6310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
6320    Exit Sub

ERRH:
6330    Select Case ERR.Number
        Case Else
6340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6350    End Select
6360    Resume EXITP

End Sub

Private Sub description_masterasset_Enter()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "description_masterasset_Enter"

6410    With Me
6420      If Left(.description_masterasset, 3) = "HA-" Then
6430        blnIsHid = True
6440      Else
6450        blnIsHid = False
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

Private Sub description_masterasset_AfterUpdate()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "description_masterasset_AfterUpdate"

        Dim intPos01 As Integer
        Dim strTmp01 As String

6610    With Me
6620      If IsNull(.description_masterasset) = False Then
6630        If Trim(.description_masterasset) <> vbNullString Then
6640          strTmp01 = Trim(.description_masterasset)
6650          intPos01 = InStr(strTmp01, Chr(34))  ' ** No double-quotes allowed, so just replace with single-quote, apostrophe.
6660          If intPos01 > 0 Then
6670            Do While intPos01 > 0
6680              strTmp01 = Left(strTmp01, (intPos01 - 1)) & "'" & Mid(strTmp01, (intPos01 + 1))
6690              strTmp01 = Trim(strTmp01)
6700              intPos01 = InStr(strTmp01, Chr(34))
6710            Loop
6720            .description_masterasset = strTmp01
6730            Beep
6740          End If
6750          If Left(strTmp01, 3) = "HA-" Then
6760            If .Parent.chkShowHidden.Enabled = False Then
6770              .Parent.chkShowHidden_lbl.Visible = True
6780              .Parent.chkShowHidden_lbl2_dim.Visible = False
6790              .Parent.chkShowHidden_lbl2_dim_hi.Visible = False
6800              .Parent.chkShowHidden_lbl3_dim.Visible = False
6810              .Parent.chkShowHidden_lbl3_dim_hi.Visible = False
6820              .Parent.chkShowHidden.Enabled = True
6830              .Parent.chkShowHidden = True
6840              .Parent.chkShowHidden_lbl.FontBold = True
6850              DoEvents
6860            End If
6870          Else
6880            If blnIsHid = True Then
                  ' ** This was hidden, now it no longer is.
6890              .IsHid_Sum.Requery
6900              DoEvents
6910              If .IsHid_Sum = 0 Then  ' ** May not have registered yet
6920                .Parent.chkShowHidden = False
6930                .Parent.chkShowHidden_lbl2_dim.Visible = True
6940                .Parent.chkShowHidden_lbl2_dim_hi.Visible = True
6950                .Parent.chkShowHidden_lbl3_dim.Visible = True
6960                .Parent.chkShowHidden_lbl3_dim_hi.Visible = True
6970                .Parent.chkShowHidden_lbl.FontBold = False
6980                .Parent.chkShowHidden_lbl.Visible = False
6990                .Parent.chkShowHidden.Enabled = False
7000                DoEvents
7010              End If
                  ' ** Leave blnIsHid = True.
7020              .TimerInterval = 500&
7030            End If
7040            If Left(strTmp01, 3) = "SW-" Then
7050              Select Case .cmbAssetType
                  Case "60", "80", "81"
                    ' ** These are currently the only ones that can receive Sweep funds.
7060              Case Else
7070                Beep
7080                MsgBox "Only Asset Types 60 and 80 can receive Sweep funds.", vbInformation + vbOKOnly, "Asset Type Inappropriate For Sweep"
7090              End Select
7100            End If
7110          End If
7120        End If
7130      End If
7140    End With

EXITP:
7150    Exit Sub

ERRH:
7160    Select Case ERR.Number
        Case Else
7170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7180    End Select
7190    Resume EXITP

End Sub

Private Sub description_masterasset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "description_masterasset_MouseMove"

7210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
7220    Exit Sub

ERRH:
7230    Select Case ERR.Number
        Case Else
7240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7250    End Select
7260    Resume EXITP

End Sub

Private Sub description_masterasset_Exit(Cancel As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "description_masterasset_Exit"

7310    cmdSave_Click  ' ** Procedure: Above.

EXITP:
7320    Exit Sub

ERRH:
7330    Select Case ERR.Number
        Case Else
7340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7350    End Select
7360    Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

7410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7420    Exit Sub

ERRH:
7430    Select Case ERR.Number
        Case Else
7440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7450    End Select
7460    Resume EXITP

End Sub

Private Sub curr_id_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_MouseMove"

7510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
7520    Exit Sub

ERRH:
7530    Select Case ERR.Number
        Case Else
7540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7550    End Select
7560    Resume EXITP

End Sub

Private Sub curr_id_Enter()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_Enter"

7610    With Me
7620      intRecChanged = .RecChanged
7630    End With

EXITP:
7640    Exit Sub

ERRH:
7650    THAT_PROC = THIS_PROC
7660    That_Erl = Erl
7670    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7680    Resume EXITP

End Sub

Private Sub curr_id_BeforeUpdate(Cancel As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_BeforeUpdate"

        Dim strDocName As String

7710    With Me
7720      If IsNull(.curr_id) = False Then
7730        lngCurrID_New = .curr_id.Column(CBX_C_CID)
7740        If lngCurrID_New <> lngCurrID_Old Then
7750          strDocName = "frmAssets_Currency"
7760          strCurrCode_New = .curr_id.Column(CBX_C_CODE)
7770          strCurrName_New = .curr_id.Column(CBX_C_NAME)
7780          gblnMessage = True  ' ** If user cancels, returns False.
7790          gblnForeignCurrencies = False  ' ** If returns True, conver all values to new currency.
7800          Beep
7810          DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME & "~" & CStr(lngCurrID_Old) & "~" & CStr(lngCurrID_New) & "~" & _
                strCurrCode_Old & "~" & strCurrCode_New & "~" & strCurrName_Old & "~" & strCurrName_New
7820          Select Case gblnMessage
              Case True
7830            .curr_code = strCurrCode_New
7840            .curr_id_old = lngCurrID_Old  ' ** Set in Form_Current().
7850            .curr_code_old = strCurrCode_Old
7860            lngCurrID_Old = lngCurrID_New
7870            strCurrCode_Old = strCurrCode_New
7880            .RecChanged = 1
7890            .curr_convert = gblnForeignCurrencies
7900          Case False
7910            Cancel = -1
7920            .Undo
7930            lngCurrID_New = 0&
7940            strCurrCode_New = vbNullString
7950            strCurrName_New = vbNullString
7960            gblnForeignCurrencies = False
7970          End Select
7980        End If
7990      End If
8000      DoCmd.Hourglass False
8010    End With

EXITP:
8020    Exit Sub

ERRH:
8030    DoCmd.Hourglass False
8040    THAT_PROC = THIS_PROC
8050    That_Erl = Erl
8060    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8070    Resume EXITP

End Sub

Private Sub curr_id_AfterUpdate()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_AfterUpdate"

8110    With Me
8120      cmdSave_Click  ' ** Procedure: Above.
8130    End With

EXITP:
8140    Exit Sub

ERRH:
8150    THAT_PROC = THIS_PROC
8160    That_Erl = Erl
8170    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8180    Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

8210    With Me
8220      MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
8230      .curr_id.SetFocus
8240      Response = acDataErrContinue
8250    End With

EXITP:
8260    Exit Sub

ERRH:
8270    THAT_PROC = THIS_PROC
8280    That_Erl = Erl
8290    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8300    Resume EXITP

End Sub

Private Sub curr_id_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_MouseMove"

8410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
8420    Exit Sub

ERRH:
8430    Select Case ERR.Number
        Case Else
8440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8450    End Select
8460    Resume EXITP

End Sub

Private Sub curr_id_Exit(Cancel As Integer)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_Exit"

8510    cmdSave_Click  ' ** Procedure: Above.

EXITP:
8520    Exit Sub

ERRH:
8530    Select Case ERR.Number
        Case Else
8540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8550    End Select
8560    Resume EXITP

End Sub

Private Sub rate_entry_lbl_DblClick(Cancel As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "rate_entry_lbl_DblClick"

8610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
8620    Exit Sub

ERRH:
8630    Select Case ERR.Number
        Case Else
8640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8650    End Select
8660    Resume EXITP

End Sub

Private Sub rate_entry_Exit(Cancel As Integer)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "rate_entry_Exit"

        Dim intPos01 As Integer
        Dim strTmp01 As String, dblTmp02 As Double

8710    With Me
8720      If IsNull(.rate_entry) = True Then
8730        .rate_entry = 0#
8740        .rate = 0#
8750      End If
8760      strTmp01 = CStr(.rate_entry)
8770      strTmp01 = Trim(strTmp01)
8780      If Right(strTmp01, 1) = "%" Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
8790      If IsNumeric(strTmp01) = True Then
8800        dblTmp02 = CDbl(strTmp01)
8810        If dblTmp02 >= 100# Then
8820          MsgBox "Rate must be less than 100 percent.", vbInformation + vbOKOnly, "Invalid Entry"
8830          .rate_entry = Null
8840          .rate = Null
8850          cmdSave_Click  ' ** Procedure: Above.
8860          .rate_entry.SetFocus
8870          blnSetToRate = True
8880          .TimerInterval = 100&
8890        Else
8900          intPos01 = InStr(strTmp01, ".")
8910          If intPos01 > 0 Then
                ' ** The user entered a decimal number.
8920            If Len(Mid(strTmp01, (intPos01 + 1))) > 3 Then
8930              MsgBox "You can only enter a rate percentage with up to 3 decimal places.", vbInformation + vbOKOnly, "Invalid Entry"
8940              .rate_entry = Null
8950              .rate = Null
8960              cmdSave_Click  ' ** Procedure: Above.
8970              .rate_entry.SetFocus
8980              blnSetToRate = True
8990              .TimerInterval = 100&
9000            Else
9010              If (dblTmp02 / 100#) <> .rate Then
9020                .rate = dblTmp02 / 100#  ' ** Saved as a percentage, i.e., must be less than 1!
9030                cmdSave_Click  ' ** Procedure: Above.
9040              End If
9050            End If
9060          Else
9070            If dblTmp02 <> 0# Then
9080              If (dblTmp02 / 100#) <> .rate Then
9090                .rate = dblTmp02 / 100#  ' ** Saved as a percentage, i.e., must be less than 1!
9100                cmdSave_Click  ' ** Procedure: Above.
9110              End If
9120            Else
9130              If .rate <> 0# Then
9140                .rate = 0#
9150                cmdSave_Click  ' ** Procedure: Above.
9160              End If
9170            End If
9180          End If
9190        End If
9200      Else
9210        MsgBox "Please enter a valid rate percentage.", vbInformation + vbOKOnly, "Invalid Entry"
9220        .rate_entry = Null
9230        .rate = Null
9240        cmdSave_Click  ' ** Procedure: Above.
9250        .rate_entry.SetFocus
9260        blnSetToRate = True
9270        .TimerInterval = 100&
9280      End If
9290    End With

EXITP:
9300    Exit Sub

ERRH:
9310    THAT_PROC = THIS_PROC
9320    That_Erl = Erl
9330    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9340    Resume EXITP

End Sub

Private Sub due_lbl_DblClick(Cancel As Integer)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "due_lbl_DblClick"

9410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
9420    Exit Sub

ERRH:
9430    Select Case ERR.Number
        Case Else
9440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9450    End Select
9460    Resume EXITP

End Sub

Private Sub due_Exit(Cancel As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "due_Exit"

9510    cmdSave_Click  ' ** Procedure: Above.

EXITP:
9520    Exit Sub

ERRH:
9530    Select Case ERR.Number
        Case Else
9540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9550    End Select
9560    Resume EXITP

End Sub

Private Sub shareface_lbl_DblClick(Cancel As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_lbl_DblClick"

9610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
9620    Exit Sub

ERRH:
9630    Select Case ERR.Number
        Case Else
9640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9650    End Select
9660    Resume EXITP

End Sub

Private Sub cmbAssetType_lbl_DblClick(Cancel As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssetType_lbl_DblClick"

9710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
9720    Exit Sub

ERRH:
9730    Select Case ERR.Number
        Case Else
9740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9750    End Select
9760    Resume EXITP

End Sub

Private Sub cmbAssetType_GotFocus()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssetType_GotFocus"

9810    With Me
9820      ExpandCombo .cmbAssetType  ' ** Module Procedure: modUtilities
9830    End With

EXITP:
9840    Exit Sub

ERRH:
9850    Select Case ERR.Number
        Case Else
9860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9870    End Select
9880    Resume EXITP

End Sub

Private Sub cmbAssetType_BeforeUpdate(Cancel As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssetType_BeforeUpdate"

        Dim strTmp01 As String

9910    With Me
9920      If Not .NewRecord And (IsNull(.cmbAssetType) Or (.cmbAssetType = vbNullString)) Then
9930        MsgBox "You must select an asset type to continue.", vbInformation + vbOKOnly, "Entry Required"
9940        Cancel = -1
9950        DoCmd.CancelEvent
9960      Else
9970  On Error Resume Next
9980        strTmp01 = .cmbAssetType.OldValue
9990  On Error GoTo ERRH
10000       If ((strTmp01 = "90" Or strThisAssetType = "90") And .cmbAssetType <> "90") Or _
                ((strThisAssetType <> "90" And .cmbAssetType = "90") Or _
                (strTmp01 <> vbNullString And strTmp01 <> "90" And .cmbAssetType = "90")) Then
10010         If strTmp01 <> vbNullString Then
10020           strTmp01 = strTmp01 & " - " & DLookup("[description]", "AssetType", "[assettype] = '" & strTmp01 & "'")
10030         Else
10040           strTmp01 = strThisAssetType & " - " & DLookup("[description]", "AssetType", "[assettype] = '" & strThisAssetType & "'")
10050         End If

10060         MsgBox "A Liability Asset Type cannot be changed, nor a non-Liability changed to a Liability." & vbCrLf & vbCrLf & _
                "If an Asset was coded in error, you must add a new Asset of the correct type," & vbCrLf & _
                "then back out any transactions to the incorrect one and re-enter them correctly." & vbCrLf & vbCrLf & vbCrLf & _
                "Original Asset Type: " & strTmp01, vbInformation + vbOKOnly, "Liability Error"
10070         Cancel = -1
10080       End If
10090     End If
10100   End With

EXITP:
10110   Exit Sub

ERRH:
10120   Select Case ERR.Number
        Case 3162  ' ** Tried to be blank.
10130     ERR.Clear
10140     DoCmd.CancelEvent
10150   Case Else
10160     THAT_PROC = THIS_PROC
10170     That_Erl = Erl
10180     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10190   End Select
10200   Resume EXITP

End Sub

Private Sub cmbAssetType_Exit(Cancel As Integer)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssetType_Exit"

10310   With Me
10320     If Not .NewRecord And (IsNull(.cmbAssetType) Or (.cmbAssetType = vbNullString)) Then
10330       MsgBox "You must select an asset type to continue.", vbInformation + vbOKOnly, "Entry Required"
10340       DoCmd.CancelEvent
10350     Else
10360       cmdSave_Click  ' ** Procedure: Above.
10370     End If
10380   End With

EXITP:
10390   Exit Sub

ERRH:
10400   Select Case ERR.Number
        Case 3162  ' ** Tried to be blank.
10410     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10420     ERR.Clear
10430     DoCmd.CancelEvent
10440   Case Else
10450     THAT_PROC = THIS_PROC
10460     That_Erl = Erl
10470     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10480   End Select
10490   Resume EXITP

End Sub

Private Sub cmbAssetType_KeyDown(KeyCode As Integer, Shift As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssetType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10520   intShiftDown = (Shift And acShiftMask) > 0
10530   intAltDown = (Shift And acAltMask) > 0
10540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10570       With Me
10580         intRetVal = 0
10590         lngRecsCur = RecCnt  ' ** Function: Below.
10600         If .CurrentRecord < lngRecsCur Then
10610           MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
10620           .cusip.SetFocus
10630         Else
10640           DoCmd.SelectObject acForm, .Parent.Name, False
10650           .Parent.opgShow.SetFocus
10660         End If
10670       End With
10680     End Select
10690   End If

        ' ** Shift keys.
10700   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10710     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10720       With Me
10730         intRetVal = 0
10740         .due.SetFocus
10750       End With
10760     End Select
10770   End If

EXITP:
10780   KeyCode = intRetVal
10790   Exit Sub

ERRH:
10800   intRetVal = 0
10810   THAT_PROC = THIS_PROC
10820   That_Erl = Erl
10830   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10840   Resume EXITP

End Sub

Private Sub cmdShorten_GotFocus()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_GotFocus"

10910   With Me
10920     blnShorten_Focus = True
10930     .cmdShorten_raised_semifocus_dots_img.Visible = True
10940     .cmdShorten_raised_img.Visible = False
10950     .cmdShorten_raised_focus_img.Visible = False
10960     .cmdShorten_raised_focus_dots_img.Visible = False
10970     .cmdShorten_sunken_focus_dots_img.Visible = False
10980     .cmdShorten_raised_img_dis.Visible = False
10990   End With

EXITP:
11000   Exit Sub

ERRH:
11010   Select Case ERR.Number
        Case Else
11020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11030   End Select
11040   Resume EXITP

End Sub

Private Sub cmdShorten_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseDown"

11110   With Me
11120     blnShorten_MouseDown = True
11130     .cmdShorten_sunken_focus_dots_img.Visible = True
11140     .cmdShorten_raised_img.Visible = False
11150     .cmdShorten_raised_semifocus_dots_img.Visible = False
11160     .cmdShorten_raised_focus_img.Visible = False
11170     .cmdShorten_raised_focus_dots_img.Visible = False
11180     .cmdShorten_raised_img_dis.Visible = False
11190   End With

EXITP:
11200   Exit Sub

ERRH:
11210   Select Case ERR.Number
        Case Else
11220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11230   End Select
11240   Resume EXITP

End Sub

Public Sub cmdShorten_Click()
' ** lngTpp * 4 = 60  : 1 grid point on form (my screen is 15 Twips Per Pixel).
' ** 60 * 4 = 240     : 4 grid points, a nice noticeable amount.
' ** 1440 Twips per Inch.

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_Click"

        Dim lngMinusedWidth As Long, blnSortHere As Boolean
        Dim lngDiff1 As Long, lngDiff2 As Long

11310   With Me
11320     blnSortHere = False
11330     lngMinusedWidth = lngTpp * WIDEN_MULT
11340     If (.description_masterasset.Width <= IIf(.Parent.chkShowAssetNo = True, lngAssetDescAstno_Width, lngAssetDesc_Width)) Then
11350       Beep
11360     ElseIf ((.description_masterasset.Width - lngMinusedWidth) < IIf(.Parent.chkShowAssetNo = True, lngAssetDescAstno_Width, lngAssetDesc_Width)) Then
11370       .description_masterasset.Width = IIf(.Parent.chkShowAssetNo = True, lngAssetDescAstno_Width, lngAssetDesc_Width)
11380       .description_masterasset_lbl.Width = IIf(.Parent.chkShowAssetNo = True, lngAssetDescAstno_Width, lngAssetDesc_Width)
11390       .description_masterasset_lbl_line.Width = (.description_masterasset_lbl.Width + lngTpp)
11400       If .Sort_line.Left = .description_masterasset_lbl.Left Then
11410         .Sort_line.Width = (.description_masterasset_lbl.Width + lngTpp)
11420         .Sort_lbl.Left = (((.description_masterasset_lbl.Left + .description_masterasset_lbl.Width) - lngSortLbl_Width) + (2& * lngTpp))
11430       End If
11440       lngDiff1 = (.rate_entry.Left - .curr_id.Left)
11450       If .Sort_line.Left = .curr_id.Left Then blnSortHere = True
11460       .curr_id = lngAssetNext_Left
11470       .curr_id_lbl.Left = (lngAssetNext_Left - (4& * lngTpp))
11480       .curr_id_lbl_line.Left = lngAssetNext_Left
11490       If blnSortHere = True Then
11500         .Sort_line.Left = .curr_id.Left
11510         .Sort_lbl.Left = ((.curr_id.Left + .curr_id.Width) - lngSortLbl_Width)  'THIS WILL DEFINITELY HAVE TO CHANGE!
11520       End If
11530       blnSortHere = False
11540       lngDiff2 = (.due.Left - .rate_entry.Left)
11550       If .Sort_line.Left = .rate_entry_lbl.Left Then blnSortHere = True
11560       .rate_entry.Left = (.curr_id.Left + lngDiff1)
11570       .rate_pct_lbl.Left = (.rate_entry.Left + lngRatePct_Offset)
11580       .rate_entry_lbl.Left = .rate_entry.Left
11590       .rate_entry_lbl_line.Left = .rate_entry_lbl.Left
11600       If blnSortHere = True Then
11610         .Sort_line.Left = .rate_entry_lbl.Left
11620         .Sort_lbl.Left = ((.rate_entry_lbl.Left + .rate_entry_lbl.Width) - lngSortLbl_Width)
11630       End If
11640       blnSortHere = False
11650       lngDiff1 = (.txtLgrUse.Left - .due.Left)
11660       If .Sort_line.Left = .due_lbl.Left Then blnSortHere = True
11670       .due.Left = (.rate_entry.Left + lngDiff2)
11680       .due_lbl.Left = .due.Left
11690       .due_lbl_line.Left = .due_lbl.Left
11700       If blnSortHere = True Then
11710         .Sort_line.Left = .due_lbl.Left
11720         .Sort_lbl.Left = ((.due_lbl.Left + .due_lbl.Width) - lngSortLbl_Width)
11730       End If
11740       blnSortHere = False
11750       lngDiff2 = (.Deleted.Left - .txtLgrUse.Left)
11760       .txtLgrUse.Left = (.due.Left + lngDiff1)
11770       .txtMAType.Left = txtLgrUse.Left
11780       lngDiff1 = (.cmbAssetType.Left - .Deleted.Left)
11790       .Deleted.Left = (.txtLgrUse.Left + lngDiff2)
11800       If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
11810       .shareface.Left = .Deleted.Left
11820       .shareface_lbl.Left = .Deleted.Left
11830       .shareface_lbl_line.Left = .shareface_lbl.Left
11840       If blnSortHere = True Then
11850         .Sort_line.Left = .shareface_lbl.Left
11860         .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
11870       End If
11880       blnSortHere = False
11890       If .Sort_line.Left = .cmbAssetType_lbl.Left Then blnSortHere = True
11900       .cmbAssetType.Left = (.shareface.Left + lngDiff1)
11910       .cmbAssetType_lbl.Left = .cmbAssetType.Left
11920       .cmbAssetType_lbl_line.Left = .cmbAssetType_lbl.Left
11930       If blnSortHere = True Then
11940         .Sort_line.Left = .cmbAssetType_lbl.Left
11950         .Sort_lbl.Left = ((.cmbAssetType_lbl.Left + .cmbAssetType_lbl.Width) - lngSortLbl_Width)
11960       End If
11970       blnSortHere = False
11980       .Width = lngParForm_Width
11990     Else
12000       .description_masterasset.Width = (.description_masterasset.Width - lngMinusedWidth)
12010       .description_masterasset_lbl.Width = .description_masterasset.Width
12020       .description_masterasset_lbl_line.Width = (.description_masterasset_lbl.Width + lngTpp)
12030       If .Sort_line.Left = .description_masterasset_lbl.Left Then
12040         .Sort_line.Width = (.description_masterasset_lbl.Width + lngTpp)
12050         .Sort_lbl.Left = (((.description_masterasset_lbl.Left + .description_masterasset_lbl.Width) - lngSortLbl_Width) + (2& * lngTpp))
12060       End If
12070       If .Sort_line.Left = .curr_id.Left Then blnSortHere = True
12080       .curr_id.Left = (.curr_id.Left - lngMinusedWidth)
12090       .curr_id_usd_box.Left = (.curr_id.Left - lngTpp)
12100       .curr_id_forex_box.Left = (.curr_id.Left - lngTpp)
12110       .curr_id_lbl.Left = (.curr_id.Left - (4& * lngTpp))
12120       .curr_id_lbl_line.Left = .curr_id.Left
12130       If blnSortHere = True Then
12140         .Sort_line.Left = .curr_id.Left
12150         .Sort_lbl.Left = ((.curr_id.Left + .curr_id.Width) - lngSortLbl_Width)
12160       End If
12170       blnSortHere = False
12180       If .Sort_line.Left = .rate_entry_lbl.Left Then blnSortHere = True
12190       .rate_entry.Left = (.rate_entry.Left - lngMinusedWidth)
12200       .rate_pct_lbl.Left = (.rate_entry.Left + lngRatePct_Offset)
12210       .rate_entry_lbl.Left = .rate_entry.Left
12220       .rate_entry_lbl_line.Left = .rate_entry_lbl.Left
12230       If blnSortHere = True Then
12240         .Sort_line.Left = .rate_entry_lbl.Left
12250         .Sort_lbl.Left = ((.rate_entry_lbl.Left + .rate_entry_lbl.Width) - lngSortLbl_Width)
12260       End If
12270       blnSortHere = False
12280       If .Sort_line.Left = .due_lbl.Left Then blnSortHere = True
12290       .due.Left = .due.Left - lngMinusedWidth
12300       .due_lbl.Left = .due.Left
12310       .due_lbl_line.Left = .due_lbl.Left
12320       If blnSortHere = True Then
12330         .Sort_line.Left = .due_lbl.Left
12340         .Sort_lbl.Left = ((.due_lbl.Left + .due_lbl.Width) - lngSortLbl_Width)
12350       End If
12360       blnSortHere = False
12370       .txtLgrUse.Left = .txtLgrUse.Left - lngMinusedWidth
12380       .txtMAType.Left = txtLgrUse.Left
12390       .Deleted.Left = .Deleted.Left - lngMinusedWidth
12400       If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
12410       .shareface.Left = .Deleted.Left
12420       .shareface_lbl.Left = .Deleted.Left
12430       .shareface_lbl_line.Left = .shareface_lbl.Left
12440       If blnSortHere = True Then
12450         .Sort_line.Left = .shareface_lbl.Left
12460         .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
12470       End If
12480       blnSortHere = False
12490       If .Sort_line.Left = .cmbAssetType_lbl.Left Then blnSortHere = True
12500       .cmbAssetType.Left = .cmbAssetType.Left - lngMinusedWidth
12510       .cmbAssetType_lbl.Left = .cmbAssetType.Left
12520       .cmbAssetType_lbl_line.Left = .cmbAssetType_lbl.Left
12530       If blnSortHere = True Then
12540         .Sort_line.Left = .cmbAssetType_lbl.Left
12550         .Sort_lbl.Left = ((.cmbAssetType_lbl.Left + .cmbAssetType_lbl.Width) - lngSortLbl_Width)
12560       End If
12570       blnSortHere = False
12580       .Width = .Width - lngMinusedWidth
12590     End If
12600     .FocusHolder.SetFocus
12610     If .cmdWiden.Enabled = False Then
            ' ** Since cmdShorten was clicked, cmdWiden is certainly available.
12620       .cmdWiden.Enabled = True
12630       .cmdWiden_raised_img.Visible = True
12640       .cmdWiden_raised_semifocus_dots_img.Visible = False
12650       .cmdWiden_raised_focus_img.Visible = False
12660       .cmdWiden_raised_focus_dots_img.Visible = False
12670       .cmdWiden_sunken_focus_dots_img.Visible = False
12680       .cmdWiden_raised_img_dis.Visible = False
12690     End If
12700     If (.description_masterasset.Width <= IIf(.Parent.chkShowAssetNo = True, lngAssetDescAstno_Width, lngAssetDesc_Width)) Then
            ' ** If there's nowhere to go, disable it.
12710       .cmdShorten.Enabled = False
12720       .cmdShorten_raised_img_dis.Visible = True
12730       .cmdShorten_raised_img.Visible = False
12740       .cmdShorten_raised_semifocus_dots_img.Visible = False
12750       .cmdShorten_raised_focus_img.Visible = False
12760       .cmdShorten_raised_focus_dots_img.Visible = False
12770       .cmdShorten_sunken_focus_dots_img.Visible = False
12780     End If
12790   End With

EXITP:
12800   Exit Sub

ERRH:
12810   Select Case ERR.Number
        Case Else
12820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12830   End Select
12840   Resume EXITP

End Sub

Private Sub cmdShorten_KeyDown(KeyCode As Integer, Shift As Integer)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_KeyDown"

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
12990         .cmdWiden.SetFocus
13000       End With
13010     End Select
13020   End If

        ' ** Shift keys.
13030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13040     Select Case intRetVal
          Case vbKeyTab
13050       With Me
13060         intRetVal = 0
13070         If .cmbAssetType.Enabled = True Then
13080           .cmbAssetType.SetFocus
13090         Else
13100           .FocusHolder.SetFocus
13110         End If
13120       End With
13130     End Select
13140   End If

EXITP:
13150   KeyCode = intRetVal
13160   Exit Sub

ERRH:
13170   intRetVal = 0
13180   THAT_PROC = THIS_PROC
13190   That_Erl = Erl
13200   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13210   Resume EXITP

End Sub

Private Sub cmdShorten_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseMove"

13310   With Me
13320     If blnShorten_MouseDown = False Then
13330       Select Case blnShorten_Focus
            Case True
13340         .cmdShorten_raised_focus_dots_img.Visible = True
13350         .cmdShorten_raised_focus_img.Visible = False
13360       Case False
13370         .cmdShorten_raised_focus_img.Visible = True
13380         .cmdShorten_raised_focus_dots_img.Visible = False
13390       End Select
13400       .cmdShorten_raised_img.Visible = False
13410       .cmdShorten_raised_semifocus_dots_img.Visible = False
13420       .cmdShorten_sunken_focus_dots_img.Visible = False
13430       .cmdShorten_raised_img_dis.Visible = False
13440     End If
13450     If .cmdWiden_raised_focus_dots_img.Visible = True Or .cmdWiden_raised_focus_img.Visible = True Then
13460       Select Case blnWiden_Focus
            Case True
13470         .cmdWiden_raised_semifocus_dots_img.Visible = True
13480         .cmdWiden_raised_img.Visible = False
13490       Case False
13500         .cmdWiden_raised_img.Visible = True
13510         .cmdWiden_raised_semifocus_dots_img.Visible = False
13520       End Select
13530       .cmdWiden_raised_focus_img.Visible = False
13540       .cmdWiden_raised_focus_dots_img.Visible = False
13550       .cmdWiden_sunken_focus_dots_img.Visible = False
13560       .cmdWiden_raised_img_dis.Visible = False
13570     End If
13580   End With

EXITP:
13590   Exit Sub

ERRH:
13600   Select Case ERR.Number
        Case Else
13610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13620   End Select
13630   Resume EXITP

End Sub

Private Sub cmdShorten_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseUp"

13710   With Me
13720     .cmdShorten_raised_focus_dots_img.Visible = True
13730     .cmdShorten_raised_img.Visible = False
13740     .cmdShorten_raised_semifocus_dots_img.Visible = False
13750     .cmdShorten_raised_focus_img.Visible = False
13760     .cmdShorten_sunken_focus_dots_img.Visible = False
13770     .cmdShorten_raised_img_dis.Visible = False
13780     blnShorten_MouseDown = False
13790   End With

EXITP:
13800   Exit Sub

ERRH:
13810   Select Case ERR.Number
        Case Else
13820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13830   End Select
13840   Resume EXITP

End Sub

Private Sub cmdShorten_LostFocus()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_LostFocus"

13910   With Me
13920     .cmdShorten_raised_img.Visible = True
13930     .cmdShorten_raised_semifocus_dots_img.Visible = False
13940     .cmdShorten_raised_focus_img.Visible = False
13950     .cmdShorten_raised_focus_dots_img.Visible = False
13960     .cmdShorten_sunken_focus_dots_img.Visible = False
13970     .cmdShorten_raised_img_dis.Visible = False
13980     blnShorten_Focus = False
13990   End With

EXITP:
14000   Exit Sub

ERRH:
14010   Select Case ERR.Number
        Case Else
14020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14030   End Select
14040   Resume EXITP

End Sub

Private Sub cmdWiden_GotFocus()

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_GotFocus"

14110   With Me
14120     blnWiden_Focus = True
14130     .cmdWiden_raised_semifocus_dots_img.Visible = True
14140     .cmdWiden_raised_img.Visible = False
14150     .cmdWiden_raised_focus_img.Visible = False
14160     .cmdWiden_raised_focus_dots_img.Visible = False
14170     .cmdWiden_sunken_focus_dots_img.Visible = False
14180     .cmdWiden_raised_img_dis.Visible = False
14190   End With

EXITP:
14200   Exit Sub

ERRH:
14210   Select Case ERR.Number
        Case Else
14220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14230   End Select
14240   Resume EXITP

End Sub

Private Sub cmdWiden_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseDown"

14310   With Me
14320     blnWiden_MouseDown = True
14330     .cmdWiden_sunken_focus_dots_img.Visible = True
14340     .cmdWiden_raised_img.Visible = False
14350     .cmdWiden_raised_semifocus_dots_img.Visible = False
14360     .cmdWiden_raised_focus_img.Visible = False
14370     .cmdWiden_raised_focus_dots_img.Visible = False
14380     .cmdWiden_raised_img_dis.Visible = False
14390   End With

EXITP:
14400   Exit Sub

ERRH:
14410   Select Case ERR.Number
        Case Else
14420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14430   End Select
14440   Resume EXITP

End Sub

Public Sub cmdWiden_Click()
' ** lngTpp * 4 = 60  : 1 grid point on form (my screen is 15 Twips Per Pixel).
' ** 60 * 4 = 240     : 4 grid points, a nice noticeable amount.
' ** 1440 Twips per Inch.
' ** Arbitrarily cap it at 10 in. wide.

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_Click"

        Dim lngAddedWidth As Long, blnSortHere As Boolean

14510   With Me
14520     blnSortHere = False
14530     lngAddedWidth = lngTpp * WIDEN_MULT
14540     If (.Width >= WIDEN_MAX) Or ((.Width + lngAddedWidth) > WIDEN_MAX) Then
14550       Beep
14560     Else
14570       .Width = .Width + lngAddedWidth
14580       If .Sort_line.Left = .cmbAssetType_lbl.Left Then blnSortHere = True
14590       .cmbAssetType.Left = (.cmbAssetType.Left + lngAddedWidth)
14600       .cmbAssetType_lbl.Left = .cmbAssetType.Left
14610       .cmbAssetType_lbl_line.Left = .cmbAssetType_lbl.Left
14620       If blnSortHere = True Then
14630         .Sort_line.Left = .cmbAssetType_lbl.Left
14640         .Sort_lbl.Left = ((.cmbAssetType_lbl.Left + .cmbAssetType_lbl.Width) - lngSortLbl_Width)
14650       End If
14660       blnSortHere = False
14670       .Deleted.Left = (.Deleted.Left + lngAddedWidth)
14680       If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
14690       .shareface.Left = .Deleted.Left
14700       .shareface_lbl.Left = .Deleted.Left
14710       .shareface_lbl_line.Left = .shareface_lbl.Left
14720       If blnSortHere = True Then
14730         .Sort_line.Left = .shareface_lbl.Left
14740         .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
14750       End If
14760       blnSortHere = False
14770       .txtLgrUse.Left = (.txtLgrUse.Left + lngAddedWidth)
14780       .txtMAType.Left = txtLgrUse.Left
14790       If .Sort_line.Left = .due_lbl.Left Then blnSortHere = True
14800       .due.Left = (.due.Left + lngAddedWidth)
14810       .due_lbl.Left = .due.Left
14820       .due_lbl_line.Left = .due_lbl.Left
14830       If blnSortHere = True Then
14840         .Sort_line.Left = .due_lbl.Left
14850         .Sort_lbl.Left = ((.due_lbl.Left + .due_lbl.Width) - lngSortLbl_Width)
14860       End If
14870       blnSortHere = False
14880       If .Sort_line.Left = .rate_entry_lbl.Left Then blnSortHere = True
14890       .rate_entry.Left = (.rate_entry.Left + lngAddedWidth)
14900       .rate_pct_lbl.Left = (.rate_entry.Left + lngRatePct_Offset)
14910       .rate_entry_lbl.Left = .rate_entry.Left
14920       .rate_entry_lbl_line.Left = .rate_entry_lbl.Left
14930       If blnSortHere = True Then
14940         .Sort_line.Left = .rate_entry_lbl.Left
14950         .Sort_lbl.Left = ((.rate_entry_lbl.Left + .rate_entry_lbl.Width) - lngSortLbl_Width)
14960       End If
14970       blnSortHere = False
14980       If .Sort_line.Left = .curr_id.Left Then blnSortHere = True
14990       .curr_id.Left = (.curr_id.Left + lngAddedWidth)
15000       .curr_id_usd_box.Left = (.curr_id.Left - lngTpp)
15010       .curr_id_forex_box.Left = (.curr_id.Left - lngTpp)
15020       .curr_id_lbl.Left = (.curr_id.Left - (4& * lngTpp))
15030       .curr_id_lbl_line.Left = .curr_id.Left
15040       If blnSortHere = True Then
15050         .Sort_line.Left = .curr_id.Left
15060         .Sort_lbl.Left = ((.curr_id.Left + .curr_id.Width) - lngSortLbl_Width)
15070       End If
15080       blnSortHere = False
15090       .description_masterasset.Width = .description_masterasset.Width + lngAddedWidth
15100       .description_masterasset_lbl.Width = .description_masterasset.Width
15110       .description_masterasset_lbl_line.Width = (.description_masterasset_lbl.Width + lngTpp)
15120       If .Sort_line.Left = .description_masterasset_lbl.Left Then
15130         .Sort_line.Width = (.description_masterasset_lbl.Width + lngTpp)
15140         .Sort_lbl.Left = (((.description_masterasset_lbl.Left + .description_masterasset_lbl.Width) - lngSortLbl_Width) + (2& * lngTpp))
15150       End If
15160     End If
15170     .FocusHolder.SetFocus
15180     If .cmdShorten.Enabled = False Then
            ' ** Since cmdWiden was clicked, cmdShorten is certainly available.
15190       .cmdShorten.Enabled = True
15200       .cmdShorten_raised_img.Visible = True
15210       .cmdShorten_raised_semifocus_dots_img.Visible = False
15220       .cmdShorten_raised_focus_img.Visible = False
15230       .cmdShorten_raised_focus_dots_img.Visible = False
15240       .cmdShorten_sunken_focus_dots_img.Visible = False
15250       .cmdShorten_raised_img_dis.Visible = False
15260     End If
          ' ** If there's nowhere to go, disable it.
15270     If (.Width >= WIDEN_MAX) Or ((.Width + lngAddedWidth) > WIDEN_MAX) Then
15280       .cmdWiden.Enabled = False
15290       .cmdWiden_raised_img_dis.Visible = True
15300       .cmdWiden_raised_img.Visible = False
15310       .cmdWiden_raised_semifocus_dots_img.Visible = False
15320       .cmdWiden_raised_focus_img.Visible = False
15330       .cmdWiden_raised_focus_dots_img.Visible = False
15340       .cmdWiden_sunken_focus_dots_img.Visible = False
15350     End If
15360   End With

EXITP:
15370   Exit Sub

ERRH:
15380   Select Case ERR.Number
        Case Else
15390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15400   End Select
15410   Resume EXITP

End Sub

Private Sub cmdWiden_KeyDown(KeyCode As Integer, Shift As Integer)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15520   intShiftDown = (Shift And acShiftMask) > 0
15530   intAltDown = (Shift And acAltMask) > 0
15540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15560     Select Case intRetVal
          Case vbKeyTab
15570       With Me
15580         intRetVal = 0
15590         If .cusip.Enabled = True Then
15600           .cusip.SetFocus
15610         Else
15620           .FocusHolder.SetFocus
15630         End If
15640       End With
15650     End Select
15660   End If

        ' ** Shift keys.
15670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15680     Select Case intRetVal
          Case vbKeyTab
15690       With Me
15700         intRetVal = 0
15710         .cmdShorten.SetFocus
15720       End With
15730     End Select
15740   End If

EXITP:
15750   KeyCode = intRetVal
15760   Exit Sub

ERRH:
15770   intRetVal = 0
15780   THAT_PROC = THIS_PROC
15790   That_Erl = Erl
15800   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15810   Resume EXITP

End Sub

Private Sub cmdWiden_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseMove"

15910   With Me
15920     If blnWiden_MouseDown = False Then
15930       Select Case blnWiden_Focus
            Case True
15940         .cmdWiden_raised_focus_dots_img.Visible = True
15950         .cmdWiden_raised_focus_img.Visible = False
15960       Case False
15970         .cmdWiden_raised_focus_img.Visible = True
15980         .cmdWiden_raised_focus_dots_img.Visible = False
15990       End Select
16000       .cmdWiden_raised_img.Visible = False
16010       .cmdWiden_raised_semifocus_dots_img.Visible = False
16020       .cmdWiden_sunken_focus_dots_img.Visible = False
16030       .cmdWiden_raised_img_dis.Visible = False
16040     End If
16050     If .cmdShorten_raised_focus_dots_img.Visible = True Or .cmdShorten_raised_focus_img.Visible = True Then
16060       Select Case blnShorten_Focus
            Case True
16070         .cmdShorten_raised_semifocus_dots_img.Visible = True
16080         .cmdShorten_raised_img.Visible = False
16090       Case False
16100         .cmdShorten_raised_img.Visible = True
16110         .cmdShorten_raised_semifocus_dots_img.Visible = False
16120       End Select
16130       .cmdShorten_raised_focus_img.Visible = False
16140       .cmdShorten_raised_focus_dots_img.Visible = False
16150       .cmdShorten_sunken_focus_dots_img.Visible = False
16160       .cmdShorten_raised_img_dis.Visible = False
16170     End If
16180   End With

EXITP:
16190   Exit Sub

ERRH:
16200   Select Case ERR.Number
        Case Else
16210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16220   End Select
16230   Resume EXITP

End Sub

Private Sub cmdWiden_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseUp"

16310   With Me
16320     .cmdWiden_raised_focus_dots_img.Visible = True
16330     .cmdWiden_raised_img.Visible = False
16340     .cmdWiden_raised_semifocus_dots_img.Visible = False
16350     .cmdWiden_raised_focus_img.Visible = False
16360     .cmdWiden_sunken_focus_dots_img.Visible = False
16370     .cmdWiden_raised_img_dis.Visible = False
16380     blnWiden_MouseDown = False
16390   End With

EXITP:
16400   Exit Sub

ERRH:
16410   Select Case ERR.Number
        Case Else
16420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16430   End Select
16440   Resume EXITP

End Sub

Private Sub cmdWiden_LostFocus()

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_LostFocus"

16510   With Me
16520     .cmdWiden_raised_img.Visible = True
16530     .cmdWiden_raised_semifocus_dots_img.Visible = False
16540     .cmdWiden_raised_focus_img.Visible = False
16550     .cmdWiden_raised_focus_dots_img.Visible = False
16560     .cmdWiden_sunken_focus_dots_img.Visible = False
16570     .cmdWiden_raised_img_dis.Visible = False
16580     blnWiden_Focus = False
16590   End With

EXITP:
16600   Exit Sub

ERRH:
16610   Select Case ERR.Number
        Case Else
16620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16630   End Select
16640   Resume EXITP

End Sub

Public Sub ShowAssetNo(blnShow As Boolean)

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAssetNo"

        Dim blnSortHere As Boolean

16710   With Me
16720     blnSortHere = False
16730     Select Case blnShow
          Case True
16740       If .Sort_line.Left = .description_masterasset_lbl.Left Then blnSortHere = True
16750       .description_masterasset.Width = (.description_masterasset.Width - lngAssetNo_Width)
16760       .description_masterasset.Left = (lngAssetDesc_left + lngAssetNo_Width)
16770       .description_masterasset_lbl.Width = .description_masterasset.Width
16780       .description_masterasset_lbl.Left = .description_masterasset.Left
16790       .description_masterasset_lbl_line.Width = (.description_masterasset_lbl.Width + lngTpp)
16800       .description_masterasset_lbl_line.Left = .description_masterasset_lbl.Left
16810       If blnSortHere = True Then
16820         .Sort_line.Width = (.description_masterasset_lbl.Width + lngTpp)
16830         .Sort_line.Left = .description_masterasset_lbl.Left
16840       End If
16850       blnSortHere = False
16860       If .Sort_line.Left = .cusip_lbl.Left Then blnSortHere = True
16870       .cusip.Left = (.assetno.Left + lngAssetNo_Width)
16880       .IsHid1.Left = .cusip.Left - (4& * lngTpp)
16890       .IsHid2.Left = .IsHid1.Left
16900       .cusip_lbl.Left = .cusip.Left
16910       .cusip_lbl_line.Left = .cusip_lbl.Left
16920       If blnSortHere = True Then
16930         .Sort_line.Width = (.cusip_lbl.Width + lngTpp)
16940         .Sort_line.Left = .cusip_lbl.Left
16950       End If
16960       blnSortHere = False
16970       .assetno.Visible = True
16980       .assetno_lbl.Visible = True
16990       .assetno_lbl2.Visible = True
17000       .assetno_lbl_line.Visible = True
17010     Case False
17020       If .description_masterasset.Left <> lngAssetDesc_left Then
17030         If .Sort_line.Left = .assetno_lbl.Left Then blnSortHere = True
17040         .assetno.Visible = False
17050         .assetno_lbl.Visible = False
17060         .assetno_lbl2.Visible = False
17070         .assetno_lbl_line.Visible = False
17080         If blnSortHere = True Then
17090           .Sort_line.Visible = False
17100           .Sort_lbl.Visible = False
17110         End If
17120         blnSortHere = False
17130         If .Sort_line.Left = .cusip_lbl.Left Then blnSortHere = True
17140         .cusip.Left = .assetno.Left
17150         .IsHid1.Left = 0&
17160         .IsHid2.Left = 0&
17170         .cusip_lbl.Left = .cusip.Left
17180         .cusip_lbl_line.Left = .cusip_lbl.Left
17190         If blnSortHere = True Then
17200           .Sort_line.Left = .cusip_lbl.Left
17210           .Sort_lbl.Left = ((.cusip_lbl.Left + .cusip_lbl.Width) - lngSortLbl_Width)
17220         End If
17230         blnSortHere = False
17240         If .Sort_line.Left = .description_masterasset_lbl.Left Then blnSortHere = True
17250         .description_masterasset.Left = lngAssetDesc_left
17260         .description_masterasset_lbl.Left = .description_masterasset.Left
17270         .description_masterasset_lbl_line.Left = .description_masterasset_lbl.Left
17280         .description_masterasset.Width = (.description_masterasset.Width + lngAssetNo_Width)
17290         .description_masterasset_lbl.Width = .description_masterasset.Width
17300         .description_masterasset_lbl_line.Width = (.description_masterasset_lbl.Width + lngTpp)
17310         If blnSortHere = True Then
17320           .Sort_line.Left = .description_masterasset_lbl.Left
17330           .Sort_line.Width = (.description_masterasset_lbl.Width + lngTpp)
17340         End If
17350         blnSortHere = False
17360       End If

17370     End Select
17380   End With

EXITP:
17390   Exit Sub

ERRH:
17400   Select Case ERR.Number
        Case Else
17410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17420   End Select
17430   Resume EXITP

End Sub

Public Function IAHasChanged_Get() As Boolean

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "IAHasChanged_Get"

        Dim blnRetVal As Boolean

17510   blnRetVal = blnIAHasChanged

EXITP:
17520   IAHasChanged_Get = blnRetVal
17530   Exit Function

ERRH:
17540   blnRetVal = False
17550   Select Case ERR.Number
        Case Else
17560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17570   End Select
17580   Resume EXITP

End Function

Public Sub IAHasChanged_Set(blnSet As Boolean)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "IAHasChanged_Set"

17610   blnIAHasChanged = blnSet

EXITP:
17620   Exit Sub

ERRH:
17630   Select Case ERR.Number
        Case Else
17640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17650   End Select
17660   Resume EXITP

End Sub

Public Sub IsHid_Width_Set(lngWidth As Long)

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "IsHid_Width_Set"

17710   lngIsHid_Width = lngWidth

EXITP:
17720   Exit Sub

ERRH:
17730   Select Case ERR.Number
        Case Else
17740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17750   End Select
17760   Resume EXITP

End Sub

Public Sub ParForm_Width_Set(lngWidth As Long)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "ParForm_Width_Set"

17810   lngParForm_Width = lngWidth

EXITP:
17820   Exit Sub

ERRH:
17830   Select Case ERR.Number
        Case Else
17840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17850   End Select
17860   Resume EXITP

End Sub

Public Sub IncludeCurrency(blnShow As Boolean)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim blnSortHere As Boolean
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, blnTmp04 As Boolean

17910   With Me
17920     blnSortHere = False: blnTmp04 = False
17930     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
17940       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
17950     End If
17960     lngTmp01 = (.curr_id.Width + (4& * lngTpp))
17970     lngTmp02 = (.description_masterasset.Width - lngAssetDesc_Width)  ' ** May be Zero.
17980     lngTmp03 = .rate_entry.Left
17990     Select Case blnShow
          Case True
18000       If .curr_id.Visible = False Then
18010         .Width = (lngForm_Width + lngTmp02)
18020         If .Sort_line.Left = .cmbAssetType_lbl.Left Then blnSortHere = True
18030         .cmbAssetType.Left = (.cmbAssetType.Left + lngTmp01)
18040         .cmbAssetType_lbl.Left = .cmbAssetType.Left
18050         .cmbAssetType_lbl_line.Left = .cmbAssetType_lbl.Left
18060         If blnSortHere = True Then
18070           .Sort_line.Left = .cmbAssetType_lbl.Left
18080           .Sort_lbl.Left = ((.cmbAssetType_lbl.Left + .cmbAssetType_lbl.Width) - lngSortLbl_Width)
18090         End If
18100         blnSortHere = False
18110         If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
18120         .shareface.Left = (.shareface.Left + lngTmp01)
18130         .shareface_lbl.Left = .shareface.Left
18140         .shareface_lbl_line.Left = .shareface_lbl.Left
18150         .Deleted.Left = .shareface.Left
18160         If blnSortHere = True Then
18170           .Sort_line.Left = .shareface_lbl.Left
18180           .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
18190         End If
18200         blnSortHere = False
18210         .txtLgrUse.Left = (.txtLgrUse.Left + lngTmp01)
18220         .txtMAType.Left = .txtLgrUse.Left
18230         If .Sort_line.Left = .due_lbl.Left Then blnSortHere = True
18240         .due.Left = (.due.Left + lngTmp01)
18250         .due_lbl.Left = .due.Left
18260         .due_lbl_line.Left = .due_lbl.Left
18270         If blnSortHere = True Then
18280           .Sort_line.Left = .due_lbl.Left
18290           .Sort_lbl.Left = ((.due_lbl.Left + .due_lbl.Width) - lngSortLbl_Width)
18300         End If
18310         blnSortHere = False
18320         If .Sort_line.Left = .rate_entry_lbl.Left Then blnSortHere = True
18330         .rate_entry.Left = (.curr_id.Left + lngTmp01)
18340         .rate_pct_lbl.Left = (.rate_entry.Left + lngRatePct_Offset)
18350         .rate_entry_lbl.Left = .rate_entry.Left
18360         .rate_entry_lbl_line.Left = .rate_entry_lbl.Left
18370         If blnSortHere = True Then
18380           .Sort_line.Left = .rate_entry_lbl.Left
18390           .Sort_lbl.Left = ((.rate_entry_lbl.Left + .rate_entry_lbl.Width) - lngSortLbl_Width)
18400         End If
18410         blnSortHere = False
18420         .curr_id.Left = lngTmp03
18430         .curr_id_usd_box.Left = (.curr_id.Left - lngTpp)
18440         .curr_id_forex_box.Left = (.curr_id.Left - lngTpp)
18450         .curr_id_lbl.Left = (.curr_id.Left - (4& * lngTpp))
18460         .curr_id_lbl_line.Left = .curr_id.Left
18470         .curr_id.Visible = True
18480         .curr_id_usd_box.Visible = True
18490         .curr_id_forex_box.Visible = True
18500         .curr_id_lbl.Visible = True
18510         .curr_id_lbl_line.Visible = True
18520       End If
18530     Case False
18540       If .curr_id.Visible = True Then
18550         If .Sort_line.Left = .curr_id.Left Then blnSortHere = True
18560         .curr_id.Visible = False
18570         .curr_id_usd_box.Visible = False
18580         .curr_id_forex_box.Visible = False
18590         .curr_id_lbl.Visible = False
18600         .curr_id_lbl_line.Visible = False
18610         If blnSortHere = True Then
18620           blnTmp04 = True
18630         End If
18640         If .Sort_line.Left = .rate_entry_lbl.Left Then blnSortHere = True
18650         .rate_entry.Left = .curr_id.Left
18660         .rate_pct_lbl.Left = (.rate_entry.Left + lngRatePct_Offset)
18670         .rate_entry_lbl.Left = .rate_entry.Left
18680         .rate_entry_lbl_line.Left = .rate_entry_lbl.Left
18690         If blnSortHere = True Then
18700           .Sort_line.Left = .rate_entry_lbl.Left
18710           .Sort_lbl.Left = ((.rate_entry_lbl.Left + .rate_entry_lbl.Width) - lngSortLbl_Width)
18720         End If
18730         blnSortHere = False
18740         If .Sort_line.Left = .due_lbl.Left Then blnSortHere = True
18750         .due.Left = (.due.Left - lngTmp01)
18760         .due_lbl.Left = .due.Left
18770         .due_lbl_line.Left = .due_lbl.Left
18780         If blnSortHere = True Then
18790           .Sort_line.Left = .due_lbl.Left
18800           .Sort_lbl.Left = ((.due_lbl.Left + .due_lbl.Width) - lngSortLbl_Width)
18810         End If
18820         blnSortHere = False
18830         .txtLgrUse.Left = (.txtLgrUse.Left - lngTmp01)
18840         .txtMAType.Left = .txtLgrUse.Left
18850         If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
18860         .shareface.Left = (.shareface.Left - lngTmp01)
18870         .shareface_lbl.Left = .shareface.Left
18880         .shareface_lbl_line.Left = .shareface_lbl.Left
18890         .Deleted.Left = .shareface.Left
18900         If blnSortHere = True Then
18910           .Sort_line.Left = .shareface_lbl.Left
18920           .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
18930         End If
18940         blnSortHere = False
18950         If .Sort_line.Left = .cmbAssetType_lbl.Left Then blnSortHere = True
18960         .cmbAssetType.Left = (.cmbAssetType.Left - lngTmp01)
18970         .cmbAssetType_lbl.Left = .cmbAssetType.Left
18980         .cmbAssetType_lbl_line.Left = .cmbAssetType_lbl.Left
18990         If blnSortHere = True Then
19000           .Sort_line.Left = .cmbAssetType_lbl.Left
19010           .Sort_lbl.Left = ((.cmbAssetType_lbl.Left + .cmbAssetType_lbl.Width) - lngSortLbl_Width)
19020         End If
19030         blnSortHere = False
19040         .Width = ((lngForm_Width - lngTmp01) + lngTmp02)
19050       End If
19060       If blnTmp04 = True Then
19070         SortNow "Form_Load"  ' ** Procedure: Below.
19080       End If
19090     End Select
19100   End With

EXITP:
19110   Exit Sub

ERRH:
19120   Select Case ERR.Number
        Case Else
19130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19140   End Select
19150   Resume EXITP

End Sub

Public Function RecCnt() As Long

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

19210   lngRetVal = 0&

19220   With Me
19230     Set rst = .RecordsetClone
19240     With rst
19250       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
19260       Else
19270         .MoveLast
19280         lngRetVal = .RecordCount
19290       End If
19300       .Close
19310     End With
19320   End With

EXITP:
19330   Set rst = Nothing
19340   RecCnt = lngRetVal
19350   Exit Function

ERRH:
19360   lngRetVal = 0&
19370   Select Case ERR.Number
        Case Else
19380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19390   End Select
19400   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

19510   Select Case intWhere
        Case acCmdRecordsGoToFirst
19520     DoCmd.RunCommand acCmdRecordsGoToFirst
19530   Case acCmdRecordsGoToLast
19540     DoCmd.RunCommand acCmdRecordsGoToLast
19550   Case acCmdRecordsGoToPrevious
19560     DoCmd.RunCommand acCmdRecordsGoToPrevious
19570   Case acCmdRecordsGoToNext
19580     DoCmd.RunCommand acCmdRecordsGoToNext
19590   Case Else
19600     If IsMissing(varID) = False Then
19610       With Me
19620         Set rst = .RecordsetClone
19630         With rst
19640           .FindFirst "[assetno] = " & CStr(varID)
19650           If .NoMatch = False Then
19660             Me.Bookmark = .Bookmark
19670           End If
19680           .Close
19690         End With
19700       End With
19710     End If
19720   End Select

EXITP:
19730   Set rst = Nothing
19740   Exit Sub

ERRH:
19750   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
19760   Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
19770   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
19780   Case Else
19790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19800   End Select
19810   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intFlds As Integer, intPos01 As Integer

        Const strStdAsc As String = ", [description_masterasset_sort], [assettype]"

19910   With Me
19920     .assetno_Sort_line1.Visible = False
19930     .assetno_Sort_line2.Visible = False
19940     .Controls(strSortLbl).Visible = False
19950     .Controls(strSortLbl).Top = lngSortLbl_Top
19960     .Controls(strSortLine).Visible = False
19970     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push of the right side of the form.
19980     If strProc = "Form_Load" Then
19990       strCalled = "cmbAssetType"
20000       strSortNow = strSortOrig
20010       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
20020       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
20030       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
20040       .Controls(strSortLbl).Left = lngSortLbl_Left
20050       .Controls(strSortLbl).Caption = strArwUp
20060       .Controls(strSortLbl).ForeColor = CLR_DKBLU
20070       .Controls(strSortLine).Left = lngSortLine_Left
20080       .Controls(strSortLine).Width = lngSortLine_Width
20090     Else
20100       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
20110       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
20120       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
20130       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
20140       .Controls(strSortLbl).Left = lngSortLbl_Left
20150       If strCalled = "cmbAssetType" Then
20160         strSortAsc = strSortOrig
20170         If strSortNow = strSortAsc Then
20180           strSortNow = "[assettype] DESC, [description_masterasset_sort] DESC, [cusip] DESC"
20190           .Controls(strSortLbl).Caption = strArwDn
20200           .Controls(strSortLbl).ForeColor = CLR_DKRED
20210         Else
20220           strSortNow = strSortAsc
20230           .Controls(strSortLbl).Caption = strArwUp
20240           .Controls(strSortLbl).ForeColor = CLR_DKBLU
20250         End If
20260         .Controls(strSortLine).Left = lngSortLine_Left
20270         .Controls(strSortLine).Width = lngSortLine_Width
20280       Else
20290         .Controls(strSortLbl).Caption = strArwUp
20300         .Controls(strSortLbl).ForeColor = CLR_DKBLU
20310         .Controls(strSortLine).Left = lngSortLine_Left
20320         .Controls(strSortLine).Width = lngSortLine_Width
20330         Select Case strCalled
              Case "assetno"
20340           strSortAsc = "[assetno]"
20350           .Controls(strSortLbl).Top = (lngSortLbl_Top - (2& * lngTpp))
20360           lngSortLbl_Left = (lngSortLbl_Left + (2& * lngTpp))
20370           .Controls(strSortLbl).Left = lngSortLbl_Left
20380         Case "cusip"
20390           strSortAsc = "[cusip], [description_masterasset_sort], [assettype]"
20400         Case "description_masterasset"
20410           strSortAsc = "[description_masterasset_sort], [cusip], [assettype]"
20420           lngSortLbl_Left = (lngSortLbl_Left + (2& * lngTpp))
20430           .Controls(strSortLbl).Left = lngSortLbl_Left
20440         Case "curr_id"
20450           strSortAsc = "[curr_code]" & strStdAsc
20460           lngSortLine_Left = .Controls(strCalled).Left
20470           lngSortLine_Width = (.Controls(strCalled).Width + lngTpp)
20480           lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
20490           .Controls(strSortLbl).Left = lngSortLbl_Left
20500           .Controls(strSortLine).Left = lngSortLine_Left
20510           .Controls(strSortLine).Width = lngSortLine_Width
20520         Case "rate", "rate_entry"
20530           strSortAsc = "[rate]" & strStdAsc
20540         Case "due"
20550           strSortAsc = "[due]" & strStdAsc
20560         Case "shareface"
20570           strSortAsc = "[shareface]" & strStdAsc
20580         End Select
20590         If strSortNow = strSortAsc Then
                ' ** They're currently all 3 fields long.
20600           intFlds = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
20610           Select Case intFlds
                Case 1
20620             strSortNow = strSortAsc & " DESC"
20630           Case 2
20640             intPos01 = InStr(strSortAsc, ",")
20650             strSortNow = Left(strSortAsc, (intPos01 - 1)) & " DESC" & Mid(strSortAsc, intPos01) & " DESC"
20660           Case 3
20670             intPos01 = InStr(strSortAsc, ",")
20680             strSortNow = Left(strSortAsc, (intPos01 - 1)) & " DESC" & Mid(strSortAsc, intPos01)
20690             intPos01 = InStr(strSortNow, ",")
20700             intPos01 = InStr((intPos01 + 1), strSortNow, ",")
20710             strSortNow = Left(strSortNow, (intPos01 - 1)) & " DESC" & Mid(strSortNow, intPos01) & " DESC"
20720           End Select
20730           .Controls(strSortLbl).Caption = strArwDn
20740           .Controls(strSortLbl).ForeColor = CLR_DKRED
20750         Else
20760           strSortNow = strSortAsc
20770         End If
20780       End If
20790     End If
20800     .Controls(strSortLbl).Visible = True
20810     If strCalled = "assetno" Then
20820       .assetno_Sort_line1.Visible = True
20830       .assetno_Sort_line2.Visible = True
20840     Else
20850       .Controls(strSortLine).Visible = True
20860     End If
20870     .OrderBy = strSortNow
20880     .OrderByOn = True
20890     If .Parent.chkOrderBy = True Then
20900       Debug.Print "'" & strSortNow
20910     End If
20920   End With

        ' ** [cusip], [description_masterasset_sort], [assettype]
        ' ** [cusip] DESC, [description_masterasset_sort] DESC, [assettype] DESC
        ' ** [description_masterasset_sort], [cusip], [assettype]
        ' ** [description_masterasset_sort] DESC, [cusip] DESC, [assettype] DESC
        ' ** [rate], [description_masterasset_sort], [assettype]
        ' ** [rate] DESC, [description_masterasset_sort] DESC, [assettype] DESC
        ' ** [due], [description_masterasset_sort], [assettype]
        ' ** [due] DESC, [description_masterasset_sort] DESC, [assettype] DESC
        ' ** [shareface], [description_masterasset_sort], [assettype]
        ' ** [shareface] DESC, [description_masterasset_sort] DESC, [assettype] DESC
        ' ** [assettype], [description_masterasset_sort], [cusip]
        ' ** [assettype] DESC, [description_masterasset_sort] DESC, [cusip] DESC

EXITP:
20930   Exit Sub

ERRH:
20940   Select Case ERR.Number
        Case Else
20950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20960   End Select
20970   Resume EXITP

End Sub

Public Function SortNow_Get() As String

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

21010   strRetVal = strSortNow

EXITP:
21020   SortNow_Get = strRetVal
21030   Exit Function

ERRH:
21040   strRetVal = strSortOrig
21050   Select Case ERR.Number
        Case Else
21060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21070   End Select
21080   Resume EXITP

End Function
