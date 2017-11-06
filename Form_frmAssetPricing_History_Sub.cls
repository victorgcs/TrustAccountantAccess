VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssetPricing_History_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssetPricing_History_Sub"

'VGC 03/16/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmAssetPricing_History}
' **   Date:             D {lbxHistoryDate on frmAssetPricing_History}
' **   One:              E {ckgDeleteDates_chkOne on frmAssetPricing_History}
' **   Show Hidden:      H {chkShowHidden on frmAssetPricing_History}
' **   Include Currency: L {chkIncludeCurrency on frmAssetPricing_History}
' **   Market Value:     M {marketvaluecurrent}
' **   Show AssetNo:     N {chkShowAssetNo on frmAssetPricing_History}
' **   Oldest:           O {ckgDeleteDates_chkOldest on frmAssetPricing_History}
' **   Print Report:     P {cmdPrintReport on frmAssetPricing_History}
' **   Restore:          R {cmdRestore on frmAssetPricing_History}
' **   Shorten:          S {cmdShorten}
' **   Delete:           T {cmdHistoryDeleteOldest, cmdHistoryDeleteOne on frmAssetPricing_History}
' **   Widen:            W {cmdWiden}
' **   Exit:             X {cmdClose on frmAssetPricing_History}
' **   Show Description: Y {chkShowAssetTypeDesc on frmAssetPricing_History}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[assettype], [totdesc], [cusip]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private blnShorten_Focus As Boolean, blnShorten_MouseDown As Boolean
Private blnWiden_Focus As Boolean, blnWiden_MouseDown As Boolean

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngForm_Width As Long, lngTotDesc_Width As Long, lngTotDescAstno_Width As Long, lngTotDescNext_Left As Long
Private lngYieldPct_Offset As Long, lngAssetNo_Width As Long, lngCurrCode_Left As Long, lngCurrCode_Width As Long
Private lngAssetTypeDesc_Left As Long, lngAssetTypeDesc_Width As Long, lngMktValCurUSD_Left As Long, lngMarketValueCurrentLbl_Top As Long

Private strCallingForm As String, lngRecsCur As Long, lngTpp As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAssetPricing_History" Then

160         blnShorten_Focus = False: blnShorten_MouseDown = False
170         blnWiden_Focus = False: blnWiden_MouseDown = False

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190         lngSortLbl_Width = .Sort_lbl.Width

200         .form_width_line.Left = 0&  ' ** Just get this out of the way.

210         lngForm_Width = .Width
220         lngTotDesc_Width = .totdesc.Width
230         lngTotDescNext_Left = .assettype.Left
240         lngYieldPct_Offset = (.yield_pct_lbl.Left - .yield_entry.Left)
250         lngAssetNo_Width = (.assetno.Width + (4& * lngTpp))
260         lngTotDescAstno_Width = (.totdesc.Width - lngAssetNo_Width)
270         lngCurrCode_Left = .curr_code.Left
280         lngCurrCode_Width = (.curr_code.Width + (4& * lngTpp))
290         lngAssetTypeDesc_Left = .assettype_description.Left
300         lngAssetTypeDesc_Width = (.assettype_description.Width + (4& * lngTpp))
310         lngMktValCurUSD_Left = .MarketValueCurrent_usd.Left
320         lngMarketValueCurrentLbl_Top = .marketvaluecurrent_lbl.Top

            ' ** Since there's nowhere to go, disable cmdShorten.
330         .cmdShorten.Enabled = False
340         .cmdShorten_raised_img_dis.Visible = True
350         .cmdShorten_raised_img.Visible = False
360         .cmdShorten_raised_semifocus_dots_img.Visible = False
370         .cmdShorten_raised_focus_img.Visible = False
380         .cmdShorten_raised_focus_dots_img.Visible = False
390         .cmdShorten_sunken_focus_dots_img.Visible = False

            'lngForm_Width
            'lngTotDesc_Width
            'lngTotDescAstno_Width
            'lngTotDescNext_Left
            'lngYieldPct_Offset
            'lngAssetNo_Width
            'lngCurrCode_Left
            'lngCurrCode_Width
            'lngAssetTypeDesc_Left
            'lngAssetTypeDesc_Width
            'lngMktValCurUSD_Left
            'lngMarketValueCurrentLbl_Top

400       Else
410         Cancel = -1
420       End If
430     End With

440     If Cancel = -1 Then
450       DoCmd.Hourglass False
460       Beep
470       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
480     End If

EXITP:
490     Exit Sub

ERRH:
500     Select Case ERR.Number
        Case Else
510       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
520     End Select
530     Resume EXITP

End Sub

Public Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     With Me

620       CLR_DISABLED_FG = CLR_DKGRY
630       CLR_DISABLED_BG = CLR_LTTEAL

640       .assetno.ForeColor = CLR_DISABLED_FG
650       .assetno.BackColor = CLR_DISABLED_BG
660       .cusip.ForeColor = CLR_DISABLED_FG
670       .cusip.BackColor = CLR_DISABLED_BG
680       .totdesc.ForeColor = CLR_DISABLED_FG
690       .totdesc.BackColor = CLR_DISABLED_BG
700       .assettype.ForeColor = CLR_DISABLED_FG
710       .assettype.BackColor = CLR_DISABLED_BG
720       .assettype_description.ForeColor = CLR_DISABLED_FG
730       .assettype_description.BackColor = CLR_DISABLED_BG
740       .curr_code.ForeColor = CLR_DISABLED_FG
750       .curr_code.BackColor = CLR_DISABLED_BG
760       .yield_entry.ForeColor = CLR_DISABLED_FG
770       .yield_entry.BackColor = CLR_DISABLED_BG
780       .marketvaluecurrent_str.ForeColor = CLR_DISABLED_FG
790       .marketvaluecurrent_str.BackColor = CLR_DISABLED_BG
800       .MarketValueCurrent_usd.ForeColor = CLR_DISABLED_FG
810       .MarketValueCurrent_usd.BackColor = CLR_DISABLED_BG

820     End With

830     SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
840     Exit Sub

ERRH:
850     Select Case ERR.Number
        Case Else
860       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
870     End Select
880     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
910     intShiftDown = (Shift And acShiftMask) > 0
920     intAltDown = (Shift And acAltMask) > 0
930     intCtrlDown = (Shift And acCtrlMask) > 0

940     intRetVal = KeyCode

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Market Value:     M {marketvaluecurrent}
        ' **   Shorten:          S {cmdShorten}
        ' **   Delete:           T {cmdHistoryDeleteOldest, cmdHistoryDeleteOne on frmAssetPricing_History}
        ' **   Widen:            W {cmdWiden}
        ' **   Exit:             X {cmdClose on frmAssetPricing_History}

        ' ** Plain keys.
950     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyDown
970         intRetVal = 0
980         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
990       Case vbKeyUp
1000        intRetVal = 0
1010        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1020      Case vbKeyEscape
1030        With Me
1040          intRetVal = 0
1050          .Parent.cmdClose_Click  ' ** Form Procedure: frmAssetPricing_History.
1060        End With
1070      End Select
1080    End If

        ' ** Alt keys.
1090    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1100      Select Case intRetVal
          Case vbKeyM
1110        With Me
1120          intRetVal = 0
1130          If .marketvaluecurrent_str.Enabled = True Then
1140            .marketvaluecurrent_str.SetFocus
1150          End If
1160        End With
1170      Case vbKeyS
            ' ** Invoke the shorten-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
1180        With Me
1190          intRetVal = 0
1200          If .cmdShorten.Enabled = True Then
1210            cmdShorten_Click  ' ** Procedure: Below.
1220          Else
1230            Beep
1240          End If
1250        End With
1260      Case vbKeyT
1270        With Me
1280          If .Parent.cmdHistoryDeleteOldest.Visible = True And .Parent.cmdHistoryDeleteOldest.Enabled = True Then
1290            .Parent.cmdHistoryDeleteOldest.SetFocus
1300            .Parent.cmdHistoryDeleteOldest_Click  ' ** Form Procedure: frmAssetPricing_History.
1310          ElseIf .Parent.cmdHistoryDeleteOne.Visible = True And .Parent.cmdHistoryDeleteOne.Enabled = True Then
1320            .Parent.cmdHistoryDeleteOne.SetFocus
1330            .Parent.cmdHistoryDeleteOne_Click  ' ** Form Procedure: frmAssetPricing_History.
1340          Else
1350            Beep
1360          End If
1370        End With
1380      Case vbKeyW
            ' ** Invoke the widen-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
1390        With Me
1400          intRetVal = 0
1410          If .cmdWiden.Enabled = True Then
1420            cmdWiden_Click  ' ** Procedure: Below.
1430          Else
1440            Beep
1450          End If
1460        End With
1470      Case vbKeyX
1480        With Me
1490          intRetVal = 0
1500          .Parent.cmdClose_Click  ' ** Form Procedure: frmAssetPricing_History.
1510        End With
1520      End Select
1530    End If

        ' ** Ctrl keys.
1540    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1550      Select Case intRetVal
          Case vbKeyTab
1560        With Me
1570          intRetVal = 0
1580          DoCmd.SelectObject acForm, .Parent.Name, False
1590          If .Parent.chkIncludeCurrency.Enabled = True Then
1600            .Parent.chkIncludeCurrency.SetFocus
1610          ElseIf .Parent.chkShowAssetNo.Visible = True And .Parent.chkShowAssetNo.Enabled = True Then
1620            .Parent.chkShowAssetNo.SetFocus
1630          ElseIf .Parent.chkShowHidden.Enabled = True Then
1640            .Parent.chkShowHidden.SetFocus
1650          Else
1660            .Parent.chkShowAssetTypeDesc.SetFocus
1670          End If
1680        End With
1690      Case vbKeyPageUp, vbKeyHome
1700        intRetVal = 0
1710        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1720      Case vbKeyPageDown, vbKeyEnd
1730        intRetVal = 0
1740        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1750      End Select
1760    End If

        ' ** Ctrl-Shift keys.
1770    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1780      Select Case intRetVal
          Case vbKeyF
1790        With Me
1800          intRetVal = 0
1810          DoCmd.SelectObject acForm, .Parent.Name, False
1820          .Parent.FocusHolder.SetFocus
1830        End With
1840      Case vbKeyTab
1850        With Me
1860          intRetVal = 0
1870          DoCmd.SelectObject acForm, .Parent.Name, False
1880          .Parent.lbxHistoryDate.SetFocus
1890        End With
1900      End Select
1910    End If

EXITP:
1920    KeyCode = intRetVal
1930    Exit Sub

ERRH:
1940    intRetVal = 0
1950    Select Case ERR.Number
        Case Else
1960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1970    End Select
1980    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2010    With Me
2020      If .cmdShorten_raised_focus_dots_img.Visible = True Or .cmdShorten_raised_focus_img.Visible = True Then
2030        Select Case blnShorten_Focus
            Case True
2040          .cmdShorten_raised_semifocus_dots_img.Visible = True
2050          .cmdShorten_raised_img.Visible = False
2060        Case False
2070          .cmdShorten_raised_img.Visible = True
2080          .cmdShorten_raised_semifocus_dots_img.Visible = False
2090        End Select
2100        .cmdShorten_raised_focus_img.Visible = False
2110        .cmdShorten_raised_focus_dots_img.Visible = False
2120        .cmdShorten_sunken_focus_dots_img.Visible = False
2130        .cmdShorten_raised_img_dis.Visible = False
2140      End If
2150      If .cmdWiden_raised_focus_dots_img.Visible = True Or .cmdWiden_raised_focus_img.Visible = True Then
2160        Select Case blnWiden_Focus
            Case True
2170          .cmdWiden_raised_semifocus_dots_img.Visible = True
2180          .cmdWiden_raised_img.Visible = False
2190        Case False
2200          .cmdWiden_raised_img.Visible = True
2210          .cmdWiden_raised_semifocus_dots_img.Visible = False
2220        End Select
2230        .cmdWiden_raised_focus_img.Visible = False
2240        .cmdWiden_raised_focus_dots_img.Visible = False
2250        .cmdWiden_sunken_focus_dots_img.Visible = False
2260        .cmdWiden_raised_img_dis.Visible = False
2270      End If
2280    End With

EXITP:
2290    Exit Sub

ERRH:
2300    Select Case ERR.Number
        Case Else
2310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2320    End Select
2330    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

2410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
2420    Exit Sub

ERRH:
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Sub

Private Sub assetno_lbl_DblClick(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_lbl_DblClick"

2510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub assetno_GotFocus()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_GotFocus"

2610    With Me
2620      .assetno.SelLength = 0
2630      .assetno.SelStart = 0
2640    End With

EXITP:
2650    Exit Sub

ERRH:
2660    Select Case ERR.Number
        Case Else
2670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2680    End Select
2690    Resume EXITP

End Sub

Private Sub assetno_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
2710    intShiftDown = (Shift And acShiftMask) > 0
2720    intAltDown = (Shift And acAltMask) > 0
2730    intCtrlDown = (Shift And acCtrlMask) > 0

2740    intRetVal = KeyCode

        ' ** Plain keys.
2750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2770        With Me
2780          intRetVal = 0
2790          .cusip.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2850        With Me
2860          intRetVal = 0
2870          If .CurrentRecord > 1 Then
2880            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2890            .MarketValueCurrent_usd.SetFocus
2900          Else
2910            DoCmd.SelectObject acForm, .Parent.Name, False
2920            .Parent.lbxHistoryDate.SetFocus
2930          End If
2940        End With
2950      End Select
2960    End If

EXITP:
2970    KeyCode = intRetVal
2980    Exit Sub

ERRH:
2990    intRetVal = 0
3000    Select Case ERR.Number
        Case Else
3010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3020    End Select
3030    Resume EXITP

End Sub

Private Sub cusip_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_lbl_DblClick"

3110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub cusip_GotFocus()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_GotFocus"

3210    With Me
3220      .cusip.SelLength = 0
3230      .cusip.SelStart = 0
3240    End With

EXITP:
3250    Exit Sub

ERRH:
3260    Select Case ERR.Number
        Case Else
3270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3280    End Select
3290    Resume EXITP

End Sub

Private Sub cusip_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
3310    intShiftDown = (Shift And acShiftMask) > 0
3320    intAltDown = (Shift And acAltMask) > 0
3330    intCtrlDown = (Shift And acCtrlMask) > 0

3340    intRetVal = KeyCode

        ' ** Plain keys.
3350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3370        With Me
3380          intRetVal = 0
3390          .totdesc.SetFocus
3400        End With
3410      End Select
3420    End If

        ' ** Shift keys.
3430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3450        With Me
3460          intRetVal = 0
3470          If .assetno.Visible = True And .assetno.Enabled = True Then
3480            .assetno.SetFocus
3490          Else
3500            If .CurrentRecord > 1 Then
3510              MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3520              .marketvaluecurrent_str.SetFocus
3530            Else
3540              DoCmd.SelectObject acForm, .Parent.Name, False
3550              .Parent.lbxHistoryDate.SetFocus
3560            End If
3570          End If
3580        End With
3590      End Select
3600    End If

EXITP:
3610    KeyCode = intRetVal
3620    Exit Sub

ERRH:
3630    intRetVal = 0
3640    Select Case ERR.Number
        Case Else
3650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3660    End Select
3670    Resume EXITP

End Sub

Private Sub totdesc_lbl_DblClick(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_DblClick"

3710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub totdesc_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_MouseMove"

3810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub totdesc_GotFocus()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_GotFocus"

3910    With Me
3920      .totdesc.SelLength = 0
3930      .totdesc.SelStart = 0
3940    End With

EXITP:
3950    Exit Sub

ERRH:
3960    Select Case ERR.Number
        Case Else
3970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3980    End Select
3990    Resume EXITP

End Sub

Private Sub totdesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_MouseMove"

4010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Sub assettype_lbl_DblClick(Cancel As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_lbl_DblClick"

4110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Private Sub assettype_GotFocus()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_GotFocus"

4210    With Me
4220      .assettype.SelLength = 0
4230      .assettype.SelStart = 0
4240    End With

EXITP:
4250    Exit Sub

ERRH:
4260    Select Case ERR.Number
        Case Else
4270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4280    End Select
4290    Resume EXITP

End Sub

Private Sub assettype_description_lbl_DblClick(Cancel As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_description_lbl_DblClick"

4310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Private Sub assettype_description_GotFocus()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_description_GotFocus"

4410    With Me
4420      .assettype_description.SelLength = 0
4430      .assettype_description.SelStart = 0
4440    End With

EXITP:
4450    Exit Sub

ERRH:
4460    Select Case ERR.Number
        Case Else
4470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4480    End Select
4490    Resume EXITP

End Sub

Private Sub curr_code_lbl_DblClick(Cancel As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_lbl_DblClick"

4510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub curr_code_GotFocus()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_GotFocus"

4610    With Me
4620      .curr_code.SelLength = 0
4630      .curr_code.SelStart = 0
4640    End With

EXITP:
4650    Exit Sub

ERRH:
4660    Select Case ERR.Number
        Case Else
4670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4680    End Select
4690    Resume EXITP

End Sub

Private Sub yield_entry_lbl_DblClick(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "yield_entry_lbl_DblClick"

4710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4720    Exit Sub

ERRH:
4730    Select Case ERR.Number
        Case Else
4740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4750    End Select
4760    Resume EXITP

End Sub

Private Sub yield_entry_GotFocus()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "yield_entry_GotFocus"

4810    With Me
4820      .yield_entry.SelLength = 0
4830      .yield_entry.SelStart = 0
4840    End With

EXITP:
4850    Exit Sub

ERRH:
4860    Select Case ERR.Number
        Case Else
4870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4880    End Select
4890    Resume EXITP

End Sub

Private Sub marketvaluecurrent_str_lbl_DblClick(Cancel As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "marketvaluecurrent_str_lbl_DblClick"

4910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4920    Exit Sub

ERRH:
4930    Select Case ERR.Number
        Case Else
4940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4950    End Select
4960    Resume EXITP

End Sub

Private Sub marketvaluecurrent_str_GotFocus()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "marketvaluecurrent_str_GotFocus"

5010    With Me
5020      .marketvaluecurrent_str.SelLength = 0
5030      .marketvaluecurrent_str.SelStart = 0
5040    End With

EXITP:
5050    Exit Sub

ERRH:
5060    Select Case ERR.Number
        Case Else
5070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5080    End Select
5090    Resume EXITP

End Sub

Private Sub marketvaluecurrent_str_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "marketvaluecurrent_str_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
5110    intShiftDown = (Shift And acShiftMask) > 0
5120    intAltDown = (Shift And acAltMask) > 0
5130    intCtrlDown = (Shift And acCtrlMask) > 0

5140    intRetVal = KeyCode

        ' ** Plain keys.
5150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5170        With Me
5180          intRetVal = 0
5190          If .MarketValueCurrent_usd.Visible = True And .MarketValueCurrent_usd.Enabled = True Then
5200            .MarketValueCurrent_usd.SetFocus
5210          Else
5220            lngRecsCur = RecCnt  ' ** Function: Below.
5230            If .CurrentRecord < lngRecsCur Then
5240              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5250              If .assetno.Visible = True And .assetno.Enabled = True Then
5260                .assetno.SetFocus
5270              Else
5280                .cusip.SetFocus
5290              End If
5300            Else
5310              DoCmd.SelectObject acForm, .Parent.Name, False
5320              If .Parent.chkIncludeCurrency.Enabled = True Then
5330                .Parent.chkIncludeCurrency.SetFocus
5340              ElseIf .Parent.chkShowAssetNo.Visible = True And .Parent.chkShowAssetNo.Enabled = True Then
5350                .Parent.chkShowAssetNo.SetFocus
5360              ElseIf .Parent.chkShowHidden.Enabled = True Then
5370                .Parent.chkShowHidden.SetFocus
5380              Else
5390                .Parent.chkShowAssetTypeDesc.SetFocus
5400              End If
5410            End If
5420          End If
5430        End With
5440      End Select
5450    End If

        ' ** Shift keys.
5460    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5470      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5480        With Me
5490          intRetVal = 0
5500          .yield_entry.SetFocus
5510        End With
5520      End Select
5530    End If

EXITP:
5540    KeyCode = intRetVal
5550    Exit Sub

ERRH:
5560    intRetVal = 0
5570    Select Case ERR.Number
        Case Else
5580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5590    End Select
5600    Resume EXITP

End Sub

Private Sub marketvaluecurrent_usd_lbl_DblClick(Cancel As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "marketvaluecurrent_usd_lbl_DblClick"

5710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5720    Exit Sub

ERRH:
5730    Select Case ERR.Number
        Case Else
5740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5750    End Select
5760    Resume EXITP

End Sub

Private Sub marketvaluecurrent_usd_GotFocus()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "marketvaluecurrent_usd_GotFocus"

5810    With Me
5820      .MarketValueCurrent_usd.SelLength = 0
5830      .MarketValueCurrent_usd.SelStart = 0
5840    End With

EXITP:
5850    Exit Sub

ERRH:
5860    Select Case ERR.Number
        Case Else
5870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5880    End Select
5890    Resume EXITP

End Sub

Private Sub marketvaluecurrent_usd_KeyDown(KeyCode As Integer, Shift As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "marketvaluecurrent_usd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
5910    intShiftDown = (Shift And acShiftMask) > 0
5920    intAltDown = (Shift And acAltMask) > 0
5930    intCtrlDown = (Shift And acCtrlMask) > 0

5940    intRetVal = KeyCode

        ' ** Plain keys.
5950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5970        With Me
5980          intRetVal = 0
5990          lngRecsCur = RecCnt  ' ** Function: Below.
6000          If .CurrentRecord < lngRecsCur Then
6010            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
6020            If .assetno.Visible = True And .assetno.Enabled = True Then
6030              .assetno.SetFocus
6040            Else
6050              .cusip.SetFocus
6060            End If
6070          Else
6080            DoCmd.SelectObject acForm, .Parent.Name, False
6090            If .Parent.chkIncludeCurrency.Enabled = True Then
6100              .Parent.chkIncludeCurrency.SetFocus
6110            ElseIf .Parent.chkShowAssetNo.Visible = True And .Parent.chkShowAssetNo.Enabled = True Then
6120              .Parent.chkShowAssetNo.SetFocus
6130            ElseIf .Parent.chkShowHidden.Enabled = True Then
6140              .Parent.chkShowHidden.SetFocus
6150            Else
6160              .Parent.chkShowAssetTypeDesc.SetFocus
6170            End If
6180          End If
6190        End With
6200      End Select
6210    End If

        ' ** Shift keys.
6220    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6230      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6240        With Me
6250          intRetVal = 0
6260          .marketvaluecurrent_str.SetFocus
6270        End With
6280      End Select
6290    End If

EXITP:
6300    KeyCode = intRetVal
6310    Exit Sub

ERRH:
6320    intRetVal = 0
6330    Select Case ERR.Number
        Case Else
6340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6350    End Select
6360    Resume EXITP

End Sub

Private Sub cmdShorten_GotFocus()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_GotFocus"

6410    With Me
6420      blnShorten_Focus = True
6430      .cmdShorten_raised_semifocus_dots_img.Visible = True
6440      .cmdShorten_raised_img.Visible = False
6450      .cmdShorten_raised_focus_img.Visible = False
6460      .cmdShorten_raised_focus_dots_img.Visible = False
6470      .cmdShorten_sunken_focus_dots_img.Visible = False
6480      .cmdShorten_raised_img_dis.Visible = False
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

Private Sub cmdShorten_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseDown"

6610    With Me
6620      blnShorten_MouseDown = True
6630      .cmdShorten_sunken_focus_dots_img.Visible = True
6640      .cmdShorten_raised_img.Visible = False
6650      .cmdShorten_raised_semifocus_dots_img.Visible = False
6660      .cmdShorten_raised_focus_img.Visible = False
6670      .cmdShorten_raised_focus_dots_img.Visible = False
6680      .cmdShorten_raised_img_dis.Visible = False
6690    End With

EXITP:
6700    Exit Sub

ERRH:
6710    Select Case ERR.Number
        Case Else
6720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6730    End Select
6740    Resume EXITP

End Sub

Public Sub cmdShorten_Click()
' ** lngTpp * 4 = 60  : 1 grid point on form (my screen is 15 Twips Per Pixel).
' ** 60 * 4 = 240     : 4 grid points, a nice noticeable amount.
' ** 1440 Twips per Inch.

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_Click"

        Dim lngMinusedWidth As Long
        Dim lngDiff1 As Long
        Dim blnSortHere As Boolean
        Dim lngTmp01 As Long

6810    With Me
6820      blnSortHere = False
6830      lngMinusedWidth = lngTpp * WIDEN_MULT
6840      Select Case .Parent.chkShowAssetNo
          Case True
6850        lngTmp01 = lngTotDescAstno_Width
6860      Case False
6870        lngTmp01 = lngTotDesc_Width
6880      End Select
6890      If (.totdesc.Width <= lngTmp01) Then
6900        Beep
6910      ElseIf ((.totdesc.Width - lngMinusedWidth) < IIf(.Parent.chkShowAssetNo = True, lngTotDescAstno_Width, lngTotDesc_Width)) Then
6920        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
6930        .totdesc.Width = IIf(.Parent.chkShowAssetNo = True, lngTotDescAstno_Width, lngTotDesc_Width)
6940        .totdesc_lbl.Width = .totdesc.Width
6950        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
6960        If blnSortHere = True Then
6970          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
6980          .Sort_lbl.Left = (((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width) + (3& * lngTpp))
6990        End If
7000        blnSortHere = False
7010        If .Sort_line.Left = .assettype_lbl.Left Then blnSortHere = True
7020        .assettype.Left = lngTotDescNext_Left
7030        .assettype_lbl.Left = .assettype.Left
7040        .assettype_lbl_line.Left = .assettype_lbl.Left
7050        If blnSortHere = True Then
7060          .Sort_line.Left = .assettype_lbl.Left
7070          .Sort_lbl.Left = (((.assettype_lbl.Left + .assettype_lbl.Width) - lngSortLbl_Width) + (5& * lngTpp))
7080        End If
7090        blnSortHere = False
7100        If .Parent.chkShowAssetTypeDesc = True Then
7110          If .Sort_line.Left = .assettype_description_lbl.Left Then blnSortHere = True
7120          .assettype_description.Left = ((.assettype.Left + .assettype.Width) + (4& * lngTpp))
7130          .assettype_description_lbl.Left = .assettype_description.Left
7140          .assettype_description_lbl_line.Left = .assettype_description_lbl.Left
7150          If blnSortHere = True Then
7160            .Sort_line.Left = .assettype_description_lbl.Left
7170            .Sort_lbl.Left = ((.assettype_description_lbl.Left + .assettype_description_lbl.Width) - lngSortLbl_Width)
7180          End If
7190        End If
7200        blnSortHere = False
7210        If .Parent.chkIncludeCurrency = True Then
7220          If .Sort_line.Left = .curr_code.Left Then blnSortHere = True
7230          Select Case .Parent.chkShowAssetTypeDesc
              Case True
7240            .curr_code.Left = ((.assettype_description.Left + .assettype_description.Width) + (4& * lngTpp))
7250          Case False
7260            .curr_code.Left = ((.assettype.Left + .assettype.Width) + (4& * lngTpp))
7270          End Select
7280          .curr_code_usd_box.Left = (.curr_code.Left - lngTpp)
7290          .curr_code_forex_box.Left = (.curr_code.Left - lngTpp)
7300          .curr_code_lbl.Left = IIf((.curr_code.Left - (4& * lngTpp)) < 0&, 0&, (.curr_code.Left - (4& * lngTpp)))
7310          .curr_code_lbl_line.Left = .curr_code.Left
7320          If blnSortHere = True Then
7330            .Sort_line.Left = .curr_code.Left
7340            .Sort_lbl.Left = (((.curr_code_lbl.Left + .curr_code_lbl.Width) - lngSortLbl_Width) + (8& * lngTpp))
7350          End If
7360        End If
7370        blnSortHere = False
7380        lngDiff1 = (.marketvaluecurrent_str.Left - .yield_entry.Left)
7390        If .Sort_line.Left = .yield_entry.Left Then blnSortHere = True
7400        Select Case .Parent.chkIncludeCurrency
            Case True
7410          .yield_entry.Left = ((.curr_code.Left + .curr_code.Width) + (4& * lngTpp))
7420        Case False
7430          Select Case .Parent.chkShowAssetTypeDesc
              Case True
7440            .yield_entry.Left = ((.assettype_description.Left + .assettype_description.Width) + (4& * lngTpp))
7450          Case False
7460            .yield_entry.Left = ((.assettype.Left + .assettype.Width) + (4& * lngTpp))
7470          End Select
7480        End Select
7490        .yield_pct_lbl.Left = (.yield_entry.Left + lngYieldPct_Offset)
7500        .yield_entry_lbl_line.Left = .yield_entry_lbl.Left
7510        If blnSortHere = True Then
7520          .Sort_line.Left = .yield_entry_lbl.Left
7530          .Sort_lbl.Left = ((.yield_entry_lbl.Left + .yield_entry_lbl.Width) - lngSortLbl_Width)
7540        End If
7550        blnSortHere = False
7560        If .Sort_line.Left = .marketvaluecurrent_str.Left Then blnSortHere = True
7570        .marketvaluecurrent_str.Left = (.yield_entry.Left + lngDiff1)
7580        .marketvaluecurrent_str_usd_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
7590        .marketvaluecurrent_str_forex_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
7600        .marketvaluecurrent_str_lbl.Left = .marketvaluecurrent_str.Left
7610        .marketvaluecurrent_str_lbl_line.Left = .marketvaluecurrent_str_lbl.Left
7620        .marketvaluecurrent_lbl.Left = .marketvaluecurrent_str_lbl.Left
7630        If blnSortHere = True Then
7640          .Sort_line.Left = .marketvaluecurrent_str_lbl.Left
7650          .Sort_lbl.Left = ((.marketvaluecurrent_str_lbl.Left + .marketvaluecurrent_str_lbl.Width) - lngSortLbl_Width)
7660        End If
7670        blnSortHere = False
7680        If .Parent.chkIncludeCurrency = True Then
7690          If .Sort_line.Left = .marketvaluecurrent_usd_lbl.Left Then blnSortHere = True
7700          .MarketValueCurrent_usd.Left = ((.marketvaluecurrent_str.Left + .marketvaluecurrent_str.Width) + (4& * lngTpp))
7710          .marketvaluecurrent_usd_lbl.Left = .MarketValueCurrent_usd.Left
7720          .marketvaluecurrent_usd_lbl_line.Left = .marketvaluecurrent_usd_lbl.Left
7730          If blnSortHere = True Then
7740            .Sort_line.Left = .marketvaluecurrent_usd_lbl.Left
7750            .Sort_lbl.Left = ((.marketvaluecurrent_usd_lbl.Left + .marketvaluecurrent_usd_lbl.Width) - lngSortLbl_Width)
7760          End If
7770        End If
7780        blnSortHere = False
7790        .Width = lngForm_Width
7800      Else
7810        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
7820        .totdesc.Width = (.totdesc.Width - lngMinusedWidth)
7830        .totdesc_lbl.Width = .totdesc.Width
7840        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7850        If blnSortHere = True Then
7860          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
7870          .Sort_lbl.Left = (((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width) + (3& * lngTpp))
7880        End If
7890        blnSortHere = False
7900        If .Sort_line.Left = .assettype_lbl.Left Then blnSortHere = True
7910        .assettype.Left = (.assettype.Left - lngMinusedWidth)
7920        .assettype_lbl.Left = .assettype.Left
7930        .assettype_lbl_line.Left = .assettype_lbl.Left
7940        If blnSortHere = True Then
7950          .Sort_line.Left = .assettype_lbl.Left
7960          .Sort_lbl.Left = (((.assettype_lbl.Left + .assettype_lbl.Width) - lngSortLbl_Width) + (5& * lngTpp))
7970        End If
7980        blnSortHere = False
7990        If .Parent.chkShowAssetTypeDesc = True Then
8000          If .Sort_line.Left = .assettype_description.Left Then blnSortHere = True
8010          .assettype_description.Left = ((.assettype.Left + .assettype.Width) + (4& * lngTpp))
8020          .assettype_description_lbl.Left = .assettype_description.Left
8030          .assettype_description_lbl_line.Left = .assettype_description_lbl.Left
8040          If blnSortHere = True Then
8050            .Sort_line.Left = .assettype_description_lbl.Left
8060            .Sort_lbl.Left = ((.assettype_description_lbl.Left + .assettype_description_lbl.Width) - lngSortLbl_Width)
8070          End If
8080        End If
8090        blnSortHere = False
8100        If .Parent.chkIncludeCurrency = True Then
8110          If .Sort_line.Left = .curr_code.Left Then blnSortHere = True
8120          Select Case .Parent.chkShowAssetTypeDesc
              Case True
8130            .curr_code.Left = ((.assettype_description.Left + .assettype_description.Width) + (4& * lngTpp))
8140          Case False
8150            .curr_code.Left = ((.assettype_lbl.Left + .assettype_lbl.Width) + (4& * lngTpp))
8160          End Select
8170          .curr_code_usd_box.Left = (.curr_code.Left - lngTpp)
8180          .curr_code_forex_box.Left = (.curr_code.Left - lngTpp)
8190          .curr_code_lbl.Left = IIf((.curr_code.Left - (4& * lngTpp)) < 0&, 0&, (.curr_code.Left - (4& * lngTpp)))
8200          .curr_code_lbl_line.Left = .curr_code.Left
8210          If blnSortHere = True Then
8220            .Sort_line.Left = .curr_code.Left
8230            .Sort_lbl.Left = (((.curr_code_lbl.Left + .curr_code_lbl.Width) - lngSortLbl_Width) + (8& * lngTpp))
8240          End If
8250        End If
8260        blnSortHere = False
8270        lngDiff1 = (.marketvaluecurrent_str.Left - .yield_entry.Left)
8280        If .Sort_line.Left = .yield_entry_lbl.Left Then blnSortHere = True
8290        Select Case .Parent.chkIncludeCurrency
            Case True
8300          .yield_entry.Left = ((.curr_code.Left + .curr_code.Width) + (4& * lngTpp))
8310        Case False
8320          Select Case .Parent.chkShowAssetTypeDesc
              Case True
8330            .yield_entry.Left = ((.assettype_description.Left + .assettype_description.Width) + (4& * lngTpp))
8340          Case False
8350            .yield_entry.Left = ((.assettype_lbl.Left + .assettype_lbl.Width) + (4& * lngTpp))
8360          End Select
8370        End Select
8380        .yield_pct_lbl.Left = (.yield_entry.Left + lngYieldPct_Offset)
8390        .yield_entry_lbl.Left = .yield_entry.Left
8400        .yield_entry_lbl_line.Left = .yield_entry_lbl.Left
8410        If blnSortHere = True Then
8420          .Sort_line.Left = .yield_entry_lbl.Left
8430          .Sort_lbl.Left = ((.yield_entry_lbl.Left + .yield_entry_lbl.Width) - lngSortLbl_Width)
8440        End If
8450        blnSortHere = False
8460        If .Sort_line.Left = .marketvaluecurrent_str_lbl.Left Then blnSortHere = True
8470        .marketvaluecurrent_str.Left = (.yield_entry.Left + lngDiff1)
8480        .marketvaluecurrent_str_usd_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
8490        .marketvaluecurrent_str_forex_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
8500        .marketvaluecurrent_str_lbl.Left = .marketvaluecurrent_str.Left
8510        .marketvaluecurrent_str_lbl_line.Left = .marketvaluecurrent_str_lbl.Left
8520        .marketvaluecurrent_lbl.Left = .marketvaluecurrent_str_lbl.Left
8530        If blnSortHere = True Then
8540          .Sort_line.Left = .marketvaluecurrent_str_lbl.Left
8550          .Sort_lbl.Left = ((.marketvaluecurrent_str_lbl.Left + .marketvaluecurrent_str_lbl.Width) - lngSortLbl_Width)
8560        End If
8570        blnSortHere = False
8580        If .Parent.chkIncludeCurrency = True Then
8590          If .Sort_line.Left = .marketvaluecurrent_usd_lbl.Left Then blnSortHere = True
8600          .MarketValueCurrent_usd.Left = ((.marketvaluecurrent_str.Left + .marketvaluecurrent_str.Width) + (4& * lngTpp))
8610          .marketvaluecurrent_usd_lbl.Left = .MarketValueCurrent_usd.Left
8620          .marketvaluecurrent_usd_lbl_line.Left = .marketvaluecurrent_usd_lbl.Left
8630          If blnSortHere = True Then
8640            .Sort_line.Left = .marketvaluecurrent_usd_lbl.Left
8650            .Sort_lbl.Left = ((.marketvaluecurrent_usd_lbl.Left + .marketvaluecurrent_usd_lbl.Width) - lngSortLbl_Width)
8660          End If
8670        End If
8680        blnSortHere = False
8690        .Width = (.Width - lngMinusedWidth)
8700      End If
8710      .FocusHolder.SetFocus
8720      If .cmdWiden.Enabled = False Then
            ' ** Since cmdShorten was clicked, cmdWiden is certainly available.
8730        .cmdWiden.Enabled = True
8740        .cmdWiden_raised_img.Visible = True
8750        .cmdWiden_raised_semifocus_dots_img.Visible = False
8760        .cmdWiden_raised_focus_img.Visible = False
8770        .cmdWiden_raised_focus_dots_img.Visible = False
8780        .cmdWiden_sunken_focus_dots_img.Visible = False
8790        .cmdWiden_raised_img_dis.Visible = False
8800      End If
8810      If (.totdesc.Width <= lngTmp01) Then
            ' ** If there's nowhere to go, disable it.
8820        .cmdShorten.Enabled = False
8830        .cmdShorten_raised_img_dis.Visible = True
8840        .cmdShorten_raised_img.Visible = False
8850        .cmdShorten_raised_semifocus_dots_img.Visible = False
8860        .cmdShorten_raised_focus_img.Visible = False
8870        .cmdShorten_raised_focus_dots_img.Visible = False
8880        .cmdShorten_sunken_focus_dots_img.Visible = False
8890      End If
8900    End With

EXITP:
8910    Exit Sub

ERRH:
8920    Select Case ERR.Number
        Case Else
8930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8940    End Select
8950    Resume EXITP

End Sub

Private Sub cmdShorten_KeyDown(KeyCode As Integer, Shift As Integer)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_KeyDown"

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
9090          .cmdWiden.SetFocus
9100        End With
9110      End Select
9120    End If

        ' ** Shift keys.
9130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9140      Select Case intRetVal
          Case vbKeyTab
9150        With Me
9160          intRetVal = 0
9170          If .marketvaluecurrent_str.Enabled = True Then
9180            .marketvaluecurrent_str.SetFocus
9190          Else
9200            .FocusHolder.SetFocus
9210          End If
9220        End With
9230      End Select
9240    End If

EXITP:
9250    KeyCode = intRetVal
9260    Exit Sub

ERRH:
9270    intRetVal = 0
9280    Select Case ERR.Number
        Case Else
9290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9300    End Select
9310    Resume EXITP

End Sub

Private Sub cmdShorten_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseMove"

9410    With Me
9420      If blnShorten_MouseDown = False Then
9430        Select Case blnShorten_Focus
            Case True
9440          .cmdShorten_raised_focus_dots_img.Visible = True
9450          .cmdShorten_raised_focus_img.Visible = False
9460        Case False
9470          .cmdShorten_raised_focus_img.Visible = True
9480          .cmdShorten_raised_focus_dots_img.Visible = False
9490        End Select
9500        .cmdShorten_raised_img.Visible = False
9510        .cmdShorten_raised_semifocus_dots_img.Visible = False
9520        .cmdShorten_sunken_focus_dots_img.Visible = False
9530        .cmdShorten_raised_img_dis.Visible = False
9540      End If
9550      If .cmdWiden_raised_focus_dots_img.Visible = True Or .cmdWiden_raised_focus_img.Visible = True Then
9560        Select Case blnWiden_Focus
            Case True
9570          .cmdWiden_raised_semifocus_dots_img.Visible = True
9580          .cmdWiden_raised_img.Visible = False
9590        Case False
9600          .cmdWiden_raised_img.Visible = True
9610          .cmdWiden_raised_semifocus_dots_img.Visible = False
9620        End Select
9630        .cmdWiden_raised_focus_img.Visible = False
9640        .cmdWiden_raised_focus_dots_img.Visible = False
9650        .cmdWiden_sunken_focus_dots_img.Visible = False
9660        .cmdWiden_raised_img_dis.Visible = False
9670      End If
9680    End With

EXITP:
9690    Exit Sub

ERRH:
9700    Select Case ERR.Number
        Case Else
9710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9720    End Select
9730    Resume EXITP

End Sub

Private Sub cmdShorten_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseUp"

9810    With Me
9820      .cmdShorten_raised_focus_dots_img.Visible = True
9830      .cmdShorten_raised_img.Visible = False
9840      .cmdShorten_raised_semifocus_dots_img.Visible = False
9850      .cmdShorten_raised_focus_img.Visible = False
9860      .cmdShorten_sunken_focus_dots_img.Visible = False
9870      .cmdShorten_raised_img_dis.Visible = False
9880      blnShorten_MouseDown = False
9890    End With

EXITP:
9900    Exit Sub

ERRH:
9910    Select Case ERR.Number
        Case Else
9920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9930    End Select
9940    Resume EXITP

End Sub

Private Sub cmdShorten_LostFocus()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_LostFocus"

10010   With Me
10020     .cmdShorten_raised_img.Visible = True
10030     .cmdShorten_raised_semifocus_dots_img.Visible = False
10040     .cmdShorten_raised_focus_img.Visible = False
10050     .cmdShorten_raised_focus_dots_img.Visible = False
10060     .cmdShorten_sunken_focus_dots_img.Visible = False
10070     .cmdShorten_raised_img_dis.Visible = False
10080     blnShorten_Focus = False
10090   End With

EXITP:
10100   Exit Sub

ERRH:
10110   Select Case ERR.Number
        Case Else
10120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10130   End Select
10140   Resume EXITP

End Sub

Private Sub cmdWiden_GotFocus()

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_GotFocus"

10210   With Me
10220     blnWiden_Focus = True
10230     .cmdWiden_raised_semifocus_dots_img.Visible = True
10240     .cmdWiden_raised_img.Visible = False
10250     .cmdWiden_raised_focus_img.Visible = False
10260     .cmdWiden_raised_focus_dots_img.Visible = False
10270     .cmdWiden_sunken_focus_dots_img.Visible = False
10280     .cmdWiden_raised_img_dis.Visible = False
10290   End With

EXITP:
10300   Exit Sub

ERRH:
10310   Select Case ERR.Number
        Case Else
10320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10330   End Select
10340   Resume EXITP

End Sub

Private Sub cmdWiden_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseDown"

10410   With Me
10420     blnWiden_MouseDown = True
10430     .cmdWiden_sunken_focus_dots_img.Visible = True
10440     .cmdWiden_raised_img.Visible = False
10450     .cmdWiden_raised_semifocus_dots_img.Visible = False
10460     .cmdWiden_raised_focus_img.Visible = False
10470     .cmdWiden_raised_focus_dots_img.Visible = False
10480     .cmdWiden_raised_img_dis.Visible = False
10490   End With

EXITP:
10500   Exit Sub

ERRH:
10510   Select Case ERR.Number
        Case Else
10520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10530   End Select
10540   Resume EXITP

End Sub

Public Sub cmdWiden_Click()
' ** lngTpp * 4 = 60  : 1 grid point on form (my screen is 15 Twips Per Pixel).
' ** 60 * 4 = 240     : 4 grid points, a nice noticeable amount.
' ** 1440 Twips per Inch.
' ** Arbitrarily cap it at 10 in. wide.

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_Click"

        Dim lngAddedWidth As Long
        Dim blnSortHere As Boolean

        'totdesc +3
        'curr_code + 8
10610   With Me
10620     blnSortHere = False
10630     lngAddedWidth = lngTpp * WIDEN_MULT
10640     If (.Width >= WIDEN_MAX) Or ((.Width + lngAddedWidth) > WIDEN_MAX) Then
10650       Beep
10660     Else
10670       .Width = (.Width + lngAddedWidth)
10680       If .Sort_line.Left = .assettype_lbl.Left Then blnSortHere = True
10690       .assettype.Left = (.assettype.Left + lngAddedWidth)
10700       .assettype_lbl.Left = .assettype.Left
10710       .assettype_lbl_line.Left = .assettype_lbl.Left
10720       If blnSortHere = True Then
10730         .Sort_line.Left = .assettype_lbl.Left
10740         .Sort_lbl.Left = (((.assettype_lbl.Left + .assettype_lbl.Width) - lngSortLbl_Width) + (5& * lngTpp))
10750       End If
10760       blnSortHere = False
10770       If .Parent.chkShowAssetTypeDesc = True Then
10780         If .Sort_line.Left = .assettype_description_lbl.Left Then blnSortHere = True
10790         .assettype_description.Left = (.assettype_description.Left + lngAddedWidth)
10800         .assettype_description_lbl.Left = .assettype_description.Left
10810         .assettype_description_lbl_line.Left = .assettype_description_lbl.Left
10820         If blnSortHere = True Then
10830           .Sort_line.Left = .assettype_description_lbl.Left
10840           .Sort_lbl.Left = ((.assettype_description_lbl.Left + .assettype_description_lbl.Width) - lngSortLbl_Width)
10850         End If
10860       End If
10870       blnSortHere = False
10880       If .Parent.chkIncludeCurrency = True Then
10890         If .Sort_line.Left = .curr_code.Left Then blnSortHere = True
10900         .curr_code.Left = (.curr_code.Left + lngAddedWidth)
10910         .curr_code_usd_box.Left = (.curr_code.Left - lngTpp)
10920         .curr_code_forex_box.Left = (.curr_code.Left - lngTpp)
10930         .curr_code_lbl.Left = IIf((.curr_code.Left - (4& * lngTpp)) < 0&, 0&, (.curr_code.Left - (4& * lngTpp)))
10940         .curr_code_lbl_line.Left = .curr_code.Left
10950         If blnSortHere = True Then
10960           .Sort_line.Left = .curr_code.Left
10970           .Sort_lbl.Left = ((.curr_code_lbl.Left + .curr_code_lbl.Width) - lngSortLbl_Width) '+ (0& * lngTpp))
10980         End If
10990       End If
11000       blnSortHere = False
11010       If .Sort_line.Left = .yield_entry_lbl.Left Then blnSortHere = True
11020       .yield_entry.Left = (.yield_entry.Left + lngAddedWidth)
11030       .yield_pct_lbl.Left = (.yield_entry.Left + lngYieldPct_Offset)
11040       .yield_entry_lbl.Left = .yield_entry.Left
11050       .yield_entry_lbl_line.Left = .yield_entry_lbl.Left
11060       If blnSortHere = True Then
11070         .Sort_line.Left = .yield_entry_lbl.Left
11080         .Sort_lbl.Left = ((.yield_entry_lbl.Left + .yield_entry_lbl.Width) - lngSortLbl_Width)
11090       End If
11100       blnSortHere = False
11110       If .Sort_line.Left = .marketvaluecurrent_str_lbl.Left Then blnSortHere = True
11120       .marketvaluecurrent_str.Left = (.marketvaluecurrent_str.Left + lngAddedWidth)
11130       .marketvaluecurrent_str_usd_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
11140       .marketvaluecurrent_str_forex_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
11150       .marketvaluecurrent_str_lbl.Left = .marketvaluecurrent_str.Left
11160       .marketvaluecurrent_str_lbl_line.Left = .marketvaluecurrent_str_lbl.Left
11170       .marketvaluecurrent_lbl.Left = .marketvaluecurrent_str_lbl.Left
11180       If blnSortHere = True Then
11190         .Sort_line.Left = .marketvaluecurrent_str_lbl.Left
11200         .Sort_lbl.Left = ((.marketvaluecurrent_str_lbl.Left + .marketvaluecurrent_str_lbl.Width) - lngSortLbl_Width)
11210       End If
11220       blnSortHere = False
11230       If .Parent.chkIncludeCurrency = True Then
11240         If .Sort_line.Left = .marketvaluecurrent_usd_lbl.Left Then blnSortHere = True
11250         .MarketValueCurrent_usd.Left = (.MarketValueCurrent_usd.Left + lngAddedWidth)
11260         .marketvaluecurrent_usd_lbl.Left = .MarketValueCurrent_usd.Left
11270         .marketvaluecurrent_usd_lbl_line.Left = .marketvaluecurrent_usd_lbl.Left
11280         If blnSortHere = True Then
11290           .Sort_line.Left = .marketvaluecurrent_usd_lbl.Left
11300           .Sort_lbl.Left = ((.marketvaluecurrent_usd_lbl.Left + .marketvaluecurrent_usd_lbl.Width) - lngSortLbl_Width)
11310         End If
11320       End If
11330       blnSortHere = False
11340       If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
11350       .totdesc.Width = (.totdesc.Width + lngAddedWidth)
11360       .totdesc_lbl.Width = .totdesc.Width
11370       .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
11380       If blnSortHere = True Then
11390         .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
11400         .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
11410       End If
11420       blnSortHere = False
11430     End If
11440     .FocusHolder.SetFocus
11450     If .cmdShorten.Enabled = False Then
            ' ** Since cmdWiden was clicked, cmdShorten is certainly available.
11460       .cmdShorten.Enabled = True
11470       .cmdShorten_raised_img.Visible = True
11480       .cmdShorten_raised_semifocus_dots_img.Visible = False
11490       .cmdShorten_raised_focus_img.Visible = False
11500       .cmdShorten_raised_focus_dots_img.Visible = False
11510       .cmdShorten_sunken_focus_dots_img.Visible = False
11520       .cmdShorten_raised_img_dis.Visible = False
11530     End If
          ' ** If there's nowhere to go, disable it.
11540     If (.Width >= WIDEN_MAX) Or ((.Width + lngAddedWidth) > WIDEN_MAX) Then
11550       .cmdWiden.Enabled = False
11560       .cmdWiden_raised_img_dis.Visible = True
11570       .cmdWiden_raised_img.Visible = False
11580       .cmdWiden_raised_semifocus_dots_img.Visible = False
11590       .cmdWiden_raised_focus_img.Visible = False
11600       .cmdWiden_raised_focus_dots_img.Visible = False
11610       .cmdWiden_sunken_focus_dots_img.Visible = False
11620     End If
11630   End With

EXITP:
11640   Exit Sub

ERRH:
11650   Select Case ERR.Number
        Case Else
11660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11670   End Select
11680   Resume EXITP

End Sub

Private Sub cmdWiden_KeyDown(KeyCode As Integer, Shift As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11720   intShiftDown = (Shift And acShiftMask) > 0
11730   intAltDown = (Shift And acAltMask) > 0
11740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11760     Select Case intRetVal
          Case vbKeyTab
11770       With Me
11780         intRetVal = 0
11790         If .marketvaluecurrent_str.Enabled = True Then
11800           .marketvaluecurrent_str.SetFocus
11810         Else
11820           .FocusHolder.SetFocus
11830         End If
11840       End With
11850     End Select
11860   End If

        ' ** Shift keys.
11870   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11880     Select Case intRetVal
          Case vbKeyTab
11890       With Me
11900         intRetVal = 0
11910         .cmdShorten.SetFocus
11920       End With
11930     End Select
11940   End If

EXITP:
11950   KeyCode = intRetVal
11960   Exit Sub

ERRH:
11970   intRetVal = 0
11980   Select Case ERR.Number
        Case Else
11990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12000   End Select
12010   Resume EXITP

End Sub

Private Sub cmdWiden_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseMove"

12110   With Me
12120     If blnWiden_MouseDown = False Then
12130       Select Case blnWiden_Focus
            Case True
12140         .cmdWiden_raised_focus_dots_img.Visible = True
12150         .cmdWiden_raised_focus_img.Visible = False
12160       Case False
12170         .cmdWiden_raised_focus_img.Visible = True
12180         .cmdWiden_raised_focus_dots_img.Visible = False
12190       End Select
12200       .cmdWiden_raised_img.Visible = False
12210       .cmdWiden_raised_semifocus_dots_img.Visible = False
12220       .cmdWiden_sunken_focus_dots_img.Visible = False
12230       .cmdWiden_raised_img_dis.Visible = False
12240     End If
12250     If .cmdShorten_raised_focus_dots_img.Visible = True Or .cmdShorten_raised_focus_img.Visible = True Then
12260       Select Case blnShorten_Focus
            Case True
12270         .cmdShorten_raised_semifocus_dots_img.Visible = True
12280         .cmdShorten_raised_img.Visible = False
12290       Case False
12300         .cmdShorten_raised_img.Visible = True
12310         .cmdShorten_raised_semifocus_dots_img.Visible = False
12320       End Select
12330       .cmdShorten_raised_focus_img.Visible = False
12340       .cmdShorten_raised_focus_dots_img.Visible = False
12350       .cmdShorten_sunken_focus_dots_img.Visible = False
12360       .cmdShorten_raised_img_dis.Visible = False
12370     End If
12380   End With

EXITP:
12390   Exit Sub

ERRH:
12400   Select Case ERR.Number
        Case Else
12410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12420   End Select
12430   Resume EXITP

End Sub

Private Sub cmdWiden_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseUp"

12510   With Me
12520     .cmdWiden_raised_focus_dots_img.Visible = True
12530     .cmdWiden_raised_img.Visible = False
12540     .cmdWiden_raised_semifocus_dots_img.Visible = False
12550     .cmdWiden_raised_focus_img.Visible = False
12560     .cmdWiden_sunken_focus_dots_img.Visible = False
12570     .cmdWiden_raised_img_dis.Visible = False
12580     blnWiden_MouseDown = False
12590   End With

EXITP:
12600   Exit Sub

ERRH:
12610   Select Case ERR.Number
        Case Else
12620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12630   End Select
12640   Resume EXITP

End Sub

Private Sub cmdWiden_LostFocus()

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_LostFocus"

12710   With Me
12720     .cmdWiden_raised_img.Visible = True
12730     .cmdWiden_raised_semifocus_dots_img.Visible = False
12740     .cmdWiden_raised_focus_img.Visible = False
12750     .cmdWiden_raised_focus_dots_img.Visible = False
12760     .cmdWiden_sunken_focus_dots_img.Visible = False
12770     .cmdWiden_raised_img_dis.Visible = False
12780     blnWiden_Focus = False
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

Public Sub ShowAssetNo_Sub(blnShow As Boolean)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAssetNo_Sub"

        Dim lngSubform_Width As Long
        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim lngTmp01 As Long

12910   With Me
12920     lngSubform_Width = .Width
12930     blnSortHere = False
12940     Select Case blnShow
          Case True
12950       If .assetno.Visible = False Then
12960         If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
12970         lngTmp01 = .totdesc.Width
12980         .totdesc.Width = (lngTmp01 - lngAssetNo_Width) 'lngTotDescAstno_Width
12990         .totdesc_lbl.Width = .totdesc.Width
13000         .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
13010         .totdesc.Left = (.totdesc.Left + lngAssetNo_Width)
13020         .totdesc_lbl.Left = .totdesc.Left
13030         .totdesc_lbl_line.Left = .totdesc_lbl.Left
13040         If blnSortHere = True Then
13050           .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
13060           .Sort_line.Left = .totdesc_lbl.Left
13070         End If
13080         blnSortHere = False
13090         If .Sort_line.Left = .cusip_lbl.Left Then blnSortHere = True
13100         .cusip.Left = (.cusip.Left + lngAssetNo_Width)
13110         .cusip_lbl.Left = .cusip.Left
13120         .cusip_lbl_line.Left = .cusip_lbl.Left
13130         If blnSortHere = True Then
13140           .Sort_line.Left = .cusip_lbl.Left
13150           .Sort_lbl.Left = ((.cusip_lbl.Left + .cusip_lbl.Width) - lngSortLbl_Width)
13160         End If
13170         blnSortHere = False
13180         .assetno.Visible = True
13190         .assetno_lbl.Visible = True
13200         .assetno_lbl2.Visible = True
13210         .assetno_lbl_line.Visible = True
13220       End If
13230     Case False
13240       If .assetno.Visible = True Then
13250         If .Sort_line.Left = .assetno_lbl.Left Then blnSortHere = True
13260         .assetno.Visible = False
13270         .assetno_lbl.Visible = False
13280         .assetno_lbl2.Visible = False
13290         .assetno_lbl_line.Visible = False
13300         If blnSortHere = True Then
13310           blnResort = True
13320         End If
13330         blnSortHere = False
13340         If .Sort_line.Left = .cusip_lbl.Left Then blnSortHere = True
13350         .cusip.Left = .assetno.Left
13360         .cusip_lbl.Left = .cusip.Left
13370         .cusip_lbl_line.Left = .cusip_lbl.Left
13380         If blnSortHere = True Then
13390           .Sort_line.Left = .cusip_lbl.Left
13400           .Sort_lbl.Left = ((.cusip_lbl.Left + .cusip_lbl.Width) - lngSortLbl_Width)
13410         End If
13420         blnSortHere = False
13430         lngTmp01 = .totdesc.Width
13440         If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
13450         .totdesc.Left = (.totdesc.Left - lngAssetNo_Width)
13460         .totdesc_lbl.Left = .totdesc.Left
13470         .totdesc_lbl_line.Left = .totdesc_lbl.Left
13480         .totdesc.Width = (lngTmp01 + lngAssetNo_Width)
13490         .totdesc_lbl.Width = .totdesc.Width
13500         .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
13510         If blnSortHere = True Then
13520           .Sort_line.Left = .totdesc_lbl.Left
13530           .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
13540         End If
13550         blnSortHere = False
13560       End If
13570     End Select
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

Public Sub ShowAssetTypeDesc_Sub(blnShow As Boolean)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAssetTypeDesc_Sub"

        Dim lngNewCurrCode_Left As Long, lngNewYield_Left As Long, lngNewPctLbl_Left As Long
        Dim lngNewMktValCurStr_Left As Long, lngNewMktValCurUSD_Left As Long, lngSubform_Width As Long, lngNewSubform_Width As Long
        Dim blnSortHere As Boolean, blnResort As Boolean

13710   With Me
13720     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
13730       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
13740     End If
13750     lngSubform_Width = .Width
13760     blnSortHere = False
13770     Select Case blnShow
          Case True
13780       If .assettype_description.Visible = False Then
13790         Select Case .Parent.chkIncludeCurrency
              Case True
13800           lngNewCurrCode_Left = (.curr_code.Left + lngAssetTypeDesc_Width)
13810         Case False
13820           lngNewCurrCode_Left = 0&
13830         End Select
13840         lngNewYield_Left = (.yield_entry.Left + lngAssetTypeDesc_Width)
13850         lngNewPctLbl_Left = (lngNewYield_Left + lngYieldPct_Offset)
13860         lngNewMktValCurStr_Left = (.marketvaluecurrent_str.Left + lngAssetTypeDesc_Width)
13870         Select Case .Parent.chkIncludeCurrency
              Case True
13880           lngNewMktValCurUSD_Left = (.MarketValueCurrent_usd.Left + lngAssetTypeDesc_Width)
13890         Case False
13900           lngNewMktValCurUSD_Left = 0&
13910         End Select
13920         lngNewSubform_Width = (lngSubform_Width + lngAssetTypeDesc_Width)
13930         .Width = lngNewSubform_Width
13940         If .Sort_line.Left = .marketvaluecurrent_usd_lbl.Left And .marketvaluecurrent_usd_lbl.Visible = True Then blnSortHere = True
13950         .MarketValueCurrent_usd.Left = lngNewMktValCurUSD_Left
13960         .marketvaluecurrent_usd_lbl.Left = .MarketValueCurrent_usd.Left
13970         .marketvaluecurrent_usd_lbl_line.Left = .marketvaluecurrent_usd_lbl.Left
13980         If blnSortHere = True Then
13990           .Sort_line.Left = .marketvaluecurrent_usd_lbl.Left
14000           .Sort_lbl.Left = ((.marketvaluecurrent_usd_lbl.Left + .marketvaluecurrent_usd_lbl.Width) - lngSortLbl_Width)
14010         End If
14020         blnSortHere = False
14030         If .Sort_line.Left = .marketvaluecurrent_str_lbl.Left Then blnSortHere = True
14040         .marketvaluecurrent_str.Left = lngNewMktValCurStr_Left
14050         .marketvaluecurrent_str_usd_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
14060         .marketvaluecurrent_str_forex_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
14070         .marketvaluecurrent_str_lbl.Left = .marketvaluecurrent_str.Left
14080         .marketvaluecurrent_lbl.Left = .marketvaluecurrent_str_lbl.Left
14090         If .Parent.chkIncludeCurrency = True Then
14100           .marketvaluecurrent_lbl_line.Left = .marketvaluecurrent_lbl.Left
14110         End If
14120         .marketvaluecurrent_str_lbl_line.Left = .marketvaluecurrent_str_lbl.Left
14130         If blnSortHere = True Then
14140           .Sort_line.Left = .marketvaluecurrent_str_lbl.Left
14150           .Sort_lbl.Left = ((.marketvaluecurrent_str_lbl.Left + .marketvaluecurrent_str_lbl.Width) - lngSortLbl_Width)
14160         End If
14170         blnSortHere = False
14180         If .Sort_line.Left = .yield_entry_lbl.Left Then blnSortHere = True
14190         .yield_entry.Left = lngNewYield_Left
14200         .yield_entry_lbl.Left = .yield_entry.Left
14210         .yield_entry_lbl_line.Left = .yield_entry_lbl.Left
14220         .yield_pct_lbl.Left = lngNewPctLbl_Left
14230         blnSortHere = False
14240         If .Sort_line.Left = .curr_code.Left And .curr_code.Visible = True Then blnSortHere = True
14250         .curr_code.Left = lngNewCurrCode_Left
14260         .curr_code_usd_box.Left = IIf(.curr_code.Visible = True, (.curr_code.Left - lngTpp), 0&)
14270         .curr_code_forex_box.Left = IIf(.curr_code.Visible = True, (.curr_code.Left - lngTpp), 0&)
14280         .curr_code_lbl.Left = IIf((.curr_code.Left - (4& * lngTpp)) < 0&, 0&, (.curr_code.Left - (4& * lngTpp)))
14290         .curr_code_lbl_line.Left = .curr_code.Left
14300         If blnSortHere = True Then
14310           .Sort_line.Left = .curr_code.Left
14320           .Sort_lbl.Left = (((.curr_code_lbl.Left + .curr_code_lbl.Width) - lngSortLbl_Width) + (8& * lngTpp))
14330         End If
14340         blnSortHere = False
14350         .assettype_description.Left = ((.assettype.Left + .assettype.Width) + (4& * lngTpp))
14360         .assettype_description_lbl.Left = .assettype_description.Left
14370         .assettype_description_lbl_line.Left = .assettype_description_lbl.Left
14380         .assettype_description.Visible = True
14390         .assettype_description_lbl.Visible = True
14400         .assettype_description_lbl_line.Visible = True
14410       End If
14420     Case False
14430       If .assettype_description.Visible = True Then
14440         blnResort = False
14450         .FocusHolder.SetFocus
14460         Select Case .Parent.chkIncludeCurrency
              Case True
14470           lngNewCurrCode_Left = (.curr_code.Left - lngAssetTypeDesc_Width)
14480         Case False
14490           lngNewCurrCode_Left = 0&
14500         End Select
14510         lngNewYield_Left = (.yield_entry.Left - lngAssetTypeDesc_Width)
14520         lngNewPctLbl_Left = (lngNewYield_Left + lngYieldPct_Offset)
14530         lngNewMktValCurStr_Left = (.marketvaluecurrent_str.Left - lngAssetTypeDesc_Width)
14540         Select Case .Parent.chkIncludeCurrency
              Case True
14550           lngNewMktValCurUSD_Left = (.MarketValueCurrent_usd.Left - lngAssetTypeDesc_Width)
14560         Case False
14570           lngNewMktValCurUSD_Left = 0&
14580         End Select
14590         lngNewSubform_Width = (lngSubform_Width - lngAssetTypeDesc_Width)
14600         If .Sort_line.Left = .assettype_description_lbl.Left Then blnSortHere = True
14610         .assettype_description.Visible = False
14620         .assettype_description_lbl.Visible = False
14630         .assettype_description_lbl_line.Visible = False
14640         .assettype_description.Left = 0&
14650         .assettype_description_lbl.Left = 0&
14660         .assettype_description_lbl_line.Left = 0&
14670         If blnSortHere = True Then
14680           blnResort = True
14690         End If
14700         blnSortHere = False
14710         If .Sort_line.Left = .curr_code_lbl.Left And .curr_code_lbl.Visible = True Then blnSortHere = True
14720         .curr_code.Left = lngNewCurrCode_Left
14730         Select Case .curr_code.Visible
              Case True
14740           If .curr_code.Left > 0& Then
14750             .curr_code_usd_box.Left = (.curr_code.Left - lngTpp)
14760             .curr_code_forex_box.Left = (.curr_code.Left - lngTpp)
14770           End If
14780         Case False
14790           .curr_code_usd_box.Left = 0&
14800           .curr_code_forex_box.Left = 0&
14810         End Select
14820         .curr_code_lbl.Left = IIf((.curr_code.Left - (4& * lngTpp)) < 0&, 0&, (.curr_code.Left - (4& * lngTpp)))
14830         .curr_code_lbl_line.Left = .curr_code.Left
14840         If blnSortHere = True Then
14850           .Sort_line.Left = .curr_code.Left
14860           .Sort_lbl.Left = (((.curr_code_lbl.Left + .curr_code_lbl.Width) - lngSortLbl_Width) + (8& * lngTpp))
14870         End If
14880         blnSortHere = False
14890         If .Sort_line.Left = .yield_entry_lbl.Left Then blnSortHere = True
14900         .yield_entry.Left = lngNewYield_Left
14910         .yield_entry_lbl.Left = .yield_entry.Left
14920         .yield_entry_lbl_line.Left = .yield_entry_lbl.Left
14930         .yield_pct_lbl.Left = lngNewPctLbl_Left
14940         If blnSortHere = True Then
14950           .Sort_line.Left = .yield_entry_lbl.Left
14960           .Sort_lbl.Left = ((.yield_entry_lbl.Left + .yield_entry_lbl.Width) - lngSortLbl_Width)
14970         End If
14980         blnSortHere = False
14990         If .Sort_line.Left = .marketvaluecurrent_str_lbl.Left Then blnSortHere = True
15000         .marketvaluecurrent_str.Left = lngNewMktValCurStr_Left
15010         .marketvaluecurrent_str_usd_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
15020         .marketvaluecurrent_str_forex_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
15030         .marketvaluecurrent_str_lbl.Left = .marketvaluecurrent_str.Left
15040         .marketvaluecurrent_lbl.Left = .marketvaluecurrent_str_lbl.Left
15050         If .Parent.chkIncludeCurrency = True Then
15060           .marketvaluecurrent_lbl_line.Left = .marketvaluecurrent_lbl.Left
15070         End If
15080         .marketvaluecurrent_str_lbl_line.Left = .marketvaluecurrent_str_lbl.Left
15090         If blnSortHere = True Then
15100           .Sort_line.Left = .marketvaluecurrent_str_lbl.Left
15110           .Sort_lbl.Left = ((.marketvaluecurrent_str_lbl.Left + .marketvaluecurrent_str_lbl.Width) - lngSortLbl_Width)
15120         End If
15130         blnSortHere = False
15140         If .Sort_line.Left = .marketvaluecurrent_usd_lbl.Left And .marketvaluecurrent_usd_lbl.Visible = True Then blnSortHere = True
15150         .MarketValueCurrent_usd.Left = lngNewMktValCurUSD_Left
15160         .marketvaluecurrent_usd_lbl.Left = .MarketValueCurrent_usd.Left
15170         .marketvaluecurrent_usd_lbl_line.Left = .marketvaluecurrent_usd_lbl.Left
15180         If blnSortHere = True Then
15190           .Sort_line.Left = .marketvaluecurrent_usd_lbl.Left
15200           .Sort_lbl.Left = ((.marketvaluecurrent_usd_lbl.Left + .marketvaluecurrent_usd_lbl.Width) - lngSortLbl_Width)
15210         End If
15220         blnSortHere = False
15230         .Width = lngNewSubform_Width
15240         If blnResort = True Then
15250           SortNow "Form_Load"  ' ** Procedure: Below.
15260         End If
15270       End If
15280     End Select
15290   End With

EXITP:
15300   Exit Sub

ERRH:
15310   Select Case ERR.Number
        Case Else
15320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15330   End Select
15340   Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnShow As Boolean)

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim lngSubform_Width As Long, lngNewSubform_Width As Long
        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim lngDiff1 As Long
        Dim lngTmp01 As Long

15410   With Me
15420     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
15430       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
15440     End If
15450     lngSubform_Width = .Width
15460     blnSortHere = False
15470     Select Case .Parent.chkShowAssetNo
          Case True
15480       lngDiff1 = (.totdesc.Width - lngTotDescAstno_Width)
15490     Case False
15500       lngDiff1 = (.totdesc.Width - lngTotDesc_Width)
15510     End Select
15520     Select Case blnShow
          Case True
15530       If .curr_code.Visible = False Then
15540         lngTmp01 = (lngSubform_Width + lngDiff1)
15550         lngNewSubform_Width = (lngTmp01 + lngCurrCode_Width + .MarketValueCurrent_usd.Width + (4& * lngTpp))
15560         .Width = lngNewSubform_Width
15570         Select Case .Parent.chkShowAssetTypeDesc
              Case True
15580           .MarketValueCurrent_usd.Left = (lngMktValCurUSD_Left + lngDiff1)
15590           .marketvaluecurrent_usd_lbl.Left = .MarketValueCurrent_usd.Left
15600           .marketvaluecurrent_usd_lbl_line.Left = .marketvaluecurrent_usd_lbl.Left
15610         Case False
15620           .MarketValueCurrent_usd.Left = ((lngMktValCurUSD_Left + lngDiff1) - (.assettype_description.Width + (4& * lngTpp)))
15630           .marketvaluecurrent_usd_lbl.Left = .MarketValueCurrent_usd.Left
15640           .marketvaluecurrent_usd_lbl_line.Left = .marketvaluecurrent_usd_lbl.Left
15650         End Select
15660         .MarketValueCurrent_usd.Visible = True
15670         .marketvaluecurrent_usd_lbl.Visible = True
15680         .marketvaluecurrent_usd_lbl_line.Visible = True
15690         If .Sort_line.Left = .marketvaluecurrent_str_lbl.Left Then blnSortHere = True
15700         .marketvaluecurrent_str.Left = (.marketvaluecurrent_str.Left + lngCurrCode_Width)
15710         .marketvaluecurrent_str_usd_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
15720         .marketvaluecurrent_str_forex_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
15730         .marketvaluecurrent_str_lbl.Left = .marketvaluecurrent_str.Left
15740         .marketvaluecurrent_str_lbl_line.Left = .marketvaluecurrent_str_lbl.Left
15750         If blnSortHere = True Then
15760           .Sort_line.Left = .marketvaluecurrent_str_lbl.Left
15770           .Sort_lbl.Left = ((.marketvaluecurrent_str_lbl.Left + .marketvaluecurrent_str_lbl.Width) - lngSortLbl_Width)
15780         End If
15790         blnSortHere = False
15800         .marketvaluecurrent_lbl.Left = .marketvaluecurrent_str_lbl.Left
15810         .marketvaluecurrent_lbl.Width = ((.marketvaluecurrent_usd_lbl.Left + .marketvaluecurrent_usd_lbl.Width) - _
                .marketvaluecurrent_str_lbl.Left)
15820         .marketvaluecurrent_lbl.Caption = "&Market Unit Value"
15830         .marketvaluecurrent_lbl.Top = lngMarketValueCurrentLbl_Top
15840         .marketvaluecurrent_lbl_line.Left = .marketvaluecurrent_lbl.Left
15850         .marketvaluecurrent_lbl_line.Visible = True
15860         .marketvaluecurrent_str_lbl.Caption = "Local"
15870         If .Sort_line.Left = .yield_entry_lbl.Left Then blnSortHere = True
15880         .yield_entry.Left = .yield_entry.Left + lngCurrCode_Width
15890         .yield_pct_lbl.Left = (.yield_entry.Left + lngYieldPct_Offset)
15900         .yield_entry_lbl.Left = .yield_entry.Left
15910         .yield_entry_lbl_line.Left = .yield_entry_lbl.Left
15920         If blnSortHere = True Then
15930           .Sort_line.Left = .yield_entry_lbl.Left
15940           .Sort_lbl.Left = ((.yield_entry_lbl.Left + .yield_entry_lbl.Width) - lngSortLbl_Width)
15950         End If
15960         blnSortHere = False
15970         Select Case .Parent.chkShowAssetTypeDesc
              Case True
15980           .curr_code.Left = ((.assettype_description.Left + .assettype_description.Width) + (4& * lngTpp))
15990         Case False
16000           .curr_code.Left = ((.assettype.Left + .assettype.Width) + (4& * lngTpp))
16010         End Select
16020         .curr_code_usd_box.Left = (.curr_code.Left - lngTpp)
16030         .curr_code_forex_box.Left = (.curr_code.Left - lngTpp)
16040         .curr_code_lbl.Left = IIf((.curr_code.Left - (4& * lngTpp)) < 0&, 0&, (.curr_code.Left - (4& * lngTpp)))
16050         .curr_code_lbl_line.Left = .curr_code.Left
16060         .curr_code.Visible = True
16070         .curr_code_usd_box.Visible = True
16080         .curr_code_forex_box.Visible = True
16090         .curr_code_lbl.Visible = True
16100         .curr_code_lbl_line.Visible = True
16110       End If
16120     Case False
16130       If .curr_code.Visible = True Then
16140         If .Sort_line.Left = .marketvaluecurrent_usd_lbl.Left Then blnSortHere = True
16150         .MarketValueCurrent_usd.Visible = False
16160         .marketvaluecurrent_usd_lbl.Visible = False
16170         .marketvaluecurrent_usd_lbl_line.Visible = False
16180         .MarketValueCurrent_usd.Left = 0&
16190         .marketvaluecurrent_usd_lbl.Left = 0&
16200         .marketvaluecurrent_usd_lbl_line.Left = 0&
16210         .marketvaluecurrent_lbl_line.Visible = False
16220         .marketvaluecurrent_lbl_line.Left = 0&
16230         .marketvaluecurrent_lbl.Caption = "&Market"
16240         .marketvaluecurrent_lbl.Top = (lngMarketValueCurrentLbl_Top + (7& * lngTpp))
16250         .marketvaluecurrent_lbl.Width = .marketvaluecurrent_str_lbl.Width
16260         .marketvaluecurrent_str_lbl.Caption = "Unit Value"
16270         If blnSortHere = True Then
16280           blnResort = True
16290         End If
16300         blnSortHere = False
16310         If .Sort_line.Left = .curr_code.Left Then blnSortHere = True
16320         .curr_code.Visible = False
16330         .curr_code_usd_box.Visible = False
16340         .curr_code_forex_box.Visible = False
16350         .curr_code_lbl.Visible = False
16360         .curr_code_lbl_line.Visible = False
16370         .curr_code.Left = 0&
16380         .curr_code_usd_box.Left = 0&
16390         .curr_code_forex_box.Left = 0&
16400         .curr_code_lbl.Left = 0&
16410         .curr_code_lbl_line.Left = 0&
16420         If blnSortHere = True Then
16430           blnResort = True
16440         End If
16450         blnSortHere = False
16460         If .Sort_line.Left = .yield_entry_lbl.Left Then blnSortHere = True
16470         .yield_entry.Left = (.yield_entry.Left - lngCurrCode_Width)
16480         .yield_pct_lbl.Left = (.yield_entry.Left + lngYieldPct_Offset)
16490         .yield_entry_lbl.Left = .yield_entry.Left
16500         .yield_entry_lbl_line.Left = .yield_entry_lbl.Left
16510         If blnSortHere = True Then
16520           .Sort_line.Left = .yield_entry_lbl.Left
16530           .Sort_lbl.Left = ((.yield_entry_lbl.Left + .yield_entry_lbl.Width) - lngSortLbl_Width)
16540         End If
16550         blnSortHere = False
16560         If .Sort_line.Left = .marketvaluecurrent_str_lbl.Left Then blnSortHere = True
16570         .marketvaluecurrent_str.Left = (.marketvaluecurrent_str.Left - lngCurrCode_Width)
16580         .marketvaluecurrent_str_usd_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
16590         .marketvaluecurrent_str_forex_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
16600         .marketvaluecurrent_str_lbl.Left = .marketvaluecurrent_str.Left
16610         .marketvaluecurrent_lbl.Left = .marketvaluecurrent_str_lbl.Left
16620         .marketvaluecurrent_str_lbl_line.Left = .marketvaluecurrent_str_lbl.Left
16630         If blnSortHere = True Then
16640           .Sort_line.Left = .marketvaluecurrent_str_lbl.Left
16650           .Sort_lbl.Left = ((.marketvaluecurrent_str_lbl.Left + .marketvaluecurrent_str_lbl.Width) - lngSortLbl_Width)
16660         End If
16670         blnSortHere = False
16680         lngNewSubform_Width = ((.marketvaluecurrent_str.Left + .marketvaluecurrent_str.Width) + (4& * lngTpp))
16690         .Width = lngNewSubform_Width
16700       End If
16710     End Select
16720   End With

EXITP:
16730   Exit Sub

ERRH:
16740   Select Case ERR.Number
        Case Else
16750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16760   End Select
16770   Resume EXITP

End Sub

Public Function RecCnt() As Long

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

16810   lngRetVal = 0&

16820   With Me
16830     Set rst = .RecordsetClone
16840     With rst
16850       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
16860       Else
16870         .MoveLast
16880         lngRetVal = .RecordCount
16890       End If
16900       .Close
16910     End With
16920   End With

EXITP:
16930   Set rst = Nothing
16940   RecCnt = lngRetVal
16950   Exit Function

ERRH:
16960   lngRetVal = 0&
16970   Select Case ERR.Number
        Case Else
16980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16990   End Select
17000   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

17110   Select Case intWhere
        Case acCmdRecordsGoToFirst
17120     DoCmd.RunCommand acCmdRecordsGoToFirst
17130   Case acCmdRecordsGoToLast
17140     DoCmd.RunCommand acCmdRecordsGoToLast
17150   Case acCmdRecordsGoToPrevious
17160     DoCmd.RunCommand acCmdRecordsGoToPrevious
17170   Case acCmdRecordsGoToNext
17180     DoCmd.RunCommand acCmdRecordsGoToNext
17190   Case Else
17200     If IsMissing(varID) = False Then
17210       With Me
17220         Set rst = .RecordsetClone
17230         With rst
17240           .FindFirst "[AssetPricing_ID] = " & CStr(varID)
17250           If .NoMatch = False Then
17260             Me.Bookmark = .Bookmark
17270           End If
17280           .Close
17290         End With
17300       End With
17310     End If
17320   End Select

EXITP:
17330   Set rst = Nothing
17340   Exit Sub

ERRH:
17350   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
17360   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
17370   Case Else
17380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17390   End Select
17400   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim intPos01 As Integer, intCnt As Integer
        Dim strCalled As String, strSortAsc As String
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [totdesc], [assettype]"

17510   With Me
17520     .Controls(strSortLbl).Visible = False
17530     .Controls(strSortLine).Visible = False
17540     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
17550     If strProc = "Form_Load" Then
17560       strCalled = "assettype"
17570       strSortNow = strSortOrig
17580       lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
17590       lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (5& * lngTpp))
17600       lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
17610       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
17620       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
17630       .Controls(strSortLbl).Top = lngSortLbl_Top
17640       .Controls(strSortLbl).Left = lngSortLbl_Left
17650       .Controls(strSortLine).Top = lngSortLine_Top
17660       .Controls(strSortLine).Left = lngSortLine_Left
17670       .Controls(strSortLine).Width = lngSortLine_Width
17680       .Controls(strSortLbl).Caption = strArwUp
17690       .Controls(strSortLbl).ForeColor = CLR_DKBLU
17700     Else
17710       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
17720       lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
17730       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
17740       lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
17750       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
17760       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
17770       .Controls(strSortLbl).Top = lngSortLbl_Top
17780       .Controls(strSortLbl).Left = lngSortLbl_Left
17790       .Controls(strSortLine).Top = lngSortLine_Top
17800       .Controls(strSortLine).Left = lngSortLine_Left
17810       .Controls(strSortLine).Width = lngSortLine_Width
17820       If strCalled = "assettype" Then
17830         strSortAsc = strSortOrig
17840         If strSortNow = strSortAsc Then
17850           strSortNow = "[assettype] DESC, [totdesc] DESC, [cusip] DESC"
17860           .Controls(strSortLbl).Caption = strArwDn
17870           .Controls(strSortLbl).ForeColor = CLR_DKRED
17880         Else
17890           strSortNow = strSortAsc
17900           .Controls(strSortLbl).Caption = strArwUp
17910           .Controls(strSortLbl).ForeColor = CLR_DKBLU
17920         End If
17930         lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
17940         .Controls(strSortLbl).Left = lngSortLbl_Left
17950       Else
17960         .Controls(strSortLbl).Caption = strArwUp
17970         .Controls(strSortLbl).ForeColor = CLR_DKBLU
17980         Select Case strCalled
              Case "assetno"
17990           strSortAsc = "[assetno]"
18000           lngSortLbl_Left = (lngSortLbl_Left + (2& * lngTpp))
18010           .Controls(strSortLbl).Left = lngSortLbl_Left
18020         Case "cusip"
18030           strSortAsc = "[cusip]" & strStdAsc
18040         Case "totdesc"
18050           strSortAsc = "[totdesc], [cusip], [assettype]"
18060           lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
18070           .Controls(strSortLbl).Left = lngSortLbl_Left
18080         Case "assettype_description"
18090           strSortAsc = "[assettype_description], [totdesc], [cusip]"
18100         Case "curr_code"
18110           strSortAsc = "[curr_code]" & strStdAsc
18120           lngSortLine_Left = .Controls(strCalled).Left
18130           lngSortLine_Width = (.Controls(strCalled).Width + lngTpp)
18140           .Controls(strSortLine).Left = lngSortLine_Left
18150           .Controls(strSortLine).Width = lngSortLine_Width
18160           lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
18170           .Controls(strSortLbl).Left = lngSortLbl_Left
18180         Case "yield", "yield_entry"
18190           strSortAsc = "[yield]" & strStdAsc
18200         Case "marketvaluecurrent_str"
18210           strSortAsc = "[marketvaluecurrent]" & strStdAsc
18220         Case "marketvaluecurrent_usd"
18230           strSortAsc = "[marketvaluecurrent_usd]" & strStdAsc
18240         End Select
18250         If strSortNow = strSortAsc Then
18260           intCnt = (CharCnt(strSortAsc, ",") + 1)
18270           Select Case intCnt
                Case 1
18280             strTmp01 = strSortAsc & " DESC"
18290           Case 2
18300             intPos01 = InStr(strSortAsc, ",")
18310             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
18320             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
18330           Case 3
18340             intPos01 = InStr(strSortAsc, ",")
18350             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
18360             strTmp02 = Mid(strSortAsc, intPos01)
18370             intPos01 = InStr(2, strTmp02, ",")
18380             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
18390             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
18400           Case 4
18410             intPos01 = InStr(strSortAsc, ",")
18420             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
18430             strTmp02 = Mid(strSortAsc, intPos01)
18440             intPos01 = InStr(2, strTmp02, ",")
18450             strTmp01 = strTmp01 & Left(strSortAsc, (intPos01 - 1)) & " DESC"
18460             strTmp02 = Mid(strSortAsc, intPos01)
18470             intPos01 = InStr(2, strTmp02, ",")
18480             strTmp01 = strTmp01 & Left(strSortAsc, (intPos01 - 1)) & " DESC"
18490             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
18500           End Select
18510           strSortNow = strTmp01
18520           .Controls(strSortLbl).Caption = strArwDn
18530           .Controls(strSortLbl).ForeColor = CLR_DKRED
18540         Else
18550           strSortNow = strSortAsc
18560         End If
18570       End If
18580     End If
18590     .Controls(strSortLbl).Visible = True
18600     .Controls(strSortLine).Visible = True
18610     .OrderBy = strSortNow
18620     .OrderByOn = True
18630   End With

        ' ** Potential Sorts:
        ' **   [assettype], [totdesc], [cusip]
        ' **   [assetno]
        ' **   [cusip], [totdesc], [assettype]
        ' **   [totdesc], [cusip], [assettype]
        ' **   [assettype_description], [totdesc], [cusip]
        ' **   [curr_code], [totdesc], [assettype]
        ' **   [yield], [totdesc], [assettype]
        ' **   [marketvaluecurrent], [totdesc], [assettype]
        ' **   [marketvaluecurrent_usd], [totdesc], [assettype]

EXITP:
18640   Exit Sub

ERRH:
18650   Select Case ERR.Number
        Case Else
18660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18670   End Select
18680   Resume EXITP

End Sub

Public Function SortNow_Get() As String

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

18710   strRetVal = strSortNow

EXITP:
18720   SortNow_Get = strRetVal
18730   Exit Function

ERRH:
18740   strRetVal = strSortOrig
18750   Select Case ERR.Number
        Case Else
18760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18770   End Select
18780   Resume EXITP

End Function
