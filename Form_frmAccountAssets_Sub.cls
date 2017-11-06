VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountAssets_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountAssets_Sub"

'VGC 04/15/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmAccountAssets}
' **   Shorten:          S {cmdShorten}
' **   Include Currency: U {chkIncludeCurrency on frmAccountAssets}
' **   Widen:            W {cmdWiden}
' **   Exit:             X {cmdClose on frmAccountAssets}

'=IIf(IsNull([totalshareface])=True, "", IIf([totalshareface]=0, "", IIf([totalshareface] - CLng([totalshareface]) = 0, Format([totalshareface],"#,##0"), Format([totalshareface],"#,##0.000")))) & "  " & [totdesc]

' ** AutoRepeat on cmdShorten and cmdWiden works
' ** with the shortcut keys, but not with the mouse.
' ** Also, the disabled image doesn't show with the mouse.
' ** So it's turned off.

Private Const strSortOrig As String = "[description_masterasset]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private blnShorten_Focus As Boolean, blnShorten_MouseDown As Boolean
Private blnWiden_Focus As Boolean, blnWiden_MouseDown As Boolean

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngForm_Width As Long, lngTotDesc_wc_Width As Long, lngTotDesc_woc_Width As Long, lngAssetDesc_left As Long
Private lngShorten_Left As Long, lngWiden_Left As Long
Private lngCurrID_Left As Long, lngCurrIDLbl_Left As Long, lngCurrID_Width As Long, lngVLine_Offset As Long

Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAccountAssets" Then

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170         blnShorten_Focus = False: blnShorten_MouseDown = False
180         blnWiden_Focus = False: blnWiden_MouseDown = False

190         .curr_id_forex_bg.Visible = False
200         lngSortLbl_Width = .Sort_lbl.Width
210         lngForm_Width = .Width
220         lngAssetDesc_left = .assettype_description.Left
230         lngCurrID_Left = .curr_id.Left
240         lngCurrIDLbl_Left = .curr_id_lbl.Left
250         lngCurrID_Width = (.assettype_description.Left - .curr_id.Left)
260         lngTotDesc_wc_Width = .totdesc.Width                            ' ** This is the width with currency!
270         lngTotDesc_woc_Width = (lngTotDesc_wc_Width + lngCurrID_Width)  ' ** This is the width without currency!
280         lngVLine_Offset = (.totdesc.Left - .Detail_vline01.Left)
290         lngShorten_Left = .cmdShorten.Left
300         lngWiden_Left = .cmdWiden.Left

            ' ** Since there's nowhere to go, disable cmdShorten.
310         .cmdShorten.Enabled = False
320         .cmdShorten_raised_img_dis.Visible = True
330         .cmdShorten_raised_img.Visible = False
340         .cmdShorten_raised_semifocus_dots_img.Visible = False
350         .cmdShorten_raised_focus_img.Visible = False
360         .cmdShorten_raised_focus_dots_img.Visible = False
370         .cmdShorten_sunken_focus_dots_img.Visible = False

380       Else
390         Cancel = -1
400       End If
410     End With

420     If Cancel = -1 Then
430       Beep
440       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
450     End If

EXITP:
460     Exit Sub

ERRH:
470     Select Case ERR.Number
        Case Else
480       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
490     End Select
500     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
620     Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

710     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
720     intShiftDown = (Shift And acShiftMask) > 0
730     intAltDown = (Shift And acAltMask) > 0
740     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Shorten:          S {cmdShorten}
        ' **   Widen:            W {cmdWiden}
        ' **   Exit:             X {cmdClose on frmAccountAssets}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyUp
770         intRetVal = 0
780         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
790       Case vbKeyDown
800         intRetVal = 0
810         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
820       Case vbKeyEscape
830         With Me
840           intRetVal = 0
850           .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountAssets.
860         End With
870       End Select
880     End If

        ' ** Alt keys.
890     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
900       Select Case intRetVal
          Case vbKeyS
            ' ** Invoke the shorten-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
910         With Me
920           intRetVal = 0
930           If .cmdShorten.Enabled = True Then
940             cmdShorten_Click  ' ** Procedure: Below.
950           Else
960             Beep
970           End If
980         End With
990       Case vbKeyW
            ' ** Invoke the widen-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
1000        With Me
1010          intRetVal = 0
1020          If .cmdWiden.Enabled = True Then
1030            cmdWiden_Click  ' ** Procedure: Below.
1040          Else
1050            Beep
1060          End If
1070        End With
1080      Case vbKeyX
1090        With Me
1100          intRetVal = 0
1110          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountAssets.
1120        End With
1130      End Select
1140    End If

        ' ** Ctrl keys.
1150    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1170        With Me
1180          intRetVal = 0
1190          DoCmd.SelectObject acForm, .Parent.Name, False
1200          If .Parent.chkIncludeCurrency.Enabled = True Then
1210            .Parent.chkIncludeCurrency.SetFocus
1220          Else
1230            .Parent.cmdClose.SetFocus
1240          End If
1250        End With
1260      Case vbKeyPageUp, vbKeyHome
1270        intRetVal = 0
1280        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1290      Case vbKeyPageDown, vbKeyEnd
1300        intRetVal = 0
1310        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1320      End Select
1330    End If

        ' ** Ctrl-Shift keys.
1340    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1350      Select Case intRetVal
          Case vbKeyF
1360        With Me
1370          intRetVal = 0
1380          DoCmd.SelectObject acForm, .Parent.Name, False
1390          .Parent.FocusHolder.SetFocus
1400        End With
1410      Case vbKeyTab, vbKeyReturn
1420        With Me
1430          intRetVal = 0
1440          DoCmd.SelectObject acForm, .Parent.Name, False
1450          .Parent.cmdClose.SetFocus
1460        End With
1470      End Select
1480    End If

EXITP:
1490    KeyCode = intRetVal
1500    Exit Sub

ERRH:
1510    intRetVal = 0
1520    Select Case ERR.Number
        Case Else
1530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1540    End Select
1550    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1610    With Me
1620      If .cmdShorten_raised_focus_dots_img.Visible = True Or .cmdShorten_raised_focus_img.Visible = True Then
1630        Select Case blnShorten_Focus
            Case True
1640          .cmdShorten_raised_semifocus_dots_img.Visible = True
1650          .cmdShorten_raised_img.Visible = False
1660        Case False
1670          .cmdShorten_raised_img.Visible = True
1680          .cmdShorten_raised_semifocus_dots_img.Visible = False
1690        End Select
1700        .cmdShorten_raised_focus_img.Visible = False
1710        .cmdShorten_raised_focus_dots_img.Visible = False
1720        .cmdShorten_sunken_focus_dots_img.Visible = False
1730        .cmdShorten_raised_img_dis.Visible = False
1740      End If
1750      If .cmdWiden_raised_focus_dots_img.Visible = True Or .cmdWiden_raised_focus_img.Visible = True Then
1760        Select Case blnWiden_Focus
            Case True
1770          .cmdWiden_raised_semifocus_dots_img.Visible = True
1780          .cmdWiden_raised_img.Visible = False
1790        Case False
1800          .cmdWiden_raised_img.Visible = True
1810          .cmdWiden_raised_semifocus_dots_img.Visible = False
1820        End Select
1830        .cmdWiden_raised_focus_img.Visible = False
1840        .cmdWiden_raised_focus_dots_img.Visible = False
1850        .cmdWiden_sunken_focus_dots_img.Visible = False
1860        .cmdWiden_raised_img_dis.Visible = False
1870      End If
1880    End With

EXITP:
1890    Exit Sub

ERRH:
1900    Select Case ERR.Number
        Case Else
1910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1920    End Select
1930    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

2010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
2020    Exit Sub

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2050    End Select
2060    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2120    intShiftDown = (Shift And acShiftMask) > 0
2130    intAltDown = (Shift And acAltMask) > 0
2140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2170        With Me
2180          intRetVal = 0
2190          .totdesc.SetFocus
2200        End With
2210      End Select
2220    End If

        ' ** Shift keys.
2230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2250        With Me
2260          intRetVal = 0
2270          If .CurrentRecord > 1 Then
2280            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2290            .assettype_description.SetFocus
2300          Else
2310            DoCmd.SelectObject acForm, .Parent.Name, False
2320            .Parent.cmdClose.SetFocus
2330          End If
2340        End With
2350      End Select
2360    End If

EXITP:
2370    KeyCode = intRetVal
2380    Exit Sub

ERRH:
2390    intRetVal = 0
2400    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
2410    Case Else
2420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2430    End Select
2440    Resume EXITP

End Sub

Private Sub Sort_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Sort_lbl_MouseMove"

2510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub TotalShareFace_display_lbl_DblClick(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "TotalShareFace_display_lbl_DblClick"

2610    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
2620    Exit Sub

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub TotalShareFace_display_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "TotalShareFace_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2720    intShiftDown = (Shift And acShiftMask) > 0
2730    intAltDown = (Shift And acAltMask) > 0
2740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2770        With Me
2780          intRetVal = 0
2790          .totdesc.SetFocus
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
2890            .assettype_description.SetFocus
2900          Else
2910            DoCmd.SelectObject acForm, .Parent.Name, False
2920            .Parent.cmdClose.SetFocus
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
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
3010    Case Else
3020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3030    End Select
3040    Resume EXITP

End Sub

Private Sub totdesc_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_DblClick"

3110    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub totdesc_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_MouseMove"

3210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub totdesc_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3320    intShiftDown = (Shift And acShiftMask) > 0
3330    intAltDown = (Shift And acAltMask) > 0
3340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3370        With Me
3380          intRetVal = 0
3390          If .curr_id.Visible = True And .curr_id.Enabled = True Then
3400            .curr_id.SetFocus
3410          Else
3420            .assettype_description.SetFocus
3430          End If
3440        End With
3450      End Select
3460    End If

        ' ** Shift keys.
3470    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3480      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3490        With Me
3500          intRetVal = 0
3510          .TotalShareFace_display.SetFocus
3520        End With
3530      End Select
3540    End If

EXITP:
3550    KeyCode = intRetVal
3560    Exit Sub

ERRH:
3570    intRetVal = 0
3580    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
3590    Case Else
3600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3610    End Select
3620    Resume EXITP

End Sub

Private Sub totdesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_MouseMove"

3710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

3810    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub curr_id_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_MouseMove"

3910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub curr_id_Enter()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_Enter"

4010    With Me
4020      If .curr_id = 150& Then
4030        .curr_id.BackColor = CLR_WHT
4040      Else
4050        .curr_id.BackColor = CLR_LTGRN
4060      End If
4070    End With

EXITP:
4080    Exit Sub

ERRH:
4090    Select Case ERR.Number
        Case Else
4100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4110    End Select
4120    Resume EXITP

End Sub

Private Sub curr_id_KeyDown(KeyCode As Integer, Shift As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4220    intShiftDown = (Shift And acShiftMask) > 0
4230    intAltDown = (Shift And acAltMask) > 0
4240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4270        With Me
4280          intRetVal = 0
4290          .assettype_description.SetFocus
4300        End With
4310      End Select
4320    End If

        ' ** Shift keys.
4330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4350        With Me
4360          intRetVal = 0
4370          .totdesc.SetFocus
4380        End With
4390      End Select
4400    End If

EXITP:
4410    KeyCode = intRetVal
4420    Exit Sub

ERRH:
4430    intRetVal = 0
4440    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
4450    Case Else
4460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4470    End Select
4480    Resume EXITP

End Sub

Private Sub curr_id_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_MouseMove"

4510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub assettype_description_lbl_DblClick(Cancel As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_description_lbl_DblClick"

4610    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Sub

Private Sub assettype_description_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_description_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
4770        With Me
4780          intRetVal = 0
4790          lngRecsCur = RecCnt  ' ** Function: Below.
4800          If .CurrentRecord < lngRecsCur Then
4810            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4820            .TotalShareFace_display.SetFocus
4830          Else
4840            DoCmd.SelectObject acForm, .Parent.Name, False
4850            If .Parent.chkIncludeCurrency.Enabled = True Then
4860              .Parent.chkIncludeCurrency.SetFocus
4870            Else
4880              .Parent.cmdClose.SetFocus
4890            End If
4900          End If
4910        End With
4920      End Select
4930    End If

        ' ** Shift keys.
4940    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4950      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4960        With Me
4970          intRetVal = 0
4980          If .curr_id.Visible = True And .curr_id.Enabled = True Then
4990            .curr_id.SetFocus
5000          Else
5010            .totdesc.SetFocus
5020          End If
5030        End With
5040      End Select
5050    End If

EXITP:
5060    KeyCode = intRetVal
5070    Exit Sub

ERRH:
5080    intRetVal = 0
5090    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
5100    Case Else
5110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5120    End Select
5130    Resume EXITP

End Sub

Private Sub cmdShorten_GotFocus()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_GotFocus"

5210    With Me
5220      blnShorten_Focus = True
5230      .cmdShorten_raised_semifocus_dots_img.Visible = True
5240      .cmdShorten_raised_img.Visible = False
5250      .cmdShorten_raised_focus_img.Visible = False
5260      .cmdShorten_raised_focus_dots_img.Visible = False
5270      .cmdShorten_sunken_focus_dots_img.Visible = False
5280      .cmdShorten_raised_img_dis.Visible = False
5290    End With

EXITP:
5300    Exit Sub

ERRH:
5310    Select Case ERR.Number
        Case Else
5320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5330    End Select
5340    Resume EXITP

End Sub

Private Sub cmdShorten_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseDown"

5410    With Me
5420      blnShorten_MouseDown = True
5430      .cmdShorten_sunken_focus_dots_img.Visible = True
5440      .cmdShorten_raised_img.Visible = False
5450      .cmdShorten_raised_semifocus_dots_img.Visible = False
5460      .cmdShorten_raised_focus_img.Visible = False
5470      .cmdShorten_raised_focus_dots_img.Visible = False
5480      .cmdShorten_raised_img_dis.Visible = False
5490    End With

EXITP:
5500    Exit Sub

ERRH:
5510    Select Case ERR.Number
        Case Else
5520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5530    End Select
5540    Resume EXITP

End Sub

Public Sub cmdShorten_Click()
' ** lngTpp * 4 = 60  : 1 grid point on form (my screen is 15 Twips Per Pixel).
' ** 60 * 4 = 240     : 4 grid points, a nice noticeable amount.
' ** 1440 Twips per Inch.

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_Click"

        Dim lngMinusedWidth As Long, blnSortHere As Boolean
        Dim lngTmp01 As Long

5610    With Me
5620      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5630        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5640      End If
5650      lngMinusedWidth = lngTpp * WIDEN_MULT
5660      blnSortHere = False
5670      Select Case .Parent.chkIncludeCurrency
          Case True
5680        lngTmp01 = lngTotDesc_wc_Width
5690      Case False
5700        lngTmp01 = lngTotDesc_woc_Width
5710      End Select
5720      If (.totdesc.Width <= lngTmp01) Then
5730        Beep
5740      ElseIf ((.totdesc.Width - lngMinusedWidth) < lngTmp01) Then
5750        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
5760        .totdesc.Width = lngTmp01
5770        .totdesc_lbl.Width = .totdesc.Width
5780        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
5790        If blnSortHere = True Then
5800          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
5810          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
5820        End If
5830        blnSortHere = False
5840        If .curr_id.Visible = True Then
5850          If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
5860          .curr_id.Left = lngCurrID_Left
5870          .curr_id_forex_bg.Left = .curr_id.Left
5880          .curr_id_lbl.Left = lngCurrIDLbl_Left
5890          .curr_id_lbl_line.Left = .curr_id.Left
5900          .Detail_vline02.Left = (.curr_id.Left - lngVLine_Offset)
5910          If blnSortHere = True Then
5920            .Sort_line.Left = .curr_id_lbl_line.Left
5930            .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (4& * lngTpp))
5940          End If
5950          blnSortHere = False
5960          If .Sort_line.Left = .assettype_description_lbl.Left Then blnSortHere = True
5970          .assettype_description.Left = lngAssetDesc_left
5980          .assettype_description_lbl.Left = .assettype_description.Left
5990          .assettype_description_lbl_line.Left = .assettype_description_lbl.Left
6000          .Detail_vline03.Left = (.assettype_description.Left - lngVLine_Offset)
6010          If blnSortHere = True Then
6020            .Sort_line.Left = .assettype_description_lbl.Left
6030            .Sort_lbl.Left = ((.assettype_description_lbl.Left + .assettype_description_lbl.Width) - lngSortLbl_Width)
6040          End If
6050          blnSortHere = False
6060        Else
6070          If .Sort_line.Left = .assettype_description_lbl.Left Then blnSortHere = True
6080          .assettype_description.Left = lngAssetDesc_left
6090          .assettype_description_lbl.Left = .assettype_description.Left
6100          .assettype_description_lbl_line.Left = .assettype_description_lbl.Left
6110          .Detail_vline03.Left = (.assettype_description.Left - lngVLine_Offset)
6120          If blnSortHere = True Then
6130            .Sort_line.Left = .assettype_description_lbl.Left
6140            .Sort_lbl.Left = ((.assettype_description_lbl.Left + .assettype_description_lbl.Width) - lngSortLbl_Width)
6150          End If
6160          blnSortHere = False
6170        End If
6180        .Width = lngForm_Width
6190      Else
6200        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
6210        .totdesc.Width = (.totdesc.Width - lngMinusedWidth)
6220        .totdesc_lbl.Width = .totdesc.Width
6230        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
6240        If blnSortHere = True Then
6250          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
6260          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
6270        End If
6280        blnSortHere = False
6290        If .curr_id.Visible = True Then
6300          If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
6310          .curr_id.Left = (.curr_id.Left - lngMinusedWidth)
6320          .curr_id_forex_bg.Left = .curr_id.Left
6330          .curr_id_lbl.Left = (.curr_id_lbl.Left - lngMinusedWidth)
6340          .curr_id_lbl_line.Left = .curr_id.Left
6350          .Detail_vline02.Left = (.curr_id.Left - lngVLine_Offset)
6360          If blnSortHere = True Then
6370            .Sort_line.Left = .curr_id_lbl_line.Left
6380            .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (4& * lngTpp))
6390          End If
6400          blnSortHere = False
6410          If .Sort_line.Left = .assettype_description_lbl.Left Then blnSortHere = True
6420          .assettype_description.Left = (.assettype_description.Left - lngMinusedWidth)
6430          .assettype_description_lbl.Left = .assettype_description.Left
6440          .assettype_description_lbl_line.Left = .assettype_description_lbl.Left
6450          .Detail_vline03.Left = (.assettype_description.Left - lngVLine_Offset)
6460          If blnSortHere = True Then
6470            .Sort_line.Left = .assettype_description_lbl.Left
6480            .Sort_lbl.Left = ((.assettype_description_lbl.Left + .assettype_description_lbl.Width) - lngSortLbl_Width)
6490          End If
6500          blnSortHere = False
6510        Else
6520          If .Sort_line.Left = .assettype_description_lbl.Left Then blnSortHere = True
6530          .assettype_description.Left = (.assettype_description.Left - lngMinusedWidth)
6540          .assettype_description_lbl.Left = .assettype_description.Left
6550          .assettype_description_lbl_line.Left = .assettype_description_lbl.Left
6560          .Detail_vline03.Left = (.assettype_description.Left - lngVLine_Offset)
6570          If blnSortHere = True Then
6580            .Sort_line.Left = .assettype_description_lbl.Left
6590            .Sort_lbl.Left = ((.assettype_description_lbl.Left + .assettype_description_lbl.Width) - lngSortLbl_Width)
6600          End If
6610          blnSortHere = False
6620        End If
6630        .Width = .Width - lngMinusedWidth
6640      End If
6650      .FocusHolder.SetFocus
6660      If .cmdWiden.Enabled = False Then
            ' ** Since cmdShorten was clicked, cmdWiden is certainly available.
6670        .cmdWiden.Enabled = True
6680        .cmdWiden_raised_img.Visible = True
6690        .cmdWiden_raised_semifocus_dots_img.Visible = False
6700        .cmdWiden_raised_focus_img.Visible = False
6710        .cmdWiden_raised_focus_dots_img.Visible = False
6720        .cmdWiden_sunken_focus_dots_img.Visible = False
6730        .cmdWiden_raised_img_dis.Visible = False
6740      End If
6750      If (.totdesc.Width <= lngTmp01) Then
            ' ** If there's nowhere to go, disable it.
6760        .cmdShorten.Enabled = False
6770        .cmdShorten_raised_img_dis.Visible = True
6780        .cmdShorten_raised_img.Visible = False
6790        .cmdShorten_raised_semifocus_dots_img.Visible = False
6800        .cmdShorten_raised_focus_img.Visible = False
6810        .cmdShorten_raised_focus_dots_img.Visible = False
6820        .cmdShorten_sunken_focus_dots_img.Visible = False
6830      End If
6840    End With

EXITP:
6850    Exit Sub

ERRH:
6860    Select Case ERR.Number
        Case Else
6870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6880    End Select
6890    Resume EXITP

End Sub

Private Sub cmdShorten_KeyDown(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6920    intShiftDown = (Shift And acShiftMask) > 0
6930    intAltDown = (Shift And acAltMask) > 0
6940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6960      Select Case intRetVal
          Case vbKeyTab
6970        With Me
6980          intRetVal = 0
6990          If .cmdWiden.Enabled = True Then
7000            .cmdWiden.SetFocus
7010          Else
7020            .FocusHolder.SetFocus
7030          End If
7040        End With
7050      End Select
7060    End If

        ' ** Shift keys.
7070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7080      Select Case intRetVal
          Case vbKeyTab
7090        With Me
7100          intRetVal = 0
7110          .totdesc.SetFocus
7120        End With
7130      End Select
7140    End If

EXITP:
7150    KeyCode = intRetVal
7160    Exit Sub

ERRH:
7170    intRetVal = 0
7180    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
7190    Case Else
7200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7210    End Select
7220    Resume EXITP

End Sub

Private Sub cmdShorten_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseMove"

7310    With Me
7320      If blnShorten_MouseDown = False Then
7330        Select Case blnShorten_Focus
            Case True
7340          .cmdShorten_raised_focus_dots_img.Visible = True
7350          .cmdShorten_raised_focus_img.Visible = False
7360        Case False
7370          .cmdShorten_raised_focus_img.Visible = True
7380          .cmdShorten_raised_focus_dots_img.Visible = False
7390        End Select
7400        .cmdShorten_raised_img.Visible = False
7410        .cmdShorten_raised_semifocus_dots_img.Visible = False
7420        .cmdShorten_sunken_focus_dots_img.Visible = False
7430        .cmdShorten_raised_img_dis.Visible = False
7440      End If
7450      If .cmdWiden_raised_focus_dots_img.Visible = True Or .cmdWiden_raised_focus_img.Visible = True Then
7460        Select Case blnWiden_Focus
            Case True
7470          .cmdWiden_raised_semifocus_dots_img.Visible = True
7480          .cmdWiden_raised_img.Visible = False
7490        Case False
7500          .cmdWiden_raised_img.Visible = True
7510          .cmdWiden_raised_semifocus_dots_img.Visible = False
7520        End Select
7530        .cmdWiden_raised_focus_img.Visible = False
7540        .cmdWiden_raised_focus_dots_img.Visible = False
7550        .cmdWiden_sunken_focus_dots_img.Visible = False
7560        .cmdWiden_raised_img_dis.Visible = False
7570      End If
7580    End With

EXITP:
7590    Exit Sub

ERRH:
7600    Select Case ERR.Number
        Case Else
7610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7620    End Select
7630    Resume EXITP

End Sub

Private Sub cmdShorten_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseUp"

7710    With Me
7720      .cmdShorten_raised_focus_dots_img.Visible = True
7730      .cmdShorten_raised_img.Visible = False
7740      .cmdShorten_raised_semifocus_dots_img.Visible = False
7750      .cmdShorten_raised_focus_img.Visible = False
7760      .cmdShorten_sunken_focus_dots_img.Visible = False
7770      .cmdShorten_raised_img_dis.Visible = False
7780      blnShorten_MouseDown = False
7790    End With

EXITP:
7800    Exit Sub

ERRH:
7810    Select Case ERR.Number
        Case Else
7820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7830    End Select
7840    Resume EXITP

End Sub

Private Sub cmdShorten_LostFocus()

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_LostFocus"

7910    With Me
7920      .cmdShorten_raised_img.Visible = True
7930      .cmdShorten_raised_semifocus_dots_img.Visible = False
7940      .cmdShorten_raised_focus_img.Visible = False
7950      .cmdShorten_raised_focus_dots_img.Visible = False
7960      .cmdShorten_sunken_focus_dots_img.Visible = False
7970      .cmdShorten_raised_img_dis.Visible = False
7980      blnShorten_Focus = False
7990    End With

EXITP:
8000    Exit Sub

ERRH:
8010    Select Case ERR.Number
        Case Else
8020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8030    End Select
8040    Resume EXITP

End Sub

Private Sub cmdWiden_GotFocus()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_GotFocus"

8110    With Me
8120      blnWiden_Focus = True
8130      .cmdWiden_raised_semifocus_dots_img.Visible = True
8140      .cmdWiden_raised_img.Visible = False
8150      .cmdWiden_raised_focus_img.Visible = False
8160      .cmdWiden_raised_focus_dots_img.Visible = False
8170      .cmdWiden_sunken_focus_dots_img.Visible = False
8180      .cmdWiden_raised_img_dis.Visible = False
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

Private Sub cmdWiden_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseDown"

8310    With Me
8320      blnWiden_MouseDown = True
8330      .cmdWiden_sunken_focus_dots_img.Visible = True
8340      .cmdWiden_raised_img.Visible = False
8350      .cmdWiden_raised_semifocus_dots_img.Visible = False
8360      .cmdWiden_raised_focus_img.Visible = False
8370      .cmdWiden_raised_focus_dots_img.Visible = False
8380      .cmdWiden_raised_img_dis.Visible = False
8390    End With

EXITP:
8400    Exit Sub

ERRH:
8410    Select Case ERR.Number
        Case Else
8420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8430    End Select
8440    Resume EXITP

End Sub

Public Sub cmdWiden_Click()
' ** lngTpp * 4 = 60  : 1 grid point on form (my screen is 15 Twips Per Pixel).
' ** 60 * 4 = 240     : 4 grid points, a nice noticeable amount.
' ** 1440 Twips per Inch.
' ** Arbitrarily cap it at 10 in. wide.

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_Click"

        Dim lngAddedWidth As Long, blnSortHere As Boolean

8510    With Me
8520      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
8530        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
8540      End If
8550      lngAddedWidth = lngTpp * WIDEN_MULT
8560      blnSortHere = False
8570      If (.Width >= WIDEN_MAX) Or ((.Width + lngAddedWidth) > WIDEN_MAX) Then
8580        Beep
8590      Else
8600        .Width = .Width + lngAddedWidth
8610        If .Sort_line.Left = .assettype_description_lbl.Left Then blnSortHere = True
8620        .assettype_description.Left = (.assettype_description.Left + lngAddedWidth)
8630        .assettype_description_lbl.Left = .assettype_description.Left
8640        .assettype_description_lbl_line.Left = .assettype_description_lbl.Left
8650        .Detail_vline03.Left = (.assettype_description.Left - lngVLine_Offset)
8660        If blnSortHere = True Then
8670          .Sort_line.Left = (.Sort_line.Left + lngAddedWidth)
8680          .Sort_lbl.Left = (.Sort_lbl.Left + lngAddedWidth)
8690        End If
8700        blnSortHere = False
8710        If .curr_id.Visible = True Then
8720          If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
8730          .curr_id.Left = (.curr_id.Left + lngAddedWidth)
8740          .curr_id_forex_bg.Left = .curr_id.Left
8750          .curr_id_lbl.Left = (.curr_id_lbl.Left + lngAddedWidth)
8760          .curr_id_lbl_line.Left = .curr_id.Left
8770          .Detail_vline02.Left = (.curr_id.Left - lngVLine_Offset)
8780          If blnSortHere = True Then
8790            .Sort_line.Left = .curr_id_lbl_line.Left
8800            .Sort_lbl.Left = (.Sort_lbl.Left + lngAddedWidth)
8810          End If
8820          blnSortHere = False
8830        End If
8840        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
8850        .totdesc.Width = (.totdesc.Width + lngAddedWidth)
8860        .totdesc_lbl.Width = .totdesc.Width
8870        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
8880        If blnSortHere = True Then
8890          .Sort_lbl.Left = (.Sort_lbl.Left + lngAddedWidth)
8900          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
8910        End If
8920        blnSortHere = False
8930      End If
8940      .FocusHolder.SetFocus
8950      If .cmdShorten.Enabled = False Then
            ' ** Since cmdWiden was clicked, cmdShorten is certainly available.
8960        .cmdShorten.Enabled = True
8970        .cmdShorten_raised_img.Visible = True
8980        .cmdShorten_raised_semifocus_dots_img.Visible = False
8990        .cmdShorten_raised_focus_img.Visible = False
9000        .cmdShorten_raised_focus_dots_img.Visible = False
9010        .cmdShorten_sunken_focus_dots_img.Visible = False
9020        .cmdShorten_raised_img_dis.Visible = False
9030      End If
          ' ** If there's nowhere to go, disable it.
9040      If (.Width >= WIDEN_MAX) Or ((.Width + lngAddedWidth) > WIDEN_MAX) Then
9050        .cmdWiden.Enabled = False
9060        .cmdWiden_raised_img_dis.Visible = True
9070        .cmdWiden_raised_img.Visible = False
9080        .cmdWiden_raised_semifocus_dots_img.Visible = False
9090        .cmdWiden_raised_focus_img.Visible = False
9100        .cmdWiden_raised_focus_dots_img.Visible = False
9110        .cmdWiden_sunken_focus_dots_img.Visible = False
9120      End If
9130    End With

EXITP:
9140    Exit Sub

ERRH:
9150    Select Case ERR.Number
        Case Else
9160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9170    End Select
9180    Resume EXITP

End Sub

Private Sub cmdWiden_KeyDown(KeyCode As Integer, Shift As Integer)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9220    intShiftDown = (Shift And acShiftMask) > 0
9230    intAltDown = (Shift And acAltMask) > 0
9240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9260      Select Case intRetVal
          Case vbKeyTab
9270        With Me
9280          intRetVal = 0
9290          .totdesc.SetFocus
9300        End With
9310      End Select
9320    End If

        ' ** Shift keys.
9330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9340      Select Case intRetVal
          Case vbKeyTab
9350        With Me
9360          intRetVal = 0
9370          If .cmdShorten.Enabled = True Then
9380            .cmdShorten.SetFocus
9390          Else
9400            .FocusHolder.SetFocus
9410          End If
9420        End With
9430      End Select
9440    End If

EXITP:
9450    KeyCode = intRetVal
9460    Exit Sub

ERRH:
9470    intRetVal = 0
9480    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
9490    Case Else
9500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9510    End Select
9520    Resume EXITP

End Sub

Private Sub cmdWiden_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseMove"

9610    With Me
9620      If blnWiden_MouseDown = False Then
9630        Select Case blnWiden_Focus
            Case True
9640          .cmdWiden_raised_focus_dots_img.Visible = True
9650          .cmdWiden_raised_focus_img.Visible = False
9660        Case False
9670          .cmdWiden_raised_focus_img.Visible = True
9680          .cmdWiden_raised_focus_dots_img.Visible = False
9690        End Select
9700        .cmdWiden_raised_img.Visible = False
9710        .cmdWiden_raised_semifocus_dots_img.Visible = False
9720        .cmdWiden_sunken_focus_dots_img.Visible = False
9730        .cmdWiden_raised_img_dis.Visible = False
9740      End If
9750      If .cmdShorten_raised_focus_dots_img.Visible = True Or .cmdShorten_raised_focus_img.Visible = True Then
9760        Select Case blnShorten_Focus
            Case True
9770          .cmdShorten_raised_semifocus_dots_img.Visible = True
9780          .cmdShorten_raised_img.Visible = False
9790        Case False
9800          .cmdShorten_raised_img.Visible = True
9810          .cmdShorten_raised_semifocus_dots_img.Visible = False
9820        End Select
9830        .cmdShorten_raised_focus_img.Visible = False
9840        .cmdShorten_raised_focus_dots_img.Visible = False
9850        .cmdShorten_sunken_focus_dots_img.Visible = False
9860        .cmdShorten_raised_img_dis.Visible = False
9870      End If
9880    End With

EXITP:
9890    Exit Sub

ERRH:
9900    Select Case ERR.Number
        Case Else
9910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9920    End Select
9930    Resume EXITP

End Sub

Private Sub cmdWiden_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseUp"

10010   With Me
10020     .cmdWiden_raised_focus_dots_img.Visible = True
10030     .cmdWiden_raised_img.Visible = False
10040     .cmdWiden_raised_semifocus_dots_img.Visible = False
10050     .cmdWiden_raised_focus_img.Visible = False
10060     .cmdWiden_sunken_focus_dots_img.Visible = False
10070     .cmdWiden_raised_img_dis.Visible = False
10080     blnWiden_MouseDown = False
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

Private Sub cmdWiden_LostFocus()

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_LostFocus"

10210   With Me
10220     .cmdWiden_raised_img.Visible = True
10230     .cmdWiden_raised_semifocus_dots_img.Visible = False
10240     .cmdWiden_raised_focus_img.Visible = False
10250     .cmdWiden_raised_focus_dots_img.Visible = False
10260     .cmdWiden_sunken_focus_dots_img.Visible = False
10270     .cmdWiden_raised_img_dis.Visible = False
10280     blnWiden_Focus = False
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

Public Sub IncludeCurrency_Sub(blnShow As Boolean)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim lngTmp01 As Long

10410   With Me
10420     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
10430       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
10440     End If
10450     blnSortHere = False: blnResort = False
10460     Select Case blnShow
          Case True
10470       .FocusHolder.SetFocus
10480       DoEvents
10490       If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
10500       If .totdesc.Width = lngTotDesc_wc_Width Then  ' ** Totdesc hasn't been widened.
10510         lngTmp01 = 0&
10520         .totdesc.Width = lngTotDesc_wc_Width
10530       Else
10540         lngTmp01 = ((.totdesc.Left + (.totdesc.Width - lngCurrID_Width)) + (4& * lngTpp))
10550         .totdesc.Width = (.totdesc.Width - lngCurrID_Width)
10560       End If
10570       .totdesc_lbl.Width = .totdesc.Width
10580       .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
10590       If blnSortHere = True Then
10600         .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
10610         .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
10620       End If
10630       blnSortHere = False
10640       .cmdShorten.Left = lngShorten_Left
10650       .cmdShorten_raised_img.Left = .cmdShorten.Left
10660       .cmdShorten_raised_semifocus_dots_img.Left = .cmdShorten.Left
10670       .cmdShorten_raised_focus_img.Left = .cmdShorten.Left
10680       .cmdShorten_raised_focus_dots_img.Left = .cmdShorten.Left
10690       .cmdShorten_sunken_focus_dots_img.Left = .cmdShorten.Left
10700       .cmdShorten_raised_img_dis.Left = .cmdShorten.Left
10710       .cmdWiden.Left = lngWiden_Left
10720       .cmdWiden_raised_img.Left = .cmdWiden.Left
10730       .cmdWiden_raised_semifocus_dots_img.Left = .cmdWiden.Left
10740       .cmdWiden_raised_focus_img.Left = .cmdWiden.Left
10750       .cmdWiden_raised_focus_dots_img.Left = .cmdWiden.Left
10760       .cmdWiden_sunken_focus_dots_img.Left = .cmdWiden.Left
10770       .cmdWiden_raised_img_dis.Left = .cmdWiden.Left
10780       If lngTmp01 = 0& Then  ' ** Totdesc hasn't been widened.
10790         .curr_id.Left = lngCurrID_Left
10800         .curr_id_lbl.Left = lngCurrIDLbl_Left
10810         .curr_id_lbl_line.Left = .curr_id.Left
10820       Else
10830         .curr_id.Left = lngTmp01
10840         .curr_id_lbl.Left = (.curr_id.Left - (lngCurrID_Left - lngCurrIDLbl_Left))
10850         .curr_id_lbl_line.Left = .curr_id.Left
10860       End If
10870       .curr_id_forex_bg.Left = .curr_id.Left
10880       .curr_id.Visible = True
10890       .curr_id_forex_bg.Visible = True
10900       .curr_id_lbl.Visible = True
10910       .curr_id_lbl_line.Visible = True
10920       .Detail_vline02.Left = (.curr_id.Left - lngVLine_Offset)
10930       .Detail_vline02.Visible = True
10940     Case False
10950       .FocusHolder.SetFocus
10960       DoEvents
10970       If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
10980       .curr_id.Visible = False
10990       .curr_id_forex_bg.Visible = False
11000       .curr_id_lbl.Visible = False
11010       .curr_id_lbl_line.Visible = False
11020       .Detail_vline02.Visible = False
11030       .curr_id.Left = 0&
11040       .curr_id_forex_bg.Left = 0&
11050       .curr_id_lbl.Left = 0&
11060       .curr_id_lbl_line.Left = 0&
11070       .Detail_vline02.Left = 0&
11080       If blnSortHere = True Then
11090         .Sort_line.Visible = False
11100         .Sort_lbl.Visible = False
11110         blnResort = True
11120       End If
11130       blnSortHere = False
11140       If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
11150       .totdesc.Width = ((.assettype_description.Left - (4& * lngTpp)) - .totdesc.Left)
11160       .totdesc_lbl.Width = .totdesc.Width
11170       .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
11180       If blnSortHere = True Then
11190         .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
11200         .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
11210       End If
11220       blnSortHere = False
11230       .cmdShorten.Left = (lngShorten_Left + lngCurrID_Width)
11240       .cmdShorten_raised_img.Left = .cmdShorten.Left
11250       .cmdShorten_raised_semifocus_dots_img.Left = .cmdShorten.Left
11260       .cmdShorten_raised_focus_img.Left = .cmdShorten.Left
11270       .cmdShorten_raised_focus_dots_img.Left = .cmdShorten.Left
11280       .cmdShorten_sunken_focus_dots_img.Left = .cmdShorten.Left
11290       .cmdShorten_raised_img_dis.Left = .cmdShorten.Left
11300       .cmdWiden.Left = (lngWiden_Left + lngCurrID_Width)
11310       .cmdWiden_raised_img.Left = .cmdWiden.Left
11320       .cmdWiden_raised_semifocus_dots_img.Left = .cmdWiden.Left
11330       .cmdWiden_raised_focus_img.Left = .cmdWiden.Left
11340       .cmdWiden_raised_focus_dots_img.Left = .cmdWiden.Left
11350       .cmdWiden_sunken_focus_dots_img.Left = .cmdWiden.Left
11360       .cmdWiden_raised_img_dis.Left = .cmdWiden.Left
11370     End Select
11380     If blnResort = True Then
11390       SortNow "Form_Load"  ' ** Procedure: Below.
11400     Else
            ' ** I don't think it will need to be resorted.
11410     End If
11420   End With

EXITP:
11430   Exit Sub

ERRH:
11440   Select Case ERR.Number
        Case Else
11450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11460   End Select
11470   Resume EXITP

End Sub

Public Function RecCnt() As Long

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

11510   lngRetVal = 0&

11520   With Me
11530     Set rst = .RecordsetClone
11540     With rst
11550       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
11560       Else
11570         .MoveLast
11580         lngRetVal = .RecordCount
11590       End If
11600       .Close
11610     End With
11620   End With

EXITP:
11630   Set rst = Nothing
11640   RecCnt = lngRetVal
11650   Exit Function

ERRH:
11660   lngRetVal = 0&
11670   Select Case ERR.Number
        Case Else
11680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11690   End Select
11700   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

11810   Select Case intWhere
        Case acCmdRecordsGoToFirst
11820     DoCmd.RunCommand acCmdRecordsGoToFirst
11830   Case acCmdRecordsGoToLast
11840     DoCmd.RunCommand acCmdRecordsGoToLast
11850   Case acCmdRecordsGoToPrevious
11860     DoCmd.RunCommand acCmdRecordsGoToPrevious
11870   Case acCmdRecordsGoToNext
11880     DoCmd.RunCommand acCmdRecordsGoToNext
11890   Case Else
11900     If IsMissing(varID) = False Then
11910       With Me
11920         Set rst = .RecordsetClone
11930         With rst
11940           .FindFirst "[assetno] = " & CStr(varID)
11950           If .NoMatch = False Then
11960             Me.Bookmark = .Bookmark
11970           End If
11980           .Close
11990         End With
12000       End With
12010     End If
12020   End Select

EXITP:
12030   Set rst = Nothing
12040   Exit Sub

ERRH:
12050   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
12060   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
12070   Case Else
12080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12090   End Select
12100   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [description_masterasset]"

12210   With Me
12220     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
12230       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
12240     End If
12250     .Controls(strSortLbl).Visible = False
12260     .Controls(strSortLine).Visible = False
12270     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
12280     If strProc = "Form_Load" Then
12290       strCalled = "totdesc"
12300       strSortNow = strSortOrig
12310       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
12320       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
12330       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
12340       .Controls(strSortLbl).Left = lngSortLbl_Left
12350       .Controls(strSortLine).Left = lngSortLine_Left
12360       .Controls(strSortLine).Width = lngSortLine_Width
12370       .Controls(strSortLbl).Caption = strArwUp
12380       .Controls(strSortLbl).ForeColor = CLR_DKBLU
12390     Else
12400       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
12410       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
12420       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
12430       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
12440       .Controls(strSortLbl).Left = lngSortLbl_Left
12450       .Controls(strSortLine).Left = lngSortLine_Left
12460       .Controls(strSortLine).Width = lngSortLine_Width
12470       If strCalled = "totdesc" Then
12480         strSortAsc = strSortOrig
12490         If strSortNow = strSortAsc Then
12500           strSortNow = "[description_masterasset] DESC"
12510           .Controls(strSortLbl).Caption = strArwDn
12520           .Controls(strSortLbl).ForeColor = CLR_DKRED
12530         Else
12540           strSortNow = strSortAsc
12550           .Controls(strSortLbl).Caption = strArwUp
12560           .Controls(strSortLbl).ForeColor = CLR_DKBLU
12570         End If
12580       Else
12590         .Controls(strSortLbl).Caption = strArwUp
12600         .Controls(strSortLbl).ForeColor = CLR_DKBLU
12610         Select Case strCalled
              Case "TotalShareFace_display"
12620           strSortAsc = "[TotalShareFace]" & strStdAsc
12630         Case "curr_id"
12640           strSortAsc = "[curr_code]" & strStdAsc
12650           lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
12660           lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
12670           lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
12680           .Controls(strSortLbl).Left = lngSortLbl_Left
12690           .Controls(strSortLine).Left = lngSortLine_Left
12700           .Controls(strSortLine).Width = lngSortLine_Width
12710         Case "assettype_description"
12720           strSortAsc = "[assettype_description]" & strStdAsc
12730         End Select
12740         If strSortNow = strSortAsc Then
12750           intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
12760           Select Case intCnt
                Case 1
12770             strTmp01 = strSortAsc & " DESC"
12780           Case 2
12790             intPos01 = InStr(strSortAsc, ",")
12800             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
12810             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
12820           Case 3
12830             intPos01 = InStr(strSortAsc, ",")
12840             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
12850             strTmp02 = Mid(strSortAsc, intPos01)
12860             intPos01 = InStr(2, strTmp02, ",")
12870             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
12880             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
12890           End Select
12900           strSortNow = strTmp01
12910           .Controls(strSortLbl).Caption = strArwDn
12920           .Controls(strSortLbl).ForeColor = CLR_DKRED
12930         Else
12940           strSortNow = strSortAsc
12950         End If
12960       End If
12970     End If
12980     .Controls(strSortLbl).Visible = True
12990     .Controls(strSortLine).Visible = True
13000     .OrderBy = strSortNow
13010     .OrderByOn = True
13020   End With

EXITP:
13030   Exit Sub

ERRH:
13040   Select Case ERR.Number
        Case Else
13050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13060   End Select
13070   Resume EXITP

End Sub
