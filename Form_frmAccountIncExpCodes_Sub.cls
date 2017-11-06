VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountIncExpCodes_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountIncExpCodes_Sub"

'VGC 04/15/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Block Assign:     B {cmdBlockAssign on frmAccountIncExpCodes}
' **   Close:            C {cmdClose on frmAccountIncExpCodes}
' **   Include Archive:  H {chkIncludeArchive on frmAccountIncExpCodes}
' **   Include Currency: L {chkIncludeCurrency on frmAccountIncExpCodes}
' **   Show Cost:        O {chkShowCost on frmAccountIncExpCodes}
' **   Shorten:          S {cmdShorten}
' **   Update:           U {cmdUpdate on frmAccountIncExpCodes}
' **   Widen:            W {cmdWiden}
' **   Exit:             X {cmdClose on frmAccountIncExpCodes}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ** AutoRepeat on cmdShorten and cmdWiden works
' ** with the shortcut keys, but not with the mouse.
' ** Also, the disabled image doesn't show with the mouse.
' ** So it's turned off.

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[transdate] DESC, [journaltype], [journalno] DESC"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

' ** revcode_ID combo box constants:
'Private Const CBX_R_ID   As Integer = 0
Private Const CBX_R_DSC  As Integer = 1
Private Const CBX_R_TYP  As Integer = 2
Private Const CBX_R_COD  As Integer = 3
'Private Const CBX_R_TTYP As Integer = 4
'Private Const CBX_R_TCOD As Integer = 5

Private blnShorten_Focus As Boolean, blnShorten_MouseDown As Boolean
Private blnWiden_Focus As Boolean, blnWiden_MouseDown As Boolean

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngForm_Width As Long, lngTotDesc_Width As Long, lngRevCodeID_Left As Long, lngIsIncExp_Offset As Long
Private lngCost_Left As Long, lngCost_Width As Long, lngICash_Left As Long, lngPCash_Left As Long
Private lngCurrID_Left As Long, lngCurrID_Width As Long, lngCurrIDLbl_Offset As Long

Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAccountIncExpCodes" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

180         blnShorten_Focus = False: blnShorten_MouseDown = False
190         blnWiden_Focus = False: blnWiden_MouseDown = False

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
200         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

210         .form_width_line.Left = 0&  ' ** Just get it out of the way.

220         lngSortLbl_Width = .Sort_lbl.Width
230         lngForm_Width = .Width
240         lngTotDesc_Width = .totdesc.Width
250         lngICash_Left = .ICash.Left
260         lngPCash_Left = .PCash.Left
270         lngCost_Left = .Cost.Left
280         lngCost_Width = ((.Cost.Left + .Cost.Width) - (.PCash.Left + .PCash.Width))
290         lngCurrID_Left = .curr_id.Left
300         lngCurrID_Width = ((.curr_id.Left + .curr_id.Width) - (.Cost.Left + .Cost.Width))
310         lngCurrIDLbl_Offset = (.curr_id.Left - .curr_id_lbl.Left)
320         lngRevCodeID_Left = .revcode_ID.Left
330         lngIsIncExp_Offset = (.revcode_ID.Left - .IsIncome.Left)

            ' ** Since there's nowhere to go, disable cmdShorten.
340         .cmdShorten.Enabled = False
350         .cmdShorten_raised_img_dis.Visible = True
360         .cmdShorten_raised_img.Visible = False
370         .cmdShorten_raised_semifocus_dots_img.Visible = False
380         .cmdShorten_raised_focus_img.Visible = False
390         .cmdShorten_raised_focus_dots_img.Visible = False
400         .cmdShorten_sunken_focus_dots_img.Visible = False

410       Else
420         Cancel = -1
430       End If
440     End With

450     If Cancel = -1 Then
460       Beep
470       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
480     End If

EXITP:
490     Exit Sub

ERRH:
500     THAT_PROC = THIS_PROC
510     That_Erl = Erl
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
540     End Select
550     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     With Me

620       SortNow THIS_PROC  ' ** Procedure: Below

630       .transdate.ForeColor = CLR_DISABLED_FG
640       .transdate.BackColor = CLR_DISABLED_BG
650       .journaltype.ForeColor = CLR_DISABLED_FG
660       .journaltype.BackColor = CLR_DISABLED_BG
670       .shareface.ForeColor = CLR_DISABLED_FG
680       .shareface.BackColor = CLR_DISABLED_BG
690       .totdesc.ForeColor = CLR_DISABLED_FG
700       .totdesc.BackColor = CLR_DISABLED_BG
710       .ICash.ForeColor = CLR_DISABLED_FG
720       .ICash.BackColor = CLR_DISABLED_BG
730       .icash_str.ForeColor = CLR_DISABLED_FG
740       .icash_str.BackColor = CLR_DISABLED_BG
750       .PCash.ForeColor = CLR_DISABLED_FG
760       .PCash.BackColor = CLR_DISABLED_BG
770       .pcash_str.ForeColor = CLR_DISABLED_FG
780       .pcash_str.BackColor = CLR_DISABLED_BG
790       .Cost.ForeColor = CLR_DISABLED_FG
800       .Cost.BackColor = CLR_DISABLED_BG
810       .cost_str.ForeColor = CLR_DISABLED_FG
820       .cost_str.BackColor = CLR_DISABLED_BG
830       DoEvents

840     End With

EXITP:
850     Exit Sub

ERRH:
860     Select Case ERR.Number
        Case Else
870       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
880     End Select
890     Resume EXITP

End Sub

Private Sub Form_Current()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

910     With Me
920       Select Case .journaltype
          Case "Dividend", "Interest", "Purchase", "Deposit", "Sold", "Received"
            ' ** INCOME.
930         If .revcode_ID.RowSource <> "qryRevCodeComboI" Then
940           .revcode_ID.RowSource = "qryRevCodeComboI"
950           .revcode_ID.Requery
960         End If
970       Case "Liability", "Paid"
            ' ** EXPENSE.
980         If .revcode_ID.RowSource <> "qryRevCodeComboE" Then
990           .revcode_ID.RowSource = "qryRevCodeComboE"
1000          .revcode_ID.Requery
1010        End If
1020      Case "Withdrawn", "Cost Adj.", "Misc."
            ' ** ALL.
1030        If .revcode_ID.RowSource <> "qryRevCodeComboIE" Then
1040          .revcode_ID.RowSource = "qryRevCodeComboIE"
1050        End If
1060      End Select
1070    End With
1080    DoEvents

EXITP:
1090    Exit Sub

ERRH:
1100    THAT_PROC = THIS_PROC
1110    That_Erl = Erl
1120    Select Case ERR.Number
        Case Else
1130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1140    End Select
1150    Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Shorten:          S {cmdShorten}
        ' **   Widen:            W {cmdWiden}
        ' **   Exit:             X {cmdClose on frmAccountIncExpCodes}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyDown
1270        intRetVal = 0
1280        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1290      Case vbKeyUp
1300        intRetVal = 0
1310        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1320      Case vbKeyEscape
1330        With Me
1340          intRetVal = 0
1350          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountIncExpCodes.
1360        End With
1370      End Select
1380    End If

        ' ** Alt keys.
1390    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1400      Select Case intRetVal
          Case vbKeyS
            ' ** Invoke the shorten-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
1410        With Me
1420          intRetVal = 0
1430          If .cmdShorten.Enabled = True Then
1440            cmdShorten_Click  ' ** Procedure: Below.
1450          Else
1460            Beep
1470          End If
1480        End With
1490      Case vbKeyW
            ' ** Invoke the widen-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
1500        With Me
1510          intRetVal = 0
1520          If .cmdWiden.Enabled = True Then
1530            cmdWiden_Click  ' ** Procedure: Below.
1540          Else
1550            Beep
1560          End If
1570        End With
1580      Case vbKeyX
1590        With Me
1600          intRetVal = 0
1610          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountIncExpCodes.
1620        End With
1630      End Select
1640    End If

        ' ** Ctrl keys.
1650    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1660      Select Case intRetVal
          Case vbKeyS
1670        intRetVal = 0
1680        cmdSave_Click  ' ** Procedure: Below.
1690      Case vbKeyTab
1700        With Me
1710          intRetVal = 0
1720          DoCmd.SelectObject acForm, .Parent.Name, False
1730          If .Parent.chkIncludeCurrency.Enabled = True Then
1740            .Parent.chkIncludeCurrency.SetFocus
1750          ElseIf .Parent.chkIncludeArchive.Enabled = True Then
1760            .Parent.chkIncludeArchive.SetFocus
1770          Else
1780            .Parent.chkShowCost.SetFocus
1790          End If
1800        End With
1810      Case vbKeyPageUp, vbKeyHome
1820        intRetVal = 0
1830        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1840      Case vbKeyPageDown, vbKeyEnd
1850        intRetVal = 0
1860        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1870      End Select
1880    End If

        ' ** Ctrl-Shift keys.
1890    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1900      Select Case intRetVal
          Case vbKeyF
1910        With Me
1920          intRetVal = 0
1930          DoCmd.SelectObject acForm, .Parent.Name, False
1940          .Parent.FocusHolder.SetFocus
1950        End With
1960      Case vbKeyTab
1970        With Me
1980          intRetVal = 0
1990          DoCmd.SelectObject acForm, .Parent.Name, False
2000          .Parent.cmdClose.SetFocus
2010        End With
2020      End Select
2030    End If

EXITP:
2040    KeyCode = intRetVal
2050    Exit Sub

ERRH:
2060    intRetVal = 0
2070    Select Case ERR.Number
        Case 2105
          ' ** Do nothing.
2080    Case Else
2090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2100    End Select
2110    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

2210    With Me
2220      If IsNull(.revcode_ID) = True Then
2230        .revcode_ID = REVID_INC
2240      End If
2250    End With

EXITP:
2260    Exit Sub

ERRH:
2270    THAT_PROC = THIS_PROC
2280    That_Erl = Erl
2290    Select Case ERR.Number
        Case Else
2300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2310    End Select
2320    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim strMsg As String

2410    If DataErr <> 0 Then
2420      Select Case DataErr
          Case 3162  ' ** You tried to assign the Null value to a variable that isn't a Variant data type.
2430        Select Case Me.revcode_TYPE
            Case REVTYP_INC
2440          strMsg = "Income"
2450        Case REVTYP_EXP
2460          strMsg = "Expense"
2470        End Select
2480        MsgBox "The " & strMsg & " code cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
2490        Response = acDataErrContinue
2500      Case Else
2510        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2520        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
2530      End Select
2540    End If

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
2550    Exit Sub

ERRH:
2560    THAT_PROC = THIS_PROC
2570    That_Erl = Erl
2580    Select Case ERR.Number
        Case Else
2590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2600    End Select
2610    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2710    With Me
2720      If .cmdShorten_raised_focus_dots_img.Visible = True Or .cmdShorten_raised_focus_img.Visible = True Then
2730        Select Case blnShorten_Focus
            Case True
2740          .cmdShorten_raised_semifocus_dots_img.Visible = True
2750          .cmdShorten_raised_img.Visible = False
2760        Case False
2770          .cmdShorten_raised_img.Visible = True
2780          .cmdShorten_raised_semifocus_dots_img.Visible = False
2790        End Select
2800        .cmdShorten_raised_focus_img.Visible = False
2810        .cmdShorten_raised_focus_dots_img.Visible = False
2820        .cmdShorten_sunken_focus_dots_img.Visible = False
2830        .cmdShorten_raised_img_dis.Visible = False
2840      End If
2850      If .cmdWiden_raised_focus_dots_img.Visible = True Or .cmdWiden_raised_focus_img.Visible = True Then
2860        Select Case blnWiden_Focus
            Case True
2870          .cmdWiden_raised_semifocus_dots_img.Visible = True
2880          .cmdWiden_raised_img.Visible = False
2890        Case False
2900          .cmdWiden_raised_img.Visible = True
2910          .cmdWiden_raised_semifocus_dots_img.Visible = False
2920        End Select
2930        .cmdWiden_raised_focus_img.Visible = False
2940        .cmdWiden_raised_focus_dots_img.Visible = False
2950        .cmdWiden_sunken_focus_dots_img.Visible = False
2960        .cmdWiden_raised_img_dis.Visible = False
2970      End If
2980    End With

EXITP:
2990    Exit Sub

ERRH:
3000    Select Case ERR.Number
        Case Else
3010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3020    End Select
3030    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

3110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Public Sub cmdSave_Click()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3210    DoCmd.RunCommand acCmdSaveRecord
3220    DoEvents

EXITP:
3230    Exit Sub

ERRH:
3240    Select Case ERR.Number
        Case Else
3250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3260    End Select
3270    Resume EXITP

End Sub

Private Sub transdate_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

3310    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub journaltype_lbl_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_lbl_DblClick"

3410    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub shareface_lbl_DblClick(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_lbl_DblClick"

3510    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Sub

Private Sub totdesc_lbl_DblClick(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_DblClick"

3610    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub totdesc_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_MouseMove"

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

Private Sub totdesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_MouseMove"

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

Private Sub icash_lbl_DblClick(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "ICash_lbl_DblClick"

3910    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub ICash_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "ICash_lbl_MouseMove"

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

Private Sub icash_str_Enter()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_str_Enter"

4110    With Me
4120      If .curr_id = 150& Then
4130        .icash_str.BackColor = CLR_DISABLED_BG
4140      Else
4150        .icash_str.BackColor = CLR_LTGRN
4160      End If
4170    End With

EXITP:
4180    Exit Sub

ERRH:
4190    Select Case ERR.Number
        Case Else
4200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4210    End Select
4220    Resume EXITP

End Sub

Private Sub pcash_lbl_DblClick(Cancel As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "PCash_lbl_DblClick"

4310    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Private Sub pcash_str_Enter()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_str_Enter"

4410    With Me
4420      If .curr_id = 150& Then
4430        .pcash_str.BackColor = CLR_DISABLED_BG
4440      Else
4450        .pcash_str.BackColor = CLR_LTGRN
4460      End If
4470    End With

EXITP:
4480    Exit Sub

ERRH:
4490    Select Case ERR.Number
        Case Else
4500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4510    End Select
4520    Resume EXITP

End Sub

Private Sub cost_lbl_DblClick(Cancel As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Cost_lbl_DblClick"

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

Private Sub cost_str_Enter()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_str_Enter"

4710    With Me
4720      If .curr_id = 150& Then
4730        .cost_str.BackColor = CLR_DISABLED_BG
4740      Else
4750        .cost_str.BackColor = CLR_LTGRN
4760      End If
4770    End With

EXITP:
4780    Exit Sub

ERRH:
4790    Select Case ERR.Number
        Case Else
4800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4810    End Select
4820    Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

4910    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
4920    Exit Sub

ERRH:
4930    Select Case ERR.Number
        Case Else
4940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4950    End Select
4960    Resume EXITP

End Sub

Private Sub curr_id_Enter()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_Enter"

5010    With Me
5020      If .curr_id = 150& Then
5030        .curr_id.BackColor = CLR_DISABLED_BG
5040      Else
5050        .curr_id.BackColor = CLR_LTGRN
5060      End If
5070    End With

EXITP:
5080    Exit Sub

ERRH:
5090    Select Case ERR.Number
        Case Else
5100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5110    End Select
5120    Resume EXITP

End Sub

Private Sub revcode_ID_lbl_DblClick(Cancel As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ID_lbl_DblClick"

5210    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
5220    Exit Sub

ERRH:
5230    Select Case ERR.Number
        Case Else
5240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5250    End Select
5260    Resume EXITP

End Sub

Private Sub revcode_DESC_display_GotFocus()
' ** This control sits on top of the revcode_ID combo box.
' ** It's used to display the revcode_DESC when the
' ** RowSource is showing the other set.
' ** CONTROL IS ENABLED, BUT LOCKED.

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_display_GotFocus"

        Dim varTmp00 As Variant

5310    With Me
5320  On Error Resume Next
5330      varTmp00 = .revcode_DESC
5340      If ERR.Number <> 0 Then
5350  On Error GoTo ERRH
5360        MsgBox "No records meet the specified criteria.", vbInformation + vbOKOnly, "Not Found"
5370        DoCmd.RunCommand acCmdRemoveFilterSort
5380      Else
5390  On Error GoTo ERRH
5400        If Nz(.revcode_DESC, vbNullString) = vbNullString Then
5410          .revcode_ID.SetFocus
5420        Else
5430          Select Case .revcode_ID.RowSource
              Case "qryRevCodeComboI"
                ' ** INCOME.
5440            If .revcode_ID.Column(CBX_R_TYP) = 2 Then
5450              .revcode_ID = REVID_INC  ' ** Unspecified Income.
5460              .changed = True
5470              .revcode_DESC = vbNullString
5480              .revcode_ID.SetFocus
5490            Else
5500              .revcode_ID.SetFocus
5510            End If
5520          Case "qryRevCodeComboE"
                ' ** EXPENSE.
5530            If .revcode_ID.Column(CBX_R_TYP) = 1 Then
5540              .revcode_ID = REVID_EXP  ' ** Unspecified Expense.
5550              .changed = True
5560              .revcode_DESC = vbNullString
5570              .revcode_ID.SetFocus
5580            Else
5590              .revcode_ID.SetFocus
5600            End If
5610          Case "qryRevCodeComboIE"
                ' ** ALL.
5620            .revcode_ID.SetFocus
5630          End Select
5640        End If
5650      End If
5660    End With

EXITP:
5670    Exit Sub

ERRH:
5680    Select Case ERR.Number
        Case Else
5690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5700    End Select
5710    Resume EXITP

End Sub

Private Sub revcode_DESC_display_AfterUpdate()
' ** This control sits on top of the revcode_ID combo box.
' ** It's used to display the revcode_DESC when the
' ** RowSource is showing the other set.
' ** CONTROL IS ENABLED, BUT LOCKED.

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_display_AfterUpdate"

        ' ** Ordinarily, the GotFocus() event, above, should have quickly moved them off of this field.
        ' ** NOTE: Could they get to the control and paste something
        ' ** quicker than GotFocus() could move them to the real one?

EXITP:
5810    Exit Sub

ERRH:
5820    Select Case ERR.Number
        Case Else
5830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5840    End Select
5850    Resume EXITP

End Sub

Private Sub revcode_ID_AfterUpdate()
' ** revcode_ID:
' **   RowSource is 0-Based:
' **     Col 0: revcode_ID
' **     Col 1: revcode_DESC
' **     Col 2: revcode_TYPE
' **     Col 3: revcode_TYPE_Code (I/E)
' **     Col 4: taxcode_type
' **     Col 5: taxcode_type_Code (I/D)
' **   BoundColumn is 1-Based:
' **     Col 0: ListIndex

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ID_AfterUpdate"

        Dim strRevCode As String, intTaxCode As Integer

5910    With Me
5920      .changed = True
5930      intTaxCode = Nz(.taxcode, 0)
5940      If IsNull(.revcode_ID) = False Then
5950        If .revcode_ID > 0 Then
5960          .revcode_DESC = .revcode_ID.Column(CBX_R_DSC)
5970          If .revcode_ID = 1 Or .revcode_ID = 2 Then
5980            .revcode_DESC_display = vbNullString
5990          Else
6000            .revcode_DESC_display = .revcode_DESC
6010          End If
6020          If gblnLinkRevTaxCodes = True Then
6030            If intTaxCode > 0 Then
6040              strRevCode = .revcode_ID.Column(CBX_R_COD)
6050              If strRevCode = "I" And .taxcode_type = TAXTYP_DED Then
6060                .taxcode = TAXID_INC  ' ** Unspecified Income.
6070              ElseIf strRevCode = "E" And .taxcode_type = TAXTYP_INC Then
6080                .taxcode = TAXID_DED  ' ** Unspecified Deduction.
6090              End If
6100            End If
6110          End If
6120        Else
6130          Select Case .revcode_ID.RowSource
              Case "qryRevCodeComboI"
                ' ** INCOME.
6140            .revcode_ID = REVID_INC
6150            .revcode_DESC = "Unspecified Income"
6160            .revcode_DESC_display = vbNullString
6170          Case "qryRevCodeComboE"
                ' ** EXPENSE.
6180            .revcode_ID = REVID_EXP
6190            .revcode_DESC = "Unspecified Expense"
6200            .revcode_DESC_display = vbNullString
6210          Case "qryRevCodeComboIE"
                ' ** ALL.
6220            If gblnLinkRevTaxCodes = True Then
6230              If intTaxCode > 0 Then
6240                Select Case .taxcode_type
                    Case TAXTYP_INC
                      ' ** INCOME.
6250                  .revcode_ID = REVID_INC
6260                  .revcode_DESC = "Unspecified Income"
6270                  .revcode_DESC_display = vbNullString
6280                Case TAXTYP_DED
                      ' ** EXPENSE.
6290                  .revcode_ID = REVID_EXP
6300                  .revcode_DESC = "Unspecified Expense"
6310                  .revcode_DESC_display = vbNullString
6320                End Select
6330              Else
                    ' ** INCOME.
6340                .revcode_ID = REVID_INC
6350                .revcode_DESC = "Unspecified Income"
6360                .revcode_DESC_display = vbNullString
6370              End If
6380            Else
                  ' ** INCOME.
6390              .revcode_ID = REVID_INC
6400              .revcode_DESC = "Unspecified Income"
6410              .revcode_DESC_display = vbNullString
6420            End If
6430          End Select
6440        End If
6450      Else
6460        Select Case .revcode_ID.RowSource
            Case "qryRevCodeComboI"
              ' ** INCOME.
6470          .revcode_ID = REVID_INC
6480          .revcode_DESC = "Unspecified Income"
6490          .revcode_DESC_display = vbNullString
6500        Case "qryRevCodeComboE"
              ' ** EXPENSE.
6510          .revcode_ID = REVID_EXP
6520          .revcode_DESC = "Unspecified Expense"
6530          .revcode_DESC_display = vbNullString
6540        Case "qryRevCodeComboIE"
              ' ** ALL.
6550          If gblnLinkRevTaxCodes = True Then
6560            If intTaxCode > 0 Then
6570              Select Case .taxcode_type
                  Case TAXTYP_INC
                    ' ** INCOME.
6580                .revcode_ID = REVID_INC
6590                .revcode_DESC = "Unspecified Income"
6600                .revcode_DESC_display = vbNullString
6610              Case TAXTYP_DED
                    ' ** EXPENSE.
6620                .revcode_ID = REVID_EXP
6630                .revcode_DESC = "Unspecified Expense"
6640                .revcode_DESC_display = vbNullString
6650              End Select
6660            Else
                  ' ** INCOME.
6670              .revcode_ID = REVID_INC
6680              .revcode_DESC = "Unspecified Income"
6690              .revcode_DESC_display = vbNullString
6700            End If
6710          Else
                ' ** INCOME.
6720            .revcode_ID = REVID_INC
6730            .revcode_DESC = "Unspecified Income"
6740            .revcode_DESC_display = vbNullString
6750          End If
6760        End Select
6770      End If
6780      cmdSave_Click  ' ** Procedure: Above.
6790      .Parent.chkUpdated = True
6800      .Parent.cmdUpdate.Enabled = True
6810    End With

EXITP:
6820    Exit Sub

ERRH:
6830    THAT_PROC = THIS_PROC
6840    That_Erl = Erl
6850    Select Case ERR.Number
        Case Else
6860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6870    End Select
6880    Resume EXITP

End Sub

Private Sub revcode_ID_KeyDown(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ID_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
6970        With Me
6980          intRetVal = 0
6990          lngRecsCur = RecCnt  ' ** Function: Below.
7000          If .CurrentRecord < lngRecsCur Then
7010            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
7020          Else
7030            DoCmd.SelectObject acForm, .Parent.Name, False
7040            If .Parent.chkIncludeArchive.Enabled = True Then
7050              .Parent.chkIncludeArchive.SetFocus
7060            ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
7070              .Parent.chkIncludeCurrency.SetFocus
7080            Else
7090              .Parent.chkShowCost.SetFocus
7100            End If
7110          End If
7120        End With
7130      End Select
7140    End If

        ' ** Shift keys.
7150    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7170        With Me
7180          intRetVal = 0
7190          If .CurrentRecord > 1 Then
7200            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
7210          Else
7220            DoCmd.SelectObject acForm, .Parent.Name, False
7230            .Parent.cmdClose.SetFocus
7240          End If
7250        End With
7260      End Select
7270    End If

EXITP:
7280    KeyCode = intRetVal
7290    Exit Sub

ERRH:
7300    intRetVal = 0
7310    Select Case ERR.Number
        Case 2105
          ' ** Do nothing.
7320    Case Else
7330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7340    End Select
7350    Resume EXITP

End Sub

Private Sub cmdShorten_GotFocus()

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_GotFocus"

7410    With Me
7420      blnShorten_Focus = True
7430      .cmdShorten_raised_semifocus_dots_img.Visible = True
7440      .cmdShorten_raised_img.Visible = False
7450      .cmdShorten_raised_focus_img.Visible = False
7460      .cmdShorten_raised_focus_dots_img.Visible = False
7470      .cmdShorten_sunken_focus_dots_img.Visible = False
7480      .cmdShorten_raised_img_dis.Visible = False
7490    End With

EXITP:
7500    Exit Sub

ERRH:
7510    Select Case ERR.Number
        Case Else
7520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7530    End Select
7540    Resume EXITP

End Sub

Private Sub cmdShorten_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseDown"

7610    With Me
7620      blnShorten_MouseDown = True
7630      .cmdShorten_sunken_focus_dots_img.Visible = True
7640      .cmdShorten_raised_img.Visible = False
7650      .cmdShorten_raised_semifocus_dots_img.Visible = False
7660      .cmdShorten_raised_focus_img.Visible = False
7670      .cmdShorten_raised_focus_dots_img.Visible = False
7680      .cmdShorten_raised_img_dis.Visible = False
7690    End With

EXITP:
7700    Exit Sub

ERRH:
7710    Select Case ERR.Number
        Case Else
7720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7730    End Select
7740    Resume EXITP

End Sub

Public Sub cmdShorten_Click()
' ** lngTpp * 4 = 60  : 1 grid point on form (my screen is 15 Twips Per Pixel).
' ** 60 * 4 = 240     : 4 grid points, a nice noticeable amount.
' ** 1440 Twips per Inch.

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_Click"

        Dim lngMinusedWidth As Long
        Dim blnSortHere As Boolean, blnCurrID As Boolean, blnCost As Boolean

7810    With Me
7820      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
7830        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
7840      End If
7850      blnSortHere = False
7860      blnCurrID = .Parent.chkIncludeCurrency
7870      blnCost = .Parent.chkShowCost
7880      lngMinusedWidth = lngTpp * WIDEN_MULT
7890      If (.totdesc.Width <= lngTotDesc_Width) Then
7900        Beep
7910      ElseIf ((.totdesc.Width - lngMinusedWidth) < lngTotDesc_Width) Then
7920        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
7930        .totdesc.Width = lngTotDesc_Width
7940        .totdesc_lbl.Width = .totdesc.Width
7950        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7960        If blnSortHere = True Then
7970          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
7980          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
7990        End If
8000        blnSortHere = False
8010        If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
8020        .ICash.Left = lngICash_Left
8030        .icash_box.Left = .ICash.Left
8040        .icash_forex_bg.Left = (.icash_box.Left - lngTpp)
8050        .icash_str.Left = .ICash.Left
8060        .icash_lbl.Left = .ICash.Left
8070        .icash_lbl_line.Left = .icash_lbl.Left
8080        If blnSortHere = True Then
8090          .Sort_line.Left = .icash_lbl.Left
8100          .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
8110        End If
8120        blnSortHere = False
8130        If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
8140        .PCash.Left = lngPCash_Left
8150        .pcash_box.Left = .PCash.Left
8160        .pcash_forex_bg.Left = (.pcash_box.Left - lngTpp)
8170        .pcash_str.Left = .PCash.Left
8180        .pcash_lbl.Left = .PCash.Left
8190        .pcash_lbl_line.Left = .pcash_lbl.Left
8200        If blnSortHere = True Then
8210          .Sort_line.Left = .pcash_lbl.Left
8220          .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
8230        End If
8240        blnSortHere = False
8250        If blnCost = True Then
8260          If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
8270          .Cost.Left = lngCost_Left
8280          .cost_box.Left = .Cost.Left
8290          .cost_forex_bg.Left = (.cost_box.Left - lngTpp)
8300          .cost_str.Left = .Cost.Left
8310          .cost_lbl.Left = .Cost.Left
8320          .cost_lbl_line.Left = .cost_lbl.Left
8330          If blnSortHere = True Then
8340            .Sort_line.Left = .cost_lbl.Left
8350            .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
8360          End If
8370          blnSortHere = False
8380        End If
8390        If blnCurrID = True Then
8400          If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
8410          Select Case blnCost
              Case True
8420            .curr_id.Left = lngCurrID_Left
8430          Case False
8440            .curr_id.Left = (lngCurrID_Left - lngCost_Width)
8450          End Select
8460          .curr_id_box.Left = .curr_id.Left
8470          .curr_id_forex_bg.Left = (.curr_id_box.Left - lngTpp)
8480          .curr_id_lbl.Left = (.curr_id.Left - lngCurrIDLbl_Offset)
8490          .curr_id_lbl_line.Left = .curr_id.Left
8500          If blnSortHere = True Then
8510            .Sort_line.Left = .curr_id_lbl_line.Left
8520            .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
8530          End If
8540          blnSortHere = False
8550        End If
8560        If .Sort_line.Left = .revcode_ID_lbl.Left Then blnSortHere = True
8570        Select Case blnCost
            Case True
8580          Select Case blnCurrID
              Case True
8590            .revcode_ID.Left = lngRevCodeID_Left
8600          Case False
8610            .revcode_ID.Left = (lngRevCodeID_Left - lngCurrID_Width)
8620          End Select
8630        Case False
8640          Select Case blnCurrID
              Case True
8650            .revcode_ID.Left = (lngRevCodeID_Left - lngCost_Width)
8660          Case False
8670            .revcode_ID.Left = (lngRevCodeID_Left - (lngCost_Width + lngCurrID_Width))
8680          End Select
8690        End Select
8700        .revcode_DESC_display = .revcode_ID.Left
8710        .revcode_ID_lbl.Left = .revcode_ID.Left
8720        .revcode_ID_lbl_line.Left = .revcode_ID_lbl.Left
8730        If blnSortHere = True Then
8740          .Sort_line.Left = .revcode_ID_lbl.Left
8750          .Sort_lbl.Left = (((.revcode_ID_lbl.Left + .revcode_ID_lbl.Width) - lngSortLbl_Width) + (4& * lngTpp))
8760        End If
8770        blnSortHere = False
8780        .IsIncome.Left = (.revcode_ID.Left - lngIsIncExp_Offset)
8790        .IsExpense.Left = .IsIncome.Left
8800        .IsArch_bg1.Width = lngForm_Width
8810        .IsArch_bg2.Width = lngForm_Width
8820        .Width = lngForm_Width
8830      Else
8840        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
8850        .totdesc.Width = (.totdesc.Width - lngMinusedWidth)
8860        .totdesc_lbl.Width = .totdesc.Width
8870        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
8880        If blnSortHere = True Then
8890          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
8900          .Sort_lbl = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
8910        End If
8920        blnSortHere = False
8930        If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
8940        .ICash.Left = (.ICash.Left - lngMinusedWidth)
8950        .icash_box.Left = .ICash.Left
8960        .icash_forex_bg.Left = (.icash_box.Left - lngTpp)
8970        .icash_str.Left = .ICash.Left
8980        .icash_lbl.Left = .ICash.Left
8990        .icash_lbl_line.Left = .icash_lbl.Left
9000        If blnSortHere = True Then
9010          .Sort_line.Left = .icash_lbl.Left
9020          .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
9030        End If
9040        blnSortHere = False
9050        If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
9060        .PCash.Left = (.PCash.Left - lngMinusedWidth)
9070        .pcash_box.Left = .PCash.Left
9080        .pcash_forex_bg.Left = (.pcash_box.Left - lngTpp)
9090        .pcash_str.Left = .PCash.Left
9100        .pcash_lbl.Left = .PCash.Left
9110        .pcash_lbl_line.Left = .pcash_lbl.Left
9120        If blnSortHere = True Then
9130          .Sort_line.Left = .pcash_lbl.Left
9140          .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
9150        End If
9160        blnSortHere = False
9170        If blnCost = True Then
9180          If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
9190          .Cost.Left = (.Cost.Left - lngMinusedWidth)
9200          .cost_box.Left = .Cost.Left
9210          .cost_forex_bg.Left = (.cost_box.Left - lngTpp)
9220          .cost_str.Left = .Cost.Left
9230          .cost_lbl.Left = .Cost.Left
9240          .cost_lbl_line.Left = .cost_lbl.Left
9250          If blnSortHere = True Then
9260            .Sort_line.Left = .cost_lbl.Left
9270            .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
9280          End If
9290          blnSortHere = False
9300        End If
9310        If blnCurrID = True Then
9320          If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
9330          .curr_id.Left = (.curr_id.Left - lngMinusedWidth)
9340          .curr_id_box.Left = .curr_id.Left
9350          .curr_id_forex_bg.Left = (.curr_id_box.Left - lngTpp)
9360          .curr_id_lbl.Left = (.curr_id.Left - lngCurrIDLbl_Offset)
9370          .curr_id_lbl_line.Left = .curr_id.Left
9380          If blnSortHere = True Then
9390            .Sort_line.Left = .curr_id_lbl_line.Left
9400            .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
9410          End If
9420          blnSortHere = False
9430        End If
9440        If .Sort_line.Left = .revcode_ID_lbl.Left Then blnSortHere = True
9450        .revcode_ID.Left = (.revcode_ID.Left - lngMinusedWidth)
9460        .revcode_DESC_display.Left = .revcode_ID.Left
9470        .revcode_ID_lbl.Left = .revcode_ID.Left
9480        .revcode_ID_lbl_line.Left = .revcode_ID_lbl.Left
9490        If blnSortHere = True Then
9500          .Sort_line.Left = .revcode_ID_lbl.Left
9510          .Sort_lbl.Left = (((.revcode_ID_lbl.Left + .revcode_ID_lbl.Width) - lngSortLbl_Width) + (4& * lngTpp))
9520        End If
9530        blnSortHere = False
9540        .IsIncome.Left = (.revcode_ID.Left - lngIsIncExp_Offset)
9550        .IsExpense.Left = .IsIncome.Left
9560        .IsArch_bg1.Width = (.Width - lngMinusedWidth)
9570        .IsArch_bg2.Width = (.Width - lngMinusedWidth)
9580        .Width = (.Width - lngMinusedWidth)
9590      End If
9600      .FocusHolder.SetFocus
9610      If .cmdWiden.Enabled = False Then
            ' ** Since cmdShorten was clicked, cmdWiden is certainly available.
9620        .cmdWiden.Enabled = True
9630        .cmdWiden_raised_img.Visible = True
9640        .cmdWiden_raised_semifocus_dots_img.Visible = False
9650        .cmdWiden_raised_focus_img.Visible = False
9660        .cmdWiden_raised_focus_dots_img.Visible = False
9670        .cmdWiden_sunken_focus_dots_img.Visible = False
9680        .cmdWiden_raised_img_dis.Visible = False
9690      End If
9700      If (.totdesc.Width <= lngTotDesc_Width) Then
            ' ** If there's nowhere to go, disable it.
9710        .cmdShorten.Enabled = False
9720        .cmdShorten_raised_img_dis.Visible = True
9730        .cmdShorten_raised_img.Visible = False
9740        .cmdShorten_raised_semifocus_dots_img.Visible = False
9750        .cmdShorten_raised_focus_img.Visible = False
9760        .cmdShorten_raised_focus_dots_img.Visible = False
9770        .cmdShorten_sunken_focus_dots_img.Visible = False
9780      End If
9790    End With

EXITP:
9800    Exit Sub

ERRH:
9810    Select Case ERR.Number
        Case Else
9820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9830    End Select
9840    Resume EXITP

End Sub

Private Sub cmdShorten_KeyDown(KeyCode As Integer, Shift As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_KeyDown"

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
9990          .cmdWiden.SetFocus
10000       End With
10010     End Select
10020   End If

        ' ** Shift keys.
10030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10040     Select Case intRetVal
          Case vbKeyTab
10050       With Me
10060         intRetVal = 0
10070 On Error Resume Next
10080         .revcode_DESC_display.SetFocus
10090         If ERR.Number <> 0 Then
10100 On Error GoTo ERRH
10110           .revcode_ID.SetFocus
10120         Else
10130 On Error GoTo ERRH
10140         End If
10150       End With
10160     End Select
10170   End If

EXITP:
10180   KeyCode = intRetVal
10190   Exit Sub

ERRH:
10200   intRetVal = 0
10210   Select Case ERR.Number
        Case 2105
          ' ** Do nothing.
10220   Case Else
10230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10240   End Select
10250   Resume EXITP

End Sub

Private Sub cmdShorten_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseMove"

10310   With Me
10320     If blnShorten_MouseDown = False Then
10330       Select Case blnShorten_Focus
            Case True
10340         .cmdShorten_raised_focus_dots_img.Visible = True
10350         .cmdShorten_raised_focus_img.Visible = False
10360       Case False
10370         .cmdShorten_raised_focus_img.Visible = True
10380         .cmdShorten_raised_focus_dots_img.Visible = False
10390       End Select
10400       .cmdShorten_raised_img.Visible = False
10410       .cmdShorten_raised_semifocus_dots_img.Visible = False
10420       .cmdShorten_sunken_focus_dots_img.Visible = False
10430       .cmdShorten_raised_img_dis.Visible = False
10440     End If
10450     If .cmdWiden_raised_focus_dots_img.Visible = True Or .cmdWiden_raised_focus_img.Visible = True Then
10460       Select Case blnWiden_Focus
            Case True
10470         .cmdWiden_raised_semifocus_dots_img.Visible = True
10480         .cmdWiden_raised_img.Visible = False
10490       Case False
10500         .cmdWiden_raised_img.Visible = True
10510         .cmdWiden_raised_semifocus_dots_img.Visible = False
10520       End Select
10530       .cmdWiden_raised_focus_img.Visible = False
10540       .cmdWiden_raised_focus_dots_img.Visible = False
10550       .cmdWiden_sunken_focus_dots_img.Visible = False
10560       .cmdWiden_raised_img_dis.Visible = False
10570     End If
10580   End With

EXITP:
10590   Exit Sub

ERRH:
10600   Select Case ERR.Number
        Case Else
10610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10620   End Select
10630   Resume EXITP

End Sub

Private Sub cmdShorten_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseUp"

10710   With Me
10720     .cmdShorten_raised_focus_dots_img.Visible = True
10730     .cmdShorten_raised_img.Visible = False
10740     .cmdShorten_raised_semifocus_dots_img.Visible = False
10750     .cmdShorten_raised_focus_img.Visible = False
10760     .cmdShorten_sunken_focus_dots_img.Visible = False
10770     .cmdShorten_raised_img_dis.Visible = False
10780     blnShorten_MouseDown = False
10790   End With

EXITP:
10800   Exit Sub

ERRH:
10810   Select Case ERR.Number
        Case Else
10820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10830   End Select
10840   Resume EXITP

End Sub

Private Sub cmdShorten_LostFocus()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_LostFocus"

10910   With Me
10920     .cmdShorten_raised_img.Visible = True
10930     .cmdShorten_raised_semifocus_dots_img.Visible = False
10940     .cmdShorten_raised_focus_img.Visible = False
10950     .cmdShorten_raised_focus_dots_img.Visible = False
10960     .cmdShorten_sunken_focus_dots_img.Visible = False
10970     .cmdShorten_raised_img_dis.Visible = False
10980     blnShorten_Focus = False
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

Private Sub cmdWiden_GotFocus()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_GotFocus"

11110   With Me
11120     blnWiden_Focus = True
11130     .cmdWiden_raised_semifocus_dots_img.Visible = True
11140     .cmdWiden_raised_img.Visible = False
11150     .cmdWiden_raised_focus_img.Visible = False
11160     .cmdWiden_raised_focus_dots_img.Visible = False
11170     .cmdWiden_sunken_focus_dots_img.Visible = False
11180     .cmdWiden_raised_img_dis.Visible = False
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

Private Sub cmdWiden_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseDown"

11310   With Me
11320     blnWiden_MouseDown = True
11330     .cmdWiden_sunken_focus_dots_img.Visible = True
11340     .cmdWiden_raised_img.Visible = False
11350     .cmdWiden_raised_semifocus_dots_img.Visible = False
11360     .cmdWiden_raised_focus_img.Visible = False
11370     .cmdWiden_raised_focus_dots_img.Visible = False
11380     .cmdWiden_raised_img_dis.Visible = False
11390   End With

EXITP:
11400   Exit Sub

ERRH:
11410   Select Case ERR.Number
        Case Else
11420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11430   End Select
11440   Resume EXITP

End Sub

Public Sub cmdWiden_Click()
' ** lngTpp * 4 = 60  : 1 grid point on form (my screen is 15 Twips Per Pixel).
' ** 60 * 4 = 240     : 4 grid points, a nice noticeable amount.
' ** 1440 Twips per Inch.
' ** Arbitrarily cap it at 10 in. wide.

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_Click"

        Dim lngAddedWidth As Long
        Dim blnSortHere As Boolean, blnCurrID As Boolean, blnCost As Boolean

11510   With Me
11520     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
11530       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
11540     End If
11550     blnSortHere = False
11560     blnCurrID = .Parent.chkIncludeCurrency
11570     blnCost = .Parent.chkShowCost
11580     lngAddedWidth = (lngTpp * WIDEN_MULT)
11590     If (.Width >= WIDEN_MAX) Or ((.Width + lngAddedWidth) > WIDEN_MAX) Then
11600       Beep
11610     Else
11620       .Width = (.Width + lngAddedWidth)
11630       .IsArch_bg1.Width = .Width
11640       .IsArch_bg2.Width = .Width
11650       If .Sort_line.Left = .revcode_ID_lbl.Left Then blnSortHere = True
11660       .revcode_DESC_display.Left = (.revcode_DESC_display.Left + lngAddedWidth)
11670       .revcode_ID.Left = .revcode_DESC_display.Left
11680       .revcode_ID_lbl.Left = .revcode_ID.Left
11690       .revcode_ID_lbl_line.Left = .revcode_ID_lbl.Left
11700       If blnSortHere = True Then
11710         .Sort_line.Left = .revcode_ID_lbl.Left
11720         .Sort_lbl.Left = (((.revcode_ID_lbl.Left + .revcode_ID_lbl.Width) - lngSortLbl_Width) + (4& * lngTpp))
11730       End If
11740       blnSortHere = False
11750       .IsIncome.Left = (.IsIncome.Left + lngAddedWidth)
11760       .IsExpense.Left = .IsIncome.Left
11770       If blnCurrID = True Then
11780         If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
11790         .curr_id.Left = (.curr_id.Left + lngAddedWidth)
11800         .curr_id_box.Left = .curr_id.Left
11810         .curr_id_forex_bg.Left = (.curr_id_box.Left - lngTpp)
11820         .curr_id_lbl.Left = (.curr_id.Left - lngCurrIDLbl_Offset)
11830         .curr_id_lbl_line.Left = .curr_id.Left
11840         If blnSortHere = True Then
11850           .Sort_line.Left = .curr_id_lbl_line.Left
11860           .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
11870         End If
11880         blnSortHere = False
11890       End If
11900       If blnCost = True Then
11910         If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
11920         .Cost.Left = (.Cost.Left + lngAddedWidth)
11930         .cost_box.Left = .Cost.Left
11940         .cost_forex_bg.Left = (.cost_box.Left - lngTpp)
11950         .cost_str.Left = .Cost.Left
11960         .cost_lbl.Left = .Cost.Left
11970         .cost_lbl_line.Left = .cost_lbl.Left
11980         If blnSortHere = True Then
11990           .Sort_line.Left = .cost_lbl.Left
12000           .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
12010         End If
12020         blnSortHere = False
12030       End If
12040       If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
12050       .PCash.Left = (.PCash.Left + lngAddedWidth)
12060       .pcash_box.Left = .PCash.Left
12070       .pcash_forex_bg.Left = (.pcash_box.Left - lngTpp)
12080       .pcash_str.Left = .PCash.Left
12090       .pcash_lbl.Left = .PCash.Left
12100       .pcash_lbl_line.Left = .pcash_lbl.Left
12110       If blnSortHere = True Then
12120         .Sort_line.Left = .pcash_lbl.Left
12130         .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
12140       End If
12150       blnSortHere = False
12160       If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
12170       .ICash.Left = (.ICash.Left + lngAddedWidth)
12180       .icash_box.Left = .ICash.Left
12190       .icash_forex_bg.Left = (.icash_box.Left - lngTpp)
12200       .icash_str.Left = .ICash.Left
12210       .icash_lbl.Left = .ICash.Left
12220       .icash_lbl_line.Left = .icash_lbl.Left
12230       If blnSortHere = True Then
12240         .Sort_line.Left = .icash_lbl.Left
12250         .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
12260       End If
12270       blnSortHere = False
12280       If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
12290       .totdesc.Width = (.totdesc.Width + lngAddedWidth)
12300       .totdesc_lbl.Width = .totdesc.Width
12310       .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
12320       If blnSortHere = True Then
12330         .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
12340         .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
12350       End If
12360       blnSortHere = False
12370     End If
12380     .FocusHolder.SetFocus
12390     If .cmdShorten.Enabled = False Then
            ' ** Since cmdWiden was clicked, cmdShorten is certainly available.
12400       .cmdShorten.Enabled = True
12410       .cmdShorten_raised_img.Visible = True
12420       .cmdShorten_raised_semifocus_dots_img.Visible = False
12430       .cmdShorten_raised_focus_img.Visible = False
12440       .cmdShorten_raised_focus_dots_img.Visible = False
12450       .cmdShorten_sunken_focus_dots_img.Visible = False
12460       .cmdShorten_raised_img_dis.Visible = False
12470     End If
          ' ** If there's nowhere to go, disable it.
12480     If (.Width >= WIDEN_MAX) Or ((.Width + lngAddedWidth) > WIDEN_MAX) Then
12490       .cmdWiden.Enabled = False
12500       .cmdWiden_raised_img_dis.Visible = True
12510       .cmdWiden_raised_img.Visible = False
12520       .cmdWiden_raised_semifocus_dots_img.Visible = False
12530       .cmdWiden_raised_focus_img.Visible = False
12540       .cmdWiden_raised_focus_dots_img.Visible = False
12550       .cmdWiden_sunken_focus_dots_img.Visible = False
12560     End If
12570   End With

EXITP:
12580   Exit Sub

ERRH:
12590   Select Case ERR.Number
        Case Else
12600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12610   End Select
12620   Resume EXITP

End Sub

Private Sub cmdWiden_KeyDown(KeyCode As Integer, Shift As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12720   intShiftDown = (Shift And acShiftMask) > 0
12730   intAltDown = (Shift And acAltMask) > 0
12740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12760     Select Case intRetVal
          Case vbKeyTab
12770       With Me
12780         intRetVal = 0
12790 On Error Resume Next
12800         .revcode_DESC_display.SetFocus
12810         If ERR.Number <> 0 Then
12820 On Error GoTo ERRH
12830           .revcode_ID.SetFocus
12840         Else
12850 On Error GoTo ERRH
12860         End If
12870       End With
12880     End Select
12890   End If

        ' ** Shift keys.
12900   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12910     Select Case intRetVal
          Case vbKeyTab
12920       With Me
12930         intRetVal = 0
12940         .cmdShorten.SetFocus
12950       End With
12960     End Select
12970   End If

EXITP:
12980   KeyCode = intRetVal
12990   Exit Sub

ERRH:
13000   intRetVal = 0
13010   Select Case ERR.Number
        Case 2105
          ' ** Do nothing.
13020   Case Else
13030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13040   End Select
13050   Resume EXITP

End Sub

Private Sub cmdWiden_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseMove"

13110   With Me
13120     If blnWiden_MouseDown = False Then
13130       Select Case blnWiden_Focus
            Case True
13140         .cmdWiden_raised_focus_dots_img.Visible = True
13150         .cmdWiden_raised_focus_img.Visible = False
13160       Case False
13170         .cmdWiden_raised_focus_img.Visible = True
13180         .cmdWiden_raised_focus_dots_img.Visible = False
13190       End Select
13200       .cmdWiden_raised_img.Visible = False
13210       .cmdWiden_raised_semifocus_dots_img.Visible = False
13220       .cmdWiden_sunken_focus_dots_img.Visible = False
13230       .cmdWiden_raised_img_dis.Visible = False
13240     End If
13250     If .cmdShorten_raised_focus_dots_img.Visible = True Or .cmdShorten_raised_focus_img.Visible = True Then
13260       Select Case blnShorten_Focus
            Case True
13270         .cmdShorten_raised_semifocus_dots_img.Visible = True
13280         .cmdShorten_raised_img.Visible = False
13290       Case False
13300         .cmdShorten_raised_img.Visible = True
13310         .cmdShorten_raised_semifocus_dots_img.Visible = False
13320       End Select
13330       .cmdShorten_raised_focus_img.Visible = False
13340       .cmdShorten_raised_focus_dots_img.Visible = False
13350       .cmdShorten_sunken_focus_dots_img.Visible = False
13360       .cmdShorten_raised_img_dis.Visible = False
13370     End If
13380   End With

EXITP:
13390   Exit Sub

ERRH:
13400   Select Case ERR.Number
        Case Else
13410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13420   End Select
13430   Resume EXITP

End Sub

Private Sub cmdWiden_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseUp"

13510   With Me
13520     .cmdWiden_raised_focus_dots_img.Visible = True
13530     .cmdWiden_raised_img.Visible = False
13540     .cmdWiden_raised_semifocus_dots_img.Visible = False
13550     .cmdWiden_raised_focus_img.Visible = False
13560     .cmdWiden_sunken_focus_dots_img.Visible = False
13570     .cmdWiden_raised_img_dis.Visible = False
13580     blnWiden_MouseDown = False
13590   End With

EXITP:
13600   Exit Sub

ERRH:
13610   Select Case ERR.Number
        Case Else
13620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13630   End Select
13640   Resume EXITP

End Sub

Private Sub cmdWiden_LostFocus()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_LostFocus"

13710   With Me
13720     .cmdWiden_raised_img.Visible = True
13730     .cmdWiden_raised_semifocus_dots_img.Visible = False
13740     .cmdWiden_raised_focus_img.Visible = False
13750     .cmdWiden_raised_focus_dots_img.Visible = False
13760     .cmdWiden_sunken_focus_dots_img.Visible = False
13770     .cmdWiden_raised_img_dis.Visible = False
13780     blnWiden_Focus = False
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

Public Sub ShowCost_Sub(blnShow As Boolean)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowCost_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean, blnCurrID As Boolean
        Dim lngNewFormWidth As Long
        Dim lngDiff1 As Long

13910   With Me
13920     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
13930       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
13940     End If
13950     blnCurrID = .Parent.chkIncludeCurrency
13960     lngDiff1 = (.totdesc.Width - lngTotDesc_Width)
13970     blnSortHere = False: blnResort = False
13980     Select Case blnShow
          Case True
13990       Select Case blnCurrID
            Case True
14000         lngNewFormWidth = lngForm_Width
14010       Case False
14020         lngNewFormWidth = (lngForm_Width - lngCurrID_Width)
14030       End Select
14040       .Width = lngNewFormWidth
14050       If .Sort_line.Left = .revcode_ID_lbl.Left Then blnSortHere = True
14060       Select Case blnCurrID
            Case True
14070         .revcode_ID.Left = (lngRevCodeID_Left + lngDiff1)
14080       Case False
14090         .revcode_ID.Left = ((lngRevCodeID_Left + lngDiff1) - lngCurrID_Width)
14100       End Select
14110       .revcode_DESC_display.Left = .revcode_ID.Left
14120       .revcode_ID_lbl.Left = .revcode_ID.Left
14130       .revcode_ID_lbl_line.Left = .revcode_ID_lbl.Left
14140       If blnSortHere = True Then
14150         .Sort_line.Left = .revcode_ID_lbl.Left
14160         .Sort_lbl.Left = (((.revcode_ID_lbl.Left + .revcode_ID_lbl.Width) - lngSortLbl_Width) + (4& * lngTpp))
14170       End If
14180       blnSortHere = False
14190       .IsIncome.Left = (.revcode_ID.Left - lngIsIncExp_Offset)
14200       .IsExpense.Left = .IsIncome.Left
14210       If blnCurrID = True Then
14220         If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
14230         .curr_id.Left = (lngCurrID_Left + lngDiff1)
14240         .curr_id_lbl.Left = (.curr_id.Left - lngCurrIDLbl_Offset)
14250         .curr_id_lbl_line.Left = .curr_id.Left
14260         .curr_id_box.Left = .curr_id.Left
14270         .curr_id_forex_bg.Left = (.curr_id.Left - lngTpp)
14280         If blnSortHere = True Then
14290           .Sort_line.Left = .curr_id_lbl_line.Left
14300           .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
14310         End If
14320         blnSortHere = False
14330       End If
14340       .Cost.Left = (lngCost_Left + lngDiff1)
14350       .cost_str.Left = .Cost.Left
14360       .cost_lbl.Left = .Cost.Left
14370       .cost_lbl_line.Left = .cost_lbl.Left
14380       .cost_box.Left = .cost_str.Left
14390       .cost_forex_bg.Left = (.cost_str.Left - lngTpp)
14400       Select Case blnCurrID
            Case True
14410         .cost_str.Visible = True
14420         .Cost.Visible = False
14430       Case False
14440         Select Case gblnHasForExThis
              Case True
14450           .cost_str.Visible = True
14460           .Cost.Visible = False
14470         Case False
14480           .Cost.Visible = True
14490           .cost_str.Visible = False
14500         End Select
14510       End Select
14520       .cost_lbl.Visible = True
14530       .cost_lbl_line.Visible = True
14540       .cost_box.Visible = True
14550       .cost_forex_bg.Visible = True
14560       .IsArch_bg1.Width = lngNewFormWidth
14570       .IsArch_bg2.Width = lngNewFormWidth
14580     Case False
14590       Select Case blnCurrID
            Case True
14600         lngNewFormWidth = ((lngForm_Width + lngDiff1) - lngCost_Width)
14610       Case False
14620         lngNewFormWidth = ((lngForm_Width + lngDiff1) - (lngCurrID_Width + lngCost_Width))
14630       End Select
14640       If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
14650       .Cost.Visible = False
14660       .cost_str.Visible = False
14670       .cost_lbl.Visible = False
14680       .cost_lbl_line.Visible = False
14690       .cost_box.Visible = False
14700       .cost_forex_bg.Visible = False
14710       .Cost.Left = 0&
14720       .cost_str.Left = 0&
14730       .cost_lbl.Left = 0&
14740       .cost_lbl_line.Left = 0&
14750       .cost_box.Left = 0&
14760       .cost_forex_bg.Left = 0&
14770       If blnSortHere = True Then
14780         .Sort_lbl.Visible = False
14790         .Sort_line.Visible = False
14800         blnResort = True
14810       End If
14820       blnSortHere = False
14830       If blnCurrID = True Then
14840         If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
14850         .curr_id.Left = ((lngCurrID_Left + lngDiff1) - lngCost_Width)
14860         .curr_id_lbl.Left = (.curr_id.Left - lngCurrIDLbl_Offset)
14870         .curr_id_lbl_line.Left = .curr_id.Left
14880         .curr_id_box.Left = .curr_id.Left
14890         .curr_id_forex_bg.Left = (.curr_id.Left - lngTpp)
14900         If blnSortHere = True Then
14910           .Sort_line.Left = .curr_id_lbl_line.Left
14920           .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
14930         End If
14940         blnSortHere = False
14950       End If
14960       If .Sort_line.Left = .revcode_ID_lbl.Left Then blnSortHere = True
14970       Select Case blnCurrID
            Case True
14980         .revcode_ID.Left = ((lngRevCodeID_Left + lngDiff1) - lngCost_Width)
14990       Case False
15000         .revcode_ID.Left = ((lngRevCodeID_Left + lngDiff1) - (lngCost_Width + lngCurrID_Width))
15010       End Select
15020       .revcode_DESC_display.Left = .revcode_ID.Left
15030       .revcode_ID_lbl.Left = .revcode_ID.Left
15040       .revcode_ID_lbl_line.Left = .revcode_ID_lbl.Left
15050       If blnSortHere = True Then
15060         .Sort_line.Left = .revcode_ID_lbl.Left
15070         .Sort_lbl.Left = (((.revcode_ID_lbl.Left + .revcode_ID_lbl.Width) - lngSortLbl_Width) + (4& * lngTpp))
15080       End If
15090       blnSortHere = False
15100       .IsIncome.Left = (.revcode_ID.Left - lngIsIncExp_Offset)
15110       .IsExpense.Left = .IsIncome.Left
15120       .IsArch_bg1.Width = lngNewFormWidth
15130       .IsArch_bg2.Width = lngNewFormWidth
15140       .Width = lngNewFormWidth
15150     End Select
15160     If blnResort = True Then
15170       SortNow "Form_Load"  ' ** Procedure: Below.
15180     End If
15190   End With

EXITP:
15200   Exit Sub

ERRH:
15210   Select Case ERR.Number
        Case Else
15220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15230   End Select
15240   Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnShow As Boolean)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean, blnShowCost As Boolean
        Dim lngNewFormWidth As Long
        Dim lngDiff1 As Long

15310   With Me
15320     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
15330       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
15340     End If
15350     blnShowCost = .Parent.chkShowCost
15360     blnSortHere = False: blnResort = False
15370     lngDiff1 = (.totdesc.Width - lngTotDesc_Width)
15380     Select Case blnShow
          Case True
15390       Select Case blnShowCost
            Case True
15400         lngNewFormWidth = (lngForm_Width + lngDiff1)
15410       Case False
15420         lngNewFormWidth = ((lngForm_Width + lngDiff1) - lngCost_Width)
15430       End Select
15440       .Width = lngNewFormWidth
15450       If .Sort_line.Left = .revcode_ID_lbl.Left Then blnSortHere = True
15460       Select Case blnShowCost
            Case True
15470         .revcode_ID.Left = (lngRevCodeID_Left + lngDiff1)
15480       Case False
15490         .revcode_ID.Left = ((lngRevCodeID_Left + lngDiff1) - lngCost_Width)
15500       End Select
15510       .revcode_DESC_display.Left = .revcode_ID.Left
15520       .revcode_ID_lbl.Left = .revcode_ID.Left
15530       .revcode_ID_lbl_line.Left = .revcode_ID_lbl.Left
15540       If blnSortHere = True Then
15550         .Sort_line.Left = .revcode_ID_lbl.Left
15560         .Sort_lbl.Left = (((.revcode_ID_lbl.Left + .revcode_ID_lbl.Width) - lngSortLbl_Width) + (4& * lngTpp))
15570       End If
15580       blnSortHere = False
15590       .IsExpense.Left = (.revcode_ID.Left - lngIsIncExp_Offset)
15600       .IsIncome.Left = .IsExpense.Left
15610       Select Case blnShowCost
            Case True
15620         .curr_id.Left = (lngCurrID_Left + lngDiff1)
15630       Case False
15640         .curr_id.Left = ((lngCurrID_Left + lngDiff1) - lngCost_Width)
15650       End Select
15660       .curr_id_lbl.Left = (.curr_id.Left - lngCurrIDLbl_Offset)
15670       .curr_id_lbl_line.Left = .curr_id.Left
15680       .curr_id_box.Left = .curr_id.Left
15690       .curr_id_forex_bg.Left = (.curr_id.Left - lngTpp)
15700       .curr_id.Visible = True
15710       .curr_id_lbl.Visible = True
15720       .curr_id_lbl_line.Visible = True
15730       .curr_id_box.Visible = True
15740       .curr_id_forex_bg.Visible = True
15750       .IsArch_bg1.Width = lngNewFormWidth
15760       .IsArch_bg2.Width = lngNewFormWidth
15770     Case False
15780       Select Case blnShowCost  ' ** lngDiff1 will usually be Zero; difference when cmdWiden used.
            Case True
15790         lngNewFormWidth = ((lngForm_Width + lngDiff1) - lngCurrID_Width)
15800       Case False
15810         lngNewFormWidth = ((lngForm_Width + lngDiff1) - (lngCost_Width + lngCurrID_Width))
15820       End Select
15830       If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
15840       .curr_id.Visible = False
15850       .curr_id_lbl.Visible = False
15860       .curr_id_lbl_line.Visible = False
15870       .curr_id_box.Visible = False
15880       .curr_id_forex_bg.Visible = False
15890       .curr_id.Left = 0&
15900       .curr_id_lbl.Left = 0&
15910       .curr_id_lbl_line.Left = 0&
15920       .curr_id_box.Left = 0&
15930       .curr_id_forex_bg.Left = 0&
15940       If blnSortHere = True Then
15950         .Sort_line.Visible = False
15960         .Sort_line.Left = 0&
15970         .Sort_lbl.Visible = False
15980         .Sort_lbl.Left = 0&
15990         blnResort = True
16000       End If
16010       blnSortHere = False
16020       If .Sort_line.Left = .revcode_ID_lbl.Left Then blnSortHere = True
16030       Select Case blnShowCost
            Case True
16040         .revcode_ID.Left = ((lngRevCodeID_Left + lngDiff1) - lngCurrID_Width)
16050       Case False
16060         .revcode_ID.Left = ((lngRevCodeID_Left + lngDiff1) - (lngCost_Width + lngCurrID_Width))
16070       End Select
16080       .revcode_DESC_display.Left = .revcode_ID.Left
16090       .revcode_ID_lbl.Left = .revcode_ID.Left
16100       .revcode_ID_lbl_line.Left = .revcode_ID_lbl.Left
16110       If blnSortHere = True Then
16120         .Sort_line.Left = .revcode_ID_lbl.Left
16130         .Sort_lbl.Left = (((.revcode_ID_lbl.Left + .revcode_ID_lbl.Width) - lngSortLbl_Width) + (4& * lngTpp))
16140       End If
16150       blnSortHere = False
16160       .IsExpense.Left = (.revcode_ID.Left - lngIsIncExp_Offset)
16170       .IsIncome.Left = .IsExpense.Left
16180       .IsArch_bg1.Width = lngNewFormWidth
16190       .IsArch_bg2.Width = lngNewFormWidth
16200       .Width = lngNewFormWidth
16210       If blnResort = True Then
16220         SortNow "Form_Load"  ' ** Procedure: Below.
16230       End If
16240     End Select
16250   End With

EXITP:
16260   Exit Sub

ERRH:
16270   Select Case ERR.Number
        Case Else
16280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16290   End Select
16300   Resume EXITP

End Sub

Public Function RecCnt() As Long

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

16410   lngRetVal = 0&

16420   With Me
16430     Set rst = .RecordsetClone
16440     With rst
16450       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
16460       Else
16470         .MoveLast
16480         lngRetVal = .RecordCount
16490       End If
16500       .Close
16510     End With
16520   End With

EXITP:
16530   Set rst = Nothing
16540   RecCnt = lngRetVal
16550   Exit Function

ERRH:
16560   lngRetVal = 0&
16570   Select Case ERR.Number
        Case Else
16580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16590   End Select
16600   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

16710   Select Case intWhere
        Case acCmdRecordsGoToFirst
16720     DoCmd.RunCommand acCmdRecordsGoToFirst
16730   Case acCmdRecordsGoToLast
16740     DoCmd.RunCommand acCmdRecordsGoToLast
16750   Case acCmdRecordsGoToPrevious
16760     DoCmd.RunCommand acCmdRecordsGoToPrevious
16770   Case acCmdRecordsGoToNext
16780     DoCmd.RunCommand acCmdRecordsGoToNext
16790   Case Else
16800     If IsMissing(varID) = False Then
16810       With Me
16820         Set rst = .RecordsetClone
16830         With rst
16840           .FindFirst "[journalno] = " & CStr(varID)
16850           If .NoMatch = False Then
16860             Me.Bookmark = .Bookmark
16870           End If
16880           .Close
16890         End With
16900       End With
16910     End If
16920   End Select

EXITP:
16930   Set rst = Nothing
16940   Exit Sub

ERRH:
16950   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
16960   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
16970   Case Else
16980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16990   End Select
17000   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [transdate], [journaltype], [journalno]"

17110   With Me
17120     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
17130       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
17140     End If
17150     .Controls(strSortLbl).Visible = False
17160     .Controls(strSortLine).Visible = False
17170     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
17180     If strProc = "Form_Load" Then
17190       strCalled = "transdate"
17200       strSortNow = strSortOrig
17210       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
17220       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
17230       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
17240       .Controls(strSortLbl).Left = lngSortLbl_Left
17250       .Controls(strSortLine).Left = lngSortLine_Left
17260       .Controls(strSortLine).Width = lngSortLine_Width
17270       .Controls(strSortLbl).Caption = strArwDn   ' ** Original is descending.
17280       .Controls(strSortLbl).ForeColor = CLR_DKRED
17290     Else
17300       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: revcode_ID_lbl_DblClick
17310       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
17320       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
17330       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
17340       .Controls(strSortLbl).Left = lngSortLbl_Left
17350       .Controls(strSortLine).Left = lngSortLine_Left
17360       .Controls(strSortLine).Width = lngSortLine_Width
17370       If strCalled = "transdate" Then
17380         strSortAsc = "[transdate], [journaltype], [journalno]"
17390         If strSortNow = strSortAsc Then
17400           strSortNow = strSortOrig   ' ** Original is descending: "[transdate] DESC, [journaltype] ASC, [journalno] DESC"
17410           .Controls(strSortLbl).Caption = strArwDn
17420           .Controls(strSortLbl).ForeColor = CLR_DKRED
17430         Else
17440           strSortNow = strSortAsc
17450           .Controls(strSortLbl).Caption = strArwUp
17460           .Controls(strSortLbl).ForeColor = CLR_DKBLU
17470         End If
17480       Else
17490         .Controls(strSortLbl).Caption = strArwUp
17500         .Controls(strSortLbl).ForeColor = CLR_DKBLU
17510         Select Case strCalled
              Case "journaltype"
17520           strSortAsc = "[journaltype], [transdate], [journalno]"
17530         Case "shareface"
17540           strSortAsc = "[sharefacex]" & strStdAsc
17550         Case "totdesc"
17560           strSortAsc = "[totdesc]" & strStdAsc
17570         Case "ICash"
17580           strSortAsc = "[ICash]" & strStdAsc
17590         Case "PCash"
17600           strSortAsc = "[PCash]" & strStdAsc
17610         Case "Cost"
17620           strSortAsc = "[Cost]" & strStdAsc
17630         Case "curr_id"
17640           strSortAsc = "[curr_code]" & strStdAsc
17650           lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
17660           lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
17670           lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
17680           .Controls(strSortLbl).Left = lngSortLbl_Left
17690           .Controls(strSortLine).Left = lngSortLine_Left
17700           .Controls(strSortLine).Width = lngSortLine_Width
17710         Case "revcode_ID"
17720           strSortAsc = "[revcode_DESC]" & strStdAsc
17730           lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
17740           .Controls(strSortLbl).Left = lngSortLbl_Left
17750         End Select
17760         If strSortNow = strSortAsc Then
17770           intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
17780           Select Case intCnt
                Case 1
17790             strTmp01 = strSortAsc & " DESC"
17800           Case 2
17810             intPos01 = InStr(strSortAsc, ",")
17820             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
17830             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
17840           Case 3
17850             intPos01 = InStr(strSortAsc, ",")
17860             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
17870             strTmp02 = Mid(strSortAsc, intPos01)
17880             intPos01 = InStr(2, strTmp02, ",")
17890             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
17900             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
17910           Case 4
17920             intPos01 = InStr(strSortAsc, ",")
17930             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
17940             strTmp02 = Mid(strSortAsc, intPos01)
17950             intPos01 = InStr(2, strTmp02, ",")
17960             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
17970             strTmp02 = Mid(strTmp02, intPos01)
17980             intPos01 = InStr(2, strTmp02, ",")
17990             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
18000             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
18010           Case 5
18020             intPos01 = InStr(strSortAsc, ",")
18030             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
18040             strTmp02 = Mid(strSortAsc, intPos01)
18050             intPos01 = InStr(2, strTmp02, ",")
18060             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
18070             strTmp02 = Mid(strTmp02, intPos01)
18080             intPos01 = InStr(2, strTmp02, ",")
18090             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
18100             strTmp02 = Mid(strTmp02, intPos01)
18110             intPos01 = InStr(2, strTmp02, ",")
18120             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
18130             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
18140           End Select
18150           strSortNow = strTmp01
18160           .Controls(strSortLbl).Caption = strArwDn
18170           .Controls(strSortLbl).ForeColor = CLR_DKRED
18180         Else
18190           strSortNow = strSortAsc
18200         End If
18210       End If
18220     End If
18230     .Controls(strSortLbl).Visible = True
18240     .Controls(strSortLine).Visible = True
18250     .OrderBy = strSortNow
18260     .OrderByOn = True
18270   End With

EXITP:
18280   Exit Sub

ERRH:
18290   Select Case ERR.Number
        Case Else
18300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18310   End Select
18320   Resume EXITP

End Sub
