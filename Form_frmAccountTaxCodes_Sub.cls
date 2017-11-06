VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountTaxCodes_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountTaxCodes_Sub"

'VGC 04/15/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Block Assign:     B {cmdBlockAssign on frmAccountTaxCodes}
' **   Close:            C {cmdClose on frmAccountTaxCodes}
' **   Include Archive:  H {chkIncludeArchive on frmAccountTaxCodes}
' **   Include Currency: L {chkIncludeCurrency on frmAccountTaxCodes}
' **   Show Cost:        O {chkShowCost on frmAccountTaxCodes}
' **   Shorten:          S {cmdShorten}
' **   Update:           U {cmdUpdate on frmAccountTaxCodes}
' **   Widen:            W {cmdWiden}
' **   Exit:             X {cmdClose on frmAccountTaxCodes}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ** .taxcode.ListWidth = 2280
' ** .taxcode.ColumnWidths = "0;1785;0;390;0;0" + 255 (17 Pixels) for ScrollBar.

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

Private blnShorten_Focus As Boolean, blnShorten_MouseDown As Boolean
Private blnWiden_Focus As Boolean, blnWiden_MouseDown As Boolean

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngForm_Width As Long, lngTotDesc_Width As Long, lngTaxCode_Left As Long, lngIsIncDed_Offset As Long
Private lngCost_Left As Long, lngCost_Width As Long, lngICash_Left As Long, lngPCash_Left As Long
Private lngCurrID_Left As Long, lngCurrID_Width As Long, lngCurrIDLbl_Offset As Long

Private lngIncs As Long, arr_varInc As Variant
Private lngDeds As Long, arr_varDed As Variant
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
150       If strCallingForm = "frmAccountTaxCodes" Then

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
320         lngTaxCode_Left = .taxcode.Left
330         lngIsIncDed_Offset = (.taxcode.Left - .IsIncome.Left)

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
500     Select Case ERR.Number
        Case Else
510       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
520     End Select
530     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

610     With Me

620       .transdate.ForeColor = CLR_DISABLED_FG
630       .transdate.BackColor = CLR_DISABLED_BG
640       .journaltype.ForeColor = CLR_DISABLED_FG
650       .journaltype.BackColor = CLR_DISABLED_BG
660       .shareface.ForeColor = CLR_DISABLED_FG
670       .shareface.BackColor = CLR_DISABLED_BG
680       .totdesc.ForeColor = CLR_DISABLED_FG
690       .totdesc.BackColor = CLR_DISABLED_BG
700       .ICash.ForeColor = CLR_DISABLED_FG
710       .ICash.BackColor = CLR_DISABLED_BG
720       .icash_str.ForeColor = CLR_DISABLED_FG
730       .icash_str.BackColor = CLR_DISABLED_BG
740       .PCash.ForeColor = CLR_DISABLED_FG
750       .PCash.BackColor = CLR_DISABLED_BG
760       .pcash_str.ForeColor = CLR_DISABLED_FG
770       .pcash_str.BackColor = CLR_DISABLED_BG
780       .Cost.ForeColor = CLR_DISABLED_FG
790       .Cost.BackColor = CLR_DISABLED_BG
800       .cost_str.ForeColor = CLR_DISABLED_FG
810       .cost_str.BackColor = CLR_DISABLED_BG
820       Set dbs = CurrentDb
830       With dbs

            ' ** INCOME taxcodes.                 Number, Description, w/IncExp, INCOME
840         Set qdf = .QueryDefs("qryTaxCode_07")
850         Set rst = qdf.OpenRecordset
860         With rst
870           .MoveLast
880           lngIncs = .RecordCount
890           .MoveFirst
900           arr_varInc = .GetRows(lngIncs)
              ' *************************************************
              ' ** Array: arr_varInc()
              ' **
              ' **   Field  Element  Name            Constant
              ' **   =====  =======  ==============  ==========
              ' **     1       0     taxcode
              ' **     2       1     taxcode_description
              ' **     3       2     revcode_TYPE
              ' **     4       3     IE
              ' **
              ' *************************************************
910           .Close
920         End With
930         Set rst = Nothing
940         Set qdf = Nothing

            ' ** EXPENSE taxcodes.                Number, Description, w/IncExp, EXPENSE
950         Set qdf = .QueryDefs("qryTaxCode_08")
960         Set rst = qdf.OpenRecordset
970         With rst
980           .MoveLast
990           lngDeds = .RecordCount
1000          .MoveFirst
1010          arr_varDed = .GetRows(lngDeds)
              ' *************************************************
              ' ** Array: arr_varDed()
              ' **
              ' **   Field  Element  Name            Constant
              ' **   =====  =======  ==============  ==========
              ' **     1       0     taxcode
              ' **     2       1     taxcode_description
              ' **     3       2     revcode_TYPE
              ' **     4       3     IE
              ' **
              ' *************************************************
1020          .Close
1030        End With
1040        Set rst = Nothing
1050        Set qdf = Nothing

1060        .Close
1070      End With
1080      Set dbs = Nothing

1090      SortNow THIS_PROC  ' ** Procedure: Below

1100    End With

EXITP:
1110    Set rst = Nothing
1120    Set qdf = Nothing
1130    Set dbs = Nothing
1140    Exit Sub

ERRH:
1150    Select Case ERR.Number
        Case Else
1160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1170    End Select
1180    Resume EXITP

End Sub

Private Sub Form_Current()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

1210    With Me
1220      Select Case .journaltype
          Case "Dividend", "Interest", "Purchase", "Deposit", "Sold", "Received"
            ' ** INCOME.
1230        If .taxcode.RowSource <> "qryTaxCode_07" Then
1240          .taxcode.RowSource = "qryTaxCode_07"
1250          .taxcode.Requery
1260        End If
1270      Case "Liability", "Paid"
            ' ** DEDUCTION.
1280        If .taxcode.RowSource <> "qryTaxCode_08" Then
1290          .taxcode.RowSource = "qryTaxCode_08"
1300          .taxcode.Requery
1310        End If
1320      Case "Withdrawn", "Cost Adj.", "Misc."
            ' ** ALL.
1330        If .taxcode.RowSource <> "qryTaxCode_09" Then
1340          .taxcode.RowSource = "qryTaxCode_09"
1350          .taxcode.Requery
1360        End If
1370      End Select
1380    End With

EXITP:
1390    Exit Sub

ERRH:
1400    Select Case ERR.Number
        Case Else
1410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1420    End Select
1430    Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Shorten:          S {cmdShorten}
        ' **   Widen:            W {cmdWiden}
        ' **   Exit:             X {cmdClose on frmAccountTaxCodes}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1560      Select Case intRetVal
          Case vbKeyDown
1570        intRetVal = 0
1580        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1590      Case vbKeyUp
1600        intRetVal = 0
1610        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1620      Case vbKeyEscape
1630        With Me
1640          intRetVal = 0
1650          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountTaxCodes.
1660        End With
1670      End Select
1680    End If

        ' ** Alt keys.
1690    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1700      Select Case intRetVal
          Case vbKeyS
            ' ** Invoke the shorten-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
1710        With Me
1720          intRetVal = 0
1730          If .cmdShorten.Enabled = True Then
1740            cmdShorten_Click  ' ** Procedure: Below.
1750          Else
1760            Beep
1770          End If
1780        End With
1790      Case vbKeyW
            ' ** Invoke the widen-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
1800        With Me
1810          intRetVal = 0
1820          If .cmdWiden.Enabled = True Then
1830            cmdWiden_Click  ' ** Procedure: Below.
1840          Else
1850            Beep
1860          End If
1870        End With
1880      Case vbKeyX
1890        With Me
1900          intRetVal = 0
1910          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountTaxCodes.
1920        End With
1930      End Select
1940    End If

        ' ** Ctrl keys.
1950    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1960      Select Case intRetVal
          Case vbKeyS
1970        intRetVal = 0
1980        cmdSave_Click  ' ** Procedure: Below.
1990      Case vbKeyTab
2000        With Me
2010          intRetVal = 0
2020          DoCmd.SelectObject acForm, .Parent.Name, False
2030          If .Parent.chkIncludeArchive.Enabled = True Then
2040            .Parent.chkIncludeArchive.SetFocus
2050          ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
2060            .Parent.chkIncludeCurrency.SetFocus
2070          Else
2080            .Parent.chkShowCost.SetFocus
2090          End If
2100        End With
2110      Case vbKeyPageUp, vbKeyHome
2120        intRetVal = 0
2130        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2140      Case vbKeyPageDown, vbKeyEnd
2150        intRetVal = 0
2160        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2170      End Select
2180    End If

        ' ** Ctrl-Shift keys.
2190    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2200      Select Case intRetVal
          Case vbKeyF
2210        With Me
2220          intRetVal = 0
2230          DoCmd.SelectObject acForm, .Parent.Name, False
2240          .Parent.FocusHolder.SetFocus
2250        End With
2260      Case vbKeyTab
2270        With Me
2280          intRetVal = 0
2290          DoCmd.SelectObject acForm, .Parent.Name, False
2300          .Parent.cmdClose.SetFocus
2310        End With
2320      End Select
2330    End If

EXITP:
2340    KeyCode = intRetVal
2350    Exit Sub

ERRH:
2360    intRetVal = 0
2370    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
2380    Case Else
2390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2400    End Select
2410    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

2510    With Me
2520      If IsNull(.taxcode) = True Then
2530        Select Case .revcode_TYPE
            Case REVTYP_INC
2540          .taxcode = TAXID_INC
2550          .taxcode_DESC = "Unspecified Income"
2560          .taxcode_DESC_display = vbNullString
2570        Case REVTYP_EXP
2580          .taxcode = TAXID_DED
2590          .taxcode_DESC = "Unspecified Deduction"
2600          .taxcode_DESC_display = vbNullString
2610        End Select
2620      End If
2630    End With

EXITP:
2640    Exit Sub

ERRH:
2650    Select Case ERR.Number
        Case Else
2660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2670    End Select
2680    Resume EXITP

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

Private Sub taxcode_lbl_DblClick(Cancel As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_lbl_DblClick"

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

Private Sub taxcode_DESC_display_GotFocus()
' ** This control sits on top of the taxcode combo box.
' ** It's used to display the taxcode 'taxcode_description' [sic]
' ** when the combo box RowSource is showing the other set.
' ** CONTROL IS ENABLED, BUT LOCKED.

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_DESC_display_GotFocus"

        Dim varTmp00 As Variant

5310    With Me
5320  On Error Resume Next
5330      varTmp00 = .taxcode_DESC
5340      If ERR.Number <> 0 Then
5350  On Error GoTo ERRH
5360        MsgBox "No records meet the specified criteria.", vbInformation + vbOKOnly, "Not Found"
5370        DoCmd.RunCommand acCmdRemoveFilterSort
5380      Else
5390  On Error GoTo ERRH
5400        If Nz(.taxcode_DESC, vbNullString) = vbNullString Then
5410          .taxcode.SetFocus
5420        Else
5430          Select Case .taxcode.RowSource
              Case "qryTaxCode_07"
                ' ** INCOME.
5440            If .taxcode_type = TAXTYP_DED Then
5450              .taxcode = TAXID_INC
5460              .changed = True
5470              .taxcode_DESC = "Unspecified Income"
5480              .taxcode_DESC_display = vbNullString
5490              .taxcode.SetFocus
5500            Else
5510              .taxcode.SetFocus
5520            End If
5530          Case "qryTaxCode_08"
                ' ** DEDUCTION.
5540            If .taxcode_type = TAXTYP_INC Then
5550              .taxcode = TAXID_DED
5560              .changed = True
5570              .taxcode_DESC = "Unspecified Deduction"
5580              .taxcode_DESC_display = vbNullString
5590              .taxcode.SetFocus
5600            Else
5610              .taxcode.SetFocus
5620            End If
5630          Case "qryTaxCode_09"
                ' ** ALL.
5640            .taxcode.SetFocus
5650          End Select
5660        End If
5670      End If
5680    End With

EXITP:
5690    Exit Sub

ERRH:
5700    Select Case ERR.Number
        Case Else
5710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5720    End Select
5730    Resume EXITP

End Sub

Private Sub taxcode_DESC_display_AfterUpdate()
' ** This control sits on top of the taxcode combo box.
' ** It's used to display the taxcode 'taxcode_description'
' ** when the combo box RowSource is showing the other set.
' ** CONTROL IS ENABLED, BUT LOCKED.

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_DESC_display_AfterUpdate"

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

Private Sub taxcode_AfterUpdate()

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_AfterUpdate"

5910    With Me
5920      .changed = True
5930      If IsNull(.taxcode) = False Then
5940        If IsNull(.taxcode.Column(1)) = False Then
5950          .taxcode_DESC = .taxcode.Column(1)
5960          Select Case .taxcode.Column(0)
              Case 1, 2
5970            .taxcode_DESC_display = vbNullString
5980          Case Else
5990            .taxcode_DESC_display = .taxcode.Column(1)
6000          End Select
6010        Else
6020          .taxcode = TAXID_INC
6030          .taxcode_DESC = "Unspecified Income"
6040          .taxcode_DESC_display = vbNullString
6050        End If
6060      Else
6070        .taxcode = TAXID_INC
6080        .taxcode_DESC = "Unspecified Income"
6090        .taxcode_DESC_display = vbNullString
6100      End If
6110      If gblnLinkRevTaxCodes = True Then
6120        Select Case .taxcode_type
            Case TAXTYP_INC
6130          If .revcode_TYPE = REVTYP_EXP Then
6140            .revcode_ID = REVID_INC  ' ** Unspecified Income.
6150            .revcode_TYPE = REVTYP_INC  ' ** INCOME.
6160          End If
6170        Case TAXTYP_DED
6180          If .revcode_TYPE = REVTYP_INC Then
6190            .revcode_ID = REVID_EXP  ' ** Unspecified Expense.
6200            .revcode_TYPE = REVTYP_EXP  ' ** EXPENSE.
6210          End If
6220        End Select
6230      End If
6240      cmdSave_Click  ' ** Procedure: Above.
6250      .Parent.chkUpdated = True
6260      .Parent.cmdUpdate.Enabled = True
6270    End With

EXITP:
6280    Exit Sub

ERRH:
6290    Select Case ERR.Number
        Case Else
6300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6310    End Select
6320    Resume EXITP

End Sub

Private Sub taxcode_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6420    intShiftDown = (Shift And acShiftMask) > 0
6430    intAltDown = (Shift And acAltMask) > 0
6440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6470        With Me
6480          intRetVal = 0
6490          lngRecsCur = RecCnt  ' ** Function: Below.
6500          If .CurrentRecord < lngRecsCur Then
6510            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
6520  On Error Resume Next
6530            .taxcode_DESC.SetFocus
6540            If ERR.Number <> 0 Then
6550  On Error GoTo ERRH
6560              .taxcode.SetFocus
6570            Else
6580  On Error GoTo ERRH
6590            End If
6600          Else
6610            DoCmd.SelectObject acForm, .Parent.Name, False
6620            If .Parent.chkIncludeArchive.Enabled = True Then
6630              .Parent.chkIncludeArchive.SetFocus
6640            ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
6650              .Parent.chkIncludeCurrency.SetFocus
6660            Else
6670              .Parent.chkShowCost.SetFocus
6680            End If
6690          End If
6700        End With
6710      End Select
6720    End If

        ' ** Shift keys.
6730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6750        With Me
6760          intRetVal = 0
6770          If .CurrentRecord > 1 Then
6780            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
6790  On Error Resume Next
6800            .taxcode_DESC.SetFocus
6810            If ERR.Number <> 0 Then
6820  On Error GoTo ERRH
6830              .taxcode.SetFocus
6840            Else
6850  On Error GoTo ERRH
6860            End If
6870          Else
6880            DoCmd.SelectObject acForm, .Parent.Name, False
6890            .Parent.cmdClose.SetFocus
6900          End If
6910        End With
6920      End Select
6930    End If

EXITP:
6940    KeyCode = intRetVal
6950    Exit Sub

ERRH:
6960    intRetVal = 0
6970    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
6980    Case Else
6990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7000    End Select
7010    Resume EXITP

End Sub

Private Sub cmdShorten_GotFocus()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_GotFocus"

7110    With Me
7120      blnShorten_Focus = True
7130      .cmdShorten_raised_semifocus_dots_img.Visible = True
7140      .cmdShorten_raised_img.Visible = False
7150      .cmdShorten_raised_focus_img.Visible = False
7160      .cmdShorten_raised_focus_dots_img.Visible = False
7170      .cmdShorten_sunken_focus_dots_img.Visible = False
7180      .cmdShorten_raised_img_dis.Visible = False
7190    End With

EXITP:
7200    Exit Sub

ERRH:
7210    Select Case ERR.Number
        Case Else
7220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7230    End Select
7240    Resume EXITP

End Sub

Private Sub cmdShorten_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseDown"

7310    With Me
7320      blnShorten_MouseDown = True
7330      .cmdShorten_sunken_focus_dots_img.Visible = True
7340      .cmdShorten_raised_img.Visible = False
7350      .cmdShorten_raised_semifocus_dots_img.Visible = False
7360      .cmdShorten_raised_focus_img.Visible = False
7370      .cmdShorten_raised_focus_dots_img.Visible = False
7380      .cmdShorten_raised_img_dis.Visible = False
7390    End With

EXITP:
7400    Exit Sub

ERRH:
7410    Select Case ERR.Number
        Case Else
7420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7430    End Select
7440    Resume EXITP

End Sub

Public Sub cmdShorten_Click()
' ** lngTpp * 4 = 60  : 1 grid point on form (my screen is 15 Twips Per Pixel).
' ** 60 * 4 = 240     : 4 grid points, a nice noticeable amount.
' ** 1440 Twips per Inch.

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_Click"

        Dim lngMinusedWidth As Long
        Dim blnSortHere As Boolean, blnCurrID As Boolean, blnCost As Boolean

7510    With Me
7520      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
7530        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
7540      End If
7550      blnSortHere = False
7560      blnCurrID = .Parent.chkIncludeCurrency
7570      blnCost = .Parent.chkShowCost
7580      lngMinusedWidth = lngTpp * WIDEN_MULT
7590      If (.totdesc.Width <= lngTotDesc_Width) Then
7600        Beep
7610      ElseIf ((.totdesc.Width - lngMinusedWidth) < lngTotDesc_Width) Then
7620        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
7630        .totdesc.Width = lngTotDesc_Width
7640        .totdesc_lbl.Width = .totdesc.Width
7650        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7660        If blnSortHere = True Then
7670          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
7680          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
7690        End If
7700        blnSortHere = False
7710        If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
7720        .ICash.Left = lngICash_Left
7730        .icash_box.Left = .ICash.Left
7740        .icash_forex_bg.Left = (.icash_box.Left - lngTpp)
7750        .icash_str.Left = .ICash.Left
7760        .icash_lbl.Left = .ICash.Left
7770        .icash_lbl_line.Left = .icash_lbl.Left
7780        If blnSortHere = True Then
7790          .Sort_line.Left = .icash_lbl.Left
7800          .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
7810        End If
7820        blnSortHere = False
7830        If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
7840        .PCash.Left = lngPCash_Left
7850        .pcash_box.Left = .PCash.Left
7860        .pcash_forex_bg.Left = (.pcash_box.Left - lngTpp)
7870        .pcash_str.Left = .PCash.Left
7880        .pcash_lbl.Left = .PCash.Left
7890        .pcash_lbl_line.Left = .pcash_lbl.Left
7900        If blnSortHere = True Then
7910          .Sort_line.Left = .pcash_lbl.Left
7920          .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
7930        End If
7940        blnSortHere = False
7950        If blnCost = True Then
7960          If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
7970          .Cost.Left = lngCost_Left
7980          .cost_box.Left = .Cost.Left
7990          .cost_forex_bg.Left = (.cost_box.Left - lngTpp)
8000          .cost_str.Left = .Cost.Left
8010          .cost_lbl.Left = .Cost.Left
8020          .cost_lbl_line.Left = .cost_lbl.Left
8030          If blnSortHere = True Then
8040            .Sort_line.Left = .cost_lbl.Left
8050            .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
8060          End If
8070          blnSortHere = False
8080        End If
8090        If blnCurrID = True Then
8100          If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
8110          Select Case blnCost
              Case True
8120            .curr_id.Left = lngCurrID_Left
8130          Case False
8140            .curr_id.Left = (lngCurrID_Left - lngCost_Width)
8150          End Select
8160          .curr_id_box.Left = .curr_id.Left
8170          .curr_id_forex_bg.Left = (.curr_id_box.Left - lngTpp)
8180          .curr_id_lbl.Left = (.curr_id.Left - lngCurrIDLbl_Offset)
8190          .curr_id_lbl_line.Left = .curr_id.Left
8200          If blnSortHere = True Then
8210            .Sort_line.Left = .curr_id_lbl_line.Left
8220            .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
8230          End If
8240          blnSortHere = False
8250        End If
8260        If .Sort_line.Left = .taxcode_lbl.Left Then blnSortHere = True
8270        Select Case blnCost
            Case True
8280          Select Case blnCurrID
              Case True
8290            .taxcode.Left = lngTaxCode_Left
8300          Case False
8310            .taxcode.Left = (lngTaxCode_Left - lngCurrID_Width)
8320          End Select
8330        Case False
8340          Select Case blnCurrID
              Case True
8350            .taxcode.Left = (lngTaxCode_Left - lngCost_Width)
8360          Case False
8370            .taxcode.Left = (lngTaxCode_Left - (lngCost_Width + lngCurrID_Width))
8380          End Select
8390        End Select
8400        .taxcode_DESC_display = .taxcode.Left
8410        .taxcode_lbl.Left = .taxcode.Left
8420        .taxcode_lbl_line.Left = .taxcode_lbl.Left
8430        If blnSortHere = True Then
8440          .Sort_line.Left = .taxcode_lbl.Left
8450          .Sort_lbl.Left = ((.taxcode_lbl.Left + .taxcode_lbl.Width) - lngSortLbl_Width)
8460        End If
8470        blnSortHere = False
8480        .IsIncome.Left = (.taxcode.Left - lngIsIncDed_Offset)
8490        .IsExpense.Left = .IsIncome.Left
8500        .IsArch_bg1.Width = lngForm_Width
8510        .IsArch_bg2.Width = lngForm_Width
8520        .Width = lngForm_Width
8530      Else
8540        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
8550        .totdesc.Width = (.totdesc.Width - lngMinusedWidth)
8560        .totdesc_lbl.Width = .totdesc.Width
8570        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
8580        If blnSortHere = True Then
8590          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
8600          .Sort_lbl = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
8610        End If
8620        blnSortHere = False
8630        If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
8640        .ICash.Left = (.ICash.Left - lngMinusedWidth)
8650        .icash_box.Left = .ICash.Left
8660        .icash_forex_bg.Left = (.icash_box.Left - lngTpp)
8670        .icash_str.Left = .ICash.Left
8680        .icash_lbl.Left = .ICash.Left
8690        .icash_lbl_line.Left = .icash_lbl.Left
8700        If blnSortHere = True Then
8710          .Sort_line.Left = .icash_lbl.Left
8720          .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
8730        End If
8740        blnSortHere = False
8750        If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
8760        .PCash.Left = (.PCash.Left - lngMinusedWidth)
8770        .pcash_box.Left = .PCash.Left
8780        .pcash_forex_bg.Left = (.pcash_box.Left - lngTpp)
8790        .pcash_str.Left = .PCash.Left
8800        .pcash_lbl.Left = .PCash.Left
8810        .pcash_lbl_line.Left = .pcash_lbl.Left
8820        If blnSortHere = True Then
8830          .Sort_line.Left = .pcash_lbl.Left
8840          .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
8850        End If
8860        blnSortHere = False
8870        If blnCost = True Then
8880          If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
8890          .Cost.Left = (.Cost.Left - lngMinusedWidth)
8900          .cost_box.Left = .Cost.Left
8910          .cost_forex_bg.Left = (.cost_box.Left - lngTpp)
8920          .cost_str.Left = .Cost.Left
8930          .cost_lbl.Left = .Cost.Left
8940          .cost_lbl_line.Left = .cost_lbl.Left
8950          If blnSortHere = True Then
8960            .Sort_line.Left = .cost_lbl.Left
8970            .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
8980          End If
8990          blnSortHere = False
9000        End If
9010        If blnCurrID = True Then
9020          If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
9030          .curr_id.Left = (.curr_id.Left - lngMinusedWidth)
9040          .curr_id_box.Left = .curr_id.Left
9050          .curr_id_forex_bg.Left = (.curr_id_box.Left - lngTpp)
9060          .curr_id_lbl.Left = (.curr_id.Left - lngCurrIDLbl_Offset)
9070          .curr_id_lbl_line.Left = .curr_id.Left
9080          If blnSortHere = True Then
9090            .Sort_line.Left = .curr_id_lbl_line.Left
9100            .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
9110          End If
9120          blnSortHere = False
9130        End If
9140        If .Sort_line.Left = .taxcode_lbl.Left Then blnSortHere = True
9150        .taxcode.Left = (.taxcode.Left - lngMinusedWidth)
9160        .taxcode_DESC_display.Left = .taxcode.Left
9170        .taxcode_lbl.Left = .taxcode.Left
9180        .taxcode_lbl_line.Left = .taxcode_lbl.Left
9190        If blnSortHere = True Then
9200          .Sort_line.Left = .taxcode_lbl.Left
9210          .Sort_lbl.Left = ((.taxcode_lbl.Left + .taxcode_lbl.Width) - lngSortLbl_Width)
9220        End If
9230        blnSortHere = False
9240        .IsIncome.Left = (.taxcode.Left - lngIsIncDed_Offset)
9250        .IsExpense.Left = .IsIncome.Left
9260        .IsArch_bg1.Width = (.Width - lngMinusedWidth)
9270        .IsArch_bg2.Width = (.Width - lngMinusedWidth)
9280        .Width = (.Width - lngMinusedWidth)
9290      End If
9300      .FocusHolder.SetFocus
9310      If .cmdWiden.Enabled = False Then
            ' ** Since cmdShorten was clicked, cmdWiden is certainly available.
9320        .cmdWiden.Enabled = True
9330        .cmdWiden_raised_img.Visible = True
9340        .cmdWiden_raised_semifocus_dots_img.Visible = False
9350        .cmdWiden_raised_focus_img.Visible = False
9360        .cmdWiden_raised_focus_dots_img.Visible = False
9370        .cmdWiden_sunken_focus_dots_img.Visible = False
9380        .cmdWiden_raised_img_dis.Visible = False
9390      End If
9400      If (.totdesc.Width <= lngTotDesc_Width) Then
            ' ** If there's nowhere to go, disable it.
9410        .cmdShorten.Enabled = False
9420        .cmdShorten_raised_img_dis.Visible = True
9430        .cmdShorten_raised_img.Visible = False
9440        .cmdShorten_raised_semifocus_dots_img.Visible = False
9450        .cmdShorten_raised_focus_img.Visible = False
9460        .cmdShorten_raised_focus_dots_img.Visible = False
9470        .cmdShorten_sunken_focus_dots_img.Visible = False
9480      End If
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

Private Sub cmdShorten_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_KeyDown"

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
9690          .cmdWiden.SetFocus
9700        End With
9710      End Select
9720    End If

        ' ** Shift keys.
9730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9740      Select Case intRetVal
          Case vbKeyTab
9750        With Me
9760          intRetVal = 0
9770  On Error Resume Next
9780          .taxcode_DESC_display.SetFocus
9790          If ERR.Number <> 0 Then
9800  On Error GoTo ERRH
9810            .taxcode.SetFocus
9820          Else
9830  On Error GoTo ERRH
9840          End If
9850        End With
9860      End Select
9870    End If

EXITP:
9880    KeyCode = intRetVal
9890    Exit Sub

ERRH:
9900    intRetVal = 0
9910    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
9920    Case Else
9930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9940    End Select
9950    Resume EXITP

End Sub

Private Sub cmdShorten_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseMove"

10010   With Me
10020     If blnShorten_MouseDown = False Then
10030       Select Case blnShorten_Focus
            Case True
10040         .cmdShorten_raised_focus_dots_img.Visible = True
10050         .cmdShorten_raised_focus_img.Visible = False
10060       Case False
10070         .cmdShorten_raised_focus_img.Visible = True
10080         .cmdShorten_raised_focus_dots_img.Visible = False
10090       End Select
10100       .cmdShorten_raised_img.Visible = False
10110       .cmdShorten_raised_semifocus_dots_img.Visible = False
10120       .cmdShorten_sunken_focus_dots_img.Visible = False
10130       .cmdShorten_raised_img_dis.Visible = False
10140     End If
10150     If .cmdWiden_raised_focus_dots_img.Visible = True Or .cmdWiden_raised_focus_img.Visible = True Then
10160       Select Case blnWiden_Focus
            Case True
10170         .cmdWiden_raised_semifocus_dots_img.Visible = True
10180         .cmdWiden_raised_img.Visible = False
10190       Case False
10200         .cmdWiden_raised_img.Visible = True
10210         .cmdWiden_raised_semifocus_dots_img.Visible = False
10220       End Select
10230       .cmdWiden_raised_focus_img.Visible = False
10240       .cmdWiden_raised_focus_dots_img.Visible = False
10250       .cmdWiden_sunken_focus_dots_img.Visible = False
10260       .cmdWiden_raised_img_dis.Visible = False
10270     End If
10280   End With

EXITP:
10290   Exit Sub

ERRH:
10300   Select Case ERR.Number
        Case Else
10310     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10320   End Select
10330   Resume EXITP

End Sub

Private Sub cmdShorten_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseUp"

10410   With Me
10420     .cmdShorten_raised_focus_dots_img.Visible = True
10430     .cmdShorten_raised_img.Visible = False
10440     .cmdShorten_raised_semifocus_dots_img.Visible = False
10450     .cmdShorten_raised_focus_img.Visible = False
10460     .cmdShorten_sunken_focus_dots_img.Visible = False
10470     .cmdShorten_raised_img_dis.Visible = False
10480     blnShorten_MouseDown = False
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

Private Sub cmdShorten_LostFocus()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_LostFocus"

10610   With Me
10620     .cmdShorten_raised_img.Visible = True
10630     .cmdShorten_raised_semifocus_dots_img.Visible = False
10640     .cmdShorten_raised_focus_img.Visible = False
10650     .cmdShorten_raised_focus_dots_img.Visible = False
10660     .cmdShorten_sunken_focus_dots_img.Visible = False
10670     .cmdShorten_raised_img_dis.Visible = False
10680     blnShorten_Focus = False
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

Private Sub cmdWiden_GotFocus()

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_GotFocus"

10810   With Me
10820     blnWiden_Focus = True
10830     .cmdWiden_raised_semifocus_dots_img.Visible = True
10840     .cmdWiden_raised_img.Visible = False
10850     .cmdWiden_raised_focus_img.Visible = False
10860     .cmdWiden_raised_focus_dots_img.Visible = False
10870     .cmdWiden_sunken_focus_dots_img.Visible = False
10880     .cmdWiden_raised_img_dis.Visible = False
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

Private Sub cmdWiden_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseDown"

11010   With Me
11020     blnWiden_MouseDown = True
11030     .cmdWiden_sunken_focus_dots_img.Visible = True
11040     .cmdWiden_raised_img.Visible = False
11050     .cmdWiden_raised_semifocus_dots_img.Visible = False
11060     .cmdWiden_raised_focus_img.Visible = False
11070     .cmdWiden_raised_focus_dots_img.Visible = False
11080     .cmdWiden_raised_img_dis.Visible = False
11090   End With

EXITP:
11100   Exit Sub

ERRH:
11110   Select Case ERR.Number
        Case Else
11120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11130   End Select
11140   Resume EXITP

End Sub

Public Sub cmdWiden_Click()
' ** lngTpp * 4 = 60  : 1 grid point on form (my screen is 15 Twips Per Pixel).
' ** 60 * 4 = 240     : 4 grid points, a nice noticeable amount.
' ** 1440 Twips per Inch.
' ** Arbitrarily cap it at 10 in. wide.

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_Click"

        Dim lngAddedWidth As Long
        Dim blnSortHere As Boolean, blnCurrID As Boolean, blnCost As Boolean

11210   With Me
11220     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
11230       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
11240     End If
11250     blnSortHere = False
11260     blnCurrID = .Parent.chkIncludeCurrency
11270     blnCost = .Parent.chkShowCost
11280     lngAddedWidth = (lngTpp * WIDEN_MULT)
11290     If (.Width >= WIDEN_MAX) Or ((.Width + lngAddedWidth) > WIDEN_MAX) Then
11300       Beep
11310     Else
11320       .Width = (.Width + lngAddedWidth)
11330       .IsArch_bg1.Width = .Width
11340       .IsArch_bg2.Width = .Width
11350       If .Sort_line.Left = .taxcode_lbl.Left Then blnSortHere = True
11360       .taxcode_DESC_display.Left = (.taxcode_DESC_display.Left + lngAddedWidth)
11370       .taxcode.Left = .taxcode_DESC_display.Left
11380       .taxcode_lbl.Left = .taxcode.Left
11390       .taxcode_lbl_line.Left = .taxcode_lbl.Left
11400       If blnSortHere = True Then
11410         .Sort_line.Left = .taxcode_lbl.Left
11420         .Sort_lbl.Left = ((.taxcode_lbl.Left + .taxcode_lbl.Width) - lngSortLbl_Width)
11430       End If
11440       blnSortHere = False
11450       .IsIncome.Left = (.IsIncome.Left + lngAddedWidth)
11460       .IsExpense.Left = .IsIncome.Left
11470       If blnCurrID = True Then
11480         If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
11490         .curr_id.Left = (.curr_id.Left + lngAddedWidth)
11500         .curr_id_box.Left = .curr_id.Left
11510         .curr_id_forex_bg.Left = (.curr_id_box.Left - lngTpp)
11520         .curr_id_lbl.Left = (.curr_id.Left - lngCurrIDLbl_Offset)
11530         .curr_id_lbl_line.Left = .curr_id.Left
11540         If blnSortHere = True Then
11550           .Sort_line.Left = .curr_id_lbl_line.Left
11560           .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
11570         End If
11580         blnSortHere = False
11590       End If
11600       If blnCost = True Then
11610         If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
11620         .Cost.Left = (.Cost.Left + lngAddedWidth)
11630         .cost_box.Left = .Cost.Left
11640         .cost_forex_bg.Left = (.cost_box.Left - lngTpp)
11650         .cost_str.Left = .Cost.Left
11660         .cost_lbl.Left = .Cost.Left
11670         .cost_lbl_line.Left = .cost_lbl.Left
11680         If blnSortHere = True Then
11690           .Sort_line.Left = .cost_lbl.Left
11700           .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
11710         End If
11720         blnSortHere = False
11730       End If
11740       If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
11750       .PCash.Left = (.PCash.Left + lngAddedWidth)
11760       .pcash_box.Left = .PCash.Left
11770       .pcash_forex_bg.Left = (.pcash_box.Left - lngTpp)
11780       .pcash_str.Left = .PCash.Left
11790       .pcash_lbl.Left = .PCash.Left
11800       .pcash_lbl_line.Left = .pcash_lbl.Left
11810       If blnSortHere = True Then
11820         .Sort_line.Left = .pcash_lbl.Left
11830         .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
11840       End If
11850       blnSortHere = False
11860       If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
11870       .ICash.Left = (.ICash.Left + lngAddedWidth)
11880       .icash_box.Left = .ICash.Left
11890       .icash_forex_bg.Left = (.icash_box.Left - lngTpp)
11900       .icash_str.Left = .ICash.Left
11910       .icash_lbl.Left = .ICash.Left
11920       .icash_lbl_line.Left = .icash_lbl.Left
11930       If blnSortHere = True Then
11940         .Sort_line.Left = .icash_lbl.Left
11950         .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
11960       End If
11970       blnSortHere = False
11980       If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
11990       .totdesc.Width = (.totdesc.Width + lngAddedWidth)
12000       .totdesc_lbl.Width = .totdesc.Width
12010       .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
12020       If blnSortHere = True Then
12030         .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
12040         .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
12050       End If
12060       blnSortHere = False
12070     End If
12080     .FocusHolder.SetFocus
12090     If .cmdShorten.Enabled = False Then
            ' ** Since cmdWiden was clicked, cmdShorten is certainly available.
12100       .cmdShorten.Enabled = True
12110       .cmdShorten_raised_img.Visible = True
12120       .cmdShorten_raised_semifocus_dots_img.Visible = False
12130       .cmdShorten_raised_focus_img.Visible = False
12140       .cmdShorten_raised_focus_dots_img.Visible = False
12150       .cmdShorten_sunken_focus_dots_img.Visible = False
12160       .cmdShorten_raised_img_dis.Visible = False
12170     End If
          ' ** If there's nowhere to go, disable it.
12180     If (.Width >= WIDEN_MAX) Or ((.Width + lngAddedWidth) > WIDEN_MAX) Then
12190       .cmdWiden.Enabled = False
12200       .cmdWiden_raised_img_dis.Visible = True
12210       .cmdWiden_raised_img.Visible = False
12220       .cmdWiden_raised_semifocus_dots_img.Visible = False
12230       .cmdWiden_raised_focus_img.Visible = False
12240       .cmdWiden_raised_focus_dots_img.Visible = False
12250       .cmdWiden_sunken_focus_dots_img.Visible = False
12260     End If
12270   End With

EXITP:
12280   Exit Sub

ERRH:
12290   Select Case ERR.Number
        Case Else
12300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12310   End Select
12320   Resume EXITP

End Sub

Private Sub cmdWiden_KeyDown(KeyCode As Integer, Shift As Integer)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12420   intShiftDown = (Shift And acShiftMask) > 0
12430   intAltDown = (Shift And acAltMask) > 0
12440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12460     Select Case intRetVal
          Case vbKeyTab
12470       With Me
12480         intRetVal = 0
12490 On Error Resume Next
12500         .taxcode_DESC_display.SetFocus
12510         If ERR.Number <> 0 Then
12520 On Error GoTo ERRH
12530           .taxcode.SetFocus
12540         Else
12550 On Error GoTo ERRH
12560         End If
12570       End With
12580     End Select
12590   End If

        ' ** Shift keys.
12600   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12610     Select Case intRetVal
          Case vbKeyTab
12620       With Me
12630         intRetVal = 0
12640         .cmdShorten.SetFocus
12650       End With
12660     End Select
12670   End If

EXITP:
12680   KeyCode = intRetVal
12690   Exit Sub

ERRH:
12700   intRetVal = 0
12710   Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
12720   Case Else
12730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12740   End Select
12750   Resume EXITP

End Sub

Private Sub cmdWiden_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseMove"

12810   With Me
12820     If blnWiden_MouseDown = False Then
12830       Select Case blnWiden_Focus
            Case True
12840         .cmdWiden_raised_focus_dots_img.Visible = True
12850         .cmdWiden_raised_focus_img.Visible = False
12860       Case False
12870         .cmdWiden_raised_focus_img.Visible = True
12880         .cmdWiden_raised_focus_dots_img.Visible = False
12890       End Select
12900       .cmdWiden_raised_img.Visible = False
12910       .cmdWiden_raised_semifocus_dots_img.Visible = False
12920       .cmdWiden_sunken_focus_dots_img.Visible = False
12930       .cmdWiden_raised_img_dis.Visible = False
12940     End If
12950     If .cmdShorten_raised_focus_dots_img.Visible = True Or .cmdShorten_raised_focus_img.Visible = True Then
12960       Select Case blnShorten_Focus
            Case True
12970         .cmdShorten_raised_semifocus_dots_img.Visible = True
12980         .cmdShorten_raised_img.Visible = False
12990       Case False
13000         .cmdShorten_raised_img.Visible = True
13010         .cmdShorten_raised_semifocus_dots_img.Visible = False
13020       End Select
13030       .cmdShorten_raised_focus_img.Visible = False
13040       .cmdShorten_raised_focus_dots_img.Visible = False
13050       .cmdShorten_sunken_focus_dots_img.Visible = False
13060       .cmdShorten_raised_img_dis.Visible = False
13070     End If
13080   End With

EXITP:
13090   Exit Sub

ERRH:
13100   Select Case ERR.Number
        Case Else
13110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13120   End Select
13130   Resume EXITP

End Sub

Private Sub cmdWiden_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseUp"

13210   With Me
13220     .cmdWiden_raised_focus_dots_img.Visible = True
13230     .cmdWiden_raised_img.Visible = False
13240     .cmdWiden_raised_semifocus_dots_img.Visible = False
13250     .cmdWiden_raised_focus_img.Visible = False
13260     .cmdWiden_sunken_focus_dots_img.Visible = False
13270     .cmdWiden_raised_img_dis.Visible = False
13280     blnWiden_MouseDown = False
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

Private Sub cmdWiden_LostFocus()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_LostFocus"

13410   With Me
13420     .cmdWiden_raised_img.Visible = True
13430     .cmdWiden_raised_semifocus_dots_img.Visible = False
13440     .cmdWiden_raised_focus_img.Visible = False
13450     .cmdWiden_raised_focus_dots_img.Visible = False
13460     .cmdWiden_sunken_focus_dots_img.Visible = False
13470     .cmdWiden_raised_img_dis.Visible = False
13480     blnWiden_Focus = False
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

Public Sub ShowCost_Sub(blnShow As Boolean)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowCost_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean, blnCurrID As Boolean
        Dim lngNewFormWidth As Long
        Dim lngDiff1 As Long

13610   With Me
13620     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
13630       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
13640     End If
13650     blnCurrID = .Parent.chkIncludeCurrency
13660     lngDiff1 = (.totdesc.Width - lngTotDesc_Width)
13670     blnSortHere = False: blnResort = False
13680     Select Case blnShow
          Case True
13690       Select Case blnCurrID
            Case True
13700         lngNewFormWidth = lngForm_Width
13710       Case False
13720         lngNewFormWidth = (lngForm_Width - lngCurrID_Width)
13730       End Select
13740       .Width = lngNewFormWidth
13750       If .Sort_line.Left = .taxcode_lbl.Left Then blnSortHere = True
13760       Select Case blnCurrID
            Case True
13770         .taxcode.Left = (lngTaxCode_Left + lngDiff1)
13780       Case False
13790         .taxcode.Left = ((lngTaxCode_Left + lngDiff1) - lngCurrID_Width)
13800       End Select
13810       .taxcode_DESC_display.Left = .taxcode.Left
13820       .taxcode_lbl.Left = .taxcode.Left
13830       .taxcode_lbl_line.Left = .taxcode_lbl.Left
13840       If blnSortHere = True Then
13850         .Sort_line.Left = .taxcode_lbl.Left
13860         .Sort_lbl.Left = ((.taxcode_lbl.Left + .taxcode_lbl.Width) - lngSortLbl_Width)
13870       End If
13880       blnSortHere = False
13890       .IsIncome.Left = (.taxcode.Left - lngIsIncDed_Offset)
13900       .IsExpense.Left = .IsIncome.Left
13910       If blnCurrID = True Then
13920         If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
13930         .curr_id.Left = (lngCurrID_Left + lngDiff1)
13940         .curr_id_lbl.Left = (.curr_id.Left - lngCurrIDLbl_Offset)
13950         .curr_id_lbl_line.Left = .curr_id.Left
13960         .curr_id_box.Left = .curr_id.Left
13970         .curr_id_forex_bg.Left = (.curr_id.Left - lngTpp)
13980         If blnSortHere = True Then
13990           .Sort_line.Left = .curr_id_lbl_line.Left
14000           .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
14010         End If
14020         blnSortHere = False
14030       End If
14040       .Cost.Left = (lngCost_Left + lngDiff1)
14050       .cost_str.Left = .Cost.Left
14060       .cost_lbl.Left = .Cost.Left
14070       .cost_lbl_line.Left = .cost_lbl.Left
14080       .cost_box.Left = .cost_lbl.Left
14090       .cost_forex_bg.Left = (.cost_lbl.Left - lngTpp)
14100       Select Case blnCurrID
            Case True
14110         .cost_str.Visible = True
14120         .Cost.Visible = False
14130       Case False
14140         Select Case gblnHasForExThis
              Case True
14150           .cost_str.Visible = True
14160           .Cost.Visible = False
14170         Case False
14180           .Cost.Visible = True
14190           .cost_str.Visible = False
14200         End Select
14210       End Select
14220       .cost_lbl.Visible = True
14230       .cost_lbl_line.Visible = True
14240       .cost_box.Visible = True
14250       .cost_forex_bg.Visible = True
14260       .IsArch_bg1.Width = lngNewFormWidth
14270       .IsArch_bg2.Width = lngNewFormWidth
14280     Case False
14290       Select Case blnCurrID
            Case True
14300         lngNewFormWidth = ((lngForm_Width + lngDiff1) - lngCost_Width)
14310       Case False
14320         lngNewFormWidth = ((lngForm_Width + lngDiff1) - (lngCurrID_Width + lngCost_Width))
14330       End Select
14340       If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
14350       .Cost.Visible = False
14360       .cost_str.Visible = False
14370       .cost_lbl.Visible = False
14380       .cost_lbl_line.Visible = False
14390       .cost_box.Visible = False
14400       .cost_forex_bg.Visible = False
14410       .Cost.Left = 0&
14420       .cost_str.Left = 0&
14430       .cost_lbl.Left = 0&
14440       .cost_lbl_line.Left = 0&
14450       .cost_box.Left = 0&
14460       .cost_forex_bg.Left = 0&
14470       If blnSortHere = True Then
14480         .Sort_lbl.Visible = False
14490         .Sort_line.Visible = False
14500         blnResort = True
14510       End If
14520       blnSortHere = False
14530       If blnCurrID = True Then
14540         If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
14550         .curr_id.Left = ((lngCurrID_Left + lngDiff1) - lngCost_Width)
14560         .curr_id_lbl.Left = (.curr_id.Left - lngCurrIDLbl_Offset)
14570         .curr_id_lbl_line.Left = .curr_id.Left
14580         .curr_id_box.Left = .curr_id.Left
14590         .curr_id_forex_bg.Left = (.curr_id.Left - lngTpp)
14600         If blnSortHere = True Then
14610           .Sort_line.Left = .curr_id_lbl_line.Left
14620           .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
14630         End If
14640         blnSortHere = False
14650       End If
14660       If .Sort_line.Left = .taxcode_lbl.Left Then blnSortHere = True
14670       Select Case blnCurrID
            Case True
14680         .taxcode.Left = ((lngTaxCode_Left + lngDiff1) - lngCost_Width)
14690       Case False
14700         .taxcode.Left = ((lngTaxCode_Left + lngDiff1) - (lngCost_Width + lngCurrID_Width))
14710       End Select
14720       .taxcode_DESC_display.Left = .taxcode.Left
14730       .taxcode_lbl.Left = .taxcode.Left
14740       .taxcode_lbl_line.Left = .taxcode_lbl.Left
14750       If blnSortHere = True Then
14760         .Sort_line.Left = .taxcode_lbl.Left
14770         .Sort_lbl.Left = ((.taxcode_lbl.Left + .taxcode_lbl.Width) - lngSortLbl_Width)
14780       End If
14790       blnSortHere = False
14800       .IsIncome.Left = (.taxcode.Left - lngIsIncDed_Offset)
14810       .IsExpense.Left = .IsIncome.Left
14820       .IsArch_bg1.Width = lngNewFormWidth
14830       .IsArch_bg2.Width = lngNewFormWidth
14840       .Width = lngNewFormWidth
14850     End Select
14860     If blnResort = True Then
14870       SortNow "Form_Load"  ' ** Procedure: Below.
14880     End If
14890   End With

EXITP:
14900   Exit Sub

ERRH:
14910   Select Case ERR.Number
        Case Else
14920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14930   End Select
14940   Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnShow As Boolean)

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean, blnShowCost As Boolean
        Dim lngNewFormWidth As Long
        Dim lngDiff1 As Long

15010   With Me
15020     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
15030       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
15040     End If
15050     blnShowCost = .Parent.chkShowCost
15060     blnSortHere = False: blnResort = False
15070     lngDiff1 = (.totdesc.Width - lngTotDesc_Width)
15080     Select Case blnShow
          Case True
15090       Select Case blnShowCost
            Case True
15100         lngNewFormWidth = (lngForm_Width + lngDiff1)
15110       Case False
15120         lngNewFormWidth = ((lngForm_Width + lngDiff1) - lngCost_Width)
15130       End Select
15140       .Width = lngNewFormWidth
15150       If .Sort_line.Left = .taxcode_lbl.Left Then blnSortHere = True
15160       Select Case blnShowCost
            Case True
15170         .taxcode.Left = (lngTaxCode_Left + lngDiff1)
15180       Case False
15190         .taxcode.Left = ((lngTaxCode_Left + lngDiff1) - lngCost_Width)
15200       End Select
15210       .taxcode_DESC_display.Left = .taxcode.Left
15220       .taxcode_lbl.Left = .taxcode.Left
15230       .taxcode_lbl_line.Left = .taxcode_lbl.Left
15240       If blnSortHere = True Then
15250         .Sort_line.Left = .taxcode_lbl.Left
15260         .Sort_lbl.Left = ((.taxcode_lbl.Left + .taxcode_lbl.Width) - lngSortLbl_Width)
15270       End If
15280       blnSortHere = False
15290       .IsExpense.Left = (.taxcode.Left - lngIsIncDed_Offset)
15300       .IsIncome.Left = .IsExpense.Left
15310       Select Case blnShowCost
            Case True
15320         .curr_id.Left = (lngCurrID_Left + lngDiff1)
15330       Case False
15340         .curr_id.Left = ((lngCurrID_Left + lngDiff1) - lngCost_Width)
15350       End Select
15360       .curr_id_lbl.Left = (.curr_id.Left - lngCurrIDLbl_Offset)
15370       .curr_id_lbl_line.Left = .curr_id.Left
15380       .curr_id_box.Left = .curr_id.Left
15390       .curr_id_forex_bg.Left = (.curr_id.Left - lngTpp)
15400       .curr_id.Visible = True
15410       .curr_id_lbl.Visible = True
15420       .curr_id_lbl_line.Visible = True
15430       .curr_id_box.Visible = True
15440       .curr_id_forex_bg.Visible = True
15450       .IsArch_bg1.Width = lngNewFormWidth
15460       .IsArch_bg2.Width = lngNewFormWidth
15470     Case False
15480       Select Case blnShowCost  ' ** lngDiff1 will usually be Zero; difference when cmdWiden used.
            Case True
15490         lngNewFormWidth = ((lngForm_Width + lngDiff1) - lngCurrID_Width)
15500       Case False
15510         lngNewFormWidth = ((lngForm_Width + lngDiff1) - (lngCost_Width + lngCurrID_Width))
15520       End Select
15530       If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
15540       .curr_id.Visible = False
15550       .curr_id_lbl.Visible = False
15560       .curr_id_lbl_line.Visible = False
15570       .curr_id_box.Visible = False
15580       .curr_id_forex_bg.Visible = False
15590       .curr_id.Left = 0&
15600       .curr_id_lbl.Left = 0&
15610       .curr_id_lbl_line.Left = 0&
15620       .curr_id_box.Left = 0&
15630       .curr_id_forex_bg.Left = 0&
15640       If blnSortHere = True Then
15650         .Sort_line.Visible = False
15660         .Sort_line.Left = 0&
15670         .Sort_lbl.Visible = False
15680         .Sort_lbl.Left = 0&
15690         blnResort = True
15700       End If
15710       blnSortHere = False
15720       If .Sort_line.Left = .taxcode_lbl.Left Then blnSortHere = True
15730       Select Case blnShowCost
            Case True
15740         .taxcode.Left = ((lngTaxCode_Left + lngDiff1) - lngCurrID_Width)
15750       Case False
15760         .taxcode.Left = ((lngTaxCode_Left + lngDiff1) - (lngCost_Width + lngCurrID_Width))
15770       End Select
15780       .taxcode_DESC_display.Left = .taxcode.Left
15790       .taxcode_lbl.Left = .taxcode.Left
15800       .taxcode_lbl_line.Left = .taxcode_lbl.Left
15810       If blnSortHere = True Then
15820         .Sort_line.Left = .taxcode_lbl.Left
15830         .Sort_lbl.Left = ((.taxcode_lbl.Left + .taxcode_lbl.Width) - lngSortLbl_Width)
15840       End If
15850       blnSortHere = False
15860       .IsExpense.Left = (.taxcode.Left - lngIsIncDed_Offset)
15870       .IsIncome.Left = .IsExpense.Left
15880       .IsArch_bg1.Width = lngNewFormWidth
15890       .IsArch_bg2.Width = lngNewFormWidth
15900       .Width = lngNewFormWidth
15910       If blnResort = True Then
15920         SortNow "Form_Load"  ' ** Procedure: Below.
15930       End If
15940     End Select
15950   End With

EXITP:
15960   Exit Sub

ERRH:
15970   Select Case ERR.Number
        Case Else
15980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15990   End Select
16000   Resume EXITP

End Sub

Public Function RecCnt() As Long

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

16110   lngRetVal = 0&

16120   With Me
16130     Set rst = .RecordsetClone
16140     With rst
16150       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
16160       Else
16170         .MoveLast
16180         lngRetVal = .RecordCount
16190       End If
16200       .Close
16210     End With
16220   End With

EXITP:
16230   Set rst = Nothing
16240   RecCnt = lngRetVal
16250   Exit Function

ERRH:
16260   lngRetVal = 0&
16270   Select Case ERR.Number
        Case Else
16280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16290   End Select
16300   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

16410   Select Case intWhere
        Case acCmdRecordsGoToFirst
16420     DoCmd.RunCommand acCmdRecordsGoToFirst
16430   Case acCmdRecordsGoToLast
16440     DoCmd.RunCommand acCmdRecordsGoToLast
16450   Case acCmdRecordsGoToPrevious
16460     DoCmd.RunCommand acCmdRecordsGoToPrevious
16470   Case acCmdRecordsGoToNext
16480     DoCmd.RunCommand acCmdRecordsGoToNext
16490   Case Else
16500     If IsMissing(varID) = False Then
16510       With Me
16520         Set rst = .RecordsetClone
16530         With rst
16540           .FindFirst "[journalno] = " & CStr(varID)
16550           If .NoMatch = False Then
16560             Me.Bookmark = .Bookmark
16570           End If
16580           .Close
16590         End With
16600       End With
16610     End If
16620   End Select

EXITP:
16630   Set rst = Nothing
16640   Exit Sub

ERRH:
16650   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
16660   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
16670   Case Else
16680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16690   End Select
16700   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [transdate], [journaltype], [journalno]"

16810   With Me
16820     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
16830       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
16840     End If
16850     .Controls(strSortLbl).Visible = False
16860     .Controls(strSortLine).Visible = False
16870     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
16880     If strProc = "Form_Load" Then
16890       strCalled = "transdate"
16900       strSortNow = strSortOrig
16910       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
16920       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
16930       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
16940       .Controls(strSortLbl).Left = lngSortLbl_Left
16950       .Controls(strSortLine).Left = lngSortLine_Left
16960       .Controls(strSortLine).Width = lngSortLine_Width
16970       .Controls(strSortLbl).Caption = strArwDn   ' ** Original is descending.
16980       .Controls(strSortLbl).ForeColor = CLR_DKRED
16990     Else
17000       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
17010       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
17020       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
17030       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
17040       .Controls(strSortLbl).Left = lngSortLbl_Left
17050       .Controls(strSortLine).Left = lngSortLine_Left
17060       .Controls(strSortLine).Width = lngSortLine_Width
17070       If strCalled = "transdate" Then
17080         strSortAsc = "[transdate], [journaltype], [journalno]"
17090         If strSortNow = strSortAsc Then
17100           strSortNow = strSortOrig   ' ** Original is descending: "[transdate] DESC, [journaltype] ASC, [journalno] DESC"
17110           .Controls(strSortLbl).Caption = strArwDn
17120           .Controls(strSortLbl).ForeColor = CLR_DKRED
17130         Else
17140           strSortNow = strSortAsc
17150           .Controls(strSortLbl).Caption = strArwUp
17160           .Controls(strSortLbl).ForeColor = CLR_DKBLU
17170         End If
17180       Else
17190         .Controls(strSortLbl).Caption = strArwUp
17200         .Controls(strSortLbl).ForeColor = CLR_DKBLU
17210         Select Case strCalled
              Case "journaltype"
17220           strSortAsc = "[journaltype], [transdate], [journalno]"
17230         Case "shareface"
17240           strSortAsc = "[sharefacex]" & strStdAsc
17250         Case "totdesc"
17260           strSortAsc = "[totdesc]" & strStdAsc
17270         Case "ICash"
17280           strSortAsc = "[ICash]" & strStdAsc
17290         Case "PCash"
17300           strSortAsc = "[PCash]" & strStdAsc
17310         Case "Cost"
17320           strSortAsc = "[Cost]" & strStdAsc
17330         Case "curr_id"
17340           strSortAsc = "[curr_code]" & strStdAsc
17350           lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
17360           lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
17370           lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
17380           .Controls(strSortLbl).Left = lngSortLbl_Left
17390           .Controls(strSortLine).Left = lngSortLine_Left
17400           .Controls(strSortLine).Width = lngSortLine_Width
17410         Case "taxcode"
17420           strSortAsc = "[taxcode_DESC]" & strStdAsc
17430         End Select
17440         If strSortNow = strSortAsc Then
17450           intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
17460           Select Case intCnt
                Case 1
17470             strTmp01 = strSortAsc & " DESC"
17480           Case 2
17490             intPos01 = InStr(strSortAsc, ",")
17500             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
17510             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
17520           Case 3
17530             intPos01 = InStr(strSortAsc, ",")
17540             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
17550             strTmp02 = Mid(strSortAsc, intPos01)
17560             intPos01 = InStr(2, strTmp02, ",")
17570             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
17580             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
17590           Case 4
17600             intPos01 = InStr(strSortAsc, ",")
17610             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
17620             strTmp02 = Mid(strSortAsc, intPos01)
17630             intPos01 = InStr(2, strTmp02, ",")
17640             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
17650             strTmp02 = Mid(strTmp02, intPos01)
17660             intPos01 = InStr(2, strTmp02, ",")
17670             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
17680             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
17690           Case 5
17700             intPos01 = InStr(strSortAsc, ",")
17710             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
17720             strTmp02 = Mid(strSortAsc, intPos01)
17730             intPos01 = InStr(2, strTmp02, ",")
17740             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
17750             strTmp02 = Mid(strTmp02, intPos01)
17760             intPos01 = InStr(2, strTmp02, ",")
17770             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
17780             strTmp02 = Mid(strTmp02, intPos01)
17790             intPos01 = InStr(2, strTmp02, ",")
17800             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
17810             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
17820           End Select
17830           strSortNow = strTmp01
17840           .Controls(strSortLbl).Caption = strArwDn
17850           .Controls(strSortLbl).ForeColor = CLR_DKRED
17860         Else
17870           strSortNow = strSortAsc
17880         End If
17890       End If
17900     End If
17910     .Controls(strSortLbl).Visible = True
17920     .Controls(strSortLine).Visible = True
17930     .OrderBy = strSortNow
17940     .OrderByOn = True
17950   End With

EXITP:
17960   Exit Sub

ERRH:
17970   Select Case ERR.Number
        Case Else
17980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17990   End Select
18000   Resume EXITP

End Sub
