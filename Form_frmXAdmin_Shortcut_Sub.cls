VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmXAdmin_Shortcut_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmXAdmin_Shortcut_Sub"

'VGC 04/17/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Show All:         A {opgSearchMode_optAll on frmXAdmin_Shortcut}
' **   Combo Keys:       B {fsp_keycombo}
' **   Close:            C {cmdClose on frmXAdmin_Shortcut}
' **   Description:      D {fsp_description}
' **   Forms:            F {lbxForms on frmXAdmin_Shortcut}
' **   Key:              K {fsp_keycode}
' **   Menus:            M {cmbSections on frmXAdmin_Shortcut}
' **   Show Form Name:   N {chkShowFormName on frmXAdmin_Shortcut}
' **   Filter Only:      O {opgSearchMode_optSpecified on frmXAdmin_Shortcut}
' **   KeyDown Type:     T {keydowntype_type}
' **   Include Misc:     U {chkIncludeMiscForms on frmXAdmin_Shortcut}
' **   Exit:             X {cmdClose on frmXAdmin_Shortcut}

' ** cmbSections combo box constants:
Private Const CBX_S_SEC   As Integer = 0  'fsfd_section
'Private Const CBX_S_CAP   As Integer = 1  'frm_header_section
'Private Const CBX_S_NAM   As Integer = 2  'frm_name_section
'Private Const CBX_S_DID   As Integer = 3  'dbs_id
'Private Const CBX_S_FID   As Integer = 4  'frm_id
'Private Const CBX_S_FSFID As Integer = 5  'fsp_order

Private Const strSortOrig As String = "[fsfd_section], [fsfd_order], [fsfd_level], [keydowntype_order], [fsp_keycode]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngForm_Width As Long, lngFrmHeader_Left As Long, lngKeyDownType_Left As Long
Private lngKeyCombo_Left As Long, lngKeyCode_Left As Long, lngDescription_Left As Long

Private blnFromAfterUpdate As Boolean
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
150       If strCallingForm = "frmXAdmin_Shortcut" Then

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170         blnFromAfterUpdate = False

180         .form_width_line.Left = 0&  ' ** Just get it out of the way.

190         lngForm_Width = .Width
200         lngSortLbl_Width = .Sort_lbl.Width
210         lngFrmHeader_Left = .frm_header_section.Left
220         lngKeyDownType_Left = .keydowntype_type.Left
230         lngKeyCombo_Left = .fsp_keycombo.Left
240         lngKeyCode_Left = .fsp_keycode.Left
250         lngDescription_Left = .fsp_description.Left

260         .FocusHolder.SetFocus

270       Else
280         Cancel = -1
290       End If
300     End With

310     If Cancel = -1 Then
320       Beep
330       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
340     End If

EXITP:
350     Exit Sub

ERRH:
360     Select Case ERR.Number
        Case Else
370       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
380     End Select
390     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
420     Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub Form_Current()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim lngFsfID As Long, lngFsfdSec As Long

510     With Me
520       If IsNull(.Parent.lbxForms) = False Then
530         lngFsfID = .Parent.lbxForms
540         If lngFsfID > 0& Then
550           lngFsfdSec = .Parent.lbxForms.Column(5)
560           If .fsfd_section <> lngFsfdSec Or .fsf_id <> lngFsfID Then
570             .Parent.lbxForms = 0  ' ** Clear the list box.
580           End If
590         End If
600         If IsNull(.Parent.cmbSections) = False Then
610           lngFsfdSec = .Parent.cmbSections.Column(CBX_S_SEC)
620           If .fsfd_section <> lngFsfdSec Then
630             .Parent.cmbSections = Null  ' ** Clear the combo box.
640           End If
650         End If
660       End If
670       DoCmd.Hourglass False
680     End With

EXITP:
690     Exit Sub

ERRH:
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

810     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
820     intShiftDown = (Shift And acShiftMask) > 0
830     intAltDown = (Shift And acAltMask) > 0
840     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionalities:
        ' **   Combo Keys:       B {fsp_keycombo}
        ' **   Description:      D {fsp_description}
        ' **   Key:              K {fsp_keycode}
        ' **   KeyDown Type:     T {keydowntype_type}
        ' **   Exit:             X {cmdClose on frmXAdmin_Shortcut}

        ' ** Plain keys.
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyUp
870         With Me
880           intRetVal = 0
890           MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
900         End With
910       Case vbKeyDown
920         With Me
930           intRetVal = 0
940           MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
950         End With
960       Case vbKeyEscape
970         With Me
980           intRetVal = 0
990           .Parent.cmdClose_Click  ' ** Form Procedure: frmXAdmin_Shortcut.
1000        End With
1010      End Select
1020    End If

        ' ** Alt keys.
1030    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1040      Select Case intRetVal
          Case vbKeyB
1050        With Me
1060          intRetVal = 0
1070          .fsp_keycombo.SetFocus
1080        End With
1090      Case vbKeyD
1100        With Me
1110          intRetVal = 0
1120          .fsp_description.SetFocus
1130        End With
1140      Case vbKeyK
1150        With Me
1160          intRetVal = 0
1170          .fsp_keycode.SetFocus
1180        End With
1190      Case vbKeyT
1200        With Me
1210          intRetVal = 0
1220          .keydowntype_type.SetFocus
1230        End With
1240      Case vbKeyX
1250        With Me
1260          intRetVal = 0
1270          .Parent.cmdClose_Click  ' ** Form Procedure: frmXAdmin_Shortcut.
1280        End With
1290      End Select
1300    End If

        ' ** Ctrl keys.
1310    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1320      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1330        With Me
1340          intRetVal = 0
1350          DoCmd.SelectObject acForm, .Parent.Name, False
1360          .Parent.cmdClose.SetFocus
1370        End With
1380      Case vbKeyPageUp, vbKeyHome
1390        With Me
1400          intRetVal = 0
1410          MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1420        End With
1430      Case vbKeyPageDown, vbKeyEnd
1440        With Me
1450          intRetVal = 0
1460          MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1470        End With
1480      End Select
1490    End If

        ' ** Ctrl-Shift keys.
1500    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1510      Select Case intRetVal
          Case vbKeyF
1520        With Me
1530          intRetVal = 0
1540          DoCmd.SelectObject acForm, .Parent.Name, False
1550          .Parent.FocusHolder.SetFocus
1560        End With
1570      Case vbKeyTab, vbKeyReturn
1580        With Me
1590          intRetVal = 0
1600          DoCmd.SelectObject acForm, .Parent.Name, False
1610          .Parent.chkIncludeMiscForms.SetFocus
1620        End With
1630      End Select
1640    End If

EXITP:
1650    KeyCode = intRetVal
1660    Exit Sub

ERRH:
1670    intRetVal = 0
1680    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1690    Case Else
1700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1710    End Select
1720    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1820    intShiftDown = (Shift And acShiftMask) > 0
1830    intAltDown = (Shift And acAltMask) > 0
1840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1870        With Me
1880          intRetVal = 0
1890          lngRecsCur = RecCnt  ' ** Function: Below.
1900          If .CurrentRecord < lngRecsCur Then
1910            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1920            .FocusHolder.SetFocus
1930          Else
1940            DoCmd.SelectObject acForm, .Parent.Name, False
1950            .Parent.cmdClose.SetFocus
1960          End If
1970        End With
1980      End Select
1990    End If

        ' ** Shift keys.
2000    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2010      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2020        With Me
2030          intRetVal = 0
2040          If .CurrentRecord > 1 Then
2050            .MoveRec acCmdRecordsGoToPrevious
2060            .FocusHolder.SetFocus
2070          Else
2080            DoCmd.SelectObject acForm, .Parent.Name, False
2090            .Parent.chkIncludeMiscForms.SetFocus
2100          End If
2110        End With
2120      End Select
2130    End If

EXITP:
2140    KeyCode = intRetVal
2150    Exit Sub

ERRH:
2160    intRetVal = 0
2170    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2180    Case Else
2190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2200    End Select
2210    Resume EXITP

End Sub

Private Sub frm_name_section_lbl_DblClick(Cancel As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "frm_name_section_lbl_DblClick"

2310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

End Sub

Private Sub frm_header_section_lbl_DblClick(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "frm_header_section_lbl_DblClick"

2410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2420    Exit Sub

ERRH:
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Sub

Private Sub keydowntype_type_lbl_DblClick(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "keydowntype_type_lbl_DblClick"

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

Private Sub keydowntype_type_Enter()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "keydowntype_type_Enter"

2610    With Me
2620      Select Case .fsfd_order
          Case 0
2630        .keydowntype_type.BackColor = MY_CLR_BGE
2640      Case 2
2650        .keydowntype_type.BackColor = CLR_WHT
2660      End Select
2670    End With

EXITP:
2680    Exit Sub

ERRH:
2690    Select Case ERR.Number
        Case Else
2700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2710    End Select
2720    Resume EXITP

End Sub

Private Sub keydowntype_type_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "keydowntype_type_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2820    intShiftDown = (Shift And acShiftMask) > 0
2830    intAltDown = (Shift And acAltMask) > 0
2840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2870        With Me
2880          intRetVal = 0
2890          .fsp_keycombo.SetFocus
2900        End With
2910      End Select
2920    End If

        ' ** Shift keys.
2930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2950        With Me
2960          intRetVal = 0
2970          If .CurrentRecord > 1 Then
2980            .MoveRec acCmdRecordsGoToPrevious
2990            .fsp_description.SetFocus
3000          Else
3010            DoCmd.SelectObject acForm, .Parent.Name, False
3020            .Parent.chkIncludeMiscForms.SetFocus
3030          End If
3040        End With
3050      End Select
3060    End If

EXITP:
3070    KeyCode = intRetVal
3080    Exit Sub

ERRH:
3090    intRetVal = 0
3100    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3110    Case Else
3120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3130    End Select
3140    Resume EXITP

End Sub

Private Sub keydowntype_type_Exit(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "keydowntype_type_Exit"

3210    With Me
3220      .keydowntype_type.BackColor = CLR_WHT
3230    End With

EXITP:
3240    Exit Sub

ERRH:
3250    Select Case ERR.Number
        Case Else
3260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3270    End Select
3280    Resume EXITP

End Sub

Private Sub fsp_keycombo_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "fsp_keycombo_lbl_DblClick"

3310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub fsp_keycombo_Enter()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "fsp_keycombo_Enter"

3410    With Me
3420      Select Case .fsfd_order
          Case 0
3430        .fsp_keycombo.BackColor = MY_CLR_BGE
3440      Case 2
3450        .fsp_keycombo.BackColor = CLR_WHT
3460      End Select
3470    End With

EXITP:
3480    Exit Sub

ERRH:
3490    Select Case ERR.Number
        Case Else
3500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3510    End Select
3520    Resume EXITP

End Sub

Private Sub fsp_keycombo_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "fsp_keycombo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3620    intShiftDown = (Shift And acShiftMask) > 0
3630    intAltDown = (Shift And acAltMask) > 0
3640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3670        With Me
3680          intRetVal = 0
3690          .fsp_keycode.SetFocus
3700        End With
3710      End Select
3720    End If

        ' ** Shift keys.
3730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3750        With Me
3760          intRetVal = 0
3770          .keydowntype_type.SetFocus
3780        End With
3790      End Select
3800    End If

EXITP:
3810    KeyCode = intRetVal
3820    Exit Sub

ERRH:
3830    intRetVal = 0
3840    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3850    Case Else
3860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3870    End Select
3880    Resume EXITP

End Sub

Private Sub fsp_keycombo_Exit(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "fsp_keycombo_Exit"

3910    With Me
3920      .fsp_keycombo.BackColor = CLR_WHT
3930    End With

EXITP:
3940    Exit Sub

ERRH:
3950    Select Case ERR.Number
        Case Else
3960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3970    End Select
3980    Resume EXITP

End Sub

Private Sub fsp_keycode_lbl_DblClick(Cancel As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "fsp_keycode_lbl_DblClick"

4010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Sub fsp_keycode_Enter()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "fsp_keycode_Enter"

4110    With Me
4120      Select Case .fsfd_order
          Case 0
4130        .fsp_keycode.BackColor = MY_CLR_BGE
4140      Case 2
4150        .fsp_keycode.BackColor = CLR_WHT
4160      End Select
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

Private Sub fsp_keycode_KeyDown(KeyCode As Integer, Shift As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "fsp_keycode_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4320    intShiftDown = (Shift And acShiftMask) > 0
4330    intAltDown = (Shift And acAltMask) > 0
4340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4370        With Me
4380          intRetVal = 0
4390          .fsp_description.SetFocus
4400        End With
4410      End Select
4420    End If

        ' ** Shift keys.
4430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4450        With Me
4460          intRetVal = 0
4470          .fsp_keycombo.SetFocus
4480        End With
4490      End Select
4500    End If

EXITP:
4510    KeyCode = intRetVal
4520    Exit Sub

ERRH:
4530    intRetVal = 0
4540    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4550    Case Else
4560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4570    End Select
4580    Resume EXITP

End Sub

Private Sub fsp_keycode_Exit(Cancel As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "fsp_keycode_Exit"

4610    With Me
4620      .fsp_keycode.BackColor = CLR_WHT
4630    End With

EXITP:
4640    Exit Sub

ERRH:
4650    Select Case ERR.Number
        Case Else
4660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4670    End Select
4680    Resume EXITP

End Sub

Private Sub fsp_description_lbl_DblClick(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "fsp_description_lbl_DblClick"

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

Private Sub fsp_description_Enter()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "fsp_description_Enter"

4810    With Me
4820      Select Case .fsfd_order
          Case 0
4830        .fsp_description.BackColor = MY_CLR_BGE
4840      Case 2
4850        .fsp_description.BackColor = CLR_WHT
4860      End Select
4870    End With

EXITP:
4880    Exit Sub

ERRH:
4890    Select Case ERR.Number
        Case Else
4900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4910    End Select
4920    Resume EXITP

End Sub

Private Sub fsp_description_KeyDown(KeyCode As Integer, Shift As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "fsp_description_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5020    intShiftDown = (Shift And acShiftMask) > 0
5030    intAltDown = (Shift And acAltMask) > 0
5040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5070        With Me
5080          intRetVal = 0
5090          lngRecsCur = RecCnt  ' ** Function: Below.
5100          If .CurrentRecord < lngRecsCur Then
5110            .MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5120            .keydowntype_type.SetFocus
5130          Else
5140            DoCmd.SelectObject acForm, .Parent.Name, False
5150            .Parent.cmdClose.SetFocus
5160          End If
5170        End With
5180      End Select
5190    End If

        ' ** Shift keys.
5200    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5210      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5220        With Me
5230          intRetVal = 0
5240          .fsp_keycode.SetFocus
5250        End With
5260      End Select
5270    End If

EXITP:
5280    KeyCode = intRetVal
5290    Exit Sub

ERRH:
5300    intRetVal = 0
5310    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5320    Case Else
5330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5340    End Select
5350    Resume EXITP

End Sub

Private Sub fsp_description_Exit(Cancel As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "fsp_description_Exit"

5410    With Me
5420      .fsp_description.BackColor = CLR_WHT
5430    End With

EXITP:
5440    Exit Sub

ERRH:
5450    Select Case ERR.Number
        Case Else
5460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5470    End Select
5480    Resume EXITP

End Sub

Public Sub AfterUpdate_Set(blnAfter As Boolean)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "AfterUpdate_Set"

5510    blnFromAfterUpdate = blnAfter

EXITP:
5520    Exit Sub

ERRH:
5530    Select Case ERR.Number
        Case Else
5540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5550    End Select
5560    Resume EXITP

End Sub

Public Sub ShowFormName_Sub(blnShow As Boolean)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowFormName_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim lngDiff As Long

5610    With Me
5620      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5630        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5640      End If
5650      blnSortHere = False: blnResort = False
5660      Select Case blnShow
          Case True
5670        .Width = lngForm_Width
5680        If .Sort_line.Left = .fsp_description_lbl.Left Then blnSortHere = True
5690        .fsp_description.Left = lngDescription_Left
5700        .fsp_description_bg1.Left = (.fsp_description.Left - lngTpp)
5710        .fsp_description_bg2.Left = .fsp_description_bg1.Left
5720        .fsp_description_lbl.Left = .fsp_description.Left
5730        .fsp_description_lbl_line.Left = .fsp_description_lbl.Left
5740        If blnSortHere = True Then
5750          .Sort_line.Left = .fsp_description_lbl.Left
5760          .Sort_lbl.Left = ((.fsp_description_lbl.Left + .fsp_description_lbl.Width) - lngSortLbl_Width)
5770        End If
5780        blnSortHere = False
5790        If .Sort_line.Left = .fsp_keycode_lbl.Left Then blnSortHere = True
5800        .fsp_keycode.Left = lngKeyCode_Left
5810        .fsp_keycode_bg1.Left = (.fsp_keycode.Left - lngTpp)
5820        .fsp_keycode_bg2.Left = .fsp_keycode_bg1.Left
5830        .fsp_keycode_lbl.Left = .fsp_keycode.Left
5840        .fsp_keycode_lbl_line.Left = .fsp_keycode_lbl.Left
5850        If blnSortHere = True Then
5860          .Sort_line.Left = .fsp_keycode_lbl.Left
5870          .Sort_lbl.Left = ((.fsp_keycode_lbl.Left + .fsp_keycode_lbl.Width) - lngSortLbl_Width)
5880        End If
5890        blnSortHere = False
5900        If .Sort_line.Left = .fsp_keycombo_lbl.Left Then blnSortHere = True
5910        .fsp_keycombo.Left = (lngKeyCombo_Left - lngDiff)
            '.fsp_keycombo_bg1.Left = (.fsp_keycombo.Left - lngTpp)
            '.fsp_keycombo_bg2.Left = .fsp_keycombo_bg1.Left
5920        .fsp_keycombo_lbl.Left = .fsp_keycombo.Left
5930        .fsp_keycombo_lbl_line.Left = .fsp_keycombo_lbl.Left
5940        .fsp_keycombo_plain_bg1.Left = (.fsp_keycombo.Left - lngTpp)
5950        .fsp_keycombo_plain_bg2.Left = .fsp_keycombo_plain_bg1.Left
5960        .fsp_keycombo_alt_bg1.Left = .fsp_keycombo_plain_bg1.Left
5970        .fsp_keycombo_alt_bg2.Left = .fsp_keycombo_plain_bg1.Left
5980        .fsp_keycombo_shift_bg1.Left = .fsp_keycombo_plain_bg1.Left
5990        .fsp_keycombo_shift_bg2.Left = .fsp_keycombo_plain_bg1.Left
6000        .fsp_keycombo_ctrl_bg1.Left = .fsp_keycombo_plain_bg1.Left
6010        .fsp_keycombo_ctrl_bg2.Left = .fsp_keycombo_plain_bg1.Left
6020        .fsp_keycombo_ctrl_alt_bg1.Left = .fsp_keycombo_plain_bg1.Left
6030        .fsp_keycombo_ctrl_alt_bg2.Left = .fsp_keycombo_plain_bg1.Left
6040        .fsp_keycombo_ctrl_shift_bg1.Left = .fsp_keycombo_plain_bg1.Left
6050        .fsp_keycombo_ctrl_shift_bg2.Left = .fsp_keycombo_plain_bg1.Left
6060        .fsp_keycombo_alt_shift_bg1.Left = .fsp_keycombo_plain_bg1.Left
6070        .fsp_keycombo_alt_shift_bg2.Left = .fsp_keycombo_plain_bg1.Left
6080        .fsp_keycombo_ctrl_alt_shift_bg1.Left = .fsp_keycombo_plain_bg1.Left
6090        .fsp_keycombo_ctrl_alt_shift_bg2.Left = .fsp_keycombo_plain_bg1.Left
6100        If blnSortHere = True Then
6110          .Sort_line.Left = .fsp_keycombo_lbl.Left
6120          .Sort_lbl.Left = ((.fsp_keycombo_lbl.Left + .fsp_keycombo_lbl.Width) - lngSortLbl_Width)
6130        End If
6140        blnSortHere = False
6150        If .Sort_line.Left = .keydowntype_type_lbl.Left Then blnSortHere = True
6160        .keydowntype_type.Left = lngKeyDownType_Left
            '.keydowntype_type_bg1.Left = (.keydowntype_type.Left - lngTpp)
            '.keydowntype_type_bg2.Left = .keydowntype_type_bg1.Left
6170        .keydowntype_type_lbl.Left = .keydowntype_type.Left
6180        .keydowntype_type_lbl_line.Left = .keydowntype_type_lbl.Left
6190        .keydowntype_plain_bg1.Left = (.keydowntype_type.Left - lngTpp)
6200        .keydowntype_plain_bg2.Left = .keydowntype_plain_bg1.Left
6210        .keydowntype_alt_bg1.Left = .keydowntype_plain_bg1.Left
6220        .keydowntype_alt_bg2.Left = .keydowntype_plain_bg1.Left
6230        .keydowntype_shift_bg1.Left = .keydowntype_plain_bg1.Left
6240        .keydowntype_shift_bg2.Left = .keydowntype_plain_bg1.Left
6250        .keydowntype_ctrl_bg1.Left = .keydowntype_plain_bg1.Left
6260        .keydowntype_ctrl_bg2.Left = .keydowntype_plain_bg1.Left
6270        .keydowntype_ctrl_alt_bg1.Left = .keydowntype_plain_bg1.Left
6280        .keydowntype_ctrl_alt_bg2.Left = .keydowntype_plain_bg1.Left
6290        .keydowntype_ctrl_shift_bg1.Left = .keydowntype_plain_bg1.Left
6300        .keydowntype_ctrl_shift_bg2.Left = .keydowntype_plain_bg1.Left
6310        .keydowntype_alt_shift_bg1.Left = .keydowntype_plain_bg1.Left
6320        .keydowntype_alt_shift_bg2.Left = .keydowntype_plain_bg1.Left
6330        .keydowntype_ctrl_alt_shift_bg1.Left = .keydowntype_plain_bg1.Left
6340        .keydowntype_ctrl_alt_shift_bg2.Left = .keydowntype_plain_bg1.Left
6350        If blnSortHere = True Then
6360          .Sort_line.Left = .keydowntype_type_lbl.Left
6370          .Sort_lbl.Left = ((.keydowntype_type_lbl.Left + .keydowntype_type_lbl.Width) - lngSortLbl_Width)
6380        End If
6390        blnSortHere = False
6400        If .Sort_line.Left = .frm_header_section_lbl.Left Then blnSortHere = True
6410        .frm_header_section.Left = lngFrmHeader_Left
6420        .frm_header_section_bg1.Left = (.frm_header_section.Left - lngTpp)
6430        .frm_header_section_bg2.Left = .frm_header_section_bg1.Left
6440        .frm_header_section_lbl.Left = .frm_header_section.Left
6450        .frm_header_section_lbl_line.Left = .frm_header_section_lbl.Left
6460        .frm_caption_item.Left = .frm_header_section.Left
6470        .frm_caption_item_bg1.Left = (.frm_caption_item.Left - lngTpp)
6480        .frm_caption_item_bg2.Left = .frm_caption_item_bg1.Left
6490        If blnSortHere = True Then
6500          .Sort_line.Left = .frm_header_section_lbl.Left
6510          .Sort_lbl.Left = ((.frm_header_section_lbl.Left + .frm_header_section_lbl.Width) - lngSortLbl_Width)
6520        End If
6530        blnSortHere = False
6540        .frm_name_section.Visible = True
6550        .frm_name_section_bg1.Visible = True
6560        .frm_name_section_bg2.Visible = True
6570        .frm_name_section_lbl.Visible = True
6580        .frm_name_section_lbl_line.Visible = True
6590        .frm_name_item.Visible = True
6600        .frm_name_item_bg1.Visible = True
6610        .frm_name_item_bg2.Visible = True
6620        .frm_name_item_04.Visible = True
6630        .frm_name_item_04_line.Visible = True
6640        .frm_name_item_04_bg1.Visible = True
6650        .frm_name_item_04_bg2.Visible = True
6660        .frm_name_item_08.Visible = True
6670        .frm_name_item_08_line.Visible = True
6680        .frm_name_item_08_bg1.Visible = True
6690        .frm_name_item_08_bg2.Visible = True
6700        .frm_name_item_12.Visible = True
6710        .frm_name_item_12_line.Visible = True
6720        .frm_name_item_12_bg1.Visible = True
6730        .frm_name_item_12_bg2.Visible = True
6740        .frm_name_item_16.Visible = True
6750        .frm_name_item_16_line.Visible = True
6760        .frm_name_item_16_bg1.Visible = True
6770        .frm_name_item_16_bg2.Visible = True
6780        .frm_name_item_20.Visible = True
6790        .frm_name_item_20_line.Visible = True
6800        .frm_name_item_20_bg1.Visible = True
6810        .frm_name_item_20_bg2.Visible = True
6820        .frm_name_item_24.Visible = True
6830        .frm_name_item_24_line.Visible = True
6840        .frm_name_item_24_bg1.Visible = True
6850        .frm_name_item_24_bg2.Visible = True
6860        .Width = lngForm_Width
6870      Case False
6880        lngDiff = (.frm_header_section.Left - .frm_name_section.Left)
6890        If .Sort_line.Left = .frm_name_section_lbl.Left Then blnSortHere = True
6900        .frm_name_section.Visible = False
6910        .frm_name_section_bg1.Visible = False
6920        .frm_name_section_bg2.Visible = False
6930        .frm_name_section_lbl.Visible = False
6940        .frm_name_section_lbl_line.Visible = False
6950        .frm_name_item.Visible = False
6960        .frm_name_item_bg1.Visible = False
6970        .frm_name_item_bg2.Visible = False
6980        .frm_name_item_04.Visible = False
6990        .frm_name_item_04_line.Visible = False
7000        .frm_name_item_04_bg1.Visible = False
7010        .frm_name_item_04_bg2.Visible = False
7020        .frm_name_item_08.Visible = False
7030        .frm_name_item_08_line.Visible = False
7040        .frm_name_item_08_bg1.Visible = False
7050        .frm_name_item_08_bg2.Visible = False
7060        .frm_name_item_12.Visible = False
7070        .frm_name_item_12_line.Visible = False
7080        .frm_name_item_12_bg1.Visible = False
7090        .frm_name_item_12_bg2.Visible = False
7100        .frm_name_item_16.Visible = False
7110        .frm_name_item_16_line.Visible = False
7120        .frm_name_item_16_bg1.Visible = False
7130        .frm_name_item_16_bg2.Visible = False
7140        .frm_name_item_20.Visible = False
7150        .frm_name_item_20_line.Visible = False
7160        .frm_name_item_20_bg1.Visible = False
7170        .frm_name_item_20_bg2.Visible = False
7180        .frm_name_item_24.Visible = False
7190        .frm_name_item_24_line.Visible = False
7200        .frm_name_item_24_bg1.Visible = False
7210        .frm_name_item_24_bg2.Visible = False
7220        If blnSortHere = True Then
7230          .Sort_line.Visible = False
7240          .Sort_lbl.Visible = False
7250          blnResort = True
7260        End If
7270        blnSortHere = False
7280        If .Sort_line.Left = .frm_header_section_lbl.Left Then blnSortHere = True
7290        .frm_header_section.Left = (lngFrmHeader_Left - lngDiff)
7300        .frm_header_section_bg1.Left = (.frm_header_section.Left - lngTpp)
7310        .frm_header_section_bg2.Left = .frm_header_section_bg1.Left
7320        .frm_header_section_lbl.Left = .frm_header_section.Left
7330        .frm_header_section_lbl_line.Left = .frm_header_section_lbl.Left
7340        .frm_caption_item.Left = .frm_header_section.Left
7350        .frm_caption_item_bg1.Left = (.frm_caption_item.Left - lngTpp)
7360        .frm_caption_item_bg2.Left = .frm_caption_item_bg1.Left
7370        If blnSortHere = True Then
7380          .Sort_line.Left = .frm_header_section_lbl.Left
7390          .Sort_lbl.Left = ((.frm_header_section_lbl.Left + .frm_header_section_lbl.Width) - lngSortLbl_Width)
7400        End If
7410        blnSortHere = False
7420        If .Sort_line.Left = .keydowntype_type_lbl.Left Then blnSortHere = True
7430        .keydowntype_type.Left = (lngKeyDownType_Left - lngDiff)
            '.keydowntype_type_bg1.Left = (.keydowntype_type.Left - lngTpp)
            '.keydowntype_type_bg2.Left = .keydowntype_type_bg1.Left
7440        .keydowntype_type_lbl.Left = .keydowntype_type.Left
7450        .keydowntype_type_lbl_line.Left = .keydowntype_type_lbl.Left
7460        .keydowntype_plain_bg1.Left = (.keydowntype_type.Left - lngTpp)
7470        .keydowntype_plain_bg2.Left = .keydowntype_plain_bg1.Left
7480        .keydowntype_alt_bg1.Left = .keydowntype_plain_bg1.Left
7490        .keydowntype_alt_bg2.Left = .keydowntype_plain_bg1.Left
7500        .keydowntype_shift_bg1.Left = .keydowntype_plain_bg1.Left
7510        .keydowntype_shift_bg2.Left = .keydowntype_plain_bg1.Left
7520        .keydowntype_ctrl_bg1.Left = .keydowntype_plain_bg1.Left
7530        .keydowntype_ctrl_bg2.Left = .keydowntype_plain_bg1.Left
7540        .keydowntype_ctrl_alt_bg1.Left = .keydowntype_plain_bg1.Left
7550        .keydowntype_ctrl_alt_bg2.Left = .keydowntype_plain_bg1.Left
7560        .keydowntype_ctrl_shift_bg1.Left = .keydowntype_plain_bg1.Left
7570        .keydowntype_ctrl_shift_bg2.Left = .keydowntype_plain_bg1.Left
7580        .keydowntype_alt_shift_bg1.Left = .keydowntype_plain_bg1.Left
7590        .keydowntype_alt_shift_bg2.Left = .keydowntype_plain_bg1.Left
7600        .keydowntype_ctrl_alt_shift_bg1.Left = .keydowntype_plain_bg1.Left
7610        .keydowntype_ctrl_alt_shift_bg2.Left = .keydowntype_plain_bg1.Left
7620        If blnSortHere = True Then
7630          .Sort_line.Left = .keydowntype_type_lbl.Left
7640          .Sort_lbl.Left = ((.keydowntype_type_lbl.Left + .keydowntype_type_lbl.Width) - lngSortLbl_Width)
7650        End If
7660        blnSortHere = False
7670        If .Sort_line.Left = .fsp_keycombo_lbl.Left Then blnSortHere = True
7680        .fsp_keycombo.Left = (lngKeyCombo_Left - lngDiff)
            '.fsp_keycombo_bg1.Left = (.fsp_keycombo.Left - lngTpp)
            '.fsp_keycombo_bg2.Left = .fsp_keycombo_bg1.Left
7690        .fsp_keycombo_lbl.Left = .fsp_keycombo.Left
7700        .fsp_keycombo_lbl_line.Left = .fsp_keycombo_lbl.Left
7710        .fsp_keycombo_plain_bg1.Left = (.fsp_keycombo.Left - lngTpp)
7720        .fsp_keycombo_plain_bg2.Left = .fsp_keycombo_plain_bg1.Left
7730        .fsp_keycombo_alt_bg1.Left = .fsp_keycombo_plain_bg1.Left
7740        .fsp_keycombo_alt_bg2.Left = .fsp_keycombo_plain_bg1.Left
7750        .fsp_keycombo_shift_bg1.Left = .fsp_keycombo_plain_bg1.Left
7760        .fsp_keycombo_shift_bg2.Left = .fsp_keycombo_plain_bg1.Left
7770        .fsp_keycombo_ctrl_bg1.Left = .fsp_keycombo_plain_bg1.Left
7780        .fsp_keycombo_ctrl_bg2.Left = .fsp_keycombo_plain_bg1.Left
7790        .fsp_keycombo_ctrl_alt_bg1.Left = .fsp_keycombo_plain_bg1.Left
7800        .fsp_keycombo_ctrl_alt_bg2.Left = .fsp_keycombo_plain_bg1.Left
7810        .fsp_keycombo_ctrl_shift_bg1.Left = .fsp_keycombo_plain_bg1.Left
7820        .fsp_keycombo_ctrl_shift_bg2.Left = .fsp_keycombo_plain_bg1.Left
7830        .fsp_keycombo_alt_shift_bg1.Left = .fsp_keycombo_plain_bg1.Left
7840        .fsp_keycombo_alt_shift_bg2.Left = .fsp_keycombo_plain_bg1.Left
7850        .fsp_keycombo_ctrl_alt_shift_bg1.Left = .fsp_keycombo_plain_bg1.Left
7860        .fsp_keycombo_ctrl_alt_shift_bg2.Left = .fsp_keycombo_plain_bg1.Left
7870        If blnSortHere = True Then
7880          .Sort_line.Left = .fsp_keycombo_lbl.Left
7890          .Sort_lbl.Left = ((.fsp_keycombo_lbl.Left + .fsp_keycombo_lbl.Width) - lngSortLbl_Width)
7900        End If
7910        blnSortHere = False
7920        If .Sort_line.Left = .fsp_keycode_lbl.Left Then blnSortHere = True
7930        .fsp_keycode.Left = (lngKeyCode_Left - lngDiff)
7940        .fsp_keycode_bg1.Left = (.fsp_keycode.Left - lngTpp)
7950        .fsp_keycode_bg2.Left = .fsp_keycode_bg1.Left
7960        .fsp_keycode_lbl.Left = .fsp_keycode.Left
7970        .fsp_keycode_lbl_line.Left = .fsp_keycode_lbl.Left
7980        If blnSortHere = True Then
7990          .Sort_line.Left = .fsp_keycode_lbl.Left
8000          .Sort_lbl.Left = ((.fsp_keycode_lbl.Left + .fsp_keycode_lbl.Width) - lngSortLbl_Width)
8010        End If
8020        blnSortHere = False
8030        If .Sort_line.Left = .fsp_description_lbl.Left Then blnSortHere = True
8040        .fsp_description.Left = (lngDescription_Left - lngDiff)
8050        .fsp_description_bg1.Left = (.fsp_description.Left - lngTpp)
8060        .fsp_description_bg2.Left = .fsp_description_bg1.Left
8070        .fsp_description_lbl.Left = .fsp_description.Left
8080        .fsp_description_lbl_line.Left = .fsp_description_lbl.Left
8090        If blnSortHere = True Then
8100          .Sort_line.Left = .fsp_description_lbl.Left
8110          .Sort_lbl.Left = ((.fsp_description_lbl.Left + .fsp_description_lbl.Width) - lngSortLbl_Width)
8120        End If
8130        blnSortHere = False
8140        .Width = (lngForm_Width - lngDiff)
8150      End Select
8160    End With

        'VERT LINE FOR INDENTS?

EXITP:
8170    Exit Sub

ERRH:
8180    Select Case ERR.Number
        Case Else
8190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8200    End Select
8210    Resume EXITP

End Sub

Public Function RecCnt() As Long

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

8310    lngRetVal = 0&

8320    With Me
8330      Set rst = .RecordsetClone
8340      With rst
8350        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
8360        Else
8370          .MoveLast
8380          lngRetVal = .RecordCount
8390        End If
8400        .Close
8410      End With
8420    End With

EXITP:
8430    Set rst = Nothing
8440    RecCnt = lngRetVal
8450    Exit Function

ERRH:
8460    lngRetVal = 0&
8470    Select Case ERR.Number
        Case Else
8480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8490    End Select
8500    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant, Optional varKey As Variant)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

8610    Select Case intWhere
        Case acCmdRecordsGoToFirst
8620      DoCmd.RunCommand acCmdRecordsGoToFirst
8630    Case acCmdRecordsGoToLast
8640      DoCmd.RunCommand acCmdRecordsGoToLast
8650    Case acCmdRecordsGoToPrevious
8660      DoCmd.RunCommand acCmdRecordsGoToPrevious
8670    Case acCmdRecordsGoToNext
8680      DoCmd.RunCommand acCmdRecordsGoToNext
8690    Case Else
8700      If IsMissing(varID) = False Then
8710        With Me
8720          Set rst = .RecordsetClone
8730          With rst
8740            Select Case varKey
                Case "fsf_id"
8750              .FindFirst "[fsf_id] = " & CStr(varID)
8760              If .NoMatch = False Then
8770                Me.Bookmark = .Bookmark
8780              End If
8790            Case "fsfd_section"
8800              .FindFirst "[fsfd_section] = " & CStr(varID) & " And [fsfd_order] = 0"
8810              If .NoMatch = False Then
8820                Me.Bookmark = .Bookmark
8830              End If
8840            End Select
8850            .Close
8860          End With
8870        End With
8880      End If
8890    End Select

EXITP:
8900    Set rst = Nothing
8910    Exit Sub

ERRH:
8920    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8930    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8940    Case Else
8950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8960    End Select
8970    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        'Const strStdAsc As String = ", [fsfd_section], [fsfd_order], [fsfd_level], [keydowntype_order], [fsp_keycode]"
        ' ** [fsfd_section], [fsfd_order], [fsfd_level], [keydowntype_order], [fsp_keycode]

9010    With Me
9020      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
9030        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
9040      End If
9050      .Controls(strSortLbl).Visible = False
9060      .Controls(strSortLine).Visible = False
9070      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the the right side of the form.
9080      If strProc = "Form_Load" Then
9090        strCalled = "frm_name_section"
9100        strSortNow = strSortOrig
9110        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9120        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
9130        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9140        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
9150        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
9160        .Controls(strSortLbl).Top = lngSortLbl_Top
9170        .Controls(strSortLbl).Left = lngSortLbl_Left
9180        .Controls(strSortLine).Top = lngSortLine_Top
9190        .Controls(strSortLine).Left = lngSortLine_Left
9200        .Controls(strSortLine).Width = lngSortLine_Width
9210        .Controls(strSortLbl).Caption = strArwUp
9220        .Controls(strSortLbl).ForeColor = CLR_DKBLU
9230      Else
9240        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
9250        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9260        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
9270        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9280        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
9290        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
9300        .Controls(strSortLbl).Top = lngSortLbl_Top
9310        .Controls(strSortLbl).Left = lngSortLbl_Left
9320        .Controls(strSortLine).Top = lngSortLine_Top
9330        .Controls(strSortLine).Left = lngSortLine_Left
9340        .Controls(strSortLine).Width = lngSortLine_Width
9350        If strCalled = "frm_name_section" Then
              ' ** frm_name_section, frm_name_item
9360          strSortAsc = strSortOrig
9370          If strSortNow = strSortAsc Then
9380            strSortNow = "[fsfd_section] DESC, [fsfd_order] DESC, [fsfd_level] DESC, [keydowntype_order] DESC, [fsp_keycode] DESC"
9390            .Controls(strSortLbl).Caption = strArwDn
9400            .Controls(strSortLbl).ForeColor = CLR_DKRED
9410          Else
9420            strSortNow = strSortAsc
9430            .Controls(strSortLbl).Caption = strArwUp
9440            .Controls(strSortLbl).ForeColor = CLR_DKBLU
9450          End If
9460        Else
9470          .Controls(strSortLbl).Caption = strArwUp
9480          .Controls(strSortLbl).ForeColor = CLR_DKBLU
9490          Select Case strCalled
              Case "frm_header_section"
                ' ** frm_header_section, frm_caption_item
9500            strSortAsc = "[frm_caption_sort], [fsfd_section]"
9510          Case "keydowntype_type"
                ' ** keydowntype_type
9520            strSortAsc = "[keydowntype_order], [fsp_keycode], [fsfd_section], [fsfd_level]"
9530          Case "fsp_keycombo"
                ' ** fsp_keycombo
9540            strSortAsc = "[fsp_keycombo], [fsp_keycode], [fsfd_section], [fsfd_level]"
9550            lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
9560            .Controls(strSortLbl).Left = lngSortLbl_Left
9570          Case "fsp_keycode"
                ' ** fsp_keycode
9580            strSortAsc = "[fsp_keycode], [keydowntype_order], [fsfd_section], [fsfd_level]"
9590          Case "fsp_description"
                ' ** fsp_description
9600            strSortAsc = "[fsp_description], [keydowntype_order], [fsp_keycode], [fsfd_section], [fsfd_level]"
9610          End Select
9620          If strSortNow = strSortAsc Then
9630            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
9640            Select Case intCnt
                Case 1
9650              strTmp01 = strSortAsc & " DESC"
9660            Case 2
9670              intPos01 = InStr(strSortAsc, ",")
9680              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9690              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
9700            Case 3
9710              intPos01 = InStr(strSortAsc, ",")
9720              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9730              strTmp02 = Mid(strSortAsc, intPos01)
9740              intPos01 = InStr(2, strTmp02, ",")
9750              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
9760              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
9770            Case 4
9780              intPos01 = InStr(strSortAsc, ",")
9790              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9800              strTmp02 = Mid(strSortAsc, intPos01)
9810              intPos01 = InStr(2, strTmp02, ",")
9820              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
9830              strTmp02 = Mid(strTmp02, intPos01)
9840              intPos01 = InStr(2, strTmp02, ",")
9850              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
9860              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
9870            Case 5
9880              intPos01 = InStr(strSortAsc, ",")
9890              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9900              strTmp02 = Mid(strSortAsc, intPos01)
9910              intPos01 = InStr(2, strTmp02, ",")
9920              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
9930              strTmp02 = Mid(strTmp02, intPos01)
9940              intPos01 = InStr(2, strTmp02, ",")
9950              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
9960              strTmp02 = Mid(strTmp02, intPos01)
9970              intPos01 = InStr(2, strTmp02, ",")
9980              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
9990              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
10000           Case 6
10010             intPos01 = InStr(strSortAsc, ",")
10020             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
10030             strTmp02 = Mid(strSortAsc, intPos01)
10040             intPos01 = InStr(2, strTmp02, ",")
10050             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10060             strTmp02 = Mid(strTmp02, intPos01)
10070             intPos01 = InStr(2, strTmp02, ",")
10080             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10090             strTmp02 = Mid(strTmp02, intPos01)
10100             intPos01 = InStr(2, strTmp02, ",")
10110             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10120             strTmp02 = Mid(strTmp02, intPos01)
10130             intPos01 = InStr(2, strTmp02, ",")
10140             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10150             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
10160           Case 7
10170             intPos01 = InStr(strSortAsc, ",")
10180             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
10190             strTmp02 = Mid(strSortAsc, intPos01)
10200             intPos01 = InStr(2, strTmp02, ",")
10210             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10220             strTmp02 = Mid(strTmp02, intPos01)
10230             intPos01 = InStr(2, strTmp02, ",")
10240             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10250             strTmp02 = Mid(strTmp02, intPos01)
10260             intPos01 = InStr(2, strTmp02, ",")
10270             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10280             strTmp02 = Mid(strTmp02, intPos01)
10290             intPos01 = InStr(2, strTmp02, ",")
10300             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10310             strTmp02 = Mid(strTmp02, intPos01)
10320             intPos01 = InStr(2, strTmp02, ",")
10330             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10340             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
10350           End Select
10360           strSortNow = strTmp01
10370           .Controls(strSortLbl).Caption = strArwDn
10380           .Controls(strSortLbl).ForeColor = CLR_DKRED
10390         Else
10400           strSortNow = strSortAsc
10410         End If
10420       End If
10430     End If
10440     .Controls(strSortLbl).Visible = True
10450     .Controls(strSortLine).Visible = True
10460     .OrderBy = strSortNow
10470     .OrderByOn = True
10480   End With

        ' ** Potential Sorts:
        ' **   [fsfd_section], [fsfd_order], [fsfd_level], [keydowntype_order], [fsp_keycode]
        ' **   [frm_caption_sort], [fsfd_section]
        ' **   [keydowntype_order], [fsp_keycode], [fsfd_section], [fsfd_level]
        ' **   [fsp_keycombo], [fsp_keycode], [fsfd_section], [fsfd_level]
        ' **   [fsp_keycode], [keydowntype_order], [fsfd_section], [fsfd_level]
        ' **   [fsp_description], [keydowntype_order], [fsp_keycode], [fsfd_section], [fsfd_level]

EXITP:
10490   Exit Sub

ERRH:
10500   Select Case ERR.Number
        Case Else
10510     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10520   End Select
10530   Resume EXITP

End Sub

Public Function SortNow_Get() As String

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

10610   strRetVal = strSortNow

EXITP:
10620   SortNow_Get = strRetVal
10630   Exit Function

ERRH:
10640   strRetVal = vbNullString
10650   Select Case ERR.Number
        Case Else
10660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10670   End Select
10680   Resume EXITP

End Function
