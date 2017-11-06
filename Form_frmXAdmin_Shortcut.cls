VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmXAdmin_Shortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmXAdmin_Shortcut"

'VGC 04/21/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   ShowFormName()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Show All:         A {opgSearchMode_optAll}
' **   Combo Keys:       B {fsp_keycombo on frmXAdmin_Shortcut_Sub}
' **   Close:            C {cmdClose}
' **   Description:      D {fsp_description on frmXAdmin_Shortcut_Sub}
' **   Forms:            F {lbxForms}
' **   Key:              K {fsp_keycode on frmXAdmin_Shortcut_Sub}
' **   Menus:            M {cmbSections}
' **   Show Form Name:   N {chkShowFormName}
' **   Filter Only:      O {opgSearchMode_optSpecified}
' **   KeyDown Type:     T {keydowntype_type on frmXAdmin_Shortcut_Sub}
' **   Include Misc:     U {chkIncludeMiscForms}
' **   Exit:             X {cmdClose}

'zzz_qry_xForm_Shortcut_33_02 -> qryForm_Shortcut_03_02
'zzz_qry_xForm_Shortcut_33_06 -> qryForm_Shortcut_03_06
'zzz_qry_xForm_Shortcut_34_01 -> qryForm_Shortcut_04_01
'zzz_qry_xForm_Shortcut_34_02 -> qryForm_Shortcut_04_02
'zzz_qry_xForm_Shortcut_34_03 -> qryForm_Shortcut_04_03
'zzz_qry_xForm_Shortcut_36_02 -> qryForm_Shortcut_06_02

' ** cmbSections combo box constants:
Private Const CBX_S_SEC   As Integer = 0  'fsfd_section
'Private Const CBX_S_CAP   As Integer = 1  'frm_header_section
'Private Const CBX_S_NAM   As Integer = 2  'frm_name_section
'Private Const CBX_S_DID   As Integer = 3  'dbs_id
'Private Const CBX_S_FID   As Integer = 4  'frm_id
'Private Const CBX_S_FSFID As Integer = 5  'fsp_order

' ** Array: arr_varCtl().
Private lngCtls As Long, arr_varCtl() As Variant
Private Const C_ELEMS As Integer = 1  ' ** Array's first-element UBound().
Private Const C_CNAM As Integer = 0
Private Const C_LFT  As Integer = 1

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngSub_Width As Long, lngClose_Left As Long

Private lngLbxForms_Width As Long, lngLbxFormsBox_Width As Long, strLbxCol_Widths As String
Private lngOpgSearchMode_Left As Long, lngOptLbl_Offset As Long, lngBox2_Offset As Long, lngOpt_Offset As Long
Private strCmbCol_Widths As String, strList_Width As String, lngFrmName_Width As Long, lngOrder_Left As Long

Private blnRecover As Boolean, blnHourglass As Boolean
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           blnRecover = False: blnHourglass = False

170           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190           lngForm_Width = .Width
200           lngLbxForms_Width = .lbxForms.Width
210           lngLbxFormsBox_Width = .lbxForms_box.Width
220           lngOrder_Left = .fsp_order_lbl.Left
230           lngOpgSearchMode_Left = .opgSearchMode.Left
240           lngOptLbl_Offset = (.opgSearchMode_optAll_lbl.Left - .opgSearchMode_optAll.Left)
250           lngBox2_Offset = (.opgSearchMode_box2.Left - .opgSearchMode_box.Left)
260           lngOpt_Offset = (.opgSearchMode_optAll.Left - .opgSearchMode.Left)
270           strLbxCol_Widths = .lbxForms.ColumnWidths
280           strCmbCol_Widths = .cmbSections.ColumnWidths
290           strList_Width = .cmbSections.ListWidth
300           lngClose_Left = .cmdClose.Left
310           lngSub_Width = .frmXAdmin_Shortcut_Sub.Width
320           lngFrmName_Width = (.frmXAdmin_Shortcut_Sub.Form.frm_header_section.Left - .frmXAdmin_Shortcut_Sub.Form.frm_name_section.Left)

330           LoadKeyArray  ' ** Procedure: Below.

              ' ** Variables are fed empty, then populated ByRef.
340           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

350           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

360           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
370           If lngMonitorCnt > 1& Then
380             EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
390             lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
400             If lngMonitorNum > 1& Then
410               LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
420             End If
430           End If

440           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
450           DoEvents

460           chkIncludeMiscForms_AfterUpdate  ' ** Procedure: Below.
470           opgSearchMode_AfterUpdate  ' ** Procedure: Below.
480           chkShowFormName_AfterUpdate  ' ** Procedure: Below.
490           cmbSections_AfterUpdate  ' ** Procedure: Below.

500           SetFrmNumbering  ' ** Procedure: Below.

510           blnIsOpen = False

520         Else
530           Cancel = -1
540         End If
550       Else
560         Cancel = -1
570       End If

580     End With

590     DoCmd.Hourglass False

600     If Cancel = -1 Then
610       Beep
620       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
630     End If

EXITP:
640     Exit Sub

ERRH:
650     DoCmd.Hourglass False
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionalities:
        ' **   Combo Keys:       B {fsp_keycombo on frmXAdmin_Shortcut_Sub}
        ' **   Description:      D {fsp_description on frmXAdmin_Shortcut_Sub}
        ' **   Key:              K {fsp_keycode on frmXAdmin_Shortcut_Sub}
        ' **   KeyDown Type:     T {keydowntype_type on frmXAdmin_Shortcut_Sub}
        ' **   Exit:             X {cmdClose}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyF5
770         intRetVal = 0
780         SetFrmNumbering  ' ** Procedure: Below.
790       Case vbKeyEscape
800         intRetVal = 0
810         cmdClose_Click  ' ** Procedure: Below.
820       End Select
830     End If

        ' ** Alt keys.
840     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
850       Select Case intRetVal
          Case vbKeyB
860         With Me
870           intRetVal = 0
880           .frmXAdmin_Shortcut_Sub.SetFocus
890           .frmXAdmin_Shortcut_Sub.Form.fsp_keycombo.SetFocus
900         End With
910       Case vbKeyD
920         With Me
930           intRetVal = 0
940           .frmXAdmin_Shortcut_Sub.SetFocus
950           .frmXAdmin_Shortcut_Sub.Form.fsp_description.SetFocus
960         End With
970       Case vbKeyK
980         With Me
990           intRetVal = 0
1000          .frmXAdmin_Shortcut_Sub.SetFocus
1010          .frmXAdmin_Shortcut_Sub.Form.fsp_keycode.SetFocus
1020        End With
1030      Case vbKeyT
1040        With Me
1050          intRetVal = 0
1060          .frmXAdmin_Shortcut_Sub.SetFocus
1070          .frmXAdmin_Shortcut_Sub.Form.keydowntype_type.SetFocus
1080        End With
1090      Case vbKeyX
1100        intRetVal = 0
1110        cmdClose_Click  ' ** Procedure: Below.
1120      End Select
1130    End If

        ' ** Ctrl-Shift keys.
1140    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1150      Select Case intRetVal
          Case vbKeyF
1160        With Me
1170          intRetVal = 0
1180          .FocusHolder.SetFocus
1190        End With
1200      End Select
1210    End If

EXITP:
1220    KeyCode = intRetVal
1230    Exit Sub

ERRH:
1240    intRetVal = 0
1250    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1260    Case Else
1270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1280    End Select
1290    Resume EXITP

End Sub

Private Sub Form_Timer()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngFsfID As Long, lngFsfdSec As Long
        Dim strFormName As String
        Static lngTmp01 As Long

1310    With Me
1320      .TimerInterval = 0&
1330      If gblnSetFocus = True Then
1340        gblnSetFocus = False
1350        strFormName = vbNullString
1360  On Error Resume Next
1370        Do Until strFormName = THIS_NAME
1380          DoCmd.SelectObject acForm, THIS_NAME, False
1390          strFormName = Screen.ActiveForm.Name
1400        Loop
1410  On Error GoTo ERRH
1420        DoCmd.Hourglass False
1430      End If
1440      If blnRecover = True Then
1450        blnRecover = False
1460        lngFsfdSec = .fsfd_section
1470        lngFsfID = .fsf_id
1480        If lngFsfdSec > 0& Then
1490          .cmbSections = lngFsfdSec
1500          .frmXAdmin_Shortcut_Sub.Form.MoveRec 0, lngFsfdSec, "fsfd_section"  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
1510        End If
1520        If lngFsfID > 0& Then
1530          .lbxForms = lngFsfID
1540          .frmXAdmin_Shortcut_Sub.Form.MoveRec 0, lngFsfID, "fsf_id"  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
1550        End If
1560        DoCmd.Hourglass False
1570      End If
          ' ** This is all because I just couldn't get the damn hourglass to stop.
1580      If lngTmp01 = 0 Then
1590        lngTmp01 = lngTmp01 + 1&
1600        blnHourglass = True
1610        .TimerInterval = 200&
1620      Else
1630        If blnHourglass = True Then
1640          If lngTmp01 < 6& Then
1650            lngTmp01 = lngTmp01 + 1&
1660            DoCmd.Hourglass False
1670            .TimerInterval = 200&
1680          Else
1690            DoCmd.Hourglass False
1700            blnHourglass = False
1710          End If
1720        End If
1730      End If
1740    End With

EXITP:
1750    Exit Sub

ERRH:
1760    DoCmd.Hourglass False
1770    Select Case ERR.Number
        Case Else
1780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1790    End Select
1800    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String

1910    With Me

          'DoCmd.Hourglass True
1920      DoEvents

1930      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1940      gstrReturningForm = THIS_NAME
1950      gblnSetFocus = True
1960      strDocName = strCallingForm
1970      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Utility"

1980    End With

EXITP:
1990    Exit Sub

ERRH:
2000    DoCmd.Hourglass False
2010    Select Case ERR.Number
        Case Else
2020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2030    End Select
2040    Resume EXITP

End Sub

Public Sub cmdClose_Click()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        'DoCmd.Hourglass True
2110    DoEvents

2120    DoCmd.Close acForm, THIS_NAME

EXITP:
2130    Exit Sub

ERRH:
2140    DoCmd.Hourglass False
2150    Select Case ERR.Number
        Case Else
2160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2170    End Select
2180    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2220    intShiftDown = (Shift And acShiftMask) > 0
2230    intAltDown = (Shift And acAltMask) > 0
2240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2260      Select Case intRetVal
          Case vbKeyTab
2270        With Me
2280          intRetVal = 0
2290          .cmbSections.SetFocus
2300        End With
2310      End Select
2320    End If

        ' ** Shift keys.
2330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2340      Select Case intRetVal
          Case vbKeyTab
2350        With Me
2360          intRetVal = 0
2370          lngRecsCur = .frmXAdmin_Shortcut_Sub.Form.RecCnt  ' ** Form Function: frmXAdmin_Shortcut_Sub.
2380          If lngRecsCur > 0& Then
2390            .frmXAdmin_Shortcut_Sub.SetFocus
2400          Else
2410            .chkIncludeMiscForms.SetFocus
2420          End If
2430        End With
2440      End Select
2450    End If

EXITP:
2460    KeyCode = intRetVal
2470    Exit Sub

ERRH:
2480    intRetVal = 0
2490    Select Case ERR.Number
        Case Else
2500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2510    End Select
2520    Resume EXITP

End Sub

Private Sub cmbSections_AfterUpdate()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbSections_AfterUpdate"

        Dim lngFsfdSec As Long
        Dim strFilter As String

2610    With Me
2620      If IsNull(.cmbSections) = False Then
2630        lngFsfdSec = .cmbSections.Column(CBX_S_SEC)
2640        Select Case .opgSearchMode
            Case .opgSearchMode_optAll.OptionValue
2650          If .frmXAdmin_Shortcut_Sub.Form.Filter <> vbNullString Then
2660            Select Case .chkIncludeMiscForms
                Case True
2670              .frmXAdmin_Shortcut_Sub.Form.Filter = vbNullString
2680              .frmXAdmin_Shortcut_Sub.Form.FilterOn = False
2690            Case False
2700              strFilter = "[fsf_show2]=False"
2710              .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
2720              .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
2730            End Select
2740          End If
2750          .frmXAdmin_Shortcut_Sub.Form.MoveRec 0, lngFsfdSec, "fsfd_section"  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
2760        Case .opgSearchMode_optSpecified.OptionValue
2770          strFilter = "[fsfd_section] = " & CStr(lngFsfdSec)
2780          If .frmXAdmin_Shortcut_Sub.Form.Filter <> vbNullString Then
2790            .frmXAdmin_Shortcut_Sub.Form.Filter = vbNullString
2800            .frmXAdmin_Shortcut_Sub.Form.FilterOn = False
2810          End If
2820          .frmXAdmin_Shortcut_Sub.Form.AfterUpdate_Set True  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
2830          If .chkIncludeMiscForms = False Then
2840            strFilter = strFilter & " And [fsf_show2]=False"
2850          End If
2860          .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
2870          .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
2880          .frmXAdmin_Shortcut_Sub.Form.MoveRec 0, lngFsfdSec, "fsfd_section"  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
2890          DoEvents
2900        End Select
2910        DoEvents
2920        .cmbSections = lngFsfdSec
2930        .lbxForms = Null
2940        .frmXAdmin_Shortcut_Sub.SetFocus
2950      End If
2960    End With

EXITP:
2970    Exit Sub

ERRH:
2980    Select Case ERR.Number
        Case Else
2990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3000    End Select
3010    Resume EXITP

End Sub

Private Sub cmbSections_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbSections_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3120    intShiftDown = (Shift And acShiftMask) > 0
3130    intAltDown = (Shift And acAltMask) > 0
3140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3160      Select Case intRetVal
          Case vbKeyTab
3170        With Me
3180          intRetVal = 0
3190          .lbxForms.SetFocus
3200        End With
3210      End Select
3220    End If

        ' ** Shift keys.
3230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3240      Select Case intRetVal
          Case vbKeyTab
3250        With Me
3260          intRetVal = 0
3270          .cmdClose.SetFocus
3280        End With
3290      End Select
3300    End If

EXITP:
3310    KeyCode = intRetVal
3320    Exit Sub

ERRH:
3330    intRetVal = 0
3340    Select Case ERR.Number
        Case Else
3350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3360    End Select
3370    Resume EXITP

End Sub

Private Sub lbxForms_AfterUpdate()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxForms_AfterUpdate"

        Dim strFilter As String
        Dim lngFsfID As Long, lngFsfdSec As Long

3410    With Me
3420      If IsNull(.lbxForms) = False Then
3430        lngFsfID = .lbxForms
3440        lngFsfdSec = .lbxForms.Column(5)
3450        Select Case .opgSearchMode
            Case .opgSearchMode_optAll.OptionValue
3460          .frmXAdmin_Shortcut_Sub.Form.MoveRec 0, lngFsfID, "fsf_id"  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
3470        Case .opgSearchMode_optSpecified.OptionValue
3480          strFilter = .frmXAdmin_Shortcut_Sub.Form.Filter
3490          If strFilter = vbNullString Then
3500            strFilter = "[fsfd_section] = " & CStr(lngFsfdSec)
3510            .frmXAdmin_Shortcut_Sub.Form.AfterUpdate_Set True  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
3520            If .chkIncludeMiscForms = False Then
3530              strFilter = strFilter & " And [fsf_show2]=False"
3540            End If
3550            .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
3560            .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
3570            DoEvents
3580            .frmXAdmin_Shortcut_Sub.Form.MoveRec 0, lngFsfID, "fsf_id"  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
3590            DoEvents
3600          Else
3610            If .frmXAdmin_Shortcut_Sub.Form.fsfd_section = lngFsfdSec Then
3620              .frmXAdmin_Shortcut_Sub.Form.MoveRec 0, lngFsfID, "fsf_id"  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
3630            Else
3640              .frmXAdmin_Shortcut_Sub.Form.Filter = vbNullString
3650              .frmXAdmin_Shortcut_Sub.Form.FilterOn = False
3660              strFilter = "[fsfd_section] = " & CStr(lngFsfdSec)
3670              If .chkIncludeMiscForms = False Then
3680                strFilter = strFilter & " And [fsf_show2]=False"
3690              End If
3700              .frmXAdmin_Shortcut_Sub.Form.AfterUpdate_Set True  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
3710              .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
3720              .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
3730              DoEvents
3740              .frmXAdmin_Shortcut_Sub.Form.MoveRec 0, lngFsfID, "fsf_id"  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
3750              DoEvents
3760            End If
3770          End If
3780        End Select
3790        DoEvents
3800        .lbxForms = lngFsfID
3810        .frmXAdmin_Shortcut_Sub.SetFocus
3820      End If
3830    End With

EXITP:
3840    Exit Sub

ERRH:
3850    Select Case ERR.Number
        Case Else
3860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3870    End Select
3880    Resume EXITP

End Sub

Private Sub lbxForms_DblClick(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxForms_DblClick"

3910    lbxForms_AfterUpdate  ' ** Procedure: Above.

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub lbxForms_KeyDown(KeyCode As Integer, Shift As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxForms_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4020    intShiftDown = (Shift And acShiftMask) > 0
4030    intAltDown = (Shift And acAltMask) > 0
4040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4070        With Me
4080          intRetVal = 0
4090          .opgSearchMode.SetFocus
4100        End With
4110      End Select
4120    End If

        ' ** Shift keys.
4130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4140      Select Case intRetVal
          Case vbKeyTab
4150        With Me
4160          intRetVal = 0
4170          .cmbSections.SetFocus
4180        End With
4190      End Select
4200    End If

EXITP:
4210    KeyCode = intRetVal
4220    Exit Sub

ERRH:
4230    intRetVal = 0
4240    Select Case ERR.Number
        Case Else
4250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4260    End Select
4270    Resume EXITP

End Sub

Public Sub opgSearchMode_AfterUpdate()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgSearchMode_AfterUpdate"

        Dim lngFsfID As Long, lngFsfdSec As Long
        Dim strFilter As String

4310    With Me
4320      lngFsfID = .frmXAdmin_Shortcut_Sub.Form.fsf_id
4330      lngFsfdSec = .frmXAdmin_Shortcut_Sub.Form.fsfd_section
4340      .fsfd_section = 0
4350      .fsf_id = 0
4360      If IsNull(.cmbSections) = False Then
4370        .fsfd_section = lngFsfdSec
4380      End If
4390      If IsNull(.lbxForms) = False Then
4400        If .lbxForms > 0 Then
4410          .fsf_id = lngFsfID
4420        End If
4430      End If
4440      Select Case .opgSearchMode
          Case .opgSearchMode_optAll.OptionValue
4450        .opgSearchMode_optAll_lbl.FontBold = True
4460        .opgSearchMode_optSpecified_lbl.FontBold = False
4470        If .frmXAdmin_Shortcut_Sub.Form.Filter <> vbNullString Then
4480          Select Case .chkIncludeMiscForms
              Case True
4490            .frmXAdmin_Shortcut_Sub.Form.Filter = vbNullString
4500            .frmXAdmin_Shortcut_Sub.Form.FilterOn = False
4510          Case False
4520            strFilter = "[fsf_show2]=False"
4530            .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
4540            .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
4550          End Select
4560        End If
4570        .frmXAdmin_Shortcut_Sub.Form.MoveRec 0, lngFsfdSec, "fsfd_section"  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
4580        .frmXAdmin_Shortcut_Sub.Form.MoveRec 0, lngFsfID, "fsf_id"  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
4590      Case .opgSearchMode_optSpecified.OptionValue
4600        .opgSearchMode_optAll_lbl.FontBold = False
4610        .opgSearchMode_optSpecified_lbl.FontBold = True
4620        If .frmXAdmin_Shortcut_Sub.Form.Filter <> vbNullString Then
4630          .frmXAdmin_Shortcut_Sub.Form.Filter = vbNullString
4640          .frmXAdmin_Shortcut_Sub.Form.FilterOn = False
4650        End If
4660        strFilter = "[fsfd_section] = " & CStr(lngFsfdSec)
4670        If .chkIncludeMiscForms = False Then
4680          strFilter = strFilter & " And [fsf_show2]=False"
4690        End If
4700        .frmXAdmin_Shortcut_Sub.Form.AfterUpdate_Set True  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
4710        .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
4720        .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
4730        .frmXAdmin_Shortcut_Sub.Form.MoveRec 0, lngFsfID, "fsf_id"  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
4740        DoEvents
4750      End Select
4760      SetFrmNumbering  ' ** Procedure: Below.
4770      blnRecover = True
4780      .TimerInterval = 500&
4790    End With

EXITP:
4800    Exit Sub

ERRH:
4810    Select Case ERR.Number
        Case Else
4820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4830    End Select
4840    Resume EXITP

End Sub

Private Sub opgSearchMode_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgSearchMode_optAll_KeyDown"

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
4990          .chkShowFormName.SetFocus
5000        End With
5010      End Select
5020    End If

        ' ** Shift keys.
5030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5050        With Me
5060          intRetVal = 0
5070          .lbxForms.SetFocus
5080        End With
5090      End Select
5100    End If

EXITP:
5110    KeyCode = intRetVal
5120    Exit Sub

ERRH:
5130    intRetVal = 0
5140    Select Case ERR.Number
        Case Else
5150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5160    End Select
5170    Resume EXITP

End Sub

Private Sub opgSearchMode_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgSearchMode_optSpecified_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5220    intShiftDown = (Shift And acShiftMask) > 0
5230    intAltDown = (Shift And acAltMask) > 0
5240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5270        With Me
5280          intRetVal = 0
5290          .chkShowFormName.SetFocus
5300        End With
5310      End Select
5320    End If

        ' ** Shift keys.
5330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5350        With Me
5360          intRetVal = 0
5370          .lbxForms.SetFocus
5380        End With
5390      End Select
5400    End If

EXITP:
5410    KeyCode = intRetVal
5420    Exit Sub

ERRH:
5430    intRetVal = 0
5440    Select Case ERR.Number
        Case Else
5450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5460    End Select
5470    Resume EXITP

End Sub

Private Sub chkShowFormName_AfterUpdate()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowFormName_AfterUpdate"

        Dim lngDiff As Long
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String, lngTmp03 As Long

5510    With Me
5520      Select Case .chkShowFormName
          Case True
5530        .chkShowFormName_lbl.FontBold = True
5540        .frmXAdmin_Shortcut_Sub.Form.ShowFormName_Sub True  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
5550        DoEvents
5560        ShowFormName True  ' ** Procedure: Below.
5570        DoEvents
5580        .cmbSections.ColumnWidths = strCmbCol_Widths
5590        lngTmp03 = Val(strList_Width)
5600        .cmbSections.ListWidth = lngTmp03  ' ** Property is actually a string.
5610      Case False
5620        .chkShowFormName_lbl.FontBold = False
5630        .frmXAdmin_Shortcut_Sub.Form.ShowFormName_Sub False  ' ** Form Procedure: frmXAdmin_Shortcut_Sub.
5640        DoEvents
5650        ShowFormName False  ' ** Procedure: Below.
5660        DoEvents
            ' ** .cmbSections.ColumnWidths = "0;2685;315;2505;0;0;0"
5670        intPos01 = CharPos(strCmbCol_Widths, 4, ";")  ' ** Module Function: modStringFuncs.
5680        strTmp01 = Left(strCmbCol_Widths, (intPos01 - 1))
5690        intPos01 = CharPos(strTmp01, 3, ";")  ' ** Module Function: modStringFuncs.
5700        strTmp01 = Mid(strTmp01, (intPos01 + 1))  ' ** Width of form name column, in Twips.
5710        lngDiff = Val(strTmp01)
5720        strTmp02 = StringReplace(strCmbCol_Widths, strTmp01, "0")  ' ** Depends on it being unique.
5730        .cmbSections.ColumnWidths = strTmp02
5740        lngTmp03 = (Val(strList_Width) - Val(strTmp01))
5750        .cmbSections.ListWidth = lngTmp03  ' ** Property is actually a string.
5760      End Select
5770      .frmXAdmin_Shortcut_Sub.SetFocus
5780    End With

EXITP:
5790    Exit Sub

ERRH:
5800    Select Case ERR.Number
        Case Else
5810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5820    End Select
5830    Resume EXITP

End Sub

Private Sub chkShowFormName_KeyDown(KeyCode As Integer, Shift As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowFormName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5920    intShiftDown = (Shift And acShiftMask) > 0
5930    intAltDown = (Shift And acAltMask) > 0
5940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5970        With Me
5980          intRetVal = 0
5990          .chkIncludeMiscForms.SetFocus
6000        End With
6010      End Select
6020    End If

        ' ** Shift keys.
6030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6050        With Me
6060          intRetVal = 0
6070          .opgSearchMode.SetFocus
6080        End With
6090      End Select
6100    End If

EXITP:
6110    KeyCode = intRetVal
6120    Exit Sub

ERRH:
6130    intRetVal = 0
6140    Select Case ERR.Number
        Case Else
6150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6160    End Select
6170    Resume EXITP

End Sub

Private Sub chkIncludeMiscForms_AfterUpdate()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeMiscForms_AfterUpdate"

        Dim lngFsfID As Long, lngFsfdSec As Long
        Dim strFilter As String
        Dim intPos01 As Integer

6210    With Me
          ' ** Why does filtering have to get so big and messy!
6220      lngFsfID = .frmXAdmin_Shortcut_Sub.Form.fsf_id
6230      lngFsfdSec = .frmXAdmin_Shortcut_Sub.Form.fsfd_section
6240      .fsfd_section = 0
6250      .fsf_id = 0
6260      If IsNull(.cmbSections) = False Then
6270        .fsfd_section = lngFsfdSec
6280      End If
6290      If IsNull(.lbxForms) = False Then
6300        If .lbxForms > 0 Then
6310          .fsf_id = lngFsfID
6320        End If
6330      End If
6340      Select Case .chkIncludeMiscForms
          Case True
6350        .chkIncludeMiscForms_lbl.FontBold = True
6360        If .lbxForms.RowSource <> "qryForm_Shortcut_04_02" Then
              ' ** qryForm_Shortcut_03_02 (tblVBComponent_Shortcut_New, grouped,
              ' ** with Min(fsp_id), for fsfd_order = 2), linked to tblForm_Shortcut_Form,
              ' ** fsf_id, frm_caption_item, frm_name_item.
6370          .lbxForms.RowSource = "qryForm_Shortcut_04_02"
6380          .lbxForms.Requery
6390          DoEvents
6400        End If
6410        strFilter = .frmXAdmin_Shortcut_Sub.Form.Filter
6420        If strFilter <> vbNullString Then
6430          intPos01 = InStr(strFilter, "[fsf_show2]")
6440          If intPos01 > 0 Then
6450            If intPos01 > 1 Then
6460              strFilter = Trim(Left(strFilter, (intPos01 - 1)))
6470              If Right(strFilter, 3) = "And" Then strFilter = Trim(Left(strFilter, (Len(strFilter) - 3)))
6480              If strFilter <> vbNullString Then
6490                .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
6500                .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
6510              Else
6520                .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
6530                .frmXAdmin_Shortcut_Sub.Form.FilterOn = False
6540              End If
6550            Else
                  ' ** At the beginning of the filter.
6560              intPos01 = InStr(strFilter, "And")
6570              If intPos01 > 0 Then
6580                strFilter = Trim(Mid(strFilter, (intPos01 + 3)))
6590                .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
6600                .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
6610              Else
6620                strFilter = vbNullString
6630                .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
6640                .frmXAdmin_Shortcut_Sub.Form.FilterOn = False
6650              End If
6660            End If
6670          Else
                ' ** Leave as-is.
6680          End If
6690        Else
              ' ** We're good to go.
6700        End If
6710      Case False
6720        .chkIncludeMiscForms_lbl.FontBold = False
6730        If .lbxForms.RowSource <> "qryForm_Shortcut_04_03" Then
              ' ** qryForm_Shortcut_03_02 (tblVBComponent_Shortcut_New, grouped,
              ' ** with Min(fsp_id), for fsfd_order = 2), linked to tblForm_Shortcut_Form,
              ' ** tblVBComponent_Shortcut_Form, above, but without fsf_show2.
6740          .lbxForms.RowSource = "qryForm_Shortcut_04_03"
6750          .lbxForms.Requery
6760          DoEvents
6770        End If
6780        strFilter = .frmXAdmin_Shortcut_Sub.Form.Filter
6790        If strFilter <> vbNullString Then
6800          intPos01 = InStr(strFilter, "[fsf_show2]")
6810          If intPos01 > 0 Then
                ' ** What's it doing here!
6820            If InStr(strFilter, "[fsf_show2]=False") > 0 Then
                  ' ** Just leave it alone, but make sure the filter is on!
6830              .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
6840            ElseIf intPos01 > 1 Then
6850              strFilter = Trim(Left(strFilter, (intPos01 - 1)))
6860              If Right(strFilter, 3) = "And" Then
6870                strFilter = strFilter & " [fsf_show2]=False"
6880              Else
6890                strFilter = strFilter & " And [fsf_show2]=False"
6900              End If
6910              .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
6920              .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
6930            Else
6940              intPos01 = InStr(strFilter, "And")
6950              If intPos01 > 0 Then
6960                strFilter = Trim(Mid(strFilter, (intPos01 + 3))) & " And [fsf_show2]=False"
6970                .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
6980                .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
6990              Else
                    ' ** I dont know what it is!
7000                strFilter = "[fsf_show2]=False"
7010                .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
7020                .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
7030              End If
7040            End If
7050          Else
7060            strFilter = strFilter & " And [fsf_show2]=False"
7070            .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
7080            .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
7090          End If
7100        Else
7110          strFilter = "[fsf_show2]=False"
7120          .frmXAdmin_Shortcut_Sub.Form.Filter = strFilter
7130          .frmXAdmin_Shortcut_Sub.Form.FilterOn = True
7140        End If
7150      End Select
7160      .frmXAdmin_Shortcut_Sub.Form.Requery
7170      blnRecover = True
7180      .TimerInterval = 500&
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

Private Sub chkIncludeMiscForms_KeyDown(KeyCode As Integer, Shift As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeMiscForms_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7320    intShiftDown = (Shift And acShiftMask) > 0
7330    intAltDown = (Shift And acAltMask) > 0
7340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7370        With Me
7380          intRetVal = 0
7390          lngRecsCur = .frmXAdmin_Shortcut_Sub.Form.RecCnt  ' ** Form Function: frmXAdmin_Shortcut_Sub.
7400          If lngRecsCur > 0& Then
7410            .frmXAdmin_Shortcut_Sub.SetFocus
7420          Else
7430            .cmdClose.SetFocus
7440          End If
7450        End With
7460      End Select
7470    End If

        ' ** Shift keys.
7480    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7490      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7500        With Me
7510          intRetVal = 0
7520          .opgSearchMode.SetFocus
7530        End With
7540      End Select
7550    End If

EXITP:
7560    KeyCode = intRetVal
7570    Exit Sub

ERRH:
7580    intRetVal = 0
7590    Select Case ERR.Number
        Case Else
7600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7610    End Select
7620    Resume EXITP

End Sub

Private Sub ShowFormName(blnShow As Boolean)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowFormName"

        Dim lngDiff As Long, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String, lngTmp03 As Long, lngTmp04 As Long
        Dim lngX As Long

7710    With Me

7720      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
7730        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
7740      End If

          ' ** Variables are fed empty, then populated ByRef.
7750      GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Function: modWindowFunctions.

7760      lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
7770      lngMonitorNum = 1&: lngTmp04 = 0&
7780      EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
7790      If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

7800      Select Case blnShow
          Case True

7810        If lngMonitorNum = 1& Then lngTmp04 = lngFrm_Top
7820        DoCmd.MoveSize lngFrm_Left, lngTmp04, lngFrm_Width, lngFrm_Height  'lngFrm_Top
7830        If lngMonitorNum > 1& Then
7840          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
7850        End If

7860        .Width = lngForm_Width

7870        .opgSearchMode_box.Left = lngOpgSearchMode_Left
7880        .opgSearchMode_box2.Left = (.opgSearchMode_box.Left + lngBox2_Offset)
7890        .opgSearchMode_lbl.Left = (.opgSearchMode_box2.Left + lngTpp)
7900        .opgSearchMode_hline02.Left = (.opgSearchMode_box2.Left - lngTpp)
7910        .opgSearchMode_hline01.Left = (.opgSearchMode_hline02.Left - lngTpp)
7920        .opgSearchMode_hline03.Left = .opgSearchMode_box2.Left
7930        .opgSearchMode_vline02.Left = (.opgSearchMode_box2.Left - lngTpp)
7940        .opgSearchMode_vline01.Left = (.opgSearchMode_vline02.Left - lngTpp)
7950        .opgSearchMode_vline03.Left = ((.opgSearchMode_box2.Left + .opgSearchMode_box2.Width) + lngTpp)
7960        .opgSearchMode_vline04.Left = (.opgSearchMode_vline03.Left + lngTpp)

7970        lngDiff = (.opgSearchMode_box.Left - .opgSearchMode.Left)
7980        lngTmp03 = (lngDiff / lngTpp)
            ' ** Move is incremental because frame will stretch if moved without options.
7990        For lngX = 1& To lngTmp03
8000          .opgSearchMode_optAll.Left = (.opgSearchMode_optAll.Left + lngTpp)
8010          .opgSearchMode_optAll_lbl.Left = (.opgSearchMode_optAll_lbl.Left + lngTpp)
8020          .opgSearchMode_optSpecified.Left = (.opgSearchMode_optSpecified.Left + lngTpp)
8030          .opgSearchMode_optSpecified_lbl.Left = (.opgSearchMode_optSpecified_lbl.Left + lngTpp)
8040          .opgSearchMode.Left = (.opgSearchMode.Left + lngTpp)
8050        Next
            ' ** Now make sure it's back to normal.
8060        .opgSearchMode.Left = lngOpgSearchMode_Left
8070        .opgSearchMode_optAll.Left = (.opgSearchMode.Left + lngOpt_Offset)
8080        .opgSearchMode_optAll_lbl.Left = (.opgSearchMode_optAll.Left + lngOptLbl_Offset)
8090        .opgSearchMode_optSpecified.Left = (.opgSearchMode.Left + lngOpt_Offset)
8100        .opgSearchMode_optSpecified_lbl.Left = (.opgSearchMode_optSpecified.Left + lngOptLbl_Offset)
8110        .opgSearchMode.Width = .opgSearchMode_box.Width

8120        For lngX = 0& To (lngCtls - 1&)
8130          .Controls(arr_varCtl(C_CNAM, lngX)).Left = arr_varCtl(C_LFT, lngX)
8140        Next

8150        .lbxForms_box.Width = lngLbxFormsBox_Width
8160        .lbxForms.Width = lngLbxForms_Width
8170        .frm_name_item_lbl.Visible = True
8180        .frm_name_item_lbl_line.Visible = True
8190        .fsp_order_lbl.Left = lngOrder_Left
8200        .fsp_order_lbl_line.Left = (.fsp_order_lbl.Left + lngTpp)
8210        .lbxForms.ColumnWidths = strLbxCol_Widths

8220        .frmXAdmin_Shortcut_Sub.Width = lngSub_Width
8230        .frmXAdmin_Shortcut_Sub_box.Width = (.frmXAdmin_Shortcut_Sub.Width + (2& * lngTpp))
8240        .Nav_box01.Width = .frmXAdmin_Shortcut_Sub.Width
8250        .cmdClose.Left = lngClose_Left

8260        .Header_vline01.Left = lngForm_Width
8270        .Header_vline02.Left = .Header_vline01.Left
8280        .Header_hline01.Width = lngForm_Width
8290        .Header_hline02.Width = .Header_hline01.Width
8300        .Footer_vline01.Left = lngForm_Width
8310        .Footer_vline02.Left = .Footer_vline01.Left
8320        .Footer_hline01.Width = lngForm_Width
8330        .Footer_hline02.Width = .Footer_hline01.Width

8340        .Width = lngForm_Width

8350      Case False

            ' ** .lbxForms.ColumnWidths = "0;3150;315;3735;0;0;0"
8360        intPos01 = CharPos(strLbxCol_Widths, 4, ";")  ' ** Module Function: modStringFuncs.
8370        strTmp01 = Left(strLbxCol_Widths, (intPos01 - 1))
8380        intPos01 = CharPos(strTmp01, 3, ";")  ' ** Module Function: modStringFuncs.
8390        strTmp01 = Mid(strTmp01, (intPos01 + 1))  ' ** Width of form name column, in Twips.
8400        lngDiff = Val(strTmp01)
8410        strTmp02 = StringReplace(strLbxCol_Widths, strTmp01, "0")  ' ** Depends on it being unique.
8420        .lbxForms.ColumnWidths = strTmp02
8430        .frm_name_item_lbl.Visible = False
8440        .frm_name_item_lbl_line.Visible = False

8450        .fsfd_section_lbl.Left = (.fsfd_section_lbl.Left - (3& * lngTpp))
8460        .fsfd_section_lbl_line.Left = (.fsfd_section_lbl_line.Left - (3& * lngTpp))
8470        .frm_caption_item_lbl_line.Width = (.frm_caption_item_lbl_line.Width - (3& * lngTpp))

8480        .fsp_order_lbl.Left = (lngOrder_Left - lngDiff)
8490        .fsp_order_lbl_line.Left = (.fsp_order_lbl.Left + lngTpp)
8500        .lbxForms.Width = (lngLbxForms_Width - lngDiff)
8510        .lbxForms_box.Width = (lngLbxFormsBox_Width - lngDiff)

            ' ** lbl to box2: 15, lngTpp
            ' ** box2 to hline01: 30, 2 * lngTpp

8520        lngTmp03 = (lngOpgSearchMode_Left - lngDiff)
8530        .opgSearchMode_box.Left = lngTmp03
8540        .opgSearchMode_box2.Left = (.opgSearchMode_box.Left + lngBox2_Offset)
8550        .opgSearchMode_lbl.Left = (.opgSearchMode_box2.Left + lngTpp)
8560        .opgSearchMode_hline02.Left = (.opgSearchMode_box2.Left - lngTpp)
8570        .opgSearchMode_hline01.Left = (.opgSearchMode_hline02.Left - lngTpp)
8580        .opgSearchMode_hline03.Left = .opgSearchMode_box2.Left
8590        .opgSearchMode_vline02.Left = (.opgSearchMode_box2.Left - lngTpp)
8600        .opgSearchMode_vline01.Left = (.opgSearchMode_vline02.Left - lngTpp)
8610        .opgSearchMode_vline03.Left = ((.opgSearchMode_box2.Left + .opgSearchMode_box2.Width) + lngTpp)
8620        .opgSearchMode_vline04.Left = (.opgSearchMode_vline03.Left + lngTpp)

8630        lngTmp03 = (lngDiff / lngTpp)  ' ** Rounding is intentional.
            ' ** Move is incremental because frame will stretch if moved without options.
8640        For lngX = 1& To lngTmp03
8650          .opgSearchMode_optAll.Left = (.opgSearchMode_optAll.Left - lngTpp)
8660          .opgSearchMode_optAll_lbl.Left = (.opgSearchMode_optAll_lbl.Left - lngTpp)
8670          .opgSearchMode_optSpecified.Left = (.opgSearchMode_optSpecified.Left - lngTpp)
8680          .opgSearchMode_optSpecified_lbl.Left = (.opgSearchMode_optSpecified_lbl.Left - lngTpp)
8690          .opgSearchMode.Left = (.opgSearchMode.Left - lngTpp)
8700        Next

8710        lngDiff = (arr_varCtl(C_LFT, 0) - .opgSearchMode_box.Left)
8720        For lngX = 0& To (lngCtls - 1&)
8730          .Controls(arr_varCtl(C_CNAM, lngX)).Left = (arr_varCtl(C_LFT, lngX) - lngDiff)
8740        Next

8750        lngDiff = lngFrmName_Width
8760        .cmdClose.Left = (lngClose_Left - lngDiff)
8770        .frmXAdmin_Shortcut_Sub.Width = (lngSub_Width - lngDiff)
8780        .frmXAdmin_Shortcut_Sub_box.Width = (.frmXAdmin_Shortcut_Sub.Width + (2& * lngTpp))
8790        .Nav_box01.Width = .frmXAdmin_Shortcut_Sub.Width

8800        .Header_vline01.Left = (lngForm_Width - lngDiff)
8810        .Header_vline02.Left = .Header_vline01.Left
8820        .Header_hline01.Width = (lngForm_Width - lngDiff)
8830        .Header_hline02.Width = .Header_hline01.Width
8840        .Footer_vline01.Left = (lngForm_Width - lngDiff)
8850        .Footer_vline02.Left = .Footer_vline01.Left
8860        .Footer_hline01.Width = (lngForm_Width - lngDiff)
8870        .Footer_hline02.Width = .Footer_hline01.Width

8880        .Width = (lngForm_Width - lngDiff)

8890        lngTmp03 = (lngDiff / 2)  ' ** The rounding here is OK.

            ' ** Center the new size.
8900        If lngMonitorNum = 1& Then lngTmp04 = lngFrm_Top
8910        DoCmd.MoveSize (lngFrm_Left + lngTmp03), lngTmp04, (lngFrm_Width - lngDiff), lngFrm_Height  'lngFrm_Top
8920        If lngMonitorNum > 1& Then
8930          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
8940        End If

8950      End Select

8960    End With

EXITP:
8970    Exit Sub

ERRH:
8980    Select Case ERR.Number
        Case Else
8990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9000    End Select
9010    Resume EXITP

End Sub

Private Sub SetFrmNumbering()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "SetFrmNumbering"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngItems As Long, arr_varItem() As Variant
        Dim lngRecs As Long
        Dim lngX As Long, lngE As Long

        ' ** Array: arr_varItem().
        Const I_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const I_ORD As Integer = 0
        Const I_FID As Integer = 1
        Const I_FND As Integer = 2

9110    With Me

9120      Set dbs = CurrentDb

          ' ** Update tblVBComponent_Shortcut_Form, for fsp_order = 0.
9130      Set qdf = dbs.QueryDefs("qryForm_Shortcut_06_02")
9140      qdf.Execute
9150      Set qdf = Nothing
9160      DoEvents

9170      lngItems = 0&
9180      ReDim arr_varItem(I_ELEMS, 0)

9190      Select Case .chkIncludeMiscForms
          Case True
            ' ** qryForm_Shortcut_03_02 (tblVBComponent_Shortcut_New,
            ' ** grouped, with Min(fsp_id), for fsfd_order = 2), linked to
            ' ** tblForm_Shortcut_Form, fsf_id, frm_caption_item, frm_name_item.
9200        Set qdf = dbs.QueryDefs("qryForm_Shortcut_04_02")
9210      Case False
            ' ** qryForm_Shortcut_03_02 (tblVBComponent_Shortcut_New,
            ' ** grouped, with Min(fsp_id), for fsfd_order = 2), linked to
            ' ** tblForm_Shortcut_Form, tblVBComponent_Shortcut_Form, above,
            ' ** but without fsf_show2.
9220        Set qdf = dbs.QueryDefs("qryForm_Shortcut_04_03")
9230      End Select
9240      Set rst = qdf.OpenRecordset
9250      With rst
9260        .MoveLast
9270        lngRecs = .RecordCount
9280        .MoveFirst
9290        For lngX = 1& To lngRecs
9300          lngItems = lngItems + 1&
9310          lngE = lngItems - 1&
9320          ReDim Preserve arr_varItem(I_ELEMS, lngE)
9330          arr_varItem(I_ORD, lngE) = lngX
9340          arr_varItem(I_FID, lngE) = ![frm_id]
9350          arr_varItem(I_FND, lngE) = CBool(False)
9360          If lngX < lngRecs Then .MoveNext
9370        Next
9380        .Close
9390      End With
9400      Set rst = Nothing
9410      DoEvents

9420      Set rst = dbs.OpenRecordset("tblVBComponent_Shortcut_Form", dbOpenDynaset, dbConsistent)
9430      With rst
9440        .MoveFirst
9450        For lngX = 0& To (lngItems - 1&)
9460          If ![frm_id] = arr_varItem(I_FID, lngX) Then
9470            arr_varItem(I_FND, lngX) = CBool(True)
9480            .Edit
9490            ![fsp_order] = arr_varItem(I_ORD, lngX)
9500            ![vbcomscf_datemodified] = Now()
9510            .Update
9520          Else
9530            .FindFirst "[frm_id] = " & CStr(arr_varItem(I_FID, lngX))
9540            If .NoMatch = False Then
9550              arr_varItem(I_FND, lngX) = CBool(True)
9560              .Edit
9570              ![fsp_order] = arr_varItem(I_ORD, lngX)
9580              ![vbcomscf_datemodified] = Now()
9590              .Update
9600            End If
9610          End If
9620        Next
9630        .Close
9640      End With
9650      Set rst = Nothing
9660      DoEvents

9670      dbs.Close
9680      Set dbs = Nothing
9690      DoEvents

9700      .frmXAdmin_Shortcut_Sub.Form.Requery
9710      DoEvents

9720    End With

EXITP:
9730    Set rst = Nothing
9740    Set qdf = Nothing
9750    Set dbs = Nothing
9760    Exit Sub

ERRH:
9770    Select Case ERR.Number
        Case Else
9780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9790    End Select
9800    Resume EXITP

End Sub

Private Sub LoadKeyArray()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "LoadKeyArray"

        Dim lngE As Long

9910    With Me

9920      lngCtls = 0&
9930      ReDim arr_varCtl(C_ELEMS, 0)

9940      lngCtls = lngCtls + 1&
9950      lngE = lngCtls - 1&
9960      ReDim Preserve arr_varCtl(C_ELEMS, lngE)
9970      arr_varCtl(C_CNAM, lngE) = .KeyDownType_box.Name
9980      arr_varCtl(C_LFT, lngE) = .KeyDownType_box.Left

9990      lngCtls = lngCtls + 1&
10000     lngE = lngCtls - 1&
10010     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10020     arr_varCtl(C_CNAM, lngE) = .KeyDownType_box2.Name
10030     arr_varCtl(C_LFT, lngE) = .KeyDownType_box2.Left

10040     lngCtls = lngCtls + 1&
10050     lngE = lngCtls - 1&
10060     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10070     arr_varCtl(C_CNAM, lngE) = .KeyDownType_lbl.Name
10080     arr_varCtl(C_LFT, lngE) = .KeyDownType_lbl.Left

10090     lngCtls = lngCtls + 1&
10100     lngE = lngCtls - 1&
10110     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10120     arr_varCtl(C_CNAM, lngE) = .KeyDownType_hline01.Name
10130     arr_varCtl(C_LFT, lngE) = .KeyDownType_hline01.Left

10140     lngCtls = lngCtls + 1&
10150     lngE = lngCtls - 1&
10160     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10170     arr_varCtl(C_CNAM, lngE) = .KeyDownType_hline02.Name
10180     arr_varCtl(C_LFT, lngE) = .KeyDownType_hline02.Left

10190     lngCtls = lngCtls + 1&
10200     lngE = lngCtls - 1&
10210     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10220     arr_varCtl(C_CNAM, lngE) = .KeyDownType_hline03.Name
10230     arr_varCtl(C_LFT, lngE) = .KeyDownType_hline03.Left

10240     lngCtls = lngCtls + 1&
10250     lngE = lngCtls - 1&
10260     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10270     arr_varCtl(C_CNAM, lngE) = .KeyDownType_vline01.Name
10280     arr_varCtl(C_LFT, lngE) = .KeyDownType_vline01.Left

10290     lngCtls = lngCtls + 1&
10300     lngE = lngCtls - 1&
10310     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10320     arr_varCtl(C_CNAM, lngE) = .KeyDownType_vline02.Name
10330     arr_varCtl(C_LFT, lngE) = .KeyDownType_vline02.Left

10340     lngCtls = lngCtls + 1&
10350     lngE = lngCtls - 1&
10360     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10370     arr_varCtl(C_CNAM, lngE) = .KeyDownType_vline03.Name
10380     arr_varCtl(C_LFT, lngE) = .KeyDownType_vline03.Left

10390     lngCtls = lngCtls + 1&
10400     lngE = lngCtls - 1&
10410     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10420     arr_varCtl(C_CNAM, lngE) = .KeyDownType_vline04.Name
10430     arr_varCtl(C_LFT, lngE) = .KeyDownType_vline04.Left

10440     lngCtls = lngCtls + 1&
10450     lngE = lngCtls - 1&
10460     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10470     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Plain_box.Name
10480     arr_varCtl(C_LFT, lngE) = .KeyDownType_Plain_box.Left

10490     lngCtls = lngCtls + 1&
10500     lngE = lngCtls - 1&
10510     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10520     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Plain_lbl.Name
10530     arr_varCtl(C_LFT, lngE) = .KeyDownType_Plain_lbl.Left

10540     lngCtls = lngCtls + 1&
10550     lngE = lngCtls - 1&
10560     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10570     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Alt_box.Name
10580     arr_varCtl(C_LFT, lngE) = .KeyDownType_Alt_box.Left

10590     lngCtls = lngCtls + 1&
10600     lngE = lngCtls - 1&
10610     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10620     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Alt_lbl.Name
10630     arr_varCtl(C_LFT, lngE) = .KeyDownType_Alt_lbl.Left

10640     lngCtls = lngCtls + 1&
10650     lngE = lngCtls - 1&
10660     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10670     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Shift_box.Name
10680     arr_varCtl(C_LFT, lngE) = .KeyDownType_Shift_box.Left

10690     lngCtls = lngCtls + 1&
10700     lngE = lngCtls - 1&
10710     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10720     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Shift_lbl.Name
10730     arr_varCtl(C_LFT, lngE) = .KeyDownType_Shift_lbl.Left

10740     lngCtls = lngCtls + 1&
10750     lngE = lngCtls - 1&
10760     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10770     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Ctrl_box.Name
10780     arr_varCtl(C_LFT, lngE) = .KeyDownType_Ctrl_box.Left

10790     lngCtls = lngCtls + 1&
10800     lngE = lngCtls - 1&
10810     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10820     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Ctrl_lbl.Name
10830     arr_varCtl(C_LFT, lngE) = .KeyDownType_Ctrl_lbl.Left

10840     lngCtls = lngCtls + 1&
10850     lngE = lngCtls - 1&
10860     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10870     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Ctrl_Alt_box.Name
10880     arr_varCtl(C_LFT, lngE) = .KeyDownType_Ctrl_Alt_box.Left

10890     lngCtls = lngCtls + 1&
10900     lngE = lngCtls - 1&
10910     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10920     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Ctrl_Alt_lbl.Name
10930     arr_varCtl(C_LFT, lngE) = .KeyDownType_Ctrl_Alt_lbl.Left

10940     lngCtls = lngCtls + 1&
10950     lngE = lngCtls - 1&
10960     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
10970     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Ctrl_Shift_box.Name
10980     arr_varCtl(C_LFT, lngE) = .KeyDownType_Ctrl_Shift_box.Left

10990     lngCtls = lngCtls + 1&
11000     lngE = lngCtls - 1&
11010     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
11020     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Ctrl_Shift_lbl.Name
11030     arr_varCtl(C_LFT, lngE) = .KeyDownType_Ctrl_Shift_lbl.Left

11040     lngCtls = lngCtls + 1&
11050     lngE = lngCtls - 1&
11060     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
11070     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Alt_Shift_box.Name
11080     arr_varCtl(C_LFT, lngE) = .KeyDownType_Alt_Shift_box.Left

11090     lngCtls = lngCtls + 1&
11100     lngE = lngCtls - 1&
11110     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
11120     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Alt_Shift_lbl.Name
11130     arr_varCtl(C_LFT, lngE) = .KeyDownType_Alt_Shift_lbl.Left

11140     lngCtls = lngCtls + 1&
11150     lngE = lngCtls - 1&
11160     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
11170     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Ctrl_Alt_Shift_box.Name
11180     arr_varCtl(C_LFT, lngE) = .KeyDownType_Ctrl_Alt_Shift_box.Left

11190     lngCtls = lngCtls + 1&
11200     lngE = lngCtls - 1&
11210     ReDim Preserve arr_varCtl(C_ELEMS, lngE)
11220     arr_varCtl(C_CNAM, lngE) = .KeyDownType_Ctrl_Alt_Shift_lbl.Name
11230     arr_varCtl(C_LFT, lngE) = .KeyDownType_Ctrl_Alt_Shift_lbl.Left

11240   End With

EXITP:
11250   Exit Sub

ERRH:
11260   Select Case ERR.Number
        Case Else
11270     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11280   End Select
11290   Resume EXITP

End Sub
