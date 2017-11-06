VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLoadTimer_Mgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmLoadTimer_Mgmt"

'VGC 10/12/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Latest Only:      O {chkLatestOnly}
' **   Clear:            R {cmdClear}
' **   Elapsed Time:     T {xadload_time on frmLoadTimer_Mgmt_Sub}
' **   Exit:             X {cmdClose}

Private lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         strCallingForm = Trim(.OpenArgs)
140       End If

150       blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

160       Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

170       lngRecsCur = .frmLoadTimer_Mgmt_Sub.Form.RecCnt  ' ** Form Function: frmLoadTimer_Mgmt_Sub.
180       If lngRecsCur > 0 Then
190         .frmLoadTimer_Mgmt_Sub.SetFocus
200       Else
210         .frmLoadTimer_Mgmt_Sub.Form.ColHdrsSet False  ' ** Form Procedure: frmLoadTimer_Mgmt_Sub.
220         .cmdClose.SetFocus
230       End If

240       chkLatestOnly_AfterUpdate  ' ** Procedure: Below.
250       ckgVersion_chkAll_AfterUpdate  ' ** Procedure: Below.
260       ckgVersion_chkAcc2003_AfterUpdate  ' ** Procedure: Below.
270       ckgVersion_chkAcc2007_AfterUpdate  ' ** Procedure: Below.
280       ckgVersion_chkAcc2010_AfterUpdate  ' ** Procedure: Below.
290       ckgVersion_chkAcc2013_AfterUpdate  ' ** Procedure: Below.
300       ckgVersion_chkAcc2016_AfterUpdate  ' ** Procedure: Below.

310       DoCmd.Hourglass False

320     End With

EXITP:
330     Exit Sub

ERRH:
340     DoCmd.Hourglass False
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

410     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
420     intShiftDown = (Shift And acShiftMask) > 0
430     intAltDown = (Shift And acAltMask) > 0
440     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Elapsed Time:     T {xadload_time on frmLoadTimer_Mgmt_Sub}
        ' **   Exit:             X {cmdClose}

        ' ** Alt keys.
450     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
460       Select Case intRetVal
          Case vbKeyT
470         With Me
480           intRetVal = 0
490           lngRecsCur = .frmLoadTimer_Mgmt_Sub.Form.RecCnt  ' ** Form Function: frmLoadTimer_Mgmt_Sub.
500           If lngRecsCur > 0& Then
510             .frmLoadTimer_Mgmt_Sub.SetFocus
520             .frmLoadTimer_Mgmt_Sub.Form.xadload_time.SetFocus
530           Else
540             Beep
550           End If
560         End With
570       Case vbKeyX
580         intRetVal = 0
590         cmdClose_Click  ' ** Procedure: Below.
600       End Select
610     End If

        ' ** Ctrl-Shift keys.
620     If intCtrlDown And (Not intAltDown) And intShiftDown Then
630       Select Case intRetVal
          Case vbKeyF
640         With Me
650           intRetVal = 0
660           .FocusHolder.SetFocus
670         End With
680       End Select
690     End If

EXITP:
700     KeyCode = intRetVal
710     Exit Sub

ERRH:
720     intRetVal = 0
730     Select Case ERR.Number
        Case Else
740       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
750     End Select
760     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim blnRetVal As Boolean

810     With Me

820       Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

830       blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel

840       If strCallingForm <> vbNullString Then
850         If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
860           gblnSetFocus = True
870           gstrReturningForm = THIS_NAME
880           DoCmd.SelectObject acForm, strCallingForm, False
890           Forms(strCallingForm).TimerInterval = 200&
900         Else
910           gblnSetFocus = True
920           gstrReturningForm = THIS_NAME
930           DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Utility"
940         End If
950       End If

960     End With

EXITP:
970     Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Public Sub cmdClose_Click()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

1110    With Me

1120      .TimerInterval = 0&

1130      DoCmd.Close acForm, THIS_NAME

1140    End With

EXITP:
1150    Exit Sub

ERRH:
1160    Select Case ERR.Number
        Case Else
1170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1180    End Select
1190    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1220    intShiftDown = (Shift And acShiftMask) > 0
1230    intAltDown = (Shift And acAltMask) > 0
1240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyTab
1270        With Me
1280          intRetVal = 0
1290          .frmLoadTimer_Mgmt_Sub.SetFocus
1300        End With
1310      End Select
1320    End If

        ' ** Shift keys.
1330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1340      Select Case intRetVal
          Case vbKeyTab
1350        With Me
1360          intRetVal = 0
1370          If .cmdClear.Enabled = True Then
1380            .cmdClear.SetFocus
1390          Else
1400            .chkLatestOnly.SetFocus
1410          End If
1420        End With
1430      End Select
1440    End If

EXITP:
1450    KeyCode = intRetVal
1460    Exit Sub

ERRH:
1470    intRetVal = 0
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub cmdClear_Click()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClear_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim msgResponse As VbMsgBoxResult

1610    With Me
1620      msgResponse = MsgBox("This will delete all timing records and empty the table." & vbCrLf & vbCrLf & _
            "Proceed?", vbQuestion + vbOKCancel, "Clear All Timing Recordsa")
1630      DoCmd.Hourglass True
1640      DoEvents
1650      If msgResponse = vbOK Then
1660        Set dbs = CurrentDb
            ' ** Empty tblXAdmin_Load.
1670        Set qdf = dbs.QueryDefs("qryFormLoad_02")
1680        qdf.Execute
1690        .frmLoadTimer_Mgmt_Sub.Form.Requery
1700        DoEvents
1710        .cmdClose.SetFocus
1720        DoEvents
1730        .cmdClear.Enabled = False
1740        Beep
1750      End If
1760      DoCmd.Hourglass False
1770    End With

EXITP:
1780    Set qdf = Nothing
1790    Set dbs = Nothing
1800    Exit Sub

ERRH:
1810    DoCmd.Hourglass False
1820    Select Case ERR.Number
        Case Else
1830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1840    End Select
1850    Resume EXITP

End Sub

Private Sub cmdClear_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClear_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1920    intShiftDown = (Shift And acShiftMask) > 0
1930    intAltDown = (Shift And acAltMask) > 0
1940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1960      Select Case intRetVal
          Case vbKeyTab
1970        With Me
1980          intRetVal = 0
1990          .cmdClose.SetFocus
2000        End With
2010      End Select
2020    End If

        ' ** Shift keys.
2030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2040      Select Case intRetVal
          Case vbKeyTab
2050        With Me
2060          intRetVal = 0
2070          .chkLatestOnly.SetFocus
2080        End With
2090      End Select
2100    End If

EXITP:
2110    KeyCode = intRetVal
2120    Exit Sub

ERRH:
2130    intRetVal = 0
2140    Select Case ERR.Number
        Case Else
2150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2160    End Select
2170    Resume EXITP

End Sub

Private Sub chkLatestOnly_AfterUpdate()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkLatestOnly_AfterUpdate"

        Dim frm As Access.Form

2210    With Me
2220      Set frm = .frmLoadTimer_Mgmt_Sub.Form
2230      Select Case .chkLatestOnly
          Case True
2240        .chkLatestOnly_lbl.FontBold = True
2250        Select Case .ckgVersion_chkAll
            Case True
2260          If frm.RecordSource <> "qryFormLoad_05" Then
                ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_04
                ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                ' ** with Max(xadload_datemodified)).
2270            frm.RecordSource = "qryFormLoad_05"
2280          End If
2290        Case False
2300          If frm.RecordSource <> "qryFormLoad_07" Then
                ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_06
                ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                ' ** xadload_access_version, with Max(xadload_datemodified)).
2310            frm.RecordSource = "qryFormLoad_07"
2320          End If
2330        End Select
2340        DoEvents
2350        frm.SortNow "Form_Load"
2360      Case False
2370        .chkLatestOnly_lbl.FontBold = False
2380        If frm.RecordSource <> "qryFormLoad_03" Then
              ' ** tblXAdmin_Load, all records.
2390          frm.RecordSource = "qryFormLoad_03"
2400          DoEvents
2410          frm.SortNow "Form_Load"
2420        End If
2430      End Select
2440      lngRecsCur = frm.RecCnt  ' ** Form Function: frmLoadTimer_Mgmt_Sub.
2450      If lngRecsCur > 0& Then
2460        .frmLoadTimer_Mgmt_Sub.SetFocus
2470      Else
2480        .FocusHolder.SetFocus
2490      End If
2500    End With

EXITP:
2510    Set frm = Nothing
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub chkLatestOnly_KeyDown(KeyCode As Integer, Shift As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkLatestOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2620    intShiftDown = (Shift And acShiftMask) > 0
2630    intAltDown = (Shift And acAltMask) > 0
2640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2660      Select Case intRetVal
          Case vbKeyTab
2670        With Me
2680          intRetVal = 0
2690          If .cmdClear.Enabled = True Then
2700            .cmdClear.SetFocus
2710          Else
2720            .cmdClose.SetFocus
2730          End If
2740        End With
2750      End Select
2760    End If

        ' ** Shift keys.
2770    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2780      Select Case intRetVal
          Case vbKeyTab
2790        With Me
2800          intRetVal = 0
2810          .frmLoadTimer_Mgmt_Sub.SetFocus
2820        End With
2830      End Select
2840    End If

EXITP:
2850    KeyCode = intRetVal
2860    Exit Sub

ERRH:
2870    intRetVal = 0
2880    Select Case ERR.Number
        Case Else
2890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2900    End Select
2910    Resume EXITP

End Sub

Private Sub ckgVersion_chkAll_AfterUpdate()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgVersion_chkAll_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String ', strOrderBy As String

3010    With Me
3020      Set frm = .frmLoadTimer_Mgmt_Sub.Form
3030      strFilter = vbNullString
3040      Select Case .ckgVersion_chkAll
          Case True
3050        .ckgVersion_chkAll_lbl.FontBold = True
            ' ** Uses either qryFormLoad_03 or qryFormLoad_05.
3060        Select Case .chkLatestOnly
            Case True
3070          If frm.RecordSource <> "qryFormLoad_05" Then
                ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_04
                ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                ' ** with Max(xadload_datemodified)).
3080            frm.RecordSource = "qryFormLoad_05"
3090          End If
3100        Case False
3110          If frm.RecordSource <> "qryFormLoad_03" Then
                ' ** tblXAdmin_Load, all records.
3120            frm.RecordSource = "qryFormLoad_03"
3130          End If
3140        End Select
3150        DoEvents
3160        If .ckgVersion_chkAcc2003 = True Then
3170          .ckgVersion_chkAcc2003 = False
3180          .ckgVersion_chkAcc2003_lbl.FontBold = False
3190        End If
3200        If .ckgVersion_chkAcc2007 = True Then
3210          .ckgVersion_chkAcc2007 = False
3220          .ckgVersion_chkAcc2007_lbl.FontBold = False
3230        End If
3240        If .ckgVersion_chkAcc2010 = True Then
3250          .ckgVersion_chkAcc2010 = False
3260          .ckgVersion_chkAcc2010_lbl.FontBold = False
3270        End If
3280        If .ckgVersion_chkAcc2013 = True Then
3290          .ckgVersion_chkAcc2013 = False
3300          .ckgVersion_chkAcc2013_lbl.FontBold = False
3310        End If
3320        If .ckgVersion_chkAcc2016 = True Then
3330          .ckgVersion_chkAcc2016 = False
3340          .ckgVersion_chkAcc2016_lbl.FontBold = False
3350        End If
3360        frm.Filter = strFilter  ' ** Null String.
3370        frm.FilterOn = False
3380        frm.SortNow "Form_Load"
3390      Case False
3400        .ckgVersion_chkAll_lbl.FontBold = False
            ' ** Uses either qryFormLoad_03 or qryFormLoad_07.
3410        Select Case .chkLatestOnly
            Case True
3420          If frm.RecordSource <> "qryFormLoad_07" Then
                ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_06
                ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                ' ** xadload_access_version, with Max(xadload_datemodified)).
3430            frm.RecordSource = "qryFormLoad_07"
3440          End If
3450        Case False
3460          If frm.RecordSource <> "qryFormLoad_03" Then
                ' ** tblXAdmin_Load, all records.
3470            frm.RecordSource = "qryFormLoad_03"
3480          End If
3490        End Select
3500        DoEvents
3510        If .ckgVersion_chkAcc2003 = False And .ckgVersion_chkAcc2007 = False And .ckgVersion_chkAcc2010 = False And _
                .ckgVersion_chkAcc2013 = False And .ckgVersion_chkAcc2016 = False Then
              ' ** Not all, but not anything else either.
3520          strFilter = "[xadload_access_version] = 'Access 0000'"  ' ** Intended to return no records.
3530          frm.Filter = strFilter
3540          frm.FilterOn = True
3550        Else
3560          If .ckgVersion_chkAcc2003 = True Then
3570            strFilter = "[xadload_access_version] = 'Access 2003'"
3580          End If
3590          If .ckgVersion_chkAcc2007 = True Then
3600            If strFilter <> vbNullString Then strFilter = strFilter & " Or "
3610            strFilter = strFilter & "[xadload_access_version] = 'Access 2007'"
3620          End If
3630          If .ckgVersion_chkAcc2010 = True Then
3640            If strFilter <> vbNullString Then strFilter = strFilter & " Or "
3650            strFilter = strFilter & "[xadload_access_version] = 'Access 2010'"
3660          End If
3670          If .ckgVersion_chkAcc2013 = True Then
3680            If strFilter <> vbNullString Then strFilter = strFilter & " Or "
3690            strFilter = strFilter & "[xadload_access_version] = 'Access 2013'"
3700          End If
3710          If .ckgVersion_chkAcc2016 = True Then
3720            If strFilter <> vbNullString Then strFilter = strFilter & " Or "
3730            strFilter = strFilter & "[xadload_access_version] = 'Access 2016'"
3740          End If
3750          frm.Filter = strFilter
3760          frm.FilterOn = True
3770        End If
3780        frm.SortNow "Form_Load"
3790      End Select
3800      lngRecsCur = frm.RecCnt  ' ** Form Function: frmLoadTimer_Mgmt_Sub.
3810      If lngRecsCur > 0& Then
3820        .frmLoadTimer_Mgmt_Sub.SetFocus
3830      Else
3840        .FocusHolder.SetFocus
3850      End If
3860    End With

        '.ckgVersion_chkAll
        '.ckgVersion_chkAcc2003
        '.ckgVersion_chkAcc2007
        '.ckgVersion_chkAcc2010
        '.ckgVersion_chkAcc2013
        '.ckgVersion_chkAcc2016

EXITP:
3870    Set frm = Nothing
3880    Exit Sub

ERRH:
3890    Select Case ERR.Number
        Case Else
3900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3910    End Select
3920    Resume EXITP

End Sub

Private Sub ckgVersion_chkAcc2003_AfterUpdate()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgVersion_chkAcc2003_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String ', strOrderBy As String

4010    With Me
4020      Set frm = .frmLoadTimer_Mgmt_Sub.Form
4030      strFilter = vbNullString
4040      Select Case .ckgVersion_chkAcc2003
          Case True
4050        .ckgVersion_chkAcc2003_lbl.FontBold = True
            ' ** Uses either qryFormLoad_03 or qryFormLoad_07.
4060        Select Case .chkLatestOnly
            Case True
4070          If frm.RecordSource <> "qryFormLoad_07" Then
                ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_06
                ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                ' ** xadload_access_version, with Max(xadload_datemodified)).
4080            frm.RecordSource = "qryFormLoad_07"
4090          End If
4100        Case False
4110          If frm.RecordSource <> "qryFormLoad_03" Then
                ' ** tblXAdmin_Load, all records.
4120            frm.RecordSource = "qryFormLoad_03"
4130          End If
4140        End Select
4150        DoEvents
4160        If .ckgVersion_chkAll = True Then
4170          .ckgVersion_chkAll = False
4180          .ckgVersion_chkAll_lbl.FontBold = False
4190        End If
4200        strFilter = "[xadload_access_version] = 'Access 2003'"
4210        If .ckgVersion_chkAcc2007 = True Then
4220          strFilter = strFilter & " Or "
4230          strFilter = strFilter & "[xadload_access_version] = 'Access 2007'"
4240        End If
4250        If .ckgVersion_chkAcc2010 = True Then
4260          strFilter = strFilter & " Or "
4270          strFilter = strFilter & "[xadload_access_version] = 'Access 2010'"
4280        End If
4290        If .ckgVersion_chkAcc2013 = True Then
4300          strFilter = strFilter & " Or "
4310          strFilter = strFilter & "[xadload_access_version] = 'Access 2013'"
4320        End If
4330        If .ckgVersion_chkAcc2016 = True Then
4340          strFilter = strFilter & " Or "
4350          strFilter = strFilter & "[xadload_access_version] = 'Access 2016'"
4360        End If
4370        .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
4380        .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
4390        .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
4400      Case False
4410        .ckgVersion_chkAcc2003_lbl.FontBold = False
            ' ** Uses qryFormLoad_03 or qryFormLoad_05 Or qryFormLoad_07.
4420        Select Case .chkLatestOnly
            Case True
4430          Select Case .ckgVersion_chkAll
              Case True
                ' ** If all is true, qryFormLoad_05.
4440            If frm.RecordSource <> "qryFormLoad_05" Then
                  ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_04
                  ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                  ' ** with Max(xadload_datemodified)).
4450              frm.RecordSource = "qryFormLoad_05"
4460            End If
4470          Case False
4480            If .ckgVersion_chkAcc2007 = True Or .ckgVersion_chkAcc2010 = True Or _
                    .ckgVersion_chkAcc2013 = True Or .ckgVersion_chkAcc2016 = True Then
                  ' ** If other versions are on, qryFormLoad_07.
4490              If frm.RecordSource <> "qryFormLoad_07" Then
                    ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_06
                    ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                    ' ** xadload_access_version, with Max(xadload_datemodified)).
4500                frm.RecordSource = "qryFormLoad_07"
4510              End If
4520            Else
                  ' ** If no versions are on, qryFormLoad_03 {arbitrary}.
4530              If frm.RecordSource <> "qryFormLoad_03" Then
                    ' ** tblXAdmin_Load, all records.
4540                frm.RecordSource = "qryFormLoad_03"
4550              End If
4560            End If
4570          End Select
4580        Case False
              ' ** If all is true, qryFormLoad_03.
              ' ** If other versions are on, qryFormLoad_03.
              ' ** If no versions are on, qryFormLoad_03 {arbitrary}.
4590          If frm.RecordSource <> "qryFormLoad_03" Then
                ' ** tblXAdmin_Load, all records.
4600            frm.RecordSource = "qryFormLoad_03"
4610          End If
4620        End Select
4630        DoEvents
4640        If .ckgVersion_chkAcc2007 = True Then
4650          strFilter = strFilter & "[xadload_access_version] = 'Access 2007'"
4660        End If
4670        If .ckgVersion_chkAcc2010 = True Then
4680          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
4690          strFilter = strFilter & "[xadload_access_version] = 'Access 2010'"
4700        End If
4710        If .ckgVersion_chkAcc2013 = True Then
4720          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
4730          strFilter = strFilter & "[xadload_access_version] = 'Access 2013'"
4740        End If
4750        If .ckgVersion_chkAcc2016 = True Then
4760          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
4770          strFilter = strFilter & "[xadload_access_version] = 'Access 2016'"
4780        End If
4790        If strFilter <> vbNullString Then
4800          If .ckgVersion_chkAll = True Then  ' ** It shouldn't be, anyway.
4810            .ckgVersion_chkAll = False
4820            .ckgVersion_chkAll_lbl.FontBold = False
4830          End If
4840          .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
4850          .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
4860          .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
4870        Else
4880          Select Case .ckgVersion_chkAll
              Case True
4890            .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter  ' ** Null String.
4900            .frmLoadTimer_Mgmt_Sub.Form.FilterOn = False
4910          Case False
4920            strFilter = "[xadload_access_version] = 'Access 0000'"  ' ** Intended to return no records.
4930            .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
4940            .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
4950          End Select
4960          .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
4970        End If
4980      End Select
4990      lngRecsCur = frm.RecCnt  ' ** Form Function: frmLoadTimer_Mgmt_Sub.
5000      If lngRecsCur > 0& Then
5010        .frmLoadTimer_Mgmt_Sub.SetFocus
5020      Else
5030        .FocusHolder.SetFocus
5040      End If
5050    End With

        '.ckgVersion_chkAll
        '.ckgVersion_chkAcc2003
        '.ckgVersion_chkAcc2007
        '.ckgVersion_chkAcc2010
        '.ckgVersion_chkAcc2013
        '.ckgVersion_chkAcc2016

EXITP:
5060    Set frm = Nothing
5070    Exit Sub

ERRH:
5080    Select Case ERR.Number
        Case Else
5090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5100    End Select
5110    Resume EXITP

End Sub

Private Sub ckgVersion_chkAcc2007_AfterUpdate()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgVersion_chkAcc2007_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String ', strOrderBy As String

5210    With Me
5220      Set frm = .frmLoadTimer_Mgmt_Sub.Form
5230      strFilter = vbNullString
5240      Select Case .ckgVersion_chkAcc2007
          Case True
5250        .ckgVersion_chkAcc2007_lbl.FontBold = True
            ' ** Uses either qryFormLoad_03 or qryFormLoad_07.
5260        Select Case .chkLatestOnly
            Case True
5270          If frm.RecordSource <> "qryFormLoad_07" Then
                ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_06
                ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                ' ** xadload_access_version, with Max(xadload_datemodified)).
5280            frm.RecordSource = "qryFormLoad_07"
5290          End If
5300        Case False
5310          If frm.RecordSource <> "qryFormLoad_03" Then
                ' ** tblXAdmin_Load, all records.
5320            frm.RecordSource = "qryFormLoad_03"
5330          End If
5340        End Select
5350        DoEvents
5360        If .ckgVersion_chkAll = True Then
5370          .ckgVersion_chkAll = False
5380          .ckgVersion_chkAll_lbl.FontBold = False
5390        End If
5400        If .ckgVersion_chkAcc2003 = True Then
5410          strFilter = "[xadload_access_version] = 'Access 2003'"
5420        End If
5430        If strFilter <> vbNullString Then strFilter = strFilter & " Or "
5440        strFilter = strFilter & "[xadload_access_version] = 'Access 2007'"
5450        If .ckgVersion_chkAcc2010 = True Then
5460          strFilter = strFilter & " Or "
5470          strFilter = strFilter & "[xadload_access_version] = 'Access 2010'"
5480        End If
5490        If .ckgVersion_chkAcc2013 = True Then
5500          strFilter = strFilter & " Or "
5510          strFilter = strFilter & "[xadload_access_version] = 'Access 2013'"
5520        End If
5530        If .ckgVersion_chkAcc2016 = True Then
5540          strFilter = strFilter & " Or "
5550          strFilter = strFilter & "[xadload_access_version] = 'Access 2016'"
5560        End If
5570        .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
5580        .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
5590        .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
5600      Case False
5610        .ckgVersion_chkAcc2007_lbl.FontBold = False
            ' ** Uses qryFormLoad_03 or qryFormLoad_05 Or qryFormLoad_07.
5620        Select Case .chkLatestOnly
            Case True
5630          Select Case .ckgVersion_chkAll
              Case True
                ' ** If all is true, qryFormLoad_05.
5640            If frm.RecordSource <> "qryFormLoad_05" Then
                  ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_04
                  ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                  ' ** with Max(xadload_datemodified)).
5650              frm.RecordSource = "qryFormLoad_05"
5660            End If
5670          Case False
5680            If .ckgVersion_chkAcc2007 = True Or .ckgVersion_chkAcc2010 = True Or _
                    .ckgVersion_chkAcc2013 = True Or .ckgVersion_chkAcc2016 = True Then
                  ' ** If other versions are on, qryFormLoad_07.
5690              If frm.RecordSource <> "qryFormLoad_07" Then
                    ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_06
                    ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                    ' ** xadload_access_version, with Max(xadload_datemodified)).
5700                frm.RecordSource = "qryFormLoad_07"
5710              End If
5720            Else
                  ' ** If no versions are on, qryFormLoad_03 {arbitrary}.
5730              If frm.RecordSource <> "qryFormLoad_03" Then
                    ' ** tblXAdmin_Load, all records.
5740                frm.RecordSource = "qryFormLoad_03"
5750              End If
5760            End If
5770          End Select
5780        Case False
              ' ** If all is true, qryFormLoad_03.
              ' ** If other versions are on, qryFormLoad_03.
              ' ** If no versions are on, qryFormLoad_03 {arbitrary}.
5790          If frm.RecordSource <> "qryFormLoad_03" Then
                ' ** tblXAdmin_Load, all records.
5800            frm.RecordSource = "qryFormLoad_03"
5810          End If
5820        End Select
5830        DoEvents
5840        If .ckgVersion_chkAcc2003 = True Then
5850          strFilter = "[xadload_access_version] = 'Access 2003'"
5860        End If
5870        If .ckgVersion_chkAcc2010 = True Then
5880          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
5890          strFilter = strFilter & "[xadload_access_version] = 'Access 2010'"
5900        End If
5910        If .ckgVersion_chkAcc2013 = True Then
5920          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
5930          strFilter = strFilter & "[xadload_access_version] = 'Access 2013'"
5940        End If
5950        If .ckgVersion_chkAcc2016 = True Then
5960          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
5970          strFilter = strFilter & "[xadload_access_version] = 'Access 2016'"
5980        End If
5990        If strFilter <> vbNullString Then
6000          If .ckgVersion_chkAll = True Then  ' ** It shouldn't be, anyway.
6010            .ckgVersion_chkAll = False
6020            .ckgVersion_chkAll_lbl.FontBold = False
6030          End If
6040          .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
6050          .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
6060          .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
6070        Else
6080          Select Case .ckgVersion_chkAll
              Case True
6090            .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter  ' ** Null String.
6100            .frmLoadTimer_Mgmt_Sub.Form.FilterOn = False
6110          Case False
6120            strFilter = "[xadload_access_version] = 'Access 0000'"  ' ** Intended to return no records.
6130            .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
6140            .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
6150          End Select
6160          .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
6170        End If
6180      End Select
6190      lngRecsCur = frm.RecCnt  ' ** Form Function: frmLoadTimer_Mgmt_Sub.
6200      If lngRecsCur > 0& Then
6210        .frmLoadTimer_Mgmt_Sub.SetFocus
6220      Else
6230        .FocusHolder.SetFocus
6240      End If
6250    End With

        '.ckgVersion_chkAll
        '.ckgVersion_chkAcc2003
        '.ckgVersion_chkAcc2007
        '.ckgVersion_chkAcc2010
        '.ckgVersion_chkAcc2013
        '.ckgVersion_chkAcc2016

EXITP:
6260    Set frm = Nothing
6270    Exit Sub

ERRH:
6280    Select Case ERR.Number
        Case Else
6290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6300    End Select
6310    Resume EXITP

End Sub

Private Sub ckgVersion_chkAcc2010_AfterUpdate()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgVersion_chkAcc2010_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String ', strOrderBy As String

6410    With Me
6420      Set frm = .frmLoadTimer_Mgmt_Sub.Form
6430      strFilter = vbNullString
6440      Select Case .ckgVersion_chkAcc2010
          Case True
6450        .ckgVersion_chkAcc2010_lbl.FontBold = True
            ' ** Uses either qryFormLoad_03 or qryFormLoad_07.
6460        Select Case .chkLatestOnly
            Case True
6470          If frm.RecordSource <> "qryFormLoad_07" Then
                ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_06
                ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                ' ** xadload_access_version, with Max(xadload_datemodified)).
6480            frm.RecordSource = "qryFormLoad_07"
6490          End If
6500        Case False
6510          If frm.RecordSource <> "qryFormLoad_03" Then
                ' ** tblXAdmin_Load, all records.
6520            frm.RecordSource = "qryFormLoad_03"
6530          End If
6540        End Select
6550        DoEvents
6560        If .ckgVersion_chkAll = True Then
6570          .ckgVersion_chkAll = False
6580          .ckgVersion_chkAll_lbl.FontBold = False
6590        End If
6600        If .ckgVersion_chkAcc2003 = True Then
6610          strFilter = "[xadload_access_version] = 'Access 2003'"
6620        End If
6630        If .ckgVersion_chkAcc2007 = True Then
6640          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
6650          strFilter = strFilter & "[xadload_access_version] = 'Access 2007'"
6660        End If
6670        If strFilter <> vbNullString Then strFilter = strFilter & " Or "
6680        strFilter = strFilter & "[xadload_access_version] = 'Access 2010'"
6690        If .ckgVersion_chkAcc2013 = True Then
6700          strFilter = strFilter & " Or "
6710          strFilter = strFilter & "[xadload_access_version] = 'Access 2013'"
6720        End If
6730        If .ckgVersion_chkAcc2016 = True Then
6740          strFilter = strFilter & " Or "
6750          strFilter = strFilter & "[xadload_access_version] = 'Access 2016'"
6760        End If
6770        .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
6780        .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
6790        .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
6800      Case False
6810        .ckgVersion_chkAcc2010_lbl.FontBold = False
            ' ** Uses qryFormLoad_03 or qryFormLoad_05 Or qryFormLoad_07.
6820        Select Case .chkLatestOnly
            Case True
6830          Select Case .ckgVersion_chkAll
              Case True
                ' ** If all is true, qryFormLoad_05.
6840            If frm.RecordSource <> "qryFormLoad_05" Then
                  ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_04
                  ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                  ' ** with Max(xadload_datemodified)).
6850              frm.RecordSource = "qryFormLoad_05"
6860            End If
6870          Case False
6880            If .ckgVersion_chkAcc2007 = True Or .ckgVersion_chkAcc2010 = True Or _
                    .ckgVersion_chkAcc2013 = True Or .ckgVersion_chkAcc2016 = True Then
                  ' ** If other versions are on, qryFormLoad_07.
6890              If frm.RecordSource <> "qryFormLoad_07" Then
                    ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_06
                    ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                    ' ** xadload_access_version, with Max(xadload_datemodified)).
6900                frm.RecordSource = "qryFormLoad_07"
6910              End If
6920            Else
                  ' ** If no versions are on, qryFormLoad_03 {arbitrary}.
6930              If frm.RecordSource <> "qryFormLoad_03" Then
                    ' ** tblXAdmin_Load, all records.
6940                frm.RecordSource = "qryFormLoad_03"
6950              End If
6960            End If
6970          End Select
6980        Case False
              ' ** If all is true, qryFormLoad_03.
              ' ** If other versions are on, qryFormLoad_03.
              ' ** If no versions are on, qryFormLoad_03 {arbitrary}.
6990          If frm.RecordSource <> "qryFormLoad_03" Then
                ' ** tblXAdmin_Load, all records.
7000            frm.RecordSource = "qryFormLoad_03"
7010          End If
7020        End Select
7030        DoEvents
7040        If .ckgVersion_chkAcc2003 = True Then
7050          strFilter = "[xadload_access_version] = 'Access 2003'"
7060        End If
7070        If .ckgVersion_chkAcc2007 = True Then
7080          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
7090          strFilter = strFilter & "[xadload_access_version] = 'Access 2007'"
7100        End If
7110        If .ckgVersion_chkAcc2013 = True Then
7120          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
7130          strFilter = strFilter & "[xadload_access_version] = 'Access 2013'"
7140        End If
7150        If .ckgVersion_chkAcc2016 = True Then
7160          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
7170          strFilter = strFilter & "[xadload_access_version] = 'Access 2016'"
7180        End If
7190        If strFilter <> vbNullString Then
7200          If .ckgVersion_chkAll = True Then  ' ** It shouldn't be, anyway.
7210            .ckgVersion_chkAll = False
7220            .ckgVersion_chkAll_lbl.FontBold = False
7230          End If
7240          .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
7250          .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
7260          .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
7270        Else
7280          Select Case .ckgVersion_chkAll
              Case True
7290            .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter  ' ** Null String.
7300            .frmLoadTimer_Mgmt_Sub.Form.FilterOn = False
7310          Case False
7320            strFilter = "[xadload_access_version] = 'Access 0000'"  ' ** Intended to return no records.
7330            .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
7340            .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
7350          End Select
7360          .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
7370        End If
7380      End Select
7390      lngRecsCur = frm.RecCnt  ' ** Form Function: frmLoadTimer_Mgmt_Sub.
7400      If lngRecsCur > 0& Then
7410        .frmLoadTimer_Mgmt_Sub.SetFocus
7420      Else
7430        .FocusHolder.SetFocus
7440      End If
7450    End With

        '.ckgVersion_chkAll
        '.ckgVersion_chkAcc2003
        '.ckgVersion_chkAcc2007
        '.ckgVersion_chkAcc2010
        '.ckgVersion_chkAcc2013
        '.ckgVersion_chkAcc2016

EXITP:
7460    Set frm = Nothing
7470    Exit Sub

ERRH:
7480    Select Case ERR.Number
        Case Else
7490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7500    End Select
7510    Resume EXITP

End Sub

Private Sub ckgVersion_chkAcc2013_AfterUpdate()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgVersion_chkAcc2013_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String ', strOrderBy As String

7610    With Me
7620      Set frm = .frmLoadTimer_Mgmt_Sub.Form
7630      strFilter = vbNullString
7640      Select Case .ckgVersion_chkAcc2013
          Case True
7650        .ckgVersion_chkAcc2013_lbl.FontBold = True
            ' ** Uses either qryFormLoad_03 or qryFormLoad_07.
7660        Select Case .chkLatestOnly
            Case True
7670          If frm.RecordSource <> "qryFormLoad_07" Then
                ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_06
                ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                ' ** xadload_access_version, with Max(xadload_datemodified)).
7680            frm.RecordSource = "qryFormLoad_07"
7690          End If
7700        Case False
7710          If frm.RecordSource <> "qryFormLoad_03" Then
                ' ** tblXAdmin_Load, all records.
7720            frm.RecordSource = "qryFormLoad_03"
7730          End If
7740        End Select
7750        DoEvents
7760        If .ckgVersion_chkAll = True Then
7770          .ckgVersion_chkAll = False
7780          .ckgVersion_chkAll_lbl.FontBold = False
7790        End If
7800        If .ckgVersion_chkAcc2003 = True Then
7810          strFilter = "[xadload_access_version] = 'Access 2003'"
7820        End If
7830        If .ckgVersion_chkAcc2007 = True Then
7840          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
7850          strFilter = strFilter & "[xadload_access_version] = 'Access 2007'"
7860        End If
7870        If .ckgVersion_chkAcc2010 = True Then
7880          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
7890          strFilter = strFilter & "[xadload_access_version] = 'Access 2010'"
7900        End If
7910        If strFilter <> vbNullString Then strFilter = strFilter & " Or "
7920        strFilter = strFilter & "[xadload_access_version] = 'Access 2013'"
7930        If .ckgVersion_chkAcc2016 = True Then
7940          strFilter = strFilter & " Or "
7950          strFilter = strFilter & "[xadload_access_version] = 'Access 2016'"
7960        End If
7970        .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
7980        .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
7990        .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
8000      Case False
8010        .ckgVersion_chkAcc2013_lbl.FontBold = False
            ' ** Uses qryFormLoad_03 or qryFormLoad_05 Or qryFormLoad_07.
8020        Select Case .chkLatestOnly
            Case True
8030          Select Case .ckgVersion_chkAll
              Case True
                ' ** If all is true, qryFormLoad_05.
8040            If frm.RecordSource <> "qryFormLoad_05" Then
                  ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_04
                  ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                  ' ** with Max(xadload_datemodified)).
8050              frm.RecordSource = "qryFormLoad_05"
8060            End If
8070          Case False
8080            If .ckgVersion_chkAcc2007 = True Or .ckgVersion_chkAcc2010 = True Or _
                    .ckgVersion_chkAcc2013 = True Or .ckgVersion_chkAcc2016 = True Then
                  ' ** If other versions are on, qryFormLoad_07.
8090              If frm.RecordSource <> "qryFormLoad_07" Then
                    ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_06
                    ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                    ' ** xadload_access_version, with Max(xadload_datemodified)).
8100                frm.RecordSource = "qryFormLoad_07"
8110              End If
8120            Else
                  ' ** If no versions are on, qryFormLoad_03 {arbitrary}.
8130              If frm.RecordSource <> "qryFormLoad_03" Then
                    ' ** tblXAdmin_Load, all records.
8140                frm.RecordSource = "qryFormLoad_03"
8150              End If
8160            End If
8170          End Select
8180        Case False
              ' ** If all is true, qryFormLoad_03.
              ' ** If other versions are on, qryFormLoad_03.
              ' ** If no versions are on, qryFormLoad_03 {arbitrary}.
8190          If frm.RecordSource <> "qryFormLoad_03" Then
                ' ** tblXAdmin_Load, all records.
8200            frm.RecordSource = "qryFormLoad_03"
8210          End If
8220        End Select
8230        DoEvents
8240        If .ckgVersion_chkAcc2003 = True Then
8250          strFilter = "[xadload_access_version] = 'Access 2003'"
8260        End If
8270        If .ckgVersion_chkAcc2007 = True Then
8280          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
8290          strFilter = strFilter & "[xadload_access_version] = 'Access 2007'"
8300        End If
8310        If .ckgVersion_chkAcc2010 = True Then
8320          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
8330          strFilter = strFilter & "[xadload_access_version] = 'Access 2010'"
8340        End If
8350        If .ckgVersion_chkAcc2016 = True Then
8360          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
8370          strFilter = strFilter & "[xadload_access_version] = 'Access 2016'"
8380        End If
8390        If strFilter <> vbNullString Then
8400          If .ckgVersion_chkAll = True Then  ' ** It shouldn't be, anyway.
8410            .ckgVersion_chkAll = False
8420            .ckgVersion_chkAll_lbl.FontBold = False
8430          End If
8440          .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
8450          .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
8460          .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
8470        Else
8480          Select Case .ckgVersion_chkAll
              Case True
8490            .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter  ' ** Null String.
8500            .frmLoadTimer_Mgmt_Sub.Form.FilterOn = False
8510          Case False
8520            strFilter = "[xadload_access_version] = 'Access 0000'"  ' ** Intended to return no records.
8530            .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
8540            .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
8550          End Select
8560          .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
8570        End If
8580      End Select
8590      lngRecsCur = frm.RecCnt  ' ** Form Function: frmLoadTimer_Mgmt_Sub.
8600      If lngRecsCur > 0& Then
8610        .frmLoadTimer_Mgmt_Sub.SetFocus
8620      Else
8630        .FocusHolder.SetFocus
8640      End If
8650    End With

        '.ckgVersion_chkAll
        '.ckgVersion_chkAcc2003
        '.ckgVersion_chkAcc2007
        '.ckgVersion_chkAcc2010
        '.ckgVersion_chkAcc2013
        '.ckgVersion_chkAcc2016

EXITP:
8660    Set frm = Nothing
8670    Exit Sub

ERRH:
8680    Select Case ERR.Number
        Case Else
8690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8700    End Select
8710    Resume EXITP

End Sub

Private Sub ckgVersion_chkAcc2016_AfterUpdate()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgVersion_chkAcc2016_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String ', strOrderBy As String

8810    With Me
8820      Set frm = .frmLoadTimer_Mgmt_Sub.Form
8830      strFilter = vbNullString
8840      Select Case .ckgVersion_chkAcc2016
          Case True
8850        .ckgVersion_chkAcc2016_lbl.FontBold = True
            ' ** Uses either qryFormLoad_03 or qryFormLoad_07.
8860        Select Case .chkLatestOnly
            Case True
8870          If frm.RecordSource <> "qryFormLoad_07" Then
                ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_06
                ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                ' ** xadload_access_version, with Max(xadload_datemodified)).
8880            frm.RecordSource = "qryFormLoad_07"
8890          End If
8900        Case False
8910          If frm.RecordSource <> "qryFormLoad_03" Then
                ' ** tblXAdmin_Load, all records.
8920            frm.RecordSource = "qryFormLoad_03"
8930          End If
8940        End Select
8950        DoEvents
8960        If .ckgVersion_chkAll = True Then
8970          .ckgVersion_chkAll = False
8980          .ckgVersion_chkAll_lbl.FontBold = False
8990        End If
9000        If .ckgVersion_chkAcc2003 = True Then
9010          strFilter = "[xadload_access_version] = 'Access 2003'"
9020        End If
9030        If .ckgVersion_chkAcc2007 = True Then
9040          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
9050          strFilter = strFilter & "[xadload_access_version] = 'Access 2007'"
9060        End If
9070        If .ckgVersion_chkAcc2010 = True Then
9080          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
9090          strFilter = strFilter & "[xadload_access_version] = 'Access 2010'"
9100        End If
9110        If .ckgVersion_chkAcc2013 = True Then
9120          strFilter = strFilter & " Or "
9130          strFilter = strFilter & "[xadload_access_version] = 'Access 2013'"
9140        End If
9150        If strFilter <> vbNullString Then strFilter = strFilter & " Or "
9160        strFilter = strFilter & "[xadload_access_version] = 'Access 2016'"
9170        .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
9180        .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
9190        .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
9200      Case False
9210        .ckgVersion_chkAcc2016_lbl.FontBold = False
            ' ** Uses qryFormLoad_03 or qryFormLoad_05 Or qryFormLoad_07.
9220        Select Case .chkLatestOnly
            Case True
9230          Select Case .ckgVersion_chkAll
              Case True
                ' ** If all is true, qryFormLoad_05.
9240            If frm.RecordSource <> "qryFormLoad_05" Then
                  ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_04
                  ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                  ' ** with Max(xadload_datemodified)).
9250              frm.RecordSource = "qryFormLoad_05"
9260            End If
9270          Case False
9280            If .ckgVersion_chkAcc2007 = True Or .ckgVersion_chkAcc2010 = True Or _
                    .ckgVersion_chkAcc2013 = True Or .ckgVersion_chkAcc2016 = True Then
                  ' ** If other versions are on, qryFormLoad_07.
9290              If frm.RecordSource <> "qryFormLoad_07" Then
                    ' ** qryFormLoad_03 (tblXAdmin_Load, all records), linked to qryFormLoad_06
                    ' ** (qryFormLoad_03 (tblXAdmin_Load, all records), grouped by forms,
                    ' ** xadload_access_version, with Max(xadload_datemodified)).
9300                frm.RecordSource = "qryFormLoad_07"
9310              End If
9320            Else
                  ' ** If no versions are on, qryFormLoad_03 {arbitrary}.
9330              If frm.RecordSource <> "qryFormLoad_03" Then
                    ' ** tblXAdmin_Load, all records.
9340                frm.RecordSource = "qryFormLoad_03"
9350              End If
9360            End If
9370          End Select
9380        Case False
              ' ** If all is true, qryFormLoad_03.
              ' ** If other versions are on, qryFormLoad_03.
              ' ** If no versions are on, qryFormLoad_03 {arbitrary}.
9390          If frm.RecordSource <> "qryFormLoad_03" Then
                ' ** tblXAdmin_Load, all records.
9400            frm.RecordSource = "qryFormLoad_03"
9410          End If
9420        End Select
9430        DoEvents
9440        If .ckgVersion_chkAcc2003 = True Then
9450          strFilter = "[xadload_access_version] = 'Access 2003'"
9460        End If
9470        If .ckgVersion_chkAcc2007 = True Then
9480          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
9490          strFilter = strFilter & "[xadload_access_version] = 'Access 2007'"
9500        End If
9510        If .ckgVersion_chkAcc2010 = True Then
9520          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
9530          strFilter = strFilter & "[xadload_access_version] = 'Access 2010'"
9540        End If
9550        If .ckgVersion_chkAcc2013 = True Then
9560          If strFilter <> vbNullString Then strFilter = strFilter & " Or "
9570          strFilter = strFilter & "[xadload_access_version] = 'Access 2013'"
9580        End If
9590        If strFilter <> vbNullString Then
9600          If .ckgVersion_chkAll = True Then  ' ** It shouldn't be, anyway.
9610            .ckgVersion_chkAll = False
9620            .ckgVersion_chkAll_lbl.FontBold = False
9630          End If
9640          .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
9650          .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
9660          .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
9670        Else
9680          Select Case .ckgVersion_chkAll
              Case True
9690            .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter  ' ** Null String.
9700            .frmLoadTimer_Mgmt_Sub.Form.FilterOn = False
9710          Case False
9720            strFilter = "[xadload_access_version] = 'Access 0000'"  ' ** Intended to return no records.
9730            .frmLoadTimer_Mgmt_Sub.Form.Filter = strFilter
9740            .frmLoadTimer_Mgmt_Sub.Form.FilterOn = True
9750          End Select
9760          .frmLoadTimer_Mgmt_Sub.Form.SortNow "Form_Load"
9770        End If
9780      End Select
9790      lngRecsCur = frm.RecCnt  ' ** Form Function: frmLoadTimer_Mgmt_Sub.
9800      If lngRecsCur > 0& Then
9810        .frmLoadTimer_Mgmt_Sub.SetFocus
9820      Else
9830        .FocusHolder.SetFocus
9840      End If
9850    End With

        '.ckgVersion_chkAll
        '.ckgVersion_chkAcc2003
        '.ckgVersion_chkAcc2007
        '.ckgVersion_chkAcc2010
        '.ckgVersion_chkAcc2013
        '.ckgVersion_chkAcc2016

EXITP:
9860    Set frm = Nothing
9870    Exit Sub

ERRH:
9880    Select Case ERR.Number
        Case Else
9890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9900    End Select
9910    Resume EXITP

End Sub
