VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCountry_Currency_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCountry_Currency_Sub"

'VGC 03/23/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmCountry_Currency}
' **   Currency:         N {curr_name}
' **   Exit:             X {cmdClose on frmCountry_Currency}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[curr_name]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmCountry_Currency" Then

160         blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
170         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

180         CLR_DISABLED_FG = CLR_DKGRY
190         CLR_DISABLED_BG = CLR_LTTEAL

200         lngSortLbl_Width = .Sort_lbl.Width

210       Else
220         Cancel = -1
230       End If
240     End With

250     If Cancel = -1 Then
260       Beep
270       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
280     End If

EXITP:
290     Exit Sub

ERRH:
300     Select Case ERR.Number
        Case Else
310       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
320     End Select
330     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       .curr_code.ForeColor = CLR_DISABLED_FG
430       .curr_code.BackColor = CLR_DISABLED_BG
440       .curr_name.ForeColor = CLR_DISABLED_FG
450       .curr_name.BackColor = CLR_DISABLED_BG
460       SortNow THIS_PROC  ' ** Procedure: Below.
470     End With

EXITP:
480     Exit Sub

ERRH:
490     Select Case ERR.Number
        Case Else
500       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
510     End Select
520     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

610     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
620     intShiftDown = (Shift And acShiftMask) > 0
630     intAltDown = (Shift And acAltMask) > 0
640     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Country:          N {country_name}

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyUp
670         intRetVal = 0
680         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
690       Case vbKeyDown
700         intRetVal = 0
710         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
720       Case vbKeyEscape
730         With Me
740           intRetVal = 0
750           .Parent.cmdClose_Click  ' ** Form Procedure: frmCountryCode.
760         End With
770       End Select
780     End If

        ' ** Alt keys.
790     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
800       Select Case intRetVal
          Case vbKeyN
810         With Me
820           intRetVal = 0
830           .curr_name.SetFocus
840         End With
850       Case vbKeyX
860         With Me
870           intRetVal = 0
880           .Parent.cmdClose_Click  ' ** Form Procedure: frmCountryCode.
890         End With
900       End Select
910     End If

        ' ** Ctrl keys.
920     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
930       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
940         With Me
950           intRetVal = 0
960           DoCmd.SelectObject acForm, .Parent.Name, False
970           .Parent.cmdClose.SetFocus
980         End With
990       Case vbKeyUp, vbKeyPageUp
1000        intRetVal = 0
1010        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1020      Case vbKeyDown, vbKeyPageDown
1030        intRetVal = 0
1040        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1050      End Select
1060    End If

        ' ** Ctrl-Shift keys.
1070    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1080      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1090        With Me
1100          intRetVal = 0
1110          DoCmd.SelectObject acForm, .Parent.Name, False
1120          .Parent.cmdClose.SetFocus
1130        End With
1140      End Select
1150    End If

EXITP:
1160    KeyCode = intRetVal
1170    Exit Sub

ERRH:
1180    intRetVal = 0
1190    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1200    Case Else
1210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1220    End Select
1230    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1320    intShiftDown = (Shift And acShiftMask) > 0
1330    intAltDown = (Shift And acAltMask) > 0
1340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1370        With Me
1380          intRetVal = 0
1390          lngRecsCur = RecCnt  ' ** Function: Below.
1400          If .CurrentRecord < lngRecsCur Then
1410            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1420            .FocusHolder.SetFocus
1430          Else
1440            DoCmd.SelectObject acForm, .Parent.Name, False
1450            .Parent.cmdClose.SetFocus
1460          End If
1470        End With
1480      End Select
1490    End If

        ' ** Shift keys.
1500    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1510      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1520        With Me
1530          intRetVal = 0
1540          If .CurrentRecord > 1 Then
1550            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1560            .FocusHolder.SetFocus
1570          Else
1580            DoCmd.SelectObject acForm, .Parent.Name, False
1590            .Parent.cmdClose.SetFocus
1600          End If
1610        End With
1620      End Select
1630    End If

EXITP:
1640    KeyCode = intRetVal
1650    Exit Sub

ERRH:
1660    intRetVal = 0
1670    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1680    Case Else
1690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1700    End Select
1710    Resume EXITP

End Sub

Private Sub curr_code_lbl_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_lbl_DblClick"

1810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub curr_code_GotFocus()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_GotFocus"

1910    With Me
1920      .curr_code.SelLength = 0
1930      .curr_code.SelStart = 9
1940    End With

EXITP:
1950    Exit Sub

ERRH:
1960    Select Case ERR.Number
        Case Else
1970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1980    End Select
1990    Resume EXITP

End Sub

Private Sub curr_code_KeyDown(KeyCode As Integer, Shift As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2020    intShiftDown = (Shift And acShiftMask) > 0
2030    intAltDown = (Shift And acAltMask) > 0
2040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2070        With Me
2080          intRetVal = 0
2090          .curr_name.SetFocus
2100        End With
2110      End Select
2120    End If

        ' ** Shift keys.
2130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2150        With Me
2160          intRetVal = 0
2170          If .CurrentRecord > 1 Then
2180            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2190            .FocusHolder.SetFocus
2200          Else
2210            DoCmd.SelectObject acForm, .Parent.Name, False
2220            .Parent.cmdClose.SetFocus
2230          End If
2240        End With
2250      End Select
2260    End If

EXITP:
2270    KeyCode = intRetVal
2280    Exit Sub

ERRH:
2290    intRetVal = 0
2300    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2310    Case Else
2320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2330    End Select
2340    Resume EXITP

End Sub

Private Sub curr_name_lbl_DblClick(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_name_lbl_DblClick"

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

Private Sub curr_name_GotFocus()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_name_GotFocus"

2510    With Me
2520      .curr_name.SelLength = 0
2530      .curr_name.SelStart = 99
2540    End With

EXITP:
2550    Exit Sub

ERRH:
2560    Select Case ERR.Number
        Case Else
2570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2580    End Select
2590    Resume EXITP

End Sub

Private Sub curr_name_KeyDown(KeyCode As Integer, Shift As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_name_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
2670        With Me
2680          intRetVal = 0
2690          lngRecsCur = RecCnt  ' ** Function: Below.
2700          If .CurrentRecord < lngRecsCur Then
2710            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2720            .FocusHolder.SetFocus
2730          Else
2740            DoCmd.SelectObject acForm, .Parent.Name, False
2750            .Parent.cmdClose.SetFocus
2760          End If
2770        End With
2780      End Select
2790    End If

        ' ** Shift keys.
2800    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2810      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2820        With Me
2830          intRetVal = 0
2840          .curr_name.SetFocus
2850        End With
2860      End Select
2870    End If

EXITP:
2880    KeyCode = intRetVal
2890    Exit Sub

ERRH:
2900    intRetVal = 0
2910    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2920    Case Else
2930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2940    End Select
2950    Resume EXITP

End Sub

Public Function RecCnt() As Long

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

3010    lngRetVal = 0&

3020    With Me
3030      Set rst = .RecordsetClone
3040      With rst
3050        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
3060        Else
3070          .MoveLast
3080          lngRetVal = .RecordCount
3090        End If
3100        .Close
3110      End With
3120    End With

EXITP:
3130    Set rst = Nothing
3140    RecCnt = lngRetVal
3150    Exit Function

ERRH:
3160    lngRetVal = 0&
3170    Select Case ERR.Number
        Case Else
3180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3190    End Select
3200    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

3310    Select Case intWhere
        Case acCmdRecordsGoToFirst
3320      DoCmd.RunCommand acCmdRecordsGoToFirst
3330    Case acCmdRecordsGoToLast
3340      DoCmd.RunCommand acCmdRecordsGoToLast
3350    Case acCmdRecordsGoToPrevious
3360      DoCmd.RunCommand acCmdRecordsGoToPrevious
3370    Case acCmdRecordsGoToNext
3380      DoCmd.RunCommand acCmdRecordsGoToNext
3390    Case Else
3400      If IsMissing(varID) = False Then
3410        With Me
3420          Set rst = .RecordsetClone
3430          With rst
3440            .FindFirst "[curr_id] = " & varID
3450            If .NoMatch = False Then
3460              Me.Bookmark = .Bookmark
3470            End If
3480            .Close
3490          End With
3500        End With
3510      End If
3520    End Select

EXITP:
3530    Set rst = Nothing
3540    Exit Sub

ERRH:
3550    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
3560    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
3570    Case Else
3580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3590    End Select
3600    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim ctl As Access.Control
        Dim strCalled As String, strSortAsc As String

        'Const strStdAsc As String = ", [curr_name]"

3710    With Me
3720      .Controls(strSortLbl).Visible = False
3730      .Controls(strSortLine).Visible = False
3740      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
3750      If strProc = "Form_Load" Then
3760        strCalled = "curr_name"
3770        strSortNow = strSortOrig
3780        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
3790        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
3800        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
3810        .Controls(strSortLbl).Left = lngSortLbl_Left
3820        .Controls(strSortLbl).Caption = strArwUp
3830        .Controls(strSortLbl).ForeColor = CLR_DKBLU
3840        .Controls(strSortLine).Left = lngSortLine_Left
3850        .Controls(strSortLine).Width = lngSortLine_Width
3860      Else
3870        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
3880        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
3890        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
3900        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
3910        .Controls(strSortLbl).Left = lngSortLbl_Left
3920        .Controls(strSortLine).Left = lngSortLine_Left
3930        .Controls(strSortLine).Width = lngSortLine_Width
3940        If strCalled = "curr_name" Then
3950          strSortAsc = strSortOrig
3960          If strSortNow = strSortAsc Then
3970            strSortNow = "[curr_name] DESC"
3980            .Controls(strSortLbl).Caption = strArwDn
3990            .Controls(strSortLbl).ForeColor = CLR_DKRED
4000          Else
4010            strSortNow = strSortAsc
4020            .Controls(strSortLbl).Caption = strArwUp
4030            .Controls(strSortLbl).ForeColor = CLR_DKBLU
4040          End If
4050        Else
4060          .Controls(strSortLbl).Caption = strArwUp
4070          .Controls(strSortLbl).ForeColor = CLR_DKBLU
4080          Select Case strCalled
              Case "curr_code"
4090            strSortAsc = "[curr_code]"
4100            lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
4110            .Controls(strSortLbl).Left = lngSortLbl_Left
4120          End Select
4130          If strSortNow = strSortAsc Then
4140            If InStr(strSortAsc, ",") > 0 Then
4150              strSortNow = Left(strSortAsc, (InStr(strSortAsc, ",") - 1)) & " DESC" & Mid(strSortAsc, InStr(strSortAsc, ",")) & " DESC"
4160            Else
4170              strSortNow = strSortNow & " DESC"
4180            End If
4190            .Controls(strSortLbl).Caption = strArwDn
4200            .Controls(strSortLbl).ForeColor = CLR_DKRED
4210          Else
4220            strSortNow = strSortAsc
4230          End If
4240        End If
4250      End If
4260      .Controls(strSortLbl).Visible = True
4270      .Controls(strSortLine).Visible = True
4280      .OrderBy = strSortNow
4290      .OrderByOn = True
4300    End With

EXITP:
4310    Set ctl = Nothing
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub
