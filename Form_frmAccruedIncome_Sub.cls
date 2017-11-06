VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccruedIncome_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccruedIncome_Sub"

'VGC 12/30/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Tax Codes Purch:  A {cmbTaxCodes_Purchase on frmAccruedIncome}
' **   Close:            C {cmdClose on frmAccruedIncome}
' **   Posting Date:     D {AccrualDate on frmAccruedIncome}
' **   I/E Codes Purch:  E {cmbRevenueCodes_Purchase on frmAccruedIncome}
' **   Copy First:       F {cmdCopyFactor on frmAccruedIncome}
' **   Generate Dates:   G {cmdGenerateDates on frmAccruedIncome}
' **   Create Jrnl Ent:  J {cmdCreateEntries on frmAccruedIncome}
' **   Daily Factor:     L {tmpRate}
' **   Month:            M {AccrualMonth on frmAccruedIncome}
' **   I/E Codes Int:    N {cmbRevenueCodes_Interest on frmAccruedIncome}
' **   Comments:         O {ReinvestComments on frmAccruedIncome}
' **   Reinvest:         R {chkReinvest on frmAccruedIncome}
' **   Per Share:        S {ReinvestPerShare on frmAccruedIncome}
' **   Tax Codes Int:    T {cmbTaxCodes_Interest on frmAccruedIncome}
' **   Exit:             X {cmdClose on frmAccruedIncome}
' **   Year:             Y {AccrualYear on frmAccruedIncome}

' ** Shortcut F-Keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar on frmAccruedIncome}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[tmpDate]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine1 As String = "Sort_line1"
Private Const strSortLine2 As String = "Sort_line2"
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

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAccruedIncome" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190         lngSortLbl_Width = .Sort_lbl.Width

200       Else
210         Cancel = -1
220       End If
230     End With

240     If Cancel = -1 Then
250       Beep
260       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
270     End If

EXITP:
280     Exit Sub

ERRH:
290     Select Case ERR.Number
        Case Else
300       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
310     End Select
320     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me

420       .tmpDate.ForeColor = CLR_DISABLED_FG
430       .tmpDate.BackColor = CLR_DISABLED_BG

440       SortNow THIS_PROC  ' ** Procedure: Below.

450     End With

EXITP:
460     Exit Sub

ERRH:
470     Select Case ERR.Number
        Case Else
480       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
490     End Select
500     Resume EXITP

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

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar on frmAccruedIncome}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Daily Factor:     A {tmpRate}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyF7
670         With Me
680           intRetVal = 0
690           DoCmd.SelectObject acForm, .Parent.Name, False
700           .Parent.cmdCalendar.SetFocus
710           .Parent.cmdCalendar_Click  ' ** Form Procedure: frmAccruedIncome.
720         End With
730       Case vbKeyUp
740         intRetVal = 0
750         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
760       Case vbKeyDown
770         intRetVal = 0
780         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
790       Case vbKeyEscape
800         With Me
810           intRetVal = 0
820           .Parent.cmdClose_Click  ' ** Form Procedure: frmAccruedIncome.
830         End With
840       End Select
850     End If

        ' ** Alt keys.
860     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
870       Select Case intRetVal
          Case vbKeyL
880         With Me
890           intRetVal = 0
900           If .tmpDate_lbl.Visible = True Then
910             .tmpRate.SetFocus
920           Else
930             Beep
940           End If
950         End With
960       Case vbKeyX
970         With Me
980           intRetVal = 0
990           .Parent.cmdClose_Click  ' ** Form Procedure: frmAccruedIncome.
1000        End With
1010      End Select
1020    End If

        ' ** Ctrl keys.
1030    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1040      Select Case intRetVal
          Case vbKeyS
1050        intRetVal = 0
1060        cmdSave_Click  ' ** Procedure: Below.
1070      Case vbKeyTab, vbKeyReturn
1080        With Me
1090          intRetVal = 0
1100          DoCmd.SelectObject acForm, .Parent.Name, False
1110          If .Parent.cmdCopyFactor.Enabled = True Then
1120            .Parent.cmdCopyFactor.SetFocus
1130          ElseIf .Parent.AccrualDate.Enabled = True Then
1140            .Parent.AccrualDate.SetFocus
1150          ElseIf .Parent.chkReinvest.Enabled = True Then
1160            .Parent.chkReinvest.SetFocus
1170          Else
1180            .Parent.cmdClose.SetFocus
1190          End If
1200        End With
1210      Case vbKeyPageUp
1220        intRetVal = 0
1230        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1240      Case vbKeyPageDown
1250        intRetVal = 0
1260        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1270      End Select
1280    End If

        ' ** Ctrl-Shift keys.
1290    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1300      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1310        With Me
1320          intRetVal = 0
1330          DoCmd.SelectObject acForm, .Parent.Name, False
1340          .Parent.cmdClose.SetFocus
1350        End With
1360      End Select
1370    End If

EXITP:
1380    KeyCode = intRetVal
1390    Exit Sub

ERRH:
1400    intRetVal = 0
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Public Sub cmdSave_Click()
' ** This is a fake click event.

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

1510    DoCmd.RunCommand acCmdSaveRecord

EXITP:
1520    Exit Sub

ERRH:
1530    Select Case ERR.Number
        Case Else
1540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1550    End Select
1560    Resume EXITP

End Sub

Private Sub tmpDate_lbl_DblClick(Cancel As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "tmpDate_lbl_DblClick"

1610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1620    Exit Sub

ERRH:
1630    Select Case ERR.Number
        Case Else
1640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1650    End Select
1660    Resume EXITP

End Sub

Private Sub tmpRate_lbl_DblClick(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "tmpRate_lbl_DblClick"

1710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub

Private Sub tmpRate_GotFocus()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "tmpRate_GotFocus"

1810    With Me
1820      .tmpRate.SelLength = 0
1830      .tmpRate.SelStart = 0
1840    End With

EXITP:
1850    Exit Sub

ERRH:
1860    Select Case ERR.Number
        Case Else
1870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1880    End Select
1890    Resume EXITP

End Sub

Private Sub tmpRate_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "tmpRate_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
1970        With Me
1980          intRetVal = 0
1990          lngRecsCur = RecCnt  ' ** Function: Below.
2000          If .CurrentRecord < lngRecsCur Then
2010            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2020            .tmpRate.SetFocus
2030          Else
2040            DoCmd.SelectObject acForm, .Parent.Name, False
2050            .Parent.AccrualDate.SetFocus
2060          End If
2070        End With
2080      End Select
2090    End If

        ' ** Shift keys.
2100    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2110      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2120        With Me
2130          intRetVal = 0
2140          If .CurrentRecord > 1 Then
2150            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2160            .tmpRate.SetFocus
2170          Else
2180            DoCmd.SelectObject acForm, .Parent.Name, False
2190            .Parent.cmdCopyFactor.SetFocus
2200          End If
2210        End With
2220      End Select
2230    End If

EXITP:
2240    KeyCode = intRetVal
2250    Exit Sub

ERRH:
2260    intRetVal = 0
2270    Select Case ERR.Number
        Case Else
2280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2290    End Select
2300    Resume EXITP

End Sub

Public Function RecCnt() As Long

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

2410    lngRetVal = 0&

2420    With Me
2430      Set rst = .RecordsetClone
2440      With rst
2450        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
2460        Else
2470          .MoveLast
2480          lngRetVal = .RecordCount
2490        End If
2500        .Close
2510      End With
2520    End With

EXITP:
2530    Set rst = Nothing
2540    RecCnt = lngRetVal
2550    Exit Function

ERRH:
2560    lngRetVal = 0&
2570    Select Case ERR.Number
        Case Else
2580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2590    End Select
2600    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

2710    Select Case intWhere
        Case acCmdRecordsGoToFirst
2720      DoCmd.RunCommand acCmdRecordsGoToFirst
2730    Case acCmdRecordsGoToLast
2740      DoCmd.RunCommand acCmdRecordsGoToLast
2750    Case acCmdRecordsGoToPrevious
2760      DoCmd.RunCommand acCmdRecordsGoToPrevious
2770    Case acCmdRecordsGoToNext
2780      DoCmd.RunCommand acCmdRecordsGoToNext
2790    Case Else
2800      If IsMissing(varID) = False Then
2810        With Me
2820          Set rst = .RecordsetClone
2830          With rst
2840            .FindFirst "[tmpDate] = #" & CStr(varID) & "#"
2850            If .NoMatch = False Then
2860              Me.Bookmark = .Bookmark
2870            End If
2880            .Close
2890          End With
2900        End With
2910      End If
2920    End Select

EXITP:
2930    Set rst = Nothing
2940    Exit Sub

ERRH:
2950    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
2960    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
2970    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
2980    Case Else
2990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3000    End Select
3010    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [tmpDate]"

3110    With Me
3120      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
3130        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
3140      End If
3150      .Controls(strSortLbl).Visible = False
3160      .Controls(strSortLine1).Visible = False
3170      .Controls(strSortLine2).Visible = False
3180      .Controls(strSortLine1).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
3190      .Controls(strSortLine2).Width = lngTpp
3200      If strProc = "Form_Load" Then
3210        strCalled = "tmpDate"
3220        strSortNow = strSortOrig
3230        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
3240        lngSortLine_Left = .tmpDate_lbl_line.Left
3250        lngSortLine_Width = .tmpDate_lbl_line.Width
3260        .Controls(strSortLbl).Left = lngSortLbl_Left
3270        .Controls(strSortLine1).Left = lngSortLine_Left
3280        .Controls(strSortLine2).Left = lngSortLine_Left
3290        .Controls(strSortLine1).Width = lngSortLine_Width
3300        .Controls(strSortLine2).Width = lngSortLine_Width
3310        .Controls(strSortLbl).Caption = strArwUp
3320        .Controls(strSortLbl).ForeColor = CLR_DKBLU
3330      Else
3340        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
3350        lngSortLbl_Left = (.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width
3360        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
3370        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
3380        .Controls(strSortLbl).Left = lngSortLbl_Left
3390        .Controls(strSortLine1).Left = lngSortLine_Left
3400        .Controls(strSortLine2).Left = lngSortLine_Left
3410        .Controls(strSortLine1).Width = lngSortLine_Width
3420        .Controls(strSortLine2).Width = lngSortLine_Width
3430        If strCalled = "tmpDate" Then
3440          strSortAsc = strSortOrig
3450          If strSortNow = strSortAsc Then
3460            strSortNow = "[tmpDate] DESC"
3470            lngSortLine_Left = .tmpDate_lbl_line.Left
3480            lngSortLine_Width = .tmpDate_lbl_line.Width
3490            .Controls(strSortLine1).Left = lngSortLine_Left
3500            .Controls(strSortLine2).Left = lngSortLine_Left
3510            .Controls(strSortLine1).Width = lngSortLine_Width
3520            .Controls(strSortLine2).Width = lngSortLine_Width
3530            .Controls(strSortLbl).Caption = strArwDn
3540            .Controls(strSortLbl).ForeColor = CLR_DKRED
3550          Else
3560            strSortNow = strSortAsc
3570            .Controls(strSortLbl).Caption = strArwUp
3580            .Controls(strSortLbl).ForeColor = CLR_DKBLU
3590          End If
3600        Else
3610          .Controls(strSortLbl).Caption = strArwUp
3620          .Controls(strSortLbl).ForeColor = CLR_DKBLU
3630          Select Case strCalled
              Case "tmpRate"
3640            strSortAsc = "[tmpRate]" & strStdAsc
3650          End Select
3660          If strSortNow = strSortAsc Then
3670            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
3680            Select Case intCnt
                Case 1
3690              strTmp01 = strSortAsc & " DESC"
3700            Case 2
3710              intPos01 = InStr(strSortAsc, ",")
3720              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
3730              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
3740            Case 3
3750              intPos01 = InStr(strSortAsc, ",")
3760              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
3770              strTmp02 = Mid(strSortAsc, intPos01)
3780              intPos01 = InStr(2, strTmp02, ",")
3790              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
3800              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
3810            End Select
3820            strSortNow = strTmp01
3830            .Controls(strSortLbl).Caption = strArwDn
3840            .Controls(strSortLbl).ForeColor = CLR_DKRED
3850          Else
3860            strSortNow = strSortAsc
3870          End If
3880        End If
3890      End If
3900      .Controls(strSortLbl).Visible = True
3910      .Controls(strSortLine1).Visible = True
3920      .Controls(strSortLine2).Visible = True
3930      .OrderBy = strSortNow
3940      .OrderByOn = True
3950    End With

EXITP:
3960    Exit Sub

ERRH:
3970    Select Case ERR.Number
        Case Else
3980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3990    End Select
4000    Resume EXITP

End Sub
