VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Checks_Bank2_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Checks_Bank2_Sub"

'VGC 07/24/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Account Number:   A {Bank_AccountNumber}
' **   Copy Bank Info:   C {cmdCopy on frmRpt_Checks_Bank2}
' **   Bank Name:        N {Bank_Name}
' **   OK:               O {cmdOK on frmRpt_Checks_Bank2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[alphasort]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private strSortNow As String, lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngTpp As Long, lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmRpt_Checks_Bank2" Then

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

420       .accountno.ForeColor = CLR_DISABLED_FG
430       .accountno.BackColor = CLR_DISABLED_BG
440       .shortname.ForeColor = CLR_DISABLED_FG
450       .shortname.BackColor = CLR_DISABLED_BG

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

Private Sub Form_Current()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim blnAble As Boolean

610     With Me
620       blnAble = False
630       If IsNull(.Bank_Name) = False And IsNull(.Bank_AccountNumber) = False Then  ' ** Bank info is present.
640         .Parent.cmdCopy.Enabled = True
650         blnAble = True
660       Else
670         .Parent.cmdCopy.Enabled = False
680       End If
690       Select Case blnAble
          Case True
700         .Parent.cmdCopy_raised_img.Visible = True
710         .Parent.cmdCopy_raised_semifocus_dots_img.Visible = False
720         .Parent.cmdCopy_raised_focus_img.Visible = False
730         .Parent.cmdCopy_raised_focus_dots_img.Visible = False
740         .Parent.cmdCopy_sunken_focus_dots_img.Visible = False
750         .Parent.cmdCopy_raised_img_dis.Visible = False
760       Case False
770         .Parent.cmdCopy_raised_img_dis.Visible = True
780         .Parent.cmdCopy_raised_img.Visible = False
790         .Parent.cmdCopy_raised_semifocus_dots_img.Visible = False
800         .Parent.cmdCopy_raised_focus_img.Visible = False
810         .Parent.cmdCopy_raised_focus_dots_img.Visible = False
820         .Parent.cmdCopy_sunken_focus_dots_img.Visible = False
830       End Select
840       DoEvents
850     End With

EXITP:
860     Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1020    intShiftDown = (Shift And acShiftMask) > 0
1030    intAltDown = (Shift And acAltMask) > 0
1040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Account Number:   A {Bank_AccountNumber}
        ' **   Bank Name:        N {Bank_Name}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1060      Select Case intRetVal
          Case vbKeyUp
1070        intRetVal = 0
1080        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1090      Case vbKeyDown
1100        intRetVal = 0
1110        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1120      End Select
1130    End If

        ' ** Alt keys.
1140    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1150      Select Case intRetVal
          Case vbKeyA
1160        With Me
1170          intRetVal = 0
1180          .Bank_AccountNumber.SetFocus
1190        End With
1200      Case vbKeyN
1210        With Me
1220          intRetVal = 0
1230          .Bank_Name.SetFocus
1240        End With
1250      End Select
1260    End If

        ' ** Ctrl keys.
1270    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1280      Select Case intRetVal
          Case vbKeyS
1290        intRetVal = 0
1300        cmdSave_Click  ' ** Procedure: Below.
1310      Case vbKeyTab, vbKeyReturn
1320        With Me
1330          intRetVal = 0
1340          DoCmd.SelectObject acForm, .Parent.Name, False
1350          If .Parent.cmdCopy.Enabled = True Then
1360            .Parent.cmdCopy.SetFocus
1370          Else
1380            .Parent.cmdOK.SetFocus
1390          End If
1400        End With
1410      Case vbKeyPageUp, vbKeyHome
1420        intRetVal = 0
1430        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1440      Case vbKeyPageDown, vbKeyEnd
1450        intRetVal = 0
1460        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1470      End Select
1480    End If

        ' ** Ctrl-Shift keys.
1490    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1500      Select Case intRetVal
          Case vbKeyF
1510        With Me
1520          intRetVal = 0
1530          DoCmd.SelectObject acForm, .Parent.Name, False
1540          .Parent.FocusHolder.SetFocus
1550        End With
1560      Case vbKeyTab, vbKeyReturn
1570        With Me
1580          DoCmd.SelectObject acForm, .Parent.Name, False
1590          .Parent.cmdOK.SetFocus
1600        End With
1610      End Select
1620    End If

EXITP:
1630    KeyCode = intRetVal
1640    Exit Sub

ERRH:
1650    intRetVal = 0
1660    Select Case ERR.Number
        Case Else
1670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1680    End Select
1690    Resume EXITP

End Sub

Public Sub cmdSave_Click()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

1710    DoCmd.RunCommand acCmdSaveRecord

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case 3022  ' ** The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.
          ' ** Ignore! I don't know why it happens on some machines and not mine.
1740    Case Else
1750      DoCmd.Hourglass False
1760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1770    End Select
1780    Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

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

Private Sub shortname_lbl_DblClick(Cancel As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

1910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1920    Exit Sub

ERRH:
1930    Select Case ERR.Number
        Case Else
1940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1950    End Select
1960    Resume EXITP

End Sub

Private Sub Bank_Name_lbl_DblClick(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_Name_lbl_DblClick"

2010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2020    Exit Sub

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2050    End Select
2060    Resume EXITP

End Sub

Private Sub Bank_Name_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_Name_KeyDown"

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
          Case vbKeyTab
2170        With Me
2180          intRetVal = 0
2190          .Bank_AccountNumber.SetFocus
2200        End With
2210      End Select
2220    End If

        ' ** Shift keys.
2230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2240      Select Case intRetVal
          Case vbKeyTab
2250        With Me
2260          intRetVal = 0
2270          If .CurrentRecord > 1 Then
2280            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2290            .Bank_AccountNumber.SetFocus
2300          Else
2310            DoCmd.SelectObject acForm, .Parent.Name, False
2320            .Parent.cmdOK.SetFocus
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
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2410    Case Else
2420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2430    End Select
2440    Resume EXITP

End Sub

Private Sub Bank_AccountNumber_lbl_DblClick(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_AccountNumber_lbl_DblClick"

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

Private Sub Bank_AccountNumber_KeyDown(KeyCode As Integer, Shift As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_AccountNumber_KeyDown"

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
2690          lngRecsCur = RecCnt  ' ** Function: Below.
2700          If .CurrentRecord < lngRecsCur Then
2710            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2720            .Bank_Name.SetFocus
2730          Else
2740            DoCmd.SelectObject acForm, .Parent.Name, False
2750            If .Parent.cmdCopy.Enabled = True Then
2760              .Parent.cmdCopy.SetFocus
2770            Else
2780              .Parent.cmdOK.SetFocus
2790            End If
2800          End If
2810        End With
2820      End Select
2830    End If

        ' ** Shift keys.
2840    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2850      Select Case intRetVal
          Case vbKeyTab
2860        With Me
2870          intRetVal = 0
2880          .Bank_Name.SetFocus
2890        End With
2900      End Select
2910    End If

EXITP:
2920    KeyCode = intRetVal
2930    Exit Sub

ERRH:
2940    intRetVal = 0
2950    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2960    Case Else
2970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2980    End Select
2990    Resume EXITP

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
3440            .FindFirst "[accountno] = '" & varID & "'"
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
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
3560    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
3570    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
3580    Case Else
3590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3600    End Select
3610    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [alphasort]"

3710    With Me
3720      .Controls(strSortLbl).Visible = False
3730      .Controls(strSortLine).Visible = False
3740      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form
3750      If strProc = "Form_Load" Then
3760        strCalled = "accountno"
3770        strSortNow = strSortOrig
3780        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
3790        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
3800        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
3810        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
3820        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
3830        .Controls(strSortLbl).Left = lngSortLbl_Left
3840        .Controls(strSortLbl).Top = lngSortLbl_Top
3850        .Controls(strSortLine).Top = lngSortLine_Top
3860        .Controls(strSortLine).Left = lngSortLine_Left
3870        .Controls(strSortLine).Width = lngSortLine_Width
3880        .Controls(strSortLbl).Caption = strArwUp
3890        .Controls(strSortLbl).ForeColor = CLR_DKBLU
3900      Else
3910        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
3920        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
3930        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
3940        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
3950        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
3960        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
3970        .Controls(strSortLbl).Left = lngSortLbl_Left
3980        .Controls(strSortLbl).Top = lngSortLbl_Top
3990        .Controls(strSortLine).Top = lngSortLine_Top
4000        .Controls(strSortLine).Left = lngSortLine_Left
4010        .Controls(strSortLine).Width = lngSortLine_Width
4020        If strCalled = "accountno" Then
4030          strSortAsc = strSortOrig
4040          If strSortNow = strSortAsc Then
4050            strSortNow = "[alphasort] DESC"
4060            .Controls(strSortLbl).Caption = strArwDn
4070            .Controls(strSortLbl).ForeColor = CLR_DKRED
4080          Else
4090            strSortNow = strSortAsc
4100            .Controls(strSortLbl).Caption = strArwUp
4110            .Controls(strSortLbl).ForeColor = CLR_DKBLU
4120          End If
4130        Else
4140          .Controls(strSortLbl).Caption = strArwUp
4150          .Controls(strSortLbl).ForeColor = CLR_DKBLU
4160          Select Case strCalled
              Case "shortname"
4170            strSortAsc = "[shortname]"
4180          Case "Bank_Name"
4190            strSortAsc = "[Bank_Name]" & strStdAsc
4200          Case "Bank_AccountNumber"
4210            strSortAsc = "[Bank_AccountNumber]" & strStdAsc
4220          End Select
4230          If strSortNow = strSortAsc Then
4240            intCnt = CharCnt(strSortAsc, ",") + 1
4250            Select Case intCnt
                Case 1
4260              strTmp01 = strSortAsc & " DESC"
4270            Case 2
4280              intPos01 = InStr(strSortAsc, ",")
4290              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
4300              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
4310            Case 3
4320              intPos01 = InStr(strSortAsc, ",")
4330              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
4340              strTmp02 = Mid(strSortAsc, intPos01)
4350              intPos01 = InStr(2, strTmp02, ",")
4360              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
4370              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
4380            End Select
4390            strSortNow = strTmp01
4400            .Controls(strSortLbl).Caption = strArwDn
4410            .Controls(strSortLbl).ForeColor = CLR_DKRED
4420          Else
4430            strSortNow = strSortAsc
4440          End If
4450        End If
4460      End If
4470      .Controls(strSortLbl).Visible = True
4480      .Controls(strSortLine).Visible = True
4490      .OrderBy = strSortNow
4500      .OrderByOn = True
4510      .TimerInterval = 500&
4520    End With

EXITP:
4530    Exit Sub

ERRH:
4540    Select Case ERR.Number
        Case Else
4550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4560    End Select
4570    Resume EXITP

End Sub
