VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPortfolioModeling_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmPortfolioModeling_Sub"

'VGC 10/19/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Account Select:   A {opgReturn_optSelect on frmPortfolioModeling}
' **   Close:            C {cmdClose on frmPortfolioModeling}
' **   Model:            M {Model_input}
' **   Invest Objective: O {cmbInvestmentObjectives on frmPortfolioModeling}
' **   Print Report:     P {cmdPrintReport on frmPortfolioModeling}
' **   Asset Menu:       S {opgReturn_optMenu on frmPortfolioModeling}
' **   Exit:             X {cmdClose on frmPortfolioModeling}

' ** Shortcut F-Keys responsive from this form:
' **   Recalc:           F5 {RecalcTots}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[assettype]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private blnFromEnter As Boolean, blnChanged As Boolean
Private strAssetType_First As String, strAssetType_Last As String
Private lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmPortfolioModeling" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190         blnFromEnter = False: blnChanged = False

200         lngSortLbl_Width = .Sort_lbl.Width

210         Set dbs = CurrentDb
220         With dbs
230           Set qdf = .QueryDefs("qryAssetType_03")
240           Set rst = qdf.OpenRecordset
250           With rst
260             .MoveFirst
270             strAssetType_First = ![assettype]
280             .MoveLast
290             strAssetType_Last = ![assettype]
300             .Close
310           End With
320           .Close
330         End With

340       Else
350         Cancel = -1
360       End If
370     End With

380     If Cancel = -1 Then
390       Beep
400       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
410     End If

EXITP:
420     Set rst = Nothing
430     Set qdf = Nothing
440     Set dbs = Nothing
450     Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me

520       .assettype_description.ForeColor = CLR_DISABLED_FG
530       .assettype_description.BackColor = CLR_DISABLED_BG
540       .Current.ForeColor = CLR_DISABLED_FG
550       .Current.BackColor = CLR_DISABLED_BG
560       .Difference.ForeColor = CLR_DISABLED_FG
570       .Difference.BackColor = CLR_DISABLED_BG

580       .Parent.accountno = .accountno
590       .Parent.shortname = .shortname

600       Select Case IsNull(.investmentobj)
          Case True
610         .Parent.cmbInvestmentObjectives = Null
620       Case False
630         .Parent.cmbInvestmentObjectives = .investmentobj
640       End Select

650       SortNow THIS_PROC  ' ** Procedure: Below.

660     End With

EXITP:
670     Exit Sub

ERRH:
680     Select Case ERR.Number
        Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub Form_Current()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

810     With Me
820       RecalcTots  ' ** Procedure: Below.
830       DoEvents
840       If blnChanged = True Then
850         .TimerInterval = 500&
860       End If
870     End With

EXITP:
880     Exit Sub

ERRH:
890     Select Case ERR.Number
        Case Else
900       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
910     End Select
920     Resume EXITP

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
        ' **   Account Select:   A {opgReturn_optSelect on frmPortfolioModeling}
        ' **   Model:            M {Model_input}
        ' **   Asset Menu:       S {opgReturn_optMenu on frmPortfolioModeling}

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
1120      Case vbKeyF5
1130        intRetVal = 0
1140        RecalcTots  ' ** Procedure: Below.
1150      Case vbKeyEscape
1160        With Me
1170          intRetVal = 0
1180          .Parent.cmdClose_Click  ' ** Form Procedure: frmPortfolioModeling.
1190        End With
1200      End Select
1210    End If

        ' ** Alt keys.
1220    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1230      Select Case intRetVal
          Case vbKeyA
1240        With Me
1250          intRetVal = 0
1260          DoCmd.SelectObject acForm, .Parent.Name, False
1270          .Parent.opgReturn.SetFocus
1280          .Parent.opgReturn = .Parent.opgReturn_optSelect.OptionValue
1290          .Parent.opgReturn_AfterUpdate  ' ** Form Procedure: frmPortfolioModeling.
1300        End With
1310      Case vbKeyM
1320        With Me
1330          intRetVal = 0
1340          .Model_input.SetFocus
1350        End With
1360      Case vbKeyS
1370        With Me
1380          intRetVal = 0
1390          DoCmd.SelectObject acForm, .Parent.Name, False
1400          .Parent.opgReturn.SetFocus
1410          .Parent.opgReturn = .Parent.opgReturn_optMenu.OptionValue
1420          .Parent.opgReturn_AfterUpdate  ' ** Form Procedure: frmPortfolioModeling.
1430        End With
1440      Case vbKeyX
1450        With Me
1460          intRetVal = 0
1470          .Parent.cmdClose_Click  ' ** Form Procedure: frmPortfolioModeling.
1480        End With
1490      End Select
1500    End If

        ' ** Ctrl keys.
1510    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1520      Select Case intRetVal
          Case vbKeyS
1530        intRetVal = 0
1540        cmdSave_Click  ' ** Procedure: Below.
1550      Case vbKeyTab, vbKeyReturn
1560        With Me
1570          intRetVal = 0
1580          DoCmd.SelectObject acForm, .Parent.Name, False
1590          .Parent.opgReturn.SetFocus
1600        End With
1610      Case vbKeyPageUp, vbKeyHome
1620        intRetVal = 0
1630        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1640      Case vbKeyPageDown, vbKeyEnd
1650        intRetVal = 0
1660        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1670      End Select
1680    End If

        ' ** Ctrl-Shift keys.
1690    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1700      Select Case intRetVal
          Case vbKeyF
1710        With Me
1720          intRetVal = 0
1730          DoCmd.SelectObject acForm, .Parent.Name, False
1740          .Parent.FocusHolder.SetFocus
1750        End With
1760      Case vbKeyTab, vbKeyReturn
1770        With Me
1780          DoCmd.SelectObject acForm, .Parent.Name, False
1790          .Parent.cmbInvestmentObjectives.SetFocus
1800        End With
1810      End Select
1820    End If

EXITP:
1830    KeyCode = intRetVal
1840    Exit Sub

ERRH:
1850    intRetVal = 0
1860    Select Case ERR.Number
        Case Else
1870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1880    End Select
1890    Resume EXITP

End Sub

Private Sub Form_Timer()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

1910    With Me
1920      .TimerInterval = 0&
1930      If blnChanged = True Then
1940        blnChanged = False
1950        RecalcTots  ' ** Procedure: Below.
1960        DoEvents
1970      End If
1980    End With

EXITP:
1990    Exit Sub

ERRH:
2000    Select Case ERR.Number
        Case Else
2010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2020    End Select
2030    Resume EXITP

End Sub

Private Sub Form_AfterUpdate()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_AfterUpdate"

2110    RecalcTots  ' ** Procedure: Below.

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Public Sub cmdSave_Click()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

2210    DoCmd.RunCommand acCmdSaveRecord

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case 3022  ' ** The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.
          ' ** Ignore! I don't know why it happens on some machines and not mine.
2240    Case Else
2250      DoCmd.Hourglass False
2260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2270    End Select
2280    Resume EXITP

End Sub

Private Sub assettype_lbl_DblClick(Cancel As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_lbl_DblClick"

2310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl   ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

End Sub

Private Sub assettype_lbl2_DblClick(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_lbl2_DblClick"

2410    assettype_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
2420    Exit Sub

ERRH:
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl   ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Sub

Private Sub assettype_description_lbl_DblClick(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_description_lbl_DblClick"

2510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl   ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub Model_input_lbl_DblClick(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Model_input_lbl_DblClick"

2610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2620    Exit Sub

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl   ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub Model_input_GotFocus()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Model_input_GotFocus"

2710    With Me
2720      If blnFromEnter = True Then
2730        blnFromEnter = False
2740        .Model_input.SelLength = 0
2750        .Model_input.SelStart = 0
2760      End If
2770    End With

EXITP:
2780    Exit Sub

ERRH:
2790    Select Case ERR.Number
        Case Else
2800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl   ' ** Module Function: modErrorHandler.
2810    End Select
2820    Resume EXITP

End Sub

Private Sub Model_input_BeforeUpdate(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Model_input_BeforeUpdate"

2910    With Me
2920      If IsNull(.Model_input) = False Then
2930        If .Model_input < 0 Then
2940          Beep
2950          MsgBox "Percentages must be positive.", vbInformation + vbOKOnly, "Invalid Entry"
2960          Cancel = -1
2970        Else
2980          If .Model_input > 100 Then
2990            Beep
3000            MsgBox "Percentages must not exceed 100%.", vbInformation + vbOKOnly, "Invalid Entry"
3010            Cancel = -1
3020          End If
3030        End If
3040      End If
3050    End With

EXITP:
3060    Exit Sub

ERRH:
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub Model_input_AfterUpdate()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Model_input_AfterUpdate"

3210    With Me
3220      If IsNull(.Model_input) = False Then
3230        .Model = (.Model_input / 100)
3240      Else
3250        .Model = 0#
3260      End If
3270      cmdSave_Click  ' ** Procedure: Above.
3280      blnChanged = True
3290    End With

EXITP:
3300    Exit Sub

ERRH:
3310    Select Case ERR.Number
        Case Else
3320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3330    End Select
3340    Resume EXITP

End Sub

Private Sub Model_input_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Model_input_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3420    intShiftDown = (Shift And acShiftMask) > 0
3430    intAltDown = (Shift And acAltMask) > 0
3440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3470        With Me
3480          intRetVal = 0
3490          If .assettype = strAssetType_Last Then
3500            DoCmd.SelectObject acForm, .Parent.Name, False
3510            .Parent.opgReturn.SetFocus
3520          Else
3530            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3540          End If
3550        End With
3560      End Select
3570    End If

        ' ** Shift keys.
3580    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3590      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3600        With Me
3610          intRetVal = 0
3620          If .CurrentRecord > 1 Then
3630            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3640          Else
3650            DoCmd.SelectObject acForm, .Parent.Name, False
3660            .Parent.cmbInvestmentObjectives.SetFocus
3670          End If
3680        End With
3690      End Select
3700    End If

        ' ** Ctrl-Shift keys.
3710    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3720      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3730        With Me
3740          intRetVal = 0
3750          DoCmd.SelectObject acForm, .Parent.Name, False
3760          .Parent.cmbInvestmentObjectives.SetFocus
3770        End With
3780      End Select
3790    End If

EXITP:
3800    KeyCode = intRetVal
3810    Exit Sub

ERRH:
3820    intRetVal = 0
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub Model_input_LostFocus()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Model_input_LostFocus"

3910    With Me
3920      If IsNull(.Model_input) = True Then
3930        .Model_input = 0
3940      End If
3950      .DifferenceX = (.Current - .Model_input)
3960      .Refresh
3970      DoEvents
3980      RecalcTots  ' ** Procedure: Below.
3990    End With

EXITP:
4000    Exit Sub

ERRH:
4010    Select Case ERR.Number
        Case Else
4020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl   ' ** Module Function: modErrorHandler.
4030    End Select
4040    Resume EXITP

End Sub

Private Sub Current_lbl_DblClick(Cancel As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Current_lbl_DblClick"

4110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl   ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Private Sub Difference_lbl_DblClick(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Difference_lbl_DblClick"

4210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4220    Exit Sub

ERRH:
4230    Select Case ERR.Number
        Case Else
4240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl   ' ** Module Function: modErrorHandler.
4250    End Select
4260    Resume EXITP

End Sub

Public Sub RecalcTots()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

        Dim varTmp00 As Variant

        Static intReady As Integer

4310    With Me
4320      If intReady = 0 Then
4330        If IsNull(.ModelTot) = True Then
4340          varTmp00 = "Null"
4350        Else
4360          varTmp00 = CStr(.ModelTot)
4370          If Trim(varTmp00) = vbNullString Then
4380            varTmp00 = "NullString"
4390          End If
4400        End If
4410        If varTmp00 = "Null" Or varTmp00 = "NullString" Then
4420          .Parent.RecalcTots_NotReady True  ' ** Form Procedure: frmPortfolioModeling.
4430        Else
4440          intReady = 1
4450          .Parent.RecalcTots_NotReady False  ' ** Form Procedure: frmPortfolioModeling.
4460        End If
4470      End If
4480      .Parent.ModelTot = .ModelTot
4490      .Parent.CurrentTot = .CurrentTot
4500      .Parent.DifferenceTot = .DifferenceTot
4510      .Parent.ICash = .cash
4520      If .ModelTot > 100 Then
4530        .Parent.ModelTot.BackColor = CLR_LTRED
4540      Else
4550        .Parent.ModelTot.BackColor = CLR_DISABLED_BG
4560      End If
4570    End With

EXITP:
4580    Exit Sub

ERRH:
4590    Select Case ERR.Number
        Case Else
4600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4610    End Select
4620    Resume EXITP

End Sub

Public Sub FromEnter_Set(blnEnter As Boolean)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "FromEnter_Set"

4710    blnFromEnter = blnEnter

EXITP:
4720    Exit Sub

ERRH:
4730    Select Case ERR.Number
        Case Else
4740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4750    End Select
4760    Resume EXITP

End Sub

Public Function RecCnt() As Long

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

4810    lngRetVal = 0&

4820    With Me
4830      Set rst = .RecordsetClone
4840      With rst
4850        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
4860        Else
4870          .MoveLast
4880          lngRetVal = .RecordCount
4890        End If
4900        .Close
4910      End With
4920    End With

EXITP:
4930    Set rst = Nothing
4940    RecCnt = lngRetVal
4950    Exit Function

ERRH:
4960    lngRetVal = 0&
4970    Select Case ERR.Number
        Case Else
4980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4990    End Select
5000    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

5110    Select Case intWhere
        Case acCmdRecordsGoToFirst
5120      DoCmd.RunCommand acCmdRecordsGoToFirst
5130    Case acCmdRecordsGoToLast
5140      DoCmd.RunCommand acCmdRecordsGoToLast
5150    Case acCmdRecordsGoToPrevious
5160      DoCmd.RunCommand acCmdRecordsGoToPrevious
5170    Case acCmdRecordsGoToNext
5180      DoCmd.RunCommand acCmdRecordsGoToNext
5190    Case Else
5200      If IsMissing(varID) = False Then
5210        With Me
5220          Set rst = .RecordsetClone
5230          With rst
5240            .FindFirst "[assettype_description] = '" & varID & "'"
5250            If .NoMatch = False Then
5260              Me.Bookmark = .Bookmark
5270            End If
5280            .Close
5290          End With
5300        End With
5310      End If
5320    End Select

EXITP:
5330    Set rst = Nothing
5340    Exit Sub

ERRH:
5350    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
5360    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5370    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5380    Case Else
5390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5400    End Select
5410    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [assettype]"

5510    With Me
5520      .Controls(strSortLbl).Visible = False
5530      .Controls(strSortLine).Visible = False
5540      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form
5550      If strProc = "Form_Load" Then
5560        strCalled = "assettype"
5570        strSortNow = strSortOrig
5580        lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (8& * 15&))
5590        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5600        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5610        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5620        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5630        .Controls(strSortLbl).Left = lngSortLbl_Left
5640        .Controls(strSortLbl).Top = lngSortLbl_Top
5650        .Controls(strSortLine).Top = lngSortLine_Top
5660        .Controls(strSortLine).Left = lngSortLine_Left
5670        .Controls(strSortLine).Width = lngSortLine_Width
5680        .Controls(strSortLbl).Caption = strArwUp
5690        .Controls(strSortLbl).ForeColor = CLR_DKBLU
5700      Else
5710        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
5720        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
5730        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5740        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5750        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5760        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5770        .Controls(strSortLbl).Left = lngSortLbl_Left
5780        .Controls(strSortLbl).Top = lngSortLbl_Top
5790        .Controls(strSortLine).Top = lngSortLine_Top
5800        .Controls(strSortLine).Left = lngSortLine_Left
5810        .Controls(strSortLine).Width = lngSortLine_Width
5820        If strCalled = "assettype" Then
5830          strSortAsc = strSortOrig
5840          If strSortNow = strSortAsc Then
5850            strSortNow = "[assettype] DESC"
5860            .Controls(strSortLbl).Caption = strArwDn
5870            .Controls(strSortLbl).ForeColor = CLR_DKRED
5880          Else
5890            strSortNow = strSortAsc
5900            .Controls(strSortLbl).Caption = strArwUp
5910            .Controls(strSortLbl).ForeColor = CLR_DKBLU
5920          End If
5930          lngSortLbl_Left = (lngSortLbl_Left + (8& * 15&))
5940          .Controls(strSortLbl).Left = lngSortLbl_Left
5950        Else
5960          .Controls(strSortLbl).Caption = strArwUp
5970          .Controls(strSortLbl).ForeColor = CLR_DKBLU
5980          Select Case strCalled
              Case "assettype_description"
5990            strSortAsc = "[assettype_description]"
6000          Case "Model", "Model_input"
6010            strSortAsc = "[Model]" & strStdAsc
6020          Case "Current"
6030            strSortAsc = "[Current]" & strStdAsc
6040          Case "Difference"
6050            strSortAsc = "[Difference]" & strStdAsc
6060          End Select
6070          If strSortNow = strSortAsc Then
6080            intCnt = CharCnt(strSortAsc, ",") + 1
6090            Select Case intCnt
                Case 1
6100              strTmp01 = strSortAsc & " DESC"
6110            Case 2
6120              intPos01 = InStr(strSortAsc, ",")
6130              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6140              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
6150            Case 3
6160              intPos01 = InStr(strSortAsc, ",")
6170              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6180              strTmp02 = Mid(strSortAsc, intPos01)
6190              intPos01 = InStr(2, strTmp02, ",")
6200              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
6210              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
6220            End Select
6230            strSortNow = strTmp01
6240            .Controls(strSortLbl).Caption = strArwDn
6250            .Controls(strSortLbl).ForeColor = CLR_DKRED
6260          Else
6270            strSortNow = strSortAsc
6280          End If
6290        End If
6300      End If
6310      .Controls(strSortLbl).Visible = True
6320      .Controls(strSortLine).Visible = True
6330      .OrderBy = strSortNow
6340      .OrderByOn = True
6350      .TimerInterval = 500&
6360    End With

EXITP:
6370    Exit Sub

ERRH:
6380    Select Case ERR.Number
        Case Else
6390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6400    End Select
6410    Resume EXITP

End Sub
