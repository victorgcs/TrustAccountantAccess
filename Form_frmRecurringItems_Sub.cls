VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRecurringItems_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRecurringItems_Sub"

'VGC 04/16/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd on frmRecurringItems}
' **   Close:            C {cmdClose on frmRecurringItems}
' **   Delete:           D {cmdDelete on frmRecurringItems}
' **   Enable Country:   E {chkEnableCountry on frmRecurringItems}
' **   Form 1099:        F {RecurringItemAux_1099}
' **   Paid:             I {opgType_optPaid on frmRecurringItems}
' **   All:              L {opgType_optAll on frmRecurringItems}
' **   Miscellaneous:    M {opgType_optMisc on frmRecurringItems}
' **   Description:      N {Recur_Name}
' **   Print Report:     P {cmdPrintReport on frmRecurringItems}
' **   Received:         R {opgType_optReceived on frmRecurringItems}
' **   Exit:             X {cmdClose on frmRecurringItems}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[Recur_Name], [Recur_Address]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngForm_Width As Long, lngCountry_Left As Long, lngPostalCode_Left As Long
Private lngForm1099_Left As Long, lngForm1099_Diff As Long, lngTaxID_Left As Long
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
150       If strCallingForm = "frmRecurringItems" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190         .form_width_line.Left = 0&  ' ** Just get it out of the way.

200         lngSortLbl_Top = .Sort_lbl.Top
210         lngSortLbl_Width = .Sort_lbl.Width

220         lngForm_Width = .Width
230         lngCountry_Left = .Recur_Country.Left
240         lngPostalCode_Left = .Recur_PostalCode.Left
250         lngForm1099_Left = .RecurringItemAux_1099_lbl.Left  ' ** Because it's a check box.
260         lngForm1099_Diff = (.RecurringItemAux_1099.Left - .RecurringItemAux_1099_lbl.Left)  ' ** Check box offset from label.
270         lngTaxID_Left = .RecurringItemAux_TaxNo.Left

            'blnResetAll = True
            '.TimerInterval = 500&

280       Else
290         Cancel = -1
300       End If
310     End With

320     If Cancel = -1 Then
330       Beep
340       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
350     End If

EXITP:
360     Exit Sub

ERRH:
370     Select Case ERR.Number
        Case Else
380       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
390     End Select
400     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me

520       .RecurringItemAux_1099.Enabled = False
530       .RecurringItemAux_1099_lbl.ForeColor = WIN_CLR_DISF
540       .RecurringItemAux_1099_lbl2.ForeColor = WIN_CLR_DISF
550       .RecurringItemAux_1099_lbl_dim_hi.Visible = True
560       .RecurringItemAux_1099_lbl2_dim_hi.Visible = True
570       .RecurringItemAux_1099_lbl_line.BorderColor = WIN_CLR_DISR
580       .RecurringItemAux_1099_lbl_line_dim_hi.Visible = True
590       .RecurringItemAux_TaxNo.Enabled = False
600       .RecurringItemAux_TaxNo.BorderColor = WIN_CLR_DISR
610       .RecurringItemAux_TaxNo_lbl.ForeColor = WIN_CLR_DISF
620       .RecurringItemAux_TaxNo_lbl_dim_hi.Visible = True
630       .RecurringItemAux_TaxNo_lbl_line.BorderColor = WIN_CLR_DISR
640       .RecurringItemAux_TaxNo_lbl_line_dim_hi.Visible = True

650       .Recur_Type.ForeColor = CLR_DISABLED_FG
660       .Recur_Type.BackColor = CLR_DISABLED_BG
670       SortNow THIS_PROC  ' ** Procedure: Below

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

Private Sub Form_Current()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

810     With Me
820       If .Recur_Name = RECUR_I_TO_P Or .Recur_Name = RECUR_P_TO_I Or _
              .RecurringItem_ID = RECUR_I_TO_P_ID Or .RecurringItem_ID = RECUR_P_TO_I_ID Then
830         .Recur_Name.Locked = True
840         .Recur_Name.BackColor = CLR_DISABLED_BG
850         If .Recur_Address.Enabled = True Then
860           .Recur_Address.Locked = True
870         End If
880         If .Recur_City.Enabled = True Then
890           .Recur_City.Locked = True
900         End If
910         If .Recur_State.Enabled = True Then
920           .Recur_State.Locked = True
930         End If
940         If .Recur_Zip.Enabled = True Then
950           .Recur_Zip.Locked = True
960         End If
970         If .Parent.chkEnableCountry = True Then
980           If .Recur_Country.Enabled = True Then
990             .Recur_Country.Locked = True
1000          End If
1010          If .Recur_PostalCode.Enabled = True Then
1020            .Recur_PostalCode.Locked = True
1030          End If
1040        End If
            'If .RecurringItemAux_1099.Enabled = True Then
            '  .RecurringItemAux_1099.Locked = True
            'End If
            'If .RecurringItemAux_TaxNo.Enabled = True Then
            '  .RecurringItemAux_TaxNo.Locked = True
            'End If
1050      Else
1060        .Recur_Name.Locked = False
1070        .Recur_Name.BackColor = CLR_WHT
1080        .Recur_Address.Locked = False
1090        .Recur_City.Locked = False
1100        .Recur_State.Locked = False
1110        .Recur_Zip.Locked = False
1120        If .Parent.chkEnableCountry = True Then
1130          .Recur_Country.Locked = False
1140          .Recur_PostalCode.Locked = False
1150        End If
1160        Select Case .Recur_Type
            Case "Misc"
1170          .Recur_Address.Enabled = False
1180          .Recur_Address.BorderColor = WIN_CLR_DISR
1190          .Recur_City.Enabled = False
1200          .Recur_City.BorderColor = WIN_CLR_DISR
1210          .Recur_State.Enabled = False
1220          .Recur_State.BorderColor = WIN_CLR_DISR
1230          .Recur_Zip.Enabled = False
1240          .Recur_Zip.BorderColor = WIN_CLR_DISR
1250          If .Parent.chkEnableCountry = True Then
1260            .Recur_Country.Enabled = False
1270            .Recur_Country.BorderColor = WIN_CLR_DISR
1280            .Recur_PostalCode.Enabled = False
1290            .Recur_PostalCode.BorderColor = WIN_CLR_DISR
1300          End If
              '.RecurringItemAux_1099.Enabled = False
              '.RecurringItemAux_TaxNo.Enabled = False
              '.RecurringItemAux_1099.Locked = False
              '.RecurringItemAux_TaxNo.Locked = False
1310          RecurHeaders False  ' ** Procedure: Below.
1320        Case "Payee"
1330          .Recur_Address.Enabled = True
1340          .Recur_Address.BorderColor = CLR_LTBLU2
1350          .Recur_City.Enabled = True
1360          .Recur_City.BorderColor = CLR_LTBLU2
1370          .Recur_State.Enabled = True
1380          .Recur_State.BorderColor = CLR_LTBLU2
1390          .Recur_Zip.Enabled = True
1400          .Recur_Zip.BorderColor = CLR_LTBLU2
1410          If .Parent.chkEnableCountry = True Then
1420            .Recur_Country.Enabled = True
1430            .Recur_Country.BorderColor = CLR_LTBLU2
1440            .Recur_PostalCode.Enabled = True
1450            .Recur_PostalCode.BorderColor = CLR_LTBLU2
1460          End If
              '.RecurringItemAux_1099.Enabled = True
              '.RecurringItemAux_TaxNo.Enabled = True
              'Select Case IsNull(.RecurringItemAux_1099)
              'Case True
              '  .RecurringItemAux_TaxNo.Locked = True
              'Case False
              '  Select Case .RecurringItemAux_1099
              '  Case True
              '    .RecurringItemAux_TaxNo.Locked = False
              '  Case False
              '    .RecurringItemAux_TaxNo.Locked = True
              '  End Select
              'End Select
1470          RecurHeaders True  ' ** Procedure: Below.
1480        Case "Payor"
1490          .Recur_Address.Enabled = False
1500          .Recur_Address.BorderColor = WIN_CLR_DISR
1510          .Recur_City.Enabled = False
1520          .Recur_City.BorderColor = WIN_CLR_DISR
1530          .Recur_State.Enabled = False
1540          .Recur_State.BorderColor = WIN_CLR_DISR
1550          .Recur_Zip.Enabled = False
1560          .Recur_Zip.BorderColor = WIN_CLR_DISR
1570          If .Parent.chkEnableCountry = True Then
1580            .Recur_Country.Enabled = False
1590            .Recur_Country.BorderColor = WIN_CLR_DISR
1600            .Recur_PostalCode.Enabled = False
1610            .Recur_PostalCode.BorderColor = WIN_CLR_DISR
1620          End If
              '.RecurringItemAux_1099.Enabled = False
              '.RecurringItemAux_TaxNo.Enabled = False
              '.RecurringItemAux_1099.Locked = False
              '.RecurringItemAux_TaxNo.Locked = False
1630          RecurHeaders False  ' ** procedure: Below.
1640        End Select
1650      End If
1660    End With

EXITP:
1670    Exit Sub

ERRH:
1680    Select Case ERR.Number
        Case Else
1690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1700    End Select
1710    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1820    intShiftDown = (Shift And acShiftMask) > 0
1830    intAltDown = (Shift And acAltMask) > 0
1840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Form 1099:        F {RecurringItemAux_1099}
        ' **   Description:      N {Recur_Name}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1860      Select Case intRetVal
          Case vbKeyDown
1870        intRetVal = 0
1880        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1890      Case vbKeyUp
1900        intRetVal = 0
1910        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1920      Case vbKeyEscape
1930        With Me
1940          intRetVal = 0
1950          .Parent.cmdClose_Click  ' ** Form Procedure: frmRecurringItems.
1960        End With
1970      End Select
1980    End If

        ' ** Alt keys.
1990    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2000      Select Case intRetVal
          Case vbKeyF
2010        With Me
2020          intRetVal = 0
2030          If .RecurringItemAux_1099.Enabled = True Then
2040            .RecurringItemAux_1099.SetFocus
2050            .RecurringItemAux_1099 = (Not .RecurringItemAux_1099)
2060            RecurringItemAux_1099_AfterUpdate  ' ** Procedure: Below.
2070          Else
2080            Beep
2090          End If
2100        End With
2110      Case vbKeyN
2120        With Me
2130          intRetVal = 0
2140          .Recur_Name.SetFocus
2150        End With
2160      Case vbKeyX
2170        With Me
2180          intRetVal = 0
2190          .Parent.cmdClose_Click  ' ** Form Procedure: frmRecurringItems.
2200        End With
2210      End Select
2220    End If

        ' ** Ctrl keys.
2230    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2240      Select Case intRetVal
          Case vbKeyS
2250        intRetVal = 0
2260        cmdSave_Click  ' ** Procedure: Below.
2270      Case vbKeyTab
2280        With Me
2290          intRetVal = 0
2300          DoCmd.SelectObject acForm, .Parent.Name, False
2310          If .Parent.cmdPrintReport.Enabled = True Then
2320            .Parent.cmdPrintReport.SetFocus
2330          Else
2340            .Parent.opgType.SetFocus
2350          End If
2360        End With
2370      Case vbKeyPageUp, vbKeyHome
2380        intRetVal = 0
2390        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2400      Case vbKeyPageDown, vbKeyEnd
2410        intRetVal = 0
2420        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2430      End Select
2440    End If

        ' ** Ctrl-Shift keys.
2450    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2460      Select Case intRetVal
          Case vbKeyF
2470        With Me
2480          intRetVal = 0
2490          DoCmd.SelectObject acForm, .Parent.Name, False
2500          .Parent.FocusHolder.SetFocus
2510        End With
2520      Case vbKeyTab
2530        With Me
2540          intRetVal = 0
2550          DoCmd.SelectObject acForm, .Parent.Name, False
2560          .Parent.cmdClose.SetFocus
2570        End With
2580      End Select
2590    End If

EXITP:
2600    KeyCode = intRetVal
2610    Exit Sub

ERRH:
2620    intRetVal = 0
2630    THAT_PROC = THIS_PROC
2640    That_Erl = Erl
2650    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2660    Resume EXITP

End Sub

Private Sub Form_Timer()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

2710    With Me
2720      .TimerInterval = 0&
          'If blnResetAll = True Then
          '  blnResetAll = False
          '  If .Parent.opgType = .Parent.opgType_optAll.OptionValue Then
          '    If .Recur_Type = "Payee" Then
          '      RecurHeaders True  ' ** Procedure: Below.
          '    Else
          '      RecurHeaders False  ' ** Procedure: Below.
          '    End If
          '  End If
          'End If
2730    End With

EXITP:
2740    Exit Sub

ERRH:
2750    Select Case ERR.Number
        Case Else
2760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2770    End Select
2780    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2810    If DataErr <> 0 Then
2820      Select Case DataErr
          Case 2501, 3101, 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2237, 2427
            ' ** 2501: The '|' action was Canceled.
2830        Response = acDataErrContinue
2840      Case 2046  ' ** The command or action isn't available now (first or last record).
2850        MsgBox "There are no records to delete.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
2860        Response = acDataErrContinue
2870      Case 3162
2880        MsgBox "The " & Me.ActiveControl.Name & " field can not be left blank.", vbInformation + vbOKOnly, "Entry Required"
2890        Response = acDataErrContinue
2900      Case 3058
2910        MsgBox "You must enter the Description to continue.", vbInformation + vbOKOnly, "Entry Required"
2920        Me.Recur_Name.SetFocus
2930        Response = acDataErrContinue
2940      Case Else
2950        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2960        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
2970      End Select
2980    End If

EXITP:
2990    Exit Sub

ERRH:
3000    Select Case ERR.Number
        Case Else
3010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3020    End Select
3030    Resume EXITP

End Sub

Public Sub cmdSave_Click()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3110    DoCmd.RunCommand acCmdSaveRecord

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3220    intShiftDown = (Shift And acShiftMask) > 0
3230    intAltDown = (Shift And acAltMask) > 0
3240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3270        With Me
3280          intRetVal = 0
3290          .Recur_Name.SetFocus
3300        End With
3310      End Select
3320    End If

        ' ** Shift keys.
3330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3350        With Me
3360          intRetVal = 0
3370          If .CurrentRecord > 1 Then
3380            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3390            If .Recur_Zip.Enabled = True And .Recur_Zip.Visible = True Then
3400              If .RecurringItemAux_1099.Enabled = True Then
3410                .RecurringItemAux_TaxNo.SetFocus
3420              Else
3430                If IsNull(.Recur_State) = True And IsNull(.Recur_Zip) = True Then
3440                  .Recur_PostalCode.SetFocus
3450                Else
3460                  .Recur_Zip.SetFocus
3470                End If
3480              End If
3490            Else
3500              .Recur_Name.SetFocus
3510            End If
3520          Else
3530            DoCmd.SelectObject acForm, .Parent.Name, False
3540            .Parent.cmdClose.SetFocus
3550          End If
3560        End With
3570      End Select
3580    End If

EXITP:
3590    KeyCode = intRetVal
3600    Exit Sub

ERRH:
3610    intRetVal = 0
3620    THAT_PROC = THIS_PROC
3630    That_Erl = Erl
3640    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3650    Resume EXITP

End Sub

Private Sub FocusHolder2_KeyDown(KeyCode As Integer, Shift As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3720    intShiftDown = (Shift And acShiftMask) > 0
3730    intAltDown = (Shift And acAltMask) > 0
3740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3770        With Me
3780          intRetVal = 0
3790          lngRecsCur = RecCnt  ' ** Function: Below.
3800          If lngRecsCur > 0& Then
3810            .Recur_Name.SetFocus
3820          Else
3830            DoCmd.SelectObject acForm, .Parent.Name, False
3840            If .Parent.cmdPrintReport.Enabled = True Then
3850              .Parent.cmdPrintReport.SetFocus
3860            Else
3870              .Parent.opgType.SetFocus
3880            End If
3890          End If
3900        End With
3910      End Select
3920    End If

        ' ** Shift keys.
3930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3950        With Me
3960          intRetVal = 0
3970          lngRecsCur = RecCnt  ' ** Function: Below.
3980          If lngRecsCur > 0& Then
3990            If .CurrentRecord > 1 Then
4000              MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4010              If .Recur_Zip.Enabled = True And .Recur_Zip.Visible = True Then
4020                If .RecurringItemAux_TaxNo.Enabled = True Then
4030                  .RecurringItemAux_TaxNo.SetFocus
4040                Else
4050                  If IsNull(.Recur_State) = True And IsNull(.Recur_Zip) = True Then
4060                    .Recur_PostalCode.SetFocus
4070                  Else
4080                    .Recur_Zip.SetFocus
4090                  End If
4100                End If
4110              Else
4120                .Recur_Name.SetFocus
4130              End If
4140            Else
4150              DoCmd.SelectObject acForm, .Parent.Name, False
4160              .Parent.cmdClose.SetFocus
4170            End If
4180          Else
4190            DoCmd.SelectObject acForm, .Parent.Name, False
4200            .Parent.cmdClose.SetFocus
4210          End If
4220        End With
4230      End Select
4240    End If

EXITP:
4250    KeyCode = intRetVal
4260    Exit Sub

ERRH:
4270    intRetVal = 0
4280    THAT_PROC = THIS_PROC
4290    That_Erl = Erl
4300    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4310    Resume EXITP

End Sub

Private Sub Recur_Type_lbl_DblClick(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Type_lbl_DblClick"

4410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub Recur_Name_lbl_DblClick(Cancel As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_lbl_DblClick"

4510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub Recur_Name_GotFocus()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_GotFocus"

4610    With Me
4620      .Recur_Name.SelLength = 0
4630      .Recur_Name.SelStart = 0
4640    End With

EXITP:
4650    Exit Sub

ERRH:
4660    Select Case ERR.Number
        Case Else
4670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4680    End Select
4690    Resume EXITP

End Sub

Private Sub Recur_Name_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_KeyDown"

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
4790          If .Recur_Address.Enabled = True And .Recur_Address.Visible = True Then
4800            .Recur_Address.SetFocus
4810          Else
4820            lngRecsCur = RecCnt  ' ** Function: Below.
4830            If .CurrentRecord < lngRecsCur Then
4840              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4850              .Recur_Name.SetFocus
4860            Else
4870              DoCmd.SelectObject acForm, .Parent.Name, False
4880              If .Parent.cmdPrintReport.Enabled = True Then
4890                .Parent.cmdPrintReport.SetFocus
4900              Else
4910                .Parent.opgType.SetFocus
4920              End If
4930            End If
4940          End If
4950        End With
4960      End Select
4970    End If

        ' ** Shift keys.
4980    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4990      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5000        With Me
5010          intRetVal = 0
5020          If .CurrentRecord > 1 Then
5030            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
5040            If .Recur_Zip.Enabled = True And .Recur_Zip.Visible = True Then
5050              If .RecurringItemAux_TaxNo.Enabled = True Then
5060                .RecurringItemAux_TaxNo.SetFocus
5070              Else
5080                If IsNull(.Recur_State) = True And IsNull(.Recur_Zip) = True Then
5090                  .Recur_PostalCode.SetFocus
5100                Else
5110                  .Recur_Zip.SetFocus
5120                End If
5130              End If
5140            Else
5150              .Recur_Name.SetFocus
5160            End If
5170          Else
5180            DoCmd.SelectObject acForm, .Parent.Name, False
5190            .Parent.cmdClose.SetFocus
5200          End If
5210        End With
5220      End Select
5230    End If

EXITP:
5240    KeyCode = intRetVal
5250    Exit Sub

ERRH:
5260    intRetVal = 0
5270    THAT_PROC = THIS_PROC
5280    That_Erl = Erl
5290    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5300    Resume EXITP

End Sub

Private Sub Recur_Address_lbl_DblClick(Cancel As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Address_lbl_DblClick"

5410    With Me
5420      If .Recur_Type = "Payee" Then
5430        SortNow THIS_PROC  ' ** Procedure: Below.
5440      End If
5450    End With

EXITP:
5460    Exit Sub

ERRH:
5470    Select Case ERR.Number
        Case Else
5480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5490    End Select
5500    Resume EXITP

End Sub

Private Sub Recur_Address_AfterUpdate()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Address_AfterUpdate"

        Dim strTmp01 As String

5610    With Me
5620      If IsNull(.Recur_Address) = False Then
5630        strTmp01 = FixQuotes(.Recur_Address)  ' ** Module Function: modStringFuncs.
5640        If Compare_StringA_StringB(.Recur_Address, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
5650          .Recur_Address = strTmp01
5660          cmdSave_Click  ' ** Procedure: Above.
5670        End If
5680      End If
5690    End With

EXITP:
5700    Exit Sub

ERRH:
5710    Select Case ERR.Number
        Case Else
5720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5730    End Select
5740    Resume EXITP

End Sub

Private Sub Recur_City_lbl_DblClick(Cancel As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_City_lbl_DblClick"

5810    With Me
5820      If .Recur_Type = "Payee" Then
5830        SortNow THIS_PROC  ' ** Procedure: Below.
5840      End If
5850    End With

EXITP:
5860    Exit Sub

ERRH:
5870    Select Case ERR.Number
        Case Else
5880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5890    End Select
5900    Resume EXITP

End Sub

Private Sub Recur_City_AfterUpdate()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_City_AfterUpdate"

        Dim strTmp01 As String

6010    With Me
6020      If IsNull(.Recur_City) = False Then
6030        strTmp01 = FixQuotes(.Recur_City)  ' ** Module Function: modStringFuncs.
6040        If Compare_StringA_StringB(.Recur_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
6050          .Recur_City = strTmp01
6060          cmdSave_Click  ' ** Procedure: Above.
6070        End If
6080      End If
6090    End With

EXITP:
6100    Exit Sub

ERRH:
6110    Select Case ERR.Number
        Case Else
6120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6130    End Select
6140    Resume EXITP

End Sub

Private Sub Recur_City_KeyDown(KeyCode As Integer, Shift As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_City_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6220    intShiftDown = (Shift And acShiftMask) > 0
6230    intAltDown = (Shift And acAltMask) > 0
6240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6270        With Me
6280          intRetVal = 0
6290          If IsNull(.Recur_Country) = False Or IsNull(.Recur_PostalCode) = False Then
6300            .Recur_Country.SetFocus
6310          Else
6320            .Recur_State.SetFocus
6330          End If
6340        End With
6350      End Select
6360    End If

        ' ** Shift keys.
6370    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6380      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6390        With Me
6400          intRetVal = 0
6410          .Recur_Address.SetFocus
6420        End With
6430      End Select
6440    End If

EXITP:
6450    KeyCode = intRetVal
6460    Exit Sub

ERRH:
6470    intRetVal = 0
6480    THAT_PROC = THIS_PROC
6490    That_Erl = Erl
6500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6510    Resume EXITP

End Sub

Private Sub Recur_State_lbl_DblClick(Cancel As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_State_lbl_DblClick"

6610    With Me
6620      If .Recur_Type = "Payee" Then
6630        SortNow THIS_PROC  ' ** Procedure: Below.
6640      End If
6650    End With

EXITP:
6660    Exit Sub

ERRH:
6670    Select Case ERR.Number
        Case Else
6680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6690    End Select
6700    Resume EXITP

End Sub

Private Sub Recur_State_Enter()

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_State_Enter"

6810    With Me
6820      If IsNull(.Recur_State) = True Then
6830        ExpandCombo .Recur_State  ' ** Module Procedure: modUtilities.
6840      End If
6850    End With

EXITP:
6860    Exit Sub

ERRH:
6870    Select Case ERR.Number
        Case Else
6880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6890    End Select
6900    Resume EXITP

End Sub

Private Sub Recur_State_KeyUp(KeyCode As Integer, Shift As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_State_KeyUp"

7010    With Me
7020      If .Recur_State.SelStart <> 0 Then
7030        .Recur_State.SelStart = 1
7040        .Recur_State.SelLength = 1
7050      End If
7060    End With

EXITP:
7070    Exit Sub

ERRH:
7080    Select Case ERR.Number
        Case Else
7090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7100    End Select
7110    Resume EXITP

End Sub

Private Sub Recur_State_AfterUpdate()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_State_AfterUpdate"

7210    With Me
7220      If IsNull(.Recur_State) = False Then
7230        If Trim(.Recur_State) <> vbNullString Then
7240          Select Case IsNull(.Parent.Recur_State_Pref)
              Case True
7250            .Parent.Recur_State_Pref = .Recur_State
7260            .Recur_State.DefaultValue = Chr(34) & .Recur_State & Chr(34)
7270          Case False
7280            If .Parent.Recur_State_Pref <> .Recur_State Then
7290              .Parent.Recur_State_Pref = .Recur_State
7300              .Recur_State.DefaultValue = Chr(34) & .Recur_State & Chr(34)
7310            End If
7320          End Select
7330        End If
7340      End If
7350    End With

EXITP:
7360    Exit Sub

ERRH:
7370    Select Case ERR.Number
        Case Else
7380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7390    End Select
7400    Resume EXITP

End Sub

Private Sub Recur_State_NotInList(NewData As String, Response As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_State_NotInList"

7510    With Me
7520      .Recur_State = Null
7530      MsgBox "Must select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
7540      Response = acDataErrContinue
7550    End With

EXITP:
7560    Exit Sub

ERRH:
7570    Select Case ERR.Number
        Case Else
7580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7590    End Select
7600    Resume EXITP

End Sub

Private Sub Recur_Zip_lbl_DblClick(Cancel As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Zip_lbl_DblClick"

7710    With Me
7720      If .Recur_Type = "Payee" Then
7730        SortNow THIS_PROC  ' ** Procedure: Below.
7740      End If
7750    End With

EXITP:
7760    Exit Sub

ERRH:
7770    Select Case ERR.Number
        Case Else
7780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7790    End Select
7800    Resume EXITP

End Sub

Private Sub Recur_Zip_KeyDown(KeyCode As Integer, Shift As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Zip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7920    intShiftDown = (Shift And acShiftMask) > 0
7930    intAltDown = (Shift And acAltMask) > 0
7940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7970        With Me
7980          intRetVal = 0
7990          Select Case .Parent.chkEnableCountry
              Case True
8000            If IsNull(.Recur_State) = True And IsNull(.Recur_Zip) = True Then
8010              .Recur_Country.SetFocus
8020            ElseIf .RecurringItemAux_1099.Enabled = True Then
8030              .RecurringItemAux_1099.SetFocus
8040            Else
8050              lngRecsCur = RecCnt  ' ** Function: Below.
8060              If .CurrentRecord < lngRecsCur Then
8070                MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
8080                .Recur_Name.SetFocus
8090              Else
8100                DoCmd.SelectObject acForm, .Parent.Name, False
8110                If .Parent.cmdPrintReport.Enabled = True Then
8120                  .Parent.cmdPrintReport.SetFocus
8130                Else
8140                  .Parent.opgType.SetFocus
8150                End If
8160              End If
8170            End If
8180          Case False
8190            If .RecurringItemAux_1099.Enabled = True Then
8200              .RecurringItemAux_1099.SetFocus
8210            Else
8220              lngRecsCur = RecCnt  ' ** Function: Below.
8230              If .CurrentRecord < lngRecsCur Then
8240                MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
8250                .Recur_Name.SetFocus
8260              Else
8270                DoCmd.SelectObject acForm, .Parent.Name, False
8280                If .Parent.cmdPrintReport.Enabled = True Then
8290                  .Parent.cmdPrintReport.SetFocus
8300                Else
8310                  .Parent.opgType.SetFocus
8320                End If
8330              End If
8340            End If
8350          End Select
8360        End With
8370      End Select
8380    End If

        ' ** Shift keys.
8390    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8400      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8410        With Me
8420          intRetVal = 0
8430          .Recur_State.SetFocus
8440        End With
8450      End Select
8460    End If

EXITP:
8470    KeyCode = intRetVal
8480    Exit Sub

ERRH:
8490    intRetVal = 0
8500    THAT_PROC = THIS_PROC
8510    That_Erl = Erl
8520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8530    Resume EXITP

End Sub

Private Sub Recur_Country_lbl_DblClick(Cancel As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Country_lbl_DblClick"

8610    With Me
8620      If .Recur_Type = "Payee" Then
8630        SortNow THIS_PROC  ' ** Procedure: Below.
8640      End If
8650    End With

EXITP:
8660    Exit Sub

ERRH:
8670    Select Case ERR.Number
        Case Else
8680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8690    End Select
8700    Resume EXITP

End Sub

Private Sub Recur_Country_AfterUpdate()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Country_AfterUpdate"

        Dim strTmp01 As String

8810    With Me
8820      If IsNull(.Recur_Country) = False Then
8830        strTmp01 = FixQuotes(.Recur_Country)  ' ** Module Function: modStringFuncs.
8840        If Compare_StringA_StringB(.Recur_Country, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
8850          .Recur_Country = strTmp01
8860          cmdSave_Click  ' ** Procedure: Above.
8870        End If
8880      End If
8890    End With

EXITP:
8900    Exit Sub

ERRH:
8910    Select Case ERR.Number
        Case Else
8920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8930    End Select
8940    Resume EXITP

End Sub

Private Sub Recur_Country_KeyDown(KeyCode As Integer, Shift As Integer)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Country_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9020    intShiftDown = (Shift And acShiftMask) > 0
9030    intAltDown = (Shift And acAltMask) > 0
9040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9070        With Me
9080          intRetVal = 0
9090          .Recur_PostalCode.SetFocus
9100        End With
9110      End Select
9120    End If

        ' ** Shift keys.
9130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9150        With Me
9160          intRetVal = 0
9170          If IsNull(.Recur_Country) = False Or IsNull(.Recur_PostalCode) = False Then
9180            .Recur_City.SetFocus
9190          Else
9200            .Recur_Zip.SetFocus
9210          End If
9220        End With
9230      End Select
9240    End If

EXITP:
9250    KeyCode = intRetVal
9260    Exit Sub

ERRH:
9270    intRetVal = 0
9280    THAT_PROC = THIS_PROC
9290    That_Erl = Erl
9300    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9310    Resume EXITP

End Sub

Private Sub Recur_PostalCode_lbl_DblClick(Cancel As Integer)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_PostalCode_lbl_DblClick"

9410    With Me
9420      If .Recur_Type = "Payee" Then
9430        SortNow THIS_PROC  ' ** Procedure: Below.
9440      End If
9450    End With

EXITP:
9460    Exit Sub

ERRH:
9470    Select Case ERR.Number
        Case Else
9480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9490    End Select
9500    Resume EXITP

End Sub

Private Sub Recur_PostalCode_AfterUpdate()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_PostalCode_AfterUpdate"

9610    With Me

9620    End With

EXITP:
9630    Exit Sub

ERRH:
9640    Select Case ERR.Number
        Case Else
9650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9660    End Select
9670    Resume EXITP

End Sub

Private Sub Recur_PostalCode_KeyDown(KeyCode As Integer, Shift As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_PostalCode_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9720    intShiftDown = (Shift And acShiftMask) > 0
9730    intAltDown = (Shift And acAltMask) > 0
9740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9770        With Me
9780          intRetVal = 0
9790          If .RecurringItemAux_1099.Enabled = True Then
9800            .RecurringItemAux_1099.SetFocus
9810          Else
9820            lngRecsCur = RecCnt  ' ** Function: Below.
9830            If .CurrentRecord < lngRecsCur Then
9840              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
9850              .Recur_Name.SetFocus
9860            Else
9870              DoCmd.SelectObject acForm, .Parent.Name, False
9880              If .Parent.cmdPrintReport.Enabled = True Then
9890                .Parent.cmdPrintReport.SetFocus
9900              Else
9910                .Parent.opgType.SetFocus
9920              End If
9930            End If
9940          End If
9950        End With
9960      End Select
9970    End If

        ' ** Shift keys.
9980    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9990      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10000       With Me
10010         intRetVal = 0
10020         .Recur_Country.SetFocus
10030       End With
10040     End Select
10050   End If

EXITP:
10060   KeyCode = intRetVal
10070   Exit Sub

ERRH:
10080   intRetVal = 0
10090   THAT_PROC = THIS_PROC
10100   That_Erl = Erl
10110   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10120   Resume EXITP

End Sub

Private Sub RecurringItemAux_1099_lbl_DblClick(Cancel As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_1099_lbl_DblClick"

10210   With Me
10220     If .Recur_Type = "Payee" Then
10230       SortNow THIS_PROC  ' ** Procedure: Below.
10240     End If
10250   End With

EXITP:
10260   Exit Sub

ERRH:
10270   Select Case ERR.Number
        Case Else
10280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10290   End Select
10300   Resume EXITP

End Sub

Private Sub RecurringItemAux_1099_Enter()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_1099_Enter"

10410   With Me
10420     If IsNull(.RecurringItemAux_1099) = True Then
10430       .RecurringItemAux_1099 = False
10440       .RecurringItem_ID2 = .RecurringItem_ID
10450       .RecurringItemAux_DateModified = Now()
10460     End If
10470   End With

EXITP:
10480   Exit Sub

ERRH:
10490   Select Case ERR.Number
        Case Else
10500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10510   End Select
10520   Resume EXITP

End Sub

Public Sub RecurringItemAux_1099_AfterUpdate()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_1099_AfterUpdate"

10610   With Me
10620     Select Case .RecurringItemAux_1099
          Case True
10630       .RecurringItemAux_TaxNo.Locked = False
10640     Case False
10650       .RecurringItemAux_TaxNo.Locked = True
10660       If IsNull(.RecurringItemAux_TaxNo) = False Then
10670         .RecurringItemAux_TaxNo = Null
10680       End If
10690     End Select
10700     cmdSave_Click  ' ** Procedure: Above.
10710   End With

EXITP:
10720   Exit Sub

ERRH:
10730   Select Case ERR.Number
        Case Else
10740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10750   End Select
10760   Resume EXITP

End Sub

Private Sub RecurringItemAux_1099_KeyDown(KeyCode As Integer, Shift As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_1099_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10820   intShiftDown = (Shift And acShiftMask) > 0
10830   intAltDown = (Shift And acAltMask) > 0
10840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10870       With Me
10880         intRetVal = 0
10890         Select Case .RecurringItemAux_1099
              Case True
10900           .RecurringItemAux_TaxNo.SetFocus
10910         Case False
10920           lngRecsCur = RecCnt  ' ** Function: Below.
10930           If .CurrentRecord < lngRecsCur Then
10940             MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
10950             .Recur_Name.SetFocus
10960           Else
10970             DoCmd.SelectObject acForm, .Parent.Name, False
10980             If .Parent.cmdPrintReport.Enabled = True Then
10990               .Parent.cmdPrintReport.SetFocus
11000             Else
11010               .Parent.opgType.SetFocus
11020             End If
11030           End If
11040         End Select
11050       End With
11060     End Select
11070   End If

        ' ** Shift keys.
11080   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11090     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11100       With Me
11110         intRetVal = 0
11120         Select Case .Parent.chkEnableCountry
              Case True
11130           If IsNull(.Recur_State) = True And IsNull(.Recur_Zip) = True Then
11140             .Recur_Country.SetFocus
11150           Else
11160             .Recur_Zip.SetFocus
11170           End If
11180         Case False
11190           .Recur_Zip.SetFocus
11200         End Select
11210       End With
11220     End Select
11230   End If

EXITP:
11240   KeyCode = intRetVal
11250   Exit Sub

ERRH:
11260   intRetVal = 0
11270   THAT_PROC = THIS_PROC
11280   That_Erl = Erl
11290   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11300   Resume EXITP

End Sub

Private Sub RecurringItemAux_1099_Exit(Cancel As Integer)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_1099_Exit"

11410   With Me
11420     If .Dirty = True Then
11430       cmdSave_Click  ' ** Procedure: Above.
11440     End If
11450   End With

EXITP:
11460   Exit Sub

ERRH:
11470   Select Case ERR.Number
        Case Else
11480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11490   End Select
11500   Resume EXITP

End Sub

Private Sub RecurringItemAux_TaxNo_lbl_DblClick(Cancel As Integer)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_TaxNo_lbl_DblClick"

11610   With Me
11620     If .Recur_Type = "Payee" Then
11630       SortNow THIS_PROC  ' ** Procedure: Below.
11640     End If
11650   End With

EXITP:
11660   Exit Sub

ERRH:
11670   Select Case ERR.Number
        Case Else
11680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11690   End Select
11700   Resume EXITP

End Sub

Private Sub RecurringItemAux_TaxNo_AfterUpdate()

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_TaxNo_AfterUpdate"

11810   cmdSave_Click  ' ** Procedure: Above.

EXITP:
11820   Exit Sub

ERRH:
11830   Select Case ERR.Number
        Case Else
11840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11850   End Select
11860   Resume EXITP

End Sub

Private Sub RecurringItemAux_TaxNo_KeyDown(KeyCode As Integer, Shift As Integer)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_TaxNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11920   intShiftDown = (Shift And acShiftMask) > 0
11930   intAltDown = (Shift And acAltMask) > 0
11940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11970       With Me
11980         intRetVal = 0
11990         lngRecsCur = RecCnt  ' ** Function: Below.
12000         If .CurrentRecord < lngRecsCur Then
12010           MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
12020           .Recur_Name.SetFocus
12030         Else
12040           DoCmd.SelectObject acForm, .Parent.Name, False
12050           If .Parent.cmdPrintReport.Enabled = True Then
12060             .Parent.cmdPrintReport.SetFocus
12070           Else
12080             .Parent.opgType.SetFocus
12090           End If
12100         End If
12110       End With
12120     End Select
12130   End If

        ' ** Shift keys.
12140   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12150     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12160       With Me
12170         intRetVal = 0
12180         .RecurringItemAux_1099.SetFocus
12190       End With
12200     End Select
12210   End If

EXITP:
12220   KeyCode = intRetVal
12230   Exit Sub

ERRH:
12240   intRetVal = 0
12250   THAT_PROC = THIS_PROC
12260   That_Erl = Erl
12270   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12280   Resume EXITP

End Sub

Private Sub RecurringItemAux_TaxNo_Exit(Cancel As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_TaxNo_Exit"

12310   With Me
12320     If IsNull(.RecurringItemAux_TaxNo) = False Then
12330       If IsNull(.RecurringItemAux_1099) = True Then
12340         .RecurringItemAux_1099 = False
12350         .RecurringItem_ID2 = .RecurringItem_ID
12360         .RecurringItemAux_DateModified = Now()
12370         .RecurringItemAux_TaxNo = Null
12380       Else
12390         If .RecurringItemAux_1099 = False Then
12400           .RecurringItemAux_TaxNo = Null
12410         End If
12420       End If
12430     End If
12440     If .Dirty = True Then
12450       cmdSave_Click  ' ** Procedure: Above.
12460     End If
12470   End With

EXITP:
12480   Exit Sub

ERRH:
12490   Select Case ERR.Number
        Case Else
12500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12510   End Select
12520   Resume EXITP

End Sub

Public Sub EnableAddressInfo()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "EnableAddressInfo"

12610   With Me
12620     lngRecsCur = RecCnt  ' ** Function: Below.
12630     If lngRecsCur > 0& Then
12640       .Recur_Name.SetFocus
12650     Else
12660       .FocusHolder2.SetFocus
12670     End If
12680     Select Case .Parent.opgType
          Case .Parent.opgType_optMisc.OptionValue
12690       .Recur_Address.Enabled = False
12700       .Recur_Address.Locked = False
12710       .Recur_Address.BorderColor = WIN_CLR_DISR
12720       .Recur_City.Enabled = False
12730       .Recur_City.Locked = False
12740       .Recur_City.BorderColor = WIN_CLR_DISR
12750       .Recur_State.Enabled = False
12760       .Recur_State.Locked = False
12770       .Recur_State.BorderColor = WIN_CLR_DISR
12780       .Recur_Zip.Enabled = False
12790       .Recur_Zip.Locked = False
12800       .Recur_Zip.BorderColor = WIN_CLR_DISR
12810       .Recur_Country.Enabled = False
12820       .Recur_Country.Locked = False
12830       .Recur_Country.BorderColor = WIN_CLR_DISR
12840       .Recur_PostalCode.Enabled = False
12850       .Recur_PostalCode.Locked = False
12860       .Recur_PostalCode.BorderColor = WIN_CLR_DISR
            '.RecurringItemAux_1099.Enabled = False
            '.RecurringItemAux_1099.Locked = False
            '.RecurringItemAux_TaxNo.Enabled = False
            '.RecurringItemAux_TaxNo.Locked = False
12870       RecurHeaders False  ' ** Procedure: Below.
12880     Case .Parent.opgType_optPaid.OptionValue
12890       .Recur_Address.Enabled = True
12900       .Recur_Address.BorderColor = CLR_LTBLU2
12910       .Recur_City.Enabled = True
12920       .Recur_City.BorderColor = CLR_LTBLU2
12930       .Recur_State.Enabled = True
12940       .Recur_State.BorderColor = CLR_LTBLU2
12950       .Recur_Zip.Enabled = True
12960       .Recur_Zip.BorderColor = CLR_LTBLU2
12970       If .Parent.chkEnableCountry = True Then
12980         .Recur_Country.Enabled = True
12990         .Recur_Country.BorderColor = CLR_LTBLU2
13000         .Recur_PostalCode.Enabled = True
13010         .Recur_PostalCode.BorderColor = CLR_LTBLU2
13020       End If
            '.RecurringItemAux_1099.Enabled = True
            '.RecurringItemAux_1099.Locked = False
            '.RecurringItemAux_TaxNo.Enabled = True
            'Select Case .RecurringItemAux_1099
            'Case True
            '  .RecurringItemAux_TaxNo.Locked = False
            'Case False
            '  .RecurringItemAux_TaxNo.Locked = True
            'End Select
13030       RecurHeaders True  ' ** Procedure: Below.
13040     Case .Parent.opgType_optReceived.OptionValue
13050       .Recur_Address.Enabled = False
13060       .Recur_Address.Locked = False
13070       .Recur_Address.BorderColor = WIN_CLR_DISR
13080       .Recur_City.Enabled = False
13090       .Recur_City.Locked = False
13100       .Recur_City.BorderColor = WIN_CLR_DISR
13110       .Recur_State.Enabled = False
13120       .Recur_State.Locked = False
13130       .Recur_State.BorderColor = WIN_CLR_DISR
13140       .Recur_Zip.Enabled = False
13150       .Recur_Zip.Locked = False
13160       .Recur_Zip.BorderColor = WIN_CLR_DISR
13170       .Recur_Country.Enabled = False
13180       .Recur_Country.Locked = False
13190       .Recur_Country.BorderColor = WIN_CLR_DISR
13200       .Recur_PostalCode.Enabled = False
13210       .Recur_PostalCode.Locked = False
13220       .Recur_PostalCode.BorderColor = WIN_CLR_DISR
            '.RecurringItemAux_1099.Enabled = False
            '.RecurringItemAux_1099.Locked = False
            '.RecurringItemAux_TaxNo.Enabled = False
            '.RecurringItemAux_TaxNo.Locked = False
13230       RecurHeaders False  ' ** Procedure: Below.
13240     Case .Parent.opgType_optAll.OptionValue
            ' ** Set as non-Payee, then let Form_Current() deal with it.
13250       .Recur_Address.Enabled = False
13260       .Recur_Address.Locked = False
13270       .Recur_Address.BorderColor = WIN_CLR_DISR
13280       .Recur_City.Enabled = False
13290       .Recur_City.Locked = False
13300       .Recur_City.BorderColor = WIN_CLR_DISR
13310       .Recur_State.Enabled = False
13320       .Recur_State.Locked = False
13330       .Recur_State.BorderColor = WIN_CLR_DISR
13340       .Recur_Zip.Enabled = False
13350       .Recur_Zip.Locked = False
13360       .Recur_Zip.BorderColor = WIN_CLR_DISR
13370       .Recur_Country.Enabled = False
13380       .Recur_Country.Locked = False
13390       .Recur_Country.BorderColor = WIN_CLR_DISR
13400       .Recur_PostalCode.Enabled = False
13410       .Recur_PostalCode.Locked = False
13420       .Recur_PostalCode.BorderColor = WIN_CLR_DISR
            '.RecurringItemAux_1099.Enabled = False
            '.RecurringItemAux_1099.Locked = False
            '.RecurringItemAux_TaxNo.Enabled = False
            '.RecurringItemAux_TaxNo.Locked = False
13430       RecurHeaders False  ' ** Procedure: Below.
13440     End Select
13450     If .Parent.opgType <> .Parent.opgType_optPaid.OptionValue Then
13460       If .Sort_line.Left >= .Recur_Address_lbl.Left Then
13470         strSortNow = strSortOrig
13480         .OrderBy = strSortNow
13490         .OrderByOn = True
13500         .Sort_lbl.Left = ((.Recur_Name_lbl.Left + .Recur_Name_lbl.Width) - lngSortLbl_Width)
13510         .Sort_lbl.Caption = strArwUp
13520         .Sort_lbl.ForeColor = CLR_DKBLU
13530         .Sort_lbl.Top = lngSortLbl_Top
13540         .Sort_line.Left = .Recur_Name_lbl.Left
13550         .Sort_line.Width = (.Recur_Name_lbl.Width + lngTpp)
              'If .Sort_line.Visible = False Then
              '  .Sort_line.Visible = True
              '  .RecurringItemAux_1099_Sort_line1.Visible = False
              '  .RecurringItemAux_1099_Sort_line2.Visible = False
              'End If
13560       End If
13570     End If
13580   End With

EXITP:
13590   Exit Sub

ERRH:
13600   Select Case ERR.Number
        Case Else
13610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13620   End Select
13630   Resume EXITP

End Sub

Public Sub EnableCountry(blnEnable As Boolean)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "EnableCountry"

        Dim lngTmp01 As Long, lngTmp02 As Long

13710   With Me
13720     Select Case blnEnable
          Case True
13730       .Width = lngForm_Width
            ' ** First move form 1099 right.
13740       .RecurringItemAux_TaxNo.Left = lngTaxID_Left
13750       .RecurringItemAux_TaxNo_box.Left = lngTaxID_Left
13760       .RecurringItemAux_TaxNo_locked.Left = (lngTaxID_Left - lngTpp)
13770       .RecurringItemAux_TaxNo_lbl.Left = lngTaxID_Left
13780       .RecurringItemAux_TaxNo_lbl_dim_hi.Left = (lngTaxID_Left + lngTpp)
13790       .RecurringItemAux_TaxNo_lbl_line.Left = lngTaxID_Left
13800       .RecurringItemAux_TaxNo_lbl_line_dim_hi.Left = (lngTaxID_Left + lngTpp)
13810       .RecurringItemAux_1099.Left = (lngForm1099_Left + lngForm1099_Diff)
13820       .RecurringItemAux_1099_lbl.Left = lngForm1099_Left
13830       .RecurringItemAux_1099_lbl_dim_hi.Left = (lngForm1099_Left + lngTpp)
13840       .RecurringItemAux_1099_lbl2.Left = lngForm1099_Left
13850       .RecurringItemAux_1099_lbl2_dim_hi.Left = (lngForm1099_Left + lngTpp)
13860       .RecurringItemAux_1099_lbl_line.Left = lngForm1099_Left
13870       .RecurringItemAux_1099_lbl_line_dim_hi.Left = (lngForm1099_Left + lngTpp)
13880       .RecurringItemAux_1099_Sort_line1.Left = lngForm1099_Left
13890       .RecurringItemAux_1099_Sort_line2.Left = lngForm1099_Left
            ' ** Then move country into place and show it.
13900       .Recur_PostalCode.Left = lngPostalCode_Left
13910       .Recur_PostalCode_box.Left = lngPostalCode_Left
13920       .Recur_PostalCode_locked.Left = (lngPostalCode_Left - lngTpp)
13930       .Recur_PostalCode_lbl.Left = lngPostalCode_Left
13940       .Recur_PostalCode_lbl_dim_hi.Left = (lngPostalCode_Left + lngTpp)
13950       .Recur_PostalCode_lbl_line.Left = lngPostalCode_Left
13960       .Recur_PostalCode_lbl_line_dim_hi.Left = (lngPostalCode_Left + lngTpp)
13970       .Recur_PostalCode_box.Visible = True
13980       .Recur_PostalCode_locked.Visible = True
13990       .Recur_PostalCode.Visible = True
14000       .Recur_PostalCode_lbl.Visible = True
14010       .Recur_PostalCode_lbl_dim_hi.Visible = False
14020       .Recur_PostalCode_lbl_line.Visible = True
14030       .Recur_PostalCode_lbl_line_dim_hi.Visible = False
14040       .Recur_PostalCode.Enabled = True
14050       .Recur_Country.Left = lngCountry_Left
14060       .Recur_Country_box.Left = lngCountry_Left
14070       .Recur_Country_locked.Left = (lngCountry_Left - lngTpp)
14080       .Recur_Country_lbl.Left = lngCountry_Left
14090       .Recur_Country_lbl_dim_hi.Left = (lngCountry_Left + lngTpp)
14100       .Recur_Country_lbl_line.Left = lngCountry_Left
14110       .Recur_Country_lbl_line_dim_hi.Left = (lngCountry_Left + lngTpp)
14120       .Recur_Country_box.Visible = True
14130       .Recur_Country_locked.Visible = True
14140       .Recur_Country.Visible = True
14150       .Recur_Country_lbl.Visible = True
14160       .Recur_Country_lbl_dim_hi.Visible = False
14170       .Recur_Country_lbl_line.Visible = True
14180       .Recur_Country_lbl_line_dim_hi.Visible = False
14190       .Recur_Country.Enabled = True
14200     Case False
            ' ** First move country and postal code out of the way.
14210       .Recur_Country.Enabled = False
14220       .Recur_Country.Visible = False
14230       .Recur_Country.Left = 0&
14240       .Recur_Country_box.Visible = False
14250       .Recur_Country_box.Left = 0&
14260       .Recur_Country_locked.Visible = False
14270       .Recur_Country_locked.Left = 0&
14280       .Recur_Country_lbl.Visible = False
14290       .Recur_Country_lbl.Left = 0&
14300       .Recur_Country_lbl_dim_hi.Visible = False
14310       .Recur_Country_lbl_dim_hi.Left = 0&
14320       .Recur_Country_lbl_line.Visible = False
14330       .Recur_Country_lbl_line.Left = 0&
14340       .Recur_Country_lbl_line_dim_hi.Visible = False
14350       .Recur_Country_lbl_line_dim_hi.Left = 0&
14360       .Recur_PostalCode.Enabled = False
14370       .Recur_PostalCode.Visible = False
14380       .Recur_PostalCode.Left = 0&
14390       .Recur_PostalCode_box.Visible = False
14400       .Recur_PostalCode_box.Left = 0&
14410       .Recur_PostalCode_locked.Visible = False
14420       .Recur_PostalCode_locked.Left = 0&
14430       .Recur_PostalCode_lbl.Visible = False
14440       .Recur_PostalCode_lbl.Left = 0&
14450       .Recur_PostalCode_lbl_dim_hi.Visible = False
14460       .Recur_PostalCode_lbl_dim_hi.Left = 0&
14470       .Recur_PostalCode_lbl_line.Visible = False
14480       .Recur_PostalCode_lbl_line.Left = 0&
14490       .Recur_PostalCode_lbl_line_dim_hi.Visible = False
14500       .Recur_PostalCode_lbl_line_dim_hi.Left = 0
            ' ** Then move form 1099 left.
14510       lngTmp01 = (lngForm1099_Left - lngCountry_Left)  ' ** Difference.
14520       lngTmp02 = (lngForm1099_Left - lngTmp01)  ' ** New left.
14530       .RecurringItemAux_1099.Left = (lngTmp02 + lngForm1099_Diff)
14540       .RecurringItemAux_1099_lbl.Left = lngTmp02
14550       .RecurringItemAux_1099_lbl_dim_hi.Left = (lngTmp02 + lngTpp)
14560       .RecurringItemAux_1099_lbl2.Left = lngTmp02
14570       .RecurringItemAux_1099_lbl2_dim_hi.Left = (lngTmp02 + lngTpp)
14580       .RecurringItemAux_1099_lbl_line.Left = lngTmp02
14590       .RecurringItemAux_1099_lbl_line_dim_hi.Left = (lngTmp02 + lngTpp)
14600       .RecurringItemAux_1099_Sort_line1.Left = lngTmp02
14610       .RecurringItemAux_1099_Sort_line2.Left = lngTmp02
14620       lngTmp02 = (lngTaxID_Left - lngTmp01)  ' ** New left.
14630       .RecurringItemAux_TaxNo.Left = lngTmp02
14640       .RecurringItemAux_TaxNo_box.Left = lngTmp02
14650       .RecurringItemAux_TaxNo_locked.Left = (lngTmp02 - lngTpp)
14660       .RecurringItemAux_TaxNo_lbl.Left = lngTmp02
14670       .RecurringItemAux_TaxNo_lbl_dim_hi.Left = (lngTmp02 + lngTpp)
14680       .RecurringItemAux_TaxNo_lbl_line.Left = lngTmp02
14690       .RecurringItemAux_TaxNo_lbl_line_dim_hi.Left = (lngTmp02 + lngTpp)
14700       lngTmp02 = (lngForm_Width - lngTmp01)  ' ** New width.
14710       .Width = lngTmp02
14720     End Select
14730     If InStr(strSortNow, "Country") > 0 Then
14740       Select Case blnEnable
            Case True
14750         .Sort_line.Left = .Recur_Country_lbl.Left
14760         .Sort_lbl.Left = ((.Recur_Country_lbl.Left + .Recur_Country_lbl.Width) - lngSortLbl_Width)
14770       Case False
14780         SortNow "Form_Load"  ' ** Procedure: Below.
14790       End Select
14800     ElseIf InStr(strSortNow, "PostalCode") > 0 Then
14810       Select Case blnEnable
            Case True
14820         .Sort_line.Left = .Recur_PostalCode_lbl.Left
14830         .Sort_lbl.Left = ((.Recur_PostalCode_lbl.Left + .Recur_PostalCode_lbl.Width) - lngSortLbl_Width)
14840       Case False
14850         SortNow "Form_Load"  ' ** Procedure: Below.
14860       End Select
14870     ElseIf InStr(strSortNow, "Aux_1099") > 0 Then
14880       .Sort_line.Left = .RecurringItemAux_1099_lbl.Left
14890       .Sort_lbl.Left = (((.RecurringItemAux_1099_lbl.Left + .RecurringItemAux_1099_lbl.Width) - lngSortLbl_Width) + (5& * lngTpp))
14900     ElseIf InStr(strSortNow, "Aux_TaxNo") > 0 Then
14910       .Sort_line.Left = .RecurringItemAux_TaxNo_lbl.Left
14920       .Sort_lbl.Left = ((.RecurringItemAux_TaxNo_lbl.Left + .RecurringItemAux_TaxNo_lbl.Width) - lngSortLbl_Width)
14930     End If
14940     DoEvents
14950   End With

EXITP:
14960   Exit Sub

ERRH:
14970   Select Case ERR.Number
        Case 2100  ' ** The control or subform control is too large for this location.
          ' ** Confusion.
14980   Case Else
14990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15000   End Select
15010   Resume EXITP

End Sub

Public Sub RecurHeaders(blnEnable As Boolean)

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurHeaders"

15110   With Me
15120     Select Case blnEnable
          Case True
15130       .Recur_Address_lbl.ForeColor = CLR_DKGRY2
15140       .Recur_Address_lbl_dim_hi.Visible = False
15150       .Recur_Address_lbl_line.BorderColor = CLR_DKGRY
15160       .Recur_Address_lbl_line_dim_hi.Visible = False
15170       .Recur_City_lbl.ForeColor = CLR_DKGRY2
15180       .Recur_City_lbl_dim_hi.Visible = False
15190       .Recur_City_lbl_line.BorderColor = CLR_DKGRY
15200       .Recur_City_lbl_line_dim_hi.Visible = False
15210       .Recur_State_lbl.ForeColor = CLR_DKGRY2
15220       .Recur_State_lbl_dim_hi.Visible = False
15230       .Recur_State_lbl_line.BorderColor = CLR_DKGRY
15240       .Recur_State_lbl_line_dim_hi.Visible = False
15250       .Recur_Zip_lbl.ForeColor = CLR_DKGRY2
15260       .Recur_Zip_lbl_dim_hi.Visible = False
15270       .Recur_Zip_lbl_line.BorderColor = CLR_DKGRY
15280       .Recur_Zip_lbl_line_dim_hi.Visible = False
15290       If .Parent.chkEnableCountry = True Then
15300         .Recur_Country_lbl.ForeColor = CLR_DKGRY2
15310         .Recur_Country_lbl_dim_hi.Visible = False
15320         .Recur_Country_lbl_line.BorderColor = CLR_DKGRY
15330         .Recur_Country_lbl_line_dim_hi.Visible = False
15340         .Recur_PostalCode_lbl.ForeColor = CLR_DKGRY2
15350         .Recur_PostalCode_lbl_dim_hi.Visible = False
15360         .Recur_PostalCode_lbl_line.BorderColor = CLR_DKGRY
15370         .Recur_PostalCode_lbl_line_dim_hi.Visible = False
15380       End If
            '.RecurringItemAux_1099_lbl.ForeColor = CLR_DKGRY2
            '.RecurringItemAux_1099_lbl_dim_hi.Visible = False
            '.RecurringItemAux_1099_lbl2.ForeColor = CLR_DKGRY2
            '.RecurringItemAux_1099_lbl2_dim_hi.Visible = False
            '.RecurringItemAux_1099_lbl_line.BorderColor = CLR_DKGRY
            '.RecurringItemAux_1099_lbl_line_dim_hi.Visible = False
            '.RecurringItemAux_TaxNo_lbl.ForeColor = CLR_DKGRY2
            '.RecurringItemAux_TaxNo_lbl_dim_hi.Visible = False
            '.RecurringItemAux_TaxNo_lbl_line.BorderColor = CLR_DKGRY
            '.RecurringItemAux_TaxNo_lbl_line_dim_hi.Visible = False
15390     Case False
15400       .Recur_Address_lbl.ForeColor = WIN_CLR_DISF
15410       .Recur_Address_lbl_dim_hi.Visible = True
15420       .Recur_Address_lbl_line.BorderColor = WIN_CLR_DISR
15430       .Recur_Address_lbl_line_dim_hi.Visible = True
15440       .Recur_City_lbl.ForeColor = WIN_CLR_DISF
15450       .Recur_City_lbl_dim_hi.Visible = True
15460       .Recur_City_lbl_line.BorderColor = WIN_CLR_DISR
15470       .Recur_City_lbl_line_dim_hi.Visible = True
15480       .Recur_State_lbl.ForeColor = WIN_CLR_DISF
15490       .Recur_State_lbl_dim_hi.Visible = True
15500       .Recur_State_lbl_line.BorderColor = WIN_CLR_DISR
15510       .Recur_State_lbl_line_dim_hi.Visible = True
15520       .Recur_Zip_lbl.ForeColor = WIN_CLR_DISF
15530       .Recur_Zip_lbl_dim_hi.Visible = True
15540       .Recur_Zip_lbl_line.BorderColor = WIN_CLR_DISR
15550       .Recur_Zip_lbl_line_dim_hi.Visible = True
15560       If .Parent.chkEnableCountry = True Then
15570         .Recur_Country_lbl.ForeColor = WIN_CLR_DISF
15580         .Recur_Country_lbl_dim_hi.Visible = True
15590         .Recur_Country_lbl_line.BorderColor = WIN_CLR_DISR
15600         .Recur_Country_lbl_line_dim_hi.Visible = True
15610         .Recur_PostalCode_lbl.ForeColor = WIN_CLR_DISF
15620         .Recur_PostalCode_lbl_dim_hi.Visible = True
15630         .Recur_PostalCode_lbl_line.BorderColor = WIN_CLR_DISR
15640         .Recur_PostalCode_lbl_line_dim_hi.Visible = True
15650       End If
            '.RecurringItemAux_1099_lbl.ForeColor = WIN_CLR_DISF
            '.RecurringItemAux_1099_lbl_dim_hi.Visible = True
            '.RecurringItemAux_1099_lbl2.ForeColor = WIN_CLR_DISF
            '.RecurringItemAux_1099_lbl2_dim_hi.Visible = True
            '.RecurringItemAux_1099_lbl_line.BorderColor = WIN_CLR_DISR
            '.RecurringItemAux_1099_lbl_line_dim_hi.Visible = True
            '.RecurringItemAux_TaxNo_lbl.ForeColor = WIN_CLR_DISF
            '.RecurringItemAux_TaxNo_lbl_dim_hi.Visible = True
            '.RecurringItemAux_TaxNo_lbl_line.BorderColor = WIN_CLR_DISR
            '.RecurringItemAux_TaxNo_lbl_line_dim_hi.Visible = True
15660     End Select
15670   End With

EXITP:
15680   Exit Sub

ERRH:
15690   Select Case ERR.Number
        Case Else
15700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15710   End Select
15720   Resume EXITP

End Sub

Public Sub DeleteRec(lngRecurID As Long)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "DeleteRec"

15810   With Me
15820     DoCmd.SelectObject acForm, "frmRecurringItems", False
15830     Forms("frmRecurringItems").frmRecurringItems_Sub.SetFocus
15840     If .RecurringItem_ID = lngRecurID Then
15850       If .NewRecord Then
15860         MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Above.
15870       Else
15880         .AllowDeletions = True
15890         DoCmd.RunCommand acCmdDeleteRecord
15900         .AllowDeletions = False
15910       End If
15920     End If
15930   End With

EXITP:
15940   Exit Sub

ERRH:
15950   Select Case ERR.Number
        Case Else
15960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15970   End Select
15980   Resume EXITP

End Sub

Public Function RecCnt() As Long

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

16010   lngRetVal = 0&

16020   With Me
16030     Set rst = .RecordsetClone
16040     With rst
16050       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
16060       Else
16070         .MoveLast
16080         lngRetVal = .RecordCount
16090       End If
16100       .Close
16110     End With
16120   End With

EXITP:
16130   Set rst = Nothing
16140   RecCnt = lngRetVal
16150   Exit Function

ERRH:
16160   lngRetVal = 0&
16170   Select Case ERR.Number
        Case Else
16180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16190   End Select
16200   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

16310   Select Case intWhere
        Case acCmdRecordsGoToFirst
16320     DoCmd.RunCommand acCmdRecordsGoToFirst
16330   Case acCmdRecordsGoToLast
16340     DoCmd.RunCommand acCmdRecordsGoToLast
16350   Case acCmdRecordsGoToPrevious
16360     DoCmd.RunCommand acCmdRecordsGoToPrevious
16370   Case acCmdRecordsGoToNext
16380     DoCmd.RunCommand acCmdRecordsGoToNext
16390   Case Else
16400     If IsMissing(varID) = False Then
16410       With Me
16420         Set rst = .RecordsetClone
16430         With rst
16440           .FindFirst "[RecurringItem_ID] = " & CStr(varID)
16450           If .NoMatch = False Then
16460             Me.Bookmark = .Bookmark
16470           End If
16480           .Close
16490         End With
16500       End With
16510     End If
16520   End Select

EXITP:
16530   Set rst = Nothing
16540   Exit Sub

ERRH:
16550   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
16560   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
16570   Case Else
16580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16590   End Select
16600   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [Recur_Name], [Recur_Address]"

16710   If lngTpp = 0& Then
          'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
16720     lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
16730   End If

16740   With Me
16750     .Controls(strSortLbl).Top = lngSortLbl_Top
16760     .Controls(strSortLbl).Visible = False
16770     .Controls(strSortLine).Visible = False
16780     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
16790     .RecurringItemAux_1099_Sort_line1.Visible = False
16800     .RecurringItemAux_1099_Sort_line2.Visible = False
16810     If strProc = "Form_Load" Then
16820       strCalled = "Recur_Name"
16830       strSortNow = strSortOrig
16840       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
16850       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
16860       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
16870       .Controls(strSortLbl).Left = lngSortLbl_Left
16880       .Controls(strSortLbl).Caption = strArwUp
16890       .Controls(strSortLbl).ForeColor = CLR_DKBLU
16900       .Controls(strSortLine).Left = lngSortLine_Left
16910       .Controls(strSortLine).Width = lngSortLine_Width
16920     Else
16930       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
16940       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
16950       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
16960       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
16970       .Controls(strSortLbl).Left = lngSortLbl_Left
16980       .Controls(strSortLine).Left = lngSortLine_Left
16990       .Controls(strSortLine).Width = lngSortLine_Width
17000       If strCalled = "Recur_Name" Then
17010         strSortAsc = strSortOrig
17020         If strSortNow = strSortAsc Then
17030           strSortNow = "[Recur_Name] DESC, [Recur_Address]"
17040           .Controls(strSortLbl).Caption = strArwDn
17050           .Controls(strSortLbl).ForeColor = CLR_DKRED
17060         Else
17070           strSortNow = strSortAsc
17080           .Controls(strSortLbl).Caption = strArwUp
17090           .Controls(strSortLbl).ForeColor = CLR_DKBLU
17100         End If
17110       Else
17120         .Controls(strSortLbl).Caption = strArwUp
17130         .Controls(strSortLbl).ForeColor = CLR_DKBLU
17140         Select Case strCalled
              Case "Recur_Type"
17150           strSortAsc = "[Recur_Type], [Recur_Name]"
17160         Case "Recur_Address"
17170           strSortAsc = "[Recur_Address], [Recur_Name]"
17180         Case "Recur_City"
17190           strSortAsc = "[Recur_City]" & strStdAsc
17200         Case "Recur_State"
17210           strSortAsc = "[Recur_State]" & strStdAsc
17220         Case "Recur_Zip"
17230           strSortAsc = "[Recur_Zip]" & strStdAsc
17240         Case "Recur_Country"
17250           strSortAsc = "[Recur_Country]" & strStdAsc
17260         Case "Recur_PostalCode"
17270           strSortAsc = "[Recur_PostalCode]" & strStdAsc
17280         Case "RecurringItemAux_1099"
17290           strSortAsc = "[RecurringItemAux_1099]" & strStdAsc
17300         Case "RecurringItemAux_TaxNo"
17310           strSortAsc = "[RecurringItemAux_TaxNo]" & strStdAsc
17320         End Select
17330         If strSortNow = strSortAsc Then
17340           intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
17350           Select Case intCnt
                Case 1
17360             strTmp01 = strSortAsc & " DESC"
17370           Case 2
17380             intPos01 = InStr(strSortAsc, ",")
17390             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
17400             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
17410           Case 3
17420             intPos01 = InStr(strSortAsc, ",")
17430             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
17440             strTmp02 = Mid(strSortAsc, intPos01)
17450             intPos01 = InStr(2, strTmp02, ",")
17460             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
17470             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
17480           Case 4
17490             intPos01 = InStr(strSortAsc, ",")
17500             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
17510             strTmp02 = Mid(strSortAsc, intPos01)
17520             intPos01 = InStr(2, strTmp02, ",")
17530             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
17540             strTmp02 = Mid(strTmp02, intPos01)
17550             intPos01 = InStr(2, strTmp02, ",")
17560             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
17570             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
17580           End Select
17590           strSortNow = strTmp01
17600           .Controls(strSortLbl).Caption = strArwDn
17610           .Controls(strSortLbl).ForeColor = CLR_DKRED
17620         Else
17630           strSortNow = strSortAsc
17640         End If
17650       End If
17660     End If
17670     If strCalled = "RecurringItemAux_1099" Then
17680       .RecurringItemAux_1099_Sort_line1.Visible = True
17690       .RecurringItemAux_1099_Sort_line2.Visible = True
17700       .Controls(strSortLbl).Top = (.Controls(strSortLbl).Top - (4& * lngTpp))
17710       .Controls(strSortLbl).Left = (.Controls(strSortLbl).Left + (5& * lngTpp))
17720       .Controls(strSortLbl).Visible = True
17730     Else
17740       .Controls(strSortLbl).Visible = True
17750       .Controls(strSortLine).Visible = True
17760     End If
17770     .OrderBy = strSortNow
17780     .OrderByOn = True
17790   End With

        ' ** Potential Sorts:
        ' **   [Recur_Name], [Recur_Address]
        ' **   [Recur_Type], [Recur_Name]
        ' **   [Recur_Address], [Recur_Name]
        ' **   [Recur_City], [Recur_Name], [Recur_Address]
        ' **   [Recur_State], [Recur_Name], [Recur_Address]
        ' **   [Recur_Zip], [Recur_Name], [Recur_Address]
        ' **   [Recur_Country], [Recur_Name], [Recur_Address]
        ' **   [Recur_PostalCode], [Recur_Name], [Recur_Address]
        ' **   [RecurringItemAux_1099], [Recur_Name], [Recur_Address]
        ' **   [RecurringItemAux_TaxNo], [Recur_Name], [Recur_Address]

EXITP:
17800   Exit Sub

ERRH:
17810   Select Case ERR.Number
        Case Else
17820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17830   End Select
17840   Resume EXITP

End Sub

Public Function SortNow_Get() As String

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

17910   strRetVal = strSortNow

EXITP:
17920   SortNow_Get = strRetVal
17930   Exit Function

ERRH:
17940   strRetVal = vbNullString
17950   Select Case ERR.Number
        Case Else
17960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17970   End Select
17980   Resume EXITP

End Function
