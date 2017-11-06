VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountHideTrans2_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountHideTrans2_Select"

'VGC 10/02/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Auto Reverse:     A {opgHide_optOne}
' **   Cancel:           C {cmdCancel}
' **   OK:               O {cmdOK}
' **   Select Hide:      S {opgHide_optChoose}

Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long, blnWasGTR As Boolean
Private lngHids As Long, lngTotTrans_L As Long, lngTotTrans_LA As Long
Private blnJustClose As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim strTmp01 As String
        Dim intPos01 As Integer

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strTmp01 = Trim(.OpenArgs)
150           intPos01 = InStr(strTmp01, "~")
160           strCallingForm = Left(strTmp01, (intPos01 - 1))
170           strTmp01 = Mid(strTmp01, (intPos01 + 1))
180           intPos01 = InStr(strTmp01, "~")
190           lngTotTrans_L = CLng(Left(strTmp01, (intPos01 - 1)))
200           lngTotTrans_LA = CLng(Mid(strTmp01, (intPos01 + 1)))

210           blnJustClose = False
220           blnWasGTR = gblnGoToReport
230           blnGTR_Emblem = False: lngGTR_Stat = 0&

              ' ** gstrAccountNo and gstrAccountName should already be populated.

240           .PictureData = .Form_HideTrans2_Select_img

250           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

260           lngHids = Forms(strCallingForm).hidden_trans

270           opgHide_AfterUpdate  ' ** Procedure: Below.

280           .cmdOK.SetFocus

290         Else
300           Cancel = -1
310         End If
320       Else
330         Cancel = -1
340       End If
350     End With

360     DoCmd.Hourglass False

370     If Cancel = -1 Then
380       Beep
390       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
400     End If

EXITP:
410     Exit Sub

ERRH:
420     DoCmd.Hourglass False
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
570         intRetVal = 0
580         cmdCancel_Click  ' ** Procedure: Below.
590       End Select
600     End If

        ' ** Ctrl-Shift keys.
610     If intCtrlDown And (Not intAltDown) And intShiftDown Then
620       Select Case intRetVal
          Case vbKeyF
630         With Me
640           intRetVal = 0
650           .FocusHolder.SetFocus
660         End With
670       End Select
680     End If

EXITP:
690     KeyCode = intRetVal
700     Exit Sub

ERRH:
710     intRetVal = 0
720     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
730     Case Else
740       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
750     End Select
760     Resume EXITP

End Sub

Private Sub Form_Timer()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Static lngTmp01 As Long, lngTmp02 As Long

810     With Me
820       .TimerInterval = 0&
830       If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
840         lngGTR_Stat = lngGTR_Stat + 1&
850         If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
860         lngTmp01 = lngTmp01 + 1&
870         .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
880         If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
890         If lngTmp01 > 1& Then
900           .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
910         End If
920         If lngTmp01 = 24& Then
930           lngTmp01 = 0&
940           lngTmp02 = lngTmp02 + 1&
950           If lngTmp02 < 8& Then
960             .TimerInterval = 50&
970           Else
980             blnGTR_Emblem = False
990           End If
1000        Else
1010          .TimerInterval = 50&
1020        End If
1030      End If  ' ** blnGTR_Emblem.
1040      If gblnGoToReport = True Then
1050        DoCmd.Hourglass True  ' ** Make sure it's still running.
1060        If ((lngGTR_Stat > 0& And lngGTR_Stat < 24&) Or (lngGTR_Stat > 24& And lngGTR_Stat < 48&) Or _
                (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
1070        Else
1080          DoCmd.Hourglass True  ' ** Make sure it's still running.
1090          If blnGTR_Emblem = False Then
1100            blnGTR_Emblem = True
1110            lngTmp01 = 0&: lngTmp02 = 0&
1120          End If
1130          DoEvents
1140          Select Case blnGoingToReport
              Case True
1150            DoCmd.Hourglass True
1160            If .GoToReport_arw_hideblu_ok_img.Visible = True Then
1170              .TimerInterval = 50&
1180              cmdOK_Click  ' ** Procedure: Below.
1190            Else
1200              .cmdOK.SetFocus
1210              .GoToReport_arw_hideblu_sel_img.Visible = False
1220              .GoToReport_arw_hideblu_ok_img.Visible = True
1230              DoEvents
1240              .TimerInterval = 50&
1250            End If
1260          Case False
1270            blnGoingToReport = True
1280            .opgHide.SetFocus
1290            .GoToReport_arw_hideblu_sel_img.Visible = True
1300            If .opgHide <> .opgHide_optChoose.OptionValue Then
1310              .opgHide = .opgHide_optChoose.OptionValue
1320              opgHide_AfterUpdate  ' ** Procedure: Below.
1330            End If
1340            DoEvents
1350            Forms(strCallingForm).GTREmblem_Off  ' ** Form Procedure: frmMenu_Accounrt.
1360            DoEvents
1370            .TimerInterval = 50&
1380          End Select
1390        End If
1400      End If
1410    End With

EXITP:
1420    Exit Sub

ERRH:
1430    DoCmd.Hourglass False
1440    Select Case ERR.Number
        Case Else
1450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1460    End Select
1470    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1510    If blnWasGTR = False Then
1520      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
1530    End If

1540    If blnJustClose = False Then
1550      gstrReturningForm = THIS_NAME
1560      gblnSetFocus = True
1570      Select Case strCallingForm
          Case "frmMenu_Account"
1580        If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
1590          Forms(strCallingForm).TimerInterval = 200&
1600          DoCmd.SelectObject acForm, strCallingForm, False
1610        Else
1620          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
1630        End If
1640      End Select
1650    End If

EXITP:
1660    Exit Sub

ERRH:
1670    DoCmd.Hourglass False
1680    Select Case ERR.Number
        Case Else
1690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1700    End Select
1710    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String, strEarliest As String, strHasCurrent As String
        Dim lngThisDbsID As Long, lngTotHid As Long
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant

1810    With Me

1820      DoCmd.Hourglass True
1830      DoEvents

1840      msgResponse = vbOK
1850      If .opgHide = .opgHide_optChoose.OptionValue And lngHids > 0& Then
1860        If Pref_GetBln("frmAccountHideTrans2", "chkHiddenFirstUse") = False Then  ' ** Module Function: modPreferenceFuncs.
              ' ** qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a (Ledger,
              ' ** just needed fields), qryAccountHideTrans2_24b (LedgerArchive,
              ' ** just just needed fields)), just ledger_HIDDEN = True.
1870          varTmp00 = DCount("*", "qryAccountHideTrans2_25")
1880          lngTotHid = Nz(varTmp00, lngHids)
              ' ** Since the first use will collect all the hidden transactions,
              ' ** give them a message if it might take a while.
1890          If lngTotHid > 50& Then  ' ** Arbitrary.
1900            DoCmd.Hourglass False
1910            msgResponse = MsgBox("If you have recently upgraded Trust Accountant, first use of this feature" & vbCrLf & _
                  "may require several minutes to collate existing hidden transactions." & vbCrLf & vbCrLf & _
                  "Do you wish to continue at this time?", vbQuestion + vbOKCancel, "Hide Transactions First Use")
1920          End If
1930        End If
1940      End If

1950      If msgResponse = vbOK Then

1960        DoCmd.Hourglass True
1970        DoEvents

1980        lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

1990        Set dbs = CurrentDb

2000        Select Case .opgHide
            Case .opgHide_optOne.OptionValue

2010          strDocName = "frmAccountHideTrans2_One"
2020          gstrFormQuerySpec = strDocName

2030          With dbs

                ' ** Empty tblLedgerHidden_Staging1.
2040            Set qdf = .QueryDefs("qryAccountHideTrans2_One_Sub_List_04")
2050            qdf.Execute
2060            Set qdf = Nothing
2070            DoEvents
                ' ** Empty tblLedgerHidden_Staging2.
2080            Set qdf = .QueryDefs("qryAccountHideTrans2_One_Sub_Pick_03")
2090            qdf.Execute
2100            Set qdf = Nothing
2110            DoEvents

                ' ** Append qryAccountHideTrans2_One_Sub_List_03 (qryAccountHideTrans2_One_Sub_List_01
                ' ** (Union of qryAccountHideTrans2_One_Sub_List_01_01 (Ledger, just ledger_HIDDEN = False,
                ' ** by specified FormRef('AcctNum')), qryAccountHideTrans2_One_Sub_List_01_02 (LedgerArchive,
                ' ** just ledger_HIDDEN = False, by specified FormRef('AcctNum'))), linked to tblJournalType,
                ' ** with shareface_0 - shareface_4) to tblLedgerHidden_Staging1.
2120            Set qdf = .QueryDefs("qryAccountHideTrans2_One_Sub_List_05")
2130            qdf.Execute
2140            Set qdf = Nothing
2150            DoEvents

                ' ** Balance, grouped by accountno, with MaxBD, by specified [actno].
2160            Set qdf = .QueryDefs("qryAccountHideTrans2_30")
2170            With qdf.Parameters
2180              ![actno] = gstrAccountNo
2190            End With
2200            Set rst = qdf.OpenRecordset
2210            With rst
2220              If .BOF = True And .EOF = True Then  ' ** No balances.
2230                strEarliest = "0"
2240              Else
2250                .MoveFirst
2260                If IsNull(![MaxBD]) Then  ' ** No balance.
2270                  strEarliest = "0"
2280                Else
2290                  strEarliest = Format$(![MaxBD], "mm/dd/yyyy")  ' ** Earliest, as in start of current statement period. (Which would be +1 day)
2300                End If
2310              End If
2320              .Close
2330            End With  ' ** rst.
2340            Set rst = Nothing
2350            Set qdf = Nothing
2360            DoEvents

                ' ** VGC 07/08/2009: Per Rich, always allow them in, and allow older with warning.
2370            strHasCurrent = "True"
                ' ** tmpAccountHideTrans_02, just entries after last balance date, by specified [tdat].
2380            Set qdf = .QueryDefs("qryAccountHideTrans2_31")
2390            With qdf.Parameters
2400              ![tdat] = CDate(strEarliest)  ' ** Criteria is: >[tdat]
2410            End With
2420            Set rst = qdf.OpenRecordset
2430            With rst
2440              If .BOF = True And .EOF = True Then
2450                strHasCurrent = "False"
2460              End If
2470              .Close
2480            End With  ' ** rst.
2490            Set rst = Nothing
2500            Set qdf = Nothing
2510            DoEvents

2520            .Close
2530          End With  ' ** dbs.
2540          Set dbs = Nothing
2550          DoEvents

2560          Forms(strCallingForm).JustClose_Set True  ' ** Form Procedure: frmMenu_Account.

2570          gblnSetFocus = True
2580          blnJustClose = True
2590          DoCmd.Close acForm, strCallingForm
2600          DoCmd.Close acForm, THIS_NAME
2610          DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & strCallingForm & "~" & gstrAccountNo & "~" & _
                gstrAccountName & "~" & strHasCurrent & "~" & strEarliest

2620        Case .opgHide_optChoose.OptionValue

2630          strDocName = "frmAccountHideTrans2"
2640          gstrFormQuerySpec = strDocName

2650          With dbs

                ' ** Empty tmpAccountHideTrans_01.
2660            Set qdf = .QueryDefs("qryAccountHideTrans2_01a")
2670            qdf.Execute
2680            Set qdf = Nothing
2690            DoEvents
                ' ** Empty tmpAccountHideTrans_02.
2700            Set qdf = .QueryDefs("qryAccountHideTrans2_01b")
2710            qdf.Execute
2720            Set qdf = Nothing
2730            DoEvents
                ' ** Empty tmpAccountHideTrans_03.
2740            Set qdf = .QueryDefs("qryAccountHideTrans2_01c")
2750            qdf.Execute
2760            Set qdf = Nothing
2770            DoEvents
                ' ** Empty tmpAccountHideTrans_04.
2780            Set qdf = .QueryDefs("qryAccountHideTrans2_01d")
2790            qdf.Execute
2800            Set qdf = Nothing
2810            DoEvents

                ' ** Append qryAccountHideTrans2_10 (qryAccountHideTrans2_08 (Account, linked to qryAccountHideTrans2_05
                ' ** (LedgerArchive, grouped by accountno, with cnt, by specified [actno]), qryAccountHideTrans2_06
                ' ** (Union of qryAccountHideTrans2_06a (Ledger, grouped by accountno, just ledger_HIDDEN = True, with cnt,
                ' ** by specified [actno]), qryAccountHideTrans2_06b (LedgerArchive, grouped by accountno, just ledger_HIDDEN = True,
                ' ** with cnt, by specified [actno])), by specified [actno]), linked to qryAccountHideTrans2_09 (tblForm_Graphics,
                ' ** for frmAccountHideTrans2, by specified [dbid]), with add'l fields) to tmpAccountHideTrans_01.
2820            Set qdf = .QueryDefs("qryAccountHideTrans2_11")
2830            With qdf.Parameters
2840              ![dbid] = lngThisDbsID
2850              ![actno] = gstrAccountNo
2860            End With
2870            qdf.Execute
2880            Set qdf = Nothing
2890            DoEvents
                ' ** tmpAccountHideTrans_01 has:
                ' **   hidden_trans: Total number of hidden in both Ledger and LedgerArchive.
                ' **   archive_trans: Total number of all transactions in LedgerArchive.

                ' ** Append qryAccountHideTrans2_12 (Union of qryAccountHideTrans2_12a (Ledger, by specified [actno]),
                ' ** qryAccountHideTrans2_12b (LedgerArchive, by specified [actno])) to tmpAccountHideTrans_02.
2900            Set qdf = .QueryDefs("qryAccountHideTrans2_13")
2910            With qdf.Parameters
2920              ![actno] = gstrAccountNo
2930            End With
2940            qdf.Execute
2950            Set qdf = Nothing
2960            DoEvents

                ' ** Balance, grouped by accountno, with MaxBD, by specified [actno].
2970            Set qdf = .QueryDefs("qryAccountHideTrans2_30")
2980            With qdf.Parameters
2990              ![actno] = gstrAccountNo
3000            End With
3010            Set rst = qdf.OpenRecordset
3020            With rst
3030              If .BOF = True And .EOF = True Then  ' ** No balances.
3040                strEarliest = vbNullString
3050              Else
3060                .MoveFirst
3070                If IsNull(![MaxBD]) Then  ' ** No balance.
3080                  strEarliest = vbNullString
3090                Else
3100                  strEarliest = Format$(![MaxBD], "mm/dd/yyyy")  ' ** Earliest, as in start of current statement period. (Which would be +1 day)
3110                End If
3120              End If
3130              .Close
3140            End With  ' ** rst.
3150            Set rst = Nothing
3160            Set qdf = Nothing
3170            DoEvents

                ' ** VGC 07/08/2009: Per Rich, always allow them in, and allow older with warning.
3180            strHasCurrent = "True"
                ' ** tmpAccountHideTrans_02, just entries after last balance date, by specified [tdat].
3190            Set qdf = .QueryDefs("qryAccountHideTrans2_31")
3200            With qdf.Parameters
3210              ![tdat] = CDate(strEarliest)  ' ** Criteria is: >[tdat]
3220            End With
3230            Set rst = qdf.OpenRecordset
3240            With rst
3250              If .BOF = True And .EOF = True Then
3260                strHasCurrent = "False"
3270              End If
3280              .Close
3290            End With  ' ** rst.
3300            Set rst = Nothing
3310            Set qdf = Nothing
3320            DoEvents

3330            .Close
3340          End With  ' ** dbs.
3350          Set dbs = Nothing
3360          DoEvents

3370          Forms(strCallingForm).JustClose_Set True  ' ** Form Procedure: frmMenu_Account.

3380          gblnSetFocus = True
3390          blnJustClose = True
3400          DoCmd.Close acForm, strCallingForm
3410          DoCmd.Close acForm, THIS_NAME
3420          DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & strCallingForm & "~" & gstrAccountNo & "~" & _
                strHasCurrent & "~" & CStr(lngHids)
3430          If gblnGoToReport = True Then
3440            Forms(strDocName).TimerInterval = 50&
3450            blnGoingToReport = False
3460          End If

3470        End Select

3480      End If  ' ** msgResponse.

3490    End With

EXITP:
3500    Set rst = Nothing
3510    Set qdf = Nothing
3520    Set dbs = Nothing
3530    Exit Sub

ERRH:
3540    DoCmd.Hourglass False
3550    Select Case ERR.Number
        Case 2450  ' ** Microsoft Access can't find the form 'frmMenu_Account' referred to in a macro expression or Visual Basic code.
          ' ** Ignore.
3560    Case Else
3570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3580    End Select
3590    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
          Case vbKeyTab
3670        With Me
3680          intRetVal = 0
3690          .cmdCancel.SetFocus
3700        End With
3710      End Select
3720    End If

        ' ** Shift keys.
3730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3740      Select Case intRetVal
          Case vbKeyTab
3750        With Me
3760          intRetVal = 0
3770          .opgHide.SetFocus
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

Private Sub cmdCancel_Click()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

3910    blnJustClose = False
3920    DoCmd.Close acForm, THIS_NAME

EXITP:
3930    Exit Sub

ERRH:
3940    Select Case ERR.Number
        Case Else
3950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3960    End Select
3970    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
          Case vbKeyTab
4070        With Me
4080          intRetVal = 0
4090          .opgHide.SetFocus
4100        End With
4110      End Select
4120    End If

        ' ** Shift keys.
4130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4140      Select Case intRetVal
          Case vbKeyTab
4150        With Me
4160          intRetVal = 0
4170          .cmdOK.SetFocus
4180        End With
4190      End Select
4200    End If

EXITP:
4210    KeyCode = intRetVal
4220    Exit Sub

ERRH:
4230    intRetVal = 0
4240    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4250    Case Else
4260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4270    End Select
4280    Resume EXITP

End Sub

Private Sub opgHide_AfterUpdate()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgHide_AfterUpdate"

4310    With Me
4320      Select Case .opgHide
          Case .opgHide_optOne.OptionValue
4330        .opgHide_optOne_lbl.FontBold = True
4340        .opgHide_optChoose_lbl.FontBold = False
4350      Case .opgHide_optChoose.OptionValue
4360        .opgHide_optChoose_lbl.FontBold = True
4370        .opgHide_optOne_lbl.FontBold = False
4380      End Select
4390    End With

EXITP:
4400    Exit Sub

ERRH:
4410    Select Case ERR.Number
        Case Else
4420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4430    End Select
4440    Resume EXITP

End Sub

Private Sub opgHide_optOne_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgHide_optOne_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4520    intShiftDown = (Shift And acShiftMask) > 0
4530    intAltDown = (Shift And acAltMask) > 0
4540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4560      Select Case intRetVal
          Case vbKeyTab
4570        With Me
4580          intRetVal = 0
4590          .cmdOK.SetFocus
4600        End With
4610      End Select
4620    End If

        ' ** Shift keys.
4630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4640      Select Case intRetVal
          Case vbKeyTab
4650        With Me
4660          intRetVal = 0
4670          .cmdCancel.SetFocus
4680        End With
4690      End Select
4700    End If

EXITP:
4710    KeyCode = intRetVal
4720    Exit Sub

ERRH:
4730    intRetVal = 0
4740    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4750    Case Else
4760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4770    End Select
4780    Resume EXITP

End Sub

Private Sub opgHide_optChoose_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgHide_optChoose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4820    intShiftDown = (Shift And acShiftMask) > 0
4830    intAltDown = (Shift And acAltMask) > 0
4840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4860      Select Case intRetVal
          Case vbKeyTab
4870        With Me
4880          intRetVal = 0
4890          .cmdOK.SetFocus
4900        End With
4910      End Select
4920    End If

        ' ** Shift keys.
4930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4940      Select Case intRetVal
          Case vbKeyTab
4950        With Me
4960          intRetVal = 0
4970          .cmdCancel.SetFocus
4980        End With
4990      End Select
5000    End If

EXITP:
5010    KeyCode = intRetVal
5020    Exit Sub

ERRH:
5030    intRetVal = 0
5040    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5050    Case Else
5060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5070    End Select
5080    Resume EXITP

End Sub

Public Sub GTREmblem_Off()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

5110    With Me
5120      blnGTR_Emblem = False
5130      For lngX = 1& To 24&
5140        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
5150      Next
5160    End With

EXITP:
5170    Exit Sub

ERRH:
5180    Select Case ERR.Number
        Case Else
5190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5200    End Select
5210    Resume EXITP

End Sub
