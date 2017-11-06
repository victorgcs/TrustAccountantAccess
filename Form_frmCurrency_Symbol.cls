VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCurrency_Symbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCurrency_Symbol"

'VGC 04/17/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Active Exclude:   E {opgActive_optExclude}
' **   Active Include:   I {opgActive_optInclude}
' **   Country:          N {country_name1 on frmCurrency_Symbol_Sub}
' **   Active Only:      O {opgActive_optOnly}
' **   Print Report:     P {cmdPrintReport}
' **   Currency:         R {curr_word1 on frmCurrency_Symbol_Sub}
' **   Exit:             X {cmdClose}

' ******************************************************
' ** NOTE: Code2000 Font cannot be used in the report.
' **       If sent to PDF, it will not print.
' ******************************************************

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean, lngPrintReportLbl_Left As Long
Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

160           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
170           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

180           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

190           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

200           opgActive_AfterUpdate  ' ** Procedure: Below.
210           opgFont_AfterUpdate  ' ** Procedure: Below.

              ' ** StrConv() conversion argument enumeration:
              ' **   1  vbUpperCase      Converts the string to uppercase characters.
              ' **   2  vbLowerCase      Converts the string to lowercase characters.
              ' **   3  vbProperCase     Converts the first letter of every word in string to uppercase.
              ' **   4  vbWide        1  Converts narrow (single-byte) characters in string to wide (double-byte) characters.
              ' **   8  vbNarrow      1  Converts wide (double-byte) characters in string to narrow (single-byte) characters.
              ' **  16  vbKatakana    2  Converts Hiragana characters in string to Katakana characters.
              ' **  32  vbHiragana    2  Converts Katakana characters in string to Hiragana characters.
              ' **  64  vbUnicode        Converts the string to Unicode using the default code page of the system.
              ' ** 128  vbFromUnicode    Converts the string from Unicode to the default code page of the system.
              ' **                       1: Applies to Far East locales.  2: Applies to Japan only.

              'If CurrentUser = "Superuser" Then
              '  .opgFont.Visible = True
              '  .opgFont_box.Visible = True
              'Else
220           .opgFont.Visible = False
230           .opgFont_box.Visible = False
              'End If

240           .frmCurrency_Symbol_Sub.SetFocus

250         Else
260           Cancel = -1
270         End If
280       Else
290         Cancel = -1
300       End If
310     End With

320     If Cancel = -1 Then
330       Beep
340       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
350     End If

360     DoCmd.Hourglass False

EXITP:
370     Exit Sub

ERRH:
380     DoCmd.Hourglass False
390     Select Case ERR.Number
        Case Else
400       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
410     End Select
420     Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Name:             N {country_name1 on frmCurrency_Symbol_Sub}
        ' **   Currency:         R {curr_word1 on frmCurrency_Symbol_Sub}

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
570         intRetVal = 0
580         cmdClose_Click  ' ** Procedure: Below.
590       End Select
600     End If

        ' ** Alt keys.
610     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
620       Select Case intRetVal
          Case vbKeyN
630         With Me
640           intRetVal = 0
650           .frmCurrency_Symbol_Sub.SetFocus
660           .frmCurrency_Symbol_Sub.Form.country_name1.SetFocus
670         End With
680       Case vbKeyR
690         With Me
700           intRetVal = 0
710           .frmCurrency_Symbol_Sub.SetFocus
720           .frmCurrency_Symbol_Sub.Form.curr_word1.SetFocus
730         End With
740       Case vbKeyX
750         intRetVal = 0
760         cmdClose_Click  ' ** Procedure: Below.
770       End Select
780     End If

        ' ** Ctrl-Shift keys.
790     If intCtrlDown And (Not intAltDown) And intShiftDown Then
800       Select Case intRetVal
          Case vbKeyF
810         With Me
820           intRetVal = 0
830           .FocusHolder.SetFocus
840         End With
850       End Select
860     End If

EXITP:
870     KeyCode = intRetVal
880     Exit Sub

ERRH:
890     intRetVal = 0
900     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
910     Case Else
920       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
930     End Select
940     Resume EXITP

End Sub

Private Sub Form_Timer()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1010    With Me
1020      .TimerInterval = 0&
1030      If gblnSetFocus = True Then
1040        gblnSetFocus = False
1050        strFormName = vbNullString
1060  On Error Resume Next
1070        Do Until strFormName = THIS_NAME
1080          DoCmd.SelectObject acForm, THIS_NAME, False
1090          strFormName = Screen.ActiveForm.Name
1100        Loop
1110  On Error GoTo ERRH
1120        .frmCurrency_Symbol_Sub.SetFocus
1130      End If
1140      If gblnGoToReport = True Then
1150        DoCmd.Hourglass True  ' ** Make sure it's still running.
1160        DoEvents
1170        Select Case blnGoingToReport
            Case True
1180          DoCmd.Hourglass False  ' ** It should already be off.
1190          gblnGoToReport = False
1200          blnGoingToReport = False
1210          .GoToReport_arw_pr_img.Visible = False
1220          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1230          DoEvents
1240        Case False
1250          blnGoingToReport = True
1260          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1270          .GoToReport_arw_pr_img.Visible = True
1280          DoEvents
1290          Beep
1300          DoCmd.Hourglass False
1310          DoEvents
1320          .TimerInterval = (GTR_WAIT * 2&)
1330        End Select
1340      Else
1350        DoCmd.Hourglass False
1360      End If
1370    End With

EXITP:
1380    Exit Sub

ERRH:
1390    DoCmd.Hourglass False
1400    Select Case ERR.Number
        Case Else
1410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1420    End Select
1430    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim blnRetVal As Boolean

1510    With Me

1520      DoCmd.Hourglass True
1530      DoEvents

          ' ** Make sure all reports are closed.
1540      Do While Reports.Count > 0
1550        DoCmd.Close acReport, Reports(0).Name
1560        DoEvents
1570      Loop

1580      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

          'blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.

1590      Select Case strCallingForm
          Case "frmMenu_ForeignExchange"
1600        gstrReturningForm = THIS_NAME
1610        gblnSetFocus = True
1620        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Maintenance"
1630      End Select

1640    End With

EXITP:
1650    Exit Sub

ERRH:
1660    DoCmd.Hourglass False
1670    Select Case ERR.Number
        Case Else
1680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1690    End Select
1700    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1810    With Me
1820      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
1830        Select Case blnPrintReport_Focus
            Case True
1840          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
1850          .cmdPrintReport_raised_img.Visible = False
1860        Case False
1870          .cmdPrintReport_raised_img.Visible = True
1880          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
1890        End Select
1900        .cmdPrintReport_raised_focus_img.Visible = False
1910        .cmdPrintReport_raised_focus_dots_img.Visible = False
1920        .cmdPrintReport_sunken_focus_dots_img.Visible = False
1930        .cmdPrintReport_raised_img_dis.Visible = False
1940      End If
1950    End With

EXITP:
1960    Exit Sub

ERRH:
1970    Select Case ERR.Number
        Case Else
1980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1990    End Select
2000    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

2110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2210    DoCmd.Close acForm, THIS_NAME

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2320    intShiftDown = (Shift And acShiftMask) > 0
2330    intAltDown = (Shift And acAltMask) > 0
2340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2360      Select Case intRetVal
          Case vbKeyTab
2370        With Me
2380          intRetVal = 0
2390          .frmCurrency_Symbol_Sub.SetFocus
2400          .frmCurrency_Symbol_Sub.Form.FocusHolder.SetFocus
2410        End With
2420      End Select
2430    End If

        ' ** Shift keys.
2440    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2450      Select Case intRetVal
          Case vbKeyTab
2460        With Me
2470          intRetVal = 0
2480          If .cmdPrintReport.Enabled = True Then
2490            .cmdPrintReport.SetFocus
2500          Else
2510            .opgActive.SetFocus
2520          End If
2530        End With
2540      End Select
2550    End If

EXITP:
2560    KeyCode = intRetVal
2570    Exit Sub

ERRH:
2580    intRetVal = 0
2590    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2600    Case Else
2610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2620    End Select
2630    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

2710    With Me
2720      blnPrintReport_Focus = True
2730      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2740      .cmdPrintReport_raised_img.Visible = False
2750      .cmdPrintReport_raised_focus_img.Visible = False
2760      .cmdPrintReport_raised_focus_dots_img.Visible = False
2770      .cmdPrintReport_sunken_focus_dots_img.Visible = False
2780      .cmdPrintReport_raised_img_dis.Visible = False
2790    End With

EXITP:
2800    Exit Sub

ERRH:
2810    Select Case ERR.Number
        Case Else
2820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2830    End Select
2840    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

2910    With Me
2920      blnPrintReport_MouseDown = True
2930      .cmdPrintReport_sunken_focus_dots_img.Visible = True
2940      .cmdPrintReport_raised_img.Visible = False
2950      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2960      .cmdPrintReport_raised_focus_img.Visible = False
2970      .cmdPrintReport_raised_focus_dots_img.Visible = False
2980      .cmdPrintReport_raised_img_dis.Visible = False
2990    End With

EXITP:
3000    Exit Sub

ERRH:
3010    Select Case ERR.Number
        Case Else
3020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3030    End Select
3040    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

3110    DoCmd.Hourglass True
3120    DoEvents

3130    Set dbs = CurrentDb
3140    With dbs
          ' ** Empty tmpCurrencySymbolReport.
3150      Set qdf = .QueryDefs("qryRpt_ListOfCurrencySymbols_04")
3160      qdf.Execute
3170      Set qdf = Nothing
          ' ** Append qryRpt_ListOfCurrencySymbols_02 (tblCurrency_Symbol, linked to
          ' ** tblCurrency, qryRpt_ListOfCurrencySymbols_01 (tblCurrency_Country_Primary,
          ' ** linked to tblCountry)) to tmpCurrencySymbolReport.
3180      Set qdf = .QueryDefs("qryRpt_ListOfCurrencySymbols_05")
3190      qdf.Execute
3200      Set qdf = Nothing
          ' ** Append qryRpt_ListOfCurrencySymbols_02 (tblCurrency_Symbol, linked to tblCurrency,
          ' ** qryRpt_ListOfCurrencySymbols_01 (tblCurrency_Country_Primary, linked to tblCountry,
          ' ** with country_name_sort)), linked to qryRpt_ListOfCurrencySymbols_03 (tblReport_Graphics,
          ' ** just 'rptListOfCurrencySymbols'), to tmpCurrencySymbolReport, just 'Armenian Dram'.
3210      Set qdf = .QueryDefs("qryRpt_ListOfCurrencySymbols_05_01")
3220      qdf.Execute
3230      Set qdf = Nothing
          ' ** Append qryRpt_ListOfCurrencySymbols_02 (tblCurrency_Symbol, linked to tblCurrency,
          ' ** qryRpt_ListOfCurrencySymbols_01 (tblCurrency_Country_Primary, linked to tblCountry,
          ' ** with country_name_sort)), linked to qryRpt_ListOfCurrencySymbols_03 (tblReport_Graphics,
          ' ** just 'rptListOfCurrencySymbols'), to tmpCurrencySymbolReport, just 'Bitcoin'.
3240      Set qdf = .QueryDefs("qryRpt_ListOfCurrencySymbols_05_02")
3250      qdf.Execute
3260      Set qdf = Nothing
          ' ** Append qryRpt_ListOfCurrencySymbols_02 (tblCurrency_Symbol, linked to tblCurrency,
          ' ** qryRpt_ListOfCurrencySymbols_01 (tblCurrency_Country_Primary, linked to tblCountry,
          ' ** with country_name_sort)), linked to qryRpt_ListOfCurrencySymbols_03 (tblReport_Graphics,
          ' ** just 'rptListOfCurrencySymbols'), to tmpCurrencySymbolReport, just 'Cape Verdean Escudo'.
3270      Set qdf = .QueryDefs("qryRpt_ListOfCurrencySymbols_05_03")
3280      qdf.Execute
3290      Set qdf = Nothing
          ' ** Append qryRpt_ListOfCurrencySymbols_02 (tblCurrency_Symbol, linked to tblCurrency,
          ' ** qryRpt_ListOfCurrencySymbols_01 (tblCurrency_Country_Primary, linked to tblCountry,
          ' ** with country_name_sort)), linked to qryRpt_ListOfCurrencySymbols_03 (tblReport_Graphics,
          ' ** just 'rptListOfCurrencySymbols'), to tmpCurrencySymbolReport, just 'Litecoin'.
3300      Set qdf = .QueryDefs("qryRpt_ListOfCurrencySymbols_05_04")
3310      qdf.Execute
3320      Set qdf = Nothing
          ' ** Append qryRpt_ListOfCurrencySymbols_02 (tblCurrency_Symbol, linked to tblCurrency,
          ' ** qryRpt_ListOfCurrencySymbols_01 (tblCurrency_Country_Primary, linked to tblCountry,
          ' ** with country_name_sort)), linked to qryRpt_ListOfCurrencySymbols_03 (tblReport_Graphics,
          ' ** just 'rptListOfCurrencySymbols'), to tmpCurrencySymbolReport, just 'Pridnestrovie Ruble'.
3330      Set qdf = .QueryDefs("qryRpt_ListOfCurrencySymbols_05_05")
3340      qdf.Execute
3350      Set qdf = Nothing
          ' ** Append qryRpt_ListOfCurrencySymbols_02 (tblCurrency_Symbol, linked to tblCurrency,
          ' ** qryRpt_ListOfCurrencySymbols_01 (tblCurrency_Country_Primary, linked to tblCountry,
          ' ** with country_name_sort)), linked to qryRpt_ListOfCurrencySymbols_03 (tblReport_Graphics,
          ' ** just 'rptListOfCurrencySymbols'), to tmpCurrencySymbolReport, just 'Portuguese Escudo'.
3360      Set qdf = .QueryDefs("qryRpt_ListOfCurrencySymbols_05_06")
3370      qdf.Execute
3380      Set qdf = Nothing
3390      .Close
3400    End With
3410    Set dbs = Nothing

3420    With Me
3430      strDocName = "rptListOfCurrencySymbols"
3440      If gblnDev_Debug Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
3450        DoCmd.OpenReport strDocName, acViewPreview
3460        DoCmd.Maximize
3470        DoCmd.RunCommand acCmdFitToWindow
3480        .Visible = False
3490      Else
            '##GTR_Ref: rptListOfCurrencySymbols
3500        DoCmd.OpenReport strDocName, acViewNormal
3510      End If
3520    End With

3530    DoCmd.Hourglass False

EXITP:
3540    Set qdf = Nothing
3550    Set dbs = Nothing
3560    Exit Sub

ERRH:
3570    DoCmd.Hourglass False
3580    Select Case ERR.Number
        Case Else
3590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3600    End Select
3610    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

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
          Case vbKeyTab
3770        With Me
3780          intRetVal = 0
3790          .cmdClose.SetFocus
3800        End With
3810      End Select
3820    End If

        ' ** Shift keys.
3830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3840      Select Case intRetVal
          Case vbKeyTab
3850        With Me
3860          intRetVal = 0
3870          .opgActive.SetFocus
3880        End With
3890      End Select
3900    End If

EXITP:
3910    KeyCode = intRetVal
3920    Exit Sub

ERRH:
3930    intRetVal = 0
3940    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3950    Case Else
3960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3970    End Select
3980    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

4010    With Me
4020      If blnPrintReport_MouseDown = False Then
4030        Select Case blnPrintReport_Focus
            Case True
4040          .cmdPrintReport_raised_focus_dots_img.Visible = True
4050          .cmdPrintReport_raised_focus_img.Visible = False
4060        Case False
4070          .cmdPrintReport_raised_focus_img.Visible = True
4080          .cmdPrintReport_raised_focus_dots_img.Visible = False
4090        End Select
4100        .cmdPrintReport_raised_img.Visible = False
4110        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4120        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4130        .cmdPrintReport_raised_img_dis.Visible = False
4140      End If
4150    End With

EXITP:
4160    Exit Sub

ERRH:
4170    Select Case ERR.Number
        Case Else
4180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4190    End Select
4200    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

4310    With Me
4320      .cmdPrintReport_raised_focus_dots_img.Visible = True
4330      .cmdPrintReport_raised_img.Visible = False
4340      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4350      .cmdPrintReport_raised_focus_img.Visible = False
4360      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4370      .cmdPrintReport_raised_img_dis.Visible = False
4380      blnPrintReport_MouseDown = False
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

Private Sub cmdPrintReport_LostFocus()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

4510    With Me
4520      .cmdPrintReport_raised_img.Visible = True
4530      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4540      .cmdPrintReport_raised_focus_img.Visible = False
4550      .cmdPrintReport_raised_focus_dots_img.Visible = False
4560      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4570      .cmdPrintReport_raised_img_dis.Visible = False
4580      blnPrintReport_Focus = False
4590    End With

EXITP:
4600    Exit Sub

ERRH:
4610    Select Case ERR.Number
        Case Else
4620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4630    End Select
4640    Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

4710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4720    Exit Sub

ERRH:
4730    Select Case ERR.Number
        Case Else
4740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4750    End Select
4760    Resume EXITP

End Sub

Private Sub opgActive_AfterUpdate()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_AfterUpdate"

        Dim strFilter As String

4810    With Me
4820      Select Case .opgActive
          Case .opgActive_optInclude.OptionValue
4830        .opgActive_optInclude_lbl.FontBold = True
4840        .opgActive_optExclude_lbl.FontBold = False
4850        .opgActive_optOnly_lbl.FontBold = False
4860        strFilter = vbNullString
4870      Case .opgActive_optExclude.OptionValue
4880        .opgActive_optInclude_lbl.FontBold = False
4890        .opgActive_optExclude_lbl.FontBold = True
4900        .opgActive_optOnly_lbl.FontBold = False
4910        strFilter = "[curr_active] = False"
4920      Case .opgActive_optOnly.OptionValue
4930        .opgActive_optInclude_lbl.FontBold = False
4940        .opgActive_optExclude_lbl.FontBold = False
4950        .opgActive_optOnly_lbl.FontBold = True
4960        strFilter = "[curr_active] = True"
4970      End Select

4980      .frmCurrency_Symbol_Sub.Form.Filter = strFilter
4990      If strFilter = vbNullString Then
5000        .frmCurrency_Symbol_Sub.Form.FilterOn = False
5010      Else
5020        .frmCurrency_Symbol_Sub.Form.FilterOn = True
5030      End If
5040      DoEvents

5050      lngRecsCur = .frmCurrency_Symbol_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_Symbol_Sub.
5060      If lngRecsCur > 0& Then
5070        .frmCurrency_Symbol_Sub.SetFocus
5080      End If

5090    End With

EXITP:
5100    Exit Sub

ERRH:
5110    Select Case ERR.Number
        Case Else
5120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5130    End Select
5140    Resume EXITP

End Sub

Private Sub opgActive_optInclude_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_optInclude_KeyDown"

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
          Case vbKeyTab
5270        With Me
5280          intRetVal = 0
5290          If .cmdPrintReport.Enabled = True Then
5300            .cmdPrintReport.SetFocus
5310          Else
5320            .cmdClose.SetFocus
5330          End If
5340        End With
5350      End Select
5360    End If

        ' ** Shift keys.
5370    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5380      Select Case intRetVal
          Case vbKeyTab
5390        With Me
5400          intRetVal = 0
5410          .frmCurrency_Symbol_Sub.SetFocus
5420          .frmCurrency_Symbol_Sub.Form.FocusHolder.SetFocus
5430        End With
5440      End Select
5450    End If

EXITP:
5460    KeyCode = intRetVal
5470    Exit Sub

ERRH:
5480    intRetVal = 0
5490    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5500    Case Else
5510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5520    End Select
5530    Resume EXITP

End Sub

Private Sub opgActive_optExclude_KeyDown(KeyCode As Integer, Shift As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_optExclude_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5620    intShiftDown = (Shift And acShiftMask) > 0
5630    intAltDown = (Shift And acAltMask) > 0
5640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5660      Select Case intRetVal
          Case vbKeyTab
5670        With Me
5680          intRetVal = 0
5690          If .cmdPrintReport.Enabled = True Then
5700            .cmdPrintReport.SetFocus
5710          Else
5720            .cmdClose.SetFocus
5730          End If
5740        End With
5750      End Select
5760    End If

        ' ** Shift keys.
5770    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5780      Select Case intRetVal
          Case vbKeyTab
5790        With Me
5800          intRetVal = 0
5810          .frmCurrency_Symbol_Sub.SetFocus
5820          .frmCurrency_Symbol_Sub.Form.FocusHolder.SetFocus
5830        End With
5840      End Select
5850    End If

EXITP:
5860    KeyCode = intRetVal
5870    Exit Sub

ERRH:
5880    intRetVal = 0
5890    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5900    Case Else
5910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5920    End Select
5930    Resume EXITP

End Sub

Private Sub opgActive_optOnly_KeyDown(KeyCode As Integer, Shift As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgActive_optOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6020    intShiftDown = (Shift And acShiftMask) > 0
6030    intAltDown = (Shift And acAltMask) > 0
6040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6060      Select Case intRetVal
          Case vbKeyTab
6070        With Me
6080          intRetVal = 0
6090          If .cmdPrintReport.Enabled = True Then
6100            .cmdPrintReport.SetFocus
6110          Else
6120            .cmdClose.SetFocus
6130          End If
6140        End With
6150      End Select
6160    End If

        ' ** Shift keys.
6170    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6180      Select Case intRetVal
          Case vbKeyTab
6190        With Me
6200          intRetVal = 0
6210          .frmCurrency_Symbol_Sub.SetFocus
6220          .frmCurrency_Symbol_Sub.Form.FocusHolder.SetFocus
6230        End With
6240      End Select
6250    End If

EXITP:
6260    KeyCode = intRetVal
6270    Exit Sub

ERRH:
6280    intRetVal = 0
6290    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6300    Case Else
6310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6320    End Select
6330    Resume EXITP

End Sub

Private Sub opgFont_AfterUpdate()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgFont_AfterUpdate"

6410    With Me
6420      Select Case .opgFont
          Case .opgFont_optCode2000.OptionValue
6430        .opgFont_optCode2000_lbl.FontBold = True
6440        .opgFont_optArial_lbl.FontBold = False
6450        .opgFont_optMSSansSeif_lbl.FontBold = False
6460        If .frmCurrency_Symbol_Sub.Form.currsym_symbol_display.FontName <> .opgFont_optCode2000_lbl.Caption Then
6470          .frmCurrency_Symbol_Sub.Form.currsym_symbol_display.FontName = .opgFont_optCode2000_lbl.Caption
6480        End If
6490      Case .opgFont_optArial.OptionValue
6500        .opgFont_optCode2000_lbl.FontBold = False
6510        .opgFont_optArial_lbl.FontBold = True
6520        .opgFont_optMSSansSeif_lbl.FontBold = False
6530        If .frmCurrency_Symbol_Sub.Form.currsym_symbol_display.FontName <> .opgFont_optArial_lbl.Caption Then
6540          .frmCurrency_Symbol_Sub.Form.currsym_symbol_display.FontName = .opgFont_optArial_lbl.Caption
6550        End If
6560      Case .opgFont_optMSSansSeif.OptionValue
6570        .opgFont_optCode2000_lbl.FontBold = False
6580        .opgFont_optArial_lbl.FontBold = False
6590        .opgFont_optMSSansSeif_lbl.FontBold = True
6600        If .frmCurrency_Symbol_Sub.Form.currsym_symbol_display.FontName <> .opgFont_optMSSansSeif_lbl.Caption Then
6610          .frmCurrency_Symbol_Sub.Form.currsym_symbol_display.FontName = .opgFont_optMSSansSeif_lbl.Caption
6620        End If
6630      End Select
6640      .frmCurrency_Symbol_Sub.SetFocus
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
