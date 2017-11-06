VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_AssetHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_AssetHistory"

'VGC 07/15/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, sets options, then stops.

' ** Shortcut Alt keys responsive from this form:
' **   Account:          A {cmbAccounts}
' **   Acct Number:      B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Hidden Assets:    D {chkIncludeHiddenAssets}
' **   Excel:            E {cmdExcel}
' **   From:             F {DateStart}
' **   Purchase Date:    G {chkShowPurchaseDate}
' **   Archived:         H {chkIncludeArchive}
' **   Asset CUSIP:      I {opgAssetSource_optCusip}
' **   All:              L {opgAccountNumber_optAll}
' **   Acct Name:        M {opgAccountSource_optName}
' **   Asset Name:       N {opgAssetSource_optName}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Hidden Trans:     R {chkIncludeHiddenTrans}
' **   Asset:            S {cmbAssets}
' **   To:               T {DateEnd}
' **   Specified:        U {opgAccountNumber_optSpecified}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}
' **   Asset Type:       Y {opgAssetSource_optType}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of Calendar Class.
Public clsMonthClass As clsMonthCal

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean

' ** cmbAssets combo box constants:
Private Const CBX_AST_ASTNO As Integer = 0
Private Const CBX_AST_TDESC As Integer = 1
'Private Const CBX_AST_CUSIP As Integer = 2
'Private Const CBX_AST_ADESC As Integer = 3
'Private Const CBX_AST_ATYPE As Integer = 4

' ** Array: arr_varAcctArch().
Private lngAcctArchs As Long, arr_varAcctArch As Variant
Private Const A_ACTNO As Integer = 0
'Private Const A_TDATE As Integer = 1
'Private Const A_CNT   As Integer = 2

Private blnGoingToReport As Boolean, blnWasGTR As Boolean
Private lngTpp As Long
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           DoCmd.Hourglass True
160           DoEvents

170           blnWasGTR = gblnGoToReport
180           ArrowsOff  ' ** Procedure: below.

190           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
200           blnCalendar2_Focus = False: blnCalendar2_MouseDown = False

210           gstrFormQuerySpec = THIS_NAME
220           gstrAccountNo = vbNullString
230           gdatStartDate = 0
240           gdatEndDate = 0

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
250           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

260           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

      #If NoExcel Then
270           .cmdExcel.Enabled = False
      #Else
280           .cmdExcel.Enabled = True
      #End If

              ' ** EVENT CHECK: chkRememberMe!
290           cmbAccounts_AfterUpdate  ' ** Procedure: Below.
              ' ** EVENT CHECK: chkRememberDates!
300           If IsNull(.DateStart) = False And IsNull(.DateEnd) = False Then
310             DateStart_Exit Cancel  ' ** Procedure: Below.
320             DateEnd_Exit Cancel  ' ** Procedure: Below.
330           End If
              ' ** EVENT CHECK: chkRememberAsset!
              ' ** NONE!

340           opgAccountSource_AfterUpdate  ' ** Procedure: Below.
350           opgAssetSource_AfterUpdate  ' ** Procedure: Below.
360           chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.
370           chkShowPurchaseDate_AfterUpdate  ' ** Procedure: Below.
380           chkIncludeHiddenAssets_AfterUpdate  ' ** Procedure: Below.
390           chkIncludeHiddenTrans_AfterUpdate  ' ** Procedure: Below.
400           opgAccountNumber_AfterUpdate  ' ** Procedure: Below.
410           chkRememberMe_AfterUpdate  ' ** Procedure: Below.
420           chkRememberDates_AfterUpdate  ' ** Procedure: Below.
430           chkRememberAsset_AfterUpdate  ' ** Procedure: Below.
440           chkPageOf_AfterUpdate  ' ** Procedure: Below.

              ' ** Get a list of accounts with archive transactions.
450           Set dbs = CurrentDb
460           With dbs
470             Set qdf = .QueryDefs("qryArchiveTransactions_05")
480             Set rst = qdf.OpenRecordset
490             With rst
500               If .BOF = True And .EOF = True Then
                    ' ** They've not used TrstArch.mdb.
510                 lngAcctArchs = 0&
520               Else
530                 .MoveLast
540                 lngAcctArchs = .RecordCount
550                 .MoveFirst
560                 arr_varAcctArch = .GetRows(lngAcctArchs)
                    ' ***********************************************
                    ' ** Array: arr_varAcctArch()
                    ' **
                    ' **   Field  Element  Name         Constant
                    ' **   =====  =======  ===========  ===========
                    ' **     1       0     accountno    A_ACTNO
                    ' **     2       1     transdate    A_TDATE
                    ' **     3       2     cnt          A_CNT
                    ' **
                    ' ***********************************************
570               End If
580               .Close
590             End With
600             .Close
610           End With

620           SetArchiveOption True  ' ** Procedure: Below.
630           chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.
640           SetHiddenAssets  ' ** Procedure: Below.
650           SetHiddenTrans  ' ** Procedure: Below.

660           DoCmd.Hourglass False

670         Else
680           Cancel = -1
690         End If
700       Else
710         Cancel = -1
720       End If
730     End With

740     If Cancel = -1 Then
750       Beep
760       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
770     End If

EXITP:
780     Set rst = Nothing
790     Set qdf = Nothing
800     Set dbs = Nothing
810     Exit Sub

ERRH:
820     DoCmd.Hourglass False
830     Select Case ERR.Number
        Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
850     End Select
860     Resume EXITP

End Sub

Private Sub Form_Load()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

910     With Me

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
920       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
930       clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
940       clsMonthClass.PositionAtCursor = True

950       .UserReportPath = Pref_ReportPath(.UserReportPath, THIS_NAME)  ' ** Module Function: modPreferenceFuncs.

960     End With

EXITP:
970     Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     Select Case ERR.Number
        Case Else
1000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1010    End Select
1020    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1120    intShiftDown = (Shift And acShiftMask) > 0
1130    intAltDown = (Shift And acAltMask) > 0
1140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Acct Number:      B {opgAccountSource_optNumber}
        ' **   Asset CUSIP:      I {opgAssetSource_optCusip}
        ' **   Acct Name:        M {opgAccountSource_optName}
        ' **   Asset Name:       N {opgAssetSource_optName}
        ' **   Asset Type:       Y {opgAssetSource_optType}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

        ' ** Plain keys.
1150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyEscape
1170        intRetVal = 0
1180        cmdClose_Click  ' ** Procedure: Below.
1190      Case vbKeyF7
1200        intRetVal = 0
1210        cmdCalendar1_Click  ' ** Procedure: Below.
1220      Case vbKeyF8
1230        intRetVal = 0
1240        cmdCalendar2_Click  ' ** Procedure: Below.
1250      End Select
1260    End If

        ' ** Alt keys.
1270    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1280      Select Case intRetVal
          Case vbKeyB
1290        With Me
1300          intRetVal = 0
1310          If .opgAccountSource.Enabled = True Then
1320            .opgAccountSource.SetFocus
1330            .opgAccountSource = .opgAccountSource_optNumber.OptionValue
1340            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1350          Else
1360            Beep
1370          End If
1380        End With
1390      Case vbKeyI
1400        With Me
1410          intRetVal = 0
1420          If .opgAssetSource.Enabled = True Then
1430            .opgAssetSource.SetFocus
1440            .opgAssetSource = .opgAssetSource_optCusip.OptionValue
1450            opgAssetSource_AfterUpdate  ' ** Procedure: Below.
1460          Else
1470            Beep
1480          End If
1490        End With
1500      Case vbKeyM
1510        With Me
1520          intRetVal = 0
1530          If .opgAccountSource.Enabled = True Then
1540            .opgAccountSource.SetFocus
1550            .opgAccountSource = .opgAccountSource_optName.OptionValue
1560            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1570          Else
1580            Beep
1590          End If
1600        End With
1610      Case vbKeyN
1620        With Me
1630          intRetVal = 0
1640          If .opgAssetSource.Enabled = True Then
1650            .opgAssetSource.SetFocus
1660            .opgAssetSource = .opgAssetSource_optName.OptionValue
1670            opgAssetSource_AfterUpdate  ' ** Procedure: Below.
1680          Else
1690            Beep
1700          End If
1710        End With
1720      Case vbKeyX
1730        intRetVal = 0
1740        cmdClose_Click  ' ** Procedure: Below.
1750      Case vbKeyY
1760        With Me
1770          intRetVal = 0
1780          If .opgAssetSource.Enabled = True Then
1790            .opgAssetSource.SetFocus
1800            .opgAssetSource = .opgAssetSource_optType.OptionValue
1810            opgAssetSource_AfterUpdate  ' ** Procedure: Below.
1820          Else
1830            Beep
1840          End If
1850        End With
1860      End Select
1870    End If

        ' ** Ctrl keys.
1880    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1890      Select Case intRetVal
          Case vbKeyG
1900        With Me
1910          intRetVal = 0
1920          If .chkPageOf.Enabled = True Then
1930            .chkPageOf.SetFocus
1940            .chkPageOf = (Not .chkPageOf)
1950            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1960          End If
1970        End With
1980      End Select
1990    End If

        ' ** Ctrl-Shift keys.
2000    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2010      Select Case intRetVal
          Case vbKeyF
2020        With Me
2030          intRetVal = 0
2040          .FocusHolder.SetFocus
2050          .DateStart = Null
2060        End With
2070      End Select
2080    End If

EXITP:
2090    KeyCode = intRetVal
2100    Exit Sub

ERRH:
2110    intRetVal = 0
2120    THAT_PROC = THIS_PROC
2130    That_Erl = Erl
2140    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2150    Resume EXITP

End Sub

Private Sub Form_Timer()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2210    With Me
2220      .TimerInterval = 0&
2230      If gblnSetFocus = True Then
2240        gblnSetFocus = False
2250        strFormName = vbNullString
2260  On Error Resume Next
2270        Do Until Screen.ActiveForm.Name = strFormName
2280          DoCmd.SelectObject acForm, THIS_NAME, False
2290          strFormName = Screen.ActiveForm.Name
2300        Loop
2310  On Error GoTo ERRH
2320      End If
2330      If gblnGoToReport = True Then
2340        Select Case blnGoingToReport
            Case True
2350          If .cmdPrint.Enabled = True Then
2360            .cmdPrint.SetFocus
2370          Else
2380            .cmdClose.SetFocus
2390          End If
2400          ArrowsOff  ' ** Procedure: Below.
2410          DoEvents
2420          blnGoingToReport = False
2430          gblnGoToReport = False
2440          Beep
2450          DoCmd.Hourglass False
2460          DoEvents
2470        Case False
2480          DoCmd.Hourglass True  ' ** Make sure it's still running.
2490          DoEvents
2500          blnGoingToReport = True
2510          Select Case garr_varGoToReport(GTR_RNAM)
              Case "rptAssetHistory"
2520            Select Case .chkShowPurchaseDate
                Case True
2530              Select Case .chkIncludeArchive
                  Case True
2540                .GoToReport_arw_ah_spec_onon_img.Visible = True
2550              Case False
2560                Select Case .chkIncludeArchive.Enabled
                    Case True
2570                  .GoToReport_arw_ah_spec_onoff_img.Visible = True
2580                Case False
2590                  Select Case lngAcctArchs
                      Case 0
2600                    .GoToReport_arw_ah_spec_ondisnoarch_img.Visible = True
2610                  Case Else
2620                    .GoToReport_arw_ah_spec_ondisnothis_img.Visible = True
2630                  End Select
2640                End Select
2650              End Select
2660            Case False
2670              Select Case .chkIncludeArchive
                  Case True
2680                .GoToReport_arw_ah_spec_offon_img.Visible = True
2690              Case False
2700                Select Case .chkIncludeArchive.Enabled
                    Case True
2710                  .GoToReport_arw_ah_spec_offoff_img.Visible = True
2720                Case False
2730                  Select Case lngAcctArchs
                      Case 0
2740                    .GoToReport_arw_ah_spec_offdisnoarch_img.Visible = True
2750                  Case Else
2760                    .GoToReport_arw_ah_spec_offdisnothis_img.Visible = True
2770                  End Select
2780                End Select
2790              End Select
2800            End Select
2810            DoEvents
2820            If .opgAccountNumber <> .opgAccountNumber_optSpecified.OptionValue Then
2830              .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue
2840              opgAccountNumber_AfterUpdate  ' ** Procedure: Below.
2850              DoEvents
2860            End If
2870          Case "rptAssetHistory_All"
2880            Select Case .chkIncludeHiddenAssets
                Case True
2890              Select Case .chkIncludeHiddenTrans.Enabled
                  Case True
2900                .GoToReport_arw_ah_all_on_img.Visible = True
2910              Case False
2920                .GoToReport_arw_ah_all_ondis_img.Visible = True
2930              End Select
2940            Case False
2950              Select Case .chkIncludeHiddenAssets.Enabled
                  Case True
2960                Select Case .chkIncludeHiddenTrans.Enabled
                    Case True
2970                  .GoToReport_arw_ah_all_off_img.Visible = True
2980                Case False
2990                  .GoToReport_arw_ah_all_offdis_img.Visible = True
3000                End Select
3010              Case False
3020                Select Case .chkIncludeHiddenTrans.Enabled
                    Case True
3030                  .GoToReport_arw_ah_all_dis_img.Visible = True
3040                Case False
3050                  .GoToReport_arw_ah_all_disdis_img.Visible = True
3060                End Select
3070              End Select
3080            End Select
3090            DoEvents
3100            If .opgAccountNumber <> .opgAccountNumber_optAll.OptionValue Then
3110              .opgAccountNumber = .opgAccountNumber_optAll.OptionValue
3120              opgAccountNumber_AfterUpdate  ' ** Procedure: Below.
3130              DoEvents
3140            End If
3150          End Select
3160          .TimerInterval = (GTR_WAIT * 2&)
3170        End Select
3180      End If
3190    End With

EXITP:
3200    Exit Sub

ERRH:
3210    DoCmd.Hourglass False
3220    Select Case ERR.Number
        Case Else
3230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3240    End Select
3250    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3310    Select Case DataErr
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
3320      Response = acDataErrContinue
3330    Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
3340      If Left(Me.ActiveControl.Name, 4) = "Date" Then
3350        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3360        Response = acDataErrContinue
3370      End If
3380    Case 2143  ' ** You didn't specify search criteria with a FindRecord action.
          ' ** I have no idea what triggers this!
3390      Response = acDataErrContinue
3400    Case 2237  ' ** The text you entered isn't an item in the list.
3410      Response = acDataErrDisplay
3420    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
3430    Case Else
3440      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3450      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3460    End Select

EXITP:
3470    Exit Sub

ERRH:
3480    Select Case ERR.Number
        Case Else
3490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3500    End Select
3510    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3610    With Me
3620      If .chkRememberDates = False Then
3630        .DateStart = Null  ' ** So it won't be there when Pref_Save() runs.
3640        .DateEnd = Null
3650      End If
3660      If .chkRememberMe = False Then
3670        .cmbAccounts = Null
3680      End If
3690      If .chkRememberAsset = False Then
3700        .cmbAssets = Null
3710      End If
3720    End With

3730    Select Case blnWasGTR
        Case True
3740      blnWasGTR = False
3750    Case False
          ' ** Don't save settings imposed by GoToReport!
3760      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
3770    End Select

3780    gstrFormQuerySpec = vbNullString
3790    gblnSetFocus = True
3800    gstrReturningForm = THIS_NAME
3810    Select Case strCallingForm
        Case "frmMenu_Asset"
3820      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
3830    End Select

EXITP:
3840    Exit Sub

ERRH:
3850    Select Case ERR.Number
        Case Else
3860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3870    End Select
3880    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

3910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4010    With Me
4020      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
4030        Select Case blnCalendar1_Focus
            Case True
4040          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
4050          .cmdCalendar1_raised_img.Visible = False
4060        Case False
4070          .cmdCalendar1_raised_img.Visible = True
4080          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
4090        End Select
4100        .cmdCalendar1_raised_focus_dots_img.Visible = False
4110        .cmdCalendar1_raised_focus_img.Visible = False
4120        .cmdCalendar1_sunken_focus_dots_img.Visible = False
4130        .cmdCalendar1_raised_img_dis.Visible = False
4140      End If
4150      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
4160        Select Case blnCalendar2_Focus
            Case True
4170          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
4180          .cmdCalendar2_raised_img.Visible = False
4190        Case False
4200          .cmdCalendar2_raised_img.Visible = True
4210          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
4220        End Select
4230        .cmdCalendar2_raised_focus_dots_img.Visible = False
4240        .cmdCalendar2_raised_focus_img.Visible = False
4250        .cmdCalendar2_sunken_focus_dots_img.Visible = False
4260        .cmdCalendar2_raised_img_dis.Visible = False
4270      End If
4280    End With

EXITP:
4290    Exit Sub

ERRH:
4300    Select Case ERR.Number
        Case Else
4310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4320    End Select
4330    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

4410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub cmdClose_Click()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

4510    DoCmd.Close acForm, THIS_NAME

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4620    intShiftDown = (Shift And acShiftMask) > 0
4630    intAltDown = (Shift And acAltMask) > 0
4640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4660      Select Case intRetVal
          Case vbKeyTab
4670        With Me
4680          intRetVal = 0
4690          .DateStart.SetFocus
4700        End With
4710      End Select
4720    End If

        ' ** Shift keys.
4730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4740      Select Case intRetVal
          Case vbKeyTab
4750        With Me
4760          intRetVal = 0
4770          If .cmdExcel.Enabled = True Then
4780            .cmdExcel.SetFocus
4790          Else
4800            .cmdWord.SetFocus
4810          End If
4820        End With
4830      End Select
4840    End If

EXITP:
4850    KeyCode = intRetVal
4860    Exit Sub

ERRH:
4870    intRetVal = 0
4880    THAT_PROC = THIS_PROC
4890    That_Erl = Erl
4900    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4910    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

5010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
5020    Exit Sub

ERRH:
5030    Select Case ERR.Number
        Case Else
5040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5050    End Select
5060    Resume EXITP

End Sub

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5120    intShiftDown = (Shift And acShiftMask) > 0
5130    intAltDown = (Shift And acAltMask) > 0
5140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5170        With Me
5180          intRetVal = 0
5190          .DateEnd.SetFocus
5200        End With
5210      End Select
5220    End If

        ' ** Shift keys.
5230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5250        With Me
5260          intRetVal = 0
5270          .cmdClose.SetFocus
5280        End With
5290      End Select
5300    End If

        ' ** Ctrl keys.
5310    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
5320      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5330        With Me
5340          intRetVal = 0
5350          .cmdCalendar1.SetFocus
5360        End With
5370      End Select
5380    End If

EXITP:
5390    KeyCode = intRetVal
5400    Exit Sub

ERRH:
5410    intRetVal = 0
5420    THAT_PROC = THIS_PROC
5430    That_Erl = Erl
5440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5450    Resume EXITP

End Sub

Private Sub DateStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_MouseMove"

5510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
5520    Exit Sub

ERRH:
5530    Select Case ERR.Number
        Case Else
5540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5550    End Select
5560    Resume EXITP

End Sub

Private Sub DateStart_Exit(Cancel As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_Exit"

5610    With Me
5620      If IsDate(.DateEnd) And IsDate(.DateStart) Then
5630        If CDate(.DateStart) > CDate(.DateEnd) Then
5640          MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
5650          Cancel = -1
5660        End If
5670      End If
5680    End With

EXITP:
5690    Exit Sub

ERRH:
5700    THAT_PROC = THIS_PROC
5710    That_Erl = Erl
5720    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5730    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

5810    With Me
5820      blnCalendar1_Focus = True
5830      .cmdCalendar1_raised_semifocus_dots_img.Visible = True
5840      .cmdCalendar1_raised_img.Visible = False
5850      .cmdCalendar1_raised_focus_img.Visible = False
5860      .cmdCalendar1_raised_focus_dots_img.Visible = False
5870      .cmdCalendar1_sunken_focus_dots_img.Visible = False
5880      .cmdCalendar1_raised_img_dis.Visible = False
5890    End With

EXITP:
5900    Exit Sub

ERRH:
5910    Select Case ERR.Number
        Case Else
5920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5930    End Select
5940    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

6010    With Me
6020      blnCalendar1_MouseDown = True
6030      .cmdCalendar1_sunken_focus_dots_img.Visible = True
6040      .cmdCalendar1_raised_img.Visible = False
6050      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6060      .cmdCalendar1_raised_focus_img.Visible = False
6070      .cmdCalendar1_raised_focus_dots_img.Visible = False
6080      .cmdCalendar1_raised_img_dis.Visible = False
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

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

6210    With Me
6220      datStartDate = Date
6230      datEndDate = 0
6240      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
6250      If blnRetVal = True Then
6260        .DateStart = datStartDate
6270      Else
6280        .DateStart = CDate(Format(Date, "mm/dd/yyyy"))
6290      End If
6300      .DateStart.SetFocus
6310    End With

EXITP:
6320    Exit Sub

ERRH:
6330    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
6340    Case Else
6350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6360    End Select
6370    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

6410    With Me
6420      If blnCalendar1_MouseDown = False Then
6430        Select Case blnCalendar1_Focus
            Case True
6440          .cmdCalendar1_raised_focus_dots_img.Visible = True
6450          .cmdCalendar1_raised_focus_img.Visible = False
6460        Case False
6470          .cmdCalendar1_raised_focus_img.Visible = True
6480          .cmdCalendar1_raised_focus_dots_img.Visible = False
6490        End Select
6500        .cmdCalendar1_raised_img.Visible = False
6510        .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6520        .cmdCalendar1_sunken_focus_dots_img.Visible = False
6530        .cmdCalendar1_raised_img_dis.Visible = False
6540      End If
6550    End With

EXITP:
6560    Exit Sub

ERRH:
6570    Select Case ERR.Number
        Case Else
6580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6590    End Select
6600    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

6710    With Me
6720      .cmdCalendar1_raised_focus_dots_img.Visible = True
6730      .cmdCalendar1_raised_img.Visible = False
6740      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6750      .cmdCalendar1_raised_focus_img.Visible = False
6760      .cmdCalendar1_sunken_focus_dots_img.Visible = False
6770      .cmdCalendar1_raised_img_dis.Visible = False
6780      blnCalendar1_MouseDown = False
6790    End With

EXITP:
6800    Exit Sub

ERRH:
6810    Select Case ERR.Number
        Case Else
6820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6830    End Select
6840    Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

6910    With Me
6920      .cmdCalendar1_raised_img.Visible = True
6930      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6940      .cmdCalendar1_raised_focus_img.Visible = False
6950      .cmdCalendar1_raised_focus_dots_img.Visible = False
6960      .cmdCalendar1_sunken_focus_dots_img.Visible = False
6970      .cmdCalendar1_raised_img_dis.Visible = False
6980      blnCalendar1_Focus = False
6990    End With

EXITP:
7000    Exit Sub

ERRH:
7010    Select Case ERR.Number
        Case Else
7020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7030    End Select
7040    Resume EXITP

End Sub

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7120    intShiftDown = (Shift And acShiftMask) > 0
7130    intAltDown = (Shift And acAltMask) > 0
7140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7170        With Me
7180          intRetVal = 0
7190          .cmbAssets.SetFocus
7200        End With
7210      End Select
7220    End If

        ' ** Shift keys.
7230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7250        With Me
7260          intRetVal = 0
7270          .DateStart.SetFocus
7280        End With
7290      End Select
7300    End If

        ' ** Ctrl keys.
7310    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
7320      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7330        With Me
7340          intRetVal = 0
7350          .cmdCalendar2.SetFocus
7360        End With
7370      End Select
7380    End If

        ' ** Ctrl-Shift keys.
7390    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7400      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7410        With Me
7420          intRetVal = 0
7430          .cmdCalendar1.SetFocus
7440        End With
7450      End Select
7460    End If

EXITP:
7470    KeyCode = intRetVal
7480    Exit Sub

ERRH:
7490    intRetVal = 0
7500    THAT_PROC = THIS_PROC
7510    That_Erl = Erl
7520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7530    Resume EXITP

End Sub

Private Sub DateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_MouseMove"

7610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
7620    Exit Sub

ERRH:
7630    Select Case ERR.Number
        Case Else
7640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7650    End Select
7660    Resume EXITP

End Sub

Private Sub DateEnd_Exit(Cancel As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

7710    With Me
7720      If IsNull(.DateEnd) = True And IsNull(.DateStart) = True Then
            ' ** Just let them tab through
7730      ElseIf IsNull(.DateEnd) = True Then
            ' ** Populate it with today's date.
7740        .DateEnd = Date
7750      Else
7760        If IsNull(.DateStart) = True Then
              ' ** Populate it with today's date.
7770          .DateStart = Date
7780        Else
7790          If CDate(.DateStart) > CDate(.DateEnd) Then
7800            MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
7810            .DateStart.SetFocus
7820          End If
7830        End If
7840      End If
7850    End With

EXITP:
7860    Exit Sub

ERRH:
7870    THAT_PROC = THIS_PROC
7880    That_Erl = Erl
7890    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7900    Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

8010    With Me
8020      blnCalendar2_Focus = True
8030      .cmdCalendar2_raised_semifocus_dots_img.Visible = True
8040      .cmdCalendar2_raised_img.Visible = False
8050      .cmdCalendar2_raised_focus_img.Visible = False
8060      .cmdCalendar2_raised_focus_dots_img.Visible = False
8070      .cmdCalendar2_sunken_focus_dots_img.Visible = False
8080      .cmdCalendar2_raised_img_dis.Visible = False
8090    End With

EXITP:
8100    Exit Sub

ERRH:
8110    Select Case ERR.Number
        Case Else
8120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8130    End Select
8140    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

8210    With Me
8220      blnCalendar2_MouseDown = True
8230      .cmdCalendar2_sunken_focus_dots_img.Visible = True
8240      .cmdCalendar2_raised_img.Visible = False
8250      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8260      .cmdCalendar2_raised_focus_img.Visible = False
8270      .cmdCalendar2_raised_focus_dots_img.Visible = False
8280      .cmdCalendar2_raised_img_dis.Visible = False
8290    End With

EXITP:
8300    Exit Sub

ERRH:
8310    Select Case ERR.Number
        Case Else
8320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8330    End Select
8340    Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

8410    With Me
8420      datStartDate = Date
8430      datEndDate = 0
8440      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
8450      If blnRetVal = True Then
8460        .DateEnd = datStartDate
8470      Else
8480        .DateEnd = CDate(Format(Date, "mm/dd/yyyy"))
8490      End If
8500      .DateEnd.SetFocus
8510      Cancel = 0
8520      DateEnd_Exit Cancel  ' ** Procedure: Above.
8530      If Cancel = 0 Then
8540        .cmdPreview.SetFocus
8550      End If
8560    End With

EXITP:
8570    Exit Sub

ERRH:
8580    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
8590    Case Else
8600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8610    End Select
8620    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

8710    With Me
8720      If blnCalendar2_MouseDown = False Then
8730        Select Case blnCalendar2_Focus
            Case True
8740          .cmdCalendar2_raised_focus_dots_img.Visible = True
8750          .cmdCalendar2_raised_focus_img.Visible = False
8760        Case False
8770          .cmdCalendar2_raised_focus_img.Visible = True
8780          .cmdCalendar2_raised_focus_dots_img.Visible = False
8790        End Select
8800        .cmdCalendar2_raised_img.Visible = False
8810        .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8820        .cmdCalendar2_sunken_focus_dots_img.Visible = False
8830        .cmdCalendar2_raised_img_dis.Visible = False
8840      End If
8850    End With

EXITP:
8860    Exit Sub

ERRH:
8870    Select Case ERR.Number
        Case Else
8880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8890    End Select
8900    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

9010    With Me
9020      .cmdCalendar2_raised_focus_dots_img.Visible = True
9030      .cmdCalendar2_raised_img.Visible = False
9040      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
9050      .cmdCalendar2_raised_focus_img.Visible = False
9060      .cmdCalendar2_sunken_focus_dots_img.Visible = False
9070      .cmdCalendar2_raised_img_dis.Visible = False
9080      blnCalendar2_MouseDown = False
9090    End With

EXITP:
9100    Exit Sub

ERRH:
9110    Select Case ERR.Number
        Case Else
9120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9130    End Select
9140    Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

9210    With Me
9220      .cmdCalendar2_raised_img.Visible = True
9230      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
9240      .cmdCalendar2_raised_focus_img.Visible = False
9250      .cmdCalendar2_raised_focus_dots_img.Visible = False
9260      .cmdCalendar2_sunken_focus_dots_img.Visible = False
9270      .cmdCalendar2_raised_img_dis.Visible = False
9280      blnCalendar2_Focus = False
9290    End With

EXITP:
9300    Exit Sub

ERRH:
9310    Select Case ERR.Number
        Case Else
9320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9330    End Select
9340    Resume EXITP

End Sub

Private Sub chkRememberDates_AfterUpdate()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

9410    With Me
9420      Select Case .chkRememberDates
          Case True
9430        .chkRememberDates_lbl.FontBold = True
9440      Case False
9450        .chkRememberDates_lbl.FontBold = False
9460      End Select
9470    End With

EXITP:
9480    Exit Sub

ERRH:
9490    Select Case ERR.Number
        Case Else
9500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9510    End Select
9520    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9620    intShiftDown = (Shift And acShiftMask) > 0
9630    intAltDown = (Shift And acAltMask) > 0
9640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9670        With Me
9680          intRetVal = 0
9690          .opgAssetSource.SetFocus
9700        End With
9710      End Select
9720    End If

        ' ** Shift keys.
9730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9750        With Me
9760          intRetVal = 0
9770          .cmdCalendar2.SetFocus
9780        End With
9790      End Select
9800    End If

EXITP:
9810    KeyCode = intRetVal
9820    Exit Sub

ERRH:
9830    intRetVal = 0
9840    THAT_PROC = THIS_PROC
9850    That_Erl = Erl
9860    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9870    Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

9910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
9920    Exit Sub

ERRH:
9930    Select Case ERR.Number
        Case Else
9940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9950    End Select
9960    Resume EXITP

End Sub

Private Sub opgAssetSource_AfterUpdate()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_AfterUpdate"

        Dim lngAssetNo As Long

10010   With Me
10020     .opgAssetSource_optType_lbl2_dim_hi.Visible = False
10030     .opgAssetSource_optName_lbl2_dim_hi.Visible = False
10040     .opgAssetSource_optCusip_lbl2_dim_hi.Visible = False
10050     If IsNull(.cmbAssets) = False Then
10060       lngAssetNo = .cmbAssets.Column(CBX_AST_ASTNO)
10070     End If
10080     Select Case .opgAssetSource
          Case .opgAssetSource_optType.OptionValue
10090       .opgAssetSource_optType_lbl2.FontBold = True
10100       .opgAssetSource_optType_lbl2_dim_hi.FontBold = True
10110       .opgAssetSource_optName_lbl2.FontBold = False
10120       .opgAssetSource_optName_lbl2_dim_hi.FontBold = False
10130       .opgAssetSource_optCusip_lbl2.FontBold = False
10140       .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = False
10150       Select Case .chkIncludeHiddenAssets
            Case True
              ' ** qryAssetHistory_01 (MasterAsset, with totdesc), with hidden assets; by assettype.
10160         If .cmbAssets.RowSource <> "qryAssetHistory_02b" Then
10170           .cmbAssets.RowSource = "qryAssetHistory_02b"
10180           .cmbAssets.Requery
10190         End If
10200       Case False
              ' ** qryAssetHistory_01 (MasterAsset, with totdesc), without hidden assets; by assettype.
10210         If .cmbAssets.RowSource <> "qryAssetHistory_02a" Then
10220           .cmbAssets.RowSource = "qryAssetHistory_02a"
10230           .cmbAssets.Requery
10240         End If
10250       End Select
10260     Case .opgAssetSource_optName.OptionValue
10270       .opgAssetSource_optType_lbl2.FontBold = False
10280       .opgAssetSource_optType_lbl2_dim_hi.FontBold = False
10290       .opgAssetSource_optName_lbl2.FontBold = True
10300       .opgAssetSource_optName_lbl2_dim_hi.FontBold = True
10310       .opgAssetSource_optCusip_lbl2.FontBold = False
10320       .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = False
10330       Select Case .chkIncludeHiddenAssets
            Case True
              ' **qryAssetHistory_01 (MasterAsset, with totdesc), with hidden assets, by totdesc.
10340         If .cmbAssets.RowSource <> "qryAssetHistory_02d" Then
10350           .cmbAssets.RowSource = "qryAssetHistory_02d"
10360           .cmbAssets.Requery
10370         End If
10380       Case False
              ' ** qryAssetHistory_01 (MasterAsset, with totdesc), without hidden assets, by totdesc.
10390         If .cmbAssets.RowSource <> "qryAssetHistory_02c" Then
10400           .cmbAssets.RowSource = "qryAssetHistory_02c"
10410           .cmbAssets.Requery
10420         End If
10430       End Select
10440     Case .opgAssetSource_optCusip.OptionValue
10450       .opgAssetSource_optType_lbl2.FontBold = False
10460       .opgAssetSource_optType_lbl2_dim_hi.FontBold = False
10470       .opgAssetSource_optName_lbl2.FontBold = False
10480       .opgAssetSource_optName_lbl2_dim_hi.FontBold = False
10490       .opgAssetSource_optCusip_lbl2.FontBold = True
10500       .opgAssetSource_optCusip_lbl2_dim_hi.FontBold = True
10510       Select Case .chkIncludeHiddenAssets
            Case True
              ' ** qryAssetHistory_01 (MasterAsset, with totdesc), with qryAssetHistory_02h
              ' ** (qryAssetHistory_02g (MasterAsset, with cusip_len), grouped, with
              ' ** Max(cusip_len)), with hidden assets, by cusip; Cartesian.
10520         If .cmbAssets.RowSource <> "qryAssetHistory_02f" Then
10530           .cmbAssets.RowSource = "qryAssetHistory_02f"
10540           .cmbAssets.Requery
10550         End If
10560       Case False
              ' ** qryAssetHistory_01 (MasterAsset, with totdesc), with qryAssetHistory_02h
              ' ** (qryAssetHistory_02g (MasterAsset, with cusip_len), grouped, with
              ' ** Max(cusip_len)), without hidden assets, by cusip; Cartesian.
10570         If .cmbAssets.RowSource <> "qryAssetHistory_02e" Then
10580           .cmbAssets.RowSource = "qryAssetHistory_02e"
10590           .cmbAssets.Requery
10600         End If
10610       End Select
10620     End Select
10630   End With

EXITP:
10640   Exit Sub

ERRH:
10650   Select Case ERR.Number
        Case Else
10660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10670   End Select
10680   Resume EXITP

End Sub

Private Sub opgAssetSource_optType_KeyDown(KeyCode As Integer, Shift As Integer)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10720   intShiftDown = (Shift And acShiftMask) > 0
10730   intAltDown = (Shift And acAltMask) > 0
10740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10770       With Me
10780         intRetVal = 0
10790         .chkRememberAsset.SetFocus
10800       End With
10810     End Select
10820   End If

        ' ** Shift keys.
10830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10850       With Me
10860         intRetVal = 0
10870         .chkRememberDates.SetFocus
10880       End With
10890     End Select
10900   End If

EXITP:
10910   KeyCode = intRetVal
10920   Exit Sub

ERRH:
10930   intRetVal = 0
10940   THAT_PROC = THIS_PROC
10950   That_Erl = Erl
10960   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10970   Resume EXITP

End Sub

Private Sub opgAssetSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11020   intShiftDown = (Shift And acShiftMask) > 0
11030   intAltDown = (Shift And acAltMask) > 0
11040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11070       With Me
11080         intRetVal = 0
11090         .chkRememberAsset.SetFocus
11100       End With
11110     End Select
11120   End If

        ' ** Shift keys.
11130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11150       With Me
11160         intRetVal = 0
11170         .chkRememberDates.SetFocus
11180       End With
11190     End Select
11200   End If

EXITP:
11210   KeyCode = intRetVal
11220   Exit Sub

ERRH:
11230   intRetVal = 0
11240   THAT_PROC = THIS_PROC
11250   That_Erl = Erl
11260   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11270   Resume EXITP

End Sub

Private Sub opgAssetSource_optCusip_KeyDown(KeyCode As Integer, Shift As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optCusip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11320   intShiftDown = (Shift And acShiftMask) > 0
11330   intAltDown = (Shift And acAltMask) > 0
11340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11370       With Me
11380         intRetVal = 0
11390         .chkRememberAsset.SetFocus
11400       End With
11410     End Select
11420   End If

        ' ** Shift keys.
11430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11450       With Me
11460         intRetVal = 0
11470         .chkRememberDates.SetFocus
11480       End With
11490     End Select
11500   End If

EXITP:
11510   KeyCode = intRetVal
11520   Exit Sub

ERRH:
11530   intRetVal = 0
11540   THAT_PROC = THIS_PROC
11550   That_Erl = Erl
11560   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11570   Resume EXITP

End Sub

Private Sub chkRememberAsset_AfterUpdate()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberAsset_AfterUpdate"

11610   With Me
11620     Select Case .chkRememberAsset
          Case True
11630       .chkRememberAsset_lbl.FontBold = True
11640     Case False
11650       .chkRememberAsset_lbl.FontBold = False
11660     End Select
11670   End With

EXITP:
11680   Exit Sub

ERRH:
11690   Select Case ERR.Number
        Case Else
11700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11710   End Select
11720   Resume EXITP

End Sub

Private Sub chkRememberAsset_KeyDown(KeyCode As Integer, Shift As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberAsset_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11820   intShiftDown = (Shift And acShiftMask) > 0
11830   intAltDown = (Shift And acAltMask) > 0
11840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11870       With Me
11880         intRetVal = 0
11890         .cmbAssets.SetFocus
11900       End With
11910     End Select
11920   End If

        ' ** Shift keys.
11930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11950       With Me
11960         intRetVal = 0
11970         .opgAssetSource.SetFocus
11980       End With
11990     End Select
12000   End If

EXITP:
12010   KeyCode = intRetVal
12020   Exit Sub

ERRH:
12030   intRetVal = 0
12040   THAT_PROC = THIS_PROC
12050   That_Erl = Erl
12060   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12070   Resume EXITP

End Sub

Private Sub cmbAssets_Enter()

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_Enter"

12110   With Me
12120     If .cmbAssets.text = vbNullString Then
12130       ExpandCombo  ' ** Module Procedure: modUtilities.
12140       .cmbAssets = .cmbAssets.ItemData(0)
12150     End If
12160   End With

EXITP:
12170   Exit Sub

ERRH:
12180   Select Case ERR.Number
        Case Else
12190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12200   End Select
12210   Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12320   intShiftDown = (Shift And acShiftMask) > 0
12330   intAltDown = (Shift And acAltMask) > 0
12340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12370       With Me
12380         intRetVal = 0
12390         .opgAccountNumber.SetFocus
12400       End With
12410     End Select
12420   End If

        ' ** Shift keys.
12430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12450       With Me
12460         intRetVal = 0
12470         .DateEnd.SetFocus
12480       End With
12490     End Select
12500   End If

        ' ** Ctrl-Shift keys.
12510   If intCtrlDown And (Not intAltDown) And intShiftDown Then
12520     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12530       With Me
12540         intRetVal = 0
12550         .chkRememberAsset.SetFocus
12560       End With
12570     End Select
12580   End If

EXITP:
12590   KeyCode = intRetVal
12600   Exit Sub

ERRH:
12610   intRetVal = 0
12620   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
12630   Case Else
12640     THAT_PROC = THIS_PROC
12650     That_Erl = Erl
12660     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12670   End Select
12680   Resume EXITP

End Sub

Public Sub opgAccountNumber_AfterUpdate()
' ** Public for GoToReport.

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_AfterUpdate"

12710   With Me
12720     Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
12730       .opgAccountNumber_optSpecified_lbl.FontBold = True
12740       .opgAccountNumber_optAll_lbl.FontBold = False
12750       .opgAccountNumber_optSpecified_lbl_box.Visible = True
12760       .opgAccountNumber_optAll_lbl_box.Visible = False
12770       .cmbAccounts.Enabled = True
12780       .cmbAccounts.BorderColor = CLR_LTBLU2
12790       .cmbAccounts.BackStyle = acBackStyleNormal
12800       .opgAccountSource.Enabled = True
12810       .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
12820       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
12830       .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
12840       .opgAccountSource_optName_lbl2_dim_hi.Visible = False
12850       .chkRememberMe.Enabled = True
12860       .chkRememberMe_lbl.Visible = True
12870       .chkRememberMe_lbl2_dim.Visible = False
12880       .chkRememberMe_lbl2_dim_hi.Visible = False
12890     Case .opgAccountNumber_optAll.OptionValue
12900       .opgAccountNumber_optSpecified_lbl.FontBold = False
12910       .opgAccountNumber_optAll_lbl.FontBold = True
12920       .opgAccountNumber_optSpecified_lbl_box.Visible = False
12930       .opgAccountNumber_optAll_lbl_box.Visible = True
12940       .cmbAccounts.Enabled = False
12950       .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
12960       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
12970       .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
12980       .opgAccountSource_optName_lbl2_dim_hi.Visible = True
12990       .cmbAccounts.BorderColor = WIN_CLR_DISR
13000       .cmbAccounts.BackStyle = acBackStyleTransparent
13010       .opgAccountSource.Enabled = False
13020       .chkRememberMe.Enabled = False
13030       .chkRememberMe_lbl.Visible = False
13040       .chkRememberMe_lbl2_dim.Visible = True
13050       .chkRememberMe_lbl2_dim_hi.Visible = True
13060       SetArchiveOption True  ' ** Procedure: Below.
13070     End Select
13080   End With

EXITP:
13090   Exit Sub

ERRH:
13100   Select Case ERR.Number
        Case Else
13110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13120   End Select
13130   Resume EXITP

End Sub

Private Sub opgAccountNumber_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optSpecified_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13220   intShiftDown = (Shift And acShiftMask) > 0
13230   intAltDown = (Shift And acAltMask) > 0
13240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13270       With Me
13280         intRetVal = 0
13290         If .cmbAccounts.Enabled = True Then
13300           .cmbAccounts.SetFocus
13310         Else
13320           .cmdPreview.SetFocus
13330         End If
13340       End With
13350     End Select
13360   End If

        ' ** Shift keys.
13370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13390       With Me
13400         intRetVal = 0
13410         .cmbAssets.SetFocus
13420       End With
13430     End Select
13440   End If

        ' ** Ctrl keys.
13450   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
13460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13470       With Me
13480         intRetVal = 0
13490         If .chkIncludeArchive.Enabled = True Then
13500           .chkIncludeArchive.SetFocus
13510         Else
13520           .chkShowPurchaseDate.SetFocus
13530         End If
13540       End With
13550     End Select
13560   End If

EXITP:
13570   KeyCode = intRetVal
13580   Exit Sub

ERRH:
13590   intRetVal = 0
13600   THAT_PROC = THIS_PROC
13610   That_Erl = Erl
13620   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13630   Resume EXITP

End Sub

Private Sub opgAccountNumber_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13720   intShiftDown = (Shift And acShiftMask) > 0
13730   intAltDown = (Shift And acAltMask) > 0
13740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13770       With Me
13780         intRetVal = 0
13790         If .cmbAccounts.Enabled = True Then
13800           .cmbAccounts.SetFocus
13810         Else
13820           .cmdPreview.SetFocus
13830         End If
13840       End With
13850     End Select
13860   End If

        ' ** Shift keys.
13870   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13880     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13890       With Me
13900         intRetVal = 0
13910         .cmbAssets.SetFocus
13920       End With
13930     End Select
13940   End If

        ' ** Ctrl keys.
13950   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
13960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13970       With Me
13980         intRetVal = 0
13990         If .chkIncludeArchive.Enabled = True Then
14000           .chkIncludeArchive.SetFocus
14010         Else
14020           .chkShowPurchaseDate.SetFocus
14030         End If
14040       End With
14050     End Select
14060   End If

EXITP:
14070   KeyCode = intRetVal
14080   Exit Sub

ERRH:
14090   intRetVal = 0
14100   THAT_PROC = THIS_PROC
14110   That_Erl = Erl
14120   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14130   Resume EXITP

End Sub

Private Sub chkIncludeArchive_AfterUpdate()

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_AfterUpdate"

14210   With Me
14220     Select Case .chkIncludeArchive
          Case True  ' ** Checked for archive transactions.
14230       .chkIncludeArchive_lbl.FontBold = True
            ' ** Force specific account. WHY?
            '.opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue
            '.opgAccountNumber_optAll.Enabled = False
            'opgAccountNumber_AfterUpdate  ' ** Procedure: Below.
14240     Case Else  ' ** Regular.
14250       .chkIncludeArchive_lbl.FontBold = False
            '.opgAccountNumber_optAll.Enabled = True
14260     End Select
14270   End With

EXITP:
14280   Exit Sub

ERRH:
14290   Select Case ERR.Number
        Case Else
14300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14310   End Select
14320   Resume EXITP

End Sub

Private Sub chkIncludeArchive_KeyDown(KeyCode As Integer, Shift As Integer)

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14420   intShiftDown = (Shift And acShiftMask) > 0
14430   intAltDown = (Shift And acAltMask) > 0
14440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14470       With Me
14480         intRetVal = 0
14490         .chkShowPurchaseDate.SetFocus
14500       End With
14510     End Select
14520   End If

        ' ** Shift keys.
14530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14550       With Me
14560         intRetVal = 0
14570         .opgAccountNumber.SetFocus
14580       End With
14590     End Select
14600   End If

EXITP:
14610   KeyCode = intRetVal
14620   Exit Sub

ERRH:
14630   intRetVal = 0
14640   THAT_PROC = THIS_PROC
14650   That_Erl = Erl
14660   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14670   Resume EXITP

End Sub

Private Sub chkShowPurchaseDate_AfterUpdate()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowPurchaseDate_AfterUpdate"

14710   With Me
14720     Select Case .chkShowPurchaseDate
          Case True
14730       .chkShowPurchaseDate_lbl.FontBold = True
14740     Case Else
14750       .chkShowPurchaseDate_lbl.FontBold = False
14760     End Select
14770   End With

EXITP:
14780   Exit Sub

ERRH:
14790   Select Case ERR.Number
        Case Else
14800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14810   End Select
14820   Resume EXITP

End Sub

Private Sub chkShowPurchaseDate_KeyDown(KeyCode As Integer, Shift As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowPurchaseDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14920   intShiftDown = (Shift And acShiftMask) > 0
14930   intAltDown = (Shift And acAltMask) > 0
14940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14970       With Me
14980         intRetVal = 0
14990         If .chkIncludeHiddenAssets.Enabled = True Then
15000           .chkIncludeHiddenAssets.SetFocus
15010         ElseIf .chkIncludeHiddenTrans.Enabled = True Then
15020           .chkIncludeHiddenTrans.SetFocus
15030         ElseIf .cmbAccounts.Enabled = True Then
15040           .cmbAccounts.SetFocus
15050         Else
15060           .cmdPreview.SetFocus
15070         End If
15080       End With
15090     End Select
15100   End If

        ' ** Shift keys.
15110   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15120     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15130       With Me
15140         intRetVal = 0
15150         If .chkIncludeArchive.Enabled = True Then
15160           .chkIncludeArchive.SetFocus
15170         Else
15180           .opgAccountNumber.SetFocus
15190         End If
15200       End With
15210     End Select
15220   End If

EXITP:
15230   KeyCode = intRetVal
15240   Exit Sub

ERRH:
15250   intRetVal = 0
15260   THAT_PROC = THIS_PROC
15270   That_Erl = Erl
15280   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15290   Resume EXITP

End Sub

Private Sub chkIncludeHiddenAssets_AfterUpdate()

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeHiddenAssets_AfterUpdate"

15310   With Me
15320     Select Case .chkIncludeHiddenAssets
          Case True
15330       .chkIncludeHiddenAssets_lbl.FontBold = True
15340       Select Case .opgAssetSource
            Case .opgAssetSource_optType.OptionValue
              ' ** qryAssetHistory_01 (MasterAsset, with totdesc), with hidden assets; by assettype.
15350         If .cmbAssets.RowSource <> "qryAssetHistory_02b" Then
15360           .cmbAssets.RowSource = "qryAssetHistory_02b"
15370           .cmbAssets.Requery
15380         End If
15390       Case .opgAssetSource_optName.OptionValue
              ' **qryAssetHistory_01 (MasterAsset, with totdesc), with hidden assets, by totdesc.
15400         If .cmbAssets.RowSource <> "qryAssetHistory_02d" Then
15410           .cmbAssets.RowSource = "qryAssetHistory_02d"
15420           .cmbAssets.Requery
15430         End If
15440       Case .opgAssetSource_optCusip.OptionValue
              ' ** qryAssetHistory_01 (MasterAsset, with totdesc), with qryAssetHistory_02h
              ' ** (qryAssetHistory_02g (MasterAsset, with cusip_len), grouped, with
              ' ** Max(cusip_len)), with hidden assets, by cusip; Cartesian.
15450         If .cmbAssets.RowSource <> "qryAssetHistory_02f" Then
15460           .cmbAssets.RowSource = "qryAssetHistory_02f"
15470           .cmbAssets.Requery
15480         End If
15490       End Select
15500     Case Else
15510       .chkIncludeHiddenAssets_lbl.FontBold = False
15520       Select Case .opgAssetSource
            Case .opgAssetSource_optType.OptionValue
              ' ** qryAssetHistory_01 (MasterAsset, with totdesc), without hidden assets; by assettype.
15530         If .cmbAssets.RowSource <> "qryAssetHistory_02a" Then
15540           .cmbAssets.RowSource = "qryAssetHistory_02a"
15550           .cmbAssets.Requery
15560         End If
15570       Case .opgAssetSource_optName.OptionValue
              ' ** qryAssetHistory_01 (MasterAsset, with totdesc), without hidden assets, by totdesc.
15580         If .cmbAssets.RowSource <> "qryAssetHistory_02c" Then
15590           .cmbAssets.RowSource = "qryAssetHistory_02c"
15600           .cmbAssets.Requery
15610         End If
15620       Case .opgAssetSource_optCusip.OptionValue
              ' ** qryAssetHistory_01 (MasterAsset, with totdesc), with qryAssetHistory_02h
              ' ** (qryAssetHistory_02g (MasterAsset, with cusip_len), grouped, with
              ' ** Max(cusip_len)), without hidden assets, by cusip; Cartesian.
15630         If .cmbAssets.RowSource <> "qryAssetHistory_02e" Then
15640           .cmbAssets.RowSource = "qryAssetHistory_02e"
15650           .cmbAssets.Requery
15660         End If
15670       End Select
15680     End Select
15690   End With

EXITP:
15700   Exit Sub

ERRH:
15710   Select Case ERR.Number
        Case Else
15720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15730   End Select
15740   Resume EXITP

End Sub

Private Sub chkIncludeHiddenAssets_KeyDown(KeyCode As Integer, Shift As Integer)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeHiddenAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15820   intShiftDown = (Shift And acShiftMask) > 0
15830   intAltDown = (Shift And acAltMask) > 0
15840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15870       With Me
15880         intRetVal = 0
15890         If .chkIncludeHiddenTrans.Enabled = True Then
15900           .chkIncludeHiddenTrans.SetFocus
15910         ElseIf .cmbAccounts.Enabled = True Then
15920           .cmbAccounts.SetFocus
15930         Else
15940           .cmdPreview.SetFocus
15950         End If
15960       End With
15970     End Select
15980   End If

        ' ** Shift keys.
15990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16000     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16010       With Me
16020         intRetVal = 0
16030         .chkShowPurchaseDate.SetFocus
16040       End With
16050     End Select
16060   End If

EXITP:
16070   KeyCode = intRetVal
16080   Exit Sub

ERRH:
16090   intRetVal = 0
16100   THAT_PROC = THIS_PROC
16110   That_Erl = Erl
16120   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16130   Resume EXITP

End Sub

Private Sub chkIncludeHiddenTrans_AfterUpdate()

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeHiddenTrans_AfterUpdate"

16210   With Me
16220     Select Case .chkIncludeHiddenTrans
          Case True
16230       .chkIncludeHiddenTrans_lbl.FontBold = True
16240     Case Else
16250       .chkIncludeHiddenTrans_lbl.FontBold = False
16260     End Select
16270   End With

EXITP:
16280   Exit Sub

ERRH:
16290   Select Case ERR.Number
        Case Else
16300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16310   End Select
16320   Resume EXITP

End Sub

Private Sub chkIncludeHiddenTrans_KeyDown(KeyCode As Integer, Shift As Integer)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeHiddenTrans_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16420   intShiftDown = (Shift And acShiftMask) > 0
16430   intAltDown = (Shift And acAltMask) > 0
16440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16470       With Me
16480         intRetVal = 0
16490         If .cmbAccounts.Enabled = True Then
16500           .cmbAccounts.SetFocus
16510         Else
16520           .cmdPreview.SetFocus
16530         End If
16540       End With
16550     End Select
16560   End If

        ' ** Shift keys.
16570   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16580     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16590       With Me
16600         intRetVal = 0
16610         If .chkIncludeHiddenAssets.Enabled = True Then
16620           .chkIncludeHiddenAssets.SetFocus
16630         Else
16640           .chkShowPurchaseDate.SetFocus
16650         End If
16660       End With
16670     End Select
16680   End If

        ' ** Ctrl keys.
16690   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
16700     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16710       With Me
16720         intRetVal = 0
16730         .opgAccountSource.SetFocus
16740       End With
16750     End Select
16760   End If

EXITP:
16770   KeyCode = intRetVal
16780   Exit Sub

ERRH:
16790   intRetVal = 0
16800   THAT_PROC = THIS_PROC
16810   That_Erl = Erl
16820   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16830   Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

16910   strAccountNo = vbNullString

16920   With Me
16930     If IsNull(.cmbAccounts) = False Then
16940       If Len(.cmbAccounts.Column(0)) > 0 Then
16950         strAccountNo = .cmbAccounts.Column(0)
16960       End If
16970     End If
16980     .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
16990     .opgAccountSource_optName_lbl2_dim_hi.Visible = False
17000     Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
17010       .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
17020       .opgAccountSource_optNumber_lbl2.FontBold = True
17030       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = True
17040       .opgAccountSource_optName_lbl2.FontBold = False
17050       .opgAccountSource_optName_lbl2_dim_hi.FontBold = False
17060     Case .opgAccountSource_optName.OptionValue
17070       .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
17080       .opgAccountSource_optNumber_lbl2.FontBold = False
17090       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = False
17100       .opgAccountSource_optName_lbl2.FontBold = True
17110       .opgAccountSource_optName_lbl2_dim_hi.FontBold = True
17120     End Select
17130     DoEvents
17140     If strAccountNo <> vbNullString Then
17150       .cmbAccounts = strAccountNo
17160     End If
17170   End With

EXITP:
17180   Exit Sub

ERRH:
17190   Select Case ERR.Number
        Case Else
17200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17210   End Select
17220   Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17320   intShiftDown = (Shift And acShiftMask) > 0
17330   intAltDown = (Shift And acAltMask) > 0
17340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17370       With Me
17380         intRetVal = 0
17390         .chkRememberMe.SetFocus
17400       End With
17410     End Select
17420   End If

        ' ** Shift keys.
17430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17450       With Me
17460         intRetVal = 0
17470         If .chkIncludeHiddenTrans.Enabled = True Then
17480           .chkIncludeHiddenTrans.SetFocus
17490         ElseIf .chkIncludeHiddenAssets.Enabled = True Then
17500           .chkIncludeHiddenAssets.SetFocus
17510         Else
17520           .chkShowPurchaseDate.SetFocus
17530         End If
17540       End With
17550     End Select
17560   End If

EXITP:
17570   KeyCode = intRetVal
17580   Exit Sub

ERRH:
17590   intRetVal = 0
17600   THAT_PROC = THIS_PROC
17610   That_Erl = Erl
17620   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17630   Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17720   intShiftDown = (Shift And acShiftMask) > 0
17730   intAltDown = (Shift And acAltMask) > 0
17740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17770       With Me
17780         intRetVal = 0
17790         .chkRememberMe.SetFocus
17800       End With
17810     End Select
17820   End If

        ' ** Shift keys.
17830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17850       With Me
17860         intRetVal = 0
17870         If .chkIncludeHiddenTrans.Enabled = True Then
17880           .chkIncludeHiddenTrans.SetFocus
17890         ElseIf .chkIncludeHiddenAssets.Enabled = True Then
17900           .chkIncludeHiddenAssets.SetFocus
17910         Else
17920           .chkShowPurchaseDate.SetFocus
17930         End If
17940       End With
17950     End Select
17960   End If

EXITP:
17970   KeyCode = intRetVal
17980   Exit Sub

ERRH:
17990   intRetVal = 0
18000   THAT_PROC = THIS_PROC
18010   That_Erl = Erl
18020   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18030   Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

18110   With Me
18120     Select Case .chkRememberMe
          Case True
18130       .chkRememberMe_lbl.FontBold = True
18140       .chkRememberMe_lbl2_dim.FontBold = True
18150       .chkRememberMe_lbl2_dim_hi.FontBold = True
18160     Case False
18170       .chkRememberMe_lbl.FontBold = False
18180       .chkRememberMe_lbl2_dim.FontBold = False
18190       .chkRememberMe_lbl2_dim_hi.FontBold = False
18200     End Select
18210   End With

EXITP:
18220   Exit Sub

ERRH:
18230   Select Case ERR.Number
        Case Else
18240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18250   End Select
18260   Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18320   intShiftDown = (Shift And acShiftMask) > 0
18330   intAltDown = (Shift And acAltMask) > 0
18340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18370       With Me
18380         intRetVal = 0
18390         .cmbAccounts.SetFocus
18400       End With
18410     End Select
18420   End If

        ' ** Shift keys.
18430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18450       With Me
18460         intRetVal = 0
18470         .opgAccountSource.SetFocus
18480       End With
18490     End Select
18500   End If

EXITP:
18510   KeyCode = intRetVal
18520   Exit Sub

ERRH:
18530   intRetVal = 0
18540   THAT_PROC = THIS_PROC
18550   That_Erl = Erl
18560   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18570   Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

18610   With Me
18620     If .cmbAccounts.text = vbNullString Then
18630       ExpandCombo  ' ** Module Procedure: modUtilities.
18640       .cmbAccounts = .cmbAccounts.ItemData(0)
18650     End If
18660   End With

EXITP:
18670   Exit Sub

ERRH:
18680   Select Case ERR.Number
        Case Else
18690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18700   End Select
18710   Resume EXITP

End Sub

Private Sub cmbAccounts_AfterUpdate()

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

18810   SetArchiveOption True  ' ** Procedure: Below.

EXITP:
18820   Exit Sub

ERRH:
18830   Select Case ERR.Number
        Case Else
18840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18850   End Select
18860   Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18920   intShiftDown = (Shift And acShiftMask) > 0
18930   intAltDown = (Shift And acAltMask) > 0
18940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18970       With Me
18980         intRetVal = 0
18990         .cmdPreview.SetFocus
19000       End With
19010     End Select
19020   End If

        ' ** Shift keys.
19030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19050       With Me
19060         intRetVal = 0
19070         .opgAccountNumber.SetFocus
19080       End With
19090     End Select
19100   End If

        ' ** Ctrl keys.
19110   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
19120     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19130       With Me
19140         intRetVal = 0
19150         .chkPageOf.SetFocus
19160       End With
19170     End Select
19180   End If

        ' ** Ctrl-Shift keys.
19190   If intCtrlDown And (Not intAltDown) And intShiftDown Then
19200     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19210       With Me
19220         intRetVal = 0
19230         .chkRememberMe.SetFocus
19240       End With
19250     End Select
19260   End If

EXITP:
19270   KeyCode = intRetVal
19280   Exit Sub

ERRH:
19290   intRetVal = 0
19300   THAT_PROC = THIS_PROC
19310   That_Erl = Erl
19320   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19330   Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

19410   With Me
19420     Select Case .chkPageOf
          Case True
19430       .chkPageOf_lbl.FontBold = True
19440     Case False
19450       .chkPageOf_lbl.FontBold = False
19460     End Select
19470   End With

EXITP:
19480   Exit Sub

ERRH:
19490   Select Case ERR.Number
        Case Else
19500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19510   End Select
19520   Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19620   intShiftDown = (Shift And acShiftMask) > 0
19630   intAltDown = (Shift And acAltMask) > 0
19640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19670       With Me
19680         intRetVal = 0
19690         .cmdPreview.SetFocus
19700       End With
19710     End Select
19720   End If

        ' ** Shift keys.
19730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19750       With Me
19760         intRetVal = 0
19770         If .cmbAccounts.Enabled = True Then
19780           .cmbAccounts.SetFocus
19790         ElseIf .chkIncludeHiddenTrans.Enabled = True Then
19800           .chkIncludeHiddenTrans.SetFocus
19810         ElseIf .chkIncludeHiddenAssets.Enabled = True Then
19820           .chkIncludeHiddenAssets.SetFocus
19830         Else
19840           .chkShowPurchaseDate.SetFocus
19850         End If
19860       End With
19870     End Select
19880   End If

EXITP:
19890   KeyCode = intRetVal
19900   Exit Sub

ERRH:
19910   intRetVal = 0
19920   THAT_PROC = THIS_PROC
19930   That_Erl = Erl
19940   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19950   Resume EXITP

End Sub

Private Sub cmdPreview_Click()

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

20010   With Me
20020     If DoReport = True Then  ' ** Function: Below.
20030       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
20040         strDocName = "rptAssetHistory"
20050       Case .opgAccountNumber_optAll.OptionValue
20060         strDocName = "rptAssetHistory_All"
20070       End Select
20080       DoCmd.OpenReport strDocName, acViewPreview
20090       DoCmd.Maximize
20100       DoCmd.RunCommand acCmdFitToWindow
20110     End If
20120   End With

EXITP:
20130   Exit Sub

ERRH:
20140   Select Case ERR.Number
        Case Else
20150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20160   End Select
20170   Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20220   intShiftDown = (Shift And acShiftMask) > 0
20230   intAltDown = (Shift And acAltMask) > 0
20240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20260     Select Case intRetVal
          Case vbKeyTab
20270       With Me
20280         intRetVal = 0
20290         .cmdPrint.SetFocus
20300       End With
20310     End Select
20320   End If

        ' ** Shift keys.
20330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20340     Select Case intRetVal
          Case vbKeyTab
20350       With Me
20360         intRetVal = 0
20370         If .cmbAccounts.Enabled = True Then
20380           .cmbAccounts.SetFocus
20390         Else
20400           .opgAccountNumber.SetFocus
20410         End If
20420       End With
20430     End Select
20440   End If

        ' ** Ctlr-Shift keys.
20450   If intCtrlDown And (Not intAltDown) And intShiftDown Then
20460     Select Case intRetVal
          Case vbKeyTab
20470       With Me
20480         intRetVal = 0
20490         .chkRememberMe.SetFocus
20500       End With
20510     End Select
20520   End If

EXITP:
20530   KeyCode = intRetVal
20540   Exit Sub

ERRH:
20550   intRetVal = 0
20560   THAT_PROC = THIS_PROC
20570   That_Erl = Erl
20580   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20590   Resume EXITP

End Sub

Private Sub cmdPrint_Click()

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

20610   With Me
20620     If DoReport = True Then  ' ** Procedure: Below.
20630       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
20640         strDocName = "rptAssetHistory"
20650       Case .opgAccountNumber_optAll.OptionValue
20660         strDocName = "rptAssetHistory_All"
20670       End Select
20680       If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
20690         DoCmd.OpenReport strDocName, acViewPreview
20700         DoCmd.Maximize
20710         DoCmd.RunCommand acCmdFitToWindow
20720       Else
              '##GTR_Ref: rptAssetHistory
              '##GTR_Ref: rptAssetHistory_All
20730         DoCmd.OpenReport strDocName, acViewNormal
20740       End If
20750     End If
20760   End With

EXITP:
20770   Exit Sub

ERRH:
20780   THAT_PROC = THIS_PROC
20790   That_Erl = Erl
20800   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20810   Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20920   intShiftDown = (Shift And acShiftMask) > 0
20930   intAltDown = (Shift And acAltMask) > 0
20940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20960     Select Case intRetVal
          Case vbKeyTab
20970       With Me
20980         intRetVal = 0
20990         .cmdWord.SetFocus
21000       End With
21010     End Select
21020   End If

        ' ** Shift keys.
21030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21040     Select Case intRetVal
          Case vbKeyTab
21050       With Me
21060         intRetVal = 0
21070         .cmdPreview.SetFocus
21080       End With
21090     End Select
21100   End If

EXITP:
21110   KeyCode = intRetVal
21120   Exit Sub

ERRH:
21130   intRetVal = 0
21140   THAT_PROC = THIS_PROC
21150   That_Erl = Erl
21160   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
21170   Resume EXITP

End Sub

Private Sub cmdWord_Click()

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String, strRptAcct As String
        Dim strRptPath As String, strRptPathFile As String
        Dim strCusip As String

21210   With Me
21220     If DoReport = True Then  ' ** Procedure: Below.

21230       If IsNull(.UserReportPath) = True Then
21240         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
21250       Else
21260         strRptPath = .UserReportPath
21270       End If

21280       strCusip = .cmbAssets.Column(2)

21290       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
21300         strRpt = "rptAssetHistory"
21310         strRptAcct = .cmbAccounts
21320         strRptCap = "Asset_History_" & strRptAcct & "_" & strCusip & "_" & _
                Format(.DateStart, "yymmdd") & "_To_" & Format(.DateEnd, "yymmdd")
21330       Case .opgAccountNumber_optAll.OptionValue
21340         strRpt = "rptAssetHistory_All"
21350         strRptCap = "Asset_History_All_" & strCusip & "_" & _
                Format(.DateStart, "yymmdd") & "_To_" & Format(.DateEnd, "yymmdd")
21360       End Select

21370       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

21380       If strRptPathFile <> vbNullString Then
21390         DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
21400         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
21410       End If

21420     End If
21430   End With

EXITP:
21440   Exit Sub

ERRH:
21450   Select Case ERR.Number
        Case Else
21460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21470   End Select
21480   Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21520   intShiftDown = (Shift And acShiftMask) > 0
21530   intAltDown = (Shift And acAltMask) > 0
21540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21560     Select Case intRetVal
          Case vbKeyTab
21570       With Me
21580         intRetVal = 0
21590         If .cmdExcel.Enabled = True Then
21600           .cmdExcel.SetFocus
21610         Else
21620           .cmdClose.SetFocus
21630         End If
21640       End With
21650     End Select
21660   End If

        ' ** Shift keys.
21670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21680     Select Case intRetVal
          Case vbKeyTab
21690       With Me
21700         intRetVal = 0
21710         .cmdPrint.SetFocus
21720       End With
21730     End Select
21740   End If

EXITP:
21750   KeyCode = intRetVal
21760   Exit Sub

ERRH:
21770   intRetVal = 0
21780   THAT_PROC = THIS_PROC
21790   That_Erl = Erl
21800   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
21810   Resume EXITP

End Sub

Private Sub cmdExcel_Click()

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String, strRptCap As String, strRptAcct As String
        Dim strRptPath As String, strRptPathFile As String
        Dim strCusip As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

21910   With Me

21920     DoCmd.Hourglass True
21930     DoEvents

21940     blnContinue = True

21950     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
21960       DoCmd.Hourglass False
21970       msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
21980       If msgResponse <> vbRetry Then
21990         blnContinue = False
22000       End If
22010     End If

22020     If blnContinue = True Then

22030       DoCmd.Hourglass True
22040       DoEvents

22050       If DoReport = True Then  ' ** Procedure: Below.

22060         blnContinue = True

22070         If IsNull(.UserReportPath) = True Then
22080           strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
22090         Else
22100           strRptPath = .UserReportPath
22110         End If

22120         strCusip = .cmbAssets.Column(2)

22130         Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
                ' ** qryRpt_AssetHistory_43 (xx), For Export, Specified Account.
22140           strQry = "qryRpt_AssetHistory_44"
22150           strRptAcct = .cmbAccounts
22160           strRptCap = "Asset_History_" & strRptAcct & "_" & strCusip & "_" & _
                  Format(.DateStart, "yymmdd") & "_To_" & Format(.DateEnd, "yymmdd")
22170         Case .opgAccountNumber_optAll.OptionValue
                ' ** qryRpt_AssetHistory_All_43 (xx), For Export, All Accounts.
22180           strQry = "qryRpt_AssetHistory_All_44"
22190           strRptAcct = vbNullString
22200           strRptCap = "Asset_History_All_" & strCusip & "_" & _
                  Format(.DateStart, "yymmdd") & "_To_" & Format(.DateEnd, "yymmdd")
22210         End Select

22220         strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

22230         If strRptPathFile <> vbNullString Then
22240           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
22250             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
22260           End If
22270           DoEvents
22280           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
22290             Kill strRptPathFile
22300           End If
22310           DoEvents
22320           gstrFormQuerySpec = THIS_NAME
22330           DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
22340           If Excel_NameOnly(strRptPathFile, "History") = True Then  ' ** Module Function: modExcelFuncs.
22350             If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
22360               EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
22370             End If
22380             OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
22390           End If
22400           .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
22410         End If

22420       End If  ' ** DoReport().
22430     End If  ' ** blnContinue.

22440     DoCmd.Hourglass False

22450   End With

EXITP:
22460   Exit Sub

ERRH:
22470   DoCmd.Hourglass False
22480   Select Case ERR.Number
        Case Else
22490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22500   End Select
22510   Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22620   intShiftDown = (Shift And acShiftMask) > 0
22630   intAltDown = (Shift And acAltMask) > 0
22640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22660     Select Case intRetVal
          Case vbKeyTab
22670       With Me
22680         intRetVal = 0
22690         .cmdClose.SetFocus
22700       End With
22710     End Select
22720   End If

        ' ** Shift keys.
22730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22740     Select Case intRetVal
          Case vbKeyTab
22750       With Me
22760         intRetVal = 0
22770         .cmdWord.SetFocus
22780       End With
22790     End Select
22800   End If

EXITP:
22810   KeyCode = intRetVal
22820   Exit Sub

ERRH:
22830   intRetVal = 0
22840   THAT_PROC = THIS_PROC
22850   That_Erl = Erl
22860   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
22870   Resume EXITP

End Sub

Private Function DoReport() As Boolean
' ** CBX_AST_ASTNO As Integer = 0
' ** CBX_AST_TDESC As Integer = 1
' ** CBX_AST_CUSIP As Integer = 2
' ** CBX_AST_ADESC As Integer = 3
' ** CBX_AST_ATYPE As Integer = 4

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, rst As DAO.Recordset, qdf As DAO.QueryDef
        Dim blnRetVal As Boolean

22910   blnRetVal = True

22920   With Me

22930     If IsNull(.cmbAssets) = True Then
22940       blnRetVal = False
22950       MsgBox "You must select an asset to continue.", vbInformation + vbOKOnly, "Entry Required"
22960     Else
22970       If IsNull(.cmbAssets.Column(CBX_AST_TDESC)) = True Then
22980         blnRetVal = False
22990         MsgBox "You must select an asset to continue.", vbInformation + vbOKOnly, "Entry Required"
23000       Else
23010         If IsNull(.DateStart) Or IsNull(.DateEnd) Then
23020           blnRetVal = False
23030           MsgBox "You must enter appropriate dates to continue.", vbInformation + vbOKOnly, "Entry Required"
23040         Else
23050           If cmbAccounts.Enabled = True Then
23060             cmbAccounts.SetFocus
23070             If cmbAccounts.text = vbNullString Then
23080               blnRetVal = False
23090               MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
23100             End If
23110           End If
23120         End If
23130       End If
23140     End If

23150     If blnRetVal = True Then

23160       DoCmd.Hourglass True
23170       DoEvents

23180       gstrFormQuerySpec = THIS_NAME
23190       gdatStartDate = .DateStart
23200       gdatEndDate = .DateEnd

23210       Set dbs = CurrentDb
23220       Select Case .opgAccountNumber.Value
            Case .opgAccountNumber_optSpecified.OptionValue
              ' ** Specific account.
              ' **
23230         Set qdf = dbs.QueryDefs("qryAssetHistory_11")
23240         GetBal_Beg Me.cmbAccounts.Column(0)  ' ** Procedure: Below.
23250         GetBal_End Me.cmbAccounts.Column(0)  ' ** Procedure: Below.
23260       Case .opgAccountNumber_optAll.OptionValue
              ' ** All.
              ' **
23270         Set qdf = dbs.QueryDefs("qryAssetHistory_20")
23280         gstrAccountNo = vbNullString
23290       End Select

            ' ** Archived records are included if chkIncludeArchive = True.
            ' ** These queries use chkIncludeArchive:
            ' **   qryAssetHistory_03
            ' **   qryAssetHistory_03_bak
            ' **   qryAssetHistory_07
            ' **   qryAssetHistory_07_bak
            ' **   qryAssetHistory_12
            ' **   qryAssetHistory_16

23300       Set rst = qdf.OpenRecordset
23310       With rst
23320         If .BOF = True And .EOF Then
23330           blnRetVal = False
23340           DoCmd.Hourglass False
23350           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
23360         End If
23370         .Close
23380       End With
23390       dbs.Close

23400       DoCmd.Hourglass False

23410     End If  ' ** blnRetVal.

23420   End With

EXITP:
23430   Set rst = Nothing
23440   Set qdf = Nothing
23450   Set dbs = Nothing
23460   DoReport = blnRetVal
23470   Exit Function

ERRH:
23480   DoCmd.Hourglass False
23490   blnRetVal = False
23500   THAT_PROC = THIS_PROC
23510   That_Erl = Erl
23520   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23530   Resume EXITP

End Function

Public Sub GetBal_Beg(strAccountNo As String)

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "GetBal_Beg"

        Dim dbs As DAO.Database, rst As DAO.Recordset, qdf As DAO.QueryDef

23610   With Me
23620     gstrAccountNo = strAccountNo
23630     gstrFormQuerySpec = THIS_NAME
23640     Set dbs = CurrentDb
          ' ** Get a beginning balance.
23650     Set qdf = dbs.QueryDefs("qryAssetHistory_24")  ' ** Ledger, by accountno.
23660     Set rst = qdf.OpenRecordset
23670     If rst.BOF = True And rst.EOF = True Then
            ' ** No transactions listed in Ledger.
23680       .PreviousShareface = 0#
23690       .PreviousIcash = 0@
23700       .PreviousPcash = 0@
23710       .PreviousCost = 0@
23720     Else
23730       rst.MoveFirst
23740       If IsNull(rst![Sumshareface_dp]) = False Then
23750         .PreviousShareface = (rst![Sumshareface_dp] - rst![Sumshareface_ws])
23760         .PreviousIcash = rst![SumIcash]
23770         .PreviousPcash = rst![SumPcash]
23780         .PreviousCost = rst![SumCost]
23790       Else
23800         .PreviousShareface = 0#
23810         .PreviousIcash = 0@
23820         .PreviousPcash = 0@
23830         .PreviousCost = 0@
23840       End If
23850     End If
23860     rst.Close
23870     Set qdf = dbs.QueryDefs("qryAssetHistory_28")  ' ** Archive, by accountno.
23880     Set rst = qdf.OpenRecordset
23890     If rst.BOF = True And rst.EOF = True Then
            ' ** No transactions listed in LedgerArchive.
23900     Else
23910       rst.MoveFirst
23920       If IsNull(rst![Sumshareface_dp]) = False Then
23930         .PreviousShareface = (.PreviousShareface + (rst![Sumshareface_dp] - rst![Sumshareface_ws]))
23940         .PreviousIcash = (.PreviousIcash + rst![SumIcash])
23950         .PreviousPcash = (.PreviousPcash + rst![SumPcash])
23960         .PreviousCost = (.PreviousCost + rst![SumCost])
23970       End If
23980     End If
23990     rst.Close
24000   End With

EXITP:
24010   Set rst = Nothing
24020   Set qdf = Nothing
24030   Set dbs = Nothing
24040   Exit Sub

ERRH:
24050   THAT_PROC = THIS_PROC
24060   That_Erl = Erl
24070   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
24080   Resume EXITP

End Sub

Public Sub GetBal_End(strAccountNo As String)

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "GetBal_End"

        Dim dbs As DAO.Database, rst As DAO.Recordset, qdf As DAO.QueryDef

24110   With Me
24120     gstrAccountNo = strAccountNo
24130     gstrFormQuerySpec = THIS_NAME
24140     Set dbs = CurrentDb
          ' ** Get an ending balance.
24150     Set qdf = dbs.QueryDefs("qryAssetHistory_44")  ' ** Ledger, by accountno.
24160     Set rst = qdf.OpenRecordset
24170     If rst.BOF = True And rst.EOF = True Then
            ' ** No transactions listed in Ledger.
24180       .EndingShareface = 0#
24190       .EndingICash = 0@
24200       .EndingPCash = 0@
24210       .EndingCost = 0@
24220     Else
24230       rst.MoveFirst
24240       If IsNull(rst![Sumshareface_dp]) = False Then
24250         .EndingShareface = (rst![Sumshareface_dp] - rst![Sumshareface_ws])
24260         .EndingICash = rst![SumIcash]
24270         .EndingPCash = rst![SumPcash]
24280         .EndingCost = rst![SumCost]
24290       Else
24300         .EndingShareface = 0#
24310         .EndingICash = 0@
24320         .EndingPCash = 0@
24330         .EndingCost = 0@
24340       End If
24350     End If
24360     rst.Close
24370     Set qdf = dbs.QueryDefs("qryAssetHistory_48")  ' ** Archive, by accountno.
24380     Set rst = qdf.OpenRecordset
24390     If rst.BOF = True And rst.EOF = True Then
            ' ** No transactions listed in LedgerArchive.
24400     Else
24410       rst.MoveFirst
24420       If IsNull(rst![Sumshareface_dp]) = False Then
24430         .EndingShareface = (.EndingShareface + (rst![Sumshareface_dp] - rst![Sumshareface_ws]))
24440         .EndingICash = (.EndingICash + rst![SumIcash])
24450         .EndingPCash = (.EndingPCash + rst![SumPcash])
24460         .EndingCost = (.EndingCost + rst![SumCost])
24470       End If
24480     End If
24490     rst.Close
24500   End With

EXITP:
24510   Set rst = Nothing
24520   Set qdf = Nothing
24530   Set dbs = Nothing
24540   Exit Sub

ERRH:
24550   THAT_PROC = THIS_PROC
24560   That_Erl = Erl
24570   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
24580   Resume EXITP

End Sub

Private Sub SetArchiveOption(blnEnabled As Boolean)

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "SetArchiveOption"

        Dim blnArchOpts As Boolean
        Dim lngX As Long

24610   With Me
          ' ** Only enable if there are transactions in LedgerArchive.
24620     If lngAcctArchs > 0 Then
24630       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
              ' ** Specified account.
24640         blnArchOpts = False
24650         If Nz(.cmbAccounts.Column(0), vbNullString) <> vbNullString Then
                ' ***********************************************
                ' ** Array: arr_varAcctArch()
                ' **
                ' **   Field  Element  Name         Constant
                ' **   =====  =======  ===========  ===========
                ' **     1       0     accountno    A_ACTNO
                ' **     2       1     transdate    A_TDATE
                ' **     3       2     cnt          A_CNT
                ' **
                ' ***********************************************
24660           For lngX = 0& To (lngAcctArchs - 1&)
24670             If arr_varAcctArch(A_ACTNO, lngX) = .cmbAccounts.Column(0) Then
                    ' ** This account's got archived transactions, so normal settings.
24680               blnArchOpts = True
24690               Exit For
24700             End If
24710           Next
24720         Else
                ' ** Nothing selected yet, so normal settings.
24730           blnArchOpts = True
24740         End If
24750       Case .opgAccountNumber_optAll.OptionValue
              ' ** All accounts, normal settings.
24760         blnArchOpts = True
24770       End Select
            ' ** Normal settings.
24780       Select Case blnArchOpts
            Case True
24790         .chkIncludeArchive.Enabled = True
24800         .chkIncludeArchive_lbl.Visible = True
24810         .chkIncludeArchive_lbl2_dim.Visible = False
24820         .chkIncludeArchive_lbl2_dim_hi.Visible = False
24830         .chkIncludeArchive_lbl3_dim.Visible = False
24840         .chkIncludeArchive_lbl3_dim_hi.Visible = False
24850       Case False
24860         .chkIncludeArchive.Enabled = False
24870         .chkIncludeArchive_lbl.Visible = False
24880         .chkIncludeArchive_lbl2_dim.Visible = False
24890         .chkIncludeArchive_lbl2_dim_hi.Visible = False
24900         .chkIncludeArchive_lbl3_dim.Visible = True
24910         .chkIncludeArchive_lbl3_dim_hi.Visible = True
24920       End Select
24930     Else
24940       .chkIncludeArchive.Enabled = False
24950       .chkIncludeArchive_lbl.Visible = False
24960       .chkIncludeArchive_lbl2_dim.Visible = True
24970       .chkIncludeArchive_lbl2_dim_hi.Visible = True
24980       .chkIncludeArchive_lbl3_dim.Visible = False
24990       .chkIncludeArchive_lbl3_dim_hi.Visible = False
25000     End If
25010   End With

EXITP:
25020   Exit Sub

ERRH:
25030   Select Case ERR.Number
        Case Else
25040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25050   End Select
25060   Resume EXITP

End Sub

Private Sub SetHiddenAssets()

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "SetHiddenAssets"

        Dim varTmp00 As Variant

25110   With Me
          ' ** MasterAsset, just Hidden Assets.
25120     varTmp00 = DCount("*", "qryRpt_AssetHistory_02")
25130     If IsNull(varTmp00) = True Then
25140       .chkIncludeHiddenAssets.Enabled = False
25150       .chkIncludeHiddenAssets_lbl.Visible = False
25160       .chkIncludeHiddenAssets_lbl2_dim.Visible = True
25170       .chkIncludeHiddenAssets_lbl2_dim_hi.Visible = True
25180     Else
25190       If varTmp00 = 0 Then
25200         .chkIncludeHiddenAssets.Enabled = False
25210         .chkIncludeHiddenAssets_lbl.Visible = False
25220         .chkIncludeHiddenAssets_lbl2_dim.Visible = True
25230         .chkIncludeHiddenAssets_lbl2_dim_hi.Visible = True
25240       Else
25250         .chkIncludeHiddenAssets.Enabled = True
25260         .chkIncludeHiddenAssets_lbl.Visible = True
25270         .chkIncludeHiddenAssets_lbl2_dim.Visible = False
25280         .chkIncludeHiddenAssets_lbl2_dim_hi.Visible = False
25290       End If
25300     End If
25310   End With

EXITP:
25320   Exit Sub

ERRH:
25330   Select Case ERR.Number
        Case Else
25340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25350   End Select
25360   Resume EXITP

End Sub

Private Sub SetHiddenTrans()

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "SetHiddenTrans"

        Dim varTmp00 As Variant

25410   With Me
          ' ** qryRpt_AssetHistory_03_03 (Union of qryRpt_AssetHistory_03_01
          ' ** (Ledger, just Hidden Transactions), qryRpt_AssetHistory_03_02
          ' ** (LedgerArchive, just Hidden Transactions)), grouped, with cnt_hid.
25420     varTmp00 = DLookup("[cnt_hid]", "qryRpt_AssetHistory_03_04")
25430     If IsNull(varTmp00) = True Then
25440       .chkIncludeHiddenTrans.Enabled = False
25450       .chkIncludeHiddenTrans_lbl.Visible = False
25460       .chkIncludeHiddenTrans_lbl2_dim.Visible = True
25470       .chkIncludeHiddenTrans_lbl2_dim_hi.Visible = True
25480     Else
25490       If varTmp00 = 0 Then
25500         .chkIncludeHiddenTrans.Enabled = False
25510         .chkIncludeHiddenTrans_lbl.Visible = False
25520         .chkIncludeHiddenTrans_lbl2_dim.Visible = True
25530         .chkIncludeHiddenTrans_lbl2_dim_hi.Visible = True
25540       Else
25550         .chkIncludeHiddenTrans.Enabled = True
25560         .chkIncludeHiddenTrans_lbl.Visible = True
25570         .chkIncludeHiddenTrans_lbl2_dim.Visible = False
25580         .chkIncludeHiddenTrans_lbl2_dim_hi.Visible = False
25590       End If
25600     End If
25610   End With

EXITP:
25620   Exit Sub

ERRH:
25630   Select Case ERR.Number
        Case Else
25640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25650   End Select
25660   Resume EXITP

End Sub

Private Sub ArrowsOff()

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "ArrowsOff"

25710   With Me
25720     .GoToReport_arw_ah_spec_onon_img.Visible = False
25730     .GoToReport_arw_ah_spec_onoff_img.Visible = False
25740     .GoToReport_arw_ah_spec_ondisnothis_img.Visible = False
25750     .GoToReport_arw_ah_spec_ondisnoarch_img.Visible = False
25760     .GoToReport_arw_ah_spec_offon_img.Visible = False
25770     .GoToReport_arw_ah_spec_offoff_img.Visible = False
25780     .GoToReport_arw_ah_spec_offdisnothis_img.Visible = False
25790     .GoToReport_arw_ah_spec_offdisnoarch_img.Visible = False
25800     .GoToReport_arw_ah_all_on_img.Visible = False
25810     .GoToReport_arw_ah_all_ondis_img.Visible = False
25820     .GoToReport_arw_ah_all_off_img.Visible = False
25830     .GoToReport_arw_ah_all_offdis_img.Visible = False
25840     .GoToReport_arw_ah_all_dis_img.Visible = False
25850     .GoToReport_arw_ah_all_disdis_img.Visible = False
25860   End With

EXITP:
25870   Exit Sub

ERRH:
25880   DoCmd.Hourglass False
25890   Select Case ERR.Number
        Case Else
25900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25910   End Select
25920   Resume EXITP

End Sub
