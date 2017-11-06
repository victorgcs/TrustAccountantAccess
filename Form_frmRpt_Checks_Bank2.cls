VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Checks_Bank2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Checks_Bank2"

'VGC 07/24/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Account Number:   A {Bank_AccountNumber on frmRpt_Checks_Bank2_Sub}
' **   Copy Bank Info:   C {cmdCopy}
' **   Bank Name:        N {Bank_Name on frmRpt_Checks_Bank2_Sub}
' **   OK:               O {cmdOK}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmRpt_Checks_Bank2_Sub}

Private blnCopy_Focus As Boolean, blnCopy_MouseDown As Boolean

Private lngRecsCur As Long, lngChkVoid_Set As Long
Private strAccountNo As String, blnAll As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         strTmp01 = Trim(.OpenArgs)
140         If strTmp01 <> vbNullString Then

150           intPos01 = InStr(strTmp01, "~")
160           strCallingForm = Left(strTmp01, (intPos01 - 1))
170           strTmp01 = Mid(strTmp01, (intPos01 + 1))
180           intPos01 = InStr(strTmp01, "~")
190           strAccountNo = Left(strTmp01, (intPos01 - 1))  ' ** "All".
200           strTmp01 = Mid(strTmp01, (intPos01 + 1))
210           intPos01 = InStr(strTmp01, "~")
220           blnAll = CBool(Left(strTmp01, (intPos01 - 1)))
230           lngChkVoid_Set = CLng(Mid(strTmp01, (intPos01 + 1)))

240           .PictureData = .Form_Rpt_Checks_Bank2_img

250           blnCopy_Focus = False: blnCopy_MouseDown = False

260           .chkvoid_set = lngChkVoid_Set
270           gstrFormQuerySpec = THIS_NAME

280           .frmRpt_Checks_Bank2_Sub.SetFocus

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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Account Number:   A {Bank_AccountNumber on frmRpt_Checks_Bank2_Sub}
        ' **   Bank Name:        N {Bank_Name on frmRpt_Checks_Bank2_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmRpt_Checks_Bank2_Sub}

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
            ' ** Do not allow Escape (what would it mean?).
570         intRetVal = 0
580       End Select
590     End If

        ' ** Alt keys.
600     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
610       Select Case intRetVal
          Case vbKeyA
620         With Me
630           intRetVal = 0
640           .frmRpt_Checks_Bank2_Sub.SetFocus
650           .frmRpt_Checks_Bank2_Sub.Form.Bank_AccountNumber.SetFocus
660         End With
670       Case vbKeyN
680         With Me
690           intRetVal = 0
700           .frmRpt_Checks_Bank2_Sub.SetFocus
710           .frmRpt_Checks_Bank2_Sub.Form.Bank_Name.SetFocus
720         End With
730       End Select
740     End If

        ' ** Ctrl keys.
750     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyS
770         With Me
780           intRetVal = 0
790           .frmRpt_Checks_Bank2_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmRpt_Checks_Bank2_Sub.
800         End With
810       End Select
820     End If

        ' ** Ctrl-Shift keys.
830     If intCtrlDown And (Not intAltDown) And intShiftDown Then
840       Select Case intRetVal
          Case vbKeyF
850         With Me
860           intRetVal = 0
870           .FocusHolder.SetFocus
880         End With
890       Case vbKeyX
            ' ** To close the form in an emergency.
900         intRetVal = 0
910         gblnMessage = False
920         DoCmd.Close acForm, THIS_NAME
930       End Select
940     End If

EXITP:
950     KeyCode = intRetVal
960     Exit Sub

ERRH:
970     intRetVal = 0
980     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
990     Case Else
1000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1010    End Select
1020    Resume EXITP

End Sub

Private Sub Form_Timer()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

1110    With Me
1120      .TimerInterval = 0&

1130    End With

EXITP:
1140    Exit Sub

ERRH:
1150    Select Case ERR.Number
        Case Else
1160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1170    End Select
1180    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1210    gstrReturningForm = THIS_NAME
1220    gblnSetFocus = True
1230    Select Case strCallingForm
        Case "frmRpt_Checks"
1240      If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
1250        DoCmd.SelectObject acForm, strCallingForm, False
1260      Else
1270        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
1280      End If
1290    End Select

EXITP:
1300    Exit Sub

ERRH:
1310    DoCmd.Hourglass False
1320    Select Case ERR.Number
        Case Else
1330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1340    End Select
1350    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1410    With Me
1420      If .cmdCopy_raised_focus_dots_img.Visible = True Or .cmdCopy_raised_focus_img.Visible = True Then
1430        Select Case blnCopy_Focus
            Case True
1440          .cmdCopy_raised_semifocus_dots_img.Visible = True
1450          .cmdCopy_raised_img.Visible = False
1460        Case False
1470          .cmdCopy_raised_img.Visible = True
1480          .cmdCopy_raised_semifocus_dots_img.Visible = False
1490        End Select
1500        .cmdCopy_raised_focus_img.Visible = False
1510        .cmdCopy_raised_focus_dots_img.Visible = False
1520        .cmdCopy_sunken_focus_dots_img.Visible = False
1530        .cmdCopy_raised_img_dis.Visible = False
1540      End If
1550    End With

EXITP:
1560    Exit Sub

ERRH:
1570    Select Case ERR.Number
        Case Else
1580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1590    End Select
1600    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strAccountNo As String, strBankName As String, strBankAcctNum As String, strMsg As String
        Dim lngChecks As Long, arr_varCheck() As Variant
        Dim lngItems As Long, arr_varItem() As Variant
        Dim lngRecs As Long, lngChkBankID As Long
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean, blnFound As Boolean, blnAdd As Boolean, blnAddAll As Boolean
        Dim lngX As Long, lngY As Long, lngE As Long

        ' ** Array: arr_varCheck().
        Const C_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const C_ACTNO As Integer = 0
        Const C_BID   As Integer = 1
        Const C_BANK  As Integer = 2
        Const C_ACCT  As Integer = 3

1710    With Me

1720      DoCmd.Hourglass True
1730      DoEvents

1740      blnContinue = True

          ' ** The arr_varItem() array is only those missing bank info!
1750      lngItems = 0&
1760      ReDim arr_varItem(C_ELEMS, 0)

          ' ** The arr_varCheck() array is all checks!
1770      lngChecks = 0&
1780      ReDim arr_varCheck(C_ELEMS, 0)

1790      Set rst = .frmRpt_Checks_Bank2_Sub.Form.RecordsetClone
1800      With rst
1810        .MoveLast
1820        lngRecs = .RecordCount
1830        .MoveFirst
1840        For lngX = 1& To lngRecs
1850          lngE = -1&
1860          If IsNull(![Bank_Name]) = True Then
1870            lngItems = lngItems + 1&
1880            lngE = lngItems - 1&
1890            ReDim Preserve arr_varItem(C_ELEMS, lngE)
1900            arr_varItem(C_ACTNO, lngE) = ![accountno]
1910            arr_varItem(C_BID, lngE) = Null
1920            arr_varItem(C_BANK, lngE) = ![Bank_Name]
1930            arr_varItem(C_ACCT, lngE) = ![Bank_AccountNumber]
1940          Else
1950            If Trim(![Bank_Name]) = vbNullString Then
1960              lngItems = lngItems + 1&
1970              lngE = lngItems - 1&
1980              ReDim Preserve arr_varItem(C_ELEMS, lngE)
1990              arr_varItem(C_ACTNO, lngE) = ![accountno]
2000              arr_varItem(C_BID, lngE) = Null
2010              arr_varItem(C_BANK, lngE) = ![Bank_Name]
2020              arr_varItem(C_ACCT, lngE) = ![Bank_AccountNumber]
2030            End If
2040          End If
2050          If lngE = -1& Then
2060            If IsNull(![Bank_AccountNumber]) = True Then
2070              lngItems = lngItems + 1&
2080              lngE = lngItems - 1&
2090              ReDim Preserve arr_varItem(C_ELEMS, lngE)
2100              arr_varItem(C_ACTNO, lngE) = ![accountno]
2110              arr_varItem(C_BID, lngE) = Null
2120              arr_varItem(C_BANK, lngE) = ![Bank_Name]
2130              arr_varItem(C_ACCT, lngE) = ![Bank_AccountNumber]
2140            Else
2150              If Trim(![Bank_AccountNumber]) = vbNullString Then
2160                lngItems = lngItems + 1&
2170                lngE = lngItems - 1&
2180                ReDim Preserve arr_varItem(C_ELEMS, lngE)
2190                arr_varItem(C_ACTNO, lngE) = ![accountno]
2200                arr_varItem(C_BID, lngE) = Null
2210                arr_varItem(C_BANK, lngE) = ![Bank_Name]
2220                arr_varItem(C_ACCT, lngE) = ![Bank_AccountNumber]
2230              End If
2240            End If
2250          End If
2260          lngChecks = lngChecks + 1&
2270          lngE = lngChecks - 1&
2280          ReDim Preserve arr_varCheck(C_ELEMS, lngE)
2290          arr_varCheck(C_ACTNO, lngE) = ![accountno]
2300          arr_varCheck(C_BID, lngE) = Null
2310          arr_varCheck(C_BANK, lngE) = ![Bank_Name]
2320          arr_varCheck(C_ACCT, lngE) = ![Bank_AccountNumber]
2330          If lngX < lngRecs Then .MoveNext
2340        Next
2350        .Close
2360      End With
2370      Set rst = Nothing

2380      If lngItems > 0& Then
2390        blnContinue = False
2400        strMsg = CStr(lngItems) & " account" & IIf(lngItems = 1&, " is ", "s are ") & "missing bank information." & vbCrLf & vbCrLf & _
              "You must enter the data to continue."
2410        Beep
2420        DoCmd.Hourglass False
2430        MsgBox strMsg, vbInformation + vbOKOnly, "Entry Required"
2440        .frmRpt_Checks_Bank2_Sub.SetFocus
2450        .frmRpt_Checks_Bank2_Sub.Form.MoveRec 0, arr_varCheck(C_ACTNO, 0)  ' ** Form Procedure: frmRpt_Checks_Bank2_Sub.
2460        If IsNull(arr_varItem(C_BANK, 0)) = True Then
2470          .frmRpt_Checks_Bank2_Sub.Form.Bank_Name.SetFocus
2480        Else
2490          If Trim(arr_varItem(C_BANK, 0)) = vbNullString Then
2500            .frmRpt_Checks_Bank2_Sub.Form.Bank_Name.SetFocus
2510          Else
2520            If IsNull(arr_varItem(C_ACCT, 0)) = True Then
2530              .frmRpt_Checks_Bank2_Sub.Form.Bank_AccountNumber.SetFocus
2540            Else
2550              If Trim(arr_varItem(C_ACCT, 0)) = vbNullString Then
2560                .frmRpt_Checks_Bank2_Sub.Form.Bank_AccountNumber.SetFocus
2570              End If
2580            End If
2590          End If
2600        End If
2610      End If  ' ** lngChecks.

2620      If blnContinue = True Then

2630        Set dbs = CurrentDb
2640        With dbs

2650          blnFound = False: blnAdd = False: blnAddAll = False
2660          Set rst = .OpenRecordset("tblCheckBank", dbOpenDynaset, dbConsistent)
2670          With rst
2680            If .BOF = True And .EOF = True Then
2690              blnAddAll = True
2700            End If
2710            For lngX = 0& To (lngChecks - 1&)
2720              strAccountNo = arr_varCheck(C_ACTNO, lngX)
2730              strBankName = arr_varCheck(C_BANK, lngX)
2740              strBankAcctNum = arr_varCheck(C_ACCT, lngX)
2750              Select Case blnAddAll
                  Case True
2760                blnAdd = True
2770              Case False
2780                .FindFirst "[accountno] = '" & strAccountNo & "' And [chkbank_acctnum] = '" & strBankAcctNum & "'"
2790                Select Case .NoMatch
                    Case True
2800                  blnAdd = True
2810                Case False
2820                  arr_varCheck(C_BID, lngX) = ![chkbank_id]
2830                  If ![chkbank_name] = strBankName Then
                        ' ** Hmm... It's already there.
2840                    blnFound = True
2850                  Else
2860                    strMsg = "The account number you entered, " & strBankAcctNum & ", matches one for the bank:" & vbCrLf & _
                          ![chkbank_name] & vbCrLf & vbCrLf & "Should the name you entered, '" & strBankName & "', replace the name above?"
2870                    DoCmd.Hourglass False
2880                    msgResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Similar Entry Found")
2890                    DoCmd.Hourglass True
2900                    DoEvents
2910                    Select Case msgResponse
                        Case vbYes
2920                      .Edit
2930                      ![chkbank_name] = strBankName
2940                      ![chkbank_datemodified] = Now()
2950                      .Update
2960                      blnFound = True
2970                    Case Else
2980                      blnAdd = True
2990                      arr_varCheck(C_BID, lngX) = Null
3000                    End Select
3010                  End If
3020                End Select
3030              End Select
3040              If blnAdd = True Then
3050                .AddNew
                    ' ** ![chkbank_id] : AutoNumber.
3060                ![chkbank_name] = strBankName
3070                ![chkbank_acctnum] = strBankAcctNum
3080                ![accountno] = strAccountNo
3090                ![chkbank_active] = True
3100                ![chkbank_datemodified] = Now()
3110                .Update
3120                .Bookmark = .LastModified
3130                arr_varCheck(C_BID, lngX) = ![chkbank_id]
3140              End If
3150            Next  ' ** lngX.

3160            .Close
3170          End With  ' ** rst.
3180          Set rst = Nothing

3190          .Close
3200        End With  ' ** dbs.

            ' ** I'm going to borrow this variables to get the bank assignments back to frmRpt_Checks.
3210        glngPrintRpts = 0&
3220        ReDim garr_varPrintRpt(C_ELEMS, 0)
3230        For lngX = 0& To (lngChecks - 1&)
3240          glngPrintRpts = glngPrintRpts + 1&
3250          lngE = glngPrintRpts - 1&
3260          ReDim Preserve garr_varPrintRpt(C_ELEMS, lngE)
3270          For lngY = 0 To C_ELEMS
3280            garr_varPrintRpt(lngY, lngE) = arr_varCheck(lngY, lngX)
3290          Next
3300        Next

3310        DoCmd.Close acForm, THIS_NAME

3320      End If  ' ** blnContinue.

3330    End With

EXITP:
        'Set frm = Nothing
3340    Set rst = Nothing
3350    Set dbs = Nothing
3360    Exit Sub

ERRH:
3370    DoCmd.Hourglass False
3380    Select Case ERR.Number
        Case Else
3390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3400    End Select
3410    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3520    intShiftDown = (Shift And acShiftMask) > 0
3530    intAltDown = (Shift And acAltMask) > 0
3540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3560      Select Case intRetVal
          Case vbKeyTab
3570        With Me
3580          intRetVal = 0
3590          .frmRpt_Checks_Bank2_Sub.SetFocus
3600        End With
3610      End Select
3620    End If

        ' ** Shift keys.
3630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3640      Select Case intRetVal
          Case vbKeyTab
3650        With Me
3660          intRetVal = 0
3670          If .cmdCopy.Enabled = True Then
3680            .cmdCopy.SetFocus
3690          Else
3700            .frmRpt_Checks_Bank2_Sub.SetFocus
3710          End If
3720        End With
3730      End Select
3740    End If

EXITP:
3750    KeyCode = intRetVal
3760    Exit Sub

ERRH:
3770    intRetVal = 0
3780    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3790    Case Else
3800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3810    End Select
3820    Resume EXITP

End Sub

Private Sub cmdCopy_GotFocus()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_GotFocus"

3910    With Me
3920      blnCopy_Focus = True
3930      .cmdCopy_raised_semifocus_dots_img.Visible = True
3940      .cmdCopy_raised_img.Visible = False
3950      .cmdCopy_raised_focus_img.Visible = False
3960      .cmdCopy_raised_focus_dots_img.Visible = False
3970      .cmdCopy_sunken_focus_dots_img.Visible = False
3980      .cmdCopy_raised_img_dis.Visible = False
3990    End With

EXITP:
4000    Exit Sub

ERRH:
4010    Select Case ERR.Number
        Case Else
4020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4030    End Select
4040    Resume EXITP

End Sub

Private Sub cmdCopy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_MouseDown"

4110    With Me
4120      blnCopy_MouseDown = True
4130      .cmdCopy_sunken_focus_dots_img.Visible = True
4140      .cmdCopy_raised_img.Visible = False
4150      .cmdCopy_raised_semifocus_dots_img.Visible = False
4160      .cmdCopy_raised_focus_img.Visible = False
4170      .cmdCopy_raised_focus_dots_img.Visible = False
4180      .cmdCopy_raised_img_dis.Visible = False
4190    End With

EXITP:
4200    Exit Sub

ERRH:
4210    Select Case ERR.Number
        Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Sub

Private Sub cmdCopy_Click()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_Click"

        Dim frm As Access.Form, rst As DAO.Recordset
        Dim strBankName As String, strBankAcctNum As String, strAccountNo As String
        Dim msgResponse As VbMsgBoxResult
        Dim lngRecs As Long
        Dim lngX As Long

4310    With Me
4320      Set frm = .frmRpt_Checks_Bank2_Sub.Form
4330      With frm
4340        strAccountNo = .accountno
4350        strBankName = .Bank_Name
4360        strBankAcctNum = .Bank_AccountNumber
4370        msgResponse = MsgBox("Copy bank information from this record to" & vbCrLf & _
              "other voided checks that are missing bank info." & vbCrLf & vbCrLf & _
              String(10, Chr(160)) & "Bank:  " & strBankName & vbCrLf & String(4, Chr(160)) & "Account:  " & strBankAcctNum & vbCrLf & vbCrLf & _
              "Continue?", vbQuestion + vbOKCancel + vbDefaultButton1, "Copy Bank Information")
4380        If msgResponse = vbOK Then
4390          DoCmd.Hourglass True
4400          DoEvents
4410          If IsNull(.Bank_Name) = False And IsNull(.Bank_AccountNumber) = False Then
4420            Set rst = .RecordsetClone
4430            With rst
4440              .MoveLast
4450              lngRecs = .RecordCount
4460              .MoveFirst
4470              For lngX = 1& To lngRecs
4480                If ![accountno] <> strAccountNo And (IsNull(![Bank_Name]) = True Or IsNull(![Bank_AccountNumber]) = True) Then
4490                  .Edit
4500                  ![Bank_Name] = strBankName
4510                  ![Bank_AccountNumber] = strBankAcctNum
4520                  .Update
4530                End If
4540                If lngX < lngRecs Then .MoveNext
4550              Next
4560              .Close
4570            End With
4580            Set rst = Nothing
4590            .Refresh
4600            Beep
4610            DoCmd.Hourglass False
4620          Else
4630            DoCmd.Hourglass False
4640            MsgBox "Bank data is missing.", vbInformation + vbOKOnly, "Nothing To Do"
4650          End If
4660        End If  ' ** msgResponse.
4670      End With  ' ** frm.
4680    End With

EXITP:
4690    Set rst = Nothing
4700    Set frm = Nothing
4710    Exit Sub

ERRH:
4720    DoCmd.Hourglass False
4730    Select Case ERR.Number
        Case Else
4740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4750    End Select
4760    Resume EXITP

End Sub

Private Sub cmdCopy_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_KeyDown"

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
4970          .frmRpt_Checks_Bank2_Sub.SetFocus
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

Private Sub cmdCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_MouseMove"

5110    With Me
5120      If blnCopy_MouseDown = False Then
5130        Select Case blnCopy_Focus
            Case True
5140          .cmdCopy_raised_focus_dots_img.Visible = True
5150          .cmdCopy_raised_focus_img.Visible = False
5160        Case False
5170          .cmdCopy_raised_focus_img.Visible = True
5180          .cmdCopy_raised_focus_dots_img.Visible = False
5190        End Select
5200        .cmdCopy_raised_img.Visible = False
5210        .cmdCopy_raised_semifocus_dots_img.Visible = False
5220        .cmdCopy_sunken_focus_dots_img.Visible = False
5230        .cmdCopy_raised_img_dis.Visible = False
5240      End If
5250    End With

EXITP:
5260    Exit Sub

ERRH:
5270    Select Case ERR.Number
        Case Else
5280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5290    End Select
5300    Resume EXITP

End Sub

Private Sub cmdCopy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_MouseUp"

5410    With Me
5420      .cmdCopy_raised_focus_dots_img.Visible = True
5430      .cmdCopy_raised_img.Visible = False
5440      .cmdCopy_raised_semifocus_dots_img.Visible = False
5450      .cmdCopy_raised_focus_img.Visible = False
5460      .cmdCopy_sunken_focus_dots_img.Visible = False
5470      .cmdCopy_raised_img_dis.Visible = False
5480      blnCopy_MouseDown = False
5490    End With

EXITP:
5500    Exit Sub

ERRH:
5510    Select Case ERR.Number
        Case Else
5520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5530    End Select
5540    Resume EXITP

End Sub

Private Sub cmdCopy_LostFocus()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_LostFocus"

5610    With Me
5620      .cmdCopy_raised_img.Visible = True
5630      .cmdCopy_raised_semifocus_dots_img.Visible = False
5640      .cmdCopy_raised_focus_img.Visible = False
5650      .cmdCopy_raised_focus_dots_img.Visible = False
5660      .cmdCopy_sunken_focus_dots_img.Visible = False
5670      .cmdCopy_raised_img_dis.Visible = False
5680      blnCopy_Focus = False
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
