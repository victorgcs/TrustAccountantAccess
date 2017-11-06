VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Checks_Bank1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Checks_Bank1"

'VGC 07/08/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Account Number:   A {chkbank_acctnum}
' **   Bank Name:        N {chkbank_name}
' **   OK:               O {cmdOK}

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
190           strAccountNo = Left(strTmp01, (intPos01 - 1))
200           blnAll = CBool(Mid(strTmp01, (intPos01 + 1)))

210           .PictureData = .Form_Rpt_Checks_Bank1_img

220           .chkbank_name.SetFocus

230         Else
240           Cancel = -1
250         End If
260       Else
270         Cancel = -1
280       End If
290     End With

300     DoCmd.Hourglass False

310     If Cancel = -1 Then
320       Beep
330       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
340     End If

EXITP:
350     Exit Sub

ERRH:
360     DoCmd.Hourglass False
370     Select Case ERR.Number
        Case Else
380       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
390     End Select
400     Resume EXITP

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
            ' ** Do not allow Escape (what would it mean?).
570         intRetVal = 0
580       End Select
590     End If

        ' ** Ctrl-Shift keys.
600     If intCtrlDown And (Not intAltDown) And intShiftDown Then
610       Select Case intRetVal
          Case vbKeyF
620         With Me
630           intRetVal = 0
640           .FocusHolder.SetFocus
650         End With
660       Case vbKeyX
            ' ** To close the form in an emergency.
670         intRetVal = 0
680         gblnMessage = False
690         DoCmd.Close acForm, THIS_NAME
700       End Select
710     End If

EXITP:
720     KeyCode = intRetVal
730     Exit Sub

ERRH:
740     intRetVal = 0
750     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
760     Case Else
770       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
780     End Select
790     Resume EXITP

End Sub

Private Sub Form_Timer()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

810     With Me
820       .TimerInterval = 0&

830     End With

EXITP:
840     Exit Sub

ERRH:
850     Select Case ERR.Number
        Case Else
860       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
870     End Select
880     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

910     gstrReturningForm = THIS_NAME
920     gblnSetFocus = True
930     Select Case strCallingForm
        Case "frmRpt_Checks"
940       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
950         DoCmd.SelectObject acForm, strCallingForm, False
960       Else
970         DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
980       End If
990     End Select

EXITP:
1000    Exit Sub

ERRH:
1010    DoCmd.Hourglass False
1020    Select Case ERR.Number
        Case Else
1030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1040    End Select
1050    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form
        Dim strBankName As String, strBankAcctNum As String, strMsg As String
        Dim lngAccts As Long, arr_varAcct As Variant
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean, blnFound As Boolean, blnAdd As Boolean
        Dim lngX As Long

        ' ** Array: arr_varAcct().
        Const A_ACTNO As Integer = 0
        Const A_CNT   As Integer = 1

1110    With Me

1120      DoCmd.Hourglass True
1130      DoEvents

1140      blnContinue = True

1150      If IsNull(.chkbank_name) = True Then
1160        blnContinue = False
1170        DoCmd.Hourglass False
1180        MsgBox "A bank name is required to continue.", vbInformation + vbOKOnly, "Entry Required"
1190        .chkbank_name.SetFocus
1200      Else
1210        If Trim(.chkbank_name) = vbNullString Then
1220          blnContinue = False
1230          DoCmd.Hourglass False
1240          MsgBox "A bank name is required to continue.", vbInformation + vbOKOnly, "Entry Required"
1250          .chkbank_name.SetFocus
1260        Else
1270          strBankName = Trim(.chkbank_name)
1280          If IsNull(.chkbank_acctnum) = True Then
1290            blnContinue = False
1300            DoCmd.Hourglass False
1310            MsgBox "An account number is required to continue.", vbInformation + vbOKOnly, "Entry Required"
1320            .chkbank_acctnum.SetFocus
1330          Else
1340            If Trim(.chkbank_acctnum) = vbNullString Then
1350              blnContinue = False
1360              DoCmd.Hourglass False
1370              MsgBox "An account number is required to continue.", vbInformation + vbOKOnly, "Entry Required"
1380              .chkbank_acctnum.SetFocus
1390            Else
                  ' ** Good to go.
1400              strBankAcctNum = Trim(.chkbank_acctnum)
1410            End If
1420          End If
1430        End If
1440      End If

1450      If blnContinue = True Then

1460        Set frm = Forms(strCallingForm)

1470        Set dbs = CurrentDb
1480        With dbs

1490          If strAccountNo = "All" Then
                ' ** qryPrintChecks_05_07_04 (xx), grouped by accountno, with cnt_chk.
1500            Set qdf = dbs.QueryDefs("qryPrintChecks_05_07_08")
1510            Set rst = qdf.OpenRecordset
1520            With rst
1530              .MoveLast
1540              lngAccts = .RecordCount
1550              .MoveFirst
1560              arr_varAcct = .GetRows(lngAccts)
1570              .Close
1580            End With
1590            Set rst = Nothing
1600            Set qdf = Nothing
1610            DoEvents
1620          End If  ' ** All.

              ' ** This form is always for just 1 accountno. (Except non-separate checking accounts!)
1630          blnFound = False: blnAdd = False
1640          Set rst = .OpenRecordset("tblCheckBank", dbOpenDynaset, dbConsistent)
1650          With rst
1660            If strAccountNo = "All" Then
                  ' ** Used for gblnSeparateCheckingAccounts = False, so
                  ' ** all accounno's will get the same bank and accountno.
1670              For lngX = 0& To (lngAccts - 1&)
1680                blnAdd = False
1690                If .BOF = True And .EOF = True Then
1700                  blnAdd = True
1710                Else
1720                  .FindFirst "[accountno] = '" & arr_varAcct(A_ACTNO, lngX) & "'"
1730                  Select Case .NoMatch
                      Case True
1740                    blnAdd = True
1750                  Case False
1760                    If ![chkbank_name] <> strBankName Then
1770                      .Edit
1780                      ![chkbank_name] = strBankName
1790                      ![chkbank_datemodified] = Now()
1800                      .Update
1810                    End If
1820                    If ![chkbank_acctnum] <> strBankAcctNum Then
1830                      .Edit
1840                      ![chkbank_acctnum] = strBankAcctNum
1850                      ![chkbank_datemodified] = Now()
1860                      .Update
1870                    End If
1880                  End Select
1890                End If
1900                If blnAdd = True Then
1910                  .AddNew
                      ' ** ![chkbank_id] : AutoNumber.
1920                  ![chkbank_name] = strBankName
1930                  ![chkbank_acctnum] = strBankAcctNum
1940                  ![accountno] = arr_varAcct(A_ACTNO, lngX)
1950                  ![chkbank_active] = True
1960                  ![chkbank_datemodified] = Now()
1970                  .Update
1980                End If
1990              Next  ' ** lngX.
2000              frm.chkbank_name = strBankName
2010              frm.chkbank_acctnum = strBankAcctNum
2020            Else
2030              If .BOF = True And .EOF = True Then
2040                blnAdd = True
2050              Else
2060                .FindFirst "[accountno] = '" & strAccountNo & "' And [chkbank_acctnum] = '" & strBankAcctNum & "'"
2070                Select Case .NoMatch
                    Case True
2080                  blnAdd = True
2090                Case False
2100                  If ![chkbank_name] = strBankName Then
                        ' ** Hmm... It's already there.
2110                    blnFound = True
2120                    frm.chkbank_name = ![chkbank_name]
2130                    frm.chkbank_acctnum = ![chkbank_acctnum]
2140                  Else
2150                    strMsg = "The account number you entered, " & strBankAcctNum & ", matches one for the bank:" & vbCrLf & _
                          ![chkbank_name] & vbCrLf & vbCrLf & "Should the name you entered, '" & strBankName & "', replace the name above?"
2160                    DoCmd.Hourglass False
2170                    msgResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Similar Entry Found")
2180                    DoCmd.Hourglass True
2190                    DoEvents
2200                    Select Case msgResponse
                        Case vbYes
2210                      .Edit
2220                      ![chkbank_name] = strBankName
2230                      ![chkbank_datemodified] = Now()
2240                      .Update
2250                      blnFound = True
2260                      frm.chkbank_name = ![chkbank_name]
2270                      frm.chkbank_acctnum = ![chkbank_acctnum]
2280                    Case Else
2290                      blnAdd = True
2300                    End Select
2310                  End If
2320                End Select
2330              End If
2340              If blnAdd = True Then
2350                .AddNew
                    ' ** ![chkbank_id] : AutoNumber.
2360                ![chkbank_name] = strBankName
2370                ![chkbank_acctnum] = strBankAcctNum
2380                ![accountno] = strAccountNo
2390                ![chkbank_active] = True
2400                ![chkbank_datemodified] = Now()
2410                .Update
2420                .Bookmark = .LastModified
2430                frm.chkbank_name = ![chkbank_name]
2440                frm.chkbank_acctnum = ![chkbank_acctnum]
2450                glngTaxCode_Distribution = ![chkbank_id]  ' ** Borrowing this variable.
2460              End If
2470            End If  ' ** All.
2480            .Close
2490          End With  ' ** rst.
2500          Set rst = Nothing

2510          .Close
2520        End With  ' ** dbs.

2530        DoCmd.Close acForm, THIS_NAME

2540      End If  ' ** blnContinue.

2550    End With

EXITP:
2560    Set frm = Nothing
2570    Set rst = Nothing
2580    Set qdf = Nothing
2590    Set dbs = Nothing
2600    Exit Sub

ERRH:
2610    DoCmd.Hourglass False
2620    Select Case ERR.Number
        Case Else
2630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2640    End Select
2650    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2720    intShiftDown = (Shift And acShiftMask) > 0
2730    intAltDown = (Shift And acAltMask) > 0
2740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyTab
2770        With Me
2780          intRetVal = 0
2790          .chkbank_name.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab
2850        With Me
2860          intRetVal = 0
2870          .chkbank_acctnum.SetFocus
2880        End With
2890      End Select
2900    End If

EXITP:
2910    KeyCode = intRetVal
2920    Exit Sub

ERRH:
2930    intRetVal = 0
2940    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2950    Case Else
2960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2970    End Select
2980    Resume EXITP

End Sub

Private Sub chkbank_name_KeyDown(KeyCode As Integer, Shift As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3020    intShiftDown = (Shift And acShiftMask) > 0
3030    intAltDown = (Shift And acAltMask) > 0
3040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3060      Select Case intRetVal
          Case vbKeyTab
3070        With Me
3080          intRetVal = 0
3090          .chkbank_acctnum.SetFocus
3100        End With
3110      End Select
3120    End If

        ' ** Shift keys.
3130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3140      Select Case intRetVal
          Case vbKeyTab
3150        With Me
3160          intRetVal = 0
3170          .cmdOK.SetFocus
3180        End With
3190      End Select
3200    End If

EXITP:
3210    KeyCode = intRetVal
3220    Exit Sub

ERRH:
3230    intRetVal = 0
3240    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3250    Case Else
3260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3270    End Select
3280    Resume EXITP

End Sub

Private Sub chkbank_acctnum_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkbank_acctnum_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3320    intShiftDown = (Shift And acShiftMask) > 0
3330    intAltDown = (Shift And acAltMask) > 0
3340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3360      Select Case intRetVal
          Case vbKeyTab
3370        With Me
3380          intRetVal = 0
3390          .cmdOK.SetFocus
3400        End With
3410      End Select
3420    End If

        ' ** Shift keys.
3430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3440      Select Case intRetVal
          Case vbKeyTab
3450        With Me
3460          intRetVal = 0
3470          .chkbank_name.SetFocus
3480        End With
3490      End Select
3500    End If

EXITP:
3510    KeyCode = intRetVal
3520    Exit Sub

ERRH:
3530    intRetVal = 0
3540    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3550    Case Else
3560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3570    End Select
3580    Resume EXITP

End Sub
