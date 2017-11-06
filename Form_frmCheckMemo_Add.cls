VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCheckMemo_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCheckMemo_Add"

'VGC 06/11/2012: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Memo:             M {ChkMemo_Memo}
' **   OK:               O {cmdOK}

Private lngNextMemoID As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           .ChkMemo_Memo.SetFocus

160         Else
170           Cancel = -1
180         End If
190       Else
200         Cancel = -1
210       End If
220     End With

230     If Cancel = -1 Then
240       Beep
250       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
260     End If

EXITP:
270     Exit Sub

ERRH:
280     Select Case ERR.Number
        Case Else
290       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
300     End Select
310     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

410     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
420     intShiftDown = (Shift And acShiftMask) > 0
430     intAltDown = (Shift And acAltMask) > 0
440     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
450     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
460       Select Case intRetVal
          Case vbKeyEscape
470         intRetVal = 0
480         DoCmd.Close acForm, THIS_NAME
490         DoCmd.SelectObject acForm, strCallingForm, False
500       End Select
510     End If

EXITP:
520     KeyCode = intRetVal
530     Exit Sub

ERRH:
540     intRetVal = 0
550     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
560     Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdOK_Click()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strMemo As String, lngMemoType As Long
        Dim blnContinue As Boolean

610     blnContinue = True

620     With Me
630       If IsNull(.ChkMemo_Memo) = False Then
640         If Trim(.ChkMemo_Memo) <> vbNullString Then
650           .ChkMemo_Memo.SetFocus
660           If InStr(.ChkMemo_Memo.text, Chr(34)) > 0 Then
670             blnContinue = False
680             MsgBox "Cannot have quote marks in Memo.", vbInformation + vbOKOnly, "Invalid Characters"
690             DoCmd.CancelEvent
700           Else
710             strMemo = Trim(.ChkMemo_Memo)
720             If IsNull(.ChkMemoType_Type) = True Then
730               lngMemoType = 0&  ' ** Default: Misc.
740             Else
750               lngMemoType = .ChkMemoType_Type
760             End If
770             Set dbs = CurrentDb
780             With dbs
790               Set rst = .OpenRecordset("tblCheckMemo", dbOpenDynaset, dbConsistent)
800               With rst
810                 .FindFirst "[ChkMemo_Memo] = '" & strMemo & "'"
820                 If .NoMatch = True Then
830                   .AddNew
840                   ![ChkMemoType_Type] = lngMemoType
850                   ![ChkMemo_Memo] = strMemo
860                   ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
870                   ![ChkMemo_DateModified] = Now()
880                   .Update
890                   .Bookmark = .LastModified
900                   lngNextMemoID = ![ChkMemo_ID]
910                   Forms("frmCheckMemo").frmCheckMemo_Sub.Form.Refresh
920                   Forms("frmCheckMemo").frmCheckMemo_Sub.Form.Requery
930                 Else
940                   blnContinue = False
950                   MsgBox "The Memo, '" & strMemo & "', is already present.", vbInformation + vbOKOnly, "Duplicate Check Memo"
960                   DoCmd.CancelEvent
970                 End If
980               End With
990               .Close
1000            End With
1010          End If
1020        Else
1030          blnContinue = False
1040          MsgBox "Memo must be filled to continue.", vbInformation + vbOKOnly, "Entry Required"
1050          DoCmd.CancelEvent
1060        End If
1070      Else
1080        blnContinue = False
1090        MsgBox "Memo must be filled to continue.", vbInformation + vbOKOnly, "Entry Required"
1100        DoCmd.CancelEvent
1110      End If

1120    End With

1130    If blnContinue = True Then
1140      DoCmd.Close acForm, THIS_NAME
1150    End If

EXITP:
1160    Set rst = Nothing
1170    Set dbs = Nothing
1180    Exit Sub

ERRH:
1190    Select Case ERR.Number
        Case Else
1200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1210    End Select
1220    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
          Case vbKeyTab
1370        With Me
1380          intRetVal = 0
1390          .cmdCancel.SetFocus
1400        End With
1410      End Select
1420    End If

        ' ** Shift keys.
1430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1440      Select Case intRetVal
          Case vbKeyTab
1450        With Me
1460          intRetVal = 0
1470          .ChkMemo_Memo.SetFocus
1480        End With
1490      End Select
1500    End If

EXITP:
1510    KeyCode = intRetVal
1520    Exit Sub

ERRH:
1530    intRetVal = 0
1540    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
1550    Case Else
1560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1570    End Select
1580    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

1610    DoCmd.Close acForm, THIS_NAME

EXITP:
1620    Exit Sub

ERRH:
1630    Select Case ERR.Number
        Case Else
1640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1650    End Select
1660    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1720    intShiftDown = (Shift And acShiftMask) > 0
1730    intAltDown = (Shift And acAltMask) > 0
1740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1760      Select Case intRetVal
          Case vbKeyTab
1770        With Me
1780          intRetVal = 0
1790          .ChkMemoType_Type.SetFocus
1800        End With
1810      End Select
1820    End If

        ' ** Shift keys.
1830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1840      Select Case intRetVal
          Case vbKeyTab
1850        With Me
1860          intRetVal = 0
1870          .cmdOK.SetFocus
1880        End With
1890      End Select
1900    End If

EXITP:
1910    KeyCode = intRetVal
1920    Exit Sub

ERRH:
1930    intRetVal = 0
1940    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
1950    Case Else
1960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1970    End Select
1980    Resume EXITP

End Sub

Private Sub ChkMemoType_Type_GotFocus()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkMemoType_Type_GotFocus"

2010    With Me
2020      .ChkMemoType_Type.SelLength = 0
2030      .ChkMemoType_Type.SelStart = 19
2040    End With

EXITP:
2050    Exit Sub

ERRH:
2060    Select Case ERR.Number
        Case Else
2070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2080    End Select
2090    Resume EXITP

End Sub

Private Sub ChkMemoType_Type_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkMemoType_Type_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
2170        With Me
2180          intRetVal = 0
2190          .ChkMemo_Memo.SetFocus
2200        End With
2210      End Select
2220    End If

        ' ** Shift keys.
2230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2250        With Me
2260          intRetVal = 0
2270          .cmdCancel.SetFocus
2280        End With
2290      End Select
2300    End If

EXITP:
2310    KeyCode = intRetVal
2320    Exit Sub

ERRH:
2330    intRetVal = 0
2340    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
2350    Case Else
2360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2370    End Select
2380    Resume EXITP

End Sub

Private Sub ChkMemo_Memo_KeyDown(KeyCode As Integer, Shift As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkMemo_Memo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2420    intShiftDown = (Shift And acShiftMask) > 0
2430    intAltDown = (Shift And acAltMask) > 0
2440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2470        With Me
2480          intRetVal = 0
2490          .cmdOK.SetFocus
2500        End With
2510      End Select
2520    End If

        ' ** Shift keys.
2530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2550        With Me
2560          intRetVal = 0
2570          .ChkMemoType_Type.SetFocus
2580        End With
2590      End Select
2600    End If

EXITP:
2610    KeyCode = intRetVal
2620    Exit Sub

ERRH:
2630    intRetVal = 0
2640    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
2650    Case Else
2660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2670    End Select
2680    Resume EXITP

End Sub
