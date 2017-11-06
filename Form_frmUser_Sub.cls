VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmUser_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME  As String = "frmUser_Sub"

'VGC 04/16/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd on frmUser}
' **   About User:       B {cmdHelp on frmUser}
' **   Close:            C {cmdClose on frmUser}
' **   Delete:           D {cmdDelete on frmUser}
' **   GUID Test:        G {cmdGUIDTest on frmUser}
' **   Show All Users:   H {cmdShowAll on frmUser}
' **   Print Report:     P {cmdPrintReport on frmUser}
' **   User Name:        U {Username}
' **   Change Password:  W {cmdChangePassword on frmUser}
' **   Exit:             X {cmdClose on frmUser}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

'###########################################################################################
'## FOR SOME REASON, THE RECORDSET CLONE METHOD WILL NOT WORK AFTER CHANGING A PASSWORD!
'###########################################################################################

Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[Username]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private usr As DAO.User
Private blnUserNameEdit As Boolean, blnAdmin As Boolean, strSecLev As String
Private strCallingForm As String, lngRecsCur As Long, lngTpp As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmUser" Then

160         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
170         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

180         lngSortLbl_Width = .Sort_lbl.Width

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

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
420     Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub Form_Current()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

510     With Me
520       lngRecsCur = RecCnt  ' ** Function: Below.
530       If lngRecsCur > 0& Then
540         .FocusHolder.SetFocus
550         If .secusr_defaultx = True Then
560           .Parent.cmdClose.SetFocus
570           .Parent.cmdDelete.Enabled = False
580           .Parent.cmdChangePassword.Enabled = False
590           .Employee_Name.Locked = True
600           .Primary_Group.Enabled = False
610           .Primary_Group.Locked = True
620         Else
630           .Parent.cmdDelete.Enabled = True
640           .Parent.cmdChangePassword.Enabled = True
650           .Employee_Name.Locked = False
660           .Primary_Group.Enabled = True
670           .Primary_Group.Locked = False
680         End If
690       Else
700         If .Parent.cmdAdd.Enabled = True Then
710           .Parent.cmdAdd.SetFocus
720         Else
730           .Parent.cmdClose.SetFocus
740         End If
750       End If
760     End With

EXITP:
770     Exit Sub

ERRH:
780     Select Case ERR.Number
        Case Else
790       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
800     End Select
810     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strControl As String
        Dim intRetVal As Integer

910     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
920     intShiftDown = (Shift And acShiftMask) > 0
930     intAltDown = (Shift And acAltMask) > 0
940     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

950   On Error Resume Next
960     strControl = Screen.ActiveControl.Name
970   On Error GoTo ERRH

        ' ** Plain keys.
980     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
990       Select Case intRetVal
          Case vbKeyUp
1000        intRetVal = 0
1010        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1020      Case vbKeyDown
1030        intRetVal = 0
1040        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1050      Case vbKeyDelete
1060        With Me
1070          Select Case strControl
              Case "Username", "Employee_Name", "Primary_Group"
                ' ** Let them delete the contents of the field.
1080          Case Else
1090            intRetVal = 0
1100            If .Parent.cmdDelete.Enabled = True Then
1110              .Parent.cmdDelete_Click  ' ** Form Procedure: frmUser.
1120            Else
1130              Beep
1140            End If
1150          End Select
1160        End With
1170      Case vbKeyEscape
1180        With Me
1190          intRetVal = 0
1200          .Parent.cmdClose_Click  ' ** Form Procedure: frmUser.
1210        End With
1220      End Select
1230    End If

        ' ** Alt keys.
1240    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1250      Select Case intRetVal
          Case vbKeyU
1260        With Me
1270          intRetVal = 0
1280          .Username.SetFocus
1290        End With
1300      Case vbKeyX
1310        With Me
1320          intRetVal = 0
1330          .Parent.cmdClose_Click  ' ** Form Procedure: frmUser.
1340        End With
1350      End Select
1360    End If

        ' ** Ctrl keys.
1370    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1380      Select Case intRetVal
          Case vbKeyS
1390        cmdSave_Click  ' ** Procedure: Below.
1400      Case vbKeyTab, vbKeyReturn
1410        With Me
1420          intRetVal = 0
1430          DoCmd.SelectObject acForm, .Parent.Name, False
1440          If .Parent.cmdChangePassword.Enabled = True Then
1450            .Parent.cmdChangePassword.SetFocus
1460          ElseIf .Parent.cmdAdd.Enabled = True Then
1470            .Parent.cmdAdd.SetFocus
1480          ElseIf .Parent.cmdDelete.Enabled = True Then
1490            .Parent.cmdDelete.SetFocus
1500          Else
1510            .Parent.cmdClose.SetFocus
1520          End If
1530        End With
1540      Case vbKeyPageUp
1550        intRetVal = 0
1560        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1570      Case vbKeyPageDown
1580        intRetVal = 0
1590        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1600      End Select
1610    End If

        ' ** Ctrl-Shift keys.
1620    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1630      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1640        With Me
1650          intRetVal = 0
1660          DoCmd.SelectObject acForm, .Parent.Name, False
1670          .Parent.cmdClose.SetFocus
1680        End With
1690      End Select
1700    End If

EXITP:
1710    KeyCode = intRetVal
1720    Exit Sub

ERRH:
1730    intRetVal = 0
1740    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1750    Case Else
1760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1770    End Select
1780    Resume EXITP

End Sub

Public Sub cmdSave_Click()
' ** This is a fake click event!

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

1810    DoCmd.RunCommand acCmdSelectRecord
1820    DoCmd.RunCommand acCmdSaveRecord

EXITP:
1830    Exit Sub

ERRH:
1840    Select Case ERR.Number
        Case Else
1850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1860    End Select
1870    Resume EXITP

End Sub

Private Sub Username_lbl_DblClick(Cancel As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Username_lbl_DblClick"

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

Private Sub Username_GotFocus()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Username_GotFocus"

2010    With Me
2020      .Username.SelLength = 0
2030      .Username.SelStart = 99
2040      If CLR_DISABLED_BG = 0 Then CLR_DISABLED_BG = CLR_LTTEAL
2050      If .secusr_defaultx = True Then
2060        .Username.BackColor = CLR_DISABLED_BG
2070      Else
2080        .Username.BackColor = CLR_WHT
2090      End If
2100    End With

EXITP:
2110    Exit Sub

ERRH:
2120    Select Case ERR.Number
        Case Else
2130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2140    End Select
2150    Resume EXITP

End Sub

Private Sub Username_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Username_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2220    intShiftDown = (Shift And acShiftMask) > 0
2230    intAltDown = (Shift And acAltMask) > 0
2240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2270        With Me
2280          intRetVal = 0
2290          .Employee_Name.SetFocus
2300        End With
2310      End Select
2320    End If

        ' ** Shift keys.
2330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2350        With Me
2360          intRetVal = 0
2370          If .CurrentRecord > 1 Then
2380            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2390            .Primary_Group.SetFocus
2400          Else
2410            DoCmd.SelectObject acForm, .Parent.Name, False
2420            .Parent.cmdClose.SetFocus
2430          End If
2440        End With
2450      End Select
2460    End If

        ' ** Prevent editing.
2470    If (Not intCtrlDown) And (Not intAltDown) Then  ' ** And (Not intShiftDown)
2480      If (intRetVal >= vbKey0 And intRetVal <= vbKey9) Or _
              (intRetVal >= vbKeyA And intRetVal <= vbKeyZ) Or (intRetVal = vbKeySpace) Then
2490        intRetVal = 0
2500        If blnUserNameEdit = False Then
2510          blnUserNameEdit = True
2520          MsgBox "User names cannot be edited." & vbCrLf & vbCrLf & _
                "To change a user name, delete the user and re-enter.", vbInformation + vbOKOnly, "Invalid Action"
2530        Else
2540          SysCmd acSysCmdSetStatus, "User names cannot be edited"
2550        End If
2560      End If
2570    End If

EXITP:
2580    KeyCode = intRetVal
2590    Exit Sub

ERRH:
2600    intRetVal = 0
2610    Select Case ERR.Number
        Case Else
2620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2630    End Select
2640    Resume EXITP

End Sub

Private Sub Username_LostFocus()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Username_LostFocus"

2710    SysCmd acSysCmdClearStatus

EXITP:
2720    Exit Sub

ERRH:
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Sub

Private Sub Employee_Name_lbl_DblClick(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Employee_Name_lbl_DblClick"

2810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub Employee_Name_GotFocus()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Employee_Name_GotFocus"

2910    With Me
2920      If CLR_DISABLED_BG = 0 Then CLR_DISABLED_BG = CLR_LTTEAL
2930      If .secusr_defaultx = True Then
2940        .Employee_Name.BackColor = CLR_DISABLED_BG
2950      Else
2960        .Employee_Name.BackColor = CLR_WHT
2970      End If
2980      .Employee_Name.SelLength = 0
2990      .Employee_Name.SelStart = 99
3000    End With

EXITP:
3010    Exit Sub

ERRH:
3020    Select Case ERR.Number
        Case Else
3030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3040    End Select
3050    Resume EXITP

End Sub

Private Sub Employee_Name_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Employee_Name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3120    intShiftDown = (Shift And acShiftMask) > 0
3130    intAltDown = (Shift And acAltMask) > 0
3140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3170        With Me
3180          intRetVal = 0
3190          If .Primary_Group.Enabled = True Then
3200            .Primary_Group.SetFocus
3210          Else
3220            lngRecsCur = RecCnt  ' ** Function: below.
3230            If .CurrentRecord < lngRecsCur Then
3240              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3250              .Username.SetFocus
3260            Else
3270              DoCmd.SelectObject acForm, .Parent.Name, False
3280              If .Parent.cmdDelete.Enabled = True Then
3290                .Parent.cmdDelete.SetFocus
3300              ElseIf .Parent.cmdAdd.Enabled = True Then
3310                .Parent.cmdAdd.SetFocus
3320              ElseIf .Parent.cmdChangePassword.Enabled = True Then
3330                .Parent.cmdChangePassword.SetFocus
3340              Else
3350                .Parent.cmdClose.SetFocus
3360              End If
3370            End If
3380          End If
3390        End With
3400      End Select
3410    End If

        ' ** Shift keys.
3420    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3430      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3440        With Me
3450          intRetVal = 0
3460          .Username.SetFocus
3470        End With
3480      End Select
3490    End If

EXITP:
3500    KeyCode = intRetVal
3510    Exit Sub

ERRH:
3520    Select Case ERR.Number
        Case Else
3530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3540    End Select
3550    Resume EXITP

End Sub

Private Sub Primary_Group_lbl_DblClick(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Primary_Group_lbl_DblClick"

3610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub Primary_Group_GotFocus()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Primary_Group_GotFocus"

3710    With Me
3720      .Primary_Group.SelLength = 0
3730      .Primary_Group.SelStart = 99
3740    End With

EXITP:
3750    Exit Sub

ERRH:
3760    Select Case ERR.Number
        Case Else
3770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3780    End Select
3790    Resume EXITP

End Sub

Private Sub Primary_Group_Enter()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Primary_Group_Enter"

        Dim strGroupName As String
        Dim intX As Integer

3810    With Me
3820      If blnAdmin = True Then
3830        Set usr = DBEngine.Workspaces(0).Users(.Username)
3840        strSecLev = .Primary_Group
3850  On Error Resume Next
3860        strGroupName = usr.Groups(strSecLev).Name
3870        If ERR <> 0 Then
3880  On Error GoTo ERRH
              ' ** Clean up user's group list.
3890          For intX = 0 To (usr.Groups.Count - 1)  ' ** First delete out non-"User" groups.
3900            If usr.Groups(intX).Name <> "Users" Then
3910              usr.Groups.Delete usr.Groups(intX).Name
3920            End If
3930          Next
              ' ** Now append correct group.
3940          usr.Groups.Append usr.CreateGroup(.Primary_Group.text)
3950        Else
3960  On Error GoTo ERRH
3970        End If
3980      End If
3990    End With

EXITP:
4000    Exit Sub

ERRH:
4010    Select Case ERR.Number
        Case 3265 ' ** Item not found in this collection.
4020      lngRecsCur = RecCnt  ' ** Function: Below.
4030      If lngRecsCur > 0& Then
4040        MsgBox "The TA and Access security files are out of sync." & vbCrLf & _
              "Please delete and re-add the user " & Me.Username & ".", _
              vbCritical + vbOKOnly + vbDefaultButton1, "Out of Sync Security Files"
4050      Else
            ' ** Shouldn't be here anyway! Move on.
4060      End If
4070    Case Else
4080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4090    End Select
4100    Resume EXITP

End Sub

Private Sub Primary_Group_KeyDown(KeyCode As Integer, Shift As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Primary_Group_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4220    intShiftDown = (Shift And acShiftMask) > 0
4230    intAltDown = (Shift And acAltMask) > 0
4240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4270        With Me
4280          intRetVal = 0
4290          lngRecsCur = RecCnt  ' ** Function: below.
4300          If .CurrentRecord < lngRecsCur Then
4310            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4320            .Username.SetFocus
4330          Else
4340            DoCmd.SelectObject acForm, .Parent.Name, False
4350            If .Parent.cmdDelete.Enabled = True Then
4360              .Parent.cmdDelete.SetFocus
4370            ElseIf .Parent.cmdAdd.Enabled = True Then
4380              .Parent.cmdAdd.SetFocus
4390            ElseIf .Parent.cmdChangePassword.Enabled = True Then
4400              .Parent.cmdChangePassword.SetFocus
4410            Else
4420              .Parent.cmdClose.SetFocus
4430            End If
4440          End If
4450        End With
4460      End Select
4470    End If

        ' ** Shift keys.
4480    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4490      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4500        With Me
4510          intRetVal = 0
4520          .Employee_Name.SetFocus
4530        End With
4540      End Select
4550    End If

EXITP:
4560    KeyCode = intRetVal
4570    Exit Sub

ERRH:
4580    Select Case ERR.Number
        Case Else
4590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4600    End Select
4610    Resume EXITP

End Sub

Private Sub Primary_Group_Exit(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "Primary_Group_Exit"

4710    With Me
4720      If blnAdmin = True Then
4730        If .Primary_Group <> strSecLev Then
4740          If UCase$(usr.Name) = "SUPERUSER" Then
4750            If .Primary_Group <> "Admins" Then
4760              .Primary_Group = "Admins"
4770              MsgBox "Cannot take 'Superuser' out of 'Admins' group.", vbExclamation + vbOKOnly, "Invalid Action"
4780            End If
4790          Else
4800            If UCase$(usr.Name) = UCase$(DBEngine.Workspaces(0).Username) Then
4810              .Primary_Group = strSecLev
4820              MsgBox "Cannot edit your own security level.", vbExclamation + vbOKOnly, "Invalid Action"
4830            Else
4840              usr.Groups.Delete strSecLev
4850              usr.Groups.Append usr.CreateGroup(.Primary_Group)
4860              strSecLev = .Primary_Group
4870            End If
4880          End If
4890        End If
4900      End If
4910    End With

EXITP:
4920    Exit Sub

ERRH:
4930    Select Case ERR.Number
        Case 91
          ' ** The user in TA.Users was not in the MDW file when the Group textbox
          ' ** was entered, therefore no usr was instantiated. Skip this message.
4940    Case 2427  ' ** You entered an expression that has no value.
          ' ** Shouldn't be here anyway! Move on.
4950    Case Else
4960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4970    End Select
4980    Resume EXITP

End Sub

Public Function RecCnt() As Long

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

5010    lngRetVal = 0&

5020    With Me
5030      Set rst = .RecordsetClone
5040      If rst.BOF = True And rst.EOF = True Then
            ' ** Shouldn't really be here if that's the case.
5050      Else
5060        rst.MoveLast
5070        lngRetVal = rst.RecordCount
5080      End If
5090      rst.Close
5100    End With

EXITP:
5110    Set rst = Nothing
5120    RecCnt = lngRetVal
5130    Exit Function

ERRH:
5140    lngRetVal = 0&
5150    Select Case ERR.Number
        Case 3420  ' ** Object invalid or no longer set.
          ' ** No records?
5160    Case Else
5170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5180    End Select
5190    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

5210    Select Case intWhere
        Case acCmdRecordsGoToFirst
5220      DoCmd.RunCommand acCmdRecordsGoToFirst
5230    Case acCmdRecordsGoToLast
5240      DoCmd.RunCommand acCmdRecordsGoToLast
5250    Case acCmdRecordsGoToPrevious
5260      DoCmd.RunCommand acCmdRecordsGoToPrevious
5270    Case acCmdRecordsGoToNext
5280      DoCmd.RunCommand acCmdRecordsGoToNext
5290    Case Else
5300      If IsMissing(varID) = False Then
5310        With Me
5320          Set rst = .RecordsetClone
5330          With rst
5340            .FindFirst "[secusr_id] = " & CStr(varID)
5350            If .NoMatch = False Then
5360              Me.Bookmark = .Bookmark
5370            End If
5380            .Close
5390          End With
5400        End With
5410      End If
5420    End Select

EXITP:
5430    Set rst = Nothing
5440    Exit Sub

ERRH:
5450    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
5460    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5470    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5480    Case Else
5490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5500    End Select
5510    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [Username]"

5610    With Me
5620      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5630        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5640      End If
5650      .Controls(strSortLbl).Visible = False
5660      .Controls(strSortLine).Visible = False
5670      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
5680      If strProc = "Form_Load" Then
5690        strCalled = "Username"
5700        strSortNow = strSortOrig
5710        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
5720        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5730        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5740        .Controls(strSortLbl).Left = lngSortLbl_Left
5750        .Controls(strSortLine).Left = lngSortLine_Left
5760        .Controls(strSortLine).Width = lngSortLine_Width
5770        .Controls(strSortLbl).Caption = strArwUp
5780        .Controls(strSortLbl).ForeColor = CLR_DKBLU
5790      Else
5800        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
5810        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
5820        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5830        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5840        .Controls(strSortLbl).Left = lngSortLbl_Left
5850        .Controls(strSortLine).Left = lngSortLine_Left
5860        .Controls(strSortLine).Width = lngSortLine_Width
5870        If strCalled = "Username" Then
5880          strSortAsc = strSortOrig
5890          If strSortNow = strSortAsc Then
5900            strSortNow = "[Username] DESC"
5910            .Controls(strSortLbl).Caption = strArwDn
5920            .Controls(strSortLbl).ForeColor = CLR_DKRED
5930          Else
5940            strSortNow = strSortAsc
5950            .Controls(strSortLbl).Caption = strArwUp
5960            .Controls(strSortLbl).ForeColor = CLR_DKBLU
5970          End If
5980        Else
5990          .Controls(strSortLbl).Caption = strArwUp
6000          .Controls(strSortLbl).ForeColor = CLR_DKBLU
6010          Select Case strCalled
              Case "Employee_Name"
6020            strSortAsc = "[Employee_Name]" & strStdAsc
6030          Case "Primary_Group"
6040            strSortAsc = "[Primary_Group]" & strStdAsc
6050          End Select
6060          If strSortNow = strSortAsc Then
6070            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
6080            Select Case intCnt
                Case 1
6090              strTmp01 = strSortAsc & " DESC"
6100            Case 2
6110              intPos01 = InStr(strSortAsc, ",")
6120              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6130              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
6140            Case 3
6150              intPos01 = InStr(strSortAsc, ",")
6160              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6170              strTmp02 = Mid(strSortAsc, intPos01)
6180              intPos01 = InStr(2, strTmp02, ",")
6190              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
6200              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
6210            End Select
6220            strSortNow = strTmp01
6230            .Controls(strSortLbl).Caption = strArwDn
6240            .Controls(strSortLbl).ForeColor = CLR_DKRED
6250          Else
6260            strSortNow = strSortAsc
6270          End If
6280        End If
6290      End If
6300      .Controls(strSortLbl).Visible = True
6310      .Controls(strSortLine).Visible = True
6320      .OrderBy = strSortNow
6330      .OrderByOn = True
6340    End With

        ' ** Potential Sorts:
        ' **   [Username]
        ' **   [Employee_Name], [Username]
        ' **   [Primary_Group], [Username]

EXITP:
6350    Exit Sub

ERRH:
6360    Select Case ERR.Number
        Case Else
6370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6380    End Select
6390    Resume EXITP

End Sub

Public Function SortNow_Get() As String

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

6410    strRetVal = strSortNow

EXITP:
6420    SortNow_Get = strRetVal
6430    Exit Function

ERRH:
6440    strRetVal = vbNullString
6450    Select Case ERR.Number
        Case Else
6460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6470    End Select
6480    Resume EXITP

End Function
