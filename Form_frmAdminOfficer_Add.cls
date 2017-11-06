VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAdminOfficer_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAdminOfficer_Add"

'VGC 04/16/2011: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   OK:               O {cmdOK}
' **   Cancel:           C {cmdCancel}
' **   Officer:          F {AdminName}

Private lngNextAdminNo As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim varTmp00 As Variant

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           varTmp00 = DMax("[adminno]", "adminofficer")
160           If IsNull(varTmp00) = False Then

170             lngNextAdminNo = (CLng(varTmp00) + 1&)
180             .adminno = lngNextAdminNo

190             .AdminName.SetFocus

200           Else
                ' ** There must me a #1 for 'Unassigned'!
210             Cancel = -1
220           End If
230         Else
240           Cancel = -1
250         End If
260       Else
270         Cancel = -1
280       End If
290     End With

300     If Cancel = -1 Then
310       Beep
320       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
330     End If

EXITP:
340     Exit Sub

ERRH:
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

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
560     Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
570     Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Sub

Private Sub cmdOK_Click()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strOfficer As String
        Dim blnContinue As Boolean

710     blnContinue = True

720     With Me
730       If IsNull(.AdminName) = False Then
740         If Trim(.AdminName) <> vbNullString Then
750           .AdminName.SetFocus
760           If InStr(.AdminName.text, Chr(34)) > 0 Then
770             blnContinue = False
780             MsgBox "Cannot have quote marks in Name.", vbInformation + vbOKOnly, "Invalid Characters"
790             DoCmd.CancelEvent
800           Else
810             strOfficer = Trim(.AdminName)
820             Set dbs = CurrentDb
830             With dbs
840               Set rst = .OpenRecordset("adminofficer", dbOpenDynaset, dbConsistent)
850               With rst
860                 .FindFirst "[officer] = '" & strOfficer & "'"
870                 If .NoMatch = True Then
880                   .AddNew
890                   ![officer] = strOfficer
900                   .Update
910                   .Bookmark = .LastModified
920                   lngNextAdminNo = ![adminno]
930                   If Forms("frmAdminOfficer").chkShowAccounts = True Then
940                     Forms("frmAdminOfficer").chkShowAccounts = False
950                     Forms("frmAdminOfficer").chkShowAccounts_AfterUpdate  ' ** Form Procedure: frmAdminOfficer.
960                   End If
970                   Forms("frmAdminOfficer").frmAdminOfficer_Sub.Form.Refresh
980                   Forms("frmAdminOfficer").frmAdminOfficer_Sub.Form.Requery
990                 Else
1000                  blnContinue = False
1010                  MsgBox "The Officer, '" & strOfficer & "', is already present.", vbInformation + vbOKOnly, "Duplicate Officer Name"
1020                  DoCmd.CancelEvent
1030                End If
1040              End With
1050              .Close
1060            End With
1070          End If
1080        Else
1090          blnContinue = False
1100          MsgBox "Officer name must be filled to continue.", vbInformation + vbOKOnly, "Entry Required"
1110          DoCmd.CancelEvent
1120        End If
1130      Else
1140        blnContinue = False
1150        MsgBox "Officer name must be filled to continue.", vbInformation + vbOKOnly, "Entry Required"
1160        DoCmd.CancelEvent
1170      End If

1180    End With

1190    If blnContinue = True Then
1200      DoCmd.Close acForm, THIS_NAME
1210    End If

EXITP:
1220    Set rst = Nothing
1230    Set dbs = Nothing
1240    Exit Sub

ERRH:
1250    Select Case ERR.Number
        Case Else
1260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1270    End Select
1280    Resume EXITP

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
1470          .AdminName.SetFocus
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
1790          .AdminName.SetFocus
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

Private Sub AdminName_KeyDown(KeyCode As Integer, Shift As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "AdminName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2020    intShiftDown = (Shift And acShiftMask) > 0
2030    intAltDown = (Shift And acAltMask) > 0
2040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2070        With Me
2080          intRetVal = 0
2090          .cmdOK.SetFocus
2100        End With
2110      End Select
2120    End If

        ' ** Shift keys.
2130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2150        With Me
2160          intRetVal = 0
2170          .cmdCancel.SetFocus
2180        End With
2190      End Select
2200    End If

EXITP:
2210    KeyCode = intRetVal
2220    Exit Sub

ERRH:
2230    intRetVal = 0
2240    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
2250    Case Else
2260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2270    End Select
2280    Resume EXITP

End Sub
