VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmReferences"

'VGC 11/23/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Refs:             R {frmReferences_Sub}
' **   Exit:             X {cmdClose}

' ** Array: arr_varRef().
Private lngRefs As Long, arr_varRef() As Variant
Private Const R_ELEMS As Integer = 5  ' ** Array's first-element UBound().
Private Const R_RNAM    As Integer = 0
Private Const R_DSC     As Integer = 1
Private Const R_PATH    As Integer = 2
Private Const R_ISBROKE As Integer = 3
Private Const R_BUILTIN As Integer = 4
Private Const R_GUID    As Integer = 5

Private blnJustClose As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me

120       strCallingForm = vbNullString
130       If IsNull(.OpenArgs) = False Then
140         If Trim(.OpenArgs) <> vbNullString Then
150           strCallingForm = Trim(.OpenArgs)
160         End If
170       End If

180       blnJustClose = False

190       blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

200       ShowRefs  ' ** Procedure: Below.

210     End With

220     If Cancel = -1 Then
          ' ** Just for the record.
230       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Allowed"
240     End If

EXITP:
250     Exit Sub

ERRH:
260     Select Case ERR.Number
        Case Else
270       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
280     End Select
290     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

300   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

310     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
320     intShiftDown = (Shift And acShiftMask) > 0
330     intAltDown = (Shift And acAltMask) > 0
340     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Form:       X {cmdCloseForm}

        ' ** Alt keys.
350     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
360       Select Case intRetVal
          Case vbKeyX
370         intRetVal = 0
380         cmdClose_Click  ' ** Procedure: Below.
390       End Select
400     End If

        ' ** Ctrl-Shift keys.
410     If intCtrlDown And (Not intAltDown) And intShiftDown Then
420       Select Case intRetVal
          Case vbKeyX
430         intRetVal = 0
440         blnJustClose = True
450         SetOption_Dev  ' ** Module Function: modUtilities.
460         DoCmd.Close acForm, THIS_NAME
470       End Select
480     End If

EXITP:
490     KeyCode = intRetVal
500     Exit Sub

ERRH:
510     intRetVal = 0
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

610     If blnJustClose = False Then
620       If strCallingForm <> vbNullString Then
630         Select Case strCallingForm
            Case "frmMenu_Title"
640           gstrReturningForm = THIS_NAME
650           gblnSetFocus = True
660           DoCmd.SelectObject acForm, strCallingForm, False
670         End Select
680       End If
690     End If

EXITP:
700     Exit Sub

ERRH:
710     Select Case ERR.Number
        Case Else
720       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
730     End Select
740     Resume EXITP

End Sub

Public Sub cmdClose_Click()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

810     DoCmd.Close acForm, THIS_NAME

EXITP:
820     Exit Sub

ERRH:
830     Select Case ERR.Number
        Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
850     End Select
860     Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

910     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
920     intShiftDown = (Shift And acShiftMask) > 0
930     intAltDown = (Shift And acAltMask) > 0
940     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
950     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
970         With Me
980           intRetVal = 0
990           .frmReferences_Sub.SetFocus
1000        End With
1010      End Select
1020    End If

        ' ** Shift keys.
1030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1050        With Me
1060          intRetVal = 0
1070          .frmReferences_Sub.SetFocus
1080        End With
1090      End Select
1100    End If

EXITP:
1110    KeyCode = intRetVal
1120    Exit Sub

ERRH:
1130    intRetVal = 0
1140    Select Case ERR.Number
        Case Else
1150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1160    End Select
1170    Resume EXITP

End Sub

Private Sub ShowRefs()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowRefs"

        Dim ref As Access.Reference
        Dim strPath As String, blnIsBroken As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim lngX As Long, lngE As Long

1210    With Me

1220      lngRefs = 0&
1230      ReDim arr_varRef(R_ELEMS, 0)
1240      For Each ref In Application.References
1250        lngRefs = lngRefs + 1&
1260        lngE = lngRefs - 1&
1270        ReDim Preserve arr_varRef(R_ELEMS, lngE)
1280        arr_varRef(R_RNAM, lngE) = ref.Name
1290        arr_varRef(R_DSC, lngE) = Application.VBE.ActiveVBProject.References(lngRefs).description
1300        blnIsBroken = ref.IsBroken
1310  On Error Resume Next
1320        strPath = ref.FullPath
1330        If ERR.Number <> 0 Then
1340          strPath = "Error: " & CStr(ERR.Number) & "  " & ERR.description
1350  On Error GoTo ERRH
1360          If blnIsBroken = False Then blnIsBroken = True
1370        Else
1380  On Error GoTo ERRH
1390        End If
1400        arr_varRef(R_PATH, lngE) = strPath
1410        arr_varRef(R_ISBROKE, lngE) = blnIsBroken
1420        arr_varRef(R_BUILTIN, lngE) = ref.BuiltIn
1430        arr_varRef(R_GUID, lngE) = ref.Guid
1440      Next ref
1450      For lngX = 0& To (lngRefs - 1&)
1460        intPos01 = InStr(arr_varRef(R_PATH, lngX), " ")
1470        If intPos01 > 0 Then
1480          Do While intPos01 > 0  ' ** Replace all spaces with hard spaces.
1490            arr_varRef(R_PATH, lngX) = Left(arr_varRef(R_PATH, lngX), (intPos01 - 1)) & Chr(160) & _
                  Mid(arr_varRef(R_PATH, lngX), (intPos01 + 1))
1500            intPos01 = InStr(arr_varRef(R_PATH, lngX), " ")
1510          Loop
1520        End If
            ' ** Special treatment.
1530        intPos01 = InStr(arr_varRef(R_PATH, lngX), "Microsoft" & Chr(160) & "Shared")
1540        If intPos01 > 0 Then
              ' ** Microsoft Shared.
1550          arr_varRef(R_PATH, lngX) = Left(arr_varRef(R_PATH, lngX), (intPos01 + Len("Microsoft Shared"))) & _
                Mid(arr_varRef(R_PATH, lngX), (intPos01 + Len("Microsoft Shared") + 1))
1560        End If
1570        intPos01 = InStr(arr_varRef(R_PATH, lngX), "Microsoft" & Chr(160) & "Office")
1580        If intPos01 > 0 Then
              ' ** Microsoft Office.
1590          arr_varRef(R_PATH, lngX) = Left(arr_varRef(R_PATH, lngX), (intPos01 + Len("Microsoft Office"))) & _
                Mid(arr_varRef(R_PATH, lngX), (intPos01 + Len("Microsoft Office") + 1))
1600        End If
1610        intPos01 = InStr(arr_varRef(R_PATH, lngX), "Common" & Chr(160) & "Files\System")
1620        If intPos01 > 0 Then
              ' ** Common Files\System.
1630          arr_varRef(R_PATH, lngX) = Left(arr_varRef(R_PATH, lngX), (intPos01 + Len("Common Files\System"))) & _
                Mid(arr_varRef(R_PATH, lngX), (intPos01 + Len("Common Files\System") + 1))
1640        End If
1650      Next
1660      strTmp01 = vbNullString
1670      For lngX = 0& To (lngRefs - 1&) ' ** Chr(160) is the hard-space.
1680        strTmp01 = strTmp01 & (CStr(lngX + 1&) & ".") & Chr(160) & Chr(160) & Chr(160) & _
              IIf((lngX + 1&) < 10&, Chr(160), vbNullString) & arr_varRef(R_RNAM, lngX) & vbCrLf
1690        strTmp01 = strTmp01 & String(7, Chr(160)) & arr_varRef(R_DSC, lngX) & vbCrLf
1700        strTmp01 = strTmp01 & String(7, Chr(160)) & "IsBroken: " & Format(arr_varRef(R_ISBROKE, lngX), "True/False") & vbCrLf
1710        strTmp01 = strTmp01 & String(7, Chr(160)) & arr_varRef(R_PATH, lngX) & vbCrLf & vbCrLf
1720      Next
1730      If Right(strTmp01, 4) = (vbCrLf & vbCrLf) Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 4))
1740      If Right(strTmp01, 2) = vbCrLf Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 2))

1750      .frmReferences_Sub.Form.Ref_display = strTmp01

1760    End With

EXITP:
1770    Set ref = Nothing
1780    Exit Sub

ERRH:
1790    Select Case ERR.Number
        Case Else
1800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1810    End Select
1820    Resume EXITP

End Sub
