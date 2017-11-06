VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCalendar"

'VGC 03/23/2017: CHANGES!

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar}
' **   Date Picker:      F8 {cmdCalendar}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Form:       X {cmdCloseForm}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of the Calendar class.
Public clsMonthClass As clsMonthCal

'Private Const opgDateType_optDateOnly As Integer = 1
'Private Const opgDateType_optDateHour As Integer = 2

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean

Private blnTwoSubs As Boolean, blnGoodDate As Boolean
Private datStartDate As Date, datEndDate As Date
Private strCallingSub1 As String, strCallingSub2 As String, strTag1 As String, strTag2 As String
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm <> vbNullString Then  'frmReinvest_Dividend, frmReinvest_Interest

160         blnCalendar1_Focus = False: blnCalendar1_MouseDown = False

            ' ** Copyright Lebans Holdings 1999 Ltd.
            ' ** Create an instance of the Calendar class.
170         Set clsMonthClass = New clsMonthCal
            ' ** You MUST SET the class hWndForm prop!!!
180         clsMonthClass.hWndForm = Me.hwnd
            ' ** Let's default to PositionAtCursor.
190         clsMonthClass.PositionAtCursor = True

200         .DateOnly = Null
210         .DateHour = Null

220       Else
230         Cancel = -1
240       End If
250     End With

260     If Cancel = -1 Then
270       Beep
280       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
290     End If

EXITP:
300     Exit Sub

ERRH:
310     Select Case ERR.Number
        Case Else
320       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
330     End Select
340     Resume EXITP

End Sub

Private Sub Form_Load()
' ** We want to make sure the ControlTipText matches,
' ** and which subform this is.

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim ctl As Access.Control
        Dim intPos01 As Integer
        Dim strTmp01 As String

410     With Me
420       If strCallingForm <> vbNullString Then
430         strCallingSub1 = vbNullString: strCallingSub2 = vbNullString: blnTwoSubs = False
440         For Each ctl In .Parent.Controls
450           With ctl
460             If .ControlType = acSubform Then
470               If .SourceObject = THIS_NAME Then
480                 If strCallingSub1 = vbNullString Then
490                   strCallingSub1 = .Name
500                 Else
510                   blnTwoSubs = True
520                   strCallingSub2 = .Name
530                   Exit For
540                 End If
550               End If
560             End If
570           End With
580         Next
            ' ** Calendar Subform; assetdate; F8
590         strTag1 = .Parent.Controls(strCallingSub1).Tag
600         Select Case blnTwoSubs
            Case True
              ' ** Just collect the tags if we don't yet know which one this is.
610           strTag2 = .Parent.Controls(strCallingSub2).Tag
620         Case False
630           intPos01 = InStr(strTag1, ";")
640           If intPos01 > 0 Then
650             strTmp01 = Trim(Mid(strTag1, (intPos01 + 1)))
660             intPos01 = InStr(strTmp01, ";")
670             If intPos01 > 0 Then
680               .DateDest = Trim(Left(strTmp01, (intPos01 - 1)))
690               strTmp01 = Trim(Mid(strTmp01, (intPos01 + 1)))
700               If Right(.cmdCalendar.ControlTipText, 2) <> strTmp01 Then
710                 .cmdCalendar.ControlTipText = "Date Picker - " & strTmp01
720               End If
730             End If
740           End If
750         End Select
760       End If
770     End With

EXITP:
780     Set ctl = Nothing
790     Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

900   On Error GoTo ERRH

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        Const THIS_PROC As String = "Form_KeyDown"

910     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
920     intShiftDown = (Shift And acShiftMask) > 0
930     intAltDown = (Shift And acAltMask) > 0
940     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar}
        ' **   Date Picker:      F8 {cmdCalendar}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Form:       X {cmdCloseForm}

        ' ** Plain keys.
950     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyF7
970         With Me
980           intRetVal = 0
990           If InStr(.cmdCalendar.ControlTipText, "F7") > 0 Then
1000            cmdCalendar_Click  ' ** Procedure: Below.
1010          End If
1020        End With
1030      Case vbKeyF8
1040        With Me
1050          intRetVal = 0
1060          If InStr(.cmdCalendar.ControlTipText, "F8") > 0 Then
1070            cmdCalendar_Click  ' ** Procedure: Below.
1080          End If
1090        End With
1100      End Select
1110    End If

        ' ** Ctrl-Shift keys.
1120    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1130      Select Case intRetVal
          Case vbKeyX
1140        If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
1150          intRetVal = 0
1160          If strCallingForm = vbNullString Then
1170            DoCmd.Close acForm, THIS_NAME
1180          End If
1190        End If
1200      End Select
1210    End If

EXITP:
1220    KeyCode = intRetVal
1230    Exit Sub

ERRH:
1240    intRetVal = 0
1250    Select Case ERR.Number
        Case Else
1260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1270    End Select
1280    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1310    If Not clsMonthClass Is Nothing Then
1320      If clsMonthClass.IsCalendar Then
1330        Cancel = -1
1340      Else
1350        Set clsMonthClass = Nothing
1360      End If
1370    End If

EXITP:
1380    Exit Sub

ERRH:
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub

Public Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1510    With Me
1520      If .cmdCalendar_raised_focus_dots_img.Visible = True Or .cmdCalendar_raised_focus_img.Visible = True Then
1530        Select Case blnCalendar1_Focus
            Case True
1540          .cmdCalendar_raised_semifocus_dots_img.Visible = True
1550          .cmdCalendar_raised_img.Visible = False
1560        Case False
1570          .cmdCalendar_raised_img.Visible = True
1580          .cmdCalendar_raised_semifocus_dots_img.Visible = False
1590        End Select
1600        .cmdCalendar_raised_focus_dots_img.Visible = False
1610        .cmdCalendar_raised_focus_img.Visible = False
1620        .cmdCalendar_sunken_focus_dots_img.Visible = False
1630        .cmdCalendar_raised_img_dis.Visible = False
1640      End If
1650    End With

EXITP:
1660    Exit Sub

ERRH:
1670    Select Case ERR.Number
        Case Else
1680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1690    End Select
1700    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

1800  On Error GoTo ERRH

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        Const THIS_PROC As String = "FocusHolder_KeyDown"

1810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1820    intShiftDown = (Shift And acShiftMask) > 0
1830    intAltDown = (Shift And acAltMask) > 0
1840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1870        With Me
1880          intRetVal = 0
1890          If strCallingForm = vbNullString Then
1900            .cmdCalendar.SetFocus
1910          End If
1920        End With
1930      End Select
1940    End If

EXITP:
1950    KeyCode = intRetVal
1960    Exit Sub

ERRH:
1970    intRetVal = 0
1980    Select Case ERR.Number
        Case Else
1990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2000    End Select
2010    Resume EXITP

End Sub

Private Sub cmdCalendar_Enter()
' ** We want to make sure the ControlTipText matches,
' ** and which subform this is.

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Enter"

        Dim intPos01 As Integer
        Dim strTmp01 As String

2110    With Me
2120      blnGoodDate = False
2130      If blnTwoSubs = True Then
2140        strTmp01 = vbNullString
2150        If IsNull(.DateDest) = False Then
2160          If InStr(strTag1, .DateDest) > 0 Then
2170            strCallingSub2 = vbNullString: strTag2 = vbNullString
2180          ElseIf InStr(strTag2, .DateDest) > 0 Then
2190            strCallingSub1 = strCallingSub2: strTag1 = strTag2
2200            strCallingSub2 = vbNullString: strTag2 = vbNullString
2210          End If
2220          intPos01 = InStr(strTag1, ";")
2230          If intPos01 > 0 Then
2240            strTmp01 = Trim(Mid(strTag1, (intPos01 + 1)))
2250            intPos01 = InStr(strTmp01, ";")
2260            If intPos01 > 0 Then
2270              strTmp01 = Trim(Mid(strTmp01, (intPos01 + 1)))
2280              If Right(.cmdCalendar.ControlTipText, 2) <> strTmp01 Then
2290                .cmdCalendar.ControlTipText = "Date Picker - " & strTmp01
2300              End If
2310            End If
2320          End If
2330        End If
2340      End If
2350    End With

EXITP:
2360    Exit Sub

ERRH:
2370    Select Case ERR.Number
        Case Else
2380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2390    End Select
2400    Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

2510    With Me
2520      blnCalendar1_Focus = True
2530      .cmdCalendar_raised_semifocus_dots_img.Visible = True
2540      .cmdCalendar_raised_img.Visible = False
2550      .cmdCalendar_raised_focus_img.Visible = False
2560      .cmdCalendar_raised_focus_dots_img.Visible = False
2570      .cmdCalendar_sunken_focus_dots_img.Visible = False
2580      .cmdCalendar_raised_img_dis.Visible = False
2590    End With

EXITP:
2600    Exit Sub

ERRH:
2610    Select Case ERR.Number
        Case Else
2620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2630    End Select
2640    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

2710    With Me
2720      blnCalendar1_MouseDown = True
2730      .cmdCalendar_sunken_focus_dots_img.Visible = True
2740      .cmdCalendar_raised_img.Visible = False
2750      .cmdCalendar_raised_semifocus_dots_img.Visible = False
2760      .cmdCalendar_raised_focus_img.Visible = False
2770      .cmdCalendar_raised_focus_dots_img.Visible = False
2780      .cmdCalendar_raised_img_dis.Visible = False
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

Public Sub cmdCalendar_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

2910    With Me
2920      datStartDate = Date
2930      datEndDate = 0
2940      blnGoodDate = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
2950      Select Case blnGoodDate
          Case True
2960        .DateOnly = CDate(Format(datStartDate, "mm/dd/yyyy"))
2970        .DateHour = CDate(Format(datStartDate + time(), "mm/dd/yyyy hh:nn:ss"))
2980      Case False
2990        .DateOnly = CDate(Format(Date, "mm/dd/yyyy"))
3000        .DateHour = CDate(Format(Now(), "mm/dd/yyyy hh:nn:ss"))
3010      End Select
3020      ParentSet  ' ** Procedure: Below.
3030    End With

EXITP:
3040    Exit Sub

ERRH:
3050    Select Case ERR.Number
        Case Else
3060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3070    End Select
3080    Resume EXITP

End Sub

Private Sub cmdCalendar_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        Const THIS_PROC As String = "cmdCalendar_KeyDown"

3110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3120    intShiftDown = (Shift And acShiftMask) > 0
3130    intAltDown = (Shift And acAltMask) > 0
3140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3160      Select Case intRetVal
          Case vbKeyTab
3170        With Me
3180          intRetVal = 0
3190          If strCallingForm = vbNullString Then
3200            ParentMove  ' ** Procedure: Below.
3210          End If
3220        End With
3230      End Select
3240    End If

EXITP:
3250    KeyCode = intRetVal
3260    Exit Sub

ERRH:
3270    intRetVal = 0
3280    Select Case ERR.Number
        Case Else
3290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3300    End Select
3310    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseMove"

3410    With Me
3420      If blnCalendar1_MouseDown = False Then
3430        Select Case blnCalendar1_Focus
            Case True
3440          .cmdCalendar_raised_focus_dots_img.Visible = True
3450          .cmdCalendar_raised_focus_img.Visible = False
3460        Case False
3470          .cmdCalendar_raised_focus_img.Visible = True
3480          .cmdCalendar_raised_focus_dots_img.Visible = False
3490        End Select
3500        .cmdCalendar_raised_img.Visible = False
3510        .cmdCalendar_raised_semifocus_dots_img.Visible = False
3520        .cmdCalendar_sunken_focus_dots_img.Visible = False
3530        .cmdCalendar_raised_img_dis.Visible = False
3540      End If
3550    End With

EXITP:
3560    Exit Sub

ERRH:
3570    Select Case ERR.Number
        Case Else
3580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3590    End Select
3600    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseUp"

3710    With Me
3720      .cmdCalendar_raised_focus_dots_img.Visible = True
3730      .cmdCalendar_raised_img.Visible = False
3740      .cmdCalendar_raised_semifocus_dots_img.Visible = False
3750      .cmdCalendar_raised_focus_img.Visible = False
3760      .cmdCalendar_sunken_focus_dots_img.Visible = False
3770      .cmdCalendar_raised_img_dis.Visible = False
3780      blnCalendar1_MouseDown = False
3790    End With

EXITP:
3800    Exit Sub

ERRH:
3810    Select Case ERR.Number
        Case Else
3820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3830    End Select
3840    Resume EXITP

End Sub

Private Sub cmdCalendar_LostFocus()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_LostFocus"

3910    With Me
3920      .cmdCalendar_raised_img.Visible = True
3930      .cmdCalendar_raised_semifocus_dots_img.Visible = False
3940      .cmdCalendar_raised_focus_img.Visible = False
3950      .cmdCalendar_raised_focus_dots_img.Visible = False
3960      .cmdCalendar_sunken_focus_dots_img.Visible = False
3970      .cmdCalendar_raised_img_dis.Visible = False
3980      blnCalendar1_Focus = False
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

Public Sub ParentSet()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "ParentSet"

        Dim Button As Integer, Shift As Integer, X As Single, Y As Single

4110    With Me
4120      If strCallingForm <> vbNullString Then
4130        Select Case strCallingSub1
            Case "cmdCalendar"
4140          .Parent.CheckCalendar blnGoodDate, datStartDate  ' ** Form Procedure: Parent.
4150        Case "cmdCalendar1"
4160          .Parent.CheckCalendar1 blnGoodDate, datStartDate  ' ** Form Procedure: Parent.
4170        Case "cmdCalendar2"
4180          .Parent.CheckCalendar2 blnGoodDate, datStartDate  ' ** Form Procedure: Parent.
4190        End Select
4200        .FocusHolder.SetFocus
4210        Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.
4220      End If
4230    End With

EXITP:
4240    Exit Sub

ERRH:
4250    Select Case ERR.Number
        Case Else
4260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4270    End Select
4280    Resume EXITP

End Sub

Public Sub ParentMove()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "ParentMove"

        Dim ctl As Access.Control
        Dim lngCtls As Long, arr_varCtl() As Variant
        Dim lngDestSec As Long, lngDestTab As Long
        Dim strDestCtl As String, strNextCtl As String
        Dim varTmp00 As Variant, lngTmp01 As Long
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long

        ' ** Array: arr_varCtl().
        Const C_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const C_NAM As Integer = 0
        Const C_TYP As Integer = 1
        Const C_TAB As Integer = 2

4310    With Me
4320      If strCallingForm <> vbNullString Then
4330        If IsNull(.DateDest) = False Then

4340          strDestCtl = .DateDest
4350          lngDestSec = .Parent.Controls(strDestCtl).Section

4360          lngCtls = 0&
4370          ReDim arr_varCtl(C_ELEMS, 0)

              ' ** Collect the parent controls.
4380          For Each ctl In .Parent.Section(lngDestSec).Controls
4390            With ctl
4400              Select Case .ControlType
                  Case acLabel, acLine, acRectangle, acImage, acPageBreak
                    ' ** Skip these.
4410              Case Else
4420                lngTmp01 = .TabIndex
4430                If .Visible = True And .Enabled = True And .TabStop = True Then
4440                  lngCtls = lngCtls + 1&
4450                  lngE = lngCtls - 1&
4460                  ReDim Preserve arr_varCtl(C_ELEMS, lngE)
4470                  arr_varCtl(C_NAM, lngE) = .Name
4480                  arr_varCtl(C_TYP, lngE) = .ControlType
4490                  arr_varCtl(C_TAB, lngE) = lngTmp01
4500                End If
4510              End Select
4520            End With
4530          Next

              ' ** Binary Sort arr_varCtl() array by tab order.
4540          For lngX = UBound(arr_varCtl, 2) To 1 Step -1
4550            For lngY = 0 To (lngX - 1)
4560              If arr_varCtl(C_TAB, lngY) > arr_varCtl(C_TAB, (lngY + 1)) Then
4570                For lngZ = 0& To C_ELEMS
4580                  varTmp00 = arr_varCtl(lngZ, lngY)
4590                  arr_varCtl(lngZ, lngY) = arr_varCtl(lngZ, (lngY + 1))
4600                  arr_varCtl(lngZ, (lngY + 1)) = varTmp00
4610                  varTmp00 = Empty
4620                Next  ' ** lngZ.
4630              End If
4640            Next  ' ** lngY.
4650          Next  ' ** lngX.

4660          lngDestTab = -1&: strNextCtl = vbNullString
4670          For lngX = 0& To (lngCtls - 1&)
4680            If arr_varCtl(C_NAM, lngX) = strDestCtl Then
4690              lngDestTab = arr_varCtl(C_TAB, lngX)
4700            ElseIf lngDestTab <> -1& Then
4710              If arr_varCtl(C_TYP, lngX) = acSubform Then
4720                If .Parent.Controls(arr_varCtl(C_NAM, lngX)).SourceObject = THIS_NAME Then
                      ' ** Let's try not to just come back here!
4730                Else
4740                  strNextCtl = arr_varCtl(C_NAM, lngX)
4750                  Exit For
4760                End If
4770              Else
4780                strNextCtl = arr_varCtl(C_NAM, lngX)
4790                Exit For
4800              End If
4810            End If
4820          Next

4830          .FocusHolder.SetFocus
4840          DoCmd.SelectObject acForm, .Parent.Name, False
4850          .Parent.Controls(strNextCtl).SetFocus
4860          DoEvents

4870        End If
4880      End If
4890    End With

EXITP:
4900    Set ctl = Nothing
4910    Exit Sub

ERRH:
4920    Select Case ERR.Number
        Case Else
4930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4940    End Select
4950    Resume EXITP

End Sub
