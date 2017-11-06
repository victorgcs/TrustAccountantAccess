VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmVersion_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmVersion_Input"

'VGC 10/27/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Acct w/Type:      A {chkAccountNoWithType}
' **   Cancel:           C {cmdCancel}
' **   Dev Data:         D {cmdDevData}
' **   Cap Gain/Loss:    G {chkSpecialCapGainLoss}
' **   Inc/Exp Track:    I {chkRevenueExpenseTracking}
' **   Tab Key Copy:     K {chkTabCopy}
' **   Link Tax-Inc/Exp: L {chkLinkRevTaxCodes}
' **   Name:             M {CoInfo_Name}
' **   Next:             N {cmdNext}
' **   Other:            O {opgLoc_optOther}
' **   Phone Number:     P {CoInfo_Phone}
' **   Separate Check:   S {chkSeparateCheckingAccounts}
' **   Tax Track:        T {chkIncomeTaxCoding}
' **   USA:              U {opgLoc_optUSA}

' ** AcBorderStyle enumeration (my own):
' **   0  acBorderStyleTransparent  Transparent   Transparent (Default only for label, chart, and subreport)
' **   1  acBorderStyleSolid        Solid         Solid line (Default)
' **   2  acBorderStyleDash         Dashes        Dashed line
' **   3  acBorderStyleShortDash    Short dashes  Dashed line with short dashes
' **   4  acBorderStyleDot          Dots          Dotted line
' **   5  acBorderStyleSparseDot    Sparse dots   Dotted line with dots spaced far apart
' **   6  acBorderStyleDashDot      Dash dot      Line with a dash-dot combination
' **   7  acBorderStyleDashDotDot   Dash dot dot  Line with a dash-dot-dot combination
' **   8  acBorderStyleDoubleSolid  Double solid  Double solid lines

Private blnIsOpen As Boolean
Private strOpt As String
Private intTmp_SpecialCapGainLossOpt As Integer
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strOpt = Trim(.OpenArgs)
150           blnIsOpen = True

160           StateCodeQrySet Me  ' ** Module Function: modPreferenceFuncs.

170           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

180           If IsNull(.CoInfo_State_Pref) = False Then
190             .CoInfo_State.DefaultValue = Chr(34) & .CoInfo_State_Pref & Chr(34)
200           End If

210           If GetUserName = gstrDevUserName Then    ' ** Module Function: modFileUtilities.
220             .cmdDevData.Enabled = True
230             .cmdDevData.Visible = True
240           End If

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

EXITP:
360     Exit Sub

ERRH:
370     Select Case ERR.Number
        Case Else
380       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
390     End Select
400     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim ctl As Access.Control
        Dim arr_varRetVal As Variant

510     With Me

520       If strOpt = vbNullString Then
530         If IsNull(.OpenArgs) = False Then
540           If Trim(.OpenArgs) <> vbNullString Then
550             strOpt = Trim(.OpenArgs)
560           End If
570         End If
580       End If

590       If strOpt <> vbNullString Then
600         arr_varRetVal = Version_DataXFer("Get", "CoInfo")  ' ** Module Function: modVersionConvertFuncs2.
610         If Left(arr_varRetVal(0, 0), 1) <> "#" Then  ' ** Covers both #ERROR and #EMPTY.
620           .CoInfo_ID = CLng(1)
630           .CoInfo_Name = arr_varRetVal(0, 0)
640           .CoInfo_Address1 = arr_varRetVal(1, 0)
650           .CoInfo_Address2 = arr_varRetVal(2, 0)
660           .CoInfo_City = arr_varRetVal(3, 0)
670           .CoInfo_State = arr_varRetVal(4, 0)
680           .CoInfo_Zip = arr_varRetVal(5, 0)
690           Select Case IsNull(arr_varRetVal(15, 0))
              Case True
700             .opgLoc = .opgLoc_optUSA.OptionValue
710             .CoInfo_Country = Null
720           Case False
730             If Trim(arr_varRetVal(15, 0)) = vbNullString Then
740               .opgLoc = .opgLoc_optUSA.OptionValue
750               .CoInfo_Country = Null
760             Else
770               .opgLoc = .opgLoc_optOther.OptionValue
780               .CoInfo_Country = arr_varRetVal(15, 0)
790             End If
800           End Select
810           .CoInfo_PostalCode = arr_varRetVal(16, 0)
820           .CoInfo_Phone = arr_varRetVal(6, 0)
830           .chkIncomeTaxCoding = False 'arr_varRetVal(7, 0)
840           .chkRevenueExpenseTracking = False 'arr_varRetVal(8, 0)
850           .chkAccountNoWithType = False 'arr_varRetVal(9, 0)
860           .chkSeparateCheckingAccounts = False 'arr_varRetVal(10, 0)
870           .chkTabCopy = False 'arr_varRetVal(11, 0)
880           .chkLinkRevTaxCodes = False 'arr_varRetVal(12, 0)
890           .chkSpecialCapGainLoss = False 'arr_varRetVal(13, 0)
900           intTmp_SpecialCapGainLossOpt = arr_varRetVal(14, 0)  ' ** We won't be setting this here.
910           .Username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
920           .CoInfo_DateModified = Now()
930           For Each ctl In .Controls
940             With ctl
950               Select Case .ControlType
                  Case acTextBox, acComboBox, acCheckBox, acCommandButton
960                 .Enabled = False
970                 If .ControlType = acTextBox Or .ControlType = acComboBox Then
980                   .BorderColor = WIN_CLR_DISR
990                   .BackStyle = acBackStyleTransparent
1000                End If
1010              Case Else
                    ' ** Doesn't matter.
1020              End Select
1030            End With
1040          Next
1050          .opgLoc.Enabled = False
1060          .CoInfo_lbl.ForeColor = WIN_CLR_DISF
1070          .CoInfo_lbl_dim_hi.Visible = True
1080          .CoInfo_Name_lbl2.ForeColor = WIN_CLR_DISF
1090          .CoInfo_Name_lbl2_dim_hi.Visible = True
1100          .CoInfo_Address1_lbl2.ForeColor = WIN_CLR_DISF
1110          .CoInfo_Address1_lbl2_dim_hi.Visible = True
1120          .CoInfo_City_lbl2.ForeColor = WIN_CLR_DISF
1130          .CoInfo_City_lbl2_dim_hi.Visible = True
1140          .CoInfo_State_lbl2.ForeColor = WIN_CLR_DISF
1150          .CoInfo_State_lbl2_dim_hi.Visible = True
1160          .CoInfo_Zip_lbl2.ForeColor = WIN_CLR_DISF
1170          .CoInfo_Zip_lbl2_dim_hi.Visible = True
1180          .CoInfo_Country_lbl2.ForeColor = WIN_CLR_DISF
1190          .CoInfo_Country_lbl2_dim_hi.Visible = True
1200          .CoInfo_ReqMsg_lbl1.ForeColor = WIN_CLR_DISF
1210          .CoInfo_ReqMsg_lbl1_dim_hi.Visible = True
1220          .CoInfo_ReqMsg_lbl2.ForeColor = WIN_CLR_DISF
1230          .CoInfo_ReqMsg_lbl2_dim_hi.Visible = True
1240          .Options_lbl.ForeColor = WIN_CLR_DISF
1250          .Options_lbl_dim_hi.Visible = True
1260          .TimerInterval = 100&
1270        Else
1280          .cmdCancel.SetFocus
1290        End If
1300      Else
1310        .opgLoc = .opgLoc_optUSA.OptionValue
1320      End If
1330      opgLoc_AfterUpdate  ' ** Procedure: Below.

1340      If .chkIncomeTaxCoding = False Or .chkRevenueExpenseTracking = False Then
1350        .chkLinkRevTaxCodes = False
1360        .chkLinkRevTaxCodes.Enabled = False
1370      End If

1380      chkIncomeTaxCoding_AfterUpdate  ' ** Procedure: Below.
1390      chkRevenueExpenseTracking_AfterUpdate  ' ** Procedure: Below.
1400      chkAccountNoWithType_AfterUpdate  ' ** Procedure: Below.
1410      chkSeparateCheckingAccounts_AfterUpdate  ' ** Procedure: Below.
1420      chkTabCopy_AfterUpdate  ' ** Procedure: Below.
1430      chkLinkRevTaxCodes_AfterUpdate  ' ** Procedure: Below.
1440      chkSpecialCapGainLoss_AfterUpdate  ' ** Procedure: Below.

1450    End With

EXITP:
1460    Set ctl = Nothing
1470    Exit Sub

ERRH:
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1620    intShiftDown = (Shift And acShiftMask) > 0
1630    intAltDown = (Shift And acAltMask) > 0
1640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Other:            O {opgLoc_optOther}
        ' **   USA:              U {opgLoc_optUSA}

        ' ** Alt keys.
1650    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1660      Select Case intRetVal
          Case vbKeyO
1670        With Me
1680          intRetVal = 0
1690          If .opgLoc.Enabled = True Then
1700            .opgLoc = .opgLoc_optOther.OptionValue
1710            .opgLoc.SetFocus
1720            opgLoc_AfterUpdate  ' ** Procedure: Below.
1730          Else
1740            Beep
1750          End If
1760        End With
1770      Case vbKeyU
1780        With Me
1790          intRetVal = 0
1800          If .opgLoc.Enabled = True Then
1810            .opgLoc = .opgLoc_optUSA.OptionValue
1820            .opgLoc.SetFocus
1830            opgLoc_AfterUpdate  ' ** Procedure: Below.
1840          Else
1850            Beep
1860          End If
1870        End With
1880      End Select
1890    End If

EXITP:
1900    KeyCode = intRetVal
1910    Exit Sub

ERRH:
1920    intRetVal = 0
1930    Select Case ERR.Number
        Case Else
1940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1950    End Select
1960    Resume EXITP

End Sub

Private Sub Form_Timer()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

2010    With Me
2020      .TimerInterval = 0&
2030      Beep
2040      DoCmd.OpenForm "frmVersion_Input_Notice", , , , , acDialog, THIS_NAME
2050      ResetDisabled  ' ** Procedure: Below.
2060    End With

EXITP:
2070    Exit Sub

ERRH:
2080    Select Case ERR.Number
        Case Else
2090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2100    End Select
2110    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2210    If DataErr <> 0 Then
2220      Select Case DataErr
          Case 2046, 2501, 3101, 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2237
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
2230      Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
2240        Select Case Me.ActiveControl.Name
            Case "CoInfo_State"
2250          MsgBox "Must enter a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
2260        Case "CoInfo_Zip"
2270          MsgBox "Must enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
2280        Case "CoInfo_Phone"
2290          MsgBox "Must enter a valid phone number.", vbInformation + vbOKOnly, "Invalid Entry"
2300        End Select
2310        Response = acDataErrContinue
2320      Case 3162
2330        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
2340        DoCmd.CancelEvent
2350      Case Else
2360        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2370        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
2380      End Select
2390    End If

EXITP:
2400    Exit Sub

ERRH:
2410    Select Case ERR.Number
        Case Else
2420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2430    End Select
2440    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2510    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim arr_varRetVal() As Variant
        Dim blnRetVal As Boolean

2610    ReDim arr_varRetVal(0, 0)
2620    arr_varRetVal(0, 0) = "#CANCEL"
2630    blnRetVal = Version_DataXFer("Ret", "CoInfo", arr_varRetVal)  ' ** Module Function: modVersionConvertFuncs2.
2640    DoCmd.Close acForm, THIS_NAME

EXITP:
2650    Exit Sub

ERRH:
2660    Select Case ERR.Number
        Case Else
2670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2680    End Select
2690    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
2790          .CoInfo_Name.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab
2850        With Me
2860          intRetVal = 0
2870          .cmdNext.SetFocus
2880        End With
2890      End Select
2900    End If

EXITP:
2910    KeyCode = intRetVal
2920    Exit Sub

ERRH:
2930    intRetVal = 0
2940    Select Case ERR.Number
        Case Else
2950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2960    End Select
2970    Resume EXITP

End Sub

Private Sub cmdDevData_Click()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDevData_Click"

3010    Version_Input1_EnterData  ' ** Module Function: modVersionConvertFuncs3.

EXITP:
3020    Exit Sub

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub cmdNext_Click()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdNext_Click"

        Dim arr_varRetVal() As Variant
        Dim blnRetVal As Boolean

3110    blnRetVal = True

3120    With Me
3130      If Nz(.CoInfo_Name, vbNullString) = vbNullString Then
3140        blnRetVal = False
3150        MsgBox "Please enter a Name to continue.", vbInformation + vbOKOnly, "Entry Required"
3160        .CoInfo_Name.SetFocus
3170      Else
3180        If Nz(.CoInfo_Address1, vbNullString) = vbNullString Then
3190          blnRetVal = False
3200          MsgBox "Please enter an Address 1 to continue.", vbInformation + vbOKOnly, "Entry Required"
3210          .CoInfo_Address1.SetFocus
3220        Else
3230          If Nz(.CoInfo_City, vbNullString) = vbNullString Then
3240            blnRetVal = False
3250            MsgBox "Please enter a City to continue.", vbInformation + vbOKOnly, "Entry Required"
3260            .CoInfo_City.SetFocus
3270          Else
3280            If IsNull(.CoInfo_Country) = True Then
3290              If Nz(.CoInfo_State, vbNullString) = vbNullString Then
3300                blnRetVal = False
3310                MsgBox "Please enter a State to continue.", vbInformation + vbOKOnly, "Entry Required"
3320                .CoInfo_State.SetFocus
3330              Else
3340                If Nz(.CoInfo_Zip, vbNullString) = vbNullString Then
3350                  blnRetVal = False
3360                  MsgBox "Please enter a Zip Code to continue.", vbInformation + vbOKOnly, "Entry Required"
3370                  .CoInfo_Zip.SetFocus
3380                Else
                      ' ** Everything required is present.
3390                End If
3400              End If
3410            Else
3420              If Nz(.CoInfo_Country, vbNullString) = vbNullString Then
3430                blnRetVal = False
3440                MsgBox "Please enter a Country to continue.", vbInformation + vbOKOnly, "Entry Required"
3450                .CoInfo_Country.SetFocus
3460              Else
                    ' ** Everything required is present.
3470              End If
3480            End If
3490          End If
3500        End If
3510      End If

3520      If blnRetVal = True Then
3530        ReDim arr_varRetVal(16, 0)  ' ** 1 Record of 16 fields.
3540        arr_varRetVal(0, 0) = Trim(.CoInfo_Name)
3550        arr_varRetVal(1, 0) = Trim(.CoInfo_Address1)
3560        arr_varRetVal(2, 0) = Trim(Nz(.CoInfo_Address2, vbNullString))
3570        arr_varRetVal(3, 0) = Trim(.CoInfo_City)
3580        arr_varRetVal(4, 0) = Trim(.CoInfo_State)
3590        arr_varRetVal(5, 0) = Trim(.CoInfo_Zip)
3600        arr_varRetVal(6, 0) = Trim(Nz(.CoInfo_Phone, vbNullString))
3610        arr_varRetVal(7, 0) = .chkIncomeTaxCoding
3620        arr_varRetVal(8, 0) = .chkRevenueExpenseTracking
3630        arr_varRetVal(9, 0) = .chkAccountNoWithType
3640        arr_varRetVal(10, 0) = .chkSeparateCheckingAccounts
3650        arr_varRetVal(11, 0) = .chkTabCopy
3660        arr_varRetVal(12, 0) = .chkLinkRevTaxCodes
3670        arr_varRetVal(13, 0) = .chkSpecialCapGainLoss
3680        arr_varRetVal(14, 0) = intTmp_SpecialCapGainLossOpt  ' ** We won't be setting this here.
3690        arr_varRetVal(15, 0) = Trim(.CoInfo_Country)
3700        arr_varRetVal(16, 0) = Trim(.CoInfo_PostalCode)
3710        blnRetVal = Version_DataXFer("Ret", "CoInfo", arr_varRetVal)  ' ** Module Function: modVersionConvertFuncs2.
3720        DoCmd.Close acForm, THIS_NAME
3730      End If

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

Private Sub cmdNext_KeyDown(KeyCode As Integer, Shift As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdNext_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3820    intShiftDown = (Shift And acShiftMask) > 0
3830    intAltDown = (Shift And acAltMask) > 0
3840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3860      Select Case intRetVal
          Case vbKeyTab
3870        With Me
3880          intRetVal = 0
3890          .cmdCancel.SetFocus
3900        End With
3910      End Select
3920    End If

        ' ** Shift keys.
3930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3940      Select Case intRetVal
          Case vbKeyTab
3950        With Me
3960          intRetVal = 0
3970          .chkSpecialCapGainLoss.SetFocus
3980        End With
3990      End Select
4000    End If

EXITP:
4010    KeyCode = intRetVal
4020    Exit Sub

ERRH:
4030    intRetVal = 0
4040    Select Case ERR.Number
        Case Else
4050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4060    End Select
4070    Resume EXITP

End Sub

Private Sub CoInfo_Name_AfterUpdate()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Name_AfterUpdate"

4110    With Me
4120      If IsNull(.CoInfo_Name) = False Then
            ' ** 30 chars max.
4130        .CoInfo_Name = Left(Trim(.CoInfo_Name), 30)
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

Private Sub CoInfo_Name_KeyDown(KeyCode As Integer, Shift As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4320    intShiftDown = (Shift And acShiftMask) > 0
4330    intAltDown = (Shift And acAltMask) > 0
4340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4370        With Me
4380          intRetVal = 0
4390          .CoInfo_Address1.SetFocus
4400        End With
4410      End Select
4420    End If

        ' ** Shift keys.
4430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4450        With Me
4460          intRetVal = 0
4470          .cmdCancel.SetFocus
4480        End With
4490      End Select
4500    End If

EXITP:
4510    KeyCode = intRetVal
4520    Exit Sub

ERRH:
4530    intRetVal = 0
4540    Select Case ERR.Number
        Case Else
4550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4560    End Select
4570    Resume EXITP

End Sub

Private Sub CoInfo_Address1_AfterUpdate()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Address1_AfterUpdate"

4610    With Me
4620      If IsNull(.CoInfo_Address1) = False Then
            ' ** 30 chars max.
4630        .CoInfo_Address1 = Left(Trim(.CoInfo_Address1), 30)
4640      End If
4650    End With

EXITP:
4660    Exit Sub

ERRH:
4670    Select Case ERR.Number
        Case Else
4680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4690    End Select
4700    Resume EXITP

End Sub

Private Sub CoInfo_Address2_AfterUpdate()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Address2_AfterUpdate"

4810    With Me
4820      If IsNull(.CoInfo_Address2) = False Then
            ' ** 30 chars max.
4830        .CoInfo_Address2 = Left(Trim(.CoInfo_Address2), 30)
4840      End If
4850    End With

EXITP:
4860    Exit Sub

ERRH:
4870    Select Case ERR.Number
        Case Else
4880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4890    End Select
4900    Resume EXITP

End Sub

Private Sub CoInfo_City_AfterUpdate()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_City_AfterUpdate"

5010    With Me
5020      If IsNull(.CoInfo_City) = False Then
            ' ** 30 chars max.
5030        .CoInfo_City = Left(Trim(.CoInfo_City), 30)
5040      End If
5050    End With

EXITP:
5060    Exit Sub

ERRH:
5070    Select Case ERR.Number
        Case Else
5080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5090    End Select
5100    Resume EXITP

End Sub

Private Sub opgLoc_AfterUpdate()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_AfterUpdate"

5210    With Me
5220      Select Case .opgLoc
          Case .opgLoc_optUSA.OptionValue
5230        .CoInfo_State.Enabled = True
5240        .CoInfo_State.BorderColor = CLR_LTBLU2
5250        .CoInfo_State.BackStyle = acBackStyleNormal
5260        .CoInfo_State_lbl2.ForeColor = CLR_RED
5270        .CoInfo_State_lbl2_dim_hi.Visible = False
5280        .CoInfo_Zip.Enabled = True
5290        .CoInfo_Zip.BorderColor = CLR_LTBLU2
5300        .CoInfo_Zip.BackStyle = acBackStyleNormal
5310        .CoInfo_Zip_lbl2.ForeColor = CLR_RED
5320        .CoInfo_Zip_lbl2_dim_hi.Visible = False
5330        .CoInfo_Country.Enabled = False
5340        .CoInfo_Country.BorderColor = WIN_CLR_DISR
5350        .CoInfo_Country.BackStyle = acBackStyleTransparent
5360        .CoInfo_Country = Null
5370        .CoInfo_Country_lbl2.ForeColor = WIN_CLR_DISF
5380        .CoInfo_Country_lbl2_dim_hi.Visible = True
5390        .CoInfo_PostalCode.Enabled = False
5400        .CoInfo_PostalCode.BorderColor = WIN_CLR_DISR
5410        .CoInfo_PostalCode.BackStyle = acBackStyleTransparent
5420        .CoInfo_PostalCode = Null
            '.CoInfo_PostalCode_lbl2.ForeColor = WIN_CLR_DISF  ' ** Not required for non-USA.
            '.CoInfo_PostalCode_lbl2_dim_hi.Visible = True
5430        Select Case blnIsOpen
            Case True
5440          blnIsOpen = False
5450        Case False
5460          .CoInfo_State.SetFocus
5470        End Select
5480      Case .opgLoc_optOther.OptionValue
5490        .CoInfo_State.Enabled = False
5500        .CoInfo_State.BorderColor = WIN_CLR_DISR
5510        .CoInfo_State.BackStyle = acBackStyleTransparent
5520        .CoInfo_State = Null
5530        .CoInfo_Zip.Enabled = False
5540        .CoInfo_Zip.BorderColor = WIN_CLR_DISR
5550        .CoInfo_Zip.BackStyle = acBackStyleTransparent
5560        .CoInfo_Zip = Null
5570        .CoInfo_Country.Enabled = True
5580        .CoInfo_Country.BorderColor = CLR_LTBLU2
5590        .CoInfo_Country.BackStyle = acBackStyleNormal
5600        .CoInfo_PostalCode.Enabled = True
5610        .CoInfo_PostalCode.BorderColor = CLR_LTBLU2
5620        .CoInfo_PostalCode.BackStyle = acBackStyleNormal
5630        .CoInfo_State_lbl2.ForeColor = WIN_CLR_DISF
5640        .CoInfo_State_lbl2_dim_hi.Visible = True
5650        .CoInfo_Zip_lbl2.ForeColor = WIN_CLR_DISF
5660        .CoInfo_Zip_lbl2_dim_hi.Visible = True
5670        .CoInfo_Country_lbl2.ForeColor = CLR_RED
5680        .CoInfo_Country_lbl2_dim_hi.Visible = False
            '.CoInfo_PostalCode_lbl2.ForeColor = CLR_RED  ' ** Not required for non-USA.
            '.CoInfo_PostalCode_lbl2_dim_hi.Visible = False
5690        Select Case blnIsOpen
            Case True
5700          blnIsOpen = False
5710        Case False
5720          .CoInfo_Country.SetFocus
5730        End Select
5740      End Select
5750    End With

EXITP:
5760    Exit Sub

ERRH:
5770    Select Case ERR.Number
        Case Else
5780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
5790    End Select
5800    Resume EXITP

End Sub

Private Sub opgLoc_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optUSA_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5920    intShiftDown = (Shift And acShiftMask) > 0
5930    intAltDown = (Shift And acAltMask) > 0
5940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5970        With Me
5980          intRetVal = 0
5990          If .CoInfo_State.Enabled = True Then
6000            .CoInfo_State.SetFocus
6010          ElseIf .CoInfo_Country.Enabled = True Then
6020            .CoInfo_Country.SetFocus
6030          Else
6040            .CoInfo_Phone.SetFocus
6050          End If
6060        End With
6070      End Select
6080    End If

        ' ** Shift keys.
6090    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6100      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6110        With Me
6120          intRetVal = 0
6130          .CoInfo_City.SetFocus
6140        End With
6150      End Select
6160    End If

EXITP:
6170    KeyCode = intRetVal
6180    Exit Sub

ERRH:
6190    intRetVal = 0
6200    Select Case ERR.Number
        Case Else
6210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6220    End Select
6230    Resume EXITP

End Sub

Private Sub opgLoc_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optOther_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6320    intShiftDown = (Shift And acShiftMask) > 0
6330    intAltDown = (Shift And acAltMask) > 0
6340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6370        With Me
6380          intRetVal = 0
6390          If .CoInfo_State.Enabled = True Then
6400            .CoInfo_State.SetFocus
6410          ElseIf .CoInfo_Country.Enabled = True Then
6420            .CoInfo_Country.SetFocus
6430          Else
6440            .CoInfo_Phone.SetFocus
6450          End If
6460        End With
6470      End Select
6480    End If

        ' ** Shift keys.
6490    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6500      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6510        With Me
6520          intRetVal = 0
6530          .CoInfo_City.SetFocus
6540        End With
6550      End Select
6560    End If

EXITP:
6570    KeyCode = intRetVal
6580    Exit Sub

ERRH:
6590    intRetVal = 0
6600    Select Case ERR.Number
        Case Else
6610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6620    End Select
6630    Resume EXITP

End Sub

Private Sub CoInfo_State_Enter()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_State_Enter"

6710    With Me
6720      If IsNull(.CoInfo_State) Then
6730        ExpandCombo .CoInfo_State  ' ** Module Procedure: modUtilities.
6740      End If
6750    End With

EXITP:
6760    Exit Sub

ERRH:
6770    THAT_PROC = THIS_PROC
6780    That_Erl = Erl
6790    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6800    Resume EXITP

End Sub

Private Sub CoInfo_State_KeyUp(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_State_KeyUp"

6910    With Me
6920      If .CoInfo_State.SelStart <> 0 Then
6930        .CoInfo_State.SelStart = 1
6940        .CoInfo_State.SelLength = 0
6950      End If
6960    End With

EXITP:
6970    Exit Sub

ERRH:
6980    THAT_PROC = THIS_PROC
6990    That_Erl = Erl
7000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7010    Resume EXITP

End Sub

Private Sub CoInfo_State_AfterUpdate()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_State_AfterUpdate"

7110    With Me
7120      If IsNull(.CoInfo_State) = False Then
7130        If Trim(.CoInfo_State) <> vbNullString Then
7140          Select Case IsNull(.CoInfo_State_Pref)
              Case True
7150            .CoInfo_State_Pref = .CoInfo_State
7160          Case False
7170            If .CoInfo_State_Pref <> .CoInfo_State Then
7180              .CoInfo_State_Pref = .CoInfo_State
7190            End If
7200          End Select
7210        End If
7220      End If
7230    End With

EXITP:
7240    Exit Sub

ERRH:
7250    Select Case ERR.Number
        Case Else
7260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
7270    End Select
7280    Resume EXITP

End Sub

Private Sub CoInfo_State_NotInList(NewData As String, Response As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_State_NotInList"

7310    With Me
7320      .CoInfo_State = Null
7330      MsgBox "You Must select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
7340      Response = acDataErrContinue
7350    End With

EXITP:
7360    Exit Sub

ERRH:
7370    THAT_PROC = THIS_PROC
7380    That_Erl = Erl
7390    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7400    Resume EXITP

End Sub

Private Sub CoInfo_Zip_Exit(Cancel As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Zip_Exit"

7510    With Me
7520      If IsNull(.CoInfo_Zip) = False Then
7530        If IsNumeric(Left(.CoInfo_Zip, 5)) = False Or Len(.CoInfo_Zip) < 5 Then
7540          MsgBox "You Must enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
7550          DoCmd.CancelEvent
7560        Else
7570          If InStr(1, .CoInfo_Zip, "-", vbTextCompare) > 0 Then
                ' ** The user typed a long Zip code.
7580            If Len(.CoInfo_Zip) < 10 Then
7590              MsgBox "You Must enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
7600              DoCmd.CancelEvent
7610            Else
7620              If IsNumeric(Right(.CoInfo_Zip, 4)) = False Then
7630                MsgBox "You Must enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
7640                DoCmd.CancelEvent
7650              End If
7660            End If
7670          Else
7680            If Len(.CoInfo_Zip) < 9 And Len(.CoInfo_Zip) <> 5 Then
7690              MsgBox "You Must enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
7700              DoCmd.CancelEvent
7710            End If
7720          End If
7730        End If
7740      End If
7750    End With

EXITP:
7760    Exit Sub

ERRH:
7770    THAT_PROC = THIS_PROC
7780    That_Erl = Erl
7790    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7800    Resume EXITP

End Sub

Private Sub CoInfo_Country_AfterUpdate()

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Country_AfterUpdate"

7910    With Me
7920      If IsNull(.CoInfo_Country) = False Then
            ' ** 30 chars max.
7930        .CoInfo_Country = Left(Trim(.CoInfo_Country), 30)
7940      End If
7950    End With

EXITP:
7960    Exit Sub

ERRH:
7970    Select Case ERR.Number
        Case Else
7980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7990    End Select
8000    Resume EXITP

End Sub

Private Sub CoInfo_PostalCode_AfterUpdate()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_PostalCode_AfterUpdate"

8110    With Me
8120      If IsNull(.CoInfo_PostalCode) = False Then
            ' ** 10 chars max.
8130        .CoInfo_PostalCode = Left(Trim(.CoInfo_PostalCode), 10)
8140      End If
8150    End With

EXITP:
8160    Exit Sub

ERRH:
8170    Select Case ERR.Number
        Case Else
8180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8190    End Select
8200    Resume EXITP

End Sub

Private Sub chkIncomeTaxCoding_AfterUpdate()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncomeTaxCoding_AfterUpdate"

8310    With Me
8320      Select Case .chkIncomeTaxCoding
          Case True
8330        .chkIncomeTaxCoding_lbl.FontBold = True
8340        If .chkRevenueExpenseTracking = True Then
8350          .chkLinkRevTaxCodes.Enabled = True
8360        End If
8370      Case False
8380        .chkIncomeTaxCoding_lbl.FontBold = False
8390        .chkLinkRevTaxCodes = False
8400        .chkLinkRevTaxCodes.Enabled = False
8410      End Select
8420    End With

EXITP:
8430    Exit Sub

ERRH:
8440    Select Case ERR.Number
        Case Else
8450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8460    End Select
8470    Resume EXITP

End Sub

Private Sub chkIncomeTaxCoding_KeyDown(KeyCode As Integer, Shift As Integer)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncomeTaxCoding_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8520    intShiftDown = (Shift And acShiftMask) > 0
8530    intAltDown = (Shift And acAltMask) > 0
8540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8570        With Me
8580          intRetVal = 0
8590          .chkRevenueExpenseTracking.SetFocus
8600        End With
8610      End Select
8620    End If

        ' ** Shift keys.
8630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8650        With Me
8660          intRetVal = 0
8670          .CoInfo_Phone.SetFocus
8680        End With
8690      End Select
8700    End If

EXITP:
8710    KeyCode = intRetVal
8720    Exit Sub

ERRH:
8730    intRetVal = 0
8740    Select Case ERR.Number
        Case Else
8750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8760    End Select
8770    Resume EXITP

End Sub

Private Sub chkRevenueExpenseTracking_AfterUpdate()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRevenueExpenseTracking_AfterUpdate"

8810    With Me
8820      Select Case .chkRevenueExpenseTracking
          Case True
8830        .chkRevenueExpenseTracking_lbl.FontBold = True
8840        If .chkIncomeTaxCoding = True Then
8850          .chkLinkRevTaxCodes.Enabled = True
8860        End If
8870      Case False
8880        .chkRevenueExpenseTracking_lbl.FontBold = False
8890        .chkLinkRevTaxCodes = False
8900        .chkLinkRevTaxCodes.Enabled = False
8910      End Select
8920    End With

EXITP:
8930    Exit Sub

ERRH:
8940    Select Case ERR.Number
        Case Else
8950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8960    End Select
8970    Resume EXITP

End Sub

Private Sub chkSeparateCheckingAccounts_AfterUpdate()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkSeparateCheckingAccounts_AfterUpdate"

9010    With Me
9020      Select Case .chkSeparateCheckingAccounts
          Case True
9030        .chkSeparateCheckingAccounts_lbl.FontBold = True
9040      Case False
9050        .chkSeparateCheckingAccounts_lbl.FontBold = False
9060      End Select
9070    End With

EXITP:
9080    Exit Sub

ERRH:
9090    Select Case ERR.Number
        Case Else
9100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9110    End Select
9120    Resume EXITP

End Sub

Private Sub chkTabCopy_AfterUpdate()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkTabCopy_AfterUpdate"

9210    With Me
9220      Select Case .chkTabCopy
          Case True
9230        .chkTabCopy_lbl.FontBold = True
9240      Case False
9250        .chkTabCopy_lbl.FontBold = False
9260      End Select
9270    End With

EXITP:
9280    Exit Sub

ERRH:
9290    Select Case ERR.Number
        Case Else
9300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9310    End Select
9320    Resume EXITP

End Sub

Private Sub chkAccountNoWithType_AfterUpdate()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkAccountNoWithType_AfterUpdate"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim blnSkip As Boolean

9410    blnSkip = True

9420    With Me
9430      Select Case .chkAccountNoWithType
          Case True
9440        .chkAccountNoWithType_lbl.FontBold = True
9450        If blnSkip = False Then
9460          Set dbs = CurrentDb
9470          Set rst = dbs.OpenRecordset("SELECT account.* FROM account;")
9480          If rst.BOF = True And rst.EOF = True Then
                ' ** No accounting for taste.
9490          Else
                ' ** There are current accounts in the old format.
                ' ** Go through the account table and edit the
                ' ** account number to be account type + account number.
9500            rst.MoveLast
9510            rst.MoveFirst
9520            While Not rst.EOF
9530              rst.Edit
9540              rst![accountno] = rst![accounttype] + "-" + rst![accountno]
9550              rst![related_accountno] = rst![accounttype] + "-" + rst![related_accountno]
9560              rst.Update
9570              rst.MoveNext
9580            Wend
9590          End If
9600          rst.Close
9610          dbs.Close
9620        End If  ' ** blnSkip.
9630      Case False
9640        .chkAccountNoWithType_lbl.FontBold = False
9650        If blnSkip = False Then
9660          Set dbs = CurrentDb
9670          Set rst = dbs.OpenRecordset("SELECT [account].* FROM [account]")
9680          If rst.BOF = True And rst.EOF = True Then
                ' ** Taste the difference.
9690          Else
                ' ** There are current accounts in the old format.
                ' ** Go through the account table and take out
                ' ** the account type from the account number
9700            rst.MoveLast
9710            rst.MoveFirst
9720            While Not rst.EOF
9730              If Len(rst![accountno]) > 4 Or (Len(rst![related_accountno]) > 0 And Len(rst![related_accountno]) > 4) Then
9740                rst.Edit ' ** If one is less than 4 characters, something else is up and that one should not be changed.
9750                rst![accountno] = Mid(rst![accountno], 4)
9760                rst![related_accountno] = Mid(rst![related_accountno], 4)
9770                rst.Update
9780              End If
9790              rst.MoveNext
9800            Wend
9810          End If
9820          rst.Close
9830          dbs.Close
9840        End If  ' ** blnSkip.
9850      End Select

9860    End With

EXITP:
9870    Set rst = Nothing
9880    Set dbs = Nothing
9890    Exit Sub

ERRH:
9900    Select Case ERR.Number
        Case Else
9910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9920    End Select
9930    Resume EXITP

End Sub

Private Sub chkAccountNoWithType_KeyDown(KeyCode As Integer, Shift As Integer)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAccountNoWithType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10020   intShiftDown = (Shift And acShiftMask) > 0
10030   intAltDown = (Shift And acAltMask) > 0
10040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10070       With Me
10080         intRetVal = 0
10090         If .chkLinkRevTaxCodes.Enabled = True Then
10100           .chkLinkRevTaxCodes.SetFocus
10110         Else
10120           .chkSpecialCapGainLoss.SetFocus
10130         End If
10140       End With
10150     End Select
10160   End If

        ' ** Shift keys.
10170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10180     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10190       With Me
10200         intRetVal = 0
10210         .chkTabCopy.SetFocus
10220       End With
10230     End Select
10240   End If

EXITP:
10250   KeyCode = intRetVal
10260   Exit Sub

ERRH:
10270   intRetVal = 0
10280   Select Case ERR.Number
        Case Else
10290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10300   End Select
10310   Resume EXITP

End Sub

Private Sub chkLinkRevTaxCodes_AfterUpdate()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLinkRevTaxCodes_AfterUpdate"

10410   With Me
10420     Select Case .chkLinkRevTaxCodes
          Case True
10430       .chkLinkRevTaxCodes_lbl.FontBold = True
10440     Case False
10450       .chkLinkRevTaxCodes_lbl.FontBold = False
10460     End Select
10470   End With

EXITP:
10480   Exit Sub

ERRH:
10490   Select Case ERR.Number
        Case Else
10500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10510   End Select
10520   Resume EXITP

End Sub

Private Sub chkLinkRevTaxCodes_KeyDown(KeyCode As Integer, Shift As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLinkRevTaxCodes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10620   intShiftDown = (Shift And acShiftMask) > 0
10630   intAltDown = (Shift And acAltMask) > 0
10640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10670       With Me
10680         intRetVal = 0
10690         .chkSpecialCapGainLoss.SetFocus
10700       End With
10710     End Select
10720   End If

        ' ** Shift keys.
10730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10750       With Me
10760         intRetVal = 0
10770         .chkAccountNoWithType.SetFocus
10780       End With
10790     End Select
10800   End If

EXITP:
10810   KeyCode = intRetVal
10820   Exit Sub

ERRH:
10830   intRetVal = 0
10840   Select Case ERR.Number
        Case Else
10850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10860   End Select
10870   Resume EXITP

End Sub

Private Sub chkSpecialCapGainLoss_AfterUpdate()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkSpecialCapGainLoss_AfterUpdate"

10910   With Me
10920     Select Case .chkSpecialCapGainLoss
          Case True
10930       .chkSpecialCapGainLoss_lbl.FontBold = True
10940     Case False
10950       .chkSpecialCapGainLoss_lbl.FontBold = False
10960     End Select
10970   End With

EXITP:
10980   Exit Sub

ERRH:
10990   Select Case ERR.Number
        Case Else
11000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11010   End Select
11020   Resume EXITP

End Sub

Private Sub chkSpecialCapGainLoss_KeyDown(KeyCode As Integer, Shift As Integer)

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkSpecialCapGainLoss_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11120   intShiftDown = (Shift And acShiftMask) > 0
11130   intAltDown = (Shift And acAltMask) > 0
11140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11170       With Me
11180         intRetVal = 0
11190         .cmdNext.SetFocus
11200       End With
11210     End Select
11220   End If

        ' ** Shift keys.
11230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11250       With Me
11260         intRetVal = 0
11270         If .chkLinkRevTaxCodes.Enabled = True Then
11280           .chkLinkRevTaxCodes.SetFocus
11290         Else
11300           .chkAccountNoWithType.SetFocus
11310         End If
11320       End With
11330     End Select
11340   End If

EXITP:
11350   KeyCode = intRetVal
11360   Exit Sub

ERRH:
11370   intRetVal = 0
11380   Select Case ERR.Number
        Case Else
11390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11400   End Select
11410   Resume EXITP

End Sub

Private Sub ResetDisabled()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "ResetDisabled"

        Dim ctl As Access.Control

11510   With Me
11520     For Each ctl In .Controls
11530       With ctl
11540         Select Case .ControlType
              Case acTextBox, acComboBox, acCheckBox, acCommandButton
11550           .Enabled = True
11560           If .ControlType = acTextBox Or .ControlType = acComboBox Then
11570             .BorderColor = CLR_LTBLU2
11580             .BackStyle = acBackStyleNormal
11590           End If
11600         Case Else
                ' ** Doesn't matter.
11610         End Select
11620       End With
11630     Next
11640     .CoInfo_lbl.ForeColor = CLR_DKGRY2
11650     .CoInfo_lbl_dim_hi.Visible = False
11660     .CoInfo_Name_lbl2.ForeColor = CLR_RED
11670     .CoInfo_Name_lbl2_dim_hi.Visible = False
11680     .CoInfo_Address1_lbl2.ForeColor = CLR_RED
11690     .CoInfo_Address1_lbl2_dim_hi.Visible = False
11700     .CoInfo_City_lbl2.ForeColor = CLR_RED
11710     .CoInfo_City_lbl2_dim_hi.Visible = False
11720     Select Case .opgLoc
          Case .opgLoc_optUSA.OptionValue
11730       .CoInfo_State_lbl2.ForeColor = CLR_RED
11740       .CoInfo_State_lbl2_dim_hi.Visible = False
11750       .CoInfo_Zip_lbl2.ForeColor = CLR_RED
11760       .CoInfo_Zip_lbl2_dim_hi.Visible = False
11770       .CoInfo_Country.Enabled = False
11780       .CoInfo_Country.BorderColor = WIN_CLR_DISR
11790       .CoInfo_Country.BackStyle = acBackStyleTransparent
11800       .CoInfo_Country_lbl2.ForeColor = WIN_CLR_DISF
11810       .CoInfo_Country_lbl2_dim_hi.Visible = True
11820       .CoInfo_PostalCode.Enabled = False
11830       .CoInfo_PostalCode.BorderColor = WIN_CLR_DISR
11840       .CoInfo_PostalCode.BackStyle = acBackStyleTransparent
11850     Case .opgLoc_optOther.OptionValue
11860       .CoInfo_Country_lbl2.ForeColor = CLR_RED  ' ** Postal Code not required for non-USA.
11870       .CoInfo_Country_lbl2_dim_hi.Visible = False
11880       .CoInfo_State.Enabled = False
11890       .CoInfo_State.BorderColor = WIN_CLR_DISR
11900       .CoInfo_State.BackStyle = acBackStyleTransparent
11910       .CoInfo_Zip.Enabled = False
11920       .CoInfo_Zip.BorderColor = WIN_CLR_DISR
11930       .CoInfo_Zip.BackStyle = acBackStyleTransparent
11940     End Select
11950     .CoInfo_ReqMsg_lbl1.ForeColor = CLR_RED
11960     .CoInfo_ReqMsg_lbl1_dim_hi.Visible = False
11970     .CoInfo_ReqMsg_lbl2.ForeColor = CLR_RED
11980     .CoInfo_ReqMsg_lbl2_dim_hi.Visible = False
11990     .Options_lbl.ForeColor = CLR_DKGRY2
12000     .Options_lbl_dim_hi.Visible = False
12010     .CoInfo_Name.SetFocus

12020     .Repaint
12030   End With

EXITP:
12040   Set ctl = Nothing
12050   Exit Sub

ERRH:
12060   Select Case ERR.Number
        Case Else
12070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12080   End Select
12090   Resume EXITP

End Sub
