VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCurrency_History_Date_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCurrency_History_Date_Sub"

'VGC 11/17/2015: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel on frmCurrency_History_Date}
' **   OK:               O {cmdOK on frmCurrency_History_Date}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[curr_code], [curr_date] DESC"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine1 As String = "Sort_line1"
Private Const strSortLine2 As String = "Sort_line2"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmCurrency_History_Date" Then

160         blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
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

410     With Me

420       CLR_DISABLED_FG = CLR_DKGRY
430       CLR_DISABLED_BG = CLR_LTTEAL

440       .curr_code.ForeColor = CLR_DISABLED_FG
450       .curr_code.BackColor = CLR_DISABLED_BG
460       .curr_name.ForeColor = CLR_DISABLED_FG
470       .curr_name.BackColor = CLR_DISABLED_BG
480       .curr_date.ForeColor = CLR_DISABLED_FG
490       .curr_date.BackColor = CLR_DISABLED_BG
500       .curr_rate1.ForeColor = CLR_DISABLED_FG
510       .curr_rate1.BackColor = CLR_DISABLED_BG
520       .curr_rate1_hist.ForeColor = CLR_DISABLED_FG
530       .curr_rate1_hist.BackColor = CLR_DISABLED_BG
540       .curr_rate2.ForeColor = CLR_DISABLED_FG
550       .curr_rate2.BackColor = CLR_DISABLED_BG
560       .curr_rate2_hist.ForeColor = CLR_DISABLED_FG
570       .curr_rate2_hist.BackColor = CLR_DISABLED_BG

580       SortNow THIS_PROC  ' ** Procedure: Below.

590     End With

EXITP:
600     Exit Sub

ERRH:
610     Select Case ERR.Number
        Case Else
620       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
630     End Select
640     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

710     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
720     intShiftDown = (Shift And acShiftMask) > 0
730     intAltDown = (Shift And acAltMask) > 0
740     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         With Me
780           intRetVal = 0
790           .Parent.cmdCancel_Click  ' ** Form Procedure: frmCurrency_History_Date.
800         End With
810       Case vbKeyUp
820         intRetVal = 0
830         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
840       Case vbKeyDown
850         intRetVal = 0
860         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
870       End Select
880     End If

        ' ** Ctrl keys.
890     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
900       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
910         With Me
920           intRetVal = 0
930           DoCmd.SelectObject acForm, .Parent.Name, False
940           .Parent.cmdOK.SetFocus
950         End With
960       End Select
970     End If

        ' ** Ctrl-Shift keys.
980     If intCtrlDown And (Not intAltDown) And intShiftDown Then
990       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1000        With Me
1010          intRetVal = 0
1020          DoCmd.SelectObject acForm, .Parent.Name, False
1030          .Parent.cmdCancel.SetFocus
1040        End With
1050      End Select
1060    End If

EXITP:
1070    KeyCode = intRetVal
1080    Exit Sub

ERRH:
1090    intRetVal = 0
1100    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1110    Case Else
1120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1130    End Select
1140    Resume EXITP

End Sub

Private Sub curr_code_lbl_DblClick(Cancel As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_lbl_DblClick"

1210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1220    Exit Sub

ERRH:
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub curr_name_lbl_DblClick(Cancel As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_name_lbl_DblClick"

1310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1320    Exit Sub

ERRH:
1330    Select Case ERR.Number
        Case Else
1340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1350    End Select
1360    Resume EXITP

End Sub

Private Sub curr_date_lbl_DblClick(Cancel As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_date_lbl_DblClick"

1410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1420    Exit Sub

ERRH:
1430    Select Case ERR.Number
        Case Else
1440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1450    End Select
1460    Resume EXITP

End Sub

Private Sub curr_rate1_lbl_DblClick(Cancel As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate1_lbl_DblClick"

1510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1520    Exit Sub

ERRH:
1530    Select Case ERR.Number
        Case Else
1540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1550    End Select
1560    Resume EXITP

End Sub

Private Sub curr_rate1_hist_lbl_DblClick(Cancel As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate1_hist_lbl_DblClick"

1610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1620    Exit Sub

ERRH:
1630    Select Case ERR.Number
        Case Else
1640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1650    End Select
1660    Resume EXITP

End Sub

Private Sub R1x_lbl_DblClick(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "R1x_lbl_DblClick"

1710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub

Private Sub curr_rate2_lbl_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate2_lbl_DblClick"

1810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub curr_rate2_hist_lbl_DblClick(Cancel As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate2_hist_lbl_DblClick"

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

Private Sub R2x_lbl_DblClick(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "R2x_lbl_DblClick"

2010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2020    Exit Sub

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2050    End Select
2060    Resume EXITP

End Sub

Public Function RecCnt() As Long

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

2110    lngRetVal = 0&

2120    With Me
2130      Set rst = .RecordsetClone
2140      With rst
2150        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
2160        Else
2170          .MoveLast
2180          lngRetVal = .RecordCount
2190        End If
2200        .Close
2210      End With
2220    End With

EXITP:
2230    Set rst = Nothing
2240    RecCnt = lngRetVal
2250    Exit Function

ERRH:
2260    lngRetVal = 0&
2270    Select Case ERR.Number
        Case Else
2280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2290    End Select
2300    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

2410    Select Case intWhere
        Case acCmdRecordsGoToFirst
2420      DoCmd.RunCommand acCmdRecordsGoToFirst
2430    Case acCmdRecordsGoToLast
2440      DoCmd.RunCommand acCmdRecordsGoToLast
2450    Case acCmdRecordsGoToPrevious
2460      DoCmd.RunCommand acCmdRecordsGoToPrevious
2470    Case acCmdRecordsGoToNext
2480      DoCmd.RunCommand acCmdRecordsGoToNext
2490    Case Else
2500      If IsMissing(varID) = False Then
2510        With Me
2520          Set rst = .RecordsetClone
2530          With rst
2540            .FindFirst "[currhist_id] = '" & varID & "'"
2550            If .NoMatch = False Then
2560              Me.Bookmark = .Bookmark
2570            End If
2580            .Close
2590          End With
2600        End With
2610      End If
2620    End Select

EXITP:
2630    Set rst = Nothing
2640    Exit Sub

ERRH:
2650    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
2660    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
2670    Case Else
2680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2690    End Select
2700    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [curr_code], [curr_date] DESC"

2810    With Me
2820      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
2830        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
2840      End If
2850      .Controls(strSortLbl).Visible = False
2860      .Controls(strSortLine1).Visible = False
2870      .Controls(strSortLine2).Visible = False
2880      .Controls(strSortLine1).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
2890      .Controls(strSortLine2).Width = lngTpp
2900      If strProc = "Form_Load" Then
2910        strCalled = "curr_code"
2920        strSortNow = strSortOrig
2930        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
2940        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
2950        lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
2960        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
2970        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
2980        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
2990        .Controls(strSortLbl).Top = lngSortLbl_Top
3000        .Controls(strSortLbl).Left = lngSortLbl_Left
3010        .Controls(strSortLine1).Top = lngSortLine_Top
3020        .Controls(strSortLine2).Top = (lngSortLine_Top + lngTpp)
3030        .Controls(strSortLine1).Left = lngSortLine_Left
3040        .Controls(strSortLine2).Left = lngSortLine_Left
3050        .Controls(strSortLine1).Width = lngSortLine_Width
3060        .Controls(strSortLine2).Width = lngSortLine_Width
3070        .Controls(strSortLbl).Caption = strArwUp
3080        .Controls(strSortLbl).ForeColor = CLR_DKBLU
3090      Else
3100        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
3110        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
3120        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
3130        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
3140        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
3150        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
3160        .Controls(strSortLbl).Top = lngSortLbl_Top
3170        .Controls(strSortLbl).Left = lngSortLbl_Left
3180        .Controls(strSortLine1).Top = lngSortLine_Top
3190        .Controls(strSortLine2).Top = (lngSortLine_Top + lngTpp)
3200        .Controls(strSortLine1).Left = lngSortLine_Left
3210        .Controls(strSortLine2).Left = lngSortLine_Left
3220        .Controls(strSortLine1).Width = lngSortLine_Width
3230        .Controls(strSortLine2).Width = lngSortLine_Width
3240        If strCalled = "curr_code" Then
3250          strSortAsc = strSortOrig
3260          If strSortNow = strSortAsc Then
3270            strSortNow = "[curr_code] DESC, [curr_date]"
3280            .Controls(strSortLbl).Caption = strArwDn
3290            .Controls(strSortLbl).ForeColor = CLR_DKRED
3300          Else
3310            strSortNow = strSortAsc
3320            .Controls(strSortLbl).Caption = strArwUp
3330            .Controls(strSortLbl).ForeColor = CLR_DKBLU
3340          End If
3350          lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
3360          .Controls(strSortLbl).Left = lngSortLbl_Left
3370        Else
3380          .Controls(strSortLbl).Caption = strArwUp
3390          .Controls(strSortLbl).ForeColor = CLR_DKBLU
3400          Select Case strCalled
              Case "curr_name"
3410            strSortAsc = "[curr_name], [curr_date] DESC"
3420          Case "curr_date"
3430            strSortAsc = "[curr_date]" & strStdAsc
3440            lngSortLbl_Left = (lngSortLbl_Left + lngTpp)
3450            .Controls(strSortLbl).Left = lngSortLbl_Left
3460          Case "curr_rate1"
3470            strSortAsc = "[curr_rate1]" & strStdAsc
3480          Case "curr_rate1_hist"
3490            strSortAsc = "[curr_rate1_hist]" & strStdAsc
3500          Case "R1x"
3510            strSortAsc = "[R1x]" & strStdAsc
3520            lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
3530            .Controls(strSortLbl).Left = lngSortLbl_Left
3540          Case "curr_rate2"
3550            strSortAsc = "[curr_rate2]" & strStdAsc
3560          Case "curr_rate2_hist"
3570            strSortAsc = "[curr_rate2_hist]" & strStdAsc
3580          Case "R2x"
3590            strSortAsc = "[R2x]" & strStdAsc
3600            lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
3610            .Controls(strSortLbl).Left = lngSortLbl_Left
3620          End Select
3630          If strSortNow = strSortAsc Then
3640            intCnt = CharCnt(strSortAsc, ",") + 1
3650            Select Case intCnt
                Case 1
3660              strTmp01 = strSortAsc & " DESC"
3670            Case 2
3680              intPos01 = InStr(strSortAsc, ",")
3690              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
3700              strTmp02 = Mid(strSortAsc, intPos01)
3710              If Right(strTmp02, 5) = " DESC" Then
3720                strTmp02 = Trim(Left(strTmp02, (Len(strTmp02) - 4)))
3730                strTmp01 = strTmp01 & strTmp02
3740              Else
3750                strTmp01 = strTmp01 & strTmp02 & " DESC"
3760              End If
3770            Case 3
3780              intPos01 = InStr(strSortAsc, ",")
3790              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
3800              strTmp02 = Mid(strSortAsc, intPos01)
3810              intPos01 = InStr(2, strTmp02, ",")
3820              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
3830              strTmp02 = Mid(strTmp02, intPos01)
3840              If Right(strTmp02, 5) = " DESC" Then
3850                strTmp02 = Trim(Left(strTmp02, (Len(strTmp02) - 4)))
3860                strTmp01 = strTmp01 & strTmp02
3870              Else
3880                strTmp01 = strTmp01 & strTmp02 & " DESC"
3890              End If
3900            End Select
3910            strSortNow = strTmp01
3920            .Controls(strSortLbl).Caption = strArwDn
3930            .Controls(strSortLbl).ForeColor = CLR_DKRED
3940          Else
3950            strSortNow = strSortAsc
3960          End If
3970        End If
3980      End If
3990      .Controls(strSortLbl).Visible = True
4000      .Controls(strSortLine1).Visible = True
4010      .Controls(strSortLine2).Visible = True
4020      .OrderBy = strSortNow
4030      .OrderByOn = True
4040    End With

EXITP:
4050    Exit Sub

ERRH:
4060    Select Case ERR.Number
        Case Else
4070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4080    End Select
4090    Resume EXITP

End Sub
