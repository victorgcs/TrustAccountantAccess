Attribute VB_Name = "modJrnlCol_Keys"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modJrnlCol_Keys"

'VGC 09/09/2017: CHANGES!

'tblJOURNAL_FIELD DOESN'T GET AUTOMATICALLY UPDATED!!!

' ** Array: arr_varJType().
Private lngJTypes As Long, arr_varJType As Variant
Private Const J_JTYPE As Integer = 0
'Private Const J_ORD   As Integer = 1
Private Const J_E1    As Integer = 2
Private Const J_E2    As Integer = 3

' ** Array: arr_varTab().
Private lngTabs As Long, arr_varTab As Variant
Private Const T_JTYPE  As Integer = 0
Private Const T_CTLNAM As Integer = 1
Private Const T_CTLTYP As Integer = 2
Private Const T_ACTIVE As Integer = 3

' ** Array: arr_varTab2().
Private lngTabs2 As Long, arr_varTab2() As Variant
Private Const T2_ELEMS As Integer = 4  ' ** Array's first-element UBound().
Private Const T2_JTYPE  As Integer = 0
Private Const T2_CTLNAM As Integer = 1
Private Const T2_CTLTYP As Integer = 2
Private Const T2_ACTIVE As Integer = 3
Private Const T2_ACTNOW As Integer = 4

' ** Dummy variable for the subform's JC_Key_Sub_Next() function.
Private blnD1 As Boolean, blnD2 As Boolean

Private lngRecsCur As Long
Private strPageMoveCtl As String
' **

Public Function JC_Key_Sub(frmSub As Access.Form, strProc As String, intKeyCode As Integer, intAux As Integer, blnNextRec As Boolean, blnFromZero As Boolean, blnToTaxLot As Boolean, Optional varAccountNo_Tab As Variant, Optional varAccountNo_Next As Variant) As String
' ** intAux:
' **   0 : Plain keys.
' **   1 : Shift keys.
' **   2 : Tab copies accountno.
' **   3 : Tax Lot screen.
' **   4 : CommitRec.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Key_Sub"

        Dim strThis As String, strNext As String, strThisJType As String
        Dim blnAcctNoTab As Boolean, blnContinue As Boolean, blnSpecCashNo As Boolean, blnWarnZeroCost As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String
        Dim strRetVal As String, lngRetVal As Long, intRetVal As Integer

110     strRetVal = vbNullString
120     strPageMoveCtl = vbNullString
130     strThis = Left(strProc, (InStr(strProc, "_KeyDown") - 1))
140     blnSpecCashNo = False: blnWarnZeroCost = False

150     With frmSub  ' ** This is the subform, frmJournal_Columns_Sub.
160       Select Case intAux

          Case 0
            ' **************************
            ' ** Plain keys.
            ' **************************
170         Select Case intKeyCode
            Case vbKeyTab
180           strNext = JC_Key_Sub_Next(strProc, blnNextRec, blnFromZero)  ' ** Function: Below.
190           Select Case blnNextRec
              Case True
200             blnNextRec = False
210             lngRecsCur = .RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
220             If .CurrentRecord < lngRecsCur Then
230               .MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Columns_Sub.
240               .Controls(strNext).SetFocus
250             Else
260               strNext = JC_Key_Par_Next(.Name & "_Exit", .Parent)  ' ** Function: Below.
270               DoCmd.SelectObject acForm, .Parent.Name, False
280               .Parent.cmdAdd.SetFocus  '.Parent.Controls(strNext).SetFocus
290             End If
300             lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
310           Case False
320   On Error Resume Next
330             .Controls(strNext).SetFocus
340   On Error GoTo ERRH
350             If strNext = "shareface" Or strNext = "icash" Or strNext = "pcash" Or strNext = "cost" Then
360               lngRetVal = fSetScrollBarPosHZ(frmSub, 999&)  ' ** Module Function: modScrollBarFuncs.
370             End If
380           End Select
390         Case vbKeyReturn
400           strNext = JC_Key_Sub_Next(strProc, blnNextRec, blnFromZero)  ' ** Function: Below.
410           Select Case blnNextRec
              Case True
420             blnNextRec = False
430             lngRecsCur = .RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
440             Select Case .Parent.opgEnterKey
                Case .Parent.opgEnterKey_optRight.OptionValue
450               If .CurrentRecord < lngRecsCur Then
460                 .MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Columns_Sub.
470                 .Controls(strNext).SetFocus
480               Else
490                 strNext = JC_Key_Par_Next(.Name & "_Exit", .Parent)  ' ** Function: Below.
500                 DoCmd.SelectObject acForm, .Parent.Name, False
510                 .Parent.cmdAdd.SetFocus  '.Parent.Controls(strNext).SetFocus
520               End If
530             Case .Parent.opgEnterKey_optDown.OptionValue
540               strPageMoveCtl = strThis
550               If .CurrentRecord < lngRecsCur Then
560                 .MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Columns_Sub.
570                 Select Case strThis
                    Case "cmdCalendar1"
580                   strThis = "transdate"
590                 Case "cmdCalendar2"
600                   strThis = "assetdate_display"
610                 End Select
620                 strPageMoveCtl = strThis
630                 .Controls(strThis).SetFocus
640               Else
650                 strNext = JC_Key_Par_Next(.Name & "_Exit", .Parent)  ' ** Function: Below.
660                 DoCmd.SelectObject acForm, .Parent.Name, False
670                 .Parent.cmdAdd.SetFocus  '.Parent.Controls(strNext).SetFocus
680               End If
690             End Select
700             lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
710           Case False
720             Select Case .Parent.opgEnterKey
                Case .Parent.opgEnterKey_optRight.OptionValue
730   On Error Resume Next
740               .Controls(strNext).SetFocus
750   On Error GoTo ERRH
760               If strNext = "shareface" Or strNext = "icash" Or strNext = "pcash" Or strNext = "cost" Then
770                 lngRetVal = fSetScrollBarPosHZ(frmSub, 999&)  ' ** Module Function: modScrollBarFuncs.
780               End If
790             Case .Parent.opgEnterKey_optDown.OptionValue
800               strPageMoveCtl = strThis
810               lngRecsCur = .RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
820               If .CurrentRecord < lngRecsCur Then
830                 .MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Columns_Sub.
840                 Select Case strThis
                    Case "cmdCalendar1"
850                   strThis = "transdate"
860                 Case "cmdCalendar2"
870                   strThis = "assetdate_display"
880                 End Select
890                 strPageMoveCtl = strThis
900                 .Controls(strThis).SetFocus
910               Else
920                 .Controls(strNext).SetFocus
930               End If
940             End Select
950           End Select
960         End Select

970       Case 1
            ' **************************
            ' ** Shift keys.
            ' **************************
980         Select Case intKeyCode
            Case vbKeyTab
990           strNext = JC_Key_Sub_Next(strProc, blnNextRec, blnFromZero, False)  ' ** Function: Below.
1000          Select Case blnNextRec
              Case True
1010            blnNextRec = False
1020            If .CurrentRecord > 1 Then
1030              .MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmJournal_Columns_Sub.
1040              .Controls(strNext).SetFocus
1050            Else
1060              strNext = JC_Key_Par_Next(.Name & "_Exit", .Parent, False)  ' ** Function: Below.
1070              DoCmd.SelectObject acForm, .Parent.Name, False
1080              .Parent.cmdAdd.SetFocus  '.Parent.Controls(strNext).SetFocus
1090              lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
1100            End If
1110          Case False
1120  On Error Resume Next
1130            .Controls(strNext).SetFocus
1140  On Error GoTo ERRH
1150          End Select
1160        Case vbKeyReturn
1170          strNext = JC_Key_Sub_Next(strProc, blnNextRec, blnFromZero, False)  ' ** Function: Below.
1180          Select Case blnNextRec
              Case True
1190            blnNextRec = False
1200            If .CurrentRecord > 1 Then
1210              .MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmJournal_Columns_Sub.
1220              .Controls(strNext).SetFocus
1230            Else
1240              strNext = JC_Key_Par_Next(.Name & "_Exit", .Parent, False)  ' ** Function: Below.
1250              DoCmd.SelectObject acForm, .Parent.Name, False
1260              .Parent.cmdAdd.SetFocus  '.Parent.Controls(strNext).SetFocus
1270              lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
1280            End If
1290          Case False
1300  On Error Resume Next
1310            .Controls(strNext).SetFocus
1320  On Error GoTo ERRH
1330          End Select
1340        End Select

1350      Case 2
            ' **************************
            ' ** Tab copies accountno.
            ' **************************
1360        blnAcctNoTab = False
1370        strNext = JC_Key_Sub_Next(strProc, blnNextRec, blnFromZero)  ' ** Function: Below.
1380        If strNext = varAccountNo_Next Then
1390          Select Case varAccountNo_Next
              Case "accountno"
1400            If gblnTabCopyAccount = True And varAccountNo_Tab <> vbNullString Then
1410              If IsNull(.accountno) = True Then
1420                blnAcctNoTab = True
1430                .accountno = varAccountNo_Tab
1440                .accountno2.Requery
1450                .shortname = .accountno.Column(1)
1460              End If
1470            End If
1480          Case "accountno2"
1490            If gblnTabCopyAccount = True And varAccountNo_Tab <> vbNullString Then
1500              If IsNull(.accountno) = True And IsNull(.accountno2) = True Then
1510                blnAcctNoTab = True
1520                .accountno2 = varAccountNo_Tab
1530                .accountno.Requery
1540                .shortname = .accountno2.Column(1)
1550              End If
1560            End If
1570          End Select
1580        End If
1590        Select Case blnAcctNoTab
            Case True
1600          .journaltype.SetFocus
1610        Case False
1620          .Controls(strNext).SetFocus
1630        End Select

1640      Case 3
            ' **************************
            ' ** Tax Lot screen.
            ' **************************
1650        Select Case intKeyCode
            Case vbKeyTab
1660          blnContinue = True
1670          strNext = JC_Key_Sub_Next(strProc, blnNextRec, blnFromZero)  ' ** Function: Below.
1680          Select Case .posted
              Case True
                ' ** Proceed normally.
1690          Case False
1700            strThisJType = Nz(.journaltype, vbNullString)
1710            Select Case strThis
                Case "shareface"
1720              strTmp01 = "Withdrawn"
1730              strTmp02 = "Withdrawn"
1740              strTmp03 = "Withdrawn"
1750            Case "pcash"
1760              strTmp01 = "Withdrawn"
1770              strTmp02 = "Sold"
1780              strTmp03 = "Liability (-)"
1790            End Select
1800            Select Case strThisJType
                Case strTmp01, strTmp02, strTmp03
1810              If .Cost = 0@ And .shareface <> 0# Then
1820                Select Case strThisJType
                    Case "Withdrawn"
1830                  If .ICash <> 0@ Or .PCash <> 0@ Then
1840                    If gblnClosing = False And gblnDeleting = False Then
1850                      Beep
1860                      DoCmd.Hourglass False
1870                      MsgBox "No cash is allowed for a Withdrawn transaction.", vbInformation + vbOKOnly, "Invalid Entry"
1880                      blnSpecCashNo = True
1890                    End If
1900                  End If
1910                Case "Sold"
1920                  If .ICash = 0@ And .PCash = 0@ Then
1930  On Error Resume Next
1940                    strTmp04 = .PCash.text
1950  On Error GoTo ERRH
1960                    If Val(strTmp04) = 0 Then
1970                      If gblnClosing = False And gblnDeleting = False Then
1980                        blnWarnZeroCost = True
1990                        DoCmd.Hourglass False
2000                        msgResponse = MsgBox("Are you sure you want Income and Principal Cash to be ZERO?" & vbCrLf & vbCrLf & _
                              "As would be the case for the sale of worthless shares.", vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
2010                        If msgResponse <> vbYes Then
2020                          blnSpecCashNo = True
2030                        Else
                              ' ** Let sub know they've been warned!
                              ' ** blnZeroCash     : Local to PCash_Exit() only.
                              ' ** blnWarnZeroCost : Not checked during Commit.
2040                        End If
2050                      End If
2060                    End If
2070                  ElseIf .PCash = 0@ Then
2080  On Error Resume Next
2090                    strTmp04 = .PCash.text
2100  On Error GoTo ERRH
2110                    If Val(strTmp04) = 0 Then
2120                      If gblnClosing = False And gblnDeleting = False Then
2130                        blnWarnZeroCost = True
2140                        DoCmd.Hourglass False
2150                        msgResponse = MsgBox("Are you sure you want Principal Cash to be ZERO?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
2160                        If msgResponse <> vbYes Then
2170                          blnSpecCashNo = True
2180                        Else
                              ' ** Let sub know they've been warned!
                              ' ** blnZeroCash     : Local to PCash_Exit() only.
                              ' ** blnWarnZeroCost : Not checked during Commit.
2190                        End If
2200                      End If
2210                    End If
2220                  End If
2230                Case "Liability (-)"
                      ' ** Handled elsewhere.
2240                  If .PCash = 0@ Then
2250  On Error Resume Next
2260                    strTmp04 = .PCash.text
2270  On Error GoTo ERRH
2280                    If Val(strTmp04) = 0 Then
2290                      If gblnClosing = False And gblnDeleting = False Then
2300                        blnWarnZeroCost = True
2310                        DoCmd.Hourglass False
2320                        msgResponse = MsgBox("Are you sure you want Income and Principal Cash to be ZERO?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
2330                        If msgResponse <> vbYes Then
2340                          blnSpecCashNo = True
2350                        Else
                              ' ** Let sub know they've been warned!
                              ' ** blnZeroCash     : Local to PCash_Exit() only.
                              ' ** blnWarnZeroCost : Not checked during Commit.
2360                        End If
2370                      End If
2380                    End If
2390                  End If
2400                End Select  ' ** strThisJType.
2410                If blnSpecCashNo = False Then
2420                  blnContinue = False
2430                  DoCmd.Hourglass True
2440                  DoEvents
2450                  blnToTaxLot = True
2460                  .Parent.JrnlCol_ID = .JrnlCol_ID
2470                  .Parent.ToTaxLot = 1&  ' ** 0 = Nothing; 1 = To Tax Lot; 2 = OK Single; 3 = OK Multi; -4 = Cancel Return; PLUS...
2480                  .Parent.TaxLotFrom = strThis
2490                  intRetVal = OpenLotInfoForm(False, .Name)  ' ** Module Function: modPurchaseSold.
                      ' ** Since the OpenLotInfoForm() function calls the form as a Dialog,
                      ' ** it should return right to here.
                      ' ** Return Values:
                      ' **    0 OK.
                      ' **   -1 Input missing.
                      ' **   -2 No holdings.
                      ' **   -3 Insufficient holdings.
                      ' **   -4 Zero shares.
                      ' **   -9 Data problem.
2500                  If intRetVal <> 0 Then
                        ' ** A non-zero here indicates a problem within the OpenLotInfoForm() function,
                        ' ** and it never went to the form.
2510                    .Parent.ToTaxLot = CLng(intRetVal)  ' ** Just 'cause.
2520                    gblnSetFocus = True
2530                    .Parent.TimerInterval = 250&
2540                  End If
2550                End If  ' ** blnSpecCashNo.
2560              Else
                    ' ** Proceed normally.
2570              End If
2580            Case "Liability (+)"
                  ' ** Proceed normally.
2590            Case Else
                  ' ** Proceed normally.
2600            End Select  ' ** strThisJType.
2610          End Select  ' ** posted.
2620          If blnContinue = True Then
2630            .Controls(strNext).SetFocus
2640          End If
2650        Case vbKeyReturn
2660          blnContinue = True
2670          strNext = JC_Key_Sub_Next(strProc, blnNextRec, blnFromZero)  ' ** Function: Below.
2680          Select Case .posted
              Case True
                ' ** Proceed normally.
2690          Case False
2700            strThisJType = Nz(.journaltype, vbNullString)
2710            Select Case strThis
                Case "shareface"
2720              strTmp01 = "Withdrawn"
2730              strTmp02 = "Withdrawn"
2740              strTmp03 = "Withdrawn"
2750            Case "pcash"
2760              strTmp01 = "Withdrawn"
2770              strTmp02 = "Sold"
2780              strTmp03 = "Liability (-)"
2790            End Select
2800            Select Case strThisJType
                Case strTmp01, strTmp02, strTmp03
2810              If .Cost = 0@ And .shareface <> 0# Then
2820                Select Case strThisJType
                    Case "Withdrawn"
2830                  If .ICash <> 0@ Or .PCash <> 0@ Then
2840                    If gblnClosing = False And gblnDeleting = False Then
2850                      Beep
2860                      DoCmd.Hourglass False
2870                      MsgBox "No cash is allowed for a Withdrawn transaction.", vbInformation + vbOKOnly, "Invalid Entry"
2880                      blnSpecCashNo = True
2890                    End If
2900                  End If
2910                Case "Sold"
2920                  If .ICash = 0@ And .PCash = 0@ Then
2930  On Error Resume Next
2940                    strTmp04 = .PCash.text
2950  On Error GoTo ERRH
2960                    If Val(strTmp04) = 0 Then
2970                      If gblnClosing = False And gblnDeleting = False Then
2980                        blnWarnZeroCost = True
2990                        DoCmd.Hourglass False
3000                        msgResponse = MsgBox("Are you sure you want Income and Principal Cash to be ZERO?" & vbCrLf & vbCrLf & _
                              "As would be the case for the sale of worthless shares.", vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
3010                        If msgResponse <> vbYes Then
3020                          blnSpecCashNo = True
3030                        Else
                              ' ** Let sub know they've been warned!
                              ' ** blnZeroCash     : Local to PCash_Exit() only.
                              ' ** blnWarnZeroCost : Not checked during Commit.
3040                        End If
3050                      End If
3060                    End If
3070                  ElseIf .PCash = 0@ Then
3080  On Error Resume Next
3090                    strTmp04 = .PCash.text
3100  On Error GoTo ERRH
3110                    If Val(strTmp04) = 0 Then
3120                      If gblnClosing = False And gblnDeleting = False Then
3130                        blnWarnZeroCost = True
3140                        DoCmd.Hourglass False
3150                        msgResponse = MsgBox("Are you sure you want Principal Cash to be ZERO?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
3160                        If msgResponse <> vbYes Then
3170                          blnSpecCashNo = True
3180                        Else
                              ' ** Let sub know they've been warned!
                              ' ** blnZeroCash     : Local to PCash_Exit() only.
                              ' ** blnWarnZeroCost : Not checked during Commit.
3190                        End If
3200                      End If
3210                    End If
3220                  End If
3230                Case "Liability (-)"
                      ' ** Handled elsewhere.
3240                  If .PCash = 0@ Then
3250  On Error Resume Next
3260                    strTmp04 = .PCash.text
3270  On Error GoTo ERRH
3280                    If Val(strTmp04) = 0 Then
3290                      If gblnClosing = False And gblnDeleting = False Then
3300                        blnWarnZeroCost = True
3310                        DoCmd.Hourglass False
3320                        msgResponse = MsgBox("Are you sure you want Income and Principal Cash to be ZERO?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
3330                        If msgResponse <> vbYes Then
3340                          blnSpecCashNo = True
3350                        Else
                              ' ** Let sub know they've been warned!
                              ' ** blnZeroCash     : Local to PCash_Exit() only.
                              ' ** blnWarnZeroCost : Not checked during Commit.
3360                        End If
3370                      End If
3380                    End If
3390                  End If
3400                End Select  ' ** strThisJType.
3410                If blnSpecCashNo = False Then
3420                  blnContinue = False
3430                  DoCmd.Hourglass True
3440                  DoEvents
3450                  blnToTaxLot = True
3460                  .Parent.JrnlCol_ID = .JrnlCol_ID
3470                  .Parent.ToTaxLot = 1&  ' ** 0 = Nothing; 1 = To Tax Lot; 2 = OK Single; 3 = OK Multi; -4 = Cancel Return; PLUS...
3480                  .Parent.TaxLotFrom = strThis
3490                  intRetVal = OpenLotInfoForm(False, .Name)  ' ** Module Function: modPurchaseSold.
                      ' ** Since the OpenLotInfoForm() function calls the form as a Dialog,
                      ' ** it should return right to here.
                      ' ** Return Values:
                      ' **    0 OK.
                      ' **   -1 Input missing.
                      ' **   -2 No holdings.
                      ' **   -3 Insufficient holdings.
                      ' **   -4 Zero shares.
                      ' **   -9 Data problem.
3500                  If intRetVal <> 0 Then
                        ' ** A non-zero here indicates a problem within the OpenLotInfoForm() function,
                        ' ** and it never went to the form.
3510                    .Parent.ToTaxLot = CLng(intRetVal)  ' ** Just 'cause.
3520                    gblnSetFocus = True
3530                    .Parent.TimerInterval = 250&
3540                  End If
3550                End If  ' ** blnSpecCashNo.
3560              Else
                    ' ** Proceed normally.
3570              End If
3580            Case "Liability (+)"
                  ' ** Proceed normally.
3590            Case Else
                  ' ** Proceed normally.
3600            End Select  ' ** strThisJType.
3610          End Select  ' ** posted.
3620          If blnContinue = True Then
3630            Select Case .Parent.opgEnterKey
                Case .Parent.opgEnterKey_optRight.OptionValue
3640              .Controls(strNext).SetFocus
3650            Case .Parent.opgEnterKey_optDown.OptionValue
3660              strPageMoveCtl = strThis
3670              lngRecsCur = .RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
3680              If .CurrentRecord < lngRecsCur Then
3690                .MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Columns_Sub.
3700                .shareface.SetFocus
3710              Else
3720                .Controls(strNext).SetFocus
3730              End If
3740            End Select
3750          End If
3760        End Select  ' ** intKeyCode.
3770        If blnWarnZeroCost = True Then
3780          .WarnZeroCash_GetSet False, blnWarnZeroCost  ' ** Form Procedure: frmJournal_Columns_Sub.
3790        End If

3800      Case 4
            ' **************************
            ' ** CommitRec.
            ' **************************
3810        Select Case intKeyCode
            Case vbKeyTab
3820          strNext = JC_Key_Sub_Next(strProc, blnNextRec, blnFromZero)  ' ** Function: Below.
3830          Select Case blnNextRec
              Case True
3840            blnNextRec = False
3850            Select Case .posted
                Case True
3860              lngRecsCur = .RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
3870              If .CurrentRecord < lngRecsCur Then
3880                .MoveRec acCmdRecordsGoToNext  ' ** Form Procedure:  frmJournal_Columns_Sub.
3890                .Controls(strNext).SetFocus
3900              Else
3910                strNext = JC_Key_Par_Next(.Name & "_Exit", .Parent)  ' ** Function: Below.
3920                DoCmd.SelectObject acForm, .Parent.Name, False
3930                .Parent.cmdAdd.SetFocus  '.Parent.Controls(strNext).SetFocus
3940              End If
3950              lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
3960            Case False
3970              strThisJType = Nz(.journaltype, vbNullString)
3980              Select Case strThisJType
                  Case "Cost Adj."
3990                If .Cost <> 0@ And .assetno > 0& Then
                      '.CostAdjRec  ' ** Form Procedure: frmJournal_Columns_Sub.
4000                  JC_Rec_CostAdjRec frmSub  ' ** Module Procedure: modJrnlCol_Recs.
4010                End If
4020              Case Else
4030                CommitRec frmSub, blnNextRec, blnFromZero  ' ** Module Function: modJrnlCol_Recs.
4040              End Select
4050            End Select
4060          Case False
4070  On Error Resume Next
4080            .Controls(strNext).SetFocus
4090  On Error GoTo ERRH
4100          End Select
4110        Case vbKeyReturn
4120          strNext = JC_Key_Sub_Next(strProc, blnNextRec, blnFromZero)  ' ** Function: Below.
4130          lngRecsCur = .RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
4140          Select Case blnNextRec
              Case True
4150            blnNextRec = False
4160            Select Case .posted
                Case True
4170              Select Case .Parent.opgEnterKey
                  Case .Parent.opgEnterKey_optRight.OptionValue
4180                If .CurrentRecord < lngRecsCur Then
4190                  .MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Columns_Sub.
4200                  .Controls(strNext).SetFocus
4210                Else
4220                  strNext = JC_Key_Par_Next(.Name & "_Exit", .Parent)  ' ** Function: Below.
4230                  DoCmd.SelectObject acForm, .Parent.Name, False
4240                  .Parent.cmdAdd.SetFocus  '.Parent.Controls(strNext).SetFocus
4250                End If
4260              Case .Parent.opgEnterKey_optDown.OptionValue
4270                If .CurrentRecord < lngRecsCur Then
4280                  strPageMoveCtl = strThis
4290                  .MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Columns_Sub.
4300                  .Controls(strThis).SetFocus
4310                Else
4320                  strNext = JC_Key_Par_Next(.Name & "_Exit", .Parent)  ' ** Function: Below.
4330                  DoCmd.SelectObject acForm, .Parent.Name, False
4340                  .Parent.cmdAdd.SetFocus  '.Parent.Controls(strNext).SetFocus
4350                End If
4360              End Select
4370              lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
4380            Case False
4390              strThisJType = Nz(.journaltype, vbNullString)
4400              Select Case strThisJType
                  Case "Cost Adj."
4410                If .Cost <> 0@ And .assetno > 0& Then
                      '.CostAdjRec  ' ** Form Procedure: frmJournal_Columns_Sub.
4420                  JC_Rec_CostAdjRec frmSub  ' ** Module Procedure: modJrnlCol_Recs.
4430                End If
4440              Case Else
4450                CommitRec frmSub, blnNextRec, blnFromZero  ' ** Module Function: modJrnlCol_Recs.
4460              End Select
4470            End Select
4480          Case False
4490            Select Case .Parent.opgEnterKey
                Case .Parent.opgEnterKey_optRight.OptionValue
4500  On Error Resume Next
4510              .Controls(strNext).SetFocus
4520  On Error GoTo ERRH
4530            Case .Parent.opgEnterKey_optDown.OptionValue
4540              strPageMoveCtl = strThis
4550              If .CurrentRecord < lngRecsCur Then
4560                .MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Columns_Sub.
4570                .Controls(strThis).SetFocus
4580              Else
4590                .Controls(strNext).SetFocus
4600              End If
4610            End Select
4620          End Select
4630        Case vbKeyPageUp, vbKeyPageDown
4640          strNext = JC_Key_Sub_Next(strProc, blnNextRec, blnFromZero, True, "Last")  ' ** Function: Below.
4650          Select Case strNext
              Case strThis
4660            Select Case .posted
                Case True
4670              strPageMoveCtl = strThis
4680            Case False
4690              strThisJType = Nz(.journaltype, vbNullString)
4700              Select Case strThisJType
                  Case "Cost Adj."
4710                If .Cost <> 0@ And .assetno > 0& Then
                      '.CostAdjRec  ' ** Form Procedure: frmJournal_Columns_Sub.
4720                  JC_Rec_CostAdjRec frmSub  ' ** Module Procedure: modJrnlCol_Recs.
4730                End If
4740              Case Else
4750                CommitRec frmSub, blnNextRec, blnFromZero  ' ** Module Function: modJrnlCol_Recs.
4760              End Select
4770            End Select
4780          Case Else
4790            strPageMoveCtl = strThis
4800          End Select
4810        End Select  ' ** intKeyCode.

4820      End Select  ' ** intAux.
4830    End With  ' ** frmSub.

4840    strRetVal = strPageMoveCtl

EXITP:
4850    JC_Key_Sub = strRetVal
4860    Exit Function

ERRH:
4870    strRetVal = vbNullString
4880    Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
4890    Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
4900    Case Else
4910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4920    End Select
4930    Resume EXITP

End Function

Public Function JC_Key_Sub_Next(strProc As String, blnNextRec As Boolean, blnFromZero As Boolean, Optional varForward As Variant, Optional varSpecial As Variant) As String
' ** Called by:
' **   Oodles!

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Key_Sub_Next"

        Dim frm As Access.Form
        Dim strThisField As String, strThisJType As String
        Dim blnForward As Boolean, blnFound As Boolean
        Dim strSpecial As String
        Dim lngRecsCur As Long, lngAssetNo As Long
        Dim lngPos01 As Long, lngLen As Long
        Dim lngX As Long, lngY As Long, lngE1 As Long, lngE2 As Long
        Dim strRetVal As String

        'MAKE SURE ALL OF THESE PROCEDURES ARE GETTING
        'THE ORIGINAL CALLING PROC, AND NOT SOME INTERMEDIARY!

5010    strRetVal = vbNullString
5020    blnNextRec = False: blnFromZero = False

5030    If lngTabs = 0& Then JC_Key_Sub_Load  ' ** Function: Below.

5040    Set frm = Forms("frmJournal_Columns").frmJournal_Columns_Sub.Form

        ' ** Get the calling field's name.
5050    lngPos01 = 0&
5060    lngLen = Len(strProc)
5070    For lngX = lngLen To 1& Step -1&
5080      If Mid(strProc, lngX, 1) = "_" Then
5090        lngPos01 = lngX
5100        Exit For
5110      End If
5120    Next
5130    strThisField = Left(strProc, (lngPos01 - 1))

        ' ** Get the tabbing direction.
5140    Select Case IsMissing(varForward)
        Case True
5150      blnForward = True
5160    Case False
5170      blnForward = CBool(varForward)
5180    End Select

        ' ** Get First or Last.
5190    Select Case IsMissing(varSpecial)
        Case True
5200      strSpecial = vbNullString
5210    Case False
5220      strSpecial = varSpecial
5230    End Select

5240    lngRecsCur = frm.RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
5250    If lngRecsCur = 0& And strSpecial <> "AddRec" Then
          ' ** OnOpen, this is where blnFromZero is set to True.
          ' ** This is supposed to be then setting it to True in the calling procedure.
          ' ** However, since the first time, it's Form_Timer that calls this, the
          ' ** subform, where AddRec() is, doesn't know it's set!
5260      blnFromZero = True
          ' ** AddRec() will call JC_Key_Sub_Next() again, recursively.
          ' ** Since blnFromZero is only a viable variable in frmJournal_Columns_Sub
          ' ** (where AddRec() is located), should the dummy variables used in
          ' ** frmJournal_Columns, blnD1 and blnD2, be watched more carefully?
          'frm.AddRec blnFromZero  ' ** Form Function: frmJournal_Columns_Sub.
5270      frm.AddRec_Send blnFromZero  ' ** Form Procedure: frmJournal_Columns_Sub.
          ' ** On a completely empty record (save transdate),
          ' ** I want to keep the top buttons active!
5280      DoEvents
5290    End If
5300    If strSpecial = "AddRec" Then strSpecial = vbNullString

5310    blnFromZero = False: lngAssetNo = 0&

        ' ** Get the current JournalType.
5320    Select Case IsNull(frm.journaltype)
        Case True
5330      strThisJType = "Dividend"  ' ** Default to the first JournalType.
5340    Case False
5350      strThisJType = frm.journaltype
5360      If strThisJType = "Received" Then
5370        lngAssetNo = Nz(frm.assetno, 0)
5380      End If
5390    End Select

        ' ** Get the array's start and end.
5400    lngE1 = -1&: lngE2 = -1&
5410    For lngX = 0& To (lngJTypes - 1&)
5420      If arr_varJType(J_JTYPE, lngX) = strThisJType Then
5430        lngE1 = arr_varJType(J_E1, lngX)
5440        lngE2 = arr_varJType(J_E2, lngX)
5450        Exit For
5460      End If
5470    Next

        ' ** Now get the next valid field.
5480    For lngX = lngE1 To lngE2
5490      Select Case strSpecial
          Case "First"
5500        If arr_varTab(T_ACTIVE, lngX) = True Then
5510          strRetVal = arr_varTab(T_CTLNAM, lngX)
5520          Exit For
5530        End If
5540      Case "Last"
5550        If arr_varTab(T_ACTIVE, lngX) = True Then
5560          strRetVal = arr_varTab(T_CTLNAM, lngX)  ' ** It'll just keep updating untill the end of the list.
5570        End If
5580      Case Else
5590        If arr_varTab(T_CTLNAM, lngX) = strThisField Then
5600          blnFound = False
5610          Select Case blnForward
              Case True
                ' ** Forward.
5620            For lngY = (lngX + 1&) To lngE2
5630              If (arr_varTab(T_ACTIVE, lngY) = True) Then
5640                If ((lngAssetNo > 0&) And ((arr_varTab(T_CTLNAM, lngY) = "Recur_Name") Or (arr_varTab(T_CTLNAM, lngY) = "shareface") Or _
                        (arr_varTab(T_CTLNAM, lngY) = "icash") Or (arr_varTab(T_CTLNAM, lngY) = "pcash"))) Then
                      ' ** Nope, keep going.
5650                Else
5660                  blnFound = True
5670                  strRetVal = arr_varTab(T_CTLNAM, lngY)
5680                  Exit For
5690                End If
5700              End If
5710            Next
5720            If blnFound = False Then
                  ' ** Next record, first field.
5730              blnNextRec = True
5740              For lngY = lngE1 To lngE2
5750                If arr_varTab(T_ACTIVE, lngY) = True Then
5760                  blnFound = True
5770                  strRetVal = arr_varTab(T_CTLNAM, lngY)
5780                  Exit For
5790                End If
5800              Next
5810            End If
5820          Case False
                ' ** Backward.
5830            For lngY = (lngX - 1&) To lngE1 Step -1&
5840              If (strThisField = "revcode_ID" And arr_varTab(T_CTLNAM, lngY) = "revcode_DESC_display") Or _
                      (strThisField = "taxcode" And arr_varTab(T_CTLNAM, lngY) = "taxcode_description_display") Or _
                      (strThisField = "Location_ID" And arr_varTab(T_CTLNAM, lngY) = "Loc_Name_display") Or _
                      (strThisField = "assetno" And arr_varTab(T_CTLNAM, lngY) = "assetno_description") Then
                    ' ** Don't send them back to their display counterpart!
5850              Else
5860                If (arr_varTab(T_ACTIVE, lngY) = True) Then
5870                  If ((lngAssetNo > 0&) And ((arr_varTab(T_CTLNAM, lngY) = "Recur_Name") Or (arr_varTab(T_CTLNAM, lngY) = "shareface") Or _
                          (arr_varTab(T_CTLNAM, lngY) = "icash") Or (arr_varTab(T_CTLNAM, lngY) = "pcash"))) Then
                        ' ** Nope, keep going.
5880                  Else
5890                    blnFound = True
5900                    strRetVal = arr_varTab(T_CTLNAM, lngY)
5910                    Exit For
5920                  End If
5930                End If
5940              End If
5950            Next
5960            If blnFound = False Then
                  ' ** Previous record, last field.
5970              blnNextRec = True
5980              For lngY = lngE2 To lngE1 Step -1&
5990                If arr_varTab(T_ACTIVE, lngY) = True Then
6000                  blnFound = True
6010                  strRetVal = arr_varTab(T_CTLNAM, lngY)
6020                  Exit For
6030                End If
6040              Next
6050            End If
6060          End Select
6070        End If  ' ** strThisField.
6080      End Select  ' ** strSpecial.
6090    Next  ' ** lngX.

EXITP:
6100    Set frm = Nothing
6110    JC_Key_Sub_Next = strRetVal
6120    Exit Function

ERRH:
6130    strRetVal = RET_ERR
6140    Select Case ERR.Number
        Case Else
6150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6160    End Select
6170    Resume EXITP

End Function

Public Function JC_Key_Par_Next(strProc As String, frm As Access.Form, Optional varForward As Variant) As String

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Key_Par_Next"

        Dim strThisField As String, strThisJType As String
        Dim blnForward As Boolean, blnFound As Boolean
        Dim lngPos01 As Long, lngLen As Long
        Dim lngX As Long, lngY As Long, lngE1 As Long, lngE2 As Long
        Dim strRetVal As String

6210    strRetVal = vbNullString

6220    If lngTabs2 = 0& Then JC_Key_Par_Load  ' ** Function: Below.

6230    With frm

          ' ** Get the calling control's name.
6240      lngPos01 = 0&
6250      lngLen = Len(strProc)
6260      For lngX = lngLen To 1& Step -1&
6270        If Mid(strProc, lngX, 1) = "_" Then
6280          lngPos01 = lngX
6290          Exit For
6300        End If
6310      Next
6320      strThisField = Left(strProc, (lngPos01 - 1))

          ' ** Get the tabbing direction.
6330      Select Case IsMissing(varForward)
          Case True
6340        blnForward = True
6350      Case False
6360        blnForward = CBool(varForward)
6370      End Select

          ' ** Get the current (pseudo) JournalType.
6380      strThisJType = "{None}"

          ' ** Get the array's start and end.
6390      lngE1 = -1&: lngE2 = -1&
6400      For lngX = 0& To (lngJTypes - 1&)
6410        If arr_varJType(J_JTYPE, lngX) = strThisJType Then
6420          lngE1 = arr_varJType(J_E1, lngX)
6430          lngE2 = arr_varJType(J_E2, lngX)
6440          Exit For
6450        End If
6460      Next

          ' ** Get the current Enabled status.
6470      For lngX = lngE1 To lngE2
6480        If arr_varTab2(T2_ACTIVE, lngX) = True Then
              ' ** Only check those controls that were active to begin with.
6490          arr_varTab2(T2_ACTNOW, lngX) = .Controls(arr_varTab2(T2_CTLNAM, lngX)).Enabled
6500        End If
6510      Next

          ' ** Now get the next valid field.
6520      For lngX = lngE1 To lngE2
6530        If arr_varTab2(T2_CTLNAM, lngX) = strThisField Then
6540          blnFound = False
6550          Select Case blnForward
              Case True
6560            For lngY = (lngX + 1&) To lngE2
6570              If arr_varTab2(T2_ACTNOW, lngY) = True Then
6580                blnFound = True
6590                strRetVal = arr_varTab2(T2_CTLNAM, lngY)
6600                Exit For
6610              End If
6620            Next
6630            If blnFound = False Then
                  ' ** First field.
6640              For lngY = lngE1 To lngE2
6650                If arr_varTab2(T2_ACTNOW, lngY) = True Then
6660                  blnFound = True
6670                  strRetVal = arr_varTab2(T2_CTLNAM, lngY)
6680                  Exit For
6690                End If
6700              Next
6710            End If
6720          Case False
6730            For lngY = (lngX - 1&) To lngE1 Step -1&
6740              If arr_varTab2(T2_ACTNOW, lngY) = True Then
6750                blnFound = True
6760                strRetVal = arr_varTab2(T2_CTLNAM, lngY)
6770                Exit For
6780              End If
6790            Next
6800            If blnFound = False Then
                  ' ** Last field.
6810              For lngY = lngE2 To lngE1 Step -1&
6820                If arr_varTab2(T2_ACTNOW, lngY) = True Then
6830                  blnFound = True
6840                  strRetVal = arr_varTab2(T2_CTLNAM, lngY)
6850                  Exit For
6860                End If
6870              Next
6880            End If
6890          End Select
6900        End If
6910      Next  ' ** lngX.

6920    End With  ' ** Me.

EXITP:
6930    JC_Key_Par_Next = strRetVal
6940    Exit Function

ERRH:
6950    strRetVal = RET_ERR
6960    Select Case ERR.Number
        Case Else
6970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6980    End Select
6990    Resume EXITP

End Function

Private Function JC_Key_Sub_Load() As Variant
' ** Called by:
' **   JC_Key_Sub_Next(), Above

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Key_Sub_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim arr_varRetVal() As Variant
        Dim lngX As Long, lngY As Long, lngZ As Long
        Dim blnContinue As Boolean

        ' ** Array: arr_varRetVal().
        Const RV_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const RV_ERR    As Integer = 0
        Const RV_TABS   As Integer = 1
        Const RV_JTYPES As Integer = 2

7010    blnContinue = True

7020    ReDim arr_varRetVal(RV_ELEMS, 0)
7030    arr_varRetVal(RV_ERR, 0) = vbNullString

7040    Set dbs = CurrentDb
7050    With dbs

          ' ** tblJournal_Field, sorted by JournalType_Order, ctlspec_tabindex.
7060      Set qdf = .QueryDefs("qryJournal_Columns_11")
7070      Set rst = qdf.OpenRecordset
7080      With rst
7090        If .BOF = True And .EOF = True Then
              ' ** Something's horribly wrong!
7100          blnContinue = False
7110          MsgBox "Data necessary for Journal input is missing.", vbCritical + vbOKOnly, "Data Not Found"
7120        Else
7130          .MoveLast
7140          lngTabs = .RecordCount
7150          .MoveFirst
7160          arr_varTab = .GetRows(lngTabs)
              ' *****************************************************
              ' ** Array: arr_varTab()
              ' **
              ' **   Field  Element  Name                Constant
              ' **   =====  =======  ==================  ==========
              ' **     1       0     JournalType         T_JTYPE
              ' **     2       1     ctl_name            T_CTLNAM
              ' **     3       2     ctltype_type        T_CTLTYP
              ' **     4       3     jfld_active         T_ACTIVE
              ' **
              ' *****************************************************
7170        End If
7180        .Close
7190      End With  ' ** rst.

          ' ** If Income/Expense Tracking isn't on, turn revcode_ID off.
7200      If gblnRevenueExpenseTracking = False Then
7210        For lngX = 0& To (lngTabs - 1&)
7220          If arr_varTab(T_CTLNAM, lngX) = "revcode_DESC_display" Or arr_varTab(T_CTLNAM, lngX) = "revcode_ID" Then
7230            arr_varTab(T_ACTIVE, lngX) = CBool(False)
7240          End If
7250        Next
7260      End If

          ' ** If Income Tax Tracking isn't on, turn taxcode off.
7270      If gblnIncomeTaxCoding = False Then
7280        For lngX = 0& To (lngTabs - 1&)
7290          If arr_varTab(T_CTLNAM, lngX) = "taxcode_description_display" Or arr_varTab(T_CTLNAM, lngX) = "taxcode" Then
7300            arr_varTab(T_ACTIVE, lngX) = CBool(False)
7310          End If
7320        Next
7330      End If

          ' ** tblJournalType, with ElementStart, ElementEnd.
7340      Set qdf = .QueryDefs("qryJournal_Columns_12")
7350      Set rst = qdf.OpenRecordset
7360      With rst
7370        If .BOF = True And .EOF = True Then
              ' ** Something's horribly wrong!
7380          blnContinue = False
7390          MsgBox "Data necessary for Journal input is missing.", vbCritical + vbOKOnly, "Data Not Found"
7400        Else
7410          .MoveLast
7420          lngJTypes = .RecordCount
7430          .MoveFirst
7440          arr_varJType = .GetRows(lngJTypes)
              ' ******************************************************
              ' ** Array: arr_varJType()
              ' **
              ' **   Field  Element  Name                 Constant
              ' **   =====  =======  ===================  ==========
              ' **     1       0     JournalType          J_JTYPE
              ' **     2       1     JournalType_Order    J_ORD
              ' **     3       2     ElementStart         J_E1
              ' **     4       3     ElementEnd           J_E2
              ' **
              ' ******************************************************
7450        End If
7460        .Close
7470      End With  ' ** rst.

7480      .Close
7490    End With  ' ** dbs.

        ' ** Add starting and ending array elements for each JournalType.
7500    For lngX = 0& To (lngJTypes - 1&)
7510      For lngY = 0& To (lngTabs - 1&)
7520        If arr_varTab(T_JTYPE, lngY) = arr_varJType(J_JTYPE, lngX) Then
7530          arr_varJType(J_E1, lngX) = lngY
7540          arr_varJType(J_E2, lngX) = (lngTabs - 1&)  ' ** Default to last element.
7550          For lngZ = lngY To (lngTabs - 1&)
7560            If arr_varTab(T_JTYPE, lngZ) <> arr_varJType(J_JTYPE, lngX) Then
7570              arr_varJType(J_E2, lngX) = (lngZ - 1&)
7580              Exit For
7590            End If
7600          Next
7610          Exit For
7620        End If
7630      Next
7640    Next

7650    If blnContinue = True Then
7660      arr_varRetVal(RV_TABS, 0) = arr_varTab
7670      arr_varRetVal(RV_JTYPES, 0) = arr_varJType
7680    Else
7690      arr_varRetVal(RV_ERR, 0) = RET_ERR
7700    End If

        'For lngX = 0& To (lngJTypes - 1&)
        '  Debug.Print "'" & Left(arr_varJType(J_JTYPE, lngX) & Space(10), 10) & "  E1: " & CStr(arr_varJType(J_E1, lngX)) & "  E2: " & CStr(arr_varJType(J_E2, lngX))
        'Next

EXITP:
7710    Set rst = Nothing
7720    Set qdf = Nothing
7730    Set dbs = Nothing
7740    JC_Key_Sub_Load = arr_varRetVal
7750    Exit Function

ERRH:
7760    arr_varRetVal(RV_ERR, 0) = RET_ERR
7770    Select Case ERR.Number
        Case Else
7780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7790    End Select
7800    Resume EXITP

End Function

Private Function JC_Key_Par_Load() As Boolean
' ** Called by:
' **   JC_Key_Par_Next(), Above

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Key_Par_Load"

        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

7910    blnRetVal = True

7920    lngTabs2 = lngTabs
7930    lngE = lngTabs2 - 1&
7940    ReDim arr_varTab2(T2_ELEMS, lngE)  ' ** With the extra element.
        ' ******************************************************
        ' ** Array: arr_varTab2()
        ' **
        ' **   Field  Element  Name                Constant
        ' **   =====  =======  ==================  ===========
        ' **     1       0     JournalType         T2_JTYPE
        ' **     2       1     ctl_name            T2_CTLNAM
        ' **     3       2     ctltype_type        T2_CTLTYP
        ' **     4       3     jfld_active         T2_ACTIVE
        ' **     5       4     active now          T2_ACTNOW
        ' **
        ' ******************************************************
7950    For lngX = 0& To (lngE)
7960      arr_varTab2(T2_JTYPE, lngX) = arr_varTab(T_JTYPE, lngX)
7970      arr_varTab2(T2_CTLNAM, lngX) = arr_varTab(T_CTLNAM, lngX)
7980      arr_varTab2(T2_CTLTYP, lngX) = arr_varTab(T_CTLTYP, lngX)
7990      arr_varTab2(T2_ACTIVE, lngX) = arr_varTab(T_ACTIVE, lngX)
8000      arr_varTab2(T2_ACTNOW, lngX) = arr_varTab(T_ACTIVE, lngX)
8010    Next

        'For lngX = 0& To (lngTabs2 - 1&)
        '  If arr_varTab2(T2_JTYPE, lngX) = "{None}" Then
        '    Debug.Print "'" & arr_varTab2(T2_CTLNAM, lngX)
        '  End If
        'Next

EXITP:
8020    JC_Key_Par_Load = blnRetVal
8030    Exit Function

ERRH:
8040    blnRetVal = False
8050    Select Case ERR.Number
        Case Else
8060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8070    End Select
8080    Resume EXITP

End Function

Public Function JC_Key_JType_Set(strThisJType As String, frmSub As Access.Form) As Boolean
' ** The locking done here is just the broad categories covering which JournalTypes require which fields.
' ** Called by:
' **   frmJournal_Columns_Sub:
' **     Form_Timer()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Key_JType_Set"

        Dim lngE1 As Long, lngE2 As Long, lngENext As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

8110    blnRetVal = True

        ' ** Get the arr_varTab() array's start and end.
8120    lngE1 = -1&: lngE2 = -1&
8130    For lngX = 0& To (lngJTypes - 1&)
8140      If arr_varJType(J_JTYPE, lngX) = strThisJType Then
8150        lngE1 = arr_varJType(J_E1, lngX)
8160        lngE2 = arr_varJType(J_E2, lngX)
8170        Exit For
8180      End If
8190    Next

        ' *****************************************************
        ' ** Array: arr_varTab()
        ' **
        ' **   Field  Element  Name                Constant
        ' **   =====  =======  ==================  ==========
        ' **     1       0     JournalType         T_JTYPE
        ' **     2       1     ctl_name            T_CTLNAM
        ' **     3       2     ctltype_type        T_CTLTYP
        ' **     4       3     jfld_active         T_ACTIVE
        ' **
        ' *****************************************************

8200    If IsEmpty(arr_varTab) = False Then  ' ** It sometimes gets here as the form is closing.

8210      For lngX = lngE1 To lngE2
8220        If arr_varTab(T_CTLNAM, lngX) = "journaltype" Then
8230          lngENext = lngX + 1&  ' ** Only look at fields to the right.
8240          Exit For
8250        End If
8260      Next

          ' ** Lock or unlock as appropriate.
8270      For lngX = lngENext To lngE2
8280        If arr_varTab(T_CTLTYP, lngX) <> acCommandButton Then
8290          frmSub.Controls(arr_varTab(T_CTLNAM, lngX)).Locked = Not (arr_varTab(T_ACTIVE, lngX))  ' ** If ACTIVE then NOT LOCKED.
8300        End If
8310      Next

8320    End If

EXITP:
8330    JC_Key_JType_Set = blnRetVal
8340    Exit Function

ERRH:
8350    blnRetVal = False
8360    Select Case ERR.Number
        Case Else
8370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8380    End Select
8390    Resume EXITP

End Function

Public Sub JC_Key_Clear()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Key_Clear"

8410    lngJTypes = 0&
8420    arr_varJType = Empty
8430    lngTabs = 0&
8440    arr_varTab = Empty
8450    lngTabs2 = 0&
8460    ReDim arr_varTab2(T2_ELEMS, 0)

EXITP:
8470    Exit Sub

ERRH:
8480    Select Case ERR.Number
        Case Else
8490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8500    End Select
8510    Resume EXITP

End Sub

Public Function JC_Key_Frm(KeyCode As Integer, intMode As Integer, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form) As Integer

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Key_Frm"

        Dim intRetVal As Integer, lngRetVal As Long

8610    intRetVal = KeyCode

8620    Select Case intMode
        Case 1  ' ** Plain keys.
8630      Select Case intRetVal
            ' ** No vbKeyEscape on this form.
          Case vbKeyF5
8640        With frm
8650          intRetVal = 0
8660          .frmJournal_Columns_Sub.Form.RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.
8670        End With
8680      Case vbKeyF7
8690        With frm
8700          intRetVal = 0
8710          If IsNull(.frmJournal_Columns_Sub.Form.JrnlCol_ID) = False Then
8720            If .frmJournal_Columns_Sub.Form.transdate.Locked = False Then
8730              .frmJournal_Columns_Sub.SetFocus
8740              .frmJournal_Columns_Sub.Form.cmdCalendar1.SetFocus
8750              .frmJournal_Columns_Sub.Form.cmdCalendar1_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
8760            End If
8770          End If
8780        End With
8790      Case vbKeyF8
8800        With frm
8810          intRetVal = 0
8820          If IsNull(.frmJournal_Columns_Sub.Form.JrnlCol_ID) = False Then
8830            If .frmJournal_Columns_Sub.Form.assetdate_display.Locked = False Then
8840              .frmJournal_Columns_Sub.SetFocus
8850              .frmJournal_Columns_Sub.Form.cmdCalendar2.SetFocus
8860              .frmJournal_Columns_Sub.Form.cmdCalendar2_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
8870            End If
8880          End If
8890        End With
8900      End Select
8910    Case 2  ' ** Alt keys.
8920      Select Case intRetVal
          Case vbKeyD
8930        With frm
8940          intRetVal = 0
8950          .frmJournal_Columns_Sub.SetFocus
8960          lngRetVal = fSetScrollBarPosHZ(.frmJournal_Columns_Sub.Form, 1&)  ' ** Module Function: modScrollBarFuncs.
8970          If .frmJournal_Columns_Sub.Form.transdate.Locked = False Then
8980            .frmJournal_Columns_Sub.Form.transdate.SetFocus
8990          End If
9000        End With
9010      Case vbKeyH
9020        With frm
9030          intRetVal = 0
9040          .frmJournal_Columns_Sub.SetFocus
9050          If .frmJournal_Columns_Sub.Form.ICash.Locked = False Then
9060            .frmJournal_Columns_Sub.Form.ICash.SetFocus
9070          Else
9080            .frmJournal_Columns_Sub.Form.FocusHolder.SetFocus
9090          End If
9100        End With
9110      Case vbKeyJ
9120        With frm
9130          intRetVal = 0
9140          .frmJournal_Columns_Sub.SetFocus
9150          If .frmJournal_Columns_Sub.Form.journaltype.Locked = False Then
9160            .frmJournal_Columns_Sub.Form.journaltype.SetFocus
9170          Else
9180            .frmJournal_Columns_Sub.Form.FocusHolder.SetFocus
9190          End If
9200        End With
9210      Case vbKeyK
9220        With frm
9230          intRetVal = 0
9240          .frmJournal_Columns_Sub.SetFocus
9250          If .frmJournal_Columns_Sub.Form.PrintCheck.Locked = False Then
9260            .frmJournal_Columns_Sub.Form.PrintCheck.SetFocus
9270          Else
9280            lngRetVal = fSetScrollBarPosHZ(.frmJournal_Columns_Sub.Form, 999&)  ' ** Module Function: modScrollBarFuncs.
9290          End If
9300        End With
9310      Case vbKeyN
9320        With frm
9330          intRetVal = 0
9340          .frmJournal_Columns_Sub.SetFocus
9350          If .frmJournal_Columns_Sub.Form.accountno.Locked = False Then
9360            .frmJournal_Columns_Sub.Form.accountno.SetFocus
9370          Else
9380            .frmJournal_Columns_Sub.Form.FocusHolder.SetFocus
9390          End If
9400        End With
9410      Case vbKeyO
9420        With frm
9430          intRetVal = 0
9440          If .frmJournal_Columns_Sub.Form.posted.Locked = False Then
9450            lngRetVal = fSetScrollBarPosHZ(.frmJournal_Columns_Sub.Form, 1&)  ' ** Module Function: modScrollBarFuncs.
9460            .frmJournal_Columns_Sub.Form.posted.SetFocus
9470            If .frmJournal_Columns_Sub.Form.posted.Locked = False Then
9480              .frmJournal_Columns_Sub.Form.posted = Not .frmJournal_Columns_Sub.Form.posted
9490            End If
9500            .frmJournal_Columns_Sub.Form.posted_AfterUpdate  ' ** Form Procedure: frmJournal_Columns_Sub.
9510          End If
9520        End With
9530      Case vbKeyR
9540        With frm
9550          intRetVal = 0
9560          .frmJournal_Columns_Sub.SetFocus
9570          If .frmJournal_Columns_Sub.Form.Recur_Name.Locked = False Then
9580            .frmJournal_Columns_Sub.Form.Recur_Name.SetFocus
9590          Else
9600            .frmJournal_Columns_Sub.Form.FocusHolder.SetFocus
9610          End If
9620        End With
9630      Case vbKeyS
9640        With frm
9650          intRetVal = 0
9660          .frmJournal_Columns_Sub.SetFocus
9670          If .frmJournal_Columns_Sub.Form.assetno.Locked = False Then
9680            .frmJournal_Columns_Sub.Form.assetno.SetFocus
9690          Else
9700            .frmJournal_Columns_Sub.Form.FocusHolder.SetFocus
9710          End If
9720        End With
9730      Case vbKeyT
9740        With frm
9750          intRetVal = 0
9760          .frmJournal_Columns_Sub.SetFocus
9770        End With
9780      Case vbKeyX
9790        With frm
9800          intRetVal = 0
9810          .cmdClose.SetFocus
9820          .cmdClose_Click  ' ** Form Procedure: frmJournal_Columns.
9830        End With
9840      End Select
9850    Case 3  ' ** Ctrl keys.
9860      Select Case intRetVal
          Case vbKeyA
9870        With frm
9880          intRetVal = 0
9890          If .cmdAdd.Enabled = True Then
9900            .cmdAdd.SetFocus
9910            .cmdAdd_Click  ' ** Form Procedure: frmJournal_Columns.
9920          End If
9930        End With
9940      Case vbKeyD
9950        With frm
9960          intRetVal = 0
9970          If .cmdDelete.Enabled = True Then
9980            .cmdDelete.SetFocus
9990            .cmdDelete_Click  ' ** Form Procedure: frmJournal_Columns.
10000         End If
10010       End With
10020     Case vbKeyE
10030       With frm
10040         intRetVal = 0
10050         If .cmdEdit.Enabled = True Then
10060           .cmdEdit.SetFocus
10070           .cmdEdit_Click  ' ** Form Procedure: frmJournal_Columns.
10080         End If
10090       End With
10100     Case vbKeyH
10110       With frm
10120         intRetVal = 0
10130         If .cmdRefresh.Enabled = True Then
10140           .cmdRefresh.SetFocus
10150           .cmdRefresh_Click  ' ** Form Procedure: frmJournal_Columns.
10160         End If
10170       End With
10180     Case vbKeyL
10190       With frm
10200         intRetVal = 0
10210         lngRecsCur = .frmJournal_Columns_Sub.Form.RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
10220         If lngRecsCur > 0& Then
                ' ** But don't SetFocus.
10230           .cmdScrollLeft_Click  ' ** Form Procedure: frmJournal_Columns.
10240         End If
10250       End With
10260     Case vbKeyM
10270       With frm
10280         intRetVal = 0
10290         If .cmdMemoReveal.Visible = True And .cmdMemoReveal.Enabled = True Then
10300           .cmdMemoReveal.SetFocus
10310           .cmdMemoReveal_Click  ' ** Form Procedure: frmJournal_Columns.
10320         Else
10330           Beep
10340         End If
10350       End With
10360     Case vbKeyR
10370       With frm
10380         intRetVal = 0
10390         lngRecsCur = .frmJournal_Columns_Sub.Form.RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
10400         If lngRecsCur > 0& Then
                ' ** But don't SetFocus.
10410           .cmdScrollRight_Click  ' ** Form Procedure: frmJournal_Columns.
10420         End If
10430       End With
10440     Case vbKeyS
10450       With frm
10460         intRetVal = 0
10470         .frmJournal_Columns_Sub.Form.Refresh
10480         .frmJournal_Columns_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
10490       End With
10500     Case vbKeyT
10510       With frm
10520         intRetVal = 0
10530         .cmdSwitch.SetFocus
10540         DoEvents
10550         ForcePause 1  ' ** Module Function: modCodeUtilities.
10560         .cmdSwitch_Click  ' ** Form Procedure: frmJournal_Columns.
10570       End With
10580     End Select
10590   Case 4  ' ** Ctrl-Shift keys.
10600     Select Case intRetVal
          Case vbKeyA
10610       With frm
10620         intRetVal = 0
10630         If .cmdUncomComAll.Enabled = True And .cmdUncomComAll.Visible = True Then
10640           .cmdUncomComAll.SetFocus
10650           .cmdUncomComAll_Click  ' ** Form Procedure: frmJournal_Columns.
10660         End If
10670       End With
10680     Case vbKeyD
10690       With frm
10700         intRetVal = 0
10710         If .cmdUncomDelAll.Enabled = True And .cmdUncomDelAll.Visible = True Then
10720           .cmdUncomDelAll.SetFocus
10730           .cmdUncomDelAll_Click  ' ** Form Procedure: frmJournal_Columns.
10740         End If
10750       End With
10760     Case vbKeyF
10770       With frm
10780         intRetVal = 0
10790         .FocusHolder.SetFocus
10800       End With
10810     End Select
10820   Case 5  ' ** Alt-Shift keys.
10830     Select Case intRetVal
          Case vbKeyA
10840       With frm
10850         intRetVal = 0
10860         .opgFilter.SetFocus
10870         .opgFilter = .opgFilter_optAll.OptionValue
10880         .opgFilter_AfterUpdate  ' ** Form Procedure: frmJournal_Columns.
10890       End With
10900     Case vbKeyC
10910       With frm
10920         intRetVal = 0
10930         .opgFilter.SetFocus
10940         .opgFilter = .opgFilter_optCommitted.OptionValue
10950         .opgFilter_AfterUpdate  ' ** Form Procedure: frmJournal_Columns.
10960       End With
10970     Case vbKeyD
10980       With frm
10990         intRetVal = 0
11000         If .cmdSpecPurp_Div_Map.Enabled = True Then
11010           .cmdSpecPurp_Div_Map.SetFocus
11020           .cmdSpecPurp_Div_Map_Click  ' ** Form Procedure: frmJournal_Columns.
11030         End If
11040       End With
11050     Case vbKeyG
11060       With frm
11070         intRetVal = 0
11080         If .cmdSpecPurp_Misc_MapSTCGL.Enabled = True Then
11090           .cmdSpecPurp_Misc_MapSTCGL.SetFocus
11100           .cmdSpecPurp_Misc_MapSTCGL_Click  ' ** Form Procedure: frmJournal_Columns.
11110         End If
11120       End With
11130     Case vbKeyI
11140       With frm
11150         intRetVal = 0
11160         If .cmdSpecPurp_Sold_PaidTotal.Enabled = True Then
11170           .cmdSpecPurp_Sold_PaidTotal.SetFocus
11180           .cmdSpecPurp_Sold_PaidTotal_Click  ' ** Form Procedure: frmJournal_Columns.
11190         End If
11200       End With
11210     Case vbKeyL
11220       With frm
11230         intRetVal = 0
11240         If .cmdSpecPurp_Misc_MapLTCG.Enabled = True Then
11250           .cmdSpecPurp_Misc_MapLTCG.SetFocus
11260           .cmdSpecPurp_Misc_MapLTCG_Click  ' ** Form Procedure: frmJournal_Columns.
11270         End If
11280       End With
11290     Case vbKeyP
11300       With frm
11310         intRetVal = 0
11320         If .cmdSpecPurp_Int_Map.Enabled = True Then
11330           .cmdSpecPurp_Int_Map.SetFocus
11340           .cmdSpecPurp_Int_Map_Click  ' ** Form Procedure: frmJournal_Columns.
11350         End If
11360       End With
11370     Case vbKeyS
11380       With frm
11390         intRetVal = 0
11400         If .cmdSpecPurp_Purch_MapSplit.Enabled = True Then
11410           .cmdSpecPurp_Purch_MapSplit.SetFocus
11420           .cmdSpecPurp_Purch_MapSplit_Click  ' ** Form Procedure: frmJournal_Columns.
11430         End If
11440       End With
11450     Case vbKeyT
11460       With frm
11470         intRetVal = 0
11480         If .cmdSpecPurp_Misc_MapLTCL.Enabled = True Then
11490           .cmdSpecPurp_Misc_MapLTCL.SetFocus
11500           .cmdSpecPurp_Misc_MapLTCL_Click  ' ** Form Procedure: frmJournal_Columns.
11510         End If
11520       End With
11530     Case vbKeyU
11540       With frm
11550         intRetVal = 0
11560         .opgFilter.SetFocus
11570         .opgFilter = .opgFilter_optUncommitted.OptionValue
11580         .opgFilter_AfterUpdate  ' ** Form Procedure: frmJournal_Columns.
11590       End With
11600     End Select
11610   End Select

EXITP:
11620   JC_Key_Frm = intRetVal
11630   Exit Function

ERRH:
11640   intRetVal = 0
11650   THAT_PROC = THIS_PROC
11660   That_Erl = Erl: That_Desc = ERR.description
11670   frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns.
        'Case 2046  ' ** The command or action isn't available now (first or last record).
        'Case 2110  ' ** Access can't move the focus to the control '|'.
11680   Resume EXITP

End Function

Public Function JC_Key_SubFrm(KeyCode As Integer, Shift As Integer, strSaveMoveCtl As String, blnF4Invoked As Boolean, blnF4InvokedMouse As Boolean, strF4LastControl As String, blnNextRec As Boolean, blnFromZero As Boolean, frmSub As Access.Form) As Integer

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Key_SubFrm"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strControl As String, strNext As String
        Dim intRetVal As Integer, lngRetVal As Long

11710   intRetVal = KeyCode
11720   strPageMoveCtl = vbNullString

        ' ** Use bit masks to determine which key was pressed.
11730   intShiftDown = (Shift And acShiftMask) > 0
11740   intAltDown = (Shift And acAltMask) > 0
11750   intCtrlDown = (Shift And acCtrlMask) > 0

11760 On Error Resume Next
11770   strControl = Screen.ActiveControl.Name
11780 On Error GoTo ERRH
11790   If strControl = "assetno_description" Then strControl = "assetno"
11800   If strControl = "Loc_Name_display" Then strControl = "Location_ID"
11810   If strControl = "revcode_DESC_display" Then strControl = "revcode_ID"
11820   If strControl = "taxcode_description_display" Then strControl = "taxcode"

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Recalc:           F5 {RecsTotUpdate}
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   TransDate:        D {transdate}
        ' **   Income Cash:      H {icash}
        ' **   Journal Type:     J {journaltype}
        ' **   Print Check:      K {PrintCheck}
        ' **   Account Number:   N {accountno}
        ' **   Commit:           O {posted}
        ' **   Recurring Item:   R {Recur_Name}
        ' **   Asset:            S {assetno}
        ' **   Exit:             X {cmdClose on frmJournal_Columns}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Add:              A {cmdAdd on frmJournal_Columns}
        ' **   Delete:           D {cmdDelete on frmJournal_Columns}
        ' **   Edit:             E {cmdEdit on frmJournal_Columns}
        ' **   Refresh:          H {cmdRefresh on frmJournal_Columns}
        ' **   Scroll Left:      L {cmdScrollLeft on frmJournal_Columns}
        ' **   Check Memo:       M {cmdMemoReveal on frmJournal_Columns}
        ' **   Scroll Right:     R {cmdScrollRight on frmJournal_Columns}
        ' **   Save:             S {cmdSave}
        ' **   Switch:           T {cmdSwitch on frmJournal_Columns}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Commit All:       A {cmdUncomComAll on frmJournal_Columns}
        ' **   Delete All:       D {cmdUncomDelAll on frmJournal_Columns}

        ' ** Shortcut Alt-Shift keys to other forms and functionality:
        ' **   Show All:         A {opgFilter_optAll on frmJournal_Columns}
        ' **   Show Committed:   C {opgFilter_optCommitted on frmJournal_Columns}
        ' **   Map Div:          D {cmdSpecPurp_Div_Map on frmJournal_Columns}
        ' **   Map STCG/L:       G {cmdSpecPurp_Misc_MapSTCGL on frmJournal_Columns}
        ' **   Paid Total:       I {cmdSpecPurp_Sold_PaidTotal on frmJournal_Columns}
        ' **   Map LTCG:         L {cmdSpecPurp_Misc_MapLTCG on frmJournal_Columns}
        ' **   Map Int:          P {cmdSpecPurp_Int_Map on frmJournal_Columns}
        ' **   Map Split:        S {cmdSpecPurp_Purch_MapSplit on frmJournal_Columns}
        ' **   Map LTCL:         T {cmdSpecPurp_Misc_MapLTCL on frmJournal_Columns}
        ' **   Show Uncommitted: U {opgFilter_optUncommitted on frmJournal_Columns}

        ' ** Plain keys.
11830   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11840     Select Case intRetVal
          Case vbKeyF4
11850       blnF4Invoked = True
11860     Case vbKeyF5
11870       With frmSub
11880         intRetVal = 0
11890         .RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.
11900       End With
11910     Case vbKeyF7
11920       With frmSub
11930         intRetVal = 0
11940         .cmdCalendar1.SetFocus
11950         .cmdCalendar1_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
11960       End With
11970     Case vbKeyF8
11980       With frmSub
11990         intRetVal = 0
12000         .cmdCalendar2.SetFocus
12010         .cmdCalendar2_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
12020       End With
12030     Case vbKeyUp
12040       With frmSub
12050         If strControl <> vbNullString Then
12060           If .Controls(strControl).ControlType = acComboBox And (blnF4Invoked = True Or blnF4InvokedMouse = True) Then
                  ' ** Let it scroll through the dropdown, normally.
12070           Else
12080             blnF4Invoked = False: blnF4InvokedMouse = False  ' ** If it's not on a combo box.
12090           End If
12100         Else
12110           blnF4Invoked = False: blnF4InvokedMouse = False  ' ** If it's not on a control.
12120         End If
12130         If blnF4Invoked = False And blnF4InvokedMouse = False Then
12140           intRetVal = 0
12150           If .CurrentRecord > 1& Then
                  ' ** Focus handled by Form_Current().
12160             strPageMoveCtl = strControl
12170             .MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmJournal_Columns_Sub.
12180           Else
12190             strPageMoveCtl = vbNullString
12200           End If
12210         End If
12220       End With
12230     Case vbKeyDown
12240       With frmSub
12250         If strControl <> vbNullString Then
12260           If .Controls(strControl).ControlType = acComboBox And (blnF4Invoked = True Or blnF4InvokedMouse = True) Then
                  ' ** Let it scroll through the dropdown, normally.
12270           Else
12280             blnF4Invoked = False: blnF4InvokedMouse = False  ' ** If it's not on a combo box.
12290           End If
12300         Else
12310           blnF4Invoked = False: blnF4InvokedMouse = False  ' ** If it's not on a control.
12320         End If
12330         If blnF4Invoked = False Then
12340           intRetVal = 0
12350           lngRecsCur = .RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
12360           If .CurrentRecord < lngRecsCur Then
                  ' ** Focus handled by Form_Current().
12370             strPageMoveCtl = strControl
12380             .MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Columns_Sub.
12390           Else
12400             strPageMoveCtl = vbNullString
12410           End If
12420         End If
12430       End With
12440     Case vbKeyDelete
12450       If strControl = vbNullString Then
              ' ** If this shows up empty, it should mean
              ' ** they've selected the whole record for deletion.
12460         With frmSub
12470           intRetVal = 0
12480           .Parent.cmdDelete_Click  ' ** Form Procedure: frmJournal_Columns.
12490         End With
12500       Else
              ' ** Let it proceed normally.
12510       End If
12520     Case Else
12530       If strControl <> strF4LastControl Then
12540         blnF4Invoked = False: blnF4InvokedMouse = False  ' ** If it's not an Up or Down arrow.
12550       End If
12560     End Select
12570   End If

        ' ** Alt keys.
12580   If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
12590     blnF4Invoked = False: blnF4InvokedMouse = False
12600     Select Case intRetVal
          Case vbKeyD
12610       With frmSub
12620         intRetVal = 0
12630         lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
12640         If .transdate.Locked = False Then
12650           .transdate.SetFocus
12660         Else
12670           .FocusHolder.SetFocus
12680         End If
12690       End With
12700     Case vbKeyH
12710       With frmSub
12720         intRetVal = 0
12730         If .ICash.Locked = False Then
12740           .ICash.SetFocus
12750         Else
12760           .FocusHolder.SetFocus
12770         End If
12780       End With
12790     Case vbKeyJ
12800       With frmSub
12810         intRetVal = 0
12820         If .journaltype.Locked = False Then
12830           .journaltype.SetFocus
12840         Else
12850           .FocusHolder.SetFocus
12860         End If
12870       End With
12880     Case vbKeyK
12890       With frmSub
12900         intRetVal = 0
12910         If .PrintCheck.Locked = False Then
12920           .PrintCheck.SetFocus
12930         Else
12940           lngRetVal = fSetScrollBarPosHZ(frmSub, 999&)  ' ** Module Function: modScrollBarFuncs.
12950         End If
12960       End With
12970     Case vbKeyN
12980       With frmSub
12990         intRetVal = 0
13000         If .accountno.Locked = False Then
13010           .accountno.SetFocus
13020         Else
13030           .FocusHolder.SetFocus
13040         End If
13050       End With
13060     Case vbKeyO
13070       With frmSub
13080         intRetVal = 0
13090         If .posted.Locked = False Then
13100           lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
13110           .posted.SetFocus
13120           If .posted.Locked = False Then
13130             .posted = Not .posted
13140           End If
13150           .posted_AfterUpdate  ' ** Form Procedure: frmJournal_Columns_Sub.
13160         End If
13170       End With
13180     Case vbKeyR
13190       With frmSub
13200         intRetVal = 0
13210         If .Recur_Name.Locked = False Then
13220           .Recur_Name.SetFocus
13230         Else
13240           .FocusHolder.SetFocus
13250         End If
13260       End With
13270     Case vbKeyS
13280       With frmSub
13290         intRetVal = 0
13300         If .assetno.Locked = False Then
13310           .assetno.SetFocus
13320         Else
13330           .FocusHolder.SetFocus
13340         End If
13350       End With
13360     Case vbKeyX
13370       With frmSub
13380         intRetVal = 0
13390         DoCmd.SelectObject acForm, .Parent.Name, False
13400         .Parent.cmdClose.SetFocus
13410         .Parent.cmdClose_Click  ' ** Form Procedure: frmJournal_Columns.
13420       End With
13430     End Select
13440   End If

        ' ** Ctrl keys.
13450   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
13460     blnF4Invoked = False: blnF4InvokedMouse = False
13470     Select Case intRetVal
          Case vbKeyA
13480       With frmSub
13490         intRetVal = 0
13500         If .Parent.cmdAdd.Enabled = True Then
13510           lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
13520           DoCmd.SelectObject acForm, .Parent.Name, False
13530           .Parent.cmdAdd.SetFocus
13540           .Parent.cmdAdd_Click  ' ** Form Procedure: frmJournal_Columns.
13550         End If
13560       End With
13570     Case vbKeyD
13580       With frmSub
13590         intRetVal = 0
13600         If .Parent.cmdDelete.Enabled = True Then
13610           lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
13620           DoCmd.SelectObject acForm, .Parent.Name, False
13630           .Parent.cmdDelete.SetFocus
13640           .Parent.cmdDelete_Click  ' ** Form Procedure: frmJournal_Columns.
13650         End If
13660       End With
13670     Case vbKeyE
13680       With frmSub
13690         intRetVal = 0
13700         If .Parent.cmdEdit.Enabled = True Then
13710           lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
13720           DoCmd.SelectObject acForm, .Parent.Name, False
13730           .Parent.cmdEdit.SetFocus
13740           .Parent.cmdEdit_Click  ' ** Form Procedure: frmJournal_Columns.
13750         End If
13760       End With
13770     Case vbKeyH
13780       With frmSub
13790         intRetVal = 0
13800         If .Parent.cmdRefresh.Enabled = True Then
13810           lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
13820           DoCmd.SelectObject acForm, .Parent.Name, False
13830           .Parent.cmdRefresh.SetFocus
13840           .Parent.cmdRefresh_Click  ' ** Form Procedure: frmJournal_Columns.
13850         End If
13860       End With
13870     Case vbKeyL
13880       With frmSub
13890         intRetVal = 0
13900         lngRecsCur = .RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
13910         If lngRecsCur > 0& Then
                ' ** But don't SetFocus.
13920           .Parent.cmdScrollLeft_Click  ' ** Form Procedure: frmJournal_Columns.
13930         End If
13940       End With
13950     Case vbKeyM
13960       With frmSub
13970         intRetVal = 0
13980         If .Parent.cmdMemoReveal.Visible = True And .Parent.cmdMemoReveal.Enabled = True Then
13990           .Parent.cmdMemoReveal_Click  ' ** Form Procedure: frmJournal_Columns.
14000         Else
14010           Beep
14020         End If
14030       End With
14040     Case vbKeyR
14050       With frmSub
14060         intRetVal = 0
14070         lngRecsCur = .RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
14080         If lngRecsCur > 0& Then
                ' ** But don't SetFocus.
14090           .Parent.cmdScrollRight_Click  ' ** Form Procedure: frmJournal_Columns.
14100         End If
14110       End With
14120     Case vbKeyS
14130       With frmSub
14140         intRetVal = 0
14150         strSaveMoveCtl = strControl
14160         .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
14170       End With
14180     Case vbKeyT
14190       With frmSub
14200         intRetVal = 0
14210         DoCmd.SelectObject acForm, .Parent.Name, False
14220         .Parent.cmdSwitch.SetFocus
14230         DoEvents
14240         ForcePause 1  ' ** Module Function: modCodeUtilities.
14250         .Parent.cmdSwitch_Click  ' ** Form Procedure: frmJournal_Columns.
14260       End With
14270     Case vbKeyTab, vbKeyReturn
14280       With frmSub
14290         intRetVal = 0
14300         strNext = JC_Key_Par_Next(THIS_NAME & "_Exit", .Parent)  ' ** Function: Above.
14310         DoCmd.SelectObject acForm, .Parent.Name, False
14320         .Parent.Controls(strNext).SetFocus
14330       End With
14340     Case vbKeyHome
14350       With frmSub
14360         intRetVal = 0
14370         strNext = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero, False, "First")  ' ** Function: Above.
14380         lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
14390         .Controls(strNext).SetFocus
14400       End With
14410     Case vbKeyEnd
14420       With frmSub
14430         intRetVal = 0
14440         strNext = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero, True, "Last")  ' ** Function: Above.
14450         .Controls(strNext).SetFocus
14460         lngRetVal = fSetScrollBarPosHZ(frmSub, 999&)  ' ** Module Function: modScrollBarFuncs.
14470       End With
14480     Case vbKeyPageUp
14490       With frmSub
14500         If strControl <> vbNullString Then
                ' ** Focus handled by Form_Current().
14510           strPageMoveCtl = strControl
14520         Else
14530           strPageMoveCtl = vbNullString
14540         End If
14550         .MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
14560       End With
14570     Case vbKeyPageDown
14580       With frmSub
14590         If strControl <> vbNullString Then
                ' ** Focus handled by Form_Current().
14600           strPageMoveCtl = strControl
14610         Else
14620           strPageMoveCtl = vbNullString
14630         End If
14640         .MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
14650       End With
14660     End Select
14670   End If

        ' ** Ctrl-Shift keys.
14680   If intCtrlDown And (Not intAltDown) And intShiftDown Then
14690     blnF4Invoked = False: blnF4InvokedMouse = False
14700     Select Case intRetVal
          Case vbKeyA
14710       With frmSub
14720         intRetVal = 0
14730         If .Parent.cmdUncomComAll.Enabled = True And .Parent.cmdUncomComAll.Visible = True Then
14740           DoCmd.SelectObject acForm, .Parent.Name, False
14750           .Parent.cmdUncomComAll.SetFocus
14760           .Parent.cmdUncomComAll_Click  ' ** Form Procedure: frmJournal_Columns.
14770         End If
14780       End With
14790     Case vbKeyD
14800       With frmSub
14810         intRetVal = 0
14820         If .Parent.cmdUncomDelAll.Enabled = True And .Parent.cmdUncomDelAll.Visible = True Then
14830           DoCmd.SelectObject acForm, .Parent.Name, False
14840           .Parent.cmdUncomDelAll.SetFocus
14850           .Parent.cmdUncomDelAll_Click  ' ** Form Procedure: frmJournal_Columns.
14860         End If
14870       End With
14880     Case vbKeyF
14890       With frmSub
14900         intRetVal = 0
14910         DoCmd.SelectObject acForm, .Parent.Name, False
14920         .Parent.FocusHolder.SetFocus
14930       End With
14940     Case vbKeyTab, vbKeyReturn
14950       With frmSub
14960         intRetVal = 0
14970         strNext = JC_Key_Par_Next(THIS_NAME & "_Exit", .Parent, False)  ' ** Function: Above.
14980         DoCmd.SelectObject acForm, .Parent.Name, False
14990         .Parent.Controls(strNext).SetFocus
15000       End With
15010     End Select
15020   End If

        ' ** Alt-Shift keys.
15030   If (Not intCtrlDown) And intAltDown And intShiftDown Then
15040     Select Case intRetVal
          Case vbKeyA
15050       With frmSub
15060         intRetVal = 0
15070         DoCmd.SelectObject acForm, .Parent.Name, False
15080         .Parent.opgFilter.SetFocus
15090         .Parent.opgFilter = .Parent.opgFilter_optAll.OptionValue
15100         .Parent.opgFilter_AfterUpdate  ' ** Form Procedure: frmJournal_Columns.
15110       End With
15120     Case vbKeyC
15130       With frmSub
15140         intRetVal = 0
15150         DoCmd.SelectObject acForm, .Parent.Name, False
15160         .Parent.opgFilter.SetFocus
15170         .Parent.opgFilter = .Parent.opgFilter_optCommitted.OptionValue
15180         .Parent.opgFilter_AfterUpdate  ' ** Form Procedure: frmJournal_Columns.
15190       End With
15200     Case vbKeyD
15210       With frmSub
15220         intRetVal = 0
15230         If .Parent.cmdSpecPurp_Div_Map.Enabled = True Then
15240           DoCmd.SelectObject acForm, .Parent.Name, False
15250           .Parent.cmdSpecPurp_Div_Map.SetFocus
15260           .Parent.cmdSpecPurp_Div_Map_Click  ' ** Form Procedure: frmJournal_Columns.
15270         End If
15280       End With
15290     Case vbKeyG
15300       With frmSub
15310         intRetVal = 0
15320         If .Parent.cmdSpecPurp_Misc_MapSTCGL.Enabled = True Then
15330           DoCmd.SelectObject acForm, .Parent.Name, False
15340           .Parent.cmdSpecPurp_Misc_MapSTCGL.SetFocus
15350           .Parent.cmdSpecPurp_Misc_MapSTCGL_Click  ' ** Form Procedure: frmJournal_Columns.
15360         End If
15370       End With
15380     Case vbKeyI
15390       With frmSub
15400         intRetVal = 0
15410         If .Parent.cmdSpecPurp_Sold_PaidTotal.Enabled = True Then
15420           DoCmd.SelectObject acForm, .Parent.Name, False
15430           .Parent.cmdSpecPurp_Sold_PaidTotal.SetFocus
15440           .Parent.cmdSpecPurp_Sold_PaidTotal_Click  ' ** Form Procedure: frmJournal_Columns.
15450         End If
15460       End With
15470     Case vbKeyL
15480       With frmSub
15490         intRetVal = 0
15500         If .Parent.cmdSpecPurp_Misc_MapLTCG.Enabled = True Then
15510           DoCmd.SelectObject acForm, .Parent.Name, False
15520           .Parent.cmdSpecPurp_Misc_MapLTCG.SetFocus
15530           .Parent.cmdSpecPurp_Misc_MapLTCG_Click  ' ** Form Procedure: frmJournal_Columns.
15540         End If
15550       End With
15560     Case vbKeyP
15570       With frmSub
15580         intRetVal = 0
15590         If .Parent.cmdSpecPurp_Int_Map.Enabled = True Then
15600           DoCmd.SelectObject acForm, .Parent.Name, False
15610           .Parent.cmdSpecPurp_Int_Map.SetFocus
15620           .Parent.cmdSpecPurp_Int_Map_Click  ' ** Form Procedure: frmJournal_Columns.
15630         End If
15640       End With
15650     Case vbKeyS
15660       With frmSub
15670         intRetVal = 0
15680         If .Parent.cmdSpecPurp_Purch_MapSplit.Enabled = True Then
15690           DoCmd.SelectObject acForm, .Parent.Name, False
15700           .Parent.cmdSpecPurp_Purch_MapSplit.SetFocus
15710           .Parent.cmdSpecPurp_Purch_MapSplit_Click  ' ** Form Procedure: frmJournal_Columns.
15720         End If
15730       End With
15740     Case vbKeyT
15750       With frmSub
15760         intRetVal = 0
15770         If .Parent.cmdSpecPurp_Misc_MapLTCL.Enabled = True Then
15780           DoCmd.SelectObject acForm, .Parent.Name, False
15790           .Parent.cmdSpecPurp_Misc_MapLTCL.SetFocus
15800           .Parent.cmdSpecPurp_Misc_MapLTCL_Click  ' ** Form Procedure: frmJournal_Columns.
15810         End If
15820       End With
15830     Case vbKeyU
15840       With frmSub
15850         intRetVal = 0
15860         DoCmd.SelectObject acForm, .Parent.Name, False
15870         .Parent.opgFilter.SetFocus
15880         .Parent.opgFilter = .Parent.opgFilter_optUncommitted.OptionValue
15890         .Parent.opgFilter_AfterUpdate  ' ** Form Procedure: frmJournal_Columns.
15900       End With
15910     End Select
15920   End If

EXITP:
15930   JC_Key_SubFrm = intRetVal
15940   Exit Function

ERRH:
15950   intRetVal = 0
15960   Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
15970   Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
15980   Case Else
15990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16000   End Select
16010   Resume EXITP

End Function

Public Function JC_Key_FocusDown(KeyCode As Integer, Shift As Integer, strProc As String, blnNextRec As Boolean, blnFromZero As Boolean, frmSub As Access.Form) As Integer

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Key_FocusDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strEvent As String, strCtlName As String, strNext As String
        Dim intRetVal As Integer, lngRetVal As Long

16110   intRetVal = KeyCode
16120   strPageMoveCtl = vbNullString

16130   strEvent = Mid(strProc, (CharPos(strProc, 1, "_") + 1))  ' ** Module Function: modStringFuncs.
16140   strCtlName = Left(strProc, (CharPos(strProc, 1, "_") - 1))  ' ** Module Function: modStringFuncs.

        ' ** Use bit masks to determine which key was pressed.
16150   intShiftDown = (Shift And acShiftMask) > 0
16160   intAltDown = (Shift And acAltMask) > 0
16170   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16180   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16190     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16200       With frmSub
16210         intRetVal = 0
16220         Select Case strCtlName
              Case "FocusHolder"
16230           strNext = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero)  ' ** Function: Above.
16240           .Controls(strNext).SetFocus
16250         Case "FocusHolder2"
16260           lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
16270           .FocusHolder.SetFocus
16280         End Select
16290       End With
16300     Case vbKeyPageUp, vbKeyPageDown
16310       Select Case strCtlName
            Case "FocusHolder"
16320         strPageMoveCtl = "FocusHolder"
16330       Case "FocusHolder2"
16340         strPageMoveCtl = "FocusHolder2"
16350       End Select
16360     End Select
16370   End If

        ' ** Shift keys.
16380   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16390     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16400       With frmSub
16410         intRetVal = 0
16420         Select Case strCtlName
              Case "FocusHolder"
16430           strNext = JC_Key_Sub_Next(THIS_PROC, blnNextRec, blnFromZero, False)  ' ** Function: Above.
16440           Select Case blnNextRec
                Case True
16450             blnNextRec = False
16460             If .CurrentRecord > 1 Then
16470               .MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmJournal_Columns_Sub.
16480               .Controls(strNext).SetFocus
16490             Else
16500               strNext = JC_Key_Par_Next(THIS_NAME & "_Exit", .Parent, False)  ' ** Function: Above.
16510               DoCmd.SelectObject acForm, .Parent.Name, False
16520               .Parent.Controls(strNext).SetFocus
16530             End If
16540           Case False
16550             .Controls(strNext).SetFocus
16560           End Select
16570         Case "FocusHolder2"
16580           strNext = JC_Key_Par_Next(THIS_NAME & "_Exit", .Parent, False)  ' ** Function: Above.
16590           DoCmd.SelectObject acForm, .Parent.Name, False
16600           .Parent.Controls(strNext).SetFocus
16610         End Select
16620       End With
16630     End Select
16640   End If

EXITP:
16650   JC_Key_FocusDown = intRetVal
16660   Exit Function

ERRH:
16670   intRetVal = 0
16680   Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
16690   Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
16700   Case Else
16710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16720   End Select
16730   Resume EXITP

End Function

Public Sub JC_Key_CalendarClick(strProc As String, blnNoMove As Boolean, blnNextRec As Boolean, blnFromZero As Boolean, strSaveMoveCtl As String, clsMonthClass As clsMonthCal, frmSub As Access.Form)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Key_CalendarClick"

        Dim datStartDate As Date, datEndDate As Date
        Dim strEvent As String, strCtlName As String
        Dim blnRetVal As Boolean

16810   With frmSub

16820     strEvent = Mid(strProc, (CharPos(strProc, 1, "_") + 1))  ' ** Module Function: modStringFuncs.
16830     strCtlName = Left(strProc, (CharPos(strProc, 1, "_") - 1))  ' ** Module Function: modStringFuncs.

16840     Select Case strCtlName
          Case "Calendar1"
16850       .CalendarCheck   ' ** Form Procedure: frmJournal_Columns_Sub.
16860       If .transdate.Locked = False Then
16870         datStartDate = Date
16880         datEndDate = 0
16890         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
16900         If blnRetVal = True Then
16910           If DateCheck_Post(datStartDate) = True Then  ' ** Module Function: modUtilities.
16920             .transdate = datStartDate
16930           Else
16940             datStartDate = Date
16950             .transdate = CDate(Format(datStartDate, "mm/dd/yyyy"))
16960           End If
16970         Else
16980           datStartDate = Date
16990           .transdate = CDate(Format(datStartDate, "mm/dd/yyyy"))
17000         End If
17010         blnNoMove = True
17020         .transdate.SetFocus
17030         strSaveMoveCtl = JC_Key_Sub_Next("transdate_AfterUpdate", blnNextRec, blnFromZero)  ' ** Function:Above.
17040         .cmdSave_Click   ' ** Form Procedure: frmJournal_Columns_Sub.
17050       End If
17060     Case "Calendar2"
17070       .CalendarCheck   ' ** Form Procedure: frmJournal_Columns_Sub.
17080       If .assetdate_display.Locked = False Then
17090         datStartDate = Date
17100         datEndDate = 0
17110         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
17120         If blnRetVal = True Then
17130           If DateCheck_Trade(datStartDate) = True Then  ' ** Module Function: modUtilities.
17140             datStartDate = datStartDate + time
17150             .assetdate = datStartDate
17160             .assetdate_display = CDate(Format(datStartDate, "mm/dd/yyyy"))
17170           Else
17180             datStartDate = Now()
17190             .assetdate = datStartDate
17200             .assetdate_display = CDate(Format(datStartDate, "mm/dd/yyyy"))
17210           End If
17220         Else
17230           datStartDate = Now()
17240           .assetdate = datStartDate
17250           .assetdate_display = CDate(Format(datStartDate, "mm/dd/yyyy"))
17260         End If
17270         blnNoMove = True
17280         .assetdate_display.SetFocus
17290         strSaveMoveCtl = JC_Key_Sub_Next("assetdate_display_AfterUpdate", blnNextRec, blnFromZero)  ' ** Function: Above.
17300         .cmdSave_Click   ' ** Form Procedure: frmJournal_Columns_Sub.
17310       End If
17320     End Select
17330   End With

EXITP:
17340   Exit Sub

ERRH:
17350   Select Case ERR.Number
        Case Else
17360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17370   End Select
17380   Resume EXITP

End Sub

Public Function JC_Key_SwitchScroll(KeyCode As Integer, Shift As Integer, strProc As String, THIS_NAME_SUB As String, frm As Access.Form) As Integer

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Key_SwitchScroll"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strEvent As String, strCtlName As String
        Dim strNext As String, strNextSub As String
        Dim intPos01 As Integer, lngCnt As Long
        Dim intRetVal As Integer

17410   intRetVal = KeyCode

17420   lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
17430   intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
17440   strEvent = Mid(strProc, (intPos01 + 1))
17450   strCtlName = Left(strProc, (intPos01 - 1))

        ' ** Use bit masks to determine which key was pressed.
17460   intShiftDown = (Shift And acShiftMask) > 0
17470   intAltDown = (Shift And acAltMask) > 0
17480   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17490   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17500     Select Case intRetVal
          Case vbKeyTab
17510       With frm
17520         intRetVal = 0
17530         lngRecsCur = .frmJournal_Columns_Sub.Form.RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
17540         Select Case strCtlName
              Case "cmdSwitch"
17550           If lngRecsCur > 0& Then
17560             .frmJournal_Columns_Sub.SetFocus  ' ** Wherever it lands.
17570           Else
17580             If .cmdPreviewReport.Enabled = True And .cmdPreviewReport.Visible = True Then
17590               .cmdPreviewReport.SetFocus
17600             ElseIf .cmdUncomComAll.Enabled = True And .cmdUncomComAll.Visible = True Then
17610               .cmdUncomComAll.SetFocus
17620             ElseIf .cmdAssetNew.Enabled = True Then
17630               .cmdAssetNew.SetFocus
17640             ElseIf .cmdLocNew.Enabled = True Then
17650               .cmdLocNew.SetFocus
17660             ElseIf .cmdRecurNew.Enabled = True Then
17670               .cmdRecurNew.SetFocus
17680             ElseIf .cmdAdd.Enabled = True Then
17690               .cmdClose.SetFocus
17700             End If
17710           End If
17720         Case "cmdScrollLeft"
17730           If lngRecsCur > 0& Then
17740             strNext = THIS_NAME_SUB
17750             strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Function: Above.
17760             .Controls(strNext).SetFocus
17770             .Controls(strNext).Form.Controls(strNextSub).SetFocus
17780           Else
17790             .cmdClose.SetFocus
17800           End If
17810         Case "cmdScrollRight"
17820           If lngRecsCur > 0& Then
17830             strNext = THIS_NAME_SUB
17840             strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "Last")  ' ** Function: Above.
17850             .Controls(strNext).SetFocus
17860             .Controls(strNext).Form.Controls(strNextSub).SetFocus
17870           Else
17880             .cmdClose.SetFocus
17890           End If
17900         End Select
17910       End With
17920     End Select
17930   End If

        ' ** Shift keys.
17940   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17950     Select Case intRetVal
          Case vbKeyTab
17960       With frm
17970         intRetVal = 0
17980         Select Case strCtlName
              Case "cmdSwitch"
17990           If .cmdSpecPurp_Sold_PaidTotal.Enabled = True Then
18000             .cmdSpecPurp_Sold_PaidTotal.SetFocus
18010           ElseIf .cmdSpecPurp_Purch_MapSplit.Enabled = True Then
18020             .cmdSpecPurp_Purch_MapSplit.SetFocus
18030           ElseIf .cmdSpecPurp_Misc_MapLTCG.Enabled = True Then
18040             .cmdSpecPurp_Misc_MapLTCG.SetFocus
18050           ElseIf .cmdSpecPurp_Int_Map.Enabled = True Then
18060             .cmdSpecPurp_Int_Map.SetFocus
18070           ElseIf .cmdSpecPurp_Div_Map.Enabled = True Then
18080             .cmdSpecPurp_Div_Map.SetFocus
18090           Else
18100             .cmdClose.SetFocus
18110           End If
18120         Case "cmdScrollLeft"
18130           .cmdClose.SetFocus
18140         Case "cmdScrollRight"
18150           .cmdClose.SetFocus
18160         End Select
18170       End With
18180     End Select
18190   End If

EXITP:
18200   JC_Key_SwitchScroll = intRetVal
18210   Exit Function

ERRH:
18220   intRetVal = 0
18230   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
18240   Case Else
18250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18260   End Select
18270   Resume EXITP

End Function

Public Function JC_Key_EnterMemo(KeyCode As Integer, Shift As Integer, strProc As String, THIS_NAME_SUB As String, frm As Access.Form) As Integer

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Key_EnterMemo"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strEvent As String, strCtlName As String
        Dim strNext As String, strNextSub As String
        Dim intPos01 As Integer, lngCnt As Long
        Dim intRetVal As Integer, lngRetVal As Long

18310   intRetVal = KeyCode

18320   lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
18330   intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
18340   strEvent = Mid(strProc, (intPos01 + 1))
18350   strCtlName = Left(strProc, (intPos01 - 1))

        ' ** Use bit masks to determine which key was pressed.
18360   intShiftDown = (Shift And acShiftMask) > 0
18370   intAltDown = (Shift And acAltMask) > 0
18380   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18390   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18400     Select Case strCtlName
          Case "cmdMemoReveal"
18410       Select Case intRetVal
            Case vbKeyTab
18420         With frm
18430           intRetVal = 0
18440           If .JrnlMemo_Memo.Visible = True And .JrnlMemo_Memo.Enabled = True Then
18450             .JrnlMemo_Memo.SetFocus
18460           ElseIf .cmdPreviewReport.Visible = True And .cmdPreviewReport.Enabled = True Then
18470             .cmdPreviewReport.SetFocus
18480           ElseIf .cmdAssetNew.Enabled = True Then
18490             .cmdAssetNew.SetFocus
18500           ElseIf .cmdAdd.Enabled = True Then
18510             .cmdAdd.SetFocus
18520           Else
18530             .cmdClose.SetFocus
18540           End If
18550         End With
18560       End Select
18570     Case Else
18580       Select Case intRetVal
            Case vbKeyTab, vbKeyReturn
18590         With frm
18600           intRetVal = 0
18610           Select Case strCtlName
                Case "opgEnterKey_optRight"
18620             strNext = JC_Key_Par_Next(strProc, frm)  ' ** Function: Above.
18630             If strNext = (THIS_NAME_SUB) Then
18640               .Controls(strNext).SetFocus
18650               .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
18660               lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
18670               strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Function: Above.
18680               .Controls(strNext).Form.Controls(strNextSub).SetFocus
18690             Else
18700               .Controls(strNext).SetFocus
18710             End If
18720           Case "opgEnterKey_optDown"
18730             strNext = JC_Key_Par_Next(strProc, frm)  ' ** Function: Above.
18740             If strNext = (THIS_NAME_SUB) Then
18750               .Controls(strNext).SetFocus
18760               .Controls(strNext).Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_Sub.
18770               lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 1&)  ' ** Module Function: modScrollBarFuncs.
18780               strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, True, "First")  ' ** Function: Above.
18790               .Controls(strNext).Form.Controls(strNextSub).SetFocus
18800             Else
18810               .Controls(strNext).SetFocus
18820             End If
18830           End Select
18840         End With
18850       Case vbKeyLeft
18860         With frm
18870           intRetVal = 0
18880           Select Case strCtlName
                Case "opgEnterKey_optRight"
18890             .opgFilter_optUncommitted.SetFocus
18900           Case "opgEnterKey_optDown"
18910             .opgEnterKey_optRight.SetFocus
18920           End Select
18930         End With
18940       Case vbKeyRight
18950         With frm
18960           intRetVal = 0
18970           Select Case strCtlName
                Case "opgEnterKey_optRight"
18980             .opgEnterKey_optDown.SetFocus
18990           Case "opgEnterKey_optDown"
19000             .cmdClose.SetFocus
19010           End Select
19020         End With
19030       End Select
19040     End Select
19050   End If

        ' ** Shift keys.
19060   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19070     Select Case strCtlName
          Case "cmdMemoReveal"
19080       Select Case intRetVal
            Case vbKeyTab
19090         With frm
19100           intRetVal = 0
19110           If .JrnlMemo_Memo.Visible = True And .JrnlMemo_Memo.Enabled = True Then
19120             .JrnlMemo_Memo.SetFocus
19130           ElseIf .cmdPreviewReport.Visible = True And .cmdPreviewReport.Enabled = True Then
19140             .cmdPreviewReport.SetFocus
19150           ElseIf .cmdAssetNew.Enabled = True Then
19160             .cmdAssetNew.SetFocus
19170           ElseIf .cmdAdd.Enabled = True Then
19180             .cmdAdd.SetFocus
19190           Else
19200             .cmdClose.SetFocus
19210           End If
19220         End With
19230       End Select
19240     Case Else
19250       Select Case intRetVal
            Case vbKeyTab, vbKeyReturn
19260         With frm
19270           intRetVal = 0
19280           Select Case strCtlName
                Case "opgEnterKey_optRight"
19290             strNext = JC_Key_Par_Next(strProc, frm, False)  ' ** Function: Above.
19300             If strNext = (THIS_NAME_SUB) Then
19310               .Controls(strNext).SetFocus
19320               .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
19330               strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Function: Above.
19340               .Controls(strNext).Form.Controls(strNextSub).SetFocus
19350               lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
19360             Else
19370               .Controls(strNext).SetFocus
19380             End If
19390           Case "opgEnterKey_optDown"
19400             strNext = JC_Key_Par_Next(strProc, frm, False)  ' ** Function: Above.
19410             If strNext = (THIS_NAME_SUB) Then
19420               .Controls(strNext).SetFocus
19430               .Controls(strNext).Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
19440               strNextSub = JC_Key_Sub_Next("FocusHolder_KeyDown", blnD1, blnD2, False, "Last")  ' ** Function: Above.
19450               .Controls(strNext).Form.Controls(strNextSub).SetFocus
19460               lngRetVal = fSetScrollBarPosHZ(.Controls(strNext).Form, 999&)  ' ** Module Function: modScrollBarFuncs.
19470             Else
19480               .Controls(strNext).SetFocus
19490             End If
19500           End Select
19510         End With
19520       End Select
19530     End Select
19540   End If

EXITP:
19550   JC_Key_EnterMemo = intRetVal
19560   Exit Function

ERRH:
19570   intRetVal = 0
19580   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
19590   Case Else
19600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19610   End Select
19620   Resume EXITP

End Function
