Attribute VB_Name = "modJrnlCol_Controls"
Option Compare Database
Option Explicit

'VGC 06/27/2017: CHANGES!

'tblJOURNAL_FIELD DOESN'T GET AUTOMATICALLY UPDATED!!!

Private Const THIS_NAME As String = "modJrnlCol_Controls"
' **

Public Sub JC_Ctl_Posted_Update(frmSub As Access.Form, blnWarnZeroCost As Boolean, strSaveMoveCtl As String, blnNoMove As Boolean, blnNextRec As Boolean, blnFromZero As Boolean)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Ctl_Posted_Update"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, fld As DAO.Field
        Dim strThisJType As String
        Dim lngJrnlColID As Long, lngJrnlID As Long
        Dim blnUpdate As Boolean, blnSave As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String
        Dim blnContinue As Boolean

110     blnContinue = True
120     strThisJType = vbNullString

130     With frmSub
140       Select Case .posted
          Case True

            ' **********************************
            ' ** TransDate.
            ' **********************************
150         If IsNull(.transdate) = True Then
160           blnContinue = False
170           MsgBox "A Posting Date is required to commit this Journal transaction.", vbInformation + vbOKOnly, "Entry Required"
180         Else
              ' **********************************
              ' ** AccountNo.
              ' **********************************
190           If IsNull(.accountno) = True Then
200             blnContinue = False
210             MsgBox "An Account Number is required to commit this Journal transaction.", vbInformation + vbOKOnly, "Entry Required"
220           Else
                ' **********************************
                ' ** JournalType.
                ' **********************************
230             If IsNull(.journaltype) = True Then
240               blnContinue = False
250               MsgBox "A Journal Type is required to commit this Journal transaction.", vbInformation + vbOKOnly, "Entry Required"
260             Else
270               strThisJType = .journaltype
                  ' **********************************
                  ' ** Journal_USER.
                  ' **********************************
280               If IsNull(.journal_USER) = True Then
290                 .journal_USER = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
300               End If
310             End If
320           End If
330         End If

            ' **********************************
            ' ** AssetNo.
            ' **********************************
340         If blnContinue = True Then
350           Select Case strThisJType
              Case "Dividend", "Interest", "Purchase", "Deposit", "Sold", "Withdrawn", "Liability (+)", "Liability (-)", "Cost Adj."
360             If IsNull(.assetno) = True Then
370               blnContinue = False
380               MsgBox "An Asset is required to commit this Journal transaction.", vbInformation + vbOKOnly, "Entry Required"
390             End If
400           Case "Misc.", "Paid", "Received"
410             If IsNull(.assetno) = True Then
420               .assetno = 0&
430             End If
440           End Select
450         End If  ' ** blnContinue.

            ' **********************************
            ' ** AssetDate.
            ' **********************************
460         If blnContinue = True Then
470           Select Case strThisJType
              Case "Dividend", "Interest", "Purchase", "Deposit", "Cost Adj.", "Sold", "Withdrawn", "Liability (+)", "Liability (-)"
480             If IsNull(.assetdate) = True Then
490               If IsNull(.assetdate_display) = True Then
500                 blnContinue = .ADateFromTDate  ' ** Form Function: frmJournal_Columns_Sub.
510               Else
520                 blnContinue = .ADateFromADateDisp  ' ** Form Function: frmJournal_Columns_Sub.
530               End If
540             End If  ' ** assetdate.
550             If blnContinue = True Then
560               If IsNull(.assetdate_display) = True Then
570                 blnContinue = .ADateDispFromADate  ' ** Form Function: frmJournal_Columns_Sub.
580               End If  ' ** assetdate_display.
590             End If  ' ** blnContinue.
600           Case "Received"
610             If IsNull(.assetno) = False Then
620               If .assetno > 0& Then
630                 If IsNull(.assetdate) = True Then
640                   If IsNull(.assetdate_display) = True Then
650                     blnContinue = .ADateFromTDate  ' ** Form Function: frmJournal_Columns_Sub.
660                   Else
670                     blnContinue = .ADateFromADateDisp  ' ** Form Function: frmJournal_Columns_Sub.
680                   End If
690                 End If  ' ** assetdate.
700                 If blnContinue = True Then
710                   If IsNull(.assetdate_display) = True Then
720                     blnContinue = .ADateDispFromADate  ' ** Form Function: frmJournal_Columns_Sub.
730                   End If  ' ** assetdate_display.
740                 End If  ' ** blnContinue.
750               Else
760                 blnContinue = .ADateNull  ' ** Form Function: frmJournal_Columns_Sub.
770               End If
780             Else
790               blnContinue = .ADateNull  ' ** Form Function: frmJournal_Columns_Sub.
800             End If
810           Case "Misc.", "Paid"
820             blnContinue = .ADateNull  ' ** Form Function: frmJournal_Columns_Sub.
830           End Select
840         End If  ' ** blnContinue.

            ' **********************************
            ' ** PurchaseDate.
            ' **********************************
850         If blnContinue = True Then
860           Select Case strThisJType
              Case "Sold", "Withdrawn"
870             If IsNull(.PurchaseDate) = True Then
880               blnContinue = False
890               MsgBox "An Original Trade Date is required to commit this Journal transaction." & vbCrLf & vbCrLf & _
                    "Click the Edit button to choose the related Tax Lot.", vbInformation + vbOKOnly, "Entry Required"
900             End If
910           Case "Cost Adj."
920             If IsNull(.PurchaseDate) = True Then
930               .PurchaseDate = Now()
940             End If
950           Case "Liability (-)"
960             If IsNull(.Cost) = False Then
970               If .Cost > 0@ Then
                    ' ** Liability decrease (debt paid off).
980                 If IsNull(.PurchaseDate) = True Then
990                   blnContinue = False
1000                  MsgBox "An Original Trade Date (date Liability increased)" & vbCrLf & _
                        "is required to commit this Journal transaction." & vbCrLf & vbCrLf & _
                        "Click the Assets button to choose the related Tax Lot.", vbInformation + vbOKOnly, "Entry Required"
1010                End If
1020              ElseIf .Cost < 0@ Then
                    ' ** Let Cost handle this.
1030              End If
1040            Else
                  ' ** Let Cost handle this.
1050            End If
1060          Case "Dividend", "Interest", "Purchase", "Deposit", "Liability (+)", "Misc.", "Paid", "Received"
1070            blnContinue = .PDateNull  ' ** Form Function: frmJournal_Columns_Sub.
1080          End Select
1090        End If  ' ** blnContinue.

            ' **********************************
            ' ** Recur_Name (RecurringItem).
            ' **********************************
1100        If blnContinue = True Then
1110          Select Case strThisJType
              Case "Misc.", "Paid"
1120            If IsNull(.Recur_Name) = True Then
1130              blnContinue = False
1140              MsgBox "A Description or Recurring Item is required to commit this Journal transaction.", _
                    vbInformation + vbOKOnly, "Entry Required"
1150            Else
1160              If IsNull(.Recur_Type) = True Then
1170                Select Case strThisJType
                    Case "Misc."
1180                  .Recur_Type = "Misc"
1190                Case "Paid"
1200                  .Recur_Type = "Payee"
1210                End Select
1220              End If
1230              If IsNull(.RecurringItem_ID) = True Then
                    ' ** This may remain Null if Recur_Name was entered manually.
1240                blnContinue = JC_Sort_Recur_Chk(frmSub)  ' ** Module Function: modJrnlCol_Sort.
1250              End If
1260            End If
1270          Case "Received"
1280            If .assetno = 0& Then
1290              If IsNull(.Recur_Name) = True Then
1300                blnContinue = False
1310                MsgBox "A Description or Recurring Item is required to commit this Journal transaction.", _
                      vbInformation + vbOKOnly, "Entry Required"
1320              Else
1330                If IsNull(.Recur_Type) = True Then
1340                  .Recur_Type = "Payor"
1350                End If
1360                If IsNull(.RecurringItem_ID) = True Then
                      ' ** This may remain Null if Recur_Name was entered manually.
1370                  blnContinue = JC_Sort_Recur_Chk(frmSub)  ' ** Module Function: modJrnlCol_Sort.
1380                End If
1390              End If
1400            End If
1410          Case "Dividend", "Interest", "Purchase", "Deposit", "Sold", "Withdrawn", "Liability (+)", "Liability (-)", "Cost Adj."
1420            blnContinue = JC_Sort_Recur_Null(frmSub)  ' ** Module Function: modJrnlCol_Sort.
1430          End Select
1440        End If  ' ** blnContinue.

            ' **********************************
            ' ** Shareface.
            ' **********************************
1450        If blnContinue = True Then
1460          If IsNull(.shareface) = True Then
1470            .shareface = 0#
1480          End If
1490          Select Case strThisJType
              Case "Dividend", "Interest"
1500            If .shareface = 0# Then
1510              If .accountno = "INCOME O/U" Or .accountno = "99-INCOME O/U" Then
                    ' ** Zero shareface allowed.
1520              Else
1530                blnContinue = False
1540                MsgBox "A Share/Face value is required to commit this Journal transaction.", _
                      vbInformation + vbOKOnly, "Entry Required"
1550              End If
1560            ElseIf .shareface < 0# Then
1570              blnContinue = False
1580              MsgBox "A negative Share/Face is not allowed." & vbCrLf & vbCrLf & _
                    "The reduction of Share/Face is reflected in" & vbCrLf & _
                    "the choice of Journal Type, cash and cost", vbInformation + vbOKOnly, "Invalid Entry"
1590            End If
1600          Case "Purchase", "Deposit", "Sold", "Withdrawn", "Liability (+)", "Liability (-)"
1610            If .shareface = 0# Then
1620              If .accountno = "INCOME O/U" Or .accountno = "99-INCOME O/U" Then
                    ' ** Zero shareface allowed.
1630              Else
1640                blnContinue = False
1650                MsgBox "A Share/Face value is required to commit this Journal transaction.", _
                      vbInformation + vbOKOnly, "Entry Required"
1660              End If
1670            ElseIf .shareface < 0# Then
1680              blnContinue = False
1690              MsgBox "A negative Share/Face is not allowed." & vbCrLf & vbCrLf & _
                    "The reduction of Share/Face is reflected in" & vbCrLf & _
                    "the choice of Journal Type, cash and cost", vbInformation + vbOKOnly, "Invalid Entry"
1700            End If
1710          Case "Received"
1720            If .assetno > 0& Then
1730              If .shareface = 0# Then
1740                If .accountno = "INCOME O/U" Or .accountno = "99-INCOME O/U" Then
                      ' ** Zero shareface allowed.
1750                Else
1760                  blnContinue = False
1770                  MsgBox "A Share/Face value is required to commit this Journal transaction.", _
                        vbInformation + vbOKOnly, "Entry Required"
1780                End If
1790              ElseIf .shareface < 0# Then
1800                blnContinue = False
1810                MsgBox "A negative Share/Face is not allowed." & vbCrLf & vbCrLf & _
                      "The reduction of Share/Face is reflected in" & vbCrLf & _
                      "the choice of Journal Type, cash and cost", vbInformation + vbOKOnly, "Invalid Entry"
1820              End If
1830            Else
1840              If .shareface <> 0# Then
1850                blnContinue = False
1860                MsgBox "Share/Face must be ZERO for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
1870              End If
1880            End If
1890          Case "Cost Adj.", "Misc.", "Paid"
1900            If .shareface <> 0# Then
1910              blnContinue = False
1920              MsgBox "Share/Face must be ZERO for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
1930            End If
1940          End Select
1950        End If  ' ** blnContinue.

            ' **********************************
            ' ** ICash.
            ' **********************************
1960        If blnContinue = True Then
1970          If IsNull(.ICash) = True Then
1980            .ICash = 0@
1990          End If
2000          Select Case strThisJType
              Case "Dividend", "Interest"
2010            If .ICash = 0@ Then
2020              blnContinue = False
2030              MsgBox "An Income Cash value is required to commit this Journal transaction.", vbInformation + vbOKOnly, "Invalid Entry"
2040            ElseIf .ICash < 0@ Then
2050              If .accountno = "INCOME O/U" Or .accountno = "99-INCOME O/U" Then
                    ' ** Negative icash allowed.
2060              Else
2070                blnContinue = False
2080                MsgBox "Income Cash must be greater than ZERO for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
2090              End If
2100            End If
2110          Case "Purchase", "Paid"
2120            If .ICash > 0@ Then
2130              blnContinue = False
2140              MsgBox "Income Cash cannot be positive for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
2150            End If
2160          Case "Sold", "Received"
2170            If .ICash < 0@ Then
2180              blnContinue = False
2190              MsgBox "Income Cash cannot be negative for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
2200            End If
2210          Case "Liability (+)"
                ' ** Liability increase (debt incurred).
2220            If IsNull(.Cost) = False Then
2230              If .ICash <> 0@ Then
2240                blnContinue = False
2250                MsgBox "Income Cash must be ZERO for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
2260              End If
2270            Else
                  ' ** Let Cost handle this.
2280            End If
2290          Case "Liability (-)"
2300            If IsNull(.Cost) = False Then
                  ' ** Liability decrease (debt paid off).
2310              If .ICash > 0@ Then
2320                blnContinue = False
2330                MsgBox "Income Cash cannot be positive for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
2340              End If
2350            Else
                  ' ** Let Cost handle this.
2360            End If
2370          Case "Cost Adj.", "Deposit", "Withdrawn"
2380            If .ICash <> 0@ Then
2390              blnContinue = False
2400              MsgBox "Income Cash must be ZERO for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
2410            End If
2420          Case "Misc."
                ' ** Anything goes.
2430          End Select
2440        End If  ' ** blnContinue.

            ' **********************************
            ' ** PCash.
            ' **********************************
2450        If blnContinue = True Then
2460          If IsNull(.PCash) = True Then
2470            .PCash = 0@
2480          End If
2490          Select Case strThisJType
              Case "Dividend", "Interest", "Cost Adj.", "Deposit", "Withdrawn"
2500            If .PCash <> 0@ Then
2510              blnContinue = False
2520              MsgBox "Principal Cash must be ZERO for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
2530            End If
2540          Case "Purchase", "Paid"
2550            If .PCash > 0@ Then
2560              blnContinue = False
2570              MsgBox "Principal Cash cannot be positive for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
2580            End If
2590          Case "Sold", "Received"
2600            If .PCash < 0@ Then
2610              If .accountno = "INCOME O/U" Or .accountno = "99-INCOME O/U" Then
                    ' ** Negative pcash allowed.
2620              Else
2630                blnContinue = False
2640                MsgBox "Principal Cash cannot be negative for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
2650              End If
2660            End If
2670          Case "Liability (+)"
                ' ** Liability increase (debt incurred).
2680            If IsNull(.Cost) = False Then
2690              If .PCash = 0@ Then
2700                If blnWarnZeroCost = False Then
2710                  msgResponse = MsgBox("Are you sure you want Principal Cash to be ZERO?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
2720                  If msgResponse <> vbYes Then
2730                    blnContinue = False
2740                  Else
2750                    blnWarnZeroCost = False
2760                  End If
2770                End If
2780              ElseIf .PCash < 0@ Then
2790                If .accountno = "INCOME O/U" Or .accountno = "99-INCOME O/U" Then
                      ' ** Negative pcash allowed.
2800                Else
2810                  blnContinue = False
2820                  MsgBox "Principal Cash cannot be negative for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
2830                End If
2840              End If
2850            Else
                  ' ** Let Cost handle this.
2860            End If
2870          Case "Liability (-)"
                ' ** Liability decrease (debt paid off).
2880            If IsNull(.Cost) = False Then
2890              If .PCash = 0@ Then
2900                If blnWarnZeroCost = False Then
2910                  msgResponse = MsgBox("Are you sure you want Principal Cash to be ZERO?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
2920                  If msgResponse <> vbYes Then
2930                    blnContinue = False
2940                  End If
2950                Else
2960                  blnWarnZeroCost = False
2970                End If
2980              ElseIf .PCash > 0@ Then
2990                blnContinue = False
3000                MsgBox "Principal Cash cannot be positive for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3010              End If
3020            Else
                  ' ** Let Cost handle this.
3030            End If
3040          Case "Misc."
                ' ** Anything goes.
3050          End Select
3060        End If  ' ** blnContinue.

            ' **********************************
            ' ** Cost.
            ' **********************************
3070        If blnContinue = True Then
3080          If IsNull(.Cost) = True Then
3090            .Cost = 0@
3100          End If
3110          Select Case strThisJType
              Case "Dividend", "Interest"
3120            If .Cost <> 0@ Then
3130              blnContinue = False
3140              MsgBox "Cost must be ZERO for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3150            End If
3160          Case "Purchase"
3170            If .Cost = 0@ Then
3180              blnContinue = False
3190              MsgBox "A positive Cost value is required to commit this Journal transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3200            ElseIf .Cost < 0@ Then
3210              blnContinue = False
3220              MsgBox "Cost cannot be negative for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3230            End If
3240          Case "Deposit"
3250            If .Cost = 0@ Then
3260              strTmp01 = Nz(.description, vbNullString)
3270              If InStr(strTmp01, "Stock Split") > 0 Then
                    ' ** No need to ask.
3280                msgResponse = vbYes
3290              ElseIf blnWarnZeroCost = False Then
3300                msgResponse = MsgBox("Are you sure you want this Deposit to have ZERO cost?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cost Basis")
3310              Else
3320                blnWarnZeroCost = False
3330                msgResponse = vbYes
3340              End If
3350              Select Case msgResponse
                  Case vbYes
                    ' ** Let it stand.
3360              Case Else
3370                blnContinue = False
3380              End Select
3390            ElseIf .Cost < 0@ Then
3400              blnContinue = False
3410              MsgBox "Cost cannot be negative for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3420            End If
3430          Case "Sold"
3440            If .Cost = 0@ Then
3450              blnContinue = False
3460              MsgBox "A negative Cost value is required to commit this Journal transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3470            ElseIf .Cost > 0@ Then
3480              blnContinue = False
3490              MsgBox "Cost cannot be positive for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3500            End If
3510          Case "Withdrawn"
3520            If .Cost = 0@ And blnWarnZeroCost = False Then
3530              msgResponse = MsgBox("Are you sure you want this Withdrawn to have ZERO Cost?", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cost Basis  1")
3540              Select Case msgResponse
                  Case vbYes
                    ' ** Let it stand.
3550              Case Else
3560                blnContinue = False
3570              End Select
3580            ElseIf .Cost = 0@ Then
3590              blnWarnZeroCost = False
3600            ElseIf .Cost > 0@ Then
3610              blnContinue = False
3620              MsgBox "Cost cannot be positive for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3630            End If
3640          Case "Liability (+)"
                ' ** Liability increase (debt incurred).
3650            If IsNull(.Cost) = False Then
3660              If .Cost = 0@ Then
3670                blnContinue = False
3680                MsgBox "A negative Cost value is required to commit this Journal transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3690              ElseIf .Cost > 0@ Then
3700                blnContinue = False
3710                MsgBox "Cost cannot be positive for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3720              End If
3730            Else
3740              blnContinue = False
3750              MsgBox "A negative Cost value is required to commit this Journal transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3760            End If
3770          Case "Liability (-)"
                ' ** Liability decrease (debt paid off).
3780            If IsNull(.Cost) = False Then
3790              If .Cost = 0@ Then
3800                blnContinue = False
3810                MsgBox "A positive Cost value is required to commit this Journal transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3820              ElseIf .Cost < 0@ Then
3830                blnContinue = False
3840                MsgBox "Cost cannot be negative for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3850              End If
3860            Else
3870              blnContinue = False
3880              MsgBox "A positive Cost value is required to commit this Journal transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3890            End If
3900          Case "Misc.", "Paid", "Received"
3910            If .Cost <> 0@ Then
3920              blnContinue = False
3930              MsgBox "Cost must be ZERO for this type of transaction.", vbInformation + vbOKOnly, "Invalid Entry"
3940            End If
3950          End Select
3960        End If  ' ** blnContinue.

            ' **********************************
            ' ** Description.
            ' **********************************
3970        If blnContinue = True Then
3980          If IsNull(.description) = False Then
3990            If Trim(.description) = vbNullString Then
4000              .description = Null
4010            Else
4020              If InStr(.description, Chr(34)) > 0 Then
4030                MsgBox "The Description cannot contain standard quote marks.", vbInformation + vbOKOnly, "Invalid Characters"
4040              End If
4050            End If
4060          End If
4070        End If  ' ** blnContinue.

            ' **********************************
            ' ** PrintCheck.
            ' **********************************
4080        If blnContinue = True Then
4090          Select Case strThisJType
              Case "Paid"
                ' ** Either way is fine.
4100          Case "Dividend", "Interest", "Purchase", "Deposit", "Sold", "Withdrawn", "Liability (+)", "Liability (-)", _
                  "Cost Adj.", "Misc.", "Received"
4110            If .PrintCheck = True Then
4120              .PrintCheck = False
4130            End If
4140          End Select
4150        End If  ' ** blnContinue.

            ' **********************************
            ' ** Location_ID.
            ' **********************************
4160        If blnContinue = True Then
4170          If IsNull(.Location_ID) = True Then
4180            .Location_ID = 1&  ' ** {Unassigned}, {no entry}.
4190          Else
4200            Select Case strThisJType
                Case "Purchase", "Deposit"
                  ' ** Anything's fine.
4210            Case "Liability (+)", "Liability (-)"
4220              If .Location_ID <> 1& Then
4230                .Location_ID = 1&
4240              End If
4250            Case "Dividend", "Interest", "Sold", "Withdrawn", "Cost Adj.", "Misc.", "Paid", "Received"
4260              If .Location_ID <> 1& Then
4270                .Location_ID = 1&
4280              End If
4290            End Select
4300          End If
4310        End If  ' ** blnContinue.

            ' **********************************
            ' ** revcode_ID.
            ' **********************************
4320        If blnContinue = True Then
4330          If IsNull(.revcode_ID) = True Then
4340            Select Case strThisJType
                Case "Dividend", "Interest", "Purchase", "Deposit", "Sold", "Received"
                  ' ** INCOME.
4350              .revcode_ID = REVID_INC  ' ** It happens that these both are the same.
4360              .revcode_TYPE = REVTYP_INC
4370            Case "Liability (+)", "Liability (-)", "Paid"
                  ' ** EXPENSE.
4380              .revcode_ID = REVID_EXP
4390              .revcode_TYPE = REVTYP_EXP
4400            Case "Withdrawn", "Cost Adj.", "Misc."
                  ' ** ALL.
4410              .revcode_ID = REVID_INC  ' ** Default to Income.
4420              .revcode_TYPE = REVTYP_INC
4430            End Select
4440          Else
4450            Select Case gblnRevenueExpenseTracking
                Case True
4460              Select Case gblnLinkRevTaxCodes
                  Case True
                    ' ** Just check for consistency.
4470                If IsNull(.taxcode) = True Then
4480                  blnContinue = False
4490                  MsgBox "A Tax Code is required to commit this Journal Transaction," & vbCrLf & _
                        "because you've chosen to link Income/Expense Codes with Tax Codes.", _
                        vbInformation + vbOKOnly, "Entry Required"
4500                Else
4510                  If IsNull(.revcode_ID.Column(2)) = True Or IsNull(.taxcode.Column(2)) = True Then
                        ' ** They just shouldn't be!
4520                  Else
4530                    If CLng(.revcode_ID.Column(2)) <> CLng(.taxcode.Column(2)) Then  ' ** revcode_TYPE, taxcode_type.
4540                      blnContinue = False
4550                      MsgBox "The Tax Code type does not match the Income/Expense Code type," & vbCrLf & _
                            "and you've chosen to link Income/Expense Codes with Tax Codes.", _
                            vbInformation + vbOKOnly, "Invalid Entry"
4560                    End If
4570                  End If
4580                End If
4590              Case False
4600                If IsNull(.revcode_ID.Column(2)) = True Then
                      ' ** Shouldn't be!
4610                Else
                      ' ** Pretty much anything goes, it's up to them.
4620                  If IsNull(.revcode_TYPE) = True Then
4630                    .revcode_TYPE = CLng(.revcode_ID.Column(2))  ' ** Zero-based 3rd column.
4640                  End If
4650                End If
4660              End Select
4670            Case False
4680              Select Case strThisJType
                  Case "Dividend", "Interest", "Purchase", "Deposit", "Sold", "Received"
                    ' ** INCOME.
4690                If .revcode_ID <> REVID_INC Then
4700                  .revcode_ID = REVID_INC
4710                End If
4720                .revcode_TYPE = REVTYP_INC
4730              Case "Liability (+)", "Liability (-)", "Paid"
                    ' ** EXPENSE.
4740                If .revcode_ID <> REVID_EXP Then
4750                  .revcode_ID = REVID_EXP
4760                  .revcode_TYPE = REVTYP_EXP
4770                End If
4780                .revcode_TYPE = REVTYP_EXP
4790              Case "Withdrawn", "Cost Adj.", "Misc."
                    ' ** ALL.
4800                If .revcode_ID <> REVID_INC Then
4810                  .revcode_ID = REVID_INC  ' ** Default to Income.
4820                End If
4830                .revcode_TYPE = REVTYP_INC
4840              End Select
4850            End Select
4860          End If
4870        End If  ' ** blnContinue.

            ' **********************************
            ' ** Taxcode.
            ' **********************************
4880        If blnContinue = True Then
4890          If IsNull(.taxcode) = True Then
                ' ** Set the default as if they've invoked gblnLinkRevTaxCodes,
                ' ** whether or not they actually have.
4900            Select Case .revcode_TYPE
                Case REVTYP_INC
4910              .taxcode = TAXID_INC
4920              .taxcode_description = "Unspecified Income"
4930              .taxcode_description_display = Null
4940              .taxcode_type = TAXTYP_INC
4950            Case REVTYP_EXP
4960              .taxcode = TAXID_DED
4970              .taxcode_description = "Unspecified Deduction"
4980              .taxcode_description_display = Null
4990              .taxcode_type = TAXTYP_DED
5000            End Select
5010          Else
5020            Select Case gblnIncomeTaxCoding
                Case True
5030              Select Case gblnLinkRevTaxCodes
                  Case True
5040                If .taxcode = 0& Then
5050                  blnContinue = False
5060                  MsgBox "A Tax Code is required to commit this Journal Transaction," & vbCrLf & _
                        "because you've chosen to link Income/Expense Codes with Tax Codes.", _
                        vbInformation + vbOKOnly, "Entry Required"
5070                Else
5080                  If IsNull(.taxcode.Column(2)) = True Or IsNull(.revcode_ID.Column(2)) = True Then
                        ' ** They just shouldn't be!
5090                  Else
5100                    If CLng(.taxcode.Column(2)) <> CLng(.revcode_ID.Column(2)) Then  ' ** revcode_TYPE, taxcode_type.
5110                      blnContinue = False
5120                      MsgBox "The Income/Expense Code type does not match the Tax Code type," & vbCrLf & _
                            "and you've chosen to link Tax Codes with Income/Expense Codes.", _
                            vbInformation + vbOKOnly, "Invalid Entry"
5130                    End If
5140                  End If
5150                End If
5160              Case False
5170                If IsNull(.taxcode.Column(2)) = True Then
                      ' ** Shouldn't be!
5180                Else
5190                  If .taxcode = 0& Then
5200                    Select Case .revcode_TYPE
                        Case REVTYP_INC
5210                      .taxcode = TAXID_INC
5220                      .taxcode_description = "Unspecified Income"
5230                      .taxcode_description_display = Null
5240                      .taxcode_type = TAXTYP_INC
5250                    Case REVTYP_EXP
5260                      .taxcode = TAXID_DED
5270                      .taxcode_description = "Unspecified Deduction"
5280                      .taxcode_description_display = Null
5290                      .taxcode_type = TAXTYP_DED
5300                    End Select
5310                  ElseIf .taxcode_type <> CLng(.taxcode.Column(2)) Then  ' ** Zero-based 3rd column.
5320                    .taxcode_type = CLng(.taxcode.Column(2))
5330                  End If
5340                End If
5350              End Select
5360            Case False
5370              If IsNull(.taxcode.Column(2)) = True Then
                    ' ** Shouldn't be!
5380              Else
5390                If IsNull(.taxcode_type) = True Then
5400                  .taxcode_type = CLng(.taxcode.Column(2))
5410                Else
5420                  If .taxcode = 0& Then
5430                    Select Case .revcode_TYPE
                        Case REVTYP_INC
5440                      .taxcode = TAXID_INC
5450                      .taxcode_description = "Unspecified Income"
5460                      .taxcode_description_display = Null
5470                      .taxcode_type = TAXTYP_INC
5480                    Case REVTYP_EXP
5490                      .taxcode = TAXID_DED
5500                      .taxcode_description = "Unspecified Deduction"
5510                      .taxcode_description_display = Null
5520                      .taxcode_type = TAXTYP_DED
5530                    End Select
5540                  ElseIf .taxcode_type <> CLng(.taxcode.Column(2)) Then  ' ** Zero-based 3rd column.
5550                    .taxcode_type = CLng(.taxcode.Column(2))
5560                  End If
5570                End If
5580              End If
5590            End Select
5600          End If
5610        End If  ' ** blnContinue.

5620        If blnContinue = False Then
5630          gblnCrtRpt_Zero = False  ' ** Indicating it failed a test.
5640          gblnMessage = False      ' ** Even if they wanted one, don't.
5650          .Undo
5660        Else

5670          lngJrnlColID = .JrnlCol_ID
5680          lngJrnlID = Nz(.Journal_ID, 0)
5690          blnUpdate = False

5700          Set dbs = CurrentDb

              'TRY A QUERY INSTEAD!
              ' ** I'm not sure why I chose to do it this way, but what the hey...
5710          Set rst = dbs.OpenRecordset("journal", dbOpenDynaset, dbConsistent)
5720          If lngJrnlID = 0& Then
5730            rst.AddNew
5740            For Each fld In rst.Fields
5750              Select Case fld.Type
                  Case dbBoolean
                    ' ** posted, IsAverage, Reinvested, PrintCheck.
5760                If fld.Name <> "posted" Then
5770                  fld = .Controls(fld.Name)
5780                Else
                      ' ** If this is a Map, and there's a reinvest, then yes, set this True.
5790                  If IsNull(.journalSubtype) = False Then
5800                    If Left(.journalSubtype, 3) = "Map" And .Reinvested = True Then
5810                      fld = .Controls(fld.Name)
5820                    End If
5830                  End If
5840                End If
5850              Case dbLong
                    ' ** ID (Journal_ID), assetno, Location_ID, revcode_ID, CheckNum.
5860                If fld.Name = "Location_ID" Then
5870                  fld = .Location_ID
5880                ElseIf fld.Name <> "ID" Then
5890                  fld = .Controls(fld.Name)
5900                End If
5910              Case dbCurrency
                    ' ** icash, pcash, cost.
5920                fld = .Controls(fld.Name)
5930              Case dbDouble
                    ' ** shareface, rate, pershare.
5940                fld = .Controls(fld.Name)
5950              Case dbDate
                    ' ** due, assetdate, transdate, purchaseDate
5960                fld = .Controls(fld.Name)
5970              Case dbInteger
                    ' ** taxcode.
5980                fld = .Controls(fld.Name)
5990              Case dbText
                    ' ** accountno, assettype, journaltype, journalSubtype, description, RecurringItem, journal_USER
6000                If fld.Name = "journaltype" Then
6010                  Select Case .journaltype
                      Case "Liability (+)", "Liability (-)"
6020                    fld = "Liability"
6030                  Case Else
6040                    fld = .journaltype
6050                  End Select
6060                ElseIf fld.Name = "RecurringItem" Then
6070                  If IsNull(.Recur_Name) = False Then
6080                    fld = .Recur_Name
6090                  End If
6100                ElseIf IsNull(.Controls(fld.Name)) = False Then
6110                  fld = .Controls(fld.Name)
6120                End If
6130              End Select
6140            Next
6150            rst.Update
6160            rst.Bookmark = rst.LastModified
6170            lngJrnlID = rst![ID]
6180          Else
6190            blnUpdate = True
6200            rst.FindFirst "[ID] = " & CStr(lngJrnlID)
6210            If rst.NoMatch = False Then
6220              For Each fld In rst.Fields
6230                Select Case fld.Type
                    Case dbBoolean
                      ' ** posted, IsAverage, Reinvested, PrintCheck.
6240                  If fld.Name <> "posted" Then
6250                    If fld <> .Controls(fld.Name) Then
6260                      rst.Edit
6270                      fld = .Controls(fld.Name)
6280                      rst.Update
6290                    End If
6300                  Else
                        ' ** If this is a Map, and there's a reinvest, then yes, set this True.
6310                    If IsNull(.journalSubtype) = False Then
6320                      If Left(.journalSubtype, 3) = "Map" And .Reinvested = True Then
6330                        rst.Edit
6340                        fld = .Controls(fld.Name)
6350                        rst.Update
6360                      End If
6370                    End If
6380                  End If
6390                Case dbLong
                      ' ** assetno, Location_ID, revcode_ID, CheckNum.
6400                  If fld.Name = "Location_ID" Then
6410                    If IsNull(fld) = True And IsNull(.Location_ID) = True Then
                          ' ** Skip it!
6420                    ElseIf IsNull(fld) = False And IsNull(.Location_ID) = False Then
6430                      If fld <> .Location_ID Then
6440                        rst.Edit
6450                        fld = .Location_ID
6460                        rst.Update
6470                      End If
6480                    ElseIf IsNull(fld) = True Then
6490                      rst.Edit
6500                      fld = .Location_ID
6510                      rst.Update
6520                    ElseIf IsNull(.Location_ID) = True Then
6530                      rst.Edit
6540                      fld = Null
6550                      rst.Update
6560                    End If
6570                  ElseIf fld.Name <> "ID" Then
6580                    If IsNull(fld) = True And IsNull(.Controls(fld.Name)) = True Then
                          ' ** Skip it!
6590                    ElseIf IsNull(fld) = False And IsNull(.Controls(fld.Name)) = False Then
6600                      If fld <> .Controls(fld.Name) Then
6610                        rst.Edit
6620                        fld = .Controls(fld.Name)
6630                        rst.Update
6640                      End If
6650                    ElseIf IsNull(fld) = True Then
6660                      rst.Edit
6670                      fld = .Controls(fld.Name)
6680                      rst.Update
6690                    ElseIf IsNull(.Controls(fld.Name)) = True Then
6700                      rst.Edit
6710                      fld = Null
6720                      rst.Update
6730                    End If
6740                  End If
6750                Case dbCurrency
                      ' ** icash, pcash, cost. (None should be Null.)
6760                  If IsNull(fld) = True Then
6770                    rst.Edit
6780                    fld = .Controls(fld.Name)
6790                    rst.Update
6800                  Else
6810                    If fld <> .Controls(fld.Name) Then
6820                      rst.Edit
6830                      fld = .Controls(fld.Name)
6840                      rst.Update
6850                    End If
6860                  End If
6870                Case dbDouble
                      ' ** shareface, rate, pershare. (None should be Null.)
6880                  If IsNull(fld) = True Then
6890                    rst.Edit
6900                    fld = .Controls(fld.Name)
6910                    rst.Update
6920                  Else
6930                    If fld <> .Controls(fld.Name) Then
6940                      rst.Edit
6950                      fld = .Controls(fld.Name)
6960                      rst.Update
6970                    End If
6980                  End If
6990                Case dbDate
                      ' ** due, assetdate, transdate, purchaseDate
7000                  If IsNull(fld) = True And IsNull(.Controls(fld.Name)) = True Then
                        ' ** Skip it!
7010                  ElseIf IsNull(fld) = False And IsNull(.Controls(fld.Name)) = False Then
7020                    If fld <> .Controls(fld.Name) Then
7030                      rst.Edit
7040                      fld = .Controls(fld.Name)
7050                      rst.Update
7060                    End If
7070                  ElseIf IsNull(fld) = True Then
7080                    rst.Edit
7090                    fld = .Controls(fld.Name)
7100                    rst.Update
7110                  ElseIf IsNull(.Controls(fld.Name)) = True Then
7120                    rst.Edit
7130                    fld = Null
7140                    rst.Update
7150                  End If
7160                Case dbInteger
                      ' ** taxcode.
7170                  If IsNull(fld) = True Then
7180                    rst.Edit
7190                    fld = .Controls(fld.Name)
7200                    rst.Update
7210                  Else
7220                    If fld <> .Controls(fld.Name) Then
7230                      rst.Edit
7240                      fld = .Controls(fld.Name)
7250                      rst.Update
7260                    End If
7270                  End If
7280                Case dbText
7290                  If fld.Name = "journaltype" Then
7300                    Select Case .journaltype
                        Case "Liability (+)", "Liability (-)"
7310                      If fld <> "Liability" Then
7320                        rst.Edit
7330                        fld = "Liability"
7340                        rst.Update
7350                      End If
7360                    Case Else
7370                      If fld <> .journaltype Then
7380                        rst.Edit
7390                        fld = .journaltype
7400                        rst.Update
7410                      End If
7420                    End Select
7430                  ElseIf fld.Name = "RecurringItem" Then
7440                    If IsNull(fld) = True And IsNull(.Recur_Name) = True Then
                          ' ** Skip it!
7450                    ElseIf IsNull(fld) = False And IsNull(.Recur_Name) = False Then
7460                      If fld <> .Recur_Name Then
7470                        rst.Edit
7480                        fld = .Recur_Name
7490                        rst.Update
7500                      End If
7510                    ElseIf IsNull(fld) = True Then
7520                      rst.Edit
7530                      fld = .Recur_Name
7540                      rst.Update
7550                    ElseIf IsNull(.Recur_Name) = True Then
7560                      rst.Edit
7570                      fld = Null
7580                      rst.Update
7590                    End If
7600                  Else
7610                    If IsNull(fld) = True And IsNull(.Controls(fld.Name)) = True Then
                          ' ** Skip it!
7620                    ElseIf IsNull(fld) = False And IsNull(.Controls(fld.Name)) = False Then
7630                      If fld <> .Controls(fld.Name) Then
7640                        rst.Edit
7650                        fld = .Controls(fld.Name)
7660                        rst.Update
7670                      End If
7680                    ElseIf IsNull(fld) = True Then
7690                      rst.Edit
7700                      fld = .Controls(fld.Name)
7710                      rst.Update
7720                    ElseIf IsNull(.Controls(fld.Name)) = True Then
7730                      rst.Edit
7740                      fld = Null
7750                      rst.Update
7760                    End If
7770                  End If
7780                End Select
7790              Next
7800            Else
                  ' ** Shouldn't happen, but what if it does?
                  ' ** (Admin editing, but user deleted.)
7810              MsgBox "Another user deleted this Journal entry.", vbCritical + vbOKOnly, "Journal Update Failed"
                  'DO YOU WISH TO ADD IT?
7820            End If
7830          End If
7840          rst.Close

7850          blnSave = False

7860          If .JrnlMemo_HasMemo = True Then
7870            If IsNull(.JrnlMemo_Memo) = True Then
                  ' ** Delete tblJournal_Memo, by specified [jid].
7880              Set qdf = dbs.QueryDefs("qryJournal_Columns_28_01")
7890              With qdf.Parameters
7900                ![jid] = lngJrnlID
7910              End With
7920              qdf.Execute
7930            Else
                  ' ** tblJournal_Memo, by specified [jid].
7940              Set qdf = dbs.QueryDefs("qryJournal_Columns_28_04")
7950              With qdf.Parameters
7960                ![jid] = lngJrnlID
7970              End With
7980              Set rst = qdf.OpenRecordset
7990              With rst
8000                If .BOF = True And .EOF = True Then
                      ' ** Oops! It's not there!
8010                  .AddNew
8020                  ![Journal_ID] = lngJrnlID
8030                Else
8040                  .MoveFirst
8050                  .Edit
8060                End If
8070                ![journaltype] = frmSub.journaltype
8080                ![accountno] = frmSub.accountno
8090                ![transdate] = frmSub.transdate
8100                ![JrnlMemo_Memo] = frmSub.JrnlMemo_Memo
8110                ![JrnlMemo_DateModified] = Now()
8120                .Update
8130                .Close
8140              End With
8150            End If
8160          ElseIf IsNull(.JrnlMemo_Memo) = False Then
                ' ** Append tblJournal_Column to tblJournal_Memo, by specified [jid].
8170            Set qdf = dbs.QueryDefs("qryJournal_Columns_28_03")
8180            With qdf.Parameters
8190              ![jid] = lngJrnlID
8200            End With
8210            qdf.Execute
8220            .JrnlMemo_HasMemo = True
8230            blnSave = True
8240          End If

8250          dbs.Close

8260          If blnUpdate = False Then
8270            .Journal_ID = lngJrnlID
8280            blnSave = True
8290          End If

8300          If blnSave = True Then
8310            strSaveMoveCtl = vbNullString
8320            blnNoMove = True
8330            .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
8340            DoEvents
8350          End If

8360          JC_Btn_Set strThisJType, True, .Parent  ' ** Module Procedure: modJrnlCol_Buttons.

              'SEE IF I CAN DO THIS ONLY IF IT'S NOT DONE WITH CommitRec()!
              'AddRec  ' ** Procedure: Below.

8370        End If  ' ** blnContinue.

8380      Case False
            ' ** I think I'm going to lock this to True,
            ' ** then it goes off if they make any changes.
8390        strSaveMoveCtl = JC_Key_Sub_Next("posted_AfterUpdate", blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
8400        .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
8410      End Select
8420    End With  ' ** Me.

EXITP:
8430    Set fld = Nothing
8440    Set rst = Nothing
8450    Set qdf = Nothing
8460    Set dbs = Nothing
8470    Exit Sub

ERRH:
8480    frmSub.posted = False
8490    Select Case ERR.Number
        Case Else
8500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8510    End Select
8520    Resume EXITP

End Sub

Public Sub JC_Ctl_Delete(blnGoneToReport As Boolean, frm As Access.Form)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Ctl_Delete"

        Dim rst As DAO.Recordset
        Dim lngJrnlColID As Long, lngRecsCur As Long
        Dim blnEmpty As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String, lngTmp02 As Long

8610    With frm
8620      If blnGoneToReport = True Then
8630        blnGoneToReport = False
8640      End If
8650      Select Case .frmJournal_Columns_Sub.Form.posted
          Case True
8660        strTmp01 = "committed"
8670        blnEmpty = False
8680      Case False
8690        strTmp01 = "incomplete"
8700        If IsNull(.frmJournal_Columns_Sub.Form.accountno) = True And _
                IsNull(.frmJournal_Columns_Sub.Form.journaltype) = True Then
8710          blnEmpty = True
8720        Else
8730          blnEmpty = False
8740        End If
8750      End Select
8760      Beep
8770      Select Case blnEmpty
          Case True
8780        msgResponse = vbOK
8790      Case False
8800        Select Case blnGoneToReport
            Case True
8810          msgResponse = vbYes
8820        Case False
8830          msgResponse = MsgBox("Are you sure you want to delete this " & strTmp01 & " transaction?", _
                vbQuestion + vbYesNo, "Delete " & strTmp01 & " Transaction")
8840        End Select
8850      End Select
8860      Select Case msgResponse
          Case vbYes
8870        lngJrnlColID = 0&
8880        gblnDeleting = True
8890        lngRecsCur = .frmJournal_Columns_Sub.Form.RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
8900        If lngRecsCur > 1& Then
8910          lngTmp02 = .frmJournal_Columns_Sub.Form.JrnlCol_ID
8920          Set rst = .frmJournal_Columns_Sub.Form.RecordsetClone
8930          lngJrnlColID = JC_Msc_Find_Adjascent(lngTmp02, rst)  ' ** Module Function: modJrnlCol_Misc.
8940          Set rst = Nothing
8950        End If
8960        .frmJournal_Columns_Sub.SetFocus
            '.frmJournal_Columns_Sub.Form.DelRec  ' ** Form Procedure: frmJournal_Columns_Sub.
8970        .frmJournal_Columns_Sub.Form.DelRec_Send  ' ** Form Procedure: frmJournal_Columns_Sub.
8980        DoEvents
8990        If lngJrnlColID > 0& Then
9000          .frmJournal_Columns_Sub.Form.MoveRec 0, lngJrnlColID  ' ** Form Procedure: frmJournal_Columns_Sub.
9010        Else
              ' ** Enable all the special purpose buttons, and leave the subform empty.
9020          JC_Btn_Set vbNullString, True, frm  ' ** Module Procedure: modJrnlCol_Buttons.
9030          .cmdAdd.SetFocus
9040        End If
9050      Case Else
            ' ** Nothing.
9060      End Select
9070    End With

EXITP:
9080    Set rst = Nothing
9090    Exit Sub

ERRH:
9100    Select Case ERR.Number
        Case Else
9110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9120    End Select
9130    Resume EXITP

End Sub

Public Sub JC_Ctl_UnCommitOne(frm As Access.Form)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Ctl_UnCommitOne"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngJrnlColID As Long, lngJrnlID As Long
        Dim lngRecsCur As Long

9210    With frm
9220      DoCmd.Hourglass True
9230      DoEvents
9240      lngRecsCur = .frmJournal_Columns_Sub.Form.RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
9250      If lngRecsCur > 0& Then
9260        If IsNull(.frmJournal_Columns_Sub.Form.JrnlCol_ID) = False Then
9270          lngJrnlColID = .frmJournal_Columns_Sub.Form.JrnlCol_ID
9280          If .frmJournal_Columns_Sub.Form.posted = True Then
9290            lngJrnlID = .frmJournal_Columns_Sub.Form.Journal_ID
9300            Set dbs = CurrentDb
9310            With dbs
                  ' ** Update qryJournal_Columns_02l (tblJournal_Column, to uncommit,
                  ' ** with Journal_ID_new, posted_new, by specified [jcolid]).
9320              Set qdf = .QueryDefs("qryJournal_Columns_02m")
9330              With qdf.Parameters
9340                ![jcolid] = lngJrnlColID
9350              End With
9360              qdf.Execute
9370              Set qdf = Nothing
9380              DoEvents
                  ' ** Delete Journal, by specified [jid].
9390              Set qdf = .QueryDefs("qryJournal_Columns_02n")
9400              With qdf.Parameters
9410                ![jid] = lngJrnlID
9420              End With
9430              qdf.Execute
9440              Set qdf = Nothing
9450              .Close
9460            End With
9470            Set dbs = Nothing
9480            DoEvents
9490            .frmJournal_Columns_Sub.SetFocus
9500            .frmJournal_Columns_Sub.Form.Requery
9510          Else
9520            Beep
9530          End If
9540        Else
9550          Beep
9560        End If
9570      Else
9580        Beep
9590      End If
9600      DoCmd.Hourglass False
9610    End With

EXITP:
9620    Set qdf = Nothing
9630    Set dbs = Nothing
9640    Exit Sub

ERRH:
9650    DoCmd.Hourglass False
9660    Select Case ERR.Number
        Case Else
9670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9680    End Select
9690    Resume EXITP

End Sub

Public Sub JC_Ctl_UncomDelAll(frm As Access.Form)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Ctl_UncomDelAll"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String

9710    With frm
9720      If .RecsTot_Uncommitted > 0& Then
9730        If .RecsTot_Uncommitted = 1& Then
9740          strTmp01 = "that entry"
9750        Else
9760          strTmp01 = "those " & CStr(.RecsTot_Uncommitted) & " entries"
9770        End If
9780        msgResponse = MsgBox("Are you sure you want to delete " & strTmp01 & "?", vbQuestion + vbYesNo, "Delete All Uncommitted Entries")
9790        If msgResponse = vbYes Then
9800          Set dbs = CurrentDb
9810          With dbs
9820            Select Case gblnAdmin
                Case True
                  ' ** Delete tblJournal_Column, just uncommitted, all Users.
9830              Set qdf = .QueryDefs("qryJournal_Columns_02h")
9840            Case False
                  ' ** Delete tblJournal_Column, just uncommitted, by specified [usr].
9850              Set qdf = .QueryDefs("qryJournal_Columns_02i")
9860              With qdf.Parameters
9870                ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
9880              End With
9890            End Select
9900            qdf.Execute
9910            .Close
9920          End With
9930          If .opgFilter <> .opgFilter_optAll.OptionValue Then
9940            .opgFilter = .opgFilter_optAll.OptionValue
9950            .opgFilter_AfterUpdate  ' ** Form Procedure: frmJournal_Columns.
9960          Else
9970            .frmJournal_Columns_Sub.Form.Requery
9980          End If
              ' ** Enable all the special purpose buttons, and leave the subform empty.
9990          JC_Btn_Set vbNullString, True, frm  ' ** Module Procedure: modJrnlCol_Buttons.
10000         .cmdAdd.SetFocus
10010       End If
10020     Else
10030       MsgBox "There are no uncommitted entries to delete.", vbInformation + vbOKOnly, "Nothing To Do"
10040     End If
10050   End With

EXITP:
10060   Set qdf = Nothing
10070   Set dbs = Nothing
10080   Exit Sub

ERRH:
10090   Select Case ERR.Number
        Case Else
10100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10110   End Select
10120   Resume EXITP

End Sub

Public Sub JC_Ctl_MemoReveal(frm As Access.Form)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Ctl_MemoReveal"

10210   With frm
10220     Select Case .JrnlMemo_Memo.Visible
          Case True
            ' ** Switch to buttons.

10230       .JrnlMemo_Memo.Enabled = False
10240       .JrnlMemo_Memo.Visible = False
10250       .JrnlMemo_Memo_lbl2.Visible = False
10260       .JrnlMemo_Memo_box.Visible = False
10270       .JrnlMemo_Memo_box2.Visible = False
10280       .JrnlMemo_Memo_box3.Visible = False

10290       .cmdPreviewReport.Visible = True
10300       Select Case .cmdPreviewReport.Enabled
            Case True
10310         .cmdPreviewReport_raised_img.Visible = True
10320       Case False
10330         .cmdPreviewReport_raised_img_dis.Visible = True
10340       End Select
10350       .cmdPrintReport.Visible = True
10360       Select Case .cmdPrintReport.Enabled
            Case True
10370         .cmdPrintReport_raised_img.Visible = True
10380       Case False
10390         .cmdPrintReport_raised_img_dis.Visible = True
10400       End Select

10410       .RecsTot_Committed.Visible = True
10420       .RecsTot_Uncommitted.Visible = True
10430       .RecsTot_box.Visible = True
10440       .RecsTot_vline01.Visible = True
10450       .RecsTot_vline02.Visible = True
10460       .RecsTot_vline03.Visible = True
10470       .RecsTot_vline04.Visible = True

10480       .cmdUncom_box.Visible = True
10490       .cmdUncom_vline01.Visible = True
10500       .cmdUncom_vline02.Visible = True
10510       .cmdUncom_vline03.Visible = True
10520       .cmdUncom_vline04.Visible = True
10530       .cmdUncom_lbl.Visible = True

10540       .cmdUncomComAll.Visible = True
10550       Select Case .cmdUncomComAll.Enabled
            Case True
10560         .cmdUncomComAll_raised_img.Visible = True
10570         .cmdUncomComAll_raised_img_dis.Visible = False
10580       Case False
10590         .cmdUncomComAll_raised_img_dis.Visible = True
10600         .cmdUncomComAll_raised_img.Visible = False
10610       End Select
10620       .cmdUncomComAll_raised_focus_dots_img.Visible = False
10630       .cmdUncomComAll_sunken_focus_dots_img.Visible = False

10640       .cmdUncomDelAll.Visible = True
10650       Select Case .cmdUncomDelAll.Enabled
            Case True
10660         .cmdUncomDelAll_raised_img.Visible = True
10670         .cmdUncomDelAll_raised_img_dis.Visible = False
10680       Case False
10690         .cmdUncomDelAll_raised_img_dis.Visible = True
10700         .cmdUncomDelAll_raised_img.Visible = False
10710       End Select
10720       .cmdUncomDelAll_raised_focus_dots_img.Visible = False
10730       .cmdUncomDelAll_sunken_focus_dots_img.Visible = False

10740       .cmdUnCommitOne.Visible = True
10750       Select Case .cmdUnCommitOne.Enabled
            Case True
10760         .cmdUnCommitOne_raised_img.Visible = True
10770         .cmdUnCommitOne_raised_img_dis.Visible = False
10780       Case False
10790         .cmdUnCommitOne_raised_img_dis.Visible = True
10800         .cmdUnCommitOne_raised_img.Visible = False
10810       End Select
10820       .cmdUnCommitOne_raised_focus_dots_img.Visible = False
10830       .cmdUnCommitOne_sunken_focus_dots_img.Visible = False

10840       .cmdMemoReveal_R_raised_focus_img.Visible = True
10850       .cmdMemoReveal_R_raised_img.Visible = False
10860       .cmdMemoReveal_R_raised_semifocus_img.Visible = False
10870       .cmdMemoReveal_R_sunken_focus_img.Visible = False
10880       .cmdMemoReveal_R_raised_img_dis.Visible = False

10890       .cmdMemoReveal_L_raised_focus_img.Visible = False
10900       .cmdMemoReveal_L_raised_img.Visible = False
10910       .cmdMemoReveal_L_raised_semifocus_img.Visible = False
10920       .cmdMemoReveal_L_sunken_focus_img.Visible = False
10930       .cmdMemoReveal_L_raised_img_dis.Visible = False

10940       If .cmdPreviewReport.Enabled = True Then
10950         .cmdPreviewReport.SetFocus
10960       ElseIf .cmdAssetNew.Enabled = True Then
10970         .cmdAssetNew.SetFocus
10980       ElseIf .cmdAdd.Enabled = True Then
10990         .cmdAdd.SetFocus
11000       Else
11010         .cmdClose.SetFocus
11020       End If

11030     Case False
            ' ** Switch to memo.

11040       .cmdUncomComAll.Visible = False
11050       .cmdUncomComAll_raised_img.Visible = False
11060       .cmdUncomComAll_raised_img_dis.Visible = False
11070       .cmdUncomComAll_raised_focus_dots_img.Visible = False
11080       .cmdUncomComAll_sunken_focus_dots_img.Visible = False

11090       .cmdUncomDelAll.Visible = False
11100       .cmdUncomDelAll_raised_img.Visible = False
11110       .cmdUncomDelAll_raised_img_dis.Visible = False
11120       .cmdUncomDelAll_raised_focus_dots_img.Visible = False
11130       .cmdUncomDelAll_sunken_focus_dots_img.Visible = False

11140       .cmdUnCommitOne.Visible = False
11150       .cmdUnCommitOne_raised_img.Visible = False
11160       .cmdUnCommitOne_raised_img_dis.Visible = False
11170       .cmdUnCommitOne_raised_focus_dots_img.Visible = False
11180       .cmdUnCommitOne_sunken_focus_dots_img.Visible = False

11190       .cmdUncom_box.Visible = False
11200       .cmdUncom_vline01.Visible = False
11210       .cmdUncom_vline02.Visible = False
11220       .cmdUncom_vline03.Visible = False
11230       .cmdUncom_vline04.Visible = False
11240       .cmdUncom_lbl.Visible = False

11250       .RecsTot_Committed.Visible = False
11260       .RecsTot_Uncommitted.Visible = False
11270       .RecsTot_box.Visible = False
11280       .RecsTot_vline01.Visible = False
11290       .RecsTot_vline02.Visible = False
11300       .RecsTot_vline03.Visible = False
11310       .RecsTot_vline04.Visible = False

11320       .cmdPreviewReport.Visible = False
11330       .cmdPreviewReport_raised_img.Visible = False
11340       .cmdPreviewReport_raised_semifocus_dots_img.Visible = False
11350       .cmdPreviewReport_raised_focus_img.Visible = False
11360       .cmdPreviewReport_raised_focus_dots_img.Visible = False
11370       .cmdPreviewReport_sunken_focus_dots_img.Visible = False
11380       .cmdPreviewReport_raised_img_dis.Visible = False
11390       .cmdPrintReport.Visible = False
11400       .cmdPrintReport_raised_img.Visible = False
11410       .cmdPrintReport_raised_semifocus_dots_img.Visible = False
11420       .cmdPrintReport_raised_focus_img.Visible = False
11430       .cmdPrintReport_raised_focus_dots_img.Visible = False
11440       .cmdPrintReport_sunken_focus_dots_img.Visible = False
11450       .cmdPrintReport_raised_img_dis.Visible = False

11460       .JrnlMemo_Memo.Enabled = True
11470       .JrnlMemo_Memo.Visible = True
11480       .JrnlMemo_Memo_lbl2.Visible = True
11490       .JrnlMemo_Memo_box.Visible = True
11500       .JrnlMemo_Memo_box2.Visible = True
11510       .JrnlMemo_Memo_box3.Visible = True

11520       .cmdMemoReveal_L_raised_focus_img.Visible = True
11530       .cmdMemoReveal_L_raised_img.Visible = False
11540       .cmdMemoReveal_L_raised_semifocus_img.Visible = False
11550       .cmdMemoReveal_L_sunken_focus_img.Visible = False
11560       .cmdMemoReveal_L_raised_img_dis.Visible = False

11570       .cmdMemoReveal_R_raised_focus_img.Visible = False
11580       .cmdMemoReveal_R_raised_img.Visible = False
11590       .cmdMemoReveal_R_raised_semifocus_img.Visible = False
11600       .cmdMemoReveal_R_sunken_focus_img.Visible = False
11610       .cmdMemoReveal_R_raised_img_dis.Visible = False

11620       .JrnlMemo_Memo.SetFocus

11630     End Select

11640   End With

EXITP:
11650   Exit Sub

ERRH:
11660   Select Case ERR.Number
        Case Else
11670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11680   End Select
11690   Resume EXITP

End Sub
