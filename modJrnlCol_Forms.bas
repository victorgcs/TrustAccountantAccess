Attribute VB_Name = "modJrnlCol_Forms"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modJrnlCol_Forms"

'VGC 09/09/2017: CHANGES!

'tblJOURNAL_FIELD DOESN'T GET AUTOMATICALLY UPDATED!!!

Private Const FRM_PAR As String = "frmJournal_Columns"
'Private Const FRM_SUB As String = "frmJournal_Columns_Sub"

' ** Array: arr_varNewRec().
Private lngNewRecs As Long, arr_varNewRec() As Variant
Private Const N_ELEMS As Integer = 3  ' ** Array's first-element UBound().
Private Const N_ID   As Integer = 0
Private Const N_CMTD As Integer = 1
Private Const N_JTYP As Integer = 2
Private Const N_CHK  As Integer = 3

Private lngRecsCur As Long
' **

Public Sub JC_Frm_Report(lngRecCnt As Long, intMode As Integer)
' ** Called by:
' **   frmJournal_Columns:
' **     cmdPreviewReport_Click()
' **     cmdPrintReport_Click()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_Report"

        Dim strDocName As String
        Dim varTmp00 As Variant
        Dim blnContinue As Boolean

110     blnContinue = True

120     gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
130     gstrFormQuerySpec = FRM_PAR
140     gstrReportCallingForm = FRM_PAR

150     lngRecsCur = lngRecCnt
160     If lngRecsCur > 0& Then
          ' ** tblJournal_Column, just totally empty records.
170       varTmp00 = DCount("*", "qryJournal_Columns_02f")
180       If IsNull(varTmp00) = False Then
190         If CLng(varTmp00) = lngRecsCur Then
200           blnContinue = False
210           Beep
220           MsgBox "There are no entries to report.", vbInformation + vbOKOnly, "Nothing To Do"
230         End If
240       End If
250     Else
260       blnContinue = False
270       Beep
280       MsgBox "There are no entries to report.", vbInformation + vbOKOnly, "Nothing To Do"
290     End If

300     If blnContinue = True Then
310       strDocName = "rptPostingJournal_Column"
320       Select Case gblnAdmin
          Case True
330         DoCmd.OpenReport strDocName, intMode
340       Case False
350         DoCmd.OpenReport strDocName, intMode, , "[journal_USER] = '" & gstrJournalUser & "'"
360       End Select
370       If intMode = acViewPreview Then
380         DoCmd.Maximize
390         DoCmd.RunCommand acCmdFitToWindow
400       End If
410     End If

EXITP:
420     Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Public Sub JC_Frm_Assets(frmPar As Access.Form, frmSub As Access.Form, blnToTaxLot As Boolean)
' ** Called by:
' **   frmJournal_Columns:
' **     cmdAssetNew_Click()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_Assets"

        Dim strDocName As String, strThisJType As String
        Dim lngAssetNo As Long
        Dim blnJustAdd As Boolean, blnAddAndSet As Boolean
        Dim varTmp00 As Variant
        Dim intRetVal As Integer

510     With frmPar

520       strDocName = "frmAssets_Add_Purchase"
530       lngAssetNo = 0&

540       .cmdLocNew.Enabled = False
550       .cmdRecurNew.Enabled = False

560       lngRecsCur = frmSub.RecCnt  ' ** Form Function: frmJournal_Columns_sub.
570       If lngRecsCur > 0& Then
580         Select Case IsNull(frmSub.accountno)
            Case True
590           blnJustAdd = True
600           blnAddAndSet = False
610         Case False
620           Select Case IsNull(frmSub.journaltype)
              Case True
630             blnJustAdd = True
640             blnAddAndSet = False
650           Case False
660             strThisJType = frmSub.journaltype
670             Select Case strThisJType
                Case "Dividend", "Interest", "Deposit", "Purchase", "Withdrawn", "Sold", "Liability (+)", "Liability (-)", "Cost Adj."
680               blnJustAdd = False
690             Case "Misc.", "Paid", "Received"
700               blnJustAdd = True
710               blnAddAndSet = False
720             End Select
730           End Select
740         End Select
750       Else
760         blnJustAdd = True
770         blnAddAndSet = False
780       End If

790       If blnJustAdd = False Then
800         varTmp00 = frmSub.assetno
810         If IsNull(varTmp00) = False Then
820           lngAssetNo = CLng(varTmp00)
830           If lngAssetNo > 0& Then
840             If frmSub.posted = False Then
850               strThisJType = frmSub.journaltype
860               Select Case strThisJType
                  Case "Sold", "Withdrawn", "Liability (-)"
870                 If Nz(frmSub.sharface, 0) > 0& And Nz(frmSub.Cost, 0) = 0& Then
880                   strDocName = vbNullString
890                   blnJustAdd = False
900                   blnAddAndSet = False
910                   DoCmd.Hourglass True
920                   DoEvents
930                   blnToTaxLot = True
940                   .JrnlCol_ID = frmSub.JrnlCol_ID
950                   .ToTaxLot = 1&  ' ** 0 = Nothing; 1 = To Tax Lot; 2 = OK Single; 3 = OK Multi; -4 = Cancel Return; PLUS...
960                   .Parent.TaxLotFrom = "cmdAssetNew"
970                   .frmJournal_Columns_Sub.SetFocus
980                   intRetVal = OpenLotInfoForm(False, FRM_PAR)  ' ** Module Function: modPurchaseSold.
                      ' ** Since the OpenLotInfoForm() function calls the form as a Dialog,
                      ' ** it should return right to here.
                      ' ** Return Values:
                      ' **    0 OK.
                      ' **   -1 Input missing.
                      ' **   -2 No holdings.
                      ' **   -3 Insufficient holdings.
                      ' **   -4 Zero shares.
                      ' **   -9 Data problem.
990                   If intRetVal <> 0 Then
                        ' ** A non-zero here indicates a problem within the OpenLotInfoForm() function,
                        ' ** and it never went to the form.
1000                    .ToTaxLot = CLng(intRetVal)  ' ** Just 'cause.
1010                    gblnSetFocus = True
1020                    .TimerInterval = 250&
1030                  End If
1040                Else
1050                  blnJustAdd = True
1060                  blnAddAndSet = True
1070                End If
1080              Case Else
1090                blnJustAdd = True
1100                blnAddAndSet = False
1110              End Select
1120            Else
                  ' ** If it's really to be changed, they'll have to set it.
1130              blnJustAdd = True
1140              blnAddAndSet = False
1150            End If
1160          Else
1170            blnJustAdd = True
1180            blnAddAndSet = True
1190          End If
1200        Else
1210          blnJustAdd = True
1220          blnAddAndSet = True
1230        End If
1240      End If  ' ** blnJustAdd.

1250      If blnJustAdd = True Then

1260        gblnMessage = True  ' ** If it comes back false, they canceled it.
1270        gstrPurchaseAsset = vbNullString
1280        gdblCrtRpt_CostTot = 0#  ' ** Borrowing this variable from the Court Reports.
1290        DoCmd.OpenForm strDocName, , , , acFormAdd, acDialog, FRM_PAR & "~" & CStr(lngAssetNo)

1300        .frmJournal_Columns_Sub.SetFocus
1310        DoEvents
1320        Select Case gblnMessage
            Case True
              ' ** A new asset has been added.
1330          frmSub.assetno.Requery
1340          If gdblCrtRpt_CostTot > 0# And blnAddAndSet = True Then
1350            frmSub.assetno = CLng(gdblCrtRpt_CostTot)
1360            frmSub.assetno_description = gstrPurchaseAsset
1370          End If
1380        Case False
              ' ** They canceled it.
1390        End Select

1400        If blnAddAndSet = True Then
1410          frmSub.assetno.SetFocus
1420        Else
1430          DoCmd.SelectObject acForm, frmPar.Name, False
1440          .cmdAssetNew.SetFocus
1450        End If

1460        gblnMessage = False
1470        gdblCrtRpt_CostTot = 0#
1480        gstrPurchaseAsset = vbNullString

1490        .cmdLocNew.Enabled = True
1500        .cmdRecurNew.Enabled = True

1510      End If  ' ** blnJustAdd.

1520    End With

EXITP:
1530    Exit Sub

ERRH:
1540    DoCmd.Hourglass False
1550    frmPar.cmdLocNew.Enabled = True
1560    frmPar.cmdRecurNew.Enabled = True
1570    Select Case ERR.Number
        Case Else
1580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1590    End Select
1600    Resume EXITP

End Sub

Public Sub JC_Frm_Locations(frmPar As Access.Form, frmSub As Access.Form)
' ** Called by:
' **   frmJournal_Columns:
' **     cmdLocNew_Click()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_Locations"

        Dim strDocName As String, strThisJType As String
        Dim blnJustAdd As Boolean, blnAddAndSet As Boolean

1710    With frmPar

1720      .cmdAssetNew.Enabled = False
1730      .cmdRecurNew.Enabled = False

1740      blnJustAdd = True
1750      lngRecsCur = frmSub.RecCnt  ' ** Form Function: frmJournal_Columns_sub.
1760      If lngRecsCur > 0& Then
1770        Select Case IsNull(frmSub.accountno)
            Case True
1780          blnAddAndSet = False
1790        Case False
1800          Select Case IsNull(frmSub.journaltype)
              Case True
1810            blnAddAndSet = False
1820          Case False
1830            strThisJType = frmSub.journaltype
1840            Select Case strThisJType
                Case "Deposit", "Purchase", "Liability (+)", "Cost Adj."
1850              blnAddAndSet = True
1860            Case "Received"
1870              Select Case IsNull(frmSub.assetno)
                  Case True
1880                blnAddAndSet = False
1890              Case False
1900                If frmSub.assetno > 0& Then
1910                  blnAddAndSet = True
1920                Else
1930                  blnAddAndSet = False
1940                End If
1950              End Select
1960            Case "Dividend", "Interest", "Withdrawn", "Sold", "Liability (-)", "Misc.", "Paid"
1970              blnAddAndSet = False
1980            End Select
1990          End Select
2000        End Select
2010      Else
2020        blnAddAndSet = False
2030      End If

2040      If blnJustAdd = True Then

2050        gblnMessage = True  ' ** If it comes back false, they canceled it.
2060        gdblCrtRpt_CostTot = 0#  ' ** Borrowing this variable from the Court Reports.
2070        gstrCrtRpt_NetLoss = vbNullString  ' ** Borrowing this variable from the Court Reports.
2080        strDocName = "frmLocations_Add_Purchase"
2090        DoCmd.OpenForm strDocName, , , , acFormAdd, acDialog, FRM_PAR

2100        .frmJournal_Columns_Sub.SetFocus
2110        DoEvents
2120        Select Case gblnMessage
            Case True
              ' ** A new location has been added.
2130          frmSub.Location_ID.Requery
2140          If gdblCrtRpt_CostTot > 0# And blnAddAndSet = True Then
2150            frmSub.Location_ID = CLng(gdblCrtRpt_CostTot)
2160            frmSub.Loc_Name = gstrCrtRpt_NetLoss
2170            frmSub.Loc_Name_display = gstrCrtRpt_NetLoss
2180          End If
2190        Case False
              ' ** They canceled it.
2200        End Select

2210        If blnAddAndSet = True Then
2220          frmSub.Location_ID.SetFocus
2230        Else
2240          DoCmd.SelectObject acForm, frmPar.Name, False
2250          .cmdLocNew.SetFocus
2260        End If

2270      End If  ' ** blnJustAdd.

2280      gblnMessage = False
2290      gdblCrtRpt_CostTot = 0#
2300      gstrCrtRpt_NetLoss = vbNullString

2310      .cmdAssetNew.Enabled = True
2320      .cmdRecurNew.Enabled = True

2330    End With

EXITP:
2340    Exit Sub

ERRH:
2350    DoCmd.Hourglass False
2360    frmPar.cmdAssetNew.Enabled = True
2370    frmPar.cmdRecurNew.Enabled = True
2380    Select Case ERR.Number
        Case Else
2390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2400    End Select
2410    Resume EXITP

End Sub

Public Sub JC_Frm_RecurringItems(frmPar As Access.Form, frmSub As Access.Form)
' ** Called by:
' **   frmJournal_Columns:
' **     cmdRecurNew_Click()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_RecurringItems"

        Dim strDocName As String, strThisJType As String
        Dim blnJustAdd As Boolean, blnAddAndSet As Boolean
        Dim blnContinue As Boolean

2510    blnContinue = True

2520    With frmPar

2530      .cmdAssetNew.Enabled = False
2540      .cmdLocNew.Enabled = False

2550      blnJustAdd = True
2560      lngRecsCur = frmSub.RecCnt  ' ** Form Function: frmJournal_Columns_sub.
2570      If lngRecsCur > 0& Then
2580        Select Case IsNull(frmSub.accountno)
            Case True
2590          blnAddAndSet = False
2600        Case False
2610          Select Case IsNull(frmSub.journaltype)
              Case True
2620            blnAddAndSet = False
2630          Case False
2640            strThisJType = frmSub.journaltype
2650            Select Case strThisJType
                Case "Misc.", "Paid", "Received"
2660              blnAddAndSet = True
2670            Case "Dividend", "Interest", "Deposit", "Purcahse", "Withdrawn", "Sold", "Liability (+)", "Liability (-)", "Cost Adj."
2680              blnAddAndSet = False
2690            End Select
2700          End Select
2710        End Select
2720      Else
2730        blnAddAndSet = False
2740      End If

2750      If blnJustAdd = True Then

2760        gblnSetFocus = True
2770        gblnMessage = True  ' ** If it comes back false, they canceled it.
2780        gstrCrtRpt_NetLoss = vbNullString  ' ** Borrowing this variable from the Court Reports.
2790        strDocName = "frmRecurringItems_Add_Misc"
2800        DoCmd.OpenForm strDocName, , , , acFormAdd, acDialog, FRM_PAR & "~" & strThisJType

2810        .frmJournal_Columns_Sub.SetFocus
2820        DoEvents
2830        Select Case gblnMessage
            Case True
              ' ** A new item has been added.
2840          frmSub.Recur_Name.Requery
2850          If gstrCrtRpt_NetLoss <> vbNullString And blnAddAndSet = True Then
2860            frmSub.Recur_Name = gstrCrtRpt_NetLoss
2870          End If
2880        Case False
              ' ** They canceled it.
2890        End Select

2900        If blnAddAndSet = True Then
2910          frmSub.Recur_Name.SetFocus
2920        Else
2930          DoCmd.SelectObject acForm, frmPar.Name, False
2940          .cmdRecurNew.SetFocus
2950        End If

2960        gblnMessage = False
2970        gstrCrtRpt_NetLoss = vbNullString

2980        .cmdAssetNew.Enabled = True
2990        .cmdLocNew.Enabled = True

3000      Else
3010        blnContinue = False
3020      End If  ' ** blnJustAdd.

3030    End With

EXITP:
3040    Exit Sub

ERRH:
3050    DoCmd.Hourglass False
3060    frmPar.cmdAssetNew.Enabled = True
3070    frmPar.cmdLocNew.Enabled = True
3080    Select Case ERR.Number
        Case Else
3090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3100    End Select
3110    Resume EXITP

End Sub

Public Sub JC_Frm_Refresh(frmSub As Access.Form)
' ** Called by:
' **   frmJournal_Columns:
' **     cmdRefresh_Click()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_Refresh"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strThisUser As String
        Dim lngTotalJrnlRecs As Long
        Dim lngUserRecsHere As Long, lngNonUserRecsHere As Long, lngUserRecsNotHere As Long, lngNonUserRecsNotHere As Long
        Dim msgResponse As VbMsgBoxResult, blnAskResponse As Boolean
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, lngTmp04 As Long
        Dim blnContinue As Boolean

3210    blnContinue = True

3220    Set dbs = CurrentDb
3230    With dbs

3240      Set rst = .OpenRecordset("journal", dbOpenDynaset, dbReadOnly)
3250      With rst
3260        If .BOF = True And .EOF = True Then
              ' ** Nothing to refresh!
3270          blnContinue = False
3280          lngTotalJrnlRecs = 0&
3290          MsgBox "There are no other currently pending transactions in the Journal.", _
                vbInformation + vbOKOnly, "Nothing To Do"
3300        Else
3310          .MoveLast
3320          lngTotalJrnlRecs = .RecordCount
3330        End If
3340        .Close
3350      End With

3360      If blnContinue = True Then

3370        strThisUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
3380        blnAskResponse = False

3390        Select Case gblnAdmin
            Case True
3400          Select Case gblnLocalData
              Case True
3410            Select Case gblnSingleUser
                Case True
                  ' ** If this user is an Admin, and this is a single-user environment,
                  ' ** and the data is local, we're free to check and replace everything.
                  ' ** NOT AVAILABLE!
3420              blnContinue = False
3430            Case False
                  ' ** If this user is an Admin, and the data
                  ' ** is local, but it's not single-user, ask.
3440              blnAskResponse = True
3450            End Select
3460          Case False
                ' ** If this user is an Admin, and the data is on a network, ask.
3470            blnAskResponse = True
3480          End Select
3490        Case False
              ' ** If this isn't an Admin, only check and replace this
              ' ** user's transactions, regardless of where the data is.
              ' ** NOT AVAILABLE!
3500          blnContinue = False
3510        End Select

3520      End If  ' ** blnContinue.

3530      If blnContinue = True Then

            ' ** Journal, linked to tblJournal_Column, by specified [usr].
3540        Set qdf = .QueryDefs("qryJournal_Columns_03a")
3550        With qdf.Parameters
3560          ![usr] = strThisUser
3570        End With
3580        Set rst = qdf.OpenRecordset
3590        With rst
3600          If .BOF = True And .EOF = True Then
3610            lngUserRecsHere = 0&
3620          Else
3630            .MoveLast
3640            lngUserRecsHere = .RecordCount
3650          End If
3660          .Close
3670        End With
            ' ** This user's recs already here, don't replace!

            ' ** Journal, linked to tblJournal_Column, not belonging to specified [usr].
3680        Set qdf = .QueryDefs("qryJournal_Columns_03b")
3690        With qdf.Parameters
3700          ![usr] = strThisUser
3710        End With
3720        Set rst = qdf.OpenRecordset
3730        With rst
3740          If .BOF = True And .EOF = True Then
3750            lngNonUserRecsHere = 0&
3760          Else
3770            .MoveLast
3780            lngNonUserRecsHere = .RecordCount
3790          End If
3800          .Close
3810        End With
            ' ** Replace non-user recs already here?

            ' ** Journal, not in tblJournal_Column, by specified [usr].
3820        Set qdf = .QueryDefs("qryJournal_Columns_04a")
3830        With qdf.Parameters
3840          ![usr] = strThisUser
3850        End With
3860        Set rst = qdf.OpenRecordset
3870        With rst
3880          If .BOF = True And .EOF = True Then
3890            lngUserRecsNotHere = 0&
3900          Else
3910            .MoveLast
3920            lngUserRecsNotHere = .RecordCount
3930          End If
3940          .Close
3950        End With
            ' ** Add new recs, no question...

            ' ** Journal, not in tblJournal_Column, not belonging to specified [usr].
3960        Set qdf = .QueryDefs("qryJournal_Columns_04b")
3970        With qdf.Parameters
3980          ![usr] = strThisUser
3990        End With
4000        Set rst = qdf.OpenRecordset
4010        With rst
4020          If .BOF = True And .EOF = True Then
4030            lngNonUserRecsNotHere = 0&
4040          Else
4050            .MoveLast
4060            lngNonUserRecsNotHere = .RecordCount
4070          End If
4080          .Close
4090        End With
            ' ** Add new recs, no question...

4100        If blnAskResponse = True Then
4110          If lngNonUserRecsHere = 0& And lngUserRecsNotHere = 0& And lngNonUserRecsNotHere = 0& Then
4120            If lngUserRecsHere > 0& Then
4130              MsgBox "There are no committed transactions in the Journal other than your own.", _
                    vbInformation + vbOKOnly, "Nothing To Do"
4140            Else
                  ' ** Should've been caught at the outset!
4150              MsgBox "There are no other currently pending transactions in the Journal to refresh.", _
                    vbInformation + vbOKOnly, "Nothing To Do"
4160            End If
4170            msgResponse = vbCancel
4180          Else
4190            If lngNonUserRecsHere > 0& Then
                  ' ** For example:
                  ' **   In addition to 5 new transactions, there are 2 that you already see. Do you wish to refresh those as well?
4200              If lngUserRecsNotHere > 0& Or lngNonUserRecsNotHere > 0& Then
4210                lngTmp04 = (lngUserRecsNotHere + lngNonUserRecsNotHere)
4220                If lngTmp04 > 1& Then strTmp01 = "s" Else strTmp01 = vbNullString
4230                If lngNonUserRecsHere > 1& Then
4240                  strTmp02 = "are": strTmp03 = "those"
4250                Else
4260                  strTmp02 = "is": strTmp03 = "that"
4270                End If
4280                msgResponse = MsgBox("In addition to " & CStr(lngUserRecsNotHere + lngNonUserRecsNotHere) & " " & _
                      "new transaction" & strTmp01 & "," & vbCrLf & _
                      "there " & strTmp02 & " " & CStr(lngNonUserRecsHere) & " that you already see." & vbCrLf & _
                      "Do you wish to refresh " & strTmp03 & " as well?" & vbCrLf & vbCrLf & _
                      "Any changes you may have made in " & strTmp03 & " will be lost." & vbCrLf & vbCrLf & _
                      "Your own transactions will not be affected.", _
                      vbQuestion + vbYesNoCancel, "Refresh Non-User Transactions")
                    ' ** vbYes replaces non-user trans already here, along with adding the new ones.
                    ' ** vbNo only adds the new ones.
                    ' ** vbCancel aborts process.
4290              Else
4300                If lngNonUserRecsHere > 1& Then strTmp03 = "those" Else strTmp03 = "that"
4310                msgResponse = MsgBox("There are no new transactions, and " & CStr(lngNonUserRecsHere) & " " & _
                      "that you already see." & vbCrLf & _
                      "Do you wish to refresh " & strTmp03 & "?" & vbCrLf & vbCrLf & _
                      "Any changes you may have made in " & strTmp03 & " will be lost." & vbCrLf & vbCrLf & _
                      "Your own transactions will not be affected.", _
                      vbQuestion + vbYesNoCancel, "Refresh Non-User Transactions")
                    ' ** vbYes replaces non-user trans already here, along with adding the new ones.
                    ' ** vbNo only adds the new ones.
                    ' ** vbCancel aborts process.
4320              End If
4330            Else
                  ' ** Add new records.
4340              msgResponse = vbNo  ' ** Don't replace recs that weren't here anyway!
4350            End If
4360          End If
4370          If msgResponse <> vbYes And msgResponse <> vbNo Then
4380            blnContinue = False
4390          End If
4400        End If  ' ** blnAskResponse.

4410      End If  ' ** blnContinue.

4420      If blnContinue = True Then

4430        If msgResponse = vbYes Then
              ' ** Delete tblJournal_Column, not belonging to specified [usr].
4440          Set qdf = .QueryDefs("qryJournal_Columns_02e")
4450          With qdf.Parameters
4460            ![usr] = strThisUser
4470          End With
4480          qdf.Execute
4490        End If

            ' ** Append qryJournal_Columns_07 (qryJournal_Columns_06 (Journal, not
            ' ** in tblJournal_Column), with add'l fields) to tblJournal_Column.
4500        Set qdf = .QueryDefs("qryJournal_Columns_08")
4510        qdf.Execute

4520      End If  ' ** blnContinue.

4530      .Close
4540    End With  ' ** dbs.

4550    With frmSub
4560      If blnContinue = True Then
4570        .Requery
4580        .MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
4590      End If
4600    End With

EXITP:
4610    Set rst = Nothing
4620    Set qdf = Nothing
4630    Set dbs = Nothing
4640    Exit Sub

ERRH:
4650    Select Case ERR.Number
        Case Else
4660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4670    End Select
4680    Resume EXITP

End Sub

Public Sub JC_Frm_Dividend(lngRecCnt As Long, lngJrnlColID_Max As Long)
' ** Called by:
' **   frmJournal_Columns:
' **     cmdSpecPurp_Div_Map_Click()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_Dividend"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim varTmp00 As Variant

4710    DoCmd.Hourglass True
4720    DoEvents

4730    Set dbs = CurrentDb
4740    With dbs
          ' ** Empty Journal Map.
4750      Set qdf = .QueryDefs("qryJournal_Columns_30_04")
4760      qdf.Execute
4770      .Close
4780    End With

4790    JC_Frm_LoadMapDropdowns  ' ** Procedure: Below.

        ' ** This thread can track through 4 additional forms before returning here!
        ' **   frmMap_Div
        ' **   frmMap_Div_Detail
        ' **   frmMap_Reinvest_DivInt_Price
        ' **   frmMap_Reinvest_DivInt_Detail

4800    DoEvents

4810    JC_Frm_Map_Reset  ' ** Procedure: Below.

4820    DoEvents

4830    lngRecsCur = lngRecCnt
4840    If lngRecsCur > 0& Then
4850      lngJrnlColID_Max = DMax("[JrnlCol_ID]", "tblJournal_Column")
4860    Else
4870      lngJrnlColID_Max = -1&
4880    End If

4890    If gstrFormQuerySpec = vbNullString Then gstrFormQuerySpec = FRM_PAR

4900    strDocName = "frmMap_Div"
4910    DoCmd.OpenForm strDocName, , , , , , FRM_PAR

4920    If gblnGoToReport = True Then
4930      DoCmd.Hourglass True  ' ** Make sure it's still running.
4940      DoEvents
4950      varTmp00 = DMax("[JrnlCol_ID]", "tbljournal_Column")  ' ** Save this so we can delete any fake Journal records.
4960      Select Case IsNull(varTmp00)
          Case True
4970        glngTaxCode_Distribution = 0&  ' ** Borrowing this variable from the Court Reports.
4980      Case False
4990        glngTaxCode_Distribution = varTmp00
5000      End Select
5010      Forms(FRM_PAR).Header_lbl.Visible = True
5020      Forms(FRM_PAR).GoToReport_arw_mapdiv_img.Visible = False
5030      Forms(FRM_PAR).GoToReport_lbl_mapdiv_img.Visible = False
5040      Forms(FRM_PAR).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
5050      DoEvents
5060      Forms(strDocName).TimerInterval = 100&
5070    End If

EXITP:
5080    Set qdf = Nothing
5090    Set dbs = Nothing
5100    Exit Sub

ERRH:
5110    DoCmd.Hourglass False
5120    Select Case ERR.Number
        Case Else
5130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5140    End Select
5150    Resume EXITP

End Sub

Public Sub JC_Frm_Interest(lngRecCnt As Long, lngJrnlColID_Max As Long)
' ** Called by:
' **   frmJournal_Columns:
' **     cmdSpecPurp_Int_Map_Click()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_Interest"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim varTmp00 As Variant

5210    DoCmd.Hourglass True
5220    DoEvents

5230    Set dbs = CurrentDb
5240    With dbs
          ' ** Empty Journal Map.
5250      Set qdf = .QueryDefs("qryJournal_Columns_31_04")
5260      qdf.Execute
5270      .Close
5280    End With

5290    JC_Frm_LoadMapDropdowns  ' ** Procedure: Below.

        ' ** This thread can track through 4 additional forms before returning here!
        ' **   frmMap_Int
        ' **   frmMap_Int_Detail
        ' **   frmMap_Reinvest_DivInt_Price
        ' **   frmMap_Reinvest_DivInt_Detail

5300    DoEvents

5310    JC_Frm_Map_Reset  ' ** Procedure: Below.

5320    DoEvents

5330    lngRecsCur = lngRecCnt
5340    If lngRecsCur > 0& Then
5350      lngJrnlColID_Max = DMax("[JrnlCol_ID]", "tblJournal_Column")
5360    Else
5370      lngJrnlColID_Max = -1&
5380    End If

5390    If gstrFormQuerySpec = vbNullString Then gstrFormQuerySpec = FRM_PAR

5400    strDocName = "frmMap_Int"
5410    DoCmd.OpenForm strDocName, , , , , , FRM_PAR

5420    If gblnGoToReport = True Then
5430      DoCmd.Hourglass True  ' ** Make sure it's still running.
5440      DoEvents
5450      varTmp00 = DMax("[JrnlCol_ID]", "tbljournal_Column")  ' ** Save this so we can delete any fake Journal records.
5460      Select Case IsNull(varTmp00)
          Case True
5470        glngTaxCode_Distribution = 0&  ' ** Borrowing this variable from the Court Reports.
5480      Case False
5490        glngTaxCode_Distribution = varTmp00
5500      End Select
5510      Forms(FRM_PAR).cmdSpecPurp_Div_Map.Visible = True
5520      Forms(FRM_PAR).cmdSpecPurp_Div_Map_raised_img.Visible = True
5530      Forms(FRM_PAR).GoToReport_arw_mapint_img.Visible = False
5540      Forms(FRM_PAR).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
5550      DoEvents
5560      Forms(strDocName).TimerInterval = 100&
5570    End If

EXITP:
5580    Set qdf = Nothing
5590    Set dbs = Nothing
5600    Exit Sub

ERRH:
5610    DoCmd.Hourglass False
5620    Select Case ERR.Number
        Case Else
5630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5640    End Select
5650    Resume EXITP

End Sub

Public Sub JC_Frm_Split(lngRecCnt As Long, lngJrnlColID_Max As Long)
' ** Called by:
' **   frmJournal_Columns:
' **     cmdSpecPurp_Purch_MapSplit_Click()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_Split"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

5710    DoCmd.Hourglass True
5720    DoEvents

5730    Set dbs = CurrentDb
5740    With dbs
          ' ** Empty Journal Map.
5750      Set qdf = .QueryDefs("qryJournal_Columns_34_04")
5760      qdf.Execute
5770      .Close
5780    End With

5790    JC_Frm_LoadMapDropdowns  ' ** Procedure: Below.

        ' ** This thread can track through 2 additional forms before returning here!
        ' **   frmMap_Split
        ' **   frmMap_Split_Detail

5800    DoEvents

5810    JC_Frm_Map_Reset  ' ** Procedure: Below.

5820    DoEvents

5830    lngRecsCur = lngRecCnt
5840    If lngRecsCur > 0& Then
5850      lngJrnlColID_Max = DMax("[JrnlCol_ID]", "tblJournal_Column")
5860    Else
5870      lngJrnlColID_Max = -1&
5880    End If

5890    If gstrFormQuerySpec = vbNullString Then gstrFormQuerySpec = FRM_PAR

5900    strDocName = "frmMap_Split"
5910    DoCmd.OpenForm strDocName, , , , , , FRM_PAR

5920    If gblnGoToReport = True Then
5930      DoCmd.Hourglass True  ' ** Make sure it's still running.
5940      DoEvents
5950      Forms(FRM_PAR).cmdSpecPurp_Sold_PaidTotal.Visible = True
5960      Forms(FRM_PAR).cmdSpecPurp_Sold_PaidTotal_raised_img_dis.Visible = True
5970      Forms(FRM_PAR).GoToReport_arw_mapsplit_img.Visible = False
5980      Forms(FRM_PAR).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
5990      DoEvents
6000      Forms(strDocName).TimerInterval = 100&
6010    End If

EXITP:
6020    Set qdf = Nothing
6030    Set dbs = Nothing
6040    Exit Sub

ERRH:
6050    DoCmd.Hourglass False
6060    Select Case ERR.Number
        Case Else
6070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6080    End Select
6090    Resume EXITP

End Sub

Public Sub JC_Frm_Paid(frmSub As Access.Form)
' ** Called by:
' **   frmJournal_Columns:
' **     cmdSpecPurp_Sold_PaidTotal_Click()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_Paid"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strJrnlUser As String, lngJrnlColID As Long, strAccountNo As String
        Dim curPaidTot As Currency, lngAssetNo As Long
        Dim strTmp01 As String
        Dim blnContinue As Boolean

6110    blnContinue = True
6120    curPaidTot = 0@
6130    lngJrnlColID = 0&

6140    strJrnlUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.

6150    With frmSub

          ' ** Check the current record.
6160      If .journaltype = "Sold" Then
6170        If .posted = False Then
              ' ** Proceed.
6180        Else
6190          blnContinue = False
6200          MsgBox "This requires an uncommitted Sold transaction to proceed.", vbInformation + vbOKOnly, "Entry Required"
6210        End If
6220      Else
6230        Set rst = .RecordsetClone
6240        lngJrnlColID = JC_Msc_Find_JType("Sold", "False", rst)   ' ** Module Function: modJrnlCol_Misc.
6250        Set rst = Nothing
6260        If lngJrnlColID > 0& Then
6270          .MoveRec 0, lngJrnlColID  ' ** Form Procedure: frmJournal_Columns_Sub.
6280        Else
6290          blnContinue = False
6300          MsgBox "This requires an uncommitted Sold transaction to proceed.", vbInformation + vbOKOnly, "Entry Required"
6310        End If
6320      End If

6330      If blnContinue = True Then

6340        Set dbs = CurrentDb
6350        With dbs
6360          If strJrnlUser = "TAAdmin" Or strJrnlUser = "Superuser" Then
                ' ** Journal, grouped and summed, for 'Paid', all.
6370            Set qdf = .QueryDefs("qryJournal_Columns_14b")
6380          Else
                ' ** tblJournal_Column, grouped and summed, for 'Paid', by specified [jusr].
6390            Set qdf = .QueryDefs("qryJournal_Columns_14a")
6400            With qdf.Parameters
6410              ![jusr] = frmSub.journal_USER
6420            End With
6430          End If
6440          Set rst = qdf.OpenRecordset
6450          With rst
6460            If .BOF = True And .EOF = True Then
6470              curPaidTot = 0@
6480            Else
6490              .MoveFirst
6500              curPaidTot = CCur(Abs(Nz(![ICash], 0) + Nz(![PCash], 0)))
6510            End If
6520            .Close
6530          End With
6540          .Close
6550        End With

6560        If curPaidTot = 0@ Then
6570          blnContinue = False
6580          Select Case strJrnlUser
              Case "TAAdmin", "Superuser"
6590            MsgBox "There are no available 'Paid' transactions in the Journal.", vbInformation + vbOKOnly, "Nothing To Do"
6600          Case Else
6610            MsgBox "There are no available 'Paid' transactions for User " & strJrnlUser & " in the Journal.", vbInformation + vbOKOnly, "Nothing To Do"
6620          End Select
6630          .shareface = 0#
6640          .ICash = 0@
6650          .PCash = 0@
6660          .Cost = 0@
6670        Else
6680          If IsNull(.accountno) = False And IsNull(.accountno2) = False Then
6690            strAccountNo = .accountno
6700          Else
6710            blnContinue = False
6720            MsgBox "Please enter a valid account number to continue.", vbInformation + vbOKOnly, "Entry Required"
6730          End If
6740        End If  ' ** curPaidTot.

6750        If blnContinue = True Then

6760          .shareface = CDbl(curPaidTot)
6770          .PCash = curPaidTot
6780          If IsNull(.Cost) = False Then
6790            If .Cost <> 0@ Then
6800              .Cost = 0@
6810            End If
6820          End If

6830          Set dbs = CurrentDb
6840          With dbs
                ' ** Account, just taxlot, by specified [actno].
6850            Set qdf = .QueryDefs("qryJournal_Columns_14c")
6860            With qdf.Parameters
6870              ![actno] = strAccountNo
6880            End With
6890            Set rst = qdf.OpenRecordset
6900            With rst
6910              .MoveFirst
6920              If IsNull(![taxlot]) = False Then
6930                strTmp01 = ![taxlot]
6940              Else
6950                strTmp01 = vbNullString
6960              End If
6970              .Close
6980            End With
6990            If strTmp01 <> vbNullString Then
7000              If Val(strTmp01) > 0 Then
7010                lngAssetNo = Val(strTmp01)
7020              Else
7030                lngAssetNo = 0&
7040              End If
7050            End If
7060            If lngAssetNo > 0& Then
                  ' ** MasterAsset, just description, by specified [astno].
7070              Set qdf = .QueryDefs("qryJournal_Columns_14d")
7080              With qdf.Parameters
7090                ![astno] = lngAssetNo
7100              End With
7110              Set rst = qdf.OpenRecordset
7120              With rst
7130                If .BOF = True And .EOF = True Then
                      ' ** Oops! something's messed up!
7140                  strTmp01 = vbNullString
7150                Else
7160                  .MoveFirst
7170                  strTmp01 = ![description]
7180                End If
7190                .Close
7200              End With
7210            End If
7220            .Close
7230          End With

7240          If lngAssetNo > 0& And strTmp01 <> vbNullString Then
7250            .assetno = lngAssetNo
7260            .assetno_description = strTmp01
7270          End If

7280          DoCmd.SelectObject acForm, .Parent.Name, False
7290          .Parent.frmJournal_Columns_Sub.SetFocus
7300          .shareface.SetFocus

7310        End If  ' ** blnContinue.

7320      End If

7330    End With

EXITP:
7340    Set rst = Nothing
7350    Set qdf = Nothing
7360    Set dbs = Nothing
7370    Exit Sub

ERRH:
7380    DoCmd.Hourglass False
7390    Select Case ERR.Number
        Case Else
7400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7410    End Select
7420    Resume EXITP

End Sub

Public Sub JC_Frm_LTCG(lngRecCnt As Long, lngJrnlColID_Max As Long)
' ** Called by:
' **   frmJournal_Columns:
' **     cmdSpecPurp_Misc_MapLTCG_Click()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_LTCG"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim varTmp00 As Variant

7510    DoCmd.Hourglass True
7520    DoEvents

7530    Set dbs = CurrentDb
7540    With dbs
          ' ** Empty Journal Map.
7550      Set qdf = .QueryDefs("qryJournal_Columns_35_04")
7560      qdf.Execute
7570      .Close
7580    End With

7590    JC_Frm_LoadMapDropdowns  ' ** Procedure: Below.

        ' ** This thread can track through 4 additional forms before returning here!
        ' **   frmMap_Rec
        ' **   frmMap_Rec_Detail
        ' **   frmMap_Reinvest_Rec_Price
        ' **   frmMap_Reinvest_Rec_Detail

7600    DoEvents

7610    JC_Frm_Map_Reset  ' ** Module Procedure: modJrnlCol_Misc.

7620    DoEvents

7630    lngRecsCur = lngRecCnt
7640    If lngRecsCur > 0& Then
7650      lngJrnlColID_Max = DMax("[JrnlCol_ID]", "tblJournal_Column")
7660    Else
7670      lngJrnlColID_Max = -1&
7680    End If

7690    If gstrFormQuerySpec = vbNullString Then gstrFormQuerySpec = FRM_PAR

7700    strDocName = "frmMap_Rec"
7710    DoCmd.OpenForm strDocName, , , , , , FRM_PAR

7720    If gblnGoToReport = True Then
7730      DoCmd.Hourglass True  ' ** Make sure it's still running.
7740      DoEvents
7750      varTmp00 = DMax("[JrnlCol_ID]", "tbljournal_Column")  ' ** Save this so we can delete any fake Journal records.
7760      Select Case IsNull(varTmp00)
          Case True
7770        glngTaxCode_Distribution = 0&  ' ** Borrowing this variable from the Court Reports.
7780      Case False
7790        glngTaxCode_Distribution = varTmp00
7800      End Select
7810      Forms(FRM_PAR).cmdSpecPurp_Purch_MapSplit.Visible = False
7820      Forms(FRM_PAR).cmdSpecPurp_Purch_MapSplit_raised_img.Visible = False
7830      Forms(FRM_PAR).GoToReport_arw_maprec_img.Visible = False
7840      Forms(FRM_PAR).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
7850      DoEvents
7860      Forms(strDocName).TimerInterval = 100&
7870    End If

EXITP:
7880    Set qdf = Nothing
7890    Set dbs = Nothing
7900    Exit Sub

ERRH:
7910    DoCmd.Hourglass False
7920    Select Case ERR.Number
        Case Else
7930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7940    End Select
7950    Resume EXITP

End Sub

Public Sub JC_Frm_LTCL(lngRecCnt As Long, lngJrnlColID_Max As Long)
' ** Called by:
' **   frmJournal_Columns:
' **     cmdSpecPurp_Misc_MapLTCG_Click()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_LTCL"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim varTmp00 As Variant

8010    DoCmd.Hourglass True
8020    DoEvents

8030    Set dbs = CurrentDb
8040    With dbs
          ' ** Empty Journal Map.
8050      Set qdf = .QueryDefs("qryJournal_Columns_35_04")
8060      qdf.Execute
8070      .Close
8080    End With

8090    JC_Frm_LoadMapDropdowns  ' ** Procedure: Below.

        ' ** This thread can track through 2 additional forms before returning here!
        ' **   frmMap_Misc_LTCL
        ' **   frmMap_Misc_LTCL_Detail

8100    DoEvents

8110    JC_Frm_Map_Reset  ' ** Procedure: Below.

8120    DoEvents

8130    lngRecsCur = lngRecCnt
8140    If lngRecsCur > 0& Then
8150      lngJrnlColID_Max = DMax("[JrnlCol_ID]", "tblJournal_Column")
8160    Else
8170      lngJrnlColID_Max = -1&
8180    End If

8190    If gstrFormQuerySpec = vbNullString Then gstrFormQuerySpec = FRM_PAR

8200    strDocName = "frmMap_Misc_LTCL"
8210    DoCmd.OpenForm strDocName, , , , , , FRM_PAR

8220    If gblnGoToReport = True Then
8230      DoCmd.Hourglass True  ' ** Make sure it's still running.
8240      DoEvents
8250      varTmp00 = DMax("[JrnlCol_ID]", "tbljournal_Column")  ' ** Save this so we can delete any fake Journal records.
8260      Select Case IsNull(varTmp00)
          Case True
8270        glngTaxCode_Distribution = 0&  ' ** Borrowing this variable from the Court Reports.
8280      Case False
8290        glngTaxCode_Distribution = varTmp00
8300      End Select
8310      Forms(FRM_PAR).cmdSpecPurp_Purch_MapSplit.Visible = False
8320      Forms(FRM_PAR).cmdSpecPurp_Purch_MapSplit_raised_img.Visible = False
8330      Forms(FRM_PAR).GoToReport_arw_maprec_img.Visible = False
8340      Forms(FRM_PAR).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
8350      DoEvents
8360      Forms(strDocName).TimerInterval = 100&
8370    End If

EXITP:
8380    Set qdf = Nothing
8390    Set dbs = Nothing
8400    Exit Sub

ERRH:
8410    DoCmd.Hourglass False
8420    Select Case ERR.Number
        Case Else
8430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8440    End Select
8450    Resume EXITP

End Sub

Public Sub JC_Frm_STCGL(lngRecCnt As Long, lngJrnlColID_Max As Long)
' ** Called by:
' **   frmJournal_Columns:
' **     cmdSpecPurp_Misc_MapLTCG_Click()

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_STCGL"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim varTmp00 As Variant

8510    DoCmd.Hourglass True
8520    DoEvents

8530    Set dbs = CurrentDb
8540    With dbs
          ' ** Empty Journal Map.
8550      Set qdf = .QueryDefs("qryJournal_Columns_35_04")
8560      qdf.Execute
8570      .Close
8580    End With

8590    JC_Frm_LoadMapDropdowns  ' ** Procedure: Below.

        ' ** This thread can track through 2 additional forms before returning here!
        ' **   frmMap_Misc_STCGL
        ' **   frmMap_Misc_STCGL_Detail

8600    DoEvents

8610    JC_Frm_Map_Reset  ' ** Module Procedure: modJrnlCol_Misc.

8620    DoEvents

8630    lngRecsCur = lngRecCnt
8640    If lngRecsCur > 0& Then
8650      lngJrnlColID_Max = DMax("[JrnlCol_ID]", "tblJournal_Column")
8660    Else
8670      lngJrnlColID_Max = -1&
8680    End If

8690    If gstrFormQuerySpec = vbNullString Then gstrFormQuerySpec = FRM_PAR

8700    strDocName = "frmMap_Misc_STCGL"
8710    DoCmd.OpenForm strDocName, , , , , , FRM_PAR

8720    If gblnGoToReport = True Then
8730      DoCmd.Hourglass True  ' ** Make sure it's still running.
8740      DoEvents
8750      varTmp00 = DMax("[JrnlCol_ID]", "tbljournal_Column")  ' ** Save this so we can delete any fake Journal records.
8760      Select Case IsNull(varTmp00)
          Case True
8770        glngTaxCode_Distribution = 0&  ' ** Borrowing this variable from the Court Reports.
8780      Case False
8790        glngTaxCode_Distribution = varTmp00
8800      End Select
8810      Forms(FRM_PAR).cmdSpecPurp_Purch_MapSplit.Visible = False
8820      Forms(FRM_PAR).cmdSpecPurp_Purch_MapSplit_raised_img.Visible = False
8830      Forms(FRM_PAR).GoToReport_arw_maprec_img.Visible = False
8840      Forms(FRM_PAR).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
8850      DoEvents
8860      Forms(strDocName).TimerInterval = 100&
8870    End If

EXITP:
8880    Set qdf = Nothing
8890    Set dbs = Nothing
8900    Exit Sub

ERRH:
8910    DoCmd.Hourglass False
8920    Select Case ERR.Number
        Case Else
8930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8940    End Select
8950    Resume EXITP

End Sub

Public Sub JC_Frm_Map_Return(lngJrnlColID_Max As Long, blnNextRec As Boolean, blnFromZero As Boolean)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_Map_Return"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frmPar As Access.Form, frmSub As Access.Form
        Dim lngJrnlColID_Empty As Long
        Dim lngRecs As Long
        Dim blnContinue As Boolean, blnAddNew As Boolean
        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim lngX As Long, lngE As Long

9010    blnContinue = True

9020    Set frmPar = Forms("frmJournal_Columns")

9030    With frmPar

9040      If lngJrnlColID_Max <> 0& Then

9050        DoCmd.Hourglass True
9060        DoEvents

9070        If .opgFilter <> .opgFilter_optAll.OptionValue Then
9080          .opgFilter = .opgFilter_optAll.OptionValue
9090          .opgFilter_AfterUpdate  ' ** Procedure: Above.
9100          DoEvents
9110        End If

9120        Set frmSub = .frmJournal_Columns_Sub.Form

9130        Set dbs = CurrentDb
9140        With dbs

9150          lngRecs = 0&
9160          lngJrnlColID_Empty = 0&

              ' ** tblJournal_Column, grouped, just empty records, with cnt, Max(JrnlCol_ID).
9170          Set qdf = .QueryDefs("qryJournal_Columns_40")
9180          Set rst = qdf.OpenRecordset
9190          With rst
9200            If .BOF = True And .EOF = True Then
                  ' ** No empties.
9210            Else
9220              .MoveFirst
9230              lngRecs = ![cnt]  ' ** May be 0.
9240              If lngRecs > 0& Then
9250                lngJrnlColID_Empty = ![JrnlCol_ID]
9260              End If
9270            End If
9280            .Close
9290          End With

9300          If lngRecs > 1& Then
                ' ** Delete tblJournal_Column, just extra empties, by specified [jcolid].
9310            Set qdf = .QueryDefs("qryJournal_Columns_40")
9320            With qdf.Parameters
9330              ![jcolid] = lngJrnlColID_Empty  ' ** JrnlCol_ID <> lngJrnlColID_Empty.
9340            End With
9350            qdf.Execute
9360          End If

9370          lngRecs = 0&
9380          If lngJrnlColID_Max > 0& Then  ' ** Previous records exist in tblJournal_Column.
                ' ** tblJournal_Column, by specified [jcolid].
9390            Set qdf = .QueryDefs("qryJournal_Columns_42a")
9400            With qdf.Parameters
9410              ![jcolid] = lngJrnlColID_Max  ' ** JrnlCol_ID > lngJrnlColID_Max.
9420            End With
9430            Set rst = qdf.OpenRecordset
9440            With rst
9450              If .BOF = True And .EOF = True Then
9460                blnContinue = False
9470              Else
9480                .MoveLast
9490                lngRecs = .RecordCount
9500                .MoveFirst
9510              End If
9520            End With
9530          ElseIf lngJrnlColID_Max = -1& Then  ' ** tblJournal_Column was empty.
                ' ** tblJournal_Column, all records.
9540            Set qdf = .QueryDefs("qryJournal_Columns_42b")
9550            Set rst = qdf.OpenRecordset
9560            With rst
9570              If .BOF = True And .EOF = True Then
9580                blnContinue = False
9590              Else
9600                .MoveLast
9610                lngRecs = .RecordCount
9620                .MoveFirst
9630              End If
9640            End With
9650          End If

9660          If blnContinue = True Then

9670            lngTmp01 = 0&: lngTmp02 = 0&
9680            With rst
9690              .MoveFirst
9700              For lngX = 1& To lngRecs
9710                If ![IsEmpty] = False Then
9720                  lngNewRecs = lngNewRecs + 1&
9730                  lngE = lngNewRecs - 1&
9740                  ReDim Preserve arr_varNewRec(N_ELEMS, lngE)
9750                  arr_varNewRec(N_ID, lngE) = ![JrnlCol_ID]
9760                  arr_varNewRec(N_CMTD, lngE) = CBool(False)
9770                  arr_varNewRec(N_JTYP, lngE) = ![journaltype]
9780                  arr_varNewRec(N_CHK, lngE) = Null
9790                  If lngTmp01 = 0& Then
9800                    lngTmp01 = ![JrnlCol_ID]
9810                  ElseIf ![JrnlCol_ID] < lngTmp01 Then
9820                    lngTmp01 = ![JrnlCol_ID]
9830                  End If
9840                  If lngTmp02 = 0& Then
9850                    lngTmp02 = ![JrnlCol_ID]
9860                  ElseIf ![JrnlCol_ID] > lngTmp02 Then
9870                    lngTmp02 = ![JrnlCol_ID]
9880                  End If
9890                End If
9900                If lngX < lngRecs Then .MoveNext
9910              Next
9920              .Close
9930            End With  ' ** rst.
9940            Set rst = Nothing
9950            DoEvents

                ' ** Reinvests already have their parent's JrnlCol_ID in CheckNum.
                ' ** If we put the parent's own JrnlCol_ID in their CheckNum,
                ' ** we'll be able to link them that way.

                ' ** Update qryJournal_Columns_55_02 (tblJournal_Column, just
                ' ** CheckNum <> Null, with CheckNum_new, by specified [jcolid1], [jcolid2]).
9960            Set qdf = dbs.QueryDefs("qryJournal_Columns_55_03")
9970            With qdf.Parameters
9980              ![jcolid1] = lngTmp01
9990              ![jcolid2] = lngTmp02
10000           End With
10010           qdf.Execute
10020           Set qdf = Nothing
10030           DoEvents

10040           blnAddNew = False
10050           For lngX = 0& To (lngNewRecs - 1&)
10060             frmSub.MoveRec 0, arr_varNewRec(N_ID, lngX)  ' ** Form Procedure: frmJournal_Columns_Sub.
10070             DoEvents
10080             If lngX = (lngNewRecs - 1&) And lngJrnlColID_Empty = 0& Then
                    ' ** When coming from JC_Frm_Map_Return(), somehow the last
                    ' ** Map entry gets unposted, and the empty entry disappears!
                    ' ** So, I'm no longer telling it to add on the last entry.
                    'blnAddNew = True  ' ** Only add new, empty record after last one.
10090             End If
                  ' ** Ideally, this would only apply to a Reverse Stock Split Withdrawn,
                  ' ** but I think that would involve a lot of other specific variables.
                  ' ** So, for now, I'm giving them a blanket pardon.
10100             frmSub.WarnZeroCash_GetSet False, True  ' ** Form Function: frmJournal_Columns_Sub.
                  ' ** Update tblJournal_Column, by specified [jcolid], [pstd].
10110             Set qdf = dbs.QueryDefs("qryJournal_Columns_55_01")
10120             With qdf.Parameters
10130               ![jcolid] = arr_varNewRec(N_ID, lngX)
10140               ![pstd] = True
10150             End With
10160             qdf.Execute
10170             Set qdf = Nothing
10180             arr_varNewRec(N_CMTD, lngX) = CBool(True)
10190             frmSub.Refresh
10200             DoEvents
                  '########################################
                  '########################################
                  'DOESN'T WORK!!
                  '########################################
                  '########################################
10210             CommitRec frmSub, blnNextRec, blnFromZero, True, blnAddNew  ' ** Module Function: modJrnlCol_Recs.
10220             DoEvents
10230           Next

                'IF ANY OF THESE ARE REINVESTMENTS...
                'AT THIS MOMENT, BOTH SIDES HAVE PARENTS' JrnlCol_ID IN CHECKNUM!

                ' ** Append qryJournal_Columns_55_07 (tblJournal_MiscSold_Staging, linked to
                ' ** qryJournal_Columns_55_04 (Journal, linked to itself, with CheckNum_new), with
                ' ** journal_id_new, jrnlms_ref_id_new, parent record) to tblJournal_MiscSold, parent record.
10240           Set qdf = .QueryDefs("qryJournal_Columns_55_09")
10250           qdf.Execute
10260           Set qdf = Nothing
10270           DoEvents

                ' ** Append qryJournal_Columns_55_08 (tblJournal_MiscSold_Staging, linked to
                ' ** qryJournal_Columns_55_04 (Journal, linked to itself, with CheckNum_new), with
                ' ** journal_id_new, jrnlms_ref_id_new, child record) to tblJournal_MiscSold, child record.
10280           Set qdf = .QueryDefs("qryJournal_Columns_55_10")
10290           qdf.Execute
10300           Set qdf = Nothing
10310           DoEvents

                ' ** Update qryJournal_Columns_55_05 (Journal, with DLookups() to
                ' ** qryJournal_Columns_55_04 (Journal, linked to itself, with CheckNum_new)).
10320           Set qdf = .QueryDefs("qryJournal_Columns_55_06")
10330           qdf.Execute
10340           Set qdf = Nothing
10350           DoEvents

                ' ** Update qryMap_Reinvest_02_15 (Journal, linked to
                ' ** tblJournal_Map_Staging3, just parents), for CheckNum = Null.
10360           Set qdf = .QueryDefs("qryMap_Reinvest_02_16")
10370           qdf.Execute
10380           Set qdf = Nothing
10390           DoEvents

                ' ** Empty tblJournal_Map_Staging3.
10400           Set qdf = .QueryDefs("qryJournal_Columns_32_03")
10410           qdf.Execute
10420           Set qdf = Nothing
10430           DoEvents

                ' ** Empty tblJournal_MiscSold_Staging.
10440           Set qdf = .QueryDefs("qryJournal_Columns_32_06")
10450           qdf.Execute
10460           Set qdf = Nothing
10470           DoEvents

10480           lngNewRecs = 0&
10490           ReDim arr_varNewRec(N_ELEMS, 0)

10500           With frmSub
10510             .Requery
10520             If lngJrnlColID_Empty > 0& Then
10530               .MoveRec 0, lngJrnlColID_Empty  ' ** Form Procedure: frmJournal_Columns_Sub.
10540             Else
10550               .MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_Sub.
10560             End If
10570             .RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.
10580           End With  ' ** frmSub.

10590           DoCmd.SelectObject acForm, frmPar.Name, False
10600           frmPar.Controls(frmSub.Name).SetFocus

10610           If lngJrnlColID_Empty = 0& Then
                  'frmSub.AddRec  ' ** Form Function: frmJournal_Columns_Sub.
10620             frmSub.AddRec_Send  ' ** Form Procedure: frmJournal_Columns_Sub.
10630           End If

10640         Else
10650           rst.Close
10660         End If  ' ** blnContinue.

10670         .Close
10680       End With  ' ** dbs.

10690     Else
10700       blnContinue = False
10710     End If  ' ** lngJrnlColID_Max.

10720   End With  ' ** Me.

10730   DoCmd.Hourglass False

EXITP:
10740   Set frmSub = Nothing
10750   Set rst = Nothing
10760   Set qdf = Nothing
10770   Set dbs = Nothing
10780   Exit Sub

ERRH:
10790   DoCmd.Hourglass False
10800   Select Case ERR.Number
        Case Else
10810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10820   End Select
10830   Resume EXITP

End Sub

Public Sub JC_Frm_Map_Reset()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_Map_Reset"

10910   lngNewRecs = 0&
10920   ReDim arr_varNewRec(N_ELEMS, 0)

EXITP:
10930   Exit Sub

ERRH:
10940   Select Case ERR.Number
        Case Else
10950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10960   End Select
10970   Resume EXITP

End Sub

Public Sub JC_Frm_Clear()

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_Clear"

11010   lngNewRecs = 0&
11020   ReDim arr_varNewRec(N_ELEMS, 0)

EXITP:
11030   Exit Sub

ERRH:
11040   Select Case ERR.Number
        Case Else
11050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11060   End Select
11070   Resume EXITP

End Sub

Public Sub JC_Frm_TaxLot_Form(frmSub As Access.Form, blnGoingThere As Boolean, blnToTaxLot As Boolean, blnGoingToReport As Boolean, blnGoneToReport As Boolean, blnNextRec As Boolean, blnFromZero As Boolean, datPostingDate As Date, strSaveMoveCtl As String, blnNoMove As Boolean, lngNewJrnlColID As Long, lngGTR_ID As Long, blnGTR_Emblem As Boolean, blnGTR_NoAdd As Boolean)
' ** Called by:
' **   frmJournal_Columns:
' **     Form_Timer()
' **   frmJournal_Columns_Sub:
' **     shareface_Exit()
' **     pcash_Exit()
' **

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_TaxLot_Form"

        Dim lngX As Long
        Dim intRetVal As Integer, blnRetVal As Boolean

11110   With frmSub
11120     Select Case blnGoingThere
          Case True

11130       DoCmd.Hourglass True
11140       DoEvents

11150       blnToTaxLot = True

11160       lngNewRecs = 0&
11170       ReDim arr_varNewRec(N_ELEMS, 0)

11180       .Parent.JrnlCol_ID = .JrnlCol_ID
11190       .Parent.ToTaxLot = 1&  ' ** 0 = Nothing; 1 = To Tax Lot; 2 = OK Single; 3 = OK Multi; -4 = Cancel Return; PLUS...
            ' ** .Parent.TaxLotFrom set by calling procedure.
11200       intRetVal = OpenLotInfoForm(False, THIS_NAME)  ' ** Module Function: modPurchaseSold.  'OK!
            ' ** Return Values:
            ' **    0 OK.
            ' **   -1 Input missing.
            ' **   -2 No holdings.
            ' **   -3 Insufficient holdings.
            ' **   -4 Zero shares.
            ' **   -9 Data problem.

11210       If gblnGoToReport = True Then
11220         blnGoneToReport = True
11230         lngGTR_ID = .JrnlCol_ID
11240         .Parent.GTRStuff 5, True, lngGTR_ID  ' ** Form Procedure: frmJournal_Columns.
11250         Forms("frmJournal_Columns_TaxLot").TimerInterval = 50&
11260         blnGoingToReport = False
11270         blnGTR_Emblem = False
11280         .GoToReport_arw_jcol_bge_img.Visible = False
11290         .GoToReport_arw_jcol_blu2_img.Visible = False
11300         .Parent.cmdDelete.SetFocus
11310       End If

11320       If intRetVal <> 0 Then
              ' ** A non-zero here indicates a problem within the OpenLotInfoForm() function,
              ' ** and it never went to the form.
11330         .Parent.ToTaxLot = CLng(intRetVal)  ' ** Just 'cause.
11340         gblnSetFocus = True
11350         .Parent.TimerInterval = 250&
11360       End If

11370     Case False

11380       DoCmd.Hourglass True
11390       DoEvents
11400       If lngNewRecs > 0& Then
11410         blnRetVal = False
11420         .Parent.CommitNoClose_Set True  ' ** Form Procedure: frmJournal_Columns.
11430         For lngX = 0& To (lngNewRecs - 1&)
11440           .MoveRec 0, arr_varNewRec(N_ID, lngX)  ' ** Form Procedure: frmJournal_Columns_Sub.
11450           DoEvents
11460           If lngX = (lngNewRecs - 1&) Then
11470             blnRetVal = True  ' ** Only add new, empty record after last one.
11480           End If
11490           CommitRec frmSub, blnNextRec, blnFromZero, True, blnRetVal  ' ** Module Function: modJrnlCol_Recs.
11500           DoEvents
11510         Next
              ' ** If not on a new record, add one!
11520         If IsNull(.journaltype) = False Then
11530           AddRec frmSub, blnGTR_NoAdd, datPostingDate, strSaveMoveCtl, blnNoMove, lngNewJrnlColID  ' ** Module Procedure: modJrnlCol_Recs.
11540         End If
11550         .Parent.CommitNoClose_Set False  ' ** Form Procedure: frmJournal_Columns.
11560       End If

11570       lngNewRecs = 0&
11580       ReDim arr_varNewRec(N_ELEMS, 0)

11590       DoCmd.Hourglass False

11600     End Select
11610   End With

EXITP:
11620   Exit Sub

ERRH:
11630   frmSub.Parent.CommitNoClose_Set False  ' ** Form Procedure: frmJournal_Columns.
11640   Select Case ERR.Number
        Case Else
11650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11660   End Select
11670   Resume EXITP

End Sub

Public Sub JC_Frm_LoadMapDropdowns()

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Frm_LoadMapDropdowns"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnContinue As Boolean

11710   blnContinue = True

11720   Set dbs = CurrentDb
11730   Set rst = dbs.OpenRecordset("tmpAccount", dbOpenDynaset, dbReadOnly)
11740   If rst.BOF = True And rst.EOF = True Then
          ' ** Continue.
11750   Else
11760     blnContinue = False
11770   End If
11780   rst.Close
11790   Set rst = Nothing
11800   DoEvents

11810   If blnContinue = True Then
          ' ** Empty tmpAccount.
11820     Set qdf = dbs.QueryDefs("qryMap_09")
11830     qdf.Execute
11840     Set qdf = Nothing
11850     DoEvents
          ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
          ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
          ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
          ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
          ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
          ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
          ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
          ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
11860     Set qdf = dbs.QueryDefs("qryMap_10")
11870     qdf.Execute
11880     Set qdf = Nothing
11890     DoEvents
11900   End If

11910   dbs.Close
11920   Set dbs = Nothing
11930   DoEvents

EXITP:
11940   Set rst = Nothing
11950   Set qdf = Nothing
11960   Set dbs = Nothing
11970   Exit Sub

ERRH:
11980   Select Case ERR.Number
        Case Else
11990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12000   End Select
12010   Resume EXITP

End Sub
