Attribute VB_Name = "modCheckPrintFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modCheckPrintFuncs"

'VGC 09/09/2017: CHANGES!

' ** cmbSortBy combo box constants:
Private Const CBX_SORT_FID  As Integer = 0  ' ** fld_id
'Private Const CBX_SORT_DESC As Integer = 1  ' ** fld_description
Private Const CBX_SORT_FNAM As Integer = 2  ' ** fld_name
'Private Const CBX_SORT_CID  As Integer = 3  ' ** ctl_id_parent
Private Const CBX_SORT_CNAM As Integer = 4  ' ** ctl_name_parent

' ** Array: arr_varItem().
Private Const I_CNAM As Integer = 0
Private Const I_TOP  As Integer = 1
Private Const I_LFT  As Integer = 2

Private lngTpp As Long
' **

Public Sub DisplaySet(blnShow As Boolean, frm As Access.Form)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "DisplaySet"

110     With frm
          ' ** ckgDisplay_opt03_CheckCnt and ckgDisplay_opt04_LastCheckNum always remain invisible.
120       Select Case blnShow
          Case True
130         .ckgDisplay_lbl.Visible = True
140         .ckgDisplay_lbl_dim.Visible = False
150         .ckgDisplay_lbl_dim_hi.Visible = False
160         .ckgDisplay_opt01_AccountNo.Enabled = True
170         .ckgDisplay_opt01_AccountNo_lbl.Visible = True
180         .ckgDisplay_opt01_AccountNo_lbl_dim.Visible = False
190         .ckgDisplay_opt01_AccountNo_lbl_dim_hi.Visible = False
200         .ckgDisplay_opt02_ShortName.Enabled = True
210         .ckgDisplay_opt02_ShortName_lbl.Visible = True
220         .ckgDisplay_opt02_ShortName_lbl_dim.Visible = False
230         .ckgDisplay_opt02_ShortName_lbl_dim_hi.Visible = False
240         .ckgDisplay_opt05_Payee.Enabled = True
250         .ckgDisplay_opt05_Payee_lbl.Visible = True
260         .ckgDisplay_opt05_Payee_lbl_dim.Visible = False
270         .ckgDisplay_opt05_Payee_lbl_dim_hi.Visible = False
280         .ckgDisplay_opt06_BankName.Enabled = True
290         .ckgDisplay_opt06_BankName_lbl.Visible = True
300         .ckgDisplay_opt06_BankName_lbl_dim.Visible = False
310         .ckgDisplay_opt06_BankName_lbl_dim_hi.Visible = False
320         .ckgDisplay_opt07_BankAcctNum.Enabled = True
330         .ckgDisplay_opt07_BankAcctNum_lbl.Visible = True
340         .ckgDisplay_opt07_BankAcctNum_lbl_dim.Visible = False
350         .ckgDisplay_opt07_BankAcctNum_lbl_dim_hi.Visible = False
360         .ckgDisplay_opt08_CheckAmount.Enabled = True
370         .ckgDisplay_opt08_CheckAmount_lbl.Visible = True
380         .ckgDisplay_opt08_CheckAmount_lbl_dim.Visible = False
390         .ckgDisplay_opt08_CheckAmount_lbl_dim_hi.Visible = False
400         .opgPayeeFont.Enabled = True
410         .opgShow.Enabled = True
420         .opgShow_lbl.Visible = True
430         .opgShow_lbl_dim.Visible = False
440         .opgShow_lbl_dim_hi.Visible = False
450         .opgShow_optAccount_lbl.Visible = True
460         .opgShow_optAccount_lbl_dim.Visible = False
470         .opgShow_optAccount_lbl_dim_hi.Visible = False
480         .opgShow_optAll_lbl.Visible = True
490         .opgShow_optAll_lbl_dim.Visible = False
500         .opgShow_optAll_lbl_dim_hi.Visible = False
510       Case False
520         .ckgDisplay_lbl.Visible = False
530         .ckgDisplay_lbl_dim.Visible = True
540         .ckgDisplay_lbl_dim_hi.Visible = True
550         .ckgDisplay_opt01_AccountNo.Enabled = False
560         .ckgDisplay_opt01_AccountNo_lbl.Visible = False
570         .ckgDisplay_opt01_AccountNo_lbl_dim.Visible = True
580         .ckgDisplay_opt01_AccountNo_lbl_dim_hi.Visible = True
590         .ckgDisplay_opt02_ShortName.Enabled = False
600         .ckgDisplay_opt02_ShortName_lbl.Visible = False
610         .ckgDisplay_opt02_ShortName_lbl_dim.Visible = True
620         .ckgDisplay_opt02_ShortName_lbl_dim_hi.Visible = True
630         .ckgDisplay_opt05_Payee.Enabled = False
640         .ckgDisplay_opt05_Payee_lbl.Visible = False
650         .ckgDisplay_opt05_Payee_lbl_dim.Visible = True
660         .ckgDisplay_opt05_Payee_lbl_dim_hi.Visible = True
670         .ckgDisplay_opt06_BankName.Enabled = False
680         .ckgDisplay_opt06_BankName_lbl.Visible = False
690         .ckgDisplay_opt06_BankName_lbl_dim.Visible = True
700         .ckgDisplay_opt06_BankName_lbl_dim_hi.Visible = True
710         .ckgDisplay_opt07_BankAcctNum.Enabled = False
720         .ckgDisplay_opt07_BankAcctNum_lbl.Visible = False
730         .ckgDisplay_opt07_BankAcctNum_lbl_dim.Visible = True
740         .ckgDisplay_opt07_BankAcctNum_lbl_dim_hi.Visible = True
750         .ckgDisplay_opt08_CheckAmount.Enabled = False
760         .ckgDisplay_opt08_CheckAmount_lbl.Visible = False
770         .ckgDisplay_opt08_CheckAmount_lbl_dim.Visible = True
780         .ckgDisplay_opt08_CheckAmount_lbl_dim_hi.Visible = True
790         .opgPayeeFont.Enabled = False
800         .opgShow.Enabled = False
810         .opgShow_lbl.Visible = False
820         .opgShow_lbl_dim.Visible = True
830         .opgShow_lbl_dim_hi.Visible = True
840         .opgShow_optAccount_lbl.Visible = False
850         .opgShow_optAccount_lbl_dim.Visible = True
860         .opgShow_optAccount_lbl_dim_hi.Visible = True
870         .opgShow_optAll_lbl.Visible = False
880         .opgShow_optAll_lbl_dim.Visible = True
890         .opgShow_optAll_lbl_dim_hi.Visible = True
900       End Select
910     End With

EXITP:
920     Exit Sub

ERRH:
930     Select Case ERR.Number
        Case Else
940       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
950     End Select
960     Resume EXITP

End Sub

Public Function AlreadyPrinted(frm As Access.Form) As Boolean
' ** Mark checks as printed if a check number is present.
' ** Show a status message.

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "AlreadyPrinted"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strMsg As String
        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim blnRetVal As Boolean

1010    With frm

1020      blnRetVal = True

1030      Set dbs = CurrentDb
          ' ** qryPrintChecks_05_30_05 (Union of qryPrintChecks_05_30_03 (qryPrintChecks_05_30_01
          ' ** (Journal, just PrintCheck = True, CheckNum = Null), grouped, with cnt_nochk),
          ' ** qryPrintChecks_05_30_04 (qryPrintChecks_05_30_02 (Journal, just PrintCheck = True,
          ' ** CheckNum <> Null), grouped, with cnt_chk)), grouped, to get 1 record.
1040      Set qdf = dbs.QueryDefs("qryPrintChecks_05_30_06")
1050      Set rst = qdf.OpenRecordset
1060      With rst
1070        If .BOF = True And .EOF = True Then
              ' ** Journal is empty!
1080          lngTmp01 = 0
1090          lngTmp02 = 0
1100          blnRetVal = False
1110        Else
1120          .MoveFirst
1130          lngTmp01 = ![cnt_chk]
1140          lngTmp02 = ![cnt_nochk]
1150        End If
1160        .Close
1170      End With
1180      Set rst = Nothing
1190      Set qdf = Nothing

1200      .AlreadyPrinted_lbl1.Visible = False
1210      .AlreadyPrinted_lbl2a.Visible = False
1220      .AlreadyPrinted_lbl2b.Visible = False
1230      .AlreadyPrinted_lbl3a.Visible = False
1240      .AlreadyPrinted_lbl3b.Visible = False

1250      If blnRetVal = True Then
1260        If lngTmp01 > 0 And lngTmp02 = 0 Then
              ' ** All have been printed.
1270          .AlreadyPrinted_lbl2a.Visible = True
1280          .AlreadyPrinted_lbl2b.Visible = True
1290          .ChkStat = "Yes"
1300        ElseIf lngTmp01 > 0 And lngTmp02 > 0 Then
              ' ** Some printed, some not.
1310          .AlreadyPrinted_lbl3a.Visible = True
1320          .AlreadyPrinted_lbl3b.Visible = True
1330          .ChkStat = "Mixed"
1340        ElseIf lngTmp01 = 0 And lngTmp02 > 0 Then
              ' ** Nothing's been printed.
1350          .AlreadyPrinted_lbl1.Visible = True
1360          .ChkStat = "No"
1370        End If
1380      End If

1390    End With

EXITP:
1400    Set rst = Nothing
1410    Set qdf = Nothing
1420    Set dbs = Nothing
1430    AlreadyPrinted = blnRetVal
1440    Exit Function

ERRH:
1450    blnRetVal = False
1460    Select Case ERR.Number
        Case Else
1470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1480    End Select
1490    Resume EXITP

End Function

Public Sub CheckVoid(strAccountNo As String, blnAll As Boolean, lngChecks As Long, lngChkVoid_Set As Long, frm As Access.Form)
' ** Handle all steps when confirm is anything other than 'Yes, all printed successfully'.

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "CheckVoid"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim intMode As Integer
        Dim strMsg As String, strDocName As String, strBankName As String, strBankAcctNum As String
        Dim lngRecs As Long
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean, blnFound As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long, lngY As Long, lngE As Long

        ' ** Array: arr_varCheck() [garr_varPrintRpt()].
        Const C_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const C_ACTNO As Integer = 0
        Const C_BID   As Integer = 1
        Const C_BANK  As Integer = 2
        Const C_ACCT  As Integer = 3

1510    With frm

1520      blnContinue = True

1530      DoCmd.Hourglass True  ' ** Make sure the hourglass is still going.
1540      DoEvents

1550      If strAccountNo = "All" Or strAccountNo = vbNullString Then
1560        intMode = 2
1570      Else
1580        intMode = 1
1590      End If

1600      varTmp00 = DMax("[chkvoid_set]", "tblCheckVoid")
1610      Select Case IsNull(varTmp00)
          Case True
1620        lngChkVoid_Set = 1&
1630      Case False
1640        lngChkVoid_Set = (varTmp00 + 1&)
1650      End Select

1660      Select Case blnAll
          Case True
            ' ** All checks are to be voided.

1670        Set dbs = CurrentDb

1680        Select Case intMode
            Case 1
              ' ** By accountno.

              ' ************************************************
              ' ** Before saving the voided check numbers,
              ' ** check for the bank name and account number.
              ' ************************************************

1690          glngTaxCode_Distribution = 0&  ' ** Borrowing this variable.

              ' ** qryPrintChecks_05_03_01 (qryPrintChecks_05_03 (Journal, just PrintCheck = True,
              ' ** by specified [actno], [cvset]), just Bank_Name = Null, Bank_AccountNumber = Null),
              ' ** linked to tblCheckBank, just '{Unknown Bank}', '000000000000000000'.
1700          Set qdf = dbs.QueryDefs("qryPrintChecks_05_03_04")
1710          With qdf.Parameters
1720            ![actno] = strAccountNo
1730            ![cvset] = lngChkVoid_Set
1740          End With
1750          Set rst1 = qdf.OpenRecordset
1760          If rst1.BOF = True And rst1.EOF = True Then
                ' ** Good, bank and account number present and accounted for.
                ' ** This is just in Account, not necessarily in tblCheckBank.
1770            rst1.Close
1780            Set rst1 = Nothing
1790            Set qdf = Nothing
1800            DoEvents

                ' *************************************************
                ' ** Save bank names and numbers in tblCheckBank.
                ' *************************************************

                ' ** qryPrintChecks_05_03_05 (qryPrintChecks_05_03_03 (qryPrintChecks_05_03_01
                ' ** (qryPrintChecks_05_03 (Journal, just PrintCheck = True, by specified [actno], [cvset]),
                ' ** just Bank_Name = Null, Bank_AccountNumber = Null), linked to tblCheckBank),
                ' ** grouped by accountno, just good bank names and numbers, with cnt_chks),
                ' ** not in tblCheckBank, as new tblCheckBank records.
1810            Set qdf = dbs.QueryDefs("qryPrintChecks_05_03_06")
1820            With qdf.Parameters
1830              ![actno] = strAccountNo
1840              ![cvset] = lngChkVoid_Set
1850            End With
1860            Set rst1 = qdf.OpenRecordset
1870            If rst1.BOF = True And rst1.EOF = True Then  ' ** THIS JUST TELLS US THAT ACCOUNT HAS A BANK & NUMBER,
                  ' ** Nothing new.                          ' ** NOT WHETHER IT'S IN tblCheckBank OR NOT!
1880              rst1.Close
1890              Set rst1 = Nothing
1900              Set qdf = Nothing
1910              DoEvents
                  ' ** qryPrintChecks_05_03_08 (qryPrintChecks_05_03 (Journal, just PrintCheck = True,
                  ' ** by specified [actno], [cvset]), just Bank_Name <> Null and Bank_AccountNumber <> Null),
                  ' ** not in  tblCheckBank, as new tblCheckBank records.
1920              Set qdf = dbs.QueryDefs("qryPrintChecks_05_03_09")
1930              With qdf.Parameters
1940                ![actno] = strAccountNo
1950                ![cvset] = lngChkVoid_Set
1960              End With
1970              Set rst1 = qdf.OpenRecordset
1980              If rst1.BOF = True And rst1.EOF = True Then
                    ' ** Really, really nothing new.
1990                rst1.Close
2000                Set rst1 = Nothing
2010                Set qdf = Nothing
2020                DoEvents
                    ' ** qryPrintChecks_05_03 (Journal, just PrintCheck = True, by specified [actno], [cvset]), linked to tblCheckBank.
2030                Set qdf = dbs.QueryDefs("qryPrintChecks_05_03_11")
2040                With qdf.Parameters
2050                  ![actno] = strAccountNo
2060                  ![cvset] = lngChkVoid_Set
2070                End With
2080                Set rst1 = qdf.OpenRecordset
2090                rst1.MoveFirst
2100                glngTaxCode_Distribution = rst1![chkbank_id]  ' ** Borrowing this variable.
2110                rst1.Close
2120                Set rst1 = Nothing
2130                Set qdf = Nothing
2140                DoEvents
2150              Else
2160                rst1.MoveFirst
2170                strBankName = rst1![chkbank_name]
2180                strBankAcctNum = rst1![chkbank_acctnum]
2190                rst1.Close
2200                Set rst1 = Nothing
2210                Set qdf = Nothing
2220                DoEvents
                    ' ** Append qryPrintChecks_05_03_09 (qryPrintChecks_05_03_08 (qryPrintChecks_05_03
                    ' ** (Journal, just PrintCheck = True, by specified [actno], [cvset]), just
                    ' ** Bank_Name <> Null and Bank_AccountNumber <> Null), not in  tblCheckBank,
                    ' ** as new tblCheckBank records) to tblCheckBank.
2230                Set qdf = dbs.QueryDefs("qryPrintChecks_05_03_10")
2240                With qdf.Parameters
2250                  ![actno] = strAccountNo
2260                  ![cvset] = lngChkVoid_Set
2270                End With
2280                qdf.Execute
2290                Set qdf = Nothing
2300                DoEvents
2310                varTmp00 = DLookup("[chkbank_id]", "tblCheckBank", "[accountno]  '" & strAccountNo & "' And " & _
                      "[chkbank_name] = '" & strBankName & "' And [chkbank_acctnum] = '" & strBankAcctNum & "' ")
2320                glngTaxCode_Distribution = Nz(varTmp00, 0)  ' ** Borrowing this variable.
2330                DoEvents
2340              End If
2350            Else
2360              rst1.MoveFirst
2370              strBankName = rst1![chkbank_name]
2380              strBankAcctNum = rst1![chkbank_acctnum]
2390              rst1.Close
2400              Set rst1 = Nothing
2410              Set qdf = Nothing
2420              DoEvents
                  ' ** Append qryPrintChecks_05_03_06 (qryPrintChecks_05_03_05 (qryPrintChecks_05_03_03
                  ' ** (qryPrintChecks_05_03_01 (qryPrintChecks_05_03 (Journal, just PrintCheck = True,
                  ' ** by specified [actno], [cvset]), just Bank_Name = Null, Bank_AccountNumber = Null), linked
                  ' ** to tblCheckBank), grouped by accountno, just good bank names and numbers, with
                  ' ** cnt_chks), not in tblCheckBank, as new tblCheckBank records) to tblCheckBank.
2430              Set qdf = dbs.QueryDefs("qryPrintChecks_05_03_07")
2440              With qdf.Parameters
2450                ![actno] = strAccountNo
2460                ![cvset] = lngChkVoid_Set
2470              End With
2480              qdf.Execute
2490              Set qdf = Nothing
2500              DoEvents
2510              varTmp00 = DLookup("[chkbank_id]", "tblCheckBank", "[accountno]  '" & strAccountNo & "' And " & _
                    "[chkbank_name] = '" & strBankName & "' And [chkbank_acctnum] = '" & strBankAcctNum & "' ")
2520              glngTaxCode_Distribution = Nz(varTmp00, 0)  ' ** Borrowing this variable.
2530              DoEvents
2540            End If
2550            DoEvents

2560          Else
                ' ** This accountno's bank name and number are missing.
2570            rst1.Close
2580            Set rst1 = Nothing
2590            Set qdf = Nothing

                ' ***************************************************************
                ' ** Require the user to supply a bank name and account number.
                ' ***************************************************************

2600            .Modal = False
2610            gblnMessage = True  ' ** For emergency closing of the form.
2620            strDocName = "frmRpt_Checks_Bank1"
2630            DoCmd.OpenForm strDocName, , , , , acDialog, frm.Name & "~" & strAccountNo & "~" & Format(blnAll, "True/False")
2640            DoEvents
2650            .Modal = True

2660            If gblnMessage = True Then

                  ' *****************************************************
                  ' ** Form puts new bank and number into tblCheckBank,
                  ' ** and sets glngTaxCode_Distribution.
                  ' *****************************************************

2670              DoCmd.Hourglass True  ' ** Make sure it's still running.
2680              DoEvents

2690            Else
2700              blnContinue = False
2710            End If  ' ** gblnMessage.

2720          End If

2730          If blnContinue = True Then

                ' ******************************************************
                ' ** Confirm our check count matches what was printed.
                ' ******************************************************

                ' ** qryPrintChecks_05_04 (qryPrintChecks_05_03 (Journal, just PrintCheck = True,
                ' ** by specified [actno], [cvset]), linked to qryPrintChecks_05_03_03 (qryPrintChecks_05_03_01
                ' ** (qryPrintChecks_05_03 (Journal, just PrintCheck = True, by specified [actno], [cvset]),
                ' ** just Bank_Name = Null, Bank_AccountNumber = Null), linked to tblCheckBank)),
                ' ** grouped, to make sure only 1 record per check, for accountno, by specified [bnkid].
2740            Set qdf = dbs.QueryDefs("qryPrintChecks_05_05")
2750            With qdf.Parameters
2760              ![actno] = strAccountNo
2770              ![cvset] = lngChkVoid_Set
2780              ![bnkid] = glngTaxCode_Distribution  ' ** Borrowing this variable.
2790            End With
2800            Set rst1 = qdf.OpenRecordset
2810            With rst1
2820              If .BOF = True And .EOF = True Then
                    ' ** Shouldn't happen.
2830                blnContinue = False
2840                Beep
2850                DoCmd.Hourglass False
2860                MsgBox "There are no checks to print.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
2870              Else
2880                .MoveLast
2890                lngRecs = .RecordCount
2900                If lngRecs <> lngChecks Then
2910                  strMsg = CStr(lngChecks) & IIf(lngChecks = 1&, " check was ", " checks were ") & "printed, but " & _
                        IIf(lngRecs < lngChecks, "only ", vbNullString) & CStr(lngRecs) & IIf(lngRecs = 1&, " check was ", " checks were ") & "found."
2920                  strMsg = strMsg & vbCrLf & vbCrLf & "Proceed anyway?"
2930                  Beep
2940                  DoCmd.Hourglass False
2950                  msgResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Check Count Discrepancy")
2960                  If msgResponse <> vbYes Then
2970                    blnContinue = False
2980                  Else
2990                    DoCmd.Hourglass True
3000                    DoEvents
3010                  End If
3020                End If
3030              End If
3040              .Close
3050            End With
3060            Set rst1 = Nothing
3070            Set qdf = Nothing
3080            DoEvents

3090          End If  ' ** blnContinue.

3100          If blnContinue = True Then

3110            DoCmd.Hourglass True  ' ** Make sure it's still running.
3120            DoEvents

                ' ***********************************************
                ' ** Save voided check numbers to tblCheckVoid.
                ' ***********************************************

                ' ** Append qryPrintChecks_05_05 (qryPrintChecks_05_04 (qryPrintChecks_05_03 (Journal,
                ' ** just PrintCheck = True, by specified [actno], [cvset]), linked to qryPrintChecks_05_03_03
                ' ** (qryPrintChecks_05_03_01 (qryPrintChecks_05_03 (Journal, just PrintCheck = True,
                ' ** by specified [actno], [cvset]), just Bank_Name = Null, Bank_AccountNumber = Null),
                ' ** linked to tblCheckBank)), grouped, to make sure only 1 record per check,
                ' ** for accountno, by specified [bnkid]) to tblCheckVoid.
3130            Set qdf = dbs.QueryDefs("qryPrintChecks_05_06")
3140            With qdf.Parameters
3150              ![actno] = strAccountNo
3160              ![cvset] = lngChkVoid_Set
3170              ![bnkid] = glngTaxCode_Distribution  ' ** Borrowing this variable.
3180            End With
3190            qdf.Execute dbFailOnError
3200            Set qdf = Nothing
3210            DoEvents

3220            DoCmd.Hourglass False
3230            MsgBox CStr(lngRecs) & " voided check " & IIf(lngRecs = 1&, "number has", "numbers have") & " been recorded.", _
                  vbInformation + vbOKOnly, "Voided Checks Saved"
3240            DoCmd.Hourglass True
3250            DoEvents

                ' ***********************************************
                ' ** Existing print routine resets the Journal.
                ' ***********************************************

3260          End If  ' ** blnContinue.

3270        Case 2
              ' ** All accounts.

              ' ************************************************
              ' ** Before saving the voided check numbers,
              ' ** check for the bank name and account number.
              ' ************************************************

              ' ** qryPrintChecks_05_07_01 (qryPrintChecks_05_07 (Journal, just PrintCheck = True,
              ' ** for all, by specified [cvset]), just Bank_Name = Null, Bank_AccountNumber = Null), linked to tblCheckBank,
              ' ** just '{Unknown Bank}', '000000000000000000'.
3280          Set qdf = dbs.QueryDefs("qryPrintChecks_05_07_04")
3290          With qdf.Parameters
3300            ![cvset] = lngChkVoid_Set
3310          End With
3320          Set rst1 = qdf.OpenRecordset
3330          If rst1.BOF = True And rst1.EOF = True Then
                ' ** Good, everyone's got a bank and account number.
3340            rst1.Close
3350            Set rst1 = Nothing
3360            Set qdf = Nothing

                ' *************************************************
                ' ** Save bank names and numbers in tblCheckBank.
                ' *************************************************

                '#################################################################
                'WHEN IT'S SINGLE CHK ACCT, ALL SHOULD GET SAME BANK AND NUMBER!
                '#################################################################

3370            DoCmd.Hourglass True  ' ** Make sure it's still running.
3380            DoEvents

                'BUT THEY'RE NOT ALL IN tblCheckBank!

                ' ** qryPrintChecks_05_07_05 (qryPrintChecks_05_07_03 (qryPrintChecks_05_07_01
                ' ** (qryPrintChecks_05_07 (Journal, just PrintCheck = True, for all, by specified [cvset]), just
                ' ** Bank_Name = Null, Bank_AccountNumber = Null), linked to tblCheckBank),
                ' ** grouped by accountno, just good bank names and numbers, with cnt_chks),
                ' ** not in tblCheckBank, as new tblCheckBank records.
3390            Set qdf = dbs.QueryDefs("qryPrintChecks_05_07_06")
3400            With qdf.Parameters
3410              ![cvset] = lngChkVoid_Set
3420            End With
3430            Set rst1 = qdf.OpenRecordset
3440            If rst1.BOF = True And rst1.EOF = True Then
                  ' ** Nothing new.
3450              rst1.Close
3460              Set rst1 = Nothing
3470              Set qdf = Nothing
                  'BECAUSE THEY'RE NOT NULL! THEY'RE FROM ACCOUNT, AND THEY'RE NOT IN tblCheckBank!
                  ' ** Append qryPrintChecks_05_07_10 (qryPrintChecks_05_07 (Journal, just PrintCheck = True,
                  ' ** for all, by specified [cvset]), not in qryPrintChecks_05_07_01 (qryPrintChecks_05_07
                  ' ** (Journal, just PrintCheck = True, for all, by specified [cvset]), just Bank_Name = Null,
                  ' ** Bank_AccountNumber = Null), grouped by accountno, just those that do have bank info),
                  ' ** not already there, to tblCheckBank.
3480              Set qdf = dbs.QueryDefs("qryPrintChecks_05_07_11")
3490              With qdf.Parameters
3500                ![cvset] = lngChkVoid_Set
3510              End With
3520              qdf.Execute
3530              Set qdf = Nothing
3540              DoEvents
3550            Else
3560              rst1.Close
3570              Set rst1 = Nothing
3580              Set qdf = Nothing
                  ' ** Append qryPrintChecks_05_07_06 (qryPrintChecks_05_07_05 (qryPrintChecks_05_07_03
                  ' ** (qryPrintChecks_05_07_01 (qryPrintChecks_05_07 (Journal, just PrintCheck = True,
                  ' ** for all, by specified [cvset]), just Bank_Name = Null, Bank_AccountNumber = Null), linked to tblCheckBank),
                  ' ** grouped by accountno, just good bank names and numbers, with cnt_chks),
                  ' ** not in tblCheckBank, as new tblCheckBank records) to tblCheckBank.
3590              Set qdf = dbs.QueryDefs("qryPrintChecks_05_07_07")
3600              With qdf.Parameters
3610                ![cvset] = lngChkVoid_Set
3620              End With
3630              qdf.Execute
3640              Set qdf = Nothing
3650            End If
3660            DoEvents

                ' *********************************************
                ' ** Get bank info for all of these accounts.
                ' *********************************************

                ' ** Using this borrowed array for all the bank info, including chkbank_id.
3670            glngPrintRpts = 0&
3680            ReDim garr_varPrintRpt(C_ELEMS, 0)

                'THIS ONLY PICKS UP WHAT IS IN tblCheckBank, AND MOST OF THEM AREN'T!
                ' ** qryPrintChecks_05_07 (Journal, just PrintCheck = True, for all, by specified [cvset]), linked to tblCheckBank.
3690            Set qdf = dbs.QueryDefs("qryPrintChecks_05_07_09")
3700            With qdf.Parameters
3710              ![cvset] = lngChkVoid_Set
3720            End With
3730            Set rst1 = qdf.OpenRecordset
3740            With rst1
3750              .MoveLast
3760              lngRecs = .RecordCount
3770              .MoveFirst
3780              For lngX = 1& To lngRecs
3790                glngPrintRpts = glngPrintRpts + 1&
3800                lngE = glngPrintRpts - 1&
3810                ReDim Preserve garr_varPrintRpt(C_ELEMS, lngE)
3820                garr_varPrintRpt(C_ACTNO, lngE) = ![accountno]
3830                garr_varPrintRpt(C_BID, lngE) = ![chkbank_id]
3840                garr_varPrintRpt(C_BANK, lngE) = ![Bank_Name]
3850                garr_varPrintRpt(C_ACCT, lngE) = ![Bank_AccountNumber]
3860                If lngX < lngRecs Then .MoveNext
3870              Next
3880              .Close
3890            End With
3900            Set rst1 = Nothing
3910            Set qdf = Nothing
3920            DoEvents

3930          Else
                ' ** Some accountno's bank name and number are missing.

3940            Set rst1 = Nothing
3950            Set qdf = Nothing

3960            .Modal = False
3970            gblnMessage = True  ' ** For emergency closing of these forms.
3980            Select Case gblnSeparateCheckingAccounts
                Case True
3990              strAccountNo = "All"
4000              strDocName = "frmRpt_Checks_Bank2"
4010              DoCmd.OpenForm strDocName, , , , , acDialog, frm.Name & "~" & strAccountNo & "~" & Format(blnAll, "True/False") & "~" & CStr(lngChkVoid_Set)
4020            Case False
4030              strAccountNo = "All"
4040              strDocName = "frmRpt_Checks_Bank1"
4050              DoCmd.OpenForm strDocName, , , , , acDialog, frm.Name & "~" & strAccountNo & "~" & Format(blnAll, "True/False")
4060            End Select
4070            DoEvents
4080            .Modal = True

4090            If gblnMessage = True Then

4100              DoCmd.Hourglass True  ' ** Make sure it's still running.
4110              DoEvents

                  ' *****************************************************
                  ' ** Form puts new bank and number into tblCheckBank.
                  ' *****************************************************

                  ' ** This borrowed array has all the bank info, including chkbank_id.
                  ' **   glngPrintRpts
                  ' **   garr_varPrintRpt()

                  ' ** Make sure array has ALL bank info, not just those supplied by frmRpt_Checks_Bank2.
4120              If glngPrintRpts < lngChecks Then
                    ' ** qryPrintChecks_05_07 (Journal, just PrintCheck = True, for all, by specified [cvset]),
                    ' ** not in qryPrintChecks_05_07_01 (qryPrintChecks_05_07 (Journal, just PrintCheck = True,
                    ' ** for all, by specified [cvset]), just Bank_Name = Null, Bank_AccountNumber = Null),
                    ' ** grouped by accountno, just those that do have bank info.
4130                Set qdf = dbs.QueryDefs("qryPrintChecks_05_07_10")
4140                With qdf.Parameters
4150                  ![cvset] = lngChkVoid_Set
4160                End With
4170                Set rst1 = qdf.OpenRecordset
4180                With rst1
4190                  If .BOF = True And .EOF = True Then
                        ' ** Well, I just don't know what to do next.
4200                  Else
4210                    .MoveLast
4220                    lngRecs = .RecordCount
4230                    .MoveFirst
4240                    For lngX = 1& To lngRecs
4250                      blnFound = False
4260                      For lngY = 0& To (glngPrintRpts - 1&)
4270                        If garr_varPrintRpt(C_ACTNO, lngY) = ![accountno] Then
4280                          blnFound = True
4290                          Exit For
4300                        End If
4310                      Next
4320                      If blnFound = False Then
4330                        glngPrintRpts = glngPrintRpts + 1&
4340                        lngE = glngPrintRpts - 1&
4350                        ReDim Preserve garr_varPrintRpt(C_ELEMS, lngE)
4360                        garr_varPrintRpt(C_ACTNO, lngE) = ![accountno]
4370                        garr_varPrintRpt(C_BID, lngE) = ![chkbank_id]  ' ** If this is NULL, we're screwed!
4380                        garr_varPrintRpt(C_BANK, lngE) = ![Bank_Name]
4390                        garr_varPrintRpt(C_ACCT, lngE) = ![Bank_AccountNumber]
4400                      End If
4410                      If lngX < lngRecs Then .MoveNext
4420                    Next
4430                  End If
4440                  .Close
4450                End With
4460                Set rst1 = Nothing
4470                Set qdf = Nothing
4480                DoEvents
4490              End If

4500            Else
4510              blnContinue = False
4520            End If  ' ** gblnMessage.

4530          End If

4540          If blnContinue = True Then

4550            DoCmd.Hourglass True  ' ** Make sure it's still running.
4560            DoEvents

                ' ******************************************************
                ' ** Confirm our check count matches what was printed.
                ' ******************************************************

                ' ** qryPrintChecks_05_08 (qryPrintChecks_05_07 (Journal, just PrintCheck = True, for all),
                ' ** linked to qryPrintChecks_05_07_03 (qryPrintChecks_05_07_01 (qryPrintChecks_05_07 (Journal,
                ' ** just PrintCheck = True, for all, by specified [cvset]), just Bank_Name = Null, Bank_AccountNumber = Null),
                ' ** linked to tblCheckBank)), grouped, to make sure only 1 record per check, for all.
4570            Set qdf = dbs.QueryDefs("qryPrintChecks_05_09")
4580            With qdf.Parameters
4590              ![cvset] = lngChkVoid_Set
4600            End With
4610            Set rst1 = qdf.OpenRecordset
4620            With rst1
4630              If .BOF = True And .EOF = True Then
                    ' ** Shouldn't happen.
4640                blnContinue = False
4650                Beep
4660                DoCmd.Hourglass False
4670                MsgBox "There are no checks to print.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
4680              Else
4690                .MoveLast
4700                lngRecs = .RecordCount
4710                If lngRecs <> lngChecks Then
4720                  strMsg = CStr(lngChecks) & IIf(lngChecks = 1&, " check was ", " checks were ") & "printed, but " & _
                        IIf(lngRecs < lngChecks, "only ", vbNullString) & CStr(lngRecs) & IIf(lngRecs = 1&, " check was ", " checks were ") & "found."
4730                  strMsg = strMsg & vbCrLf & vbCrLf & "Proceed anyway?"
4740                  Beep
4750                  DoCmd.Hourglass False
4760                  msgResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Check Count Discrepancy")
4770                  If msgResponse <> vbYes Then
4780                    blnContinue = False
4790                  Else
4800                    DoCmd.Hourglass True
4810                    DoEvents
4820                  End If
4830                End If
4840              End If
4850              .Close
4860            End With
4870            Set rst1 = Nothing
4880            Set qdf = Nothing

4890          End If  ' ** blnContinue.

4900          If blnContinue = True Then

4910            DoCmd.Hourglass True  ' ** Make sure it's still running.
4920            DoEvents

                ' ***********************************************
                ' ** Save voided check numbers to tblCheckVoid.
                ' ***********************************************

                ' ** This borrowed array has all the bank info, including chkbank_id.
                ' **   glngPrintRpts
                ' **   garr_varPrintRpt()

                ' ** qryPrintChecks_05_08 (qryPrintChecks_05_07 (Journal, just PrintCheck = True, for all),
                ' ** linked to qryPrintChecks_05_07_03 (qryPrintChecks_05_07_01 (qryPrintChecks_05_07
                ' ** (Journal, just PrintCheck = True, for all, by specified [cvset]), just Bank_Name = Null,
                ' ** Bank_AccountNumber = Null), linked to tblCheckBank)), grouped, to make sure
                ' ** only 1 record per check, for all.
4930            Set qdf = dbs.QueryDefs("qryPrintChecks_05_09")
4940            With qdf.Parameters
4950              ![cvset] = lngChkVoid_Set
4960            End With
4970            Set rst1 = qdf.OpenRecordset
4980            Set rst2 = dbs.OpenRecordset("tblCheckVoid", dbOpenDynaset, dbConsistent)
4990            If glngPrintRpts = 0& Then
                  ' ** This should not occur!
5000              MsgBox "Something bad happened!"
5010            Else
5020              With rst1
5030                .MoveLast
5040                lngRecs = .RecordCount
5050                .MoveFirst
5060                For lngX = 1& To lngRecs
5070                  glngTaxCode_Distribution = 0&: strBankName = vbNullString: strBankAcctNum = vbNullString
5080                  For lngY = 0& To (glngPrintRpts - 1&)
                        ' ** This should always hit!
5090                    If garr_varPrintRpt(C_ACTNO, lngY) = ![accountno] Then
5100                      glngTaxCode_Distribution = garr_varPrintRpt(C_BID, lngY)  ' ** Borrowing this variable.
5110                      strBankName = garr_varPrintRpt(C_BANK, lngY)
5120                      strBankAcctNum = garr_varPrintRpt(C_ACCT, lngY)
5130                      Exit For
5140                    End If
5150                  Next
5160                  rst2.AddNew
                      ' ** rst2![chkvoid_id] : AutoNumber.
5170                  rst2![chkbank_id] = glngTaxCode_Distribution
5180                  rst2![chkbank_name] = strBankName
5190                  rst2![chkbank_acctnum] = strBankAcctNum
5200                  rst2![chkvoid_chknum] = ![CheckNum]
5210                  rst2![chkvoid_date] = Date
5220                  rst2![accountno] = ![accountno]
5230                  rst2![transdate] = ![transdate]
5240                  rst2![chkvoid_payee] = ![RecurringItem]
5250                  rst2![chkvoid_amount] = CCur(Abs(![ICash] + ![PCash]))
5260                  rst2![curr_id] = ![curr_id]
5270                  rst2![chkvoid_set] = ![chkvoid_set]
5280                  rst2![Journal_ID] = ![ID]
5290                  rst2![chkvoid_datemodified] = Now()
5300                  rst2.Update
5310                  If lngX < lngRecs Then .MoveNext
5320                Next
5330              End With
5340            End If
5350            rst1.Close
5360            rst2.Close
5370            Set rst1 = Nothing
5380            Set rst2 = Nothing
5390            Set qdf = Nothing

5400            DoCmd.Hourglass False
5410            MsgBox CStr(lngRecs) & " voided check numbers have been recorded.", vbInformation + vbOKOnly, "Void Checks Saved"
5420            DoCmd.Hourglass True
5430            DoEvents

                ' ***********************************************
                ' ** Existing print routine resets the Journal.
                ' ***********************************************

5440          End If  ' ** blnContinue.

5450        End Select

5460        dbs.Close

5470      Case False
            ' ** Allow user to choose checks to void.

5480        Set dbs = CurrentDb

5490        Select Case intMode
            Case 1
              ' ** By accountno.

5500          DoCmd.Hourglass True  ' ** Make sure it's still running.
5510          DoEvents

              ' ** Empty tblCheckVoid_Staging.
5520          Set qdf = dbs.QueryDefs("qryPrintChecks_05_12_05")
5530          qdf.Execute
5540          Set qdf = Nothing
5550          DoEvents

              ' ** Append qryPrintChecks_05_12_04 (qryPrintChecks_05_12 (Journal,
              ' ** just PrintCheck = True, by specified [actno], [cvset]), linked to
              ' ** qryPrintChecks_05_12_03 (qryPrintChecks_05_12_02 (qryPrintChecks_05_12_01
              ' ** (tblCheckBank, grouped by accountno, with Max(chkbank_datemodified)),
              ' ** linked back to tblCheckBank, grouped by accountno, with Max(chkbank_id)),
              ' ** linked back to tblCheckBank)) to tblCheckVoid_Staging.
5560          Set qdf = dbs.QueryDefs("qryPrintChecks_05_12_06")
5570          With qdf.Parameters
5580            ![actno] = strAccountNo
5590            ![cvset] = lngChkVoid_Set
5600          End With
5610          qdf.Execute
5620          Set qdf = Nothing
5630          DoEvents

5640        Case 2
              ' ** All accounts.

              ' ** Empty tblCheckVoid_Staging.
5650          Set qdf = dbs.QueryDefs("qryPrintChecks_05_13_05")
5660          qdf.Execute
5670          Set qdf = Nothing
5680          DoEvents

              ' ** qryPrintChecksAppend qryPrintChecks_05_13_04 (qryPrintChecks_05_13 (Journal,
              ' ** just PrintCheck = True, for all, by specified [cvset]), linked to qryPrintChecks_05_13_03
              ' ** (qryPrintChecks_05_13_02 (qryPrintChecks_05_13_01 (tblCheckBank, grouped by
              ' ** accountno, with Max(chkbank_datemodified)), linked back to tblCheckBank, grouped
              ' ** by accountno, with Max(chkbank_id)), linked back to tblCheckBank)) to tblCheckVoid_Staging.
5690          Set qdf = dbs.QueryDefs("qryPrintChecks_05_13_06")
5700          With qdf.Parameters
5710            ![cvset] = lngChkVoid_Set
5720          End With
5730          qdf.Execute
5740          Set qdf = Nothing
5750          DoEvents

5760        End Select

            ' ***************************************
            ' ** User chooses which checks to void.
            ' ***************************************

5770        .Modal = False
5780        strDocName = "frmRpt_Checks_Void"
5790        DoCmd.OpenForm strDocName, , , , , acDialog, frm.Name & "~" & strAccountNo & "~" & Format(blnAll, "True/False")
5800        .Modal = True

            ' ** This variable signals whether user changed mind, and all printed successfully.
5810        If gblnMessage = True Then

5820          DoCmd.Hourglass True  ' ** Make sure it's still running.
5830          DoEvents

              ' ********************************************
              ' ** Bank info already saved and/or updated.
              ' ********************************************

              ' ************************************
              ' ** Voids already in tblChecksVoid.
              ' ************************************

              ' *****************************************
              ' ** Resetting the Journal handled above.
              ' *****************************************

5840        Else
              ' ** This means the user changed their mind, and now all checks are good.
5850          blnContinue = False
5860        End If

5870      End Select

5880    End With

EXITP:
5890    Set rst1 = Nothing
5900    Set rst1 = Nothing
5910    Set qdf = Nothing
5920    Set dbs = Nothing
5930    Exit Sub

ERRH:
5940    DoCmd.Hourglass False
5950    Select Case ERR.Number
        Case Else
5960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5970    End Select
5980    Resume EXITP

End Sub

Public Sub ColumnSet(strProc As String, lngFrm_Top As Long, lngFrm_Left As Long, lngFrm_Width As Long, lngFrm_Height As Long, lngThisFrm_Width As Long, lngFrmMin_Width As Long, lngFrmWidthOffset As Long, lngClose_Offset As Long, lngLbx_Width As Long, lngLbxBoxOffset As Long, strLbx_ColWidths As String, lngLbx_Col00 As Long, lngLbx_Col01 As Long, lngLbx_Col02 As Long, lngLbx_Col03 As Long, lngLbx_Col04 As Long, lngLbx_Col05 As Long, lngLbx_Col06 As Long, lngLbx_Col07 As Long, lngLbx_Col08 As Long, lngLbx_Col09 As Long, lngGTR_Offset As Long, lngMonitorCnt As Long, lngMonitorNum As Long, frm As Access.Form, Optional varShow As Variant)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "ColumnSet"

        Dim strChkBox As String
        Dim lngLbxCur_Width As Long, strLbxCur_ColWidths As String, lngLbxNew_Width As Long, strLbxNew_ColWidths As String
        Dim lngThisFrmCur_Width As Long, lngThisFrmNew_Width As Long
        Dim lngCol00 As Long, lngCol01 As Long, lngCol02 As Long, lngCol03 As Long, lngCol04 As Long
        Dim lngCol05 As Long, lngCol06 As Long, lngCol07 As Long, lngCol08 As Long, lngCol09 As Long
        Dim lngWidthDiff As Long, lngScrollBar As Long, lngColsTot As Long, lngFrmWidth_Diff As Long, lngFrmNew_Left As Long
        Dim blnShow As Boolean, blnNoChange As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String, lngTmp02 As Long, lngTmp03 As Long
        Dim lngX As Long

6010    With frm

6020      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
6030        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
6040      End If

6050      lngWidthDiff = 0&: lngLbxCur_Width = 0&: lngLbxNew_Width = 0&: lngThisFrmCur_Width = 0&: lngThisFrmNew_Width = 0&
6060      strLbxCur_ColWidths = vbNullString: strLbxNew_ColWidths = vbNullString: strChkBox = vbNullString
6070      lngScrollBar = 0&: lngColsTot = 0&
6080      blnNoChange = False

6090      If strProc = "Form_Open" Then
            ' ** Get the opening specs.
6100        lngThisFrm_Width = .Width
6110        lngLbx_Width = CLng(.lbxShortAccountName.Width)
6120        strLbx_ColWidths = .lbxShortAccountName.ColumnWidths
6130        strTmp01 = strLbx_ColWidths
6140        intPos01 = InStr(strTmp01, ";")
6150        lngLbx_Col00 = CLng(Left(strTmp01, (intPos01 - 1)))  '0
6160        strTmp01 = Mid(strTmp01, (intPos01 + 1))
6170        intPos01 = InStr(strTmp01, ";")
6180        lngLbx_Col01 = CLng(Left(strTmp01, (intPos01 - 1)))  '1335
6190        strTmp01 = Mid(strTmp01, (intPos01 + 1))
6200        intPos01 = InStr(strTmp01, ";")
6210        lngLbx_Col02 = CLng(Left(strTmp01, (intPos01 - 1)))  '3030
6220        strTmp01 = Mid(strTmp01, (intPos01 + 1))
6230        intPos01 = InStr(strTmp01, ";")
6240        lngLbx_Col03 = CLng(Left(strTmp01, (intPos01 - 1)))  '420
6250        strTmp01 = Mid(strTmp01, (intPos01 + 1))
6260        intPos01 = InStr(strTmp01, ";")
6270        lngLbx_Col04 = CLng(Left(strTmp01, (intPos01 - 1)))  '705
6280        strTmp01 = Mid(strTmp01, (intPos01 + 1))
6290        intPos01 = InStr(strTmp01, ";")
6300        lngLbx_Col05 = CLng(Left(strTmp01, (intPos01 - 1)))  '3030
6310        strTmp01 = Mid(strTmp01, (intPos01 + 1))
6320        intPos01 = InStr(strTmp01, ";")
6330        lngLbx_Col06 = CLng(Left(strTmp01, (intPos01 - 1)))  '2160
6340        strTmp01 = Mid(strTmp01, (intPos01 + 1))
6350        intPos01 = InStr(strTmp01, ";")
6360        lngLbx_Col07 = CLng(Left(strTmp01, (intPos01 - 1)))  '1770
6370        strTmp01 = Mid(strTmp01, (intPos01 + 1))
6380        intPos01 = InStr(strTmp01, ";")
6390        lngLbx_Col08 = CLng(Left(strTmp01, (intPos01 - 1)))  '1290
6400        lngLbx_Col09 = CLng(Mid(strTmp01, (intPos01 + 1)))   '705
6410        lngClose_Offset = (.Width - .cmdClose.Left)
6420        lngLbxBoxOffset = (.lbxShortAccountName_box.Width - .lbxShortAccountName.Width)
6430        lngFrmWidthOffset = (.Width - (.lbxShortAccountName_box.Left + .lbxShortAccountName_box.Width))
6440        lngFrmMin_Width = (((.cmbSortBy_box.Left + .cmbSortBy_box.Width) + (7& * lngTpp)) + (10& * lngTpp))  ' ** To accommodate Display box.
6450        blnNoChange = True
6460        For lngX = 0& To (.cmbSortBy.ListCount - 1&)
6470          .Controls(.cmbSortBy.Column(CBX_SORT_CNAM, lngX)).Tag = .cmbSortBy.Column(CBX_SORT_FID, lngX)
6480        Next

6490      Else

6500        Select Case IsMissing(varShow)
            Case True
6510          Beep
6520          blnNoChange = True
6530        Case False
6540          blnShow = CBool(varShow)

6550          strChkBox = Left(strProc, (Len(strProc) - Len("_AfterUpdate")))
6560          lngThisFrmCur_Width = .Width
6570          lngLbxCur_Width = CLng(.lbxShortAccountName.Width)
6580          strLbxCur_ColWidths = .lbxShortAccountName.ColumnWidths
6590          strTmp01 = strLbxCur_ColWidths
6600          intPos01 = InStr(strTmp01, ";")
6610          lngCol00 = CLng(Left(strTmp01, (intPos01 - 1)))
6620          strTmp01 = Mid(strTmp01, (intPos01 + 1))
6630          intPos01 = InStr(strTmp01, ";")
6640          lngCol01 = CLng(Left(strTmp01, (intPos01 - 1)))
6650          strTmp01 = Mid(strTmp01, (intPos01 + 1))
6660          intPos01 = InStr(strTmp01, ";")
6670          lngCol02 = CLng(Left(strTmp01, (intPos01 - 1)))
6680          strTmp01 = Mid(strTmp01, (intPos01 + 1))
6690          intPos01 = InStr(strTmp01, ";")
6700          lngCol03 = CLng(Left(strTmp01, (intPos01 - 1)))
6710          strTmp01 = Mid(strTmp01, (intPos01 + 1))
6720          intPos01 = InStr(strTmp01, ";")
6730          lngCol04 = CLng(Left(strTmp01, (intPos01 - 1)))
6740          strTmp01 = Mid(strTmp01, (intPos01 + 1))
6750          intPos01 = InStr(strTmp01, ";")
6760          lngCol05 = CLng(Left(strTmp01, (intPos01 - 1)))
6770          strTmp01 = Mid(strTmp01, (intPos01 + 1))
6780          intPos01 = InStr(strTmp01, ";")
6790          lngCol06 = CLng(Left(strTmp01, (intPos01 - 1)))
6800          strTmp01 = Mid(strTmp01, (intPos01 + 1))
6810          intPos01 = InStr(strTmp01, ";")
6820          lngCol07 = CLng(Left(strTmp01, (intPos01 - 1)))
6830          strTmp01 = Mid(strTmp01, (intPos01 + 1))
6840          intPos01 = InStr(strTmp01, ";")
6850          lngCol08 = CLng(Left(strTmp01, (intPos01 - 1)))
6860          lngCol09 = CLng(Mid(strTmp01, (intPos01 + 1)))
6870          lngColsTot = (lngCol01 + lngCol02 + lngCol03 + lngCol04 + lngCol05 + lngCol06 + lngCol07 + lngCol08 + lngCol09)
6880          lngScrollBar = (lngLbxCur_Width - lngColsTot)  ' ** On my machine it's 17 pixels, so (17& * 15&), 255 Twips.

              ' ** The columns 'Check_Count' and 'Last_Check_Number' aren't selectable.
6890          Select Case strChkBox
              Case "ckgDisplay_opt01_AccountNo"
6900            Select Case blnShow
                Case True
6910              If lngCol01 <> lngLbx_Col01 Then  ' ** It could only be Zero.
6920                lngWidthDiff = lngLbx_Col01
6930                lngCol01 = lngLbx_Col01
6940              Else
6950                blnNoChange = True
6960              End If
6970            Case False
6980              If lngCol01 <> 0& Then
6990                Select Case .ckgDisplay_opt02_ShortName
                    Case True
7000                  lngWidthDiff = -(lngLbx_Col01)
7010                Case False
                      ' ** Either accountno or shortname must always be visible.
7020                  lngWidthDiff = (lngLbx_Col02 - lngLbx_Col01)  ' ** It'll be positive.
7030                  .ckgDisplay_opt02_ShortName = True
7040                  .ckgDisplay_opt02_ShortName_lbl.FontBold = True
7050                  .ckgDisplay_opt02_ShortName_lbl_dim.FontBold = True
7060                  .ckgDisplay_opt02_ShortName_lbl_dim_hi.FontBold = True
7070                  lngWidthDiff = (lngWidthDiff + lngLbx_Col02)
7080                  lngCol02 = lngLbx_Col02
7090                End Select
7100                lngCol01 = 0&
7110              Else
7120                blnNoChange = True
7130              End If
7140            End Select
7150          Case "ckgDisplay_opt02_ShortName"
7160            Select Case blnShow
                Case True
7170              If lngCol02 <> lngLbx_Col02 Then
7180                lngWidthDiff = lngLbx_Col02
7190                lngCol02 = lngLbx_Col02
7200              Else
7210                blnNoChange = True
7220              End If
7230            Case False
7240              If lngCol02 <> 0& Then
7250                Select Case .ckgDisplay_opt01_AccountNo
                    Case True
7260                  lngWidthDiff = -(lngLbx_Col02)
7270                Case False
                      ' ** Either accountno or shortname must always be visible.
7280                  lngWidthDiff = -(lngLbx_Col02 - lngLbx_Col01)
7290                  .ckgDisplay_opt01_AccountNo = True
7300                  .ckgDisplay_opt01_AccountNo_lbl.FontBold = True
7310                  .ckgDisplay_opt01_AccountNo_lbl_dim.FontBold = True
7320                  .ckgDisplay_opt01_AccountNo_lbl_dim_hi.FontBold = True
7330                  lngWidthDiff = (lngWidthDiff + lngLbx_Col01)
7340                  lngCol01 = lngLbx_Col01
7350                End Select
7360                lngCol02 = 0&
7370              Else
7380                blnNoChange = True
7390              End If
7400            End Select
7410          Case "ckgDisplay_opt05_Payee"
7420            Select Case blnShow
                Case True
7430              If lngCol05 <> lngLbx_Col05 Then
7440                lngWidthDiff = lngLbx_Col05
7450                lngCol05 = lngLbx_Col05
7460              Else
7470                blnNoChange = True
7480              End If
7490            Case False
7500              If lngCol05 <> 0& Then
7510                lngWidthDiff = -(lngLbx_Col05)
7520                lngCol05 = 0&
7530              Else
7540                blnNoChange = True
7550              End If
7560            End Select
7570          Case "ckgDisplay_opt06_BankName"
7580            Select Case blnShow
                Case True
7590              If lngCol06 <> lngLbx_Col06 Then
7600                lngWidthDiff = lngLbx_Col06
7610                lngCol06 = lngLbx_Col06
7620              Else
7630                blnNoChange = True
7640              End If
7650            Case False
7660              If lngCol06 <> 0& Then
7670                lngWidthDiff = -(lngLbx_Col06)
7680                lngCol06 = 0&
7690              Else
7700                blnNoChange = True
7710              End If
7720            End Select
7730          Case "ckgDisplay_opt07_BankAcctNum"
7740            Select Case blnShow
                Case True
7750              If lngCol07 <> lngLbx_Col07 Then
7760                lngWidthDiff = lngLbx_Col07
7770                lngCol07 = lngLbx_Col07
7780              Else
7790                blnNoChange = True
7800              End If
7810            Case False
7820              If lngCol07 <> 0& Then
7830                lngWidthDiff = -(lngLbx_Col07)
7840                lngCol07 = 0&
7850              Else
7860                blnNoChange = True
7870              End If
7880            End Select
7890          Case "ckgDisplay_opt08_CheckAmount"
7900            Select Case blnShow
                Case True
7910              If lngCol08 <> lngLbx_Col08 Then
7920                lngWidthDiff = lngLbx_Col08
7930                lngCol08 = lngLbx_Col08
7940              Else
7950                blnNoChange = True
7960              End If
7970            Case False
7980              If lngCol08 <> 0& Then
7990                lngWidthDiff = -(lngLbx_Col08)
8000                lngCol08 = 0&
8010              Else
8020                blnNoChange = True
8030              End If
8040            End Select
8050          End Select

8060          Select Case .ChkStat
              Case "No"
                ' ** Don't show it.
8070            If lngCol09 <> 0& Then
                  ' ** Shorten it, so add its width to whatever else is going on.
8080              lngWidthDiff = (lngWidthDiff - lngCol09)
8090              lngCol09 = 0&
8100              If blnNoChange = True Then blnNoChange = False
8110            End If
8120          Case "Yes"
                ' ** Show it.
8130            If lngCol09 = 0& Then
8140              lngCol09 = lngLbx_Col09
8150              lngWidthDiff = (lngWidthDiff + lngCol09)
8160              If blnNoChange = True Then blnNoChange = False
8170            End If
8180          Case "Mixed"
                ' ** Show it.
8190            If lngCol09 = 0& Then
8200              lngCol09 = lngLbx_Col09
8210              lngWidthDiff = (lngWidthDiff + lngCol09)
8220              If blnNoChange = True Then blnNoChange = False
8230            End If
8240          End Select

8250          If blnNoChange = False Then

8260            lngLbxNew_Width = ((lngCol01 + lngCol02 + lngCol03 + lngCol04 + lngCol05 + lngCol06 + lngCol07 + lngCol08 + lngCol09) + lngScrollBar)
8270            strLbxNew_ColWidths = (CStr(lngCol00) & ";" & CStr(lngCol01) & ";" & CStr(lngCol02) & ";" & CStr(lngCol03) & ";" & _
                  CStr(lngCol04) & ";" & CStr(lngCol05) & ";" & CStr(lngCol06) & ";" & CStr(lngCol07) & ";" & CStr(lngCol08) & ";" & CStr(lngCol09))
                'lngThisFrmNew_Width = (lngThisFrmCur_Width + lngWidthDiff)
8280            lngThisFrmNew_Width = ((.lbxShortAccountName_box.Left + (lngLbxNew_Width + lngLbxBoxOffset)) + lngFrmWidthOffset)
8290            If lngThisFrmNew_Width < lngFrmMin_Width Then
8300              lngThisFrmNew_Width = lngFrmMin_Width
8310            End If

8320            If lngWidthDiff > 0& Then
8330              .Width = lngThisFrmNew_Width
8340            End If

8350            If lngCol01 > 0& Then
8360              .Account_Number_lbl.Visible = True
8370              .Account_Number_lbl_line.Visible = True
8380            Else
8390              .Account_Number_lbl.Visible = False
8400              .Account_Number_lbl_line.Visible = False
8410            End If
8420            .Account_Number_lbl_dim_hi.Visible = False
8430            .Account_Number_lbl_line_dim_hi.Visible = False

8440            If lngCol02 > 0& Then
8450              If lngCol01 > 0& Then
8460                .Short_Name_lbl.Left = ((.Account_Number_lbl.Left + .Account_Number_lbl.Width) + (2& * lngTpp))
8470              Else
8480                .Short_Name_lbl.Left = .Account_Number_lbl.Left
8490              End If
8500            Else
8510              .Short_Name_lbl.Left = .Account_Number_lbl.Left  ' ** Move it out of the way.
8520            End If
8530            .Short_Name_lbl_line.Left = .Short_Name_lbl.Left
8540            .Short_Name_lbl_dim_hi.Left = (.Short_Name_lbl.Left + lngTpp)
8550            .Short_Name_lbl_line_dim_hi.Left = (.Short_Name_lbl_line.Left + lngTpp)
8560            If lngCol02 > 0& Then
8570              .Short_Name_lbl.Visible = True
8580              .Short_Name_lbl_line.Visible = True
8590            Else
8600              .Short_Name_lbl.Visible = False
8610              .Short_Name_lbl_line.Visible = False
8620            End If
8630            .Short_Name_lbl_dim_hi.Visible = False
8640            .Short_Name_lbl_line_dim_hi.Visible = False

8650            If lngCol03 > 0& Then
8660              If lngCol02 > 0& Then
8670                .Check_Count_lbl.Left = ((.Short_Name_lbl.Left + .Short_Name_lbl.Width) + (2& * lngTpp))
8680              Else
8690                .Check_Count_lbl.Left = ((.Account_Number_lbl.Left + .Account_Number_lbl.Width) + (2& * lngTpp))
8700              End If
8710            Else
8720              .Check_Count_lbl.Left = .Account_Number_lbl.Left  ' ** Move it out of the way.
8730            End If
8740            .Check_Count_lbl2.Left = .Check_Count_lbl.Left
8750            .Check_Count_lbl_line.Left = .Check_Count_lbl.Left
8760            .Check_Count_lbl_dim_hi.Left = (.Check_Count_lbl.Left + lngTpp)
8770            .Check_Count_lbl2_dim_hi.Left = .Check_Count_lbl_dim_hi.Left
8780            .Check_Count_lbl_line_dim_hi.Left = (.Check_Count_lbl_line.Left + lngTpp)
8790            If lngCol03 > 0& Then
8800              .Check_Count_lbl.Visible = True
8810              .Check_Count_lbl2.Visible = True
8820              .Check_Count_lbl_line.Visible = True
8830            Else
8840              .Check_Count_lbl.Visible = False
8850              .Check_Count_lbl2.Visible = False
8860              .Check_Count_lbl_line.Visible = False
8870            End If
8880            .Check_Count_lbl_dim_hi.Visible = False
8890            .Check_Count_lbl2_dim_hi.Visible = False
8900            .Check_Count_lbl_line_dim_hi.Visible = False

8910            If lngCol04 > 0& Then
8920              If lngCol03 > 0& Then
8930                .Last_Check_Number_lbl.Left = ((.Check_Count_lbl.Left + .Check_Count_lbl.Width) + (2& * lngTpp))
8940              ElseIf lngCol02 > 0& Then
8950                .Last_Check_Number_lbl.Left = ((.Short_Name_lbl.Left + .Short_Name_lbl.Width) + (2& * lngTpp))
8960              Else
8970                .Last_Check_Number_lbl.Left = ((.Account_Number_lbl.Left + .Account_Number_lbl.Width) + (2& * lngTpp))
8980              End If
8990            Else
9000              .Last_Check_Number_lbl.Left = .Account_Number_lbl.Left  ' ** Move it out of the way.
9010            End If
9020            .Last_Check_Number_lbl2.Left = .Last_Check_Number_lbl.Left
9030            .Last_Check_Number_lbl_line.Left = .Last_Check_Number_lbl.Left
9040            .Last_Check_Number_lbl_dim_hi.Left = (.Last_Check_Number_lbl.Left + lngTpp)
9050            .Last_Check_Number_lbl2_dim_hi.Left = .Last_Check_Number_lbl_dim_hi.Left
9060            .Last_Check_Number_lbl_line_dim_hi.Left = (.Last_Check_Number_lbl_line.Left + lngTpp)
9070            If lngCol04 > 0& Then
9080              .Last_Check_Number_lbl.Visible = True
9090              .Last_Check_Number_lbl2.Visible = True
9100              .Last_Check_Number_lbl_line.Visible = True
9110            Else
9120              .Last_Check_Number_lbl.Visible = False
9130              .Last_Check_Number_lbl2.Visible = False
9140              .Last_Check_Number_lbl_line.Visible = False
9150            End If
9160            .Last_Check_Number_lbl_dim_hi.Visible = False
9170            .Last_Check_Number_lbl2_dim_hi.Visible = False
9180            .Last_Check_Number_lbl_line_dim_hi.Visible = False

9190            If lngCol05 > 0& Then
9200              If lngCol04 > 0& Then
9210                .Payee_lbl.Left = ((.Last_Check_Number_lbl.Left + .Last_Check_Number_lbl.Width) + (2& * lngTpp))
9220              ElseIf lngCol03 > 0& Then
9230                .Payee_lbl.Left = ((.Check_Count_lbl.Left + .Check_Count_lbl.Width) + (2& * lngTpp))
9240              ElseIf lngCol02 > 0& Then
9250                .Payee_lbl.Left = ((.Short_Name_lbl.Left + .Short_Name_lbl.Width) + (2& * lngTpp))
9260              Else
9270                .Payee_lbl.Left = ((.Account_Number_lbl.Left + .Account_Number_lbl.Width) + (2& * lngTpp))
9280              End If
9290            Else
9300              .Payee_lbl.Left = .Account_Number_lbl.Left  ' ** Move it out of the way.
9310            End If
9320            .Payee_lbl_line.Left = .Payee_lbl.Left
9330            .Payee_lbl_dim_hi.Left = (.Payee_lbl.Left + lngTpp)
9340            .Payee_lbl_line_dim_hi.Left = (.Payee_lbl_line.Left + lngTpp)
9350            If lngCol05 > 0& Then
9360              .Payee_lbl.Visible = True
9370              .Payee_lbl_line.Visible = True
9380            Else
9390              .Payee_lbl.Visible = False
9400              .Payee_lbl_line.Visible = False
9410            End If
9420            .Payee_lbl_dim_hi.Visible = False
9430            .Payee_lbl_line_dim_hi.Visible = False

9440            If lngCol06 > 0& Then
9450              If lngCol05 > 0& Then
9460                .Bank_Name_lbl.Left = ((.Payee_lbl.Left + .Payee_lbl.Width) + (2& * lngTpp))
9470              ElseIf lngCol04 > 0& Then
9480                .Bank_Name_lbl.Left = ((.Last_Check_Number_lbl.Left + .Last_Check_Number_lbl.Width) + (2& * lngTpp))
9490              ElseIf lngCol03 > 0& Then
9500                .Bank_Name_lbl.Left = ((.Check_Count_lbl.Left + .Check_Count_lbl.Width) + (2& * lngTpp))
9510              ElseIf lngCol02 > 0& Then
9520                .Bank_Name_lbl.Left = ((.Short_Name_lbl.Left + .Short_Name_lbl.Width) + (2& * lngTpp))
9530              Else
9540                .Bank_Name_lbl.Left = ((.Account_Number_lbl.Left + .Account_Number_lbl.Width) + (2& * lngTpp))
9550              End If
9560            Else
9570              .Bank_Name_lbl.Left = .Account_Number_lbl.Left  ' ** Move it out of the way.
9580            End If
9590            .Bank_Name_lbl_line.Left = .Bank_Name_lbl.Left
9600            .Bank_Name_lbl_dim_hi.Left = (.Bank_Name_lbl.Left + lngTpp)
9610            .Bank_Name_lbl_line_dim_hi.Left = (.Bank_Name_lbl_line.Left + lngTpp)
9620            If lngCol06 > 0& Then
9630              .Bank_Name_lbl.Visible = True
9640              .Bank_Name_lbl_line.Visible = True
9650            Else
9660              .Bank_Name_lbl.Visible = False
9670              .Bank_Name_lbl_line.Visible = False
9680            End If
9690            .Bank_Name_lbl_dim_hi.Visible = False
9700            .Bank_Name_lbl_line_dim_hi.Visible = False

9710            If lngCol07 > 0& Then
9720              If lngCol06 > 0& Then
9730                .Bank_Account_Number_lbl.Left = ((.Bank_Name_lbl.Left + .Bank_Name_lbl.Width) + (2& * lngTpp))
9740              ElseIf lngCol05 > 0& Then
9750                .Bank_Account_Number_lbl.Left = ((.Payee_lbl.Left + .Payee_lbl.Width) + (2& * lngTpp))
9760              ElseIf lngCol04 > 0& Then
9770                .Bank_Account_Number_lbl.Left = ((.Last_Check_Number_lbl.Left + .Last_Check_Number_lbl.Width) + (2& * lngTpp))
9780              ElseIf lngCol03 > 0& Then
9790                .Bank_Account_Number_lbl.Left = ((.Check_Count_lbl.Left + .Check_Count_lbl.Width) + (2& * lngTpp))
9800              ElseIf lngCol02 > 0& Then
9810                .Bank_Account_Number_lbl.Left = ((.Short_Name_lbl.Left + .Short_Name_lbl.Width) + (2& * lngTpp))
9820              Else
9830                .Bank_Account_Number_lbl.Left = ((.Account_Number_lbl.Left + .Account_Number_lbl.Width) + (2& * lngTpp))
9840              End If
9850            Else
9860              .Bank_Account_Number_lbl.Left = .Account_Number_lbl.Left  ' ** Move it out of the way.
9870            End If
9880            .Bank_Account_Number_lbl_line.Left = .Bank_Account_Number_lbl.Left
9890            .Bank_Account_Number_lbl_dim_hi.Left = (.Bank_Account_Number_lbl.Left + lngTpp)
9900            .Bank_Account_Number_lbl_line_dim_hi.Left = (.Bank_Account_Number_lbl_line.Left + lngTpp)
9910            If lngCol07 > 0& Then
9920              .Bank_Account_Number_lbl.Visible = True
9930              .Bank_Account_Number_lbl_line.Visible = True
9940            Else
9950              .Bank_Account_Number_lbl.Visible = False
9960              .Bank_Account_Number_lbl_line.Visible = False
9970            End If
9980            .Bank_Account_Number_lbl_dim_hi.Visible = False
9990            .Bank_Account_Number_lbl_line_dim_hi.Visible = False

10000           If lngCol08 > 0& Then
10010             If lngCol07 > 0& Then
10020               .Check_Amount_lbl.Left = ((.Bank_Account_Number_lbl.Left + .Bank_Account_Number_lbl.Width) + (2& * lngTpp))
10030             ElseIf lngCol06 > 0& Then
10040               .Check_Amount_lbl.Left = ((.Bank_Name_lbl.Left + .Bank_Name_lbl.Width) + (2& * lngTpp))
10050             ElseIf lngCol05 > 0& Then
10060               .Check_Amount_lbl.Left = ((.Payee_lbl.Left + .Payee_lbl.Width) + (2& * lngTpp))
10070             ElseIf lngCol04 > 0& Then
10080               .Check_Amount_lbl.Left = ((.Last_Check_Number_lbl.Left + .Last_Check_Number_lbl.Width) + (2& * lngTpp))
10090             ElseIf lngCol03 > 0& Then
10100               .Check_Amount_lbl.Left = ((.Check_Count_lbl.Left + .Check_Count_lbl.Width) + (2& * lngTpp))
10110             ElseIf lngCol02 > 0& Then
10120               .Check_Amount_lbl.Left = ((.Short_Name_lbl.Left + .Short_Name_lbl.Width) + (2& * lngTpp))
10130             Else
10140               .Check_Amount_lbl.Left = ((.Account_Number_lbl.Left + .Account_Number_lbl.Width) + (2& * lngTpp))
10150             End If
10160           Else
10170             .Check_Amount_lbl.Left = .Account_Number_lbl.Left  ' ** Move it out of the way.
10180           End If
10190           .Check_Amount_lbl_line.Left = .Check_Amount_lbl.Left
10200           .Check_Amount_lbl_dim_hi.Left = (.Check_Amount_lbl.Left + lngTpp)
10210           .Check_Amount_lbl_line_dim_hi.Left = (.Check_Amount_lbl_line.Left + lngTpp)
10220           If lngCol08 > 0& Then
10230             .Check_Amount_lbl.Visible = True
10240             .Check_Amount_lbl_line.Visible = True
10250           Else
10260             .Check_Amount_lbl.Visible = False
10270             .Check_Amount_lbl_line.Visible = False
10280           End If
10290           .Check_Amount_lbl_dim_hi.Visible = False
10300           .Check_Amount_lbl_line_dim_hi.Visible = False

10310           If lngCol09 > 0& Then
10320             If lngCol08 > 0& Then
10330               .ChkStat_lbl.Left = ((.Check_Amount_lbl.Left + .Check_Amount_lbl.Width) + (2& * lngTpp))
10340             ElseIf lngCol07 > 0& Then
10350               .ChkStat_lbl.Left = ((.Bank_Account_Number_lbl.Left + .Bank_Account_Number_lbl.Width) + (2& * lngTpp))
10360             ElseIf lngCol06 > 0& Then
10370               .ChkStat_lbl.Left = ((.Bank_Name_lbl.Left + .Bank_Name_lbl.Width) + (2& * lngTpp))
10380             ElseIf lngCol05 > 0& Then
10390               .ChkStat_lbl.Left = ((.Payee_lbl.Left + .Payee_lbl.Width) + (2& * lngTpp))
10400             ElseIf lngCol04 > 0& Then
10410               .ChkStat_lbl.Left = ((.Last_Check_Number_lbl.Left + .Last_Check_Number_lbl.Width) + (2& * lngTpp))
10420             ElseIf lngCol03 > 0& Then
10430               .ChkStat_lbl.Left = ((.Check_Count_lbl.Left + .Check_Count_lbl.Width) + (2& * lngTpp))
10440             ElseIf lngCol02 > 0& Then
10450               .ChkStat_lbl.Left = ((.Short_Name_lbl.Left + .Short_Name_lbl.Width) + (2& * lngTpp))
10460             Else
10470               .ChkStat_lbl.Left = ((.Account_Number_lbl.Left + .Account_Number_lbl.Width) + (2& * lngTpp))
10480             End If
10490           Else
10500             .ChkStat_lbl.Left = .Account_Number_lbl.Left  ' ** Move it out of the way.
10510           End If
10520           .ChkStat_lbl_line.Left = .ChkStat_lbl.Left
10530           .ChkStat_lbl_dim_hi.Left = (.ChkStat_lbl.Left + lngTpp)
10540           .ChkStat_lbl_line_dim_hi.Left = (.ChkStat_lbl_line.Left + lngTpp)
10550           If lngCol09 > 0& Then
10560             .ChkStat_lbl.Visible = True
10570             .ChkStat_lbl_line.Visible = True
10580           Else
10590             .ChkStat_lbl.Visible = False
10600             .ChkStat_lbl_line.Visible = False
10610           End If
10620           .ChkStat_lbl_dim_hi.Visible = False
10630           .ChkStat_lbl_line_dim_hi.Visible = False

10640           If lngWidthDiff > 0& Then
10650             .Width = lngThisFrmNew_Width  ' ** Already widened.
                  '.form_width_line.Left = (lngThisFrmNew_Width - .form_width_line.Width)
10660             .cmdClose.Left = (lngThisFrmNew_Width - lngClose_Offset)
10670             .Header_vline01.Left = lngThisFrmNew_Width
10680             .Header_vline02.Left = lngThisFrmNew_Width
10690             .Footer_vline01.Left = lngThisFrmNew_Width
10700             .Footer_vline02.Left = lngThisFrmNew_Width
10710             .Header_hline01.Width = lngThisFrmNew_Width
10720             .Header_hline02.Width = lngThisFrmNew_Width
10730             .Footer_hline01.Width = lngThisFrmNew_Width
10740             .Footer_hline02.Width = lngThisFrmNew_Width
10750             .lbxShortAccountName.Width = lngLbxNew_Width
10760             .lbxShortAccountName.ColumnWidths = strLbxNew_ColWidths
10770             .lbxShortAccountName_box.Width = (.lbxShortAccountName.Width + lngLbxBoxOffset)
10780             .PendingJournal_lbl.Width = .lbxShortAccountName_box.Width
10790             .PendingCheck_lbl.Width = .lbxShortAccountName_box.Width
10800             lngTmp02 = ((.lbxShortAccountName_box.Left + .lbxShortAccountName_box.Width) - .SeparateCheckingOpt_lbl.Width)
10810             If lngTmp02 < .Sort_Up_lbl.Left Then
10820               lngTmp02 = .Sort_Up_lbl.Left  ' ** Arbitrary.
10830             End If
10840             .SeparateCheckingOpt_lbl.Left = lngTmp02
10850             .cmbSortBy_lbl2.Width = lngThisFrmNew_Width
10860             For lngX = 1& To 24&
10870               .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Left = (lngThisFrmNew_Width - lngGTR_Offset)
10880             Next
10890           ElseIf lngWidthDiff < 0& Then
                  '.form_width_line.Left = (lngThisFrmNew_Width - .form_width_line.Width)
10900             .cmdClose.Left = (lngThisFrmNew_Width - lngClose_Offset)
10910             .Header_vline01.Left = lngThisFrmNew_Width
10920             .Header_vline02.Left = lngThisFrmNew_Width
10930             .Footer_vline01.Left = lngThisFrmNew_Width
10940             .Footer_vline02.Left = lngThisFrmNew_Width
10950             .Header_hline01.Width = lngThisFrmNew_Width
10960             .Header_hline02.Width = lngThisFrmNew_Width
10970             .Footer_hline01.Width = lngThisFrmNew_Width
10980             .Footer_hline02.Width = lngThisFrmNew_Width
10990             .lbxShortAccountName.Width = lngLbxNew_Width
11000             .lbxShortAccountName.ColumnWidths = strLbxNew_ColWidths
11010             .lbxShortAccountName_box.Width = (.lbxShortAccountName.Width + lngLbxBoxOffset)
11020             .PendingJournal_lbl.Width = .lbxShortAccountName_box.Width
11030             .PendingCheck_lbl.Width = .lbxShortAccountName_box.Width
11040             lngTmp02 = ((.lbxShortAccountName_box.Left + .lbxShortAccountName_box.Width) - .SeparateCheckingOpt_lbl.Width)
11050             If lngTmp02 < .Sort_Up_lbl.Left Then
11060               lngTmp02 = .Sort_Up_lbl.Left  ' ** Arbitrary.
11070             End If
11080             .SeparateCheckingOpt_lbl.Left = lngTmp02
11090             .cmbSortBy_lbl2.Width = lngThisFrmNew_Width
11100             For lngX = 1& To 24&
11110               .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Left = (lngThisFrmNew_Width - lngGTR_Offset)
11120             Next
11130             .Width = lngThisFrmNew_Width
11140           End If

11150           If blnShow = False And .chkSyncListSort = True Then
11160             If strChkBox = .cmbSortBy.Column(CBX_SORT_CNAM) Then
                    ' ** Change the sort.
11170               If lngCol01 > 0& Then
11180                 .cmbSortBy = .cmbSortBy.Column(CBX_SORT_FID, 0)
11190               Else
11200                 .cmbSortBy = .cmbSortBy.Column(CBX_SORT_FID, 1)
11210               End If
11220               .cmbSortBy_AfterUpdate  ' ** Form Procedure: frmRpt_Checks.
11230             End If
11240           End If

11250           If (lngThisFrmNew_Width <> lngThisFrmCur_Width) Then
                  ' ** Size has changed.
11260             lngFrmWidth_Diff = (lngThisFrm_Width - lngThisFrmNew_Width)
11270             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
11280             lngMonitorNum = 1&: lngTmp03 = 0&
11290             EnumMonitors frm  ' ** Module Function: modMonitorFuncs.
11300             If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
11310             If lngFrmWidth_Diff = 0& Then
                    ' ** Back to the opening dimensions.
11320               If lngMonitorNum = 1& Then lngTmp03 = lngFrm_Top
11330               DoCmd.SelectObject acForm, frm.Name, False
11340               DoCmd.MoveSize lngFrm_Left, lngTmp03, lngFrm_Width, lngFrm_Height  'lngFrm_Top
11350               If lngMonitorNum > 1& Then
11360                 LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
11370               End If
11380             Else
11390               lngFrmNew_Left = (lngFrm_Left + ((lngFrm_Width - (lngFrm_Width - lngFrmWidth_Diff)) / 2))
11400               If lngMonitorNum = 1& Then lngTmp03 = lngFrm_Top
11410               DoCmd.SelectObject acForm, frm.Name, False
11420               DoCmd.MoveSize lngFrmNew_Left, lngTmp03, (lngFrm_Width - lngFrmWidth_Diff), lngFrm_Height  'lngFrm_Top
11430               If lngMonitorNum > 1& Then
11440                 LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
11450               End If
11460             End If
11470           End If

11480         End If  ' ** blnNoChange.
11490       End Select  ' ** IsMissing.
11500     End If  ' ** strProc.

11510   End With  ' ** Me.

        ' ** All 7:
        ' ** Forms(0).lbxShortAccountName.Width = 13995  '12705
        ' ** Forms(0).lbxShortAccountName.ColumnWidths = "0;1335;3030;420;705;3030;2160;1770;1290"  '"1335;3030;420;705;3030;2160;1770"

        'OH MY! THIS COVERS WHICH COLUMNS ARE SHOWN, ALONG WITH THE COLUMN WIDTHS!

        'HOW DO I GET ChkStat INTO THIS, AND HOW SHOULD IT BE USED?
        'HOW ABOUT...
        'IF NONE HAVE BEEN PRINTED, DON'T SHOW THE ChkStat COLUMN,
        'OTHERWISE SHOW IT!
        'Printed
        '  Y  Yes, it's been printed already
        '  N  No, it hasn't been printed
        '  M  Mixed, some have some haven't

EXITP:
11520   Exit Sub

ERRH:
11530   Select Case ERR.Number
        Case Else
11540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11550   End Select
11560   Resume EXITP

End Sub

Public Sub SeparateCheckingOptSet(intMode As Integer, arr_varItem As Variant, blnSeparateChecking_R_Focus As Boolean, blnSeparateChecking_R_MouseDown As Boolean, blnSeparateChecking_L_Focus As Boolean, blnSeparateChecking_L_MouseDown As Boolean, frm As Access.Form)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOptSet"

        ' ** Array: arr_varItem().
        Const I_CNAM As Integer = 0
        Const I_TOP  As Integer = 1
        Const I_LFT  As Integer = 2

11610   With frm

11620     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
11630       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
11640     End If

          ' ** SeparateCheckingOpt:
          ' **   1  Message applies, box is open.
          ' **   2  Message applies, box is closed.
          ' **   3  Not Applicable (Separate Checking is off, or last opgPrint choice wasn't 1).

11650     Select Case intMode
          Case 1
11660       blnSeparateChecking_R_Focus = True
11670       .Controls(arr_varItem(I_CNAM, 1)).Visible = True  ' ** Dots.
11680     Case 2
11690       blnSeparateChecking_R_MouseDown = True
11700       .Controls(arr_varItem(I_CNAM, 0)).Top = (arr_varItem(I_TOP, 0) + lngTpp)
11710       .Controls(arr_varItem(I_CNAM, 0)).Left = (arr_varItem(I_LFT, 0) + lngTpp)
11720       .Controls(arr_varItem(I_CNAM, 1)).Top = (arr_varItem(I_TOP, 1) + lngTpp)
11730       .Controls(arr_varItem(I_CNAM, 1)).Left = (arr_varItem(I_LFT, 1) + lngTpp)
11740     Case 3
            ' ** Open the box.
11750       .FocusHolder.SetFocus
11760       DoEvents
11770       .opgPrint_lbl2.ForeColor = WIN_CLR_DISF
11780       .opgPrint_lbl2_box.BackStyle = acBackStyleTransparent
11790       .opgPrint_lbl3.ForeColor = WIN_CLR_DISF
11800       .opgPrint_lbl3_box.BorderColor = WIN_CLR_DISR
11810       .opgPrint_lbl4.ForeColor = WIN_CLR_DISF
11820       .opgPrint_lbl5.ForeColor = WIN_CLR_DISF
11830       .opgPrint_lbl6.ForeColor = WIN_CLR_DISF
11840       .opgPrint_lbl7.ForeColor = WIN_CLR_DISF
11850       .opgPrint_lbl8.ForeColor = WIN_CLR_DISF
11860       DoEvents
11870       .NextCheckNumber.Enabled = False
11880       .NextCheckNumber.Locked = False
11890       .NextCheckNumber.BorderColor = WIN_CLR_DISR
11900       .NextCheckNumber.BackStyle = acBackStyleTransparent
11910       .NextCheckNumber_box.BackStyle = acBackStyleTransparent
11920       .NextCheckNumber_box2.BackStyle = acBackStyleTransparent
11930       .NextCheckNumber_hline03.BorderColor = MY_CLR_BGE
11940       .ChecksTot.Locked = False
11950       .ChecksTot.BorderColor = WIN_CLR_DISR
11960       .ChecksTot.BackStyle = acBackStyleTransparent
11970       DoEvents
11980       .cmbSortBy.Enabled = False
11990       .cmbSortBy.Locked = False
12000       .cmbSortBy.BorderColor = WIN_CLR_DISR
12010       .cmbSortBy.BackStyle = acBackStyleTransparent
12020       .cmbSortBy_box.BackStyle = acBackStyleTransparent
12030       .cmbSortBy_box2.BackStyle = acBackStyleTransparent
12040       .cmbSortBy_hline03.BorderColor = MY_CLR_BGE
12050       .chkSyncListSort.Enabled = False
12060       .chkSyncListSort.Locked = False
12070       DoEvents
12080       .SeparateCheckingOpt_arw_r_cmd.Visible = False
12090       .SeparateCheckingOpt_arw_r.Visible = False
12100       .SeparateCheckingOpt_arw_r_box.Visible = False
12110       DoEvents
12120       .SeparateCheckingOpt_box.Visible = True
12130       .SeparateCheckingOpt_box2.Visible = True
12140       .SeparateCheckingOpt_arw_l_box.Visible = True
12150       .SeparateCheckingOpt_arw_l.Visible = True
12160       .SeparateCheckingOpt_arw_l_cmd.Visible = True
12170       .SeparateCheckingOpt_arw_l_hline01.Visible = True
12180       .SeparateCheckingOpt_arw_l_hline02.Visible = True
12190       .SeparateCheckingOpt_arw_l_hline03.Visible = True
12200       .SeparateCheckingOpt_arw_l_hline04.Visible = True
12210       .SeparateCheckingOpt_arw_l_vline01.Visible = True
12220       .SeparateCheckingOpt_arw_l_vline02.Visible = True
12230       .SeparateCheckingOpt_lbl1b.Visible = True
12240       .SeparateCheckingOpt_lbl1.Visible = True
12250       .SeparateCheckingOpt_lbl2.Visible = True
12260       .SeparateCheckingOpt_lbl3.Visible = True
12270       .SeparateCheckingOpt_lbl4b.Visible = True
12280       .SeparateCheckingOpt_lbl4.Visible = True
12290       .SeparateCheckingOpt_lbl5.Visible = True
12300       .SeparateCheckingOpt = 1  ' ** Open.
12310       DoEvents
12320     Case 4
12330       If blnSeparateChecking_R_MouseDown = False Then
12340         .Controls(arr_varItem(I_CNAM, 0)).ForeColor = CLR_BLU
12350         .Controls(arr_varItem(I_CNAM, 0) & "_box").BackColor = CLR_VLTBLU2
12360       End If
12370     Case 5
12380       .Controls(arr_varItem(I_CNAM, 0)).Top = arr_varItem(I_TOP, 0)
12390       .Controls(arr_varItem(I_CNAM, 0)).Left = arr_varItem(I_LFT, 0)
12400       .Controls(arr_varItem(I_CNAM, 1)).Top = arr_varItem(I_TOP, 1)
12410       .Controls(arr_varItem(I_CNAM, 1)).Left = arr_varItem(I_LFT, 1)
12420       blnSeparateChecking_R_MouseDown = False
12430     Case 6
12440       .Controls(arr_varItem(I_CNAM, 1)).Visible = False
12450       blnSeparateChecking_R_Focus = False
12460     Case 7
12470       blnSeparateChecking_L_Focus = True
12480       .Controls(arr_varItem(I_CNAM, 3)).Visible = True  ' ** Dots.
12490     Case 8
12500       blnSeparateChecking_L_MouseDown = True
12510       .Controls(arr_varItem(I_CNAM, 2)).Top = (arr_varItem(I_TOP, 2) + lngTpp)
12520       .Controls(arr_varItem(I_CNAM, 2)).Left = (arr_varItem(I_LFT, 2) + lngTpp)
12530       .Controls(arr_varItem(I_CNAM, 3)).Top = (arr_varItem(I_TOP, 3) + lngTpp)
12540       .Controls(arr_varItem(I_CNAM, 3)).Left = (arr_varItem(I_LFT, 3) + lngTpp)
12550     Case 9
            ' ** Close the box.
12560       .FocusHolder.SetFocus
12570       DoEvents
12580       .SeparateCheckingOpt_lbl1.Visible = False
12590       .SeparateCheckingOpt_lbl1b.Visible = False
12600       .SeparateCheckingOpt_lbl2.Visible = False
12610       .SeparateCheckingOpt_lbl3.Visible = False
12620       .SeparateCheckingOpt_lbl4.Visible = False
12630       .SeparateCheckingOpt_lbl4b.Visible = False
12640       .SeparateCheckingOpt_lbl5.Visible = False
12650       .SeparateCheckingOpt_box.Visible = False
12660       .SeparateCheckingOpt_box2.Visible = False
12670       .SeparateCheckingOpt_arw_l_cmd.Visible = False
12680       .SeparateCheckingOpt_arw_l.Visible = False
12690       .SeparateCheckingOpt_arw_l_box.Visible = False
12700       .SeparateCheckingOpt_arw_l_hline01.Visible = False
12710       .SeparateCheckingOpt_arw_l_hline02.Visible = False
12720       .SeparateCheckingOpt_arw_l_hline03.Visible = False
12730       .SeparateCheckingOpt_arw_l_hline04.Visible = False
12740       .SeparateCheckingOpt_arw_l_vline01.Visible = False
12750       .SeparateCheckingOpt_arw_l_vline02.Visible = False
12760       DoEvents
12770       .SeparateCheckingOpt_arw_r_box.Visible = True
12780       .SeparateCheckingOpt_arw_r.Visible = True
12790       .SeparateCheckingOpt_arw_r_cmd.Visible = True
12800       DoEvents
12810       .opgPrint_lbl2.ForeColor = CLR_VDKGRY
12820       .opgPrint_lbl2_box.BackStyle = acBackStyleNormal
12830       .opgPrint_lbl3.ForeColor = CLR_VDKGRY
12840       .opgPrint_lbl3_box.BorderColor = CLR_BLK
12850       .opgPrint_lbl4.ForeColor = CLR_DKGRY3
12860       .opgPrint_lbl5.ForeColor = CLR_DKGRY3
12870       .opgPrint_lbl6.ForeColor = CLR_DKGRY3
12880       .opgPrint_lbl7.ForeColor = CLR_DKGRY3
12890       .opgPrint_lbl8.ForeColor = CLR_VDKGRY
12900       DoEvents
12910       .NextCheckNumber.Enabled = True
12920       .NextCheckNumber.Locked = False
12930       .NextCheckNumber.BorderColor = CLR_LTBLU2
12940       .NextCheckNumber.BackStyle = acBackStyleNormal
12950       .NextCheckNumber_box.BackStyle = acBackStyleNormal
12960       .NextCheckNumber_box2.BackStyle = acBackStyleNormal
12970       .NextCheckNumber_hline03.BorderColor = MY_CLR_LTBGE
12980       .ChecksTot.Locked = True
12990       .ChecksTot.BorderColor = CLR_LTBLU2
13000       .ChecksTot.BackStyle = acBackStyleNormal
13010       DoEvents
13020       .cmbSortBy.Enabled = True
13030       .cmbSortBy.Locked = False
13040       .cmbSortBy.BorderColor = CLR_LTBLU2
13050       .cmbSortBy.BackStyle = acBackStyleNormal
13060       .cmbSortBy_box.BackStyle = acBackStyleNormal
13070       .cmbSortBy_box2.BackStyle = acBackStyleNormal
13080       .cmbSortBy_hline03.BorderColor = MY_CLR_LTBGE
13090       .chkSyncListSort.Enabled = True
13100       .chkSyncListSort.Locked = False
13110       .SeparateCheckingOpt = 2  ' ** Closed.
13120       DoEvents
13130     Case 10
13140       If blnSeparateChecking_L_MouseDown = False Then
13150         .Controls(arr_varItem(I_CNAM, 2)).ForeColor = CLR_BLU
13160         .Controls(arr_varItem(I_CNAM, 2) & "_box").BackColor = CLR_VLTBLU2
13170       End If
13180     Case 11
13190       .Controls(arr_varItem(I_CNAM, 2)).Top = arr_varItem(I_TOP, 2)
13200       .Controls(arr_varItem(I_CNAM, 2)).Left = arr_varItem(I_LFT, 2)
13210       .Controls(arr_varItem(I_CNAM, 3)).Top = arr_varItem(I_TOP, 3)
13220       .Controls(arr_varItem(I_CNAM, 3)).Left = arr_varItem(I_LFT, 3)
13230       blnSeparateChecking_L_MouseDown = False
13240     Case 12
13250       .Controls(arr_varItem(I_CNAM, 3)).Visible = False
13260       blnSeparateChecking_L_Focus = False
13270     End Select
13280   End With

EXITP:
13290   Exit Sub

ERRH:
13300   Select Case ERR.Number
        Case Else
13310     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13320   End Select
13330   Resume EXITP

End Sub

Public Sub LoadAlphaSort()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "LoadAlphaSort"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnContinue As Boolean

13410   blnContinue = True

13420   Set dbs = CurrentDb
13430   Set rst = dbs.OpenRecordset("tmpAccount", dbOpenDynaset, dbReadOnly)
13440   If rst.BOF = True And rst.EOF = True Then
          ' ** Continue.
13450   Else
13460     blnContinue = False
13470   End If
13480   rst.Close
13490   Set rst = Nothing
13500   DoEvents

13510   If blnContinue = True Then
          ' ** Empty tmpAccount.
13520     Set qdf = dbs.QueryDefs("qryPrintChecks_53_01")
13530     qdf.Execute
13540     Set qdf = Nothing
13550     DoEvents
          ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
          ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
          ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
          ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
          ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
          ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
          ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
          ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
13560     Set qdf = dbs.QueryDefs("qryPrintChecks_53_02")
13570     qdf.Execute
13580     Set qdf = Nothing
13590     DoEvents
13600   End If

13610   dbs.Close
13620   Set dbs = Nothing
13630   DoEvents

EXITP:
13640   Set rst = Nothing
13650   Set qdf = Nothing
13660   Set dbs = Nothing
13670   Exit Sub

ERRH:
13680   Select Case ERR.Number
        Case Else
13690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13700   End Select
13710   Resume EXITP

End Sub

Public Sub SubHeadSet_CP(blnShow As Boolean, frm As Access.Form)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "SubHeadSet_CP"

13810   With frm
13820     Select Case blnShow
          Case True
13830       Select Case .ckgDisplay_opt01_AccountNo
            Case True
13840         .Account_Number_lbl.ForeColor = CLR_DKGRY2
13850         .Account_Number_lbl_dim_hi.Visible = False
13860         .Account_Number_lbl_line.BorderColor = CLR_DKGRY2
13870         .Account_Number_lbl_line_dim_hi.Visible = False
13880       Case False
              ' ** Leave as-is.
13890       End Select
13900       Select Case .ckgDisplay_opt02_ShortName
            Case True
13910         .Short_Name_lbl.ForeColor = CLR_DKGRY2
13920         .Short_Name_lbl_dim_hi.Visible = False
13930         .Short_Name_lbl_line.BorderColor = CLR_DKGRY2
13940         .Short_Name_lbl_line_dim_hi.Visible = False
13950       Case False
              ' ** Leave as-is.
13960       End Select
13970       .Check_Count_lbl.ForeColor = CLR_DKGRY2
13980       .Check_Count_lbl_dim_hi.Visible = False
13990       .Check_Count_lbl2.ForeColor = CLR_DKGRY2
14000       .Check_Count_lbl2_dim_hi.Visible = False
14010       .Check_Count_lbl_line.BorderColor = CLR_DKGRY2
14020       .Check_Count_lbl_line_dim_hi.Visible = False
14030       .Last_Check_Number_lbl.ForeColor = CLR_DKGRY2
14040       .Last_Check_Number_lbl_dim_hi.Visible = False
14050       .Last_Check_Number_lbl2.ForeColor = CLR_DKGRY2
14060       .Last_Check_Number_lbl2_dim_hi.Visible = False
14070       .Last_Check_Number_lbl_line.BorderColor = CLR_DKGRY2
14080       .Last_Check_Number_lbl_line_dim_hi.Visible = False
14090       Select Case .ckgDisplay_opt05_Payee
            Case True
14100         .Payee_lbl.ForeColor = CLR_DKGRY2
14110         .Payee_lbl_dim_hi.Visible = False
14120         .Payee_lbl_line.BorderColor = CLR_DKGRY2
14130         .Payee_lbl_line_dim_hi.Visible = False
14140       Case False
              ' ** Leave as-is.
14150       End Select
14160       Select Case .ckgDisplay_opt06_BankName
            Case True
14170         .Bank_Name_lbl.ForeColor = CLR_DKGRY2
14180         .Bank_Name_lbl_dim_hi.Visible = False
14190         .Bank_Name_lbl_line.BorderColor = CLR_DKGRY2
14200         .Bank_Name_lbl_line_dim_hi.Visible = False
14210       Case False
              ' ** Leave as-is.
14220       End Select
14230       Select Case .ckgDisplay_opt07_BankAcctNum
            Case True
14240         .Bank_Account_Number_lbl.ForeColor = CLR_DKGRY2
14250         .Bank_Account_Number_lbl_dim_hi.Visible = False
14260         .Bank_Account_Number_lbl_line.BorderColor = CLR_DKGRY2
14270         .Bank_Account_Number_lbl_line_dim_hi.Visible = False
14280       Case False
              ' ** Leave as-is.
14290       End Select
14300       Select Case .ckgDisplay_opt08_CheckAmount
            Case True
14310         .Check_Amount_lbl.ForeColor = CLR_DKGRY2
14320         .Check_Amount_lbl_dim_hi.Visible = False
14330         .Check_Amount_lbl_line.BorderColor = CLR_DKGRY2
14340         .Check_Amount_lbl_line_dim_hi.Visible = False
14350       Case False
              ' ** Leave as-is.
14360       End Select
14370       Select Case .ChkStat
            Case "Yes", "Mixed"
14380         .Check_Amount_lbl.ForeColor = CLR_DKGRY2
14390         .Check_Amount_lbl_dim_hi.Visible = False
14400         .Check_Amount_lbl_line.BorderColor = CLR_DKGRY2
14410         .Check_Amount_lbl_line_dim_hi.Visible = False
14420       Case "No"
              ' ** Leave as-is.
14430       End Select
14440     Case False
14450       Select Case .ckgDisplay_opt01_AccountNo
            Case True
14460         .Account_Number_lbl.ForeColor = WIN_CLR_DISF
14470         .Account_Number_lbl_dim_hi.Visible = True
14480         .Account_Number_lbl_line.BorderColor = WIN_CLR_DISR
14490         .Account_Number_lbl_line_dim_hi.Visible = True
14500       Case False
14510         .Account_Number_lbl_dim_hi.Visible = False
14520         .Account_Number_lbl_line_dim_hi.Visible = False
14530       End Select
14540       Select Case .ckgDisplay_opt02_ShortName
            Case True
14550         .Short_Name_lbl.ForeColor = WIN_CLR_DISF
14560         .Short_Name_lbl_dim_hi.Visible = True
14570         .Short_Name_lbl_line.BorderColor = WIN_CLR_DISR
14580         .Short_Name_lbl_line_dim_hi.Visible = True
14590       Case False
14600         .Short_Name_lbl_dim_hi.Visible = False
14610         .Short_Name_lbl_line_dim_hi.Visible = False
14620       End Select
14630       .Check_Count_lbl.ForeColor = WIN_CLR_DISF
14640       .Check_Count_lbl_dim_hi.Visible = True
14650       .Check_Count_lbl2.ForeColor = WIN_CLR_DISF
14660       .Check_Count_lbl2_dim_hi.Visible = True
14670       .Check_Count_lbl_line.BorderColor = WIN_CLR_DISR
14680       .Check_Count_lbl_line_dim_hi.Visible = True
14690       .Last_Check_Number_lbl.ForeColor = WIN_CLR_DISF
14700       .Last_Check_Number_lbl_dim_hi.Visible = True
14710       .Last_Check_Number_lbl2.ForeColor = WIN_CLR_DISF
14720       .Last_Check_Number_lbl2_dim_hi.Visible = True
14730       .Last_Check_Number_lbl_line.BorderColor = WIN_CLR_DISR
14740       .Last_Check_Number_lbl_line_dim_hi.Visible = True
14750       Select Case .ckgDisplay_opt05_Payee
            Case True
14760         .Payee_lbl.ForeColor = WIN_CLR_DISF
14770         .Payee_lbl_dim_hi.Visible = True
14780         .Payee_lbl_line.BorderColor = WIN_CLR_DISR
14790         .Payee_lbl_line_dim_hi.Visible = True
14800       Case False
14810         .Payee_lbl_dim_hi.Visible = False
14820         .Payee_lbl_line_dim_hi.Visible = False
14830       End Select
14840       Select Case .ckgDisplay_opt06_BankName
            Case True
14850         .Bank_Name_lbl.ForeColor = WIN_CLR_DISF
14860         .Bank_Name_lbl_dim_hi.Visible = True
14870         .Bank_Name_lbl_line.BorderColor = WIN_CLR_DISR
14880         .Bank_Name_lbl_line_dim_hi.Visible = True
14890       Case False
14900         .Bank_Name_lbl_dim_hi.Visible = False
14910         .Bank_Name_lbl_line_dim_hi.Visible = False
14920       End Select
14930       Select Case .ckgDisplay_opt07_BankAcctNum
            Case True
14940         .Bank_Account_Number_lbl.ForeColor = WIN_CLR_DISF
14950         .Bank_Account_Number_lbl_dim_hi.Visible = True
14960         .Bank_Account_Number_lbl_line.BorderColor = WIN_CLR_DISR
14970         .Bank_Account_Number_lbl_line_dim_hi.Visible = True
14980       Case False
14990         .Bank_Account_Number_lbl_dim_hi.Visible = False
15000         .Bank_Account_Number_lbl_line_dim_hi.Visible = False
15010       End Select
15020       Select Case .ckgDisplay_opt08_CheckAmount
            Case True
15030         .Check_Amount_lbl.ForeColor = WIN_CLR_DISF
15040         .Check_Amount_lbl_dim_hi.Visible = True
15050         .Check_Amount_lbl_line.BorderColor = WIN_CLR_DISR
15060         .Check_Amount_lbl_line_dim_hi.Visible = True
15070       Case False
15080         .Check_Amount_lbl_dim_hi.Visible = False
15090         .Check_Amount_lbl_line_dim_hi.Visible = False
15100       End Select
15110       Select Case .ChkStat
            Case "Yes", "Mixed"
15120         .Check_Amount_lbl.ForeColor = WIN_CLR_DISF
15130         .Check_Amount_lbl_dim_hi.Visible = True
15140         .Check_Amount_lbl_line.BorderColor = WIN_CLR_DISR
15150         .Check_Amount_lbl_line_dim_hi.Visible = True
15160       Case "No"
15170         .Check_Amount_lbl_dim_hi.Visible = False
15180         .Check_Amount_lbl_line_dim_hi.Visible = False
15190       End Select
15200     End Select
15210   End With

EXITP:
15220   Exit Sub

ERRH:
15230   Select Case ERR.Number
        Case Else
15240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15250   End Select
15260   Resume EXITP

End Sub

Public Sub MICRSet_CP(blnShow As Boolean, frm As Access.Form)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "MICRSet_CP"

15310   With frm
15320     Select Case blnShow
          Case True
15330       .cmdMICRSet.Enabled = True
15340       If .cmdMICRSet.Visible = True Then
15350         .cmdMICRSet_raised_img.Visible = True
15360       End If
15370       .cmdMICRSet_raised_semifocus_dots_img.Visible = False
15380       .cmdMICRSet_raised_focus_img.Visible = False
15390       .cmdMICRSet_raised_focus_dots_img.Visible = False
15400       .cmdMICRSet_sunken_focus_dots_img.Visible = False
15410       .cmdMICRSet_raised_img_dis.Visible = False
15420       .cmdMICRAdjust.Enabled = True
15430       .cmdMICRAdjust_raised_img.Visible = True
15440       .cmdMICRAdjust_raised_semifocus_dots_img.Visible = False
15450       .cmdMICRAdjust_raised_focus_img.Visible = False
15460       .cmdMICRAdjust_raised_focus_dots_img.Visible = False
15470       .cmdMICRAdjust_sunken_focus_dots_img.Visible = False
15480       .cmdMICRAdjust_raised_img_dis.Visible = False
15490     Case False
15500       .cmdMICRSet.Enabled = False
15510       If .cmdMICRSet.Visible = True Then
15520         .cmdMICRSet_raised_img_dis.Visible = True
15530       End If
15540       .cmdMICRSet_raised_img.Visible = False
15550       .cmdMICRSet_raised_semifocus_dots_img.Visible = False
15560       .cmdMICRSet_raised_focus_img.Visible = False
15570       .cmdMICRSet_raised_focus_dots_img.Visible = False
15580       .cmdMICRSet_sunken_focus_dots_img.Visible = False
15590       .cmdMICRAdjust.Enabled = False
15600       .cmdMICRAdjust_raised_img_dis.Visible = True
15610       .cmdMICRAdjust_raised_img.Visible = False
15620       .cmdMICRAdjust_raised_semifocus_dots_img.Visible = False
15630       .cmdMICRAdjust_raised_focus_img.Visible = False
15640       .cmdMICRAdjust_raised_focus_dots_img.Visible = False
15650       .cmdMICRAdjust_sunken_focus_dots_img.Visible = False
15660     End Select
15670   End With

EXITP:
15680   Exit Sub

ERRH:
15690   Select Case ERR.Number
        Case Else
15700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15710   End Select
15720   Resume EXITP

End Sub

Public Function FillJournalWithCheckNums_CP(strAccountNo As String, lngStartCheckNum As Long, frm As Access.Form) As Long
' ** If process gets to here, it has been verified that there is at least one check.
' ** CheckNum's now incremented in cmbSort order.
' ** ID's are also appended to tblCheckPrint to assure only checks printed get CheckNum.

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "FillJournalWithCheckNums_CP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strFldName As String, strQryName As String
        Dim blnSync As Boolean, blnDesc As Boolean
        Dim lngChecks As Long, arr_varCheck() As Variant
        Dim lngRecs As Long, lngCheckNum As Long
        Dim lngX As Long, lngE As Long

        ' ** Array: arr_varCheck().
        Const C_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const C_JID    As Integer = 0
        Const C_CHKNUM As Integer = 1

15810   With frm

15820     strFldName = .cmbSortBy.Column(CBX_SORT_FNAM)
15830     blnSync = .chkSyncListSort
15840     blnDesc = .Sort_Dn_img.Visible

          ' ** Set up to start.
15850     lngCheckNum = lngStartCheckNum - 1&

          ' ** These all have CheckNum = Null.
15860     Set dbs = CurrentDb
15870     With dbs

15880       If strAccountNo <> vbNullString Then

15890         Select Case strFldName
              Case "Account_Number"
                ' ** Sort: [Account_Number], [Payee]
                ' ** Journal, just 'Paid', PrintCheck = True, by accountno, RecurringItem, by specified [actno].
15900           strQryName = "qryPrintChecks_40_01_01"
15910         Case "Short_Name"
                ' ** Sort: [Short_Name], [Payee]
                ' ** Journal, just 'Paid', PrintCheck = True, by shortname, RecurringItem, by specified [actno].
15920           strQryName = "qryPrintChecks_40_01_02"
15930         Case "Payee"
                ' ** Sort: [Payee], [Short_Name]
                ' ** Journal, just 'Paid', PrintCheck = True, by RecurringItem, shortname, by specified [actno].
15940           strQryName = "qryPrintChecks_40_01_05"
15950         Case "Bank_Name"
                ' ** Sort: [Bank_Name], [Payee]
                ' ** Journal, just 'Paid', PrintCheck = True, by Bank_Name, RecurringItem, by specified [actno].
15960           strQryName = "qryPrintChecks_40_01_06"
15970         Case "Bank_Account_Number"
                ' ** Sort: [Bank_Account_Number], [Payee]
                ' ** Journal, just 'Paid', PrintCheck = True, by Bank_AccountNumber, RecurringItem, by specified [actno].
15980           strQryName = "qryPrintChecks_40_01_07"
15990         Case "Check_Amount"
                ' ** Sort: [Check_Amount], [Bank_Name]
                ' ** Journal, just 'Paid', PrintCheck = True, by Check_Amount, Bank_Name, by specified [actno].
16000           strQryName = "qryPrintChecks_40_01_08"
16010         End Select
16020         If blnSync = True And blnDesc = True Then
16030           strQryName = strQryName & "d"
16040         End If
16050         Set qdf = .QueryDefs(strQryName)
16060         With qdf.Parameters
16070           ![actno] = strAccountNo
16080         End With

16090       Else

16100         Select Case strFldName
              Case "Account_Number"
                ' ** Sort: [Account_Number], [Payee]
                ' ** Journal, just 'Paid', PrintCheck = True, by accountno, RecurringItem.
16110           strQryName = "qryPrintChecks_40_02_01"
16120         Case "Short_Name"
                ' ** Sort: [Short_Name], [Payee]
                ' ** Journal, just 'Paid', PrintCheck = True, by shortname, RecurringItem.
16130           strQryName = "qryPrintChecks_40_02_02"
16140         Case "Payee"
                ' ** Sort: [Payee], [Short_Name]
                ' ** Journal, just 'Paid', PrintCheck = True, by RecurringItem, shortname.
16150           strQryName = "qryPrintChecks_40_02_05"
16160         Case "Bank_Name"
                ' ** Sort: [Bank_Name], [Payee]
                ' ** Journal, just 'Paid', PrintCheck = True, by Bank_Name, RecurringItem.
16170           strQryName = "qryPrintChecks_40_02_06"
16180         Case "Bank_Account_Number"
                ' ** Sort: [Bank_Account_Number], [Payee]
                ' ** Journal, just 'Paid', PrintCheck = True, by Bank_AccountNumber, RecurringItem.
16190           strQryName = "qryPrintChecks_40_02_07"
16200         Case "Check_Amount"
                ' ** Sort: [Check_Amount], [Bank_Name]
                ' ** Journal, just 'Paid', PrintCheck = True, by Check_Amount, Bank_Name.
16210           strQryName = "qryPrintChecks_40_02_08"
16220         End Select
16230         If blnSync = True And blnDesc = True Then
16240           strQryName = strQryName & "d"
16250         End If
16260         Set qdf = .QueryDefs(strQryName)

16270       End If

16280       lngChecks = 0&
16290       ReDim arr_varCheck(C_ELEMS, 0)

16300       Set rst = qdf.OpenRecordset
16310       With rst
16320         .MoveLast
16330         lngRecs = .RecordCount
16340         .MoveFirst
16350         For lngX = 1& To lngRecs
16360           lngChecks = lngChecks + 1&
16370           lngE = lngChecks - 1&
16380           ReDim Preserve arr_varCheck(C_ELEMS, lngE)
16390           arr_varCheck(C_JID, lngE) = ![ID]
16400           lngCheckNum = lngCheckNum + 1&
16410           .Edit
16420           ![CheckNum] = lngCheckNum
16430           .Update
16440           arr_varCheck(C_CHKNUM, lngE) = lngCheckNum
16450           If lngX < lngRecs Then .MoveNext
16460         Next
16470         .Close
16480       End With
16490       Set rst = Nothing
16500       DoEvents

            ' ** Empty tblCheckPrint.
16510       Set qdf = .QueryDefs("qryPrintChecks_05_32_01")
16520       qdf.Execute
16530       Set qdf = Nothing
16540       DoEvents

            ' ********************************************************
            ' ** Place Journal ID's in tblCheckPrint to assure we'll
            ' ** only print those that DIDN'T have a check number.
            ' ********************************************************

16550       Set rst = .OpenRecordset("tblCheckPrint", dbOpenDynaset, dbAppendOnly)
16560       With rst
16570         For lngX = 0& To (lngChecks - 1&)
16580           .AddNew
                ' ** ![chkprint_id] : AutoNumber.
16590           ![ID] = arr_varCheck(C_JID, lngX)
16600           ![CheckNum] = arr_varCheck(C_CHKNUM, lngX)
16610           ![chkprint_datemodified] = Now()
16620           .Update
16630         Next
16640         .Close
16650       End With
16660       Set rst = Nothing

16670       .Close
16680     End With

16690   End With

EXITP:
16700   Set rst = Nothing
16710   Set qdf = Nothing
16720   Set dbs = Nothing
16730   FillJournalWithCheckNums_CP = lngCheckNum
16740   Exit Function

ERRH:
16750   lngCheckNum = -999&  ' ** Error flag.
16760   Select Case ERR.Number
        Case Else
16770     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16780   End Select
16790   Resume EXITP

End Function

Public Sub Detail_Mouse_CP(blnMICRSet_Focus As Boolean, blnMICRAdjust_Focus As Boolean, blnVoidedChecks_Focus As Boolean, lngItems As Long, arr_varItem As Variant, frm As Access.Form)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Mouse_CP"

16810   With frm

16820     If .cmdMICRSet_raised_focus_dots_img.Visible = True Or .cmdMICRSet_raised_focus_img.Visible = True Then
16830       Select Case blnMICRSet_Focus
            Case True
16840         .cmdMICRSet_raised_semifocus_dots_img.Visible = True
16850         .cmdMICRSet_raised_img.Visible = False
16860       Case False
16870         .cmdMICRSet_raised_img.Visible = True
16880         .cmdMICRSet_raised_semifocus_dots_img.Visible = False
16890       End Select
16900       .cmdMICRSet_raised_focus_img.Visible = False
16910       .cmdMICRSet_raised_focus_dots_img.Visible = False
16920       .cmdMICRSet_sunken_focus_dots_img.Visible = False
16930       .cmdMICRSet_raised_img_dis.Visible = False
16940     End If
16950     If .cmdMICRAdjust_raised_focus_dots_img.Visible = True Or .cmdMICRAdjust_raised_focus_img.Visible = True Then
16960       Select Case blnMICRAdjust_Focus
            Case True
16970         .cmdMICRAdjust_raised_semifocus_dots_img.Visible = True
16980         .cmdMICRAdjust_raised_img.Visible = False
16990       Case False
17000         .cmdMICRAdjust_raised_img.Visible = True
17010         .cmdMICRAdjust_raised_semifocus_dots_img.Visible = False
17020       End Select
17030       .cmdMICRAdjust_raised_focus_img.Visible = False
17040       .cmdMICRAdjust_raised_focus_dots_img.Visible = False
17050       .cmdMICRAdjust_sunken_focus_dots_img.Visible = False
17060       .cmdMICRAdjust_raised_img_dis.Visible = False
17070     End If
17080     If .Controls(arr_varItem(I_CNAM, 0)).ForeColor = CLR_BLU Then
17090       .Controls(arr_varItem(I_CNAM, 0) & "_box").BackColor = MY_CLR_VLTBGE
17100       .Controls(arr_varItem(I_CNAM, 0)).ForeColor = CLR_DKBLU
17110     End If
17120     If .Controls(arr_varItem(I_CNAM, 2)).ForeColor = CLR_BLU Then
17130       .Controls(arr_varItem(I_CNAM, 2) & "_box").BackColor = MY_CLR_VLTBGE
17140       .Controls(arr_varItem(I_CNAM, 2)).ForeColor = CLR_DKBLU
17150     End If
17160     If .cmdVoidedChecks_raised_focus_dots_img.Visible = True Or .cmdVoidedChecks_raised_focus_img.Visible = True Then
17170       Select Case blnVoidedChecks_Focus
            Case True
17180         .cmdVoidedChecks_raised_semifocus_dots_img.Visible = True
17190         .cmdVoidedChecks_raised_img.Visible = False
17200       Case False
17210         .cmdVoidedChecks_raised_img.Visible = True
17220         .cmdVoidedChecks_raised_semifocus_dots_img.Visible = False
17230       End Select
17240       .cmdVoidedChecks_raised_focus_img.Visible = False
17250       .cmdVoidedChecks_raised_focus_dots_img.Visible = False
17260       .cmdVoidedChecks_sunken_focus_dots_img.Visible = False
17270       .cmdVoidedChecks_raised_img_dis.Visible = False
17280     End If

17290   End With

EXITP:
17300   Exit Sub

ERRH:
17310   Select Case ERR.Number
        Case Else
17320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17330   End Select
17340   Resume EXITP

End Sub
