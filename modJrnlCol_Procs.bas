Attribute VB_Name = "modJrnlCol_Procs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modJrnlCol_Procs"

'VGC 07/19/2017: CHANGES!

' ** Combo box column constants: assetno.
'Private Const CBX_AST_ASTNO As Integer = 0  'assetno
Private Const CBX_AST_DESC  As Integer = 1  'totdesc
'Private Const CBX_AST_CUSIP As Integer = 2  'cusip
'Private Const CBX_AST_TYPE  As Integer = 3  'assettype
'Private Const CBX_AST_TAX   As Integer = 4  'taxcode
'Private Const CBX_AST_RATE  As Integer = 5  'rate
'Private Const CBX_AST_DUE   As Integer = 6  'due
Private Const CBX_AST_D4D   As Integer = 7  'AssetType_D4D

' ** Memo field maximum characters.
Private Const MEMO_MAX As Integer = 40
' **

Public Sub JCol_AcctNo_BeforeUpdate(Cancel As Integer, blnContinue As Boolean, strAccountNo_OldValue As String, strShortName_OldValue As String, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_AcctNo_BeforeUpdate(
' **   Cancel As Integer, blnContinue As Boolean, strAccountNo_OldValue As String, strShortName_OldValue As String,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

100   On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_AcctNo_BeforeUpdate"

        Dim blnFound As Boolean
        Dim strTmp01 As String
        Dim intX As Integer

110     With frm

120       If IsNull(.accountno) = True Then
            ' ** Give them the opportunity to use the shortname dropdown instead.
130         blnContinue = False
140       ElseIf strAccountNo_OldValue <> vbNullString Then
150         If strAccountNo_OldValue = .accountno Then
              ' ** Let it be.
160           blnContinue = False
170         Else
180           strTmp01 = .accountno
190           blnFound = False
200           For intX = 0& To (.accountno.ListCount - 1)
210             If .accountno.Column(0, intX) = strTmp01 Then
220               blnFound = True
230               Exit For
240             End If
250           Next
260           Select Case blnFound
              Case True
                ' ** A legitimate accountno, so continue checking.
270           Case False
                ' ** Somehow not an accountno, so let them change it.
280             blnContinue = False
290           End Select
300         End If
310       Else
            ' ** Will be a NullString on a new record, so don't check further.
320         blnContinue = False
330         strAccountNo_OldValue = vbNullString
340         strShortName_OldValue = vbNullString
350       End If

360       If IsNull(.journaltype) = True And IsNull(.ICash) = True And IsNull(.PCash) = True And IsNull(.Cost) = True Then
370         If IsNull(.assetno) = True Then
              ' ** Must be a new record, so don't check further.
380           blnContinue = False
390         Else
400           If .assetno = 0& Then
                ' ** Must be a new record, so don't check further.
410             blnContinue = False
420           End If
430         End If
440       End If

450       If blnContinue = True Then
            ' ** This should mean they've changed the accountno.
460         .posted = False
470         .journaltype = Null
480         .assetno_description = Null
490         .assetno = 0&
500         .Recur_Name = Null
510         .Recur_Type = Null
520         .RecurringItem_ID = Null
530         .assetdate_display = Null
540         .assetdate = Null
550         .shareface = Null
560         .pershare = 0#
570         .ICash = Null
580         .PCash = Null
590         .Cost = Null
600         .Loc_Name_display = Null
610         .Loc_Name = "{no_entry}"
620         .Location_ID = 1&
630         .CheckNum = Null
640         .JrnlMemo_Memo = Null
650         .JrnlMemo_HasMemo = False
660         If .Parent.JrnlMemo_Memo.Visible = True Then
670           JC_Msc_Memo_Set False, .Parent  ' ** Module Procedure: modJrnlCol_Misc.
680           .Parent.JrnlMemo_Memo = Null
690         End If
700         .PrintCheck = False
710         .description = Null
720         .revcode_DESC_display = Null
730         .revcode_DESC = "Unspecified Income"
740         .revcode_ID = REVID_INC
750         .revcode_TYPE = REVTYP_INC
760         .taxcode_description_display = Null
770         .taxcode_description = "Unspecified Income"
780         .taxcode = TAXID_INC
790         .taxcode_type = TAXTYP_INC
800         .Reinvested = False
810         .journal_USER = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
820         .rate = 0#
830         .due = Null
840         .assettype = Null
850         .IsAverage = False
860         .JrnlCol_DateModified = Now()
870       End If

880     End With

EXITP:
890     Exit Sub

ERRH:
900     THAT_PROC = THIS_PROC
910     That_Erl = Erl: That_Desc = ERR.description
920     frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
930     Resume EXITP

End Sub

Public Sub JCol_AcctNo_AfterUpdate(blnDelete As Boolean, strAccountNo_OldValue As String, strShortName_OldValue As String, strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_AcctNo_AfterUpdate(
' **   blnDelete As Boolean, strAccountNo_OldValue As String, strShortName_OldValue As String,
' **   strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_AcctNo_AfterUpdate"

1010    With frm

1020      If .posted = True Then
1030        .posted = False
1040        .posted.Locked = False
1050      End If

1060      If IsNull(.accountno) = False Then
1070        If Trim(.accountno) <> vbNullString Then
1080          If Trim(.accountno) <> strAccountNo_OldValue Then  ' ** Don't keep popping up that box every time
1090            .accountno2.Requery                              ' ** they inadvertantly click the same accountno!
1100            strAccountNo_OldValue = .accountno
1110            strShortName_OldValue = .accountno.Column(1)
1120            .shortname = strShortName_OldValue
1130          Else
1140            .accountno2 = strAccountNo_OldValue  ' ** .Undo simply doesn't work!
1150            .shortname = strShortName_OldValue
1160          End If  ' ** strAccountNo_OldValue.
1170        Else
              ' ** If they've blanked it out, confirm deletion.
1180          If strAccountNo_OldValue <> vbNullString Then
1190            blnDelete = True
1200          End If  ' ** strAccountNo_OldValue
1210        End If  ' ** accountno <> vbNullString.
1220      Else
            ' ** If they've blanked it out, confirm deletion.
1230        If strAccountNo_OldValue <> vbNullString Then
1240          blnDelete = True
1250        End If
1260      End If  ' ** accountno <> Null.

1270      If blnDelete = True Then
            ' ** Give them a chance to use accountno2.
1280        strSaveMoveCtl = "accountno2"
1290        .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
1300      Else

1310        If gblnTabCopyAccount = True And IsNull(.accountno) = False Then
1320          If IsNull(.Parent.LastAcctNo) = True Then
1330            .Parent.LastAcctNo = .accountno
1340          Else
1350            If .Parent.LastAcctNo = vbNullString Then
1360              .Parent.LastAcctNo = .accountno
1370            Else
1380              If .Parent.LastAcctNo <> .accountno Then
1390                .Parent.LastAcctNo = .accountno
1400              End If
1410            End If
1420          End If
1430        End If

            ' ** If these were still active on a new record, disable them now.
1440        JC_Btn_Set Nz(.journaltype, vbNullString), .posted, .Parent  ' ** Module Procedure: modJrnlCol_Buttons.

            ' ** If they've just entered the accountno, no need to stop at shortname.
1450        strSaveMoveCtl = JC_Key_Sub_Next("accountno2_AfterUpdate", blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
1460        blnNoMove = True
1470        .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.

1480      End If  ' ** blnDelete.

1490    End With

EXITP:
1500    Exit Sub

ERRH:
1510    DoCmd.Hourglass False
1520    THAT_PROC = THIS_PROC
1530    That_Erl = Erl: That_Desc = ERR.description
1540    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
1550    Resume EXITP

End Sub

Public Sub JCol_AcctNo2_BeforeUpdate(Cancel As Integer, blnContinue As Boolean, strAccountNo_OldValue As String, strShortName_OldValue As String, blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_AcctNo2_BeforeUpdate(
' **   Cancel As Integer, blnContinue As Boolean, strAccountNo_OldValue As String, strShortName_OldValue As String,
' **   blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_AcctNo2_BeforeUpdate"

        Dim blnFound As Boolean
        Dim strTmp01 As String
        Dim intX As Integer

1610    With frm

1620      If IsNull(.accountno2) = True Then
1630        If gblnClosing = False And gblnDeleting = False Then
1640          Beep
1650          MsgBox "An Account is required to continue.", vbInformation + vbOKOnly, "Entry Required  1"
1660          Cancel = -1
1670          blnNoMove = True
1680        End If
1690      ElseIf strAccountNo_OldValue <> vbNullString Then
1700        If strAccountNo_OldValue = .accountno2 Then
              ' ** Let it be.
1710          blnContinue = False
1720        Else
1730          strTmp01 = .accountno2
1740          blnFound = False
1750          For intX = 0& To (.accountno2.ListCount - 1)
1760            If .accountno2.Column(0, intX) = strTmp01 Then
1770              blnFound = True
1780              Exit For
1790            End If
1800          Next
1810          Select Case blnFound
              Case True
                ' ** A legitimate accountno, so continue checking.
1820          Case False
                ' ** Somehow not an accountno, so let them change it.
1830            blnContinue = False
1840          End Select
1850        End If
1860      Else
            ' ** Will be a NullString on a new record, so don't check further.
1870        blnContinue = False
1880        strAccountNo_OldValue = vbNullString
1890        strShortName_OldValue = vbNullString
1900      End If

1910      If IsNull(.journaltype) = True And IsNull(.ICash) = True And IsNull(.PCash) = True And IsNull(.Cost) = True Then
1920        If IsNull(.assetno) = True Then
              ' ** Must be a new record, so don't check further.
1930          blnContinue = False
1940        Else
1950          If .assetno = 0& Then
                ' ** Must be a new record, so don't check further.
1960            blnContinue = False
1970          End If
1980        End If
1990      End If

2000      If blnContinue = True Then
            ' ** This should mean they've changed the accountno.
2010        .posted = False
2020        .journaltype = Null
2030        .assetno_description = Null
2040        .assetno = 0&
2050        .Recur_Name = Null
2060        .Recur_Type = Null
2070        .RecurringItem_ID = Null
2080        .assetdate_display = Null
2090        .assetdate = Null
2100        .shareface = Null
2110        .pershare = 0#
2120        .ICash = Null
2130        .PCash = Null
2140        .Cost = Null
2150        .Loc_Name_display = Null
2160        .Loc_Name = "{no entry}"
2170        .Location_ID = 1&
2180        .CheckNum = Null
2190        .JrnlMemo_Memo = Null
2200        .JrnlMemo_HasMemo = False
2210        If .Parent.JrnlMemo_Memo.Visible = True Then
2220          JC_Msc_Memo_Set False, .Parent  ' ** Module Procedure: modJrnlCol_Misc.
2230          .Parent.JrnlMemo_Memo = Null
2240        End If
2250        .PrintCheck = False
2260        .description = Null
2270        .revcode_DESC_display = Null
2280        .revcode_DESC = "Unspecified Income"
2290        .revcode_ID = REVID_INC
2300        .revcode_TYPE = REVTYP_INC
2310        .taxcode_description_display = Null
2320        .taxcode_description = "Unspecified Income"
2330        .taxcode = TAXID_INC
2340        .taxcode_type = TAXTYP_INC
2350        .Reinvested = False
2360        .journal_USER = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2370        .rate = 0#
2380        .due = Null
2390        .assettype = Null
2400        .IsAverage = False
2410        .JrnlCol_DateModified = Now()
2420      End If

2430    End With

EXITP:
2440    Exit Sub

ERRH:
2450    THAT_PROC = THIS_PROC
2460    That_Erl = Erl: That_Desc = ERR.description
2470    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
2480    Resume EXITP

End Sub

Public Sub JCol_AcctNo2_AfterUpdate(blnDelete As Boolean, strAccountNo_OldValue As String, strShortName_OldValue As String, strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_AcctNo2_AfterUpdate(
' **   blnDelete As Boolean, strAccountNo_OldValue As String, strShortName_OldValue As String,
' **   strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_AcctNo2_AfterUpdate"

2510    With frm

2520      If .posted = True Then
2530        .posted = False
2540        .posted.Locked = False
2550      End If

2560      If IsNull(.accountno2) = False Then
2570        If Trim(.accountno2) <> vbNullString Then
2580          If Trim(.accountno2) <> strAccountNo_OldValue Then  ' ** Don't keep popping up that box every time
2590            .accountno.Requery                                ' ** they inadvertantly click the same accountno!
2600            strAccountNo_OldValue = .accountno2
2610            strShortName_OldValue = .accountno2.Column(1)
2620            .shortname = strShortName_OldValue
2630          Else
2640            .accountno2 = strAccountNo_OldValue  ' ** .Undo simply doesn't work!
2650            .shortname = strShortName_OldValue
2660          End If  ' ** strAccountNo_OldValue.
2670        Else
              ' ** If they've blanked it out, confirm deletion.
2680          If strAccountNo_OldValue <> vbNullString Then
2690            blnDelete = True
2700          End If  ' ** strAccountNo_OldValue.
2710        End If  ' ** accountno <> vbNullString.
2720      Else
            ' ** If they've blanked it out, confirm deletion.
2730        If strAccountNo_OldValue <> vbNullString Then
2740          blnDelete = True
2750        End If
2760      End If  ' ** accountno <> Null.

2770      If blnDelete = True Then
            ' ** This will be handled OnExit.
2780      Else

2790        If gblnTabCopyAccount = True And IsNull(.accountno2) = False Then
2800          If IsNull(.Parent.LastAcctNo) = True Then
2810            .Parent.LastAcctNo = .accountno2
2820          Else
2830            If .Parent.LastAcctNo = vbNullString Then
2840              .Parent.LastAcctNo = .accountno2
2850            Else
2860              If .Parent.LastAcctNo <> .accountno2 Then
2870                .Parent.LastAcctNo = .accountno2
2880              End If
2890            End If
2900          End If
2910        End If

            ' ** If these were still active on a new record, disable them now.
2920        JC_Btn_Set Nz(.journaltype, vbNullString), .posted, .Parent  ' ** Module Procedure: modJrnlCol_Buttons.

2930        strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
            '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
2940        .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.

2950      End If  ' ** blnDelete.

2960    End With

EXITP:
2970    Exit Sub

ERRH:
2980    DoCmd.Hourglass False
2990    THAT_PROC = THIS_PROC
3000    That_Erl = Erl: That_Desc = ERR.description
3010    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
3020    Resume EXITP

End Sub

Public Sub JCol_JType_BeforeUpdate(Cancel As Integer, blnContinue As Boolean, strJournalType_OldValue As String, blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_JType_BeforeUpdate(
' **   Cancel As Integer, blnContinue As Boolean, strJournalType_OldValue As String, blnNoMove As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_JType_BeforeUpdate"

3110    With frm

3120      If IsNull(.journaltype) = True Then
3130        blnContinue = False
3140        If gblnClosing = False And gblnDeleting = False Then
3150          Beep
3160          MsgBox "A Journal Type is required to continue.", vbInformation + vbOKOnly, "Entry Required  1"
3170          Cancel = -1
3180          blnNoMove = True
3190          .journaltype.Undo
3200        End If
3210      ElseIf strJournalType_OldValue <> vbNullString Then
3220        If strJournalType_OldValue = .journaltype Then
              ' ** Let it be.
3230          blnContinue = False
3240        Else
3250          Select Case strJournalType_OldValue
              Case "Dividend", "Interest", "Purchase", "Deposit", "Sold", "Withdrawn", _
                  "Misc.", "Paid", "Received", "Liability (+)", "Liability (-)", "Cost Adj."
                ' ** A legitimate JournalType, so continue checking.
3260          Case Else
                ' ** Somehow not a JournalType, so let them change it.
3270            blnContinue = False
3280          End Select
3290        End If
3300      Else
            ' ** Will be a NullString on a new record, so don't check further.
3310        blnContinue = False
3320        strJournalType_OldValue = vbNullString
3330      End If

3340      If IsNull(.ICash) = True And IsNull(.PCash) = True And IsNull(.Cost) = True Then
            ' ** Must be a new record, so don't check further.
3350        blnContinue = False
3360      End If

3370      If blnContinue = True Then
            ' ** This should mean they've changed the JournalType.
3380        .posted = False
3390        .assetno_description = Null
3400        .assetno = 0&
3410        .Recur_Name = Null
3420        .Recur_Type = Null
3430        .RecurringItem_ID = Null
3440        .assetdate_display = Null
3450        .assetdate = Null
3460        .shareface = Null
3470        .pershare = 0#
3480        .ICash = Null
3490        .PCash = Null
3500        .Cost = Null
3510        .Loc_Name_display = Null
3520        .Loc_Name = "{no entry}"
3530        .Location_ID = 1&
3540        .CheckNum = Null
3550        .JrnlMemo_Memo = Null
3560        .JrnlMemo_HasMemo = False
3570        If .Parent.JrnlMemo_Memo.Visible = True Then
3580          JC_Msc_Memo_Set False, .Parent  ' ** Module Procedure: modJrnlCol_Misc.
3590          .Parent.JrnlMemo_Memo = Null
3600        End If
3610        .PrintCheck = False
3620        .description = Null
3630        .revcode_DESC_display = Null
3640        .revcode_DESC = "Unspecified Income"
3650        .revcode_ID = REVID_INC
3660        .revcode_TYPE = REVTYP_INC
3670        .taxcode_description_display = Null
3680        .taxcode_description = "Unspecified Income"
3690        .taxcode = TAXID_INC
3700        .taxcode_type = TAXTYP_INC
3710        .Reinvested = False
3720        .journal_USER = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
3730        .rate = 0#
3740        .due = Null
3750        .assettype = Null
3760        .IsAverage = False
3770        .JrnlCol_DateModified = Now()
3780      End If

3790    End With

EXITP:
3800    Exit Sub

ERRH:
3810    THAT_PROC = THIS_PROC
3820    That_Erl = Erl: That_Desc = ERR.description
3830    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
3840    Resume EXITP

End Sub

Public Sub JCol_JType_AfterUpdate(blnJTypeSet As Boolean, strSaveMoveCtl As String, blnWarnZeroCost As Boolean, blnReinvestment As Boolean, blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_JType_AfterUpdate(
' **   blnJTypeSet As Boolean, strSaveMoveCtl As String, blnWarnZeroCost As Boolean, blnReinvestment As Boolean,
' **   blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_JType_AfterUpdate"

        Dim strThisJType As String

        'MAKE SURE ALL OF THESE PROCEDURES ARE GETTING
        'THE ORIGINAL CALLING PROC, AND NOT SOME INTERMEDIARY!

3910    With frm

3920      blnJTypeSet = False

3930      If .posted = True Then
3940        .posted = False
3950        .posted.Locked = False
3960      End If

3970      If IsNull(.journaltype) = False And gblnClosing = False And gblnDeleting = False Then

3980        strThisJType = .journaltype
3990        blnWarnZeroCost = False

4000        .journaltype_sortorder = CLng(.journaltype.Column(2))

4010        Select Case strThisJType
            Case "Misc."
4020          .Recur_Type = "Misc"
4030          .Recur_Name.RowSource = "qryJournal_Columns_10_RecurringItems_02"
4040          .revcode_ID.DefaultValue = REVID_INC  ' ** Unspecified Income.
4050        Case "Paid"
4060          .Recur_Type = "Payee"
4070          .Recur_Name.RowSource = "qryJournal_Columns_10_RecurringItems_03"
4080          .revcode_ID.DefaultValue = REVID_EXP  ' ** Unspecified Expense.
4090        Case "Received"
4100          .Recur_Type = "Payor"
4110          .Recur_Name.RowSource = "qryJournal_Columns_10_RecurringItems_04"
4120          .revcode_ID.DefaultValue = REVID_INC  ' ** Unspecified Income.
4130        Case "Dividend"
4140          .revcode_ID.DefaultValue = REVID_ORDDIV  ' ** Ordinary Dividend.
4150        Case "Interest"
4160          .revcode_ID.DefaultValue = REVID_INTINC  ' ** Interest Income.
4170        Case "Purchase", "Deposit", "Sold", "Withdrawn", "Cost Adj."
4180          .revcode_ID.DefaultValue = REVID_INC  ' ** Unspecified Income.
4190        Case "Liability (+)", "Liability (-)"
4200          .revcode_ID.DefaultValue = REVID_EXP  ' ** Unspecified Expense.
4210        Case "Deposit", "Purchase", "Withdrawn", "Sold"
              ' ** I can't figure out where this is handled!
4220          .assetno.Locked = False
4230        Case Else
4240          .Recur_Type = "Misc"
4250          .Recur_Name.RowSource = "qryJournal_Columns_10_RecurringItems_01"  ' ** All, design view.
4260          .revcode_ID.DefaultValue = REVID_INC  ' ** Unspecified Income.
4270        End Select

4280        If IsNull(.PCash) = True Or blnReinvestment = True Then
              ' ** Must be a new record, so don't check further.
4290          blnJTypeSet = True
4300        ElseIf .PCash = 0 Then
4310          Select Case strThisJType
              Case "Purchase", "Sold"
4320            MsgBox "Principal Cash must be " & IIf(strThisJType = "Purchase", "negative", "positive") & " " & _
                  "for a " & strThisJType & " transaction.", vbInformation + vbOKOnly, "Entry Required"
4330          Case Else
                ' ** Nothing at the moment.
4340          End Select
4350        Else
4360          Select Case strThisJType
              Case "Deposit", "Withdrawn"
4370            MsgBox "Principal Cash must be ZERO " & _
                  "for a " & strThisJType & " transaction.", vbInformation + vbOKOnly, "Invalid Entry"
4380          Case Else
                ' ** Nothing at the moment.
4390          End Select
4400        End If

4410        Select Case blnReinvestment
            Case True
4420          strSaveMoveCtl = "assetno"
4430        Case False
4440          strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
              '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
4450        End Select
4460        blnNoMove = True
4470        .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.

4480      Else
4490        Beep
4500        MsgBox "A Journal Type is required to continue.", vbInformation + vbOKOnly, "Entry Required  2"
4510        blnNoMove = True
4520        .journaltype.SetFocus
4530      End If

4540    End With

EXITP:
4550    Exit Sub

ERRH:
4560    THAT_PROC = THIS_PROC
4570    That_Erl = Erl: That_Desc = ERR.description
4580    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
4590    Resume EXITP

End Sub

Public Sub JCol_Astno_AfterUpdate(blnDelete As Boolean, blnReinvestment As Boolean, lngAssetNo_OldValue As Long, strAssetDesc_OldValue As String, strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_Astno_AfterUpdate(
' **   blnDelete As Boolean, blnReinvestment As Boolean, lngAssetNo_OldValue As Long,
' **   strAssetDesc_OldValue As String, strSaveMoveCtl As String, blnNextRec As Boolean,
' **   blnFromZero As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_Astno_AfterUpdate"

        Dim strAccountNo As String
        Dim lngLocID As Long, strLocName As String

4610    With frm

4620      If .posted = True Then
4630        .posted = False
4640        .posted.Locked = False
4650      End If

4660      If IsNull(.assetno) = False Then
4670        If .assetno > 0& Then
4680          If .assetno <> lngAssetNo_OldValue Then
4690            strAccountNo = Nz(.accountno, vbNullString)
4700            lngAssetNo_OldValue = .assetno
4710            strAssetDesc_OldValue = .assetno.Column(CBX_AST_DESC)
4720            .assetno_description = strAssetDesc_OldValue
4730            If gblnLocSuggest = True Then
4740              lngLocID = JC_Msc_Loc_Set(.accountno, lngAssetNo_OldValue)  ' ** Module Function: modJrnlCol_Misc.
4750              .Location_ID = lngLocID
4760              strLocName = DLookup("[Loc_Name]", "qryJournal_Columns_10_Location_01", "[Location_ID] = " & CStr(lngLocID))
4770              .Loc_Name = strLocName
4780              If lngLocID > 1& Then
4790                .Loc_Name_display = strLocName
4800              Else
4810                .Loc_Name_display = vbNullString
4820              End If
4830            End If
                ' ** If this is a change, blank out rest of line.
4840            If blnReinvestment = False Then
4850              If IsNull(.shareface) = False Then
4860                If .shareface <> 0# Then
4870                  .shareface = 0#
4880                End If
4890              End If
4900              If IsNull(.ICash) = False Then
4910                If .ICash <> 0@ Then
4920                  .ICash = 0@
4930                End If
4940              End If
4950              If IsNull(.PCash) = False Then
4960                If .PCash <> 0@ Then
4970                  .PCash = 0@
4980                End If
4990              End If
5000              If IsNull(.Cost) = False Then
5010                If .Cost <> 0@ Then
5020                  .Cost = 0@
5030                End If
5040              End If
5050            End If
5060          End If  ' ** lngAssetNo_OldValue.
5070        Else
              ' ** If they've blanked it out, confirm deletion.
5080          If lngAssetNo_OldValue > 0& Then
5090            blnDelete = True
5100          End If  ' ** lngAssetNo_OldValue.
5110        End If  ' ** assetno = 0.
5120      Else
            ' ** If they've blanked it out, confirm deletion.
5130        If lngAssetNo_OldValue > 0& Then
5140          blnDelete = True
5150        Else
5160          .assetno = 0&
5170          .assetno_description = Null
5180        End If  ' ** lngAssetNo_OldValue.
5190      End If  ' ** assetno = Null.

5200      If blnDelete = True Then
5210        .assetno = 0&  'lngAssetNo_OldValue
5220        .assetno_description = Null  'strAssetDesc_OldValue
5230        If IsNull(.shareface) = False Then
5240          If .shareface <> 0# Then
5250            .shareface = 0#
5260          End If
5270        End If
5280        If IsNull(.ICash) = False Then
5290          If .ICash <> 0@ Then
5300            .ICash = 0@
5310          End If
5320        End If
5330        If IsNull(.PCash) = False Then
5340          If .PCash <> 0@ Then
5350            .PCash = 0@
5360          End If
5370        End If
5380        If IsNull(.Cost) = False Then
5390          If .Cost <> 0@ Then
5400            .Cost = 0@
5410          End If
5420        End If
5430      End If  ' ** blnDelete.

5440      JC_Msc_ChkAverage frm  ' ** Module Procedure: modJrnlCol_Misc.

5450      strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
          '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
5460      .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.

5470    End With

EXITP:
5480    Exit Sub

ERRH:
5490    DoCmd.Hourglass False
5500    THAT_PROC = THIS_PROC
5510    That_Erl = Erl: That_Desc = ERR.description
5520    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
5530    Resume EXITP

End Sub

Public Sub JCol_ShareFace_Exit(Cancel As Integer, blnToTaxLot As Boolean, blnGoingToReport As Boolean, blnGoneToReport As Boolean, blnNextRec As Boolean, blnFromZero As Boolean, datPostingDate As Date, strSaveMoveCtl As String, blnNoMove As Boolean, lngNewJrnlColID As Long, lngGTR_ID As Long, blnGTR_Emblem As Boolean, blnGTR_NoAdd As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_ShareFace_Exit(
' **   Cancel As Integer, blnToTaxLot As Boolean, blnGoingToReport As Boolean, blnGoneToReport As Boolean,
' **   blnNextRec As Boolean, blnFromZero As Boolean, datPostingDate As Date, strSaveMoveCtl As String,
' **   blnNoMove As Boolean, lngNewJrnlColID As Long, lngGTR_ID As Long, blnGTR_Emblem As Boolean,
' **   blnGTR_NoAdd As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_ShareFace_Exit"

        Dim strThisJType As String, strAccountNo As String
        Dim lngRetVal As Long

5610    With frm

5620      If .shareface.Locked = False Then

5630        strThisJType = Nz(.journaltype, vbNullString)
5640        strAccountNo = Nz(.accountno, vbNullString)

5650        Select Case strThisJType
            Case "Deposit", "Purchase", "Withdrawn", "Sold", "Liability (+)", "Liability (-)"
5660          Select Case strAccountNo
              Case "INCOME O/U", "99-INCOME O/U"
                ' ** OK to have Zero shareface.
                ' ** blnNoMove is True if shareface entered, so TaxLot via KeyDown not triggered!
5670            If strThisJType = "Withdrawn" Then
5680              If .Cost = 0 And .posted = False And blnToTaxLot = False And .shareface <> 0# Then
5690                If Nz(.ICash, 0) <> 0@ Or Nz(.PCash, 0) <> 0@ Then
5700                  If gblnClosing = False And gblnDeleting = False Then
5710                    Beep
5720                    MsgBox "No cash is allowed for a Withdrawn transaction.", vbInformation + vbOKOnly, "Invalid Entry"
5730                    Cancel = -1
5740                  End If
5750                ElseIf .ICash = 0@ And .PCash = 0@ And .Cost = 0@ Then
                      ' ** This should mean it hasn't gone to the Tax Lot form yet.
5760                  .Parent.TaxLotFrom = "shareface"
                      'TaxLot_Form True  ' ** Procedure: Below.
5770                  JC_Frm_TaxLot_Form frm, True, blnToTaxLot, blnGoingToReport, blnGoneToReport, _
                        blnNextRec, blnFromZero, datPostingDate, strSaveMoveCtl, blnNoMove, lngNewJrnlColID, _
                        lngGTR_ID, blnGTR_Emblem, blnGTR_NoAdd  ' ** Module Procedure: modJrnlCol_Forms.
5780                End If
5790              End If
5800            End If
5810          Case Else
5820            If Nz(.shareface, 0) = 0 Then
5830              If gblnClosing = False And gblnDeleting = False Then
5840                Beep
5850                MsgBox "A Share/Face value is required to continue.", vbInformation + vbOKOnly, "Entry Required"
5860                Cancel = -1
5870              End If
5880            Else
5890              If .shareface < 0# Then
5900                .shareface = Abs(.shareface)
5910              End If
5920              If strThisJType = "Withdrawn" Then
5930                If .Cost = 0 And .posted = False And blnToTaxLot = False And .shareface <> 0# Then
5940                  If Nz(.ICash, 0) <> 0@ Or Nz(.PCash, 0) <> 0@ Then
5950                    If gblnClosing = False And gblnDeleting = False Then
5960                      Beep
5970                      MsgBox "No cash is allowed for a Withdrawn transaction.", vbInformation + vbOKOnly, "Invalid Entry"
5980                      Cancel = -1
5990                    End If
6000                  ElseIf .ICash = 0@ And .PCash = 0@ And .Cost = 0@ Then
                        ' ** This should mean it hasn't gone to the Tax Lot form yet.
6010                    .Parent.TaxLotFrom = "shareface"
                        'TaxLot_Form True  ' ** Procedure: Below.
6020                    JC_Frm_TaxLot_Form frm, True, blnToTaxLot, blnGoingToReport, blnGoneToReport, _
                          blnNextRec, blnFromZero, datPostingDate, strSaveMoveCtl, blnNoMove, lngNewJrnlColID, _
                          lngGTR_ID, blnGTR_Emblem, blnGTR_NoAdd  ' ** Module Procedure: modJrnlCol_Forms.
6030                  End If
6040                End If
6050              End If
6060            End If
6070          End Select
6080        Case "Received"
6090          If .assetno > 0 Then
6100            Select Case strAccountNo
                Case "INCOME O/U", "99-INCOME O/U"
                  ' ** OK to have Zero shareface.
6110            Case Else
6120              If Nz(.shareface, 0) = 0 Then
6130                If gblnClosing = False And gblnDeleting = False Then
6140                  Beep
6150                  MsgBox "A Share/Face value is required to continue.", vbInformation + vbOKOnly, "Entry Required"
6160                  Cancel = -1
6170                End If
6180              End If
6190            End Select
6200          End If
6210        Case Else
6220          If IsNull(.shareface) = True Then
6230            .shareface = 0&
6240            strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
                '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
6250            .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
6260          End If
6270        End Select
6280      End If

          ' ** On Rich's machine, it doesn't seem to scroll all the way to the right!
6290      If .journaltype = "Dividend" Or .journaltype = "Interest" Or .journaltype = "Purchase" Then
6300        lngRetVal = fSetScrollBarPosHZ(frm, 999&)  ' ** Module Function: modScrollBarFuncs.
6310      End If

6320    End With

EXITP:
6330    Exit Sub

ERRH:
6340    THAT_PROC = THIS_PROC
6350    That_Erl = Erl: That_Desc = ERR.description
6360    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
6370    Resume EXITP

End Sub

Public Sub JCol_ICash_AfterUpdate(strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_ICash_AfterUpdate(
' **   strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_ICash_AfterUpdate"

        Dim strThisJType As String, strRecurName As String

6410    With frm

6420      If .posted = True Then
6430        .posted = False
6440        .posted.Locked = False
6450      End If

6460      If IsNull(.ICash) = True Then
6470        .ICash = 0@
6480      End If

6490      strThisJType = Nz(.journaltype, vbNullString)
6500      Select Case strThisJType
          Case "Purchase", "Liability (-)", "Paid"
6510        If .ICash > 0@ Then
6520          .ICash = -(.ICash)
6530        End If
6540      Case "Sold"
6550        If .ICash < 0@ Then
6560          .ICash = Abs(.ICash)  ' ** Not sure about this!
6570        End If
6580      Case "Misc."
6590        strRecurName = Nz(.Recur_Name, vbNullString)
6600        Select Case strRecurName
            Case RECUR_I_TO_P
6610          If .ICash < 0@ Then
6620            .PCash = Abs(.ICash)
6630          ElseIf .ICash > 0@ Then
6640            .ICash = -(.ICash)
6650            .PCash = Abs(.ICash)
6660          End If
6670        Case RECUR_P_TO_I
6680          If .ICash > 0@ Then
6690            .PCash = -(.ICash)
6700          ElseIf .ICash < 0@ Then
6710            .ICash = Abs(.ICash)
6720            .PCash = -(.ICash)
6730          End If
6740        End Select
6750      Case "Liability (+)", "Cost Adj.", "Deposit", "Withdrawn"
6760        If .ICash <> 0@ Then
6770          .ICash = 0@
6780          MsgBox "Income Cash must be Zero for a " & strThisJType & " transaction.", vbInformation + vbOKOnly, "Invalid Entry"
6790        End If
6800      Case Else
            ' ** Not sure what else.
6810      End Select

6820      gstrSaleICash = CStr(CDbl(Val(Nz(.ICash.text, 0))))
6830      gstrSaleICash = Rem_Dollar(gstrSaleICash)  ' ** Module Function: modStringFuncs.

6840      strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
          '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
6850      blnNoMove = True
6860      .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.

6870    End With

EXITP:
6880    Exit Sub

ERRH:
6890    THAT_PROC = THIS_PROC
6900    That_Erl = Erl: That_Desc = ERR.description
6910    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
6920    Resume EXITP

End Sub

Public Sub JCol_PCash_AfterUpdate(strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_PCash_AfterUpdate(
' **   strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_PCash_AfterUpdate"

        Dim strThisJType As String, strRecurName As String

7010    With frm

7020      If .posted = True Then
7030        .posted = False
7040        .posted.Locked = False
7050      End If

7060      If IsNull(.PCash) = True Then
7070        .PCash = 0@
7080      End If

7090      strThisJType = Nz(.journaltype, vbNullString)
7100      Select Case strThisJType
          Case "Purchase", "Liability (-)"
7110        If .PCash > 0@ Then
7120          .PCash = -(.PCash)
7130          .Cost = Abs(.PCash)
7140        End If
7150        If .ICash <> 0@ And .PCash <> 0@ Then
7160          If .assetno.Column(CBX_AST_D4D) = True Then
7170            MsgBox "A dollar-for-dollar asset may not purchased" & vbCrLf & _
                  "with both Income Cash and Principal Cash.", vbInformation + vbOKOnly, "Invalid Entry"
7180            .PCash = 0@
7190          End If
7200        End If
7210      Case "Sold"
7220        If .PCash < 0@ Then
7230          .PCash = Abs(.PCash)
7240        End If
7250      Case "Liability (+)"
7260        If .PCash <> 0@ Then
7270          If .PCash < 0@ Then
7280            .PCash = Abs(.PCash)
7290          End If
7300          .Cost = -(.PCash)
7310        Else
7320          .Cost = 0@
7330        End If
7340      Case "Misc."
7350        strRecurName = Nz(.Recur_Name, vbNullString)
7360        Select Case strRecurName
            Case RECUR_I_TO_P
7370          If .PCash > 0@ Then
7380            If .ICash = 0@ Then
7390              .ICash = -(.PCash)
7400            End If
7410          ElseIf .PCash < 0@ Then
7420            .PCash = Abs(.PCash)
7430            If .ICash = 0@ Then
7440              .ICash = -(.PCash)
7450            End If
7460          End If
7470        Case RECUR_P_TO_I
7480          If .PCash < 0@ Then
7490            If .ICash = 0@ Then
7500              .ICash = Abs(.PCash)
7510            End If
7520          ElseIf .PCash > 0@ Then
7530            .PCash = -(.PCash)
7540            If .ICash = 0@ Then
7550              .ICash = Abs(.PCash)
7560            End If
7570          End If
7580        End Select
7590      Case "Paid"
7600        If .PCash > 0@ Then
7610          .PCash = -(.PCash)
7620        End If
7630      Case "Received"
7640        If .PCash < 0@ Then
7650          .PCash = Abs(.PCash)
7660        End If
7670      Case "Dividend", "Interest", "Cost Adj.", "Deposit", "Withdrawn"
7680        If .PCash <> 0@ Then
7690          .PCash = 0@
7700          MsgBox "Principal Cash must be Zero for a " & strThisJType & " transaction.", vbInformation + vbOKOnly, "Invalid Entry"
7710        End If
7720      Case Else
            ' ** Not sure what else.
7730      End Select

7740      gstrSalePCash = CStr(CDbl(Val(Nz(.PCash.text, 0))))
7750      gstrSalePCash = Rem_Dollar(gstrSalePCash)  ' ** Module Function: modStringFuncs.

7760      strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
          '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
7770      blnNoMove = True
7780      .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.

7790    End With

EXITP:
7800    Exit Sub

ERRH:
7810    THAT_PROC = THIS_PROC
7820    That_Erl = Erl: That_Desc = ERR.description
7830    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
7840    Resume EXITP

End Sub

Public Sub JCol_PCash_Exit(Cancel As Integer, blnToTaxLot As Boolean, blnWarnZeroCost As Boolean, blnGoingToReport As Boolean, blnGoneToReport As Boolean, blnNextRec As Boolean, blnFromZero As Boolean, datPostingDate As Date, strSaveMoveCtl As String, blnNoMove As Boolean, lngNewJrnlColID As Long, lngGTR_ID As Long, blnGTR_Emblem As Boolean, blnGTR_NoAdd As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_PCash_Exit(
' **   Cancel As Integer, blnToTaxLot As Boolean, blnWarnZeroCost As Boolean, blnGoingToReport As Boolean,
' **   blnGoneToReport As Boolean, blnNextRec As Boolean, blnFromZero As Boolean, datPostingDate As Date,
' **   strSaveMoveCtl As String, blnNoMove As Boolean, lngNewJrnlColID As Long, lngGTR_ID As Long,
' **   blnGTR_Emblem As Boolean, blnGTR_NoAdd As Boolean, THAT_PROC As String,
' **   That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_PCash_Exit"

        Dim strThisJType As String, strRecurName As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnZeroCash As Boolean
        Dim blnRetVal As Boolean
        Dim strTmp03 As String

7910    With frm

7920      strThisJType = Nz(.journaltype, vbNullString)
7930      If .PCash.Locked = False Then
7940        If IsNull(.PCash) = True Then
7950          .PCash = 0@
7960        End If
7970        If .Cost = 0 And .posted = False And blnToTaxLot = False And .shareface <> 0# Then
7980          Select Case strThisJType
              Case "Deposit"  ' ** Shouldn't be here anyway!
7990            If .ICash <> 0@ Or .PCash <> 0@ Then
8000              If gblnClosing = False And gblnDeleting = False Then
8010                Beep
8020                MsgBox "No cash is allowed for a " & strThisJType & " transaction.", vbInformation + vbOKOnly, "Invalid Entry"
8030                Cancel = -1
8040              End If
8050            End If
8060          Case "Purchase"
8070            If .ICash = 0@ And .PCash = 0@ Then
8080              strTmp03 = .PCash.text
8090              If Val(strTmp03) = 0 Then
8100                If gblnClosing = False And gblnDeleting = False Then
8110                  Beep
8120                  MsgBox "A Principal Cash value is required for a " & strThisJType & " transaction.", _
                        vbInformation + vbOKOnly, "Entry Required"  ' ** Even if it's Zero.
8130                  Cancel = -1
8140                End If
8150              End If
8160            End If
8170          Case "Withdrawn"
8180            If .ICash <> 0@ Or .PCash <> 0@ Then
8190              If gblnClosing = False And gblnDeleting = False Then
8200                Beep
8210                MsgBox "No cash is allowed for a " & strThisJType & " transaction.", vbInformation + vbOKOnly, "Invalid Entry"
8220                Cancel = -1
8230              End If
8240            ElseIf .ICash = 0@ And .PCash = 0@ And .Cost = 0@ Then
                  ' ** This should mean it hasn't gone to the Tax Lot form yet.
                  ' ** Usually handled in JC_Key_Sub(), intAux = 3, via KeyDown() event, above.
8250              If gblnClosing = False And gblnDeleting = False Then
8260                .Parent.TaxLotFrom = "pcash"
                    'TaxLot_Form True  ' ** Procedure: Below.
8270                JC_Frm_TaxLot_Form frm, True, blnToTaxLot, blnGoingToReport, blnGoneToReport, _
                      blnNextRec, blnFromZero, datPostingDate, strSaveMoveCtl, blnNoMove, lngNewJrnlColID, _
                      lngGTR_ID, blnGTR_Emblem, blnGTR_NoAdd  ' ** Module Procedure: modJrnlCol_Forms.
8280              End If
8290            End If
8300          Case "Sold"
8310            blnZeroCash = False: blnRetVal = True
8320            If .ICash = 0@ And .PCash = 0@ Then
8330              strTmp03 = .PCash.text
8340              If Val(strTmp03) = 0 Then
8350                If gblnClosing = False And gblnDeleting = False Then
8360                  msgResponse = MsgBox("Are you sure you want Income and Principal Cash to be ZERO?" & vbCrLf & vbCrLf & _
                        "As would be the case for the sale of worthless shares.", vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
8370                  If msgResponse = vbYes Then
8380                    blnZeroCash = True  ' ** Local to this procedure only.
                        ' ** blnWarnZeroCost   : Not checked during Commit.
8390                  Else
8400                    blnRetVal = False
8410                    Cancel = -1
8420                  End If
8430                End If
8440              Else
                    ' ** Proceed to TaxLot.
8450              End If
8460            ElseIf .PCash = 0@ Then
8470              strTmp03 = .PCash.text
8480              If Val(strTmp03) = 0 Then
8490                If gblnClosing = False And gblnDeleting = False Then
8500                  msgResponse = MsgBox("Are you sure you want Principal Cash to be ZERO?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
8510                  If msgResponse = vbYes Then
8520                    blnZeroCash = True  ' ** Local to this procedure only.
                        ' ** blnWarnZeroCost   : Not checked during Commit.
8530                  Else
8540                    blnRetVal = False
8550                    Cancel = -1
8560                  End If
8570                End If
8580              Else
                    ' ** Proceed to TaxLot.
8590              End If
8600            End If
8610            If .ICash > 0@ Or .PCash > 0@ Or blnZeroCash = True And blnRetVal = True Then
                  ' ** This should mean it hasn't gone to the Tax Lot form yet.
                  ' ** Usually handled in JC_Key_Sub(), intAux = 3, via KeyDown() event, above.
8620              If gblnClosing = False And gblnDeleting = False Then
8630                .Parent.TaxLotFrom = "pcash"
                    'TaxLot_Form True  ' ** Procedure: Below.
8640                JC_Frm_TaxLot_Form frm, True, blnToTaxLot, blnGoingToReport, blnGoneToReport, _
                      blnNextRec, blnFromZero, datPostingDate, strSaveMoveCtl, blnNoMove, lngNewJrnlColID, _
                      lngGTR_ID, blnGTR_Emblem, blnGTR_NoAdd  ' ** Module Procedure: modJrnlCol_Forms.
8650              End If
8660            End If
8670          Case "Liability (+)"
                ' ** Other checks will cover this.
8680            If .PCash = 0@ Then
8690              strTmp03 = .PCash.text
8700              If Val(strTmp03) = 0 Then
8710                Beep
8720                msgResponse = MsgBox("Are you sure you want Principal Cash to be ZERO?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
8730                If msgResponse <> vbYes Then
8740                  Cancel = -1
8750                Else
8760                  blnWarnZeroCost = True
8770                End If
8780              End If
8790            End If
8800          Case "Liability (-)"
8810            blnRetVal = True
8820            If .PCash = 0@ Then
8830              strTmp03 = .PCash.text
8840              If Val(strTmp03) = 0 Then
8850                If gblnClosing = False And gblnDeleting = False Then
8860                  Beep
8870                  msgResponse = MsgBox("Are you sure you want Principal Cash to be ZERO?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cash Basis")
8880                  If msgResponse <> vbYes Then
8890                    blnRetVal = False
8900                    Cancel = -1
8910                  Else
8920                    blnWarnZeroCost = True
8930                  End If
8940                End If
8950              Else
                    ' ** Proceed to TaxLot.
8960              End If
8970            End If
8980            If blnRetVal = True Then
                  ' ** Usually handled in JC_Key_Sub(), intAux = 3, via KeyDown() event, above.
8990              If gblnClosing = False And gblnDeleting = False Then
9000                .Parent.TaxLotFrom = "pcash"
                    'TaxLot_Form True  ' ** Procedure: Below.
9010                JC_Frm_TaxLot_Form frm, True, blnToTaxLot, blnGoingToReport, blnGoneToReport, _
                      blnNextRec, blnFromZero, datPostingDate, strSaveMoveCtl, blnNoMove, lngNewJrnlColID, _
                      lngGTR_ID, blnGTR_Emblem, blnGTR_NoAdd  ' ** Module Procedure: modJrnlCol_Forms.
9020              End If
9030            End If
9040          Case Else
                ' ** Nothing at the moment.
9050          End Select
9060        ElseIf .Cost = 0 And .posted = False And blnToTaxLot = False And .shareface = 0# Then
9070          Select Case strThisJType
              Case "Deposit", "Purchase", "Withdrawn", "Sold", "Liability (+)", "Liability (-)"
9080            If gblnClosing = False And gblnDeleting = False Then
9090              Beep
9100              MsgBox "A Share/Face value is required for a " & strThisJType & " transaction.", _
                    vbInformation + vbOKOnly, "Entry Required"
9110              Cancel = -1
9120            End If
9130          Case Else
                ' ** Nothing else here.
9140          End Select
9150        Else
9160          Select Case strThisJType
              Case "Misc."
9170            strRecurName = Nz(.Recur_Name, vbNullString)
9180            Select Case strRecurName
                Case RECUR_I_TO_P
9190              If .PCash = 0@ And .ICash = 0@ Then
9200                If gblnClosing = False And gblnDeleting = False Then
9210                  Beep
9220                  MsgBox "Income Cash and Principal Cash must both" & vbCrLf & _
                        "have values for the chosen Recurring Item.", vbInformation + vbOKOnly, "Entry Required"
9230                  Cancel = -1
9240                End If
9250              ElseIf Abs(.PCash) = Abs(.ICash) Then
9260                If .ICash > 0@ Then
9270                  .ICash = -(.ICash)
9280                  .PCash = Abs(.ICash)
9290                ElseIf .PCash < 0@ Then
9300                  .PCash = Abs(.PCash)
9310                  .ICash = -(.PCash)
9320                End If
9330              ElseIf Abs(.PCash) <> Abs(.ICash) Then
9340                If gblnClosing = False And gblnDeleting = False Then
9350                  Beep
9360                  MsgBox "The Income Cash and Principal Cash values must" & vbCrLf & _
                        "be equal for the chosen Recurring Item.", vbInformation + vbOKOnly, "Entry Required"
9370                  Cancel = -1
9380                End If
9390              End If
9400            Case RECUR_P_TO_I
9410              If .PCash = 0@ And .ICash = 0@ Then
9420                If gblnClosing = False And gblnDeleting = False Then
9430                  Beep
9440                  MsgBox "Income Cash and Principal Cash must both" & vbCrLf & _
                        "have values for the chosen Recurring Item.", vbInformation + vbOKOnly, "Entry Required"
9450                  Cancel = -1
9460                End If
9470              ElseIf Abs(.PCash) = Abs(.ICash) Then
9480                If .PCash > 0@ Then
9490                  .PCash = -(.PCash)
9500                  .ICash = Abs(.PCash)
9510                ElseIf .ICash < 0@ Then
9520                  .ICash = Abs(.ICash)
9530                  .PCash = -(.ICash)
9540                End If
9550              ElseIf Abs(.PCash) <> Abs(.ICash) Then
9560                If gblnClosing = False And gblnDeleting = False Then
9570                  Beep
9580                  MsgBox "The Income Cash and Principal Cash values must" & vbCrLf & _
                        "be equal for the chosen Recurring Item.", vbInformation + vbOKOnly, "Entry Required"
9590                  Cancel = -1
9600                End If
9610              End If
9620            End Select
9630          Case Else
                ' ** Not sure what else.
9640          End Select
9650        End If
9660        blnToTaxLot = False
9670      End If

9680    End With

EXITP:
9690    Exit Sub

ERRH:
9700    THAT_PROC = THIS_PROC
9710    That_Erl = Erl: That_Desc = ERR.description
9720    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
9730    Resume EXITP

End Sub

Public Sub JCol_Cost_AfterUpdate(strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_Cost_AfterUpdate(
' **   strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_Cost_AfterUpdate"

        Dim strThisJType As String

9810    With frm

9820      If .posted = True Then
9830        .posted = False
9840        .posted.Locked = False
9850      End If

9860      If IsNull(.Cost) = True Then
9870        .Cost = 0@
9880      End If

9890      strThisJType = Nz(.journaltype, vbNullString)
9900      Select Case strThisJType
          Case "Liability (+)"
9910        If .Cost > 0@ Then
9920          .Cost = -(.Cost)
9930        End If
9940      Case "Deposit", "Purchase", "Liability (-)"
9950        If .Cost < 0@ Then
9960          .Cost = Abs(.Cost)
9970        End If
9980      End Select

9990      strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
          '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
10000     blnNoMove = True
10010     .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.

10020   End With

EXITP:
10030   Exit Sub

ERRH:
10040   THAT_PROC = THIS_PROC
10050   That_Erl = Erl: That_Desc = ERR.description
10060   frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
10070   Resume EXITP

End Sub

Public Sub JCol_Cost_Exit(Cancel As Integer, blnContinue As Boolean, blnWarnZeroCost As Boolean, strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_Cost_Exit(
' **   Cancel As Integer, blnContinue As Boolean, blnWarnZeroCost As Boolean, strSaveMoveCtl As String,
' **   blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_Cost_Exit"

        Dim strThisJType As String
        Dim msgResponse As VbMsgBoxResult

10110   With frm

10120     If .Cost.Locked = False Then

10130       If IsNull(.Cost) = True Then
10140         .Cost = 0@
10150         blnContinue = True
10160       Else
10170         If .Cost.text = vbNullString Then
10180           .Cost = 0@
10190         End If
10200       End If

10210       strThisJType = Nz(.journaltype, vbNullString)
10220       Select Case strThisJType
            Case "Deposit"
10230         If .Cost = 0@ Then
10240           Beep
10250           DoCmd.Hourglass False
10260           msgResponse = MsgBox("Are you sure you want this Deposit to have ZERO Cost?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cost Basis")
10270           If msgResponse <> vbYes Then
10280             blnContinue = False
10290             Cancel = -1
10300           Else
10310             blnWarnZeroCost = True
10320           End If
10330         End If
10340       Case "Purchase"
10350         If .ICash <> 0@ And .PCash = 0@ And Abs(.Cost) <> Abs(.ICash) Then
10360           blnContinue = False
10370           If gblnClosing = False And gblnDeleting = False Then
10380             Beep
10390             DoCmd.Hourglass False
10400             MsgBox "When purchasing an asset with Income Cash, Cost and Income Cash must be the same amount.", _
                    vbInformation + vbOKOnly, "Invalid Entry"
10410             Cancel = -1
10420           End If
10430         ElseIf .ICash <> 0@ And .PCash <> 0@ And Abs(.Cost) <> Abs(.PCash) Then
10440           blnContinue = False
10450           If gblnClosing = False And gblnDeleting = False Then
10460             Beep
10470             DoCmd.Hourglass False
10480             MsgBox "When purchasing an asset with Principal Cash, Cost and Principal Cash must be the same amount.", _
                    vbInformation + vbOKOnly, "Invalid Entry"
10490             Cancel = -1
10500           End If
10510         ElseIf Abs(.PCash) <> Abs(.Cost) Then  ' ** Can this even get hit? YES!
10520           Beep
10530           DoCmd.Hourglass False
10540           msgResponse = MsgBox("Are you sure you want Principal Cash to be different from Cost?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Principal And Cost Unequal")
10550           If msgResponse <> vbYes Then
10560             blnNoMove = True
10570             blnContinue = False
                  'Cancel = -1  ' ** Cancel keeps it here in Cost.
10580             .PCash.SetFocus
10590           End If
10600         End If
10610       Case "Withdrawn"
10620         If .Cost = 0@ Then
10630           DoCmd.Hourglass False
10640           If .Parent.ToTaxLot = 0 Then
10650             Beep
10660             msgResponse = MsgBox("Are you sure you want this Withdrawn to have ZERO Cost?", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "Zero Cost Basis")
10670             If msgResponse <> vbYes Then
10680               blnContinue = False
10690               Cancel = -1
10700             Else
10710               blnWarnZeroCost = True
10720             End If
10730           Else
                  ' ** Just let it leave.
10740           End If
10750         End If
10760       Case "Liability (+)"
10770         If .Cost = 0@ Then
10780           blnContinue = False
10790           If gblnClosing = False And gblnDeleting = False Then
10800             Beep
10810             DoCmd.Hourglass False
10820             MsgBox "A negative Cost value is required for this type of Liability transaction.", _
                    vbInformation + vbOKOnly, "Entry Required"
10830             Cancel = -1
10840           End If
10850         End If
10860       Case "Cost Adj."
10870         If .Cost = 0& Then
10880           blnContinue = False
10890           If gblnClosing = False And gblnDeleting = False Then
10900             Beep
10910             DoCmd.Hourglass False
10920             MsgBox "A Cost value is required for a Cost Adj. transaction.", vbInformation + vbOKOnly, "Entry Required"
10930             Cancel = -1
10940           End If
10950         End If
10960       Case "Dividend", "Interest", "Misc.", "Paid", "Received"
10970         If .Cost <> 0@ Then
10980           .Cost = 0@
10990           blnContinue = True
11000         End If
11010       Case Else
              ' ** Nothing at the moment.
11020       End Select

11030       If blnContinue = True Then
11040         strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
              '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
11050         .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
11060       End If

11070     End If

11080   End With

EXITP:
11090   Exit Sub

ERRH:
11100   THAT_PROC = THIS_PROC
11110   That_Erl = Erl: That_Desc = ERR.description
11120   frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
11130   Resume EXITP

End Sub

Public Sub JCol_Desc_AfterUpdate(blnIsPosted As Boolean, strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_Desc_AfterUpdate(
' **   blnIsPosted As Boolean, strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean,
' **   blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_Desc_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strNext As String
        Dim lngJrnlID As Long
        Dim strTmp01 As String

11210   With frm

11220     lngJrnlID = 0&: strTmp01 = vbNullString

11230     If .posted = True Then
11240       blnIsPosted = True
11250       lngJrnlID = .Journal_ID
11260     End If

11270     If InStr(.description, Chr(34)) > 0 Then
11280       MsgBox "The Description cannot contain standard quote marks.", vbInformation + vbOKOnly, "Invalid Characters"
11290       .description.Undo
11300       blnNoMove = True
11310       DoCmd.CancelEvent
11320     Else

11330       If blnIsPosted = True Then
11340         If IsNull(.description) = False Then
11350           If Trim(.description) <> vbNullString Then
11360             strTmp01 = Trim(.description)
11370           End If
11380         End If
11390         Set dbs = CurrentDb
11400         With dbs
11410           If strTmp01 = vbNullString Then
                  ' ** Update Journal, for description = Null, by specified [jid].
11420             Set qdf = .QueryDefs("qryJournal_Columns_54a")
11430             With qdf.Parameters
11440               ![jid] = lngJrnlID
11450             End With
11460           Else
                  ' ** Update Journal, by specified [jid], [jdesc].
11470             Set qdf = .QueryDefs("qryJournal_Columns_54b")
11480             With qdf.Parameters
11490               ![jid] = lngJrnlID
11500               ![jdesc] = strTmp01
11510             End With
11520           End If
11530           qdf.Execute
11540           .Close
11550         End With
11560       End If

11570       strNext = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero, True, "Last")  ' ** Module Function: modJrnlCol_Keys.
            '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
11580       Select Case strNext
            Case "description"
11590         Select Case .posted
              Case True
                ' ** Proceed normally.
11600           strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
                '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
11610           blnNoMove = True
11620           .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
11630         Case False
11640           blnNoMove = True
11650           strSaveMoveCtl = vbNullString
11660           blnNoMove = True
11670           .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
11680           CommitRec frm, blnNextRec, blnFromZero  ' ** Module Function: modJrnlCol_Recs.
11690         End Select
11700       Case Else
              ' ** Proceed normally.
11710         strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
              '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
11720         blnNoMove = True
11730         .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
11740       End Select

11750     End If

11760   End With

EXITP:
11770   Set qdf = Nothing
11780   Set dbs = Nothing
11790   Exit Sub

ERRH:
11800   THAT_PROC = THIS_PROC
11810   That_Erl = Erl: That_Desc = ERR.description
11820   frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
11830   Resume EXITP

End Sub

Public Sub JCol_RevID_AfterUpdate(blnPosted As Boolean, strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' ** Columns:
' **   0  :  revcode_ID
' **   1  :  revcode_DESCx
' **   2  :  revcode_TYPE
' **   3  :  IE  (I/E = Income/Expense)
' **
' ** JCol_RevID_AfterUpdate(
' **   blnPosted As Boolean, strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean,
' **   blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_RevID_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strRevCode As String, lngTaxcode As Long, lngTaxType As Long
        Dim strNext As String, strThisJType As String

11910   With frm

11920     lngTaxcode = Nz(.taxcode, 0&)
11930     lngTaxType = Nz(.taxcode_type, 0&)

11940     If IsNull(.revcode_ID) = False Then
11950       If .revcode_ID > 0 Then
11960         .revcode_DESC = .revcode_ID.Column(1)
11970         If .revcode_ID.Column(1) = "Unspecified Income" Or .revcode_ID.Column(1) = "Unspecified Expense" Then
11980           .revcode_DESC_display = vbNullString
11990         Else
12000           .revcode_DESC_display = .revcode_ID.Column(1)
12010         End If
12020         If gblnLinkRevTaxCodes = True Then
12030           If lngTaxcode > 0& Then
12040             strRevCode = .revcode_ID.Column(3)
12050             If strRevCode = "I" And lngTaxType = TAXTYP_DED Then
12060               .taxcode = TAXID_INC  ' ** Unspecified Income.
12070               .taxcode_description = "Unspecified Income"
12080               .taxcode_description_display = vbNullString
12090             ElseIf strRevCode = "E" And lngTaxType = TAXTYP_INC Then
12100               .taxcode = TAXID_DED  ' ** Unspecified Deduction.
12110               .taxcode_description = "Unspecified Deduction"
12120               .taxcode_description_display = vbNullString
12130             End If
12140           End If
12150         End If
12160       Else
12170         Select Case .revcode_ID.RowSource
              Case "qryJournal_Columns_10_M_RevCode_02"
                ' ** INCOME.
12180           Select Case .journaltype
                Case "Dividend"
12190             .revcode_ID = REVID_ORDDIV  ' ** Ordinary Dividend.
12200             .revcode_DESC = "Ordinary Dividend"
12210             .revcode_DESC_display = "Ordinary Dividend"
12220           Case "Interest"
12230             .revcode_ID = REVID_INTINC  ' ** Interest Income.
12240             .revcode_DESC = "Interest Income"
12250             .revcode_DESC_display = "Interest Income"
12260           Case Else
12270             .revcode_ID = REVID_INC  ' ** Unspecified Income.
12280             .revcode_DESC = "Unspecified Income"
12290             .revcode_DESC_display = vbNullString
12300           End Select
12310         Case "qryJournal_Columns_10_M_RevCode_03"
                ' ** EXPENSE.
12320           .revcode_ID = REVID_EXP  ' ** Unspecified Expense.
12330           .revcode_DESC = "Unspecified Expense"
12340           .revcode_DESC_display = vbNullString
12350         Case "qryJournal_Columns_10_M_RevCode_01"
                ' ** ALL.
12360           If gblnLinkRevTaxCodes = True Then
12370             If lngTaxcode > 0 Then
12380               Select Case lngTaxType
                    Case TAXTYP_INC
                      ' ** INCOME.
12390                 .revcode_ID = REVID_INC
12400                 .revcode_DESC = "Unspecified Income"
12410                 .revcode_DESC_display = vbNullString
12420               Case TAXTYP_DED
                      ' ** EXPENSE.
12430                 .revcode_ID = REVID_EXP
12440                 .revcode_DESC = "Unspecified Expense"
12450                 .revcode_DESC_display = vbNullString
12460               End Select
12470             Else
                    ' ** INCOME.
12480               .revcode_ID = REVID_INC
12490               .revcode_DESC = "Unspecified Income"
12500               .revcode_DESC_display = vbNullString
12510             End If
12520           Else
                  ' ** INCOME.
12530             .revcode_ID = REVID_INC
12540             .revcode_DESC = "Unspecified Income"
12550             .revcode_DESC_display = vbNullString
12560           End If
12570         End Select
12580       End If
12590     Else
12600       Select Case .revcode_ID.RowSource
            Case "qryJournal_Columns_10_M_RevCode_02"
              ' ** INCOME.
12610         Select Case .journaltype
              Case "Dividend"
12620           .revcode_ID = REVID_ORDDIV  ' ** Ordinary Dividend.
12630           .revcode_DESC = "Ordinary Dividend"
12640           .revcode_DESC_display = "Ordinary Dividend"
12650         Case "Interest"
12660           .revcode_ID = REVID_INTINC  ' ** Interest Income.
12670           .revcode_DESC = "Interest Income"
12680           .revcode_DESC_display = "Interest Income"
12690         Case Else
12700           .revcode_ID = REVID_INC
12710           .revcode_DESC = "Unspecified Income"
12720           .revcode_DESC_display = vbNullString
12730         End Select
12740       Case "qryJournal_Columns_10_M_RevCode_03"
              ' ** EXPENSE.
12750         .revcode_ID = REVID_EXP
12760         .revcode_DESC = "Unspecified Expense"
12770         .revcode_DESC_display = vbNullString
12780       Case "qryJournal_Columns_10_M_RevCode_01"
              ' ** ALL.
12790         If gblnLinkRevTaxCodes = True Then
12800           If lngTaxcode > 0 Then
12810             Select Case lngTaxType
                  Case TAXTYP_INC
                    ' ** INCOME.
12820               .revcode_ID = REVID_INC
12830               .revcode_DESC = "Unspecified Income"
12840               .revcode_DESC_display = vbNullString
12850             Case TAXTYP_DED
                    ' ** EXPENSE.
12860               .revcode_ID = REVID_EXP
12870               .revcode_DESC = "Unspecified Expense"
12880               .revcode_DESC_display = vbNullString
12890             End Select
12900           Else
                  ' ** INCOME.
12910             .revcode_ID = REVID_INC
12920             .revcode_DESC = "Unspecified Income"
12930             .revcode_DESC_display = vbNullString
12940           End If
12950         Else
                ' ** INCOME.
12960           .revcode_ID = REVID_INC
12970           .revcode_DESC = "Unspecified Income"
12980           .revcode_DESC_display = vbNullString
12990         End If
13000       End Select
13010     End If

13020     If blnPosted = True Then
13030       Set dbs = CurrentDb
13040       With dbs
              ' ** Update Journal, for revcode_ID, by specified [JrnlID], [revID].
13050         Set qdf = .QueryDefs("qryJournal_Columns_11a")
13060         With qdf.Parameters
13070           ![jrnlid] = frm.Journal_ID
13080           ![revID] = frm.revcode_ID
13090         End With
13100         qdf.Execute
13110         .Close
13120       End With
13130     End If

13140     strNext = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero, True, "Last")  ' ** Module Function: modJrnlCol_Keys.
          '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
13150     Select Case strNext
          Case "revcode_id", "revcode_DESC_display"
13160       Select Case .posted
            Case True
              ' ** Proceed normally.
13170         strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
              '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
13180         blnNoMove = True
13190         .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
13200       Case False
13210         blnNoMove = True
13220         strSaveMoveCtl = vbNullString
13230         blnNoMove = True
13240         .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
13250         strThisJType = Nz(.journaltype, vbNullString)
13260         Select Case strThisJType
              Case "Cost Adj."
13270           If .Cost <> 0@ And .assetno > 0& Then
13280             JC_Rec_CostAdjRec frm  ' ** Module Procedure: modJrnlCol_Recs.
13290           End If
13300         Case Else
13310           CommitRec frm, blnNextRec, blnFromZero  ' ** Module Function: modJrnlCol_Recs.
13320         End Select
13330       End Select
13340     Case Else
            ' ** Proceed normally.
13350       strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
            '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
13360       blnNoMove = True
13370       .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
13380     End Select

13390   End With

EXITP:
13400   Set qdf = Nothing
13410   Set dbs = Nothing
13420   Exit Sub

ERRH:
13430   THAT_PROC = THIS_PROC
13440   That_Erl = Erl: That_Desc = ERR.description
13450   frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
13460   Resume EXITP

End Sub

Public Sub JCol_TaxCode_AfterUpdate(blnPosted As Boolean, blnDontCommitTwice As Boolean, strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' ** RowSource is 0-Based:
' **   Col 0: taxcode
' **   Col 1: taxcode_description
' **   Col 2: taxcode_type
' **   Col 3: taxcode_type_code (I/D)
' **   Col 4: revcode_TYPE
' **   Col 5: revcode_TYPE_Code (I/E)
' ** BoundColumn is 1-Based:
' **   Col 0: ListIndex
' **
' ** JCol_TaxCode_AfterUpdate(
' **   blnPosted As Boolean, blnDontCommitTwice As Boolean, strSaveMoveCtl As String,
' **   blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_TaxCode_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strNext As String, strThisJType As String

13510   With frm

13520     If IsNull(.taxcode) = False Then
13530       If IsNull(.taxcode.Column(1)) = False Then
13540         .taxcode_description = .taxcode.Column(1)
13550         If .taxcode.Column(1) = "Unknown" Or _
                  .taxcode.Column(1) = "Unspecified Income" Or .taxcode.Column(1) = "Unspecified Deduction" Then
13560           .taxcode_description_display = vbNullString
13570         Else
13580           .taxcode_description_display = .taxcode.Column(1)
13590         End If
13600       Else
13610         Select Case .revcode_TYPE
              Case REVTYP_INC
13620           .taxcode = TAXID_INC
13630           .taxcode_description = "Unspecified Income"
13640           .taxcode_description_display = vbNullString
13650           .taxcode_type = TAXTYP_INC
13660         Case REVTYP_EXP
13670           .taxcode = TAXID_DED
13680           .taxcode_description = "Unspecified Deduction"
13690           .taxcode_description_display = vbNullString
13700           .taxcode_type = TAXTYP_DED
13710         End Select
13720       End If
13730     Else
13740       Select Case .revcode_TYPE
            Case REVTYP_INC
13750         .taxcode = TAXID_INC
13760         .taxcode_description = "Unspecified Income"
13770         .taxcode_description_display = vbNullString
13780         .taxcode_type = TAXTYP_INC
13790       Case REVTYP_EXP
13800         .taxcode = TAXID_DED
13810         .taxcode_description = "Unspecified Deduction"
13820         .taxcode_description_display = vbNullString
13830         .taxcode_type = TAXTYP_DED
13840       End Select
13850     End If

13860     If gblnLinkRevTaxCodes = True Then
13870       If .taxcode.Column(2) = TAXTYP_INC Then
              ' ** INCOME.
13880         If .revcode_ID.Column(2) = REVTYP_EXP Then
13890           Select Case .journaltype
                Case "Dividend"
13900             .revcode_ID = REVID_ORDDIV  ' ** Ordinary Dividend.
13910             .revcode_DESC = "Ordinary Dividend"
13920             .revcode_DESC_display = "Ordinary Dividend"
13930           Case "Interest"
13940             .revcode_ID = REVID_INTINC  ' ** Interest Income.
13950             .revcode_DESC = "Interest Income"
13960             .revcode_DESC_display = "Interest Income"
13970           Case Else
13980             .revcode_ID = REVID_INC  ' ** Unspecified Income.
13990             .revcode_DESC = "Unspecified Income"
14000             .revcode_DESC_display = vbNullString
14010           End Select
14020         End If
14030       Else
14040         If .revcode_ID.Column(2) = REVTYP_INC Then
                ' ** EXPENSE.
14050           .revcode_ID = REVID_EXP  ' ** Unspecified Expense.
14060           .revcode_DESC = "Unspecified Expense"
14070           .revcode_DESC_display = vbNullString
14080         End If
14090       End If
14100     End If

14110     If blnPosted = True Then
14120       Set dbs = CurrentDb
14130       With dbs
              ' ** Update Journal, for taxcode, by specified [JrnlID], [taxID].
14140         Set qdf = .QueryDefs("qryJournal_Columns_11b")
14150         With qdf.Parameters
14160           ![jrnlid] = frm.Journal_ID
14170           ![taxID] = frm.taxcode
14180         End With
14190         qdf.Execute
14200         .Close
14210       End With
14220     End If

14230     If blnDontCommitTwice = False Then
14240       strNext = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero, True, "Last")  ' ** Module Function: modJrnlCol_Keys.
            '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
14250       Select Case strNext
            Case "taxcode", "taxcode_description_display"
14260         Select Case .posted
              Case True
                ' ** Proceed normally.
14270           strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
                '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
14280           .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
14290         Case False
14300           blnNoMove = True
14310           strSaveMoveCtl = vbNullString
14320           .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
14330           strThisJType = Nz(.journaltype, vbNullString)
14340           Select Case strThisJType
                Case "Cost Adj."
14350             If .Cost <> 0@ And .assetno > 0& Then
14360               JC_Rec_CostAdjRec frm  ' ** Module Procedure: modJrnlCol_Recs.
14370             End If
14380           Case Else
14390             CommitRec frm, blnNextRec, blnFromZero ' ** Module Function: modJrnlCol_Recs.
14400           End Select
14410         End Select
14420       Case Else
              ' ** Proceed normally.
14430         strSaveMoveCtl = JC_Key_Sub_Next(THAT_PROC, blnNextRec, blnFromZero)  ' ** Module Function: modJrnlCol_Keys.
              '03/24/2017: CHANGED THIS_PROC TO THAT_PROC!
14440         .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
14450       End Select
14460     End If  ' ** blnDontCommitTwice.

14470   End With

EXITP:
14480   Set qdf = Nothing
14490   Set dbs = Nothing
14500   Exit Sub

ERRH:
14510   THAT_PROC = THIS_PROC
14520   That_Erl = Erl: That_Desc = ERR.description
14530   frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
14540   Resume EXITP

End Sub

Public Sub JCol_Reinvest_AfterUpdate(blnReinvestment As Boolean, strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** JCol_Reinvest_AfterUpdate(
' **   blnReinvestment As Boolean, strSaveMoveCtl As String, blnNextRec As Boolean, blnFromZero As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_Reinvest_AfterUpdate"

        Dim strThisJType As String
        Dim strAccountNo As String, lngAssetNo As Long, strAssetNo_Desc As String
        Dim datTransDate As Date, datAssetDate As Date
        Dim curICash As Currency
        Dim blnRetVal As Boolean, lngRetVal As Long

14610   With frm

14620     strThisJType = Nz(.journaltype, vbNullString)

14630     Select Case strThisJType
          Case "Dividend", "Interest"
14640       If IsNull(.transdate) = False And IsNull(.accountno) = False And IsNull(.assetno) = False Then
14650         If .assetno > 0& And IsNull(.shareface) = False And IsNull(.ICash) = False Then
14660           If .shareface > 0# And .ICash > 0@ Then

14670             Select Case gblnLinkRevTaxCodes
                  Case True
                    ' ** Make sure they match.
14680               Select Case .revcode_TYPE
                    Case REVTYP_INC
14690                 If IsNull(.taxcode) = True Then
14700                   .taxcode = TAXID_INC
14710                   .taxcode_description = "Unspecified Income"
14720                   .taxcode_description_display = Null
14730                   .taxcode_type = TAXTYP_INC
14740                 Else
14750                   If .taxcode = 0& Then
14760                     .taxcode = TAXID_INC
14770                     .taxcode_description = "Unspecified Income"
14780                     .taxcode_description_display = Null
14790                     .taxcode_type = TAXTYP_INC
14800                   ElseIf .taxcode_type <> TAXTYP_INC Then
14810                     .taxcode = TAXID_INC
14820                     .taxcode_description = "Unspecified Income"
14830                     .taxcode_description_display = Null
14840                     .taxcode_type = TAXTYP_INC
14850                   End If
14860                 End If
14870               Case REVTYP_EXP
14880                 If IsNull(.taxcode) = True Then
14890                   .taxcode = TAXID_DED
14900                   .taxcode_description = "Unspecified Deduction"
14910                   .taxcode_description_display = Null
14920                   .taxcode_type = TAXTYP_DED
14930                 Else
14940                   If .taxcode = 0& Then
14950                     .taxcode = TAXID_DED
14960                     .taxcode_description = "Unspecified Deduction"
14970                     .taxcode_description_display = Null
14980                     .taxcode_type = TAXTYP_DED
14990                   ElseIf .taxcode_type <> TAXTYP_DED Then
15000                     .taxcode = TAXID_DED
15010                     .taxcode_description = "Unspecified Deduction"
15020                     .taxcode_description_display = Null
15030                     .taxcode_type = TAXTYP_DED
15040                   End If
15050                 End If
15060               End Select
15070             Case False
                    ' ** Don't care.
15080             End Select

15090             strAccountNo = .accountno
15100             lngAssetNo = .assetno
15110             strAssetNo_Desc = .assetno_description
15120             datTransDate = .transdate
15130             datAssetDate = .assetdate
15140             curICash = .ICash

15150             blnRetVal = CommitRec(frm, blnNextRec, blnFromZero, True)  ' ** Module Function: modJrnlCol_Recs.
15160             If blnRetVal = True Then  ' ** Meaning it DID pass the tests.
                    ' ** From here on down, we're on the new Purchase record.

15170               blnReinvestment = True
15180               .journaltype = "Purchase"
15190               .journalSubtype = "Reinvest"

15200               .journaltype_AfterUpdate  ' ** Form Procedure: frmJournal_Columns_Sub.
15210               DoEvents

15220               JC_Msc_Asset_Set frm  ' ** Module Procedure: modJrnlCol_Misc.

15230               .transdate = datTransDate
15240               .accountno = strAccountNo
15250               .assetno = lngAssetNo
15260               .assetno_description = strAssetNo_Desc
15270               .assetdate = datAssetDate
15280               .assetdate_display = CDate(Format(datAssetDate, "mm/dd/yyyy"))
15290               .shareface = 0#
15300               .ICash = -(curICash)
15310               .PCash = 0@
15320               .Cost = curICash

15330               Select Case gblnLinkRevTaxCodes
                    Case True
                      ' ** Make sure they match.
15340                 .revcode_ID = REVID_INC  ' ** Unspecified Income.
15350                 .revcode_DESC = "Unspecified Income"
15360                 .revcode_DESC_display = Null
15370                 .revcode_TYPE = REVTYP_INC
15380                 .taxcode = TAXID_INC
15390                 .taxcode_description = "Unspecified Income"
15400                 .taxcode_description_display = Null
15410                 .taxcode_type = TAXTYP_INC
15420               Case False
                      ' ** Don't care.
15430               End Select

15440               .assetno.SetFocus
15450               lngRetVal = fSetScrollBarPosHZ(frm, 1&)  ' ** Module Function: modScrollBarFuncs.

15460               strSaveMoveCtl = vbNullString
15470               .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.

15480               If .assetno.Locked = True Then
15490                 JC_Key_JType_Set "Purchase", frm  ' ** Module Function: modJrnlCol_Keys.
15500               End If

15510             End If
15520           Else
                  ' ** Unlikely to get this far.
15530           End If  ' ** shareface, icash.
15540         Else
                ' ** Unlikely to get this far.
15550         End If  ' ** Values.
15560       Else
              ' ** Unlikely to get this far.
15570       End If  ' ** Null.

15580     Case Else
            ' ** Shouldn't be here.
15590       If .Reinvested = True Then .Reinvested = False
15600     End Select

15610   End With

EXITP:
15620   Exit Sub

ERRH:
15630   THAT_PROC = THIS_PROC
15640   That_Erl = Erl: That_Desc = ERR.description
15650   frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
15660   Resume EXITP

End Sub

Public Sub JCol_Shareface_Change(THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_Shareface_Change"

        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String

15710   With frm
15720     If IsNull(.shareface.text) = False Then
15730       If IsNumeric(.shareface.text) = True Then
15740         If CDbl(.shareface.text) > 999999999.9999 Then
15750           MsgBox "Share/Face is too large.", vbInformation + vbOKOnly, "Invalid Entry"
15760           .shareface.Undo
15770         ElseIf CDbl(.shareface.text) < 0 Then
15780           MsgBox "You cannot enter a negative Share/Face.", vbInformation + vbOKOnly, "Invalid Entry"
15790           .shareface.Undo
15800         Else
15810           strTmp01 = CDbl(.shareface.text)
15820           intPos01 = InStr(strTmp01, ".")
15830           If intPos01 > 0 Then
15840             strTmp02 = Mid(strTmp01, (intPos01 + 1))
15850             If Len(strTmp02) > gintShareFaceDecimals Then
15860               MsgBox "You can only enter Share/Face with up to " & Trim(str(gintShareFaceDecimals)) & " decimals.", _
                      vbInformation + vbOKOnly, "Invalid Entry"
15870               strTmp01 = Left(strTmp01, intPos01) & Left(strTmp02, gintShareFaceDecimals)
15880               .shareface.text = CDbl(strTmp01)
15890             Else
                    ' ** Let it proceed.
15900             End If
15910           Else
                  ' ** Let it proceed.
15920           End If
15930         End If
15940       Else
              'MsgBox "Share/Face must have a value.", vbInformation + vbOKOnly, "Invalid Entry"
15950         .shareface.Undo
15960       End If
15970     Else
            ' ** Let the other events handle it.
15980     End If
15990   End With

EXITP:
16000   Exit Sub

ERRH:
16010   THAT_PROC = THIS_PROC
16020   That_Erl = Erl: That_Desc = ERR.description
16030   frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
16040   Resume EXITP

End Sub

Public Sub JCol_ICash_Change(THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_ICash_Change"

        Dim blnContinue As Boolean, blnPlusMinus As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String

16110   blnContinue = True: blnPlusMinus = False
16120   With frm
16130     If IsNull(.ICash.text) = False Then
16140       If IsNumeric(.ICash.text) = False Then
16150         If Trim(.ICash.text) = vbNullString Then
16160           blnContinue = False
16170         Else
16180           If Left(.ICash.text, 1) = "-" Or Left(.ICash.text, 1) = "+" Then
                  ' ** These are OK even though they fail IsNumeric()!
16190             If .ICash.text = "-" Or .ICash.text = "+" Then
                    ' ** If they're ONLY +/-, they blow up CDbl()!
16200               blnPlusMinus = True
16210             End If
16220           Else
16230             blnContinue = False
16240             .ICash.Undo
16250           End If
16260         End If
16270       End If
16280       If blnPlusMinus = False Then
16290         If blnContinue = True Then
16300           If CDbl(.ICash.text) > 999999999.9999 Then
16310             blnContinue = False
16320             MsgBox "Income Cash is too large.", vbInformation + vbOKOnly, "Invalid Entry"
16330             .ICash.Undo
16340           End If
16350         End If
16360         If blnContinue = True Then
16370           strTmp01 = CDbl(.ICash.text)
16380           intPos01 = InStr(strTmp01, ".")
16390           If intPos01 > 0 Then
16400             strTmp02 = Mid(strTmp01, (intPos01 + 1))
16410             If Len(strTmp02) > 2 Then
16420               MsgBox "You can only enter Income Cash with up to 2 decimals.", _
                      vbInformation + vbOKOnly, "Invalid Entry"
16430               strTmp01 = Left(strTmp01, intPos01) & Left(strTmp02, 2)
16440               .ICash.text = CDbl(strTmp01)
16450             Else
                    ' ** Let it proceed.
16460             End If
16470           Else
                  ' ** Let it proceed.
16480           End If
16490         End If
16500       End If
16510     Else
            ' ** Let the other events handle it.
16520     End If
16530   End With

EXITP:
16540   Exit Sub

ERRH:
16550   THAT_PROC = THIS_PROC
16560   That_Erl = Erl: That_Desc = ERR.description
16570   frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
16580   Resume EXITP

End Sub

Public Sub JCol_PCash_Change(THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_PCash_Change"

        Dim blnContinue As Boolean, blnPlusMinus As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String

16610   blnContinue = True: blnPlusMinus = False
16620   With frm
16630     If IsNull(.PCash.text) = False Then
16640       If IsNumeric(.PCash.text) = False Then
16650         If Trim(.PCash.text) = vbNullString Then
16660           blnContinue = False
16670         Else
16680           If Left(.PCash.text, 1) = "-" Or Left(.PCash.text, 1) = "+" Then
                  ' ** These are OK even though they fail IsNumeric()!
16690             If .PCash.text = "-" Or .PCash.text = "+" Then
                    ' ** If they're ONLY +/-, they blow up CDbl()!
16700               blnPlusMinus = True
16710             End If
16720           Else
16730             blnContinue = False
16740             .PCash.Undo
16750           End If
16760         End If
16770       End If
16780       If blnPlusMinus = False Then
16790         If blnContinue = True Then
16800           If CDbl(.PCash.text) > 999999999.9999 Then
16810             MsgBox "Principal Cash is too large.", vbInformation + vbOKOnly, "Invalid Entry"
16820             .PCash.Undo
16830           Else
16840             strTmp01 = CDbl(.PCash.text)
16850             intPos01 = InStr(strTmp01, ".")
16860             If intPos01 > 0 Then
16870               strTmp02 = Mid(strTmp01, (intPos01 + 1))
16880               If Len(strTmp02) > 2 Then
16890                 MsgBox "You can only enter Principal Cash with up to 2 decimals.", _
                        vbInformation + vbOKOnly, "Invalid Entry"
16900                 strTmp01 = Left(strTmp01, intPos01) & Left(strTmp02, 2)
16910                 .PCash.text = CDbl(strTmp01)
16920               Else
                      ' ** Let it proceed.
16930               End If
16940             Else
                    ' ** Let it proceed.
16950             End If
16960           End If
16970         End If
16980       End If
16990     Else
            ' ** Let the other events handle it.
17000     End If
17010   End With

EXITP:
17020   Exit Sub

ERRH:
17030   THAT_PROC = THIS_PROC
17040   That_Erl = Erl: That_Desc = ERR.description
17050   frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
17060   Resume EXITP

End Sub

Public Sub JCol_Cost_Change(THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_Cost_Change"

        Dim blnContinue As Boolean, blnPlusMinus As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String

17110   blnContinue = True: blnPlusMinus = False
17120   With frm
17130     If IsNull(.Cost.text) = False Then
17140       If IsNumeric(.Cost.text) = False Then
17150         If Trim(.Cost.text) = vbNullString Then
17160           blnContinue = False
17170         Else
17180           If Left(.Cost.text, 1) = "-" Or Left(.Cost.text, 1) = "+" Then
                  ' ** These are OK even though they fail IsNumeric()!
17190             If .Cost.text = "-" Or .Cost.text = "+" Then
                    ' ** If they're ONLY +/-, they blow up CDbl()!
17200               blnPlusMinus = True
17210             End If
17220           Else
17230             blnContinue = False
17240             .Cost.Undo
17250           End If
17260         End If
17270       End If
17280       If blnPlusMinus = False Then
17290         If blnContinue = True Then
17300           If CDbl(.Cost.text) > 999999999.9999 Then
17310             blnContinue = False
17320             MsgBox "Cost is too large.", vbInformation + vbOKOnly, "Invalid Entry"
17330             .Cost.Undo
17340           End If
17350         End If
17360         If blnContinue = True Then
17370           strTmp01 = CDbl(.Cost)
17380           intPos01 = InStr(strTmp01, ".")
17390           If intPos01 > 0 Then
17400             strTmp02 = Mid(strTmp01, (intPos01 + 1))
17410             If Len(strTmp02) > 2 Then
17420               MsgBox "You can only enter Cost with up to 2 decimals.", _
                      vbInformation + vbOKOnly, "Invalid Entry"
17430               strTmp01 = Left(strTmp01, intPos01) & Left(strTmp02, 2)
17440               .Cost.text = CDbl(strTmp01)
17450             Else
                    ' ** Let it proceed.
17460             End If
17470           Else
                  ' ** Let it proceed.
17480           End If
17490         End If
17500       End If
17510     Else
            ' ** Let the other events handle it.
17520     End If
17530   End With

EXITP:
17540   Exit Sub

ERRH:
17550   THAT_PROC = THIS_PROC
17560   That_Erl = Erl: That_Desc = ERR.description
17570   frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns_Sub.
17580   Resume EXITP

End Sub

Public Sub JCol_Memo_AfterUpdate(THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_Memo_AfterUpdate"

        Dim strMemo_New As String

17610   With frm
17620     If .frmJournal_Columns_Sub.Form.posted = True Then
17630       .frmJournal_Columns_Sub.Form.posted = False
17640       .frmJournal_Columns_Sub.Form.posted.Locked = False
17650     End If
17660     If IsNull(.JrnlMemo_Memo) = False Then
17670       If Trim(.JrnlMemo_Memo) <> vbNullString Then
17680         strMemo_New = Trim(.JrnlMemo_Memo)
17690         If Len(strMemo_New) > MEMO_MAX Then
17700           strMemo_New = Left(strMemo_New, MEMO_MAX)
17710           .JrnlMemo_Memo = strMemo_New
17720         End If
17730         .frmJournal_Columns_Sub.Form.JrnlMemo_Memo = strMemo_New
17740         .frmJournal_Columns_Sub.Form.JrnlMemo_HasMemo = True
17750       Else
17760         .frmJournal_Columns_Sub.Form.JrnlMemo_Memo = Null
17770         .frmJournal_Columns_Sub.Form.JrnlMemo_HasMemo = False
17780       End If
17790     Else
17800       .frmJournal_Columns_Sub.Form.JrnlMemo_Memo = Null
17810       .frmJournal_Columns_Sub.Form.JrnlMemo_HasMemo = False
17820     End If
17830   End With

EXITP:
17840   Exit Sub

ERRH:
17850   THAT_PROC = THIS_PROC
17860   That_Erl = Erl: That_Desc = ERR.description
17870   frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns.
17880   Resume EXITP

End Sub

Public Sub JCol_OpgFilter_AfterUpdate(THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "JCol_OpgFilter_AfterUpdate"

        Dim strFilter As String

17910   With frm
17920     gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
17930     Select Case .opgFilter
          Case .opgFilter_optAll.OptionValue
17940       .opgFilter_optAll_lbl.FontBold = True
17950       .opgFilter_optCommitted_lbl.FontBold = False
17960       .opgFilter_optUncommitted_lbl.FontBold = False
17970       Select Case gblnAdmin
            Case True
17980         .frmJournal_Columns_Sub.Form.Filter = vbNullString
17990         .frmJournal_Columns_Sub.Form.FilterOn = False
18000       Case False
18010         strFilter = "[journal_USER] = '" & gstrJournalUser & "'"
18020         .frmJournal_Columns_Sub.Form.Filter = strFilter
18030         .frmJournal_Columns_Sub.Form.FilterOn = True
18040       End Select
18050     Case .opgFilter_optCommitted.OptionValue
18060       .opgFilter_optAll_lbl.FontBold = False
18070       .opgFilter_optCommitted_lbl.FontBold = True
18080       .opgFilter_optUncommitted_lbl.FontBold = False
18090       Select Case gblnAdmin
            Case True
18100         strFilter = "[posted] = True"
18110       Case False
18120         strFilter = "[posted] = True And [journal_USER] = '" & gstrJournalUser & "'"
18130       End Select
18140       .frmJournal_Columns_Sub.Form.Filter = strFilter
18150       .frmJournal_Columns_Sub.Form.FilterOn = True
18160     Case .opgFilter_optUncommitted.OptionValue
18170       .opgFilter_optAll_lbl.FontBold = False
18180       .opgFilter_optCommitted_lbl.FontBold = False
18190       .opgFilter_optUncommitted_lbl.FontBold = True
18200       Select Case gblnAdmin
            Case True
18210         strFilter = "[posted] = False"
18220       Case False
18230         strFilter = "[posted] = False And [journal_USER] = '" & gstrJournalUser & "'"
18240       End Select
18250       .frmJournal_Columns_Sub.Form.Filter = strFilter
18260       .frmJournal_Columns_Sub.Form.FilterOn = True
18270     End Select
18280   End With

EXITP:
18290   Exit Sub

ERRH:
18300   THAT_PROC = THIS_PROC
18310   That_Erl = Erl: That_Desc = ERR.description
18320   frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Form Procedure: frmJournal_Columns.
18330   Resume EXITP

End Sub
