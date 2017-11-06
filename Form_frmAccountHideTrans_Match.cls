VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountHideTrans_Match"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountHideTrans_Match"

'VGC 03/22/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Select Account:   A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Name:             M {opgAccountSource_optName}
' **   Update:           U {cmdUpdate}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ********************************************
' ** NOTE: RecordLocks is set to 'No Locks'.
' ********************************************

'Forms(1).Header_lbl.Height = 480
'Forms(0).Header_lbl.Height = 375
'Forms(0).Header_lbl_img.Height = 495
'Forms(0).Header_lbl_img.Height = 390

' ** frmAccountHideTrans_Match highlight colors:
' **   Light Green: 15138802
' **   Light Blue:  16774128
' **   Yellow:      8454143

Private Const TIMER_DELAY As Long = 500&  ' ** Milliseconds delay to check totals.

' ** Array: arr_varSet().
Private lngSets As Long, arr_varSet() As Variant
Private Const S_ELEMS As Integer = 15  ' ** Array's first-element UBound().
Private Const S_JNO    As Integer = 0
Private Const S_JTYPE  As Integer = 1
Private Const S_ACTNO  As Integer = 2
Private Const S_ASTNO  As Integer = 3
Private Const S_SDAT   As Integer = 4
Private Const S_ICASH  As Integer = 5
Private Const S_PCASH  As Integer = 6
Private Const S_COST   As Integer = 7
Private Const S_MTCHD  As Integer = 8
Private Const S_GRP    As Integer = 9
Private Const S_UNIQ   As Integer = 10
Private Const S_SDAT2  As Integer = 11
Private Const S_GRP2   As Integer = 12
Private Const S_UNIQ2  As Integer = 13
Private Const S_HTYP2  As Integer = 15
Private Const S_REMOV  As Integer = 15

' ** Array: arr_varOrphan().
Private lngOrphans As Long, arr_varOrphan() As Variant
Private Const O_ELEMS As Integer = 6  ' ** Array's first-element UBound().
Private Const O_JNO   As Integer = 0
Private Const O_JTYPE As Integer = 1
Private Const O_GRP   As Integer = 2
Private Const O_UNIQ  As Integer = 3
Private Const O_GRP2  As Integer = 4
Private Const O_UNIQ2 As Integer = 5
Private Const O_HTYP2 As Integer = 6

Private lngX As Long, lngE As Long
'Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

110     Set dbs = CurrentDb
120     With dbs
          ' ** Update LedgerHidden, set hid_newmatch = False.
130       Set qdf = .QueryDefs("qryAccountHide_61")
140       qdf.Execute
          ' ** Update LedgerHidden, fix that double-underscore where it doesn't belong.
150       Set qdf = .QueryDefs("qryAccountHide_63b")
160       qdf.Execute
170       .Close
180     End With

190     Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

200     opgAccountSource_AfterUpdate  ' ** Procedure: Below.

210     lngSets = 0&
220     ReDim arr_varSet(S_ELEMS, 0)

230     lngOrphans = 0&
240     ReDim arr_varOrphan(O_ELEMS, 0)

250     With Me
260       .cmdUpdate.Enabled = False
270       .assetno_chk_box_lbl.Visible = False
280       .assetno_chk_box_lbl.ForeColor = CLR_BLU
290       .assetno_chk_box.BackColor = CLR_WHT
300       .misc_chk_box_lbl.Visible = False
310       .misc_chk_box_lbl.ForeColor = CLR_BLU
320       .misc_chk_box.BackColor = CLR_WHT
          '.assetno_chk_lbl2.Visible = False
330       .hid_newmatch_sum.ForeColor = CLR_BLU
340       .hid_newmatch_sum.BackColor = CLR_WHT
350       .icash_subtot.ForeColor = CLR_BLU
360       .icash_subtot.BackColor = CLR_WHT
370       .pcash_subtot.ForeColor = CLR_BLU
380       .pcash_subtot.BackColor = CLR_WHT
390       .cost_subtot.ForeColor = CLR_BLU
400       .cost_subtot.BackColor = CLR_WHT
410     End With

420     If Cancel = -1 Then
          ' ** Just for the record.
430       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Allowed"
440     End If

EXITP:
450     Set qdf = Nothing
460     Set dbs = Nothing
470     Exit Sub

ERRH:
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

610     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
620     intShiftDown = (Shift And acShiftMask) > 0
630     intAltDown = (Shift And acAltMask) > 0
640     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyDown
670         intRetVal = 0
680         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
690       Case vbKeyUp
700         intRetVal = 0
710         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
720       End Select
730     End If

        ' ** Alt keys.
740     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
750       Select Case intRetVal
          Case vbKeyX
760         intRetVal = 0
770         cmdClose_Click  ' ** Procedure: Below.
780       End Select
790     End If

        ' ** Ctrl keys.
800     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
810       Select Case intRetVal
          Case vbKeyS
820         intRetVal = 0
830         cmdSave_Click  ' ** Procedure: Below.
840       End Select
850     End If

EXITP:
860     KeyCode = intRetVal
870     Exit Sub

ERRH:
880     intRetVal = 0
890     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
900     Case Else
910       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
920     End Select
930     Resume EXITP

End Sub

Private Sub Form_Timer()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

1010    With Me

1020      .TimerInterval = 0&

          ' ** This is done here, because it takes time to total the form, and it
          ' ** hasn't finished totaling by the end of hid_newmatch_AfterUpdate().

1030      If .hid_newmatch_sum = 1 Then
1040        .hid_newmatch_sum.ForeColor = CLR_BLK
1050        .hid_newmatch_sum.BackColor = CLR_LTRED
1060      Else
1070        .hid_newmatch_sum.ForeColor = CLR_BLU
1080        .hid_newmatch_sum.BackColor = CLR_WHT
1090      End If

1100      If lngSets = 0& Then
1110        .cmdUpdate.Enabled = False
1120      Else
1130        If lngSets = 1& Then
1140  On Error Resume Next
1150          .cmdUpdate.Enabled = False
1160          If ERR.Number <> 0 Then
1170  On Error GoTo ERRH
1180            .cmdClose.SetFocus
1190            .cmdUpdate.Enabled = False
1200          Else
1210  On Error GoTo ERRH
1220          End If
1230        Else
1240          .cmdUpdate.Enabled = True
1250        End If
1260      End If

1270      If .icash_subtot <> 0 Then
1280        .icash_subtot.ForeColor = CLR_BLK
1290        .icash_subtot.BackColor = CLR_LTRED
1300      Else
1310        .icash_subtot.ForeColor = CLR_BLU
1320        .icash_subtot.BackColor = CLR_WHT
1330      End If

1340      If .pcash_subtot <> 0 Then
1350        .pcash_subtot.ForeColor = CLR_BLK
1360        .pcash_subtot.BackColor = CLR_LTRED
1370      Else
1380        .pcash_subtot.ForeColor = CLR_BLU
1390        .pcash_subtot.BackColor = CLR_WHT
1400      End If

1410      If .cost_subtot <> 0 Then
1420        .cost_subtot.ForeColor = CLR_BLK
1430        .cost_subtot.BackColor = CLR_LTRED
1440      Else
1450        .cost_subtot.ForeColor = CLR_BLU
1460        .cost_subtot.BackColor = CLR_WHT
1470      End If

1480    End With

EXITP:
1490    Exit Sub

ERRH:
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1610    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

EXITP:
1620    Exit Sub

ERRH:
1630    Select Case ERR.Number
        Case Else
1640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1650    End Select
1660    Resume EXITP

End Sub

Private Sub cmdClose_Click()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

1710    DoCmd.Close acForm, THIS_NAME

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub

Private Sub cmdUpdate_Click()
' ** OK, all we're going to update is the grouping in LedgerHidden.
' ** There's nothing else going on that would affect real data.
' ** Note: Only 1 grouping can be done at a time.
' ** Do we really care whether items truly match?
' ** Should we prevent obvious mismatches?
' ** My thoughts are that few people will ever use this,
' ** and that we don't really care what they do with it.
' ** It's just for convenience and display anyway.
' ** PREVENT MIXING ACCOUNTNO'S!
' ** YES, PREVENT MIXING ASSETNO's!

        'Careful what you do with hid_grpnum!
        'Do we want to add an index for '[accountno], [hid_grpnum]',
        'allowing the same group number to be used by multiple accounts?
        'No! That might mess up Hide_Setup() in modHideTransactions1.
        'The on-screen sort will just have to be the order in which
        'they were hidden, regardless of their hid_sortdate.

        ' ** Sort on frmAccountHideTrans_Hidden:
        ' **   [Sortx], [SortDate], [UniqueIDx], [Ord]
        ' ** Sortx = CLng(IIf([Sort]<3,1,[Sort]))

        ' ** Sort on frmAccountHideTrans_Match:
        ' **   [accountno], [hid_sort], [hid_grpnum], [hid_sortdate], [hid_order]

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngAssetCnt As Long, lngMiscCnt As Long
        Dim lngMatched As Long, lngUnmatched As Long
        Dim strNewUniqueID As String, lngNewGrpNum As Long, strNewHidType As String, datNewSortDate As Date
        Dim lngMaxGrpNum As Long
        Dim lngAvails As Long, arr_varAvail() As Variant
        Dim blnFound As Boolean
        Dim strTmp01 As String, lngTmp02 As Long
        Dim lngY As Long, lngZ As Long

        ' *****************************************************
        ' ** Array: arr_varSet()
        ' **
        ' **   Field  Element  Name                Constant
        ' **   =====  =======  ==================  ==========
        ' **     1       0     journalno           S_JNO
        ' **     2       1     journaltype         S_JTYPE
        ' **     3       2     accountno           S_ACTNO
        ' **     4       3     assetno             S_ASTNO
        ' **     5       4     hid_sortdate        S_SDAT
        ' **     6       5     icash               S_ICASH
        ' **     7       6     pcash               S_PCASH
        ' **     8       7     cost                S_COST
        ' **     9       8     IsMatched           S_MTCHD
        ' **    10       9     hid_grpnum          S_GRP
        ' **    11      10     uniqueid            S_UNIQ
        ' **    12      11     new hid_sortdate    S_SDAT2
        ' **    12      11     new hid_grpnum      S_GRP2
        ' **    13      12     new uniqueid        S_UNIQ2
        ' **    14      13     new hidtype         S_HTYP2
        ' **    15      14     Remove              S_REMOV
        ' **
        ' *****************************************************

1810    If lngSets > 1& Then

1820      lngAssetCnt = 0&: lngMiscCnt = 0&
1830      lngMatched = 0&: lngUnmatched = 0&
1840      strNewUniqueID = vbNullString: strNewHidType = vbNullString
1850      datNewSortDate = CDate("01/01/2100")
1860      lngMaxGrpNum = 0&: lngNewGrpNum = 0&

          ' ** Count asset vs. misc.
1870      For lngX = 0& To (lngSets - 1&)
1880        If arr_varSet(S_ASTNO, lngX) <> 0& Then
1890          lngAssetCnt = lngAssetCnt + 1&
1900        Else
1910          lngMiscCnt = lngMiscCnt + 1&
1920        End If
1930      Next

          ' ** Count matched vs. unmatched.
1940      For lngX = 0& To (lngSets - 1&)
1950        If arr_varSet(S_MTCHD, lngX) = True Then
1960          lngMatched = lngMatched + 1&
1970        Else
1980          lngUnmatched = lngUnmatched + 1&
1990        End If
2000      Next

          ' ** Examples:
          ' **   accountno      astno jno1   jno2   jtype1    jtype2
          ' **   =============== ==== ====== ====== ========= =========
          ' **   000000000002300_0412_020353_020350_Withdrawn_Deposit__
          ' **   000000000002300_0000_040395_000000_Misc.______________
          ' **          15         4     6      6       9         9      ' ** Might someone get more than 9,999 assets?
          ' ** Whichever one may have an assetno gets the slot.

          ' ** Accountno.
2010      strNewUniqueID = Left(arr_varSet(S_UNIQ, 0), 16)  ' ** Include trailing underscore.

          ' ** Assetno.
2020      If lngAssetCnt = 0& Then
2030        strNewUniqueID = strNewUniqueID & "0000_"  ' ** Include the trailing underscore.
2040      Else
2050        For lngX = 0& To (lngSets - 1&)
2060          If arr_varSet(S_ASTNO, lngX) <> 0& Then
2070            strNewUniqueID = strNewUniqueID & Right(String(4, "0") & CStr(arr_varSet(S_ASTNO, lngX)), 4) & "_"
                ' ** Include the trailing underscore.
2080            Exit For
2090          End If
2100        Next
2110      End If

          ' ** Journalno.
2120      For lngX = 0& To (lngSets - 1&)
2130        strNewUniqueID = strNewUniqueID & Right(String(6, "0") & CStr(arr_varSet(S_JNO, lngX)), 6) & "_"
2140      Next

          ' ** JournalType.
2150      For lngX = 0& To (lngSets - 1&)
2160        strTmp01 = arr_varSet(S_JTYPE, lngX)
2170        strTmp01 = StringReplace(strTmp01, " ", "_")  ' ** Cost Adj.
2180        strNewUniqueID = strNewUniqueID & Left(strTmp01 & String(6, "_"), 9) & "_"
2190      Next
2200      strNewUniqueID = Left(strNewUniqueID, (Len(strNewUniqueID) - 1))  ' ** Remove last trailing underscore.

          ' ** Everyone gets the same combined uniqueid.
2210      For lngX = 0& To (lngSets - 1&)
2220        arr_varSet(S_UNIQ2, lngX) = strNewUniqueID
2230      Next

          ' ** Get the highest hid_grpnum.
2240      lngMaxGrpNum = DLookup("[hid_grpnum]", "qryAccountHide_62b")

          ' ** Use the lowest hid_grpnum among the entries.
2250      lngNewGrpNum = 999999
2260      For lngX = 0& To (lngSets - 1&)
2270        If arr_varSet(S_GRP, lngX) < lngNewGrpNum Then
2280          lngNewGrpNum = arr_varSet(S_GRP, lngX)
2290        End If
2300      Next

          ' ** And give it to everyone.
2310      For lngX = 0& To (lngSets - 1&)
2320        arr_varSet(S_GRP2, lngX) = lngNewGrpNum
2330      Next

          ' ** First collect any orphan groups.
2340      lngAvails = 0&
2350      ReDim arr_varAvail(0)
2360      For lngX = 0& To (lngSets - 1&)
2370        If arr_varSet(S_GRP, lngX) <> lngNewGrpNum Then
2380          lngAvails = lngAvails + 1&
2390          lngE = lngAvails - 1&
2400          ReDim Preserve arr_varAvail(lngE)
2410          arr_varAvail(lngE) = arr_varSet(S_GRP, lngX)
2420        End If
2430      Next

          ' ** Then see if there are any other members of those orphan groups.
2440      If lngAvails > 0& Then
2450        Set rst = Me.RecordsetClone
2460        rst.MoveLast
2470        lngTmp02 = rst.RecordCount
            ' ** Check on each group that's losing a member (because it's being reassigned to a new group).
2480        For lngX = 0& To (lngAvails - 1&)
2490          With rst
2500            .MoveFirst
                ' ** Check each record in LedgerHidden (or at least those currently on the screen).
2510            For lngY = 1& To lngTmp02
2520              If ![hid_grpnum] = arr_varAvail(lngX) Then
                    ' ** This entry is in a now-orphaned group.
2530                blnFound = False
                    ' ** See if it's one of the ones in arr_varSet().
2540                For lngZ = 0& To (lngSets - 1&)
2550                  If arr_varSet(S_JNO, lngZ) = ![journalno] Then
                        ' ** Aahh. It's one of those we're reassigning, so ignore it.
2560                    blnFound = True
2570                    Exit For
2580                  End If
2590                Next
2600                If blnFound = False Then
                      ' ** This entry is definitely an orphan.
2610                  lngOrphans = lngOrphans + 1&
2620                  lngE = lngOrphans - 1&
2630                  ReDim Preserve arr_varOrphan(O_ELEMS, lngE)
                      ' ***************************************************
                      ' ** Array: arr_varOrphan()
                      ' **
                      ' **   Field  Element  Name              Constant
                      ' **   =====  =======  ================  ==========
                      ' **     1       0     journalno         O_JNO
                      ' **     2       1     journaltype       O_JTYPE
                      ' **     3       2     hid_grpnum        O_GRP
                      ' **     4       3     uniqueid          O_UNIQ
                      ' **     5       4     new hid_grpnum    O_GRP2
                      ' **     6       5     new uniqueid      O_UNIQ2
                      ' **     7       6     new hidtype       O_HTYP2
                      ' **
                      ' ***************************************************
2640                  arr_varOrphan(O_JNO, lngE) = ![journalno]
2650                  arr_varOrphan(O_JTYPE, lngE) = ![journaltype]
2660                  arr_varOrphan(O_GRP, lngE) = ![hid_grpnum]
2670                  arr_varOrphan(O_UNIQ, lngE) = ![uniqueid]
2680                  arr_varOrphan(O_GRP2, lngE) = CLng(0)        ' ** It may or may not need a new group
2690                  arr_varOrphan(O_UNIQ2, lngE) = vbNullString  ' ** It definitely needs a new uniqueid.
2700                  arr_varOrphan(O_HTYP2, lngE) = vbNullString  ' ** And also a new Hidden Type.
2710                End If
2720              End If
2730              If lngY < lngTmp02 Then .MoveNext
2740            Next  ' ** Each record in Me.RecordsetClone: lngY.
2750          End With
2760        Next  ' ** arr_varAvail(): lngX.
2770        rst.Close
2780      End If

          ' ** Now we have to see if lngNewGrpNum had other members not in this newly constituted group.
2790      Set rst = Me.RecordsetClone
2800      With rst
2810        .MoveLast
2820        lngTmp02 = .RecordCount
2830        .MoveFirst
            ' ** Check each record in LedgerHidden (or at least those currently on the screen).
2840        For lngX = 1& To lngTmp02
2850          If ![hid_grpnum] = lngNewGrpNum Then
                ' ** It's in the new group we're working with now.
2860            blnFound = False
                ' ** See if it's one of the ones in arr_varSet().
2870            For lngY = 0& To (lngSets - 1&)
2880              If arr_varSet(S_JNO, lngY) = ![journalno] Then
                    ' ** Aahh. It's one of those we're reassigning, so ignore it.
2890                blnFound = True
2900                Exit For
2910              End If
2920            Next
2930            If blnFound = False Then
                  ' ** This entry is definitely an orphan.
2940              lngOrphans = lngOrphans + 1&
2950              lngE = lngOrphans - 1&
2960              ReDim Preserve arr_varOrphan(O_ELEMS, lngE)
2970              arr_varOrphan(O_JNO, lngE) = ![journalno]
2980              arr_varOrphan(O_JTYPE, lngE) = ![journaltype]
2990              arr_varOrphan(O_GRP, lngE) = ![hid_grpnum]
3000              arr_varOrphan(O_UNIQ, lngE) = ![uniqueid]
3010              arr_varOrphan(O_GRP2, lngE) = CLng(0)        ' ** It definitely needs a new group
3020              arr_varOrphan(O_UNIQ2, lngE) = vbNullString  ' ** It definitely needs a new uniqueid.
3030              arr_varOrphan(O_HTYP2, lngE) = vbNullString  ' ** And also a new Hidden Type.
3040            End If
3050          End If
3060          If lngX < lngTmp02 Then .MoveNext
3070        Next
3080        .Close
3090      End With

          ' ** HiddenType.
3100      If lngSets = 2& Then
3110        If lngAssetCnt = lngSets Then
3120          strNewHidType = "NORM"
3130        ElseIf lngMiscCnt = lngSets Then
3140          strNewHidType = "NORM_MISC"
3150        Else
3160          strNewHidType = "MISC_2_GRP"
3170        End If
3180      Else
3190        If lngSets = 3& And lngAssetCnt = 2& And lngMiscCnt = 1& Then
3200          strNewHidType = "MISC_3_GRP"
3210        Else
3220          strNewHidType = "MULTI_GRP"
3230        End If
3240      End If

          ' ** Everyone gets the same hidtype.
3250      For lngX = 0& To (lngSets - 1&)
3260        arr_varSet(S_HTYP2, lngX) = strNewHidType
3270      Next

          ' ** Get the earliest hid_sortdate.
3280      For lngX = 0& To (lngSets - 1&)
3290        If arr_varSet(S_SDAT, lngX) < datNewSortDate Then
3300          datNewSortDate = arr_varSet(S_SDAT, lngX)
3310        End If
3320      Next

          ' ** And give it to everyone.
3330      For lngX = 0& To (lngSets - 1&)
3340        arr_varSet(S_SDAT2, lngX) = datNewSortDate
3350      Next

          ' ** All the rest of the fields are standardized.
          ' ** There are no multi-field indexes in LedgerHidden,
          ' ** and the only unique index is hid_id.

3360      Set dbs = CurrentDb
3370      With dbs
3380        Set rst = .OpenRecordset("LedgerHidden", dbOpenDynaset, dbConsistent)
3390        With rst

              ' ** First, update all the members of this newly constituted group.
3400          For lngX = 0& To (lngSets - 1&)
3410            .MoveFirst
3420            .FindFirst "[journalno] = " & CStr(arr_varSet(S_JNO, lngX))
3430            If .NoMatch = False Then
3440              .Edit
                  '![accountno] = {not changed}
                  '![journalno] = {not changed}
3450              ![hidtype] = arr_varSet(S_HTYP2, lngX)
3460              ![hid_grpnum] = arr_varSet(S_GRP2, lngX)
3470              ![hid_sort] = 1&
3480              ![hid_sortdate] = arr_varSet(S_SDAT2, lngX)
3490              ![uniqueid] = arr_varSet(S_UNIQ2, lngX)
3500              ![hid_order] = (lngX + 1&)
3510              ![hid_newmatch] = False
3520              ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
                  '![hid_datecreated] = {not changed}
3530              ![hid_datemodified] = Now()
3540              .Update
3550            Else
                  ' ** Shouldn't happen!
3560            End If
3570          Next

3580          .Close
3590        End With

            ' ***************************************************
            ' ** Array: arr_varOrphan()
            ' **
            ' **   Field  Element  Name              Constant
            ' **   =====  =======  ================  ==========
            ' **     1       0     journalno         O_JNO
            ' **     2       1     journaltype       O_JTYPE
            ' **     3       2     hid_grpnum        O_GRP
            ' **     4       3     uniqueid          O_UNIQ
            ' **     5       4     new hid_grpnum    O_GRP2
            ' **     6       5     new uniqueid      O_UNIQ2
            ' **     7       6     new hidtype       O_HTYP2
            ' **
            ' ***************************************************

            ' ** Examples:
            ' **   accountno      astno jno1   jno2   jtype1    jtype2
            ' **   =============== ==== ====== ====== ========= =========
            ' **   000000000002300_0412_020353_020350_Withdrawn_Deposit__
            ' **   000000000002300_0000_040395_000000_Misc.______________
            ' **          15         4     6      6       9         9

            ' ** Now figure out what to do with any orphans left behind.
            ' ** Remember, if this new group is comprised entirely of
            ' ** unmatched entries, there will be no orphans to deal with.

3600        Set rst = Me.RecordsetClone
3610        With rst

3620          If lngOrphans > 0& Then
3630            For lngX = 0& To (lngOrphans - 1&)
3640              strNewUniqueID = vbNullString: strNewHidType = vbNullString: lngNewGrpNum = 0&: datNewSortDate = CDate(1)
3650              .MoveFirst
3660              .FindFirst "[journalno] = " & CStr(arr_varOrphan(O_JNO, lngX))
3670              If .NoMatch = False Then
3680                strNewUniqueID = Left(arr_varOrphan(O_UNIQ, lngX), 16)  ' ** Include the trailing underscore.
3690                If ![assetno] = 0& Then
3700                  strNewUniqueID = strNewUniqueID & "0000_"
3710                Else
3720                  strNewUniqueID = strNewUniqueID & Right(String(4, "0") & CStr(![assetno]), 4) & "_"
3730                End If
3740                strNewUniqueID = strNewUniqueID & Right(String(6, "0") & CStr(![journalno]), 6) & "_"
3750                strNewUniqueID = strNewUniqueID & "000000_"
3760                strTmp01 = ![journaltype]
3770                strTmp01 = StringReplace(strTmp01, " ", "_")  ' ** Cost Adj.
3780                strNewUniqueID = strNewUniqueID & Left(strTmp01 & String(6, "_"), 9) & "_"
3790                strNewUniqueID = strNewUniqueID & String(9, "_")  ' ** No last underscore.
3800                strNewHidType = "GRP_NONE"
3810                If lngAvails > 0& Then
3820                  For lngY = 0& To (lngAvails - 1&)
3830                    If arr_varAvail(lngY) > 0& Then
3840                      lngNewGrpNum = arr_varAvail(lngY)
3850                      arr_varAvail(lngY) = 0&
3860                      Exit For
3870                    End If
3880                  Next
3890                End If
3900                If lngNewGrpNum = 0& Then
3910                  lngNewGrpNum = lngMaxGrpNum + 1&
3920                  lngMaxGrpNum = lngMaxGrpNum + 1&
3930                End If
3940                datNewSortDate = ![transdate]
3950                .Edit
3960                If strNewHidType = "GRP_NONE" Then
                      ' ** hidtype_id  hidtype     hidtype_description
                      ' ** ==========  ==========  ======================================================================================
                      ' **      1      NORM        2 entries in hidden group, with matching assetno (which could both be zero).
                      ' **      2      NORM_MISC   2 entries in hidden group, where both are 'Misc.', to be treated like a normal pair.
                      ' **      3      MISC_2_GRP  2 entries in hidden group, one 'Misc.' and one other.
                      ' **      4      MISC_3_GRP  3 entries in hidden group, one 'Misc.' and two other matching assetno.
                      ' **      5      MULTI_GRP   3 or more entries in hidden group, with matching assetno, multi-lot group
3970                End If
3980                ![hidtype] = strNewHidType
3990                ![hid_grpnum] = lngNewGrpNum
4000                ![hid_sort] = 3&
4010                ![hid_sortdate] = datNewSortDate
4020                ![uniqueid] = strNewUniqueID
4030                ![hid_order] = 1&
4040                ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
4050                ![hid_datemodified] = Now()
4060                .Update
4070              Else
                    ' ** Shouldn't happen!
4080              End If
4090            Next
4100          End If
4110          .Close
4120        End With

4130        .Close
4140      End With

4150      With Me

4160        .Requery

4170        MsgBox "Updates completed.", vbInformation + vbOKOnly, ("Update Successful" & Space(40))

4180        .cmdClose.SetFocus
4190        .cmdUpdate.Enabled = False

4200        .TimerInterval = TIMER_DELAY

4210      End With

          ' ** HiddenType enumeration:
          ' **   hidtype       hidtype_description
          ' **   ============  ======================================================================================
          ' **   NORM          2 entries in hidden group, with matching assetno (which could both be zero).
          ' **   NORM_MISC     2 entries in hidden group, where both are 'Misc.', to be treated like a normal pair.
          ' **   MISC_2_GRP    2 entries in hidden group, one 'Misc.' and one other.
          ' **   MISC_3_GRP    3 entries in hidden group, one 'Misc.' and two other matching assetno.
          ' **   MULTI_GRP     3 or more entries in hidden group, with matching assetno, multi-lot group.
          ' **   GRP_NONE      1 unmatched entry, any journaltype.  NEW!!!!!!!!!!!!!!!!!!!!

4220    Else
          ' ** This should have been prevented anyway.
4230      Beep
4240    End If

EXITP:
4250    Set rst = Nothing
4260    Set dbs = Nothing
4270    Exit Sub

ERRH:
4280    Select Case ERR.Number
        Case Else
4290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4300    End Select
4310    Resume EXITP

End Sub

Private Sub cmdSave_Click()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

4410    DoCmd.RunCommand acCmdSelectRecord
4420    DoCmd.RunCommand acCmdSaveRecord

EXITP:
4430    Exit Sub

ERRH:
4440    Select Case ERR.Number
        Case Else
4450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4460    End Select
4470    Resume EXITP

End Sub

Private Sub cmbAccounts_AfterUpdate()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

        Dim strFilter As String

4510    With Me
4520      If IsNull(.cmbAccounts) = False Then
4530        If Trim(.cmbAccounts) <> vbNullString Then
4540          strFilter = "[accountno] = '" & .cmbAccounts & "'"
4550          .Filter = strFilter
4560          .FilterOn = True
4570        Else
4580          .Filter = vbNullString
4590          .FilterOn = False
4600        End If
4610      Else
4620        .Filter = vbNullString
4630        .FilterOn = False
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

Private Sub opgAccountSource_AfterUpdate()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

4810    strAccountNo = vbNullString

4820    With Me
4830      If IsNull(.cmbAccounts) = False Then
4840        If Len(.cmbAccounts.Column(0)) > 0 Then
4850          strAccountNo = .cmbAccounts.Column(0)
4860        End If
4870      End If
4880      Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
4890        .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
4900        .opgAccountSource_optNumber_lbl.FontBold = True
4910        .opgAccountSource_optName_lbl.FontBold = False
4920      Case .opgAccountSource_optName.OptionValue
4930        .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
4940        .opgAccountSource_optNumber_lbl.FontBold = False
4950        .opgAccountSource_optName_lbl.FontBold = True
4960      End Select
4970      DoEvents
4980      If strAccountNo <> vbNullString Then
4990        .cmbAccounts = strAccountNo
5000      End If
5010    End With

EXITP:
5020    Exit Sub

ERRH:
5030    Select Case ERR.Number
        Case Else
5040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5050    End Select
5060    Resume EXITP

End Sub

Private Sub hid_newmatch_lbl_DblClick(Cancel As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "hid_newmatch_lbl_DblClick"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long

5110    With Me
5120      Set rst = .RecordsetClone
5130      With rst
5140        If .BOF = True And .EOF = True Then
              ' ** Well, this certainly shouldn't happen!
5150        Else
5160          .MoveLast
5170          lngRecs = .RecordCount
5180          .MoveFirst
5190          For lngX = 1& To lngRecs
5200            If ![hid_newmatch] = True Then
5210              .Edit
5220              ![hid_newmatch] = False
5230              .Update
5240            End If
5250            If lngX < lngRecs Then .MoveNext
5260          Next
5270        End If
5280        .Close
5290      End With
5300      .cmdUpdate.Enabled = False
5310      .assetno_chk_box.BackColor = CLR_WHT
5320      .assetno_chk_box_lbl.ForeColor = CLR_BLU
5330      .assetno_chk_box_lbl.Visible = False
'1840      .assetno_chk_lbl2.Visible = False
5340      .misc_chk_box_lbl.Visible = False
5350      .icash_subtot = 0
5360      .pcash_subtot = 0
5370      .cost_subtot = 0
5380      lngSets = 0&
5390      ReDim arr_varSet(S_ELEMS, 0)
5400      .TimerInterval = TIMER_DELAY
5410    End With

EXITP:
5420    Set rst = Nothing
5430    Exit Sub

ERRH:
5440    Select Case ERR.Number
        Case Else
5450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5460    End Select
5470    Resume EXITP

End Sub

Private Sub hid_newmatch_AfterUpdate()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "hid_newmatch_AfterUpdate"

        Dim lngAssetCnt As Long, lngMiscCnt As Long
        Dim lngAssetNo As Long
        Dim blnFound As Boolean

5510    With Me

5520      cmdSave_Click  ' ** Procedure: Above.

5530      Select Case .hid_newmatch
          Case True

5540        .icash_subtot = (.icash_subtot + .ICash)
5550        .pcash_subtot = (.pcash_subtot + .PCash)
5560        .cost_subtot = (.cost_subtot + .Cost)

5570        lngSets = lngSets + 1&
5580        lngE = lngSets - 1&
5590        ReDim Preserve arr_varSet(S_ELEMS, lngE)
            ' *****************************************************
            ' ** Array: arr_varSet()
            ' **
            ' **   Field  Element  Name                Constant
            ' **   =====  =======  ==================  ==========
            ' **     1       0     journalno           S_JNO
            ' **     2       1     journaltype         S_JTYPE
            ' **     3       2     accountno           S_ACTNO
            ' **     4       3     assetno             S_ASTNO
            ' **     5       4     hid_sortdate        S_SDAT
            ' **     6       5     icash               S_ICASH
            ' **     7       6     pcash               S_PCASH
            ' **     8       7     cost                S_COST
            ' **     9       8     IsMatched           S_MTCHD
            ' **    10       9     hid_grpnum          S_GRP
            ' **    11      10     uniqueid            S_UNIQ
            ' **    12      11     new hid_sortdate    S_SDAT2
            ' **    12      11     new hid_grpnum      S_GRP2
            ' **    13      12     new uniqueid        S_UNIQ2
            ' **    14      13     new hidtype         S_HTYP2
            ' **    15      14     Remove              S_REMOV
            ' **
            ' *****************************************************
5600        arr_varSet(S_JNO, lngE) = .journalno
5610        arr_varSet(S_JTYPE, lngE) = .journaltype
5620        arr_varSet(S_ACTNO, lngE) = .accountno
5630        arr_varSet(S_ASTNO, lngE) = Nz(.assetno, 0)
5640        arr_varSet(S_SDAT, lngE) = .hid_sortdate
5650        arr_varSet(S_ICASH, lngE) = .ICash
5660        arr_varSet(S_PCASH, lngE) = .PCash
5670        arr_varSet(S_COST, lngE) = .Cost
5680        arr_varSet(S_MTCHD, lngE) = .IsMatched
5690        arr_varSet(S_GRP, lngE) = .hid_grpnum
5700        arr_varSet(S_UNIQ, lngE) = .uniqueid
5710        arr_varSet(S_GRP2, lngE) = CLng(0)
5720        arr_varSet(S_SDAT2, lngE) = CDate(1)
5730        arr_varSet(S_UNIQ2, lngE) = vbNullString
5740        arr_varSet(S_HTYP2, lngE) = vbNullString
5750        arr_varSet(S_REMOV, lngE) = CBool(False)

5760      Case False

5770        .icash_subtot = (.icash_subtot - .ICash)
5780        .pcash_subtot = (.pcash_subtot - .PCash)
5790        .cost_subtot = (.cost_subtot - .Cost)

5800        blnFound = False
5810        For lngX = 0& To (lngSets - 1&)
5820          If arr_varSet(S_JNO, lngX) = .journalno Then
5830            blnFound = True
5840            arr_varSet(S_REMOV, lngX) = CBool(True)
5850            Exit For
5860          End If
5870        Next

5880        If blnFound = True Then
5890          SetRemove  ' ** Procedure: Below.
5900        Else
              ' ** Else what? It should always find it!
5910        End If

5920      End Select

          ' ** Set the Asset and Misc. checkboxes.
5930      lngAssetCnt = 0&: lngMiscCnt = 0&
5940      For lngX = 0& To (lngSets - 1&)
5950        If arr_varSet(S_ASTNO, lngX) <> 0& Then
5960          lngAssetCnt = lngAssetCnt + 1&
5970        Else
5980          lngMiscCnt = lngMiscCnt + 1&
5990        End If
6000      Next

6010      If lngAssetCnt = 0& Then
6020        If .assetno_chk_box_lbl.Visible = True Then
6030          .assetno_chk_box_lbl.Visible = False
6040        End If
6050      Else
            ' ** Just check.
6060        If .assetno_chk_box_lbl.Visible = False Then
6070          .assetno_chk_box_lbl.Visible = True
6080        End If
6090      End If

6100      If lngMiscCnt = 0& Then
6110        If .misc_chk_box_lbl.Visible = True Then
6120          .misc_chk_box_lbl.Visible = False
6130        End If
6140      Else
            ' ** Just check.
6150        If .misc_chk_box_lbl.Visible = False Then
6160          .misc_chk_box_lbl.Visible = True
6170        End If
6180      End If

          ' ** Also check whether the assets are mismatched.
6190      .assetno_chk_box_lbl.ForeColor = CLR_BLU
6200      .assetno_chk_box.BackColor = CLR_WHT
'2660      .assetno_chk_lbl2.Visible = False
6210      If lngAssetCnt > 0& Then
6220        lngAssetNo = 0&
6230        For lngX = 0& To (lngSets - 1&)
6240          If arr_varSet(S_ASTNO, lngX) <> 0& Then
6250            If lngAssetNo = 0& Then
6260              lngAssetNo = arr_varSet(S_ASTNO, lngX)
6270            Else
6280              If arr_varSet(S_ASTNO, lngX) <> lngAssetNo Then
6290                Beep
6300                .assetno_chk_box_lbl.ForeColor = CLR_BLK
6310                .assetno_chk_box.BackColor = CLR_LTRED
'2780                .assetno_chk_lbl2.Visible = True
6320              End If
6330            End If
6340          End If
6350        Next
6360      End If

          ' ** The form doesn't finish totaling the sum in hid_newmatch_sum by the time
          ' ** the code reaches here, so color highlighting is done in Form_Timer().
          ' ** For slower computers, this may have to lengthened beyond 1/2 second.
6370      .TimerInterval = TIMER_DELAY

6380    End With

EXITP:
6390    Exit Sub

ERRH:
6400    Select Case ERR.Number
        Case Else
6410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6420    End Select
6430    Resume EXITP

End Sub

Private Sub SetRemove()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "SetRemove"

        Dim arr_varTmp01() As Variant, lngTmp02 As Long
        Dim lngY As Long

6510    lngTmp02 = 0&
6520    ReDim arr_varTmp01(S_ELEMS, 0)

6530    If lngSets > 1& Then

6540      For lngX = 0& To (lngSets - 1&)
6550        If arr_varSet(S_REMOV, lngX) = False Then
              ' ** Only add to the arr_varTmp01() array if it's not the one to be removed.
6560          lngTmp02 = lngTmp02 + 1&
6570          lngE = lngTmp02 - 1&
6580          ReDim Preserve arr_varTmp01(S_ELEMS, lngE)
6590          For lngY = 0& To S_ELEMS
6600            arr_varTmp01(lngY, lngE) = arr_varSet(lngY, lngX)
6610          Next
6620        End If
6630      Next

6640      If lngTmp02 > 0& Then
6650        lngSets = lngTmp02
6660        ReDim arr_varSet(S_ELEMS, (lngSets - 1&))
6670        For lngX = 0& To (lngTmp02 - 1&)
6680          For lngY = 0& To S_ELEMS
6690            arr_varSet(lngY, lngX) = arr_varTmp01(lngY, lngX)
6700          Next
6710        Next
6720      Else
            ' ** There shouldn't ever be more than one S_REMOV = True!
6730        lngSets = 0&
6740        ReDim arr_varSet(S_ELEMS, 0)
6750      End If

6760    Else
6770      If arr_varSet(S_REMOV, 0) = True Then
6780        lngSets = 0&
6790        ReDim arr_varSet(S_ELEMS, 0)
6800      End If
6810    End If

EXITP:
6820    Exit Sub

ERRH:
6830    Select Case ERR.Number
        Case Else
6840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6850    End Select
6860    Resume EXITP

End Sub

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

6910    Select Case intWhere
        Case acCmdRecordsGoToFirst
6920      DoCmd.RunCommand acCmdRecordsGoToFirst
6930    Case acCmdRecordsGoToLast
6940      DoCmd.RunCommand acCmdRecordsGoToLast
6950    Case acCmdRecordsGoToPrevious
6960      DoCmd.RunCommand acCmdRecordsGoToPrevious
6970    Case acCmdRecordsGoToNext
6980      DoCmd.RunCommand acCmdRecordsGoToNext
6990    Case Else
7000      If IsMissing(varID) = False Then
7010        With Me
7020          Set rst = .RecordsetClone
7030          With rst
7040            .FindFirst "[journalno] = " & CStr(varID)
7050            If .NoMatch = False Then
7060              Me.Bookmark = .Bookmark
7070            End If
7080            .Close
7090          End With
7100        End With
7110      End If
7120    End Select

EXITP:
7130    Set rst = Nothing
7140    Exit Sub

ERRH:
7150    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
7160    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
7170    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
7180    Case Else
7190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7200    End Select
7210    Resume EXITP

End Sub
