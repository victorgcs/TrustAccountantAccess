Attribute VB_Name = "modJrnlCol_Recs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modJrnlCol_Recs"

'VGC 10/26/2017: CHANGES!

'FROM V2.2.30: THIS IS ONLY HERE FOR frmJournal_Columns_TaxLot!

' ** Array: arr_varNewRec().
Private lngNewRecs As Long, arr_varNewRec As Variant
Private Const N_ELEMS As Integer = 1  ' ** Array's first-element UBound().
Private Const N_ID   As Integer = 0
'Private Const N_CMTD As Integer = 1

Private lngRecsCur As Long
' **

Public Sub JC_Rec_AddRec(ByRef frmSub As Access.Form, Optional varFromZero As Variant)
' ** There's some duplication between here and
' ** Form_Timer() after the JournalType is chosen.
' ** Called by:
' **   frmJournal_Columns:
' **     Form_Timer()
' **     cmdAdd_Click()
' **   modJrnlCol_Forms:
' **     JC_Frm_Map_Return()
' **     JC_Frm_TaxLot()
' **   modJrnlCol_Keys:
' **     JC_Key_Sub_Next()
' **   modJrnlCol_Recs:
' **     JC_Rec_Commit()
' **     JC_Rec_CostAdj()
' ** References to other Subs/Functions in frmJournal_Columns_Sub:
' **   FromZero_GetSet()
' **   NextRec_GetSet()
' **   GTR_NoAdd_GetSet()
' **   PostDate_GetSet()
' **   MoveRec()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Rec_AddRec"

        Dim datPostingDate As Date
        Dim lngNewJrnlColID As Long
        Dim strSaveMoveCtl As String
        Dim blnNextRec As Boolean, blnFromZero As Boolean, blnGTR_NoAdd As Boolean, blnNoMove As Boolean
        Dim lngRetVal As Long

110     With frmSub

120       Select Case IsMissing(varFromZero)
          Case True
130         blnFromZero = .FromZero_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
140       Case False
150         blnFromZero = CBool(varFromZero)
160       End Select

170       blnNextRec = .NextRec_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
180       lngNewJrnlColID = .NewJColID_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
190       blnGTR_NoAdd = .GTR_NoAdd_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
200       datPostingDate = .PostDate_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
210       strSaveMoveCtl = .SaveMoveCtl_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
220       blnNoMove = .NoMove_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.

230       If blnGTR_NoAdd = False Then
240         If datPostingDate = 0 Then
              ' ** Sometimes this seems to get lost!
250           JC_Msc_Pub_Reset .Parent, frmSub  ' ** Module Function: modJrnlCol_Misc.
260           DoEvents
270           datPostingDate = .PostDate_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
280         End If

290         JC_Msc_Memo_Set False, .Parent  ' ** Module Procedure: modJrnlCol_Misc.
300         .AllowAdditions = True
310         .MoveRec acCmdRecordsGoToNew  ' ** Form Procedure: frmJournal_Columns_Sub.
320         lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
330         .transdate.SetFocus
340         .transdate = datPostingDate
350         .posted = False
360         .pershare = 0#
370         .Reinvested = False
380         .Location_ID = 1&
390         .PrintCheck = False
400         .revcode_ID = REVID_INC
410         .revcode_TYPE = REVTYP_INC
420         .revcode_DESC = "Unspecified Income"
430         .revcode_DESC_display = Null
440         .taxcode = TAXID_INC
450         .taxcode_type = TAXTYP_INC
460         .taxcode_description = "Unspecified Income"
470         .taxcode_description_display = Null
480         .journal_USER = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
490         .JrnlCol_DateModified = Now()
500         .rate = 0#
510         .IsAverage = False
520         DoEvents
            'JC_MSC_GTR_Fill  ' ** Module Procedure: modJrnlCol_Misc.

530         If blnFromZero = False Then
              ' ** OnOpen, Form_Timer() in frmJournal_Columns calls JC_Key_Sub_Next().
              ' ** It is that call that then initiates this call to JC_Rec_AddRec().
              ' ** A call to here shouldn't always regenerate another recursive call to here.
              ' ** It would be nice if blnFromZero were already True by the time it gets here!
              ' ** At this point, the new record's already been created.
540           strSaveMoveCtl = JC_Key_Sub_Next("posted_AfterUpdate", blnNextRec, blnFromZero, True, "AddRec")  ' ** Module Function: modJrnlCol_Keys.
550           .SaveMoveCtl_GetSet False, strSaveMoveCtl  ' ** Form Function: frmJournal_Columns_Sub.
560           .NextRec_GetSet False, blnNextRec  ' ** Form Function: frmJournal_Columns_Sub.
570           .FromZero_GetSet False, blnFromZero  ' ** Form Function: frmJournal_Columns_Sub.
580         End If
            ' ** On a completely empty record (save transdate),
            ' ** I want to keep the top buttons active!
590         strSaveMoveCtl = vbNullString
600         .SaveMoveCtl_GetSet False, strSaveMoveCtl  ' ** Form Function: frmJournal_Columns_Sub.
610         blnNoMove = True
620         .NoMove_GetSet False, blnNoMove  ' ** Form Function: frmJournal_Columns_Sub.
630         DoEvents
640         .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
650         DoEvents
660         lngNewJrnlColID = .JrnlCol_ID
670         .NewJColID_GetSet False, lngNewJrnlColID  ' ** Form Function: frmJournal_Columns_Sub.
680         .AllowAdditions = False

690       End If
700     End With

        'HERE! 1  False  Detail_MouseMove
        'HERE! 6  False  Detail_MouseMove
        'Form_Timer(): 1  False
        'JC_Key_Sub_Next() 1  False
        'JC_Rec_AddRec() 1  False
        'JC_Rec_AddRec() 2  False
        'JC_Rec_AddRec() 3  True
        'JC_Rec_AddRec() 4  True
        'JC_Key_Sub_Next() 2  True
        'Form_Timer(): 2  True

EXITP:
710     Exit Sub

ERRH:
720     Select Case ERR.Number
        Case Else
730       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
740     End Select
750     Resume EXITP

End Sub

Public Function JC_Rec_Commit(ByRef blnNextRec As Boolean, ByRef blnFromZero As Boolean, ByRef frmSub As Access.Form, Optional varSilent As Variant, Optional varAddNew As Variant) As Boolean
' ** gblnCrtRpt_Zero = Commit  Yes/No
' ** gblnMessage     = New Rec Yes/No
' ** Called by:
' **   frmJournal_Columns_Sub:
' **     description_AfterUpdate()
' **     revcode_ID_AfterUpdate()
' **     taxcode_AfterUpdate()
' **     Reinvested_AfterUpdate()
' **     CommitAll()
' **   modJrnlCol_Keys:
' **     JC_Key_Sub()
' **   modJrnlCol_Forms:
' **     JC_Frm_Map_Return()
' **     JC_Frm_TaxLot()
' **   modJrnlCol_Recs:
' **     JC_Rec_CostAdj() (Above.)
' ** References to other Subs/Functions in frmJournal_Columns_Sub:
' **   posted_AfterUpdate()
' **   RecCnt()
' **   MoveRec()
' **   JC_Rec_AddRec()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Rec_Commit"

        Dim strDocName As String
        Dim strNext As String, strCallingForm As String
        Dim blnSilent As Boolean
        Dim blnRetVal As Boolean, lngRetVal As Long

810     blnRetVal = True

820     With frmSub

830       Select Case IsMissing(varSilent)
          Case True
840         blnSilent = False
850       Case False
860         blnSilent = CBool(varSilent)
870         Select Case IsMissing(varAddNew)
            Case True
880           gblnMessage = True  ' ** Add new record.
890         Case False
900           gblnMessage = CBool(varAddNew)
910         End Select
920       End Select

930       Select Case blnSilent
          Case True
940         gblnCrtRpt_Zero = True  ' ** Return value: Passed Tests.
950       Case False
960         gblnCrtRpt_Zero = False  ' ** Borrowing this variable from Court Reports.
970         strCallingForm = .Parent.Name
980         gblnMessage = False  ' ** Borrowing this variable from wherever.
990         strDocName = "frmJournal_Columns_Add"
1000        DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
1010      End Select

1020      Select Case gblnCrtRpt_Zero  ' ** Indicates whether to commit.
          Case True
1030        lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
1040  On Error Resume Next
1050        .posted.SetFocus
            ' ** Not sure why this sometimes errors!
1060  On Error GoTo ERRH
1070        .posted = True
1080        .posted_AfterUpdate  ' ** Form Procedure: frmJournal_Columns_Sub.
1090        Select Case gblnCrtRpt_Zero
            Case True
              ' ** It passed the tests.
1100          Select Case gblnMessage
              Case True
                ' ** Add a new record.
1110            gblnMessage = False
1120            JC_Rec_AddRec frmSub  ' ** Function: Above.
                '.AddRec  ' ** Form Function: frmJournal_Columns_Sub.
                'A KeyDown() event sends it to JC_Rec_Commit().
                'JC_Rec_Commit() sends it to AddRec().
                'AddRec() sets focus to transdate.
                'Then AddRec() sends it cmdSave_Click().
                'cmdSave_Click() should let it return to JC_Rec_Commit() without moving it.
                'At which point JC_Rec_Commit() should just return it to the KeyDown() event.
                ' ** When coming from JC_Frm_Map_Return(), somehow the last
                ' ** Map entry gets unposted, and the empty entry disappears!
                ' ** So, I'm no longer telling it to add on the last entry.
1130          Case False
                ' ** If there are more entries, proceed to the next,
                ' ** otherwise exit the subform
1140            lngRecsCur = .RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
1150            lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
1160            If .CurrentRecord < lngRecsCur Then
1170              .MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Columns_Sub.
1180              strNext = JC_Key_Sub_Next("posted_AfterUpdate", blnNextRec, blnFromZero, False, "First")  ' ** Module Function: modJrnlCol_Keys.
1190              .FocusHolder.SetFocus
1200              .Controls(strNext).SetFocus
1210              lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
1220            Else
                  ' ** Regardless, keep the focus in the subform!
1230              DoCmd.SelectObject acForm, .Parent.Name, False
1240              .Parent.frmJournal_Columns_Sub.SetFocus
1250              .FocusHolder.SetFocus
1260              .transdate.SetFocus
1270              lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
1280            End If
1290          End Select
1300        Case False
              ' ** It failed a test.
1310          If blnSilent = True Then blnRetVal = False
1320        End Select
1330      Case False
            ' ** Either it failed, or just, No, Don't Commit.
1340        Select Case gblnMessage
            Case True
              ' ** Add a new record.
1350          gblnMessage = False
1360          JC_Rec_AddRec frmSub  ' ** Function: Above.
              '.AddRec  ' ** Form Function: frmJournal_Columns_Sub.
1370        Case False
              ' ** Regardless, keep the focus in the subform!
1380          DoCmd.SelectObject acForm, .Parent.Name, False
1390          .Parent.frmJournal_Columns_Sub.SetFocus
1400          .transdate.SetFocus
1410          lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
1420        End Select
            'JC_MSC_GTR_Fill  ' ** Module Procedure: modJrnlCol_Misc.
1430      End Select

1440      gblnCrtRpt_Zero = False
1450      gblnMessage = False

1460    End With

EXITP:
1470    JC_Rec_Commit = blnRetVal
1480    Exit Function

ERRH:
1490    DoCmd.Hourglass False
1500    frmSub.posted = False
1510    frmSub.Parent.CommitNoClose_Set False  ' ** Form Procedure: frmJournal_Columns.
1520    blnRetVal = False
1530    Select Case ERR.Number
        Case Else
1540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1550    End Select
1560    Resume EXITP

End Function

Public Function JC_Rec_CommitAll(ByVal blnSilent As Boolean, frmSub As Access.Form) As Boolean
' ** Called by:
' **   frmJournal_Columns:
' **     cmdUncomComAll_Click()
' ** References to other Subs/Functions in frmJournal_Columns_Sub:
' **   RecCnt()
' **   MoveRec()
' **   JC_Rec_Commit()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Rec_CommitAll"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long, lngCommitted As Long, lngFails As Long
        Dim lngUncoms As Long, arr_varUnCom() As Variant
        Dim blnNextRec As Boolean, blnFromZero As Boolean
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varUnCom().
        Const U_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const U_ID   As Integer = 0
        Const U_JTYP As Integer = 1
        Const U_CMTD As Integer = 2

1610    blnRetVal = True

1620    With frmSub

1630      DoCmd.Hourglass True
1640      DoEvents

1650      lngRecsCur = .RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
1660      lngCommitted = .RecsTot_Committed
1670      If lngRecsCur > lngCommitted Then

1680        lngUncoms = 0&
1690        ReDim arr_varUnCom(U_ELEMS, 0)

1700        Set rst = .RecordsetClone
1710        With rst
1720          .MoveLast
1730          lngRecs = .RecordCount
1740          .MoveFirst
1750          For lngX = 1& To lngRecs
1760            If ![posted] = False Then
1770              If IsNull(![accountno]) = False And IsNull(![journaltype]) = False Then
1780                lngUncoms = lngUncoms + 1&
1790                lngE = lngUncoms - 1&
1800                ReDim Preserve arr_varUnCom(U_ELEMS, lngE)
1810                arr_varUnCom(U_ID, lngE) = ![JrnlCol_ID]
1820                arr_varUnCom(U_JTYP, lngE) = ![journaltype]
1830                arr_varUnCom(U_CMTD, lngE) = CBool(False)
1840              End If
1850            End If
1860            If lngX < lngRecs Then .MoveNext
1870          Next
1880          .Close
1890        End With

1900        If lngUncoms > 0& Then

1910          For lngX = 0& To (lngUncoms - 1&)
1920            .MoveRec 0, arr_varUnCom(U_ID, lngX)  ' ** Form Procedure: frmJournal_Columns_Sub.
1930            blnNextRec = frmSub.NextRec_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
1940            blnFromZero = frmSub.FromZero_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
1950            arr_varUnCom(U_CMTD, lngX) = JC_Rec_Commit(blnNextRec, blnFromZero, frmSub, True)  ' ** Module Function: modJrnlCol_Recs.
                'arr_varUnCom(U_CMTD, lngX) = CommitRec(True)  ' ** Function: Below.
1960            frmSub.NextRec_GetSet False, blnNextRec  ' ** Form Function: frmJournal_Columns_Sub.
1970            frmSub.FromZero_GetSet False, blnFromZero  ' ** Form Function: frmJournal_Columns_Sub.
1980          Next

1990          lngFails = 0&
2000          For lngX = 0& To (lngUncoms - 1&)
2010            If arr_varUnCom(U_CMTD, lngX) = False Then
2020              lngFails = lngFails + 1&
2030            End If
2040          Next

2050          If lngFails = 0& Then
2060            If blnSilent = False Then
2070              DoCmd.Hourglass False
2080              Beep
2090              MsgBox "All uncommitted entries have been committed to the Posting Journal.", _
                    vbInformation + vbOKOnly, "All Entries Successfully Committed"
2100              DoCmd.SelectObject acForm, .Parent.Name, False
2110              .Parent.cmdAdd.SetFocus
2120              DoEvents
2130            End If
2140          Else
2150            blnRetVal = False
2160            If blnSilent = False Then
2170              If lngFails = lngUncoms Then
2180                DoCmd.Hourglass False
2190                Beep
2200                MsgBox "None of the entries could be committed." & vbCrLf & _
                      "Check the entries and try again.", vbInformation + vbOKOnly, "Commit Failed"
2210              Else
2220                DoCmd.Hourglass False
2230                Beep
2240                MsgBox "Some of the entries could not be committed." & vbCrLf & _
                      "Check those remaining uncommitted and try again.", vbInformation + vbOKOnly, "Commit Incomplete"
2250              End If
2260            End If
2270          End If

2280        Else
2290          DoCmd.Hourglass False
2300          MsgBox "There are no uncommited entries.", vbInformation + vbOKOnly, "Nothing To Do"
2310          DoCmd.SelectObject acForm, .Parent.Name, False
2320          .Parent.cmdAdd.SetFocus
2330          DoEvents
2340        End If
2350      Else
2360        DoCmd.Hourglass False
2370        MsgBox "There are no uncommited entries.", vbInformation + vbOKOnly, "Nothing To Do"
2380        DoCmd.SelectObject acForm, .Parent.Name, False
2390        .Parent.cmdAdd.SetFocus
2400        DoEvents
2410      End If
2420    End With

2430    DoCmd.Hourglass False

EXITP:
2440    Set rst = Nothing
2450    JC_Rec_CommitAll = blnRetVal
2460    Exit Function

ERRH:
2470    DoCmd.Hourglass False
2480    blnRetVal = False
2490    Select Case ERR.Number
        Case Else
2500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2510    End Select
2520    Resume EXITP

End Function

Public Sub JC_Rec_DelRec(ByRef blnGoneToReport As Boolean, ByRef blnGoneToReport2 As Boolean, ByRef blnRecsTotUpdate As Boolean, ByRef frmSub As Access.Form)
' ** Called by:
' **   frmJournal_Columns:
' **     cmdDelete_Click()
' **   frmJournal_Columns_Sub:
' **     Form_Timer()
' **   frmJournal_Columns_TaxLot:
' **     cmdCancel_Click()
' ** References to other Subs/Functions in frmJournal_Columns_Sub:
' **     DoTheDeed()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Rec_DelRec"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngJrnlID As Long
        Dim strErrDesc As String, strThisJType As String
        Dim blnContinue As Boolean

2610    blnContinue = True

2620    With frmSub
2630      strThisJType = Nz(.journaltype, vbNullString)
2640      lngJrnlID = Nz(.Journal_ID, 0)
          ' ** If this entry is already in the Journal, delete it there
          ' ** first, in case someone else is currently editing it,
          ' ** in which case we won't delete it here.
2650      If lngJrnlID > 0& Then
2660        Set dbs = CurrentDb
2670        With dbs
              ' ** Delete Journal, by specified [jid].
2680          Set qdf = .QueryDefs("qryJournal_Columns_19")
2690          With qdf.Parameters
2700            ![jid] = lngJrnlID
2710          End With
2720  On Error Resume Next
2730          qdf.Execute dbFailOnError
2740          If ERR.Number <> 0 Then
2750            blnContinue = False
2760            strErrDesc = ERR.description
2770  On Error GoTo ERRH
2780            MsgBox "Trust Accountant was unable to completely delete this entry." & vbCrLf & _
                  "Someone else may be accessing it." & vbCrLf & vbCrLf & _
                  "Try again later." & vbCrLf & vbCrLf & strErrDesc, vbInformation + vbOKOnly, "Delete Failed"
2790          Else
2800  On Error GoTo ERRH
2810          End If
2820          .Close
2830        End With
2840        Set qdf = Nothing
2850        Set dbs = Nothing
2860      Else
            ' ** An uncommitted record.
2870        lngJrnlID = Nz(.JrnlCol_ID, 0)
2880        If lngJrnlID > 0& Then
2890          Set dbs = CurrentDb
2900          With dbs
                ' ** Delete tblJournal_Column, by specified [jid].
2910            Set qdf = .QueryDefs("qryJournal_Columns_19_01")
2920            With qdf.Parameters
2930              ![jid] = lngJrnlID
2940            End With
2950  On Error Resume Next
2960            qdf.Execute dbFailOnError
2970            If ERR.Number = 0 Then
                  ' ** Succeeded, so no need to do the rest.
2980  On Error GoTo ERRH
2990              blnContinue = False
3000            Else
3010  On Error GoTo ERRH
3020            End If
3030            .Close
3040          End With
3050          Set qdf = Nothing
3060          Set dbs = Nothing
3070        End If
3080      End If
3090      If blnContinue = True Then
3100        .AllowDeletions = True
3110        If blnGoneToReport = True Then
              ' ** This has become an unholy mess!
3120  On Error Resume Next
3130          DoCmd.RunCommand acCmdDeleteRecord
3140          If ERR.Number <> 0 Then
3150  On Error GoTo ERRH
3160            blnGoneToReport = False
3170            blnGoneToReport2 = True
3180            .TimerInterval = 100&
3190          Else
3200  On Error GoTo ERRH
3210            blnGoneToReport = False
3220          End If
3230        Else
3240          frmSub.DoTheDeed  ' ** Form Procedure: frmJournal_Columns_Sub.
3250        End If
3260        If blnGoneToReport = False Then
3270          .AllowDeletions = False
3280          DoEvents
3290          If strThisJType = "Paid" Then
3300            If .Parent.JrnlMemo_Memo.Visible = True Then
3310              JC_Msc_Memo_Set False, .Parent  ' ** Module Procedure: modJrnlCol_Misc.
3320              .Parent.JrnlMemo_Memo = Null
3330            End If
3340          End If
3350          blnRecsTotUpdate = True
3360          .TimerInterval = 3000&
3370        End If
3380      Else
3390        .Requery
3400      End If
3410      gblnDeleting = False
3420    End With

EXITP:
3430    Set qdf = Nothing
3440    Set dbs = Nothing
3450    Exit Sub

ERRH:
3460    Select Case ERR.Number
        Case Else
3470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3480    End Select
3490    Resume EXITP

End Sub

Public Sub JC_Rec_CostAdj(ByRef lngNewJrnlColID As Long, ByVal strCallingForm As String, ByRef frmSub As Access.Form)
' ** Called by:
' **   frmJournal_Columns_Sub:
' **     revcode_ID_AfterUpdate()
' **     taxcode_AfterUpdate()
' **   modJrnlCol_Keys:
' **     JC_Key_Sub()
' ** References to other Subs/Functions in frmJournal_Columns_Sub:
' **   CommitRec()
' **   AddRec()
' **   NewRecAdd()
' **   MoveRec()
' **   RecalcTots()
' **   NewRecRedim()
' **   NewRecGet()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Rec_CostAdj"

        Dim strDocName As String
        Dim lngJrnlColID As Long
        Dim blnNextRec As Boolean, blnFromZero As Boolean
        Dim lngTmp01 As Long, strTmp02 As String, strTmp03 As String, lngTmp04 As Long, lngTmp06 As Long, strTmp05 As String
        Dim strTmp07 As String, lngTmp08 As Long, lngTmp09 As Long, strTmp10 As String, lngTmp11 As Long, blnTmp12 As Boolean
        Dim lngX As Long
        Dim lngRetVal As Long

3510    With frmSub

3520      gblnCrtRpt_Zero = False  ' ** Borrowing this variable from Court Reports.
3530      gblnMessage = False  ' ** Borrowing this variable from wherever.
3540      strDocName = "frmJournal_Columns_Add"
3550      DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm

3560      Select Case gblnCrtRpt_Zero  ' ** Indicates whether to commit.
          Case True
            ' ** Yes, commit.

3570        DoCmd.Hourglass True
3580        DoEvents

3590        lngJrnlColID = .JrnlCol_ID  ' ** This is record 1 of the Cost Adj. series.
3600        lngTmp01 = .journaltype_sortorder
3610        strTmp02 = .shortname
3620        strTmp03 = .assetno_description
3630        lngTmp04 = .Location_ID
3640        strTmp05 = .Loc_Name
3650        lngTmp06 = .revcode_ID
3660        strTmp07 = Nz(.revcode_DESC, "Null")
3670        lngTmp08 = .revcode_TYPE
3680        lngTmp09 = .taxcode
3690        strTmp10 = Nz(.taxcode_description, "Null")
3700        lngTmp11 = Nz(.taxcode_type, 0)
3710        blnTmp12 = False  ' ** Add new, empty record.

3720        .NewRecRedim  ' ** Form Procedure: frmJournal_Columns_Sub.

            ' ** Add the 1st record to the array.
3730        .NewRecAdd lngJrnlColID  ' ** Form Procedure: frmJournal_Columns_Sub.

3740        DistributeCost frmSub, lngTmp01, strTmp02, strTmp03, lngTmp04, strTmp05, lngTmp06, strTmp07, lngTmp08, lngTmp09, strTmp10, lngTmp11  ' ** Module Procedure: modPurchaseSold.

3750        arr_varNewRec = .NewRecGet  ' ** Form Function: frmJournal_Columns_Sub.
3760        Select Case IsEmpty(arr_varNewRec(N_ID, 0))
            Case True
3770          lngNewRecs = 0&
3780        Case False
3790          lngNewRecs = UBound(arr_varNewRec, 2) + 1&
3800        End Select

3810        DoEvents
3820        If lngNewRecs > 0& Then
3830          For lngX = 0& To (lngNewRecs - 1&)
3840            .MoveRec 0, arr_varNewRec(N_ID, lngX)  ' ** Form Procedure: frmJournal_Columns_Sub.
3850            DoEvents
3860            If lngX = (lngNewRecs - 1&) Then
3870              blnTmp12 = True  ' ** Only add new, empty record after last one.
3880            End If
3890            blnNextRec = frmSub.NextRec_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
3900            blnFromZero = frmSub.FromZero_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
3910            JC_Rec_Commit blnNextRec, blnFromZero, frmSub, True, blnTmp12  ' ** Function: Above.
                '.CommitRec True, blnTmp12  ' ** Form Function: frmJournal_Columns_Sub.
3920            frmSub.NextRec_GetSet False, blnNextRec  ' ** Form Function: frmJournal_Columns_Sub.
3930            frmSub.FromZero_GetSet False, blnFromZero  ' ** Form Function: frmJournal_Columns_Sub.
3940            DoEvents
3950          Next
3960        End If

3970        lngNewRecs = 0&
3980        ReDim arr_varNewRec(N_ELEMS, 0)
3990        .NewRecRedim  ' ** Form Procedure: frmJournal_Columns_Sub.

4000        If lngNewJrnlColID > lngJrnlColID Then
4010          lngTmp01 = lngNewJrnlColID
4020        Else
4030          lngTmp01 = lngJrnlColID
4040        End If
4050        .MoveRec 0, lngTmp01  ' ** Form Procedure: frmJournal_Columns_Sub.
4060        lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.

4070        DoCmd.Hourglass False

4080        Select Case gblnMessage
            Case True
              ' ** Add a new record.
4090          JC_Rec_AddRec frmSub  ' ** Function: Above.
              '.AddRec  ' ** Form Procedure: frmJournal_Columns_Sub.
4100        Case False
              ' ** Nothing
4110        End Select

4120        .transdate.SetFocus

4130        .RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.

4140      Case False
            ' ** Either it failed, or just, No, Don't Commit.
4150        Select Case gblnMessage
            Case True
              ' ** Add a new record.
4160          JC_Rec_AddRec frmSub  ' ** Function: Above.
              '.AddRec  ' ** Form Procedure: frmJournal_Columns_Sub.
4170        Case False
              ' ** Nothing
4180          lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
4190          .transdate.SetFocus
4200        End Select
4210      End Select

4220    End With

EXITP:
4230    Exit Sub

ERRH:
4240    DoCmd.Hourglass False
4250    Select Case ERR.Number
        Case Else
4260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4270    End Select
4280    Resume EXITP

End Sub

Public Sub JC_Rec_CostAdjRec(frmSub As Access.Form)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Rec_CostAdjRec"

        Dim strDocName As String, strCallingForm As String, strSaveMoveCtl As String
        Dim lngJrnlColID As Long, lngNewJrnlColID As Long, datPostingDate As Date
        Dim blnNextRec As Boolean, blnFromZero As Boolean, blnNoMove As Boolean, blnGTR_NoAdd As Boolean
        Dim lngTmp01 As Long, strTmp02 As String, strTmp03 As String, lngTmp04 As Long, lngTmp06 As Long, strTmp05 As String
        Dim strTmp07 As String, lngTmp08 As Long, lngTmp09 As Long, strTmp10 As String, lngTmp11 As Long, blnTmp12 As Boolean
        Dim lngX As Long
        Dim lngRetVal As Long

4310    With frmSub

4320      gblnCrtRpt_Zero = False  ' ** Borrowing this variable from Court Reports.
4330      gblnMessage = False  ' ** Borrowing this variable from wherever.
4340      strCallingForm = .Parent.Name
4350      strDocName = "frmJournal_Columns_Add"
4360      DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm

4370      blnGTR_NoAdd = .GTR_NoAdd_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.
4380      datPostingDate = .PostDate_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.
4390      strSaveMoveCtl = .SaveMoveCtl_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.
4400      blnNoMove = .NoMove_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.
4410      lngNewJrnlColID = .NewJColID_GetSet(True)  ' ** Form Procedure: frmJournal_Columns_Sub.

4420      Select Case gblnCrtRpt_Zero  ' ** Indicates whether to commit.
          Case True
            ' ** Yes, commit.

4430        DoCmd.Hourglass True
4440        DoEvents

4450        lngJrnlColID = .JrnlCol_ID  ' ** This is record 1 of the Cost Adj. series.
4460        lngTmp01 = .journaltype_sortorder
4470        strTmp02 = .shortname
4480        strTmp03 = .assetno_description
4490        lngTmp04 = .Location_ID
4500        strTmp05 = .Loc_Name
4510        lngTmp06 = .revcode_ID
4520        strTmp07 = Nz(.revcode_DESC, "Null")
4530        lngTmp08 = .revcode_TYPE
4540        lngTmp09 = .taxcode
4550        strTmp10 = Nz(.taxcode_description, "Null")
4560        lngTmp11 = Nz(.taxcode_type, 0)
4570        blnTmp12 = False  ' ** Add new, empty record.

4580        lngNewRecs = 0&
4590        ReDim arr_varNewRec(N_ELEMS, 0)

            ' ** Add the 1st record to the array.
4600        .NewRecAdd lngJrnlColID  ' ** Form Procedure: frmJournal_Columns_Sub.

4610        DistributeCost frmSub, lngTmp01, strTmp02, strTmp03, lngTmp04, strTmp05, lngTmp06, strTmp07, lngTmp08, lngTmp09, strTmp10, lngTmp11  ' ** Module Procedure: modPurchaseSold.

4620        DoEvents
4630        If lngNewRecs > 0& Then
4640          For lngX = 0& To (lngNewRecs - 1&)
4650            .MoveRec 0, arr_varNewRec(N_ID, lngX)  ' ** Form Procedure: frmJournal_Columns_Sub.
4660            DoEvents
4670            If lngX = (lngNewRecs - 1&) Then
4680              blnTmp12 = True  ' ** Only add new, empty record after last one.
4690            End If
4700            blnNextRec = .NextRec_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
4710            blnFromZero = .FromZero_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
4720            CommitRec frmSub, blnNextRec, blnFromZero, True, blnTmp12  ' ** Module Function: modJrnlCol_Recs.
4730            DoEvents
4740            .NextRec_GetSet False, blnNextRec  ' ** Form Function: frmJournal_Columns_Sub.
4750            .FromZero_GetSet False, blnFromZero  ' ** Form Function: frmJournal_Columns_Sub.
4760          Next
4770        End If

4780        lngNewRecs = 0&
4790        ReDim arr_varNewRec(N_ELEMS, 0)

4800        If lngNewJrnlColID > lngJrnlColID Then
4810          lngTmp01 = lngNewJrnlColID
4820        Else
4830          lngTmp01 = lngJrnlColID
4840        End If
4850        .MoveRec 0, lngTmp01  ' ** Form Procedure: frmJournal_Columns_Sub.
4860        lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.

4870        DoCmd.Hourglass False

4880        Select Case gblnMessage
            Case True
              ' ** Add a new record.
4890          AddRec frmSub, blnGTR_NoAdd, datPostingDate, strSaveMoveCtl, blnNoMove, lngNewJrnlColID  ' ** Module Procedure: modJrnlCol_Recs.
4900        Case False
              ' ** Nothing
4910        End Select

4920        .transdate.SetFocus

4930        .RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.

4940      Case False
            ' ** Either it failed, or just, No, Don't Commit.
4950        Select Case gblnMessage
            Case True
              ' ** Add a new record.
4960          AddRec frmSub, blnGTR_NoAdd, datPostingDate, strSaveMoveCtl, blnNoMove, lngNewJrnlColID  ' ** Module Procedure: modJrnlCol_Recs.
4970        Case False
              ' ** Nothing
4980          lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
4990          .transdate.SetFocus
5000        End Select
5010      End Select

5020      .GTR_NoAdd_GetSet False, blnGTR_NoAdd  ' ** Form Procedure: frmJournal_Columns_Sub.
5030      .PostDate_GetSet False, datPostingDate  ' ** Form Procedure: frmJournal_Columns_Sub.
5040      .SaveMoveCtl_GetSet False, strSaveMoveCtl  ' ** Form Procedure: frmJournal_Columns_Sub.
5050      .NoMove_GetSet False, blnNoMove  ' ** Form Procedure: frmJournal_Columns_Sub.
5060      .NewJColID_GetSet False, lngNewJrnlColID  ' ** Form Procedure: frmJournal_Columns_Sub.

5070    End With

EXITP:
5080    Exit Sub

ERRH:
5090    DoCmd.Hourglass False
5100    Select Case ERR.Number
        Case Else
5110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5120    End Select
5130    Resume EXITP

End Sub

Public Function CommitAll(blnSilent As Boolean, frm As Access.Form) As Boolean

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "CommitAll"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long, lngCommitted As Long, lngFails As Long
        Dim lngUncoms As Long, arr_varUnCom() As Variant
        Dim blnNextRec As Boolean, blnFromZero As Boolean
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varUnCom().
        Const U_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const U_ID   As Integer = 0
        Const U_JTYP As Integer = 1
        Const U_CMTD As Integer = 2

5210    blnRetVal = True

5220    With frm

5230      blnNextRec = .NextRec_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
5240      blnFromZero = .FromZero_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.

5250      DoCmd.Hourglass True
5260      DoEvents

5270      lngRecsCur = frm.RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
5280      lngCommitted = .RecsTot_Committed
5290      If lngRecsCur > lngCommitted Then

5300        lngUncoms = 0&
5310        ReDim arr_varUnCom(U_ELEMS, 0)

5320        Set rst = .RecordsetClone
5330        With rst
5340          .MoveLast
5350          lngRecs = .RecordCount
5360          .MoveFirst
5370          For lngX = 1& To lngRecs
5380            If ![posted] = False Then
5390              If IsNull(![accountno]) = False And IsNull(![journaltype]) = False Then
5400                lngUncoms = lngUncoms + 1&
5410                lngE = lngUncoms - 1&
5420                ReDim Preserve arr_varUnCom(U_ELEMS, lngE)
5430                arr_varUnCom(U_ID, lngE) = ![JrnlCol_ID]
5440                arr_varUnCom(U_JTYP, lngE) = ![journaltype]
5450                arr_varUnCom(U_CMTD, lngE) = CBool(False)
5460              End If
5470            End If
5480            If lngX < lngRecs Then .MoveNext
5490          Next
5500          .Close
5510        End With

5520        If lngUncoms > 0& Then

5530          Set dbs = CurrentDb
5540          For lngX = 0& To (lngUncoms - 1&)
5550            frm.MoveRec 0, arr_varUnCom(U_ID, lngX)  ' ** Form Procedure: frmJournal_Columns_Sub.
                ' ** Update tblJournal_Column, by specified [jcolid], [pstd].
5560            Set qdf = dbs.QueryDefs("qryJournal_Columns_55_01")
5570            With qdf.Parameters
5580              ![jcolid] = arr_varUnCom(U_ID, lngX)
5590              ![pstd] = True
5600            End With
5610            qdf.Execute dbFailOnError
5620            Set qdf = Nothing
5630            arr_varUnCom(U_CMTD, lngX) = CBool(True)
5640            frm.Refresh
5650            DoEvents
                '########################################
                '########################################
                'DOESN'T WORK!!
                '########################################
                '########################################
5660            arr_varUnCom(U_CMTD, lngX) = CommitRec(frm, blnNextRec, blnFromZero, True)  ' ** Function: Below.
5670          Next
5680          dbs.Close
5690          Set dbs = Nothing

5700          lngFails = 0&
5710          For lngX = 0& To (lngUncoms - 1&)
5720            If arr_varUnCom(U_CMTD, lngX) = False Then
5730              lngFails = lngFails + 1&
5740            End If
5750          Next

5760          If lngFails = 0& Then
5770            If blnSilent = False Then
5780              DoCmd.Hourglass False
5790              Beep
5800              MsgBox "All uncommitted entries have been committed to the Posting Journal.", _
                    vbInformation + vbOKOnly, "All Entries Successfully Committed"
5810              DoCmd.SelectObject acForm, .Parent.Name, False
5820              .Parent.cmdAdd.SetFocus
5830              DoEvents
5840            End If
5850          Else
5860            blnRetVal = False
5870            If blnSilent = False Then
5880              If lngFails = lngUncoms Then
5890                DoCmd.Hourglass False
5900                Beep
5910                MsgBox "None of the entries could be committed." & vbCrLf & _
                      "Check the entries and try again.", vbInformation + vbOKOnly, "Commit Failed"
5920              Else
5930                DoCmd.Hourglass False
5940                Beep
5950                MsgBox "Some of the entries could not be committed." & vbCrLf & _
                      "Check those remaining uncommitted and try again.", vbInformation + vbOKOnly, "Commit Incomplete"
5960              End If
5970            End If
5980          End If

5990        Else
6000          DoCmd.Hourglass False
6010          MsgBox "There are no uncommited entries.", vbInformation + vbOKOnly, "Nothing To Do"
6020          DoCmd.SelectObject acForm, .Parent.Name, False
6030          .Parent.cmdAdd.SetFocus
6040          DoEvents
6050        End If
6060      Else
6070        DoCmd.Hourglass False
6080        MsgBox "There are no uncommited entries.", vbInformation + vbOKOnly, "Nothing To Do"
6090        DoCmd.SelectObject acForm, .Parent.Name, False
6100        .Parent.cmdAdd.SetFocus
6110        DoEvents
6120      End If

6130      .NextRec_GetSet False, blnNextRec  ' ** Form Function: frmJournal_Columns_Sub.
6140      .FromZero_GetSet False, blnFromZero  ' ** Form Function: frmJournal_Columns_Sub.

6150    End With

6160    DoCmd.Hourglass False

EXITP:
6170    Set rst = Nothing
6180    Set qdf = Nothing
6190    Set dbs = Nothing
6200    CommitAll = blnRetVal
6210    Exit Function

ERRH:
6220    DoCmd.Hourglass False
6230    blnRetVal = False
6240    Select Case ERR.Number
        Case Else
6250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6260    End Select
6270    Resume EXITP

End Function

Public Function CommitRec(frm As Access.Form, blnNextRec As Boolean, blnFromZero As Boolean, Optional varSilent As Variant, Optional varAddNew As Variant) As Boolean
' ** gblnCrtRpt_Zero = Commit  Yes/No
' ** gblnMessage     = New Rec Yes/No

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "CommitRec"

        Dim strDocName As String, strCallingForm As String
        Dim strNext As String
        Dim blnSilent As Boolean
        Dim blnRetVal As Boolean, lngRetVal As Long

6310    blnRetVal = True

6320    With frm

6330      Select Case IsMissing(varSilent)
          Case True
6340        blnSilent = False
6350      Case False
6360        blnSilent = CBool(varSilent)
6370        Select Case IsMissing(varAddNew)
            Case True
6380          gblnMessage = True  ' ** Add new record.
6390        Case False
6400          gblnMessage = CBool(varAddNew)
6410        End Select
6420      End Select

6430      Select Case blnSilent
          Case True
6440        gblnCrtRpt_Zero = True  ' ** Return value: Passed Tests.
6450      Case False
6460        gblnCrtRpt_Zero = False  ' ** Borrowing this variable from Court Reports.
6470        gblnMessage = False  ' ** Borrowing this variable from wherever.
6480        strCallingForm = .Parent.Name
6490        strDocName = "frmJournal_Columns_Add"
6500        DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
6510      End Select

6520      Select Case gblnCrtRpt_Zero  ' ** Indicates whether to commit.
          Case True
6530        lngRetVal = fSetScrollBarPosHZ(frm, 1&)  ' ** Module Function: modScrollBarFuncs.
6540  On Error Resume Next
6550        .posted.SetFocus
            ' ** Not sure why this sometimes errors!
6560  On Error GoTo ERRH
            '########################################
            '########################################
            'DOESN'T WORK!!
            '########################################
            '########################################
6570        .posted = True
6580        .posted_AfterUpdate  ' ** Form Procedure: frmJournal_Columns_Sub.
            ' ** It passed the tests.
6590        Select Case gblnMessage
            Case True
              ' ** Add a new record.
6600          gblnMessage = False
              '.AddRec  ' ** Function: Below.
6610          .AddRec_Send  ' ** Form Procedure: frmJournal_Columns_Sub.
              'A KeyDown() event sends it to CommitRec().
              'CommitRec() sends it to AddRec().
              'AddRec() sets focus to transdate.
              'Then AddRec() sends it cmdSave_Click().
              'cmdSave_Click() should let it return to CommitRec() without moving it.
              'At which point CommitRec() should just return it to the KeyDown() event.
              ' ** When coming from JC_Frm_Map_Return(), somehow the last
              ' ** Map entry gets unposted, and the empty entry disappears!
              ' ** So, I'm no longer telling it to add on the last entry.
6620        Case False
              ' ** If there are more entries, proceed to the next,
              ' ** otherwise exit the subform
6630          lngRecsCur = .RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
6640          lngRetVal = fSetScrollBarPosHZ(frm, 1&)  ' ** Module Function: modScrollBarFuncs.
6650          If .CurrentRecord < lngRecsCur Then
6660            .MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Columns_Sub.
6670            strNext = JC_Key_Sub_Next("posted_AfterUpdate", blnNextRec, blnFromZero, False, "First")  ' ** Module Function: modJrnlCol_Keys.
6680            .FocusHolder.SetFocus
6690            .Controls(strNext).SetFocus
6700            lngRetVal = fSetScrollBarPosHZ(frm, 1&)  ' ** Module Function: modScrollBarFuncs.
6710          Else
                ' ** Regardless, keep the focus in the subform!
6720            DoCmd.SelectObject acForm, .Parent.Name, False
6730            .Parent.frmJournal_Columns_Sub.SetFocus
6740            .FocusHolder.SetFocus
6750            .transdate.SetFocus
6760            lngRetVal = fSetScrollBarPosHZ(frm, 1&)  ' ** Module Function: modScrollBarFuncs.
6770          End If
6780        End Select
6790      Case False
            ' ** Either it failed, or just, No, Don't Commit.
6800        Select Case gblnMessage
            Case True
              ' ** Add a new record.
6810          gblnMessage = False
              '.AddRec  ' ** Form Function: frmJournal_Columns_Sub.
6820          .AddRec_Send  ' ** Form Procedure: frmJournal_Columns_Sub.
6830        Case False
              ' ** Regardless, keep the focus in the subform!
6840          DoCmd.SelectObject acForm, .Parent.Name, False
6850          .Parent.frmJournal_Columns_Sub.SetFocus
6860          .transdate.SetFocus
6870          lngRetVal = fSetScrollBarPosHZ(frm, 1&)  ' ** Module Function: modScrollBarFuncs.
6880        End Select
            'JC_MSC_GTR_Fill  ' ** Module Procedure: modJrnlCol_Misc.
6890      End Select

6900      gblnCrtRpt_Zero = False
6910      gblnMessage = False

6920    End With

EXITP:
6930    CommitRec = blnRetVal
6940    Exit Function

ERRH:
6950    DoCmd.Hourglass False
6960    frm.posted = False
6970    frm.Parent.CommitNoClose_Set False  ' ** Form Procedure: frmJournal_Columns.
6980    blnRetVal = False
6990    Select Case ERR.Number
        Case Else
7000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7010    End Select
7020    Resume EXITP

End Function

Public Sub AddRec(frmSub As Access.Form, blnGTR_NoAdd As Boolean, datPostingDate As Date, strSaveMoveCtl As String, blnNoMove As Boolean, lngNewJrnlColID As Long, Optional varFromZero As Variant)
' ** There's some duplication between here and
' ** Form_Timer() after the JournalType is chosen.

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "AddRec"

        Dim blnNextRec As Boolean, blnFromZero As Boolean
        Dim lngRetVal As Long

7110    With frmSub

7120      blnNextRec = .NextRec_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
7130      blnFromZero = .FromZero_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.

7140      If IsMissing(varFromZero) = False Then
7150        blnFromZero = CBool(varFromZero)
7160      End If

7170      If blnGTR_NoAdd = False Then
7180        If datPostingDate = 0 Then
              ' ** Sometimes this seems to get lost!
7190          JC_Msc_Pub_Reset .Parent, frmSub  ' ** Module Function: modJrnlCol_Misc.
7200        End If

7210        JC_Msc_Memo_Set False, .Parent  ' ** Module Procedure: modJrnlCol_Misc.
7220        .AllowAdditions = True
7230        frmSub.MoveRec acCmdRecordsGoToNew  ' ** Form Procedure: frmJournal_Columns_Sub.
7240        lngRetVal = fSetScrollBarPosHZ(frmSub, 1&)  ' ** Module Function: modScrollBarFuncs.
7250        .transdate.SetFocus
7260        .transdate = datPostingDate
7270        .posted = False
7280        .pershare = 0#
7290        .Reinvested = False
7300        .Location_ID = 1&
7310        .PrintCheck = False
7320        .revcode_ID = REVID_INC
7330        .revcode_TYPE = REVTYP_INC
7340        .revcode_DESC = "Unspecified Income"
7350        .revcode_DESC_display = Null
7360        .taxcode = TAXID_INC
7370        .taxcode_type = TAXTYP_INC
7380        .taxcode_description = "Unspecified Income"
7390        .taxcode_description_display = Null
7400        .journal_USER = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
7410        .JrnlCol_DateModified = Now()
7420        .rate = 0#
7430        .IsAverage = False
7440        DoEvents
            'JC_MSC_GTR_Fill  ' ** Module Procedure: modJrnlCol_Misc.
7450        If blnFromZero = False Then
              ' ** OnOpen, Form_Timer() in frmJournal_Columns calls JC_Key_Sub_Next().
              ' ** It is that call that then initiates this call to AddRec().
              ' ** A call to here shouldn't always regenerate another recursive call to here.
              ' ** It would be nice if blnFromZero were already True by the time it gets here!
              ' ** At this point, the new record's already been created.
7460          strSaveMoveCtl = JC_Key_Sub_Next("posted_AfterUpdate", blnNextRec, blnFromZero, True, "AddRec")  ' ** Module Function: modJrnlCol_Keys.
7470        End If
            ' ** On a completely empty record (save transdate),
            ' ** I want to keep the top buttons active!
7480        strSaveMoveCtl = vbNullString
7490        blnNoMove = True
7500        .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.
7510        lngNewJrnlColID = .JrnlCol_ID
7520        .AllowAdditions = False

7530      End If

7540      .NextRec_GetSet False, blnNextRec  ' ** Form Function: frmJournal_Columns_Sub.
7550      .FromZero_GetSet False, blnFromZero  ' ** Form Function: frmJournal_Columns_Sub.

7560    End With

EXITP:
7570    Exit Sub

ERRH:
7580    Select Case ERR.Number
        Case Else
7590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7600    End Select
7610    Resume EXITP

End Sub

Public Sub DelRec(frmSub As Access.Form, blnGoneToReport As Boolean, blnGoneToReport2 As Boolean, blnRecsTotUpdate As Boolean)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "DelRec"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngJrnlID As Long
        Dim strErrDesc As String, strThisJType As String
        Dim blnContinue As Boolean

7710    With frmSub

7720      blnContinue = True

7730      strThisJType = Nz(.journaltype, vbNullString)
7740      lngJrnlID = Nz(.Journal_ID, 0)
          ' ** If this entry is already in the Journal, delete it there
          ' ** first, in case someone else is currently editing it,
          ' ** in which case we won't delete it here.
7750      If lngJrnlID > 0& Then
7760        Set dbs = CurrentDb
7770        With dbs
              ' ** Delete Journal, by specified [jid].
7780          Set qdf = .QueryDefs("qryJournal_Columns_19")
7790          With qdf.Parameters
7800            ![jid] = lngJrnlID
7810          End With
7820  On Error Resume Next
7830          qdf.Execute dbFailOnError
7840          If ERR.Number <> 0 Then
7850            blnContinue = False
7860            strErrDesc = ERR.description
7870  On Error GoTo ERRH
7880            MsgBox "Trust Accountant was unable to completely delete this entry." & vbCrLf & _
                  "Someone else may be accessing it." & vbCrLf & vbCrLf & _
                  "Try again later." & vbCrLf & vbCrLf & strErrDesc, vbInformation + vbOKOnly, "Delete Failed"
7890          Else
7900  On Error GoTo ERRH
7910          End If
7920          .Close
7930        End With
7940      End If

7950      If blnContinue = True Then

7960        .AllowDeletions = True
7970        If blnGoneToReport = True Then
              ' ** This has become an unholy mess!
7980  On Error Resume Next
7990          DoCmd.RunCommand acCmdDeleteRecord
8000          If ERR.Number <> 0 Then
8010  On Error GoTo ERRH
8020            blnGoneToReport = False
8030            blnGoneToReport2 = True
8040            .TimerInterval = 100&
8050          Else
8060  On Error GoTo ERRH
8070            blnGoneToReport = False
8080          End If
8090        Else
8100          DoCmd.RunCommand acCmdDeleteRecord
8110        End If

8120        If blnGoneToReport = False Then
8130          .AllowDeletions = False
8140          DoEvents
8150          If strThisJType = "Paid" Then
8160            If .Parent.JrnlMemo_Memo.Visible = True Then
8170              JC_Msc_Memo_Set False, .Parent  ' ** Module Procedure: modJrnlCol_Misc.
8180              .Parent.JrnlMemo_Memo = Null
8190            End If
8200          End If
8210          blnRecsTotUpdate = True
8220          .TimerInterval = 3000&
8230        End If

8240      End If  ' ** blnContinue.
8250      gblnDeleting = False

8260    End With

EXITP:
8270    Set qdf = Nothing
8280    Set dbs = Nothing
8290    Exit Sub

ERRH:
8300    Select Case ERR.Number
        Case Else
8310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8320    End Select
8330    Resume EXITP

End Sub
