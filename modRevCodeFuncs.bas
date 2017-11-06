Attribute VB_Name = "modRevCodeFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modRevCodeFuncs"

'VGC 03/23/2017: CHANGES!

Public Const REVID_INC    As Long = 1&
Public Const REVID_EXP    As Long = 2&
Public Const REVTYP_INC   As Long = 1&
Public Const REVTYP_EXP   As Long = 2&
Public Const REVID_OCHG   As Long = 3&  ' ** OC Other Charges.
Public Const REVID_OCRED  As Long = 4&  ' ** OC Other Credits.
Public Const REVID_ORDDIV As Long = 5&  ' ** Ordinary Dividend.
Public Const REVID_INTINC As Long = 6&  ' ** Interest Income.
Public Const TAXID_INC    As Long = 1&
Public Const TAXID_DED    As Long = 2&
Public Const TAXTYP_INC   As Long = 1&
Public Const TAXTYP_DED   As Long = 2&

' ** Array garr_varRevO().
Public glngRevOs As Long, garr_varRevO As Variant
Public Const REVO_ID      As Integer = 0
Public Const REVO_SORTORD As Integer = 3
Public Const REVO_CHANGED As Integer = 5
'Public Const REVO_TYPDESC As Integer = 6

Private Const RC_FRM As String = "frmIncomeExpenseCodes"
Private Const RC_FRM_SUB As String = "frmIncomeExpenseCodes_Sub"
Private Const RC_MOD As String = "modBackendUpdate"
' **

Public Function RevCode_Setup(strCaller As String) As Boolean

100   On Error GoTo ERRH

        Const THIS_PROC As String = "RevCode_Setup"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

110     blnRetVal = True

        ' ** Make tmpRevCodeEdit table.
120     Set dbs = CurrentDb
130     With dbs
140       If TableExists("tmpRevCodeEdit") = False Then  ' ** Module Function: modFileUtilities.
            ' ** Data-Definition: Create table tmpRevCodeEdit.
150         Set qdf = .QueryDefs("qryRevCodes_40_01")
160         qdf.Execute
            ' ** Data-Definition: Create index [revcode_ID] PrimaryKey on table tmpRevCodeEdit.
170         Set qdf = .QueryDefs("qryRevCodes_40_02")
180         qdf.Execute
            ' ** Data-Definition: Create index [revcode_TYPE] on table tmpRevCodeEdit.
190         Set qdf = .QueryDefs("qryRevCodes_40_03")
200         qdf.Execute
210       End If
          ' ** Empty tmpRevCodeEdit.
220       Set qdf = .QueryDefs("qryRevCodes_12")
230       qdf.Execute
          ' ** Append qryRevCodes_01 (m_REVCODE, linked to m_REVCODE_TYPE, with add'l fields) to tmpRevCodeEdit.
240       Set qdf = dbs.QueryDefs("qryRevCodes_01a")
250       qdf.Execute
          ' ** tmpRevCodeEdit, with add'l fields.
260       Set qdf = .QueryDefs("qryRevCodes_02")
270       Set rst = qdf.OpenRecordset
          ' ** Load the garr_varRevO() array.
280       RevO_Load rst  ' ** Procedure: Below.
290       rst.Close
300       .Close
310     End With

EXITP:
320     Set rst = Nothing
330     Set qdf = Nothing
340     Set dbs = Nothing
350     RevCode_Setup = blnRetVal
360     Exit Function

ERRH:
370     blnRetVal = False
380     Select Case ERR.Number
        Case Else
390       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
400     End Select
410     Resume EXITP

End Function

Public Function RevCode_Renum(strCaller As String, ByRef blnEdited As Boolean, ByRef blnMoveAll As Boolean) As Boolean

500   On Error GoTo ERRH

        Const THIS_PROC As String = "RevCode_Renum"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRevCodes As Long, arr_varRevCode As Variant
        Dim lngLastOrd As Long, blnMoveThis As Boolean, lngMoves As Long
        Dim lngRecs As Long, intOneExp As Integer, intOneInc As Integer
        Dim blnSortChanged As Boolean
        Dim strMsg As String
        Dim lngX As Long, lngY As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varRevCode().
        Const R_RID    As Integer = 0
        'Const R_DSC    As Integer = 1
        'Const R_TYP    As Integer = 2
        Const R_ORD    As Integer = 3
        'Const R_CHGD   As Integer = 4
        'Const R_TYPDSC As Integer = 5
        Const R_ORDN   As Integer = 6
        'Const R_SORTI  As Integer = 7
        'Const R_SORTE  As Integer = 8

510     blnRetVal = True

520     Set dbs = CurrentDb

        ' ** If they've made changes to revcode_SORTORDER, and made sure the numbers were unique and continuous,
        ' ** then this renumbering won't do anything, but saying 'no changes were necessary' is misleading.
530     blnSortChanged = Forms("frmIncomeExpenseCodes").frmIncomeExpenseCodes_Sub.Form.SortOrdChanged  ' ** Form Function: frmIncomeExpenseCodes_Sub.

        ' ** Check again for integrity of Sort Order 1.
540     intOneInc = 0: intOneExp = 0
        ' ** tmpRevCodeEdit, with add'l fields.
550     Set qdf = dbs.QueryDefs("qryRevCodes_02")
560     Set rst = qdf.OpenRecordset
570     With rst
580       .MoveLast
590       lngRecs = .RecordCount
600       .MoveFirst
610       For lngX = 1& To lngRecs
620         Select Case ![revcode_TYPE]
            Case REVTYP_INC
630           If ![revcode_SORTORDER] = 1 Then intOneInc = intOneInc + 1
640         Case REVTYP_EXP
650           If ![revcode_SORTORDER] = 1 Then intOneExp = intOneExp + 1
660         End Select
670         If lngX < lngRecs Then .MoveNext
680       Next
690       .Close
700     End With

710     If intOneInc <> 1 Or intOneExp <> 1 Then
720       strMsg = vbNullString
730       If intOneInc = 0 Then
740         strMsg = "Sort Order 1, Unspecified Income, cannot be found."
750       ElseIf intOneInc > 1 Then
760         strMsg = "Sort Order 1 is locked. You cannot assign another Income code to 1."
770       End If
780       If intOneExp = 0 Then
790         strMsg = strMsg & " Sort Order 1, Unspecified Expense, cannot be found."
800       ElseIf intOneExp > 1 Then
810         strMsg = strMsg & " Sort Order 1 is locked. You cannot assign another Expense code to 1."
820       End If
830       strMsg = Trim(strMsg)
840       Select Case strCaller
          Case RC_FRM, RC_FRM_SUB
850         DoCmd.Hourglass False
860         MsgBox strMsg, vbInformation + vbOKOnly, "Invalid Entry"
870       Case RC_MOD
            'HOW SHOULD I HANDLE THIS ON STARTUP?

880       End Select
890     Else

900       lngLastOrd = 0&
910       blnMoveAll = False
920       lngMoves = 0&

930       With dbs

940         For lngX = 1& To 2&

950           lngRevCodes = 0&

              ' ** tmpRevCodeEdit, by specified [revtyp].
960           Set qdf = .QueryDefs("qryRevCodes_30")  ' ** Sorted by revcode_SORTORDER, revcode_CHANGED.
970           With qdf.Parameters
980             ![revtyp] = lngX
990           End With
1000          Set rst = qdf.OpenRecordset
1010          With rst
1020            .MoveLast
1030            lngRevCodes = .RecordCount
1040            .MoveFirst
1050            arr_varRevCode = .GetRows(lngRevCodes)
                ' *****************************************************************
                ' ** Array: arr_varRevCode()
                ' **
                ' **   Field  Element  Name (Default)              Constant
                ' **   =====  =======  ==========================  ==============
                ' **     1       0     revcode_ID                  R_RID
                ' **     2       1     revcode_DESC                R_DSC
                ' **     3       2     revcode_TYPE                R_TYP
                ' **     4       3     revcode_SORTORDER           R_ORD
                ' **     5       4     revcode_CHANGED             R_CHGD
                ' **     6       5     revcode_TYPE_Description    R_TYPDSC
                ' **     7       6     SortOrdNew (0)              R_ORDN
                ' **     8       7     revcode_SORTORDER_I         R_SORTI
                ' **     9       8     revcode_SORTORDER_E         R_SORTE
                ' **
                ' *****************************************************************
1060            .Close
1070          End With

1080          lngLastOrd = 0&
1090          blnMoveThis = False
1100          For lngY = 0& To (lngRevCodes - 1&)
1110            If arr_varRevCode(R_ORD, lngY) <> (lngLastOrd + 1&) Then
1120              arr_varRevCode(R_ORDN, lngY) = (lngLastOrd + 1&)
1130              blnMoveAll = True
1140              blnMoveThis = True
1150              lngMoves = lngMoves + 1&
1160            Else
1170              arr_varRevCode(R_ORDN, lngY) = (lngLastOrd + 1&)
1180            End If
1190            lngLastOrd = (lngLastOrd + 1&)
1200          Next

1210          If blnMoveThis = True Then

                ' ** Update qryRevCodes_30 to revcode_SORTORD * 100 to get them all out of the way.
1220            Set qdf = .QueryDefs("qryRevCodes_31")
1230            With qdf.Parameters
1240              ![revtyp] = lngX
1250            End With
1260            qdf.Execute
1270            For lngY = 0& To (lngRevCodes - 1&)
1280              arr_varRevCode(R_ORD, lngY) = arr_varRevCode(R_ORD, lngY) * 100&
1290            Next

1300            Set rst = .OpenRecordset("tmpRevCodeEdit", dbOpenDynaset, dbConsistent)
1310            With rst
1320              For lngY = 0& To (lngRevCodes - 1&)
1330                .FindFirst "[revcode_ID] = " & CStr(arr_varRevCode(R_RID, lngY))
1340                If .NoMatch = False Then
1350                  .Edit
1360                  ![revcode_SORTORDER] = arr_varRevCode(R_ORDN, lngY)
1370                  .Update
1380                Else
                      ' ** Shouldn't happen!
1390                  blnRetVal = False
1400                  DoCmd.Hourglass False
1410                  MsgBox "Problem with tmpRevCodeEdit." & vbCrLf & vbCrLf & _
                        "Module: " & THIS_NAME & vbCrLf & _
                        "Function: " & THIS_PROC, vbCritical + vbOKOnly, "Error"
1420                  Exit For
1430                End If
1440              Next
1450              .Close
1460            End With

1470          End If

1480        Next

            ' ** Update qryRevCodes_32 (tmpRevCodeEdit, with revcode_SORTORDER_I_new, revcode_SORTORDER_E_new).
1490        Set qdf = .QueryDefs("qryRevCodes_33")
1500        qdf.Execute

1510      End With  ' ** dbs.

1520      If strCaller = RC_FRM Or strCaller = RC_FRM_SUB Then
1530        Forms(strCaller).Requery
1540        DoEvents
1550      End If

          ' ** Load the garr_varRevCodeO() array.
1560      If blnRetVal = True Then
            ' ** tmpRevCodeEdit, with add'l fields.
1570        Set qdf = dbs.QueryDefs("qryRevCodes_02")
1580        Set rst = qdf.OpenRecordset
1590        RevO_Load rst  ' ** Procedure: Below.
1600      End If

1610      dbs.Close

1620      If blnMoveAll = True Then
1630        blnEdited = True
1640        If strCaller = RC_FRM Or strCaller = RC_FRM_SUB Then
1650          DoCmd.Hourglass False
1660          MsgBox "Renumbering successful." & vbCrLf & CStr(lngMoves) & " code" & _
                IIf(lngMoves = 1, vbNullString, "s") & " renumbered." & vbCrLf & vbCrLf & _
                "Click Update to save the changes!", vbInformation + vbOKOnly, "Renumbering Successful"
1670        End If
1680      Else
1690        If strCaller = RC_FRM Or strCaller = RC_FRM_SUB Then
1700          DoCmd.Hourglass False
1710          Select Case blnSortChanged
              Case True
1720            MsgBox "Sort orders are unique and continuous.", vbInformation + vbOKOnly, "Renumbering Successful"
1730          Case False
1740            MsgBox "No changes necessary.", vbExclamation + vbOKOnly, "Renumbering Unnecessary"
1750          End Select
1760        End If
1770      End If

1780    End If

EXITP:
1790    Set rst = Nothing
1800    Set qdf = Nothing
1810    Set dbs = Nothing
1820    RevCode_Renum = blnRetVal
1830    Exit Function

ERRH:
1840    DoCmd.Hourglass False
1850    blnRetVal = False
1860    Select Case ERR.Number
        Case Else
1870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1880    End Select
1890    Resume EXITP

End Function

Public Function RevCode_Update(strCaller As String, ByRef blnEdited As Boolean, Optional varWrk As Variant) As Boolean
' ** Note: All this opening and closing is because the transaction just gave me nothing but trouble!

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "RevCode_Update"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, rstFrom As DAO.Recordset, rstTo As DAO.Recordset, fld As DAO.Field
        Dim qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, qdf3 As DAO.QueryDef, tdf As DAO.TableDef
        Dim blnInTrans As Boolean
        Dim blnContinue As Boolean
        Dim lngRecs As Long
        Dim lngX As Long, lngY As Long
        Dim blnRetVal As Boolean

1910    blnRetVal = True

1920    blnContinue = True  ' ** Unless proven otherwise.
1930    blnInTrans = False  ' ** Until we get there.

1940    If strCaller = RC_FRM Or strCaller = RC_FRM_SUB Then
1950      Forms(strCaller).Refresh
1960    End If

1970    Set dbs = CurrentDb
1980    With dbs
          ' ** Make sure revcode_SORTORDER reflects values in revcode_SORTORDER_I or revcode_SORTORDER_E.
          ' ** Update qryRevCodes_34 (tmpRevCodeEdit, with revcode_SORTORDER_new).
1990      Set qdf1 = dbs.QueryDefs("qryRevCodes_35")
2000      qdf1.Execute
2010      .Close
2020    End With

2030    If RevCode_Validate(strCaller) = True Then  ' ** Function: Below.

2040      Set dbs = CurrentDb
          ' ** m_REVCODE, sorted by revcode_ID.
2050      Set qdf1 = dbs.QueryDefs("qryRevCodes_08")
2060      Set rstTo = qdf1.OpenRecordset()
2070      If rstTo.BOF = True And rstTo.EOF = True Then
            ' ** But there really would have to be at least one.
2080        blnContinue = False
2090        If strCaller = RC_FRM Or strCaller = RC_FRM_SUB Then
2100          MsgBox "No Income/Expense codes were found!", vbCritical + vbOKOnly, "No Codes Found"
2110        End If
2120        rstTo.Close
2130        dbs.Close
2140      Else
2150        rstTo.Close
2160      End If

2170      If blnContinue = True Then
            ' ** tmpRevCodeEdit, sorted by revcode_ID.
2180        Set qdf2 = dbs.QueryDefs("qryRevCodes_09")
2190        Set rstFrom = qdf2.OpenRecordset
2200        If rstFrom.BOF = True And rstFrom.EOF = True Then
              ' ** Again, there really would have to be at least one.
2210          blnContinue = False
2220          If strCaller = RC_FRM Or strCaller = RC_FRM_SUB Then
2230            MsgBox "No Income/Expense codes were found!", vbCritical + vbOKOnly, "No Codes Found"
2240          End If
2250        Else
2260        End If
2270        rstFrom.Close
2280        dbs.Close
2290      End If

2300      If blnContinue = True Then

            ' ** Remove the '>0 And <99' ValidationRule.
2310        If IsMissing(varWrk) = True Then
2320  On Error Resume Next
2330             ' ** New.
2340          If ERR.Number <> 0 Then
2350  On Error GoTo ERRH
2360  On Error Resume Next
2370            Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
2380            If ERR.Number <> 0 Then
2390  On Error GoTo ERRH
2400  On Error Resume Next
2410              Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
2420              If ERR.Number <> 0 Then
2430  On Error GoTo ERRH
2440  On Error Resume Next
2450                Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
2460                If ERR.Number <> 0 Then
2470  On Error GoTo ERRH
2480  On Error Resume Next
2490                  Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
2500                  If ERR.Number <> 0 Then
2510  On Error GoTo ERRH
2520  On Error Resume Next
2530                    Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
2540                    If ERR.Number <> 0 Then
2550  On Error GoTo ERRH
2560  On Error Resume Next
2570                      Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
2580  On Error GoTo ERRH
2590                    Else
2600  On Error GoTo ERRH
2610                    End If
2620                  Else
2630  On Error GoTo ERRH
2640                  End If
2650                Else
2660  On Error GoTo ERRH
2670                End If
2680              Else
2690  On Error GoTo ERRH
2700              End If
2710            Else
2720  On Error GoTo ERRH
2730            End If
2740          Else
2750  On Error GoTo ERRH
2760          End If
2770          Set dbs = wrk.OpenDatabase((gstrTrustDataLocation & gstrFile_DataName), False, False)  ' ** {pathfile}, {exclusive}, {read-only}
2780        Else
2790          Set dbs = varWrk.OpenDatabase((gstrTrustDataLocation & gstrFile_DataName), False, False)  ' ** {pathfile}, {exclusive}, {read-only}
2800        End If
2810        Set tdf = dbs.TableDefs![m_REVCODE]
2820        Set fld = tdf.Fields![revcode_SORTORDER]
2830        fld.ValidationRule = vbNullString
2840        dbs.Close
2850        If IsMissing(varWrk) = True Then
2860          wrk.Close
2870        End If

2880        DoCmd.DeleteObject acTable, "m_REVCODE"
2890        DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), acTable, "m_REVCODE", "m_REVCODE"
2900        CurrentDb.TableDefs.Refresh
2910        Set fld = Nothing
2920        Set tdf = Nothing
2930        Set dbs = Nothing
2940        Set wrk = Nothing

2950        Set wrk = DBEngine.Workspaces(0)

2960        wrk.BeginTrans
2970        blnInTrans = True

2980        Set dbs = wrk.Databases(0)

            ' ** m_REVCODE, sorted by revcode_ID.
2990        Set qdf1 = dbs.QueryDefs("qryRevCodes_08")
3000        Set rstTo = qdf1.OpenRecordset

            ' ** tmpRevCodeEdit, sorted by revcode_ID.
3010        Set qdf2 = dbs.QueryDefs("qryRevCodes_09")
3020        Set rstFrom = qdf2.OpenRecordset

            ' ** First, process Deletes.
3030        rstTo.MoveLast
3040        lngRecs = rstTo.RecordCount
3050        For lngX = lngRecs To 1& Step -1&
3060          If rstTo![revcode_SORTORDER] > 1 Then  ' ** Don't update sortorders #1.
3070            rstFrom.FindFirst "revcode_ID = " & Trim(CStr(rstTo![revcode_ID]))
3080            If rstFrom.NoMatch Then  ' ** ID not found in edited data.
3090              rstTo.Delete
3100            End If
3110          End If
3120          If lngX > 1& Then rstTo.MovePrevious
3130        Next
3140        wrk.CommitTrans
3150        rstTo.Close
3160        rstFrom.Close
3170        dbs.Close
3180        wrk.BeginTrans

3190        Set dbs = wrk.Databases(0)

            ' ** Bias sortorders by +/- 100 in case they are swapping 2 or more sortorders.
            ' ** Update qryRevCodes_10a (m_REVCODE, with revcode_SORTORDER_new = (revcode_SORTORDER * 100)).
3200        Set qdf3 = dbs.QueryDefs("qryRevCodes_10")
3210        qdf3.Execute dbFailOnError

            ' ** tmpRevCodeEdit, sorted by revcode_ID.
3220        Set qdf2 = dbs.QueryDefs("qryRevCodes_09")
3230        Set rstFrom = qdf2.OpenRecordset

            ' ** Now, process changes and additions.
            ' ** m_REVCODE, sorted by revcode_ID.
3240        Set qdf3 = dbs.QueryDefs("qryRevCodes_11")
3250        Set rstTo = qdf3.OpenRecordset
3260        rstFrom.MoveLast
3270        lngRecs = rstFrom.RecordCount
3280        rstFrom.MoveFirst
3290        For lngX = 1& To lngRecs
3300          If IsNull(rstFrom![revcode_ID]) Then  ' ** New record.
3310            rstTo.AddNew
3320          Else
3330            rstTo.FindFirst "revcode_ID = " & Trim(str(rstFrom![revcode_ID]))
3340            If rstTo.NoMatch Then
3350              rstTo.AddNew  ' ** Though this really shouldn't happen because it would have been added just above.
3360            Else            ' ** NEW: Before_Update() now gives it a revcode_ID!
3370              rstTo.Edit
3380            End If
3390          End If
3400          lngY = 0&
3410          For Each fld In rstTo.Fields
3420            With fld
3430              If .Name <> "revcode_ID" And .Name <> "revcode_CHANGED" And _
                      .Name <> "revcode_TYPE_Description" Then  ' ** revcode_ID is an AutoNumber field, the other two are dummies.
3440                .Value = rstFrom.Fields(.Name)
3450              End If
3460            End With
3470          Next
3480          rstTo.Update
3490          If lngX < lngRecs Then rstFrom.MoveNext
3500        Next

3510        wrk.CommitTrans
3520        blnInTrans = False
3530        rstTo.Close
3540        rstFrom.Close
3550        dbs.Close
3560        wrk.Close

            ' ** Restore the '>0 And <99' ValidationRule.
3570        If IsMissing(varWrk) = True Then
3580  On Error Resume Next
3590          Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
3600          If ERR.Number <> 0 Then
3610  On Error GoTo ERRH
3620  On Error Resume Next
3630            Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
3640            If ERR.Number <> 0 Then
3650  On Error GoTo ERRH
3660  On Error Resume Next
3670              Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
3680              If ERR.Number <> 0 Then
3690  On Error GoTo ERRH
3700  On Error Resume Next
3710                Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
3720                If ERR.Number <> 0 Then
3730  On Error GoTo ERRH
3740  On Error Resume Next
3750                  Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
3760                  If ERR.Number <> 0 Then
3770  On Error GoTo ERRH
3780  On Error Resume Next
3790                    Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
3800                    If ERR.Number <> 0 Then
3810  On Error GoTo ERRH
3820  On Error Resume Next
3830                      Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
3840  On Error GoTo ERRH
3850                    Else
3860  On Error GoTo ERRH
3870                    End If
3880                  Else
3890  On Error GoTo ERRH
3900                  End If
3910                Else
3920  On Error GoTo ERRH
3930                End If
3940              Else
3950  On Error GoTo ERRH
3960              End If
3970            Else
3980  On Error GoTo ERRH
3990            End If
4000          Else
4010  On Error GoTo ERRH
4020          End If
4030          Set dbs = wrk.OpenDatabase((gstrTrustDataLocation & gstrFile_DataName), False, False)  ' ** {pathfile}, {exclusive}, {read-only}
4040        Else
4050          Set dbs = varWrk.OpenDatabase((gstrTrustDataLocation & gstrFile_DataName), False, False)  ' ** {pathfile}, {exclusive}, {read-only}
4060        End If
4070        Set tdf = dbs.TableDefs![m_REVCODE]
4080        Set fld = tdf.Fields![revcode_SORTORDER]
4090        fld.ValidationRule = ">0 And <99"
4100        fld.ValidationText = "Sort Order must be 1 - 98." ' ** In case it wasn't there to begin with.
4110        dbs.Close
4120        If IsMissing(varWrk) = True Then
4130          wrk.Close
4140        End If
4150        DoCmd.DeleteObject acTable, "m_REVCODE"
4160        DoCmd.TransferDatabase acLink, "Microsoft Access", (gstrTrustDataLocation & gstrFile_DataName), acTable, "m_REVCODE", "m_REVCODE"
4170        CurrentDb.TableDefs.Refresh

4180        If strCaller = RC_FRM Or strCaller = RC_FRM_SUB Then
4190          Forms(strCaller).Requery
4200          Forms(strCaller).cmdClose.SetFocus
4210          MsgBox "Updates completed.", vbInformation + vbOKOnly, ("Update Successful" & Space(40))
4220        End If
            ' ** All done saving, so nothing edited now.
4230        blnEdited = False
4240        RevCode_SetEdited strCaller, False  ' ** Procedure: Below.

4250      End If

          ' ** Load the garr_varRevO() array.
4260      Set dbs = CurrentDb
4270      With dbs
            ' ** tmpRevCodeEdit, with add'l fields.
4280        Set qdf1 = dbs.QueryDefs("qryRevCodes_02")
4290        Set rstFrom = qdf1.OpenRecordset
4300        RevO_Load rstFrom  ' ** Procedure: Below.
4310        rstFrom.Close
4320        .Close
4330      End With

4340    End If

EXITP:
4350    Set fld = Nothing
4360    Set rstFrom = Nothing
4370    Set rstTo = Nothing
4380    Set qdf1 = Nothing
4390    Set qdf2 = Nothing
4400    Set qdf3 = Nothing
4410    Set tdf = Nothing
4420    Set dbs = Nothing
4430    Set wrk = Nothing
4440    RevCode_Update = blnRetVal
4450    Exit Function

ERRH:
4460    blnRetVal = False
4470    If blnInTrans Then
4480      wrk.Rollback
4490      rstTo.Close
4500      rstFrom.Close
4510      dbs.Close
4520      wrk.Close
4530    End If
4540    Select Case ERR.Number
        Case Else
4550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4560    End Select
4570    Resume EXITP

End Function

Public Function RevCode_Validate(strCaller As String) As Boolean

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "RevCode_Validate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

4610    blnRetVal = True

4620    Set dbs = CurrentDb
        ' ** tmpRevCodeEdit, for revcode_SORTORDER dupes within Income/Expense groups.
4630    Set qdf = dbs.QueryDefs("qryRevCodes_03")
4640    Set rst = qdf.OpenRecordset
4650    With rst
4660      .MoveFirst
4670      If ![NumRecs] > 1 Then
4680        blnRetVal = False
4690        MsgBox "Sort Order must be unique within the Type.", vbExclamation + vbOKOnly, "Invalid Entry"
4700      Else
            ' ** No dupes within Income/Expense groups.
4710      End If
4720      .Close
4730    End With

4740    If blnRetVal = True Then
          ' ** tmpRevCodeEdit, for 0 or Null revcode_SORTORDER.
4750      Set qdf = dbs.QueryDefs("qryRevCodes_04")
4760      Set rst = qdf.OpenRecordset
4770      With rst
4780        If .BOF = True And .EOF = True Then
              ' ** No Null or 0 revcode_SORTORDER.
4790        Else
4800          blnRetVal = False
4810          MsgBox "All revenue codes must have a Sort Order greater than 2.", vbExclamation + vbOKOnly, "Invalid Entry"
4820        End If
4830        .Close
4840      End With
4850    End If

4860    If blnRetVal = True Then
          ' ** tmpRevCodeEdit, for Null revcode_DESC.
4870      Set qdf = dbs.QueryDefs("qryRevCodes_05")
4880      Set rst = qdf.OpenRecordset
4890      With rst
4900        If .BOF = True And .EOF = True Then
              ' ** No Null or NullString revcode_DESC.
4910        Else
4920          blnRetVal = False
4930          MsgBox "All revenue codes must have a description.", vbExclamation + vbOKOnly, "Invalid Entry"
4940        End If
4950        .Close
4960      End With
4970    End If

4980    If blnRetVal = True Then
          ' ** tmpRevCodeEdit, for revcode_SORTORDER outside 3-98.
4990      Set qdf = dbs.QueryDefs("qryRevCodes_06")
5000      Set rst = qdf.OpenRecordset
5010      With rst
5020        If .BOF = True And .EOF = True Then
              ' ** No revcode_SORTORDER < 3 or > 98 (because 1 and 2 are locked, and 99 is reserved).
5030        Else
5040          blnRetVal = False
5050          MsgBox "Codes must have a Sort Order between 2 and 99.", vbExclamation + vbOKOnly, "Invalid Entry"
5060        End If
5070        .Close
5080      End With
5090    End If

EXITP:
5100    Set rst = Nothing
5110    Set qdf = Nothing
5120    Set dbs = Nothing
5130    RevCode_Validate = blnRetVal
5140    Exit Function

ERRH:
5150    blnRetVal = False
5160    Select Case ERR.Number
        Case Else
5170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5180    End Select
5190    Resume EXITP

End Function

Public Sub RevCode_SetEdited(strCaller As String, blnEditRequest As Boolean)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "RevCode_SetEdited"

5210    If strCaller = RC_FRM Then
5220      Forms(strCaller).cmdUpdate.Enabled = blnEditRequest
5230      Forms(strCaller).cmdRenumber.Enabled = blnEditRequest
5240    ElseIf strCaller = RC_FRM_SUB Then
5250      Forms(RC_FRM).cmdUpdate.Enabled = blnEditRequest
5260      Forms(RC_FRM).cmdRenumber.Enabled = blnEditRequest
5270    End If

EXITP:
5280    Exit Sub

ERRH:
5290    Select Case ERR.Number
        Case Else
5300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5310    End Select
5320    Resume EXITP

End Sub

Public Sub RevO_Load(rst As DAO.Recordset)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "RevO_Load"

5410    rst.MoveLast
5420    glngRevOs = rst.RecordCount
5430    rst.MoveFirst
5440    garr_varRevO = rst.GetRows(glngRevOs)
        ' ************************************************************************
        ' ** Array: garr_varRevO()
        ' **
        ' **   Field  Element  Name (Default)                     Constant
        ' **   =====  =======  =================================  ==============
        ' **     1       0     revcode_ID                         REVO_ID
        ' **     2       1     revcode_DESC
        ' **     3       2     revcode_TYPE
        ' **     4       3     revcode_SORTORDER                  REVO_SORTORD
        ' **     5       4     revcode_ACTIVE
        ' **     6       5     revcode_CHANGED                    REVO_CHANGED
        ' **     7       6     revcode_TYPE_Description           REVO_TYPDESC
        ' **     8       7     opgIncomeExpense_optIncome_box
        ' **     9       8     opgIncomeExpense_optExpense_box
        ' **    10       9     Unspecified
        ' **    11      10     revcode_SORTORDER_I
        ' **    12      11     revcode_SORTORDER_E
        ' **
        ' ************************************************************************

EXITP:
5450    Exit Sub

ERRH:
5460    Select Case ERR.Number
        Case Else
5470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5480    End Select
5490    Resume EXITP

End Sub
