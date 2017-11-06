Attribute VB_Name = "modHideTransactions1"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modHideTransactions1"

'VGC 07/01/2017: CHANGES!

' ######################################
' ## KEEP ALL OLD LEDGER HIDDEN STUFF!
' ## INCLUDING LedgerHidden TABLE, AND
' ## ALL THE OLD QUERIES!
' ## (At least until I can separate
' ## the old and new.)
' ######################################

'Make sure entries not matched get:
'GRP_NONE

'AHA!
'I RETAINED THE OLD LedgerHidden TABLE TO SHORTCUT THE
'CONVERSION TO THE NEW tblLedgerHidden! ON SYSTEMS
'ALREADY USING TRUST ACCOUNTANT, IT WILL SPEED SETUP
'TO THE NEW METHOD!

'modHideTransactions1.Hide_Setup() called by:
' modHideTransactions1.LedgerHiddenLoad()
'LedgerHiddenLoad() called by:
' frmAccountHideTrans2.Form_Open()

'modHideTransactions1.Hide_Group() called by
' modQueryFunctions1.FormRef()

'Hide_Type() called by:
' Hide_LoadArray()
'Hide_Max() called by:
' Hide_LoadArray()
' Hide_Setup()
'Hide_LoadArray() called by:
' Hide_Setup()
' Hide_Group()

'Hide_Count() called by:
' NOT CALLED!
'Hide_RenumGroups() called by:
' NOT CALLED!

' ** Progress bar variables.
Private lngTpp As Long  ', strSp As String
Private dblPB_Steps As Double, dblPB_StepSubs As Double
Private dblPB_Width As Double, dblPB_ThisWidth As Double, dblPB_ThisWidthSub As Double
Private dblPB_ThisStep As Double, dblPB_ThisStepSub As Double
Private arr_dblPB_ThisIncr() As Double, dblPB_ThisIncrSub As Double, strPB_ThisPct As String

Private lngGroupMax As Long

' ** Array: arr_varHide().
Private lngHides As Long, arr_varHide() As Variant
Private Const H_ELEMS As Integer = 12  ' ** Array's first-element UBound().
Private Const H_NUM   As Integer = 0
Private Const H_CNT   As Integer = 1
Private Const H_MGRP  As Integer = 2
Private Const H_ACTNO As Integer = 3
Private Const H_JNO   As Integer = 4
Private Const H_JTYPE As Integer = 5
Private Const H_UNIQ  As Integer = 6
Private Const H_UNHID As Integer = 7
Private Const H_GTYPE As Integer = 8
Private Const H_SORT  As Integer = 9
Private Const H_SDATE As Integer = 10
Private Const H_ORD   As Integer = 11
Private Const H_PREX  As Integer = 12
' **

Public Function Hide_Setup(Optional varProgBar As Variant, Optional varFrm As Variant, Optional varPB_Incr As Variant, Optional varAcct As Variant) As Integer
' ** To test queries, and for maintenance, enter this in the Immediate Window:
' **   gstrFormQuerySpec = "frmAccountHideTrans2_Hidden"
' **   gstrAccountNo = "11"
' ** It's needed by FormRef() in modQueryFunctions1, which is
' ** called by numerous queries in the qryAccountHide_00 series.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Hide_Setup"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngPreExisting As Long
        Dim lngRecs As Long
        Dim lngGrp As Long
        Dim blnMiscGroup As Boolean
        Dim lngMultis As Long, arr_varMulti As Variant
        Dim lngSkips As Long
        Dim blnProgBar As Boolean
        Dim dblPB_Incr As Double 'dblPB_ThisStep As Double, dblPB_ThisWidth As Double
        Dim intPos01 As Integer
        Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, lngTmp03 As Long, datTmp04 As Date
        Dim lngX As Long, lngY As Long
        Dim intRetVal As Integer, blnRetVal As Boolean

        'Const M_ACTNO As Integer = 0
        Const M_ASTNO As Integer = 1
        Const M_ICASH As Integer = 2
        Const M_PCASH As Integer = 3
        Const M_COST  As Integer = 4
        Const M_CNT   As Integer = 5
        Const M_SKIP  As Integer = 6

110     intRetVal = 0

120     DoCmd.Hourglass True
130     DoEvents

        ' ** This function has 8 Steps.
140     If IsMissing(varProgBar) = True Then
150       blnProgBar = False
160     Else
170       blnProgBar = CBool(varProgBar)
180       dblPB_Incr = CDbl(varPB_Incr)
190     End If

        ' ** This will be 24% of the overall progress bar width.
        ' ** dblPB_Incr is dblPB_ThisIncrSub, and should be this Step's
        ' ** increment divided by the number of Steps in this function.

200     Set dbs = CurrentDb
210     With dbs

          ' ** Get the highest existing group number from LedgerHidden.
220       blnRetVal = Hide_Max  ' ** Function: Below.

          ' ** dblPB_ThisWidth set prior to entering Function.
230       If blnProgBar = True Then
            ' ***************************************************************
            ' ** Step 2.1 Check LedgerHidden table.
240         dblPB_ThisStepSub = 1# + ((varAcct - 1#) * 8#)
            ' ** varAcct is the incrementing count of accounts.
            ' ** So, 1st time (where varAcct = 1), 1-8:
            ' **   1# + ((varAcct - 1) * 8) = 1
            ' ** 2nd time:
            ' **   1# + ((2 - 1) * 8) = 9
            ' ** 3rd time:
            ' **   1# + ((3 - 1) * 8) = 17
            ' ** etc.
            ' ***************************************************************
            ' ***************************************************************
250         dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
260         ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
            'varFrm.ProgBar_bar.Width = dblPB_ThisWidthSub
270         strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
280         varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
290         DoEvents
            ' ***************************************************************
300       End If

          ' ** Get the record count for LedgerHidden.
310       lngPreExisting = 0&
320       If lngGroupMax > 0& Then
330         varTmp00 = DCount("*", "LedgerHidden")
340         If IsNull(varTmp00) = False Then
350           lngPreExisting = CLng(varTmp00)
360         End If
370       End If

380       If blnProgBar = True Then
            ' ***************************************************************
            ' ** Step 2.2 Empty temporary tables.
390         dblPB_ThisStepSub = 2# + ((varAcct - 1#) * 8#)
            ' ***************************************************************
            ' ***************************************************************
400         dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
410         ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
            'varFrm.ProgBar_bar.Width = dblPB_ThisWidthSub
420         strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
430         varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
440         DoEvents
            ' ***************************************************************
450       End If

          ' ** Empty tmpTrx.
460       Set qdf = .QueryDefs("qryAccountHide_11a")
470       qdf.Execute
480       Set qdf = Nothing
490       DoEvents

          ' ** Empty tmpTrx1.
500       Set qdf = .QueryDefs("qryAccountHide_11b")
510       qdf.Execute
520       Set qdf = Nothing
530       DoEvents

          ' ** Empty tmpTrx2.
540       Set qdf = .QueryDefs("qryAccountHide_11c")
550       qdf.Execute
560       Set qdf = Nothing
570       DoEvents

          ' ** Empty tmpTrx3.
580       Set qdf = .QueryDefs("qryAccountHide_11d")
590       qdf.Execute
600       Set qdf = Nothing
610       DoEvents

          ' ** Empty tmpTrx4.
620       Set qdf = .QueryDefs("qryAccountHide_11e")
630       qdf.Execute
640       Set qdf = Nothing
650       DoEvents

          ' ** Empty tmpTrx5.
660       Set qdf = .QueryDefs("qryAccountHide_11f")
670       qdf.Execute
680       Set qdf = Nothing
690       DoEvents

          ' ** Empty tmpTrx6.
700       Set qdf = .QueryDefs("qryAccountHide_11g")
710       qdf.Execute
720       Set qdf = Nothing
730       DoEvents

          ' ** Empty tmpTrx7.
740       Set qdf = .QueryDefs("qryAccountHide_11h")
750       qdf.Execute
760       Set qdf = Nothing
770       DoEvents

          ' ** Empty tmpTrx8.
780       Set qdf = .QueryDefs("qryAccountHide_11i")
790       qdf.Execute
800       Set qdf = Nothing
810       DoEvents

820       If blnProgBar = True Then
            ' ***************************************************************
            ' ** Step 2.3 Populate temporary tables.
830         dblPB_ThisStepSub = 3# + ((varAcct - 1#) * 8#)
            ' ***************************************************************
            ' ***************************************************************
840         dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
850         ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
            'varFrm.ProgBar_bar.Width = dblPB_ThisWidthSub
860         strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
870         varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
880         DoEvents
            ' ***************************************************************
890       End If

900       If lngPreExisting > 0& Then
            ' ** Append LedgerHidden to tmpTrx, by specified [actno].
910         Set qdf = .QueryDefs("qryAccountHide_42")
920         With qdf.Parameters
930           ![actno] = gstrAccountNo
940         End With
950         qdf.Execute
960         Set qdf = Nothing
970         DoEvents
980       End If

          ' ** Append qryAccountHide_04 (Formatted Account numbers) to tmpTrx6.
990       Set qdf = .QueryDefs("qryAccountHide_04a")
1000      qdf.Execute
1010      Set qdf = Nothing
1020      DoEvents

          ' ** Append qryAccountHide_05c (Formatted journal types) to tmpTrx7.
1030      Set qdf = .QueryDefs("qryAccountHide_05d")
1040      qdf.Execute
1050      Set qdf = Nothing
1060      DoEvents

          ' ** Append qryAccountHide_01 (Ledger, just hidden entries not in LedgerHidden) to tmpTrx5.
          'Set qdf = .QueryDefs("qryAccountHide_01a")
1070      Set qdf = .QueryDefs("qryAccountHide_01t")
1080      With qdf.Parameters
1090        ![actno] = gstrAccountNo
1100      End With
1110      qdf.Execute
1120      Set qdf = Nothing
1130      DoEvents

          ' ** Append qryAccountHide_01b (Ledger, just hidden entries) to tmpTrx8.
          'Set qdf = .QueryDefs("qryAccountHide_01c")
1140      Set qdf = .QueryDefs("qryAccountHide_01v")
1150      With qdf.Parameters
1160        ![actno] = gstrAccountNo
1170      End With
1180      qdf.Execute
1190      Set qdf = Nothing
1200      DoEvents

1210      If blnProgBar = True Then
            ' ***************************************************************
            ' ** Step 2.4 Collect new hiddens.
1220        dblPB_ThisStepSub = 4# + ((varAcct - 1#) * 8#)
            ' ***************************************************************
            ' ***************************************************************
1230        dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
1240        ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
            'varFrm.ProgBar_bar.Width = dblPB_ThisWidthSub
1250        strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
1260        varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
1270        DoEvents
            ' ***************************************************************
1280      End If

          ' ** All the queries below exclude items already in LedgerHidden,
          ' ** so on a normal basis, these next 7 queries won't do anyting.

          ' ** Append qryAccountHide_03g (with cnt = 3 stuff) to tmpTrx4.  THIS IS ALL THE NEW ENTRIES.
1290      Set qdf = .QueryDefs("qryAccountHide_03h")
          'comes from tmpTrx5, which was specified above in qryAccountHide_01t.
1300      qdf.Execute
1310      Set qdf = Nothing
1320      DoEvents

          ' ** Append qryAccountHide_10 to tmpTrx.  THIS IS THE PERFECT MATCHES!
1330      Set qdf = .QueryDefs("qryAccountHide_12")
          'also from tmpTrx5.
1340      qdf.Execute
1350      Set qdf = Nothing
1360      DoEvents

          ' ** Delete tmpTrx4, with DLookups() from .._03_04i.
1370      Set qdf = .QueryDefs("qryAccountHide_03_04j")
1380      qdf.Execute
1390      Set qdf = Nothing
1400      DoEvents

          ' ** Delete tmpTrx, with DLookups() to .._03_04m.
1410      Set qdf = .QueryDefs("qryAccountHide_03_04n")
1420      qdf.Execute
1430      Set qdf = Nothing
1440      DoEvents

          ' ** Append .._03_04m to tmpTrx.
1450      Set qdf = .QueryDefs("qryAccountHide_03_04o")
1460      qdf.Execute
1470      Set qdf = Nothing
1480      DoEvents

          ' ** Append qryAccountHide_20a (unmatched not Misc.) to tmpTrx1.
1490      Set qdf = .QueryDefs("qryAccountHide_20b")
          'tmpTrx5.
1500      qdf.Execute
1510      Set qdf = Nothing
1520      DoEvents

          ' ** Append qryAccountHide_21a (unmatched Misc.) to tmpTrx2.
1530      Set qdf = .QueryDefs("qryAccountHide_21b")
          'tmpTrx5.
1540      qdf.Execute
1550      Set qdf = Nothing
1560      DoEvents

          ' ** Append qryAccountHide_24a to tmpTrx; 1st round of Misc. matching, one-to-one.
1570      Set qdf = .QueryDefs("qryAccountHide_24b")
          'tmpTrx1 and tmpTrx2, originally from tmpTrx5 via qryAccountHide_20b and qryAccountHide_21b, above.
1580      qdf.Execute
1590      Set qdf = Nothing
1600      DoEvents

          ' ** Append qryAccountHide_26a (remaining unmatched) to tmpTrx3.
1610      Set qdf = .QueryDefs("qryAccountHide_26b")
          'from tmpTrx5
1620      qdf.Execute
1630      Set qdf = Nothing
1640      DoEvents

          ' ** Append qryAccountHide_32d to tmpTrx; 2nd round of Misc. matching, one-to-two.
1650      Set qdf = .QueryDefs("qryAccountHide_32")
          'from tmpTrx3, populated above by qryAccountHide_26b
1660      qdf.Execute
1670      Set qdf = Nothing
1680      DoEvents

1690      If blnProgBar = True Then
            ' ***************************************************************
            ' ** Step 2.5 Begin match process.
1700        dblPB_ThisStepSub = 5# + ((varAcct - 1#) * 8#)
            ' ***************************************************************
            ' ***************************************************************
1710        dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
1720        ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
            'varFrm.ProgBar_bar.Width = dblPB_ThisWidthSub
1730        strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
1740        varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
1750        DoEvents
            ' ***************************************************************
1760      End If

          'FIND A WAY -- EARLIER -- TO ASK IF THERE'S A MULTI-LOT GROUP IN THE BUNCH.

          ' ** This next big chunk of code is for multi-lot groups.

          ' ** qryAccountHide_34 (remaining unmatched), grouped and summed.
1770      Set qdf = .QueryDefs("qryAccountHide_35")
          'from tmpTrx5
1780      Set rst = qdf.OpenRecordset
1790      If rst.BOF = True And rst.EOF = True Then
            ' ** No unmatched remain.
1800        rst.Close
1810        Set rst = Nothing
1820        Set qdf = Nothing
1830        DoEvents
1840      Else
1850        rst.MoveLast
1860        lngMultis = rst.RecordCount
1870        rst.MoveFirst
1880        arr_varMulti = rst.GetRows(lngMultis)
1890        rst.MoveFirst
1900        If (lngMultis = 1& And rst![cnt] > 0) Or lngMultis > 1& Then
1910          rst.Close
1920          Set rst = Nothing
1930          Set qdf = Nothing
1940          DoEvents

1950          lngSkips = 0
1960          If lngMultis = 1& Then
1970            If arr_varMulti(M_CNT, 0) = 1& Then
                  ' ** Very odd indeed! I don't think it'll ever hit here.
1980              lngSkips = 1
1990              arr_varMulti(M_SKIP, 0) = True
2000            Else
                  ' ** See if it qualifies.
2010              If arr_varMulti(M_ICASH, 0) = 0@ And arr_varMulti(M_PCASH, 0) = 0@ And arr_varMulti(M_COST, 0) = 0@ Then
                    ' ** If there's more than 1, and it's all zeroes, that means this group matches as a multi-lot.
2020              Else
                    ' ** This group is unknown. We'll have to let it through for now.
2030                lngSkips = 1
2040                arr_varMulti(M_SKIP, 0) = True
2050              End If
2060            End If
2070          Else
                ' ** Check each assetno.
2080            For lngX = 0& To (lngMultis - 1&)
2090              If arr_varMulti(M_ICASH, lngX) = 0@ And arr_varMulti(M_PCASH, lngX) = 0@ And arr_varMulti(M_COST, lngX) = 0@ Then
                    ' ** If there's more than 1, and it's all zeroes, that means this group matches as a multi-lot.
2100              Else
                    ' ** This group is unknown. We'll have to let it through for now.
                    ' ** Skip this one.
2110                lngSkips = lngSkips + 1
2120                arr_varMulti(M_SKIP, lngX) = True
2130              End If
2140            Next
2150          End If

2160          If lngSkips <> lngMultis Then
2170            For lngX = 0& To (lngMultis - 1&)
2180              If arr_varMulti(M_SKIP, lngX) = False Then
                    ' ** tmpTrx3, sorted, by [astno].
2190                Set qdf = dbs.QueryDefs("qryAccountHide_36")
                    'from tmpTrx3
2200                With qdf.Parameters
2210                  ![astno] = arr_varMulti(M_ASTNO, lngX)
2220                End With
2230                Set rst = qdf.OpenRecordset
2240                With rst
2250                  .MoveLast
2260                  lngRecs = .RecordCount
2270                  .MoveFirst
                      ' ** Get the base UniqueID.
                      ' ** Examples:
                      ' **   000000000000011_0055_002070_000000_Sold_______________ .  (Period needed for my formatting function.)
                      ' **   000000000000011_0055_002084_000000_Purchase___________ .  (Otherwise it thinks it's a line-continuation!)
2280                  strTmp01 = ![UniqueID1]
2290                  intPos01 = InStr((InStr(strTmp01, "_") + 1), strTmp01, "_")
2300                  strTmp01 = Left(strTmp01, intPos01)  ' ** 000000000000011_0055_ .
2310                  strTmp02 = vbNullString
2320                  For lngY = 1& To lngRecs
                        ' ** Assemble the UniqueID pieces.
2330                    strTmp01 = strTmp01 & "_" & Mid(![UniqueID1], (intPos01 + 1), 6)  '002070
                        ' ** (intPos01 + 1)  '002070_000000_Sold_______________ .
                        ' ** (InStr((intPos01 + 1), ![UniqueID1], "_") + 1)  '000000_Sold_______________ .
                        ' ** (InStr((InStr((intPos01 + 1), ![UniqueID1], "_") + 1), ![UniqueID1], "_") + 1)  'Sold_______________ .
2340                    If strTmp02 = vbNullString Then
2350                      strTmp02 = Mid(![UniqueID1], _
                            (InStr((InStr((intPos01 + 1), ![UniqueID1], "_") + 1), ![UniqueID1], "_") + 1), ![jlen])  'Sold_____ .
2360                    Else
2370                      strTmp02 = strTmp02 & "_" & Mid(![UniqueID1], _
                            (InStr((InStr((intPos01 + 1), ![UniqueID1], "_") + 1), ![UniqueID1], "_") + 1), ![jlen])
2380                    End If
2390                    If lngY < lngRecs Then .MoveNext
2400                  Next
2410                  strTmp01 = strTmp01 & "_" & strTmp02
2420                  .MoveFirst
                      ' ** Put the group's UniqueID into each of its members in tmpTrx3.
2430                  For lngY = 1& To lngRecs
2440                    .Edit
2450                    If lngY = 1& Then
2460                      datTmp04 = ![transdate]  ' ** Use for SortDate.
2470                    Else
2480                      ![transdate] = datTmp04
2490                    End If
2500                    ![Sorty] = lngY  ' ** Use for Ord.
2510                    ![UniqueIDx] = strTmp01  'Left(strTmp01, 255)  ' ** 174: "000000000000011_0055__002070_002071_002072_002073_002074_002075_002076_002077_002084_Sold______Sold______Sold______Sold______Sold______Sold______Sold______Sold______Purchase_"
2520                    .Update
2530                    If lngY < lngRecs Then .MoveNext
2540                  Next
2550                  .Close
2560                End With  ' ** tmpTrx3: rst.
2570                Set rst = Nothing
2580                Set qdf = Nothing
2590                DoEvents
2600              End If  ' ** Non-skips.
2610            Next  ' ** lngX.
2620          End If  ' ** lngSkips <> lngMultis.

2630          If lngSkips = lngMultis Then
                ' ** All are skips, so everything goes in as Sort = 3, as is.
                ' ** Append qryAccountHide_08, not in tmpTrx, to tmpTrx, for remaining unmatched.
2640            Set qdf = .QueryDefs("qryAccountHide_33")
                'from tmpTrx5
2650            qdf.Execute
2660            Set qdf = Nothing
2670            DoEvents
2680          Else
2690            For lngX = 0& To (lngMultis - 1&)
2700              If arr_varMulti(M_SKIP, lngX) = False Then
                    ' ** Append the matched ones as Sort = 1.
                    ' ** Append tmpTrx3 to tmpTrx, by specified [astno].
2710                Set qdf = .QueryDefs("qryAccountHide_37")
                    'from tmpTrx3
2720                With qdf.Parameters
2730                  ![astno] = arr_varMulti(M_ASTNO, lngX)
2740                End With
2750                qdf.Execute
2760                Set qdf = Nothing
2770                DoEvents
2780              Else
                    ' ** Append the unmatched ones as Sort = 3, as is.
                    ' ** Append qryAccountHide_08, not in tmpTrx, to tmpTrx, for remaining unmatched, by specified [astno].
2790                Set qdf = .QueryDefs("qryAccountHide_38")
                    'from tmpTrx5
2800                With qdf.Parameters
2810                  ![astno] = arr_varMulti(M_ASTNO, lngX)
2820                End With
2830                qdf.Execute
2840                Set qdf = Nothing
2850                DoEvents
2860              End If
2870            Next
2880          End If

2890        Else
              ' ** No unmatched remain.
2900          rst.Close
2910          Set rst = Nothing
2920          Set qdf = Nothing
2930          DoEvents
2940        End If  ' ** 1st unmatched filter.
2950      End If  ' ** Remaining unmatched.

2960      If blnProgBar = True Then
            ' ***************************************************************
            ' ** Step 2.6 Assign groups.
2970        dblPB_ThisStepSub = 6# + ((varAcct - 1#) * 8#)
            ' ***************************************************************
            ' ***************************************************************
2980        dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
2990        ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
            'varFrm.ProgBar_bar.Width = dblPB_ThisWidthSub
3000        strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
3010        varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
3020        DoEvents
            ' ***************************************************************
3030      End If

          ' ** From here on, the processing involves ALL hidden records, including pre-existing.
          ' ** That's because we have to set up the alternating green/blue sections on the screen.

          ' ** tmpTrx, sorted; used for assigning Groups.
3040      Set qdf = .QueryDefs("qryAccountHide_40")
3050      Set rst = qdf.OpenRecordset
3060      With rst
3070        If .BOF = True And .EOF = True Then
              ' ** No hidden transactions!
3080          intRetVal = -1
3090        Else
3100          .MoveLast
3110          lngRecs = .RecordCount
3120          .MoveFirst
              ' ** hid_grpnum is different from Grouping in tmpTrx.
              ' **   hid_grpnum and lngGroup hold the active group in the sequential series of
              ' **     hidden groups, from the very beginning; i.e., Group 1, Group2, Group3, etc.
              ' **   Grouping and lngGrp hold the alternating On/Off, green/blue that's shown on the screen.
3130          lngGrp = 0&: lngTmp03 = 0&
3140          blnMiscGroup = False
3150          For lngX = 1& To lngRecs
                ' ** Sortx is different from Sort in tmpTrx.
                ' ** Sort in tmpTrx:
                ' **   1  Normal matching pairs.
                ' **   2  Records matched in 1st and 2nd level of Misc. group matching.
                ' **   3  Remaing ledger_HIDDEN unable to be matched.
                ' ** Sortx in qryAccountHide_40:
                ' *    1  Sorts 1 and 2 combined, so that they sort together on the screen.
                ' **   3  Unmatchable ledger_HIDDEN's, all pushed to the end of the list.
3160            If ![Sortx] < 3& Then
3170              .Edit
3180              If lngGrp = 0& Then
                    ' ** 1st in group 1.
3190                lngGrp = lngGrp + 1&        ' ** On/Off (green/blue on screen), alternating groups.
3200                ![Grouping] = True          ' ** Alternates True/False signifying alternate colors.
3210                ![grp1] = String(115, "Û")  ' ** Used with Terminal font to give solid box.
3220                lngTmp03 = lngTmp03 + 1&  ' ** lngTmp03 = 1&
                    'lngTmp03 counts the entries in the current group
3230              Else
                    ' **************************************************
3240                If lngGrp = 1& Then
                      ' ** 2nd in group 1, and 3rd in Misc. or multi-lot group.
3250                  ![Grouping] = True
3260                  ![grp1] = String(115, "Û")
3270                  If ![cnt] = 2& Then     'end of a 2-item group
                        'ends here when 2nd of 2-item group (which started in the lngGrp = 0& section),
                        'and lngTmp03 still equals 1&
3280                    lngGrp = lngGrp + 1&  'next alternating group
3290                    lngTmp03 = 0&         'reset lngTmp03 for next group; doesn't go to lngTmp03 = 2 for a 2-item group.
3300                  Else
                        'gets here for item 2 of multi-lot group.
3310                    lngTmp03 = lngTmp03 + 1&  'lngTmp03 now = 2&  (OR MORE?)
                        'if it's still lngGrp = 1& (it started in the lngGrp = 0& section),
                        'lngTmp03 is incrementing, but lngGrp isn't (AND SHOULDN'T!).
3320                    If lngTmp03 = 2& Then
                          'it's item 2 of a 3-or-more-item group
                          ' ** Keep going.
3330                    Else
3340                      If ![cnt] = 3& Then  'I believe lngTmp03 will also be 3
3350                        lngGrp = lngGrp + 1&  'next alternating group
3360                        lngTmp03 = 0&
3370                      Else
3380                        If lngTmp03 = 3& Then
                              'it's item 3 of a 4-or-more-item group
                              ' ** Keep going.
3390                        Else
3400                          If ![cnt] = 4& Then
3410                            lngGrp = lngGrp + 1&
3420                            lngTmp03 = 0&
3430                          Else
3440                            If lngTmp03 = 4& Then
                                  'it's item 4 of a 5-or-more-item group
                                  ' ** Keep going.
3450                            Else
3460                              If ![cnt] = 5& Then
3470                                lngGrp = lngGrp + 1&
3480                                lngTmp03 = 0&
3490                              Else
3500                                If lngTmp03 = 5& Then
                                      'it's item 5 of a 6-or-more-item group
                                      ' ** Keep going.
3510                                Else
3520                                  If ![cnt] = 6& Then
3530                                    lngGrp = lngGrp + 1&
3540                                    lngTmp03 = 0&
3550                                  Else
3560                                    If lngTmp03 = 6& Then
                                          'it's item 6 of a 7-or-more-item group
                                          ' ** Keep going.
3570                                    Else
3580                                      If ![cnt] = 7& Then
3590                                        lngGrp = lngGrp + 1&
3600                                        lngTmp03 = 0&
3610                                      Else
3620                                        If lngTmp03 = 7& Then
                                              'it's item 7 of a 8-or-more-item group
                                              ' ** Keep going.
3630                                        Else
3640                                          If ![cnt] = 8& Then
3650                                            lngGrp = lngGrp + 1&
3660                                            lngTmp03 = 0&
3670                                          Else
3680                                            If lngTmp03 = 8& Then
                                                  'it's item 8 of a 9-or-more-item group
                                                  ' ** Keep going.
3690                                            Else
3700                                              If ![cnt] = 9& Then
3710                                                lngGrp = lngGrp + 1&
3720                                                lngTmp03 = 0&
3730                                              Else
3740                                                If lngTmp03 = 9& Then
                                                      'it's item 9 of a 10-or-more-item group
                                                      ' ** Keep going.
3750                                                Else
3760                                                  If ![cnt] = 10& Then
3770                                                    lngGrp = lngGrp + 1&
3780                                                    lngTmp03 = 0&
3790                                                  Else
3800                                                    If lngTmp03 = 10& Then
                                                          'it's item 10 of a 11-or-more-item group
                                                          ' ** Keep going.
3810                                                    Else
3820                                                      If ![cnt] = 11& Then
3830                                                        lngGrp = lngGrp + 1&
3840                                                        lngTmp03 = 0&
3850                                                      Else
3860                                                        If lngTmp03 = 11& Then
                                                              'it's item 11 of a 12-or-more-item group
                                                              ' ** Keep going.
3870                                                        Else
3880                                                          If ![cnt] = 12& Then
3890                                                            lngGrp = lngGrp + 1&
3900                                                            lngTmp03 = 0&
3910                                                          Else
3920                                                            If lngTmp03 = 12& Then
                                                                  'it's item 12 of a 13-or-more-item group
                                                                  ' ** Keep going.
3930                                                            Else
3940                                                              If ![cnt] = 13& Then
3950                                                                lngGrp = lngGrp + 1&
3960                                                                lngTmp03 = 0&
3970                                                              Else
3980                                                                If lngTmp03 = 13& Then
                                                                      'it's item 13 of a 14-or-more-item group
                                                                      ' ** Keep going.
3990                                                                Else
4000                                                                  If ![cnt] = 14& Then
4010                                                                    lngGrp = lngGrp + 1&
4020                                                                    lngTmp03 = 0&
4030                                                                  Else
4040                                                                    If lngTmp03 = 14& Then
                                                                          'it's item 14 of a 15-or-more-item group
                                                                          ' ** Keep going.
4050                                                                    Else
4060                                                                      If ![cnt] = 15& Then
4070                                                                        lngGrp = lngGrp + 1&
4080                                                                        lngTmp03 = 0&
4090                                                                      Else
4100                                                                        If lngTmp03 = 15& Then
                                                                              'it's item 15 of a 16-or-more-item group
                                                                              ' ** Keep going.
4110                                                                        Else
4120                                                                          If ![cnt] = 16& Then
4130                                                                            lngGrp = lngGrp + 1&
4140                                                                            lngTmp03 = 0&
4150                                                                          Else
4160                                                                            If lngTmp03 = 16& Then
                                                                                  'it's item 16 of a 17-or-more-item group
                                                                                  ' ** Keep going.
4170                                                                            Else
4180                                                                              If ![cnt] = 17& Then
4190                                                                                lngGrp = lngGrp + 1&
4200                                                                                lngTmp03 = 0&
4210                                                                              Else
4220                                                                                If lngTmp03 = 17& Then
                                                                                      'it's item 17 of a 18-or-more-item group
                                                                                      ' ** Keep going.
4230                                                                                Else
4240                                                                                  If ![cnt] = 18& Then
4250                                                                                    lngGrp = lngGrp + 1&
4260                                                                                    lngTmp03 = 0&
4270                                                                                  Else
4280                                                                                    If lngTmp03 = 18& Then
                                                                                          'it's item 18 of a 19-or-more-item group
                                                                                          ' ** Keep going.
4290                                                                                    Else
4300                                                                                      If ![cnt] = 19& Then
4310                                                                                        lngGrp = lngGrp + 1&
4320                                                                                        lngTmp03 = 0&
4330                                                                                      Else
4340                                                                                        If lngTmp03 = 19& Then
                                                                                              'it's item 19 of a 20-or-more-item group
                                                                                              ' ** Keep going.
4350                                                                                        Else
4360                                                                                          If ![cnt] = 20& Then
4370                                                                                            lngGrp = lngGrp + 1&
4380                                                                                            lngTmp03 = 0&
4390                                                                                          Else
4400                                                                                            If lngTmp03 = 20& Then
                                                                                                  'it's item 20 of a 21-or-more-item group
                                                                                                  ' ** Keep going.
4410                                                                                            Else
4420                                                                                              If ![cnt] = 21& Then
4430                                                                                                lngGrp = lngGrp + 1&
4440                                                                                                lngTmp03 = 0&
4450                                                                                              Else
4460                                                                                                If lngTmp03 = 21& Then
                                                                                                      'it's item 21 of a 22-or-more-item group
                                                                                                      ' ** Keep going.
4470                                                                                                Else
4480                                                                                                  If ![cnt] = 22& Then
4490                                                                                                    lngGrp = lngGrp + 1&
4500                                                                                                    lngTmp03 = 0&
4510                                                                                                  Else
4520                                                                                                    If lngTmp03 = 22& Then
                                                                                                          'it's item 22 of a 23-or-more-item group
                                                                                                          ' ** Keep going.
4530                                                                                                    Else
4540                                                                                                      If ![cnt] = 23& Then
4550                                                                                                        lngGrp = lngGrp + 1&
4560                                                                                                        lngTmp03 = 0&
4570                                                                                                      Else
4580                                                                                                        If lngTmp03 = 23& Then
                                                                                                              'it's item 23 of a 24-or-more-item group
                                                                                                              ' ** Keep going.
4590                                                                                                        Else
4600                                                                                                          If ![cnt] = 24& Then
4610                                                                                                            lngGrp = lngGrp + 1&
4620                                                                                                            lngTmp03 = 0&
4630                                                                                                          Else
4640                                                                                                            If lngTmp03 = 24& Then
                                                                                                                  'it's item 24 of a 25-or-more-item group
                                                                                                                  ' ** Keep going.
4650                                                                                                            Else
4660                                                                                                              If ![cnt] = 25& Then
4670                                                                                                                lngGrp = lngGrp + 1&
4680                                                                                                                lngTmp03 = 0&
4690                                                                                                              Else
                                                                                                                    ' ** Not set up for bigger groups.
4700                                                                                                                MsgBox "Grouping has more than 25 items.", _
                                                                                                                      vbCritical + vbOKOnly, "Group Size Not Accommodated"
4710                                                                                                              End If
4720                                                                                                            End If
4730                                                                                                          End If
4740                                                                                                        End If
4750                                                                                                      End If
4760                                                                                                    End If
4770                                                                                                  End If
4780                                                                                                End If
4790                                                                                              End If
4800                                                                                            End If
4810                                                                                          End If
4820                                                                                        End If
4830                                                                                      End If
4840                                                                                    End If
4850                                                                                  End If
4860                                                                                End If
4870                                                                              End If
4880                                                                            End If
4890                                                                          End If
4900                                                                        End If
4910                                                                      End If
4920                                                                    End If
4930                                                                  End If
4940                                                                End If
4950                                                              End If
4960                                                            End If
4970                                                          End If
4980                                                        End If
4990                                                      End If
5000                                                    End If
5010                                                  End If
5020                                                End If
5030                                              End If
5040                                            End If
5050                                          End If
5060                                        End If
5070                                      End If
5080                                    End If
5090                                  End If
5100                                End If
5110                              End If
5120                            End If
5130                          End If
5140                        End If
5150                      End If
5160                    End If
5170                  End If
5180                Else
5190                  If lngGrp = 2& Then
                        ' ** 1st in group 2.  '![journalno] = 21320 1
5200                    lngGrp = lngGrp + 1&  'lngGrp now equals 3!
                        'WHY DOES lngGrp KEEP GROWING IF IT'S SUPPOSED TO BE JUST ALTERNATING?
5210                    ![Grouping] = False          ' ** Alternates True/False signifying alternate colors.
5220                    ![grp2] = String(115, "Û")
                        'move on to item 2 of the 4-item group
5230                  Else
                        ' ** 2nd in group 2, and 3rd in Misc. or multi-lot group.
5240                    If lngGrp = 3& Then
                          'gets here for item 2 of the 4-item group
5250                      lngGrp = lngGrp + 1&  'lngGrp now equals 4!
5260                      ![Grouping] = False
5270                      ![grp2] = String(115, "Û")
5280                      If ![cnt] = 2& Then
5290                        lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
5300                        lngTmp03 = 0&
5310                      Else
                            ' ** Keep going
5320                        lngTmp03 = lngTmp03 + 1&
5330                      End If
5340                    Else
5350                      If lngGrp = 4& Then
                            'gets here for item 3 of 4-item group
5360                        lngGrp = lngGrp + 1&  'lngGrp now equals 5!
5370                        ![Grouping] = False
5380                        ![grp2] = String(115, "Û")
5390                        If ![cnt] = 3& Then
5400                          lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
5410                          lngTmp03 = 0&
5420                        Else
                              ' ** Keep going
5430                          lngTmp03 = lngTmp03 + 1&
5440                        End If
5450                      Else
5460                        If lngGrp = 5& Then
                              'gets here for item 4 of 4-item group
5470                          lngGrp = lngGrp + 1&  'lngGrp now equals 6!
5480                          ![Grouping] = False
5490                          ![grp2] = String(115, "Û")
5500                          If ![cnt] = 4& Then
5510                            lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
5520                            lngTmp03 = 0&
5530                          Else
                                ' ** Keep going
5540                            lngTmp03 = lngTmp03 + 1&
5550                          End If
5560                        Else
5570                          If lngGrp = 6& Then
                                'gets here for item 5 of 5-item group
5580                            lngGrp = lngGrp + 1&  'lngGrp now equals 7!
5590                            ![Grouping] = False
5600                            ![grp2] = String(115, "Û")
5610                            If ![cnt] = 5& Then
5620                              lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
5630                              lngTmp03 = 0&
5640                            Else
                                  ' ** Keep going
5650                              lngTmp03 = lngTmp03 + 1&
5660                            End If
5670                          Else
5680                            If lngGrp = 7& Then
                                  'gets here for item 6 of 6-item group
5690                              lngGrp = lngGrp + 1&  'lngGrp now equals 8!
5700                              ![Grouping] = False
5710                              ![grp2] = String(115, "Û")
5720                              If ![cnt] = 6& Then
5730                                lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
5740                                lngTmp03 = 0&
5750                              Else
                                    ' ** Keep going
5760                                lngTmp03 = lngTmp03 + 1&
5770                              End If
5780                            Else
5790                              If lngGrp = 8& Then
                                    'gets here for item 7 of 7-item group
5800                                lngGrp = lngGrp + 1&  'lngGrp now equals 9!
5810                                ![Grouping] = False
5820                                ![grp2] = String(115, "Û")
5830                                If ![cnt] = 7& Then
5840                                  lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
5850                                  lngTmp03 = 0&
5860                                Else
                                      ' ** Keep going
5870                                  lngTmp03 = lngTmp03 + 1&
5880                                End If
5890                              Else
5900                                If lngGrp = 9& Then
                                      'gets here for item 8 of 8-item group
5910                                  lngGrp = lngGrp + 1&  'lngGrp now equals 10!
5920                                  ![Grouping] = False
5930                                  ![grp2] = String(115, "Û")
5940                                  If ![cnt] = 8& Then
5950                                    lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
5960                                    lngTmp03 = 0&
5970                                  Else
                                        ' ** Keep going
5980                                    lngTmp03 = lngTmp03 + 1&
5990                                  End If
6000                                Else
6010                                  If lngGrp = 10& Then
                                        'gets here for item 9 of 9-item group
6020                                    lngGrp = lngGrp + 1&  'lngGrp now equals 11!
6030                                    ![Grouping] = False
6040                                    ![grp2] = String(115, "Û")
6050                                    If ![cnt] = 9& Then
6060                                      lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
6070                                      lngTmp03 = 0&
6080                                    Else
                                          ' ** Keep going
6090                                      lngTmp03 = lngTmp03 + 1&
6100                                    End If
6110                                  Else
6120                                    If lngGrp = 11& Then
                                          'gets here for item 10 of 10-item group
6130                                      lngGrp = lngGrp + 1&  'lngGrp now equals 12!
6140                                      ![Grouping] = False
6150                                      ![grp2] = String(115, "Û")
6160                                      If ![cnt] = 10& Then
6170                                        lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
6180                                        lngTmp03 = 0&
6190                                      Else
                                            ' ** Keep going
6200                                        lngTmp03 = lngTmp03 + 1&
6210                                      End If
6220                                    Else
6230                                      If lngGrp = 12& Then
                                            'gets here for item 11 of 11-item group
6240                                        lngGrp = lngGrp + 1&  'lngGrp now equals 13!
6250                                        ![Grouping] = False
6260                                        ![grp2] = String(115, "Û")
6270                                        If ![cnt] = 11& Then
6280                                          lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
6290                                          lngTmp03 = 0&
6300                                        Else
                                              ' ** Keep going
6310                                          lngTmp03 = lngTmp03 + 1&
6320                                        End If
6330                                      Else
6340                                        If lngGrp = 13& Then
                                              'gets here for item 12 of 12-item group
6350                                          lngGrp = lngGrp + 1&  'lngGrp now equals 14!
6360                                          ![Grouping] = False
6370                                          ![grp2] = String(115, "Û")
6380                                          If ![cnt] = 12& Then
6390                                            lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
6400                                            lngTmp03 = 0&
6410                                          Else
                                                ' ** Keep going
6420                                            lngTmp03 = lngTmp03 + 1&
6430                                          End If
6440                                        Else
6450                                          If lngGrp = 14& Then
                                                'gets here for item 13 of 13-item group
6460                                            lngGrp = lngGrp + 1&  'lngGrp now equals 15!
6470                                            ![Grouping] = False
6480                                            ![grp2] = String(115, "Û")
6490                                            If ![cnt] = 13& Then
6500                                              lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
6510                                              lngTmp03 = 0&
6520                                            Else
                                                  ' ** Keep going
6530                                              lngTmp03 = lngTmp03 + 1&
6540                                            End If
6550                                          Else
6560                                            If lngGrp = 15& Then
                                                  'gets here for item 14 of 14-item group
6570                                              lngGrp = lngGrp + 1&  'lngGrp now equals 16!
6580                                              ![Grouping] = False
6590                                              ![grp2] = String(115, "Û")
6600                                              If ![cnt] = 14& Then
6610                                                lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
6620                                                lngTmp03 = 0&
6630                                              Else
                                                    ' ** Keep going
6640                                                lngTmp03 = lngTmp03 + 1&
6650                                              End If
6660                                            Else
6670                                              If lngGrp = 16& Then
                                                    'gets here for item 15 of 15-item group
6680                                                lngGrp = lngGrp + 1&  'lngGrp now equals 17!
6690                                                ![Grouping] = False
6700                                                ![grp2] = String(115, "Û")
6710                                                If ![cnt] = 15& Then
6720                                                  lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
6730                                                  lngTmp03 = 0&
6740                                                Else
                                                      ' ** Keep going
6750                                                  lngTmp03 = lngTmp03 + 1&
6760                                                End If
6770                                              Else
6780                                                If lngGrp = 17& Then
                                                      'gets here for item 16 of 16-item group
6790                                                  lngGrp = lngGrp + 1&  'lngGrp now equals 18!
6800                                                  ![Grouping] = False
6810                                                  ![grp2] = String(115, "Û")
6820                                                  If ![cnt] = 16& Then
6830                                                    lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
6840                                                    lngTmp03 = 0&
6850                                                  Else
                                                        ' ** Keep going
6860                                                    lngTmp03 = lngTmp03 + 1&
6870                                                  End If
6880                                                Else
6890                                                  If lngGrp = 18& Then
                                                        'gets here for item 17 of 17-item group
6900                                                    lngGrp = lngGrp + 1&  'lngGrp now equals 19!
6910                                                    ![Grouping] = False
6920                                                    ![grp2] = String(115, "Û")
6930                                                    If ![cnt] = 17& Then
6940                                                      lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
6950                                                      lngTmp03 = 0&
6960                                                    Else
                                                          ' ** Keep going
6970                                                      lngTmp03 = lngTmp03 + 1&
6980                                                    End If
6990                                                  Else
7000                                                    If lngGrp = 19& Then
                                                          'gets here for item 18 of 18-item group
7010                                                      lngGrp = lngGrp + 1&  'lngGrp now equals 20!
7020                                                      ![Grouping] = False
7030                                                      ![grp2] = String(115, "Û")
7040                                                      If ![cnt] = 18& Then
7050                                                        lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
7060                                                        lngTmp03 = 0&
7070                                                      Else
                                                            ' ** Keep going
7080                                                        lngTmp03 = lngTmp03 + 1&
7090                                                      End If
7100                                                    Else
7110                                                      If lngGrp = 20& Then
                                                            'gets here for item 19 of 19-item group
7120                                                        lngGrp = lngGrp + 1&  'lngGrp now equals 21!
7130                                                        ![Grouping] = False
7140                                                        ![grp2] = String(115, "Û")
7150                                                        If ![cnt] = 19& Then
7160                                                          lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
7170                                                          lngTmp03 = 0&
7180                                                        Else
                                                              ' ** Keep going
7190                                                          lngTmp03 = lngTmp03 + 1&
7200                                                        End If
7210                                                      Else
7220                                                        If lngGrp = 21& Then
                                                              'gets here for item 20 of 20-item group
7230                                                          lngGrp = lngGrp + 1&  'lngGrp now equals 22!
7240                                                          ![Grouping] = False
7250                                                          ![grp2] = String(115, "Û")
7260                                                          If ![cnt] = 20& Then
7270                                                            lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
7280                                                            lngTmp03 = 0&
7290                                                          Else
                                                                ' ** Keep going
7300                                                            lngTmp03 = lngTmp03 + 1&
7310                                                          End If
7320                                                        Else
7330                                                          If lngGrp = 22& Then
                                                                'gets here for item 21 of 21-item group
7340                                                            lngGrp = lngGrp + 1&  'lngGrp now equals 23!
7350                                                            ![Grouping] = False
7360                                                            ![grp2] = String(115, "Û")
7370                                                            If ![cnt] = 21& Then
7380                                                              lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
7390                                                              lngTmp03 = 0&
7400                                                            Else
                                                                  ' ** Keep going
7410                                                              lngTmp03 = lngTmp03 + 1&
7420                                                            End If
7430                                                          Else
7440                                                            If lngGrp = 23& Then
                                                                  'gets here for item 22 of 22-item group
7450                                                              lngGrp = lngGrp + 1&  'lngGrp now equals 24!
7460                                                              ![Grouping] = False
7470                                                              ![grp2] = String(115, "Û")
7480                                                              If ![cnt] = 22& Then
7490                                                                lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
7500                                                                lngTmp03 = 0&
7510                                                              Else
                                                                    ' ** Keep going
7520                                                                lngTmp03 = lngTmp03 + 1&
7530                                                              End If
7540                                                            Else
7550                                                              If lngGrp = 24& Then
                                                                    'gets here for item 23 of 23-item group
7560                                                                lngGrp = lngGrp + 1&  'lngGrp now equals 25!
7570                                                                ![Grouping] = False
7580                                                                ![grp2] = String(115, "Û")
7590                                                                If ![cnt] = 23& Then
7600                                                                  lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
7610                                                                  lngTmp03 = 0&
7620                                                                Else
                                                                      ' ** Keep going
7630                                                                  lngTmp03 = lngTmp03 + 1&
7640                                                                End If
7650                                                              Else
7660                                                                If lngGrp = 25& Then
                                                                      'gets here for item 24 of 24-item group
7670                                                                  lngGrp = lngGrp + 1&  'lngGrp now equals 26!
7680                                                                  ![Grouping] = False
7690                                                                  ![grp2] = String(115, "Û")
7700                                                                  If ![cnt] = 24& Then
7710                                                                    lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
7720                                                                    lngTmp03 = 0&
7730                                                                  Else
                                                                        ' ** Keep going
7740                                                                    lngTmp03 = lngTmp03 + 1&
7750                                                                  End If
7760                                                                Else
7770                                                                  If lngGrp = 26& Then
                                                                        'gets here for item 25 of 25-item group
7780                                                                    lngGrp = lngGrp + 1&  'lngGrp now equals 27!
7790                                                                    ![Grouping] = False
7800                                                                    ![grp2] = String(115, "Û")
7810                                                                    If ![cnt] = 25& Then
7820                                                                      lngGrp = 0&  'the group is finished, reset so it'll be incremented to 1 on the next group
7830                                                                      lngTmp03 = 0&
7840                                                                    Else
                                                                          ' ** Keep going
7850                                                                      lngTmp03 = lngTmp03 + 1&
7860                                                                    End If
7870                                                                  Else
                                                                        ' ** Not set up to handle bigger groups.
7880                                                                    MsgBox "Grouping has more than 25 items.", _
                                                                          vbCritical + vbOKOnly, "Group Size Not Accommodated"
7890                                                                  End If
7900                                                                End If
7910                                                              End If
7920                                                            End If
7930                                                          End If
7940                                                        End If
7950                                                      End If
7960                                                    End If
7970                                                  End If
7980                                                End If
7990                                              End If
8000                                            End If
8010                                          End If
8020                                        End If
8030                                      End If
8040                                    End If
8050                                  End If
8060                                End If
8070                              End If
8080                            End If
8090                          End If
8100                        End If
8110                      End If
8120                    End If
8130                  End If
8140                End If
                    ' **************************************************
8150              End If
8160              .Update
8170            Else
8180              .Edit
8190              ![Grouping] = False
8200              ![grp1] = Null
8210              ![grp2] = Null
8220              .Update
8230            End If
8240            If lngX < lngRecs Then .MoveNext
8250          Next
8260        End If
8270        .Close
8280      End With
8290      Set rst = Nothing
8300      Set qdf = Nothing
8310      DoEvents

8320      If blnProgBar = True Then
            ' ***************************************************************
            ' ** Step 2.7 Set up the hidden array.
8330        dblPB_ThisStepSub = 7# + ((varAcct - 1#) * 8#)
            ' ***************************************************************
            ' ***************************************************************
8340        dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
8350        ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
            'varFrm.ProgBar_bar.Width = dblPB_ThisWidthSub
8360        strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
8370        varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
8380        DoEvents
            ' ***************************************************************
8390      End If

          ' ** Populate the arr_varHide() array.
          ' ** A problem there will return False.
8400      Hide_LoadArray  ' ** Function: Below.

          ' ** Append new records only to LedgerHidden from qryAccountHide_50 (hidden trx).
          ' ** Query uses Hide_Group(), below, to retrieve the group number, via FormRef(), from the arr_varHide() array.
          ' ** The accountno has already been specified, via FormRef() and gstrFormQuerySpec.
8410      Set qdf = .QueryDefs("qryAccountHide_41")
8420      qdf.Execute
8430      Set qdf = Nothing
8440      DoEvents

8450      .Close
8460    End With
8470    Set dbs = Nothing

8480    If blnProgBar = True Then
          ' ***************************************************************
          ' ** Step 2.8 Finish Hide_Setup().
8490      dblPB_ThisStepSub = 8# + ((varAcct - 1#) * 8#)
          ' ***************************************************************
          ' ***************************************************************
8500      dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
8510      ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
          'varFrm.ProgBar_bar.Width = dblPB_ThisWidthSub
8520      strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
8530      varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
8540      DoEvents
          ' ***************************************************************
8550    End If

EXITP:
8560    DoCmd.Hourglass False
8570    Set rst = Nothing
8580    Set qdf = Nothing
8590    Set dbs = Nothing
8600    Hide_Setup = intRetVal
8610    Exit Function

ERRH:
8620    DoCmd.Hourglass False
8630    intRetVal = -9
8640    Select Case ERR.Number
        Case Else
8650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8660    End Select
8670    Resume EXITP

End Function

Public Function Hide_Group(varUniqueID As Variant) As Variant
'IS THIS WHAT SHOULD HAVE BEEN USED IN QUERY?
'FormRef([UniqueIDx]) -> Hide_Group([UniqueIDx]) ?
'NO, FormRef() CALLS Hide_Group() (WHEN IT'S IN THE RIGHT SECTION!)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "Hide_Group"

        Dim lngX As Long
        Dim varRetVal As Variant

8710    varRetVal = Null

8720  On Error Resume Next
8730    lngX = arr_varHide(H_NUM, 0)
8740    If ERR.Number <> 0 Then
8750      Select Case ERR.Number
          Case 9  ' ** Subscript out of range.
8760  On Error GoTo ERRH
            ' ** Array hasn't been initialized.
8770        Hide_LoadArray  ' ** Function: Below.
8780      Case Else
8790        zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8800  On Error GoTo ERRH
8810        varRetVal = 0
8820      End Select
8830    Else
8840  On Error GoTo ERRH
8850    End If
8860    If IsNull(varUniqueID) = False Then
8870      For lngX = 0& To (lngHides - 1&)
8880        If arr_varHide(H_UNIQ, lngX) = varUniqueID Then
8890          varRetVal = arr_varHide(H_NUM, lngX)
8900          Exit For
8910        End If
8920      Next
8930    End If

EXITP:
8940    Hide_Group = varRetVal
8950    Exit Function

ERRH:
8960    Select Case ERR.Number
        Case Else
8970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8980    End Select
8990    Resume EXITP

End Function

Public Function Hide_Type(varUniqueID As Variant, Optional varTypeOnly As Variant) As Variant
' ** Group Types returned by the array are:
' **   NORM
' **   NORM_MISC
' **   MISC_2_GRP
' **   MISC_3_GRP
' **   GRP_NONE

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "Hide_Type"

        Dim intPos01 As Integer, intLen As Integer
        Dim strUniqueID As String
        Dim strJType As String
        Dim lngJTypes As Long, arr_varJType() As Variant
        Dim lngGrpCnt As Long
        Dim blnHasMisc As Boolean, blnTypeOnly As Boolean
        Dim intX As Integer, lngX As Long, lngE As Long
        Dim varRetVal As Variant

        Const JT_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const JT_JTYP As Integer = 0
        Const JT_CNT  As Integer = 1
        Const JT_TYPE As Integer = 2

9010    lngJTypes = 0&
9020    ReDim arr_varJType(JT_ELEMS, 0)
        ' **********************************************
        ' ** Array: arr_varJType()
        ' **
        ' **   Element  Name                Constant
        ' **   =======  ==================  ==========
        ' **      0     Journal Type        JT_JTYP
        ' **      1     Entries in Group    JT_CNT
        ' **      2     Group Type          JT_TYPE
        ' **
        ' **********************************************

9030    If IsNull(varUniqueID) = False Then

9040      If IsMissing(varTypeOnly) = True Then
9050        blnTypeOnly = False
9060      Else
9070        blnTypeOnly = CBool(varTypeOnly)
9080      End If

9090      strUniqueID = varUniqueID

          ' ** Examples of [UniqueID] in tmpTrx and LedgerHidden:
          ' ** Simple pair matching:
          ' **   000000000000011_0012_001593_001563_Purchase__Sold_____ .
          ' **   000000000000011_0000_001697_001742_Misc._____Misc.____ .
          ' ** Misc. pair matching (1st level Misc group matching):
          ' **   000000000000011_0012_0000_001923_001972_Dividend__Misc.____ .
          ' ** Misc. triplet matching (2nd level Misc group matching):
          ' **   000000000000011_0001_0000_002081_002082_002083_Purchase__Sold______Misc.____ .
          ' ** Multi-lot matching (3rd level matching, this one's 174 characters long!):
          ' **   000000000000011_0055__002070_002071_002072_002073_002074_002075_002076_002077_002084_Sold______Sold______Sold______Sold______Sold______Sold______Sold______Sold______Purchase_ .

          ' ** Look for the first numeral, right-to-left.
9100      intLen = Len(strUniqueID): intPos01 = 0
9110      For intX = intLen To 1 Step -1
9120        If Asc(Mid(strUniqueID, intX, 1)) >= 48 And Asc(Mid(strUniqueID, intX, 1)) <= 57 Then
              ' ** It's a numeral.
9130          intPos01 = intX
9140          Exit For
9150        End If
9160      Next

          ' ** Now parse out the JournalTypes.
          ' ** Note: 'Cost Adj.' retains its space within the UniqueID.
          ' ** 'Cost Adj.', 'Liability', and 'Withdrawn' are all 9 characters long,
          ' ** the maximum JournalType length. The rest are padded with underscores.

          ' ** qryAccountHide_05c:
          ' **   Cost Adj.  Cost Adj.
          ' **   Deposit    Deposit__
          ' **   Dividend   Dividend_
          ' **   Interest   Interest_
          ' **   Liability  Liability
          ' **   Misc.      Misc.____
          ' **   Paid       Paid_____
          ' **   Purchase   Purchase_
          ' **   Received   Received_
          ' **   Sold       Sold_____
          ' **   Withdrawn  Withdrawn

9170      If intPos01 > 0 Then

9180        strJType = vbNullString: blnHasMisc = False

9190        strUniqueID = Mid(strUniqueID, (intPos01 + 2))  ' ** Trim to the beginning of the 1st JournalType.
9200        If Left(strUniqueID, 1) = "_" Then strUniqueID = Mid(strUniqueID, 2)
9210        strJType = Left(strUniqueID, 9)  ' ** strUniqueID is 19 chars for 2-entry, 29 for 3-entry, 39, 49, etc.
9220        For intX = 9 To 1 Step -1
9230          If Mid(strJType, intX, 1) <> "_" Then
9240            strJType = Left(strJType, intX)
9250            If strJType = "Misc." Then blnHasMisc = True
9260            lngJTypes = lngJTypes + 1&
9270            lngE = lngJTypes - 1&
9280            ReDim Preserve arr_varJType(JT_ELEMS, lngE)
9290            arr_varJType(JT_JTYP, lngE) = strJType
9300            arr_varJType(JT_CNT, lngE) = CLng(0)
9310            arr_varJType(JT_TYPE, lngE) = Null
9320            Exit For
9330          End If
9340        Next

9350        strUniqueID = Mid(strUniqueID, 11)  ' ** Trim to the beginning of the 2nd JournalType (9 + '_', then next JournalType).
9360        If Left(strUniqueID, 1) = "_" Then strUniqueID = Mid(strUniqueID, 2)
9370        strJType = Left(strUniqueID, 9)
9380        For intX = 9 To 1 Step -1
9390          If Mid(strJType, intX, 1) <> "_" Then
9400            strJType = Left(strJType, intX)
9410            If strJType = "Misc." Then blnHasMisc = True
9420            lngJTypes = lngJTypes + 1&
9430            lngE = lngJTypes - 1&
9440            ReDim Preserve arr_varJType(JT_ELEMS, lngE)
9450            arr_varJType(JT_JTYP, lngE) = strJType
9460            arr_varJType(JT_CNT, lngE) = CLng(0)
9470            arr_varJType(JT_TYPE, lngE) = Null
9480            Exit For
9490          End If
9500        Next

9510        If Len(strUniqueID) = 9 Then  ' ** strUniqueID has been trimmed to the beginning of the 2nd Journaltype.
              ' ** A 2-entry grouping.
9520          lngGrpCnt = 2&
9530          arr_varJType(JT_CNT, 0) = lngGrpCnt
9540          arr_varJType(JT_CNT, 1) = lngGrpCnt
9550        Else
              ' ** At least a 3-entry grouping.
              ' ** At this point strUniqueID is 19 chars for 3-entry, 29 for 4-entry, 39, 49, etc.
9560          strUniqueID = Mid(strUniqueID, 11)  ' ** Trim to the beginning of the 3rd JournalType (9 + '_', then next JournalType).
9570          If Left(strUniqueID, 1) = "_" Then strUniqueID = Mid(strUniqueID, 2)
9580          strJType = Left(strUniqueID, 9)  ' ** Also length of strUniqueID if only a 3-entry.
9590          For intX = 9 To 1 Step -1
9600            If Mid(strJType, intX, 1) <> "_" Then
9610              strJType = Left(strJType, intX)
9620              If strJType = "Misc." Then blnHasMisc = True
9630              lngJTypes = lngJTypes + 1&
9640              lngE = lngJTypes - 1&
9650              ReDim Preserve arr_varJType(JT_ELEMS, lngE)
9660              arr_varJType(JT_JTYP, lngE) = strJType
9670              arr_varJType(JT_CNT, lngE) = CLng(0)
9680              arr_varJType(JT_TYPE, lngE) = Null
9690              Exit For
9700            End If
9710          Next
              ' ** At this point strUniqueID is 9 chars for 3-entry, 19 for 4-entry, 29, 39, etc.
9720          lngGrpCnt = (((Len(strUniqueID) + 1) / 10&) + 2&)  ' ** Group size.
9730          arr_varJType(JT_CNT, 0) = lngGrpCnt
9740          arr_varJType(JT_CNT, 1) = lngGrpCnt
9750          arr_varJType(JT_CNT, 2) = lngGrpCnt
9760          If lngGrpCnt > 3& Then
9770            lngJTypes = lngGrpCnt
9780            lngE = (lngJTypes - 1&)
9790            ReDim Preserve arr_varJType(JT_ELEMS, lngE)
                ' ** Get the the group data for entries 4 through lngGrpCnt.
9800            For lngX = (4& - 1&) To (lngGrpCnt - 1&)
9810              strUniqueID = Mid(strUniqueID, 11)  ' ** Trim to the beginning of the next JournalType.
9820              If Left(strUniqueID, 1) = "_" Then strUniqueID = Mid(strUniqueID, 2)
9830              strJType = Left(strUniqueID, 9)
9840              For intX = 9 To 1 Step -1
9850                If Mid(strJType, intX, 1) <> "_" Then
9860                  strJType = Left(strJType, intX)
9870                  If strJType = "Misc." Then blnHasMisc = True
9880                  arr_varJType(JT_JTYP, lngX) = strJType
9890                  arr_varJType(JT_CNT, lngX) = lngGrpCnt
9900                  arr_varJType(JT_TYPE, lngX) = Null
9910                  Exit For
9920                End If
9930              Next
9940            Next
9950          End If
9960        End If

            ' ** Now specify the hidden group's type.
9970        If blnHasMisc = False Then
              ' ** This could be a standard 2 entry group, or a multi-lot group.
9980          If lngGrpCnt = 2& Then
                ' ** 2 entries in hidden group, with matching Assetno (which could both be zero in a non multi-lot group).
9990            For lngX = 0& To (lngGrpCnt - 1&)
10000             arr_varJType(JT_TYPE, lngX) = "NORM"
10010           Next
10020           If blnTypeOnly = False Then
10030             varRetVal = arr_varJType
10040           Else
10050             varRetVal = "NORM"
10060           End If
10070         Else
10080           For lngX = 0& To (lngGrpCnt - 1&)
10090             arr_varJType(JT_TYPE, lngX) = "MULTI_GRP"
10100           Next
10110           If blnTypeOnly = False Then
10120             varRetVal = arr_varJType
10130           Else
10140             varRetVal = "MULTI_GRP"
10150           End If
10160         End If
10170       Else
              ' ** This group has a Misc, and shouldn't ever be a multi-lot group.
10180         If arr_varJType(JT_CNT, 0) = 2 Then
                ' ** 2 entries in hidden group.
10190           If arr_varJType(JT_JTYP, 0) = "Misc." And arr_varJType(JT_JTYP, 1) = "Misc." Then
                  ' ** Both are Misc., so they're treated like a normal pair.
10200             For lngX = 0& To (lngGrpCnt - 1&)
10210               arr_varJType(JT_TYPE, lngX) = "NORM_MISC"
10220             Next
10230             If blnTypeOnly = False Then
10240               varRetVal = arr_varJType
10250             Else
10260               varRetVal = "NORM_MISC"
10270             End If
10280           Else
                  ' ** 1 Misc. and 1 other.
10290             For lngX = 0& To (lngGrpCnt - 1&)
10300               arr_varJType(JT_TYPE, lngX) = "MISC_2_GRP"
10310             Next
10320             If blnTypeOnly = False Then
10330               varRetVal = arr_varJType
10340             Else
10350               varRetVal = "MISC_2_GRP"
10360             End If
10370           End If
10380         Else
                ' ** 3 entries in hidden group, 1 Misc. and 2 other Assets (Yes, I believe they'll always be non-zero Assets).
10390           For lngX = 0& To (lngGrpCnt - 1&)
10400             arr_varJType(JT_TYPE, lngX) = "MISC_3_GRP"
10410           Next
10420           If blnTypeOnly = False Then
10430             varRetVal = arr_varJType
10440           Else
10450             varRetVal = "MISC_3_GRP"
10460           End If
10470         End If
10480       End If

10490     Else
            ' ** No numeral found?!
10500       If blnTypeOnly = False Then
10510         arr_varJType(JT_JTYP, 0) = RET_ERR
10520         varRetVal = arr_varJType
10530       Else
10540         varRetVal = vbNullString
10550       End If
10560     End If

10570   Else
          ' ** No UniqueID sent.
10580     If blnTypeOnly = False Then
10590       arr_varJType(JT_JTYP, 0) = RET_ERR
10600       varRetVal = arr_varJType
10610     Else
10620       varRetVal = vbNullString
10630     End If
10640   End If  ' ** IsNull(varUniqueID) = False.

EXITP:
10650   Hide_Type = varRetVal
10660   Exit Function

ERRH:
10670   If blnTypeOnly = False Then
10680     arr_varJType(JT_JTYP, 0) = RET_ERR
10690     varRetVal = arr_varJType
10700   Else
10710     varRetVal = vbNullString
10720   End If
10730   Select Case ERR.Number
        Case Else
10740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10750   End Select
10760   Resume EXITP

End Function

Public Function Hide_Count(varUniqueID As Variant) As Long
' ** Count the number of entries in the specified UniqueID.

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "Hide_Count"

        Dim intPos01 As Integer, intDblUnder As Integer
        Dim strTmp01 As String, lngTmp02 As Long
        Dim lngRetVal As Long

        ' ** Examples of [UniqueID] in tmpTrx and LedgerHidden:
        ' ** Simple pair matching:
        ' **   000000000000011_0012_001593_001563_Purchase__Sold_____ .
        ' **   000000000000011_0000_001697_001742_Misc._____Misc.____ .
        ' ** Misc. pair matching (1st level Misc group matching):
        ' **   000000000000011_0012_0000_001923_001972_Dividend__Misc.____ .
        ' ** Misc. triplet matching (2nd level Misc group matching):
        ' **   000000000000011_0001_0000_002081_002082_002083_Purchase__Sold______Misc.____ .
        ' ** Multi-lot matching (3rd level matching, this one's 174 characters long!):
        ' **   000000000000011_0055__002070_002071_002072_002073_002074_002075_002076_002077_002084_Sold______Sold______Sold______Sold______Sold______Sold______Sold______Sold______Purchase_ .

10810   lngRetVal = 0&

        ' ** Count the number of journalno's, between assetno and journaltype.
10820   If IsNull(varUniqueID) = False Then
10830     If Trim(varUniqueID) <> vbNullString Then
10840       strTmp01 = Trim(varUniqueID)
10850       intDblUnder = InStr(strTmp01, "__")  ' ** Some have it, some don't.
10860       intPos01 = InStr(strTmp01, "_")
10870       strTmp01 = Mid(strTmp01, (intPos01 + IIf(intPos01 = intDblUnder, 2, 1)))  ' ** Strip accountno.
10880       intDblUnder = InStr(strTmp01, "__")
10890       intPos01 = InStr(strTmp01, "_")
10900       strTmp01 = Mid(strTmp01, (intPos01 + IIf(intPos01 = intDblUnder, 2, 1)))  ' ** Strip assetno.
10910       lngTmp02 = 0&
10920       intDblUnder = InStr(strTmp01, "__")
10930       intPos01 = InStr(strTmp01, "_")
10940       Do While intPos01 > 0
10950         If IsNumeric(Left(strTmp01, (intPos01 - 1))) = True Then
10960           If Val(Left(strTmp01, (intPos01 - 1))) = 0 Then
                  ' ** Single journalno, with 2nd one all Zeroes.
10970             Exit Do
10980           Else
10990             lngTmp02 = lngTmp02 + 1&
11000             strTmp01 = Mid(strTmp01, (intPos01 + IIf(intPos01 = intDblUnder, 2, 1)))
11010             intDblUnder = InStr(strTmp01, "__")
11020             intPos01 = InStr(strTmp01, "_")
11030           End If
11040         Else
                ' ** We've reached the journaltype's.
11050           Exit Do
11060         End If
11070       Loop
11080       lngRetVal = lngTmp02
11090     End If
11100   End If

EXITP:
11110   Hide_Count = lngRetVal
11120   Exit Function

ERRH:
11130   lngRetVal = 0&
11140   Select Case ERR.Number
        Case Else
11150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11160   End Select
11170   Resume EXITP

End Function

Public Function Hide_Max() As Boolean
' ** Get the highest existing group number from LedgerHidden.

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "Hide_Max"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

11210   blnRetVal = True

11220   Set dbs = CurrentDb
11230   With dbs
          ' ** LedgerHidden, grouped, with Max(Hidden_Group).
11240     Set qdf = dbs.QueryDefs("qryAccountHide_43")
11250     Set rst = qdf.OpenRecordset
11260     With rst
11270       If .BOF = True And .EOF = True Then
11280         lngGroupMax = 0&
11290       Else
11300         If IsNull(![hid_grpnum]) = True Then
11310           lngGroupMax = 0&
11320         Else
11330           lngGroupMax = ![hid_grpnum]
11340         End If
11350       End If
11360       .Close
11370     End With
11380     Set rst = Nothing
11390     Set qdf = Nothing
11400     .Close
11410   End With
11420   Set dbs = Nothing
11430   DoEvents

EXITP:
11440   Set rst = Nothing
11450   Set qdf = Nothing
11460   Set dbs = Nothing
11470   Hide_Max = blnRetVal
11480   Exit Function

ERRH:
11490   blnRetVal = False
11500   Select Case ERR.Number
        Case Else
11510     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11520   End Select
11530   Resume EXITP

End Function

Public Function Hide_LoadArray(Optional varReturnArray As Variant) As Variant

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "Hide_LoadArray"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strLastUniqueID As String
        Dim lngGroup As Long
        Dim blnReturnArray As Boolean, blnProblem As Boolean
        Dim lngTmp01 As Long
        Dim lngX As Long
        Dim varRetVal As Variant

11610   blnProblem = False

11620   If IsMissing(varReturnArray) = True Then
11630     blnReturnArray = False
11640   Else
11650     blnReturnArray = CBool(varReturnArray)
11660   End If

        ' ** Get the highest existing group number from LedgerHidden.
11670   If lngHides = 0& Then
          ' ** lngGroupMax = 0& if there are no hidden transactions yet.
11680     Hide_Max  ' ** Function: Above.
11690   End If

11700   Set dbs = CurrentDb
11710   With dbs

          ' ** Populate the arr_varHide() array so that queries can access it via Hide_Group(), below.
          ' ** qryAccountHide_50a (hidden trx), linked to qryAccountHide_53 (all records, with count in each group).
11720     Set qdf = dbs.QueryDefs("qryAccountHide_54a")
11730     Set rst = qdf.OpenRecordset
11740     With rst
11750       If .BOF = True And .EOF = True Then
              ' ** Now that I've prevented opening the form if there are no hidden transactions,
              ' ** I don't believe this will get hit. However...
11760         blnProblem = True
11770       Else
11780         .MoveLast
11790         lngHides = .RecordCount
11800         .MoveFirst
11810         ReDim arr_varHide(H_ELEMS, (lngHides - 1&))
              ' *************************************************
              ' ** Array: arr_varHide()
              ' **
              ' **   Field  Element  Name            Constant
              ' **   =====  =======  ==============  ==========
              ' **     1       0     hid_grpnum      H_NUM
              ' **     2       1     cnt             H_CNT
              ' **     3       2     Misc. Group     H_MGRP
              ' **     4       3     accountno       H_ACTNO
              ' **     5       4     journalno       H_JNO
              ' **     6       5     journaltype     H_JTYPE
              ' **     7       6     UniqueIDx       H_UNIQ
              ' **     8       7     Unhidden        H_UNHID
              ' **     9       8     hidtype         H_GTYPE
              ' **    10       9     hid_sort        H_SORT
              ' **    11      10     hid_sortdate    H_SDATE
              ' **    12      11     hid_order       H_ORD
              ' **    13      12     Pre-existing    H_PREX
              ' **
              ' *************************************************
11820         strLastUniqueID = vbNullString
11830         lngTmp01 = 0&
11840         For lngX = 0& To (lngHides - 1&)
11850           If IsNull(![hid_grpnum]) = False Then
11860             arr_varHide(H_NUM, lngX) = ![hid_grpnum]
11870             arr_varHide(H_CNT, lngX) = ![cnt]
11880             arr_varHide(H_ACTNO, lngX) = gstrAccountNo
11890             arr_varHide(H_JNO, lngX) = ![journalno]
11900             arr_varHide(H_JTYPE, lngX) = ![journaltypex]
11910             arr_varHide(H_UNIQ, lngX) = ![uniqueid]
11920             arr_varHide(H_UNHID, lngX) = CBool(False)
11930             arr_varHide(H_GTYPE, lngX) = ![hidtype]
11940             arr_varHide(H_MGRP, lngX) = IIf(Left(arr_varHide(H_GTYPE, lngX), 4) = "MISC", True, False)
11950             arr_varHide(H_SORT, lngX) = ![hid_sort]
11960             arr_varHide(H_SDATE, lngX) = ![hid_sortdate]
11970             arr_varHide(H_ORD, lngX) = ![hid_order]
11980             arr_varHide(H_PREX, lngX) = CBool(True)
11990           Else
12000             If ![UniqueIDx] <> strLastUniqueID Then
                    ' ** New hidden group; continue the pre-existing hidden group sequence (if any).
12010               lngTmp01 = lngTmp01 + 1&
12020               lngGroup = lngGroupMax + lngTmp01
12030               strLastUniqueID = ![UniqueIDx]
12040             End If
12050             arr_varHide(H_NUM, lngX) = lngGroup
12060             arr_varHide(H_CNT, lngX) = ![cnt]
12070             arr_varHide(H_ACTNO, lngX) = ![accountno]
12080             arr_varHide(H_JNO, lngX) = ![journalno]
12090             arr_varHide(H_JTYPE, lngX) = ![journaltypex]
12100             arr_varHide(H_UNIQ, lngX) = ![UniqueIDx]
12110             arr_varHide(H_UNHID, lngX) = CBool(False)
12120             arr_varHide(H_GTYPE, lngX) = Hide_Type(![UniqueIDx], True)  ' ** Function: Above.
12130             arr_varHide(H_MGRP, lngX) = IIf(Left(arr_varHide(H_GTYPE, lngX), 4) = "MISC", True, False)
12140             arr_varHide(H_SORT, lngX) = ![Sortx]
12150             arr_varHide(H_SDATE, lngX) = ![SortDate]
12160             arr_varHide(H_ORD, lngX) = ![ord]
12170             arr_varHide(H_PREX, lngX) = CBool(False)
12180           End If
12190           If lngX < (lngHides - 1&) Then .MoveNext
12200         Next
12210       End If
12220       .Close
12230     End With
12240     Set rst = Nothing
12250     Set qdf = Nothing
12260     DoEvents
12270     .Close
12280   End With
12290   Set dbs = Nothing
12300   DoEvents

12310   If blnReturnArray = False Then
12320     If blnProblem = False Then
12330       varRetVal = CBool(True)
12340     Else
12350       varRetVal = CBool(False)
12360     End If
12370   Else
12380     varRetVal = arr_varHide
12390   End If

EXITP:
12400   Set rst = Nothing
12410   Set qdf = Nothing
12420   Set dbs = Nothing
12430   Hide_LoadArray = varRetVal
12440   Exit Function

ERRH:
12450   If blnReturnArray = False Then
12460     varRetVal = CBool(False)
12470   Else
12480     ReDim arr_varHide(H_ELEMS, 0)
12490     arr_varHide(0, 0) = RET_ERR
12500     varRetVal = arr_varHide
12510   End If
12520   Select Case ERR.Number
        Case Else
12530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12540   End Select
12550   Resume EXITP

End Function

Public Function Hide_RenumGroups() As Boolean

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "Hide_RenumGroups"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim strAccountNo As String, strLastUniqueID As String
        Dim lngLastRealGrpNum As Long, lngThisRealGrpNum As Long, lngLastNewGrpNum As Long, lngThisNewGrpNum As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

12610   blnRetVal = True

12620   Set dbs = CurrentDb
12630   With dbs
12640     Set qdf = .QueryDefs("qryAccountHide_64")
12650     Set rst = qdf.OpenRecordset
12660     With rst
12670       If .BOF = True And .EOF = True Then
              ' ** Empty!
12680       Else
12690         .MoveLast
12700         lngRecs = .RecordCount
12710         .MoveFirst
12720         strAccountNo = vbNullString
12730         lngLastRealGrpNum = 0&: lngThisRealGrpNum = 0&
12740         lngLastNewGrpNum = 0&: lngThisNewGrpNum = 0&
12750         For lngX = 1& To lngRecs
12760           If ![accountno] <> strAccountNo Then
                  ' ** Start each account on an odd number.
12770             strAccountNo = ![accountno]
12780             If lngLastRealGrpNum = 0& Then
12790               lngLastRealGrpNum = ![hid_grpnum]
12800               lngLastNewGrpNum = 1&
12810               strLastUniqueID = ![uniqueid]
12820             Else
12830               If (lngLastNewGrpNum Mod 2) > 0 Then        ' ** Meaning the last one from the previous account was odd.
12840                 lngLastNewGrpNum = lngLastNewGrpNum + 2&  ' ** Make this one odd.
12850               Else
12860                 lngLastNewGrpNum = lngLastNewGrpNum + 1&  ' ** Last was even, so this one's odd.
12870               End If
12880             End If
12890           End If
12900           If ![hid_grpnum] = lngLastRealGrpNum Then
12910             .Edit
12920             ![hid_grpnum] = lngLastNewGrpNum
12930             .Update
12940           Else
12950             strLastUniqueID = ![uniqueid]
12960             lngLastRealGrpNum = ![hid_grpnum]
12970             lngLastNewGrpNum = lngLastNewGrpNum + 1&
12980             .Edit
12990             ![hid_grpnum] = lngLastNewGrpNum
13000             .Update
13010           End If
13020           If lngX < lngRecs Then .MoveNext
13030         Next
13040       End If
13050     End With
13060     .Close
13070   End With

        'Beep

EXITP:
13080   Set rst = Nothing
13090   Set qdf = Nothing
13100   Set dbs = Nothing
13110   Hide_RenumGroups = blnRetVal
13120   Exit Function

ERRH:
13130   blnRetVal = False
13140   Select Case ERR.Number
        Case Else
13150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13160   End Select
13170   Resume EXITP

End Function

Public Sub LedgerHiddenLoad(frm As Access.Form)
' ** Well, I certainly didn't expect this to grow so big!

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "LedgerHiddenLoad"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngHids As Long, lngOldHidRecs As Long, lngNewHidRecs As Long, lngGroupNones As Long
        Dim lngGrps As Long, arr_varGrp() As Variant
        Dim lngNones As Long, arr_varNone() As Variant
        Dim lngUniques As Long, arr_varUnique() As Variant
        Dim lngActNos As Long, arr_varActNo As Variant
        Dim lngHidType As Long, strUniqueID As String, lngGrpNum As Long
        Dim blnLoad As Boolean, blnFound As Boolean
        Dim intHideSetupResponse As Integer, intMode As Integer
        Dim lngRecs As Long
        Dim varTmp00 As Variant, lngTmp01 As Long, lngTmp02 As Long
        Dim dblZ As Double
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long, lngF As Long

        ' ** Array: arr_varGrp().
        Const G_ELEMS As Integer = 11  ' ** Array's first-element UBound().
        Const G_ACTNO As Integer = 0
        Const G_GRP   As Integer = 1
        Const G_CNT   As Integer = 2
        Const G_ICSH  As Integer = 3
        Const G_PCSH  As Integer = 4
        Const G_COST  As Integer = 5
        Const G_JNO1  As Integer = 6
        Const G_JNO2  As Integer = 7
        Const G_JNO3  As Integer = 8
        Const G_ANO1  As Integer = 9
        Const G_ANO2  As Integer = 10
        Const G_FND   As Integer = 11

        ' ** Array: arr_varNone().
        Const N_ELEMS As Integer = 7  ' ** Array's first-element UBound().
        Const N_ACTNO As Integer = 0
        Const N_JNO   As Integer = 1
        Const N_ANO   As Integer = 2
        Const N_ICSH  As Integer = 3
        Const N_PCSH  As Integer = 4
        Const N_COST  As Integer = 5
        Const N_GRP   As Integer = 6
        Const N_FND   As Integer = 7

        ' ** Array: arr_varUnique().
        Const U_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const U_ACTNO As Integer = 0
        Const U_GRP   As Integer = 1
        Const U_CNT   As Integer = 2

        ' ** Array: arr_varActNo().
        Const A_ACTNO As Integer = 0
        'Const A_HID   As Integer = 1

13210   With frm

13220     blnLoad = False
13230     lngHids = 0&: lngOldHidRecs = 0&: lngNewHidRecs = 0&

          ' ** Only run if this accountno has hidden transactions.
          ' ** If other accounts have them, this will come later.
13240     If .hidden_trans > 0& Then

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
13250       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

            ' ** Do these just because they annoy me!
13260       With .frmAccountHideTrans2_Sub_Pick.Form
13270         .curr_id_lbl.ForeColor = WIN_CLR_DISF
13280         .curr_id_lbl_line.BorderColor = WIN_CLR_DISR
13290         Select Case .Parent.chkIncludeCurrency
              Case True
13300           .curr_id_lbl_dim_hi.Visible = True
13310           .curr_id_lbl_line_dim_hi.Visible = True
13320         Case False
13330           .curr_id_lbl_dim_hi.Visible = False
13340           .curr_id_lbl_line_dim_hi.Visible = False
13350         End Select
13360         .journalno_lbl.ForeColor = WIN_CLR_DISF
13370         .journalno_lbl2.ForeColor = WIN_CLR_DISF
13380         .journalno_lbl_line.BorderColor = WIN_CLR_DISR
13390         Select Case .Parent.chkShowJournalNo
              Case True
13400           .journalno_lbl_dim_hi.Visible = True
13410           .journalno_lbl2_dim_hi.Visible = True
13420           .journalno_lbl_line_dim_hi.Visible = True
13430         Case False
13440           .journalno_lbl_dim_hi.Visible = False
13450           .journalno_lbl2_dim_hi.Visible = False
13460           .journalno_lbl_line_dim_hi.Visible = False
13470         End Select
13480         DoEvents
13490       End With
13500       .chkIncludeArchive.Enabled = False
13510       .chkIncludeCurrency.Enabled = False
13520       .chkShowAll.Enabled = False
13530       .chkShowJournalNo.Enabled = False
13540       .chkShowHiddenOnly.Enabled = False
13550       DoEvents

            ' ** Percentage labels must be left-aligned, with centering done by spaces.
            ' ** This is so that the front label's width (white Forecolor) can expand with the
            ' ** blue bar, revealing the white letters only as the bar approaches half-way.
            'strSp = Space(60)

            ' ** Initialize the progress bar.
13560       dblPB_Steps = 29#
13570       ReDim arr_dblPB_ThisIncr(dblPB_Steps)  ' ** Since arrays are zero-based, this one will only use 1-29, and not 0.
13580       dblPB_Width = .ProgBar_box.Width

            ' ** Weight the steps.
13590       arr_dblPB_ThisIncr(1) = CDbl((dblPB_Width / 100#) * 1#)   ' 1:  Initialization.
13600       arr_dblPB_ThisIncr(2) = CDbl((dblPB_Width / 100#) * 24#)  ' 2:  Original Hide_Setup(), once for each account.
13610       arr_dblPB_ThisIncr(3) = CDbl((dblPB_Width / 100#) * 1#)   ' 3:  Transfer To tblLedgerHidden.
13620       arr_dblPB_ThisIncr(4) = CDbl((dblPB_Width / 100#) * 1#)   ' 4:  Check for existing Groups.
13630       arr_dblPB_ThisIncr(5) = CDbl((dblPB_Width / 100#) * 1#)   ' 5:  Begin processing unmatched hiddens.
13640       arr_dblPB_ThisIncr(6) = CDbl((dblPB_Width / 100#) * 1#)   ' 6:  Process unmatched, pass 1.
13650       arr_dblPB_ThisIncr(7) = CDbl((dblPB_Width / 100#) * 1#)   ' 7:  Process unmatched, pass 2.
13660       arr_dblPB_ThisIncr(8) = CDbl((dblPB_Width / 100#) * 1#)   ' 8:  Process old groups, pass 1.
13670       arr_dblPB_ThisIncr(9) = CDbl((dblPB_Width / 100#) * 1#)   ' 9:  Process old groups, pass 2.
13680       arr_dblPB_ThisIncr(10) = CDbl((dblPB_Width / 100#) * 1#)  ' 10: Begin processing remaining unmatched.
13690       arr_dblPB_ThisIncr(11) = CDbl((dblPB_Width / 100#) * 1#)  ' 11: Process remaining unmatched, pass 1.
13700       arr_dblPB_ThisIncr(12) = CDbl((dblPB_Width / 100#) * 1#)  ' 12: Process remaining unmatched, pass 2.
13710       arr_dblPB_ThisIncr(13) = CDbl((dblPB_Width / 100#) * 1#)  ' 13: Process remaining unmatched, pass 3.        26  1's (26)
13720       arr_dblPB_ThisIncr(14) = CDbl((dblPB_Width / 100#) * 1#)  ' 14: Process remaining unmatched, pass 4.        20  20's (1)
13730       arr_dblPB_ThisIncr(15) = CDbl((dblPB_Width / 100#) * 1#)  ' 15: Process remaining unmatched, pass 5.        24  24's (1)
13740       arr_dblPB_ThisIncr(16) = CDbl((dblPB_Width / 100#) * 1#)  ' 16: Begin processing new hiddens.               30  50's (1)
13750       arr_dblPB_ThisIncr(17) = CDbl((dblPB_Width / 100#) * 1#)  ' 17: Process new hiddens, pass 1.               ===
13760       arr_dblPB_ThisIncr(18) = CDbl((dblPB_Width / 100#) * 1#)  ' 18: Process new hiddens, pass 2.               100%
13770       arr_dblPB_ThisIncr(19) = CDbl((dblPB_Width / 100#) * 1#)  ' 19: Process new hiddens, pass 3.
13780       arr_dblPB_ThisIncr(20) = CDbl((dblPB_Width / 100#) * 30#) ' 20: Process new hiddens, pass 4.  * THE REALLY LONG ONE!
13790       arr_dblPB_ThisIncr(21) = CDbl((dblPB_Width / 100#) * 1#)  ' 21: Check new hiddens.
13800       arr_dblPB_ThisIncr(22) = CDbl((dblPB_Width / 100#) * 1#)  ' 22: Determine hidden type.
13810       arr_dblPB_ThisIncr(23) = CDbl((dblPB_Width / 100#) * 1#)  ' 23: Update tables 1.
13820       arr_dblPB_ThisIncr(24) = CDbl((dblPB_Width / 100#) * 1#)  ' 24: Update tables 2.
13830       arr_dblPB_ThisIncr(25) = CDbl((dblPB_Width / 100#) * 1#)  ' 25: Update tables 3.
13840       arr_dblPB_ThisIncr(26) = CDbl((dblPB_Width / 100#) * 1#)  ' 26: Update tables 4.
13850       arr_dblPB_ThisIncr(27) = CDbl((dblPB_Width / 100#) * 20#) ' 27: Additional matching in Hide_AddlMatch().
13860       arr_dblPB_ThisIncr(28) = CDbl((dblPB_Width / 100#) * 1#)  ' 28: Renumber groups in Hide_RenumGroups2().
13870       arr_dblPB_ThisIncr(29) = CDbl((dblPB_Width / 100#) * 1#)  ' 29: Load finished.

            ' ** Double-check whether full loading should be done.
13880       varTmp00 = DCount("*", "tblLedgerHidden")
13890       DoEvents
13900       If IsNull(varTmp00) = False Then
13910         If varTmp00 > 0 Then
                ' ** Now compare to number of hids.
13920           If varTmp00 = .hidden_trans Then
                  ' ** Wow! Complete!
13930             If .chkHiddenFirstUse = False Then .chkHiddenFirstUse = True
13940           Else
                  ' ** Use 50% as the threshhold.
13950             If (varTmp00 / .hidden_trans) < 0.5 Then
13960               If .chkHiddenFirstUse = True Then .chkHiddenFirstUse = False
13970             Else
13980               If .chkHiddenFirstUse = False Then .chkHiddenFirstUse = True
13990             End If
14000           End If
14010         Else
14020           If .chkHiddenFirstUse = True Then .chkHiddenFirstUse = False
14030         End If
14040       Else
14050         If .chkHiddenFirstUse = True Then .chkHiddenFirstUse = False
14060       End If
14070       DoEvents

            ' ***************************************************************
            ' ** Step 1: Initialization.
14080       dblPB_ThisStep = 1#
14090       .Status2_lbl.Caption = "Initialization"
14100       .Status2_lbl.Visible = True
14110       If .chkHiddenFirstUse = False Then
14120         .Status1_lbl.Visible = True
14130       End If
14140       DoEvents
            ' ***************************************************************
            ' ***************************************************************
14150       dblPB_ThisWidth = 0#
14160       For dblZ = 1# To (dblPB_ThisStep - 1#)
              ' ** Assemble the weighted widths up to, but not including, this width.
14170         dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
14180       Next
14190       dblPB_StepSubs = 0#  ' ** No subs in this step.
14200       dblPB_ThisIncrSub = 0#
14210       dblPB_ThisStepSub = 0#
14220       ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
            '.ProgBar_bar.Width = dblPB_ThisWidth
14230       strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
14240       .ProgBar_lbl1.Caption = strPB_ThisPct

14250       DoEvents
            ' ***************************************************************

14260       .ProgBar_box.Visible = True
14270       .ProgBar_box2.Visible = True
14280       ProgBar_Width_Hide frm, True, 1  ' ** Procedure: Below.
            '.ProgBar_bar.Visible = True
14290       .ProgBar_lbl1.Visible = True
14300       DoEvents

14310       varTmp00 = DCount("*", "tblLedgerHidden")
14320       Select Case IsNull(varTmp00)
            Case True
14330         blnLoad = True
14340       Case False
14350         If varTmp00 = 0 Then
14360           blnLoad = True
14370         Else
14380           varTmp00 = DCount("*", "tblLedgerHidden", "[accountno] = '" & .accountno & "'")
14390           Select Case IsNull(varTmp00)
                Case True
14400             blnLoad = True
14410           Case False
14420             If varTmp00 <> .hidden_trans Then
14430               blnLoad = True
14440             End If
14450           End Select
14460         End If
14470       End Select
14480       DoEvents

            ' ** Only run if tblLedgerHidden is empty, or if this account's
            ' ** hidden transactions are not all in tblLedgerHidden.
            ' ** Once triggered, all hidden transactions will be loaded.
14490       If blnLoad = True Then

              ' ** If we're going to load any, load all.
14500         varTmp00 = DCount("*", "LedgerHidden")
14510         lngOldHidRecs = Nz(varTmp00, 0)
14520         varTmp00 = DCount("*", "tblLedgerHidden")
14530         lngNewHidRecs = Nz(varTmp00, 0)
              ' ** qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a (Ledger, just needed fields),
              ' ** qryAccountHideTrans2_24b (LedgerArchive, just needed fields)), just ledger_HIDDEN = True.
14540         varTmp00 = DCount("*", "qryAccountHideTrans2_25")
14550         lngHids = varTmp00  ' ** Total Hidden, including Archive.

14560         If lngNewHidRecs = 0& And lngOldHidRecs = 0& Then                ' ** 1. No tblLedgerHidden   : No LedgerHidden
                ' ** Load all the old way, then transfer.
14570           intMode = 1
14580         ElseIf lngNewHidRecs = 0& And lngOldHidRecs < lngHids Then       ' ** 2. No tblLedgerHidden   : Some LedgerHidden
                ' ** Transfer, then add.
14590           intMode = 2
14600         ElseIf lngNewHidRecs = 0& And lngOldHidRecs = lngHids Then       ' ** 3. No tblLedgerHidden   : All LedgerHidden
                ' ** Transfer all.
14610           intMode = 3
14620         ElseIf lngNewHidRecs < lngHids And lngOldHidRecs = 0& Then       ' ** 4. Some tblLedgerHidden : No LedgerHidden
                ' ** Add.
14630           intMode = 4
14640         ElseIf lngNewHidRecs < lngHids And lngOldHidRecs < lngHids Then  ' ** 5. Some tblLedgerHidden : Some LedgerHidden
                ' ** Transfer any available, then add.
14650           intMode = 5
14660         ElseIf lngNewHidRecs < lngHids And lngOldHidRecs = lngHids Then  ' ** 6. Some tblLedgerHidden : All LedgerHidden
                ' ** Transfer missing.
14670           intMode = 6
14680         Else                                                             ' ** 7. All tblLedgerHidden  : NOT HERE!
                ' ** Shouldn't be here!
14690         End If
14700         DoEvents

14710         Set dbs = CurrentDb

              ' ** LedgerHidden is empty, so use old method to load it,
              ' ** then transfer that to tblLedgerHidden.
14720         If intMode = 1 Then

                ' ** qryAccountHideTrans2_25 (qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a
                ' ** (Ledger, just needed fields), qryAccountHideTrans2_24b (LedgerArchive, just needed fields)),
                ' ** just ledger_HIDDEN = True), grouped by accountno, with cnt_jno.
14730           Set qdf = dbs.QueryDefs("qryAccountHideTrans2_25_01")
14740           Set rst = qdf.OpenRecordset
14750           With rst
14760             .MoveLast
14770             lngActNos = .RecordCount
14780             .MoveFirst
14790             arr_varActNo = .GetRows(lngActNos)
                  ' **********************************************
                  ' ** Array: arr_varActNo()
                  ' **
                  ' **   Field  Element  Name         Constant
                  ' **   =====  =======  ===========  ==========
                  ' **     1       0     accountno    A_ACTNO
                  ' **     2       1     cnt_hid      A_HID
                  ' **
                  ' **********************************************
14800             .Close
14810           End With
14820           Set rst = Nothing
14830           Set qdf = Nothing

                ' ***************************************************************
                ' ** Step 2: Original hidden setup.
14840           dblPB_ThisStep = 2#
14850           .Status2_lbl.Caption = "Original hidden setup"
14860           DoEvents
                ' ***************************************************************
                ' ***************************************************************
14870           dblPB_ThisWidth = 0#
14880           For dblZ = 1# To (dblPB_ThisStep - 1#)
                  ' ** Assemble the weighted widths up to, but not including, this width.
14890             dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
14900           Next
14910           dblPB_StepSubs = (8# * lngActNos)
14920           dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
14930           dblPB_ThisStepSub = 0#
14940           DoEvents
                ' ***************************************************************

                ' ** The easiest way might be to use the old load, then transfer.
                ' ** Hide_Setup() has 8 Steps.
                ' ** This only does what's in gstrAccountNo.

                ' ** Let's see if this'll work in a loop of all accounts.
14950           lngTmp01 = gstrAccountNo  ' ** The current accountno.
14960           lngTmp02 = 0&
                ' ** The whole Step #2 has 24%.
                ' ** Let's see if we can divide that among all the Hide_Setup() calls.
                ' ** We've got (8 * lngActNos) substeps.
14970           For lngX = 0& To (lngActNos - 1&)
14980             gstrAccountNo = arr_varActNo(A_ACTNO, lngX)
14990             intHideSetupResponse = Hide_Setup(True, frm, dblPB_ThisIncrSub, (lngX + 1&))    ' ** Function: Above.
                  ' **  0  No problem.
                  ' ** -1  No hidden transactions.
                  ' ** -9  No way! (Error)
15000             DoEvents
15010             If intHideSetupResponse <> 0 Then
15020               If intHideSetupResponse < lngTmp02 Then
15030                 lngTmp02 = intHideSetupResponse
15040               End If
15050               intHideSetupResponse = 0
15060             End If
15070           Next  ' ** lngX.
15080           gstrAccountNo = lngTmp01
15090           If lngTmp02 = -9 Then
                  ' ** What should we do if one of them errors?

15100           End If

15110           If intHideSetupResponse = 0 Then
                  ' ** Append qryAccountHideTrans2_51 (LedgerHidden, linked to qryAccountHideTrans2_25
                  ' ** (qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a (Ledger, just needed
                  ' ** fields), qryAccountHideTrans2_24b (LedgerArchive, just needed fields)), just
                  ' ** ledger_HIDDEN = True), qryAccountHideTrans2_50 (LedgerHidden, grouped by hid_grpnum,
                  ' ** with cnt), without GRP_NONE, as tblLedgerHidden records) to tblLedgerHidden.
15120             Set qdf = dbs.QueryDefs("qryAccountHideTrans2_52")
15130             qdf.Execute
15140             Set qdf = Nothing
                  ' ** LedgerHidden, just 'GRP_NONE'.
15150             varTmp00 = DCount("*", "qryAccountHideTrans2_53")
15160             lngGroupNones = Nz(varTmp00, 0)
15170           End If
15180         End If
15190         DoEvents

              ' ***************************************************************
              ' ** Step 3: Transfer To tblLedgerHidden.
15200         dblPB_ThisStep = 3#
15210         .Status2_lbl.Caption = "Transfer To tblLedgerHidden"
15220         DoEvents
              ' ***************************************************************
              ' ***************************************************************
15230         dblPB_ThisWidth = 0#
15240         For dblZ = 1# To (dblPB_ThisStep - 1#)
                ' ** Assemble the weighted widths up to, but not including, this width.
15250           dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
15260         Next
15270         dblPB_StepSubs = 0#  ' ** No subs in this step.
15280         dblPB_ThisIncrSub = 0#
15290         dblPB_ThisStepSub = 0#
15300         ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
              '.ProgBar_bar.Width = dblPB_ThisWidth
15310         strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
15320         .ProgBar_lbl1.Caption = strPB_ThisPct
15330         DoEvents
              ' ***************************************************************

              ' ** Transfer whatever is in LedgerHidden to tblLedgerHidden.
15340         If intMode = 2 Or intMode = 3 Or intMode = 5 Or intMode = 6 Then
                ' ** Append qryAccountHideTrans2_51 (LedgerHidden, linked to qryAccountHideTrans2_25
                ' ** (qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a (Ledger, just needed
                ' ** fields), qryAccountHideTrans2_24b (LedgerArchive, just needed fields)), just
                ' ** ledger_HIDDEN = True), qryAccountHideTrans2_50 (LedgerHidden, grouped by hid_grpnum,
                ' ** with cnt), without GRP_NONE, as tblLedgerHidden records) to tblLedgerHidden.
15350           Set qdf = dbs.QueryDefs("qryAccountHideTrans2_52")
15360           qdf.Execute
15370           Set qdf = Nothing
                ' ** LedgerHidden, just 'GRP_NONE'.
15380           varTmp00 = DCount("*", "qryAccountHideTrans2_53")
15390           lngGroupNones = Nz(varTmp00, 0)
15400         End If
15410         DoEvents

              ' ** LedgerHidden had some unmatched hidden transactions,
              ' ** so try now to match them up.
15420         If lngGroupNones > 0& Then
                ' ** Let's see if we can figure these out.
                ' ** The reason there are GRP_NONE's is because it was hard to do
                ' ** the matching after the fact, and I sometimes just had to give up.

15430           lngGrps = 0&
15440           ReDim arr_varGrp(G_ELEMS, 0)
15450           lngNones = 0&
15460           ReDim arr_varNone(N_ELEMS, 0)

                ' ***************************************************************
                ' ** Step 4: Check for existing Groups.
15470           dblPB_ThisStep = 4#
15480           .Status2_lbl.Caption = "Check for existing Groups"
15490           DoEvents
                ' ***************************************************************
                ' ***************************************************************
15500           dblPB_ThisWidth = 0#
15510           For dblZ = 1# To (dblPB_ThisStep - 1#)
                  ' ** Assemble the weighted widths up to, but not including, this width.
15520             dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
15530           Next
                ' ***************************************************************

                ' ** qryAccountHideTrans2_54a (tblLedgerHidden, linked to qryAccountHideTrans2_24 (Union
                ' ** of qryAccountHideTrans2_24a (Ledger, just needed fields), qryAccountHideTrans2_24b
                ' ** (LedgerArchive, just needed fields)), grouped and summed, by ledhid_grpnum), rounded.
15540           Set qdf = dbs.QueryDefs("qryAccountHideTrans2_54b")
15550           Set rst = qdf.OpenRecordset
15560           With rst
15570             If .BOF = True And .EOF = True Then
                    ' ** I guess all of them were GRP_NONE's!
                    ' ***************************************************************
15580               dblPB_StepSubs = 0#  ' ** No subs in this step.
15590               dblPB_ThisIncrSub = 0#
15600               dblPB_ThisStepSub = 0#
15610               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    'frm.ProgBar_bar.Width = dblPB_ThisWidth
15620               strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
15630               frm.ProgBar_lbl1.Caption = strPB_ThisPct
15640               DoEvents
                    ' ***************************************************************
15650             Else
15660               .MoveLast
15670               lngRecs = .RecordCount
15680               .MoveFirst
                    ' ***************************************************************
15690               dblPB_StepSubs = lngRecs
15700               dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
15710               dblPB_ThisStepSub = 0#
15720               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    'frm.ProgBar_bar.Width = dblPB_ThisWidth
15730               DoEvents
                    ' ***************************************************************
15740               For lngX = 1& To lngRecs
15750                 If ![ICash] <> 0@ Or ![PCash] <> 0@ Or ![Cost] <> 0@ Then
15760                   lngGrps = lngGrps + 1&
15770                   lngE = lngGrps - 1&
15780                   ReDim Preserve arr_varGrp(G_ELEMS, lngE)
15790                   arr_varGrp(G_ACTNO, lngE) = ![accountno]
15800                   arr_varGrp(G_GRP, lngE) = ![ledghid_grpnum]
15810                   arr_varGrp(G_CNT, lngE) = ![cnt]
15820                   arr_varGrp(G_ICSH, lngE) = ![ICash]
15830                   arr_varGrp(G_PCSH, lngE) = ![PCash]
15840                   arr_varGrp(G_COST, lngE) = ![Cost]
15850                   arr_varGrp(G_JNO1, lngE) = Null
15860                   arr_varGrp(G_JNO2, lngE) = Null
15870                   arr_varGrp(G_JNO3, lngE) = Null
15880                   arr_varGrp(G_ANO1, lngE) = ![assetno1]
15890                   arr_varGrp(G_ANO2, lngE) = ![assetno2]
15900                   arr_varGrp(G_FND, lngE) = CBool(False)
15910                 End If
                      ' ***************************************************************
15920                 dblPB_ThisStepSub = lngX
15930                 dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
15940                 ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                      'frm.ProgBar_bar.Width = dblPB_ThisWidthSub
15950                 strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
15960                 frm.ProgBar_lbl1.Caption = strPB_ThisPct
15970                 DoEvents
                      ' ***************************************************************
15980                 If lngX < lngRecs Then .MoveNext
15990               Next  ' ** lngX.
16000             End If
16010             .Close
16020           End With  ' ** rst.
16030           Set rst = Nothing
16040           Set qdf = Nothing
16050           DoEvents

                ' ** These are old groups that don't add up,
                ' ** so there should also be GRP_NONE's in LedgerHidden.
                ' ** (Unless they somehow got hidden erroneously!)
16060           If lngGrps > 0& Then
                  ' ** lngGrps are the number of groups that aren't complete.

                  ' ***************************************************************
                  ' ** Step 5: Begin processing unmatched hidden.
16070             dblPB_ThisStep = 5#
16080             .Status2_lbl.Caption = "Begin processing unmatched hidden"
16090             DoEvents
                  ' ***************************************************************
                  ' ***************************************************************
16100             dblPB_ThisWidth = 0#
16110             For dblZ = 1# To (dblPB_ThisStep - 1#)
                    ' ** Assemble the weighted widths up to, but not including, this width.
16120               dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
16130             Next
                  ' ***************************************************************

                  ' ** LedgerHidden, just 'GRP_NONE'.
16140             Set rst = dbs.OpenRecordset("qryAccountHideTrans2_53", dbOpenDynaset, dbConsistent)
16150             With rst
16160               .MoveLast
16170               lngRecs = .RecordCount
16180               .MoveFirst
                    ' ***************************************************************
16190               dblPB_StepSubs = lngRecs
16200               dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
16210               dblPB_ThisStepSub = 0#
16220               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    'frm.ProgBar_bar.Width = dblPB_ThisWidth
16230               DoEvents
                    ' ***************************************************************
16240               For lngX = 1& To lngRecs
16250                 lngNones = lngNones + 1&
16260                 lngE = lngNones - 1&
16270                 ReDim Preserve arr_varNone(N_ELEMS, lngE)
16280                 arr_varNone(N_ACTNO, lngE) = ![accountno]
16290                 arr_varNone(N_JNO, lngE) = ![journalno]
16300                 arr_varNone(N_ANO, lngE) = ![assetno]
16310                 arr_varNone(N_ICSH, lngE) = CCur(Round(![ICash], 2))
16320                 arr_varNone(N_PCSH, lngE) = CCur(Round(![PCash], 2))
16330                 arr_varNone(N_COST, lngE) = CCur(Round(![Cost], 2))
16340                 arr_varNone(N_GRP, lngE) = Null
16350                 arr_varNone(N_FND, lngE) = CBool(False)
                      ' ***************************************************************
16360                 dblPB_ThisStepSub = lngX
16370                 dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
16380                 ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                      'frm.ProgBar_bar.Width = dblPB_ThisWidthSub
16390                 strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
16400                 frm.ProgBar_lbl1.Caption = strPB_ThisPct
16410                 DoEvents
                      ' ***************************************************************
16420                 If lngX < lngRecs Then .MoveNext
16430               Next  ' ** lngX.
16440               .Close
16450             End With  ' ** rst.
16460             Set rst = Nothing
16470             Set qdf = Nothing
16480             DoEvents

                  ' ***************************************************************
                  ' ** Step 6: Processing unmatched - pass 1.
16490             dblPB_ThisStep = 6#
16500             .Status2_lbl.Caption = "Processing unmatched - pass 1"
16510             DoEvents
                  ' ***************************************************************
                  ' ***************************************************************
16520             dblPB_ThisWidth = 0#
16530             For dblZ = 1# To (dblPB_ThisStep - 1#)
                    ' ** Assemble the weighted widths up to, but not including, this width.
16540               dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
16550             Next
                  ' ***************************************************************

                  ' ** This tries to match old GRP_NONE's with incomplete old groups.
16560             lngTmp01 = 0&
                  ' ***************************************************************
16570             dblPB_StepSubs = lngGrps
16580             dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
16590             dblPB_ThisStepSub = 0#
16600             ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                  '.ProgBar_bar.Width = dblPB_ThisWidth
16610             DoEvents
                  ' ***************************************************************
16620             For lngX = 0& To (lngGrps - 1&)
                    ' ** We'll start with adding just 1 transaction from the GRP_NONE's.
16630               For lngY = 0& To (lngNones - 1&)
16640                 If arr_varNone(N_FND, lngY) = False Then
16650                   If arr_varNone(N_ACTNO, lngY) = arr_varGrp(G_ACTNO, lngX) Then
                          ' ** Make sure accountno's match.
16660                     If ((arr_varGrp(G_ANO1, lngX) = 0& And arr_varGrp(G_ANO2, lngX) = 0&) Or _
                              (arr_varGrp(G_ANO1, lngX) = 0& And arr_varNone(N_ANO, lngY) = 0&) Or _
                              (arr_varGrp(G_ANO2, lngX) = 0& And arr_varNone(N_ANO, lngY) = 0&) Or _
                              (arr_varGrp(G_ANO1, lngX) = 0& And arr_varGrp(G_ANO2, lngX) = arr_varNone(N_ANO, lngY)) Or _
                              (arr_varGrp(G_ANO2, lngX) = 0& And arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngY)) Or _
                              (arr_varNone(N_ANO, lngY) = 0& And arr_varGrp(G_ANO1, lngX) = arr_varGrp(G_ANO2, lngX)) Or _
                              (arr_varGrp(G_ANO1, lngX) = arr_varGrp(G_ANO2, lngX) And arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngY))) Then
                            ' ** Make sure assetno's are appropriate.
16670                       If (((arr_varGrp(G_ICSH, lngX) + arr_varNone(N_ICSH, lngY)) = 0@) And _
                                ((arr_varGrp(G_PCSH, lngX) + arr_varNone(N_PCSH, lngY)) = 0@) And _
                                ((arr_varGrp(G_COST, lngX) + arr_varNone(N_COST, lngY)) = 0@)) Then
                              ' ** Aha! Found one!
16680                         arr_varGrp(G_JNO1, lngX) = arr_varNone(N_JNO, lngY)
16690                         arr_varGrp(G_FND, lngX) = CBool(True)
16700                         arr_varNone(N_GRP, lngY) = arr_varGrp(G_GRP, lngX)
16710                         arr_varNone(N_FND, lngY) = CBool(True)
16720                         lngTmp01 = lngTmp01 + 1&
16730                         Exit For
16740                       End If  ' ** icash, pcash, cost.
16750                     End If  ' ** assetno.
16760                   End If  ' ** accountno.
16770                 End If  ' ** N_FND.
16780               Next  ' ** lngY.
                    ' ***************************************************************
16790               dblPB_ThisStepSub = lngX
16800               dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
16810               ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                    '.ProgBar_bar.Width = dblPB_ThisWidthSub
16820               strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
16830               .ProgBar_lbl1.Caption = strPB_ThisPct
16840               DoEvents
                    ' ***************************************************************
16850             Next  ' ** lngX.
16860             DoEvents

                  ' ***************************************************************
                  ' ** Step 7: Processing unmatched - pass 2.
16870             dblPB_ThisStep = 7#
16880             .Status2_lbl.Caption = "Processing unmatched - pass 2"
16890             DoEvents
                  ' ***************************************************************
                  ' ***************************************************************
16900             dblPB_ThisWidth = 0#
16910             For dblZ = 1# To (dblPB_ThisStep - 1#)
                    ' ** Assemble the weighted widths up to, but not including, this width.
16920               dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
16930             Next
                  ' ***************************************************************

                  ' ** This further tries to match old GRP_NONE's with incomplete old groups.
16940             If lngTmp01 < lngGrps Then
                    ' ** Alright, let's try two.
                    ' ***************************************************************
16950               dblPB_StepSubs = lngGrps
16960               dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
16970               dblPB_ThisStepSub = 0#
16980               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    '.ProgBar_bar.Width = dblPB_ThisWidth
16990               DoEvents
                    ' ***************************************************************
17000               For lngX = 0& To (lngGrps - 1&)
17010                 If arr_varGrp(G_FND, lngX) = False Then
17020                   For lngY = 0& To (lngNones - 1&)
17030                     If arr_varNone(N_FND, lngY) = False Then
17040                       If arr_varNone(N_ACTNO, lngY) = arr_varGrp(G_ACTNO, lngX) Then
                              ' ** Make sure accountno's match.
17050                         For lngZ = 0& To (lngNones - 1&)
17060                           If arr_varNone(N_FND, lngZ) = False And lngZ <> lngY Then
17070                             If arr_varNone(N_ACTNO, lngZ) = arr_varGrp(G_ACTNO, lngX) Then
                                    ' ** Make sure accountno's match.
17080                               blnFound = False
17090                               If ((arr_varGrp(G_ANO1, lngX) = 0& And arr_varGrp(G_ANO2, lngX) = 0& And arr_varNone(N_ANO, lngY) = 0&) Or _
                                        (arr_varGrp(G_ANO1, lngX) = 0& And arr_varGrp(G_ANO2, lngX) = 0& And arr_varNone(N_ANO, lngZ) = 0&) Or _
                                        (arr_varGrp(G_ANO1, lngX) = 0& And arr_varNone(N_ANO, lngY) = 0& And arr_varNone(N_ANO, lngZ) = 0&) Or _
                                        (arr_varGrp(G_ANO2, lngX) = 0& And arr_varNone(N_ANO, lngY) = 0& And arr_varNone(N_ANO, lngZ) = 0&)) Then
                                      ' ** Lines 1-4, 3 Zeroes, 1 Asset (ODD):
                                      ' **   1: lngZ ODD, 2: lngY ODD, 3: ANO2 ODD, 4: ANO1 ODD.
17100                                 blnFound = True
17110                               ElseIf ((arr_varGrp(G_ANO1, lngX) = 0& And arr_varGrp(G_ANO2, lngX) = 0& And _
                                        arr_varNone(N_ANO, lngY) = arr_varNone(N_ANO, lngZ)) Or _
                                        (arr_varGrp(G_ANO1, lngX) = 0& And arr_varNone(N_ANO, lngY) = 0& And _
                                        arr_varGrp(G_ANO2, lngX) = arr_varNone(N_ANO, lngZ)) Or _
                                        (arr_varGrp(G_ANO1, lngX) = 0& And arr_varNone(N_ANO, lngZ) = 0& And _
                                        arr_varGrp(G_ANO2, lngX) = arr_varNone(N_ANO, lngY)) Or _
                                        (arr_varNone(N_ANO, lngY) = 0& And arr_varNone(N_ANO, lngZ) = 0& And _
                                        arr_varGrp(G_ANO1, lngX) = arr_varGrp(G_ANO2, lngX)) Or _
                                        (arr_varGrp(G_ANO2, lngX) = 0& And arr_varNone(N_ANO, lngZ) = 0& And _
                                        arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngY)) Or _
                                        (arr_varGrp(G_ANO2, lngX) = 0& And arr_varNone(N_ANO, lngY) = 0& And _
                                        arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngZ))) Then
                                      ' ** Lines 1-12, 2 Zeroes, 2 Equal Assets (ODD):
                                      ' **   1/2: lngY/lngZ ODD, 3/4: ANO2/lngY ODD, 5/6: ANO2/lngZ ODD,
                                      ' **   7/8: ANO1/ANO2 ODD, 9/10: ANO1/lngY ODD, 11/12: ANO1/lngZ ODD.
17120                                 blnFound = True
17130                               ElseIf ((arr_varGrp(G_ANO1, lngX) = 0& And arr_varGrp(G_ANO2, lngX) = arr_varNone(N_ANO, lngY) And _
                                        arr_varGrp(G_ANO2, lngX) = arr_varNone(N_ANO, lngZ)) Or _
                                        (arr_varGrp(G_ANO2, lngX) = 0& And arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngY) And _
                                        arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngZ)) Or _
                                        (arr_varNone(N_ANO, lngY) = 0& And arr_varGrp(G_ANO1, lngX) = arr_varGrp(G_ANO2, lngX) And _
                                        arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngZ)) Or _
                                        (arr_varNone(N_ANO, lngZ) = 0& And arr_varGrp(G_ANO1, lngX) = arr_varGrp(G_ANO2, lngX) And _
                                        arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngY))) Then
                                      ' ** Lines 1-8, 1 Zero, 3 Equal Assets (ODD):
                                      ' **   1/2: ANO1 0, ANO2/lngY/lngZ ODD
                                      ' **   3/4: ANO2 0, ANO1/lngY/lngZ ODD
                                      ' **   5/6: lngY 0, ANO1/ANO2/lngZ ODD
                                      ' **   7/8: lngZ 0, ANO1/ANO2/lngY ODD
17140                                 blnFound = True
17150                               ElseIf ((arr_varGrp(G_ANO1, lngX) = arr_varGrp(G_ANO2, lngX) And _
                                        arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngY) And _
                                        arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngZ))) Then
                                      ' ** Lines 1-3, All Equal:
                                      ' **   1/2/3: ALL EQUAL.
17160                                 blnFound = True
17170                               End If  ' ** assetno.
17180                               If blnFound = True Then
                                      ' ** Make sure assetno's are appropriate.
17190                                 If (((arr_varGrp(G_ICSH, lngX) + arr_varNone(N_ICSH, lngY) + arr_varNone(N_ICSH, lngZ)) = 0@) And _
                                          ((arr_varGrp(G_PCSH, lngX) + arr_varNone(N_PCSH, lngY) + arr_varNone(N_PCSH, lngZ)) = 0@) And _
                                          ((arr_varGrp(G_COST, lngX) + arr_varNone(N_COST, lngY) + arr_varNone(N_COST, lngZ)) = 0@)) Then
                                        ' ** Well, maybe if I'm lucky.
17200                                   arr_varGrp(G_JNO1, lngX) = arr_varNone(N_JNO, lngY)
17210                                   arr_varGrp(G_JNO2, lngX) = arr_varNone(N_JNO, lngZ)
17220                                   arr_varGrp(G_FND, lngX) = CBool(True)
17230                                   arr_varNone(N_GRP, lngY) = arr_varGrp(G_GRP, lngX)
17240                                   arr_varNone(N_GRP, lngZ) = arr_varGrp(G_GRP, lngX)
17250                                   arr_varNone(N_FND, lngY) = CBool(True)
17260                                   arr_varNone(N_FND, lngZ) = CBool(True)
17270                                   lngTmp01 = lngTmp01 + 1&
17280                                   Exit For
17290                                 End If  ' ** icash, pcash, cost.
17300                               End If  ' ** blnFound.
17310                             End If  ' ** accountno.
17320                           End If  ' ** N_FND.
17330                         Next  ' ** lngZ.
17340                       End If  ' ** accountno.
17350                     End If
17360                     If arr_varGrp(G_FND, lngX) = True Then
17370                       Exit For
17380                     End If
17390                   Next  ' ** lngY.
17400                 End If
                      ' ***************************************************************
17410                 dblPB_ThisStepSub = lngX
17420                 dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
17430                 ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                      '.ProgBar_bar.Width = dblPB_ThisWidthSub
17440                 strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
17450                 .ProgBar_lbl1.Caption = strPB_ThisPct
17460                 DoEvents
                      ' ***************************************************************
17470               Next  ' ** lngX.
17480             End If
17490             DoEvents

                  ' ** Check again.
17500             lngTmp01 = 0&
17510             For lngX = 0& To (lngGrps - 1&)
17520               If arr_varGrp(G_FND, lngX) = False Then
17530                 lngTmp01 = lngTmp01 + 1&
17540               End If
17550             Next
17560             If lngTmp01 = 0& Then
                    ' ** Hooray!
17570             Else
                    ' ** I'm not sure what to do about these.
17580               .unmatched_groups = lngTmp01
17590             End If

                  ' ** LedgHidType enumeration:
                  ' **   0  GRP_NONE    Unmatched hidden transactions. This should only be a temporary designation.
                  ' **   1  NORM        2 entries in hidden group, with matching assetno (which could both be zero).
                  ' **   2  NORM_MISC   2 entries in hidden group, where both are 'Misc.', to be treated like a normal pair.
                  ' **   3  MISC_2_GRP  2 entries in hidden group, one 'Misc.' and one other.
                  ' **   4  MISC_3_GRP  3 entries in hidden group, one 'Misc.' and two other matching assetno.
                  ' **   5  MULTI_GRP   3 or more entries in hidden group, with matching assetno, multi-lot group

                  ' ***************************************************************
                  ' ** Step 8: Processing old groups - pass 1.
17600             dblPB_ThisStep = 8#
17610             .Status2_lbl.Caption = "Processing old groups - pass 1"
17620             DoEvents
                  ' ***************************************************************
                  ' ***************************************************************
17630             dblPB_ThisWidth = 0#
17640             For dblZ = 1# To (dblPB_ThisStep - 1#)
                    ' ** Assemble the weighted widths up to, but not including, this width.
17650               dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
17660             Next
                  ' ***************************************************************

                  ' ** This transfers the old GRP_NONE's to tblLedgerHidden in order to complete the old groups.
17670             If lngGrps - lngTmp01 > 0& Then
                    ' ** At least some were completed.
                    ' ***************************************************************
17680               dblPB_StepSubs = lngGrps
17690               dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
17700               dblPB_ThisStepSub = 0#
17710               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    '.ProgBar_bar.Width = dblPB_ThisWidth
17720               DoEvents
                    ' ***************************************************************
17730               For lngX = 0& To (lngGrps - 1&)
17740                 If arr_varGrp(G_FND, lngX) = True Then
17750                   If IsNull(arr_varGrp(G_JNO1, lngX)) = False And IsNull(arr_varGrp(G_JNO2, lngX)) = True Then
                          ' ** 1 GRP_NONE completes the group.
17760                     arr_varGrp(G_CNT, lngX) = arr_varGrp(G_CNT, lngX) + 1&
17770                     lngHidType = 0&: lngE = -1&
                          ' ** Get the element number of the new addition.
17780                     For lngY = 0& To (lngNones - 1&)
17790                       If arr_varNone(N_JNO, lngY) = arr_varGrp(G_JNO1, lngX) Then
17800                         lngE = lngY
17810                         Exit For
17820                       End If
17830                     Next  ' ** lngY.
                          ' ** Figure out the ledghidtype_type.
17840                     If arr_varGrp(G_CNT, lngX) > 3& Then
17850                       lngHidType = 5&
17860                     ElseIf arr_varGrp(G_CNT, lngX) = 3& Then
17870                       If (arr_varGrp(G_ANO1, lngX) = arr_varGrp(G_ANO2, lngX)) And _
                                (arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngE)) Then
                              ' ** All 3 assets are the same.
17880                         lngHidType = 5&
17890                       ElseIf ((arr_varGrp(G_ANO1, lngX) = 0&) And (arr_varGrp(G_ANO2, lngX) = arr_varNone(N_ANO, lngE))) Or _
                                ((arr_varGrp(G_ANO2, lngX) = 0&) And (arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngE))) Or _
                                ((arr_varNone(N_ANO, lngE) = 0&) And (arr_varGrp(G_ANO1, lngX) = arr_varGrp(G_ANO2, lngX))) Then
                              ' ** 1 Misc and 2 matching assets.
17900                         lngHidType = 4&
17910                       End If
17920                     ElseIf arr_varGrp(G_CNT, lngX) = 2& Then
17930                       If (arr_varGrp(G_ANO1, lngX) = 0& And arr_varNone(N_ANO, lngE) = 0&) Then
                              ' ** Both are Misc.
17940                         lngHidType = 2&
17950                       ElseIf (arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngE)) Then
                              ' ** Both are assets.
17960                         lngHidType = 1&
17970                       ElseIf (arr_varGrp(G_ANO1, lngX) = 0& Or arr_varNone(N_ANO, lngE) = 0&) Then
                              ' ** 1 Misc and 1 asset.
17980                         lngHidType = 3&
17990                       End If
18000                     End If
18010                     DoEvents
                          ' ** Append qryAccountHideTrans2_55b (LedgerHidden, linked to qryAccountHideTrans2_55a
                          ' ** (tblLedgerHidden, grouped, by specified [grp]), with dummy uniqueid, for 1 addition,
                          ' ** by specified [jno], [typ]) to tblLedgerHidden.
18020                     Set qdf = dbs.QueryDefs("qryAccountHideTrans2_55c")
18030                     With qdf.Parameters
18040                       ![grp] = arr_varGrp(G_GRP, lngX)
18050                       ![jno] = arr_varNone(N_JNO, lngE)
18060                       ![typ] = lngHidType
18070                     End With
18080                     qdf.Execute dbFailOnError
18090                     Set qdf = Nothing
18100                     DoEvents
18110                   ElseIf IsNull(arr_varGrp(G_JNO1, lngX)) = False And IsNull(arr_varGrp(G_JNO2, lngX)) = False Then
                          ' ** 2 GRP_NONE's complete the group.
18120                     arr_varGrp(G_CNT, lngX) = arr_varGrp(G_CNT, lngX) + 2&
                          ' ** Get the element numbers of the new additions.
18130                     lngHidType = 0&: lngE = -1&: lngF = -1&
18140                     For lngY = 0& To (lngNones - 1&)
18150                       If arr_varNone(N_JNO, lngY) = arr_varGrp(G_JNO1, lngX) Then
18160                         lngE = lngY
18170                       ElseIf arr_varNone(N_JNO, lngY) = arr_varGrp(G_JNO2, lngX) Then
18180                         lngF = lngY
18190                       End If
18200                     Next  ' ** lngY.
                          ' ** Figure out the ledghidtype_type.
18210                     If arr_varGrp(G_CNT, lngX) > 3& Then
18220                       lngHidType = 5&
18230                     ElseIf arr_varGrp(G_CNT, lngX) = 3& Then
18240                       If (arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngE)) And _
                                (arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngF)) Then
                              ' ** All 3 assets are the same.
18250                         lngHidType = 5&
18260                       ElseIf ((arr_varGrp(G_ANO1, lngX) = 0&) And (arr_varNone(N_ANO, lngE) = arr_varNone(N_ANO, lngF))) Or _
                                ((arr_varNone(N_ANO, lngE) = 0&) And (arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngF))) Or _
                                ((arr_varNone(N_ANO, lngF) = 0&) And (arr_varGrp(G_ANO1, lngX) = arr_varNone(N_ANO, lngE))) Then
                              ' ** 1 Misc and 2 matching assets.
18270                         lngHidType = 4&
18280                       End If
18290                     End If
18300                     DoEvents
                          ' ** Append qryAccountHideTrans2_55d (LedgerHidden, linked to qryAccountHideTrans2_55a
                          ' ** (tblLedgerHidden, grouped, by specified [grp]), with dummy uniqueid, for 1 of 2
                          ' ** additions, by specified [jno], [typ], [cnt], [ord]) to tblLedgerHidden.
18310                     Set qdf = dbs.QueryDefs("qryAccountHideTrans2_55f")
18320                     qdf.Execute dbFailOnError
18330                     Set qdf = Nothing
18340                     DoEvents
                          ' ** Append qryAccountHideTrans2_55e (LedgerHidden, linked to qryAccountHideTrans2_55a
                          ' ** (tblLedgerHidden, grouped, by specified [grp]), with dummy uniqueid, for 2 of 2
                          ' ** additions, by specified [jno], [typ], [cnt], [ord]) to tblLedgerHidden.
18350                     Set qdf = dbs.QueryDefs("qryAccountHideTrans2_55g")
18360                     qdf.Execute dbFailOnError
18370                     Set qdf = Nothing
18380                     DoEvents
18390                   End If
18400                 End If  ' ** G_FND.
                      ' ***************************************************************
18410                 dblPB_ThisStepSub = lngX
18420                 dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
18430                 ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                      '.ProgBar_bar.Width = dblPB_ThisWidthSub
18440                 strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
18450                 .ProgBar_lbl1.Caption = strPB_ThisPct
18460                 DoEvents
                      ' ***************************************************************
18470               Next  ' ** lngX.
18480             End If  ' ** lngGrps, lngTmp01.

                  ' ** Examples:
                  ' **   accountno      astno jno1   jno2   jtype1    jtype2
                  ' **   =============== ==== ====== ====== ========= =========
                  ' **   000000000002300_0412_020353_020350_Withdrawn_Deposit__...
                  ' **   000000000002300_0000_040395_000000_Misc.______________
                  ' **          15         4     6      6       9         9      ' ** Might someone get more than 9,999 assets?
                  ' ** Whichever one may have an assetno gets the slot.

                  ' ***************************************************************
                  ' ** Step 9: Processing old groups - pass 2.
18490             dblPB_ThisStep = 9#
18500             .Status2_lbl.Caption = "Processing old groups - pass 2"
18510             DoEvents
                  ' ***************************************************************
                  ' ***************************************************************
18520             dblPB_ThisWidth = 0#
18530             For dblZ = 1# To (dblPB_ThisStep - 1#)
                    ' ** Assemble the weighted widths up to, but not including, this width.
18540               dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
18550             Next
                  ' ***************************************************************

                  ' ** Now put in the correct uniqueid.
                  ' ***************************************************************
18560             dblPB_StepSubs = lngGrps
18570             dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
18580             dblPB_ThisStepSub = 0#
18590             ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                  '.ProgBar_bar.Width = dblPB_ThisWidth
18600             DoEvents
                  ' ***************************************************************
18610             For lngX = 0& To (lngGrps - 1&)
18620               strUniqueID = vbNullString: lngTmp01 = 0&
18630               If arr_varGrp(G_FND, lngX) = True Then
                      ' ** tblLedgerHidden, by specified [grp].
18640                 Set qdf = dbs.QueryDefs("qryAccountHideTrans2_56")
18650                 With qdf.Parameters
18660                   ![grp] = arr_varGrp(G_GRP, lngX)
18670                 End With
18680                 Set rst = qdf.OpenRecordset
18690                 With rst
18700                   .MoveLast
18710                   lngRecs = .RecordCount
18720                   .MoveFirst
18730                   strUniqueID = Right(String(15, "0") & ![accountno], 15) & "_"
                        ' ** First, get the assetno.
18740                   For lngY = 1& To lngRecs
18750                     If ![assetno] <> 0& Then
18760                       lngTmp01 = ![assetno]
18770                       Exit For
18780                     End If
18790                     If lngY < lngRecs Then .MoveNext
18800                   Next
18810                   strUniqueID = strUniqueID & Right(String(4, "0") & CStr(lngTmp01), 4) & "_"
18820                   .MoveFirst
                        ' ** Now, put in the journalno's.
18830                   For lngY = 1& To lngRecs
18840                     strUniqueID = strUniqueID & Right(String(6, "0") & CStr(![journalno]), 6) & "_"
18850                     If lngY < lngRecs Then .MoveNext
18860                   Next
18870                   .MoveFirst
                        ' ** Then, the journaltype's.
18880                   For lngY = 1& To lngRecs
18890                     strUniqueID = strUniqueID & Left(![journaltype] & String(9, "_"), 9)
18900                     If lngY < lngRecs Then strUniqueID = strUniqueID & "_"
18910                     If lngY < lngRecs Then .MoveNext
18920                   Next
18930                   .MoveFirst
                        ' ** And finally, edit the records.
18940                   For lngY = 1& To lngRecs
18950                     .Edit
18960                     ![ledghid_uniqueid] = strUniqueID
18970                     .Update
18980                     If lngY < lngRecs Then .MoveNext
18990                   Next
19000                   .Close
19010                 End With
19020                 Set rst = Nothing
19030                 Set qdf = Nothing
19040                 DoEvents
19050               End If
                    ' ***************************************************************
19060               dblPB_ThisStepSub = lngX
19070               dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
19080               ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                    '.ProgBar_bar.Width = dblPB_ThisWidthSub
19090               strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
19100               .ProgBar_lbl1.Caption = strPB_ThisPct
19110               DoEvents
                    ' ***************************************************************
19120             Next  ' ** lngX.

19130           End If  ' ** lngGrps.

                ' ** Next, we'll try to match any remaining old GRP_NONE's among themselves.
19140           lngNones = 0&
19150           ReDim arr_varNone(N_ELEMS, 0)

                ' ***************************************************************
                ' ** Step 10: Begin processing remaining unmatched.
19160           dblPB_ThisStep = 10#
19170           .Status2_lbl.Caption = "Begin processing remaining unmatched"
19180           DoEvents
                ' ***************************************************************
                ' ***************************************************************
19190           dblPB_ThisWidth = 0#
19200           For dblZ = 1# To (dblPB_ThisStep - 1#)
                  ' ** Assemble the weighted widths up to, but not including, this width.
19210             dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
19220           Next
                ' ***************************************************************

                ' ** OK, now let's try to match up the rest of the GRP_NONE's.
                ' ** qryAccountHideTrans2_53 (LedgerHidden, just 'GRP_NONE'), not in tblLedgerHidden.
19230           Set qdf = dbs.QueryDefs("qryAccountHideTrans2_57a")
19240           Set rst = qdf.OpenRecordset
19250           With rst
19260             lngTmp01 = 0&
19270             If .BOF = True And .EOF = True Then
                    ' ** Great! Nothing more to match.
                    ' ***************************************************************
19280               dblPB_StepSubs = 0#  ' ** No subs in this step.
19290               dblPB_ThisIncrSub = 0#
19300               dblPB_ThisStepSub = 0#
19310               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    'frm.ProgBar_bar.Width = dblPB_ThisWidth
19320               strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
19330               frm.ProgBar_lbl1.Caption = strPB_ThisPct
19340               DoEvents
                    ' ***************************************************************
19350             Else
19360               .MoveLast
19370               lngRecs = .RecordCount
19380               .MoveFirst
                    ' ***************************************************************
19390               dblPB_StepSubs = lngRecs
19400               dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
19410               dblPB_ThisStepSub = 0#
19420               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    'frm.ProgBar_bar.Width = dblPB_ThisWidth
19430               DoEvents
                    ' ***************************************************************
19440               If lngRecs > 1& Then
19450                 For lngX = 1& To lngRecs
19460                   lngNones = lngNones + 1&
19470                   lngE = lngNones - 1&
19480                   ReDim Preserve arr_varNone(N_ELEMS, lngE)
19490                   arr_varNone(N_ACTNO, lngE) = ![accountno]
19500                   arr_varNone(N_JNO, lngE) = ![journalno]
19510                   arr_varNone(N_ANO, lngE) = ![assetno]
19520                   arr_varNone(N_ICSH, lngE) = CCur(Round(![ICash], 2))
19530                   arr_varNone(N_PCSH, lngE) = CCur(Round(![PCash], 2))
19540                   arr_varNone(N_COST, lngE) = CCur(Round(![Cost], 2))
19550                   arr_varNone(N_GRP, lngE) = Null
19560                   arr_varNone(N_FND, lngE) = CBool(False)
19570                   If lngX < lngRecs Then .MoveNext
                        ' ***************************************************************
19580                   dblPB_ThisStepSub = lngX
19590                   dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
19600                   ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                        'frm.ProgBar_bar.Width = dblPB_ThisWidthSub
19610                   strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
19620                   frm.ProgBar_lbl1.Caption = strPB_ThisPct
19630                   DoEvents
                        ' ***************************************************************
19640                 Next
19650               Else
                      ' ** One? Don't know what to do about this.
19660                 lngTmp01 = 1&
19670               End If
19680             End If
19690             .Close
19700           End With
19710           Set rst = Nothing
19720           Set qdf = Nothing
19730           .unmatched_singles = lngTmp01
19740           DoEvents

                ' ** And it's here that matching begins.
19750           If lngNones > 0& Then

19760             varTmp00 = DMax("[ledghid_grpnum]", "tblLedgerHidden")
19770             lngGrpNum = varTmp00

                  ' ***************************************************************
                  ' ** Step 11: Processing remaining unmatched - pass 1.
19780             dblPB_ThisStep = 11#
19790             .Status2_lbl.Caption = "Processing remaining unmatched - pass 1"
19800             DoEvents
                  ' ***************************************************************
                  ' ***************************************************************
19810             dblPB_ThisWidth = 0#
19820             For dblZ = 1# To (dblPB_ThisStep - 1#)
                    ' ** Assemble the weighted widths up to, but not including, this width.
19830               dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
19840             Next
                  ' ***************************************************************

                  ' ** Start with 1 other.
                  ' ***************************************************************
19850             dblPB_StepSubs = lngNones
19860             dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
19870             dblPB_ThisStepSub = 0#
19880             ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                  '.ProgBar_bar.Width = dblPB_ThisWidth
19890             DoEvents
                  ' ***************************************************************
19900             For lngX = 0& To (lngNones - 1&)
19910               If arr_varNone(N_FND, lngX) = False Then
19920                 For lngY = 0& To (lngNones - 1&)
19930                   If arr_varNone(N_FND, lngY) = False And lngY <> lngX Then
19940                     If arr_varNone(N_ACTNO, lngY) = arr_varNone(N_ACTNO, lngX) Then
                            ' ** Make sure accountno's match.
19950                       If (arr_varNone(N_ANO, lngX) = 0&) Or (arr_varNone(N_ANO, lngY) = 0&) Or _
                                (arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngY)) Then
19960                         If ((arr_varNone(N_ICSH, lngX) + arr_varNone(N_ICSH, lngY)) = 0&) And _
                                  ((arr_varNone(N_PCSH, lngX) + arr_varNone(N_PCSH, lngY)) = 0&) And _
                                  ((arr_varNone(N_COST, lngX) + arr_varNone(N_COST, lngY)) = 0&) Then
                                ' ** We've got a match!
19970                           lngGrpNum = lngGrpNum + 1&
19980                           arr_varNone(N_GRP, lngX) = lngGrpNum
19990                           arr_varNone(N_GRP, lngY) = lngGrpNum
20000                           arr_varNone(N_FND, lngX) = CBool(True)
20010                           arr_varNone(N_FND, lngY) = CBool(True)
20020                           Exit For
20030                         End If  ' ** icash, pcash, cost.
20040                       End If  ' ** assetno.
20050                     End If  ' ** accountno.
20060                   End If  ' ** N_FND.
20070                 Next  ' ** lngY.
20080               End If  ' ** N_FND.
                    ' ***************************************************************
20090               dblPB_ThisStepSub = lngX
20100               dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
20110               ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                    '.ProgBar_bar.Width = dblPB_ThisWidthSub
20120               strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
20130               .ProgBar_lbl1.Caption = strPB_ThisPct
20140               DoEvents
                    ' ***************************************************************
20150             Next  ' ** lngX.
20160             DoEvents

                  ' ** Recount.
20170             lngTmp01 = 0&
20180             For lngX = 0& To (lngNones - 1&)
20190               If arr_varNone(N_FND, lngX) = False Then
20200                 lngTmp01 = lngTmp01 + 1&
20210               End If
20220             Next  ' ** lngX.

                  ' ***************************************************************
                  ' ** Step 12: Processing remaining unmatched - pass 2.
20230             dblPB_ThisStep = 12#
20240             .Status2_lbl.Caption = "Processing remaining unmatched - pass 2"
20250             DoEvents
                  ' ***************************************************************
                  ' ***************************************************************
20260             dblPB_ThisWidth = 0#
20270             For dblZ = 1# To (dblPB_ThisStep - 1#)
                    ' ** Assemble the weighted widths up to, but not including, this width.
20280               dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
20290             Next
                  ' ***************************************************************

                  ' ** Here that matching continues.
20300             If lngTmp01 > 1& Then
                    ' ** Still some unmatched, so let's try 2 others.
                    ' ***************************************************************
20310               dblPB_StepSubs = lngNones
20320               dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
20330               dblPB_ThisStepSub = 0#
20340               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    '.ProgBar_bar.Width = dblPB_ThisWidth
20350               DoEvents
                    ' ***************************************************************
20360               For lngX = 0& To (lngNones - 1&)
20370                 If arr_varNone(N_FND, lngX) = False Then
20380                   For lngY = 0& To (lngNones - 1&)
20390                     If arr_varNone(N_FND, lngY) = False And lngY <> lngX Then
20400                       If arr_varNone(N_ACTNO, lngY) = arr_varNone(N_ACTNO, lngX) Then
                              ' ** Make sure accountno's match.
20410                         For lngZ = 0& To (lngNones - 1&)
20420                           If arr_varNone(N_FND, lngZ) = False And lngZ <> lngX And lngZ <> lngY Then
20430                             If arr_varNone(N_ACTNO, lngZ) = arr_varNone(N_ACTNO, lngX) Then
                                    ' ** Make sure accountno's match.
20440                               If (arr_varNone(N_ANO, lngX) = 0& And arr_varNone(N_ANO, lngY) = 0&) Or _
                                        (arr_varNone(N_ANO, lngY) = 0& And arr_varNone(N_ANO, lngZ) = 0&) Or _
                                        (arr_varNone(N_ANO, lngX) = 0& And arr_varNone(N_ANO, lngZ) = 0&) Or _
                                        (arr_varNone(N_ANO, lngX) = 0& And arr_varNone(N_ANO, lngY) = arr_varNone(N_ANO, lngZ)) Or _
                                        (arr_varNone(N_ANO, lngY) = 0& And arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngZ)) Or _
                                        (arr_varNone(N_ANO, lngZ) = 0& And arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngY)) Or _
                                        (arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngY) And _
                                        arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngZ)) Then
                                      ' ** Make sure assetno's are appropriate.
20450                                 If ((arr_varNone(N_ICSH, lngX) + arr_varNone(N_ICSH, lngY) + arr_varNone(N_ICSH, lngZ)) = 0&) And _
                                          ((arr_varNone(N_PCSH, lngX) + arr_varNone(N_PCSH, lngY) + arr_varNone(N_PCSH, lngZ)) = 0&) And _
                                          ((arr_varNone(N_COST, lngX) + arr_varNone(N_COST, lngY) + arr_varNone(N_COST, lngZ)) = 0&) Then
                                        ' ** Amazingly, we've got a match!
20460                                   lngGrpNum = lngGrpNum + 1&
20470                                   arr_varNone(N_GRP, lngX) = lngGrpNum
20480                                   arr_varNone(N_GRP, lngY) = lngGrpNum
20490                                   arr_varNone(N_GRP, lngZ) = lngGrpNum
20500                                   arr_varNone(N_FND, lngX) = CBool(True)
20510                                   arr_varNone(N_FND, lngY) = CBool(True)
20520                                   arr_varNone(N_FND, lngZ) = CBool(True)
20530                                   Exit For
20540                                 End If  ' ** icash, pcash, cost.
20550                               End If  ' ** assetno.
20560                             End If  ' ** accountno.
20570                           End If  ' ** N_FND.
20580                         Next  ' ** lngX
20590                         If arr_varNone(N_FND, lngY) = True Then
20600                           Exit For
20610                         End If
20620                       End If  ' ** accountno.
20630                     End If  ' ** N_FND.
20640                   Next  ' ** lngY.
20650                 End If  ' ** N_FND.
                      ' ***************************************************************
20660                 dblPB_ThisStepSub = lngX
20670                 dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
20680                 ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                      '.ProgBar_bar.Width = dblPB_ThisWidthSub
20690                 strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
20700                 .ProgBar_lbl1.Caption = strPB_ThisPct
20710                 DoEvents
                      ' ***************************************************************
20720               Next  ' ** lngX.
20730             Else
20740               .unmatched_singles = lngTmp01
20750             End If  ' ** lngTmp01.
20760             DoEvents

                  ' ** Recount.
20770             lngTmp01 = 0&
20780             For lngX = 0& To (lngNones - 1&)
20790               If arr_varNone(N_FND, lngX) = False Then
20800                 lngTmp01 = lngTmp01 + 1&
20810               End If
20820             Next  ' ** lngX.

                  ' ***************************************************************
                  ' ** Step 13: Processing remaining unmatched - pass 3.
20830             dblPB_ThisStep = 13#
20840             .Status2_lbl.Caption = "Processing remaining unmatched - pass 3"
20850             DoEvents
                  ' ***************************************************************
                  ' ***************************************************************
20860             dblPB_ThisWidth = 0#
20870             For dblZ = 1# To (dblPB_ThisStep - 1#)
                    ' ** Assemble the weighted widths up to, but not including, this width.
20880               dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
20890             Next
                  ' ***************************************************************

                  ' ** Now matches made among old GRP_NONE's are transferred to tblLedgerHidden.
20900             If lngTmp01 < lngNones Then
                    ' ** Whew! We found some matches.
20910               lngE = -1&: lngF = -1&
                    ' ***************************************************************
20920               dblPB_StepSubs = lngNones
20930               dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
20940               dblPB_ThisStepSub = 0#
20950               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    '.ProgBar_bar.Width = dblPB_ThisWidth
20960               DoEvents
                    ' ***************************************************************
20970               For lngX = 0& To (lngNones - 1&)
20980                 If arr_varNone(N_FND, lngX) = True Then

20990                   For lngY = 0& To (lngNones - 1&)
21000                     If arr_varNone(N_FND, lngY) = True And lngY <> lngX And _
                              arr_varNone(N_GRP, lngY) = arr_varNone(N_GRP, lngX) Then
21010                       lngE = lngY
                            ' ** Is this a 2 or a 3?
21020                       For lngZ = 0& To (lngNones - 1&)
21030                         If arr_varNone(N_FND, lngZ) = True And lngZ <> lngX And lngZ <> lngY And _
                                  arr_varNone(N_GRP, lngZ) = arr_varNone(N_GRP, lngX) Then
21040                           lngF = lngZ
21050                           Exit For
21060                         End If
21070                       Next  ' ** lngZ
21080                       Exit For
21090                     End If
21100                   Next  ' ** lngY.
21110                   DoEvents

                        ' ** LedgHidType enumeration:
                        ' **   0  GRP_NONE    Unmatched hidden transactions. This should only be a temporary designation.
                        ' **   1  NORM        2 entries in hidden group, with matching assetno (which could both be zero).
                        ' **   2  NORM_MISC   2 entries in hidden group, where both are 'Misc.', to be treated like a normal pair.
                        ' **   3  MISC_2_GRP  2 entries in hidden group, one 'Misc.' and one other.
                        ' **   4  MISC_3_GRP  3 entries in hidden group, one 'Misc.' and two other matching assetno.
                        ' **   5  MULTI_GRP   3 or more entries in hidden group, with matching assetno, multi-lot group

21120                   If lngE >= 0& Then
21130                     If lngF = -1& Then
21140                       lngTmp01 = 2&
21150                       lngHidType = 0&
21160                       If arr_varNone(N_ANO, lngX) = 0& And arr_varNone(N_ANO, lngE) = 0& Then
21170                         lngHidType = 2&
21180                       ElseIf arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngE) Then
21190                         lngHidType = 1&
21200                       ElseIf (arr_varNone(N_ANO, lngX) = 0& And arr_varNone(N_ANO, lngE) > 0&) Or _
                                (arr_varNone(N_ANO, lngX) > 0& And arr_varNone(N_ANO, lngE) = 0&) Then
21210                         lngHidType = 3&
21220                       End If
21230                     Else
21240                       lngTmp01 = 3&
21250                       lngHidType = 0&
21260                       If (arr_varNone(N_ANO, lngX) = 0& And arr_varNone(N_ANO, lngE) > 0& And _
                                arr_varNone(N_ANO, lngE) = arr_varNone(N_ANO, lngF)) Or _
                                (arr_varNone(N_ANO, lngE) = 0& And arr_varNone(N_ANO, lngX) > 0& And _
                                arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngF)) Or _
                                (arr_varNone(N_ANO, lngF) = 0& And arr_varNone(N_ANO, lngX) > 0& And _
                                arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngE)) Then
21270                         lngHidType = 4&
21280                       ElseIf arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngE) And _
                                arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngF) Then
21290                         lngHidType = 5&
21300                       End If
21310                     End If
21320                     DoEvents

                          ' ** Append qryAccountHideTrans2_57b (LedgerHidden, with dummy uniqueid, for 1st,
                          ' ** by specified [jno], [cnt], [grp], [ord], [typ]) to tblLedgerHidden.
21330                     Set qdf = dbs.QueryDefs("qryAccountHideTrans2_57e")
21340                     With qdf.Parameters
21350                       ![jno] = arr_varNone(N_JNO, lngX)
21360                       ![cnt] = lngTmp01
21370                       ![grp] = arr_varNone(N_GRP, lngX)
21380                       ![ord] = 1&
21390                       ![typ] = lngHidType
21400                     End With
21410                     qdf.Execute dbFailOnError
21420                     Set qdf = Nothing
21430                     DoEvents
                          ' ** Append qryAccountHideTrans2_57c (LedgerHidden, with dummy uniqueid, for 2nd,
                          ' ** by specified [jno], [cnt], [grp], [ord], [typ]) to tblLedgerHidden.
21440                     Set qdf = dbs.QueryDefs("qryAccountHideTrans2_57f")
21450                     With qdf.Parameters
21460                       ![jno] = arr_varNone(N_JNO, lngE)
21470                       ![cnt] = lngTmp01
21480                       ![grp] = arr_varNone(N_GRP, lngE)
21490                       ![ord] = 2&
21500                       ![typ] = lngHidType
21510                     End With
21520                     qdf.Execute dbFailOnError
21530                     Set qdf = Nothing
21540                     DoEvents
21550                     If lngTmp01 = 3& Then
                            ' ** Append qryAccountHideTrans2_57d (LedgerHidden, with dummy uniqueid, for 3rd,
                            ' ** by specified [jno], [cnt], [grp], [ord], [typ]) to tblLedgerHidden.
21560                       Set qdf = dbs.QueryDefs("qryAccountHideTrans2_57g")
21570                       With qdf.Parameters
21580                         ![jno] = arr_varNone(N_JNO, lngE)
21590                         ![cnt] = lngTmp01
21600                         ![grp] = arr_varNone(N_GRP, lngE)
21610                         ![ord] = 3&
21620                         ![typ] = lngHidType
21630                       End With
21640                       qdf.Execute dbFailOnError
21650                       Set qdf = Nothing
21660                     End If
21670                     DoEvents

21680                   End If

21690                 End If  ' ** N_FND.
                      ' ***************************************************************
21700                 dblPB_ThisStepSub = lngX
21710                 dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
21720                 ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                      '.ProgBar_bar.Width = dblPB_ThisWidthSub
21730                 strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
21740                 .ProgBar_lbl1.Caption = strPB_ThisPct
21750                 DoEvents
                      ' ***************************************************************
21760               Next  ' ** lngX.

21770               lngUniques = 0&
21780               ReDim arr_varUnique(U_ELEMS, 0)

                    ' ***************************************************************
                    ' ** Step 14: Processing remaining unmatched - pass 4.
21790               dblPB_ThisStep = 14#
21800               .Status2_lbl.Caption = "Processing remaining unmatched - pass 4"
21810               DoEvents
                    ' ***************************************************************
                    ' ***************************************************************
21820               dblPB_ThisWidth = 0#
21830               For dblZ = 1# To (dblPB_ThisStep - 1#)
                      ' ** Assemble the weighted widths up to, but not including, this width.
21840                 dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
21850               Next
                    ' ***************************************************************

                    ' ** tblLedgerHidden, grouped by ledghid_grpnum, with Max([assetno]).
21860               Set qdf = dbs.QueryDefs("qryAccountHideTrans2_58")  ' ** These are only the new groups.
21870               Set rst = qdf.OpenRecordset
21880               With rst
21890                 .MoveLast
21900                 lngRecs = .RecordCount
21910                 .MoveFirst
                      ' ***************************************************************
21920                 dblPB_StepSubs = lngRecs
21930                 dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
21940                 dblPB_ThisStepSub = 0#
21950                 ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                      'frm.ProgBar_bar.Width = dblPB_ThisWidth
21960                 DoEvents
                      ' ***************************************************************
21970                 For lngX = 1& To lngRecs
21980                   lngUniques = lngUniques + 1&
21990                   lngE = lngUniques - 1&
22000                   ReDim Preserve arr_varUnique(U_ELEMS, lngE)
22010                   arr_varUnique(U_ACTNO, lngE) = ![accountno]
22020                   arr_varUnique(U_GRP, lngE) = ![ledghid_grpnum]
22030                   arr_varUnique(U_CNT, lngE) = ![cnt]
22040                   If lngX < lngRecs Then .MoveNext
                        ' ***************************************************************
22050                   dblPB_ThisStepSub = lngX
22060                   dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
22070                   ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                        'frm.ProgBar_bar.Width = dblPB_ThisWidthSub
22080                   strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
22090                   frm.ProgBar_lbl1.Caption = strPB_ThisPct
22100                   DoEvents
                        ' ***************************************************************
22110                 Next
22120                 .Close
22130               End With
22140               Set rst = Nothing
22150               Set qdf = Nothing
22160               DoEvents

                    ' ***************************************************************
                    ' ** Step 15: Processing remaining unmatched - pass 5.
22170               dblPB_ThisStep = 15#
22180               .Status2_lbl.Caption = "Processing remaining unmatched - pass 5"
22190               DoEvents
                    ' ***************************************************************
                    ' ***************************************************************
22200               dblPB_ThisWidth = 0#
22210               For dblZ = 1# To (dblPB_ThisStep - 1#)
                      ' ** Assemble the weighted widths up to, but not including, this width.
22220                 dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
22230               Next
                    ' ***************************************************************

                    ' ** Now put in the correct uniqueid.
                    ' ***************************************************************
22240               dblPB_StepSubs = lngUniques
22250               dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
22260               dblPB_ThisStepSub = 0#
22270               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    '.ProgBar_bar.Width = dblPB_ThisWidth
22280               DoEvents
                    ' ***************************************************************
22290               For lngX = 0& To (lngUniques - 1&)
22300                 strUniqueID = vbNullString: lngTmp01 = 0&
                      ' ** tblLedgerHidden, linked to qryAccountHideTrans2_25 (qryAccountHideTrans2_24 (Union
                      ' ** of qryAccountHideTrans2_24a (Ledger, just needed fields), qryAccountHideTrans2_24b
                      ' ** (LedgerArchive, just needed fields)), just ledger_HIDDEN = True),
                      ' ** qryAccountHideTrans2_58 (tblLedgerHidden, grouped by ledghid_grpnum,
                      ' ** with Max([assetno])), by specified [grp].
22310                 Set qdf = dbs.QueryDefs("qryAccountHideTrans2_59")
22320                 With qdf.Parameters
22330                   ![grp] = arr_varUnique(U_GRP, lngX)
22340                 End With
22350                 Set rst = qdf.OpenRecordset
22360                 With rst
22370                   .MoveLast
22380                   lngRecs = .RecordCount
22390                   .MoveFirst
22400                   strUniqueID = ![ledghid_uniqueidx]  ' ** This is the base, with accountno and Max(assetno).
                        ' ** Put in the journalno's.
22410                   For lngZ = 1& To lngRecs
22420                     strUniqueID = strUniqueID & Right(String(6, "0") & CStr(![journalno]), 6) & "_"
22430                     If lngZ < lngRecs Then .MoveNext
22440                   Next
22450                   .MoveFirst
                        ' ** Then, the journaltype's.
22460                   For lngZ = 1& To lngRecs
22470                     strUniqueID = strUniqueID & Left(![journaltype] & String(9, "_"), 9)
22480                     If lngZ < lngRecs Then strUniqueID = strUniqueID & "_"
22490                     If lngZ < lngRecs Then .MoveNext
22500                   Next
22510                   .Close
22520                 End With
22530                 Set rst = Nothing
22540                 Set qdf = Nothing
22550                 DoEvents
                      ' ** tblLedgerHidden, all fields, by specified [grp].
22560                 Set qdf = dbs.QueryDefs("qryAccountHideTrans2_60")
22570                 With qdf.Parameters
22580                   ![grp] = arr_varUnique(U_GRP, lngX)
22590                 End With
22600                 Set rst = qdf.OpenRecordset
22610                 With rst
22620                   .MoveLast
22630                   lngRecs = .RecordCount
22640                   .MoveFirst
                        ' ** And finally, edit the records.
22650                   For lngZ = 1& To lngRecs
22660                     .Edit
22670                     ![ledghid_uniqueid] = strUniqueID
22680                     .Update
22690                     If lngZ < lngRecs Then .MoveNext
22700                   Next
22710                   .Close
22720                 End With
22730                 Set rst = Nothing
22740                 Set qdf = Nothing
22750                 DoEvents
                      ' ***************************************************************
22760                 dblPB_ThisStepSub = lngX
22770                 dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
22780                 ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                      '.ProgBar_bar.Width = dblPB_ThisWidthSub
22790                 strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
22800                 .ProgBar_lbl1.Caption = strPB_ThisPct
22810                 DoEvents
                      ' ***************************************************************
22820               Next  ' ** lngX.

22830             End If  ' ** lngTmp01, lngNones.

22840             .unmatched_singles = lngTmp01

22850           End If  ' ** lngNones.

22860         End If  ' ** lngGroupNones.

              ' ** All of the above involved transferring the old LedgerHidden to tblLedgerHidden,
              ' ** so we can now empty LedgerHidden and never touch it again.

              ' ** Empty LedgerHidden. (We're through with it!)
22870         Set qdf = dbs.QueryDefs("qryAccountHideTrans2_62")
22880         qdf.Execute
22890         Set qdf = Nothing
22900         DoEvents

              ' ***************************************************************
              ' ** Step 16: Begin processing new hidden.
22910         dblPB_ThisStep = 16#
22920         .Status2_lbl.Caption = "Begin processing new hidden"
22930         DoEvents
              ' ***************************************************************
              ' ***************************************************************
22940         dblPB_ThisWidth = 0#
22950         For dblZ = 1# To (dblPB_ThisStep - 1#)
                ' ** Assemble the weighted widths up to, but not including, this width.
22960           dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
22970         Next
              ' ***************************************************************

              ' ** Add new hidden transactions, not found in the old LedgerHidden, from scratch.
22980         If intMode = 2 Or intMode = 4 Or intMode = 5 Then

22990           lngNones = 0&
23000           ReDim arr_varNone(N_ELEMS, 0)

                ' ** qryAccountHideTrans2_25 (qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a
                ' ** (Ledger, just needed fields), qryAccountHideTrans2_24b (LedgerArchive, just needed
                ' ** fields)), just ledger_HIDDEN = True), linked to qryAccountHideTrans2_24 (Union of
                ' ** qryAccountHideTrans2_24a (Ledger, just needed fields), qryAccountHideTrans2_24b
                ' ** (LedgerArchive, just needed fields)), not in tblLedgerHidden.
23010           Set qdf = dbs.QueryDefs("qryAccountHideTrans2_61")
23020           Set rst = qdf.OpenRecordset
23030           With rst
23040             If .BOF = True And .EOF = True Then
                    ' ** No unmatched ledger_HIDDEN.
                    ' ***************************************************************
23050               dblPB_StepSubs = 0#  ' ** No subs in this step.
23060               dblPB_ThisIncrSub = 0#
23070               dblPB_ThisStepSub = 0#
23080               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    'frm.ProgBar_bar.Width = dblPB_ThisWidth
23090               strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
23100               frm.ProgBar_lbl1.Caption = strPB_ThisPct
23110               DoEvents
                    ' ***************************************************************
23120             Else
23130               .MoveLast
23140               lngRecs = .RecordCount
23150               .MoveFirst
                    ' ***************************************************************
23160               dblPB_StepSubs = lngRecs
23170               dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
23180               dblPB_ThisStepSub = 0#
23190               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    'frm.ProgBar_bar.Width = dblPB_ThisWidth
23200               DoEvents
                    ' ***************************************************************
23210               For lngX = 1& To lngRecs
23220                 lngNones = lngNones + 1&
23230                 lngE = lngNones - 1&
23240                 ReDim Preserve arr_varNone(N_ELEMS, lngE)
23250                 arr_varNone(N_ACTNO, lngE) = ![accountno]
23260                 arr_varNone(N_JNO, lngE) = ![journalno]
23270                 arr_varNone(N_ANO, lngE) = ![assetno]
23280                 arr_varNone(N_ICSH, lngE) = CCur(Round(![ICash], 2))
23290                 arr_varNone(N_PCSH, lngE) = CCur(Round(![PCash], 2))
23300                 arr_varNone(N_COST, lngE) = CCur(Round(![Cost], 2))
23310                 arr_varNone(N_GRP, lngE) = Null
23320                 arr_varNone(N_FND, lngE) = CBool(False)
                      ' ***************************************************************
23330                 dblPB_ThisStepSub = lngX
23340                 dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
23350                 ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                      'frm.ProgBar_bar.Width = dblPB_ThisWidthSub
23360                 strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
23370                 frm.ProgBar_lbl1.Caption = strPB_ThisPct
23380                 DoEvents
                      ' ***************************************************************
23390                 If lngX < lngRecs Then .MoveNext
23400               Next  ' ** lngX.
23410             End If
23420             .Close
23430           End With
23440           Set rst = Nothing
23450           Set qdf = Nothing
23460           DoEvents

                ' ** There are indeed hidden transactions not yet in tblLedgerHidden.
23470           If lngNones > 0& Then

23480             varTmp00 = DMax("[ledghid_grpnum]", "tblLedgerHidden")
23490             lngGrpNum = varTmp00

                  ' ***************************************************************
                  ' ** Step 17: Processing new hidden - pass 1.
23500             dblPB_ThisStep = 17#
23510             .Status2_lbl.Caption = "Processing new hidden - pass 1"
23520             DoEvents
                  ' ***************************************************************
                  ' ***************************************************************
23530             dblPB_ThisWidth = 0#
23540             For dblZ = 1# To (dblPB_ThisStep - 1#)
                    ' ** Assemble the weighted widths up to, but not including, this width.
23550               dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
23560             Next
                  ' ***************************************************************

                  ' ** Let's see what we can do.
                  ' ***************************************************************
23570             dblPB_StepSubs = lngNones
23580             dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
23590             dblPB_ThisStepSub = 0#
23600             ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                  '.ProgBar_bar.Width = dblPB_ThisWidth
23610             DoEvents
                  ' ***************************************************************
23620             For lngX = 0& To (lngNones - 1&)
23630               If arr_varNone(N_FND, lngX) = False Then
                      ' ** First try just 1 other (2 total).
23640                 For lngY = 0& To (lngNones - 1&)
23650                   If arr_varNone(N_FND, lngY) = False And lngY <> lngX Then
23660                     If arr_varNone(N_ACTNO, lngY) = arr_varNone(N_ACTNO, lngX) Then
23670                       If ((arr_varNone(N_ANO, lngX) = 0& Or arr_varNone(N_ANO, lngY) = 0&) Or _
                                (arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngY))) Then
23680                         If (((arr_varNone(N_ICSH, lngX) + arr_varNone(N_ICSH, lngY)) = 0&) And _
                                  ((arr_varNone(N_PCSH, lngX) + arr_varNone(N_PCSH, lngY)) = 0&) And _
                                  ((arr_varNone(N_COST, lngX) + arr_varNone(N_COST, lngY)) = 0&)) Then
                                ' ** We've got a match!
23690                           lngGrpNum = lngGrpNum + 1&
23700                           arr_varNone(N_GRP, lngX) = lngGrpNum
23710                           arr_varNone(N_GRP, lngY) = lngGrpNum
23720                           arr_varNone(N_FND, lngX) = CBool(True)
23730                           arr_varNone(N_FND, lngY) = CBool(True)
23740                           Exit For
23750                         End If  ' ** icash, pcash, cost.
23760                       End If  ' ** assetno.
23770                     End If  ' ** accountno
23780                   End If  ' ** N_FND.
23790                 Next  ' ** lngY.
23800               End If  ' ** N_FND.
                    ' ***************************************************************
23810               dblPB_ThisStepSub = lngX
23820               dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
23830               ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                    '.ProgBar_bar.Width = dblPB_ThisWidthSub
23840               strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
23850               .ProgBar_lbl1.Caption = strPB_ThisPct
23860               DoEvents
                    ' ***************************************************************
23870             Next  ' ** lngX.
23880             DoEvents

                  ' ** Recount.
23890             lngTmp01 = 0&
23900             For lngX = 0& To (lngNones - 1&)
23910               If arr_varNone(N_FND, lngX) = False Then
23920                 lngTmp01 = lngTmp01 + 1&
23930               End If
23940             Next  ' ** lngX.

23950             If lngTmp01 > 1& Then

                    ' ***************************************************************
                    ' ** Step 18: Processing new hidden - pass 2.
23960               dblPB_ThisStep = 18#
23970               .Status2_lbl.Caption = "Processing new hidden - pass 2"
23980               DoEvents
                    ' ***************************************************************
                    ' ***************************************************************
23990               dblPB_ThisWidth = 0#
24000               For dblZ = 1# To (dblPB_ThisStep - 1#)
                      ' ** Assemble the weighted widths up to, but not including, this width.
24010                 dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
24020               Next
                    ' ***************************************************************

                    ' ***************************************************************
24030               dblPB_StepSubs = lngNones
24040               dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
24050               dblPB_ThisStepSub = 0#
24060               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    '.ProgBar_bar.Width = dblPB_ThisWidth
24070               DoEvents
                    ' ***************************************************************
24080               For lngX = 0& To (lngNones - 1&)
24090                 If arr_varNone(N_FND, lngX) = False Then
                        ' ** Then try 2 others (3 total).
24100                   For lngY = 0& To (lngNones - 1&)
24110                     If arr_varNone(N_FND, lngY) = False And lngY <> lngX Then
24120                       If arr_varNone(N_ACTNO, lngY) = arr_varNone(N_ACTNO, lngX) Then
24130                         If ((arr_varNone(N_ANO, lngX) = 0& Or arr_varNone(N_ANO, lngY) = 0&) Or _
                                  (arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngY))) Then
24140                           For lngZ = 0& To (lngNones - 1&)
24150                             If arr_varNone(N_FND, lngZ) = False And lngZ <> lngX And lngZ <> lngY Then
24160                               If arr_varNone(N_ACTNO, lngZ) = arr_varNone(N_ACTNO, lngX) Then
24170                                 If (arr_varNone(N_ANO, lngX) = 0& And arr_varNone(N_ANO, lngY) = 0&) Or _
                                          (arr_varNone(N_ANO, lngX) = 0& And arr_varNone(N_ANO, lngZ) = 0&) Or _
                                          (arr_varNone(N_ANO, lngX) = 0& And arr_varNone(N_ANO, lngY) = arr_varNone(N_ANO, lngZ)) Or _
                                          (arr_varNone(N_ANO, lngY) = 0& And arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngZ)) Or _
                                          (arr_varNone(N_ANO, lngZ) = 0& And arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngY)) Or _
                                          (arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngY) And _
                                          arr_varNone(N_ANO, lngX) = arr_varNone(N_ANO, lngZ)) Then
24180                                   If (((arr_varNone(N_ICSH, lngX) + arr_varNone(N_ICSH, lngY) + arr_varNone(N_ICSH, lngZ)) = 0&) And _
                                            ((arr_varNone(N_PCSH, lngX) + arr_varNone(N_PCSH, lngY) + arr_varNone(N_PCSH, lngZ)) = 0&) And _
                                            ((arr_varNone(N_COST, lngX) + arr_varNone(N_COST, lngY) + arr_varNone(N_COST, lngZ)) = 0&)) Then
                                          ' ** Got another set!
24190                                     lngGrpNum = lngGrpNum + 1&
24200                                     arr_varNone(N_GRP, lngX) = lngGrpNum
24210                                     arr_varNone(N_GRP, lngY) = lngGrpNum
24220                                     arr_varNone(N_GRP, lngZ) = lngGrpNum
24230                                     arr_varNone(N_FND, lngX) = CBool(True)
24240                                     arr_varNone(N_FND, lngY) = CBool(True)
24250                                     arr_varNone(N_FND, lngZ) = CBool(True)
24260                                     Exit For
24270                                   End If  ' ** icash, pcash, cost.
24280                                 End If  ' ** assetno.
24290                               End If  ' ** accountno.
24300                             End If  ' ** N_FND.
24310                           Next  ' ** lngZ.
24320                         End If  ' ** assetno.
24330                       End If  ' ** accountno
24340                     End If  ' ** N_FND.
24350                   Next  ' ** lngY.
24360                 End If  ' ** N_FND.
                      ' ***************************************************************
24370                 dblPB_ThisStepSub = lngX
24380                 dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
24390                 ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                      '.ProgBar_bar.Width = dblPB_ThisWidthSub
24400                 strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
24410                 .ProgBar_lbl1.Caption = strPB_ThisPct
24420                 DoEvents
                      ' ***************************************************************
24430               Next  ' ** lngX.
24440               DoEvents

24450             Else
24460               .unmatched_singles = lngTmp01
24470             End If  ' ** lngTmp01.

                  ' ** Recount.
24480             lngTmp01 = 0&
24490             For lngX = 0& To (lngNones - 1&)
24500               If arr_varNone(N_FND, lngX) = False Then
24510                 lngTmp01 = lngTmp01 + 1&
24520               End If
24530             Next  ' ** lngX.

24540             .unmatched_singles = lngTmp01

24550             lngUniques = 0&
24560             ReDim arr_varUnique(U_ELEMS, 0)

                  ' ***************************************************************
                  ' ** Step 19: Processing new hidden - pass 3.
24570             dblPB_ThisStep = 19#
24580             .Status2_lbl.Caption = "Processing new hidden - pass 3"
24590             DoEvents
                  ' ***************************************************************
                  ' ***************************************************************
24600             dblPB_ThisWidth = 0#
24610             For dblZ = 1# To (dblPB_ThisStep - 1#)
                    ' ** Assemble the weighted widths up to, but not including, this width.
24620               dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
24630             Next
                  ' ***************************************************************

                  ' ** Add newly matched hidden transactions to tblLedgerHidden.
24640             If lngTmp01 < lngNones Then
                    ' ** Matches found.

                    ' ***************************************************************
24650               dblPB_StepSubs = lngNones
24660               dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
24670               dblPB_ThisStepSub = 0#
24680               ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                    '.ProgBar_bar.Width = dblPB_ThisWidth
24690               DoEvents
                    ' ***************************************************************
24700               For lngX = 0& To (lngNones - 1&)
24710                 If arr_varNone(N_FND, lngX) = True Then
24720                   blnFound = False
24730                   For lngY = 0& To (lngUniques - 1&)
24740                     If arr_varUnique(U_GRP, lngY) = arr_varNone(N_GRP, lngX) Then
24750                       blnFound = True
24760                       Exit For
24770                     End If
24780                   Next
24790                   Select Case blnFound
                        Case True
24800                     arr_varUnique(U_CNT, lngE) = arr_varUnique(U_CNT, lngE) + 1&
24810                   Case False
24820                     lngUniques = lngUniques + 1&
24830                     lngE = lngUniques - 1&
24840                     ReDim Preserve arr_varUnique(U_ELEMS, lngE)
24850                     arr_varUnique(U_ACTNO, lngE) = arr_varNone(N_ACTNO, lngX)
24860                     arr_varUnique(U_GRP, lngE) = arr_varNone(N_GRP, lngX)
24870                     arr_varUnique(U_CNT, lngE) = CLng(1)
24880                   End Select
24890                 End If  ' ** N_FND.
                      ' ***************************************************************
24900                 dblPB_ThisStepSub = lngX
24910                 dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
24920                 ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                      '.ProgBar_bar.Width = dblPB_ThisWidthSub
24930                 strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
24940                 .ProgBar_lbl1.Caption = strPB_ThisPct
24950                 DoEvents
                      ' ***************************************************************
24960               Next  ' ** lngX.
24970               DoEvents

                    ' ** Binary Sort arr_varUnique() array.
24980               For lngX = UBound(arr_varUnique, 2) To 1& Step -1&
24990                 For lngY = 0& To (lngX - 1&)
25000                   If arr_varUnique(U_GRP, lngY) > arr_varUnique(U_GRP, (lngY + 1&)) Then
25010                     For lngZ = 0& To U_ELEMS
25020                       varTmp00 = arr_varUnique(lngZ, lngY)
25030                       arr_varUnique(lngZ, lngY) = arr_varUnique(lngZ, (lngY + 1&))
25040                       arr_varUnique(lngZ, (lngY + 1&)) = varTmp00
25050                       varTmp00 = Empty
25060                     Next
25070                   End If
25080                 Next
25090               Next
25100               DoEvents

                    ' ** LedgHidType enumeration:
                    ' **   0  GRP_NONE    Unmatched hidden transactions. This should only be a temporary designation.
                    ' **   1  NORM        2 entries in hidden group, with matching assetno (which could both be zero).
                    ' **   2  NORM_MISC   2 entries in hidden group, where both are 'Misc.', to be treated like a normal pair.
                    ' **   3  MISC_2_GRP  2 entries in hidden group, one 'Misc.' and one other.
                    ' **   4  MISC_3_GRP  3 entries in hidden group, one 'Misc.' and two other matching assetno.
                    ' **   5  MULTI_GRP   3 or more entries in hidden group, with matching assetno, multi-lot group

                    ' ** Examples:
                    ' **   accountno      astno jno1   jno2   jtype1    jtype2
                    ' **   =============== ==== ====== ====== ========= =========
                    ' **   000000000002300_0412_020353_020350_Withdrawn_Deposit__...
                    ' **   000000000002300_0000_040395_000000_Misc.______________
                    ' **          15         4     6      6       9         9      ' ** Might someone get more than 9,999 assets?
                    ' ** Whichever one may have an assetno gets the slot.

                    ' ***************************************************************
                    ' ** Step 20: Processing new hidden - pass 4.
25110               dblPB_ThisStep = 20#
25120               .Status2_lbl.Caption = "Processing new hidden - pass 4"
25130               DoEvents
                    ' ***************************************************************
                    ' ***************************************************************
25140               dblPB_ThisWidth = 0#
25150               For dblZ = 1# To (dblPB_ThisStep - 1#)
                      ' ** Assemble the weighted widths up to, but not including, this width.
25160                 dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
25170               Next
                    ' ***************************************************************

                    ' ** Add them to tblLedgerHidden.
25180               Set rst = dbs.OpenRecordset("tblLedgerHidden", dbOpenDynaset, dbConsistent)
25190               With rst
                      ' ***************************************************************
25200                 dblPB_StepSubs = lngUniques
25210                 dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
25220                 dblPB_ThisStepSub = 0#
25230                 ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                      'frm.ProgBar_bar.Width = dblPB_ThisWidth
25240                 DoEvents
                      ' ***************************************************************
25250                 For lngX = 0& To (lngUniques - 1&)
25260                   strUniqueID = Right(String(15, "0") & arr_varUnique(U_ACTNO, lngX), 15) & "_"
                        ' ** Get the appropriate assetno for ledghid_uniqueid.
25270                   lngTmp01 = 0&
25280                   For lngY = 0& To (lngNones - 1&)
25290                     If arr_varNone(N_GRP, lngY) = arr_varUnique(U_GRP, lngX) Then
25300                       If arr_varNone(N_ANO, lngY) > lngTmp01 Then lngTmp01 = arr_varNone(N_ANO, lngY)
25310                     End If
25320                   Next
25330                   strUniqueID = strUniqueID & Right(String(4, "0") & CStr(lngTmp01), 4) & "_"
                        ' ** Add the journalno's to ledghid_uniqueid.
25340                   For lngY = 0& To (lngNones - 1&)
25350                     If arr_varNone(N_GRP, lngY) = arr_varUnique(U_GRP, lngX) Then
25360                       strUniqueID = strUniqueID & Right(String(6, "0") & CStr(arr_varNone(N_JNO, lngY)), 6) & "_"
25370                     End If
25380                   Next
                        ' ** Add the journaltype's to ledghid_uniqueid.
25390                   lngTmp01 = 0&
25400                   For lngY = 0& To (lngNones - 1&)
25410                     If arr_varNone(N_GRP, lngY) = arr_varUnique(U_GRP, lngX) Then
25420                       lngTmp01 = lngTmp01 + 1&
25430                       varTmp00 = DLookup("[journaltype]", "qryAccountHideTrans2_25", "[journalno] = " & CStr(arr_varNone(N_JNO, lngY)))
25440                       strUniqueID = strUniqueID & Left(varTmp00 & String(9, "_"), 9)
25450                       If lngTmp01 < arr_varUnique(U_CNT, lngX) Then strUniqueID = strUniqueID & "_"
25460                     End If
25470                   Next
                        ' ** Determine the ledghidtype_type.
25480                   lngHidType = 0&
25490                   If arr_varUnique(U_CNT, lngX) > 3& Then
25500                     lngHidType = 5&  ' ** Though this section doesn't search that high.
25510                   ElseIf arr_varUnique(U_CNT, lngX) = 3& Then
25520                     lngE = -1&: lngF = -1&
25530                     For lngY = 0& To (lngNones - 1&)
25540                       If arr_varNone(N_GRP, lngY) = arr_varUnique(U_GRP, lngX) Then
25550                         If lngE = -1& Then
25560                           lngE = lngY
25570                         ElseIf lngF = -1& Then
25580                           lngF = lngY
25590                         Else
25600                           If (arr_varNone(N_ANO, lngE) = arr_varNone(N_ANO, lngF) And _
                                    arr_varNone(N_ANO, lngE) = arr_varNone(N_ANO, lngY)) Then
25610                             lngHidType = 5&
25620                           Else
25630                             lngHidType = 4&
25640                           End If
25650                           Exit For
25660                         End If
25670                       End If
25680                     Next
25690                   ElseIf arr_varUnique(U_CNT, lngX) = 2& Then
25700                     lngE = -1&
25710                     For lngY = 0& To (lngNones - 1&)
25720                       If arr_varNone(N_GRP, lngY) = arr_varUnique(U_GRP, lngX) Then
25730                         If lngE = -1& Then
25740                           lngE = lngY
25750                         Else
25760                           If arr_varNone(N_ANO, lngE) = 0& And arr_varNone(N_ANO, lngY) = 0& Then
25770                             lngHidType = 2&
25780                           ElseIf arr_varNone(N_ANO, lngE) = 0& Or arr_varNone(N_ANO, lngY) = 0& Then
25790                             lngHidType = 3&
25800                           Else
25810                             lngHidType = 1&
25820                           End If
25830                           Exit For
25840                         End If
25850                       End If
25860                     Next
25870                   End If
25880                   lngTmp01 = 0&
25890                   For lngY = 0& To (lngNones - 1&)
25900                     If arr_varNone(N_GRP, lngY) = arr_varUnique(U_GRP, lngX) Then
25910                       .AddNew
                            ' ** ![ledghid_id] : AutoNumber.
25920                       ![journalno] = arr_varNone(N_JNO, lngY)
25930                       ![accountno] = arr_varNone(N_ACTNO, lngY)
25940                       ![assetno] = arr_varNone(N_ANO, lngY)
25950                       varTmp00 = DLookup("[transdate]", "qryAccountHideTrans2_25", "[journalno] = " & CStr(arr_varNone(N_JNO, lngY)))
25960                       ![transdate] = CDate(varTmp00)
25970                       ![ledghid_cnt] = arr_varUnique(U_CNT, lngX)
25980                       ![ledghid_grpnum] = arr_varUnique(U_GRP, lngX)
25990                       lngTmp01 = lngTmp01 + 1&
26000                       ![ledghid_ord] = lngTmp01
26010                       ![ledghidtype_type] = lngHidType
26020                       ![ledghid_uniqueid] = strUniqueID
26030                       ![ledghid_username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
26040                       ![ledghid_datemodified] = Now()
26050                       .Update
26060                     End If
26070                   Next  ' ** lngY.
                        ' ***************************************************************
26080                   dblPB_ThisStepSub = lngX
26090                   dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
26100                   ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                        'frm.ProgBar_bar.Width = dblPB_ThisWidthSub
26110                   strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
26120                   frm.ProgBar_lbl1.Caption = strPB_ThisPct
26130                   DoEvents
                        ' ***************************************************************
26140                 Next  ' ** lngX.
26150                 .Close
26160               End With  ' ** rst.
26170               Set rst = Nothing
26180               DoEvents

26190             End If

26200           End If  ' ** lngNones.

26210         End If  ' ** intMode.

              ' ** Check to see that transferred LedgerHidden records
              ' ** didn't bring over groups with mixed types.
26220         lngUniques = 0&
26230         ReDim arr_varUnique(U_ELEMS, 0)

              ' ***************************************************************
              ' ** Step 21: Check new hidden.
26240         dblPB_ThisStep = 21#
26250         .Status2_lbl.Caption = "Check new hidden"
26260         DoEvents
              ' ***************************************************************
              ' ***************************************************************
26270         dblPB_ThisWidth = 0#
26280         For dblZ = 1# To (dblPB_ThisStep - 1#)
                ' ** Assemble the weighted widths up to, but not including, this width.
26290           dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
26300         Next
              ' ***************************************************************

              ' ** qryAccountHideTrans2_63a (tblLedgerHidden, grouped by ledghid_grpnum,
              ' ** ledghidtype_type, with cnt), grouped by ledghid_grpnum, with cnt > 1.
26310         Set qdf = dbs.QueryDefs("qryAccountHideTrans2_63b")
26320         Set rst = qdf.OpenRecordset
26330         With rst
26340           If .BOF = True And .EOF = True Then
                  ' ** Good, there are no mixed ledghidtype_type's among the groups!
                  ' ***************************************************************
26350             dblPB_StepSubs = 0#  ' ** No subs in this step.
26360             dblPB_ThisIncrSub = 0#
26370             dblPB_ThisStepSub = 0#
26380             ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                  'frm.ProgBar_bar.Width = dblPB_ThisWidth
26390             strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
26400             frm.ProgBar_lbl1.Caption = strPB_ThisPct
26410             DoEvents
                  ' ***************************************************************
26420           Else
26430             .MoveLast
26440             lngRecs = .RecordCount
26450             .MoveFirst
                  ' ***************************************************************
26460             dblPB_StepSubs = lngRecs
26470             dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
26480             dblPB_ThisStepSub = 0#
26490             ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                  'frm.ProgBar_bar.Width = dblPB_ThisWidth
26500             DoEvents
                  ' ***************************************************************
26510             For lngX = 1& To lngRecs
26520               lngUniques = lngUniques + 1&
26530               lngE = lngUniques - 1&
26540               ReDim Preserve arr_varUnique(U_ELEMS, lngE)
26550               arr_varUnique(U_ACTNO, lngE) = ![accountno]
26560               arr_varUnique(U_GRP, lngE) = ![ledghid_grpnum]
26570               arr_varUnique(U_CNT, lngE) = ![ledghid_cnt]
                    ' ***************************************************************
26580               dblPB_ThisStepSub = lngX
26590               dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
26600               ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                    'frm.ProgBar_bar.Width = dblPB_ThisWidthSub
26610               strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
26620               frm.ProgBar_lbl1.Caption = strPB_ThisPct
26630               DoEvents
                    ' ***************************************************************
26640               If lngX < lngRecs Then .MoveNext
26650             Next
26660           End If
26670           .Close
26680         End With
26690         Set rst = Nothing
26700         Set qdf = Nothing
26710         DoEvents

              ' ** LedgHidType enumeration:
              ' **   0  GRP_NONE    Unmatched hidden transactions. This should only be a temporary designation.
              ' **   1  NORM        2 entries in hidden group, with matching assetno (which could both be zero).
              ' **   2  NORM_MISC   2 entries in hidden group, where both are 'Misc.', to be treated like a normal pair.
              ' **   3  MISC_2_GRP  2 entries in hidden group, one 'Misc.' and one other.
              ' **   4  MISC_3_GRP  3 entries in hidden group, one 'Misc.' and two other matching assetno.
              ' **   5  MULTI_GRP   3 or more entries in hidden group, with matching assetno, multi-lot group

              ' ***************************************************************
              ' ** Step 22: Determine hidden type.
26720         dblPB_ThisStep = 22#
26730         .Status2_lbl.Caption = "Determine hidden type"
26740         DoEvents
              'I THINK THIS IS THE LONG ONE!
              ' ***************************************************************
              ' ***************************************************************
26750         dblPB_ThisWidth = 0#
26760         For dblZ = 1# To (dblPB_ThisStep - 1#)
                ' ** Assemble the weighted widths up to, but not including, this width.
26770           dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
26780         Next
              ' ***************************************************************

              ' ** We'll have to figure out the correct ledghidtype_type.
26790         If lngUniques > 0& Then
                ' ***************************************************************
26800           dblPB_StepSubs = lngUniques
26810           dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)  ' ** The total width for just this step, divided by the sub steps.
26820           dblPB_ThisStepSub = 0#
26830           ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                '.ProgBar_bar.Width = dblPB_ThisWidth
26840           DoEvents
                ' ***************************************************************
26850           For lngX = 0& To (lngUniques - 1&)
                  ' ** tblLedgerHidden, all fields, by specified [grp].
26860             Set qdf = dbs.QueryDefs("qryAccountHideTrans2_60")
26870             With qdf.Parameters
26880               ![grp] = arr_varUnique(U_GRP, lngX)
26890             End With
26900             Set rst = qdf.OpenRecordset
26910             With rst
26920               .MoveLast
26930               lngRecs = .RecordCount
26940               .MoveFirst
26950               lngHidType = 0&
26960               If arr_varUnique(U_CNT, lngX) > 3& Then
26970                 lngHidType = 5&
26980                 For lngY = 1& To lngRecs
26990                   .Edit
27000                   ![ledghidtype_type] = lngHidType
27010                   .Update
27020                   If lngY < lngRecs Then .MoveNext
27030                 Next  ' ** lngY.
27040               ElseIf arr_varUnique(U_CNT, lngX) = 3& Then
27050                 If ![assetno] = 0& Then
27060                   .MoveNext
27070                   If ![assetno] = 0& Then
27080                     .MoveNext
27090                     If ![assetno] = 0& Then
27100                       lngHidType = 5&
27110                     Else
                            ' ** I don't have a designation for 2 Misc's and 1 Asset!
27120                       lngHidType = 5&
27130                     End If
27140                   Else
27150                     lngTmp01 = ![assetno]
27160                     .MoveNext
27170                     If ![assetno] = 0& Then
                            ' ** I don't have a designation for 2 Misc's and 1 Asset!
27180                       lngHidType = 5&
27190                     Else
                            ' ** They better be equal!
27200                       If ![assetno] = lngTmp01 Then
27210                         lngHidType = 4&
27220                       End If
27230                     End If
27240                   End If
27250                 Else
27260                   lngTmp01 = ![assetno]
27270                   .MoveNext
27280                   If ![assetno] = 0& Then
27290                     .MoveNext
27300                     If ![assetno] = 0& Then
                            ' ** I don't have a designation for 2 Misc's and 1 Asset!
27310                       lngHidType = 5&
27320                     Else
                            ' ** They better be equal!
27330                       If ![assetno] = lngTmp01 Then
27340                         lngHidType = 4&
27350                       End If
27360                     End If
27370                   Else
                          ' ** They better be equal!
27380                     If ![assetno] = lngTmp01 Then
27390                       .MoveNext
27400                       If ![assetno] = 0& Then
27410                         lngHidType = 4&
27420                       Else
                              ' ** They better be equal!
27430                         If ![assetno] = lngTmp01 Then
27440                           lngHidType = 5&
27450                         End If
27460                       End If
27470                     End If
27480                   End If
27490                 End If  ' ** If's without Else's will leave lngHidType = 0, 'GRP_NONE'!
27500                 .MoveFirst
27510                 For lngY = 1& To lngRecs
27520                   .Edit
27530                   ![ledghidtype_type] = lngHidType
27540                   .Update
27550                   If lngY < lngRecs Then .MoveNext
27560                 Next  ' ** lngY.
27570               ElseIf arr_varUnique(U_CNT, lngX) = 2& Then
27580                 If ![assetno] = 0& Then
27590                   .MoveNext
27600                   If ![assetno] = 0& Then
27610                     lngHidType = 2&
27620                   Else
27630                     lngHidType = 3&
27640                   End If
27650                 Else
27660                   lngTmp01 = ![assetno]
27670                   .MoveNext
27680                   If ![assetno] = 0& Then
27690                     lngHidType = 3&
27700                   Else
                          ' ** They better be equal!
27710                     If ![assetno] = lngTmp01 Then
27720                       lngHidType = 1&
27730                     End If
27740                   End If
27750                 End If  ' ** assetno.
27760                 .MoveFirst
27770                 For lngY = 1& To lngRecs
27780                   .Edit
27790                   ![ledghidtype_type] = lngHidType
27800                   .Update
27810                   If lngY < lngRecs Then .MoveNext
27820                 Next  ' ** lngY.
27830               End If  ' ** U_CNT.
27840               .Close
27850             End With  ' ** rst.
27860             Set rst = Nothing
27870             Set qdf = Nothing
27880             DoEvents
                  ' ***************************************************************
27890             dblPB_ThisStepSub = lngX
27900             dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub))
27910             ProgBar_Width_Hide frm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
                  '.ProgBar_bar.Width = dblPB_ThisWidthSub
27920             strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
27930             .ProgBar_lbl1.Caption = strPB_ThisPct
27940             DoEvents
                  ' ***************************************************************
27950           Next  ' ** lngX.

27960         End If  ' ** lngUniques.

              ' ** Clean up 'GRP_NONE' stragglers.

              ' ***************************************************************
              ' ** Step 23: Updating tables 1.
27970         dblPB_ThisStep = 23#
27980         .Status2_lbl.Caption = "Update tables 1"
27990         DoEvents
              ' ***************************************************************
              ' ***************************************************************
28000         dblPB_ThisWidth = 0#
28010         For dblZ = 1# To (dblPB_ThisStep - 1#)
                ' ** Assemble the weighted widths up to, but not including, this width.
28020           dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
28030         Next
28040         dblPB_StepSubs = 0#  ' ** No subs in this step.
28050         dblPB_ThisIncrSub = 0#
28060         dblPB_ThisStepSub = 0#
28070         ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
              '.ProgBar_bar.Width = dblPB_ThisWidth
28080         strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
28090         .ProgBar_lbl1.Caption = strPB_ThisPct
28100         DoEvents
              ' ***************************************************************

              ' ** Empty tblLedgerHidden_Staging3.
28110         Set qdf = dbs.QueryDefs("qryAccountHideTrans2_Hidden_10_11")
28120         qdf.Execute
28130         Set qdf = Nothing
28140         DoEvents

              ' ** tblLedgerHidden, just 'GRP_NONE', with journalno1, journalno2, journalno3.
28150         Set qdf = dbs.QueryDefs("qryAccountHideTrans2_Hidden_10_01")
28160         Set rst = qdf.OpenRecordset
28170         If rst.BOF = True And rst.EOF = True Then
                ' ** All are typed properly.
28180           rst.Close
28190           Set rst = Nothing
28200           Set qdf = Nothing
28210         Else
28220           rst.Close
28230           Set rst = Nothing
28240           Set qdf = Nothing

                ' ** Append qryAccountHideTrans2_Hidden_10_10 (qryAccountHideTrans2_Hidden_10_09
                ' ** (qryAccountHideTrans2_Hidden_10_08 (qryAccountHideTrans2_Hidden_10_07
                ' ** (qryAccountHideTrans2_Hidden_10_06 (qryAccountHideTrans2_Hidden_10_03
                ' ** (qryAccountHideTrans2_Hidden_10_02 (qryAccountHideTrans2_Hidden_10_01
                ' ** (tblLedgerHidden, just 'GRP_NONE', with journalno1, journalno2, journalno3),
                ' ** with grp_pos), just grp_pos = 1), linked to qryAccountHideTrans2_Hidden_10_04
                ' ** (qryAccountHideTrans2_Hidden_10_02 (qryAccountHideTrans2_Hidden_10_01
                ' ** (tblLedgerHidden, just 'GRP_NONE', with journalno1, journalno2, journalno3),
                ' ** with grp_pos), just grp_pos = 2), with ..1, ..2 fields), linked to
                ' ** qryAccountHideTrans2_Hidden_10_05 (qryAccountHideTrans2_Hidden_10_02
                ' ** (qryAccountHideTrans2_Hidden_10_01 (tblLedgerHidden, just 'GRP_NONE',
                ' ** with journalno1, journalno2, journalno3), with grp_pos), just grp_pos = 3),
                ' ** with ..3 fields), linked to Ledger 3 times, with ledghidtype_name_new1,
                ' ** ledghidtype_name_new2), with ledghidtype_name_new), linked to
                ' ** tblLedgerHiddenType, with ledghidtype_type_new) to tblLedgerHidden_Staging3.
28250           Set qdf = dbs.QueryDefs("qryAccountHideTrans2_Hidden_10_12")
28260           qdf.Execute
28270           Set qdf = Nothing
28280           DoEvents

                ' ***************************************************************
                ' ** Step 24: Updating tables 2.
28290           dblPB_ThisStep = 24#
28300           .Status2_lbl.Caption = "Update tables 2"
28310           DoEvents
                ' ***************************************************************
                ' ***************************************************************
28320           dblPB_ThisWidth = 0#
28330           For dblZ = 1# To (dblPB_ThisStep - 1#)
                  ' ** Assemble the weighted widths up to, but not including, this width.
28340             dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
28350           Next
28360           dblPB_StepSubs = 0#  ' ** No subs in this step.
28370           dblPB_ThisIncrSub = 0#
28380           dblPB_ThisStepSub = 0#
28390           ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                '.ProgBar_bar.Width = dblPB_ThisWidth
28400           strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
28410           .ProgBar_lbl1.Caption = strPB_ThisPct
28420           DoEvents
                ' ***************************************************************

                ' ** Update qryAccountHideTrans2_Hidden_10_13 (tblLedgerHidden,
                ' ** linked to tblLedgerHidden_Staging3, for journalno1).
28430           Set qdf = dbs.QueryDefs("qryAccountHideTrans2_Hidden_10_16")
28440           qdf.Execute
28450           Set qdf = Nothing
28460           DoEvents

                ' ***************************************************************
                ' ** Step 25: Updating tables 3.
28470           dblPB_ThisStep = 25#
28480           .Status2_lbl.Caption = "Update tables 3"
28490           DoEvents
                ' ***************************************************************
                ' ***************************************************************
28500           dblPB_ThisWidth = 0#
28510           For dblZ = 1# To (dblPB_ThisStep - 1#)
                  ' ** Assemble the weighted widths up to, but not including, this width.
28520             dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
28530           Next
28540           dblPB_StepSubs = 0#  ' ** No subs in this step.
28550           dblPB_ThisIncrSub = 0#
28560           dblPB_ThisStepSub = 0#
28570           ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                '.ProgBar_bar.Width = dblPB_ThisWidth
28580           strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
28590           .ProgBar_lbl1.Caption = strPB_ThisPct
28600           DoEvents
                ' ***************************************************************

                ' ** Update qryAccountHideTrans2_Hidden_10_14 (tblLedgerHidden,
                ' ** linked to tblLedgerHidden_Staging3, for journalno2).
28610           Set qdf = dbs.QueryDefs("qryAccountHideTrans2_Hidden_10_17")
28620           qdf.Execute
28630           Set qdf = Nothing
28640           DoEvents

                ' ***************************************************************
                ' ** Step 26: Updating tables 4.
28650           dblPB_ThisStep = 26#
28660           .Status2_lbl.Caption = "Update tables 4"
28670           DoEvents
                ' ***************************************************************
                ' ***************************************************************
28680           dblPB_ThisWidth = 0#
28690           For dblZ = 1# To (dblPB_ThisStep - 1#)
                  ' ** Assemble the weighted widths up to, but not including, this width.
28700             dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
28710           Next
28720           dblPB_StepSubs = 0#  ' ** No subs in this step.
28730           dblPB_ThisIncrSub = 0#
28740           dblPB_ThisStepSub = 0#
28750           ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                '.ProgBar_bar.Width = dblPB_ThisWidth
28760           strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
28770           .ProgBar_lbl1.Caption = strPB_ThisPct
28780           DoEvents
                ' ***************************************************************

                ' ** Update qryAccountHideTrans2_Hidden_10_15 (tblLedgerHidden,
                ' ** linked to tblLedgerHidden_Staging3, for journalno3).
28790           Set qdf = dbs.QueryDefs("qryAccountHideTrans2_Hidden_10_18")
28800           qdf.Execute
28810           Set qdf = Nothing
28820           DoEvents

                ' ***************************************************************
                ' ** Step 27:
28830           dblPB_ThisStep = 27#
28840           .Status2_lbl.Caption = "Additional matching"
28850           DoEvents
                ' ***************************************************************
                ' ***************************************************************
28860           dblPB_ThisWidth = 0#
28870           For dblZ = 1# To (dblPB_ThisStep - 1#)
                  ' ** Assemble the weighted widths up to, but not including, this width.
28880             dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
28890           Next
28900           dblPB_StepSubs = 16#
28910           dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / dblPB_StepSubs)
28920           dblPB_ThisStepSub = 0#
28930           DoEvents
                ' ***************************************************************

                'SO WHEN WOULD I USE THIS WITH AN ACCOUNTNO?
28940           Hide_AddlMatch "All", True, frm, dblPB_ThisIncrSub  ' ** Procedure: Below.
                'DO I REALLY WANT TO CHECK IF THIS ACCT HAS ALL AND REDO WITH ACCT IF NOT?

                ' ***************************************************************
                ' ** Step 28:
28950           dblPB_ThisStep = 28#
28960           .Status2_lbl.Caption = "Renumber groups"
28970           DoEvents
                ' ***************************************************************
                ' ***************************************************************
28980           dblPB_ThisWidth = 0#
28990           For dblZ = 1# To (dblPB_ThisStep - 1#)
                  ' ** Assemble the weighted widths up to, but not including, this width.
29000             dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
29010           Next
29020           dblPB_StepSubs = 0#  ' ** No subs in this step.
29030           dblPB_ThisIncrSub = 0#
29040           dblPB_ThisStepSub = 0#
29050           ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
                '.ProgBar_bar.Width = dblPB_ThisWidth
29060           strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
29070           .ProgBar_lbl1.Caption = strPB_ThisPct
29080           DoEvents
                ' ***************************************************************

29090           Hide_RenumGroups2  ' ** Procedure: Below.

29100         End If  ' ** BOF, EOF.

29110         dbs.Close
29120         DoEvents

29130       End If  ' ** blnLoad.

            ' ***************************************************************
            ' ** Step 29: Hidden processing finished.
29140       dblPB_ThisStep = 29#
29150       .Status2_lbl.Caption = "Hidden processing finished"
29160       DoEvents
            ' ***************************************************************
            ' ***************************************************************
29170       dblPB_ThisWidth = 0#
29180       For dblZ = 1# To (dblPB_ThisStep - 1#)
              ' ** Assemble the weighted widths up to, but not including, this width.
29190         dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
29200       Next
29210       dblPB_StepSubs = 0#  ' ** No subs in this step.
29220       dblPB_ThisIncrSub = 0#
29230       dblPB_ThisStepSub = 0#
29240       ProgBar_Width_Hide frm, dblPB_ThisWidth, 2  ' ** Procedure: Below.
            '.ProgBar_bar.Width = dblPB_ThisWidth
29250       strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
29260       .ProgBar_lbl1.Caption = strPB_ThisPct
29270       DoEvents
            ' ***************************************************************

29280       .chkHiddenFirstUse = True

29290     End If  ' ** hidden_trans.
29300   End With

EXITP:
29310   Set rst = Nothing
29320   Set qdf = Nothing
29330   Set dbs = Nothing
29340   Exit Sub

ERRH:
29350   Select Case ERR.Number
        Case Else
29360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29370   End Select
29380   Resume EXITP

End Sub

Public Sub Hide_AddlMatch(strAccountNo As String, Optional varProgBar As Variant, Optional varFrm As Variant, Optional varPB_Incr As Variant)
' ** I'm going to try to pick up a few stragglers.

29400 On Error GoTo ERRH

        Const THIS_PROC As String = "Hide_AddlMatch"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef
        Dim rst1 As DAO.Recordset, rst2 As DAO.Recordset, rst3 As DAO.Recordset
        Dim lngHids As Long, lngMatches As Long, lngRems As Long, lngHits As Long, lngAccts As Long
        Dim lngGrps As Long, arr_varGrp As Variant
        Dim lngRecs1 As Long, lngRecs2 As Long, lngMaxGrpNum As Long, lngHidType As Long
        Dim lngGrpElem As Long
        Dim blnProgBar As Boolean, dblPB_Incr As Double
        Dim strQryName1 As String, strQryName2 As String
        Dim blnFound As Boolean, lngAssetNo As Long
        Dim dblICash As Double, dblPCash As Double, dblCost As Double
        Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String, lngTmp04 As Long
        Dim lngX As Long, lngY As Long, lngZ As Long

        ' ** Array: arr_varGrp()
        Const G_ACTNO  As Integer = 0
        Const G_JTYP1  As Integer = 1
        Const G_JTYP2  As Integer = 2
        Const G_ASTNO1 As Integer = 3
        Const G_ASTNO2 As Integer = 4
        Const G_TDAT   As Integer = 5
        'Const G_HID    As Integer = 6
        Const G_ICSH   As Integer = 7
        Const G_PCSH   As Integer = 8
        Const G_COST   As Integer = 9
        Const G_JCNT   As Integer = 10
        Const G_ACNT   As Integer = 11
        Const G_MAX    As Integer = 12
        Const G_FND    As Integer = 13

        'SOME UNIQUEID'S IN tblLedgerHidden HAVE ODD JOURNALTYPE CONCATENATION!

        ' ** This function has n Steps.
29410   If IsMissing(varProgBar) = True Then
29420     blnProgBar = False
29430   Else
29440     blnProgBar = CBool(varProgBar)
29450     dblPB_Incr = CDbl(varPB_Incr)
29460   End If

29470   Set dbs = CurrentDb
29480   With dbs

29490     If strAccountNo = "All" Then

29500       If blnProgBar = True Then
              ' ***************************************************************
              ' ** Step 27.1
29510         dblPB_ThisStepSub = 1#
              ' ***************************************************************
              ' ***************************************************************
29520         dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
29530         ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
29540         strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
29550         varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
29560         DoEvents
              ' ***************************************************************
29570       End If  ' ** blnProgBar.

            ' ** qryAccountHideTrans2_34_02 (qryAccountHideTrans2_34_01
            ' ** (qryAccountHideTrans2_25 (qryAccountHideTrans2_24 (Union of
            ' ** qryAccountHideTrans2_24a (Ledger, just needed fields),
            ' ** qryAccountHideTrans2_24b (LedgerArchive, just needed fields)),
            ' ** just ledger_HIDDEN = True), all accounts), grouped, with cnt),
            ' ** grouped and summed, with cnt_actno.
29580       Set qdf1 = .QueryDefs("qryAccountHideTrans2_34_02_01")
29590       Set rst1 = qdf1.OpenRecordset
29600       With rst1
29610         If .BOF = True And .EOF = True Then
                ' ** Not likely at this point.
29620           lngHids = 0&
29630         Else
29640           .MoveFirst
29650           lngHids = ![cnt]
29660           lngAccts = ![cnt_actno]
29670         End If
29680         .Close
29690       End With  ' ** rst1.
29700       Set rst1 = Nothing
29710       Set qdf1 = Nothing
29720       DoEvents

29730       If lngHids > 0& Then

29740         If blnProgBar = True Then
                ' ***************************************************************
                ' ** Step 27.2
29750           dblPB_ThisStepSub = 2#
                ' ***************************************************************
                ' ***************************************************************
29760           dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
29770           ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
29780           strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
29790           varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
29800           DoEvents
                ' ***************************************************************
29810         End If  ' ** blnProgBar.

              ' ** qryAccountHideTrans2_34_04 (qryAccountHideTrans2_34_03 (tblLedgerHidden,
              ' ** all accounts), grouped, with cnt, grp_max), grouped and summed, with cnt.
29820         Set qdf1 = .QueryDefs("qryAccountHideTrans2_34_04_01")
29830         Set rst1 = qdf1.OpenRecordset
29840         With rst1
29850           If .BOF = True And .EOF = True Then
                  ' ** That's a whole lot of stragglers!
29860             lngMatches = 0&
29870           Else
29880             .MoveFirst
29890             lngMatches = ![cnt]
29900           End If
29910           .Close
29920         End With  ' ** rst1.
29930         Set rst1 = Nothing
29940         Set qdf1 = Nothing
29950         DoEvents

29960         If lngHids > lngMatches Then
                ' ** We've got some stragglers.

29970           If blnProgBar = True Then
                  ' ***************************************************************
                  ' ** Step 27.3
29980             dblPB_ThisStepSub = 3#
                  ' ***************************************************************
                  ' ***************************************************************
29990             dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
30000             ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
30010             strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
30020             varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
30030             DoEvents
                  ' ***************************************************************
30040           End If  ' ** blnProgBar.

                ' ** qryAccountHideTrans2_34_01 (qryAccountHideTrans2_25 (qryAccountHideTrans2_24 (Union
                ' ** of qryAccountHideTrans2_24a (Ledger, just needed fields), qryAccountHideTrans2_24b
                ' ** (LedgerArchive, just needed fields)), just ledger_HIDDEN = True), all accounts),
                ' ** not in qryAccountHideTrans2_34_03 (tblLedgerHidden, all accounts).
30050           Set qdf1 = .QueryDefs("qryAccountHideTrans2_34_05")
30060           Set rst1 = qdf1.OpenRecordset
30070           With rst1
30080             If .BOF = True And .EOF = True Then
                    ' ** Never happen.
30090               lngRems = 0&
30100             Else
30110               .MoveLast
30120               lngRems = .RecordCount
30130             End If
30140             .Close
30150           End With  ' ** rst1.
30160           Set rst1 = Nothing
30170           Set qdf1 = Nothing
30180           DoEvents

30190           If lngRems > 0& Then
                  ' ** So what can we do with these remaining unmatched hiddens.

30200             If blnProgBar = True Then
                    ' ***************************************************************
                    ' ** Step 27.4
30210               dblPB_ThisStepSub = 4#
                    ' ***************************************************************
                    ' ***************************************************************
30220               dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
30230               ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
30240               strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
30250               varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
30260               DoEvents
                    ' ***************************************************************
30270             End If  ' ** blnProgBar.

                  ' ** Start with something simple: total by date.
                  ' ** qryAccountHideTrans2_35_02 (qryAccountHideTrans2_35_01 (qryAccountHideTrans2_34_05
                  ' ** (qryAccountHideTrans2_34_01 (qryAccountHideTrans2_25 (qryAccountHideTrans2_24 (Union
                  ' ** of qryAccountHideTrans2_24a (Ledger, just needed fields), qryAccountHideTrans2_24b
                  ' ** (LedgerArchive, just needed fields)), just ledger_HIDDEN = True), all accounts),
                  ' ** not in qryAccountHideTrans2_34_03 (tblLedgerHidden, all accounts)), linked to
                  ' ** qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a (Ledger, just needed
                  ' ** fields), qryAccountHideTrans2_24b (LedgerArchive, just needed fields)), unmatched
                  ' ** hiddens), linked to qryAccountHideTrans2_35_02_02 (qryAccountHideTrans2_35_02_01
                  ' ** (qryAccountHideTrans2_35_01 (qryAccountHideTrans2_34_05 (qryAccountHideTrans2_34_01
                  ' ** (qryAccountHideTrans2_25 (qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a
                  ' ** (Ledger, just needed fields), qryAccountHideTrans2_24b (LedgerArchive, just needed
                  ' ** fields)), just ledger_HIDDEN = True), all accounts), not in qryAccountHideTrans2_34_03
                  ' ** (tblLedgerHidden, all accounts), linked to qryAccountHideTrans2_24 (Union of
                  ' ** qryAccountHideTrans2_24a (Ledger, just needed fields), qryAccountHideTrans2_24b
                  ' ** (LedgerArchive, just needed fields)), unmatched hiddens), grouped and summed, by
                  ' ** accountno, transdate, assetno, with cnt_jno), grouped and summed, by accountno,
                  ' ** transdate, with cnt_astno), grouped and summed, by accountno, transdate, with
                  ' ** journaltype1, journaltype2,  assetno1, assetno2, cnt_jno, cnt_astno), just
                  ' **  those that Zero-Out, cnt_astno <= 2.
30280             Set qdf1 = .QueryDefs("qryAccountHideTrans2_35_03")
30290             Set rst1 = qdf1.OpenRecordset
30300             With rst1
30310               If .BOF = True And .EOF = True Then
                      ' ** Ho-hum.
30320                 lngGrps = 0&
30330               Else
30340                 .MoveLast
30350                 lngGrps = .RecordCount
30360                 .MoveFirst
30370                 arr_varGrp = .GetRows(lngGrps)
                      ' **************************************************
                      ' ** Array: arr_varGrp()
                      ' **
                      ' **   Field  Element  Name             Constant
                      ' **   =====  =======  ===============  ==========
                      ' **     1       0     accountno        G_ACTNO
                      ' **     2       1     journaltype1     G_JTYP1
                      ' **     3       2     journaltype2     G_JTYP2
                      ' **     4       3     assetno1         G_ASTNO1
                      ' **     5       4     assetno2         G_ASTNO2
                      ' **     6       5     transdate        G_TDAT
                      ' **     7       6     ledger_HIDDEN    G_HID
                      ' **     8       7     icash            G_ICSH
                      ' **     9       8     pcash            G_PCSH
                      ' **    10       9     cost             G_COST
                      ' **    11      10     cnt_jno          G_JCNT
                      ' **    12      11     cnt_astno        G_ACNT
                      ' **    13      12     grp_max          G_MAX
                      ' **    14      13     Found            G_FND
                      ' **
                      ' **************************************************
30380               End If
30390               .Close
30400             End With  ' ** rst1.
30410             Set rst1 = Nothing
30420             Set qdf1 = Nothing
30430             DoEvents

30440             If lngGrps > 0& Then

30450               If blnProgBar = True Then
                      ' ***************************************************************
                      ' ** Step 27.5
30460                 dblPB_ThisStepSub = 5#
                      ' ***************************************************************
                      ' ***************************************************************
30470                 dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
30480                 ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
30490                 strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
30500                 varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
30510                 DoEvents
                      ' ***************************************************************
30520               End If  ' ** blnProgBar.

                    ' ** This vetting as actually already been done in the query.
30530               lngHits = 0&
30540               For lngX = 0& To (lngGrps - 1&)
30550                 If arr_varGrp(G_JCNT, lngX) > 1& And arr_varGrp(G_ACNT, lngX) <= 2& Then
30560                   arr_varGrp(G_FND, lngX) = CBool(True)
30570                   lngHits = lngHits + 1&
30580                 End If
30590               Next  ' ** lngX.
30600               DoEvents

30610               If lngHits > 0& Then
                      ' ** Each group found can go into lngLedgerHidden.

                      ' ** I've divided them up into subgroups.
30620                 For lngX = 1& To 11&

30630                   If blnProgBar = True Then
                          ' ***************************************************************
                          ' ** Step 27.6 - 27.16
30640                     dblPB_ThisStepSub = 5# + lngX
                          ' ***************************************************************
                          ' ***************************************************************
30650                     dblPB_ThisWidthSub = (dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_Incr))
30660                     ProgBar_Width_Hide varFrm, dblPB_ThisWidthSub, 2  ' ** Procedure: Below.
30670                     strPB_ThisPct = Format((dblPB_ThisWidthSub / dblPB_Width), "##0%")
30680                     varFrm.ProgBar_lbl1.Caption = strPB_ThisPct
30690                     DoEvents
                          ' ***************************************************************
30700                   End If  ' ** blnProgBar.

                        ' ** qryAccountHideTrans2_35_04_01.
30710                   strQryName1 = "qryAccountHideTrans2_35_04_" & Right("00" & CStr(lngX), 2)
30720                   Set qdf1 = .QueryDefs(strQryName1)
30730                   Set rst1 = qdf1.OpenRecordset
30740                   With rst1
30750                     If .BOF = True And .EOF = True Then
                            ' ** None with this count; I'm expecting 1, 2, and 3 to be empty.
30760                       lngRecs1 = 0&
30770                       DoEvents
30780                     Else
30790                       .MoveLast
30800                       lngRecs1 = .RecordCount
30810                       .MoveFirst
30820                       strQryName2 = strQryName1 & "_01"
30830                       DoEvents
30840                       For lngY = 1& To lngRecs1
30850                         lngGrpElem = -1&
30860                         For lngZ = 0& To (lngGrps - 1&)
30870                           If arr_varGrp(G_ACTNO, lngZ) = ![accountno] And arr_varGrp(G_TDAT, lngZ) = ![transdate] Then
30880                             lngGrpElem = lngZ
30890                             Exit For
30900                           End If
30910                         Next  ' ** lngZ.
30920                         DoEvents
30930                         Set qdf2 = dbs.QueryDefs(strQryName2)
30940                         With qdf2.Parameters
30950                           ![actno] = rst1![accountno]
30960                           ![tdat] = rst1![transdate]
30970                         End With
30980                         Set rst2 = qdf2.OpenRecordset
30990                         With rst2
31000                           .MoveLast
31010                           lngRecs2 = .RecordCount
31020                           .MoveFirst
31030                           DoEvents

31040                           lngHidType = 0&
31050                           If lngRecs2 = 2& Then
31060                             If arr_varGrp(G_ASTNO1, lngGrpElem) = arr_varGrp(G_ASTNO2, lngGrpElem) Then
31070                               If arr_varGrp(G_JTYP1, lngGrpElem) = "Misc." And arr_varGrp(G_JTYP2, lngGrpElem) = "Misc." Then
31080                                 lngHidType = 2  ' ** NORM_MISC
31090                               Else
31100                                 lngHidType = 1  ' ** NORM
31110                               End If
31120                             Else
31130                               lngHidType = 3    ' ** MISC_2_GRP
31140                             End If
31150                           ElseIf lngRecs2 = 3& Then
31160                             lngHidType = 4      ' ** MISC_3_GRP
31170                           Else
31180                             lngHidType = 5      ' ** MULTI_GRP
31190                           End If
31200                           DoEvents

                                ' ** Collect all the journalno's and journaltype's.
31210                           strTmp02 = vbNullString: strTmp03 = vbNullString
31220                           For lngZ = 1& To lngRecs1
31230                             strTmp02 = strTmp02 & CStr(![journalno]) & "_"
31240                             strTmp03 = strTmp03 & Left(![journaltype] & String(9, "_"), 9) & "_"  ' ** Total 10 chars per journaltype.
31250                             If lngZ < lngRecs2 Then .MoveNext
31260                             DoEvents
31270                           Next  ' ** lngY.
31280                           If Right(strTmp03, 1) = "_" Then strTmp03 = Left(strTmp03, (Len(strTmp03) - 1))
31290                           .MoveFirst
31300                           DoEvents

31310                           Set rst3 = dbs.OpenRecordset("tblLedgerHidden", dbOpenDynaset, dbAppendOnly)
31320                           For lngZ = 1& To lngRecs2
31330                             strTmp01 = Right(String(15, "0") & ![accountno], 15) & "_"
31340                             strTmp01 = strTmp01 & Right(String(4, "0") & CStr(![assetno]), 4) & "_"
31350                             strTmp01 = strTmp01 & strTmp02 & strTmp03
31360                             With rst3
31370                               .AddNew
                                    ' ** ![ledghid_id] : AutoNumber.
31380                               ![journalno] = rst2![journalno]
31390                               ![accountno] = rst2![accountno]
31400                               ![assetno] = rst2![assetno]
31410                               ![transdate] = rst2![transdate]
31420                               ![ledghid_cnt] = lngRecs2
31430                               varTmp00 = DMax("[ledghid_grpnum]", "tblLedgerHidden", "[accountno] = '" & CStr(rst2![accountno]) & "'")
31440                               If IsNull(varTmp00) = True Then
31450                                 varTmp00 = 0&
31460                               End If
31470                               ![ledghid_grpnum] = (varTmp00 + 1&)  'arr_varGrp(G_MAX, lngGrpElem) + 1&  'lngMaxGrpNum + 1&
31480                               ![ledghid_ord] = lngZ
31490                               ![ledghidtype_type] = lngHidType
31500                               ![ledghid_uniqueid] = strTmp01  ' ** ledghid_uniqueid is Memo type.
31510                               ![ledghid_username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
31520                               ![ledghid_datemodified] = Now()
31530                               .Update
31540                             End With  ' ** rst3
31550                             DoEvents
31560                             If lngZ < lngRecs2 Then .MoveNext
31570                           Next  ' ** lngZ.
31580                           rst3.Close
31590                           Set rst3 = Nothing
31600                           arr_varGrp(G_MAX, lngGrpElem) = (varTmp00 + 1&)
31610                           DoEvents

31620                           .Close
31630                         End With  ' ** rst2.
31640                         Set rst2 = Nothing
31650                         Set qdf2 = Nothing
31660                         DoEvents

31670                         If lngY < lngRecs1 Then .MoveNext
31680                       Next  ' ** lngY.

31690                     End If  ' ** BOF, EOF.
31700                     .Close
31710                   End With  ' ** rst1.
31720                   Set rst1 = Nothing
31730                   Set qdf1 = Nothing
31740                   DoEvents

31750                 Next  ' ** lngX.

31760               End If  ' ** lngHits.
31770             End If  ' ** lngGrps.
31780           End If  ' ** lngRems.
31790         End If  ' ** lngMatches.
31800       End If  ' ** lngHids.

31810     Else

            ' ** qryAccountHideTrans2_32_01 (qryAccountHideTrans2_25 (qryAccountHideTrans2_24
            ' ** (Union of qryAccountHideTrans2_24a (Ledger, just needed fields),
            ' ** qryAccountHideTrans2_24b (LedgerArchive, just needed fields)),
            ' ** just ledger_HIDDEN = True), by specified [actno]), grouped, with cnt.
31820       Set qdf1 = .QueryDefs("qryAccountHideTrans2_32_02")
31830       With qdf1.Parameters
31840         ![actno] = strAccountNo
31850       End With
31860       Set rst1 = qdf1.OpenRecordset
31870       With rst1
31880         If .BOF = True And .EOF = True Then
                ' ** Not likely at this point.
31890           lngHids = 0&
31900         Else
31910           .MoveFirst
31920           lngHids = ![cnt]
31930         End If
31940         .Close
31950       End With
31960       Set rst1 = Nothing
31970       Set qdf1 = Nothing
31980       DoEvents

31990       If lngHids > 0& Then

              ' ** qryAccountHideTrans2_32_03 (tblLedgerHidden, by specified [actno]), grouped, with cnt.
32000         Set qdf1 = .QueryDefs("qryAccountHideTrans2_32_04")
32010         With qdf1.Parameters
32020           ![actno] = strAccountNo
32030         End With
32040         Set rst1 = qdf1.OpenRecordset
32050         With rst1
32060           If .BOF = True And .EOF = True Then
                  ' ** That's a whole lot of stragglers!
32070             lngMatches = 0&
32080           Else
32090             .MoveFirst
32100             lngMatches = ![cnt]
32110             lngMaxGrpNum = ![grp_max]
32120           End If
32130           .Close
32140         End With
32150         Set rst1 = Nothing
32160         Set qdf1 = Nothing
32170         DoEvents

32180         If lngHids > lngMatches Then
                ' ** We've got some stragglers.

                ' ** qryAccountHideTrans2_32_01 (qryAccountHideTrans2_25 (qryAccountHideTrans2_24
                ' ** (Union of qryAccountHideTrans2_24a (Ledger, just needed fields),
                ' ** qryAccountHideTrans2_24b (LedgerArchive, just needed fields)), just
                ' ** ledger_HIDDEN = True), by specified [actno]), not in qryAccountHideTrans2_32_03
                ' ** (tblLedgerHidden, by specified [actno]).
32190           Set qdf1 = .QueryDefs("qryAccountHideTrans2_32_05")
32200           With qdf1.Parameters
32210             ![actno] = strAccountNo
32220           End With
32230           Set rst1 = qdf1.OpenRecordset
32240           With rst1
32250             If .BOF = True And .EOF = True Then
                    ' ** Never happen.
32260               lngRems = 0&
32270             Else
32280               .MoveLast
32290               lngRems = .RecordCount
32300             End If
32310             .Close
32320           End With
32330           Set rst1 = Nothing
32340           Set qdf1 = Nothing
32350           DoEvents

32360           If lngRems > 0& Then
                  ' ** So what can we do with these remaining unmatched hiddens.

                  ' ** Start with something simple: total by date.
                  ' ** qryAccountHideTrans2_33_01 (qryAccountHideTrans2_32_05 (qryAccountHideTrans2_32_01
                  ' ** (qryAccountHideTrans2_25 (qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a
                  ' ** (Ledger, just needed fields), qryAccountHideTrans2_24b (LedgerArchive, just
                  ' ** needed fields)), just ledger_HIDDEN = True), by specified [actno]), not in
                  ' ** qryAccountHideTrans2_32_03 (tblLedgerHidden, by specified [actno])), linked to
                  ' ** qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a (Ledger, just needed
                  ' ** fields), qryAccountHideTrans2_24b (LedgerArchive, just needed fields)),
                  ' ** unmatched hiddens), grouped and summed, by accountno, transdate,
                  ' ** with journaltype1, journaltype2, assetno1, assetno2.
32370             Set qdf1 = .QueryDefs("qryAccountHideTrans2_33_02")
32380             With qdf1.Parameters
32390               ![actno] = strAccountNo
32400             End With
32410             Set rst1 = qdf1.OpenRecordset
32420             With rst1
32430               If .BOF = True And .EOF = True Then
                      ' ** Ho-hum.
32440                 lngGrps = 0&
32450               Else
32460                 .MoveLast
32470                 lngGrps = .RecordCount
32480                 .MoveFirst
32490                 arr_varGrp = .GetRows(lngGrps)
                      ' **************************************************
                      ' ** Array: arr_varGrp()
                      ' **
                      ' **   Field  Element  Name             Constant
                      ' **   =====  =======  ===============  ==========
                      ' **     1       0     accountno        G_ACTNO
                      ' **     2       1     journaltype1     G_JTYP1
                      ' **     3       2     journaltype2     G_JTYP2
                      ' **     4       3     assetno1         G_ASTNO1
                      ' **     5       4     assetno2         G_ASTNO2
                      ' **     6       5     transdate        G_TDAT
                      ' **     7       6     ledger_HIDDEN    G_HID
                      ' **     8       7     icash            G_ICSH
                      ' **     9       8     pcash            G_PCSH
                      ' **    10       9     cost             G_COST
                      ' **    11      10     cnt_jno          G_JCNT
                      ' **    12      11     cnt_astno        G_ACNT
                      ' **    13      12     Found            G_FND
                      ' **
                      ' **************************************************
32500               End If
32510               .Close
32520             End With
32530             Set rst1 = Nothing
32540             Set qdf1 = Nothing
32550             DoEvents

32560             If lngGrps > 0& Then

32570               lngHits = 0&
32580               For lngX = 0& To (lngGrps - 1&)
32590                 If arr_varGrp(G_JCNT, lngX) > 1& Then
32600                   If Round(arr_varGrp(G_ICSH, lngX), 2) = 0 And Round(arr_varGrp(G_PCSH, lngX), 2) = 0 And _
                            Round(arr_varGrp(G_COST, lngX), 2) = 0 Then
32610                     If (arr_varGrp(G_ASTNO1, lngX) = arr_varGrp(G_ASTNO2, lngX)) Or _
                              (arr_varGrp(G_ASTNO1, lngX) = 0 And arr_varGrp(G_ASTNO2, lngX) > 0) Or _
                              (arr_varGrp(G_ASTNO1, lngX) > 0 And arr_varGrp(G_ASTNO2, lngX) = 0) Then
                            ' ** We've got a weiner!
32620                       arr_varGrp(G_FND, lngX) = CBool(True)
32630                       lngHits = lngHits + 1&
32640                     End If
32650                   End If
32660                 End If
32670               Next  ' ** lngX.
32680               DoEvents

32690               If lngHits > 0& Then
                      ' ** Each group found can go into lngLedgerHidden.

32700                 For lngX = 0& To (lngGrps - 1&)
32710                   If arr_varGrp(G_FND, lngX) = True Then

                          ' ** qryAccountHideTrans2_33_01 (qryAccountHideTrans2_32_05 (qryAccountHideTrans2_32_01
                          ' ** (qryAccountHideTrans2_25 (qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a
                          ' ** (Ledger, just needed fields), qryAccountHideTrans2_24b (LedgerArchive, just
                          ' ** needed fields)), just ledger_HIDDEN = True), by specified [actno]), not in
                          ' ** qryAccountHideTrans2_32_03 (tblLedgerHidden, by specified [actno])), linked to
                          ' ** qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a (Ledger, just needed
                          ' ** fields), qryAccountHideTrans2_24b (LedgerArchive, just needed fields)),
                          ' ** unmatched hiddens), by specified [tdat].
32720                     Set qdf1 = .QueryDefs("qryAccountHideTrans2_33_03")
32730                     With qdf1.Parameters
32740                       ![actno] = strAccountNo
32750                       ![tdat] = arr_varGrp(G_TDAT, lngX)
32760                     End With
32770                     Set rst1 = qdf1.OpenRecordset
32780                     With rst1
32790                       .MoveLast
32800                       lngRecs1 = .RecordCount
32810                       .MoveFirst
32820                       DoEvents

32830                       lngHidType = 0&
32840                       If lngRecs1 = 2& Then
32850                         If arr_varGrp(G_ASTNO1, lngX) = arr_varGrp(G_ASTNO2, lngX) Then
32860                           If arr_varGrp(G_JTYP1, lngX) = "Misc." And arr_varGrp(G_JTYP2, lngX) = "Misc." Then
32870                             lngHidType = 2  ' ** NORM_MISC
32880                           Else
32890                             lngHidType = 1  ' ** NORM
32900                           End If
32910                         Else
32920                           lngHidType = 3    ' ** MISC_2_GRP
32930                         End If
32940                       ElseIf lngRecs1 = 3& Then
                              ' ** AssetNo's have to be checked!
32950                         lngHidType = 4      ' ** MISC_3_GRP
32960                       Else
                              ' ** AssetNo's have to be checked!
32970                         lngHidType = 5      ' ** MULTI_GRP
32980                       End If
32990                       DoEvents

                            ' ** Collect all the journalno's and journaltype's.
33000                       strTmp02 = vbNullString: strTmp03 = vbNullString
33010                       For lngY = 1& To lngRecs1
33020                         strTmp02 = strTmp02 & Right(CStr(![journalno]) & String(6, "0"), 6) & "_"  ' ** Total 6 chars per journalno.
33030                         strTmp03 = strTmp03 & Left(![journaltype] & String(9, "_"), 9) & "_"  ' ** Total 10 chars per journaltype.
33040                         If lngY < lngRecs1 Then .MoveNext
33050                       Next  ' ** lngY.
33060                       If Right(strTmp03, 1) = "_" Then strTmp03 = Left(strTmp03, (Len(strTmp03) - 1))
33070                       .MoveFirst
33080                       DoEvents

33090                       Set rst2 = dbs.OpenRecordset("tblLedgerHidden", dbOpenDynaset, dbAppendOnly)
33100                       For lngY = 1& To lngRecs1
33110                         strTmp01 = Right(String(15, "0") & ![accountno], 15) & "_"
33120                         strTmp01 = strTmp01 & Right(String(4, "0") & CStr(![assetno]), 4) & "_"
33130                         strTmp01 = strTmp01 & strTmp02 & strTmp03
33140                         With rst2
33150                           .AddNew
                                ' ** ![ledghid_id] : AutoNumber.
33160                           ![journalno] = rst1![journalno]
33170                           ![accountno] = rst1![accountno]
33180                           ![assetno] = rst1![assetno]
33190                           ![transdate] = rst1![transdate]
33200                           ![ledghid_cnt] = lngRecs1
33210                           ![ledghid_grpnum] = lngMaxGrpNum + 1&
33220                           ![ledghid_ord] = lngY
33230                           ![ledghidtype_type] = lngHidType
33240                           ![ledghid_uniqueid] = strTmp01
33250                           ![ledghid_username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
33260                           ![ledghid_datemodified] = Now()
33270                           .Update
33280                         End With  ' ** rst2
33290                         DoEvents
33300                         If lngY < lngRecs1 Then .MoveNext
33310                       Next  ' ** lngY.

33320                       rst2.Close
33330                       Set rst2 = Nothing

33340                       .Close
33350                     End With  ' ** rst1
33360                     Set rst1 = Nothing
33370                     Set qdf1 = Nothing
33380                     DoEvents

                          ' ** LedgHidType enumeration:
                          ' **   0  GRP_NONE    Unmatched hidden transactions. This should only be a temporary designation.
                          ' **   1  NORM        2 entries in hidden group, with matching assetno (which could both be zero).
                          ' **   2  NORM_MISC   2 entries in hidden group, where both are 'Misc.', to be treated like a normal pair.
                          ' **   3  MISC_2_GRP  2 entries in hidden group, one 'Misc.' and one other.
                          ' **   4  MISC_3_GRP  3 entries in hidden group, one 'Misc.' and two other matching assetno.
                          ' **   5  MULTI_GRP   3 or more entries in hidden group, with matching assetno (maybe), multi-lot group

33390                   End If  ' ** G_FND.
33400                 Next  ' ** lngX.

33410               Else
                      'ANY MORE?

33420                 lngTmp04 = 0&
33430                 For lngX = 0& To (lngGrps - 1&)
33440                   lngTmp04 = lngTmp04 + arr_varGrp(G_JCNT, lngX)
33450                 Next  ' ** lngX.
33460                 DoEvents

                      ' ** Rather confined, but it'll catch some.
33470                 If lngTmp04 = 3& Or lngTmp04 = 4& Or lngTmp04 = 5& Or lngTmp04 = 6& Or lngTmp04 = 14& Then

                        ' ** Looking for 'Dividend' and 'Misc.' Or 'Interest' and 'Misc.'.
33480                   lngAssetNo = 0&
33490                   For lngX = 0& To (lngGrps - 1&)
33500                     Select Case lngTmp04
                          Case 3&
33510                       If (arr_varGrp(G_JTYP1, lngX) = "Dividend" And arr_varGrp(G_JTYP2, lngX) = "Dividend") Or _
                                (arr_varGrp(G_JTYP1, lngX) = "Interest" And arr_varGrp(G_JTYP2, lngX) = "Interest") Or _
                                (arr_varGrp(G_JTYP1, lngX) = "Misc." And arr_varGrp(G_JTYP2, lngX) = "Misc.") Then
33520                         arr_varGrp(G_FND, lngX) = CBool(True)
33530                         If (arr_varGrp(G_JTYP1, lngX) = "Dividend" And arr_varGrp(G_JTYP2, lngX) = "Dividend") Or _
                                  (arr_varGrp(G_JTYP1, lngX) = "Interest" And arr_varGrp(G_JTYP2, lngX) = "Interest") Then
33540                           If arr_varGrp(G_ASTNO1, lngX) = arr_varGrp(G_ASTNO2, lngX) Then
33550                             lngAssetNo = arr_varGrp(G_ASTNO1, lngX)
33560                           End If
33570                         End If
33580                       End If
33590                     Case 4&, 5&
33600                       If ((arr_varGrp(G_JTYP1, lngX) = "Paid" And arr_varGrp(G_JTYP2, lngX) = "Paid") Or _
                                (arr_varGrp(G_JTYP1, lngX) = "Misc." And arr_varGrp(G_JTYP2, lngX) = "Paid") Or _
                                (arr_varGrp(G_JTYP1, lngX) = "Paid" And arr_varGrp(G_JTYP2, lngX) = "Misc.")) Or _
                                ((arr_varGrp(G_JTYP1, lngX) = "Interest" And arr_varGrp(G_JTYP2, lngX) = "Interest") Or _
                                (arr_varGrp(G_JTYP1, lngX) = "Interest" And arr_varGrp(G_JTYP2, lngX) = "Misc.") Or _
                                (arr_varGrp(G_JTYP1, lngX) = "Misc." And arr_varGrp(G_JTYP2, lngX) = "Interest")) Then
33610                         arr_varGrp(G_FND, lngX) = CBool(True)
33620                         lngAssetNo = 0&
33630                       End If
33640                     Case 6&
33650                       If (arr_varGrp(G_JTYP1, lngX) = "Paid" And arr_varGrp(G_JTYP2, lngX) = "Paid") Or _
                                (arr_varGrp(G_JTYP1, lngX) = "Misc." And arr_varGrp(G_JTYP2, lngX) = "Misc.") Then
33660                         arr_varGrp(G_FND, lngX) = CBool(True)
33670                         lngAssetNo = 0&
33680                       End If
33690                     Case 14&
33700                       If (arr_varGrp(G_JTYP1, lngX) = "Dividend" And arr_varGrp(G_JTYP2, lngX) = "Dividend") Or _
                                (arr_varGrp(G_JTYP1, lngX) = "Dividend" And arr_varGrp(G_JTYP2, lngX) = "Misc.") Or _
                                (arr_varGrp(G_JTYP1, lngX) = "Misc." And arr_varGrp(G_JTYP2, lngX) = "Dividend") Then
33710                         arr_varGrp(G_FND, lngX) = CBool(True)
33720                         lngAssetNo = 0&
33730                       End If
33740                     End Select  ' ** lngTmp04.
33750                   Next  ' ** lngX.
33760                   DoEvents

33770                   blnFound = True
33780                   For lngX = 0& To (lngGrps - 1&)
33790                     If arr_varGrp(G_FND, lngX) = False Then
                            ' ** Nope, not going to work.
33800                       blnFound = False
33810                       Exit For
33820                     End If
33830                   Next  ' ** lngX.
33840                   DoEvents

33850                   If blnFound = True Then
                          ' ** All 3 (or 4) fit the criteria, so lets add them up.

33860                     dblICash = 0#: dblPCash = 0#: dblCost = 0#
33870                     lngHits = 0&
33880                     For lngX = 0& To (lngGrps - 1&)
33890                       dblICash = dblICash + Round(arr_varGrp(G_ICSH, lngX), 2)
33900                       dblPCash = dblPCash + Round(arr_varGrp(G_PCSH, lngX), 2)
33910                       dblCost = dblCost + Round(arr_varGrp(G_COST, lngX), 2)
33920                     Next  ' ** lngX.
33930                     DoEvents

33940                     If dblICash = 0# And dblPCash = 0# And dblCost = 0# Then
                            ' ** We've got another weiner!
33950                       lngHits = lngHits + 1&
33960                     End If

33970                     If lngHits > 0& Then

33980                       lngHidType = 5  ' ** MULTI_GRP

                            ' ** qryAccountHideTrans2_32_05 (qryAccountHideTrans2_32_01 (qryAccountHideTrans2_25
                            ' ** (qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a (Ledger, just needed fields),
                            ' ** qryAccountHideTrans2_24b (LedgerArchive, just needed fields)), just ledger_HIDDEN = True),
                            ' ** by specified [actno]), not in qryAccountHideTrans2_32_03 (tblLedgerHidden, by specified
                            ' ** [actno])), linked to qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a (Ledger,
                            ' ** just needed fields), qryAccountHideTrans2_24b (LedgerArchive, just needed fields)),
                            ' ** unmatched hiddens
33990                       Set qdf1 = .QueryDefs("qryAccountHideTrans2_33_01")
34000                       With qdf1.Parameters
34010                         ![actno] = strAccountNo
34020                       End With
34030                       Set rst1 = qdf1.OpenRecordset
34040                       With rst1
34050                         .MoveLast
34060                         lngRecs1 = .RecordCount
34070                         .MoveFirst
34080                         DoEvents

                              ' ** Collect all the journalno's and journaltype's.
34090                         strTmp02 = vbNullString: strTmp03 = vbNullString
34100                         For lngY = 1& To lngRecs1
34110                           strTmp02 = strTmp02 & Right(CStr(![journalno]) & String(6, "0"), 6) & "_"  ' ** Total 7 chars per journalno.
34120                           strTmp03 = strTmp03 & Left(![journaltype] & String(9, "_"), 9) & "_"  ' ** Total 10 chars per journaltype.
34130                           If lngY < lngRecs1 Then .MoveNext
34140                         Next  ' ** lngRecs1.
34150                         If Right(strTmp03, 1) = "_" Then strTmp03 = Left(strTmp03, (Len(strTmp03) - 1))  ' ** Last one has extra underscore.
34160                         .MoveFirst
34170                         DoEvents

34180                         lngMaxGrpNum = DMax("[ledghid_grpnum]", "tblLedgerHidden", "[accountno] = '" & strAccountNo & "'")

34190                         Set rst2 = dbs.OpenRecordset("tblLedgerHidden", dbOpenDynaset, dbAppendOnly)
34200                         For lngY = 1& To lngRecs1
34210                           strTmp01 = Right(String(15, "0") & ![accountno], 15) & "_"
                                'WOULD IT HAVE ACCEPTED THIS NOW?
                                ' ** Since each in group should have identical ledghid_uniqueid,
                                ' ** they all get Zeroes if the 2 don't match.
34220                           strTmp01 = strTmp01 & Right(String(4, "0") & CStr(lngAssetNo), 4) & "_"
34230                           strTmp01 = strTmp01 & strTmp02 & strTmp03
34240                           With rst2
34250                             .AddNew
                                  ' ** ![ledghid_id] : AutoNumber.
34260                             ![journalno] = rst1![journalno]
34270                             ![accountno] = rst1![accountno]
34280                             ![assetno] = rst1![assetno]
34290                             ![transdate] = rst1![transdate]
34300                             ![ledghid_cnt] = lngRecs1
34310                             ![ledghid_grpnum] = lngMaxGrpNum + 1&
34320                             ![ledghid_ord] = lngY
34330                             ![ledghidtype_type] = lngHidType
34340                             ![ledghid_uniqueid] = strTmp01
34350                             ![ledghid_username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
34360                             ![ledghid_datemodified] = Now()
34370                             .Update
34380                           End With  ' ** rst2
34390                           DoEvents
34400                           If lngY < lngRecs1 Then .MoveNext
34410                         Next  ' ** lngY.

34420                         rst2.Close
34430                         Set rst2 = Nothing

34440                         .Close
34450                       End With  ' ** rst1
34460                       Set rst1 = Nothing
34470                       Set qdf1 = Nothing
34480                       DoEvents

34490                     End If  ' ** lngHits.
34500                   End If  ' ** blnFound.
34510                 End If  ' ** lngTmp04.

34520               End If  ' ** lngHits.

34530             End If  ' ** lngGrps.
34540           End If  ' ** lngRems.
34550         End If  ' ** lngMatches.
34560       End If  ' ** lngHids.

34570     End If  ' ** strAccountNo.

34580     .Close
34590   End With  ' ** dbs.

EXITP:
34600   Set rst1 = Nothing
34610   Set rst2 = Nothing
34620   Set rst3 = Nothing
34630   Set qdf1 = Nothing
34640   Set qdf2 = Nothing
34650   Set dbs = Nothing
34660   Exit Sub

ERRH:
34670   Select Case ERR.Number
        Case Else
34680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34690   End Select
34700   Resume EXITP

End Sub

Public Sub Hide_RenumGroups2()
' ** Renumber the groups so each accountno starts at 1.

34800 On Error GoTo ERRH

        Const THIS_PROC As String = "Hide_RenumGroups2"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngGrps As Long, arr_varGrp As Variant
        Dim lngHids As Long, arr_varHid As Variant
        Dim lngX As Long, lngY As Long

        ' ** Array: arr_varGrp()
        Const G_ACTNO As Integer = 0
        Const G_GRP1  As Integer = 1
        Const G_GRP2  As Integer = 2
        'Const G_JCNT  As Integer = 3
        Const G_GCNT  As Integer = 4

        ' ** Array: arr_varHid()
        Const H_ACTNO As Integer = 0
        Const H_GNUM  As Integer = 1
        Const H_NEW   As Integer = 2
        'Const H_CNT   As Integer = 3

34810   DoCmd.Hourglass True  ' ** Make sure this is still running.
34820   DoEvents

34830   Set dbs = CurrentDb
34840   With dbs

          ' ** qryAccountHideTrans2_36_01 (tblLedgerHidden, grouped, by accountno,
          ' ** ledghid_grpnum, with cnt_jno), grouped and summed, by accountno,
          ' ** with ledghid_grpnum1, ledghid_grpnum2, cnt_grp.
34850     Set qdf = .QueryDefs("qryAccountHideTrans2_36_02")
34860     Set rst = qdf.OpenRecordset
34870     With rst
34880       If .BOF = True And .EOF = True Then
              ' ** tblLedgerHidden is empty.
34890         lngGrps = 0&
34900       Else
34910         .MoveLast
34920         lngGrps = .RecordCount
34930         .MoveFirst
34940         arr_varGrp = .GetRows(lngGrps)
              ' ****************************************************
              ' ** Array: arr_varGrp()
              ' **
              ' **   Field  Element  Name               Constant
              ' **   =====  =======  =================  ==========
              ' **     1       0     accountno          G_ACTNO
              ' **     2       1     ledghid_grpnum1    G_GRP1
              ' **     3       2     ledghid_grpnum2    G_GRP2
              ' **     4       3     cnt_jno            G_JCNT
              ' **     5       4     cnt_grp            G_GCNT
              ' **
              ' ****************************************************
34950       End If
34960       .Close
34970     End With  ' ** rst.
34980     Set rst = Nothing
34990     Set qdf = Nothing
35000     DoEvents

35010     If lngGrps > 0& Then

35020       For lngX = 0& To (lngGrps - 1&)
35030         If arr_varGrp(G_GRP1, lngX) = 1 And arr_varGrp(G_GRP2, lngX) = arr_varGrp(G_GCNT, lngX) Then
                ' ** These are numbered as they should be.
35040         Else

35050           lngHids = 0&
35060           arr_varHid = Empty

                ' ** qryAccountHideTrans2_36_01 (tblLedgerHidden, grouped, by accountno,
                ' ** ledghid_grpnum, with cnt_jno), by specified [actno].
35070           Set qdf = .QueryDefs("qryAccountHideTrans2_36_04")
35080           With qdf.Parameters
35090             ![actno] = arr_varGrp(G_ACTNO, lngX)
35100           End With
35110           Set rst = qdf.OpenRecordset
35120           With rst
35130             .MoveLast
35140             lngHids = .RecordCount
35150             .MoveFirst
35160             arr_varHid = .GetRows(lngHids)
                  ' *******************************************************
                  ' ** Array: arr_varHid()
                  ' **
                  ' **   Field  Element  Name                  Constant
                  ' **   =====  =======  ====================  ==========
                  ' **     1       0     accountno             H_ACTNO
                  ' **     2       1     ledghid_grpnum        H_GNUM
                  ' **     3       2     ledghid_grpnum_new    H_NEW
                  ' **     4       3     cnt_jno               H_CNT
                  ' **
                  ' *******************************************************
35170             .Close
35180           End With  ' ** rst.
35190           Set rst = Nothing
35200           Set qdf = Nothing
35210           DoEvents

                ' ** Renumber groups, starting at 1.
35220           For lngY = 0& To (lngHids - 1&)
35230             arr_varHid(H_NEW, lngY) = (lngY + 1&)
35240           Next  ' ** lngY.

35250           For lngY = 0& To (lngHids - 1&)
                  ' ** Update qryAccountHideTrans2_36_05 (tblLedgerHidden,
                  ' ** by specified [actno], [grpold], [grpnew]).
35260             Set qdf = .QueryDefs("qryAccountHideTrans2_36_06")
35270             With qdf.Parameters
35280               ![actno] = arr_varHid(H_ACTNO, lngY)
35290               ![grpold] = arr_varHid(H_GNUM, lngY)
35300               ![grpnew] = arr_varHid(H_NEW, lngY)
35310             End With
35320             qdf.Execute dbFailOnError
35330             Set qdf = Nothing
35340             DoEvents
35350           Next  ' ** lngY.

35360         End If

35370       Next  ' ** lngX.

35380     End If  ' ** lngGrps.

35390     .Close
35400   End With  ' ** dbs.

EXITP:
35410   Set rst = Nothing
35420   Set qdf = Nothing
35430   Set dbs = Nothing
35440   Exit Sub

ERRH:
35450   Select Case ERR.Number
        Case Else
35460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
35470   End Select
35480   Resume EXITP

End Sub

Public Sub ProgBar_Width_Hide(varFrm As Variant, dblWidth As Double, intMode As Integer)

35500 On Error GoTo ERRH

        Const THIS_PROC As String = "ProgBar_Width_Hide"

        Dim strCtlName As String, blnVis As Boolean
        Dim lngX As Long

35510   With varFrm
35520     Select Case intMode
          Case 1
35530       blnVis = CBool(dblWidth)
35540       For lngX = 1& To 6&
35550         strCtlName = "ProgBar_bar" & Right("00" & CStr(lngX), 2)
35560         .Controls(strCtlName).Visible = blnVis
35570       Next
35580     Case 2
35590       For lngX = 1& To 6&
35600         strCtlName = "ProgBar_bar" & Right("00" & CStr(lngX), 2)
35610         .Controls(strCtlName).Width = dblWidth
35620       Next
35630     Case 3
35640       For lngX = 1& To 6&
35650         strCtlName = "ProgBar_bar" & Right("00" & CStr(lngX), 2)
35660         .Controls(strCtlName).Left = (.Controls(strCtlName).Left + dblWidth)
35670       Next
35680     End Select
35690   End With

EXITP:
35700   Exit Sub

ERRH:
35710   Select Case ERR.Number
        Case Else
35720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
35730   End Select
35740   Resume EXITP

End Sub
