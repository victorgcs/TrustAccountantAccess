Attribute VB_Name = "modHideTransactions2"
Option Compare Database
Option Explicit

'VGC 10/29/2017: CHANGES!

Private Const THIS_NAME As String = "modHideTransactions2"
' **

Public Function Hide_FixQrys() As Boolean

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Hide_FixQrys"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngActNos As Long, arr_varActNo As Variant
        Dim strQryBase1 As String, strQryBase2 As String
        Dim strQryName As String, strAccountNo1 As String, strAccountNo2 As String, strSQL As String, strDesc As String
        Dim lngQrysCreated As Long, lngRecs As Long
        Dim intPos01 As Integer
        Dim lngX As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varActNo().
        Const A_ACTNO As Integer = 0
        'Const A_CNT   As Integer = 1
        Const A_QRY1  As Integer = 2
        Const A_QRY2  As Integer = 3

110   On Error GoTo 0

120     blnRetVal = True

130     Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
140     DoEvents

150     Set dbs = CurrentDb
160     With dbs

          ' ** zzz_qry_zFiduciary_31_04_01 (zzz_qry_zFiduciary_31_04 (zzz_qry_zFiduciary_31_02
          ' ** (zzz_qry_zFiduciary_31_01 (tblLedgerHidden, grouped, by accountno, ledghid_grpnum,
          ' ** with cnt_jno), just discrepancies), just ledghid_cnt discrepancies),
          ' ** grouped by accountno, with cnt_cnt), just cnt_cnt = 1.
170       Set qdf = .QueryDefs("zzz_qry_zFiduciary_31_04_02")
          ' ** zzz_qry_zFiduciary_31_04_01 (zzz_qry_zFiduciary_31_04 (zzz_qry_zFiduciary_31_02
          ' ** (zzz_qry_zFiduciary_31_01 (tblLedgerHidden, grouped, by accountno, ledghid_grpnum,
          ' ** with cnt_jno), just discrepancies), just ledghid_cnt discrepancies),
          ' ** grouped by accountno, with cnt_cnt), just cnt_cnt = 2.
          'Set qdf = .QueryDefs("zzz_qry_zFiduciary_31_05_02")
          ' ** zzz_qry_zFiduciary_31_04_01 (zzz_qry_zFiduciary_31_04 (zzz_qry_zFiduciary_31_02
          ' ** (zzz_qry_zFiduciary_31_01 (tblLedgerHidden, grouped, by accountno, ledghid_grpnum,
          ' ** with cnt_jno), just discrepancies), just ledghid_cnt discrepancies),
          ' ** grouped by accountno, with cnt_cnt), just cnt_cnt = 3.
          'Set qdf = .QueryDefs("zzz_qry_zFiduciary_31_06_02")
          ' ** zzz_qry_zFiduciary_31_04_01 (zzz_qry_zFiduciary_31_04 (zzz_qry_zFiduciary_31_02
          ' ** (zzz_qry_zFiduciary_31_01 (tblLedgerHidden, grouped, by accountno, ledghid_grpnum,
          ' ** with cnt_jno), just discrepancies), just ledghid_cnt discrepancies),
          ' ** grouped by accountno, with cnt_cnt), just cnt_cnt = 4.
          'Set qdf = .QueryDefs("zzz_qry_zFiduciary_31_07_02")
          ' ** zzz_qry_zFiduciary_31_04_01 (zzz_qry_zFiduciary_31_04 (zzz_qry_zFiduciary_31_02
          ' ** (zzz_qry_zFiduciary_31_01 (tblLedgerHidden, grouped, by accountno, ledghid_grpnum,
          ' ** with cnt_jno), just discrepancies), just ledghid_cnt discrepancies),
          ' ** grouped by accountno, with cnt_cnt), just cnt_cnt = 5.
          'Set qdf = .QueryDefs("zzz_qry_zFiduciary_31_08_02")
          ' ** zzz_qry_zFiduciary_31_04_01 (zzz_qry_zFiduciary_31_04 (zzz_qry_zFiduciary_31_02
          ' ** (zzz_qry_zFiduciary_31_01 (tblLedgerHidden, grouped, by accountno, ledghid_grpnum,
          ' ** with cnt_jno), just discrepancies), just ledghid_cnt discrepancies),
          ' ** grouped by accountno, with cnt_cnt), just cnt_cnt = 7.
          'Set qdf = .QueryDefs("zzz_qry_zFiduciary_31_10_02")
          ' ** zzz_qry_zFiduciary_31_04_01 (zzz_qry_zFiduciary_31_04 (zzz_qry_zFiduciary_31_02
          ' ** (zzz_qry_zFiduciary_31_01 (tblLedgerHidden, grouped, by accountno, ledghid_grpnum,
          ' ** with cnt_jno), just discrepancies), just ledghid_cnt discrepancies),
          ' ** grouped by accountno, with cnt_cnt), just cnt_cnt = 8.
          'Set qdf = .QueryDefs("zzz_qry_zFiduciary_31_11_02")
180       Set rst = qdf.OpenRecordset
190       With rst
200         .MoveLast
210         lngActNos = .RecordCount
220         .MoveFirst
230         arr_varActNo = .GetRows(lngActNos)
            ' **********************************************
            ' ** Array: arr_varActNo()
            ' **
            ' **   Field  Element  Name         Constant
            ' **   =====  =======  ===========  ==========
            ' **     1       0     accountno    A_ACTNO
            ' **     2       1     cnt_cnt      A_CNT
            ' **     3       2     qry_name1    A_QRY1
            ' **     4       3     qry_name2    A_QRY2
            ' **
            ' **********************************************
240         .Close
250       End With  ' ** rst.
260       Set rst = Nothing
270       Set qdf = Nothing

280       Debug.Print "'ACTNOS: " & CStr(lngActNos)
290       DoEvents

300       If lngActNos > 0& Then

            'strQryBase1 = "zzz_QQQ_zFiduciary_31_04_BE10120101_01"
            'strQryBase2 = "zzz_QQQ_zFiduciary_31_04_BE10120101_02"
310         strAccountNo1 = "BE10120101"

320         lngQrysCreated = 0&
330         For lngX = 0& To (lngActNos - 1&)
340           strAccountNo2 = arr_varActNo(A_ACTNO, lngX)
350           strQryName = strQryBase1
360           strQryName = StringReplace(strQryName, strAccountNo1, strAccountNo2)  ' ** Module Function: modStringFuncs.
370           DoCmd.CopyObject , strQryName, acQuery, strQryBase1
380           DoEvents
390           .QueryDefs.Refresh
400           arr_varActNo(A_QRY1, lngX) = strQryName
410           Set qdf = .QueryDefs(strQryName)
420           With qdf
430             strSQL = .SQL
440             strSQL = StringReplace(strSQL, strAccountNo1, strAccountNo2)  ' ** Module Function: modStringFuncs.
450             strSQL = StringReplace(strSQL, "yyy", "0")  ' ** Module Function: modStringFuncs.
460             strSQL = StringReplace(strSQL, "xxx", "0")  ' ** Module Function: modStringFuncs.
470             .SQL = strSQL
480             strDesc = .Properties("Description")
490             strDesc = StringReplace(strDesc, strAccountNo1, strAccountNo2)  ' ** Module Function: modStringFuncs.
500             intPos01 = InStr(strDesc, ";")
510             strDesc = Left(strDesc, (intPos01 + 1))
520             .Properties("Description") = strDesc
530           End With  ' ** qdf.
540           Set qdf = Nothing
550           lngQrysCreated = lngQrysCreated + 1&
560           strQryName = strQryBase2
570           strQryName = StringReplace(strQryName, strAccountNo1, strAccountNo2)  ' ** Module Function: modStringFuncs.
580           DoCmd.CopyObject , strQryName, acQuery, strQryBase2
590           DoEvents
600           .QueryDefs.Refresh
610           arr_varActNo(A_QRY2, lngX) = strQryName
620           Set qdf = .QueryDefs(strQryName)
630           With qdf
640             strSQL = .SQL
650             strSQL = StringReplace(strSQL, strAccountNo1, strAccountNo2)  ' ** Module Function: modStringFuncs.
660             .SQL = strSQL
670             strDesc = .Properties("Description")
680             strDesc = StringReplace(strDesc, strAccountNo1, strAccountNo2)  ' ** Module Function: modStringFuncs.
690             .Properties("Description") = strDesc
700           End With  ' ** qdf.
710           Set qdf = Nothing
720           .QueryDefs.Refresh
730           lngQrysCreated = lngQrysCreated + 1&
740         Next  ' ** lngX.

750         For lngX = 0& To (lngActNos - 1&)
760           Set qdf = .QueryDefs(arr_varActNo(A_QRY1, lngX))
770           With qdf
780             Set rst = .OpenRecordset
790             With rst
800               .MoveLast
810               lngRecs = .RecordCount
820               .Close
830             End With  ' ** rst.
840             strDesc = .Properties("Description")
850             strDesc = strDesc & CStr(lngRecs) & "."
860             .Properties("Description") = strDesc
870           End With  ' ** qdf.
880           Set qdf = Nothing
890         Next  ' ** lngX.

900       End If  ' ** lngActNos.

910       .Close
920     End With  ' ** dbs.
930     Set dbs = Nothing

940     Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
950     DoEvents

960     Beep

970     Debug.Print "'DONE!"
980     DoEvents

EXITP:
990     Set rst = Nothing
1000    Set qdf = Nothing
1010    Set dbs = Nothing
1020    Hide_FixQrys = blnRetVal
1030    Exit Function

ERRH:
1040    blnRetVal = False
1050    Select Case ERR.Number
        Case Else
1060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1070    End Select
1080    Resume EXITP

End Function

Public Function Hide_RetUniq(varInput As Variant, intMode As Integer) As Variant

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Hide_RetUniq"

        Dim blnFound As Boolean
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
        Dim varRetVal As Variant

1110    varRetVal = Null

1120    If IsNull(varInput) = False Then
1130      If Trim(varInput) <> vbNullString Then
            ' ** 000000000000001_0001_033231_033230_Withdrawn_Deposit__ .

1140        strTmp01 = Trim(varInput)

1150        Select Case intMode
            Case 1  ' ** Count the number of journalno's.

1160          intPos01 = CharPos(strTmp01, 2, "_")
1170          strTmp02 = Mid(strTmp01, (intPos01 + 1))
1180          intCnt = 0: blnFound = True
1190          Do While blnFound = True
1200            blnFound = False
1210            intPos01 = InStr(strTmp02, "_")
1220            strTmp03 = Left(strTmp02, (intPos01 - 1))
1230            If IsNumeric(strTmp03) = True Then
1240              If Val(strTmp03) > 0 Then
1250                blnFound = True
1260                intCnt = intCnt + 1
1270                strTmp02 = Mid(strTmp02, (intPos01 + 1))
1280                If Left(strTmp02, 1) = "_" Then
                      ' ** At least 2 in a row, so no more.
1290                  Exit Do
1300                End If
1310              End If
1320            End If
1330          Loop
1340          varRetVal = intCnt

1350        Case 2  ' ** Count the number of assetno's.

1360          intPos01 = CharPos(strTmp01, 1, "_")
1370          strTmp02 = Mid(strTmp01, (intPos01 + 1))
1380          intCnt = 0
1390          If Mid(strTmp02, 5, 1) = "_" Then
                ' ** The 1st one is an assetno.
1400            intCnt = intCnt + 1
1410            intPos01 = CharPos(strTmp02, 1, "_")
1420            strTmp02 = Mid(strTmp02, (intPos01 + 1))
1430            If Mid(strTmp02, 5, 1) = "_" Then
                  ' ** Oops! 2 assetno's!
1440              intCnt = intCnt + 1
1450            End If
1460          End If
1470          varRetVal = intCnt

1480        End Select

1490      End If
1500    End If

EXITP:
1510    Hide_RetUniq = varRetVal
1520    Exit Function

ERRH:
1530    varRetVal = RET_ERR
1540    Select Case ERR.Number
        Case Else
1550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1560    End Select
1570    Resume EXITP

End Function

Public Function Hide_FixQrys2() As Boolean

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Hide_FixQrys2"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, prp As Object, rst As DAO.Recordset, fld As DAO.Field
        Dim lngActNos As Long, arr_varActNo As Variant
        Dim lngFlds As Long, arr_varFld() As Variant
        Dim lngQrys As Long, arr_varQry() As Variant
        Dim strQryName1 As String, strQryName2 As String, strQryName3 As String, strSQL1 As String, strSQL2 As String, strSQL3 As String
        Dim strDesc1 As String, strDesc2 As String, strQryNum As String
        Dim lngQrysCreated As Long, lngRecs As Long
        Dim blnSkip As Boolean, blnFound As Boolean
        Dim intPos01 As Integer, intLen As Integer
        Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String
        Dim dblTmp05 As Double, dblTmp06 As Double, lngTmp07 As Long
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varActNo().
        Const A_ACTNO As Integer = 0
        'Const A_CNT   As Integer = 1
        Const A_QNAM1 As Integer = 2
        Const A_SQL1  As Integer = 3
        Const A_DSC1  As Integer = 4
        Const A_QNAM2 As Integer = 5
        'Const A_SQL2  As Integer = 6
        'Const A_DSC2  As Integer = 7
        'Const A_NUM   As Integer = 8

        ' ** Array: arr_varFld().
        Const F_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const F_FNAM As Integer = 0
        Const F_TNAM As Integer = 1
        Const F_CHK  As Integer = 2

        ' ** Array: arr_varQry().
        Const Q_ELEMS As Integer = 13  ' ** Array's first-element UBound().
        Const Q_QNAM As Integer = 0
        Const Q_SQL  As Integer = 1
        Const Q_DSC  As Integer = 2
        Const Q_NUM  As Integer = 3
        Const Q_SET  As Integer = 4
        Const Q_FLD1 As Integer = 5
        Const Q_FLD2 As Integer = 6
        Const Q_FLD3 As Integer = 7
        Const Q_FLD4 As Integer = 8
        Const Q_FLD5 As Integer = 9
        Const Q_FLD6 As Integer = 10
        Const Q_FLD7 As Integer = 11
        Const Q_FLD8 As Integer = 12
        Const Q_FLD9 As Integer = 13

        Const QRY_BASE As String = "zzz_qry_MasterTrust_24_05_"

1610  On Error GoTo 0

1620    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
1630    DoEvents

1640    Set dbs = CurrentDb
1650    With dbs

1660      blnSkip = True
1670      If blnSkip = False Then

1680        Set qdf = .QueryDefs("zzz_qry_MasterTrust_24_04")
1690        Set rst = qdf.OpenRecordset
1700        With rst
1710          .MoveLast
1720          lngActNos = .RecordCount
1730          .MoveFirst
1740          arr_varActNo = .GetRows(lngActNos)
1750          .Close
1760        End With
1770        Set rst = Nothing
1780        Set qdf = Nothing

1790        Debug.Print "'ACCTS: " & CStr(lngActNos)
1800        DoEvents

1810      End If  ' ** blnSkip.

1820      blnSkip = True
1830      If blnSkip = False Then

1840        strQryName1 = "zzz_qry_MasterTrust_24_05_010_01"
1850        strSQL1 = "SELECT tblLedgerHidden.ledghid_id, tblLedgerHidden.journalno, tblLedgerHidden.accountno, tblLedgerHidden.ledghid_cnt, " & _
              "tblLedgerHidden.ledghid_grpnum, tblLedgerHidden.ledghid_ord, tblLedgerHidden.ledghidtype_type, " & _
              "IIf([tblLedgerHidden].[journalno] In (0),1,Null) AS ledghidtype_type_new, " & _
              "zzz_qry_MasterTrust_11.sharefacex, zzz_qry_MasterTrust_24_03.shareface AS shareface_tot, zzz_qry_MasterTrust_11.icash, " & _
              "zzz_qry_MasterTrust_11.pcash, zzz_qry_MasterTrust_11.cost, zzz_qry_MasterTrust_24_03.Zx, zzz_qry_MasterTrust_11.src, " & _
              "tblLedgerHidden.ledghid_uniqueid, tblLedgerHidden.assetno, tblLedgerHidden.transdate, tblLedgerHidden.ledghid_username, " & _
              "tblLedgerHidden.ledghid_datemodified" & vbCrLf
1860        strSQL1 = strSQL1 & "FROM (tblLedgerHidden INNER JOIN zzz_qry_MasterTrust_11 ON " & _
              "tblLedgerHidden.journalno = zzz_qry_MasterTrust_11.journalno) LEFT JOIN zzz_qry_MasterTrust_24_03 ON " & _
              "(tblLedgerHidden.accountno = zzz_qry_MasterTrust_24_03.accountno) AND " & _
              "(tblLedgerHidden.ledghid_grpnum = zzz_qry_MasterTrust_24_03.ledghid_grpnum)" & vbCrLf
1870        strSQL1 = strSQL1 & "WHERE (((tblLedgerHidden.accountno)='00010'))" & vbCrLf
1880        strSQL1 = strSQL1 & "ORDER BY tblLedgerHidden.accountno, tblLedgerHidden.ledghid_grpnum, tblLedgerHidden.ledghid_ord;"
1890        strDesc1 = "    tblLedgerHidden, linked to .._11, .._24_03, just accountno = '010'; "

1900        For lngX = 0& To (lngActNos - 1&)
1910          strQryNum = arr_varActNo(A_ACTNO, lngX)
1920          strQryNum = Mid(strQryNum, 3)  ' ** Lop off the 1st 2 0's.
1930          strTmp01 = StringReplace(strQryName1, "_010_", "_" & strQryNum & "_")  ' ** Module Function: modStringFuncs.
1940          arr_varActNo(A_QNAM1, lngX) = strTmp01
1950          strTmp01 = StringReplace(strSQL1, "='00010'", "='" & arr_varActNo(A_ACTNO, lngX) & "'")  ' ** Module Function: modStringFuncs.
1960          arr_varActNo(A_SQL1, lngX) = strTmp01
1970          strTmp01 = StringReplace(strDesc1, "'010'", "'" & strQryNum & "'")  ' ** Module Function: modStringFuncs.
1980          arr_varActNo(A_DSC1, lngX) = strTmp01
1990        Next

2000        lngQrysCreated = 0&

2010        For lngX = 0& To (lngActNos - 1&)
2020          Set qdf = .CreateQueryDef(arr_varActNo(A_QNAM1, lngX), arr_varActNo(A_SQL1, lngX))
2030          With qdf
2040            Set prp = .CreateProperty("Description", dbText, arr_varActNo(A_DSC1, lngX))
2050  On Error Resume Next
2060            .Properties.Append prp
2070            If ERR.Number <> 0 Then
2080  On Error GoTo 0
2090              .Properties("Description") = arr_varActNo(A_DSC1, lngX)
2100            Else
2110  On Error GoTo 0
2120            End If
2130            Set prp = Nothing
2140            Set rst = .OpenRecordset
2150            rst.MoveLast
2160            strTmp01 = .Properties("Description")
2170            strTmp01 = strTmp01 & CStr(rst.RecordCount) & "."
2180            .Properties("Description") = strTmp01
2190          End With
2200          Set qdf = Nothing
2210          lngQrysCreated = lngQrysCreated + 1&
2220        Next

2230        Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
2240        DoEvents

2250      End If  ' ** blnSkip.

2260      blnSkip = True
2270      If blnSkip = False Then

2280        Debug.Print "'|";
2290        DoEvents

2300        lngQrysCreated = 0&
2310        For lngX = 0& To (lngActNos - 1&)
2320          strQryNum = arr_varActNo(A_ACTNO, lngX)
2330          strQryNum = Mid(strQryNum, 3)  ' ** Lop off the 1st 2 0's.
2340          If Val(strQryNum) >= 19 Then
2350            strQryName1 = "zzz_qry_MasterTrust_24_05_" & strQryNum & "_01"
2360            arr_varActNo(A_QNAM1, lngX) = strQryName1
2370            Set qdf = .QueryDefs(strQryName1)
2380            Set rst = qdf.OpenRecordset
2390            With rst
2400              .MoveLast
2410              lngRecs = .RecordCount
2420              .MoveFirst
2430              strTmp01 = vbNullString
2440              For lngY = 1& To lngRecs
2450                If ![ledghidtype_type] = 0& Then
2460                  strTmp01 = strTmp01 & CStr(![journalno]) & ","
2470                End If
2480                If lngY < lngRecs Then .MoveNext
2490              Next
2500              If Right(strTmp01, 1) = "," Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
2510              .Close
2520            End With
2530            If strTmp01 <> vbNullString Then
2540              strSQL1 = qdf.SQL
2550              intPos01 = InStr(strSQL1, "(0)")
2560              If intPos01 > 0 Then
2570                strSQL1 = StringReplace(strSQL1, "(0)", "(" & strTmp01 & ")")  ' ** Module Function: modStringFuncs.
2580                arr_varActNo(A_SQL1, lngX) = strSQL1
2590                qdf.SQL = strSQL1
2600                Qry_CheckBox strQryName1, "ledghidtype_type_new", True  ' ** Module Function: modQueryFunctions1.
2610                lngQrysCreated = lngQrysCreated + 1&
2620              Else
2630                Debug.Print "'QRY: " & strQryName1
2640                DoEvents
2650              End If
2660            End If
2670            Set rst = Nothing
2680            Set qdf = Nothing
2690          End If
2700          If (lngX + 1&) Mod 100 = 0 Then
2710            Debug.Print "|  " & CStr(lngX + 1&) & " of " & CStr(lngActNos)
2720            Debug.Print "'|";
2730          ElseIf (lngX + 1&) Mod 10 = 0 Then
2740            Debug.Print "|";
2750          Else
2760            Debug.Print ".";
2770          End If
2780          DoEvents
2790        Next
2800        Debug.Print
2810        DoEvents

2820        Debug.Print "'QRYS EDITED: " & CStr(lngQrysCreated)
2830        DoEvents

2840      End If  ' ** blnSkip.

2850      blnSkip = True
2860      If blnSkip = False Then

2870        strQryName1 = "zzz_qry_MasterTrust_24_05_010_01"

2880        lngQrysCreated = 0&
2890        For lngX = 0& To (lngActNos - 1&)
2900          strQryNum = arr_varActNo(A_ACTNO, lngX)
2910          strQryNum = Mid(strQryNum, 3)  ' ** Lop off the 1st 2 0's.
2920          strQryName3 = strQryName1
2930          strQryName3 = StringReplace(strQryName3, "_010_", "_" & strQryNum & "_")  ' ** Module Function: modStringFuncs.
2940          Set qdf = .QueryDefs(strQryName3)
2950          strTmp01 = qdf.SQL
2960          intPos01 = InStr(strTmp01, "ledghid_datemodified")
2970          strTmp02 = Left(strTmp01, (intPos01 + 19))
2980          strTmp03 = Mid(strTmp01, (intPos01 + 20))
2990          strTmp02 = strTmp02 & ", zzz_qry_MasterTrust_11.journaltype"
3000          strTmp02 = strTmp02 & strTmp03
3010          qdf.SQL = strTmp02
3020          lngQrysCreated = lngQrysCreated + 1&
3030        Next

3040        Debug.Print "'QRYS EDITED: " & CStr(lngQrysCreated)
3050        DoEvents

3060      End If  ' ** blnSkip.

3070      blnSkip = True
3080      If blnSkip = False Then

3090        strQryName1 = "zzz_qry_MasterTrust_24_05_010_02"
3100        strSQL1 = "SELECT ledger.journalno, ledger.journaltype, IIf([journalno]=[jno2],'Purchase',Null) AS journaltype_new, " & _
              "ledger.accountno, ledger.assetno, ledger.transdate, ledger.assetdate, ledger.PurchaseDate, " & _
              "IIf([journaltype] In ('Deposit','Purchase'),[shareface],IIf([journaltype] In ('Withdrawn','Sold'),-[shareface]," & _
              "IIf([journaltype]='Liability',IIf(IsNull([PurchaseDate])=True,[shareface],-[shareface]),0))) AS sharefacex, " & _
              "ledger.shareface, ledger.icash, IIf([journalno]=[jno1],[icsh1],IIf([journalno]=[jno2],[icsh2],Null)) AS icash_new, " & _
              "ledger.pcash, ledger.cost, ledger.ledger_HIDDEN, IIf([journalno]=[jno2],False," & _
              "IIf([journalno]=[jno3],True,[ledger_HIDDEN])) AS ledger_HIDDEN_new, ledger.description" & vbCrLf
3110        strSQL1 = strSQL1 & "FROM ledger" & vbCrLf
3120        strSQL1 = strSQL1 & "WHERE (((ledger.journalno) In ([jno4])) AND ((ledger.accountno)='00010') AND ((ledger.assetno)=1));"
3130        strDesc1 = "        Ledger, just journalno = [jno4], with icash_new, ledger_HIDDEN_new; "

3140        strQryName2 = "zzz_qry_MasterTrust_24_05_010_03"
3150        strSQL2 = "SELECT tblLedgerHidden.ledghid_id, tblLedgerHidden.journalno, IIf([journalno]=[jno1],[jno2],Null) AS journalno_new, " & _
              "tblLedgerHidden.accountno, tblLedgerHidden.assetno, tblLedgerHidden.transdate, tblLedgerHidden.ledghid_cnt, " & _
              "tblLedgerHidden.ledghid_grpnum, tblLedgerHidden.ledghid_ord, tblLedgerHidden.ledghidtype_type, " & _
              "IIf([tblLedgerHidden].[journalno] In ([jno3]),1,Null) AS ledghidtype_type_new, tblLedgerHidden.ledghid_uniqueid, " & _
              "IIf([tblLedgerHidden].[journalno] In ([jno4]),Left([ledghid_uniqueid],CharPos([ledghid_uniqueid],2,'_')) & " & _
              "'jno5_jno6_jtyp1_jtyp2',Null) AS ledghid_uniqueid_new, tblLedgerHidden.ledghid_username, " & _
              "tblLedgerHidden.ledghid_datemodified" & vbCrLf
3160        strSQL2 = strSQL2 & "FROM tblLedgerHidden" & vbCrLf
3170        strSQL2 = strSQL2 & "WHERE (((tblLedgerHidden.journalno) In ([jno7])))" & vbCrLf
3180        strSQL2 = strSQL2 & "ORDER BY tblLedgerHidden.journalno;"
3190        strDesc2 = "        tblLedgerHidden, just [cnt] journalno's, with journalno_new, ledghid_uniqueid_new, ledghidtype_type_new; "

3200        For lngX = 0& To (lngActNos - 1&)
3210          strQryNum = arr_varActNo(A_ACTNO, lngX)
3220          strQryNum = Mid(strQryNum, 3)  ' ** Lop off the 1st 2 0's.
3230          If Val(strQryNum) >= 13 Then
3240            strTmp01 = strQryName1
3250            strTmp01 = StringReplace(strTmp01, "_010_", "_" & strQryNum & "_")  ' ** Module Function: modStringFuncs.
3260            arr_varActNo(A_QNAM1, lngX) = strTmp01
3270            strTmp01 = strQryName2
3280            strTmp01 = StringReplace(strTmp01, "_010_", "_" & strQryNum & "_")  ' ** Module Function: modStringFuncs.
3290            arr_varActNo(A_QNAM2, lngX) = strTmp01
3300            strTmp01 = strSQL1
3310            strQryName3 = arr_varActNo(A_QNAM1, lngX)
3320            strQryName3 = Left(strQryName3, (Len(strQryName3) - 1)) & "1"
3330            Set qdf = .QueryDefs(strQryName3)
3340            Set rst = qdf.OpenRecordset
3350            With rst
3360              .MoveLast
3370              lngRecs = .RecordCount
3380              .MoveFirst
3390              strTmp02 = vbNullString: strTmp03 = vbNullString: dblTmp05 = 0#: dblTmp06 = 0#
3400              For lngY = 1& To lngRecs
3410                If ![Zx] = "X" Then
3420                  strTmp02 = strTmp02 & CStr(![journalno]) & ","
3430                  If ![sharefacex] > 0 Then
3440                    dblTmp05 = ![sharefacex]
3450                  Else
3460                    dblTmp06 = ![sharefacex]
3470                  End If
3480                  If ![journaltype] = "Sold" Then
3490                    strTmp03 = CStr(![journalno])
3500                  ElseIf ![journaltype] = "Deposit" Then
3510                    strTmp04 = CStr(![journalno])
3520                  End If
3530                End If
3540                If lngY < lngRecs Then .MoveNext
3550              Next
3560              varTmp00 = DLookup("[journalno]", "zzz_qry_MasterTrust_11", "[accountno] = '" & arr_varActNo(A_ACTNO, lngX) & "' And " & _
                    "[journaltype] = 'Purchase' And [shareface] = " & CStr(Abs(dblTmp06)))
3570              If IsNull(varTmp00) = False Then
3580                strTmp02 = strTmp02 & CStr(varTmp00) & ","
3590              Else
3600                Stop
3610              End If
3620              .Close
3630            End With
3640            Set rst = Nothing
3650            If Right(strTmp02, 1) = "," Then strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
3660            Set qdf = Nothing
3670            strTmp01 = StringReplace(strTmp01, "='00010'", "='" & arr_varActNo(A_ACTNO, lngX) & "'")  ' ** Module Function: modStringFuncs.
                '[jno1] = strTmp03  '1st icash
3680            strTmp01 = StringReplace(strTmp01, "[jno1]", strTmp03)  ' ** Module Function: modStringFuncs.
                '[jno2] = strTmp04  '"Purchase" and 2nd icash and False ledger_HIDDEN
3690            strTmp01 = StringReplace(strTmp01, "[jno2]", strTmp04)  ' ** Module Function: modStringFuncs.
                '[jno3] = varTmp00  'True ledger_HIDDEN
3700            strTmp01 = StringReplace(strTmp01, "[jno3]", CStr(varTmp00))  ' ** Module Function: modStringFuncs.
                '[jno4] = strTmp02  'Criteria and Description
3710            strTmp01 = StringReplace(strTmp01, "[jno4]", strTmp02)  ' ** Module Function: modStringFuncs.
                '[icsh1] = dblTmp05
3720            strTmp01 = StringReplace(strTmp01, "[icsh1]", CStr(dblTmp05))  ' ** Module Function: modStringFuncs.
                '[icsh2] = -dblTmp05
3730            strTmp01 = StringReplace(strTmp01, "[icsh2]", CStr(-dblTmp05))  ' ** Module Function: modStringFuncs.
3740            arr_varActNo(A_SQL1, lngX) = strTmp01
3750            strTmp01 = strDesc1
3760            strTmp01 = StringReplace(strTmp01, "[jno4]", strTmp02)  ' ** Module Function: modStringFuncs.
3770            arr_varActNo(A_DSC1, lngX) = strTmp01
3780          End If
3790        Next

3800        Debug.Print "'|";
3810        DoEvents

3820        lngQrysCreated = 0&
3830        For lngX = 0& To (lngActNos - 1&)
3840          If IsNull(arr_varActNo(A_SQL1, lngX)) = False Then
3850            Set qdf = .CreateQueryDef(arr_varActNo(A_QNAM1, lngX), arr_varActNo(A_SQL1, lngX))
3860            With qdf
3870              Set prp = .CreateProperty("Description", dbText, arr_varActNo(A_DSC1, lngX))
3880  On Error Resume Next
3890              .Properties.Append prp
3900              If ERR.Number <> 0 Then
3910  On Error GoTo 0
3920                .Properties("Description") = arr_varActNo(A_DSC1, lngX)
3930              Else
3940  On Error GoTo 0
3950              End If
3960              Qry_CheckBox .Name, "ledger_HIDDEN_new", True  ' ** Module Function: modQueryFunctions1.
3970            End With
3980            Set qdf = Nothing
3990            lngQrysCreated = lngQrysCreated + 1&
4000          End If
4010          If (lngX + 1&) Mod 100 = 0 Then
4020            Debug.Print "|  " & CStr(lngX + 1&) & " of " & CStr(lngActNos)
4030            Debug.Print "'|";
4040          ElseIf (lngX + 1&) Mod 10 = 0 Then
4050            Debug.Print "|";
4060          Else
4070            Debug.Print ".";
4080          End If
4090          DoEvents
4100        Next
4110        Debug.Print
4120        DoEvents

4130        Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
4140        DoEvents

4150      End If  ' ** blnSkip.

4160      blnSkip = True
4170      If blnSkip = False Then

4180        strQryName1 = "zzz_qry_MasterTrust_24_05_010_03"

4190        Debug.Print "'|";
4200        DoEvents

4210        lngQrysCreated = 0&
4220        For lngX = 0& To (lngActNos - 1&)
4230          strQryNum = arr_varActNo(A_ACTNO, lngX)
4240          strQryNum = Mid(strQryNum, 3)  ' ** Lop off the 1st 2 0's.
4250          strQryName3 = strQryName1
4260          strQryName3 = StringReplace(strQryName3, "_010_", "_" & strQryNum & "_")  ' ** Module Function: modStringFuncs.
              'DoCmd.CopyObject , strQryName3, acQuery, strQryName1
4270          Set qdf = .QueryDefs(strQryName3)
4280          With qdf
4290            strTmp01 = .Properties("Description")
4300            intPos01 = InStr(strTmp01, ";")
4310            strTmp01 = Left(strTmp01, (intPos01 + 1))
4320            strTmp01 = StringReplace(strTmp01, "10", "00")  ' ** Module Function: modStringFuncs.
4330            .Properties("Description") = strTmp01
4340          End With
4350          Set qdf = Nothing
4360          lngQrysCreated = lngQrysCreated + 1&
4370          If (lngX + 1&) Mod 100 = 0 Then
4380            Debug.Print "|  " & CStr(lngX + 1&) & " of " & CStr(lngActNos)
4390            Debug.Print "'|";
4400          ElseIf (lngX + 1&) Mod 10 = 0 Then
4410            Debug.Print "|";
4420          Else
4430            Debug.Print ".";
4440          End If
4450          DoEvents
4460        Next
4470        Debug.Print
4480        DoEvents

4490        Debug.Print "'QRYS EDITED: " & CStr(lngQrysCreated)
4500        DoEvents

            'Debug.Print "'QRYS COPIED: " & CStr(lngQrysCreated)
            'DoEvents

4510      End If  ' ** blnSkip.

4520      blnSkip = True
4530      If blnSkip = False Then

4540        strQryName1 = "zzz_qry_MasterTrust_24_05_032_01"
4550        strQryName2 = "zzz_qry_MasterTrust_24_05_032_02"
4560        strQryName3 = "zzz_qry_MasterTrust_24_05_032_03"

4570        strSQL3 = "SELECT tblLedgerHidden.ledghid_id, tblLedgerHidden.journalno, IIf([journalno]=[jno1],[jno2],Null) AS journalno_new, " & _
              "tblLedgerHidden.accountno, tblLedgerHidden.assetno, tblLedgerHidden.transdate, tblLedgerHidden.ledghid_cnt, " & _
              "tblLedgerHidden.ledghid_grpnum, tblLedgerHidden.ledghid_ord, tblLedgerHidden.ledghidtype_type, " & _
              "IIf([tblLedgerHidden].[journalno] In (0),1,Null) AS ledghidtype_type_new, " & _
              "tblLedgerHidden.ledghid_uniqueid, IIf([tblLedgerHidden].[journalno] In ([jno3])," & _
              "Left([ledghid_uniqueid],CharPos([ledghid_uniqueid],2,'_')) & '[jno4]_[jno5]_Purchase__Sold_____',Null) AS " & _
              "ledghid_uniqueid_new, tblLedgerHidden.ledghid_username, tblLedgerHidden.ledghid_datemodified" & vbCrLf
4580        strSQL3 = strSQL3 & "FROM tblLedgerHidden" & vbCrLf
4590        strSQL3 = strSQL3 & "WHERE (((tblLedgerHidden.journalno) In ([jno3])))" & vbCrLf
4600        strSQL3 = strSQL3 & "ORDER BY tblLedgerHidden.journalno;"
            ' ** [jno1] = Existing journalno that's being reversed.
            ' ** [jno2] = New journalno that's being reversed.
            ' ** [jno3] = strTmp02, journalno's without spaces.
            ' ** [jno4] = 6-char new journalno that's being reversed.
            ' ** [jno5] = 6-char original reversing journalno.

4610        Debug.Print "'|";
4620        DoEvents

4630        lngQrysCreated = 0&
4640        For lngX = 0& To (lngActNos - 1&)
4650          strQryNum = arr_varActNo(A_ACTNO, lngX)
4660          strQryNum = Mid(strQryNum, 3)  ' ** Lop off the 1st 2 0's.
4670          If Val(strQryNum) >= 210 And Val(strQryNum) <> 224 Then
4680            strTmp01 = strQryName1
4690            strTmp01 = StringReplace(strTmp01, "_032_", "_" & strQryNum & "_")  ' ** Module Function: modStringFuncs.
4700            Set qdf = .QueryDefs(strTmp01)
4710            strTmp02 = qdf.Properties("Description")
4720            intPos01 = InStr(strTmp02, ";")
4730            strTmp02 = Mid(strTmp02, (intPos01 + 1))
4740            If Right(strTmp02, 1) = "." Then strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
4750            If Val(strTmp02) = 2 Then
4760              Set qdf = Nothing
4770              strTmp01 = strQryName2
4780              strTmp01 = StringReplace(strTmp01, "_032_", "_" & strQryNum & "_")  ' ** Module Function: modStringFuncs.
4790              Set qdf = .QueryDefs(strTmp01)
4800              strTmp01 = qdf.Properties("Description")
4810              strTmp02 = strTmp01
                  '        Ledger, just journalno = 64112,63789,61901, with icash_new, ledger_HIDDEN_new;
4820              intPos01 = InStr(strTmp02, "=")
4830              strTmp02 = Trim(Mid(strTmp02, (intPos01 + 1)))
4840              intPos01 = InStr(strTmp02, " ")
4850              strTmp02 = Trim(Left(strTmp02, intPos01))
4860              strTmp03 = strTmp02
4870              intPos01 = InStr(strTmp03, ",")
4880              strTmp03 = Left(strTmp03, intPos01) & " " & Mid(strTmp03, (intPos01 + 1))
4890              intPos01 = InStr((intPos01 + 1), strTmp03, ",")
4900              strTmp03 = Left(strTmp03, intPos01) & " " & Mid(strTmp03, (intPos01 + 1))
4910              strTmp01 = StringReplace(strTmp01, strTmp02, strTmp03)  ' ** Module Function: modStringFuncs.
4920              strTmp01 = strTmp01 & "3."
4930              qdf.Properties("Description") = strTmp01
4940              Set qdf = Nothing
4950              strTmp04 = strTmp02
4960              strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
                  ' ** strTmp04 has journalno's without spaces.
4970              strTmp01 = strQryName2
4980              strTmp01 = StringReplace(strTmp01, "_032_", "_" & strQryNum & "_")  ' ** Module Function: modStringFuncs.
4990              Set qdf = .QueryDefs(strTmp01)
5000              Set rst = qdf.OpenRecordset
5010              With rst
5020                .MoveLast
5030                lngRecs = .RecordCount  ' ** Should be 3.
5040                .MoveFirst
5050                For lngY = 1& To lngRecs
5060                  If ![journaltype] = "Purchase" Then
5070                    strTmp03 = ![journalno]  ' ** New journalno that's being reversed.
5080                  ElseIf ![journaltype] = "Sold" Then
5090                    strTmp02 = ![journalno]  ' ** Original reversing journalno.
5100                  ElseIf ![journaltype] = "Deposit" And IsNull(![journaltype_new]) = False Then
5110                    strTmp01 = ![journalno]  ' ** Existing journalno that's being reversed.
5120                  Else
5130                    Stop
5140                  End If
5150                  If lngY < lngRecs Then .MoveNext
5160                Next
5170                .Close
5180              End With
5190              Set rst = Nothing
5200              Set qdf = Nothing
5210              If strTmp01 = vbNullString Or strTmp02 = vbNullString Or strTmp03 = vbNullString Then
5220                Stop
5230              End If
                  ' ** [jno1] = strTmp01, Existing journalno that's being reversed.
                  ' ** [jno2] = strTmp03, New journalno that's being reversed.
                  ' ** [jno3] = strTmp04, journalno's without spaces.
                  ' ** [jno4] = 6-char strTmp03, new journalno that's being reversed.
                  ' ** [jno5] = 6-char strTmp02, original reversing journalno.
5240              strQryName1 = strQryName3
5250              strQryName1 = StringReplace(strQryName1, "_032_", "_" & strQryNum & "_")  ' ** Module Function: modStringFuncs.
5260              Set qdf = .QueryDefs(strQryName1)
5270              With qdf
5280                strSQL1 = strSQL3
5290                strSQL1 = StringReplace(strSQL1, "[jno1]", strTmp01)  ' ** Module Function: modStringFuncs.
5300                strSQL1 = StringReplace(strSQL1, "[jno2]", strTmp03)  ' ** Module Function: modStringFuncs.
5310                strSQL1 = StringReplace(strSQL1, "[jno3]", strTmp04)  ' ** Module Function: modStringFuncs.
5320                strSQL1 = StringReplace(strSQL1, "[jno4]", Right(String(6, "0") & strTmp03, 6))  ' ** Module Function: modStringFuncs.
5330                strSQL1 = StringReplace(strSQL1, "[jno5]", Right(String(6, "0") & strTmp02, 6))  ' ** Module Function: modStringFuncs.
5340                .SQL = strSQL1
                    '        tblLedgerHidden, just 00 journalno's, with journalno_new, ledghid_uniqueid_new, ledghidtype_type_new;
5350                strTmp01 = .Properties("Description")
5360                strTmp01 = StringReplace(strTmp01, "00", "2")  ' ** Module Function: modStringFuncs.
5370                strTmp01 = strTmp01 & "2."
5380                .Properties("Description") = strTmp01
5390              End With
5400              Set qdf = Nothing
5410              lngQrysCreated = lngQrysCreated + 1&
5420            Else
5430              Set qdf = Nothing
5440            End If
5450          End If
5460          If (lngX + 1&) Mod 100 = 0 Then
5470            Debug.Print "|  " & CStr(lngX + 1&) & " of " & CStr(lngActNos)
5480            Debug.Print "'|";
5490          ElseIf (lngX + 1&) Mod 10 = 0 Then
5500            Debug.Print "|";
5510          Else
5520            Debug.Print ".";
5530          End If
5540          DoEvents
5550        Next
5560        Debug.Print
5570        DoEvents

5580        Debug.Print "'QRYS EDITED: " & CStr(lngQrysCreated)
5590        DoEvents

5600      End If  ' ** blnSkip.

5610      blnSkip = True
5620      If blnSkip = False Then

5630        lngFlds = 0&
5640        ReDim arr_varFld(F_ELEMS, 0)

5650        intLen = Len(QRY_BASE)

5660        For Each qdf In .QueryDefs
5670          With qdf
5680            If Left(.Name, intLen) = QRY_BASE Then
5690              If .Type = dbQSelect Then
5700                If Right(.Name, 2) <> "01" Then
5710                  strQryNum = Mid(.Name, (intLen + 1), 3)
5720                  If Val(strQryNum) > 10 Then
5730                    For Each fld In .Fields
5740                      With fld
5750                        If Right(.Name, 4) = "_new" Then
5760                          blnFound = False
5770                          strTmp02 = .SourceTable
5780                          If strTmp02 = vbNullString Then
5790                            strTmp02 = qdf.Fields(0).SourceTable
5800                          End If
5810                          For lngX = 0& To (lngFlds - 1&)
5820                            If arr_varFld(F_FNAM, lngX) = .Name And arr_varFld(F_TNAM, lngX) = strTmp02 Then
5830                              blnFound = True
5840                              Exit For
5850                            End If
5860                          Next
5870                          If blnFound = False Then
5880                            lngFlds = lngFlds + 1&
5890                            lngE = lngFlds - 1&
5900                            ReDim Preserve arr_varFld(F_ELEMS, lngE)
5910                            arr_varFld(F_FNAM, lngE) = .Name
5920                            arr_varFld(F_TNAM, lngE) = strTmp02
5930                            arr_varFld(F_CHK, lngE) = CBool(False)
5940                          End If
5950                        End If
5960                      End With
5970                    Next
5980                  End If
5990                End If
6000              End If
6010            End If
6020          End With
6030        Next
6040        Set fld = Nothing
6050        Set qdf = Nothing

6060        Debug.Print "'NEW FLDS: " & CStr(lngFlds)
6070        DoEvents

6080        Debug.Print "'Ledger:"
6090        DoEvents
6100        lngY = 0&
6110        For lngX = 0& To (lngFlds - 1&)
6120          If arr_varFld(F_TNAM, lngX) = "ledger" Then
6130            lngY = lngY + 1&
6140            Debug.Print "'  " & Left(CStr(lngY) & "." & Space(4), 4) & arr_varFld(F_FNAM, lngX)
6150            DoEvents
6160            arr_varFld(F_CHK, lngX) = CBool(True)
6170          End If
6180        Next
6190        Debug.Print "'tblLedgerHidden:"
6200        DoEvents
6210        lngY = 0&
6220        For lngX = 0& To (lngFlds - 1&)
6230          If arr_varFld(F_TNAM, lngX) = "tblLedgerHidden" Then
6240            lngY = lngY + 1&
6250            Debug.Print "'  " & Left(CStr(lngY) & "." & Space(4), 4) & arr_varFld(F_FNAM, lngX)
6260            DoEvents
6270            arr_varFld(F_CHK, lngX) = CBool(True)
6280          End If
6290        Next

6300      End If  ' ** blnSkip.

6310      blnSkip = True
6320      If blnSkip = False Then

6330        lngQrys = 0&
6340        ReDim arr_varQry(Q_ELEMS, 0)

6350        intLen = Len(QRY_BASE)

6360        For Each qdf In .QueryDefs
6370          With qdf
6380            If Left(.Name, intLen) = QRY_BASE Then
6390              If .Type = dbQSelect Then
6400                strQryNum = Mid(.Name, (intLen + 1), 3)
6410                If Val(strQryNum) > 14 Then
6420                  If Right(.Name, 2) <> "01" Then
6430                    blnFound = False
6440                    For Each fld In .Fields
6450                      With fld
6460                        If Right(.Name, 4) = "_new" Then
6470                          blnFound = True
6480                          Exit For
6490                        End If
6500                      End With
6510                    Next
6520                    Set fld = Nothing
6530                    If blnFound = True Then
6540                      lngQrys = lngQrys + 1&
6550                      lngE = lngQrys - 1&
6560                      ReDim Preserve arr_varQry(Q_ELEMS, lngE)
6570                      arr_varQry(Q_QNAM, lngE) = .Name
6580                      arr_varQry(Q_SQL, lngE) = Null
6590                      arr_varQry(Q_DSC, lngE) = .Properties("Description")
6600                      arr_varQry(Q_NUM, lngE) = strQryNum
6610                      arr_varQry(Q_SET, lngE) = CLng(1)
6620                      arr_varQry(Q_FLD1, lngE) = CBool(False)
6630                      arr_varQry(Q_FLD2, lngE) = CBool(False)
6640                      arr_varQry(Q_FLD3, lngE) = CBool(False)
6650                      arr_varQry(Q_FLD4, lngE) = CBool(False)
6660                      arr_varQry(Q_FLD5, lngE) = CBool(False)
6670                      arr_varQry(Q_FLD6, lngE) = CBool(False)
6680                      arr_varQry(Q_FLD7, lngE) = CBool(False)
6690                      arr_varQry(Q_FLD8, lngE) = CBool(False)
6700                      arr_varQry(Q_FLD9, lngE) = CBool(False)
6710                      For Each fld In .Fields
6720                        With fld
6730                          Select Case .Name
                              Case "icash_new"
6740                            arr_varQry(Q_FLD1, lngE) = CBool(True)
6750                          Case "ledger_HIDDEN_new"
6760                            arr_varQry(Q_FLD2, lngE) = CBool(True)
6770                          Case "journaltype_new"
6780                            arr_varQry(Q_FLD3, lngE) = CBool(True)
6790                          Case "journalno_new"
6800                            arr_varQry(Q_FLD4, lngE) = CBool(True)
6810                          Case "ledghid_uniqueid_new"
6820                            arr_varQry(Q_FLD5, lngE) = CBool(True)
6830                          Case "ledghidtype_type_new"
6840                            arr_varQry(Q_FLD6, lngE) = CBool(True)
6850                          Case "ledghid_cnt_new"
6860                            arr_varQry(Q_FLD7, lngE) = CBool(True)
6870                          Case "ledghid_grpnum_new"
6880                            arr_varQry(Q_FLD8, lngE) = CBool(True)
6890                          Case "ledghid_ord_new"
6900                            arr_varQry(Q_FLD9, lngE) = CBool(True)
6910                          End Select
6920                        End With
6930                      Next
6940                    End If  ' ** blnFound.
6950                  End If
6960                End If
6970              End If
6980            End If
6990          End With
7000        Next
7010        Set qdf = Nothing

7020        Debug.Print "'QRYS: " & CStr(lngQrys)
7030        DoEvents

7040        For lngX = 0& To (lngQrys - 1&)
7050          lngTmp07 = Val(Right(arr_varQry(Q_QNAM, lngX), 2))
7060          Select Case lngTmp07  ' ** There are outliers, but I'll deal with them later.
              Case 2&, 3&
                ' ** 1st set of queries.
7070          Case 4&, 5&
7080            arr_varQry(Q_SET, lngX) = 2&
7090          Case 6&, 7&
7100            arr_varQry(Q_SET, lngX) = 3&
7110          Case 8&, 9&
7120            arr_varQry(Q_SET, lngX) = 4&
7130          End Select
7140        Next

            'For lngX = 0& To (lngQrys - 1&)
            '  If arr_varQry(Q_SET, lngX) = 5& Then
            '    Debug.Print "'" & arr_varQry(Q_QNAM, lngX)
            '  End If
            'Next

            'lngQrysCreated = 0&
            'For lngX = 0& To (lngQrys - 1&)
            '  If arr_varQry(Q_FLD1, lngX) = True Then
            '    lngQrysCreated = lngQrysCreated + 1&
            '  End If
            '  If arr_varQry(Q_FLD2, lngX) = True Then
            '    lngQrysCreated = lngQrysCreated + 1&
            '  End If
            '  If arr_varQry(Q_FLD3, lngX) = True Then
            '    lngQrysCreated = lngQrysCreated + 1&
            '  End If
            '  If arr_varQry(Q_FLD4, lngX) = True Then
            '    lngQrysCreated = lngQrysCreated + 1&
            '  End If
            '  If arr_varQry(Q_FLD5, lngX) = True Then
            '    lngQrysCreated = lngQrysCreated + 1&
            '  End If
            '  If arr_varQry(Q_FLD6, lngX) = True Then
            '    lngQrysCreated = lngQrysCreated + 1&
            '  End If
            '  If arr_varQry(Q_FLD7, lngX) = True Then
            '    lngQrysCreated = lngQrysCreated + 1&
            '  End If
            '  If arr_varQry(Q_FLD8, lngX) = True Then
            '    lngQrysCreated = lngQrysCreated + 1&
            '  End If
            '  If arr_varQry(Q_FLD9, lngX) = True Then
            '    lngQrysCreated = lngQrysCreated + 1&
            '  End If
            'Next

            'Debug.Print "'UPDATE QRYS: " & CStr(lngQrysCreated)
            'DoEvents

            'QRYS: 544
            'UPDATE QRYS: 1640
            'DONE!

7150        Debug.Print "'|";
7160        DoEvents

7170        lngQrysCreated = 0&
7180        For lngX = 0& To (lngQrys - 1&)

7190          strQryName1 = arr_varQry(Q_QNAM, lngX)

7200          If arr_varQry(Q_FLD1, lngX) = True Then
                ' ** Update zzz_qry_MasterTrust_24_05_001_02, for icash.  Ledger
7210            strQryName2 = "zzz_qry_MasterTrust_24_05_001_21"
7220            strQryName3 = StringReplace(strQryName2, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")  ' ** Module Function: modStringFuncs.
7230            Select Case arr_varQry(Q_SET, lngX)
                Case 1&
                  ' ** Copy as-is.
7240            Case 2&
7250              strTmp01 = Right(strQryName3, 2)
7260              strTmp01 = CStr(Val(strTmp01) + 10&)
7270              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
7280            Case 3&
7290              strTmp01 = Right(strQryName3, 2)
7300              strTmp01 = CStr(Val(strTmp01) + 20&)
7310              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
7320            Case 4&
7330              strTmp01 = Right(strQryName3, 2)
7340              strTmp01 = CStr(Val(strTmp01) + 30&)
7350              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
7360            End Select
7370            DoCmd.CopyObject , strQryName3, acQuery, strQryName2
7380            DoEvents
7390            .QueryDefs.Refresh
7400            Set qdf = .QueryDefs(strQryName3)
7410            With qdf
7420              strSQL1 = .SQL
7430              strSQL1 = StringReplace(strSQL1, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")
7440              If arr_varQry(Q_SET, lngX) = 1& Then
                    ' ** Should be fine as-is.
7450              Else
7460                strTmp01 = "_" & arr_varQry(Q_NUM, lngX) & "_02"
7470                strTmp02 = Right(strQryName1, 7)  '"_001_02"
7480                strSQL1 = StringReplace(strSQL1, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
7490              End If
7500              .SQL = strSQL1
7510              strTmp01 = .Properties("Description")
7520              strTmp01 = StringReplace(strTmp01, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")
7530              .Properties("Description") = strTmp01
7540            End With
7550            Set qdf = Nothing
7560            lngQrysCreated = lngQrysCreated + 1&
7570          End If

7580          If arr_varQry(Q_FLD2, lngX) = True Then
                ' ** Update zzz_qry_MasterTrust_24_05_001_02, for ledger_HIDDEN.  Ledger
7590            strQryName2 = "zzz_qry_MasterTrust_24_05_001_22"
7600            strQryName3 = StringReplace(strQryName2, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")  ' ** Module Function: modStringFuncs.
7610            Select Case arr_varQry(Q_SET, lngX)
                Case 1&
                  ' ** Copy as-is.
7620            Case 2&
7630              strTmp01 = Right(strQryName3, 2)
7640              strTmp01 = CStr(Val(strTmp01) + 10&)
7650              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
7660            Case 3&
7670              strTmp01 = Right(strQryName3, 2)
7680              strTmp01 = CStr(Val(strTmp01) + 20&)
7690              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
7700            Case 4&
7710              strTmp01 = Right(strQryName3, 2)
7720              strTmp01 = CStr(Val(strTmp01) + 30&)
7730              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
7740            End Select
7750            DoCmd.CopyObject , strQryName3, acQuery, strQryName2
7760            DoEvents
7770            .QueryDefs.Refresh
7780            Set qdf = .QueryDefs(strQryName3)
7790            With qdf
7800              strSQL1 = .SQL
7810              strSQL1 = StringReplace(strSQL1, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")
7820              If arr_varQry(Q_SET, lngX) = 1& Then
                    ' ** Should be fine as-is.
7830              Else
7840                strTmp01 = "_" & arr_varQry(Q_NUM, lngX) & "_02"
7850                strTmp02 = Right(strQryName1, 7)  '"_001_02"
7860                strSQL1 = StringReplace(strSQL1, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
7870              End If
7880              .SQL = strSQL1
7890              strTmp01 = .Properties("Description")
7900              strTmp01 = StringReplace(strTmp01, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")
7910              .Properties("Description") = strTmp01
7920            End With
7930            Set qdf = Nothing
7940            lngQrysCreated = lngQrysCreated + 1&
7950          End If

7960          If arr_varQry(Q_FLD3, lngX) = True Then
                ' ** Update zzz_qry_MasterTrust_24_05_001_02, for journaltype.  Ledger
7970            strQryName2 = "zzz_qry_MasterTrust_24_05_001_23"
7980            strQryName3 = StringReplace(strQryName2, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")  ' ** Module Function: modStringFuncs.
7990            Select Case arr_varQry(Q_SET, lngX)
                Case 1&
                  ' ** Copy as-is.
8000            Case 2&
8010              strTmp01 = Right(strQryName3, 2)
8020              strTmp01 = CStr(Val(strTmp01) + 10&)
8030              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
8040            Case 3&
8050              strTmp01 = Right(strQryName3, 2)
8060              strTmp01 = CStr(Val(strTmp01) + 20&)
8070              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
8080            Case 4&
8090              strTmp01 = Right(strQryName3, 2)
8100              strTmp01 = CStr(Val(strTmp01) + 30&)
8110              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
8120            End Select
8130            DoCmd.CopyObject , strQryName3, acQuery, strQryName2
8140            DoEvents
8150            .QueryDefs.Refresh
8160            Set qdf = .QueryDefs(strQryName3)
8170            With qdf
8180              strSQL1 = .SQL
8190              strSQL1 = StringReplace(strSQL1, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")
8200              If arr_varQry(Q_SET, lngX) = 1& Then
                    ' ** Should be fine as-is.
8210              Else
8220                strTmp01 = "_" & arr_varQry(Q_NUM, lngX) & "_02"
8230                strTmp02 = Right(strQryName1, 7)  '"_001_02"
8240                strSQL1 = StringReplace(strSQL1, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
8250              End If
8260              .SQL = strSQL1
8270              strTmp01 = .Properties("Description")
8280              strTmp01 = StringReplace(strTmp01, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")
8290              .Properties("Description") = strTmp01
8300            End With
8310            Set qdf = Nothing
8320            lngQrysCreated = lngQrysCreated + 1&
8330          End If

8340          If arr_varQry(Q_FLD4, lngX) = True Then
                ' ** Update zzz_qry_MasterTrust_24_05_001_03, for journalno.  tblLedgerHidden
8350            strQryName2 = "zzz_qry_MasterTrust_24_05_001_24"
8360            strQryName3 = StringReplace(strQryName2, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")  ' ** Module Function: modStringFuncs.
8370            Select Case arr_varQry(Q_SET, lngX)
                Case 1&
                  ' ** Copy as-is.
8380            Case 2&
8390              strTmp01 = Right(strQryName3, 2)
8400              strTmp01 = CStr(Val(strTmp01) + 10&)
8410              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
8420            Case 3&
8430              strTmp01 = Right(strQryName3, 2)
8440              strTmp01 = CStr(Val(strTmp01) + 20&)
8450              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
8460            Case 4&
8470              strTmp01 = Right(strQryName3, 2)
8480              strTmp01 = CStr(Val(strTmp01) + 30&)
8490              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
8500            End Select
8510            DoCmd.CopyObject , strQryName3, acQuery, strQryName2
8520            DoEvents
8530            .QueryDefs.Refresh
8540            Set qdf = .QueryDefs(strQryName3)
8550            With qdf
8560              strSQL1 = .SQL
8570              strSQL1 = StringReplace(strSQL1, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")
8580              If arr_varQry(Q_SET, lngX) = 1& Then
                    ' ** Should be fine as-is.
8590              Else
8600                strTmp01 = "_" & arr_varQry(Q_NUM, lngX) & "_03"
8610                strTmp02 = Right(strQryName1, 7)  '"_001_03"
8620                strSQL1 = StringReplace(strSQL1, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
8630              End If
8640              .SQL = strSQL1
8650              strTmp01 = .Properties("Description")
8660              strTmp01 = StringReplace(strTmp01, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")
8670              .Properties("Description") = strTmp01
8680            End With
8690            Set qdf = Nothing
8700            lngQrysCreated = lngQrysCreated + 1&
8710          End If

8720          If arr_varQry(Q_FLD5, lngX) = True Then
                ' ** Update zzz_qry_MasterTrust_24_05_001_03, for ledghid_uniqueid.  tblLedgerHidden
8730            strQryName2 = "zzz_qry_MasterTrust_24_05_001_25"
8740            strQryName3 = StringReplace(strQryName2, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")  ' ** Module Function: modStringFuncs.
8750            Select Case arr_varQry(Q_SET, lngX)
                Case 1&
                  ' ** Copy as-is.
8760            Case 2&
8770              strTmp01 = Right(strQryName3, 2)
8780              strTmp01 = CStr(Val(strTmp01) + 10&)
8790              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
8800            Case 3&
8810              strTmp01 = Right(strQryName3, 2)
8820              strTmp01 = CStr(Val(strTmp01) + 20&)
8830              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
8840            Case 4&
8850              strTmp01 = Right(strQryName3, 2)
8860              strTmp01 = CStr(Val(strTmp01) + 30&)
8870              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
8880            End Select
8890            DoCmd.CopyObject , strQryName3, acQuery, strQryName2
8900            DoEvents
8910            .QueryDefs.Refresh
8920            Set qdf = .QueryDefs(strQryName3)
8930            With qdf
8940              strSQL1 = .SQL
8950              strSQL1 = StringReplace(strSQL1, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")
8960              If arr_varQry(Q_SET, lngX) = 1& Then
                    ' ** Should be fine as-is.
8970              Else
8980                strTmp01 = "_" & arr_varQry(Q_NUM, lngX) & "_03"
8990                strTmp02 = Right(strQryName1, 7)  '"_001_03"
9000                strSQL1 = StringReplace(strSQL1, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
9010              End If
9020              .SQL = strSQL1
9030              strTmp01 = .Properties("Description")
9040              strTmp01 = StringReplace(strTmp01, "_001_", "_" & arr_varQry(Q_NUM, lngX) & "_")
9050              .Properties("Description") = strTmp01
9060            End With
9070            Set qdf = Nothing
9080            lngQrysCreated = lngQrysCreated + 1&
9090          End If

9100          If arr_varQry(Q_FLD6, lngX) = True Then
                ' ** Update zzz_qry_MasterTrust_24_05_007_03, for ledghidtype_type.  tblLedgerHidden
9110            strQryName2 = "zzz_qry_MasterTrust_24_05_007_26"
9120            strQryName3 = StringReplace(strQryName2, "_007_", "_" & arr_varQry(Q_NUM, lngX) & "_")  ' ** Module Function: modStringFuncs.
9130            Select Case arr_varQry(Q_SET, lngX)
                Case 1&
                  ' ** Copy as-is.
9140            Case 2&
9150              strTmp01 = Right(strQryName3, 2)
9160              strTmp01 = CStr(Val(strTmp01) + 10&)
9170              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
9180            Case 3&
9190              strTmp01 = Right(strQryName3, 2)
9200              strTmp01 = CStr(Val(strTmp01) + 20&)
9210              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
9220            Case 4&
9230              strTmp01 = Right(strQryName3, 2)
9240              strTmp01 = CStr(Val(strTmp01) + 30&)
9250              strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
9260            End Select
9270            DoCmd.CopyObject , strQryName3, acQuery, strQryName2
9280            DoEvents
9290            .QueryDefs.Refresh
9300            Set qdf = .QueryDefs(strQryName3)
9310            With qdf
9320              strSQL1 = .SQL
9330              strSQL1 = StringReplace(strSQL1, "_007_", "_" & arr_varQry(Q_NUM, lngX) & "_")
9340              If arr_varQry(Q_SET, lngX) = 1& Then
                    ' ** Should be fine as-is.
9350              Else
9360                strTmp01 = "_" & arr_varQry(Q_NUM, lngX) & "_03"
9370                strTmp02 = Right(strQryName1, 7)  '"_001_03"
9380                strSQL1 = StringReplace(strSQL1, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
9390              End If
9400              .SQL = strSQL1
9410              strTmp01 = .Properties("Description")
9420              strTmp01 = StringReplace(strTmp01, "_007_", "_" & arr_varQry(Q_NUM, lngX) & "_")
9430              .Properties("Description") = strTmp01
9440            End With
9450            Set qdf = Nothing
9460            lngQrysCreated = lngQrysCreated + 1&
9470          End If

9480          If arr_varQry(Q_FLD7, lngX) = True Then
9490            If arr_varQry(Q_NUM, lngX) <> "019" Then
                  ' ** Update zzz_qry_MasterTrust_24_05_019_03, for ledghid_cnt.  tblLedgerHidden
9500              strQryName2 = "zzz_qry_MasterTrust_24_05_019_27"
9510              strQryName3 = StringReplace(strQryName2, "_019_", "_" & arr_varQry(Q_NUM, lngX) & "_")  ' ** Module Function: modStringFuncs.
9520              Select Case arr_varQry(Q_SET, lngX)
                  Case 1&
                    ' ** Copy as-is.
9530              Case 2&
9540                strTmp01 = Right(strQryName3, 2)
9550                strTmp01 = CStr(Val(strTmp01) + 10&)
9560                strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
9570              Case 3&
9580                strTmp01 = Right(strQryName3, 2)
9590                strTmp01 = CStr(Val(strTmp01) + 20&)
9600                strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
9610              Case 4&
9620                strTmp01 = Right(strQryName3, 2)
9630                strTmp01 = CStr(Val(strTmp01) + 30&)
9640                strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
9650              End Select
9660              DoCmd.CopyObject , strQryName3, acQuery, strQryName2
9670              DoEvents
9680              .QueryDefs.Refresh
9690              Set qdf = .QueryDefs(strQryName3)
9700              With qdf
9710                strSQL1 = .SQL
9720                strSQL1 = StringReplace(strSQL1, "_019_", "_" & arr_varQry(Q_NUM, lngX) & "_")
9730                If arr_varQry(Q_SET, lngX) = 1& Then
                      ' ** Should be fine as-is.
9740                Else
9750                  strTmp01 = "_" & arr_varQry(Q_NUM, lngX) & "_03"
9760                  strTmp02 = Right(strQryName1, 7)  '"_001_03"
9770                  strSQL1 = StringReplace(strSQL1, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
9780                End If
9790                .SQL = strSQL1
9800                strTmp01 = .Properties("Description")
9810                strTmp01 = StringReplace(strTmp01, "_019_", "_" & arr_varQry(Q_NUM, lngX) & "_")
9820                .Properties("Description") = strTmp01
9830              End With
9840              Set qdf = Nothing
9850              lngQrysCreated = lngQrysCreated + 1&
9860            End If
9870          End If

9880          If arr_varQry(Q_FLD8, lngX) = True Then
9890            If arr_varQry(Q_NUM, lngX) <> "019" Then
                  ' ** Update zzz_qry_MasterTrust_24_05_019_03, for ledghid_grpnum.  tblLedgerHidden
9900              strQryName2 = "zzz_qry_MasterTrust_24_05_019_28"
9910              strQryName3 = StringReplace(strQryName2, "_019_", "_" & arr_varQry(Q_NUM, lngX) & "_")  ' ** Module Function: modStringFuncs.
9920              Select Case arr_varQry(Q_SET, lngX)
                  Case 1&
                    ' ** Copy as-is.
9930              Case 2&
9940                strTmp01 = Right(strQryName3, 2)
9950                strTmp01 = CStr(Val(strTmp01) + 10&)
9960                strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
9970              Case 3&
9980                strTmp01 = Right(strQryName3, 2)
9990                strTmp01 = CStr(Val(strTmp01) + 20&)
10000               strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
10010             Case 4&
10020               strTmp01 = Right(strQryName3, 2)
10030               strTmp01 = CStr(Val(strTmp01) + 30&)
10040               strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
10050             End Select
10060             DoCmd.CopyObject , strQryName3, acQuery, strQryName2
10070             DoEvents
10080             .QueryDefs.Refresh
10090             Set qdf = .QueryDefs(strQryName3)
10100             With qdf
10110               strSQL1 = .SQL
10120               strSQL1 = StringReplace(strSQL1, "_019_", "_" & arr_varQry(Q_NUM, lngX) & "_")
10130               If arr_varQry(Q_SET, lngX) = 1& Then
                      ' ** Should be fine as-is.
10140               Else
10150                 strTmp01 = "_" & arr_varQry(Q_NUM, lngX) & "_03"
10160                 strTmp02 = Right(strQryName1, 7)  '"_001_03"
10170                 strSQL1 = StringReplace(strSQL1, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
10180               End If
10190               .SQL = strSQL1
10200               strTmp01 = .Properties("Description")
10210               strTmp01 = StringReplace(strTmp01, "_019_", "_" & arr_varQry(Q_NUM, lngX) & "_")
10220               .Properties("Description") = strTmp01
10230             End With
10240             Set qdf = Nothing
10250             lngQrysCreated = lngQrysCreated + 1&
10260           End If
10270         End If

10280         If arr_varQry(Q_FLD9, lngX) = True Then
10290           If arr_varQry(Q_NUM, lngX) <> "019" Then
                  ' ** Update zzz_qry_MasterTrust_24_05_019_03, for ledghid_ord.  tblLedgerHidden
10300             strQryName2 = "zzz_qry_MasterTrust_24_05_019_29"
10310             strQryName3 = StringReplace(strQryName2, "_019_", "_" & arr_varQry(Q_NUM, lngX) & "_")  ' ** Module Function: modStringFuncs.
10320             Select Case arr_varQry(Q_SET, lngX)
                  Case 1&
                    ' ** Copy as-is.
10330             Case 2&
10340               strTmp01 = Right(strQryName3, 2)
10350               strTmp01 = CStr(Val(strTmp01) + 10&)
10360               strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
10370             Case 3&
10380               strTmp01 = Right(strQryName3, 2)
10390               strTmp01 = CStr(Val(strTmp01) + 20&)
10400               strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
10410             Case 4&
10420               strTmp01 = Right(strQryName3, 2)
10430               strTmp01 = CStr(Val(strTmp01) + 30&)
10440               strQryName3 = Left(strQryName3, (Len(strQryName3) - 2)) & strTmp01
10450             End Select
10460             DoCmd.CopyObject , strQryName3, acQuery, strQryName2
10470             DoEvents
10480             .QueryDefs.Refresh
10490             Set qdf = .QueryDefs(strQryName3)
10500             With qdf
10510               strSQL1 = .SQL
10520               strSQL1 = StringReplace(strSQL1, "_019_", "_" & arr_varQry(Q_NUM, lngX) & "_")
10530               If arr_varQry(Q_SET, lngX) = 1& Then
                      ' ** Should be fine as-is.
10540               Else
10550                 strTmp01 = "_" & arr_varQry(Q_NUM, lngX) & "_03"
10560                 strTmp02 = Right(strQryName1, 7)  '"_001_03"
10570                 strSQL1 = StringReplace(strSQL1, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
10580               End If
10590               .SQL = strSQL1
10600               strTmp01 = .Properties("Description")
10610               strTmp01 = StringReplace(strTmp01, "_019_", "_" & arr_varQry(Q_NUM, lngX) & "_")
10620               .Properties("Description") = strTmp01
10630             End With
10640             Set qdf = Nothing
10650             lngQrysCreated = lngQrysCreated + 1&
10660           End If
10670         End If

10680         If (lngX + 1&) Mod 100 = 0 Then
10690           Debug.Print "|  " & CStr(lngX + 1&) & " of "; CStr(lngQrys)
10700           Debug.Print "'|";
10710         ElseIf (lngX + 1&) Mod 10 = 0 Then
10720           Debug.Print "|";
10730         Else
10740           Debug.Print ".";
10750         End If
10760         DoEvents

10770       Next
10780       Debug.Print
10790       DoEvents

10800       Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
10810       DoEvents

            'If arr_varQry(Q_NUM, lngX) = "019" Then

            'icash_new
            'ledger_HIDDEN_new
            'journaltype_new
            'journalno_new
            'ledghid_uniqueid_new
            'ledghidtype_type_new
            'ledghid_cnt_new
            'ledghid_grpnum_new
            'ledghid_ord_new

            'Q_QNAM
            'Q_SQL
            'Q_DSC
            'Q_NUM
            'Q_SET
            'Q_FLD1
            'Q_FLD2
            'Q_FLD3
            'Q_FLD4
            'Q_FLD5
            'Q_FLD6
            'Q_FLD7
            'Q_FLD8
            'Q_FLD9

            ' ** NEW FLDS: 9
            ' ** Ledger:
            ' **   1.  icash_new
            ' **         Update zzz_qry_MasterTrust_24_05_001_02, for icash.  Ledger
            ' **         zzz_qry_MasterTrust_24_05_001_21
            ' **   2.  ledger_HIDDEN_new
            ' **         Update zzz_qry_MasterTrust_24_05_001_02, for ledger_HIDDEN.  Ledger
            ' **         zzz_qry_MasterTrust_24_05_001_22
            ' **   3.  journaltype_new
            ' **         Update zzz_qry_MasterTrust_24_05_001_02, for journaltype.  Ledger
            ' **         zzz_qry_MasterTrust_24_05_001_23
            ' ** tblLedgerHidden:
            ' **   4.  journalno_new
            ' **         Update zzz_qry_MasterTrust_24_05_001_03, for journalno.  tblLedgerHidden
            ' **         zzz_qry_MasterTrust_24_05_001_24
            ' **   5.  ledghid_uniqueid_new
            ' **         Update zzz_qry_MasterTrust_24_05_001_03, for ledghid_uniqueid.  tblLedgerHidden
            ' **         zzz_qry_MasterTrust_24_05_001_25
            ' **   6.  ledghidtype_type_new
            ' **         Update zzz_qry_MasterTrust_24_05_007_03, for ledghidtype_type.  tblLedgerHidden
            ' **         zzz_qry_MasterTrust_24_05_007_26
            ' **   7.  ledghid_cnt_new
            ' **         Update zzz_qry_MasterTrust_24_05_019_03, for ledghid_cnt.  tblLedgerHidden
            ' **         zzz_qry_MasterTrust_24_05_019_27
            ' **   8.  ledghid_grpnum_new
            ' **         Update zzz_qry_MasterTrust_24_05_019_03, for ledghid_grpnum.  tblLedgerHidden
            ' **         zzz_qry_MasterTrust_24_05_019_28
            ' **   9.  ledghid_ord_new
            ' **         Update zzz_qry_MasterTrust_24_05_019_03, for ledghid_ord.  tblLedgerHidden
            ' **         zzz_qry_MasterTrust_24_05_019_29

10820     End If  ' ** blnSkip.

10830     blnSkip = True
10840     If blnSkip = False Then

10850       lngQrys = 0&
10860       ReDim arr_varQry(Q_ELEMS, 0)

10870       intLen = Len(QRY_BASE)

10880       lngTmp07 = 0&: strTmp02 = vbNullString
10890       For Each qdf In .QueryDefs
10900         With qdf
10910           If Left(.Name, intLen) = QRY_BASE Then
10920             If .Type = dbQUpdate Then
10930               strQryNum = Mid(.Name, (intLen + 1), 3)
10940               If Val(strQryNum) >= 62 Then
10950                 strTmp01 = Right(.Name, 2)
10960                 If Right(strTmp01, 1) = "7" Or Right(strTmp01, 1) = "8" Or Right(strTmp01, 1) = "9" Then
                        ' ** ledghid_cnt, ledghid_grpnum, ledghid_ord
10970                   blnFound = False
10980                   For lngX = 0& To (lngQrys - 1&)
10990                     If arr_varQry(Q_NUM, lngX) = strQryNum Then
11000                       blnFound = True
11010                       Exit For
11020                     End If
11030                   Next
11040                   If blnFound = False Then
11050                     lngQrys = lngQrys + 1&
11060                     lngE = lngQrys - 1&
11070                     ReDim Preserve arr_varQry(Q_ELEMS, lngE)
11080                     arr_varQry(Q_NUM, lngE) = strQryNum
11090                   End If
11100                 End If
11110               End If
11120               If strQryNum <> strTmp02 Then
11130                 strTmp02 = strQryNum
11140                 lngTmp07 = lngTmp07 + 1&  ' ** Count the number of accounts.
11150               End If
11160             End If
11170           End If
11180         End With
11190       Next
11200       Set qdf = Nothing

11210       Debug.Print "'ACCTS W/ LEDGHID: " & CStr(lngQrys)
11220       DoEvents

11230       For lngX = 0& To (lngQrys - 1&)
11240         strQryName1 = QRY_BASE & arr_varQry(Q_NUM, lngX) & "_01"
11250         Set qdf = .QueryDefs(strQryName1)
11260         With qdf
11270           strTmp01 = .Properties("Description")
11280           strTmp01 = "#" & strTmp01
11290           .Properties("Description") = strTmp01
11300         End With
11310         Set qdf = Nothing
11320       Next

11330       Debug.Print "'|";
11340       DoEvents

11350       lngQrysCreated = 0&: strTmp02 = vbNullString: lngY = 0&
11360       For Each qdf In .QueryDefs
11370         With qdf
11380           If Left(.Name, intLen) = QRY_BASE Then
11390             If .Type = dbQUpdate Then
11400               strQryNum = Mid(.Name, (intLen + 1), 3)
11410               If Val(strQryNum) >= 62 Then
11420                 blnFound = False
11430                 For lngX = 0& To (lngQrys - 1&)
11440                   If arr_varQry(Q_NUM, lngX) = strQryNum Then
11450                     blnFound = True
11460                     Exit For
11470                   End If
11480                 Next
11490                 If blnFound = False Then
11500                   .Execute
11510                   lngQrysCreated = lngQrysCreated + 1&
11520                 End If
11530               End If
11540               If strQryNum <> strTmp02 Then
11550                 strTmp02 = strQryNum
11560                 lngY = lngY + 1&
11570                 If lngY Mod 100 = 0 Then
11580                   Debug.Print "|  " & CStr(lngY) & " of " & CStr(lngTmp07)
11590                   Debug.Print "'|";
11600                 ElseIf lngY Mod 10 = 0 Then
11610                   Debug.Print "|";
11620                 Else
11630                   Debug.Print ".";
11640                 End If
11650                 DoEvents
11660               End If
11670             End If
11680           End If
11690         End With

11700       Next
11710       Debug.Print
11720       DoEvents

11730       Debug.Print "'QRYS RUN: " & CStr(lngQrysCreated)
11740       DoEvents

11750     End If  ' ** blnSkip.

11760     blnSkip = False
11770     If blnSkip = False Then

11780       intLen = Len(QRY_BASE)

11790       For Each qdf In .QueryDefs
11800         With qdf
11810           If Left(.Name, intLen) = QRY_BASE Then
11820             If Right(.Name, 2) = "01" Then
11830               strTmp01 = .Properties("Description")
11840               If Left(strTmp01, 1) = "#" Then
11850                 Debug.Print "'" & .Name
11860                 DoEvents
11870               End If
11880             End If
11890           End If
11900         End With
11910       Next

            'X zzz_qry_MasterTrust_24_05_091_01
            'X zzz_qry_MasterTrust_24_05_092_01
            'X zzz_qry_MasterTrust_24_05_093_01
            'X zzz_qry_MasterTrust_24_05_098_01
            'X zzz_qry_MasterTrust_24_05_100_01
            'X zzz_qry_MasterTrust_24_05_102_01
            'X zzz_qry_MasterTrust_24_05_104_01
            'X zzz_qry_MasterTrust_24_05_106_01
            'X zzz_qry_MasterTrust_24_05_108_01
            'X zzz_qry_MasterTrust_24_05_112_01
            'X zzz_qry_MasterTrust_24_05_113_01
            'X zzz_qry_MasterTrust_24_05_114_01
            'X zzz_qry_MasterTrust_24_05_115_01
            'X zzz_qry_MasterTrust_24_05_203_01
            'X zzz_qry_MasterTrust_24_05_208_01
            'X zzz_qry_MasterTrust_24_05_224_01
            'X zzz_qry_MasterTrust_24_05_259_01
            'X zzz_qry_MasterTrust_24_05_267_01
            'X zzz_qry_MasterTrust_24_05_269_01
            'x zzz_qry_MasterTrust_24_05_341_01
            'DONE!

11920     End If  ' ** blnSkip.

11930     .Close
11940   End With

        'SPECIAL HANDLING!
        'zzz_qry_MasterTrust_24_05_019_01
        'zzz_qry_MasterTrust_24_05_025_01
        'zzz_qry_MasterTrust_24_05_058_01
        'zzz_qry_MasterTrust_24_05_091_01
        'zzz_qry_MasterTrust_24_05_092_01
        'zzz_qry_MasterTrust_24_05_093_01
        'zzz_qry_MasterTrust_24_05_098_01
        'zzz_qry_MasterTrust_24_05_100_01
        'zzz_qry_MasterTrust_24_05_102_01
        'zzz_qry_MasterTrust_24_05_104_01
        'zzz_qry_MasterTrust_24_05_106_01
        'zzz_qry_MasterTrust_24_05_108_01
        'zzz_qry_MasterTrust_24_05_112_01
        'zzz_qry_MasterTrust_24_05_113_01
        'zzz_qry_MasterTrust_24_05_114_01
        'zzz_qry_MasterTrust_24_05_115_01
        'zzz_qry_MasterTrust_24_05_203_01
        'zzz_qry_MasterTrust_24_05_208_01
        'zzz_qry_MasterTrust_24_05_213_01
        'zzz_qry_MasterTrust_24_05_224_01
        'zzz_qry_MasterTrust_24_05_259_01
        'zzz_qry_MasterTrust_24_05_267_01
        'zzz_qry_MasterTrust_24_05_269_01
        'zzz_qry_MasterTrust_24_05_341_01

11950   Beep

11960   Debug.Print "'DONE!"
11970   DoEvents

EXITP:
11980   Set fld = Nothing
11990   Set rst = Nothing
12000   Set prp = Nothing
12010   Set qdf = Nothing
12020   Set dbs = Nothing
12030   Hide_FixQrys2 = blnRetVal
12040   Exit Function

ERRH:
12050   blnRetVal = False
12060   Select Case ERR.Number
        Case Else
12070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12080   End Select
12090   Resume EXITP

End Function
