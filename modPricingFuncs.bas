Attribute VB_Name = "modPricingFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modPricingFuncs"

'VGC 03/23/2017: CHANGES!

'Private Const FT_ELEMS As Integer = 2  ' ** Array's first-element UBound().
'Private Const FT_ID  As Integer = 0  ' ** priceimport_id.
Private Const FT_EXT As Integer = 1  ' ** Extension.
Private Const FT_PF  As Integer = 2  ' ** Path\File.
' **

Public Sub GetData1(frm As Access.Form, blnUpdateSuccess As Boolean, lngFileTypes As Long, arr_varFileType As Variant, strIMPORT_FileType As String, dblProgBox_Width As Double, dblProgBar_Len As Double, intPct As Integer, dblProgBar_Incr As Double)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "GetData1"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rstEstateVal As DAO.Recordset
        Dim rstMasterAsset As DAO.Recordset, rstAppraise As DAO.Recordset, rstUpdated As DAO.Recordset
        Dim strPricePath As String, strPriceFile As String
        Dim dblProgressAmount As Double, dblProgBar_LenStart As Double
        Dim lngRecs As Long
        Dim lngUpdateCount As Long, lngSecCnt As Long
        Dim strPricePathFile_APP As String, strPricePathFile_TXT_conv As String
        Dim lngLines As Long, lngThisLine As Long
        Dim lngRowTypes As Long, arr_varRowType As Variant
        Dim strInput As String
        Dim arr_varRetVal As Variant
        Dim blnFound As Boolean
        Dim lngFileElem As Long
        Dim dblProgBar_Sec01 As Double, dblProgBar_Sec02 As Double, dblProgBar_Sec03 As Double
        Dim dblProgBar_Sec04 As Double, dblProgBar_Sec05 As Double ', dblProgBar_Sec06 As Double
        Dim intPos01 As Integer, intPos02 As Integer
        Dim strTmp01 As String, lngTmp02 As Long, dblTmp03 As Double
        Dim lngX As Long

        'Const R_TYP As Integer = 0
        'Const R_CNT As Integer = 1

        ' ** Update the progress bar.
110     With frm
120       dblProgBar_Len = dblProgBar_Len + (1# * dblProgBar_Incr)
130       .ProgBar_Width_Pric dblProgBar_Len, 2  ' ** Form Procedure: frmAssetPricing_Import.
          '.ProgBar_bar.Width = dblProgBar_Len
140       intPct = CInt((dblProgBar_Len / dblProgBox_Width) * 100#)
150       .ProgBar_lbl.Caption = CStr(intPct) & "%"
160       DoEvents
170     End With

180     lngUpdateCount = 0&
190     blnFound = False: strPricePathFile_TXT_conv = vbNullString

200     strPricePath = (gstrTrustDataLocation & gstrDir_Pricing)  ' ** The file's been copied to the standard Pricing directory.
210     strPriceFile = frm.priceimport_file

220     If FileExists(strPricePath & LNK_SEP & strPriceFile) = True Then  ' ** Module Function: modFileUtilities.
          ' ** It's where we expected it to be.
230       For lngX = 0& To (lngFileTypes - 1&)
240         If arr_varFileType(FT_EXT, lngX) = strIMPORT_FileType Then
250           blnFound = True
260           arr_varFileType(FT_PF, lngX) = strPricePath & LNK_SEP & strPriceFile
270           lngFileElem = lngX
280           Exit For
290         End If
300       Next
310     End If

320     If blnFound = True Then

330       Set dbs = CurrentDb

          ' ** Update the progress bar.
340       With frm
350         dblProgBar_Len = dblProgBar_Len + (5# * dblProgBar_Incr)
360         .ProgBar_Width_Pric dblProgBar_Len, 2  ' ** Form Procedure: frmAssetPricing_Import.
            '.ProgBar_bar.Width = dblProgBar_Len
370         intPct = CInt((dblProgBar_Len / dblProgBox_Width) * 100#)
380         .ProgBar_lbl.Caption = CStr(intPct) & "%"
390         DoEvents
400       End With

410       Select Case strIMPORT_FileType
          Case ".txt"
            ' ** In the future, there may be different types of TXT files.
420         arr_varRetVal = GetData2(frm, CStr(arr_varFileType(FT_PF, lngFileElem)), _
              dblProgBox_Width, dblProgBar_Len, intPct)  ' ** Procedure: Below.
430         lngUpdateCount = arr_varRetVal(0)
440         lngSecCnt = arr_varRetVal(1)
450         If lngUpdateCount = 0& Then
460           blnFound = False
470         End If

480       Case ".rtf"
            ' ** In the future, there may be different types of RTF files.

490         strTmp01 = ConvertRTFtoTXT(CStr(arr_varFileType(FT_PF, lngFileElem)))  ' ** Module Function: modFileUtilities.

500         If strTmp01 <> RET_ERR And strTmp01 <> vbNullString Then
510           blnFound = True
              'MsgBox "Converted RTF file:" & vbCrLf & "  " & strPricePathFile_TXT
520           strPricePathFile_TXT_conv = strTmp01
530           arr_varRetVal = GetData2(frm, CStr(arr_varFileType(FT_PF, lngFileElem)), _
                dblProgBox_Width, dblProgBar_Len, intPct)  ' ** Procedure: Below.
540           lngUpdateCount = arr_varRetVal(0)
550           lngSecCnt = arr_varRetVal(1)
560           If lngUpdateCount = -9 Then
                ' ** An error occurred.
570             blnFound = False
580           Else
590             If lngUpdateCount = 0& Then
600               blnFound = False
610             End If
620           End If
630         Else
640           blnFound = False
650           strPricePathFile_TXT_conv = vbNullString
660         End If

670       Case ".prn"
            ' ** In the future, there may be different types of PRN files.

            ' ** Copy PRN file to TXT.
680         strTmp01 = Left(arr_varFileType(FT_PF, lngFileElem), (Len(arr_varFileType(FT_PF, lngFileElem)) - 4)) & ".txt"
690         FileCopy arr_varFileType(FT_PF, lngFileElem), strTmp01
700         strPricePathFile_TXT_conv = strTmp01

            ' ** Import TXT file to tblPricing_EstateVal1 table using Import Spec 'EstateVal_PRN'.
710         DoCmd.TransferText acImportDelim, "EstateVal_PRN", "tblPricing_EstateVal1", strTmp01, False
            'DoCmd.TransferText [transfertype][, specificationname], tablename, filename[, hasfieldnames][, HTMLtablename][, codepage]

            ' ** tblPricing_EstateVal1, just needed fields.
720         Set qdf = dbs.QueryDefs("qryPricing_EstateVal_03")
730         Set rstEstateVal = qdf.OpenRecordset
740         With rstEstateVal
750           If .BOF = True And .EOF = True Then
760             blnFound = False
770             MsgBox "No valid records were found in the imported file!", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
780           End If
790           .Close
800         End With

810         If blnFound = True Then
              ' ** Append good imported records to tblPricing_EstateVal2.
820           Set qdf = dbs.QueryDefs("qryPricing_EstateVal_04")
830           qdf.Execute
              ' ** tblPricing_EstateVal2 linked to tblPricing_MasterAsset.
840           Set qdf = dbs.QueryDefs("qryPricing_EstateVal_05")
850           Set rstEstateVal = qdf.OpenRecordset
860           If rstEstateVal.BOF = True And rstEstateVal.EOF = True Then
870             blnFound = False
880             MsgBox "No valid records were found in the imported file!", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
890             rstEstateVal.Close
900           End If
910         End If

920         If blnFound = True Then

930           Set rstMasterAsset = dbs.OpenRecordset("tblPricing_MasterAsset", dbOpenDynaset, dbConsistent)

940           With rstEstateVal
950             .MoveLast
960             lngRecs = .RecordCount
970             .MoveFirst
980             dblProgressAmount = ((dblProgBox_Width - dblProgBar_Len) / lngRecs)
990             For lngX = 1& To lngRecs
1000              rstMasterAsset.FindFirst "[cusip] = '" & ![cusip] & "'"
1010              If rstMasterAsset.NoMatch = False Then
1020                rstMasterAsset.Edit
1030                rstMasterAsset![marketvalue] = ![priceperunit]
                    'WE NEED TO KNOW UNDER WHAT CONDITIONS THE Yield SHOULDN'T BE UPDATED!
1040                If IsNull(![yield_new]) = False Then
1050                  If ![yield_new] <> 0# Then
                        ' ** The saved, raw number is limited to a precision of 6 decimal places,
                        ' ** so the displayed percentage is limited to a precision of 4.
1060                    If ![yield_new] < 1# Then
1070                      rstMasterAsset![yield] = Round(![yield_new], 6)
1080                      rstMasterAsset![yield_entry] = Round((Round(![yield_new], 6) * 100#), 4)
1090                    Else
1100                      rstMasterAsset![yield] = Round((Round(![yield_new], 4) / 100#), 6)
1110                      rstMasterAsset![yield_entry] = Round(![yield_new], 4)
1120                    End If
1130                  Else
1140                    rstMasterAsset![yield] = 0#
1150                    rstMasterAsset![yield_entry] = 0#
1160                  End If
1170                Else
1180                  rstMasterAsset![yield] = 0#
1190                  rstMasterAsset![yield_entry] = 0#
1200                End If
1210                rstMasterAsset.Update
1220                lngUpdateCount = lngUpdateCount + 1&  ' ** Increment count.
1230              Else
1240                blnFound = False
1250                MsgBox "MasterAsset record not found!", vbExclamation + vbOKOnly, "Asset Not Found"
1260                Exit For
1270              End If
                  ' ** Update the progress bar.
1280              With frm
1290                dblProgBar_Len = dblProgBar_Len + dblProgressAmount
1300                .ProgBar_Width_Pric dblProgBar_Len, 2  ' ** Form Procedure: frmAssetPricing_Import.
                    '.ProgBar_bar.Width = dblProgBar_Len
1310                intPct = CInt((dblProgBar_Len / dblProgBox_Width) * 100#)
1320                .ProgBar_lbl.Caption = CStr(intPct) & "%"
1330                DoEvents
1340              End With
1350              If lngX < lngRecs Then .MoveNext
1360            Next
1370            .Close
1380          End With

1390          rstMasterAsset.Close

1400        End If

1410      Case ".ddt", ".edt", ".adt"
            ' ** Appraise export-to-data files.

            ' ** Read the file directly, and put it into tblPricing_Appraise_Raw.
1420        strPricePathFile_APP = strPricePath & LNK_SEP & strPriceFile

1430  On Error Resume Next
1440        Open strPricePathFile_APP For Input As #1
1450        If ERR.Number <> 0 Then
1460          Select Case ERR.Number
              Case 55  ' ** File already open.
1470            Close #1
1480  On Error GoTo ERRH
1490            Open strPricePathFile_APP For Input As #1
1500          Case Else
1510            GoTo ERRH
1520  On Error GoTo ERRH
1530          End Select
1540        Else
1550  On Error GoTo ERRH
1560        End If

1570        lngLines = 0&: lngThisLine = 0&: lngRecs = 0&

            ' ** Find out how many lines are in the file.
1580        Do While Not EOF(1)
1590          lngLines = lngLines + 1&
1600          Line Input #1, strInput
1610        Loop
1620        Close #1

1630        Set rstAppraise = dbs.OpenRecordset("tblPricing_Appraise_Raw", dbOpenDynaset, dbConsistent)

1640        Open strPricePathFile_APP For Input As #1

            'NEED TO WEIGHT THIS!!!!!!!!
1650        dblProgressAmount = ((dblProgBox_Width - dblProgBar_Len) / lngLines)

1660        Do While Not EOF(1)

1670          Line Input #1, strInput
1680          lngThisLine = lngThisLine + 1&

              ' ** Trim trailing and leading spaces.
1690          strInput = Trim(strInput)

1700          If strInput <> vbNullString Then
1710            lngRecs = lngRecs + 1&
1720            With rstAppraise
1730              .AddNew
                  ' ** apraw_id   : AutoNumber.
                  ' ** apraw_port : Don't know the Portfolio name yet.
1740              ![apraw_raw] = strInput  ' ** It better go in with all its quotes!
1750              lngTmp02 = CharCnt(strInput, ",")  ' ** Module Function: modStringFuncs.
1760              ![apraw_commas] = lngTmp02
1770              ![apraw_fields] = lngTmp02 + 1&
1780              .Update
1790            End With
1800          End If

              'SECTION 1:
              ' ** Update the progress bar.
1810          With frm
1820            dblProgBar_Len = dblProgBar_Len + dblProgressAmount
1830            .ProgBar_Width_Pric dblProgBar_Len, 2  ' ** Form Procedure: frmAssetPricing_Import.
                '.ProgBar_bar.Width = dblProgBar_Len
1840            intPct = CInt((dblProgBar_Len / dblProgBox_Width) * 100#)
1850            .ProgBar_lbl.Caption = CStr(intPct) & "%"
1860            DoEvents
1870          End With

1880        Loop

1890        Close #1

            'SECTION 2:
1900        If lngRecs > 0& Then
1910          blnFound = True
1920          With rstAppraise
1930            .MoveFirst
                ' ** The 2nd field of the first record should be PORT ID, the portfolio identifier.
1940            intPos01 = InStr(![apraw_raw], ",")
1950            intPos02 = InStr((intPos01 + 1), ![apraw_raw], ",")
1960            strTmp01 = Mid(![apraw_raw], (intPos01 + 1), ((intPos02 - intPos01) - 1))
1970            If strTmp01 <> (Chr(34) & Chr(34)) Then
1980              If Right(strTmp01, 1) = Chr(34) Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
1990              If Left(strTmp01, 1) = Chr(34) Then strTmp01 = Mid(strTmp01, 2)
2000            Else
                  ' ** It should always have a PORT ID.
2010              strTmp01 = "UNKNOWN"
2020            End If
2030            For lngX = 1& To lngRecs
2040              .Edit
2050              ![apraw_port] = strTmp01
2060              .Update
2070              If lngX < lngRecs Then .MoveNext
2080            Next
2090          End With
2100        End If
2110        rstAppraise.Close

            'SECTION 3:
            ' ** Append qryPricing_Appraise_10 (tblPricing_Appraise_Raw, with all fields broken out and datatyped) to tblAppraise_Header.
2120        Set qdf = dbs.QueryDefs("qryPricing_Appraise_11")
2130        qdf.Execute

            'SECTION 4:
            ' ** Append qryPricing_Appraise_12 (tblPricing_Appraise_Raw, with all fields broken out and datatyped) to tblAppraise_Data.
2140        Set qdf = dbs.QueryDefs("qryPricing_Appraise_13")
2150        qdf.Execute

            'lngUpdateCount

            'SECTION 5:

2160        lngRowTypes = 0&

            ' ** Get a list of aprtype_type's (asset types), with cnt, from tblPricing_Appraise_Data.
2170        Set qdf = dbs.QueryDefs("qryPricing_Appraise_16")
2180        Set rstAppraise = qdf.OpenRecordset
2190        With rstAppraise
2200          .MoveLast
2210          lngRowTypes = .RecordCount
2220          .MoveFirst
2230          arr_varRowType = .GetRows(lngRowTypes)
              ' *************************************************
              ' ** Array: arr_varRowType()
              ' **
              ' **   Field  Element  Name            Constant
              ' **   =====  =======  ==============  ==========
              ' **     1       0     aprtype_type    R_TYP
              ' **     2       1     cnt             R_CNT
              ' **
              ' *************************************************
2240          .Close
2250        End With

2260        For lngX = 0& To (lngRowTypes - 1&)

2270        Next

            'A bond has a 'Dated' field, and its price is price/100.

            'MasterAsset:
            '  currentdate = strPriceDate  from file: PRC DATE
            '  marketvalue = dblPrice      from file: CALC MEAN / MEAN/CLOSE
            '  yield       = ((dblIncome / intShares) / dblPrice)  from file: see Rich's fax
            'tblPricing_Cusip:
            '  for report

2280      Case ".dut", ".eut", ".aut"
            ' ** Appraise export-to-text files.

            ' ** Read the file directly, and put it into tblPricing_Appraise_Raw.
2290        strPricePathFile_APP = strPricePath & LNK_SEP & strPriceFile

2300  On Error Resume Next
2310        Open strPricePathFile_APP For Input As #1
2320        If ERR.Number <> 0 Then
2330          Select Case ERR.Number
              Case 55  ' ** File already open.
2340            Close #1
2350  On Error GoTo ERRH
2360            Open strPricePathFile_APP For Input As #1
2370          Case Else
2380            GoTo ERRH
2390  On Error GoTo ERRH
2400          End Select
2410        Else
2420  On Error GoTo ERRH
2430        End If

2440        lngLines = 0&: lngThisLine = 0&: lngRecs = 0&

            ' ** Find out how many lines are in the file.
2450        Do While Not EOF(1)
2460          lngLines = lngLines + 1&
2470          Line Input #1, strInput
2480        Loop
2490        Close #1

2500        Set rstAppraise = dbs.OpenRecordset("tblPricing_Appraise_Raw", dbOpenDynaset, dbConsistent)

2510        Open strPricePathFile_APP For Input As #1

            ' ** Weight each of the process sections.
2520        dblProgBar_Sec01 = 0.15
2530        dblProgBar_Sec02 = 0.15
2540        dblProgBar_Sec03 = 0.2
2550        dblProgBar_Sec04 = 0.2
2560        dblProgBar_Sec05 = 0.3

            ' ** SECTION 1:
2570        dblProgBar_LenStart = dblProgBar_Len
2580        dblTmp03 = (dblProgBox_Width - dblProgBar_LenStart)
2590        dblTmp03 = (dblTmp03 * dblProgBar_Sec01)
2600        dblProgressAmount = (dblTmp03 / lngLines)

2610        Do While Not EOF(1)

2620          Line Input #1, strInput
2630          lngThisLine = lngThisLine + 1&

              ' ** Trim trailing and leading spaces.
2640          strInput = Trim(strInput)

2650          If strInput <> vbNullString Then
2660            lngRecs = lngRecs + 1&
2670            With rstAppraise
2680              .AddNew
                  ' ** apraw_id     : AutoNumber.
                  ' ** apraw_port   : Don't know the Portfolio name yet.
2690              ![apraw_raw] = strInput  ' ** It better go in with all its quotes!
                  ' ** apraw_commas : Not used for text exports.
                  ' ** apraw_fields : Not used for text exports.
2700              .Update
2710            End With
2720          End If

              ' ** Update the progress bar.
2730          With frm
2740            dblProgBar_Len = dblProgBar_Len + dblProgressAmount
2750            .ProgBar_Width_Pric dblProgBar_Len, 2  ' ** Form Procedure: frmAssetPricing_Import.
                '.ProgBar_bar.Width = dblProgBar_Len
2760            intPct = CInt((dblProgBar_Len / dblProgBox_Width) * 100#)
2770            .ProgBar_lbl.Caption = CStr(intPct) & "%"
2780            DoEvents
2790          End With

2800        Loop

2810        Close #1

2820        If lngRecs > 0& Then
2830          blnFound = True

              ' ** SECTION 2:
2840          dblTmp03 = ((dblProgBox_Width - dblProgBar_LenStart) * dblProgBar_Sec02)
2850          dblProgressAmount = dblTmp03

              ' ** Update qryPricing_Appraise_32 (tblPricing_Appraise_Raw, with apraw_port_new).
2860          Set qdf = dbs.QueryDefs("qryPricing_Appraise_34")
2870          qdf.Execute

              ' ** Update the progress bar.
2880          With frm
2890            dblProgBar_Len = dblProgBar_Len + dblProgressAmount
2900            .ProgBar_Width_Pric dblProgBar_Len, 2  ' ** Form Procedure: frmAssetPricing_Import.
                '.ProgBar_bar.Width = dblProgBar_Len
2910            intPct = CInt((dblProgBar_Len / dblProgBox_Width) * 100#)
2920            .ProgBar_lbl.Caption = CStr(intPct) & "%"
2930            DoEvents
2940          End With

              ' ** SECTION 3:
2950          dblTmp03 = ((dblProgBox_Width - dblProgBar_LenStart) * dblProgBar_Sec03)
2960          dblProgressAmount = dblTmp03

              ' ** Append qryPricing_Appraise_38 (tblPricing_Appraise_Raw, with fields broken out) to tblPricing_Appraise1.
2970          Set qdf = dbs.QueryDefs("qryPricing_Appraise_40")
2980          qdf.Execute

              ' ** Update the progress bar.
2990          With frm
3000            dblProgBar_Len = dblProgBar_Len + dblProgressAmount
3010            .ProgBar_Width_Pric dblProgBar_Len, 2  ' ** Form Procedure: frmAssetPricing_Import.
                '.ProgBar_bar.Width = dblProgBar_Len
3020            intPct = CInt((dblProgBar_Len / dblProgBox_Width) * 100#)
3030            .ProgBar_lbl.Caption = CStr(intPct) & "%"
3040            DoEvents
3050          End With

              ' ** SECTION 4:
3060          dblTmp03 = ((dblProgBox_Width - dblProgBar_LenStart) * dblProgBar_Sec04)
3070          dblProgressAmount = dblTmp03

              ' ** Append qryPricing_Appraise_42(tblPricing_Appraise1, with the rest of the fields) to tblPricing_Appraise2.
3080          Set qdf = dbs.QueryDefs("qryPricing_Appraise_44")
3090          qdf.Execute

              ' ** Update the progress bar.
3100          With frm
3110            dblProgBar_Len = dblProgBar_Len + dblProgressAmount
3120            .ProgBar_Width_Pric dblProgBar_Len, 2  ' ** Form Procedure: frmAssetPricing_Import.
                '.ProgBar_bar.Width = dblProgBar_Len
3130            intPct = CInt((dblProgBar_Len / dblProgBox_Width) * 100#)
3140            .ProgBar_lbl.Caption = CStr(intPct) & "%"
3150            DoEvents
3160          End With

3170          Set rstUpdated = dbs.OpenRecordset("tblPricing_Cusip", dbOpenDynaset, dbConsistent)
3180          Set rstAppraise = dbs.OpenRecordset("tblPricing_Appraise2", dbOpenDynaset, dbConsistent)
3190          Set rstMasterAsset = dbs.OpenRecordset("tblPricing_MasterAsset", dbOpenDynaset, dbConsistent)

              ' ** SECTION 5:
3200          dblTmp03 = ((dblProgBox_Width - dblProgBar_LenStart) * dblProgBar_Sec05)
3210          dblProgressAmount = dblTmp03

3220          With rstAppraise
3230            If .BOF = True And .EOF = True Then
                  ' ** Oops! Something didn't work!
3240              blnFound = False
3250            Else
3260              .MoveLast
3270              lngSecCnt = .RecordCount
3280              .MoveFirst
3290              dblProgressAmount = (dblProgressAmount / lngSecCnt)
3300              For lngX = 1& To lngSecCnt

3310                rstMasterAsset.FindFirst "[cusip] = '" & ![cusip] & "'"
3320                If rstMasterAsset.NoMatch = True Then
                      ' ** If there is no match from the tblPricing_MasterAsset, we don't care about this one.
3330                Else
                      ' ** Found a match, so upate the tblPricing_MasterAsset.

3340                  rstUpdated.AddNew
3350                  rstUpdated![assetno] = rstMasterAsset![assetno]
3360                  rstUpdated![cusip] = Format(![cusip], "000000000")
3370                  rstUpdated![OldValue] = rstMasterAsset![marketvalue]
3380                  rstUpdated![OldDate] = rstMasterAsset![currentDate]
3390                  rstUpdated![NewValue] = ![price_close_mean]
3400                  rstUpdated![NewDate] = ![price_date]
3410                  rstUpdated![NewYield] = Round(![yield], 6)
3420                  rstUpdated.Update
                      ' ** yield_upd: IIf(IsNull([tblPricing_Cusip].[assetno])=True,False,True)
                      ' ** marketvalue_upd: IIf(IsNull([tblPricing_Cusip].[assetno])=True,False,True)

3430                  rstMasterAsset.Edit
3440                  rstMasterAsset![currentDate] = ![price_date]
3450                  rstMasterAsset![marketvalue] = ![price_close_mean]
3460                  If IsNull(![yield]) = False Then
3470                    If ![yield] <> 0# Then
3480                      If ![yield] < 1# Then
3490                        rstMasterAsset![yield] = Round(![yield], 6)
3500                        rstMasterAsset![yield_entry] = Round((Round(![yield], 6) * 100#), 4)
3510                      Else
3520                        rstMasterAsset![yield] = Round((Round(![yield], 4) / 100#), 6)
3530                        rstMasterAsset![yield_entry] = Round(![yield], 4)
3540                      End If
3550                    Else
3560                      rstMasterAsset![yield] = 0#
3570                      rstMasterAsset![yield_entry] = 0#
3580                    End If
3590                  Else
3600                    rstMasterAsset![yield] = 0#
3610                    rstMasterAsset![yield_entry] = 0#
3620                  End If
3630                  rstMasterAsset.Update

3640                  lngUpdateCount = lngUpdateCount + 1&

3650                End If

                    ' ** Update the progress bar.
3660                With frm
3670                  dblProgBar_Len = dblProgBar_Len + dblProgressAmount
3680                  .ProgBar_Width_Pric dblProgBar_Len, 2  ' ** Form Procedure: frmAssetPricing_Import.
                      '.ProgBar_bar.Width = dblProgBar_Len
3690                  intPct = CInt((dblProgBar_Len / dblProgBox_Width) * 100#)
3700                  .ProgBar_lbl.Caption = CStr(intPct) & "%"
3710                  DoEvents
3720                End With

3730                If lngX < lngSecCnt Then .MoveNext
3740              Next

3750            End If
3760          End With

3770          rstUpdated.Close
3780          rstMasterAsset.Close
3790          rstAppraise.Close

3800          If lngUpdateCount > 0& Then
3810            blnUpdateSuccess = True
3820          End If

              ' "ACCOUNT: TATEST                                   Portfolio Name: TRUST ACCOUNTANT TEST"
              ' "VALUATION DATE: Thursday, October 1 2009"
              ' "1) 369604103           1      GENERAL ELECTRIC CO            10/01    16.3900    15.9500    15.9700         16.39               5.0031"
              ' "DT 03/01/2008 5.0000% 02/15/2025"

3830        End If

3840      End Select

3850      dbs.Close

3860    End If  ' ** blnFound.

        ' ** Update underlying form.
3870    Forms("frmAssetPricing").Requery

        ' ** Update the progress bar.
3880    With frm
3890      dblProgBar_Len = dblProgBox_Width
3900      .ProgBar_Width_Pric dblProgBar_Len, 2  ' ** Form Procedure: frmAssetPricing_Import.
          '.ProgBar_bar.Width = dblProgBar_Len
3910      intPct = CInt((dblProgBar_Len / dblProgBox_Width) * 100#)
3920      .ProgBar_lbl.Caption = CStr(intPct) & "%"
3930      DoEvents
3940    End With

3950    DoCmd.Hourglass False
3960    If blnFound = True Then
3970      If lngUpdateCount > 0& Then
3980        blnUpdateSuccess = True
3990        Forms("frmAssetPricing").chkUpdated = True
4000        MsgBox CStr(lngUpdateCount) & " " & IIf(lngUpdateCount > 1&, "records were", "record was") & " updated, " & _
              "out of " & CStr(lngSecCnt) & vbCrLf & "listing" & IIf(lngSecCnt > 1&, "s ", " ") & "found in the file.", _
              vbInformation + vbOKOnly, ("Update Successful" & Space(40))
4010      Else
4020        If lngSecCnt > 0& Then
4030          strTmp01 = ", though " & CStr(lngSecCnt) & " security listing" & IIf(lngSecCnt > 1&, "s were ", " was ") & "found in the file."
4040        Else
4050          strTmp01 = "."
4060        End If
4070        Beep
4080        MsgBox "No record data was found and no Trust Accountant information was updated" & strTmp01, _
              vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
4090      End If
4100    Else
          ' ** blnFound will be False if:
          ' **   1. File not found or not chosen.
          ' **   2. No usable raw records found.
          ' **   3. None of the Cusips that were imported match a tblPricing_MasterAsset Cusip.
          ' **   4. While updating tblPricing_MasterAsset, Cusip not found that should have been found; shouldn't happen.
4110      If lngSecCnt > 0& Then
4120        strTmp01 = ", though " & CStr(lngSecCnt) & " security listing" & IIf(lngSecCnt > 1&, "s were ", " was ") & "found in the file."
4130      Else
4140        strTmp01 = "."
4150      End If
4160      Beep
4170      MsgBox "No Trust Accountant information was updated" & strTmp01, vbExclamation + vbOKOnly, "No Change Made"
4180    End If

        ' ** Delete the converted text file.
4190    If strPricePathFile_TXT_conv <> vbNullString Then
4200      If FileExists(strPricePathFile_TXT_conv) = True Then
4210        Kill strPricePathFile_TXT_conv
4220      End If
4230    End If

EXITP:
4240    Set rstAppraise = Nothing
4250    Set rstEstateVal = Nothing
4260    Set rstMasterAsset = Nothing
4270    Set rstUpdated = Nothing
4280    Set qdf = Nothing
4290    Set dbs = Nothing
4300    Exit Sub

ERRH:
4310    DoCmd.Hourglass False
4320    Select Case ERR.Number
        Case Else
4330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4340    End Select
4350    Resume EXITP

End Sub

Private Function GetData2(frm As Access.Form, strPricePathFile_TXT As String, dblProgBox_Width As Double, dblProgBar_Len As Double, intPct As Integer) As Variant
' ** From TrustAccountantPricing.mdb.
' ** Only called from GetData1(), above.

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "GetData2"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rstMasterAsset As DAO.Recordset, rstUpdated As DAO.Recordset
        Dim strInput As String
        Dim dblProgressAmount As Double
        Dim lngLines As Long, lngThisLine As Long
        Dim blnFoundSecNum As Boolean, blnFoundFileVer As Boolean, intFileVer As Integer
        Dim lngFoundSecNum As Long, lngFoundFileVer As Long
        Dim lngUpdateCount As Long, lngSecCnt As Long
        Dim blnFoundDate As Boolean, blnFoundCusip As Boolean, blnFoundBond As Boolean
        Dim lngFoundDate As Long, lngFoundCusip As Long
        Dim strCusip As String, strPriceDate As String
        Dim intShares As Integer
        Dim dblPrice As Double, dblIncome As Double
        Dim strLineTest As String, varRem As Variant  ' ** varRemainder may come back Null.
        Dim varTmp00 As Variant
        Dim arr_varRetVal(1) As Variant

        Const VER_OLD As Integer = 1
        Const VER_NEW As Integer = 2

4410    arr_varRetVal(0) = 0&
4420    arr_varRetVal(1) = 0&

4430    Set dbs = CurrentDb

4440    Set rstMasterAsset = dbs.OpenRecordset("tblPricing_MasterAsset", dbOpenDynaset)
4450    Set rstUpdated = dbs.OpenRecordset("tblPricing_Cusip", dbOpenDynaset)

4460  On Error Resume Next
4470    Open strPricePathFile_TXT For Input As #1
4480    If ERR.Number <> 0 Then
4490      Select Case ERR.Number
          Case 55  ' ** File already open.
4500        Close #1
4510  On Error GoTo ERRH
4520        Open strPricePathFile_TXT For Input As #1
4530      Case Else
4540        GoTo ERRH
4550  On Error GoTo ERRH
4560      End Select
4570    Else
4580  On Error GoTo ERRH
4590    End If

4600    lngLines = 0&: lngThisLine = 0&
4610    lngUpdateCount = 0&
4620    blnFoundSecNum = False: blnFoundFileVer = False: intFileVer = 0
4630    blnFoundCusip = False: blnFoundDate = False: blnFoundBond = False
4640    lngSecCnt = 0&
4650    lngFoundSecNum = 0&: lngFoundFileVer = 0&: lngFoundDate = 0&: lngFoundCusip = 0&

        ' ** Find out how many lines are in the file.
4660    Do While Not EOF(1)
4670      lngLines = lngLines + 1&
4680      Line Input #1, strInput
4690    Loop
4700    Close #1

4710    Open strPricePathFile_TXT For Input As #1

4720    dblProgressAmount = ((dblProgBox_Width - dblProgBar_Len) / lngLines)

4730    Do While Not EOF(1)

4740      Line Input #1, strInput
4750      lngThisLine = lngThisLine + 1&

          ' ** Trim trailing and leading spaces.
4760      strInput = Trim(strInput)

          ' ** Get whatever is first in the line up to the first space.
4770      strLineTest = GetFirstWord_AP(strInput, 0, " ")  ' ** Module Function: modPricingFuncs.

          ' ** Have we GOT the line with the number of securities?
4780      If blnFoundSecNum = False Then
4790        If strLineTest = "Number" Or InStr(strInput, "Number of Securities") > 0 Then
              '"                                                                                                   Number of Securities: 5"
              '"Processing Date: 09/17/2009                                                                        Number of Securities: 4"
4800          blnFoundSecNum = True
4810          lngSecCnt = Val(GetLastWord_AP(strInput, varRem, " "))
4820        End If
4830      End If

4840      If blnFoundSecNum = True Then

            ' *********************************************************************************************************************************
            ' ** VGC 10/04/2009: Ohana Fiduciary's test file has slightly different header and columns!
            ' *********************************************************************************************************************************
            ' ** ============================================================================================================================
            ' ** Ohana Fiduciary test file:
            ' ** ============================================================================================================================
            ' ** As-of Date:      12/01/2007                                                                      Estate of: Test Portfolio
            ' ** Valuation Date:  12/31/2008                                                                    Report Type: Appraisal Date
            ' ** Processing Date: 09/17/2009                                                                        Number of Securities: 4
            ' **                                                                                                           File ID: testing
            ' **
            ' **                 Shares      Security                                               Indicated   Security      Div and Int
            ' **     Identifier  or Par      Description                            Close           Income      Value         Accruals
            ' **     __________  ___________ __________________________             ___________     ___________ _____________ _____________
            ' **
            ' **    2) 1675923Q5           1 CHICAGO ILL O HARE INTL ARPT R 3RD LIEN
            ' **                             Financial Times Interactive Data
            ' **                             DTD: 01/31/2008 Mat: 01/01/2020 5%
            ' **                             12/31/2008                                98.67000            0.05          0.99
            ' **
            ' ** ============================================================================================================================
            ' ** My original test file:
            ' ** ============================================================================================================================
            ' ** As-of Date:      08/25/2008                                                                          Estate of: Delta Data
            ' ** Valuation Date:  07/31/2008                                                                             Account: tapricing
            ' ** Processing Date: 08/26/2008                                                                    Report Type: Appraisal Date
            ' **                                                                                                    Number of Securities: 5
            ' **                                                                                                         File ID: tapricing
            ' **
            ' **       Shares      Security                                                         Indicated                 Security
            ' **       or Par      Description                                      Close           Income                    Value
            ' **       ___________ ____________________________________             ___________     ___________               _____________
            ' **
            ' **    1)           1 AT&T CORP (001957109)
            ' **                   COM NEW
            ' **                   New York Stock Exchange
            ' **                   11/18/2005                                          20.35000                                         N/A
            ' **                   Last price available on 11/18/2005
            ' **
            ' *********************************************************************************************************************************

            ' ** Determine which version of the file we've got.
4850        If blnFoundFileVer = False Then
4860          If strLineTest = "Identifier" Or strLineTest = "or" Then
4870            Select Case strLineTest
                Case "Identifier"
4880              blnFoundFileVer = True
4890              intFileVer = VER_NEW
4900            Case "Or"
4910              If Left(strInput, 6) = "or Par" Then
4920                blnFoundFileVer = True
4930                intFileVer = VER_OLD
4940              End If
4950            End Select
4960          End If
4970        End If  ' ** blnFoundFileVer.

4980        If blnFoundFileVer = True Then

              ' ** If we have #), then we have line on the report that contains the cusip.
4990          If blnFoundCusip = False Then

5000            strLineTest = GetFirstWord_AP(strInput, varRem, " ")  ' ** Module Function: modPricingFuncs.
5010            If Right(strLineTest, 1) = ")" Then

                  ' ** Find the cusip.
5020              Select Case intFileVer
                  Case VER_NEW
                    '"   1) 20030NAA9           1 COMCAST CORP NEW"

5030                strCusip = GetFirstWord_AP(varRem, varRem, " ")  ' ** Module Function: modPricingFuncs.
5040                If Len(strCusip) = 9 Then

5050                  blnFoundCusip = True
                      ' ** Find the number of shares just in case it's not 1.
                      ' ** Do thrice to get by the 2nd set of characters.
5060                  varTmp00 = Val(GetFirstWord_AP(strInput, varRem, " "))  ' ** Module Function: modPricingFuncs.
5070                  varTmp00 = Val(GetFirstWord_AP(varRem, varRem, " "))  ' ** Module Function: modPricingFuncs.
5080                  intShares = Val(GetFirstWord_AP(varRem, 0, " "))  ' ** Module Function: modPricingFuncs.
                      '"   1) 20030NAA9           1 COMCAST CORP NEW"

5090                End If

5100              Case VER_OLD
                    '"   1)           1 GENERAL ELECTRIC CO (369604103)"

5110                strCusip = GetLastWord_AP(strInput, varRem, " ")  ' ** Module Function: modPricingFuncs.
                    ' ** Cusip will be "(#########)".
5120                If Len(strCusip) = 11 Then

                      ' ** Remove the last and first characters.
5130                  strCusip = Left(strCusip, 10)
5140                  strCusip = Right(strCusip, 9)
5150                  If Len(strCusip) = 9 Then

5160                    blnFoundCusip = True
                        ' ** Find the number of shares just in case it's not 1.
                        ' ** Do twice to get by the 1st set of characters.
5170                    varTmp00 = Val(GetFirstWord_AP(strInput, varRem, " "))  ' ** Module Function: modPricingFuncs.
5180                    intShares = Val(GetFirstWord_AP(varRem, 0, " "))  ' ** Module Function: modPricingFuncs.
                        '"   1)           1 GENERAL ELECTRIC CO (369604103)"

5190                  End If
5200                Else
                      ' ** Houston we have problem. We think we have a cusip but the string is longer than expected.
5210                End If

5220              End Select

5230            End If  ' ** Numbered line.
5240          End If  ' ** blnFoundCusip.

              ' ** Have we GOT the line with the DTD: string indicating a bond?
              ' ** This gets reset for each line during the DO loop.
5250          If strLineTest = "DTD:" Then
5260            blnFoundBond = True  '** Yep, set the flag.
5270          End If

5280          If blnFoundCusip = True Then 'And Len(strCusip) = 9 Then

                ' ** If we find the string in the format mm/dd/yyyy then we have a date.
5290            If blnFoundDate = False Then
5300              If IsDate(Left(strInput, 10)) Then  ' ** Remember, lines have been trimmed of leading and trailing spaces.

5310                blnFoundDate = True
5320                strPriceDate = GetFirstWord_AP(strInput, varRem, " ")  ' ** Module Function: modPricingFuncs.
5330                dblPrice = GetFirstWord_AP(varRem, varRem, " ")  ' ** Should be the close value.  ' ** Module Function: modPricingFuncs.
                    ' ** If it is a bond, divide by 100 to get PAR value, otherwise leave it as is.
5340                If blnFoundBond Then
5350                  dblPrice = dblPrice / 100
5360                End If
5370                Select Case intFileVer
                    Case VER_NEW
                      ' ** If new style, my only example has both 'Indicated Income' and 'Security Value'.
                      ' ** Would it ever be WITHOUT an 'Indicated Income', but WITH a 'Security Value'?
                      ' **   "12/31/2008                                98.67000            0.05          0.99"
                      ' ** And what about 'Div and Int Accruals'?
5380                  varRem = vbNullString

5390                  If Len(strInput) > 51 Then

                        ' ** 'Indicated income' should start at character: 51.
5400                    If Mid(strInput, 51, 1) = " " Then
                          ' ** Check backwards to confirm it's the first space between values.
5410                      If Mid(strInput, 50, 1) = " " Then
5420                        If Mid(strInput, 49, 1) = " " Then
5430                          If Mid(strInput, 48, 1) = " " Then
                                ' ** What gives? If there was no price, dblPrice is either zero
                                ' ** (if no values on the line), or has picked up one of the other columns!
                                ' ** Let it be, and a user will have to send us an example.
5440                          Else
5450                            varRem = Mid(strInput, 49)
5460                          End If
5470                        Else
5480                          varRem = Mid(strInput, 50)
5490                        End If
5500                      Else
5510                        varRem = Mid(strInput, 51)
5520                      End If
5530                    Else
                          ' ** Move forward to find the first space between values.
5540                      If Mid(strInput, 52, 1) = " " Then
5550                        varRem = Mid(strInput, 52)
5560                      Else
5570                        If Mid(strInput, 53, 1) = " " Then
5580                          varRem = Mid(strInput, 53)
5590                        Else
                              ' ** Oh, this could get ridiculous.
                              ' ** Let it be, and a user will have to send us an example.
5600                        End If
5610                      End If
5620                    End If

5630                    If Len(varRem) > 17 Then
5640                      If Mid(varRem, 17, 1) = " " Then
5650                        dblIncome = Val(Left(varRem, 17))
5660                      Else
5670                        If Mid(varRem, 18, 1) = " " Then
5680                          dblIncome = Val(Left(varRem, 18))
5690                        Else
                              ' ** Just see what we get!
5700                          dblIncome = Val(varRem)
5710                        End If
5720                      End If
5730                    End If

                        ' ** 'Security Value' should start at character: 67 (17 of varRem).
                        ' ** 'Div and Int Accruals' should start at character: 81 (15 of varRem).
5740                  End If  ' ** Length > 51.

5750                Case VER_OLD
                      ' ** If old style, with no 'Indicated Income', and a 'Security Value' of 'N/A', dblIncome will still come back zero.
                      ' **   "11/30/2006                                           2.55000                                         N/A"
5760                  dblIncome = Val(GetFirstWord_AP(varRem, varRem, " "))  ' ** Module Function: modPricingFuncs.
5770                End Select

5780              End If  ' ** First word IsDate().
5790            End If  ' ** blnFoundDate.

5800          End If  ' ** blnFoundCusip.

              ' ** If we think that we have all the pieces, let's process them.
5810          If blnFoundCusip And blnFoundDate Then

                ' ** Find the cusip in the master asset table.
5820            rstMasterAsset.FindFirst "[cusip] = '" & Format(strCusip, "000000000") & "'"  ' ** Yes, zeroes will still show letters.

5830            If rstMasterAsset.NoMatch Then
                  ' ** If there is no match from the tblPricing_MasterAsset, we don't care about this one.

5840              blnFoundCusip = False
5850              blnFoundDate = False
5860              blnFoundBond = False

5870            Else
                  ' ** Found a match, so upate the tblPricing_MasterAsset.

5880              rstUpdated.AddNew
5890              rstUpdated![assetno] = rstMasterAsset![assetno]
5900              rstUpdated![cusip] = Format(strCusip, "000000000")
5910              rstUpdated![OldValue] = rstMasterAsset![marketvalue]
5920              rstUpdated![OldDate] = rstMasterAsset![currentDate]
                  'rstUpdated![NewValue] = Format(dblPrice, "####.###")
5930              rstUpdated![NewValue] = dblPrice
5940              rstUpdated![NewDate] = strPriceDate
                  ' ** Yield should be income divided by shares (shares should be 1 but just in case) divided by close price.
                  'rstUpdated![NewYield] = Format(((dblIncome / intShares) / dblPrice), "####.####")
5950              rstUpdated![NewYield] = Round(((dblIncome / intShares) / dblPrice), 6)
5960              rstUpdated.Update
                  ' ** yield_upd: IIf(IsNull([tblPricing_Cusip].[assetno])=True,False,True)
                  ' ** marketvalue_upd: IIf(IsNull([tblPricing_Cusip].[assetno])=True,False,True)

5970              rstMasterAsset.Edit
5980              rstMasterAsset![currentDate] = strPriceDate
                  'rstMasterAsset![marketvalue] = Format(dblPrice, "####.###")
5990              rstMasterAsset![marketvalue] = dblPrice
                  ' ** Yield should be income divided by shares (shares should be 1 but just in case) divided by close price.
                  'rstMasterAsset![yield] = Format(((dblIncome / intShares) / dblPrice), "####.####")
6000              If ((dblIncome / intShares) / dblPrice) <> 0# Then
6010                If ((dblIncome / intShares) / dblPrice) < 1# Then
6020                  rstMasterAsset![yield] = Round(((dblIncome / intShares) / dblPrice), 6)
6030                  rstMasterAsset![yield_entry] = Round((Round(((dblIncome / intShares) / dblPrice), 6) * 100#), 4)
6040                Else
6050                  rstMasterAsset![yield] = Round((Round(((dblIncome / intShares) / dblPrice), 4) / 100#), 6)
6060                  rstMasterAsset![yield_entry] = Round(((dblIncome / intShares) / dblPrice), 4)
6070                End If
6080              Else
6090                rstMasterAsset![yield] = 0#
6100                rstMasterAsset![yield_entry] = 0#
6110              End If
6120              rstMasterAsset![AssetPricing_Changed] = True
6130              rstMasterAsset.Update

6140              lngUpdateCount = lngUpdateCount + 1  ' ** Update records updated count.

6150              blnFoundCusip = False
6160              blnFoundDate = False
6170              blnFoundBond = False

6180            End If

                ' ** Update the progress bar.
6190            With frm
6200              dblProgBar_Len = dblProgBar_Len + dblProgressAmount
6210              .ProgBar_Width_Pric dblProgBar_Len, 2  ' ** Form Procedure: frmAssetPricing_Import.
                  '.ProgBar_bar.Width = dblProgBar_Len
6220              intPct = CInt((dblProgBar_Len / dblProgBox_Width) * 100#)
6230              .ProgBar_lbl.Caption = CStr(intPct) & "%"
6240              DoEvents
6250            End With

6260          End If  ' ** blnFoundCusip, blnFoundDate.

6270        End If  ' ** blnFoundFileVer.

6280      End If  ' ** blnFoundSecNum.

6290    Loop  ' ** While Not EOF.

6300    DoEvents  ' ** Let all the screen updates catch up.

        ' ** Close the file.
6310    Close #1
6320    rstUpdated.Close
6330    rstMasterAsset.Close
6340    dbs.Close

6350    arr_varRetVal(0) = lngUpdateCount
6360    arr_varRetVal(1) = lngSecCnt

EXITP:
6370    Set rstUpdated = Nothing
6380    Set rstMasterAsset = Nothing
6390    Set qdf = Nothing
6400    Set dbs = Nothing
6410    GetData2 = arr_varRetVal
6420    Exit Function

ERRH:
6430    arr_varRetVal(0) = -9&
6440    arr_varRetVal(1) = Erl
6450    Select Case ERR.Number
        Case Else
6460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6470    End Select
6480    Resume EXITP

End Function

Public Sub CreateNewDirectory(strDirName As String)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "CreateNewDirectory"

        Dim intNewLen As Integer
        Dim intDirLen As Integer
        Dim intMaxLen As Integer

6510    intNewLen = 4
6520    intMaxLen = Len(strDirName)

6530    If Right(strDirName, 1) <> LNK_SEP Then
6540      strDirName = strDirName + LNK_SEP
6550      intMaxLen = intMaxLen + 1
6560    End If

6570    Do While True
6580      intDirLen = InStr(intNewLen, strDirName, LNK_SEP)
6590      MkDir Left(strDirName, intDirLen - 1)
6600      intNewLen = intDirLen + 1
6610      If intNewLen >= intMaxLen Then
6620        Exit Do
6630      End If
6640    Loop

EXITP:
6650    Exit Sub

ERRH:
6660    Select Case ERR.Number
        Case Else
6670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6680    End Select
6690    Resume EXITP

End Sub

Public Function GetFirstWord_AP(varStr As Variant, varRemainder As Variant, Optional strDelimiter As String = " ") As String
' ** strRetVal: returns the first word in varStr.
' ** varRemainder: returns the rest.
' ** Words are delimited by strDelimiter.
' ** SEE ALSO: GetLastWord in modStringFuncs.

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "GetFirstWord_AP"

        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim strRetVal As String

6710    strRetVal = vbNullString

6720    strTmp01 = Trim(varStr)
6730    intPos01 = InStr(strTmp01, strDelimiter)
6740    If intPos01 = 0 Then
6750      strRetVal = strTmp01
6760      varRemainder = intPos01
6770    Else
6780      strRetVal = Left(strTmp01, intPos01 - 1)
6790      varRemainder = Mid(strTmp01, intPos01 + 1)
6800    End If

EXITP:
6810    GetFirstWord_AP = strRetVal
6820    Exit Function

ERRH:
6830    Select Case ERR.Number
        Case Else
6840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6850    End Select
6860    Resume EXITP

End Function

Public Function GetLastWord_AP(varStr As Variant, varRemainder As Variant, Optional strDelimiter As String = " ") As String
' ** strRetVal: returns the last word in varStr.
' ** varRemainder: returns the rest.
' ** Words are delimited by strDelimiter.
' ** Function works from the end and works backward.
' ** SEE ALSO: GetLastWord in modStringFuncs.

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "GetLastWord_AP"

        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim intX As Integer
        Dim strRetVal As String

6910    strRetVal = vbNullString

6920    strTmp01 = Trim(varStr)
6930    intPos01 = 1
6940    For intX = Len(strTmp01) To 1 Step -1
6950      If Mid(strTmp01, intX, 1) = strDelimiter Then
6960        intPos01 = intX + 1
6970        Exit For
6980      End If
6990    Next
7000    If intPos01 = 1 Then
7010      strRetVal = strTmp01
7020      varRemainder = Null
7030    Else
7040      strRetVal = Mid(strTmp01, intPos01)
7050      varRemainder = Trim(Left(strTmp01, intPos01 - 1))
7060    End If

EXITP:
7070    GetLastWord_AP = strRetVal
7080    Exit Function

ERRH:
7090    Select Case ERR.Number
        Case Else
7100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7110    End Select
7120    Resume EXITP

End Function
