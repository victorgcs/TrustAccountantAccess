Attribute VB_Name = "modGoToReportFuncs"
Option Compare Database
Option Explicit

'VGC 10/28/2017: CHANGES!

Private Const THIS_NAME As String = "modGoToReportFuncs"
' **

Public Function GetDivAsset() As Long
' ** Return an assetno with dividends.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "GetDivAsset"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRetVal As Long

110     lngRetVal = 0&

120     Set dbs = CurrentDb
130     With dbs
          ' ** qryReport_List_02a (qryReport_List_01a (MasterAsset, linked to
          ' ** AssetType, just 'Dividend' types), linked to ActiveAssets, grouped
          ' ** by assetno, accountno, with cnt), grouped by assetno, with cnt.
140       Set qdf = .QueryDefs("qryReport_List_03a")
150       Set rst = qdf.OpenRecordset
160       With rst
170         If .BOF = True And .EOF = True Then
              ' ** None found.
180         Else
190           .MoveFirst
200           lngRetVal = ![assetno]  ' ** Just take the first one, they're sorted.
210         End If
220         .Close
230       End With
240       .Close
250     End With

EXITP:
260     Set rst = Nothing
270     Set qdf = Nothing
280     Set dbs = Nothing
290     GetDivAsset = lngRetVal
300     Exit Function

ERRH:
310     lngRetVal = 0&
320     Select Case ERR.Number
        Case Else
330       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
340     End Select
350     Resume EXITP

End Function

Public Function GetIntAsset() As Long
' ** Return an assetno with interest.

400   On Error GoTo ERRH

        Const THIS_PROC As String = "GetIntAsset"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRetVal As Long

410     lngRetVal = 0&

420     Set dbs = CurrentDb
430     With dbs
          ' ** qryReport_List_02b (qryReport_List_01b (MasterAsset, linked to
          ' ** AssetType, just 'Interest' types), linked to ActiveAssets, grouped
          ' ** by assetno, accountno, with cnt), grouped by assetno, with cnt.
440       Set qdf = .QueryDefs("qryReport_List_03b")
450       Set rst = qdf.OpenRecordset
460       With rst
470         If .BOF = True And .EOF = True Then
              ' ** None found.
480         Else
490           .MoveFirst
500           lngRetVal = ![assetno]  ' ** Just take the first one, they're sorted.
510         End If
520         .Close
530       End With
540       .Close
550     End With

EXITP:
560     Set rst = Nothing
570     Set qdf = Nothing
580     Set dbs = Nothing
590     GetIntAsset = lngRetVal
600     Exit Function

ERRH:
610     lngRetVal = 0&
620     Select Case ERR.Number
        Case Else
630       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
640     End Select
650     Resume EXITP

End Function

Public Function GetSoldAsset() As String
' ** Return an accountno and assetno that can be sold.

700   On Error GoTo ERRH

        Const THIS_PROC As String = "GetSoldAsset"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strRetVal As String

710     strRetVal = vbNullString

720     Set dbs = CurrentDb
730     With dbs
          ' ** qryReport_List_02c (qryReport_List_01c (ActiveAssets, grouped by accountno, assetno,
          ' ** with cnt), grouped by accountno, with Max(cnt)), linked back toqryReport_List_01c
          ' ** (ActiveAssets, grouped by accountno, assetno, with cnt), with assetno.
740       Set qdf = .QueryDefs("qryReport_List_03c")
750       Set rst = qdf.OpenRecordset
760       With rst
770         If .BOF = True And .EOF = True Then
              ' ** None found.
780         Else
790           .MoveFirst
800           strRetVal = ![accountno] & ";" & CStr(![assetno])  ' ** Just take the first one, they're sorted.
810         End If
820         .Close
830       End With
840       .Close
850     End With

EXITP:
860     Set rst = Nothing
870     Set qdf = Nothing
880     Set dbs = Nothing
890     GetSoldAsset = strRetVal
900     Exit Function

ERRH:
910     strRetVal = vbNullString
920     Select Case ERR.Number
        Case Else
930       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
940     End Select
950     Resume EXITP

End Function

Public Function GetHiddenTrans() As String
' ** Return an accountno with hidden transactions.

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "GetHiddenTrans"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strRetVal As String

1010    strRetVal = vbNullString

1020    Set dbs = CurrentDb
1030    With dbs
          ' ** qryReport_List_04a (Ledger, just ledger_HIDDEN = True), grouped by accountno, with cnt.
1040      Set qdf = .QueryDefs("qryReport_List_05a")
1050      Set rst = qdf.OpenRecordset
1060      With rst
1070        If .BOF = True And .EOF = True Then
              ' ** No hidden transactions.
1080        Else
1090          .MoveFirst
1100          strRetVal = ![accountno]
1110        End If
1120        .Close
1130      End With
1140      .Close
1150    End With

EXITP:
1160    Set rst = Nothing
1170    Set qdf = Nothing
1180    Set dbs = Nothing
1190    GetHiddenTrans = strRetVal
1200    Exit Function

ERRH:
1210    strRetVal = vbNullString
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Function

Public Function GetRelAcct() As String
' ** Return an accountno with related accounts.

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "GetRelAcct"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strRetVal As String

1310    strRetVal = vbNullString

1320    Set dbs = CurrentDb
1330    With dbs
          ' ** Account, just related_accountno <> Null, with cnt.
1340      Set qdf = .QueryDefs("qryReport_List_04b")
1350      Set rst = qdf.OpenRecordset
1360      With rst
1370        If .BOF = True And .EOF = True Then
              ' ** No related accounts.
1380        Else
1390          .MoveFirst
1400          strRetVal = ![cnt]
1410        End If
1420        .Close
1430      End With
1440      .Close
1450    End With

EXITP:
1460    Set rst = Nothing
1470    Set qdf = Nothing
1480    Set dbs = Nothing
1490    GetRelAcct = strRetVal
1500    Exit Function

ERRH:
1510    strRetVal = vbNullString
1520    Select Case ERR.Number
        Case Else
1530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1540    End Select
1550    Resume EXITP

End Function

Public Function GetTransCnt() As Long
' ** Return number of transactions in the Ledger.

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "GetTransCnt"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRetVal As Long

1610    lngRetVal = 0&

1620    Set dbs = CurrentDb
1630    With dbs
          ' ** Ledger, grouped, with cnt.
1640      Set qdf = .QueryDefs("qryReport_List_04c")
1650      Set rst = qdf.OpenRecordset
1660      With rst
1670        If .BOF = True And .EOF = True Then
              ' ** No transactions.
1680        Else
1690          .MoveFirst
1700          lngRetVal = ![cnt]
1710        End If
1720        .Close
1730      End With
1740      .Close
1750    End With

EXITP:
1760    Set rst = Nothing
1770    Set qdf = Nothing
1780    Set dbs = Nothing
1790    GetTransCnt = lngRetVal
1800    Exit Function

ERRH:
1810    lngRetVal = 0&
1820    Select Case ERR.Number
        Case Else
1830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1840    End Select
1850    Resume EXITP

End Function

Public Function GetAcctCnt() As Long
' ** Return number of accounts.

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "GetAcctCnt"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRetVal As Long

1910    lngRetVal = 0&

1920    Set dbs = CurrentDb
1930    With dbs
          ' ** Account, grouped, with cnt.
1940      Set qdf = .QueryDefs("qryReport_List_04d")
1950      Set rst = qdf.OpenRecordset
1960      With rst
1970        If .BOF = True And .EOF = True Then
              ' ** No transactions.
1980        Else
1990          .MoveFirst
2000          lngRetVal = ![cnt]
2010          lngRetVal = (lngRetVal - 2&)  ' ** Skip the 2 system accounts.
2020        End If
2030        .Close
2040      End With
2050      .Close
2060    End With

EXITP:
2070    Set rst = Nothing
2080    Set qdf = Nothing
2090    Set dbs = Nothing
2100    GetAcctCnt = lngRetVal
2110    Exit Function

ERRH:
2120    lngRetVal = 0&
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Function

Public Function GetJrnlColID(strFilter As String) As Long
' ** Return an empty tblJournal_Column record ID.

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "GetJrnlColID"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRetVal As Long

2210    lngRetVal = 0&

2220    Set dbs = CurrentDb
2230    With dbs
2240      If InStr(strFilter, "journal_USER") > 0 Then
            ' ** qryReport_List_01d (tblJournal_Column, with IsEmpty, by specified [usr]), just IsEmpty = True.
2250        Set qdf = .QueryDefs("qryReport_List_02d")
2260        With qdf.Parameters
2270          ![usr] = gstrJournalUser
2280        End With
2290      Else
            ' ** qryReport_List_01e (tblJournal_Column, with IsEmpty), just IsEmpty = True.
2300        Set qdf = .QueryDefs("qryReport_List_02e")
2310      End If
2320      Set rst = qdf.OpenRecordset
2330      With rst
2340        If .BOF = True And .EOF = True Then
              ' ** No new records!
2350        Else
2360          .MoveFirst
2370          lngRetVal = ![JrnlCol_ID]
2380        End If
2390      End With
2400      .Close
2410    End With

EXITP:
2420    Set rst = Nothing
2430    Set qdf = Nothing
2440    Set dbs = Nothing
2450    GetJrnlColID = lngRetVal
2460    Exit Function

ERRH:
2470    lngRetVal = 0&
2480    Select Case ERR.Number
        Case Else
2490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2500    End Select
2510    Resume EXITP

End Function

Public Function GTRImageChk() As Boolean

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "GTRImageChk"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
        Dim lngCtls As Long, arr_varCtl As Variant
        Dim strModName As String, strLine As String
        Dim lngLines As Long, lngDecLines As Long
        Dim lngHits As Long
        Dim intPos01 As Integer
        Dim lngX As Long, lngY As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varCtl().
        'Const C_DID   As Integer = 0
        'Const C_FID   As Integer = 1
        'Const C_FNAM  As Integer = 2
        'Const C_CID   As Integer = 3
        Const C_CNAM  As Integer = 4
        'Const C_TYP   As Integer = 5
        'Const C_CONST As Integer = 6

2610    blnRetVal = True

2620    Set dbs = CurrentDb
2630    With dbs
          ' ** zz_qry_Report_VBComponent_03 (zz_qry_Report_VBComponent_02 (zz_qry_Report_VBComponent_01
          ' ** (tblForm_Control, just 'GoToReport' controls), just non-Emblem controls), just image controls),
          ' ** just non-standardized names.
2640      Set qdf = .QueryDefs("zz_qry_Report_VBComponent_04")
2650      Set rst = qdf.OpenRecordset
2660      With rst
2670        .MoveLast
2680        lngCtls = .RecordCount
2690        .MoveFirst
2700        arr_varCtl = .GetRows(lngCtls)
            ' *****************************************************
            ' ** Array: arr_varCtl()
            ' **
            ' **   Field  Element  Name                Constant
            ' **   =====  =======  ==================  ==========
            ' **     1       0     dbs_id              C_DID
            ' **     2       1     frm_id              C_FID
            ' **     3       2     frm_name            C_FNAM
            ' **     4       3     ctl_id              C_CID
            ' **     5       4     ctl_name            C_CNAM
            ' **     6       5     ctltype_type        C_TYP
            ' **     7       6     ctltype_constant    C_CONST
            ' **
            ' *****************************************************
2710        .Close
2720      End With
2730      .Close
2740    End With

2750    lngHits = 0&
2760    Set vbp = Application.VBE.ActiveVBProject
2770    With vbp
2780      For Each vbc In .VBComponents
2790        With vbc
2800          strModName = .Name
2810          If strModName <> THIS_NAME Then  'OK!
2820            Set cod = .CodeModule
2830            With cod
2840              lngLines = .CountOfLines
2850              lngDecLines = .CountOfDeclarationLines
2860              For lngX = 0& To (lngCtls - 1&)
2870                For lngY = 1& To lngLines
2880                  strLine = .Lines(lngY, 1)
2890                  strLine = Trim(strLine)
2900                  If strLine <> vbNullString Then
2910                    intPos01 = InStr(strLine, arr_varCtl(C_CNAM, lngX))
2920                    If intPos01 > 0 Then
2930                      lngHits = lngHits + 1&
2940                      Debug.Print "'MOD: " & strModName & "  CTL: " & arr_varCtl(C_CNAM, lngX) & "  LINE: " & CStr(lngY)
2950                      DoEvents
2960                    End If
2970                  End If  ' ** vbNullString.
2980                Next  ' ** lngY.
2990              Next  ' ** lngX.
3000            End With  ' ** cod.
3010          End If  ' ** THIS_NAME.
3020        End With  ' ** vbc.
3030      Next  ' ** vbc.
3040    End With  ' ** vbp.

3050    If lngHits > 0& Then
3060      Debug.Print "'DONE!  " & THIS_PROC & "()  " & CStr(lngHits)
3070    Else
3080      Debug.Print "'NONE FOUND!  " & THIS_PROC & "()"
3090      Beep
3100    End If

EXITP:
3110    Set cod = Nothing
3120    Set vbc = Nothing
3130    Set vbp = Nothing
3140    Set rst = Nothing
3150    Set qdf = Nothing
3160    Set dbs = Nothing
3170    GTRImageChk = blnRetVal
3180    Exit Function

ERRH:
3190    blnRetVal = False
3200    Select Case ERR.Number
        Case Else
3210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3220    End Select
3230    Resume EXITP

End Function

Public Function Qry_XAdminGfx() As Boolean
' ** Name the graphics fields in a form query,
' ** including partially named fields.

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_XAdminGfx"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strQryName As String, strSQL As String
        Dim lngFlds As Long, arr_varFld() As Variant
        Dim intPos01 As Integer, intLen As Integer
        Dim arr_varTmp00 As Variant, strTmp01 As String, strTmp02 As String, lngTmp03 As Long
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varFld().
        Const F_ELEMS As Integer = 7  ' ** Array's first-element UBound().
        Const F_QNAM As Integer = 0
        Const F_FLDS As Integer = 1
        Const F_FNAM As Integer = 2
        Const F_FTYP As Integer = 3
        Const F_CNAM As Integer = 4  'T/F
        Const F_IMG  As Integer = 5  'T/F
        Const F_VAL  As Integer = 6
        Const F_FND  As Integer = 7

3310  On Error GoTo 0

3320    blnRetVal = True

3330    strQryName = "qryCurrency_Rate_01"

3340    lngFlds = 0&
3350    ReDim arr_varFld(F_ELEMS, 0)

3360    arr_varTmp00 = Qry_FldList_rel(strQryName)  ' ** Module Function: modQueryFunctions1.

3370    lngTmp03 = UBound(arr_varTmp00, 2) + 1&
3380    For lngX = 0& To (lngTmp03 - 1&)
3390      lngFlds = lngFlds + 1&
3400      lngE = lngFlds - 1&
3410      ReDim Preserve arr_varFld(F_ELEMS, lngE)
3420      arr_varFld(F_QNAM, lngE) = strQryName
3430      arr_varFld(F_FLDS, lngE) = lngTmp03
3440      arr_varFld(F_FNAM, lngE) = arr_varTmp00(0, lngX)
3450      arr_varFld(F_FTYP, lngE) = Null
3460      arr_varFld(F_CNAM, lngE) = CBool(False)
3470      arr_varFld(F_IMG, lngE) = CBool(False)
3480      arr_varFld(F_VAL, lngE) = Null
3490      arr_varFld(F_FND, lngE) = CBool(False)
3500    Next

3510    Set dbs = CurrentDb
3520    With dbs

3530      Set qdf = .QueryDefs(strQryName)
3540      With qdf
3550        strSQL = .SQL
3560        For lngX = 0& To (lngFlds - 1&)
3570          arr_varFld(F_FTYP, lngX) = .Fields(lngX).Type
3580        Next
3590        Set rst = .OpenRecordset
3600        With rst
3610          .MoveFirst
3620          For lngX = 0& To (lngFlds - 1&)
3630            If Left(arr_varFld(F_FNAM, lngX), 8) = "ctl_name" Then
3640              arr_varFld(F_CNAM, lngX) = CBool(True)
3650              arr_varFld(F_VAL, lngX) = .Fields(arr_varFld(F_FNAM, lngX))
3660            ElseIf Left(arr_varFld(F_FNAM, lngX), 12) = "xadgfx_image" Then
3670              arr_varFld(F_IMG, lngX) = CBool(True)
3680            ElseIf arr_varFld(F_FTYP, lngX) = dbLongBinary Then
                  ' ** This means the field has already been named.
3690              arr_varFld(F_IMG, lngX) = CBool(True)
3700              arr_varFld(F_FND, lngX) = CBool(True)
3710            End If
3720          Next
3730          .Close
3740        End With
3750        Set rst = Nothing

            ' ** SELECT tblForm_Graphics.dbs_id, tblDatabase.dbs_name, tblForm_Graphics.frm_id, tblForm.frm_name, tblForm_Graphics.frmgfx_id,
            ' **   tblForm_Graphics.ctl_name_01, tblForm_Graphics.xadgfx_image_01 AS cmdResetFilter_raised_img,
            ' **   tblForm_Graphics.ctl_name_02, tblForm_Graphics.xadgfx_image_02 AS cmdResetFilter_raised_semifocus_dots_img,
            ' **   tblForm_Graphics.ctl_name_03, tblForm_Graphics.xadgfx_image_03 AS cmdResetFilter_raised_focus_img,
            ' **   tblForm_Graphics.ctl_name_04, tblForm_Graphics.xadgfx_image_04 AS cmdResetFilter_raised_focus_dots_img,
            ' **   tblForm_Graphics.ctl_name_05, tblForm_Graphics.xadgfx_image_05 AS cmdResetFilter_sunken_focus_dots_img,
            ' **   tblForm_Graphics.ctl_name_06, tblForm_Graphics.xadgfx_image_06 AS cmdResetFilter_raised_img_dis,
            ' **   tblForm_Graphics.ctl_name_07, tblForm_Graphics.xadgfx_image_07 AS cmdCountries_raised_img,
            ' **   tblForm_Graphics.ctl_name_08, tblForm_Graphics.xadgfx_image_08 AS cmdCountries_raised_semifocus_dots_img,
            ' **   tblForm_Graphics.ctl_name_09, tblForm_Graphics.xadgfx_image_09 AS cmdCountries_raised_focus_img,
            ' **   tblForm_Graphics.ctl_name_10, tblForm_Graphics.xadgfx_image_10 AS cmdCountries_raised_focus_dots_img,
            ' **   tblForm_Graphics.ctl_name_11, tblForm_Graphics.xadgfx_image_11 AS cmdCountries_sunken_focus_dots_img,
            ' **   tblForm_Graphics.ctl_name_12, tblForm_Graphics.xadgfx_image_12 AS cmdCountries_raised_img_dis,
            ' **   tblForm_Graphics.ctl_name_13, tblForm_Graphics.xadgfx_image_13, tblForm_Graphics.ctl_name_14, tblForm_Graphics.xadgfx_image_14,
            ' **   tblForm_Graphics.ctl_name_15, tblForm_Graphics.xadgfx_image_15, tblForm_Graphics.ctl_name_16, tblForm_Graphics.xadgfx_image_16,
            ' **   tblForm_Graphics.ctl_name_17, tblForm_Graphics.xadgfx_image_17, tblForm_Graphics.ctl_name_18, tblForm_Graphics.xadgfx_image_18,
            ' **   tblForm_Graphics.ctl_name_19, tblForm_Graphics.xadgfx_image_19, tblForm_Graphics.ctl_name_20, tblForm_Graphics.xadgfx_image_20,
            ' **   tblForm_Graphics.ctl_name_21, tblForm_Graphics.xadgfx_image_21, tblForm_Graphics.ctl_name_22, tblForm_Graphics.xadgfx_image_22,
            ' **   tblForm_Graphics.ctl_name_23, tblForm_Graphics.xadgfx_image_23, tblForm_Graphics.ctl_name_24, tblForm_Graphics.xadgfx_image_24,
            ' **   tblForm_Graphics.ctl_name_25, tblForm_Graphics.xadgfx_image_25, tblForm_Graphics.ctl_name_26, tblForm_Graphics.xadgfx_image_26,
            ' **   tblForm_Graphics.ctl_name_27, tblForm_Graphics.xadgfx_image_27, tblForm_Graphics.ctl_name_28, tblForm_Graphics.xadgfx_image_28,
            ' **   tblForm_Graphics.ctl_name_29, tblForm_Graphics.xadgfx_image_29, tblForm_Graphics.ctl_name_30, tblForm_Graphics.xadgfx_image_30,
            ' **   tblForm_Graphics.ctl_name_31, tblForm_Graphics.xadgfx_image_31, tblForm_Graphics.ctl_name_32, tblForm_Graphics.xadgfx_image_32,
            ' **   tblForm_Graphics.ctl_name_33, tblForm_Graphics.xadgfx_image_33, tblForm_Graphics.ctl_name_34, tblForm_Graphics.xadgfx_image_34,
            ' **   tblForm_Graphics.ctl_name_35, tblForm_Graphics.xadgfx_image_35, tblForm_Graphics.ctl_name_36, tblForm_Graphics.xadgfx_image_36,
            ' **   tblForm_Graphics.ctl_name_37, tblForm_Graphics.xadgfx_image_37
            ' ** FROM tblDatabase INNER JOIN (tblForm INNER JOIN tblForm_Graphics ON (tblForm.frm_id = tblForm_Graphics.frm_id) AND
            ' **   (tblForm.dbs_id = tblForm_Graphics.dbs_id)) ON tblDatabase.dbs_id = tblForm.dbs_id
            ' ** WHERE (((tblDatabase.dbs_name) In ('Trust.mdb','Trust.mde')) AND ((tblForm.frm_name)='frmCurrency_Rate'));

3760        strTmp02 = strSQL
3770        strTmp01 = vbNullString
3780        For lngX = 0& To (lngFlds - 1&)
3790          If arr_varFld(F_IMG, lngX) = True Then
                ' ** It's an xadgfx_image field.
3800            Select Case arr_varFld(F_FND, lngX)
                Case True
                  ' ** The field has already be named.
3810            Case False
3820              intLen = Len(arr_varFld(F_FNAM, lngX))
3830              intPos01 = InStr(strTmp02, arr_varFld(F_FNAM, lngX))
3840              intPos01 = intPos01 + intLen
3850              If Mid(strTmp02, intPos01, 1) <> "," And lngX <> (lngFlds - 1&) Then
3860                Stop
3870              End If
3880              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1))
3890              strTmp02 = Mid(strTmp02, intPos01)
3900              strTmp01 = strTmp01 & " AS " & arr_varFld(F_VAL, lngX - 1)  ' ** Get control name from preceding field's value.
3910            End Select
3920          End If
3930        Next
3940        strTmp01 = strTmp01 & strTmp02  ' ** Get the last of the SQL.

3950        .SQL = strTmp01

3960      End With
3970      Set qdf = Nothing

3980      .Close
3990    End With

4000    Beep

EXITP:
4010    Set rst = Nothing
4020    Set qdf = Nothing
4030    Set dbs = Nothing
4040    Qry_XAdminGfx = blnRetVal
4050    Exit Function

ERRH:
4060    blnRetVal = False
4070    Select Case ERR.Number
        Case Else
4080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4090    End Select
4100    Resume EXITP

End Function

Public Function GTRStuff_Mod(frm As Access.Form, intMode As Integer, blnGTR As Boolean, blnGTR_Emblem As Boolean, blnGoingToReport As Boolean, blnGoneToReport As Boolean, lngGTR_Stat As Long, lngGTR_ID As Long, varJColID As Variant) As Boolean

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "GTRStuff_mod"

        Dim ctl As Access.Control
        Dim lngTpp As Long
        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

4210    blnRetVal = False

        ' ** GTRStuff_mod
        ' **   1 : GTREmblem_Set blnSetEmblem
        ' **   2 : GTREmblem_Get {none}
        ' **   3 : GTREmblem_Move {none}
        ' **   4 : GTREmblem_Off {none}
        ' **   5 : GTRGone_Set blnWentToReport As Boolean, lngJColID As Lo

4220    If lngTpp = 0& Then
          'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
4230      lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
4240    End If

4250    Select Case intMode
        Case 1
          ' ** GTREmblem_Set(blnSetEmblem).
4260      With frm
4270        blnGTR_Emblem = blnGTR
4280        If blnGTR_Emblem = False Then
4290          .TimerInterval = 0&
4300          blnGoingToReport = False
4310          GTRStuff_Mod frm, 4, False, blnGTR_Emblem, blnGoingToReport, blnGoneToReport, lngGTR_Stat, lngGTR_ID, 0  ' ** Recursion.
4320          GTRStuff_Mod frm, 6, False, blnGTR_Emblem, blnGoingToReport, blnGoneToReport, lngGTR_Stat, lngGTR_ID, 0  ' ** Recursion.
4330          lngGTR_Stat = 0&
4340          If .Sizable_lbl2.Visible = False Then
4350            lngTmp02 = (.GoToReport_Emblem_01_img.Width + (8& * lngTpp))
4360            .Sizable_lbl2.Visible = True
4370            .Sizable_lbl1.Visible = True
4380            .cmdClose.Left = (.cmdClose.Left + lngTmp02)
4390            .Footer_vline08.Left = (.Footer_vline08.Left + lngTmp02)
4400            .Footer_vline07.Left = (.Footer_vline07.Left + lngTmp02)
4410            .cmdDelete.Left = (.cmdDelete.Left + lngTmp02)
4420            .cmdEdit.Left = (.cmdEdit.Left + lngTmp02)
4430            .cmdAdd.Left = (.cmdAdd.Left + lngTmp02)
4440            .cmdAdd_box.Left = (.cmdAdd_box.Left + lngTmp02)
4450            .Footer_vline06.Left = (.Footer_vline06.Left + lngTmp02)
4460            .Footer_vline05.Left = (.Footer_vline05.Left + lngTmp02)
4470            .Footer_vline04.Visible = True
4480            .Footer_vline03.Visible = True
4490            .cmdRefresh.Visible = True
4500          End If
4510        End If
4520      End With
4530    Case 2
          ' ** GTREmblem_Get(). (blnGTR ignored)
4540      blnRetVal = blnGTR_Emblem
4550    Case 3
          ' ** GTREmblem_Move(). (blnGTR ignored)
4560      With frm
4570        lngTmp01 = ((.Sizable_lbl1.Left + .Sizable_lbl1.Width) - .GoToReport_Emblem_01_img.Width)
4580        For lngX = 1& To 24&
4590          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Left = lngTmp01
4600        Next
4610      End With
4620    Case 4
          ' ** GTREmblem_Off(). (blnGTR ignored)
4630      With frm
4640        blnGTR_Emblem = False
4650        For lngX = 1& To 24&
4660          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
4670        Next
4680      End With
4690    Case 5
          ' ** GTRGone_Set(blnWentToReport As Boolean, lngJColID As Long).
4700      blnGoneToReport = blnGTR
4710      lngGTR_ID = varJColID
4720    Case 6
          ' ** Turn off all GTR arrows.
4730      With frm
4740        If .GoToReport_arw_mapstcgl_img.Visible = True Then
4750          .cmdSpecPurp_Misc_MapLTCL_raised_img.Visible = True
4760          .cmdSpecPurp_Misc_MapLTCL.Enabled = True
4770        ElseIf .GoToReport_arw_mapsplit_img.Visible = True Then
4780          .cmdSpecPurp_Misc_MapSTCGL_raised_img.Visible = True
4790          .cmdSpecPurp_Misc_MapSTCGL.Enabled = True
4800        End If
4810        For Each ctl In .FormHeader.Controls
4820          With ctl
4830            If .ControlType = acBoundObjectFrame Then
4840              If Left(.Name, 15) = "GoToReport_arw_" Then
4850                .Visible = False
4860              End If
4870            End If
4880          End With
4890        Next
4900      End With
4910    End Select

EXITP:
4920    Set ctl = Nothing
4930    GTRStuff_Mod = blnRetVal
4940    Exit Function

ERRH:
4950    blnRetVal = False
4960    Select Case ERR.Number
        Case Else
4970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4980    End Select
4990    Resume EXITP

End Function

