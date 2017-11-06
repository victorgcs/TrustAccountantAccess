VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptTransaction_Audit_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptTransaction_Audit_01"

'VGC 10/06/2017: CHANGES!

Private Const strCallingForm As String = "frmTransaction_Audit"

' ** Array: arr_varFrmFld().
Private lngFrmFlds As Long, arr_varFrmFld As Variant
'Private Const F_ELEMS As Integer = 4  ' ** Array's first-element UBound().
Private Const F_FNAM As Integer = 0
'Private Const F_FTAB As Integer = 1
'Private Const F_FVIS As Integer = 2
'Private Const F_CNAM As Integer = 3
Private Const F_CVAL As Integer = 4

' ** Array: arr_varRptFld().
Private lngRptFlds As Long, arr_varRptFld() As Variant
Private Const R_ELEMS As Integer = 16  ' ** Array's first-element UBound().
Private Const R_FNAM  As Integer = 0
Private Const R_FTOP  As Integer = 1
Private Const R_FLFT  As Integer = 2
Private Const R_FWDT  As Integer = 3
Private Const R_FVIS  As Integer = 4
Private Const R_FTOPN As Integer = 5
Private Const R_FLFTN As Integer = 6
Private Const R_FWDTN As Integer = 7
Private Const R_LNAM  As Integer = 8
Private Const R_LTOP  As Integer = 9
Private Const R_LLFT  As Integer = 10
Private Const R_LWDT  As Integer = 11
Private Const R_LTOPN As Integer = 12
Private Const R_LLFTN As Integer = 13
Private Const R_LWDTN As Integer = 14
Private Const R_SHOW  As Integer = 15
Private Const R_SORT  As Integer = 16

' ** Array: arr_varNarrow().
Private lngNarrows As Long, arr_varNarrow() As Variant
Private Const N_ELEMS As Integer = 7  ' ** Array's first-element UBound().
Private Const N_NAM  As Integer = 0
Private Const N_VIS  As Integer = 1
Private Const N_ONAM As Integer = 2
Private Const N_OVIS As Integer = 3
Private Const N_TOP  As Integer = 4
Private Const N_LFT  As Integer = 5
Private Const N_NTOP As Integer = 6
Private Const N_NLFT As Integer = 7

Private lngGrandTotLbl_Offset As Long
Private lngChkCnt As Long, lngVisCnt As Long, lngTpp As Long
Private lngPageHeaderSection_Height As Long
Private lngOneLine1_Top As Long, lngTwoLine1_Top As Long, lngOneLine2_Top As Long, lngNarrow1_Top As Long, lngNarrow2_Top As Long
Private lngUnderline_Top As Long, lngHeaderLine6_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim frm As Access.Form, ctl As Access.Control
        Dim lngWidth As Long, lngGutter As Long
        Dim lng1stRowTop As Long, lng2ndRowTop As Long, lng2ndRowLeft As Long
        Dim lng1stRowLabelTop As Long, lng2ndRowLabelTop As Long, lng1stRowLabelHighTop As Long
        Dim lngFirstVisFldElem As Long
        Dim lngShareElem As Long, lngICashElem As Long, lngPCashElem As Long, lngCostElem As Long, lngJrnlUserElem As Long
        Dim lngLedgDescElem As Long, lngRecurItemElem As Long, lngLedgHidElem As Long, lngLocNameElem As Long, lngCheckNumElem As Long
        Dim lngRevDescElem As Long, lngRevTypeElem As Long, lngTaxDescElem As Long, lngTaxTypeElem As Long, lngPostedElem As Long
        Dim lngLastTop As Long, lngLastLeft As Long, lngLastWidth As Long, lngCurrentRow As Long
        Dim strOrderBy As String
        Dim blnVisRemain As Boolean, blnFound As Boolean
        Dim intPos01 As Integer, intMode As Integer
        Dim arr_varTmp00 As Variant, varTmp01 As Variant, strTmp02 As String, lngTmp03 As Long
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnContinue As Boolean

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         blnContinue = True

140         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

150         If IsNull(.OpenArgs) = False Then
160           If Trim(.OpenArgs) <> vbNullString Then
170             intPos01 = InStr(.OpenArgs, "~")
180             If intPos01 > 0 Then
190               intMode = Val(Mid(.OpenArgs, (intPos01 + 1)))
200             Else
210               intMode = 0
220             End If
230           Else
                ' ** Word.
240             .Filter = gstrCrtRpt_Ordinal
250             .FilterOn = True
260             intMode = 0
270           End If
280         Else
              ' ** Or this is Word.
290           .Filter = gstrCrtRpt_Ordinal
300           .FilterOn = True
310           intMode = 0
320         End If

330         Set frm = Forms("frmTransaction_Audit")

340         If intMode = acViewPreview Then
350           frm.Visible = False
360         End If

370         Select Case frm.chkPageOf
            Case True
380           .PageNumOf.Visible = True
390           .PageNum.Visible = False
400         Case False
410           .PageNum.Visible = True
420           .PageNumOf.Visible = False
430           .PageNumOf.ControlSource = vbNullString
440         End Select

450         If frm.ForExCnt_sum <> 0 Then
460           .ForExMsg_lbl.Visible = True
470           .ForExMsg_lbl_line.Visible = True
480         Else
490           .ForExMsg_lbl.Visible = False
500           .ForExMsg_lbl_line.Visible = False
510         End If

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
520         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

            ' ** Just get these out of the way.
530         .page_header_height_line.Height = lngTpp
540         .detail_height_line.Height = lngTpp

550         lngPageHeaderSection_Height = .Section(acPageHeader).Height
560         lngOneLine1_Top = .shortname_lbl.Top
570         lngTwoLine1_Top = .icash_lbl.Top
580         lngOneLine2_Top = .ledger_description_lbl.Top
590         lngUnderline_Top = .ledger_description_lbl_line.Top
600         lngHeaderLine6_Top = .Header_hline06.Top
610         lngNarrow1_Top = .curr_code_lbl.Top
620         lngNarrow2_Top = .CheckNum_lbl.Top

630         .AsOf_lbl.Caption = "As of " & Format(Date, "mm/dd/yyyy")
640         .AsOf_lbl2.Caption = "As of " & Format(Date, "mm/dd/yyyy")
650         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

660         strOrderBy = frm.frmTransaction_Audit_Sub.Form.SortNow_Get  ' ** Form Function: frmTransaction_Audit_Sub.
            ' ** Potential Sorts:
            ' **   [journalno]
            ' **   [journaltype], [alphasort], [journaltype], [journalno]
            ' **   [transdate], [alphasort], [journaltype], [journalno]
            ' **   [alphasort], [journaltype], [journalno]
            ' **   [shortname], [alphasort], [journaltype], [journalno]
            ' **   [cusip], [alphasort], [journaltype], [journalno]
            ' **   [asset_description], [alphasort], [journaltype], [journalno]
            ' **   [shareface], [alphasort], [journaltype], [journalno]
            ' **   [icash], [alphasort], [journaltype], [journalno]
            ' **   [pcash], [alphasort], [journaltype], [journalno]
            ' **   [cost], [alphasort], [journaltype], [journalno]
            ' **   [assetdate], [alphasort], [journaltype], [journalno]
            ' **   [PurchaseDate], [alphasort], [journaltype], [journalno]
            ' **   [ledger_description], [alphasort], [journaltype], [journalno]
            ' **   [RecurringItem], [alphasort], [journaltype], [journalno]
            ' **   [revcode_DESC], [alphasort], [journaltype], [journalno]
            ' **   [revcode_TYPE_Description], [revcode_DESC], [alphasort], [journaltype], [journalno]
            ' **   [taxcode_description], [alphasort], [journaltype], [journalno]
            ' **   [taxcode_type_description], [taxcode_description], [alphasort], [journaltype], [journalno]
            ' **   [journal_USER], [alphasort], [journaltype], [journalno]
            ' **   [posted], [alphasort], [journaltype], [journalno]
            ' **   [ledger_HIDDEN], [alphasort], [journaltype], [journalno]
670         .OrderBy = strOrderBy
680         .OrderByOn = True

690         .Caption = "Transaction_Audit_" & Format(Date, "yymmdd")

700         lngNarrows = 0&
710         ReDim arr_varNarrow(N_ELEMS, 0)

            ' ** Collect the controls that have the Arial Narrow font option.
720         For Each ctl In .Controls
730           With ctl
740             If Right(.Name, 7) = "_narrow" Then
750               lngNarrows = lngNarrows + 1&
760               lngE = lngNarrows - 1&
770               ReDim Preserve arr_varNarrow(N_ELEMS, lngE)
780               arr_varNarrow(N_NAM, lngE) = .Name
790               arr_varNarrow(N_VIS, lngE) = CBool(False)
800               arr_varNarrow(N_ONAM, lngE) = Left(.Name, (Len(.Name) - 7))
810               arr_varNarrow(N_OVIS, lngE) = CBool(True)
820               arr_varNarrow(N_TOP, lngE) = .Top
830               arr_varNarrow(N_LFT, lngE) = .Left
840               arr_varNarrow(N_NTOP, lngE) = Null
850               arr_varNarrow(N_NLFT, lngE) = Null
860             End If
870           End With
880         Next

890         lngGrandTotLbl_Offset = .shareface_tot.Left - .grand_tot_lbl.Left

900         lngRptFlds = 0&
910         ReDim arr_varRptFld(R_ELEMS, 0)

920         With frm

              ' ****************************************************
              ' ** Array: arr_varFrmFld()
              ' **
              ' **   Element  Name                    Constant
              ' **   =======  ======================  ============
              ' **      0     Field Name              F_FNAM
              ' **      1     Control Tab Index       F_FTAB
              ' **      2     Field Visible           F_FVIS
              ' **      3     Checkbox Name           F_CNAM
              ' **      4     Checkbox Value          F_CVAL
              ' **
              ' ****************************************************
930           lngFrmFlds = 0&
940           arr_varFrmFld = .FormFields_GetArr(.opgView)  ' ** Form Function: frmTransaction_Audit.
              ' ** Array: arr_varFrmFld().
              ' **   1 : journalno
              ' **   2 : journaltype
              ' **   3 : transdate
              ' **   4 : accountno
              ' **   5 : shortname
              ' **   6 : cusip
              ' **   7 : asset_description
              ' **   8 : shareface
              ' **   9 : icash
              ' **   10: pcash
              ' **   11: cost
              ' **   12: assetdate
              ' **   13: PurchaseDate
              ' **   14: ledger_description
              ' **   15: RecurringItem
              ' **   16: revcode_DESC
              ' **   17: revcode_TYPE_Description
              ' **   18: taxcode_description
              ' **   19: taxcode_type_description
              ' **   20: journal_USER
              ' **   21: posted
              ' **   22: ledger_HIDDEN
950   On Error Resume Next
960           varTmp01 = UBound(arr_varFrmFld, 2)  ' ** I don't really know why this was causing so much trouble!
970   On Error GoTo ERRH

980           If blnContinue = True Then
990             lngFrmFlds = UBound(arr_varFrmFld, 2) + 1&
1000            lngVisCnt = 0&: lngChkCnt = 0&
1010            arr_varTmp00 = VisibleCounts_Get(.opgView)  ' ** Module Function: modTransactionAuditFuncs1.
1020            lngChkCnt = arr_varTmp00(0)
1030            lngVisCnt = arr_varTmp00(1)
1040          End If

1050        End With  ' ** frm.

            ' ** These are all the fields currently available.
            'For lngX = 0& To (lngFrmFlds - 1&)
            '  Debug.Print "'" & Left(CStr(lngX + 1&) & "  ", 2) & ": " & arr_varFrmFld(F_FNAM, lngX)
            'Next
            ' ** 1 : journalno
            ' ** 2 : journaltype
            ' ** 3 : transdate
            ' ** 4 : accountno
            ' ** 5 : shortname
            ' ** 6 : cusip
            ' ** 7 : asset_description
            ' ** 8 : shareface
            ' ** 9 : icash
            ' ** 10: pcash
            ' ** 11: cost
            ' ** 12: curr_id                   ****
            ' ** 13: assetdate
            ' ** 14: PurchaseDate
            ' ** 15: ledger_description
            ' ** 16: RecurringItem
            ' ** 17: revcode_DESC
            ' ** 18: revcode_TYPE_Description
            ' ** 19: taxcode_description
            ' ** 20: taxcode_type_description
            ' ** 21: Location_Name             ****
            ' ** 22: CheckNum                  ****
            ' ** 23: journal_USER
            ' ** 24: posted
            ' ** 25: ledger_HIDDEN

1060        If lngVisCnt < lngChkCnt Then  'lngVisCnt > 100& Then

              ' **************************************************
              ' ** Array: arr_varRptFld()
              ' **
              ' **   Element  Name               Constant
              ' **   =======  =================  ===============
              ' **      0     Field Name         R_FNAM
              ' **      1     Field Top          R_FTOP
              ' **      2     Field Left         R_FLFT
              ' **      3     Field Width        R_FWDT
              ' **      4     Field Visible      R_FVIS
              ' **      5     Field New Top      R_FTOPN
              ' **      6     Field New Left     R_FLFTN
              ' **      7     Field New Width    R_FWDTN
              ' **      8     Label Name         R_LNAM
              ' **      9     Label Top          R_LTOP
              ' **     10     Label Left         R_LLFT
              ' **     11     Label Width        R_LWDT
              ' **     12     Label New Top      R_LTOPN
              ' **     13     Label New Left     R_LLFTN
              ' **     14     Label New Width    R_LWDTN
              ' **     15     Show/Hide          R_SHOW
              ' **     16     Sorting Field      R_SORT
              ' **
              ' **************************************************

              ' ** Get a list of the report data fields.
1070          For Each ctl In .Section("Detail").Controls
1080            With ctl
1090              If .Visible = True Then
1100                Select Case .Name
                    Case "assetno", "revcode_ID", "revcode_TYPE", "taxcode", "taxcode_type", "curr_id", "Location_ID", "detail_height_line"
                      ' ** Skip these.
1110                Case "shareface_narrow", "icash_narrow", "pcash_narrow", "cost_narrow", "accountno_narrow", _
                        "icash_str", "icash_str_narrow", "pcash_str", "pcash_str_narrow", "cost_str", "cost_str_narrow"
                      ' ** Skip these for now.
1120                Case Else
1130                  lngRptFlds = lngRptFlds + 1&
1140                  lngE = lngRptFlds - 1&
1150                  ReDim Preserve arr_varRptFld(R_ELEMS, lngE)
1160                  arr_varRptFld(R_FNAM, lngE) = .Name
1170                  arr_varRptFld(R_FTOP, lngE) = .Top
1180                  arr_varRptFld(R_FLFT, lngE) = .Left
1190                  If .Name = "ledger_HIDDEN" Then
1200                    arr_varRptFld(R_FWDT, lngE) = Me.Controls("ledger_HIDDEN_lbl").Width
1210                  Else
1220                    arr_varRptFld(R_FWDT, lngE) = .Width
1230                  End If
1240                  arr_varRptFld(R_FVIS, lngE) = .Visible
1250                  arr_varRptFld(R_FTOPN, lngE) = -1&
1260                  arr_varRptFld(R_FLFTN, lngE) = -1&
1270                  arr_varRptFld(R_FWDTN, lngE) = -1&
1280                  If .Visible = True Then
1290                    strTmp02 = .Name & "_lbl"
1300                    arr_varRptFld(R_LNAM, lngE) = strTmp02
1310                    arr_varRptFld(R_LLFT, lngE) = Me.Controls(strTmp02).Left
1320                    arr_varRptFld(R_LTOP, lngE) = Me.Controls(strTmp02).Top
1330                    arr_varRptFld(R_LWDT, lngE) = Me.Controls(strTmp02).Width
1340                    arr_varRptFld(R_LTOPN, lngE) = -1&
1350                    arr_varRptFld(R_LLFTN, lngE) = -1&
1360                    arr_varRptFld(R_LWDTN, lngE) = -1&
1370                  Else
1380                    arr_varRptFld(R_LNAM, lngE) = vbNullString
1390                    arr_varRptFld(R_LLFT, lngE) = -1&
1400                    arr_varRptFld(R_LTOP, lngE) = -1&
1410                    arr_varRptFld(R_LWDT, lngE) = -1&
1420                    arr_varRptFld(R_LTOPN, lngE) = -1&
1430                    arr_varRptFld(R_LLFTN, lngE) = -1&
1440                    arr_varRptFld(R_LWDTN, lngE) = -1&
1450                  End If
1460                  arr_varRptFld(R_SHOW, lngE) = CBool(False)
1470                  arr_varRptFld(R_SORT, lngE) = Right(String(8, "0") & CStr(.Top), 8) & Right(String(8, "0") & CStr(.Left), 8)
1480                End Select
1490              End If  ' ** Visible.
1500            End With  ' ** ctl.
1510          Next  ' ** ctl.

              ' ** Sort the list, left to right, top to bottom on report.
1520          RptFldSort  ' ** Procedure: Below.

              ' ** Get the report width.
1530          lngWidth = .Width

              ' ** Get the spacing between the fields and labels.
              ' ** This needs to remain consistent in order for the stuff below to work!
1540          lngGutter = arr_varRptFld(R_FLFT, 1) - (arr_varRptFld(R_FLFT, 0) + arr_varRptFld(R_FWDT, 0))

              ' ** Get the element numbers for fields needing special handling.
1550          For lngX = 0& To (lngRptFlds - 1&)
1560            Select Case arr_varRptFld(R_FNAM, lngX)
                Case "shareface"
1570              lngShareElem = lngX
1580            Case "icash"
1590              lngICashElem = lngX
1600            Case "pcash"
1610              lngPCashElem = lngX
1620            Case "cost"
1630              lngCostElem = lngX
1640            Case "revcode_DESC"
1650              lngRevDescElem = lngX
1660            Case "revcode_TYPE_Description"
1670              lngRevTypeElem = lngX
1680            Case "taxcode_description"
1690              lngTaxDescElem = lngX
1700            Case "taxcode_type_description"
1710              lngTaxTypeElem = lngX
1720            Case "ledger_description"
1730              lngLedgDescElem = lngX
1740            Case "RecurringItem"
1750              lngRecurItemElem = lngX
1760            Case "ledger_HIDDEN"
1770              lngLedgHidElem = lngX
1780            Case "Location_Name"
1790              lngLocNameElem = lngX
1800            Case "journal_USER"
1810              lngJrnlUserElem = lngX
1820            Case "posted"
1830              lngPostedElem = lngX
1840            Case "CheckNum"
1850              lngCheckNumElem = lngX
1860            End Select
1870          Next

              ' ** These are all the fields currently on the report.
              'For lngX = 0& To (lngRptFlds - 1&)
              '  Debug.Print "'" & Left(CStr(lngX + 1&) & "  ", 2) & ": " & arr_varRptFld(R_FNAM, lngX)
              'Next
              ' ** 1 : journalno
              ' ** 2 : journaltype
              ' ** 3 : transdate
              ' ** 4 : accountno
              ' ** 5 : shortname
              ' ** 6 : cusip
              ' ** 7 : asset_description
              ' ** 8 : shareface
              ' ** 9 : icash
              ' ** 10: pcash
              ' ** 11: cost
              ' ** 12: curr_code                 ****  Arial Narrow label, Top - 15&, Height + 15&
              ' ** 13: assetdate
              ' ** 14: ledger_HIDDEN
              ' ** 15: PurchaseDate
              ' ** 16: ledger_description
              ' ** 17: RecurringItem
              ' ** 18: CheckNum                  ****  Arial Narrow label, Top - 15&, Height + 15&
              ' ** 19: revcode_DESC
              ' ** 20: revcode_TYPE_Description
              ' ** 21: taxcode_description
              ' ** 22: taxcode_type_description
              ' ** 23: Location_Name
              ' ** 24: journal_USER
              ' ** 25: posted

              ' ** Get the Top for each of the 2 rows of data and labels on the report,
              ' ** as well as the left-most 2nd row field.
1880          lng1stRowTop = arr_varRptFld(R_FTOP, 0)
1890          lng1stRowLabelTop = arr_varRptFld(R_LTOP, 0)
1900          If lng1stRowLabelTop <= 0& Then
1910            lng1stRowLabelTop = .shortname_lbl.Top
1920          End If

1930          lng2ndRowTop = 0&: lng2ndRowLeft = 0&: lng2ndRowLabelTop = 0&: lng1stRowLabelHighTop = 0&
1940          For lngX = 0& To (lngRptFlds - 1&)
1950            If arr_varRptFld(R_LTOP, lngX) <> -1& Then
1960              If arr_varRptFld(R_FTOP, lngX) > lng1stRowTop And lng2ndRowTop = 0& Then
1970                lng2ndRowTop = arr_varRptFld(R_FTOP, lngX)
1980                lng2ndRowLeft = arr_varRptFld(R_FLFT, lngX)
1990                lng2ndRowLabelTop = arr_varRptFld(R_LTOP, lngX)
2000              End If
2010              If arr_varRptFld(R_LTOP, lngX) < lng1stRowLabelTop And lng1stRowLabelHighTop = 0& Then
2020                lng1stRowLabelHighTop = arr_varRptFld(R_LTOP, lngX)
2030              End If
2040            End If
2050          Next

              ' ** Transfer the checkbox values to the report list.
2060          For lngX = 0& To (lngFrmFlds - 1&)
2070            For lngY = 0& To (lngRptFlds - 1&)
2080              If arr_varRptFld(R_FNAM, lngY) = arr_varFrmFld(F_FNAM, lngX) Then
2090                If arr_varRptFld(R_FVIS, lngY) = True Then
2100                  arr_varRptFld(R_SHOW, lngY) = CBool(arr_varFrmFld(F_CVAL, lngX))
2110                Else
                      ' ** For non-visible ID fields, checkbox elements will remain vbNullString and False.
                      ' ** Their Show will always be False, but they've got to be moved up
                      ' ** if they're in the 2nd row, and we can close it up to a single row.
2120                End If
2130                Exit For
2140              Else
2150                If arr_varRptFld(R_FNAM, lngY) = "curr_code" And arr_varFrmFld(F_FNAM, lngX) = "curr_id" Then
2160                  If arr_varRptFld(R_FVIS, lngY) = True Then
2170                    arr_varRptFld(R_SHOW, lngY) = CBool(arr_varFrmFld(F_CVAL, lngX))
2180                  End If
2190                  Exit For
2200                End If
2210              End If
2220            Next
2230          Next

              ' ** Turn off all report fields not checked on form.
2240          For lngX = 0& To (lngRptFlds - 1&)
2250            .Controls(arr_varRptFld(R_FNAM, lngX)).Visible = arr_varRptFld(R_SHOW, lngX)
2260            If arr_varRptFld(R_FVIS, lngX) = True Then
                  ' ** Non-visible ID fields don't have labels.
2270              .Controls(arr_varRptFld(R_LNAM, lngX)).Visible = arr_varRptFld(R_SHOW, lngX)
2280            End If
2290          Next

              ' ** Find the first visible field.
2300          lngFirstVisFldElem = -1&: blnVisRemain = False
2310          For lngX = 0& To (lngRptFlds - 1&)
2320            If arr_varRptFld(R_SHOW, lngX) = True And lngFirstVisFldElem = -1& Then
2330              lngFirstVisFldElem = lngX
2340            ElseIf lngFirstVisFldElem <> -1& And arr_varRptFld(R_SHOW, lngX) = True Then
2350              blnVisRemain = True
2360              Exit For
2370            End If
2380          Next

2390          arr_varRptFld(R_FTOPN, lngFirstVisFldElem) = lng1stRowTop
2400          arr_varRptFld(R_FLFTN, lngFirstVisFldElem) = arr_varRptFld(R_FLFT, 0)
2410          arr_varRptFld(R_FWDTN, lngFirstVisFldElem) = arr_varRptFld(R_FWDT, lngFirstVisFldElem)

2420          lngLastTop = lng1stRowTop
2430          lngLastLeft = arr_varRptFld(R_FLFTN, lngFirstVisFldElem)
2440          lngLastWidth = arr_varRptFld(R_FWDTN, lngFirstVisFldElem)
2450          lngCurrentRow = 1&

              ' ** Now repeat until there are no more visible fields to deal with.
2460          Do While blnVisRemain = True
2470            For lngX = (lngFirstVisFldElem + 1&) To (lngRptFlds - 1&)
2480              If arr_varRptFld(R_SHOW, lngX) = True Then
2490                If (lngLastLeft + lngLastWidth + lngGutter) + arr_varRptFld(R_FWDT, lngX) <= lngWidth Then
2500                  arr_varRptFld(R_FTOPN, lngX) = lngLastTop
2510                  arr_varRptFld(R_FLFTN, lngX) = (lngLastLeft + lngLastWidth + lngGutter)
2520                  arr_varRptFld(R_FWDTN, lngX) = arr_varRptFld(R_FWDT, lngX)
2530                  lngLastTop = arr_varRptFld(R_FTOPN, lngX)
2540                  lngLastLeft = arr_varRptFld(R_FLFTN, lngX)
2550                  lngLastWidth = arr_varRptFld(R_FWDTN, lngX)
2560                Else
2570                  If lngCurrentRow = 1& Then
2580                    lngCurrentRow = 2&
2590                    arr_varRptFld(R_FTOPN, lngX) = lng2ndRowTop
2600                    arr_varRptFld(R_FLFTN, lngX) = lng2ndRowLeft
2610                    arr_varRptFld(R_FWDTN, lngX) = arr_varRptFld(R_FWDT, lngX)
2620                  Else
2630                    arr_varRptFld(R_FTOPN, lngX) = lngLastTop
2640                    arr_varRptFld(R_FLFTN, lngX) = (lngLastLeft + lngLastWidth + lngGutter)
2650                    arr_varRptFld(R_FWDTN, lngX) = arr_varRptFld(R_FWDT, lngX)
2660                  End If
2670                  lngLastTop = arr_varRptFld(R_FTOPN, lngX)
2680                  lngLastLeft = arr_varRptFld(R_FLFTN, lngX)
2690                  lngLastWidth = arr_varRptFld(R_FWDTN, lngX)
2700                End If
2710                blnVisRemain = False
2720                For lngY = (lngX + 1&) To (lngRptFlds - 1&)
2730                  If arr_varRptFld(R_SHOW, lngY) = True Then
2740                    blnVisRemain = True
2750                    Exit For
2760                  End If
2770                Next
2780              End If
2790            Next
2800          Loop

              ' ** The new positions should be calculated (correctly, I hope),
              ' ** so now start moving the fields on the report.
2810          For lngX = 0& To (lngRptFlds - 1&)
2820            If arr_varRptFld(R_FTOPN, lngX) <> -1& Then
2830              .Controls(arr_varRptFld(R_FNAM, lngX)).Top = arr_varRptFld(R_FTOPN, lngX)
2840              .Controls(arr_varRptFld(R_FNAM, lngX)).Left = arr_varRptFld(R_FLFTN, lngX)
2850            End If
2860          Next

2870          For lngX = 0& To (lngRptFlds - 1&)
2880            If arr_varRptFld(R_SHOW, lngX) = True Then
2890              Select Case arr_varRptFld(R_FNAM, lngX)
                  Case "icash"
2900                .icash_str.Top = .ICash.Top
2910                .icash_str.Left = .ICash.Left
2920              Case "pcash"
2930                .pcash_str.Top = .PCash.Top
2940                .pcash_str.Left = .PCash.Left
2950              Case "cost"
2960                .cost_str.Top = .Cost.Top
2970                .cost_str.Left = .Cost.Left
2980              Case Else
                    ' ** Nothing else.
2990              End Select
3000            End If
3010          Next

              ' ** Move the narrows to their complement's position.
3020          For lngX = 0& To (lngNarrows - 1&)
3030            Set ctl = .Controls(arr_varNarrow(N_ONAM, lngX))  ' ** Original Name (normal font).
3040            With ctl
3050              If .Top <> arr_varNarrow(N_TOP, lngX) Then
3060                If .Top = 0 Then
3070                  arr_varNarrow(N_NTOP, lngX) = .Top
3080                Else
3090                  arr_varNarrow(N_NTOP, lngX) = .Top - lngTpp  ' ** Narrow fonts sit lower, so move them up 1 pixel.
3100                End If
3110              End If
3120              If .Left <> arr_varNarrow(N_LFT, lngX) Then
3130                arr_varNarrow(N_NLFT, lngX) = .Left
3140              End If
3150            End With
3160            Set ctl = .Controls(arr_varNarrow(N_NAM, lngX))
3170            With ctl
3180              If IsNull(arr_varNarrow(N_NTOP, lngX)) = False Then
3190  On Error Resume Next
3200                If arr_varNarrow(N_NTOP, lngX) < 0& Then
3210                  .Top = 0&
3220                Else
3230                  .Top = arr_varNarrow(N_NTOP, lngX)
3240                End If
3250  On Error GoTo ERRH
3260              End If
3270              If IsNull(arr_varNarrow(N_NLFT, lngX)) = False Then
3280                .Left = arr_varNarrow(N_NLFT, lngX)
3290              End If
3300            End With
3310          Next

              ' ** If everything visible is on the 1st row, close it up.
3320          If lngCurrentRow = 1& Then
3330            For lngX = 0& To (lngRptFlds - 1&)
3340              If arr_varRptFld(R_SHOW, lngX) = False And arr_varRptFld(R_FTOP, lngX) = lng2ndRowTop Then
3350                .Controls(arr_varRptFld(R_FNAM, lngX)).Top = lng1stRowTop
3360              End If
3370            Next
3380            .detail_height_line.Top = 0&
3390            .detail_height_line.Height = .transdate.Top + .transdate.Height + 15&
3400            .Detail.Height = .transdate.Top + .transdate.Height + 15&
3410          End If

              ' ** Move the labels to match the fields.
3420          For lngX = 0& To (lngRptFlds - 1&)
3430            If arr_varRptFld(R_FTOPN, lngX) <> -1& Then
3440              arr_varRptFld(R_LLFTN, lngX) = .Controls(arr_varRptFld(R_FNAM, lngX)).Left
                  'The label for ledger_HIDDEN is wider than the field!
3450              .Controls(arr_varRptFld(R_LNAM, lngX)).Left = arr_varRptFld(R_LLFTN, lngX)
                  ' ** Move the lines for 2nd row fields.
3460              Select Case lngX
                  Case lngRevDescElem, lngRevTypeElem, lngTaxDescElem, lngTaxTypeElem, lngLedgDescElem, _
                      lngRecurItemElem, lngLedgHidElem, lngLocNameElem, lngJrnlUserElem, lngPostedElem, lngCheckNumElem
3470                If arr_varRptFld(R_LNAM, lngX) = "ledger_HIDDEN_lbl" Then
                      ' ** There is no line for ledger_HIDDEN_lbl.
3480                Else
3490                  .Controls(arr_varRptFld(R_LNAM, lngX) & "_line").Left = arr_varRptFld(R_LLFTN, lngX)
3500                End If
3510              Case Else
                    ' ** No lines to move.
3520              End Select
3530              If arr_varRptFld(R_FTOPN, lngX) = lng1stRowTop And arr_varRptFld(R_FTOP, lngX) = lng2ndRowTop Then
                    ' ** If the field's now in the 1st row, and it was in the 2nd row, move it's label up.
3540                arr_varRptFld(R_LTOPN, lngX) = lng1stRowLabelTop
3550  On Error Resume Next
3560                .Controls(arr_varRptFld(R_LNAM, lngX)).Top = arr_varRptFld(R_LTOPN, lngX)
3570                If ERR.Number <> 0 Then
3580  On Error GoTo ERRH
                      'Debug.Print "'" & arr_varRptFld(R_LNAM, lngX) & "  " & arr_varRptFld(R_LTOPN, lngX)
3590                Else
3600  On Error GoTo ERRH
3610                End If
3620              Else
3630                arr_varRptFld(R_LTOPN, lngX) = arr_varRptFld(R_LTOP, lngX)
3640              End If
3650              arr_varRptFld(R_LWDTN, lngX) = arr_varRptFld(R_LWDT, lngX)
3660            End If
3670          Next

              ' ** If we've got room, widen the cash and cost fields.
              ' ** I'd like to confine this to just when they're the last fields on the line.
3680          lngLastLeft = 0&: lngTmp03 = 0&
3690          For lngX = 0& To (lngRptFlds - 1&)
3700            If arr_varRptFld(R_SHOW, lngX) = True Then
3710              If arr_varRptFld(R_LLFTN, lngX) > lngLastLeft Then
3720                lngLastLeft = arr_varRptFld(R_LLFTN, lngX)
3730                lngLastWidth = arr_varRptFld(R_LWDTN, lngX)
3740              End If
3750            End If
3760          Next
3770          lngTmp03 = (lngLastLeft + lngLastWidth)
3780          If ((.Width - (4& * lngTpp)) - lngTmp03) > (12& * lngTpp) Then  ' ** Width - 60&, > 180&.
                ' ** There's at least enough room to widen by 60.

3790            If ((.Width - (4& * lngTpp)) - lngTmp03) > (36& * lngTpp) Then  ' ** Width - 60&, > 540&.
                  ' ** Widen by 180.
3800              lngTmp03 = (12& * lngTpp)
3810            ElseIf ((.Width - (4& * lngTpp)) - lngTmp03) > (24& * lngTpp) Then  ' ** Width - 60&, > 360&.
                  ' ** Widen by 120.
3820              lngTmp03 = (8& * lngTpp)
3830            ElseIf ((.Width - (4& * lngTpp)) - lngTmp03) > (12& * lngTpp) Then  ' ** Width - 60&, > 180&.
                  ' ** Widen by 60.
3840              lngTmp03 = (4& * lngTpp)
3850            End If

3860            If .Cost.Visible = True And .Cost.Left = lngLastLeft Then

3870              If .ICash.Visible = True And .PCash.Visible = True Then
                    ' ** Widen all three.
3880                .Cost.Left = (.Cost.Left + (2& * lngTmp03))
3890                .Cost.Width = (.Cost.Width + lngTmp03)
3900                .cost_str.Left = .Cost.Left
3910                .cost_str.Width = .Cost.Width
3920                .cost_narrow.Left = (.cost_narrow.Left + (2& * lngTmp03))
3930                .cost_narrow.Width = (.cost_narrow.Width + lngTmp03)
3940                .cost_str_narrow.Left = .cost_narrow.Left
3950                .cost_str_narrow.Width = .cost_narrow.Width
3960                .PCash.Left = (.PCash.Left + lngTmp03)
3970                .PCash.Width = (.PCash.Width + lngTmp03)
3980                .pcash_str.Left = .PCash.Left
3990                .pcash_str.Width = .PCash.Width
4000                .pcash_narrow.Left = (.pcash_narrow.Left + lngTmp03)
4010                .pcash_narrow.Width = (.pcash_narrow.Width + lngTmp03)
4020                .pcash_str_narrow.Left = .pcash_narrow.Left
4030                .pcash_str_narrow.Width = .pcash_narrow.Width
4040                .ICash.Width = (.ICash.Width + lngTmp03)
4050                .icash_str.Width = .ICash.Width
4060                .icash_narrow.Width = (.icash_narrow.Width + lngTmp03)
4070                .icash_str_narrow.Width = .icash_narrow.Width
4080                .cost_lbl.Left = .Cost.Left
4090                .cost_lbl.Width = .Cost.Width
4100                .pcash_lbl.Left = .PCash.Left
4110                .pcash_lbl.Width = .PCash.Width
4120                .icash_lbl.Width = .ICash.Width
4130                arr_varRptFld(R_FLFTN, lngCostElem) = .Cost.Left
4140                arr_varRptFld(R_FWDTN, lngCostElem) = .Cost.Width
4150                arr_varRptFld(R_LLFTN, lngCostElem) = .cost_lbl.Left
4160                arr_varRptFld(R_LWDTN, lngCostElem) = .cost_lbl.Width
4170                arr_varRptFld(R_FLFTN, lngPCashElem) = .PCash.Left
4180                arr_varRptFld(R_FWDTN, lngPCashElem) = .PCash.Width
4190                arr_varRptFld(R_LLFTN, lngPCashElem) = .pcash_lbl.Left
4200                arr_varRptFld(R_LWDTN, lngPCashElem) = .pcash_lbl.Width
4210                arr_varRptFld(R_FWDTN, lngICashElem) = .ICash.Width
4220                arr_varRptFld(R_LWDTN, lngICashElem) = .icash_lbl.Width
4230              ElseIf .ICash.Visible = True Then
                    ' ** Widen icash and cost.
4240                .Cost.Left = (.Cost.Left + lngTmp03)
4250                .Cost.Width = (.Cost.Width + lngTmp03)
4260                .cost_str.Left = .Cost.Left
4270                .cost_str.Width = .Cost.Width
4280                .cost_narrow.Left = (.cost_narrow.Left + (2& * lngTmp03))
4290                .cost_narrow.Width = (.cost_narrow.Width + lngTmp03)
4300                .cost_str_narrow.Left = .cost_narrow.Left
4310                .cost_str_narrow.Width = .cost_narrow.Width
4320                .ICash.Width = (.ICash.Width + lngTmp03)
4330                .icash_str.Width = .ICash.Width
4340                .icash_narrow.Width = (.icash_narrow.Width + lngTmp03)
4350                .icash_str_narrow.Width = .icash_narrow.Width
4360                .cost_lbl.Left = .Cost.Left
4370                .cost_lbl.Width = .Cost.Width
4380                .icash_lbl.Width = .ICash.Width
4390                arr_varRptFld(R_FLFTN, lngCostElem) = .Cost.Left
4400                arr_varRptFld(R_FWDTN, lngCostElem) = .Cost.Width
4410                arr_varRptFld(R_LLFTN, lngCostElem) = .cost_lbl.Left
4420                arr_varRptFld(R_LWDTN, lngCostElem) = .cost_lbl.Width
4430                arr_varRptFld(R_FWDTN, lngICashElem) = .ICash.Width
4440                arr_varRptFld(R_LWDTN, lngICashElem) = .icash_lbl.Width
4450              ElseIf .PCash.Visible = True Then
                    ' ** Widen pcash and cost.
4460                .Cost.Left = (.Cost.Left + lngTmp03)
4470                .Cost.Width = (.Cost.Width + lngTmp03)
4480                .cost_str.Left = .Cost.Left
4490                .cost_str.Width = .Cost.Width
4500                .cost_narrow.Left = (.cost_narrow.Left + (2& * lngTmp03))
4510                .cost_narrow.Width = (.cost_narrow.Width + lngTmp03)
4520                .cost_str_narrow.Left = .cost_narrow.Left
4530                .cost_str_narrow.Width = .cost_narrow.Width
4540                .PCash.Width = (.PCash.Width + lngTmp03)
4550                .pcash_str.Width = .PCash.Width
4560                .pcash_narrow.Width = (.pcash_narrow.Width + lngTmp03)
4570                .pcash_str_narrow.Width = .pcash_narrow.Width
4580                .cost_lbl.Left = .Cost.Left
4590                .cost_lbl.Width = .Cost.Width
4600                .pcash_lbl.Width = .PCash.Width
4610                arr_varRptFld(R_FLFTN, lngCostElem) = .Cost.Left
4620                arr_varRptFld(R_FWDTN, lngCostElem) = .Cost.Width
4630                arr_varRptFld(R_LLFTN, lngCostElem) = .cost_lbl.Left
4640                arr_varRptFld(R_LWDTN, lngCostElem) = .cost_lbl.Width
4650                arr_varRptFld(R_FWDTN, lngPCashElem) = .PCash.Width
4660                arr_varRptFld(R_LWDTN, lngPCashElem) = .pcash_lbl.Width
4670              Else
                    ' ** Widen cost alone.
4680                .Cost.Width = (.Cost.Width + lngTmp03)
4690                .cost_str.Width = .Cost.Width
4700                .cost_narrow.Width = (.cost_narrow.Width + lngTmp03)
4710                .cost_str_narrow.Width = .cost_narrow.Width
4720                .cost_lbl.Width = .Cost.Width
4730                arr_varRptFld(R_FWDTN, lngCostElem) = .Cost.Width
4740                arr_varRptFld(R_LWDTN, lngCostElem) = .cost_lbl.Width
4750              End If

4760            ElseIf .PCash.Visible = True And .PCash.Left = lngLastLeft Then

4770              If .ICash.Visible = True Then
                    ' ** Widen both icash and pcash.
4780                .PCash.Left = (.PCash.Left + lngTmp03)
4790                .PCash.Width = (.PCash.Width + lngTmp03)
4800                .pcash_str.Left = .PCash.Left
4810                .pcash_str.Width = .PCash.Width
4820                .pcash_narrow.Left = (.pcash_narrow.Left + lngTmp03)
4830                .pcash_narrow.Width = (.pcash_narrow.Width + lngTmp03)
4840                .pcash_str_narrow.Left = .pcash_narrow.Left
4850                .pcash_str_narrow.Width = .pcash_narrow.Width
4860                .ICash.Width = (.ICash.Width + lngTmp03)
4870                .icash_str.Width = .ICash.Width
4880                .icash_narrow.Width = (.icash_narrow.Width + lngTmp03)
4890                .icash_str_narrow.Width = .icash_narrow.Width
4900                .pcash_lbl.Left = .PCash.Left
4910                .pcash_lbl.Width = .PCash.Width
4920                .icash_lbl.Width = .ICash.Width
4930                arr_varRptFld(R_FLFTN, lngPCashElem) = .PCash.Left
4940                arr_varRptFld(R_FWDTN, lngPCashElem) = .PCash.Width
4950                arr_varRptFld(R_LLFTN, lngPCashElem) = .pcash_lbl.Left
4960                arr_varRptFld(R_LWDTN, lngPCashElem) = .pcash_lbl.Width
4970                arr_varRptFld(R_FWDTN, lngICashElem) = .ICash.Width
4980                arr_varRptFld(R_LWDTN, lngICashElem) = .icash_lbl.Width
4990              Else
                    ' ** Widen pcash alone.
5000                .PCash.Width = (.PCash.Width + lngTmp03)
5010                .pcash_str.Width = .PCash.Width
5020                .pcash_narrow.Width = (.pcash_narrow.Width + lngTmp03)
5030                .pcash_str_narrow.Width = .pcash_narrow.Width
5040                .pcash_lbl.Width = .PCash.Width
5050                arr_varRptFld(R_FWDTN, lngPCashElem) = .PCash.Width
5060                arr_varRptFld(R_LWDTN, lngPCashElem) = .pcash_lbl.Width
5070              End If

5080            ElseIf .ICash.Visible = True And .ICash.Left = lngLastLeft Then
                  ' ** Widen icash alone.

5090              .ICash.Width = (.ICash.Width + lngTmp03)
5100              .icash_str.Width = .ICash.Width
5110              .icash_narrow.Width = (.icash_narrow.Width + lngTmp03)
5120              .icash_str_narrow.Width = .icash_narrow.Width
5130              .icash_lbl.Width = .ICash.Width
5140              arr_varRptFld(R_FWDTN, lngICashElem) = .ICash.Width
5150              arr_varRptFld(R_LWDTN, lngICashElem) = .icash_lbl.Width

5160            End If

5170          End If

              ' ** If 2nd row fields are now on the 1st row, get rid of their label lines.
5180          For lngX = 0& To (lngRptFlds - 1&)
5190            If arr_varRptFld(R_FTOP, lngX) = lng2ndRowTop And (arr_varRptFld(R_FTOPN, lngX) = lng1stRowTop Or _
                    (arr_varRptFld(R_FVIS, lngX) = True And arr_varRptFld(R_SHOW, lngX) = False)) Then
5200  On Error Resume Next
                  ' ** It was looking for journalno_lbl_line!
5210              .Controls(arr_varRptFld(R_LNAM, lngX) & "_line").Visible = False
5220  On Error GoTo ERRH
5230            End If
5240          Next
              'assetdate_lbl_line
              'PurchaseDate_lbl_line
              'revcode_DESC_lbl_line
              'revcode_TYPE_Description_lbl_line
              'taxcode_description_lbl_line
              'taxcode_type_description_lbl_line
              'ledger_description_lbl_line
              'RecurringItem_lbl_line

              'THERE'S DATA IN SOME RECORDS!
              'ledger_description_max
              'RecurringItem_max
              ' ** If ther's 2nd row data that's visible, but the fields are empty, get rid of their labels.
              ' ** If revcode_DESC and revcode_TYPE_Description are visible,
              ' ** and they're still on the 2nd row, there will always be data in them.
5250          If arr_varRptFld(R_SHOW, lngLedgDescElem) = True And _
                  arr_varRptFld(R_FTOPN, lngLedgDescElem) = lng2ndRowTop And _
                  IsNull(Forms("frmTransaction_Audit").ledger_description_max) = True Then
5260            .Controls(arr_varRptFld(R_LNAM, lngLedgDescElem)).Visible = False
5270            .Controls(arr_varRptFld(R_LNAM, lngLedgDescElem) & "_line").Visible = False
5280          End If
5290          If arr_varRptFld(R_SHOW, lngRecurItemElem) = True And _
                  arr_varRptFld(R_FTOPN, lngRecurItemElem) = lng2ndRowTop And _
                  IsNull(Forms("frmTransaction_Audit").RecurringItem_max) = True Then
5300            .Controls(arr_varRptFld(R_LNAM, lngRecurItemElem)).Visible = False
5310            .Controls(arr_varRptFld(R_LNAM, lngRecurItemElem) & "_line").Visible = False
5320          End If
5330          If arr_varRptFld(R_SHOW, lngLedgHidElem) = True And _
                  arr_varRptFld(R_FTOPN, lngLedgHidElem) = lng2ndRowTop And _
                  Forms("frmTransaction_Audit").ledger_HIDDEN_min = 0 Then
5340            .Controls(arr_varRptFld(R_LNAM, lngLedgHidElem)).Visible = False
5350            .Controls(arr_varRptFld(R_LNAM, lngLedgHidElem) & "_line").Visible = False
5360          End If

              ' ** If there's nothing visible on the 2nd row, move the labels and lower header lines down.
5370          blnFound = False
5380          For lngX = 0& To (lngRptFlds - 1&)
5390            If arr_varRptFld(R_FTOP, lngX) = lng2ndRowTop Then
5400              If arr_varRptFld(R_SHOW, lngX) = True And arr_varRptFld(R_FTOPN, lngX) = lng2ndRowTop Then
                    ' ** It's visible, and it's still on the 2nd row.
                    ' ** Any single hit stops the move.
5410                Select Case lngX
                    Case lngLedgDescElem
5420                  If IsNull(Forms("frmTransaction_Audit").ledger_description_max) = False Then
5430                    blnFound = True
5440                    Exit For
5450                  End If
5460                Case lngRecurItemElem
5470                  If IsNull(Forms("frmTransaction_Audit").RecurringItem_max) = False Then
5480                    blnFound = True
5490                    Exit For
5500                  End If
5510                Case lngLedgHidElem
5520                  If Forms("frmTransaction_Audit").ledger_HIDDEN_min = -1 Then
5530                    blnFound = True
5540                    Exit For
5550                  End If
5560                Case Else
5570                  blnFound = True
5580                  Exit For
5590                End Select
5600              End If
5610            End If
5620          Next
5630          If blnFound = False Then
                ' ** FOR NOW, LEAVE LABELS WHERE THEY WERE!
                'For lngX = 0& To (lngRptFlds - 1&)
                '  If arr_varRptFld(R_LTOP, lngX) = lng1stRowLabelTop Or _
                '      arr_varRptFld(R_LTOP, lngX) = lng1stRowLabelHighTop Then
                '    ' ** 1st row label to be moved down.
                '    .Controls(arr_varRptFld(R_LNAM, lngX)).Top = _
                '      (.Controls(arr_varRptFld(R_LNAM, lngX)).Top + (lng2ndRowLabelTop - lng1stRowLabelTop))
                '  ElseIf arr_varRptFld(R_LTOP, lngX) = lng2ndRowLabelTop And _
                '      arr_varRptFld(R_LTOPN, lngX) = lng1stRowLabelTop Then
                '    ' ** 2nd row label, already moved up to 1st row, to be moved down again.
                '    .Controls(arr_varRptFld(R_LNAM, lngX)).Top = _
                '      (.Controls(arr_varRptFld(R_LNAM, lngX)).Top + (lng2ndRowLabelTop - lng1stRowLabelTop))
                '  End If
                'Next
5640          End If

5650          If frm.chkIncludeSharefaceTot = True Then
5660            arr_varRptFld(R_SHOW, lngShareElem) = True
5670          End If

              'WHAT ABOUT FOREIGN CURRENCY?
              ' ** Check visibility and position of totals, if applicable.
5680          If arr_varRptFld(R_SHOW, lngShareElem) = True And _
                  arr_varRptFld(R_FLFTN, lngShareElem) <> arr_varRptFld(R_FLFT, lngShareElem) Then
5690            .Controls(arr_varRptFld(R_FNAM, lngShareElem) & "_tot").Left = arr_varRptFld(R_FLFTN, lngShareElem)
5700            .Controls(arr_varRptFld(R_FNAM, lngShareElem) & "_tot_line").Left = arr_varRptFld(R_FLFTN, lngShareElem)
5710            If frm.chkIncludeSharefaceTot = True Then
5720              .shareface_tot.Visible = True
5730              .shareface_tot_line.Visible = True
5740            End If
5750          End If
5760          If arr_varRptFld(R_SHOW, lngICashElem) = True And _
                  arr_varRptFld(R_FLFTN, lngICashElem) <> arr_varRptFld(R_FLFT, lngICashElem) Then
5770            .Controls(arr_varRptFld(R_FNAM, lngICashElem) & "_tot").Left = arr_varRptFld(R_FLFTN, lngICashElem)
5780            .Controls(arr_varRptFld(R_FNAM, lngICashElem) & "_tot_line").Left = arr_varRptFld(R_FLFTN, lngICashElem)
5790          End If
5800          If arr_varRptFld(R_SHOW, lngPCashElem) = True And _
                  arr_varRptFld(R_FLFTN, lngPCashElem) <> arr_varRptFld(R_FLFT, lngPCashElem) Then
5810            .Controls(arr_varRptFld(R_FNAM, lngPCashElem) & "_tot").Left = arr_varRptFld(R_FLFTN, lngPCashElem)
5820            .Controls(arr_varRptFld(R_FNAM, lngPCashElem) & "_tot_line").Left = arr_varRptFld(R_FLFTN, lngPCashElem)
5830          End If
5840          If arr_varRptFld(R_SHOW, lngCostElem) = True And _
                  arr_varRptFld(R_FLFTN, lngCostElem) <> arr_varRptFld(R_FLFT, lngCostElem) Then
5850            .Controls(arr_varRptFld(R_FNAM, lngCostElem) & "_tot").Left = arr_varRptFld(R_FLFTN, lngCostElem)
5860            .Controls(arr_varRptFld(R_FNAM, lngCostElem) & "_tot_line").Left = arr_varRptFld(R_FLFTN, lngCostElem)
5870          End If

5880        End If  ' ** lngVisCnt < lngChkCnt.

            ' ** Coordinate the narrow total controls.
5890        For lngX = 0& To (lngNarrows - 1&)
5900          arr_varNarrow(N_OVIS, lngX) = .Controls(arr_varNarrow(N_ONAM, lngX)).Visible
5910          Select Case arr_varNarrow(N_NAM, lngX)
              Case "cost_tot_narrow"
5920            If .cost_tot_narrow.Left <> .cost_tot.Left Then
5930              .cost_tot_narrow.Top = .cost_tot.Top - lngTpp
5940              arr_varNarrow(N_NTOP, lngX) = .cost_tot.Top - lngTpp
5950              .cost_tot_narrow.Left = .cost_tot.Left
5960              arr_varNarrow(N_NLFT, lngX) = .cost_tot.Left
5970            End If
5980          Case "pcash_tot_narrow"
5990            If .pcash_tot_narrow.Left <> .pcash_tot.Left Then
6000              .pcash_tot_narrow.Top = .pcash_tot.Top - lngTpp
6010              arr_varNarrow(N_NTOP, lngX) = .pcash_tot.Top - lngTpp
6020              .pcash_tot_narrow.Left = .pcash_tot.Left
6030              arr_varNarrow(N_NLFT, lngX) = .pcash_tot.Left
6040            End If
6050          Case "icash_tot_narrow"
6060            If .icash_tot_narrow.Left <> .icash_tot.Left Then
6070              .icash_tot_narrow.Top = .icash_tot.Top - lngTpp
6080              arr_varNarrow(N_NTOP, lngX) = .icash_tot.Top - lngTpp
6090              .icash_tot_narrow.Left = .icash_tot.Left
6100              arr_varNarrow(N_NLFT, lngX) = .icash_tot.Left
6110            End If
6120          Case "shareface_tot_narrow"
6130            If .shareface_tot_narrow.Left <> .shareface_tot.Left Then
6140              .shareface_tot_narrow.Top = .shareface_tot.Top - lngTpp
6150              arr_varNarrow(N_NTOP, lngX) = .shareface_tot.Top - lngTpp
6160              .shareface_tot_narrow.Left = .shareface_tot.Left
6170              arr_varNarrow(N_NLFT, lngX) = .shareface_tot.Left
6180            End If
6190          End Select
6200        Next

            ' ** Now make the total fields much wider, pushing to the left.
6210        For lngX = 0& To (lngNarrows - 1&)
6220          Select Case Forms(strCallingForm).chkNarrowFont
              Case True
6230            If arr_varNarrow(N_OVIS, lngX) = True Then
6240              .Controls(arr_varNarrow(N_ONAM, lngX)).Visible = False
6250              .Controls(arr_varNarrow(N_NAM, lngX)).Visible = True
6260              Select Case arr_varNarrow(N_NAM, lngX)
                  Case "cost_tot_narrow"
6270                .cost_tot_narrow.Left = (.cost_tot_narrow.Left - (16& * lngTpp))  ' ** 16 Px, 240 Twips.
6280                .cost_tot_narrow.Width = (.cost_tot_narrow.Width + (16& * lngTpp))
6290                .cost_tot_line.Left = (.cost_tot_line.Left - (16& * lngTpp))
6300                .cost_tot_line.Width = (.cost_tot_line.Width + (16& * lngTpp))
6310              Case "pcash_tot_narrow"
6320                .pcash_tot_narrow.Left = (.pcash_tot_narrow.Left - (32& * lngTpp))  ' ** 32 Px, 480 Twips.
6330                .pcash_tot_narrow.Width = (.pcash_tot_narrow.Width + (16& * lngTpp))
6340                .pcash_tot_line.Left = (.pcash_tot_line.Left - (32& * lngTpp))
6350                .pcash_tot_line.Width = (.pcash_tot_line.Width + (16& * lngTpp))
6360              Case "icash_tot_narrow"
6370                .icash_tot_narrow.Left = (.icash_tot_narrow.Left - (48& * lngTpp))  ' ** 48 Px, 720 Twips.
6380                .icash_tot_narrow.Width = (.icash_tot_narrow.Width + (16& * lngTpp))
6390                .icash_tot_line.Left = (.icash_tot_line.Left - (48& * lngTpp))
6400                .icash_tot_line.Width = (.icash_tot_line.Width + (16& * lngTpp))
6410              Case "shareface_tot_narrow"
6420                .shareface_tot_narrow.Left = (.shareface_tot_narrow.Left - (64& * lngTpp))  ' ** 64 Px, 960 Twips.
6430                .shareface_tot_narrow.Width = (.shareface_tot_narrow.Width + (16& * lngTpp))
6440                .shareface_tot_line.Left = (.shareface_tot_line.Left - (64& * lngTpp))
6450                .shareface_tot_line.Width = (.shareface_tot_line.Width + (16& * lngTpp))
6460              End Select
6470            End If
6480          Case False
6490            If arr_varNarrow(N_OVIS, lngX) = True Then
6500              .Controls(arr_varNarrow(N_ONAM, lngX)).Visible = True
6510              .Controls(arr_varNarrow(N_NAM, lngX)).Visible = False
6520            End If
6530          End Select
6540        Next

6550        If frm.chkIncludeSharefaceTot = True Then
6560          Select Case frm.chkNarrowFont
              Case True
6570            .shareface_tot_narrow.Visible = True
6580          Case False
6590            .shareface_tot.Visible = True
6600          End Select
6610          .shareface_tot_line.Visible = True
6620        End If

            ' ** If the total controls are visible, make sure their label is in the right place.
6630        If (.cost_tot.Visible = True Or .cost_tot_narrow.Visible = True) Or _
                (.pcash_tot.Visible = True Or .pcash_tot_narrow.Visible = True) Or _
                (.icash_tot.Visible = True Or .icash_tot_narrow.Visible = True) Or _
                (.shareface_tot.Visible = True Or .shareface_tot_narrow.Visible = True) Then
6640          If (.shareface_tot_narrow.Left - lngGrandTotLbl_Offset) >= 0& Then
6650            .grand_tot_lbl.Left = (.shareface_tot_narrow.Left - lngGrandTotLbl_Offset)
6660          End If
6670        Else
6680          .grand_tot_lbl.Visible = False
6690        End If

6700        If .journalno.Visible = True Then
6710          .journalno_lbl.Left = .journalno_lbl.Left - 60&
6720          .journalno_lbl.Width = .journalno_lbl.Width + 120&
6730        End If

            ' ** For now, always show the narrow accountno.
6740        .accountno.Visible = False
6750        .accountno_narrow.Visible = True

6760      Else
6770        Cancel = -1
6780      End If
6790    End With

6800    If Cancel = -1 Then
6810      Beep
6820      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
6830    End If

EXITP:
6840    Set ctl = Nothing
6850    Set frm = Nothing
6860    Exit Sub

ERRH:
6870    Forms(strCallingForm).Visible = True
6880    Select Case ERR.Number
        Case Else
6890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6900    End Select
6910    Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim ctl As Access.Control
        Dim lngTmp01 As Long

7010    With Me
7020      lngTmp01 = .AsOf_lbl2.Height
7030      If .Page >= 2 Then
7040        If .Section(acPageHeader).Height <> (lngPageHeaderSection_Height + lngTmp01) Then
7050          .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
7060          For Each ctl In .Section(acPageHeader).Controls
7070            With ctl
7080              Select Case .Name
                  Case "Header_lbl2", "AsOf_lbl2", "Header_hline05", "Header_hline06", "page_header_height_line"
                    ' ** Skip, these don't move or dealt with separately.
7090              Case Else
7100                If .Visible = True Then
7110                  If Right(.Name, 4) = "_lbl" Then
7120                    Select Case .Top
                        Case lngOneLine1_Top
7130                      .Top = (lngOneLine1_Top + lngTmp01)
7140                    Case lngTwoLine1_Top
7150                      .Top = (lngTwoLine1_Top + lngTmp01)
7160                    Case lngOneLine2_Top
7170                      .Top = (lngOneLine2_Top + lngTmp01)
7180                    Case lngNarrow1_Top
7190                      .Top = (lngNarrow1_Top + lngTmp01)
7200                    Case lngNarrow2_Top
7210                      .Top = (lngNarrow2_Top + lngTmp01)
7220                    End Select
7230                  ElseIf Right(.Name, 5) = "_line" Then
7240                    .Top = (lngUnderline_Top + lngTmp01)
7250                  Else
                        ' ** Anthing else?
7260                  End If
7270                End If
7280              End Select
7290            End With
7300          Next
7310          .Header_hline06.Top = (lngHeaderLine6_Top + lngTmp01)
7320          .Header_lbl2.Visible = True
7330          .AsOf_lbl2.Visible = True
7340          .Header_hline05.Visible = True
7350        End If
7360      Else
7370        For Each ctl In .Section(acPageHeader).Controls
7380          With ctl
7390            Select Case .Name
                Case "Header_lbl2", "AsOf_lbl2", "Header_hline05", "Header_hline06", "page_header_height_line"
                  ' ** Skip, these don't move or dealt with separately.
7400            Case Else
7410              If .Visible = True Then
7420                If Right(.Name, 4) = "_lbl" Then
7430                  If .Top = (lngOneLine1_Top + lngTmp01) Then
7440                    .Top = lngOneLine1_Top
7450                  ElseIf .Top = (lngTwoLine1_Top + lngTmp01) Then
7460                    .Top = lngTwoLine1_Top
7470                  ElseIf .Top = (lngOneLine2_Top + lngTmp01) Then
7480                    .Top = lngOneLine2_Top
7490                  ElseIf .Top = (lngNarrow1_Top + lngTmp01) Then
7500                    .Top = lngNarrow1_Top
7510                  ElseIf .Top = (lngNarrow2_Top + lngTmp01) Then
7520                    .Top = lngNarrow2_Top
7530                  End If
7540                ElseIf Right(.Name, 5) = "_line" Then
7550                  .Top = lngUnderline_Top
7560                Else
                      ' ** Anthing else?
7570                End If
7580              End If
7590            End Select
7600          End With
7610        Next
7620        .Header_hline06.Top = lngHeaderLine6_Top
7630        .Header_lbl2.Visible = False
7640        .AsOf_lbl2.Visible = False
7650        .Header_hline05.Visible = False
7660        .Section(acPageHeader).Height = lngPageHeaderSection_Height
7670      End If
7680    End With

EXITP:
7690    Set ctl = Nothing
7700    Exit Sub

ERRH:
7710    Forms(strCallingForm).Visible = True
7720    Select Case ERR.Number
        Case Else
7730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7740    End Select
7750    Resume EXITP

End Sub

Private Sub Report_Close()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

        Dim blnIsMaximized As Boolean

7810    DoCmd.Restore

7820    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

7830    blnIsMaximized = Forms("frmTransaction_Audit").FrmIsMax  ' ** Form Function: frmTransaction_Audit.
7840    If blnIsMaximized = False Then
7850      gblnSetFocus = True
7860      Forms("frmTransaction_Audit").TimerInterval = 100&  ' ** 100 thousandths of a second (1/10th of a sec).
7870    End If
7880    If Forms(strCallingForm).Visible = False Then
7890      Forms(strCallingForm).Visible = True
7900    End If

EXITP:
7910    Exit Sub

ERRH:
7920    Forms(strCallingForm).Visible = True
7930    Select Case ERR.Number
        Case Else
7940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7950    End Select
7960    Resume EXITP

End Sub

Private Sub RptFldSort()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "RptFldSort"

        Dim varTmp00 As Variant
        Dim lngX As Long, lngY As Long, lngZ As Long

        ' **************************************************
        ' ** Array: arr_varRptFld()
        ' **
        ' **   Element  Name               Constant
        ' **   =======  =================  ===============
        ' **      0     Field Name         R_FNAM
        ' **      1     Field Top          R_FTOP
        ' **      2     Field Left         R_FLFT
        ' **      3     Field Width        R_FWDT
        ' **      4     Field Visible      R_FVIS
        ' **      5     Field New Top      R_FTOPN
        ' **      6     Field New Left     R_FLFTN
        ' **      7     Field New Width    R_FWDTN
        ' **      8     Label Name         R_LNAM
        ' **      9     Label Top          R_LTOP
        ' **     10     Label Left         R_LLFT
        ' **     11     Label Width        R_LWDT
        ' **     12     Label New Top      R_LTOPN
        ' **     13     Label New Left     R_LLFTN
        ' **     14     Label New Width    R_LWDTN
        ' **     15     Show/Hide          R_SHOW
        ' **     16     Sorting Field      R_SORT
        ' **
        ' **************************************************

        ' ** Binary Sort arr_varRptFld() array by position, left to right, top to bottom.
8010    For lngX = (lngRptFlds - 1&) To 1 Step -1
8020      For lngY = 0 To (lngX - 1)
8030        If arr_varRptFld(R_SORT, lngY) > arr_varRptFld(R_SORT, (lngY + 1)) Then
8040          For lngZ = 0& To R_ELEMS
8050            varTmp00 = arr_varRptFld(lngZ, lngY)
8060            arr_varRptFld(lngZ, lngY) = arr_varRptFld(lngZ, (lngY + 1))
8070            arr_varRptFld(lngZ, (lngY + 1)) = varTmp00
8080            varTmp00 = Empty
8090          Next
8100        End If
8110      Next
8120    Next

        'For lngX = 0& To (lngRptFlds - 1&)
        '  Debug.Print "'" & Left(CStr((lngX + 1&)) & "  ", 2) & ": " & arr_varRptFld(R_FNAM, lngX)
        'Next

        ' ** Array: arr_varRptFld().
        ' **   1 : journalno
        ' **   2 : journaltype
        ' **   3 : transdate
        ' **   4 : accountno
        ' **   5 : shortname
        ' **   6 : cusip
        ' **   7 : asset_description
        ' **   8 : shareface
        ' **   9 : icash
        ' **   10: pcash
        ' **   11: cost
        ' **   12: assetdate
        ' **   13: PurchaseDate
        ' **   14: ledger_description
        ' **   15: RecurringItem
        ' **   16: revcode_DESC
        ' **   17: revcode_TYPE_Description
        ' **   18: journal_USER
        ' **   19: posted
        ' **   20: ledger_HIDDEN

        ' ** Array: arr_varFrmFld().
        ' **   1 : journalno
        ' **   2 : journaltype
        ' **   3 : transdate
        ' **   4 : accountno
        ' **   5 : shortname
        ' **   6 : cusip
        ' **   7 : asset_description
        ' **   8 : shareface
        ' **   9 : icash
        ' **   10: pcash
        ' **   11: cost
        ' **   12: assetdate
        ' **   13: PurchaseDate
        ' **   14: ledger_description
        ' **   15: RecurringItem
        ' **   16: revcode_DESC
        ' **   17: revcode_TYPE_Description
        ' **   18: journal_USER
        ' **   19: posted
        ' **   20: ledger_HIDDEN

EXITP:
8130    Exit Sub

ERRH:
8140    Forms(strCallingForm).Visible = True
8150    Select Case ERR.Number
        Case Else
8160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8170    End Select
8180    Resume EXITP

End Sub
