Attribute VB_Name = "modTransactionAuditFuncs2"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modTransactionAuditFuncs2"

'VGC 10/05/2017: CHANGES!

' ** Array: arr_varFilt(), arr_varFilt_ds().
Private lngFilts As Long, arr_varFilt As Variant
Private lngFilts_ds As Long, arr_varFilt_ds As Variant
'Private Const F_ELEMS As Integer = 13  ' ** Array's first-element UBound().
'Private Const F_IDX   As Integer = 0
'Private Const F_NAM   As Integer = 1
Private Const F_CONST As Integer = 2
'Private Const F_CTL   As Integer = 3
'Private Const F_CLBL  As Integer = 4
'Private Const F_FLD   As Integer = 5
'Private Const F_FLBL  As Integer = 6
'Private Const F_CTL2  As Integer = 7
'Private Const F_CLBL2 As Integer = 8
'Private Const F_FLD2  As Integer = 9
'Private Const F_FLBL2 As Integer = 10
'Private Const F_CLBL3 As Integer = 11
'Private Const F_FLD3  As Integer = 12
'Private Const F_FLBL3 As Integer = 13

' ** Filter constants: frmTransaction_Audit_Sub.
Private Const ANDF         As String = " And "  ' ** Filter 'And'.
Private Const ORF          As String = " Or "  ' ** Filter 'Or'.
'Private Const JRNL_NUM     As String = "[journalno] = "
Private Const JRNL_TYPE    As String = "[journaltype] = '"
'Private Const TRANS_START  As String = "[transdate] >= #"
'Private Const TRANS_END    As String = "[transdate] <= #"
'Private Const ACCT_NUM     As String = "[accountno] = '"
'Private Const ASSET_NUM    As String = "[assetno] = "
'Private Const CURR_NUM     As String = "[curr_id] = "
'Private Const ASSET_START  As String = "[assetdate] >= #"
'Private Const ASSET_END    As String = "[assetdate] <= #"
'Private Const PURCH_START  As String = "[PurchaseDate] >= #"
'Private Const PURCH_END    As String = "[PurchaseDate] <= #"
'Private Const COMM_DESC    As String = "[ledger_description] Like '*"
'Private Const RECUR_ITEM   As String = "[RecurringItem] Like '*"
'Private Const REV_CODE     As String = "[revcode_ID] = "
'Private Const TAX_CODE     As String = "[taxcode] = "
'Private Const LOC_NUM      As String = "[Location_ID] = "
'Private Const CHK_NUM      As String = "[CheckNum] = "
'Private Const CHK_NUM1     As String = "[CheckNum] >= "
'Private Const CHK_NUM2     As String = "[CheckNum] <= "
'Private Const JRNL_USER    As String = "[journal_USER] = '"
'Private Const POSTED_START As String = "[posted] >= #"
'Private Const POSTED_END   As String = "[posted] <= #"
'Private Const HIDDEN_TRX1  As String = "[ledger_HIDDEN] = True"
'Private Const HIDDEN_TRX2  As String = "[ledger_HIDDEN] = False"

' ** Array: arr_varFrmFld(), arr_varFrmFld_ds().
'Private lngFrmFlds As Long, arr_varFrmFld() As Variant
'Private lngFrmFlds_ds As Long, arr_varFrmFld_ds() As Variant
Private Const FM_ELEMS As Integer = 7  ' ** Array's first-element UBound().
Private Const FM_FLD_NAM As Integer = 0
Private Const FM_FLD_TAB As Integer = 1
Private Const FM_FLD_VIS As Integer = 2
Private Const FM_CHK_NAM As Integer = 3
Private Const FM_CHK_VAL As Integer = 4
Private Const FM_VIEWCHK As Integer = 5
Private Const FM_TOPO    As Integer = 6
Private Const FM_TOPC    As Integer = 7

' ** Array: arr_varCal().
Private Const C_CNAM  As Integer = 0
Private Const C_FOCUS As Integer = 1
Private Const C_DOWN  As Integer = 2
Private Const C_ABLE  As Integer = 3
Private Const C_FLD   As Integer = 4

Private lngTpp As Long
' **

Public Sub JType_After_TA(strProc As String, strFilter01 As String, strFilter02 As String, dblFilterRecs As Double, rstAll1 As DAO.Recordset, frmCrit As Access.Form)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "JType_After_TA"

        Dim frm As Access.Form
        Dim strEvent As String, strCtlName As String
        Dim lngMultiCnt As Long, lngCnt As Long
        Dim blnErr As Boolean
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer, intCnt As Integer
        Dim lngX As Long

110     With frmCrit

120       DoCmd.Hourglass True
130       DoEvents
140       lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
150       intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
160       strEvent = Mid(strProc, (intPos01 + 1))
170       strCtlName = Left(strProc, (intPos01 - 1))

180       Set frm = .Parent

190       Select Case strCtlName
          Case "cmbJournalType1"

200         If IsNull(.cmbJournalType1) = False Then
210           If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
                ' ** This will be the only clause.
220             strFilter01 = "(" & JRNL_TYPE & .cmbJournalType1 & "')"
230             strFilter02 = "(" & JRNL_TYPE & .cmbJournalType1 & "')"
240             .cmbJournalType2.Enabled = True
250             .cmbJournalType2.BorderColor = CLR_LTBLU2
260             .cmbJournalType2.BackStyle = acBackStyleNormal
270           Else
                ' ** There are clauses present.

280             .FilterRec_GetArr  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.

290             arr_varFilt = FilterRecs_GetArr(.Parent.opgView_optForm.OptionValue)  ' ** Module Function: modTrasnsactionAuditFuncs1.
300             lngFilts = UBound(arr_varFilt, 2)
310             arr_varFilt_ds = FilterRecs_GetArr(.Parent.opgView_optDatasheet.OptionValue)  ' ** Module Function: modTrasnsactionAuditFuncs1.
320             lngFilts_ds = UBound(arr_varFilt_ds, 2)

330             intPos01 = InStr(strFilter01, JRNL_TYPE)
340             If intPos01 = 0& Then
                  ' ** This clause isn't present.
350               intPos03 = 0&: intPos04 = 0&
360               For lngX = (lngFilts - 1&) To 0& Step -1&
370                 If arr_varFilt(F_CONST, lngX) = JRNL_TYPE Then
380                   intPos04 = -1
390                 ElseIf intPos04 = -1 Then
                      ' ** Look for the next previous clause present in strFilter01.
400                   intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
410                   If intPos03 > 0 Then
420                     intPos04 = 0&
430                     Exit For
440                   End If
450                 End If
460               Next
470               If intPos03 = 0& Then
                    ' ** Add this clause at the start of the filter.
480                 strFilter01 = "(" & JRNL_TYPE & .cmbJournalType1 & "')" & ANDF & strFilter01
490               Else
                    ' ** There's a clause before this one.
500                 intPos02 = InStr(intPos03, strFilter01, ANDF)
510                 If intPos02 = 0 Then
                      ' ** Add this clause to the end of the filter.
520                   strFilter01 = strFilter01 & ANDF & "(" & JRNL_TYPE & .cmbJournalType1 & "')"
530                 Else
                      ' ** Add this clause to the middle of the filter.
540                   strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & "(" & JRNL_TYPE & .cmbJournalType1 & "')" & Mid(strFilter01, intPos02)
550                 End If
560               End If
570               .cmbJournalType2.Enabled = True
580               .cmbJournalType2.BorderColor = CLR_LTBLU2
590               .cmbJournalType2.BackStyle = acBackStyleNormal
600             Else
610               lngMultiCnt = CharCnt(strFilter01, JRNL_TYPE, True)  ' ** Module Function: modStringFuncs.
                  ' ** Because this is cmbJournalType1, it will always replace a single instance, or the 1st in a series.
620               Select Case lngMultiCnt
                  Case 1&
                    ' ** Replace this clause, whether or not it's the last one.
                    ' ** Note: If intPos01 = 1, then the Left() function is OK with
                    ' ** returning the left 0 characters (as long as it doesn't go below 0).
630                 intPos02 = InStr((intPos01 + Len(JRNL_TYPE) + 5), strFilter01, "'")  ' ** Find the closing quote.
640                 If intPos02 > 0 Then
650                   strFilter01 = Left(strFilter01, (intPos01 - 1)) & JRNL_TYPE & .cmbJournalType1 & Mid(strFilter01, (intPos02 + 1))
660                 Else
                      ' ** No quote?
670                 End If
680               Case 2&, 3&
                    ' ** Replace just the 1st occurance of the JournalType clauses.
                    ' ** The multiple Journal Type clauses need to be enclosed within parens (because they're 'Or').
690                 intPos02 = InStr(strFilter01, ORF)  ' ** Find the ' Or '.
700                 If intPos02 > 0 Then
710                   strFilter01 = Left(strFilter01, (intPos01 - 1)) & JRNL_TYPE & .cmbJournalType1 & "'" & Mid(strFilter01, intPos02)
                      ' ** Left() ends with '(', and there's an ' Or ' right after the closing quote.
720                 Else
                      ' ** There should never NOT be a closing quote!
730                 End If
740               End Select
750               .cmbJournalType2.Enabled = True
760               .cmbJournalType2.BorderColor = CLR_LTBLU2
770               .cmbJournalType2.BackStyle = acBackStyleNormal
780               Select Case IsNull(.cmbJournalType2)
                  Case True
790                 .cmbJournalType3 = Null  ' ** Should already be Null.
800                 .cmbJournalType3.Enabled = False
810                 .cmbJournalType3.BorderColor = WIN_CLR_DISR
820                 .cmbJournalType3.BackStyle = acBackStyleTransparent
830               Case False
840                 .cmbJournalType3.Enabled = True
850                 .cmbJournalType3.BorderColor = CLR_LTBLU2
860                 .cmbJournalType3.BackStyle = acBackStyleNormal
870               End Select
880             End If
890           End If
              ' ** It sometimes seems to lose the closing single quote for journaltype's!
900           intCnt = CharCnt(strFilter01, "'")
910           If intCnt > 0 Then
                '([journaltype] = 'Paid) And [transdate] >= #01/01/2013# And [transdate] <= #12/31/2013# And [assetno] = 20
920             If intCnt Mod 2 <> 0 Then  ' ** If it's odd, one's missing!
930               intPos01 = InStr(strFilter01, "[journaltype]")
940               If intPos01 > 0 Then
950                 intPos02 = InStr(intPos01, strFilter01, ")")
960                 If intPos02 > 0 Then
970                   If Mid(strFilter01, (intPos02 - 1), 1) <> "'" Then
980                     strFilter01 = Left(strFilter01, (intPos02 - 1)) & "'" & Mid(strFilter01, intPos02)
                        'Stop
990                   End If
1000                End If
1010              End If
1020            End If
1030          End If
1040          strFilter02 = strFilter01
1050          frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
1060          frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
1070          frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
1080          frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
1090          frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
1100          frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
              '.cmbJournalType2.SetFocus
              'Select Case frm.opgView
              'Case frm.opgView_optForm.OptionValue
              '  frm.frmTransaction_Audit_Sub.Form.journalno_GotFocus  ' ** Form Procedure: frmTransaction_Audit_Sub.
              'Case frm.opgView_optDatasheet.OptionValue
              '  frm.frmTransaction_Audit_Sub_ds.Form.Journal_Number_GotFocus  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
              'End Select
1110          DoEvents
1120        Else
              ' ** If there's only 1 occurance, remove it.
              ' ** If there are more than 1, remove the first only, then close up the remaining 2,
              ' ** and move 2 to cmbJournalType1, and 3 to cmbJournalType2.
1130          lngMultiCnt = CharCnt(strFilter01, JRNL_TYPE, True)  ' ** Module Function: modStringFuncs.
1140          intPos01 = InStr(strFilter01, JRNL_TYPE)
1150          If intPos01 > 0 Then
1160            Select Case lngMultiCnt  ' ** Count before removing this one.
                Case 1&
1170              intPos02 = InStr((intPos01 + Len(JRNL_TYPE) + 5), strFilter01, "'")  ' ** Find the closing quote.
1180              If intPos02 > 0 Then
1190                If intPos01 = 2 Then
                      ' ** Beginning of filter (with paren).
1200                  intPos03 = InStr(strFilter01, ANDF)
1210                  If intPos03 = 0 Then
                        ' ** Only clause in filter
1220                    strFilter01 = vbNullString
1230                  Else
1240                    strFilter01 = Mid(strFilter01, (intPos02 + 2))  ' ** Remove parens as well.
1250                  End If
1260                Else
                      ' ** There's a clause before this.
1270                  intPos03 = InStr(strFilter01, ANDF)
1280                  If intPos03 = 0 Then
                        ' ** Nothing after.
1290                    strFilter01 = Left(strFilter01, (intPos01 - 2))  ' ** Remove parens as well.
1300                  Else
1310                    strFilter01 = Left(strFilter01, (intPos01 - 2)) & Mid(strFilter01, (intPos02 + 2))  ' ** Remove parens as well.
1320                  End If
1330                End If
1340              Else
                    ' ** Screwed up.
1350              End If
1360              strFilter02 = strFilter01
1370              .cmbJournalType2.Enabled = False     ' ** When cmbJournalType1 is empty, 2 and 3 are disabled.
1380              .cmbJournalType2.BorderColor = WIN_CLR_DISR
1390              .cmbJournalType2.BackStyle = acBackStyleTransparent
1400              .cmbJournalType3.Enabled = False
1410              .cmbJournalType3.BorderColor = WIN_CLR_DISR
1420              .cmbJournalType3.BackStyle = acBackStyleTransparent
1430              frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub.
1440              frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
                  '.TransDateStart.SetFocus
1450            Case 2&
1460              intPos02 = InStr(intPos01, strFilter01, ORF)  ' ** Find the ' Or '.
1470              If intPos02 > 0 Then
1480                strFilter01 = Left(strFilter01, (intPos01 - 1)) & Mid(strFilter01, (intPos02 + Len(ORF)))  ' ** Opening paren to start of 2nd clause.
1490              Else
                    ' ** Yikes!
1500              End If
1510              strFilter02 = strFilter01
1520              .cmbJournalType1 = .cmbJournalType2  ' ** As long as cmbJournalType1 has something, 2 remains enabled.
1530              .cmbJournalType2 = Null
1540              .cmbJournalType2.Enabled = True
1550              .cmbJournalType2.BorderColor = CLR_LTBLU2
1560              .cmbJournalType2.BackStyle = acBackStyleNormal
1570              .cmbJournalType3.Enabled = False     ' ** When cmbJournalType2 is empty, 3 is disabled.
1580              .cmbJournalType3.BorderColor = WIN_CLR_DISR
1590              .cmbJournalType3.BackStyle = acBackStyleTransparent
1600              frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr "cmbJournalType2_AfterUpdate", False  ' ** Form Procedure: frmTransaction_Audit_Sub.
1610              frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
1620              frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr "cmbJournalType2_AfterUpdate", False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
1630              frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
                  '.cmbJournalType2.SetFocus
1640            Case 3&
1650              intPos02 = InStr(intPos01, strFilter01, ORF)  ' ** Find the ' Or '.
1660              If intPos02 > 0 Then
1670                strFilter01 = Left(strFilter01, (intPos01 - 1)) & Mid(strFilter01, (intPos02 + Len(ORF)))  ' ** Opening paren to start of 2nd clause.
1680              Else
                    ' ** Yikes!
1690              End If
1700              strFilter02 = strFilter01
1710              .cmbJournalType1 = .cmbJournalType2  ' ** As long as cmbJournalType1 and 2 have something, all remain enabled.
1720              .cmbJournalType2 = .cmbJournalType3
1730              .cmbJournalType3 = Null
1740              frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr "cmbJournalType3_AfterUpdate", False  ' ** Form Procedure: frmTransaction_Audit_Sub.
1750              frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr "cmbJournalType2_AfterUpdate", True  ' ** Form Procedure: frmTransaction_Audit_Sub.
1760              frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
1770              frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr "cmbJournalType3_AfterUpdate", False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
1780              frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr "cmbJournalType2_AfterUpdate", True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
1790              frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
                  '.cmbJournalType2.SetFocus
1800            End Select
1810          Else
                ' ** Shouldn't have been here.
1820          End If
              ' ** It sometimes seems to lose the closing single quote for journaltype's!
              '([journaltype] = 'Paid) And [transdate] >= #01/01/2013# And [transdate] <= #12/31/2013# And [assetno] = 20
1830          intCnt = CharCnt(strFilter01, "'")
1840          If intCnt > 0 Then
1850            If intCnt Mod 2 <> 0 Then  ' ** If it's odd, one's missing!
1860              intPos01 = InStr(strFilter01, "[journaltype]")
1870              If intPos01 > 0 Then
1880                intPos02 = InStr(intPos01, strFilter01, ")")
1890                If intPos02 > 0 Then
1900                  If Mid(strFilter01, (intPos02 - 1), 1) <> "'" Then
1910                    strFilter01 = Left(strFilter01, (intPos02 - 1)) & "'" & Mid(strFilter01, intPos02)
                        'Stop
1920                  End If
1930                End If
1940              End If
1950            End If
1960          End If
1970          frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
1980          frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
1990          frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
2000          frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
2010          DoEvents
              'Select Case frm.opgView
              'Case frm.opgView_optForm.OptionValue
              '  frm.frmTransaction_Audit_Sub.Form.journalno_GotFocus  ' ** Form Procedure: frmTransaction_Audit_Sub.
              'Case frm.opgView_optDatasheet.OptionValue
              '  frm.frmTransaction_Audit_Sub_ds.Form.Journal_Number_GotFocus  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
              'End Select
              'DoEvents
2020          frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
2030        End If  ' ** cmbJournalType1.

2040      Case "cmbJournalType2"

2050        blnErr = False
2060        If IsNull(.cmbJournalType2) = False Then
2070          If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
                ' ** Then we shouldn't be here!
2080          Else
2090            intPos01 = InStr(strFilter01, JRNL_TYPE)
2100            If intPos01 > 0& Then
2110              lngMultiCnt = CharCnt(strFilter01, JRNL_TYPE, True)  ' ** Module Function: modStringFuncs.
2120              Select Case lngMultiCnt
                  Case 1&
                    ' ** Add this one after the 1st.
2130                intPos02 = InStr((intPos01 + Len(JRNL_TYPE) + 5), strFilter01, "'")  ' ** Find the closing quote.
2140                If intPos02 > 0 Then
2150                  strFilter01 = Left(strFilter01, intPos02) & ORF & JRNL_TYPE & .cmbJournalType2 & "'" & Mid(strFilter01, (intPos02 + 1))
                      ' ** Left() ends just after 1's closing quote, so slip it in between there and the closing paren.
2160                Else
                      ' ** It should never be without the closing paren.
2170                  blnErr = True
2180                End If
2190              Case 2&
                    ' ** Replace just the 2nd occurance of the JournalType clauses.
                    ' ** The multiple Journal Type clauses need to be enclosed within parens (because they're 'Or').
2200                intPos01 = InStr(intPos01, strFilter01, ORF)  ' ** Find the ' Or '.
2210                intPos02 = InStr((intPos01 + Len(ORF) + 5), strFilter01, "'")  ' ** Find the closing quote.
2220                If intPos02 > 0 Then
2230                  strFilter01 = Left(strFilter01, (intPos01 + Len(ORF))) & JRNL_TYPE & .cmbJournalType2 & "'" & Mid(strFilter01, (intPos02 + 1))
                      ' ** Left() ends with ' Or ', and Mid() begins with the closing paren.
2240                Else
                      ' ** There should never NOT be a closing quote!
2250                  blnErr = True
2260                End If
2270              Case 3&
                    ' ** Replace the 2nd of 3 occurances.
2280                intPos01 = InStr(intPos01, strFilter01, ORF)  ' ** Find the 1st ' Or '.
2290                If intPos01 > 0 Then
2300                  intPos02 = InStr((intPos01 + Len(ORF)), strFilter01, ORF)  ' ** Find the 2nd ' Or '.
2310                  If intPos02 > 0 Then
2320                    strFilter01 = Left(strFilter01, ((intPos01 + Len(ORF)) - 1)) & JRNL_TYPE & .cmbJournalType2 & "'" & Mid(strFilter01, intPos02)
2330                  Else
                        ' ** 3 clauses with only 1 'Or'?
2340                    blnErr = True
2350                  End If
2360                Else
                      ' ** 3 clauses without even 1 'Or'?
2370                  blnErr = True
2380                End If
2390              End Select
                  ' ** It sometimes seems to lose the closing single quote for journaltype's!
2400              intCnt = CharCnt(strFilter01, "'")
2410              If intCnt > 0 Then
2420                If intCnt Mod 2 <> 0 Then
                      'Stop
2430                End If
2440              End If
2450              strFilter02 = strFilter01
2460              If blnErr = False Then
2470                .cmbJournalType2.Enabled = True
2480                .cmbJournalType2.BorderColor = CLR_LTBLU2
2490                .cmbJournalType2.BackStyle = acBackStyleNormal
2500                Select Case IsNull(.cmbJournalType2)
                    Case True
2510                  .cmbJournalType3 = Null  ' ** Should already be Null.
2520                  .cmbJournalType3.Enabled = False
2530                  .cmbJournalType3.BorderColor = WIN_CLR_DISR
2540                  .cmbJournalType3.BackStyle = acBackStyleTransparent
2550                Case False
2560                  .cmbJournalType3.Enabled = True
2570                  .cmbJournalType3.BorderColor = CLR_LTBLU2
2580                  .cmbJournalType3.BackStyle = acBackStyleNormal
2590                End Select
2600              End If
2610            Else
                  ' ** If this clause isn't present, we shouldn't be here!
2620              blnErr = True
2630            End If
2640          End If
2650          If blnErr = False Then
2660            frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
2670            frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
2680            frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
2690            frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
2700            frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
2710            frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
                '.cmbJournalType2.SetFocus
                'Select Case frm.opgView
                'Case frm.opgView_optForm.OptionValue
                '  frm.frmTransaction_Audit_Sub.Form.journalno_GotFocus  ' ** Form Procedure: frmTransaction_Audit_Sub.
                'Case frm.opgView_optDatasheet.OptionValue
                '  frm.frmTransaction_Audit_Sub_ds.Form.Journal_Number_GotFocus  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
                'End Select
2720            DoEvents
2730          End If
2740        Else
              ' ** Remove the 2nd occurance, and if there were 3, move 3 to 2.
2750          lngMultiCnt = CharCnt(strFilter01, JRNL_TYPE, True)  ' ** Module Function: modStringFuncs.
2760          intPos01 = InStr(strFilter01, JRNL_TYPE)
2770          If intPos01 > 0 Then
2780            Select Case lngMultiCnt
                Case 1&
                  ' ** Shouldn't have been here to begin with.
                  '.TransDateStart.SetFocus
2790              .cmbJournalType3 = Null  ' ** Should already be Null.
2800              .cmbJournalType3.Enabled = False
2810              .cmbJournalType3.BorderColor = WIN_CLR_DISR
2820              .cmbJournalType3.BackStyle = acBackStyleTransparent
2830            Case 2&
2840              intPos01 = InStr(intPos01, strFilter01, ORF)  ' ** Find the ' Or '.
2850              If intPos01 > 0 Then
2860                intPos02 = InStr((intPos01 + Len(ORF) + Len(JRNL_TYPE) + 5), strFilter01, "'")  ' ** Find the closing quote.
2870                If intPos02 > 0 Then
2880                  strFilter01 = Left(strFilter01, (intPos01 - 1)) & Mid(strFilter01, (intPos02 + 1))
2890                Else
                      ' ** No closing paren means no cheese.
2900                  blnErr = True
2910                End If
2920              Else
                    ' ** Messed up.
2930                blnErr = True
2940              End If
2950            Case 3&
2960              intPos01 = InStr(intPos01, strFilter01, ORF)  ' ** Find the 1st ' Or '.
2970              If intPos01 > 0 Then
2980                intPos02 = InStr((intPos01 + Len(ORF)), strFilter01, ORF)  ' ** Find the 2nd ' Or '.
2990                If intPos02 > 0 Then
3000                  strFilter01 = Left(strFilter01, (intPos01 - 1)) & Mid(strFilter01, intPos02)
3010                Else
                      ' ** Snafu.
3020                  blnErr = True
3030                End If
3040              Else
                    ' ** Disorder.
3050                blnErr = True
3060              End If
3070            End Select
3080          Else
                ' ** Unlikely to have gotten here.
3090            blnErr = True
3100          End If
              ' ** It sometimes seems to lose the closing single quote for journaltype's!
3110          intCnt = CharCnt(strFilter01, "'")
3120          If intCnt > 0 Then
3130            If intCnt Mod 2 <> 0 Then
                  'Stop
3140            End If
3150          End If
3160          strFilter02 = strFilter01
3170          If blnErr = False Then
3180            frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
3190            frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
3200            frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
3210            frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
3220            Select Case lngMultiCnt  ' ** Count before removing this one.
                Case 1&
                  ' ** Shouldn't have even gotten here!
                  '.TransDateStart.SetFocus
3230              frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub.
3240              frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
3250            Case 2&
3260              .cmbJournalType3 = Null  ' ** Should already be empty.
3270              .cmbJournalType3.Enabled = False  ' ** When cmbJournalType2 is empty, 3 is disabled.
3280              .cmbJournalType3.BorderColor = WIN_CLR_DISR
3290              .cmbJournalType3.BackStyle = acBackStyleTransparent
                  '.TransDateStart.SetFocus
3300              frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub.
3310              frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
3320            Case 3&
3330              .cmbJournalType2 = .cmbJournalType3
3340              .cmbJournalType3 = Null
                  '.cmbJournalType3.SetFocus
3350              frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
3360              frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr "cmbJournalType3_AfterUpdate", False  ' ** Form Procedure: frmTransaction_Audit_Sub.
3370              frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
3380              frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr "cmbJournalType3_AfterUpdate", False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
3390            End Select
                'Select Case frm.opgView
                'Case frm.opgView_optForm.OptionValue
                '  frm.frmTransaction_Audit_Sub.Form.journalno_GotFocus  ' ** Form Procedure: frmTransaction_Audit_Sub.
                'Case frm.opgView_optDatasheet.OptionValue
                '  frm.frmTransaction_Audit_Sub_ds.Form.Journal_Number_GotFocus  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
                'End Select
3400            DoEvents
3410            frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
3420          End If
3430        End If  ' ** cmbJournalType2.
3440        If blnErr = True Then
3450          Beep
3460          MsgBox "A problem occurred assembling your criteria." & vbCrLf & _
                "Please reenter your Journal Type parameters.", vbInformation + vbOKOnly, "Reenter Journal Types"
3470          .cmbJournalType1 = Null
3480          .cmbJournalType1.SetFocus
3490          .cmbJournalType2 = Null
3500          .cmbJournalType2.Enabled = False
3510          .cmbJournalType2.BorderColor = WIN_CLR_DISR
3520          .cmbJournalType2.BackStyle = acBackStyleTransparent
3530          .cmbJournalType3 = Null
3540          .cmbJournalType3.Enabled = False
3550          .cmbJournalType3.BorderColor = WIN_CLR_DISR
3560          .cmbJournalType3.BackStyle = acBackStyleTransparent
3570        End If

3580      Case "cmbJournalType3"

3590        blnErr = False
3600        If IsNull(.cmbJournalType3) = False Then
3610          If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
                ' ** Then we shouldn't be here!
3620          Else
3630            intPos01 = InStr(strFilter01, JRNL_TYPE)
3640            If intPos01 > 0& Then
3650              lngMultiCnt = CharCnt(strFilter01, JRNL_TYPE, True)  ' ** Module Function: modStringFuncs.
3660              Select Case lngMultiCnt
                  Case 1&
                    ' ** Shouldn't be here!
3670                blnErr = True
3680              Case 2&
                    ' ** Add this after the 2nd one.
                    ' ** The multiple Journal Type clauses need to be enclosed within parens (because they're 'Or').
3690                intPos01 = InStr(intPos01, strFilter01, ORF)  ' ** Find the ' Or '.
                    'Starting from the first JournalType clause, find the ' Or '
3700                If intPos01 > 0 Then
                      'The First, and currently only, ' Or ' was found.
3710                  intPos02 = InStr((intPos01 + Len(ORF) + Len(JRNL_TYPE) + 2), strFilter01, "'")  ' ** Find the closing quote.
                      'Because JRNL_TYPE includes the opening single-quote, don't move too far right for the closing one.
3720                  If intPos02 > 0 Then
3730                    strFilter01 = Left(strFilter01, intPos02) & ORF & JRNL_TYPE & .cmbJournalType3 & "'" & Mid(strFilter01, (intPos02 + 1))
3740                  Else
                        ' ** There should never NOT be a closing quote!
3750                    blnErr = True
3760                  End If
3770                Else
                      ' ** Oh-Oh!
3780                  blnErr = True
3790                End If
3800              Case 3&
                    ' ** Replace the 3rd occurance.
3810                intPos01 = InStr(intPos01, strFilter01, ORF)  ' ** Find the 1st ' Or '.
3820                If intPos01 > 0 Then
3830                  intPos01 = InStr((intPos01 + Len(ORF)), strFilter01, ORF)  ' ** Find the 2nd ' Or '.
3840                  If intPos01 > 0 Then
3850                    intPos02 = InStr((intPos01 + Len(ORF) + Len(JRNL_TYPE) + 5), strFilter01, "'")  ' ** Find the closing quote.
3860                    If intPos02 > 0 Then
3870                      strFilter01 = Left(strFilter01, (intPos01 + Len(ORF))) & JRNL_TYPE & .cmbJournalType3 & "'" & Mid(strFilter01, (intPos02 + 1))
3880                    Else
                          ' ** Oops!
3890                      blnErr = True
3900                    End If
3910                  Else
                        ' ** 3 clauses with only 1 'Or'?
3920                    blnErr = True
3930                  End If
3940                Else
                      ' ** 3 clauses without even 1 'Or'?
3950                  blnErr = True
3960                End If
3970              End Select
3980            Else
                  ' ** If this clause isn't present, we shouldn't be here!
3990              blnErr = True
4000            End If
4010          End If
              ' ** It sometimes seems to lose the closing single quote for journaltype's!
4020          intCnt = CharCnt(strFilter01, "'")
4030          If intCnt > 0 Then
4040            If intCnt Mod 2 <> 0 Then
                  'Stop
4050            End If
4060          End If
4070          strFilter02 = strFilter01
4080          If blnErr = False Then
4090            frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
4100            frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
4110            frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub.
4120            frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
4130            frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
4140            frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
                '.TransDateStart.SetFocus
                'Select Case frm.opgView
                'Case frm.opgView_optForm.OptionValue
                '  frm.frmTransaction_Audit_Sub.Form.journalno_GotFocus  ' ** Form Procedure: frmTransaction_Audit_Sub.
                'Case frm.opgView_optDatasheet.OptionValue
                '  frm.frmTransaction_Audit_Sub_ds.Form.Journal_Number_GotFocus  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
                'End Select
4150            DoEvents
4160          End If
4170        Else
              ' ** Remove the 3rd occurance.
4180          lngMultiCnt = CharCnt(strFilter01, JRNL_TYPE, True)  ' ** Module Function: modStringFuncs.
4190          intPos01 = InStr(strFilter01, JRNL_TYPE)
4200          If intPos01 > 0 Then
4210            Select Case lngMultiCnt
                Case 1&
                  ' ** Shouldn't have been here to begin with.
                  '.TransDateStart.SetFocus
4220              .cmbJournalType2 = Null  ' ** Should already be Null.
4230              .cmbJournalType3.Enabled = False
4240              .cmbJournalType3.BorderColor = WIN_CLR_DISR
4250              .cmbJournalType3.BackStyle = acBackStyleTransparent
4260            Case 2&
                  ' ** Ditto.
4270            Case 3&
4280              intPos01 = InStr(intPos01, strFilter01, ORF)  ' ** Find the 1st ' Or '.
4290              If intPos01 > 0 Then
4300                intPos01 = InStr((intPos01 + Len(ORF)), strFilter01, ORF)  ' ** Find the 2nd ' Or '.
4310                If intPos01 > 0 Then
4320                  intPos02 = InStr((intPos01 + Len(ORF) + Len(JRNL_TYPE) + 5), strFilter01, "'")  ' ** Find the closing quote.
4330                  If intPos02 > 0 Then
4340                    strFilter01 = Left(strFilter01, (intPos01 - 1)) & Mid(strFilter01, (intPos02 + 1))
4350                  Else
                        ' ** Holy Heck.
4360                    blnErr = True
4370                  End If
4380                Else
                      ' ** Snafu.
4390                  blnErr = True
4400                End If
4410              Else
                    ' ** Disorder.
4420                blnErr = True
4430              End If
4440            End Select
4450          Else
                ' ** Chaos.
4460            blnErr = True
4470          End If
              ' ** It sometimes seems to lose the closing single quote for journaltype's!
4480          intCnt = CharCnt(strFilter01, "'")
4490          If intCnt > 0 Then
4500            If intCnt Mod 2 <> 0 Then
                  'Stop
4510            End If
4520          End If
4530          strFilter02 = strFilter01
4540          If blnErr = False Then
4550            frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
4560            frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
4570            frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub.
4580            frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
4590            frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
4600            frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
                '.TransDateStart.SetFocus
                'Select Case frm.opgView
                'Case frm.opgView_optForm.OptionValue
                '  frm.frmTransaction_Audit_Sub.Form.journalno_GotFocus  ' ** Form Procedure: frmTransaction_Audit_Sub.
                'Case frm.opgView_optDatasheet.OptionValue
                '  frm.frmTransaction_Audit_Sub_ds.Form.Journal_Number_GotFocus  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
                'End Select
4610            DoEvents
4620            frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
4630          End If
4640        End If  ' ** cmbJournalType3.
4650        If blnErr = True Then
4660          Beep
4670          MsgBox "A problem occurred assembling your criteria." & vbCrLf & _
                "Please reenter your Journal Type parameters.", vbInformation + vbOKOnly, "Reenter Journal Types"
4680          .cmbJournalType1 = Null
4690          .cmbJournalType1.SetFocus
4700          .cmbJournalType2 = Null
4710          .cmbJournalType2.Enabled = False
4720          .cmbJournalType2.BorderColor = WIN_CLR_DISR
4730          .cmbJournalType2.BackStyle = acBackStyleTransparent
4740          .cmbJournalType3 = Null
4750          .cmbJournalType3.Enabled = False
4760          .cmbJournalType3.BorderColor = WIN_CLR_DISR
4770          .cmbJournalType3.BackStyle = acBackStyleTransparent
4780        End If

4790      End Select
4800      DoCmd.Hourglass False
4810      DoEvents
4820    End With

EXITP:
4830    Set frm = Nothing
4840    Exit Sub

ERRH:
4850    DoCmd.Hourglass False
4860    Select Case ERR.Number
        Case Else
4870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4880    End Select
4890    Resume EXITP

End Sub

Public Sub DetailMouse_Crit_TA(lngCals As Long, arr_varCal As Variant, frmCrit As Access.Form)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "DetailMouse_Crit_TA"

        Dim strCtl As String
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String, strTmp05 As String, strTmp06 As String
        Dim lngX As Long

4910    With frmCrit
4920      For lngX = 1& To lngCals
4930        If arr_varCal(C_ABLE, lngX) = True Then
4940          strCtl = arr_varCal(C_CNAM, lngX)
4950          strTmp03 = strCtl & "_raised_focus_img"
4960          strTmp04 = strCtl & "_raised_focus_dots_img"
4970          If .Controls(strTmp04).Visible = True Or .Controls(strTmp03).Visible = True Then
4980            strTmp01 = strCtl & "_raised_img"
4990            strTmp02 = strCtl & "_raised_semifocus_dots_img"
5000            strTmp05 = strCtl & "_sunken_focus_dots_img"
5010            strTmp06 = strCtl & "_raised_img_dis"
5020            Select Case arr_varCal(C_FOCUS, lngX)
                Case True
5030              .Controls(strTmp02).Visible = True
5040              .Controls(strTmp01).Visible = False
5050            Case False
5060              .Controls(strTmp01).Visible = True
5070              .Controls(strTmp02).Visible = False
5080            End Select
5090            .Controls(strTmp03).Visible = False
5100            .Controls(strTmp04).Visible = False
5110            .Controls(strTmp05).Visible = False
5120            .Controls(strTmp06).Visible = False
5130          End If
5140        End If
5150      Next
5160    End With

EXITP:
5170    Exit Sub

ERRH:
5180    Select Case ERR.Number
        Case Else
5190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5200    End Select
5210    Resume EXITP

End Sub

Public Sub DetailMouse_Main_TA(blnClearAll_Focus As Boolean, blnSelectAll_Focus As Boolean, blnSelectNone_Focus As Boolean, blnCkgFlds_Focus As Boolean, blnWidenToCrit_Focus As Boolean, frm As Access.Form)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "DetailMouse_Main_TA"

5310    With frm
5320      If .cmdClearAll_raised_focus_dots_img.Visible = True Or .cmdClearAll_raised_focus_img.Visible = True Then
5330        Select Case blnClearAll_Focus
            Case True
5340          .cmdClearAll_raised_semifocus_dots_img.Visible = True
5350          .cmdClearAll_raised_img.Visible = False
5360        Case False
5370          .cmdClearAll_raised_img.Visible = True
5380          .cmdClearAll_raised_semifocus_dots_img.Visible = False
5390        End Select
5400        .cmdClearAll_raised_focus_img.Visible = False
5410        .cmdClearAll_raised_focus_dots_img.Visible = False
5420        .cmdClearAll_sunken_focus_dots_img.Visible = False
5430        .cmdClearAll_raised_img_dis.Visible = False
5440      End If
5450      If .cmdSelectAll_raised_focus_dots_img.Visible = True Or .cmdSelectAll_raised_focus_img.Visible = True Then
5460        Select Case blnSelectAll_Focus
            Case True
5470          .cmdSelectAll_raised_semifocus_dots_img.Visible = True
5480          .cmdSelectAll_raised_img.Visible = False
5490        Case False
5500          .cmdSelectAll_raised_img.Visible = True
5510          .cmdSelectAll_raised_semifocus_dots_img.Visible = False
5520        End Select
5530        .cmdSelectAll_raised_focus_img.Visible = False
5540        .cmdSelectAll_raised_focus_dots_img.Visible = False
5550        .cmdSelectAll_sunken_focus_dots_img.Visible = False
5560        .cmdSelectAll_raised_img_dis.Visible = False
5570      End If
5580      If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
5590        Select Case blnSelectNone_Focus
            Case True
5600          .cmdSelectNone_raised_semifocus_dots_img.Visible = True
5610          .cmdSelectNone_raised_img.Visible = False
5620        Case False
5630          .cmdSelectNone_raised_img.Visible = True
5640          .cmdSelectNone_raised_semifocus_dots_img.Visible = False
5650        End Select
5660        .cmdSelectNone_raised_focus_img.Visible = False
5670        .cmdSelectNone_raised_focus_dots_img.Visible = False
5680        .cmdSelectNone_sunken_focus_dots_img.Visible = False
5690        .cmdSelectNone_raised_img_dis.Visible = False
5700      End If
5710      If .cmdSelect_Legend_tgl_off_raised_focus_img.Visible = True Then
5720        .cmdSelect_Legend_tgl_off_raised_img.Visible = True
5730        .cmdSelect_Legend_tgl_off_raised_focus_img.Visible = False
5740      End If
5750      If .cmdSelect_Legend_tgl_on_raised_focus_img.Visible = True Then
5760        .cmdSelect_Legend_tgl_on_raised_img.Visible = True
5770        .cmdSelect_Legend_tgl_on_raised_focus_img.Visible = False
5780      End If
5790      If .ckgFlds_cmd_raised_focus_dots_img.Visible = True Or .ckgFlds_cmd_raised_focus_img.Visible = True Then
5800        Select Case blnCkgFlds_Focus
            Case True
5810          .ckgFlds_cmd_raised_semifocus_dots_img.Visible = True
5820          .ckgFlds_cmd_raised_img.Visible = False
5830        Case False
5840          .ckgFlds_cmd_raised_img.Visible = True
5850          .ckgFlds_cmd_raised_semifocus_dots_img.Visible = False
5860        End Select
5870        .ckgFlds_cmd_raised_focus_img.Visible = False
5880        .ckgFlds_cmd_raised_focus_dots_img.Visible = False
5890        .ckgFlds_cmd_sunken_focus_dots_img.Visible = False
5900        .ckgFlds_cmd_raised_img_dis.Visible = False
5910      End If
5920      If .ckgFlds_lbl_box.Visible = True Then
5930        .ckgFlds_lbl_box.Visible = False
5940      End If
5950      If .cmdWidenToCriteria_raised_focus_dots_img.Visible = True Or .cmdWidenToCriteria_raised_focus_img.Visible = True Then
5960        Select Case blnWidenToCrit_Focus
            Case True
5970          .cmdWidenToCriteria_raised_semifocus_dots_img.Visible = True
5980          .cmdWidenToCriteria_raised_img.Visible = False
5990        Case False
6000          .cmdWidenToCriteria_raised_img.Visible = True
6010          .cmdWidenToCriteria_raised_semifocus_dots_img.Visible = False
6020        End Select
6030        .cmdWidenToCriteria_raised_focus_img.Visible = False
6040        .cmdWidenToCriteria_raised_focus_dots_img.Visible = False
6050        .cmdWidenToCriteria_sunken_focus_dots_img.Visible = False
6060        .cmdWidenToCriteria_raised_img_dis.Visible = False
6070      End If
6080    End With

EXITP:
6090    Exit Sub

ERRH:
6100    Select Case ERR.Number
        Case Else
6110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6120    End Select
6130    Resume EXITP

End Sub

Public Sub Widen_Handler_TA(strProc As String, blnWidenToCrit_Focus As Boolean, blnWidenToCrit_MouseDown As Boolean, frm As Access.Form)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "Widen_Handler_TA"

        Dim strEvent As String, strCtlName As String
        Dim intPos01 As Integer, lngCnt As Long

6210    With frm

6220      lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
6230      intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
6240      strEvent = Mid(strProc, (intPos01 + 1))
6250      strCtlName = Left(strProc, (intPos01 - 1))

6260      Select Case strEvent
          Case "GotFocus"
6270        blnWidenToCrit_Focus = True
6280        .cmdWidenToCriteria_raised_semifocus_dots_img.Visible = True
6290        .cmdWidenToCriteria_raised_img.Visible = False
6300        .cmdWidenToCriteria_raised_focus_img.Visible = False
6310        .cmdWidenToCriteria_raised_focus_dots_img.Visible = False
6320        .cmdWidenToCriteria_sunken_focus_dots_img.Visible = False
6330        .cmdWidenToCriteria_raised_img_dis.Visible = False
6340      Case "MouseDown"
6350        blnWidenToCrit_MouseDown = True
6360        .cmdWidenToCriteria_sunken_focus_dots_img.Visible = True
6370        .cmdWidenToCriteria_raised_img.Visible = False
6380        .cmdWidenToCriteria_raised_semifocus_dots_img.Visible = False
6390        .cmdWidenToCriteria_raised_focus_img.Visible = False
6400        .cmdWidenToCriteria_raised_focus_dots_img.Visible = False
6410        .cmdWidenToCriteria_raised_img_dis.Visible = False
6420      Case "MouseMove"
6430        If blnWidenToCrit_MouseDown = False Then
6440          Select Case blnWidenToCrit_Focus
              Case True
6450            .cmdWidenToCriteria_raised_focus_dots_img.Visible = True
6460            .cmdWidenToCriteria_raised_focus_img.Visible = False
6470          Case False
6480            .cmdWidenToCriteria_raised_focus_img.Visible = True
6490            .cmdWidenToCriteria_raised_focus_dots_img.Visible = False
6500          End Select
6510          .cmdWidenToCriteria_raised_img.Visible = False
6520          .cmdWidenToCriteria_raised_semifocus_dots_img.Visible = False
6530          .cmdWidenToCriteria_sunken_focus_dots_img.Visible = False
6540          .cmdWidenToCriteria_raised_img_dis.Visible = False
6550        End If
6560      Case "MouseUp"
6570        .cmdWidenToCriteria_raised_focus_dots_img.Visible = True
6580        .cmdWidenToCriteria_raised_img.Visible = False
6590        .cmdWidenToCriteria_raised_semifocus_dots_img.Visible = False
6600        .cmdWidenToCriteria_raised_focus_img.Visible = False
6610        .cmdWidenToCriteria_sunken_focus_dots_img.Visible = False
6620        .cmdWidenToCriteria_raised_img_dis.Visible = False
6630        blnWidenToCrit_MouseDown = False
6640      Case "LostFocus"
6650        .cmdWidenToCriteria_raised_img.Visible = True
6660        .cmdWidenToCriteria_raised_semifocus_dots_img.Visible = False
6670        .cmdWidenToCriteria_raised_focus_img.Visible = False
6680        .cmdWidenToCriteria_raised_focus_dots_img.Visible = False
6690        .cmdWidenToCriteria_sunken_focus_dots_img.Visible = False
6700        .cmdWidenToCriteria_raised_img_dis.Visible = False
6710        blnWidenToCrit_Focus = False
6720      End Select

6730    End With

EXITP:
6740    Exit Sub

ERRH:
6750    Select Case ERR.Number
        Case Else
6760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6770    End Select
6780    Resume EXITP

End Sub

Public Sub Select_Handler_TA(strProc As String, blnSelectAll_Focus As Boolean, blnSelectAll_MouseDown As Boolean, blnSelectNone_Focus As Boolean, blnSelectNone_MouseDown As Boolean, frm As Access.Form)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "Select_Handler_TA"

        Dim strEvent As String, strCtlName As String
        Dim intPos01 As Integer, lngCnt As Long

6810    With frm

6820      lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
6830      intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
6840      strEvent = Mid(strProc, (intPos01 + 1))
6850      strCtlName = Left(strProc, (intPos01 - 1))

6860      Select Case strEvent
          Case "GotFocus"
6870        Select Case strCtlName
            Case "cmdSelectAll"
6880          blnSelectAll_Focus = True
6890          .cmdSelectAll_raised_semifocus_dots_img.Visible = True
6900          .cmdSelectAll_raised_img.Visible = False
6910          .cmdSelectAll_raised_focus_img.Visible = False
6920          .cmdSelectAll_raised_focus_dots_img.Visible = False
6930          .cmdSelectAll_sunken_focus_dots_img.Visible = False
6940          .cmdSelectAll_raised_img_dis.Visible = False
6950        Case "cmdSelectNone"
6960          blnSelectNone_Focus = True
6970          .cmdSelectNone_raised_semifocus_dots_img.Visible = True
6980          .cmdSelectNone_raised_img.Visible = False
6990          .cmdSelectNone_raised_focus_img.Visible = False
7000          .cmdSelectNone_raised_focus_dots_img.Visible = False
7010          .cmdSelectNone_sunken_focus_dots_img.Visible = False
7020          .cmdSelectNone_raised_img_dis.Visible = False
7030        End Select
7040      Case "MouseDown"
7050        Select Case strCtlName
            Case "cmdSelectAll"
7060          blnSelectAll_MouseDown = True
7070          .cmdSelectAll_sunken_focus_dots_img.Visible = True
7080          .cmdSelectAll_raised_img.Visible = False
7090          .cmdSelectAll_raised_semifocus_dots_img.Visible = False
7100          .cmdSelectAll_raised_focus_img.Visible = False
7110          .cmdSelectAll_raised_focus_dots_img.Visible = False
7120          .cmdSelectAll_raised_img_dis.Visible = False
7130        Case "cmdSelectNone"
7140          blnSelectNone_MouseDown = True
7150          .cmdSelectNone_sunken_focus_dots_img.Visible = True
7160          .cmdSelectNone_raised_img.Visible = False
7170          .cmdSelectNone_raised_semifocus_dots_img.Visible = False
7180          .cmdSelectNone_raised_focus_img.Visible = False
7190          .cmdSelectNone_raised_focus_dots_img.Visible = False
7200          .cmdSelectNone_raised_img_dis.Visible = False
7210        End Select
7220      Case "MouseMove"
7230        Select Case strCtlName
            Case "cmdSelectAll"
7240          If blnSelectAll_MouseDown = False Then
7250            Select Case blnSelectAll_Focus
                Case True
7260              .cmdSelectAll_raised_focus_dots_img.Visible = True
7270              .cmdSelectAll_raised_focus_img.Visible = False
7280            Case False
7290              .cmdSelectAll_raised_focus_img.Visible = True
7300              .cmdSelectAll_raised_focus_dots_img.Visible = False
7310            End Select
7320            .cmdSelectAll_raised_img.Visible = False
7330            .cmdSelectAll_raised_semifocus_dots_img.Visible = False
7340            .cmdSelectAll_sunken_focus_dots_img.Visible = False
7350            .cmdSelectAll_raised_img_dis.Visible = False
7360          End If
7370          If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
7380            Select Case blnSelectNone_Focus
                Case True
7390              .cmdSelectNone_raised_semifocus_dots_img.Visible = True
7400              .cmdSelectNone_raised_img.Visible = False
7410            Case False
7420              .cmdSelectNone_raised_img.Visible = True
7430              .cmdSelectNone_raised_semifocus_dots_img.Visible = False
7440            End Select
7450            .cmdSelectNone_raised_focus_img.Visible = False
7460            .cmdSelectNone_raised_focus_dots_img.Visible = False
7470            .cmdSelectNone_sunken_focus_dots_img.Visible = False
7480            .cmdSelectNone_raised_img_dis.Visible = False
7490          End If
7500        Case "cmdSelectNone"
7510          If blnSelectNone_MouseDown = False Then
7520            Select Case blnSelectNone_Focus
                Case True
7530              .cmdSelectNone_raised_focus_dots_img.Visible = True
7540              .cmdSelectNone_raised_focus_img.Visible = False
7550            Case False
7560              .cmdSelectNone_raised_focus_img.Visible = True
7570              .cmdSelectNone_raised_focus_dots_img.Visible = False
7580            End Select
7590            .cmdSelectNone_raised_img.Visible = False
7600            .cmdSelectNone_raised_semifocus_dots_img.Visible = False
7610            .cmdSelectNone_sunken_focus_dots_img.Visible = False
7620            .cmdSelectNone_raised_img_dis.Visible = False
7630          End If
7640          If .cmdSelectAll_raised_focus_dots_img.Visible = True Or .cmdSelectAll_raised_focus_img.Visible = True Then
7650            Select Case blnSelectAll_Focus
                Case True
7660              .cmdSelectAll_raised_semifocus_dots_img.Visible = True
7670              .cmdSelectAll_raised_img.Visible = False
7680            Case False
7690              .cmdSelectAll_raised_img.Visible = True
7700              .cmdSelectAll_raised_semifocus_dots_img.Visible = False
7710            End Select
7720            .cmdSelectAll_raised_focus_img.Visible = False
7730            .cmdSelectAll_raised_focus_dots_img.Visible = False
7740            .cmdSelectAll_sunken_focus_dots_img.Visible = False
7750            .cmdSelectAll_raised_img_dis.Visible = False
7760          End If
7770        End Select
7780      Case "MouseUp"
7790        Select Case strCtlName
            Case "cmdSelectAll"
7800          .cmdSelectAll_raised_focus_dots_img.Visible = True
7810          .cmdSelectAll_raised_img.Visible = False
7820          .cmdSelectAll_raised_semifocus_dots_img.Visible = False
7830          .cmdSelectAll_raised_focus_img.Visible = False
7840          .cmdSelectAll_sunken_focus_dots_img.Visible = False
7850          .cmdSelectAll_raised_img_dis.Visible = False
7860          blnSelectAll_MouseDown = False
7870        Case "cmdSelectNone"
7880          .cmdSelectNone_raised_focus_dots_img.Visible = True
7890          .cmdSelectNone_raised_img.Visible = False
7900          .cmdSelectNone_raised_semifocus_dots_img.Visible = False
7910          .cmdSelectNone_raised_focus_img.Visible = False
7920          .cmdSelectNone_sunken_focus_dots_img.Visible = False
7930          .cmdSelectNone_raised_img_dis.Visible = False
7940          blnSelectNone_MouseDown = False
7950        End Select
7960      Case "LostFocus"
7970        Select Case strCtlName
            Case "cmdSelectAll"
7980          .cmdSelectAll_raised_img.Visible = True
7990          .cmdSelectAll_raised_semifocus_dots_img.Visible = False
8000          .cmdSelectAll_raised_focus_img.Visible = False
8010          .cmdSelectAll_raised_focus_dots_img.Visible = False
8020          .cmdSelectAll_sunken_focus_dots_img.Visible = False
8030          .cmdSelectAll_raised_img_dis.Visible = False
8040          blnSelectAll_Focus = False
8050        Case "cmdSelectNone"
8060          .cmdSelectNone_raised_img.Visible = True
8070          .cmdSelectNone_raised_semifocus_dots_img.Visible = False
8080          .cmdSelectNone_raised_focus_img.Visible = False
8090          .cmdSelectNone_raised_focus_dots_img.Visible = False
8100          .cmdSelectNone_sunken_focus_dots_img.Visible = False
8110          .cmdSelectNone_raised_img_dis.Visible = False
8120          blnSelectNone_Focus = False
8130        End Select
8140      End Select

8150    End With

EXITP:
8160    Exit Sub

ERRH:
8170    Select Case ERR.Number
        Case Else
8180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8190    End Select
8200    Resume EXITP

End Sub

Public Sub Clear_Handler_TA(strProc As String, blnClearAll_Focus As Boolean, blnClearAll_MouseDown As Boolean, frm As Access.Form)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "Clear_Handler_TA"

        Dim strEvent As String, strCtlName As String
        Dim intPos01 As Integer, lngCnt As Long

8310    With frm

8320      lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
8330      intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
8340      strEvent = Mid(strProc, (intPos01 + 1))
8350      strCtlName = Left(strProc, (intPos01 - 1))

8360      Select Case strEvent
          Case "GotFocus"
8370        blnClearAll_Focus = True
8380        .cmdClearAll_raised_semifocus_dots_img.Visible = True
8390        .cmdClearAll_raised_img.Visible = False
8400        .cmdClearAll_raised_focus_img.Visible = False
8410        .cmdClearAll_raised_focus_dots_img.Visible = False
8420        .cmdClearAll_sunken_focus_dots_img.Visible = False
8430        .cmdClearAll_raised_img_dis.Visible = False
8440      Case "MouseDown"
8450        blnClearAll_MouseDown = True
8460        .cmdClearAll_sunken_focus_dots_img.Visible = True
8470        .cmdClearAll_raised_img.Visible = False
8480        .cmdClearAll_raised_semifocus_dots_img.Visible = False
8490        .cmdClearAll_raised_focus_img.Visible = False
8500        .cmdClearAll_raised_focus_dots_img.Visible = False
8510        .cmdClearAll_raised_img_dis.Visible = False
8520      Case "MouseMove"
8530        If blnClearAll_MouseDown = False Then
8540          Select Case blnClearAll_Focus
              Case True
8550            .cmdClearAll_raised_focus_dots_img.Visible = True
8560            .cmdClearAll_raised_focus_img.Visible = False
8570          Case False
8580            .cmdClearAll_raised_focus_img.Visible = True
8590            .cmdClearAll_raised_focus_dots_img.Visible = False
8600          End Select
8610          .cmdClearAll_raised_img.Visible = False
8620          .cmdClearAll_raised_semifocus_dots_img.Visible = False
8630          .cmdClearAll_sunken_focus_dots_img.Visible = False
8640          .cmdClearAll_raised_img_dis.Visible = False
8650        End If
8660      Case "MouseUp"
8670        .cmdClearAll_raised_focus_dots_img.Visible = True
8680        .cmdClearAll_raised_img.Visible = False
8690        .cmdClearAll_raised_semifocus_dots_img.Visible = False
8700        .cmdClearAll_raised_focus_img.Visible = False
8710        .cmdClearAll_sunken_focus_dots_img.Visible = False
8720        .cmdClearAll_raised_img_dis.Visible = False
8730        blnClearAll_MouseDown = False
8740      Case "LostFocus"
8750        .cmdClearAll_raised_img.Visible = True
8760        .cmdClearAll_raised_semifocus_dots_img.Visible = False
8770        .cmdClearAll_raised_focus_img.Visible = False
8780        .cmdClearAll_raised_focus_dots_img.Visible = False
8790        .cmdClearAll_sunken_focus_dots_img.Visible = False
8800        .cmdClearAll_raised_img_dis.Visible = False
8810        blnClearAll_Focus = False
8820      End Select

8830    End With

EXITP:
8840    Exit Sub

ERRH:
8850    Select Case ERR.Number
        Case Else
8860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8870    End Select
8880    Resume EXITP

End Sub

Public Sub CkgFlds_Handler_TA(strProc As String, blnCkgFlds_Focus As Boolean, blnCkgFlds_MouseDown As Boolean, blnFromCkgFlds As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long, lngCkgFldsBox_Offset As Long, lngChkBoxLbl_TopOffset As Long, lngCkgFldsBox_Height As Long, lngSelectBox_Top As Long, lngCkgFldsVLine_Top As Long, lngCkgFldsVLine_Height As Long, lngFrmFlds As Long, arr_varFrmFld As Variant, frm As Access.Form)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "CkgFlds_Handler_TA"

        Dim strEvent As String, strCtlName As String
        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngDiff_Height As Long
        Dim intPos01 As Integer, lngCnt As Long
        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim lngX As Long

8910    With frm

8920      lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
8930      intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
8940      strEvent = Mid(strProc, (intPos01 + 1))
8950      strCtlName = Left(strProc, (intPos01 - 1))

8960      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
8970        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
8980      End If

8990      Select Case strEvent
          Case "GotFocus"
9000        blnCkgFlds_Focus = True
9010        Select Case .ckgFlds
            Case True
              ' ** I think if it were true, it wouldn't be here.
9020        Case False
9030          .ckgFlds_cmd_raised_semifocus_dots_img.Visible = True
9040          .ckgFlds_cmd_raised_img.Visible = False
9050          .ckgFlds_cmd_raised_focus_img.Visible = False
9060          .ckgFlds_cmd_raised_focus_dots_img.Visible = False
9070          .ckgFlds_cmd_sunken_focus_dots_img.Visible = False
9080          .ckgFlds_cmd_raised_img_dis.Visible = False
9090        End Select
9100      Case "MouseDown"
9110        blnCkgFlds_MouseDown = True
9120        Select Case .ckgFlds
            Case True
              ' ** Ditto.
9130        Case False
9140          .ckgFlds_cmd_sunken_focus_dots_img.Visible = True
9150          .ckgFlds_cmd_raised_img.Visible = False
9160          .ckgFlds_cmd_raised_semifocus_dots_img.Visible = False
9170          .ckgFlds_cmd_raised_focus_img.Visible = False
9180          .ckgFlds_cmd_raised_focus_dots_img.Visible = False
9190          .ckgFlds_cmd_raised_img_dis.Visible = False
9200        End Select
9210      Case "MouseMove"
9220        If blnCkgFlds_MouseDown = False Then
9230          Select Case .ckgFlds
              Case True
9240            .ckgFlds_lbl_box.Visible = True
9250          Case False
9260            Select Case blnCkgFlds_Focus
                Case True
9270              .ckgFlds_cmd_raised_focus_dots_img.Visible = True
9280              .ckgFlds_cmd_raised_focus_img.Visible = False
9290            Case False
9300              .ckgFlds_cmd_raised_focus_img.Visible = True
9310              .ckgFlds_cmd_raised_focus_dots_img.Visible = False
9320            End Select
9330            .ckgFlds_cmd_raised_img.Visible = False
9340            .ckgFlds_cmd_raised_semifocus_dots_img.Visible = False
9350            .ckgFlds_cmd_sunken_focus_dots_img.Visible = False
9360            .ckgFlds_cmd_raised_img_dis.Visible = False
9370          End Select
9380        End If
9390      Case "MouseUp"
9400        Select Case .ckgFlds
            Case True
              ' ** Ditto.
9410        Case False
9420          .ckgFlds_cmd_raised_focus_dots_img.Visible = True
9430          .ckgFlds_cmd_raised_img.Visible = False
9440          .ckgFlds_cmd_raised_semifocus_dots_img.Visible = False
9450          .ckgFlds_cmd_raised_focus_img.Visible = False
9460          .ckgFlds_cmd_sunken_focus_dots_img.Visible = False
9470          .ckgFlds_cmd_raised_img_dis.Visible = False
9480        End Select
9490        blnCkgFlds_MouseDown = False
9500      Case "LostFocus"
9510        Select Case .ckgFlds
            Case True
              ' ** Ditto.
9520        Case False
9530          .ckgFlds_cmd_raised_img.Visible = True
9540          .ckgFlds_cmd_raised_semifocus_dots_img.Visible = False
9550          .ckgFlds_cmd_raised_focus_img.Visible = False
9560          .ckgFlds_cmd_raised_focus_dots_img.Visible = False
9570          .ckgFlds_cmd_sunken_focus_dots_img.Visible = False
9580          .ckgFlds_cmd_raised_img_dis.Visible = False
9590        End Select
9600        blnCkgFlds_Focus = False
9610      Case "Click"
9620        .FocusHolder.SetFocus
            ' ** Variables are fed empty, then populated ByRef.
9630        GetFormDimensions frm, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Function: modWindowFunctions.
9640        lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
9650        lngMonitorNum = 1&: lngTmp02 = 0&
9660        EnumMonitors frm  ' ** Module Function: modMonitorFuncs.
9670        If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
9680        Select Case .ckgFlds
            Case True
              ' ** Close it.
9690          .ckgFlds = False
9700          lngDiff_Height = (.ckgFlds_hline04.Top + lngCkgFldsBox_Offset)  ' ** From bottom of closed box to bottom of Detail.
9710          lngDiff_Height = (.Detail.Height - lngDiff_Height)
9720          For lngX = 0& To (lngFrmFlds - 1&)
9730            .Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Visible = False
                ' ** Get the difference between its original open position and its current open position.
9740            If .Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top = arr_varFrmFld(FM_TOPO, lngX) Then
                  ' ** Form hasn't been resized.
9750              .Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top = arr_varFrmFld(FM_TOPC, lngX)
9760              .Controls(arr_varFrmFld(FM_VIEWCHK, lngX) & "_lbl").Top = _
                    (.Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top - lngChkBoxLbl_TopOffset)
9770            Else
                  ' ** Form has been resized.
9780              If .Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top < arr_varFrmFld(FM_TOPO, lngX) Then
9790                lngTmp01 = (arr_varFrmFld(FM_TOPO, lngX) - .Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top)
9800                .Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top = (arr_varFrmFld(FM_TOPC, lngX) - lngTmp01)
9810              Else
9820                lngTmp01 = (.Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top - arr_varFrmFld(FM_TOPO, lngX))
9830                .Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top = (arr_varFrmFld(FM_TOPC, lngX) + lngTmp01)
9840              End If
9850              .Controls(arr_varFrmFld(FM_VIEWCHK, lngX) & "_lbl").Top = _
                    (.Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top - lngChkBoxLbl_TopOffset)
9860            End If
9870          Next
9880          .FldCnt.Visible = False
9890          .FldCnt.Top = .ckgFlds_cmd.Top
9900          .FldCnt_lbl.Top = .ckgFlds_cmd.Top
9910          For lngX = 9& To 20&
9920            .Controls("ckgFlds_vline" & Right("00" & CStr(lngX), 2)).Visible = False
9930            .Controls("ckgFlds_vline" & Right("00" & CStr(lngX), 2)).Top = .ckgFlds_cmd.Top
9940            .Controls("ckgFlds_vline" & Right("00" & CStr(lngX), 2)).Height = .ckgFlds_cmd.Height
9950          Next
9960          .ckgFlds_lbl.Visible = False
9970          .ckgFlds_box.Visible = False
9980          .ckgFlds_box.Height = .ckgFlds_cmd.Height
9990          .ckgFlds_hline03.Visible = True
10000         .ckgFlds_hline04.Visible = True
10010         .ckgFlds_vline05.Visible = True
10020         .ckgFlds_vline06.Visible = True
10030         .ckgFlds_vline07.Visible = True
10040         .ckgFlds_vline08.Visible = True
10050         .ckgFlds_cmd_raised_focus_dots_img.Visible = True
10060         .ckgFlds_cmd_raised_img.Visible = False
10070         .ckgFlds_cmd_raised_semifocus_dots_img.Visible = False
10080         .ckgFlds_cmd_raised_focus_img.Visible = False
10090         .ckgFlds_cmd_sunken_focus_dots_img.Visible = False
10100         .ckgFlds_cmd_raised_img_dis.Visible = False
10110         .Detail.Height = (.Detail.Height - lngDiff_Height)
10120         blnFromCkgFlds = True
10130         If lngMonitorNum = 1& Then lngTmp02 = lngTop
10140         DoCmd.SelectObject acForm, "frmTransaction_Audit", False
10150         DoEvents
10160         DoCmd.MoveSize lngLeft, lngTmp02, lngWidth, (lngHeight - lngDiff_Height)  'lngTop
10170         If lngMonitorNum > 1& Then
10180           LoadPosition .hwnd, frm.Name   ' ** Module Function: modMonitorFuncs.
10190         End If
10200         DoEvents
10210         blnFromCkgFlds = False
10220       Case False
              ' ** Open it.
10230         .ckgFlds = True
10240         .ckgFlds_cmd_raised_img.Visible = False
10250         .ckgFlds_cmd_raised_semifocus_dots_img.Visible = False
10260         .ckgFlds_cmd_raised_focus_img.Visible = False
10270         .ckgFlds_cmd_raised_focus_dots_img.Visible = False
10280         .ckgFlds_cmd_sunken_focus_dots_img.Visible = False
10290         .ckgFlds_cmd_raised_img_dis.Visible = False
10300         blnFromCkgFlds = True
10310         lngDiff_Height = ((.ckgFlds_box.Top + lngCkgFldsBox_Height) + lngCkgFldsBox_Offset)  ' ** Should be new Detail.Height.
              ' ** We're mixing Detail height difference with Form height difference.
10320         lngTmp01 = (.ckgFlds_hline04.Top + lngCkgFldsBox_Offset)
10330         lngDiff_Height = (lngDiff_Height - lngTmp01)
10340         If lngMonitorNum = 1& Then lngTmp02 = lngTop
10350         DoCmd.SelectObject acForm, "frmTransaction_Audit", False
10360         DoEvents
10370         DoCmd.MoveSize lngLeft, lngTmp02, lngWidth, (lngHeight + lngDiff_Height)  'lngTop
10380         If lngMonitorNum > 1& Then
10390           LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
10400         End If
10410         DoEvents
10420         blnFromCkgFlds = False
10430         .Detail.Height = ((.ckgFlds_box.Top + lngCkgFldsBox_Height) + lngCkgFldsBox_Offset)
10440         .ckgFlds_box.Height = lngCkgFldsBox_Height
10450         .ckgFlds_box.Visible = True
10460         .ckgFlds_lbl.Visible = True
10470         For lngX = 0& To (lngFrmFlds - 1&)
                ' ** Get the difference between its original closed position and its current closed position.
10480           If .Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top = arr_varFrmFld(FM_TOPC, lngX) Then
                  ' ** Form hasn't been resized.
10490             .Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top = arr_varFrmFld(FM_TOPO, lngX)
10500             .Controls(arr_varFrmFld(FM_VIEWCHK, lngX) & "_lbl").Top = _
                    (.Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top - lngChkBoxLbl_TopOffset)
10510           Else
                  ' ** Form has been resized.
10520             lngTmp01 = (arr_varFrmFld(FM_TOPO, lngX) - arr_varFrmFld(FM_TOPC, lngX))  ' ** Difference between open and closed position.
10530             .Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top = .ckgFlds_cmd.Top + lngTmp01
10540             .Controls(arr_varFrmFld(FM_VIEWCHK, lngX) & "_lbl").Top = _
                    (.Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top - lngChkBoxLbl_TopOffset)
10550           End If
10560           .Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Visible = True
10570         Next
10580         lngTmp01 = (((lngSelectBox_Top + .cmdSelect_box.Height) + (8& * lngTpp)) + lngTpp)  ' ** Original Detail_hline04.Top.
10590         lngTmp01 = ((((lngTmp01 + (8& * lngTpp)) + lngTpp) + lngTpp) + lngTpp)  ' ** Original ckgFlds_cmd.Top.
10600         lngTmp01 = (lngCkgFldsVLine_Top - lngTmp01)  ' ** Offset between vline tops and ckgFlds_cmd.Top.
10610         For lngX = 9& To 20&
10620           If lngCkgFldsVLine_Top = (.ckgFlds_cmd.Top + lngTmp01) Then
                  ' ** Form hasn't been resized.
10630             .Controls("ckgFlds_vline" & Right("00" & CStr(lngX), 2)).Top = lngCkgFldsVLine_Top
10640           Else
                  ' ** Form has been resized.
10650             .Controls("ckgFlds_vline" & Right("00" & CStr(lngX), 2)).Top = (.ckgFlds_cmd.Top + lngTmp01)
10660           End If
10670           .Controls("ckgFlds_vline" & Right("00" & CStr(lngX), 2)).Height = lngCkgFldsVLine_Height
10680           .Controls("ckgFlds_vline" & Right("00" & CStr(lngX), 2)).Visible = True
10690         Next
10700         .FldCnt.Top = .ckgFlds_chkAccountNo_lbl.Top
10710         .FldCnt_lbl.Top = .FldCnt.Top
10720         .FldCnt.Visible = True
10730         .ckgFlds_hline03.Visible = False
10740         .ckgFlds_hline04.Visible = False
10750         .ckgFlds_vline05.Visible = False
10760         .ckgFlds_vline06.Visible = False
10770         .ckgFlds_vline07.Visible = False
10780         .ckgFlds_vline08.Visible = False
10790       End Select
10800     End Select

10810   End With

EXITP:
10820   Exit Sub

ERRH:
10830   Select Case ERR.Number
        Case Else
10840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10850   End Select
10860   Resume EXITP

End Sub

Public Sub ReloadCritPrefs_TA(frmCrit As Access.Form)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "ReloadCritPrefs_TA"

10910   With frmCrit

10920     Pref_Load frmCrit.Name  ' ** Module Procedure: modPreferenceFuncs.
10930     DoEvents

10940     If IsNull(.journalno) = False Then
10950       .journalno_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
10960       DoEvents
10970     End If
10980     If IsNull(.cmbJournalType1) = False Then
10990       .cmbJournalType1_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11000       DoEvents
11010     End If
11020     If IsNull(.cmbJournalType2) = False Then
11030       .cmbJournalType2_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11040       DoEvents
11050     End If
11060     If IsNull(.cmbJournalType3) = False Then
11070       .cmbJournalType3_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11080       DoEvents
11090     End If
11100     If IsNull(.TransDateStart) = False Then
11110       .TransDateStart_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11120       DoEvents
11130     End If
11140     If IsNull(.TransDateEnd) = False Then
11150       .TransDateEnd_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11160       DoEvents
11170     End If
11180     If IsNull(.cmbAccounts) = False Then
11190       .cmbAccounts_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11200       DoEvents
11210     End If
11220     .opgAccountSource_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11230     If IsNull(.cmbAssets) = False Then
11240       .cmbAssets_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11250     End If
11260     If IsNull(.cmbCurrencies) = False Then
11270       If .cmbCurrencies <> 150& Then
11280         .cmbCurrencies_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11290       End If
11300     End If
11310     If IsNull(.AssetDateStart) = False Then
11320       .AssetDateStart_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11330     End If
11340     If IsNull(.AssetDateEnd) = False Then
11350       .AssetDateEnd_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11360     End If
11370     If IsNull(.PurchaseDateStart) = False Then
11380       .PurchaseDateStart_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11390     End If
11400     If IsNull(.PurchaseDateEnd) = False Then
11410       .PurchaseDateEnd_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11420     End If
11430     If IsNull(.ledger_description) = False Then
11440       .ledger_description_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11450     End If
11460     If IsNull(.cmbRecurringItems) = False Then
11470       .cmbRecurringItems_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11480     End If
11490     If IsNull(.cmbRevenueCodes) = False Then
11500       .cmbRevenueCodes_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11510     End If
11520     .chkRevcodeType_Income_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11530     .chkRevcodeType_Expense_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11540     If IsNull(.cmbTaxCodes) = False Then
11550       .cmbTaxCodes_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11560     End If
11570     .chkTaxcodeType_Income_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11580     .chkTaxcodeType_Deduction_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11590     If IsNull(.cmbLocations) = False Then
11600       .cmbLocations_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11610     End If
11620     If IsNull(.CheckNum) = False Then
11630       .CheckNum_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11640     End If
11650     If IsNull(.cmbUsers) = False Then
11660       .cmbUsers_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11670     End If
11680     If IsNull(.PostedDateStart) = False Then
11690       .PostedDateStart_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11700     End If
11710     If IsNull(.PostedDateEnd) = False Then
11720       .PostedDateEnd_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
11730     End If
11740     .opgHidden_AfterUpdate

11750   End With

EXITP:
11760   Exit Sub

ERRH:
11770   Select Case ERR.Number
        Case Else
11780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11790   End Select
11800   Resume EXITP

End Sub

Public Sub AcctSource_After_TA(frmCrit As Access.Form)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "AcctSource_After_TA"

        Dim strAccountNo As String

11910   With frmCrit
11920     DoCmd.Hourglass True
11930     DoEvents
11940     strAccountNo = vbNullString
11950     If IsNull(.cmbAccounts) = False Then
11960       If Len(.cmbAccounts.Column(0)) > 0 Then
11970         strAccountNo = .cmbAccounts.Column(0)
11980       End If
11990     End If
12000     Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
12010       If .cmbAccounts.RowSource <> "qryAccountNoDropDown_03" Then
12020         .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
12030       End If
12040       .opgAccountSource_optNumber_lbl.FontBold = True
12050       .opgAccountSource_optName_lbl.FontBold = False
12060     Case .opgAccountSource_optName.OptionValue
12070       If .cmbAccounts.RowSource <> "qryAccountNoDropDown_04" Then
12080         .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
12090       End If
12100       .opgAccountSource_optNumber_lbl.FontBold = False
12110       .opgAccountSource_optName_lbl.FontBold = True
12120     End Select
12130     DoEvents
12140     If strAccountNo <> vbNullString Then
12150       .cmbAccounts = strAccountNo
12160     End If
12170     DoCmd.Hourglass False
12180     DoEvents
12190   End With

EXITP:
12200   Exit Sub

ERRH:
12210   DoCmd.Hourglass False
12220   Select Case ERR.Number
        Case Else
12230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12240   End Select
12250   Resume EXITP

End Sub

Public Sub AssetSource_After_TA(frmCrit As Access.Form)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetSource_After_TA"

        Dim lngAssetNo As Long

12310   With frmCrit
12320     DoCmd.Hourglass True
12330     DoEvents
12340     lngAssetNo = 0&
12350     If IsNull(.cmbAssets) = False Then
12360       If .cmbAssets > 0& Then
12370         lngAssetNo = .cmbAssets
12380       End If
12390     End If
12400     Select Case .opgAssetSource
          Case .opgAssetSource_optType.OptionValue
12410       If .cmbAssets.RowSource <> "qryTransaction_Audit_10_02_Asset_Type" Then
12420         .cmbAssets.RowSource = "qryTransaction_Audit_10_02_Asset_Type"
12430       End If
12440       .opgAssetSource_optType_lbl.FontBold = True
12450       .opgAssetSource_optName_lbl.FontBold = False
12460       .opgAssetSource_optCusip_lbl.FontBold = False
12470     Case .opgAssetSource_optName.OptionValue
12480       If .cmbAssets.RowSource <> "qryTransaction_Audit_10_02_Asset_Name" Then
12490         .cmbAssets.RowSource = "qryTransaction_Audit_10_02_Asset_Name"
12500       End If
12510       .opgAssetSource_optType_lbl.FontBold = False
12520       .opgAssetSource_optName_lbl.FontBold = True
12530       .opgAssetSource_optCusip_lbl.FontBold = False
12540     Case .opgAssetSource_optCusip.OptionValue
12550       If .cmbAssets.RowSource <> "qryTransaction_Audit_10_02_Asset_Cusip" Then
12560         .cmbAssets.RowSource = "qryTransaction_Audit_10_02_Asset_Cusip"
12570       End If
12580       .opgAssetSource_optType_lbl.FontBold = False
12590       .opgAssetSource_optName_lbl.FontBold = False
12600       .opgAssetSource_optCusip_lbl.FontBold = True
12610     End Select
12620     DoEvents
12630     If lngAssetNo > 0& Then
12640       .cmbAssets = lngAssetNo
12650     End If
12660     DoCmd.Hourglass False
12670     DoEvents
12680   End With

EXITP:
12690   Exit Sub

ERRH:
12700   DoCmd.Hourglass False
12710   Select Case ERR.Number
        Case Else
12720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12730   End Select
12740   Resume EXITP

End Sub
