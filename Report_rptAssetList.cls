VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAssetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAssetList"

'VGC 10/23/2017: CHANGES!

'totdesc_disp:
'=IIf(IsNull([totdesc])=True,[accountno],IIf([totdesc]='',[accountno],[totdesc]))
'totdesc:
'IIf(IsNull([ActiveAssets].[assetno])=True,'',CStr([masterasset].[description]) & IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy')))

Private Const strCallingForm As String = "frmStatementParameters"

' ** Array: arr_varFld().
Private lngFlds As Long, arr_varFld() As Variant
Private Const F_ELEMS As Integer = 9  ' ** Array's first-element UBound().
Private Const F_ACTNO As Integer = 0
Private Const F_DATE  As Integer = 1
Private Const F_FOREX As Integer = 2
Private Const F_CST   As Integer = 3
Private Const F_MKT   As Integer = 4
Private Const F_ICSH  As Integer = 5
Private Const F_PCSH  As Integer = 6
Private Const F_TCST  As Integer = 7
Private Const F_TMKT  As Integer = 8
Private Const F_RECS  As Integer = 9

Private strAccountNo As String
Private datStartDate As Date, datEndDate As Date
Private blnStatements As Boolean, blnSaveData As Boolean, blnNoAssets As Boolean
Private frm As Access.Form
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strFilter As String
        Dim blnSkip As Boolean
        Dim varTmp00 As Variant
        Dim lngZ As Long

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
            ' ** Data taken directly from form.
            ' ** .AssetListDate
            ' ** .DateEnd
            ' ** .chkStatements
            ' ** .shortname
            ' ** .legalname

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
140         blnSkip = False: blnSaveData = False

150         .group_footer1_height_line.Top = 0&  ' ** Just get it out of the way.

160         Set frm = Forms(strCallingForm)

170         If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
180           If frm.PopUp = True Then
190             frm.Visible = False
200           End If
210         End If

220         Select Case frm.chkPageOf
            Case True
230           .PageNumOf.Visible = True
240           .PageNum.Visible = False
250         Case False
260           .PageNum.Visible = True
270           .PageNumOf.Visible = False
280           .PageNumOf.ControlSource = vbNullString
290         End Select

300         .Name_lbl.Caption = vbNullString

310         varTmp00 = frm.cmbAccounts
320         If IsNull(varTmp00) = False Then
330           varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
340           varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
350           .Caption = "Asset_List_" & varTmp00 & "_" & Format(CDate(frm.AssetListDate), "yymmdd")
360           strAccountNo = varTmp00
370         Else
380           .Caption = "Asset_List_All_" & Format(CDate(frm.AssetListDate), "yymmdd")
390           strAccountNo = "All"
400         End If

410         VerifyDates  ' ** Procedure: Below.

420         datStartDate = frm.DateStart
430         datEndDate = frm.AssetListDate
440         blnNoAssets = frm.chkNoAssets

450         Select Case frm.opgAccountNumber
            Case frm.opgAccountNumber_optSpecified.OptionValue
460           .Section("GroupFooter1").ForceNewPage = acForceNewPageNone
470         Case frm.opgAccountNumber_optAll.OptionValue
480           .Section("GroupFooter1").ForceNewPage = acForceNewPageAfter
490         End Select

500         Select Case frm.chkStatements
            Case True
510           blnStatements = True
520         Case False
530           blnStatements = False
540         End Select

550         Select Case gblnMessage
            Case True
560           Select Case blnNoAssets
              Case True
                ' ** qryStatementParameters_AssetList_74_41 (xx), linked to Account, for 'NoAssets'.
570             .RecordSource = "qryStatementParameters_AssetList_15e"
580           Case False
                ' ** Rollbacks were needed.
590             Select Case frm.chkStatements
                Case True
                  ' ** qryStatementParameters_AssetList_15 (qryStatementParameters_AssetList_15a
                  ' ** (tmpAssetList2, all fields, with rollback), with Liability adjustment), linked
                  ' ** to qryStatementParameters_AssetList_19b (qryStatementParameters_AssetList_19a
                  ' ** (Account, with MonthNum = FormRef('MonthNum')), just those matching MonthNum), just scheduled accounts.
600               .RecordSource = "qryStatementParameters_AssetList_15b"
610             Case False
                  ' ** qryStatementParameters_AssetList_15a (tmpAssetList2, all fields, with rollback), with Liability adjustment.
620               .RecordSource = "qryStatementParameters_AssetList_15"
630             End Select
640           End Select
              'VGC 11/29/2014: In an effort to track down a problem I know I already fixed, but has shown up again,
              'and the source of which I now can't find, I'm adding this here to see if it limits its effects.
650           If frm.opgAccountNumber = frm.opgAccountNumber_optSpecified.OptionValue Then
660             strFilter = "[accountno] = '" & frm.cmbAccounts & "'"
670             .Filter = strFilter
680             .FilterOn = True
690           End If
700         Case False
              ' ** No Rollbacks needed.
710           Select Case frm.chkRelatedAccounts
              Case True
                ' ** With Related Accounts.
                'NO ASSETS!
720             Select Case frm.chkStatements
                Case True
                  ' ** qryStatementParameters_AssetList_28 (qryStatementParameters_AssetList_27 (tmpRelatedAccount_02,
                  ' ** with qryStatementParameters_AssetList_26b (qryStatementParameters_AssetList_26a (tmpRelatedAccount_02,
                  ' ** linked to Account, grouped and summed, by accountno), grouped and summed), by specified [ractnos];
                  ' ** Cartesian), from code, with [ractnos] replaced with actual accountno's; Cartesian), linked to
                  ' ** qryStatementParameters_AssetList_19b (qryStatementParameters_AssetList_19a (Account, with
                  ' ** MonthNum = FormRef('MonthNum')), just those matching MonthNum), just scheduled accounts.
730               .RecordSource = "qryStatementParameters_AssetList_28a"
740             Case False
                  ' ** qryStatementParameters_AssetList_27 (tmpRelatedAccount_02, with qryStatementParameters_AssetList_26b
                  ' ** (qryStatementParameters_AssetList_26a (tmpRelatedAccount_02, linked to Account, grouped and summed,
                  ' ** by accountno), grouped and summed), by specified [ractnos]; Cartesian), from code, with [ractnos]
                  ' ** replaced with actual accountno's; Cartesian.
750               .RecordSource = "qryStatementParameters_AssetList_28"  'I THINK! DON'T KNOW IF IT NEEDS A 'NoAssets'!
760             End Select
770           Case False
                ' ** Without Related Accounts.
780             Select Case frm.opgAccountNumber
                Case frm.opgAccountNumber_optSpecified.OptionValue
                  ' ** One Account.
790               Select Case frm.chkNoAssets
                  Case True
                    'Select Case frm.chkStatements
                    'Case True
                    '  ' ** qryStatementParameters_AssetList_06c (Account, linked to ActiveAssets, grouped,
                    '  ' ** with add'l fields; specified FormRef('accountno'), without assetno <> Null), linked to
                    '  ' ** qryStatementParameters_AssetList_19b (qryStatementParameters_AssetList_19a (Account, with
                    '  ' ** MonthNum = FormRef('MonthNum')), just those matching MonthNum), just scheduled accounts.
                    '  .RecordSource = "qryStatementParameters_AssetList_06g"
                    'Case False
                    '  ' ** Account, linked to ActiveAssets, grouped, with add'l fields; specified FormRef('accountno'), without assetno <> Null.
                    '  .RecordSource = "qryStatementParameters_AssetList_06c"
                    'End Select
                    ' ** qryStatementParameters_AssetList_74_41 (xx), linked to Account, for 'NoAssets'.
800                 .RecordSource = "qryStatementParameters_AssetList_15e"
810               Case False
820                 Select Case frm.chkStatements
                    Case True
                      ' ** qryStatementParameters_AssetList_06b (Account, linked to ActiveAssets, grouped, with add'l
                      ' ** fields; specified FormRef('accountno')), linked to qryStatementParameters_AssetList_19b
                      ' ** (qryStatementParameters_AssetList_19a (Account, with MonthNum = FormRef('MonthNum')),
                      ' ** just those matching MonthNum), just scheduled accounts.
830                   .RecordSource = "qryStatementParameters_AssetList_06f"
840                 Case False
                      ' ** Account, linked to ActiveAssets, grouped, with add'l fields; specified FormRef('accountno').
850                   .RecordSource = "qryStatementParameters_AssetList_06b"
860                 End Select
870               End Select
880             Case frm.opgAccountNumber_optAll.OptionValue
                  ' ** All Accounts.
                  'NO ASSETS!
890               Select Case frm.chkStatements
                  Case True
                    ' ** qryStatementParameters_AssetList_06a (Account, linked to ActiveAssets, grouped, with
                    ' ** add'l fields; all accounts), linked to qryStatementParameters_AssetList_19b
                    ' ** (qryStatementParameters_AssetList_19a (Account, with MonthNum = FormRef('MonthNum')),
                    ' ** just those matching MonthNum), just scheduled accounts.
900                 .RecordSource = "qryStatementParameters_AssetList_06e"
910               Case False
                    ' ** Account, linked to ActiveAssets, grouped, with add'l fields; all accounts.
920                 .RecordSource = "qryStatementParameters_AssetList_06a"
930               End Select
940             End Select  ' ** opgAccountNumber.
950           End Select  ' ** chkRelatedAccounts.
960         End Select  ' ** gblnMessage.
            'Debug.Print "'" & .RecordSource

970         .AsOf_lbl.Caption = "As of " & Format(CDate(frm.AssetListDate), "mmmm dd, yyyy")
980         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
990         .AsOf.Caption = "Printed " & Format(Now(), "mmmm dd"", ""yyyy")
1000        .AsOf_lbl2.Caption = "As of " & Format(CDate(frm.AssetListDate), "mmmm dd, yyyy")
1010        .CoInfoBlock2.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
1020        .AsOf2.Caption = "Printed " & Format(Now(), "mmmm dd"", ""yyyy")

1030        If frm.chkDevSaveData.Visible = True And frm.chkDevSaveData = True Then
1040          blnSaveData = True
1050          lngFlds = 1&
1060          ReDim arr_varFld(F_ELEMS, 0)
              ' ******************************************************
              ' ** Array: arr_varFld()
              ' **
              ' **   Field  Element  Name                 Constant
              ' **   =====  =======  ===================  ==========
              ' **     1       0     accountno            F_ACTNO
              ' **     2       1     st_period            F_DATE
              ' **     3       2     st_forex             F_FOREX
              ' **     4       3     sa_totcost           F_CST
              ' **     5       4     sa_totmarket         F_MKT
              ' **     6       5     sa_icash_end         F_ICSH
              ' **     7       6     sa_pcash_end         F_PCSH
              ' **     8       7     sa_grandtotcost      F_TCST
              ' **     9       8     sa_grandtotmarket    F_TMKT
              ' **    10       9     st_recs              F_RECS
              ' **
              ' ******************************************************
1070          Select Case frm.opgAccountNumber
              Case frm.opgAccountNumber_optSpecified.OptionValue
1080            arr_varFld(F_ACTNO, 0) = strAccountNo
1090            arr_varFld(F_DATE, 0) = datEndDate
1100            arr_varFld(F_FOREX, 0) = CBool(False)
1110            arr_varFld(F_CST, 0) = CDbl(0)
1120            arr_varFld(F_MKT, 0) = CDbl(0)
1130            arr_varFld(F_ICSH, 0) = CDbl(0)
1140            arr_varFld(F_PCSH, 0) = CDbl(0)
1150            arr_varFld(F_TCST, 0) = CDbl(0)
1160            arr_varFld(F_TMKT, 0) = CDbl(0)
1170            arr_varFld(F_RECS, 0) = CLng(0)
1180          Case frm.opgAccountNumber_optAll.OptionValue
1190            arr_varFld(F_ACTNO, 0) = Null
1200            arr_varFld(F_DATE, 0) = datEndDate
1210            arr_varFld(F_FOREX, 0) = CBool(False)
1220            arr_varFld(F_CST, 0) = CDbl(0)
1230            arr_varFld(F_MKT, 0) = CDbl(0)
1240            arr_varFld(F_ICSH, 0) = CDbl(0)
1250            arr_varFld(F_PCSH, 0) = CDbl(0)
1260            arr_varFld(F_TCST, 0) = CDbl(0)
1270            arr_varFld(F_TMKT, 0) = CDbl(0)
1280            arr_varFld(F_RECS, 0) = CLng(0)
1290          End Select
1300        End If

1310        Select Case blnNoAssets
            Case True
1320          .assettype_description_display.Visible = False
1330          .assettype_description_box.Visible = False
1340          .PercentMarket.Visible = False
1350          .totdesc_disp.Visible = True  ' ** Yes, show.
1360          .totdesc_disp.ForeColor = CLR_GRY
1370          .TotalPctMarket_Asset_sum.Visible = False
1380          .UnitCost.ForeColor = CLR_GRY
1390          .TotalCost.ForeColor = CLR_GRY
1400          .UnitMarket.ForeColor = CLR_GRY
1410          .TotalMarket.ForeColor = CLR_GRY
1420          .AnnualIncome.ForeColor = CLR_GRY
1430          .UnitCost_sum_line.BorderColor = CLR_GRY
1440          .UnitMarket_sum_line.BorderColor = CLR_GRY
1450          .TotalCost_Asset_sum_line.BorderColor = CLR_GRY
1460          .TotalMarket_Asset_sum_line.BorderColor = CLR_GRY
1470          .TotalPctMarket_Asset_sum_line.BorderColor = CLR_GRY
1480          .TotalCost_Asset_sum.ForeColor = CLR_GRY
1490          .TotalMarket_Asset_sum.ForeColor = CLR_GRY
1500          .TotalPctMarket_Asset_sum.ForeColor = CLR_GRY
1510        Case False
1520          Select Case gblnCombineAssets
              Case True
1530            .txtIcash.Visible = False
1540            .txtPcash.Visible = False
1550            .txtPcash2.Visible = False
1560            .txtIcash2.Visible = False
1570            .txtPcashPctMarket.Visible = False
1580            .txtIcashPctMarket.Visible = False
1590            .txtCash_lbl.Visible = True
1600            .txtCash.Visible = True
1610            .txtCash2.Visible = True
1620            .txtCashPctMarket.Visible = True
1630          Case False
1640            .txtIcash.Visible = True
1650            .txtPcash.Visible = True
1660            .txtPcash2.Visible = True
1670            .txtIcash2.Visible = True
1680            .txtPcashPctMarket.Visible = True
1690            .txtIcashPctMarket.Visible = True
1700            .txtCash_lbl.Visible = False
1710            .txtCash.Visible = False
1720            .txtCash2.Visible = False
1730            .txtCashPctMarket.Visible = False
1740          End Select  ' ** gblnCombineAssets.
1750        End Select

1760        If glngPrintRpts > 0& Then
1770          For lngZ = 0& To (glngPrintRpts - 1&)
1780            If garr_varPrintRpt(PR_ACTNO, lngZ) = gstrAccountNo Then
1790              If garr_varPrintRpt(PR_ALIST, lngZ) = True Then
1800                blnSkip = True
1810              Else
1820                garr_varPrintRpt(PR_ALIST, lngZ) = CBool(True)
1830              End If
1840              Exit For
1850            End If
1860          Next
1870        End If

1880      Else
1890        Cancel = -1
1900      End If
1910    End With

1920    If Cancel = -1 Then
1930      Beep
1940      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
1950    ElseIf blnSkip = True Then
1960      Cancel = -1
1970    End If

        ' ** VGC 02/14/2010:
        ' ** Existing equation on PercentMarket:
        '=IIf([TotalShareFace_Grand_sum]=0,'0.00%',
        'IIf((([marketvaluecurrentx]*[TotalShareface])/[TotalShareFace_Grand_sum])>1 Or
        '(([marketvaluecurrentx]*[TotalShareface])/[TotalShareFace_Grand_sum])<-1,'NA',
        'IIf(IsNull([marketvaluecurrentx])=True Or IsNull([TotalShareface])=True,'0.00%',
        'Format((([marketvaluecurrentx]*[TotalShareface])/[TotalShareFace_Grand_sum]),'Percent'))))
        ' ** Changed <-1 to <0.
        ' ** Now:
        '=IIf([TotalMarket_Grand_sum]=0,'0.00%',IIf((([MarketValueCurrentX]*[TotalShareface])/[TotalMarket_Grand_sum])>1 Or (([MarketValueCurrentX]*[TotalShareface])/[TotalMarket_Grand_sum])<0,'NA',IIf(IsNull([MarketValueCurrentX])=True Or IsNull([TotalShareface])=True,'0.00%',Format((([MarketValueCurrentX]*[TotalShareface])/[TotalMarket_Grand_sum]),'Percent'))))

        ' ** Existing equation on TotalPctMarket_Asset_sum:
        '=IIf(IsNull([TotalShareFace_Grand_sum])=True Or IsNull(Sum([marketvaluecurrentx]))=True Or IsNull(Sum([TotalShareface]))=True,0,
        'IIf([TotalShareFace_Grand_sum]=0,0,[percentMarket2x]/[TotalShareFace_Grand_sum]))
        ' ** Added IIf() to last calculation.
        ' ** Now:
        '=IIf(IsNull([TotalMarket_Grand_sum])=True Or IsNull(Sum([MarketValueCurrentX]))=True Or IsNull(Sum([TotalShareface]))=True,0,IIf([TotalMarket_Grand_sum]=0,0,IIf(([percentMarket2x]/[TotalMarket_Grand_sum])>1 Or ([percentMarket2x]/[TotalMarket_Grand_sum])<0,'NA',([percentMarket2x]/[TotalMarket_Grand_sum]))))

        ' ** Existing equation on TotalPctMarket_sum:
        '=IIf([TotalShareFace_Grand_sum]=0,0,IIf(([percentMarket3x]/[TotalShareFace_Grand_sum])=0,
        '((1-(IIf([TotalShareFace_Grand_sum]=0,0,IIf(IsNull([icash]),0,[icash])/[TotalShareFace_Grand_sum])))-
        '(IIf([TotalShareFace_Grand_sum]=0,0,IIf(IsNull([pcash]),0,[pcash])/[TotalShareFace_Grand_sum]))),
        '([percentMarket3x]/[TotalShareFace_Grand_sum])))
        ' ** Now:
        '=IIf(Nz([TotalMarket_Grand_sum],0)=0,0,IIf((IIf((Nz([percentMarket3x],0)/[TotalMarket_Grand_sum])=0,
        '((1-(Nz([icash],0)/[TotalMarket_Grand_sum]))-(Nz([pcash],0)/[TotalMarket_Grand_sum])),
        '([percentMarket3x]/[TotalMarket_Grand_sum])))>1 Or
        '(IIf((Nz([percentMarket3x],0)/[TotalMarket_Grand_sum])=0,
        '((1-(Nz([icash],0)/[TotalMarket_Grand_sum]))-(Nz([pcash],0)/[TotalMarket_Grand_sum])),
        '([percentMarket3x]/[TotalMarket_Grand_sum])))<0, 'NA',
        'IIf((Nz([percentMarket3x],0)/[TotalMarket_Grand_sum])=0,
        '((1-(Nz([icash],0)/[TotalMarket_Grand_sum]))-(Nz([pcash],0)/[TotalMarket_Grand_sum])),
        '([percentMarket3x]/[TotalMarket_Grand_sum]))))

        ' ** Existing equation on txtIcashPctMarket:
        '=IIf([TotalShareFace_Grand_sum]=0,0,IIf(IsNull([icash]),0,[icash])/[TotalShareFace_Grand_sum])
        ' ** Now:
        '=IIf([TotalMarket_Grand_sum]=0,0,IIf((IIf(IsNull([icash]),0,[icash])/[TotalMarket_Grand_sum])>1 Or
        '(IIf(IsNull([icash]),0,[icash])/[TotalMarket_Grand_sum])<0,'NA',(IIf(IsNull([icash]),0,[icash])/[TotalMarket_Grand_sum])))

        ' ** Existing equation on txtCashPctMarket:
        '=IIf([TotalShareFace_Grand_sum]=0,0,([icash]+[pcash])/[TotalShareFace_Grand_sum])
        ' ** Now:
        '=IIf([TotalMarket_Grand_sum]=0,0,IIf((([icash]+[pcash])/[TotalMarket_Grand_sum])>1 Or (([icash]+[pcash])/[TotalMarket_Grand_sum])<0,'NA',(([icash]+[pcash])/[TotalMarket_Grand_sum])))

        ' ** Existing equation on txtPcashPctMarket:
        '=IIf([TotalShareFace_Grand_sum]=0,0,IIf(IsNull([pcash]),0,[pcash])/[TotalShareFace_Grand_sum])
        ' ** Now:
        '=IIf([TotalMarket_Grand_sum]=0,0,IIf((IIf(IsNull([pcash]),0,[pcash])/[TotalMarket_Grand_sum])>1 Or (IIf(IsNull([pcash]),0,[pcash])/[TotalMarket_Grand_sum])<0,'NA',(IIf(IsNull([pcash]),0,[pcash])/[TotalMarket_Grand_sum])))

        ' ** Existing equation on TotalPctMarket_Grand_sum:
        '=IIf([TotalMarket_Grand_sum]=0,0,(Sum([MarketValueCurrentX]*[TotalShareface])+IIf(IsNull([icash]),0,[icash])+
        'IIf(IsNull([pcash]),0,[pcash]))/[TotalMarket_Grand_sum])
        ' ** Now:
        '=IIf([TotalMarket_Grand_sum]=0,0,IIf(((Sum([marketvaluecurrentx]*[TotalShareface])+
        'IIf(IsNull([icash]),0,[icash])+IIf(IsNull([pcash]),0,[pcash]))/[TotalMarket_Grand_sum])>1 Or
        '((Sum([marketvaluecurrentx]*[TotalShareface])+IIf(IsNull([icash]),0,[icash])+
        'IIf(IsNull([pcash]),0,[pcash]))/[TotalMarket_Grand_sum])<0,'NA',
        '((Sum([marketvaluecurrentx]*[TotalShareface])+
        'IIf(IsNull([icash]),0,[icash])+IIf(IsNull([pcash]),0,[pcash]))/[TotalMarket_Grand_sum])))

        ' ** The above mess for TotalPctMarket_sum, parsed:
        '      =IIf
'1         (
        ' if        [TotalShareFace_Grand_sum]=0,
        ' then      0,
        ' else      IIf
'2             (
        '     if        ([percentMarket3x]/[TotalShareFace_Grand_sum])=0,
'3    then      (
'4               (
        '                 1-
'5                  (
        '                    IIf
'6                      (
        '              if        [TotalShareFace_Grand_sum]=0,
        '              then      0,
        '              else      IIf                             PAREN THE WHOLE ELSE CLAUSE: HERE
'7                          (
        '                  if        IsNull([icash]),
        '                  then      0,
        '                  else      [icash]
'7                          )/
        '                        [TotalShareFace_Grand_sum]      TO HERE
'6                      )
'5                  )
'4               )-
'8               (
        '                 IIf
'9                   (
        '           if        [TotalShareFace_Grand_sum]=0,
        '           then      0,
        '           else      IIf                           PAREN THE WHOLE ELSE CLAUSE: HERE
'10                      (
        '               if        IsNull([pcash]),
        '               then      0,
        '               else      [pcash]
'10                      )/
        '                     [TotalShareFace_Grand_sum]    TO HERE
'9                   )
'8               )
'3              ),
        '     else      (
        '                [percentMarket3x]/[TotalShareFace_Grand_sum]         <-- GIVE THIS PARENS
        '               )
'2             )
'1         )

        ' ** The above cool diagram, broken into pieces:
        '{parens #1}
        'If [TotalShareFace_Grand_sum] = 0 Then
        '  0
        'Else
        '  {parens #2}
        'End If

        '{parens #2}
        '  If ([percentMarket3x]/[TotalShareFace_Grand_sum]) = 0 Then
        '    {parens #3}
        '  Else
        '    [percentMarket3x]/[TotalShareFace_Grand_sum]                     <-- IF THE NUMBERS AREN'T ZERO, IT GOES RIGHT TO HERE
        '  End If

        '{parens #3}
        '    ((1 - {parens #5}) - {parens #8})  {enveloped by parens #4}

        '{parens #5 envolopes parens #6}

        '{parens #6}
        '      If [TotalShareFace_Grand_sum] = 0 Then                         <-- THIS IS REDUNDENT {see parens #1}!!
        '        0
        '      Else
        '        {parens #7} / [TotalShareFace_Grand_sum]
        '      End If

        '{parens #7}
        '         If IsNull([icash]) = True Then
        '           0
        '         Else
        '           [icash]
        '         End If

        '{parens #8 envolops parens #9}

        '{parens #9}
        '         If [TotalShareFace_Grand_sum] = 0 Then                      <-- THIS IS REDUNDENT {see parens #1}!!
        '           0
        '         Else
        '           {parens #10} / [TotalShareFace_Grand_sum]
        '         End if

        '{parens #10}
        '           If IsNull([pcash]) = True Then
        '             0
        '           Else
        '             [pcash]
        '           End If

        ' ** Now, in English:
        'If [TotalShareFace_Grand_sum] is Zero, stop and return Zero.
        'Otherwise
        '  If [percentMarket3x] divided by [TotalShareFace_Grand_sum] is Zero, then
        '    Subtract {this} from one, then, from that difference, subtract {the other number}
        '      or
        '    Subtract [icash] divided by [TotalShareFace_Grand_sum] from one, then, from that difference,
        '      subtract [pcash] divided by [TotalShareFace_Grand_sum].
        '    and so we stop and return
        '      (1 - ([icash] / [TotalShareFace_Grand_sum])) - ([pcash] / [TotalShareFace_Grand_sum])
        '  Otherwise
        '    stop and return [percentMarket3x] divided by [TotalShareFace_Grand_sum].

        ' ** As programming:
        'If [TotalShareFace_Grand_sum] = 0 Then
        '  dblRetVal = 0
        'Else
        '  If ([percentMarket3x] / [TotalShareFace_Grand_sum]) = 0 Then
        '    dblRetVal = ([icash] / [TotalShareFace_Grand_sum])
        '    dblRetVal = (1 - dblRetVal)
        '    dblRetVal = (dblRetVal - ([pcash] / [TotalShareFace_Grand_sum]))
        '  Else
        '    dblRetVal = ([percentMarket3x] / [TotalShareFace_Grand_sum])
        '  End If
        'End If

        'Nz() everthing!

        'IIf( Nz([TotalShareFace_Grand_sum],0) = 0, 0,
        '  IIf( (Nz([percentMarket3x],0) / [TotalShareFace_Grand_sum]) = 0,
        '    ((1 - (Nz([icash],0) / [TotalShareFace_Grand_sum])) - (Nz([pcash],0) / [TotalShareFace_Grand_sum])) ,
        '    ([percentMarket3x] / [TotalShareFace_Grand_sum])
        '  )
        ')
        ' ** Voila!
        '=IIf(Nz([TotalShareFace_Grand_sum],0)=0,0,IIf((Nz([percentMarket3x],0)/[TotalShareFace_Grand_sum])=0,((1-(Nz([icash],0)/[TotalShareFace_Grand_sum]))-(Nz([pcash],0)/[TotalShareFace_Grand_sum])),([percentMarket3x]/[TotalShareFace_Grand_sum])))

EXITP:
1980    Exit Sub

ERRH:
1990    If Forms(strCallingForm).Visible = False Then
2000      Forms(strCallingForm).Visible = True
2010    End If
2020    Select Case ERR.Number
        Case 2427, -2147352567  ' ** You entered an expression that has no value.
          ' ** No data.
2030    Case 2585  ' ** This action can't be carried out while processing a form or report event.
          ' ** Ignore.
2040    Case Else
2050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2060    End Select
2070    Resume EXITP

End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "ReportHeader_Format"

        Dim datCurrentDate As Date
        Dim varTmp00 As Variant

2110    With Me

2120      Select Case IsNull(frm.currentDate)
          Case True
2130        varTmp00 = DMax("[currentDate]", "masterasset")
2140        datCurrentDate = varTmp00
2150      Case False
2160        datCurrentDate = frm.currentDate
2170      End Select
2180      If frm.UsePriceHistory = True Then
2190        .MarketAsOf_lbl.Caption = "Market as of varies"
2200        .MarketAsOf_lbl2.Caption = "Market as of varies"
2210      Else
2220        If datCurrentDate <> datEndDate Then
2230          .MarketAsOf_lbl.Caption = "Market as of " & Format(datCurrentDate, "mm/dd/yyyy")
2240          .MarketAsOf_lbl2.Caption = "Market as of " & Format(datCurrentDate, "mm/dd/yyyy")
2250        Else
2260          .MarketAsOf_lbl.Caption = vbNullString
2270          .MarketAsOf_lbl2.Caption = vbNullString
2280        End If
2290      End If

2300      If frm.chkStatements = True Then  ' ** Doing statements.
2310        .AsOf.Visible = False
2320        .AsOf2.Visible = False
2330      End If

2340    End With

EXITP:
2350    Exit Sub

ERRH:
2360    If Forms(strCallingForm).Visible = False Then
2370      Forms(strCallingForm).Visible = True
2380    End If
2390    Select Case ERR.Number
        Case Else
2400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2410    End Select
2420    Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

2510    With Me
2520      If .Page >= 2 Then
2530        .Section(acPageHeader).Visible = True
2540      Else
2550        .Section(acPageHeader).Visible = False
2560      End If
2570    End With

EXITP:
2580    Exit Sub

ERRH:
2590    If Forms(strCallingForm).Visible = False Then
2600      Forms(strCallingForm).Visible = True
2610    End If
2620    Select Case ERR.Number
        Case Else
2630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2640    End Select
2650    Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

        Dim blnFound As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long, lngE As Long

2710    With Me

2720      Select Case frm.chkLegalName
          Case True
2730        varTmp00 = FixAmps(Nz(.legalname, vbNullString))  ' ** Module Functions: modStringFuncs.
2740        If IsNull(varTmp00) = True Then
2750          varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
2760        End If
2770        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
2780        .Name_lbl.Caption = varTmp00
2790      Case False
2800        varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
2810        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
2820        .Name_lbl.Caption = varTmp00
2830      End Select  ' ** chkLegalName.

2840      If blnSaveData = True And blnStatements = True Then
2850        Select Case IsNull(arr_varFld(F_ACTNO, 0))
            Case True
              ' ** First accountno.
2860          arr_varFld(F_ACTNO, 0) = .accountno
2870        Case False
2880          blnFound = False
2890          For lngX = 0& To (lngFlds - 1&)
2900            If arr_varFld(F_ACTNO, lngX) = .accountno Then
2910              blnFound = True
2920              Exit For
2930            End If
2940          Next
2950          If blnFound = False Then
2960            lngFlds = lngFlds + 1&
2970            lngE = lngFlds - 1&
2980            ReDim Preserve arr_varFld(F_ELEMS, lngE)
2990            arr_varFld(F_ACTNO, lngE) = .accountno
3000            arr_varFld(F_DATE, lngE) = datEndDate
3010            arr_varFld(F_FOREX, lngE) = CBool(False)
3020            arr_varFld(F_CST, lngE) = CDbl(0)
3030            arr_varFld(F_MKT, lngE) = CDbl(0)
3040            arr_varFld(F_ICSH, lngE) = CDbl(0)
3050            arr_varFld(F_PCSH, lngE) = CDbl(0)
3060            arr_varFld(F_TCST, lngE) = CDbl(0)
3070            arr_varFld(F_TMKT, lngE) = CDbl(0)
3080            arr_varFld(F_RECS, lngE) = CLng(0)
3090          End If
3100        End Select
3110      End If

3120    End With

EXITP:
3130    Exit Sub

ERRH:
3140    If Forms(strCallingForm).Visible = False Then
3150      Forms(strCallingForm).Visible = True
3160    End If
3170    Select Case ERR.Number
        Case 2427, -2147352567  ' ** You entered an expression that has no value.
          ' ** No data.
3180    Case Else
3190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3200    End Select
3210    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

        Dim lngX As Long

3310    With Me
3320      If blnSaveData = True Then
3330        For lngX = 0& To (lngFlds - 1&)
3340          If arr_varFld(F_ACTNO, lngX) = .accountno Then
3350            arr_varFld(F_CST, lngX) = .TotalCost_sum
3360            arr_varFld(F_MKT, lngX) = .TotalMarket_sum
3370            arr_varFld(F_ICSH, lngX) = .txtIcash
3380            arr_varFld(F_PCSH, lngX) = .txtPcash
3390            arr_varFld(F_TCST, lngX) = .TotalCost_Grand_sum
3400            arr_varFld(F_TMKT, lngX) = .TotalMarket_Grand_sum
3410            Exit For
3420          End If
3430        Next
3440      End If
3450    End With

EXITP:
3460    Exit Sub

ERRH:
3470    If Forms(strCallingForm).Visible = False Then
3480      Forms(strCallingForm).Visible = True
3490    End If
3500    Select Case ERR.Number
        Case Else
3510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3520    End Select
3530    Resume EXITP

End Sub

Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "ReportFooter_Format"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strAssetListDate As String, datBalanceDate As Date
        Dim dblTotCost As Double, dblTotMkt As Double, dblICash As Double, dblPCash As Double, dblGrandCost As Double, dblGrandMkt As Double
        Dim blnSkip As Boolean

3610    With Me
3620      blnSkip = True
3630      If blnSkip = False Then
3640        strAccountNo = .accountno
3650        strAssetListDate = .AsOf_lbl.Caption
3660        strAssetListDate = GetLastWord(strAssetListDate)  ' ** Module Function: modStringFuncs.
3670        datBalanceDate = CDate(strAssetListDate)
3680        dblTotCost = .TotalCost_sum
3690        dblTotMkt = .TotalMarket_sum
3700        dblICash = .txtIcash
3710        dblPCash = .txtPcash
3720        dblGrandCost = .TotalCost_Grand_sum
3730        dblGrandMkt = .TotalMarket_Grand_sum
3740        Set dbs = CurrentDb
3750        With dbs
3760          Set rst = .OpenRecordset("zz_tbl_MasterTrust_Balance_01", dbOpenDynaset, dbAppendOnly)
3770          With rst
3780            .AddNew
                ' ** ![mtbal_id] : AutoNumber.
3790            ![accountno] = strAccountNo
3800            ![balance_date] = datBalanceDate
3810            ![total_cost] = dblTotCost
3820            ![total_market] = dblTotMkt
3830            ![ICash] = dblICash
3840            ![PCash] = dblPCash
3850            ![grand_cost] = dblGrandCost
3860            ![grand_market] = dblGrandMkt
3870            ![mtbal_datemodified] = Now()
3880            .Update
3890            .Close
3900          End With
3910          .Close
3920        End With
3930      End If  ' ** blnSkip.
3940    End With

EXITP:
3950    Set rst = Nothing
3960    Set dbs = Nothing
3970    Exit Sub

ERRH:
3980    If Forms(strCallingForm).Visible = False Then
3990      Forms(strCallingForm).Visible = True
4000    End If
4010    Select Case ERR.Number
        Case Else
4020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4030    End Select
4040    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

4110    With Me
4120      .NoData_lbl.Visible = True
4130      .assettype_description_display.Visible = False
4140      .accountno.Visible = False
4150      .Name_lbl.Visible = False
4160      .TotalShareface1.Visible = False
4170      .totdesc_disp.Visible = False
4180      .UnitCost.Visible = False
4190      .TotalCost.Visible = False
4200      .UnitMarket.Visible = False
4210      .TotalMarket.Visible = False
4220      .PercentMarket.Visible = False
4230      .AnnualIncome.Visible = False
4240      .TotalCost_Asset_sum.Visible = False
4250      .TotalMarket_Asset_sum.Visible = False
4260      .TotalPctMarket_Asset_sum.Visible = False
4270      .TotalCost_sum.Visible = False
4280      .txtIcash.Visible = False
4290      .txtCash.Visible = False
4300      .txtPcash.Visible = False
4310      .TotalCost_Grand_sum.Visible = False
4320      .TotalMarket_sum.Visible = False
4330      .txtIcash2.Visible = False
4340      .txtCash2.Visible = False
4350      .txtPcash2.Visible = False
4360      .TotalMarket_Grand_sum.Visible = False
4370      .TotalPctMarket_sum.Visible = False
4380      .txtIcashPctMarket.Visible = False
4390      .txtCashPctMarket.Visible = False
4400      .txtPcashPctMarket.Visible = False
4410      .TotalPctMarket_Grand_sum.Visible = False
4420      .marketvaluecurrent.Visible = False
4430    End With

EXITP:
4440    Exit Sub

ERRH:
4450    If Forms(strCallingForm).Visible = False Then
4460      Forms(strCallingForm).Visible = True
4470    End If
4480    Select Case ERR.Number
        Case Else
4490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4500    End Select
4510    Resume EXITP

End Sub

Private Sub Report_Close()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim blnAdd As Boolean, blnAddAll As Boolean
        Dim lngX As Long

4610    With Me

4620      Set frm = Nothing
4630      TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

4640      If blnSaveData = True Then
4650        blnAdd = False: blnAddAll = False
4660        Set dbs = CurrentDb
4670        With dbs
4680          Set rst = .OpenRecordset("zz_tbl_Statement_AssetList", dbOpenDynaset, dbConsistent)
4690          With rst
4700            If .BOF = True And .EOF = True Then
4710              blnAddAll = True
4720            End If
4730            For lngX = 0& To (lngFlds - 1&)
4740              If IsNull(arr_varFld(F_ACTNO, lngX)) = False Then
4750                If Trim(arr_varFld(F_ACTNO, lngX)) <> vbNullString Then
4760                  blnAdd = False
4770                  Select Case blnAddAll
                      Case True
4780                    blnAdd = True
4790                  Case False
4800                    .FindFirst "[accountno] = '" & arr_varFld(F_ACTNO, lngX) & "' And " & _
                          "[sa_period] = #" & Format(arr_varFld(F_DATE, lngX), "mm/dd/yyyy") & "# And " & _
                          "[sa_forex] = False"
4810                    If .NoMatch = True Then
4820                      blnAdd = True
4830                    End If
4840                  End Select
4850                  Select Case blnAdd
                      Case True
4860                    .AddNew
4870                    ![accountno] = arr_varFld(F_ACTNO, lngX)
4880                    ![sa_period] = arr_varFld(F_DATE, lngX)
4890                    ![sa_forex] = arr_varFld(F_FOREX, lngX)
4900                  Case False
4910                    .Edit
4920                  End Select
4930                  ![sa_totcost] = arr_varFld(F_CST, lngX)
4940                  ![sa_totmarket] = arr_varFld(F_MKT, lngX)
4950                  ![sa_icash_end] = arr_varFld(F_ICSH, lngX)
4960                  ![sa_pcash_end] = arr_varFld(F_PCSH, lngX)
4970                  ![sa_grandtotcost] = arr_varFld(F_TCST, lngX)
4980                  ![sa_grandtotmarket] = arr_varFld(F_TMKT, lngX)
4990                  ![sa_recs] = arr_varFld(F_RECS, lngX)
5000                  ![sa_datemodified] = Now()
5010                  .Update
5020                End If
5030              End If
5040            Next
5050            .Close
5060          End With
5070          .Close
5080        End With
5090      End If

5100      If strCallingForm <> vbNullString Then
5110        gblnReportClose = True
5120        If Forms(strCallingForm).Visible = False Then
5130          DoCmd.Restore
5140          Forms(strCallingForm).Visible = True
5150        End If
5160        Forms(strCallingForm).TimerInterval = 250&
5170      End If

5180    End With

EXITP:
5190    Set rst = Nothing
5200    Set dbs = Nothing
5210    Exit Sub

ERRH:
5220    If Forms(strCallingForm).Visible = False Then
5230      Forms(strCallingForm).Visible = True
5240    End If
5250    Select Case ERR.Number
        Case Else
5260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5270    End Select
5280    Resume EXITP

End Sub

Private Sub VerifyDates()

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "VerifyDates"

        Dim varTmp00 As Variant

5310    If IsNull(frm.DateStart) = True Then
5320      If CDbl(gdatStartDate) > 0 Then
5330        If gdatEndDate = frm.AssetListDate Then
5340          frm.DateStart = gdatStartDate
5350        Else
5360          If strAccountNo = "All" Then
5370            frm.DateStart = #1/1/1900#
5380          Else
5390            varTmp00 = DLookup("[documentdate]", "account", "[accountno] = '" & strAccountNo & "'")
5400            If IsNull(varTmp00) = True Then
5410              varTmp00 = DLookup("[appointmentdate]", "account", "[accountno] = '" & strAccountNo & "'")
5420              If IsNull(varTmp00) = True Then
5430                frm.DateStart = #1/1/1900#
5440              Else
5450                frm.DateStart = CDate(varTmp00)
5460              End If
5470            Else
5480              frm.DateStart = CDate(varTmp00)
5490            End If
5500          End If
5510        End If
5520      Else
5530        If strAccountNo = "All" Then
5540          frm.DateStart = #1/1/1900#
5550        Else
5560          varTmp00 = DLookup("[documentdate]", "account", "[accountno] = '" & strAccountNo & "'")
5570          If IsNull(varTmp00) = True Then
5580            varTmp00 = DLookup("[appointmentdate]", "account", "[accountno] = '" & strAccountNo & "'")
5590            If IsNull(varTmp00) = True Then
5600              frm.DateStart = #1/1/1900#
5610            Else
5620              frm.DateStart = CDate(varTmp00)
5630            End If
5640          Else
5650            frm.DateStart = CDate(varTmp00)
5660          End If
5670        End If
5680      End If
5690    End If
5700    DoEvents
5710    If IsNull(frm.AssetListDate) = True Then
5720      If IsNull(frm.DateEnd) = True Then
5730        If CDbl(gdatEndDate) > 0 Then
5740          If gdatEndDate > frm.DateStart Then
5750            frm.AssetListDate = gdatEndDate
5760          Else
5770            frm.AssetListDate = Date
5780          End If
5790        Else
5800          frm.AssetListDate = Date
5810        End If
5820      Else
5830        If frm.DateEnd > frm.DateStart Then
5840          frm.AssetListDate = frm.DateEnd
5850        Else
5860          frm.AssetListDate = Date
5870        End If
5880      End If
5890    End If
5900    DoEvents

EXITP:
5910    Exit Sub

ERRH:
5920    If Forms(strCallingForm).Visible = False Then
5930      Forms(strCallingForm).Visible = True
5940    End If
5950    Select Case ERR.Number
        Case 2585  ' ** This action can't be carried out while processing a form or report event.
          ' ** Ignore.
5960    Case Else
5970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5980    End Select
5990    Resume EXITP

End Sub
