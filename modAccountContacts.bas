Attribute VB_Name = "modAccountContacts"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modAccountContacts"

'VGC 07/18/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   {In Parent}
' #######################################

' ***************************
' ***************************
' ** COUNTRY NOT DONE YET!  Excel!
' ***************************
' ***************************

' ** Array: arr_varCtl().
Private lngCtls As Long, arr_varCtl() As Variant
Private Const C_ELEMS As Integer = 15  ' ** Array's first-element UBound().
Private Const C_CNAM   As Integer = 0
Private Const C_TOP    As Integer = 1
Private Const C_LFT    As Integer = 2
Private Const C_WDT    As Integer = 3
Private Const C_HGT    As Integer = 4
Private Const C_L1_NAM As Integer = 5
Private Const C_L1_LFT As Integer = 6
Private Const C_L2_NAM As Integer = 7
Private Const C_L2_LFT As Integer = 8
Private Const C_LN_NAM As Integer = 9
Private Const C_LN_TOP As Integer = 10
Private Const C_LN_LFT As Integer = 11
Private Const C_F1_NAM As Integer = 12
Private Const C_F1_WDT As Integer = 13
Private Const C_F2_NAM As Integer = 14
Private Const C_F2_WDT As Integer = 15

' ** Constants for both opgPrintCSZ and opgPrintCSZCP.
Private Const OPT_SEP As Integer = 1
Private Const OPT_COM As Integer = 2
Private Const OPT_NON As Integer = 3

' ** Array: arr_varRptFld().
Private lngRptFlds As Long, arr_varRptFld() As Variant
Private Const R_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const R_FID  As Integer = 0
Private Const R_FNAM As Integer = 1
Private Const R_VIS  As Integer = 2

Private lngMonitorCnt As Long, lngMonitorNum As Long
Private lngRecsCur As Long, lngTpp As Long
' **

Public Function SetupRptFlds(strCallingForm As String, strReport As String) As Boolean

100   On Error GoTo ERRH

        Const THIS_PROC As String = "SetupRptFlds"

        Dim frm As Access.Form, rpt As Access.Report
        Dim lngRptWidth As Long
        Dim lngFldSep As Long, strFlds As String
        Dim lngFld01 As Long, lngFld02 As Long, lngFld03 As Long, lngFld04 As Long, lngFld05 As Long
        Dim lngFld06 As Long, lngFld07 As Long, lngFld08 As Long, lngFld09 As Long, lngFld10 As Long
        Dim lngFld11 As Long, lngFld12 As Long, lngFld13 As Long
        Dim blnFld11TooWide As Boolean, blnFld12TooWide As Boolean, blnFld13TooWide As Boolean, blnFld14TooWide As Boolean
        Dim strFldTooWide As String
        Dim intOptCSZ As Integer
        Dim lngRightMost As Long
        Dim lngTmp01 As Long, blnTmp01 As Boolean
        Dim lngX As Long
        Dim blnRetVal As Boolean

110     blnRetVal = True

120     Set rpt = Reports(strReport)
130     Set frm = Forms(strCallingForm)

140     With rpt

150       lngRptWidth = .Width

160       lngFldSep = (.Contact_Number.Left - (.accountno.Left + .accountno.Width))
170       lngFld01 = 0&: lngFld02 = 0&: lngFld03 = 0&: lngFld04 = 0&: lngFld05 = 0&: lngFld06 = 0&
180       lngFld07 = 0&: lngFld08 = 0&: lngFld09 = 0&: lngFld10 = 0&: lngFld11 = 0&: lngFld12 = 0&
190       strFlds = vbNullString
200       lngRightMost = 0&

210       If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
            'frm.RightMost = 0&
            'frm.DidntShow = vbNullString
220       End If

          ' *************************************************
          ' ** Array: arr_varRptFld()
          ' **
          ' **   Field  Element  Name            Constant
          ' **   =====  =======  ==============  ==========
          ' **     1       0     ctl_name        R_FNAM
          ' **     2       1     Visible         R_VIS
          ' **
          ' *************************************************
230       RptFld_Load  ' ** Function: Below.

          ' ** Start with everything off.
240       For lngX = 0& To (lngRptFlds - 1&)
250         .Controls(arr_varRptFld(R_FNAM, lngX)).Visible = False
260         .Controls(arr_varRptFld(R_FNAM, lngX) & "_lbl").Visible = False
270         .Controls(arr_varRptFld(R_FNAM, lngX) & "_lbl_line").Visible = False
280       Next

290       lngFld01 = .accountno.Left

          ' **********************************************************
          ' ** This first section is just calculating what will fit.
          ' **********************************************************

300       Select Case frm.chkShowAcctNum
          Case True
310         strFlds = strFlds & "accountno;"
320       Case False
330         strFlds = strFlds & vbNullString
340       End Select

350       Select Case frm.chkShowShortName
          Case True
360         strFlds = strFlds & "shortname;"
370         Select Case frm.chkShowAcctNum
            Case True
380           lngTmp01 = ((lngFld01 + .accountno.Width) + lngFldSep)
390         Case False
400           lngTmp01 = lngFld01
410         End Select
420         lngFld02 = lngTmp01
430         lngTmp01 = ((lngFld02 + .shortname.Width) + lngFldSep)
440       Case False
450         strFlds = strFlds & vbNullString
460         lngFld02 = 0&
470         lngTmp01 = ((lngFld01 + .accountno.Width) + lngFldSep)
480       End Select
490       lngFld03 = lngTmp01

500       strFlds = strFlds & "Contact_Number;"
510       lngTmp01 = ((lngFld03 + .Contact_Number.Width) + lngFldSep)
520       lngFld04 = lngTmp01

530       Select Case frm.opgFormatName
          Case frm.opgFormatName_optAsWritten.OptionValue
540         strFlds = strFlds & "Contact_Name;"
550         lngTmp01 = ((lngFld04 + .Contact_Name.Width) + lngFldSep)
560       Case frm.opgFormatName_optLastFirst.OptionValue
570         strFlds = strFlds & "Contact_Name_LastFirst;"
580         lngTmp01 = ((lngFld04 + .Contact_Name_LastFirst.Width) + lngFldSep)
590       End Select
600       lngFld05 = lngTmp01

610       Select Case frm.opgPrintAddress
          Case frm.opgPrintAddress_optSeparate.OptionValue
620         strFlds = strFlds & "Contact_Address1;"
630         lngTmp01 = ((lngFld05 + .Contact_Address1.Width) + lngFldSep)
640         lngFld06 = lngTmp01
650         strFlds = strFlds & "Contact_Address2;"
660         lngTmp01 = ((lngFld06 + .Contact_Address2.Width) + lngFldSep)
670         lngFld07 = lngTmp01
680       Case frm.opgPrintAddress_optCombined.OptionValue
690         strFlds = strFlds & "Contact_Address_Combined;"
700         lngFld06 = 0&
710         lngTmp01 = ((lngFld05 + .Contact_Address_Combined.Width) + lngFldSep)
720         lngFld07 = lngTmp01
730       Case frm.opgPrintAddress_optNone.OptionValue
740         lngFld06 = 0&
750         lngTmp01 = lngFld05
760         lngFld07 = lngTmp01
770       End Select

780       blnTmp01 = False
790       If frm.chkEnableCountry_Wide.Visible = True Then
800         blnTmp01 = frm.chkEnableCountry_Wide
810       ElseIf frm.chkEnableCountry_Compact.Visible = True Then
820         blnTmp01 = frm.chkEnableCountry_Compact
830       End If

840       Select Case blnTmp01
          Case True
850         intOptCSZ = frm.opgPrintCSZCP
860       Case False
870         intOptCSZ = frm.opgPrintCSZ
880       End Select

          'COUNTRY NOT DONE!
890       Select Case intOptCSZ
          Case OPT_SEP
900         strFlds = strFlds & "Contact_City;"
910         lngTmp01 = ((lngFld07 + .Contact_City.Width) + lngFldSep)
920         lngFld08 = lngTmp01
930         strFlds = strFlds & "Contact_State;"
940         lngTmp01 = ((lngFld08 + .Contact_State.Width) + lngFldSep)
950         lngFld09 = lngTmp01
960         Select Case frm.opgFormatZip
            Case frm.opgFormatZip_optFormatted.OptionValue
970           strFlds = strFlds & "Contact_Zip_Format;"
980           lngTmp01 = ((lngFld09 + .Contact_Zip_Format.Width) + lngFldSep)
990         Case frm.opgFormatZip_optUnformatted.OptionValue
1000          strFlds = strFlds & "Contact_Zip;"
1010          lngTmp01 = ((lngFld09 + .Contact_Zip.Width) + lngFldSep)
1020        End Select
1030        lngFld10 = lngTmp01
1040      Case OPT_COM
1050        Select Case frm.opgFormatZip
            Case frm.opgFormatZip_optFormatted.OptionValue
1060          strFlds = strFlds & "Contact_CSZ_Format;"
1070          lngTmp01 = ((lngFld07 + .Contact_CSZ_Format.Width) + lngFldSep)
1080        Case frm.opgFormatZip_optUnformatted.OptionValue
1090          strFlds = strFlds & "Contact_CSZ;"
1100          lngTmp01 = ((lngFld07 + .Contact_CSZ.Width) + lngFldSep)
1110          lngFld10 = lngTmp01
1120        End Select
1130        lngFld08 = 0&
1140        lngFld09 = 0&
1150      Case OPT_NON
1160        lngFld08 = 0&
1170        lngFld09 = 0&
1180        lngTmp01 = lngFld07
1190        lngFld10 = lngTmp01
1200      End Select

          ' ** Up to this point, no combination is too wide.
1210      lngRightMost = (lngFld10 - lngFldSep)

1220      blnFld11TooWide = False: blnFld12TooWide = False
1230      Select Case frm.opgPrintPhone
          Case frm.opgPrintPhone_optSeparate.OptionValue
1240        Select Case frm.opgFormatPhone
            Case frm.opgFormatPhone_optFormatted.OptionValue
1250          strFlds = strFlds & "Contact_Phone1_Format;"
1260          lngTmp01 = ((lngFld10 + .Contact_Phone1_Format.Width) + lngFldSep)
1270          lngFld11 = lngTmp01
1280          If lngFld11 > (lngRptWidth + lngFldSep) Then
1290            blnFld11TooWide = True
1300            strFldTooWide = "Contact_Phone1_Format"
1310          Else
1320            strFlds = strFlds & "Contact_Phone2_Format;"
1330            lngTmp01 = ((lngFld11 + .Contact_Phone2_Format.Width) + lngFldSep)
1340            lngFld12 = lngTmp01
1350            If lngFld12 > (lngRptWidth + lngFldSep) Then
1360              blnFld12TooWide = True
1370              strFldTooWide = "Contact_Phone2_Format"
1380              lngRightMost = (lngFld11 - lngFldSep)
1390            Else
1400              lngRightMost = (lngFld12 - lngFldSep)
1410            End If
1420          End If
1430        Case frm.opgFormatPhone_optUnformatted.OptionValue
1440          strFlds = strFlds & "Contact_Phone1;"
1450          lngTmp01 = ((lngFld10 + .Contact_Phone1.Width) + lngFldSep)
1460          lngFld11 = lngTmp01
1470          If lngFld11 > (lngRptWidth + lngFldSep) Then
1480            blnFld11TooWide = True
1490            strFldTooWide = "Contact_Phone1"
1500          Else
1510            strFlds = strFlds & "Contact_Phone2;"
1520            lngTmp01 = ((lngFld11 + .Contact_Phone2.Width) + lngFldSep)
1530            lngFld12 = lngTmp01
1540            If lngFld12 > (lngRptWidth + lngFldSep) Then
1550              blnFld12TooWide = True
1560              strFldTooWide = "Contact_Phone2"
1570              lngRightMost = (lngFld11 - lngFldSep)
1580            Else
1590              lngRightMost = (lngFld12 - lngFldSep)
1600            End If
1610          End If
1620        End Select
1630      Case frm.opgPrintPhone_optCombined.OptionValue
1640        Select Case frm.opgFormatPhone
            Case frm.opgFormatPhone_optFormatted.OptionValue
1650          strFlds = strFlds & "Contact_Phone_Format_Combined;"
1660          lngFld11 = 0&
1670          lngTmp01 = ((lngFld10 + .Contact_Phone_Format_Combined.Width) + lngFldSep)
1680          lngFld12 = lngTmp01
1690          If lngFld12 > (lngRptWidth + lngFldSep) Then
1700            blnFld12TooWide = True
1710            strFldTooWide = "Contact_Phone_Format_Combined"
1720          Else
1730            lngRightMost = (lngFld12 - lngFldSep)
1740          End If
1750        Case frm.opgFormatPhone_optUnformatted.OptionValue
1760          strFlds = strFlds & "Contact_Phone_Combined;"
1770          lngFld11 = 0&
1780          lngTmp01 = ((lngFld10 + .Contact_Phone_Combined.Width) + lngFldSep)
1790          lngFld12 = lngTmp01
1800          If lngFld12 > (lngRptWidth + lngFldSep) Then
1810            blnFld12TooWide = True
1820            strFldTooWide = "Contact_Phone_Combined"
1830          Else
1840            lngRightMost = (lngFld12 - lngFldSep)
1850          End If
1860        End Select
1870      Case frm.opgPrintPhone_opt1Only.OptionValue
1880        Select Case frm.opgFormatPhone
            Case frm.opgFormatPhone_optFormatted.OptionValue
1890          strFlds = strFlds & "Contact_Phone1_Format;"
1900          lngTmp01 = ((lngFld10 + .Contact_Phone1_Format.Width) + lngFldSep)
1910          lngFld11 = 0&
1920          lngFld12 = lngTmp01
1930          If lngFld12 > (lngRptWidth + lngFldSep) Then
1940            blnFld12TooWide = True
1950            strFldTooWide = "Contact_Phone1_Format"
1960          Else
1970            lngRightMost = (lngFld12 - lngFldSep)
1980          End If
1990        Case frm.opgFormatPhone_optUnformatted.OptionValue
2000          strFlds = strFlds & "Contact_Phone1;"
2010          lngTmp01 = ((lngFld10 + .Contact_Phone1.Width) + lngFldSep)
2020          lngFld11 = 0&
2030          lngFld12 = lngTmp01
2040          If lngFld12 > (lngRptWidth + lngFldSep) Then
2050            blnFld12TooWide = True
2060            strFldTooWide = "Contact_Phone1"
2070          Else
2080            lngRightMost = (lngFld12 - lngFldSep)
2090          End If
2100        End Select
2110        lngTmp01 = 0&
2120      Case frm.opgPrintPhone_optNone.OptionValue
2130        lngFld11 = 0&
2140        lngTmp01 = 0&
2150        lngFld12 = lngTmp01
2160      End Select

2170      If blnFld11TooWide = False And blnFld12TooWide = False Then

2180        Select Case frm.opgPrintFax
            Case frm.opgPrintFax_optNone.OptionValue
              ' ** Do nothing, they're already off.
2190        Case frm.opgPrintFax_optPrint.OptionValue
2200          Select Case frm.opgFormatPhone
              Case frm.opgFormatPhone_optFormatted.OptionValue
2210            strFlds = strFlds & "Contact_Fax_Format;"
2220            If ((lngRightMost + lngFldSep) + .Contact_Fax_Format.Width) <= lngRptWidth Then
2230              If lngFld12 > 0& Then
2240                lngTmp01 = ((lngFld12 + .Contact_Fax_Format.Width) + lngFldSep)
2250              Else
2260                lngTmp01 = ((lngFld10 + .Contact_Fax_Format.Width) + lngFldSep)
2270              End If
2280              lngFld13 = lngTmp01
2290              lngRightMost = (lngFld13 - lngFldSep)
2300            Else
2310              blnFld13TooWide = True
2320              strFldTooWide = "Contact_Fax_Format"
2330            End If
2340          Case frm.opgFormatPhone_optUnformatted.OptionValue
2350            strFlds = strFlds & "Contact_Fax;"
2360            If ((lngRightMost + lngFldSep) + .Contact_Fax.Width) <= lngRptWidth Then
2370              If lngFld12 > 0& Then
2380                lngTmp01 = ((lngFld12 + .Contact_Fax.Width) + lngFldSep)
2390              Else
2400                lngTmp01 = ((lngFld10 + .Contact_Fax.Width) + lngFldSep)
2410              End If
2420              lngFld13 = lngTmp01
2430              lngRightMost = (lngFld13 - lngFldSep)
2440            Else
2450              blnFld13TooWide = True
2460              strFldTooWide = "Contact_Fax"
2470            End If
2480          End Select
2490        End Select

2500        If blnFld13TooWide = False Then
2510          Select Case frm.opgPrintEmail
              Case frm.opgPrintEmail_optNone.OptionValue
                ' ** Do nothing, they're already off.
2520          Case frm.opgPrintEmail_optPrint.OptionValue
2530            strFlds = strFlds & "Contact_Email;"
2540            If ((lngRightMost + lngFldSep) + .Contact_Email.Width) <= lngRptWidth Then
2550              If lngFld13 > 0& Then
2560                lngTmp01 = ((lngFld13 + .Contact_Email.Width) + lngFldSep)
2570              ElseIf lngFld12 > 0& Then
2580                lngTmp01 = ((lngFld12 + .Contact_Email.Width) + lngFldSep)
2590              Else
2600                lngTmp01 = ((lngFld10 + .Contact_Email.Width) + lngFldSep)
2610              End If
2620            Else
2630              blnFld14TooWide = True
2640              strFldTooWide = "Contact_Email"
2650            End If
2660          End Select
2670        End If  ' ** blnFld13TooWide.

2680      End If  ' ** blnFld11TooWide, blnFld12TooWide.

2690      If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
            'frm.RightMost.Visible = True
            'frm.RightMost = lngRightMost
            'frm.DidntShow.Visible = True
            'frm.DidntShow = strFldTooWide
2700      End If

          ' **********************************************************
          ' ** Now we move on to actually adjusting the fields.
          ' **********************************************************

2710      Select Case frm.chkShowAcctNum
          Case True
2720        .accountno.Visible = True
2730        .accountno_lbl.Visible = True
2740        .accountno_lbl_line.Visible = True
2750      Case False
            ' ** Already off.
2760      End Select

2770      Select Case frm.chkShowShortName
          Case True
2780        Select Case frm.chkShowAcctNum
            Case True
2790          lngTmp01 = ((.accountno.Left + .accountno.Width) + lngFldSep)
2800          .shortname.Left = lngTmp01
2810          .shortname_lbl.Left = lngTmp01
2820          .shortname_lbl_line.Left = lngTmp01
2830        Case False
2840          .shortname.Left = .accountno.Left
2850          .shortname_lbl.Left = .accountno_lbl.Left
2860          .shortname_lbl_line.Left = .accountno_lbl_line.Left
2870        End Select
2880        .shortname.Visible = True
2890        .shortname_lbl.Visible = True
2900        .shortname_lbl_line.Visible = True
2910        lngTmp01 = ((.shortname.Left + .shortname.Width) + lngFldSep)
2920      Case False
            ' ** Already off.
2930        lngTmp01 = ((.accountno.Left + .accountno.Width) + lngFldSep)
2940      End Select

2950      .Contact_Number.Left = lngTmp01
2960      .Contact_Number_lbl.Left = lngTmp01
2970      .Contact_Number_lbl_line.Left = lngTmp01
2980      lngTmp01 = ((.Contact_Number.Left + .Contact_Number.Width) + lngFldSep)

2990      Select Case frm.opgFormatName
          Case frm.opgFormatName_optAsWritten.OptionValue
3000        If .Contact_Name.Left <> lngTmp01 Then
3010          .Contact_Name.Left = lngTmp01
3020          .Contact_Name_lbl.Left = lngTmp01
3030          .Contact_Name_lbl_line.Left = lngTmp01
3040        End If
3050        .Contact_Name.Visible = True
3060        .Contact_Name_lbl.Visible = True
3070        .Contact_Name_lbl_line.Visible = True
3080        lngTmp01 = ((.Contact_Name.Left + .Contact_Name.Width) + lngFldSep)
3090      Case frm.opgFormatName_optLastFirst.OptionValue
3100        If .Contact_Name_LastFirst.Left <> lngTmp01 Then
3110          .Contact_Name_LastFirst.Left = lngTmp01
3120          .Contact_Name_LastFirst_lbl.Left = lngTmp01
3130          .Contact_Name_LastFirst_lbl_line.Left = lngTmp01
3140        End If
3150        .Contact_Name_LastFirst.Visible = True
3160        .Contact_Name_LastFirst_lbl.Visible = True
3170        .Contact_Name_LastFirst_lbl_line.Visible = True
3180        lngTmp01 = ((.Contact_Name_LastFirst.Left + .Contact_Name_LastFirst.Width) + lngFldSep)
3190      End Select

3200      Select Case frm.opgPrintAddress
          Case frm.opgPrintAddress_optSeparate.OptionValue
3210        If .Contact_Address1.Left <> lngTmp01 Then
3220          .Contact_Address1.Left = lngTmp01
3230          .Contact_Address1_lbl.Left = lngTmp01
3240          .Contact_Address1_lbl_line.Left = lngTmp01
3250        End If
3260        .Contact_Address1.Visible = True
3270        .Contact_Address1_lbl.Visible = True
3280        .Contact_Address1_lbl_line.Visible = True
3290        lngTmp01 = ((.Contact_Address1.Left + .Contact_Address1.Width) + lngFldSep)
3300        If .Contact_Address2.Left <> lngTmp01 Then
3310          .Contact_Address2.Left = lngTmp01
3320          .Contact_Address2_lbl.Left = lngTmp01
3330          .Contact_Address2_lbl_line.Left = lngTmp01
3340        End If
3350        .Contact_Address2.Visible = True
3360        .Contact_Address2_lbl.Visible = True
3370        .Contact_Address2_lbl_line.Visible = True
3380        lngTmp01 = ((.Contact_Address2.Left + .Contact_Address2.Width) + lngFldSep)
3390      Case frm.opgPrintAddress_optCombined.OptionValue
3400        If .Contact_Address_Combined.Left <> lngTmp01 Then
3410          .Contact_Address_Combined.Left = lngTmp01
3420          .Contact_Address_Combined_lbl.Left = lngTmp01
3430          .Contact_Address_Combined_lbl_line.Left = lngTmp01
3440        End If
3450        .Contact_Address_Combined.Visible = True
3460        .Contact_Address_Combined_lbl.Visible = True
3470        .Contact_Address_Combined_lbl_line.Visible = True
3480        lngTmp01 = ((.Contact_Address_Combined.Left + .Contact_Address_Combined.Width) + lngFldSep)
3490      End Select

          'COUNTRY NOT DONE!
3500      Select Case intOptCSZ
          Case OPT_SEP
3510        If .Contact_City.Left <> lngTmp01 Then
3520          .Contact_City.Left = lngTmp01
3530          .Contact_City_lbl.Left = lngTmp01
3540          .Contact_City_lbl_line.Left = lngTmp01
3550        End If
3560        .Contact_City.Visible = True
3570        .Contact_City_lbl.Visible = True
3580        .Contact_City_lbl_line.Visible = True
3590        lngTmp01 = ((.Contact_City.Left + .Contact_City.Width) + lngFldSep)
3600        If .Contact_State.Left <> lngTmp01 Then
3610          .Contact_State.Left = lngTmp01
3620          .Contact_State_lbl.Left = lngTmp01
3630          .Contact_State_lbl_line.Left = lngTmp01
3640        End If
3650        .Contact_State.Visible = True
3660        .Contact_State_lbl.Visible = True
3670        .Contact_State_lbl_line.Visible = True
3680        lngTmp01 = ((.Contact_State.Left + .Contact_State.Width) + lngFldSep)
3690        Select Case frm.opgFormatZip
            Case frm.opgFormatZip_optFormatted.OptionValue
3700          If .Contact_Zip_Format.Left <> lngTmp01 Then
3710            .Contact_Zip_Format.Left = lngTmp01
3720            .Contact_Zip_Format_lbl.Left = lngTmp01
3730            .Contact_Zip_Format_lbl_line.Left = lngTmp01
3740          End If
3750          .Contact_Zip_Format.Visible = True
3760          .Contact_Zip_Format_lbl.Visible = True
3770          .Contact_Zip_Format_lbl_line.Visible = True
3780          lngTmp01 = ((.Contact_Zip_Format.Left + .Contact_Zip_Format.Width) + lngFldSep)
3790        Case frm.opgFormatZip_optUnformatted.OptionValue
3800          If .Contact_Zip.Left <> lngTmp01 Then
3810            .Contact_Zip.Left = lngTmp01
3820            .Contact_Zip_lbl.Left = lngTmp01
3830            .Contact_Zip_lbl_line.Left = lngTmp01
3840          End If
3850          .Contact_Zip.Visible = True
3860          .Contact_Zip_lbl.Visible = True
3870          .Contact_Zip_lbl_line.Visible = True
3880          lngTmp01 = ((.Contact_Zip.Left + .Contact_Zip.Width) + lngFldSep)
3890        End Select
3900      Case OPT_COM
3910        Select Case frm.opgFormatZip
            Case frm.opgFormatZip_optFormatted.OptionValue
3920          If .Contact_CSZ_Format.Left <> lngTmp01 Then
3930            .Contact_CSZ_Format.Left = lngTmp01
3940            .Contact_CSZ_Format_lbl.Left = lngTmp01
3950            .Contact_CSZ_Format_lbl_line.Left = lngTmp01
3960          End If
3970          .Contact_CSZ_Format.Visible = True
3980          .Contact_CSZ_Format_lbl.Visible = True
3990          .Contact_CSZ_Format_lbl_line.Visible = True
4000          lngTmp01 = ((.Contact_CSZ_Format.Left + .Contact_CSZ_Format.Width) + lngFldSep)
4010        Case frm.opgFormatZip_optUnformatted.OptionValue
4020          If .Contact_CSZ.Left <> lngTmp01 Then
4030            .Contact_CSZ.Left = lngTmp01
4040            .Contact_CSZ_lbl.Left = lngTmp01
4050            .Contact_CSZ_lbl_line.Left = lngTmp01
4060          End If
4070          .Contact_CSZ.Visible = True
4080          .Contact_CSZ_lbl.Visible = True
4090          .Contact_CSZ_lbl_line.Visible = True
4100          lngTmp01 = ((.Contact_CSZ.Left + .Contact_CSZ.Width) + lngFldSep)
4110        End Select
4120      End Select

4130      Select Case frm.opgPrintPhone
          Case frm.opgPrintPhone_optSeparate.OptionValue
4140        Select Case frm.opgFormatPhone
            Case frm.opgFormatPhone_optFormatted.OptionValue
4150          If blnFld11TooWide = True Then
                ' ** Already off.
4160          Else
4170            If .Contact_Phone1_Format.Left <> lngTmp01 Then
4180              .Contact_Phone1_Format.Left = lngTmp01
4190              .Contact_Phone1_Format_lbl.Left = lngTmp01
4200              .Contact_Phone1_Format_lbl_line.Left = lngTmp01
4210            End If
4220            .Contact_Phone1_Format.Visible = True
4230            .Contact_Phone1_Format_lbl.Visible = True
4240            .Contact_Phone1_Format_lbl_line.Visible = True
4250            lngTmp01 = ((.Contact_Phone1_Format.Left + .Contact_Phone1_Format.Width) + lngFldSep)
4260            If blnFld12TooWide = True Then
                  ' ** Already off.
4270            Else
4280              If .Contact_Phone2_Format.Left <> lngTmp01 Then
4290                .Contact_Phone2_Format.Left = lngTmp01
4300                .Contact_Phone2_Format_lbl.Left = lngTmp01
4310                .Contact_Phone2_Format_lbl_line.Left = lngTmp01
4320              End If
4330              .Contact_Phone2_Format.Visible = True
4340              .Contact_Phone2_Format_lbl.Visible = True
4350              .Contact_Phone2_Format_lbl_line.Visible = True
4360              lngTmp01 = ((.Contact_Phone2_Format.Left + .Contact_Phone2_Format.Width) + lngFldSep)
4370            End If
4380          End If
4390        Case frm.opgFormatPhone_optUnformatted.OptionValue
4400          If blnFld11TooWide = True Then
                ' ** Already off.
4410          Else
4420            If .Contact_Phone1.Left <> lngTmp01 Then
4430              .Contact_Phone1.Left = lngTmp01
4440              .Contact_Phone1_lbl.Left = lngTmp01
4450              .Contact_Phone1_lbl_line.Left = lngTmp01
4460            End If
4470            .Contact_Phone1.Visible = True
4480            .Contact_Phone1_lbl.Visible = True
4490            .Contact_Phone1_lbl_line.Visible = True
4500            lngTmp01 = ((.Contact_Phone1.Left + .Contact_Phone1.Width) + lngFldSep)
4510            If blnFld12TooWide = True Then
                  ' ** Already off.
4520            Else
4530              If .Contact_Phone2.Left <> lngTmp01 Then
4540                .Contact_Phone2.Left = lngTmp01
4550                .Contact_Phone2_lbl.Left = lngTmp01
4560                .Contact_Phone2_lbl_line.Left = lngTmp01
4570              End If
4580              .Contact_Phone2.Visible = True
4590              .Contact_Phone2_lbl.Visible = True
4600              .Contact_Phone2_lbl_line.Visible = True
4610              lngTmp01 = ((.Contact_Phone2.Left + .Contact_Phone2.Width) + lngFldSep)
4620            End If
4630          End If
4640        End Select
4650      Case frm.opgPrintPhone_optCombined.OptionValue
4660        Select Case frm.opgFormatPhone
            Case frm.opgFormatPhone_optFormatted.OptionValue
4670          If blnFld11TooWide = True Or blnFld12TooWide = True Then
                ' ** Already off.
4680          Else
4690            If .Contact_Phone_Format_Combined.Left <> lngTmp01 Then
4700              .Contact_Phone_Format_Combined.Left = lngTmp01
4710              .Contact_Phone_Format_Combined_lbl.Left = lngTmp01
4720              .Contact_Phone_Format_Combined_lbl_line.Left = lngTmp01
4730            End If
4740            .Contact_Phone_Format_Combined.Visible = True
4750            .Contact_Phone_Format_Combined_lbl.Visible = True
4760            .Contact_Phone_Format_Combined_lbl_line.Visible = True
4770            lngTmp01 = ((.Contact_Phone_Format_Combined.Left + .Contact_Phone_Format_Combined.Width) + lngFldSep)
4780          End If
4790        Case frm.opgFormatPhone_optUnformatted.OptionValue
4800          If blnFld11TooWide = True Or blnFld12TooWide = True Then
                ' ** Already off.
4810          Else
4820            If .Contact_Phone_Combined.Left <> lngTmp01 Then
4830              .Contact_Phone_Combined.Left = lngTmp01
4840              .Contact_Phone_Combined_lbl.Left = lngTmp01
4850              .Contact_Phone_Combined_lbl_line.Left = lngTmp01
4860            End If
4870            .Contact_Phone_Combined.Visible = True
4880            .Contact_Phone_Combined_lbl.Visible = True
4890            .Contact_Phone_Combined_lbl_line.Visible = True
4900            lngTmp01 = ((.Contact_Phone_Combined.Left + .Contact_Phone_Combined.Width) + lngFldSep)
4910          End If
4920        End Select
4930      Case frm.opgPrintPhone_opt1Only.OptionValue
4940        Select Case frm.opgFormatPhone
            Case frm.opgFormatPhone_optFormatted.OptionValue
4950          If blnFld11TooWide = True Or blnFld12TooWide = True Then
                ' ** Already off.
4960          Else
4970            If .Contact_Phone1_Format.Left <> lngTmp01 Then
4980              .Contact_Phone1_Format.Left = lngTmp01
4990              .Contact_Phone1_Format_lbl.Left = lngTmp01
5000              .Contact_Phone1_Format_lbl_line.Left = lngTmp01
5010            End If
5020            .Contact_Phone1_Format.Visible = True
5030            .Contact_Phone1_Format_lbl.Visible = True
5040            .Contact_Phone1_Format_lbl_line.Visible = True
5050            lngTmp01 = ((.Contact_Phone1_Format.Left + .Contact_Phone1_Format.Width) + lngFldSep)
5060          End If
5070        Case frm.opgFormatPhone_optUnformatted.OptionValue
5080          If blnFld11TooWide = True Or blnFld12TooWide = True Then
                ' ** Already off.
5090          Else
5100            If .Contact_Phone1.Left <> lngTmp01 Then
5110              .Contact_Phone1.Left = lngTmp01
5120              .Contact_Phone1_lbl.Left = lngTmp01
5130              .Contact_Phone1_lbl_line.Left = lngTmp01
5140            End If
5150            .Contact_Phone1.Visible = True
5160            .Contact_Phone1_lbl.Visible = True
5170            .Contact_Phone1_lbl_line.Visible = True
5180            lngTmp01 = ((.Contact_Phone1.Left + .Contact_Phone1.Width) + lngFldSep)
5190          End If
5200        End Select
5210      Case frm.opgPrintPhone_optNone.OptionValue
            ' ** Already off.
5220      End Select

5230      If blnFld11TooWide = False And blnFld12TooWide = False And blnFld13TooWide = False Then

5240        Select Case frm.opgPrintFax
            Case frm.opgPrintFax_optNone.OptionValue
              ' ** Already off.
5250        Case frm.opgPrintFax_optPrint.OptionValue
5260          Select Case frm.opgFormatPhone
              Case frm.opgFormatPhone_optFormatted.OptionValue
5270            If .Contact_Fax_Format.Left <> lngTmp01 Then
5280              .Contact_Fax_Format.Left = lngTmp01
5290              .Contact_Fax_Format_lbl.Left = lngTmp01
5300              .Contact_Fax_Format_lbl_line.Left = lngTmp01
5310            End If
5320            .Contact_Fax_Format.Visible = True
5330            .Contact_Fax_Format_lbl.Visible = True
5340            .Contact_Fax_Format_lbl_line.Visible = True
5350            lngTmp01 = ((.Contact_Fax_Format.Left + .Contact_Fax_Format.Width) + lngFldSep)
5360          Case frm.opgFormatPhone_optUnformatted.OptionValue
5370            If .Contact_Fax.Left <> lngTmp01 Then
5380              .Contact_Fax.Left = lngTmp01
5390              .Contact_Fax_lbl.Left = lngTmp01
5400              .Contact_Fax_lbl_line.Left = lngTmp01
5410            End If
5420            .Contact_Fax.Visible = True
5430            .Contact_Fax_lbl.Visible = True
5440            .Contact_Fax_lbl_line.Visible = True
5450            lngTmp01 = ((.Contact_Fax.Left + .Contact_Fax.Width) + lngFldSep)
5460          End Select
5470        End Select

5480        If blnFld14TooWide = False Then
5490          Select Case frm.opgPrintEmail
              Case frm.opgPrintEmail_optNone.OptionValue
                ' ** Already off.
5500          Case frm.opgPrintEmail_optPrint.OptionValue
5510            If .Contact_Email.Left <> lngTmp01 Then
5520              .Contact_Email.Left = lngTmp01
5530              .Contact_Email_cmd.Left = (.Contact_Email.Left - (3& * lngTpp))
5540              .Contact_Email_lbl.Left = lngTmp01
5550              .Contact_Email_lbl2.Left = ((.Contact_Email_lbl.Left + .Contact_Email_lbl.Width) - .Contact_Email_lbl2.Width)
5560              .Contact_Email_lbl_line.Left = lngTmp01
5570            End If
5580            .Contact_Email.Visible = True
5590            .Contact_Email_lbl.Visible = True
5600            .Contact_Email_lbl2.Visible = True
5610            .Contact_Email_lbl_line.Visible = True
5620            lngTmp01 = ((.Contact_Email.Left + .Contact_Email.Width) + lngFldSep)
5630          End Select
5640        End If  ' ** blnFld13TooWide.

5650      End If  ' ** blnFld11TooWide, blnFld12TooWide.

5660    End With

EXITP:
5670    Set rpt = Nothing
5680    Set frm = Nothing
5690    SetupRptFlds = blnRetVal
5700    Exit Function

ERRH:
5710    blnRetVal = False
5720    Select Case ERR.Number
        Case Else
5730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5740    End Select
5750    Resume EXITP

End Function

Private Function RptFld_Load() As Boolean

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "RptFld_Load"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

5810    blnRetVal = True

5820    lngRptFlds = 0&
5830    ReDim arr_varRptFld(R_ELEMS, 0)

5840    Set dbs = CurrentDb
5850    With dbs
5860      Set rst = .OpenRecordset("tblAcctCon_Field", dbOpenDynaset, dbReadOnly)
5870      With rst
5880        If .BOF = True And .EOF = True Then
              ' ** Something's seriously wrong!
5890          blnRetVal = False
5900        Else
5910          .MoveLast
5920          lngRecs = .RecordCount
5930          .MoveFirst
5940          For lngX = 1& To lngRecs
5950            lngRptFlds = lngRptFlds + 1&
5960            lngE = lngRptFlds - 1&
5970            ReDim Preserve arr_varRptFld(R_ELEMS, lngE)
                ' *************************************************
                ' ** Array: arr_varRptFld()
                ' **
                ' **   Field  Element  Name            Constant
                ' **   =====  =======  ==============  ==========
                ' **     1       0     confld_id       R_FID
                ' **     2       1     confld_name     R_FNAM
                ' **     3       2     Visible         R_VIS
                ' **
                ' *************************************************
5980            arr_varRptFld(R_FID, lngE) = ![confld_id]
5990            arr_varRptFld(R_FNAM, lngE) = ![confld_name]
6000            arr_varRptFld(R_VIS, lngE) = CBool(False)
6010            If lngX < lngRecs Then .MoveNext
6020          Next
6030        End If
6040        .Close
6050      End With
6060      .Close
6070    End With

        'accountno
        'shortname
        'Contact_Number
        'Contact_Name
        'Contact_Name_LastFirst
        'Contact_Address1
        'Contact_Address2
        'Contact_Address_Combined
        'Contact_City
        'Contact_State
        'Contact_Zip
        'Contact_Zip_Format
        'Contact_CSZ
        'Contact_CSZ_Format
        'Contact_Phone1
        'Contact_Phone1_Format
        'Contact_Phone2
        'Contact_Phone2_Format
        'Contact_Phone_Combined
        'Contact_Phone_Format_Combined
        'Contact_Fax
        'Contact_Fax_Format
        'Contact_Email

EXITP:
6080    Set rst = Nothing
6090    Set dbs = Nothing
6100    RptFld_Load = blnRetVal
6110    Exit Function

ERRH:
6120    blnRetVal = False
6130    Select Case ERR.Number
        Case Else
6140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6150    End Select
6160    Resume EXITP

End Function

Public Function SetupExcelFlds(strCallingForm As String) As Long

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "SetupExcelFlds"

        Dim frm As Access.Form
        Dim lngConTypID As Long, lngConRptID As Long
        Dim lngConGrpID01 As Long, lngConGrpID02 As Long, lngConGrpID03 As Long, lngConGrpID04 As Long, lngConGrpID05 As Long
        Dim lngConGrpID06 As Long, lngConGrpID07 As Long, lngConGrpID08 As Long, lngConGrpID09 As Long
        Dim lngConGrpTypID01 As Long, lngConGrpTypID02 As Long, lngConGrpTypID03 As Long, lngConGrpTypID04 As Long, lngConGrpTypID05 As Long
        Dim lngConGrpTypID06 As Long, lngConGrpTypID07 As Long, lngConGrpTypID08 As Long, lngConGrpTypID09 As Long
        Dim lngConFmtID01 As Long, lngConFmtID02 As Long, lngConFmtID03 As Long, lngConFmtID04 As Long, lngConFmtID05 As Long
        Dim lngConFmtID06 As Long, lngConFmtID07 As Long, lngConFmtID08 As Long, lngConFmtID09 As Long
        Dim lngConOptID01 As Long, lngConOptID02 As Long, lngConOptID03 As Long
        Dim intOptCSZ As Integer
        Dim blnNoAddress As Boolean, blnNoCityStateZip As Boolean, blnNoPhone As Boolean, blnNoFax As Boolean, blnNoEmail As Boolean
        Dim strKey As String
        Dim varTmp00 As Variant, blnTmp01 As Boolean
        Dim lngRetVal As Long

6210    lngRetVal = 0&

6220    Set frm = Forms(strCallingForm)
6230    With frm

6240      blnTmp01 = False
6250      If frm.chkEnableCountry_Wide.Visible = True Then
6260        blnTmp01 = frm.chkEnableCountry_Wide
6270      ElseIf frm.chkEnableCountry_Compact.Visible = True Then
6280        blnTmp01 = frm.chkEnableCountry_Compact
6290      End If

6300      Select Case blnTmp01
          Case True
6310        intOptCSZ = .opgPrintCSZCP
6320      Case False
6330        intOptCSZ = .opgPrintCSZ
6340      End Select

          ' ** Groups 1 & 2: AccountNum, ShortName.
6350      lngConTypID = 0&
6360      lngConGrpID01 = 0&: lngConGrpID02 = 0&
6370      lngConGrpTypID01 = 0&: lngConGrpTypID02 = 0&
6380      lngConFmtID01 = 0&: lngConFmtID02 = 0&
6390      If .chkShowAcctNum = True And .chkShowShortName = False Then
6400        lngConTypID = 1&       ' ** AccountNumOnly.
6410        lngConGrpID01 = 1&     ' ** AccountNum.
6420        lngConGrpTypID01 = 1&  ' ** AccountNum.
6430        lngConFmtID01 = 2&     ' ** Unformatted.
6440      ElseIf .chkShowAcctNum = False And .chkShowShortName = True Then
6450        lngConTypID = 2&       ' ** ShortNameOnly.
6460        lngConGrpID02 = 2&     ' ** ShortName.
6470        lngConGrpTypID02 = 2&  ' ** ShortName.
6480        lngConFmtID02 = 2&     ' ** Unformatted.
6490      ElseIf .chkShowAcctNum = True And .chkShowShortName = True Then
6500        lngConTypID = 3&       ' ** AccountNumShortName.
6510        lngConGrpID01 = 1&     ' ** AccountNum.
6520        lngConGrpID02 = 2&     ' ** ShortName.
6530        lngConGrpTypID01 = 1&  ' ** AccountNum.
6540        lngConGrpTypID02 = 2&  ' ** ShortName.
6550        lngConFmtID01 = 2&     ' ** Unformatted.
6560        lngConFmtID02 = 2&     ' ** Unformatted.
6570      End If

6580      lngConOptID01 = 0&
6590      If .chkShowAcctNum = True And .chkShowShortName = False Then
6600        lngConOptID01 = 1&  ' ** AccountNumOnly.
6610      ElseIf .chkShowAcctNum = False And .chkShowShortName = True Then
6620        lngConOptID01 = 2&  ' ** ShortNameOnly.
6630      ElseIf .chkShowAcctNum = True And .chkShowShortName = True Then
6640        lngConOptID01 = 3&  ' ** AccountNumShortName.
6650      End If

          ' ** Group3: ContactNum.
6660      lngConGrpID03 = 3&     ' ** ContactNum.
6670      lngConGrpTypID03 = 3&  ' ** ContactNum.
6680      lngConFmtID03 = 2&     ' ** Unformatted.

          ' ** Group 4: Name.
6690      lngConGrpID04 = 0&: lngConGrpTypID04 = 0&: lngConFmtID04 = 0&
6700      Select Case .opgFormatName
          Case .opgFormatName_optAsWritten.OptionValue
6710        lngConGrpID04 = 4&     ' ** Name.
6720        lngConGrpTypID04 = 5&  ' ** NameCombined.
6730        lngConFmtID04 = 2&     ' ** Unformatted.
6740      Case .opgFormatName_optLastFirst.OptionValue
6750        lngConGrpID04 = 4&     ' ** Name.
6760        lngConGrpTypID04 = 5&  ' ** NameCombined.
6770        lngConFmtID04 = 1&     ' ** Formatted.
6780      End Select

          ' ** Group 5: Address.
6790      lngConGrpID05 = 0&: lngConGrpTypID05 = 0&: lngConFmtID05 = 0&
6800      blnNoAddress = False: blnNoCityStateZip = False
6810      Select Case .opgPrintAddress
          Case .opgPrintAddress_optSeparate.OptionValue
6820        lngConGrpID05 = 5&     ' ** Address.
6830        lngConGrpTypID05 = 6&  ' ** AddressSeparate.
6840        lngConFmtID05 = 2&     ' ** Unformatted.
6850      Case .opgPrintAddress_optCombined.OptionValue
6860        lngConGrpID05 = 5&     ' ** Address.
6870        lngConGrpTypID05 = 7&  ' ** AddressCombined.
6880        lngConFmtID05 = 2&     ' ** Unformatted.
6890      Case .opgPrintAddress_optNone.OptionValue
6900        blnNoAddress = True
6910      End Select

          'COUNTRY NOT DONE YET!
          ' ** Group 6: CityStateZip.
6920      lngConGrpID06 = 0&: lngConGrpTypID06 = 0&: lngConFmtID06 = 0&
6930      Select Case intOptCSZ
          Case OPT_SEP
6940        lngConGrpID06 = 6&  ' ** CityStateZip.
6950        lngConGrpTypID06 = 8&  ' ** CityStateZipSeparate.
6960        Select Case .opgFormatZip
            Case .opgFormatZip_optFormatted.OptionValue
6970          lngConFmtID06 = 1&   ' ** Formatted.
6980        Case .opgFormatZip_optUnformatted.OptionValue
6990          lngConFmtID06 = 2&   ' ** Unformatted.
7000        End Select
7010      Case OPT_COM
7020        lngConGrpID06 = 6&  ' ** CityStateZip.
7030        lngConGrpTypID06 = 9&  ' ** CityStateZipCombined.
7040        Select Case .opgFormatZip
            Case .opgFormatZip_optFormatted.OptionValue
7050          lngConFmtID06 = 1&   ' ** Formatted.
7060        Case .opgFormatZip_optUnformatted.OptionValue
7070          lngConFmtID06 = 2&   ' ** Unformatted.
7080        End Select
7090      Case OPT_NON
7100        blnNoCityStateZip = True
7110      End Select

7120      lngConOptID02 = 0&
7130      If blnNoAddress = True And blnNoCityStateZip = True Then
7140        lngConOptID02 = 4&  ' ** None.
7150      Else
7160        If blnNoCityStateZip = True Then
7170          lngConOptID02 = 5&  ' ** AddressOnly.
7180        ElseIf blnNoAddress = True Then
7190          lngConOptID02 = 6&  ' ** CityStateZipOnly.
7200        Else
7210          lngConOptID02 = 7&  ' ** AddressCityStateZip.
7220        End If
7230      End If

          ' ** Group 7: Phone.
7240      lngConGrpID07 = 0&: lngConGrpTypID07 = 0&: lngConFmtID07 = 0&
7250      blnNoPhone = False: blnNoFax = False: blnNoEmail = False
7260      Select Case .opgPrintPhone
          Case .opgPrintPhone_optSeparate.OptionValue
7270        lngConGrpID07 = 7&   ' ** Phone.
7280        lngConGrpTypID07 = 10&  ' ** PhoneSeparate.
7290        Select Case .opgFormatPhone
            Case .opgFormatPhone_optFormatted.OptionValue
7300          lngConFmtID07 = 1&   ' ** Formatted.
7310        Case .opgFormatPhone_optUnformatted.OptionValue
7320          lngConFmtID07 = 2&   ' ** Unformatted.
7330        End Select
7340      Case .opgPrintPhone_optCombined.OptionValue
7350        lngConGrpID07 = 7&   ' ** Phone.
7360        lngConGrpTypID07 = 11&  ' ** PhoneCombined.
7370        Select Case .opgFormatPhone
            Case .opgFormatPhone_optFormatted.OptionValue
7380          lngConFmtID07 = 1&   ' ** Formatted.
7390        Case .opgFormatPhone_optUnformatted.OptionValue
7400          lngConFmtID07 = 2&   ' ** Unformatted.
7410        End Select
7420      Case .opgPrintPhone_opt1Only.OptionValue
7430        lngConGrpID07 = 7&   ' ** Phone.
7440        lngConGrpTypID07 = 12&  ' ** PhoneOneOnly.
7450        Select Case .opgFormatPhone
            Case .opgFormatPhone_optFormatted.OptionValue
7460          lngConFmtID07 = 1&   ' ** Formatted.
7470        Case .opgFormatPhone_optUnformatted.OptionValue
7480          lngConFmtID07 = 2&   ' ** Unformatted.
7490        End Select
7500      Case .opgPrintPhone_optNone.OptionValue
7510        blnNoPhone = True
7520      End Select

          ' ** Group 8: Fax.
7530      lngConGrpID08 = 0&: lngConGrpTypID08 = 0&: lngConFmtID08 = 0&
7540      Select Case .opgPrintFax
          Case .opgPrintFax_optPrint.OptionValue
7550        lngConGrpID08 = 8&      ' ** Fax.
7560        lngConGrpTypID08 = 13&  ' ** Fax.
7570        Select Case .opgFormatPhone
            Case .opgFormatPhone_optFormatted.OptionValue
7580          lngConFmtID08 = 1&   ' ** Formatted.
7590        Case .opgFormatPhone_optUnformatted.OptionValue
7600          lngConFmtID08 = 2&   ' ** Unformatted.
7610        End Select
7620      Case .opgPrintFax_optNone.OptionValue
7630        blnNoFax = True
7640      End Select

          ' ** Group 9: Email.
7650      lngConGrpID09 = 0&: lngConGrpTypID09 = 0&: lngConFmtID09 = 0&
7660      Select Case .opgPrintEmail
          Case .opgPrintEmail_optPrint.OptionValue
7670        lngConGrpID09 = 9&     ' ** Email.
7680        lngConGrpTypID09 = 14  ' ** Email.
7690        lngConFmtID09 = 2&     ' ** Unformatted.
7700      Case .opgPrintEmail_optNone.OptionValue
7710        blnNoEmail = True
7720      End Select

7730      lngConOptID03 = 0&
7740      If blnNoPhone = True And blnNoFax = True And blnNoEmail = True Then
7750        lngConOptID03 = 8&        ' ** None.
7760      Else
7770        If blnNoPhone = False And blnNoFax = True And blnNoEmail = True Then
7780          lngConOptID03 = 9&      ' ** PhoneOnly.
7790        ElseIf blnNoPhone = True And blnNoFax = False And blnNoEmail = True Then
7800          lngConOptID03 = 10&      ' ** FaxOnly.
7810        ElseIf blnNoPhone = True And blnNoFax = True And blnNoEmail = False Then
7820          lngConOptID03 = 11       ' ** EmailOnly.
7830        Else
7840          If blnNoPhone = False And blnNoFax = False And blnNoEmail = True Then
7850            lngConOptID03 = 12&    ' ** PhoneFax.
7860          ElseIf blnNoPhone = False And blnNoFax = True And blnNoEmail = False Then
7870            lngConOptID03 = 13&    ' ** PhoneEmail.
7880          ElseIf blnNoPhone = True And blnNoFax = False And blnNoEmail = False Then
7890            lngConOptID03 = 14&    ' ** FaxEmail.
7900          Else
7910            If blnNoPhone = False And blnNoFax = False And blnNoEmail = False Then
7920              lngConOptID03 = 15&  ' ** PhoneFaxEmail.
7930            End If
7940          End If
7950        End If
7960      End If

7970      strKey = Right("00" & CStr(lngConTypID), 2) & "_"
7980      strKey = strKey & Right("00" & CStr(lngConGrpID01), 2) & Right("00" & CStr(lngConGrpTypID01), 2) & Right("00" & CStr(lngConFmtID01), 2)
7990      strKey = strKey & "_"
8000      strKey = strKey & Right("00" & CStr(lngConGrpID02), 2) & Right("00" & CStr(lngConGrpTypID02), 2) & Right("00" & CStr(lngConFmtID02), 2)
8010      strKey = strKey & "_"
8020      strKey = strKey & Right("00" & CStr(lngConGrpID03), 2) & Right("00" & CStr(lngConGrpTypID03), 2) & Right("00" & CStr(lngConFmtID03), 2)
8030      strKey = strKey & "_"
8040      strKey = strKey & Right("00" & CStr(lngConGrpID04), 2) & Right("00" & CStr(lngConGrpTypID04), 2) & Right("00" & CStr(lngConFmtID04), 2)
8050      strKey = strKey & "_"
8060      strKey = strKey & Right("00" & CStr(lngConGrpID05), 2) & Right("00" & CStr(lngConGrpTypID05), 2) & Right("00" & CStr(lngConFmtID05), 2)
8070      strKey = strKey & "_"
8080      strKey = strKey & Right("00" & CStr(lngConGrpID06), 2) & Right("00" & CStr(lngConGrpTypID06), 2) & Right("00" & CStr(lngConFmtID06), 2)
8090      strKey = strKey & "_"
8100      strKey = strKey & Right("00" & CStr(lngConGrpID07), 2) & Right("00" & CStr(lngConGrpTypID07), 2) & Right("00" & CStr(lngConFmtID07), 2)
8110      strKey = strKey & "_"
8120      strKey = strKey & Right("00" & CStr(lngConGrpID08), 2) & Right("00" & CStr(lngConGrpTypID08), 2) & Right("00" & CStr(lngConFmtID08), 2)
8130      strKey = strKey & "_"
8140      strKey = strKey & Right("00" & CStr(lngConGrpID09), 2) & Right("00" & CStr(lngConGrpTypID09), 2) & Right("00" & CStr(lngConFmtID09), 2)

8150      varTmp00 = DLookup("[conrpt_id]", "tblAcctCon_Report_Version", "[conrptver_key] = '" & strKey & "'")
8160      If IsNull(varTmp00) = False Then
8170        lngConRptID = varTmp00
8180      End If

8190      lngRetVal = lngConRptID

8200    End With

EXITP:
8210    Set frm = Nothing
8220    SetupExcelFlds = lngRetVal
8230    Exit Function

ERRH:
8240    lngRetVal = 0&
8250    Select Case ERR.Number
        Case Else
8260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8270    End Select
8280    Resume EXITP

End Function

Public Sub Sub_View(intMode As Integer, lngForm_Width As Long, lngFrm_Left_Orig As Long, lngFrm_Width_Orig As Long, lngFrm_Height_Orig As Long, lngWinSub_Diff As Long, frm As Access.Form)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "Sub_View"

        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngDiff_Height As Long, lngDiff_Width As Long, lngCnt As Long
        Dim varTmp00 As Variant, lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, blnTmp04 As Boolean
        Dim lngX As Long

8310    With frm

8320      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
8330        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
8340      End If

          ' ** Variables are fed empty, then populated ByRef.
8350      GetFormDimensions frm, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Function: modWindowFunctions.

8360      lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
8370      lngMonitorNum = 1&: lngTmp03 = 0&
8380      EnumMonitors frm  ' ** Module Function: modMonitorFuncs.
8390      If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

8400      .FocusHolder.SetFocus

          ' ** Original values are without shortname.

8410      Select Case intMode
          Case 1  ' ** Wide.

8420        .Detail.Height = arr_varCtl(C_HGT, 1)

8430        If lngMonitorNum = 1& Then lngTmp03 = lngTop
8440        DoCmd.SelectObject acForm, frm.Name, False
8450        DoEvents
8460        DoCmd.MoveSize lngLeft, lngTmp03, lngWidth, lngFrm_Height_Orig  'lngTop
8470        If lngMonitorNum > 1& Then
8480          LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
8490        End If

8500        .frmAccountContacts_Sub.Visible = True
8510        .frmAccountContacts_Sub_box.Visible = True
8520        .frmAccountContacts_Sub2.Visible = False
8530        .frmAccountContacts_Sub2_box.Visible = False

8540        .Nav_Sub1_box01.Visible = True
8550        .Nav_Sub1_hline01.Visible = True
8560        .Nav_Sub1_hline02.Visible = True
8570        .Nav_Sub1_vline01.Visible = True
8580        .Nav_Sub1_vline02.Visible = True
8590        .Nav_Sub1_vline03.Visible = True
8600        .Nav_Sub1_vline04.Visible = True
8610        .Nav_Sub2_box01.Visible = False
8620        .Nav_Sub2_hline01.Visible = False
8630        .Nav_Sub2_hline02.Visible = False
8640        .Nav_Sub2_vline01.Visible = False
8650        .Nav_Sub2_vline02.Visible = False
8660        .Nav_Sub2_vline03.Visible = False
8670        .Nav_Sub2_vline04.Visible = False

            ' ** tblAccount_Contact_Staging, linked to Account.
8680        varTmp00 = DCount("*", "qryAccount_Contacts_01_01")
8690        If IsNull(varTmp00) = True Then
8700          lngRecsCur = 0&
8710        Else
8720          lngRecsCur = varTmp00
8730        End If

8740        Select Case .chkShowAcctNum_Last
            Case True
8750          .chkShowAcctNum = True
8760          .chkShowAcctNum_lbl.FontBold = True
8770        Case False
8780          .chkShowAcctNum = False
8790          .chkShowAcctNum_lbl.FontBold = False
8800        End Select
8810        If lngRecsCur > 0& Then
8820          .chkShowAcctNum.Enabled = True
8830        End If
8840        Select Case .chkShowShortName_Last
            Case True
8850          .chkShowShortName = True
8860          .chkShowShortName_lbl.FontBold = True
8870        Case False
8880          .chkShowShortName = False
8890          .chkShowShortName_lbl.FontBold = False
8900        End Select
8910        If lngRecsCur > 0& Then
8920          .chkShowShortName.Enabled = True
8930        End If

8940        lngDiff_Height = (.frmAccountContacts_Sub.Height - .frmAccountContacts_Sub2.Height)

8950        .ShortcutMenu_lbl.Top = arr_varCtl(C_TOP, 11)
8960        .ShortcutMenu_up_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
8970        .ShortcutMenu_down_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
8980        lngTmp01 = (.ShortcutMenu_lbl.Top + .ShortcutMenu_lbl.Height)

8990        If .Detail_hline01.Top <> lngTmp01 Then
9000          .Detail_hline01.Top = lngTmp01
9010          .Detail_hline02.Top = (.Detail_hline01.Top + lngTpp)
9020          .Detail_vline01.Top = .Detail_hline01.Top
9030          .Detail_vline02.Top = .Detail_hline01.Top
9040        End If

9050        If .opgView_box2.Top <> arr_varCtl(C_TOP, 12) Then
9060          lngTmp01 = (arr_varCtl(C_TOP, 12) - .opgView_box2.Top)
9070          .opgView_box.Top = (.opgView_box.Top + lngTmp01)
9080          .opgView_box2.Top = (.opgView_box2.Top + lngTmp01)
9090          .opgView_lbl.Top = (.opgView_lbl.Top + lngTmp01)
9100          lngCnt = (lngTmp01 / lngTpp)
9110          For lngX = 1& To lngCnt
9120            .opgView_optWide.Top = (.opgView_optWide.Top + lngTpp)
9130            .opgView_optWide_lbl.Top = (.opgView_optWide_lbl.Top + lngTpp)
9140            .opgView_optCompact.Top = (.opgView_optCompact.Top + lngTpp)
9150            .opgView_optCompact_lbl.Top = (.opgView_optCompact_lbl.Top + lngTpp)
9160            .opgView.Top = (.opgView.Top + lngTpp)
9170          Next  ' ** lngX.
9180        End If

9190        If .opgFormat_box.Top <> arr_varCtl(C_TOP, 13) Then
9200          lngTmp01 = (arr_varCtl(C_TOP, 13) - .opgFormat_box.Top)
9210          .opgFormat_box.Top = (.opgFormat_box.Top + lngTmp01)
9220          .opgFormat_lbl.Top = (.opgFormat_lbl.Top + lngTmp01)
9230          .opgFormat_lbl_dim_hi.Top = (.opgFormat_lbl_dim_hi.Top + lngTmp01)
9240          .opgFormat_vline01.Top = (.opgFormat_vline01.Top + lngTmp01)
9250          .opgFormat_vline02.Top = (.opgFormat_vline02.Top + lngTmp01)
9260          .opgFormatName_lbl.Top = (.opgFormatName_lbl.Top + lngTmp01)
9270          .opgFormatName_box.Top = (.opgFormatName_box.Top + lngTmp01)
9280          lngCnt = (lngTmp01 / lngTpp)
9290          For lngX = 1& To lngCnt
9300            .opgFormatName_optAsWritten.Top = (.opgFormatName_optAsWritten.Top + lngTpp)
9310            .opgFormatName_optAsWritten_lbl.Top = (.opgFormatName_optAsWritten_lbl.Top + lngTpp)
9320            .opgFormatName_optLastFirst.Top = (.opgFormatName_optLastFirst.Top + lngTpp)
9330            .opgFormatName_optLastFirst_lbl.Top = (.opgFormatName_optLastFirst_lbl.Top + lngTpp)
9340            .opgFormatName.Top = (.opgFormatName.Top + lngTpp)
9350          Next  ' ** lngX.
9360          .opgFormatZip_box.Top = (.opgFormatZip_box.Top + lngTmp01)
9370          .opgFormatZip_lbl.Top = (.opgFormatZip_lbl.Top + lngTmp01)
9380          For lngX = 1& To lngCnt
9390            .opgFormatZip_optFormatted.Top = (.opgFormatZip_optFormatted.Top + lngTpp)
9400            .opgFormatZip_optFormatted_lbl.Top = (.opgFormatZip_optFormatted_lbl.Top + lngTpp)
9410            .opgFormatZip_optUnformatted.Top = (.opgFormatZip_optUnformatted.Top + lngTpp)
9420            .opgFormatZip_optUnformatted_lbl.Top = (.opgFormatZip_optUnformatted_lbl.Top + lngTpp)
9430            .opgFormatZip.Top = (.opgFormatZip.Top + lngTpp)
9440          Next  ' ** lngX.
9450          .opgFormatPhone_box.Top = (.opgFormatPhone_box.Top + lngTmp01)
9460          .opgFormatPhone_lbl.Top = (.opgFormatPhone_lbl.Top + lngTmp01)
9470          For lngX = 1& To lngCnt
9480            .opgFormatPhone_optFormatted.Top = (.opgFormatPhone_optFormatted.Top + lngTpp)
9490            .opgFormatPhone_optFormatted_lbl.Top = (.opgFormatPhone_optFormatted_lbl.Top + lngTpp)
9500            .opgFormatPhone_optUnformatted.Top = (.opgFormatPhone_optUnformatted.Top + lngTpp)
9510            .opgFormatPhone_optUnformatted_lbl.Top = (.opgFormatPhone_optUnformatted_lbl.Top + lngTpp)
9520            .opgFormatPhone.Top = (.opgFormatPhone.Top + lngTpp)
9530          Next  ' ** lngX.
9540        End If

9550        If .opgPrint_box.Top <> arr_varCtl(C_TOP, 14) Then
9560          lngTmp01 = (arr_varCtl(C_TOP, 14) - .opgPrint_box.Top)
9570          .opgPrint_box.Top = (.opgPrint_box.Top + lngTmp01)
9580          .opgPrint_lbl.Top = (.opgPrint_lbl.Top + lngTmp01)
9590          .opgPrint_lbl_dim_hi.Top = (.opgPrint_lbl_dim_hi.Top + lngTmp01)
9600          .opgPrint_hline01.Top = (.opgPrint_hline01.Top + lngTmp01)
9610          .opgPrint_hline02.Top = (.opgPrint_hline02.Top + lngTmp01)
9620          .opgPrint_vline01.Top = (.opgPrint_vline01.Top + lngTmp01)
9630          .opgPrint_vline02.Top = (.opgPrint_vline02.Top + lngTmp01)
9640          .opgPrintAddress_box.Top = (.opgPrintAddress_box.Top + lngTmp01)
9650          .opgPrintAddress_lbl.Top = (.opgPrintAddress_lbl.Top + lngTmp01)
9660          lngCnt = (lngTmp01 / lngTpp)
9670          For lngX = 1& To lngCnt
9680            .opgPrintAddress_optSeparate.Top = (.opgPrintAddress_optSeparate.Top + lngTpp)
9690            .opgPrintAddress_optSeparate_lbl.Top = (.opgPrintAddress_optSeparate_lbl.Top + lngTpp)
9700            .opgPrintAddress_optCombined.Top = (.opgPrintAddress_optCombined.Top + lngTpp)
9710            .opgPrintAddress_optCombined_lbl.Top = (.opgPrintAddress_optCombined_lbl.Top + lngTpp)
9720            .opgPrintAddress_optNone.Top = (.opgPrintAddress_optNone.Top + lngTpp)
9730            .opgPrintAddress_optNone_lbl.Top = (.opgPrintAddress_optNone_lbl.Top + lngTpp)
9740            .opgPrintAddress.Top = (.opgPrintAddress.Top + lngTpp)
9750          Next  ' ** lngX.
9760          .opgPrintCSZ_box.Top = (.opgPrintCSZ_box.Top + lngTmp01)
9770          .opgPrintCSZ_lbl.Top = (.opgPrintCSZ_lbl.Top + lngTmp01)
9780          For lngX = 1& To lngCnt
9790            .opgPrintCSZ_optSeparate.Top = (.opgPrintCSZ_optSeparate.Top + lngTpp)
9800            .opgPrintCSZ_optSeparate_lbl.Top = (.opgPrintCSZ_optSeparate_lbl.Top + lngTpp)
9810            .opgPrintCSZ_optCombined.Top = (.opgPrintCSZ_optCombined.Top + lngTpp)
9820            .opgPrintCSZ_optCombined_lbl.Top = (.opgPrintCSZ_optCombined_lbl.Top + lngTpp)
9830            .opgPrintCSZ_optNone.Top = (.opgPrintCSZ_optNone.Top + lngTpp)
9840            .opgPrintCSZ_optNone_lbl.Top = (.opgPrintCSZ_optNone_lbl.Top + lngTpp)
9850            .opgPrintCSZ.Top = (.opgPrintCSZ.Top + lngTpp)
9860          Next  ' ** lngX.
9870          .opgPrintCSZCP_box.Top = (.opgPrintCSZCP_box.Top + lngTmp01)
9880          .opgPrintCSZCP_lbl.Top = (.opgPrintCSZCP_lbl.Top + lngTmp01)
9890          For lngX = 1& To lngCnt
9900            .opgPrintCSZCP_optSeparate.Top = (.opgPrintCSZCP_optSeparate.Top + lngTpp)
9910            .opgPrintCSZCP_optSeparate_lbl.Top = (.opgPrintCSZCP_optSeparate_lbl.Top + lngTpp)
9920            .opgPrintCSZCP_optCombined.Top = (.opgPrintCSZCP_optCombined.Top + lngTpp)
9930            .opgPrintCSZCP_optCombined_lbl.Top = (.opgPrintCSZCP_optCombined_lbl.Top + lngTpp)
9940            .opgPrintCSZCP_optNone.Top = (.opgPrintCSZCP_optNone.Top + lngTpp)
9950            .opgPrintCSZCP_optNone_lbl.Top = (.opgPrintCSZCP_optNone_lbl.Top + lngTpp)
9960            .opgPrintCSZCP.Top = (.opgPrintCSZCP.Top + lngTpp)
9970          Next  ' ** lngX.
9980          .opgPrintPhone_box.Top = (.opgPrintPhone_box.Top + lngTmp01)
9990          .opgPrintPhone_lbl.Top = (.opgPrintPhone_lbl.Top + lngTmp01)
10000         For lngX = 1& To lngCnt
10010           .opgPrintPhone_optSeparate.Top = (.opgPrintPhone_optSeparate.Top + lngTpp)
10020           .opgPrintPhone_optSeparate_lbl.Top = (.opgPrintPhone_optSeparate_lbl.Top + lngTpp)
10030           .opgPrintPhone_optCombined.Top = (.opgPrintPhone_optCombined.Top + lngTpp)
10040           .opgPrintPhone_optCombined_lbl.Top = (.opgPrintPhone_optCombined_lbl.Top + lngTpp)
10050           .opgPrintPhone_opt1Only.Top = (.opgPrintPhone_opt1Only.Top + lngTpp)
10060           .opgPrintPhone_opt1Only_lbl.Top = (.opgPrintPhone_opt1Only_lbl.Top + lngTpp)
10070           .opgPrintPhone_optNone.Top = (.opgPrintPhone_optNone.Top + lngTpp)
10080           .opgPrintPhone_optNone_lbl.Top = (.opgPrintPhone_optNone_lbl.Top + lngTpp)
10090           .opgPrintPhone.Top = (.opgPrintPhone.Top + lngTpp)
10100         Next  ' ** lngX.
10110         .opgPrintFax_box.Top = (.opgPrintFax_box.Top + lngTmp01)
10120         .opgPrintFax_lbl.Top = (.opgPrintFax_lbl.Top + lngTmp01)
10130         For lngX = 1& To lngCnt
10140           .opgPrintFax_optPrint.Top = (.opgPrintFax_optPrint.Top + lngTpp)
10150           .opgPrintFax_optPrint_lbl.Top = (.opgPrintFax_optPrint_lbl.Top + lngTpp)
10160           .opgPrintFax_optNone.Top = (.opgPrintFax_optNone.Top + lngTpp)
10170           .opgPrintFax_optNone_lbl.Top = (.opgPrintFax_optNone_lbl.Top + lngTpp)
10180           .opgPrintFax.Top = (.opgPrintFax.Top + lngTpp)
10190         Next  ' ** lngX.
10200         .opgPrintEmail_lbl.Top = (.opgPrintEmail_lbl.Top + lngTmp01)
10210         .opgPrintEmail_box.Top = (.opgPrintEmail_box.Top + lngTmp01)
10220         For lngX = 1& To lngCnt
10230           .opgPrintEmail_optPrint.Top = (.opgPrintEmail_optPrint.Top + lngTpp)
10240           .opgPrintEmail_optPrint_lbl.Top = (.opgPrintEmail_optPrint_lbl.Top + lngTpp)
10250           .opgPrintEmail_optNone.Top = (.opgPrintEmail_optNone.Top + lngTpp)
10260           .opgPrintEmail_optNone_lbl.Top = (.opgPrintEmail_optNone_lbl.Top + lngTpp)
10270           .opgPrintEmail.Top = (.opgPrintEmail.Top + lngTpp)
10280         Next  ' ** lngX.
10290       End If

10300       If .cmdReset.Top <> arr_varCtl(C_TOP, 6) Then
10310         .cmdReset.Top = arr_varCtl(C_TOP, 6)
10320         .cmdReset_raised_img.Top = .cmdReset.Top
10330         .cmdReset_raised_semifocus_dots_img.Top = .cmdReset.Top
10340         .cmdReset_raised_focus_img.Top = .cmdReset.Top
10350         .cmdReset_raised_focus_dots_img.Top = .cmdReset.Top
10360         .cmdReset_sunken_focus_dots_img.Top = .cmdReset.Top
10370         .cmdReset_raised_img_dis.Top = .cmdReset.Top
10380       End If

10390       If .Detail_hline03.Top <> arr_varCtl(C_LN_TOP, 1) Then
10400         lngTmp01 = (arr_varCtl(C_LN_TOP, 1) - .Detail_hline03.Top)
10410         .Detail_hline03.Top = (.Detail_hline03.Top + lngTmp01)
10420         .Detail_hline04.Top = (.Detail_hline03.Top + lngTpp)
10430         .Detail_vline03.Top = .Detail_hline03.Top
10440         .Detail_vline04.Top = .Detail_hline03.Top
10450       End If

10460       If .chkShowAcctNum_box.Top <> arr_varCtl(C_LN_TOP, 7) Then
10470         lngTmp01 = (arr_varCtl(C_LN_TOP, 7) - .chkShowAcctNum_box.Top)
10480         .chkShowAcctNum_box.Top = (.chkShowAcctNum_box.Top + lngTmp01)
10490         .chkShowAcctNum_vline01.Top = (.chkShowAcctNum_vline01.Top + lngTmp01)
10500         .chkShowAcctNum_vline02.Top = (.chkShowAcctNum_vline02.Top + lngTmp01)
10510         .chkShowAcctNum.Top = (.chkShowAcctNum.Top + lngTmp01)
10520         .chkShowAcctNum_lbl.Top = (.chkShowAcctNum_lbl.Top + lngTmp01)
10530         .chkEnableCountry_Wide.Top = (.chkEnableCountry_Wide.Top + lngTmp01)
10540         .chkEnableCountry_Wide_lbl.Top = (.chkEnableCountry_Wide_lbl.Top + lngTmp01)
10550         .chkEnableCountry_Compact.Top = .chkEnableCountry_Wide.Top
10560         .chkEnableCountry_Compact_lbl.Top = .chkEnableCountry_Wide_lbl.Top
10570         .chkShowShortName.Top = (.chkShowShortName.Top + lngTmp01)
10580         .chkShowShortName_lbl.Top = (.chkShowShortName_lbl.Top + lngTmp01)
10590         .chkPageOf.Top = (.chkPageOf.Top + lngTmp01)
10600         .chkPageOf_lbl.Top = (.chkPageOf_lbl.Top + lngTmp01)
10610       End If

            'IF THIS OPENS COMPACT AND THEN CHANGES TO WIDE,
            'THE SUBFORM HASN'T BEEN RESIZED!

            ' ** Opening variables are w/ accountno, w/o shortname, w/o country.
10620       Select Case .chkEnableCountry_Wide
            Case True
10630         Select Case .chkShowAcctNum
              Case True
10640           Select Case .chkShowShortName
                Case True
                  ' ** w/ accountno, w/ shortname, w/ county.
10650             lngTmp01 = (.frmAccountContacts_Sub.Form.shortname.Width + (4& * lngTpp))
10660             lngTmp01 = (lngFrm_Width_Orig + arr_varCtl(C_L1_LFT, 20) + lngTmp01)
10670           Case False
                  ' ** w/ accountno, w/o shortname, w/ county.
10680             lngTmp01 = (lngFrm_Width_Orig + arr_varCtl(C_L1_LFT, 20))
10690           End Select
10700         Case False
10710           Select Case .chkShowShortName
                Case True
                  ' ** w/o accountno, w/ shortname, w/ county.
10720             lngTmp01 = (lngFrm_Width_Orig + arr_varCtl(C_L1_LFT, 16) + arr_varCtl(C_L1_LFT, 20))
10730           Case False
                  ' ** w/o accountno, w/o shortname, w/ county.
                  ' ** Option not available.
10740             lngTmp01 = 0&
10750           End Select
10760         End Select
10770         If lngWidth <> lngTmp01 Then
10780           If lngWidth > lngTmp01 Then
10790             lngDiff_Width = (lngWidth - lngTmp01)
                  ' ** Center form.
10800             lngTmp01 = (lngWidth - lngDiff_Width)  ' ** New Width.
10810             If lngTmp01 > lngFrm_Width_Orig Then
10820               lngTmp01 = ((lngTmp01 - lngFrm_Width_Orig) / 2)
10830               If lngMonitorNum = 1& Then lngTmp03 = lngTop
10840               DoCmd.SelectObject acForm, frm.Name, False
10850               DoEvents
10860               DoCmd.MoveSize (lngFrm_Left_Orig - lngTmp01), lngTmp03, (lngWidth - lngDiff_Width), (lngHeight + lngDiff_Height)  'lngTop
10870               If lngMonitorNum > 1& Then
10880                 LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
10890               End If
10900             Else
10910               lngTmp01 = ((lngFrm_Width_Orig - lngTmp01) / 2)
10920               If lngMonitorNum = 1& Then lngTmp03 = lngTop
10930               DoCmd.SelectObject acForm, frm.Name, False
10940               DoEvents
10950               DoCmd.MoveSize (lngFrm_Left_Orig + lngTmp01), lngTmp03, (lngWidth - lngDiff_Width), (lngHeight + lngDiff_Height)  'lngTop
10960               If lngMonitorNum > 1& Then
10970                 LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
10980               End If
10990             End If
11000           Else
11010             lngDiff_Width = (lngTmp01 - lngWidth)
                  ' ** Center form.
11020             lngTmp01 = (lngWidth + lngDiff_Width)
11030             If lngTmp01 > lngFrm_Width_Orig Then
11040               lngTmp01 = ((lngTmp01 - lngFrm_Width_Orig) / 2)
11050               If lngMonitorNum = 1& Then lngTmp03 = lngTop
11060               DoCmd.SelectObject acForm, frm.Name, False
11070               DoEvents
11080               DoCmd.MoveSize (lngFrm_Left_Orig - lngTmp01), lngTmp03, (lngWidth + lngDiff_Width), (lngHeight + lngDiff_Height)  'lngTop
11090               If lngMonitorNum > 1& Then
11100                 LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
11110               End If
11120             Else
11130               lngTmp01 = ((lngFrm_Width_Orig - lngTmp01) / 2)
11140               If lngMonitorNum = 1& Then lngTmp03 = lngTop
11150               DoCmd.SelectObject acForm, frm.Name, False
11160               DoEvents
11170               DoCmd.MoveSize (lngFrm_Left_Orig + lngTmp01), lngTmp03, (lngWidth + lngDiff_Width), (lngHeight + lngDiff_Height)  'lngTop
11180               If lngMonitorNum > 1& Then
11190                 LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
11200               End If
11210             End If
11220           End If
11230           Select Case .chkShowAcctNum
                Case True
11240             Select Case .chkShowShortName
                  Case True
                    ' ** w/ accountno, w/ shortname, w/ county.
11250               lngTmp01 = ((.frmAccountContacts_Sub.Form.shortname.Width + (4& * lngTpp)) + arr_varCtl(C_L1_LFT, 20))
11260             Case False
                    ' ** w/ accountno, w/o shortname, w/ county.
11270               lngTmp01 = arr_varCtl(C_L1_LFT, 20)
11280             End Select
11290           Case False
11300             Select Case .chkShowShortName
                  Case True
                    ' ** w/o accountno, w/ shortname, w/ county.
11310               lngTmp01 = (arr_varCtl(C_L1_LFT, 16) + arr_varCtl(C_L1_LFT, 20))
11320             Case False
                    ' ** w/o accountno, w/o shortname, w/ county.
                    ' ** Option not available.
11330               lngTmp01 = 0&
11340             End Select
11350           End Select
11360           .cmdClose.Left = (arr_varCtl(C_LFT, 2) + lngTmp01)
11370           .cmdAdd.Left = (arr_varCtl(C_LFT, 3) + lngTmp01)
11380           .cmdDelete.Left = (arr_varCtl(C_LFT, 4) + lngTmp01)
11390           .cmdUpdate.Left = (arr_varCtl(C_LFT, 5) = lngTmp01)
11400           .chkAlwaysUpdate.Left = (arr_varCtl(C_LFT, 10) + lngTmp01)
11410           .chkAlwaysUpdate_lbl.Left = (arr_varCtl(C_L1_LFT, 10) + lngTmp01)
11420           .chkAlwaysUpdate_lbl2.Left = .chkAlwaysUpdate_lbl.Left
11430           .chkAlwaysUpdate_lbl2_dim_hi.Left = (.chkAlwaysUpdate_lbl2.Left + lngTpp)
11440           .Header_vline01.Left = (lngForm_Width + lngTmp01)
11450           .Header_vline02.Left = .Header_vline01.Left
11460           .Header_hline01.Width = (lngForm_Width + lngTmp01)
11470           .Header_hline02.Width = .Header_hline01.Width
11480           .Detail_vline01.Left = .Header_vline01.Left
11490           .Detail_vline02.Left = .Header_vline01.Left
11500           .Detail_hline01.Width = .Header_hline01.Width
11510           .Detail_hline02.Width = .Header_hline01.Width
11520           .Detail_vline03.Left = .Header_vline01.Left
11530           .Detail_vline04.Left = .Header_vline01.Left
11540           .Detail_hline03.Width = .Header_hline01.Width
11550           .Detail_hline04.Width = .Header_hline01.Width
11560           .Footer_vline01.Left = .Header_vline01.Left
11570           .Footer_vline02.Left = .Header_vline01.Left
11580           .Footer_hline01.Width = .Header_hline01.Width
11590           .Footer_hline02.Width = .Header_hline01.Width
11600         End If
11610       Case False
11620         Select Case .chkShowAcctNum
              Case True
11630           Select Case .chkShowShortName
                Case True
                  ' ** w/ accountno, w/ shortname, w/o county.
11640             lngTmp01 = (lngFrm_Width_Orig + (.frmAccountContacts_Sub.Form.shortname.Width + (4& * lngTpp)))
11650           Case False
                  ' ** w/ accountno, w/o shortname, w/o country.
11660             lngTmp01 = lngFrm_Width_Orig
11670           End Select
11680         Case False
11690           Select Case .chkShowShortName
                Case True
                  ' ** w/o accountno, w/ shortname, w/o county.
11700             lngTmp01 = (lngFrm_Width_Orig + arr_varCtl(C_L1_LFT, 16))
11710           Case False
                  ' ** w/o accountno, w/o shortname, w/o county.
                  ' ** Option not available.
11720             lngTmp01 = 0&
11730           End Select
11740         End Select
11750         If lngWidth <> lngTmp01 Then
11760           If lngWidth > lngTmp01 Then
11770             lngDiff_Width = (lngWidth - lngTmp01)
                  ' ** Center form.
11780             lngTmp01 = (lngWidth - lngDiff_Width)
11790             If lngTmp01 > lngFrm_Width_Orig Then
11800               lngTmp01 = ((lngTmp01 - lngFrm_Width_Orig) / 2)
11810               If lngMonitorNum = 1& Then lngTmp03 = lngTop
11820               DoCmd.SelectObject acForm, frm.Name, False
11830               DoEvents
11840               DoCmd.MoveSize (lngFrm_Left_Orig - lngTmp01), lngTmp03, (lngWidth - lngDiff_Width), (lngHeight + lngDiff_Height)  'lngTop
11850               If lngMonitorNum > 1& Then
11860                 LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
11870               End If
11880             Else
11890               lngTmp01 = ((lngFrm_Width_Orig - lngTmp01) / 2)
11900               If lngMonitorNum = 1& Then lngTmp03 = lngTop
11910               DoCmd.SelectObject acForm, frm.Name, False
11920               DoEvents
11930               DoCmd.MoveSize (lngFrm_Left_Orig + lngTmp01), lngTmp03, (lngWidth - lngDiff_Width), (lngHeight + lngDiff_Height)  'lngTop
11940               If lngMonitorNum > 1& Then
11950                 LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
11960               End If
11970             End If
11980           Else
11990             lngDiff_Width = (lngTmp01 - lngWidth)
                  ' ** Center form.
12000             lngTmp01 = (lngWidth + lngDiff_Width)
12010             If lngTmp01 > lngFrm_Width_Orig Then
12020               lngTmp01 = ((lngTmp01 - lngFrm_Width_Orig) / 2)
12030               If lngMonitorNum = 1& Then lngTmp03 = lngTop
12040               DoCmd.SelectObject acForm, frm.Name, False
12050               DoEvents
12060               DoCmd.MoveSize (lngFrm_Left_Orig - lngTmp01), lngTmp03, (lngWidth + lngDiff_Width), (lngHeight + lngDiff_Height)  'lngTop
12070               If lngMonitorNum > 1& Then
12080                 LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
12090               End If
12100             Else
12110               lngTmp01 = ((lngFrm_Width_Orig - lngTmp01) / 2)
12120               If lngMonitorNum = 1& Then lngTmp03 = lngTop
12130               DoCmd.SelectObject acForm, frm.Name, False
12140               DoEvents
12150               DoCmd.MoveSize (lngFrm_Left_Orig + lngTmp01), lngTmp03, (lngWidth + lngDiff_Width), (lngHeight + lngDiff_Height)  'lngTop
12160               If lngMonitorNum > 1& Then
12170                 LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
12180               End If
12190             End If
12200           End If
12210           Select Case .chkShowAcctNum
                Case True
12220             Select Case .chkShowShortName
                  Case True
                    ' ** w/ accountno, w/ shortname, w/o county.
12230               lngTmp01 = (.frmAccountContacts_Sub.Form.shortname.Width + (4& * lngTpp))
12240             Case False
                    ' ** w/ accountno, w/o shortname, w/o country.
12250               lngTmp01 = 0&
12260             End Select
12270           Case False
12280             Select Case .chkShowShortName
                  Case True
                    ' ** w/o accountno, w/ shortname, w/o county.
12290               lngTmp01 = arr_varCtl(C_L1_LFT, 16)
12300             Case False
                    ' ** w/o accountno, w/o shortname, w/o county.
                    ' ** Option not available.
12310               lngTmp01 = 0&
12320             End Select
12330           End Select
12340           .cmdClose.Left = (arr_varCtl(C_LFT, 2) + lngTmp01)
12350           .cmdAdd.Left = (arr_varCtl(C_LFT, 3) + lngTmp01)
12360           .cmdDelete.Left = (arr_varCtl(C_LFT, 4) + lngTmp01)
12370           .cmdUpdate.Left = (arr_varCtl(C_LFT, 5) + lngTmp01)
12380           .chkAlwaysUpdate.Left = (arr_varCtl(C_LFT, 10) + lngTmp01)
12390           .chkAlwaysUpdate_lbl.Left = (arr_varCtl(C_L1_LFT, 10) + lngTmp01)
12400           .chkAlwaysUpdate_lbl2.Left = .chkAlwaysUpdate_lbl.Left
12410           .chkAlwaysUpdate_lbl2_dim_hi.Left = (.chkAlwaysUpdate_lbl2.Left + lngTpp)
12420           .Header_vline01.Left = (lngForm_Width + lngTmp01)
12430           .Header_vline02.Left = .Header_vline01.Left
12440           .Header_hline01.Width = (lngForm_Width + lngTmp01)
12450           .Header_hline02.Width = .Header_hline01.Width
12460           .Detail_vline01.Left = .Header_vline01.Left
12470           .Detail_vline02.Left = .Header_vline01.Left
12480           .Detail_hline01.Width = .Header_hline01.Width
12490           .Detail_hline02.Width = .Header_hline01.Width
12500           .Detail_vline03.Left = .Header_vline01.Left
12510           .Detail_vline04.Left = .Header_vline01.Left
12520           .Detail_hline03.Width = .Header_hline01.Width
12530           .Detail_hline04.Width = .Header_hline01.Width
12540           .Footer_vline01.Left = .Header_vline01.Left
12550           .Footer_vline02.Left = .Header_vline01.Left
12560           .Footer_hline01.Width = .Header_hline01.Width
12570           .Footer_hline02.Width = .Header_hline01.Width
12580           .frmAccountContacts_Sub.Width = (arr_varCtl(C_F1_WDT, 0) + lngTmp01)
12590           .frmAccountContacts_Sub_box.Width = ((arr_varCtl(C_F1_WDT, 0) + lngTmp01) + (2& * lngTpp))
12600         End If
12610       End Select
            ' ** Seems to be needed.
12620       .chkShowAcctNum_AfterUpdate  ' ** Procedure: Above.
12630       .chkShowShortName_AfterUpdate  ' ** Procedure: Above.
12640       .chkEnableCountry_Wide_AfterUpdate  ' ** Procedure: Above.

12650     Case 2  ' ** Compact

            'RIGHT NOW, THIS IS ALL WITH COUNTRY!

12660       .frmAccountContacts_Sub2.Visible = True
12670       .frmAccountContacts_Sub2_box.Visible = True
12680       .frmAccountContacts_Sub.Visible = False
12690       .frmAccountContacts_Sub_box.Visible = False

12700       .Nav_Sub2_box01.Visible = True
12710       .Nav_Sub2_hline01.Visible = True
12720       .Nav_Sub2_hline02.Visible = True
12730       .Nav_Sub2_vline01.Visible = True
12740       .Nav_Sub2_vline02.Visible = True
12750       .Nav_Sub2_vline03.Visible = True
12760       .Nav_Sub2_vline04.Visible = True
12770       .Nav_Sub1_box01.Visible = False
12780       .Nav_Sub1_hline01.Visible = False
12790       .Nav_Sub1_hline02.Visible = False
12800       .Nav_Sub1_vline01.Visible = False
12810       .Nav_Sub1_vline02.Visible = False
12820       .Nav_Sub1_vline03.Visible = False
12830       .Nav_Sub1_vline04.Visible = False

12840       .chkShowAcctNum = True
12850       .chkShowAcctNum_lbl.FontBold = True
12860       .chkShowAcctNum.Enabled = False
12870       .chkShowShortName = True
12880       .chkShowShortName_lbl.FontBold = True
12890       .chkShowShortName.Enabled = False

12900       lngTmp01 = ((.frmAccountContacts_Sub2_box.Top + .frmAccountContacts_Sub2_box.Height) + lngTpp)
12910       lngDiff_Height = arr_varCtl(C_TOP, 11) - lngTmp01

12920       .ShortcutMenu_lbl.Top = lngTmp01
12930       .ShortcutMenu_up_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
12940       .ShortcutMenu_down_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
12950       lngTmp01 = (.ShortcutMenu_lbl.Top + .ShortcutMenu_lbl.Height)

12960       .Detail_hline01.Top = lngTmp01
12970       .Detail_hline02.Top = (.Detail_hline01.Top + lngTpp)
12980       .Detail_vline01.Top = .Detail_hline01.Top
12990       .Detail_vline02.Top = .Detail_hline01.Top

13000       If .opgView_box2.Top <> (arr_varCtl(C_TOP, 12) - lngDiff_Height) Then
13010         lngTmp01 = (.opgView_box2.Top - (arr_varCtl(C_TOP, 12) - lngDiff_Height))
13020         .opgView_box.Top = (.opgView_box.Top - lngTmp01)
13030         .opgView_box2.Top = (.opgView_box2.Top - lngTmp01)
13040         .opgView_lbl.Top = (.opgView_lbl.Top - lngTmp01)
13050         lngCnt = (lngTmp01 / lngTpp)
13060         For lngX = 1& To lngCnt
13070           .opgView.Top = (.opgView.Top - lngTpp)
13080           .opgView_optWide.Top = (.opgView_optWide.Top - lngTpp)
13090           .opgView_optWide_lbl.Top = (.opgView_optWide_lbl.Top - lngTpp)
13100           .opgView_optCompact.Top = (.opgView_optCompact.Top - lngTpp)
13110           .opgView_optCompact_lbl.Top = (.opgView_optCompact_lbl.Top - lngTpp)
13120         Next  ' ** lngX.
13130       End If

13140       If .opgFormat_box.Top <> (arr_varCtl(C_TOP, 13) - lngDiff_Height) Then
13150         lngTmp01 = (.opgFormat_box.Top - (arr_varCtl(C_TOP, 13) - lngDiff_Height))
13160         .opgFormat_box.Top = (.opgFormat_box.Top - lngTmp01)
13170         .opgFormat_lbl.Top = (.opgFormat_lbl.Top - lngTmp01)
13180         .opgFormat_lbl_dim_hi.Top = (.opgFormat_lbl_dim_hi.Top - lngTmp01)
13190         .opgFormat_vline01.Top = (.opgFormat_vline01.Top - lngTmp01)
13200         .opgFormat_vline02.Top = (.opgFormat_vline02.Top - lngTmp01)
13210         .opgFormatName_lbl.Top = (.opgFormatName_lbl.Top - lngTmp01)
13220         .opgFormatName_box.Top = (.opgFormatName_box.Top - lngTmp01)
13230         lngCnt = (lngTmp01 / lngTpp)
13240         For lngX = 1& To lngCnt
13250           .opgFormatName.Top = (.opgFormatName.Top - lngTpp)
13260           .opgFormatName_optAsWritten.Top = (.opgFormatName_optAsWritten.Top - lngTpp)
13270           .opgFormatName_optAsWritten_lbl.Top = (.opgFormatName_optAsWritten_lbl.Top - lngTpp)
13280           .opgFormatName_optLastFirst.Top = (.opgFormatName_optLastFirst.Top - lngTpp)
13290           .opgFormatName_optLastFirst_lbl.Top = (.opgFormatName_optLastFirst_lbl.Top - lngTpp)
13300         Next  ' ** lngX.
13310         .opgFormatZip_box.Top = (.opgFormatZip_box.Top - lngTmp01)
13320         .opgFormatZip_lbl.Top = (.opgFormatZip_lbl.Top - lngTmp01)
13330         For lngX = 1& To lngCnt
13340           .opgFormatZip.Top = (.opgFormatZip.Top - lngTpp)
13350           .opgFormatZip_optFormatted.Top = (.opgFormatZip_optFormatted.Top - lngTpp)
13360           .opgFormatZip_optFormatted_lbl.Top = (.opgFormatZip_optFormatted_lbl.Top - lngTpp)
13370           .opgFormatZip_optUnformatted.Top = (.opgFormatZip_optUnformatted.Top - lngTpp)
13380           .opgFormatZip_optUnformatted_lbl.Top = (.opgFormatZip_optUnformatted_lbl.Top - lngTpp)
13390         Next  ' ** lngX.
13400         .opgFormatPhone_box.Top = (.opgFormatPhone_box.Top - lngTmp01)
13410         .opgFormatPhone_lbl.Top = (.opgFormatPhone_lbl.Top - lngTmp01)
13420         For lngX = 1& To lngCnt
13430           .opgFormatPhone.Top = (.opgFormatPhone.Top - lngTpp)
13440           .opgFormatPhone_optFormatted.Top = (.opgFormatPhone_optFormatted.Top - lngTpp)
13450           .opgFormatPhone_optFormatted_lbl.Top = (.opgFormatPhone_optFormatted_lbl.Top - lngTpp)
13460           .opgFormatPhone_optUnformatted.Top = (.opgFormatPhone_optUnformatted.Top - lngTpp)
13470           .opgFormatPhone_optUnformatted_lbl.Top = (.opgFormatPhone_optUnformatted_lbl.Top - lngTpp)
13480         Next  ' ** lngX.
13490       End If

13500       If .opgPrint_box.Top <> (arr_varCtl(C_TOP, 14) - lngDiff_Height) Then
13510         lngTmp01 = (.opgPrint_box.Top - (arr_varCtl(C_TOP, 14) - lngDiff_Height))
13520         .opgPrint_box.Top = (.opgPrint_box.Top - lngTmp01)
13530         .opgPrint_lbl.Top = (.opgPrint_lbl.Top - lngTmp01)
13540         .opgPrint_lbl_dim_hi.Top = (.opgPrint_lbl_dim_hi.Top - lngTmp01)
13550         .opgPrint_hline01.Top = (.opgPrint_hline01.Top - lngTmp01)
13560         .opgPrint_hline02.Top = (.opgPrint_hline02.Top - lngTmp01)
13570         .opgPrint_vline01.Top = (.opgPrint_vline01.Top - lngTmp01)
13580         .opgPrint_vline02.Top = (.opgPrint_vline02.Top - lngTmp01)
13590         .opgPrintAddress_box.Top = (.opgPrintAddress_box.Top - lngTmp01)
13600         .opgPrintAddress_lbl.Top = (.opgPrintAddress_lbl.Top - lngTmp01)
13610         lngCnt = (lngTmp01 / lngTpp)
13620         For lngX = 1& To lngCnt
13630           .opgPrintAddress.Top = (.opgPrintAddress.Top - lngTpp)
13640           .opgPrintAddress_optSeparate.Top = (.opgPrintAddress_optSeparate.Top - lngTpp)
13650           .opgPrintAddress_optSeparate_lbl.Top = (.opgPrintAddress_optSeparate_lbl.Top - lngTpp)
13660           .opgPrintAddress_optCombined.Top = (.opgPrintAddress_optCombined.Top - lngTpp)
13670           .opgPrintAddress_optCombined_lbl.Top = (.opgPrintAddress_optCombined_lbl.Top - lngTpp)
13680           .opgPrintAddress_optNone.Top = (.opgPrintAddress_optNone.Top - lngTpp)
13690           .opgPrintAddress_optNone_lbl.Top = (.opgPrintAddress_optNone_lbl.Top - lngTpp)
13700         Next  ' ** lngX.
13710         .opgPrintCSZ_box.Top = (.opgPrintCSZ_box.Top - lngTmp01)
13720         .opgPrintCSZ_lbl.Top = (.opgPrintCSZ_lbl.Top - lngTmp01)
13730         For lngX = 1& To lngCnt
13740           .opgPrintCSZ.Top = (.opgPrintCSZ.Top - lngTpp)
13750           .opgPrintCSZ_optSeparate.Top = (.opgPrintCSZ_optSeparate.Top - lngTpp)
13760           .opgPrintCSZ_optSeparate_lbl.Top = (.opgPrintCSZ_optSeparate_lbl.Top - lngTpp)
13770           .opgPrintCSZ_optCombined.Top = (.opgPrintCSZ_optCombined.Top - lngTpp)
13780           .opgPrintCSZ_optCombined_lbl.Top = (.opgPrintCSZ_optCombined_lbl.Top - lngTpp)
13790           .opgPrintCSZ_optNone.Top = (.opgPrintCSZ_optNone.Top - lngTpp)
13800           .opgPrintCSZ_optNone_lbl.Top = (.opgPrintCSZ_optNone_lbl.Top - lngTpp)
13810         Next  ' ** lngX.
13820         .opgPrintCSZCP_box.Top = (.opgPrintCSZCP_box.Top - lngTmp01)
13830         .opgPrintCSZCP_lbl.Top = (.opgPrintCSZCP_lbl.Top - lngTmp01)
13840         For lngX = 1& To lngCnt
13850           .opgPrintCSZCP.Top = (.opgPrintCSZCP.Top - lngTpp)
13860           .opgPrintCSZCP_optSeparate.Top = (.opgPrintCSZCP_optSeparate.Top - lngTpp)
13870           .opgPrintCSZCP_optSeparate_lbl.Top = (.opgPrintCSZCP_optSeparate_lbl.Top - lngTpp)
13880           .opgPrintCSZCP_optCombined.Top = (.opgPrintCSZCP_optCombined.Top - lngTpp)
13890           .opgPrintCSZCP_optCombined_lbl.Top = (.opgPrintCSZCP_optCombined_lbl.Top - lngTpp)
13900           .opgPrintCSZCP_optNone.Top = (.opgPrintCSZCP_optNone.Top - lngTpp)
13910           .opgPrintCSZCP_optNone_lbl.Top = (.opgPrintCSZCP_optNone_lbl.Top - lngTpp)
13920         Next  ' ** lngX.
13930         .opgPrintPhone_box.Top = (.opgPrintPhone_box.Top - lngTmp01)
13940         .opgPrintPhone_lbl.Top = (.opgPrintPhone_lbl.Top - lngTmp01)
13950         For lngX = 1& To lngCnt
13960           .opgPrintPhone.Top = (.opgPrintPhone.Top - lngTpp)
13970           .opgPrintPhone_optSeparate.Top = (.opgPrintPhone_optSeparate.Top - lngTpp)
13980           .opgPrintPhone_optSeparate_lbl.Top = (.opgPrintPhone_optSeparate_lbl.Top - lngTpp)
13990           .opgPrintPhone_optCombined.Top = (.opgPrintPhone_optCombined.Top - lngTpp)
14000           .opgPrintPhone_optCombined_lbl.Top = (.opgPrintPhone_optCombined_lbl.Top - lngTpp)
14010           .opgPrintPhone_opt1Only.Top = (.opgPrintPhone_opt1Only.Top - lngTpp)
14020           .opgPrintPhone_opt1Only_lbl.Top = (.opgPrintPhone_opt1Only_lbl.Top - lngTpp)
14030           .opgPrintPhone_optNone.Top = (.opgPrintPhone_optNone.Top - lngTpp)
14040           .opgPrintPhone_optNone_lbl.Top = (.opgPrintPhone_optNone_lbl.Top - lngTpp)
14050         Next  ' ** lngX.
14060         .opgPrintFax_box.Top = (.opgPrintFax_box.Top - lngTmp01)
14070         .opgPrintFax_lbl.Top = (.opgPrintFax_lbl.Top - lngTmp01)
14080         For lngX = 1& To lngCnt
14090           .opgPrintFax.Top = (.opgPrintFax.Top - lngTpp)
14100           .opgPrintFax_optPrint.Top = (.opgPrintFax_optPrint.Top - lngTpp)
14110           .opgPrintFax_optPrint_lbl.Top = (.opgPrintFax_optPrint_lbl.Top - lngTpp)
14120           .opgPrintFax_optNone.Top = (.opgPrintFax_optNone.Top - lngTpp)
14130           .opgPrintFax_optNone_lbl.Top = (.opgPrintFax_optNone_lbl.Top - lngTpp)
14140         Next  ' ** lngX.
14150         .opgPrintEmail_lbl.Top = (.opgPrintEmail_lbl.Top - lngTmp01)
14160         .opgPrintEmail_box.Top = (.opgPrintEmail_box.Top - lngTmp01)
14170         For lngX = 1& To lngCnt
14180           .opgPrintEmail.Top = (.opgPrintEmail.Top - lngTpp)
14190           .opgPrintEmail_optPrint.Top = (.opgPrintEmail_optPrint.Top - lngTpp)
14200           .opgPrintEmail_optPrint_lbl.Top = (.opgPrintEmail_optPrint_lbl.Top - lngTpp)
14210           .opgPrintEmail_optNone.Top = (.opgPrintEmail_optNone.Top - lngTpp)
14220           .opgPrintEmail_optNone_lbl.Top = (.opgPrintEmail_optNone_lbl.Top - lngTpp)
14230         Next  ' ** lngX.
14240       End If

14250       If .cmdReset.Top <> (arr_varCtl(C_TOP, 6) - lngDiff_Height) Then
14260         lngTmp01 = (.cmdReset.Top - (arr_varCtl(C_TOP, 6) - lngDiff_Height))
14270         .cmdReset.Top = (.cmdReset.Top - lngTmp01)
14280         .cmdReset_raised_img.Top = .cmdReset.Top
14290         .cmdReset_raised_semifocus_dots_img.Top = .cmdReset.Top
14300         .cmdReset_raised_focus_img.Top = .cmdReset.Top
14310         .cmdReset_raised_focus_dots_img.Top = .cmdReset.Top
14320         .cmdReset_sunken_focus_dots_img.Top = .cmdReset.Top
14330         .cmdReset_raised_img_dis.Top = .cmdReset.Top
14340       End If

14350       If .Detail_hline03.Top <> (arr_varCtl(C_LN_TOP, 1) - lngDiff_Height) Then
14360         lngTmp01 = (.Detail_hline03.Top - (arr_varCtl(C_LN_TOP, 1) - lngDiff_Height))
14370         .Detail_hline03.Top = (.Detail_hline03.Top - lngTmp01)
14380         .Detail_hline04.Top = (.Detail_hline03.Top + lngTpp)
14390         .Detail_vline03.Top = .Detail_hline03.Top
14400         .Detail_vline04.Top = .Detail_hline03.Top
14410       End If

14420       If .chkShowAcctNum_box.Top <> (arr_varCtl(C_LN_TOP, 7) - lngDiff_Height) Then
14430         lngTmp01 = (.chkShowAcctNum_box.Top - (arr_varCtl(C_LN_TOP, 7) - lngDiff_Height))
14440         .chkShowAcctNum_box.Top = (.chkShowAcctNum_box.Top - lngTmp01)
14450         .chkShowAcctNum_vline01.Top = (.chkShowAcctNum_vline01.Top - lngTmp01)
14460         .chkShowAcctNum_vline02.Top = (.chkShowAcctNum_vline02.Top - lngTmp01)
14470         .chkShowAcctNum.Top = (.chkShowAcctNum.Top - lngTmp01)
14480         .chkShowAcctNum_lbl.Top = (.chkShowAcctNum_lbl.Top - lngTmp01)
14490         .chkEnableCountry_Wide.Top = (.chkEnableCountry_Wide.Top - lngTmp01)
14500         .chkEnableCountry_Wide_lbl.Top = (.chkEnableCountry_Wide_lbl.Top - lngTmp01)
14510         .chkEnableCountry_Compact.Top = .chkEnableCountry_Wide.Top
14520         .chkEnableCountry_Compact_lbl.Top = .chkEnableCountry_Wide_lbl.Top
14530         .chkShowShortName.Top = (.chkShowShortName.Top - lngTmp01)
14540         .chkShowShortName_lbl.Top = (.chkShowShortName_lbl.Top - lngTmp01)
14550         .chkPageOf.Top = (.chkPageOf.Top - lngTmp01)
14560         .chkPageOf_lbl.Top = (.chkPageOf_lbl.Top - lngTmp01)
14570       End If

14580       lngTmp01 = (.Detail.Height - (arr_varCtl(C_HGT, 1) - lngDiff_Height))
14590       .Detail.Height = (arr_varCtl(C_HGT, 1) - lngDiff_Height)
14600       lngDiff_Height = lngTmp01

            ' ** Original values are w/o shortname.
            ' ** Sub1 is wide, Sub2 is compact.
14610       lngTmp01 = (.frmAccountContacts_Sub2.Width + lngWinSub_Diff)
14620       blnTmp04 = False
            ' ** Should be compact width
14630       If lngWidth <> lngTmp01 Then
              ' ** Its current width isn't equal to the compact width.
14640         If lngWidth > lngTmp01 Then
                ' ** Its current width is greater than the compact width.
14650           lngDiff_Width = (lngWidth - lngTmp01)
                ' ** This should be difference between current width
                ' ** (with whatever's included), and standard compact width.
                ' ** Center form.
14660           lngTmp02 = ((lngFrm_Width_Orig - lngTmp01) / 2)
14670           blnTmp04 = True
14680           If lngMonitorNum = 1& Then lngTmp03 = lngTop
14690           DoCmd.SelectObject acForm, frm.Name, False
14700           DoEvents
14710           DoCmd.MoveSize (lngFrm_Left_Orig + lngTmp02), lngTmp03, lngTmp01, (lngFrm_Height_Orig - lngDiff_Height)   'lngTop  '(lngHeight - lngDiff_Height)
14720           If lngMonitorNum > 1& Then
14730             LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
14740           End If
14750         Else
                ' ** Its current width is less than the compact width.  (How?)
                'lngDiff_Width = (lngTmp01 - lngWidth)
                '' ** Center form.
                'lngTmp01 = (lngWidth + lngDiff_Width)
                'If lngTmp01 > lngFrm_Width_Orig Then
                '  lngTmp01 = ((lngTmp01 - lngFrm_Width_Orig) / 2)
                '  DoCmd.MoveSize (lngFrm_Left_Orig - lngTmp01), lngTop, (lngWidth + lngDiff_Width), (lngHeight + lngDiff_Height)
                'Else
                '  lngTmp01 = ((lngFrm_Width_Orig - lngTmp01) / 2)
                '  DoCmd.MoveSize (lngFrm_Left_Orig + lngTmp01), lngTop, (lngWidth + lngDiff_Width), (lngHeight - lngDiff_Height)
                'End If
14760         End If

14770         lngTmp01 = (.frmAccountContacts_Sub2.Width - arr_varCtl(C_L2_LFT, 0))
14780         lngTmp02 = arr_varCtl(C_LFT, 2) - lngTmp01

14790         If .cmdClose.Left <> (arr_varCtl(C_LFT, 2) - lngTmp02) Then
14800           .cmdClose.Left = (arr_varCtl(C_LFT, 2) - lngTmp02)
14810           .cmdAdd.Left = (arr_varCtl(C_LFT, 3) - lngTmp02)
14820           .cmdDelete.Left = (arr_varCtl(C_LFT, 4) - lngTmp02)
14830           .cmdUpdate.Left = (arr_varCtl(C_LFT, 5) - lngTmp02)
14840           .chkAlwaysUpdate.Left = (arr_varCtl(C_LFT, 10) - lngTmp02)
14850           .chkAlwaysUpdate_lbl.Left = (arr_varCtl(C_L1_LFT, 10) - lngTmp02)
14860           .chkAlwaysUpdate_lbl2.Left = .chkAlwaysUpdate_lbl.Left
14870           .chkAlwaysUpdate_lbl2_dim_hi.Left = (.chkAlwaysUpdate_lbl2.Left + lngTpp)
14880         End If

14890       End If

14900       If blnTmp04 = False Then
14910         lngTmp02 = ((lngFrm_Width_Orig - lngTmp01) / 2)
14920         If lngMonitorNum = 1& Then lngTmp03 = lngTop
14930         DoCmd.SelectObject acForm, frm.Name, False
14940         DoEvents
14950         DoCmd.MoveSize (lngFrm_Left_Orig + lngTmp02), lngTmp03, lngWidth, (lngHeight - lngDiff_Height)  'lngTop
14960         If lngMonitorNum > 1& Then
14970           LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
14980         End If
14990       End If

15000     End Select

15010   End With

EXITP:
15020   Exit Sub

ERRH:
15030   Select Case ERR.Number
        Case 2100  ' ** The control or subform control is too large for this location.
          ' ** Ignore.
15040   Case Else
15050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15060   End Select
15070   Resume EXITP

End Sub

Public Sub CtlArrayLoad(frm As Access.Form)
' ** On this form, this array will make it extremely
' ** difficult to troubleshoot problems!

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "CtlArrayLoad"

        Dim frmSub1 As Access.Form, frmSub2 As Access.Form
        Dim lngE As Long

15110   With frm

15120     If lngCtls = 0& Or IsEmpty(arr_varCtl) = True Then

15130       lngCtls = 0&
15140       ReDim arr_varCtl(C_ELEMS, 0)

            ' ***********************************************
            ' ** Array: arr_varCtl()
            ' **
            ' **   Field  Element  Name          Constant
            ' **   =====  =======  ============  ==========
            ' **     1       0     fld name      C_CNAM
            ' **     2       1     fld top       C_TOP
            ' **     3       2     fld left      C_LFT
            ' **     4       3     fld width     C_WDT
            ' **     5       4     fld height    C_HGT
            ' **     6       5     lbl1 name     C_L1_NAM
            ' **     7       6     lbl1 left     C_L1_LFT
            ' **     8       7     lbl2 name     C_L2_NAM
            ' **     9       8     lbl2 left     C_L2_LFT
            ' **    10       9     line name     C_LN_NAM
            ' **    11      10     line top      C_LN_TOP
            ' **    12      11     line left     C_LN_LFT
            ' **    13      12     frm1 name     C_F1_NAM
            ' **    14      13     frm1 width    C_F1_WDT
            ' **    15      14     frm2 name     C_F2_NAM
            ' **    16      15     frm2 width    C_F2_WDT
            ' **
            ' ***********************************************

15150       Set frmSub1 = .frmAccountContacts_Sub.Form
15160       Set frmSub2 = .frmAccountContacts_Sub2.Form

            ' ** Leave these variables.
            ' **   lngForm_Width
            ' **   lngFrm_Top_Orig
            ' **   lngFrm_Left_Orig
            ' **   lngFrm_Width_Orig
            ' **   lngFrm_Height_Orig
            ' **   lngWinSub_Diff

            ' ** 0. lngSubFrm_Width, lngFormSub_Width, lngSub1Sub2_Offset, lngSub2Btn_Offset.
15170       lngCtls = lngCtls + 1&
15180       lngE = lngCtls - 1&
15190       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
15200       arr_varCtl(C_CNAM, lngE) = "frmAccountContacts_Sub"
15210       arr_varCtl(C_TOP, lngE) = Null
15220       arr_varCtl(C_LFT, lngE) = Null
15230       arr_varCtl(C_WDT, lngE) = Null
15240       arr_varCtl(C_HGT, lngE) = Null
15250       arr_varCtl(C_L1_NAM, lngE) = "lngSub1Sub2_Offset"
15260       arr_varCtl(C_L1_LFT, lngE) = (.frmAccountContacts_Sub.Width - .frmAccountContacts_Sub2.Width) 'lngSub1Sub2_Offset
15270       arr_varCtl(C_L2_NAM, lngE) = "lngSub2Btn_Offset"
15280       arr_varCtl(C_L2_LFT, lngE) = ((.frmAccountContacts_Sub.Left + .frmAccountContacts_Sub.Width) - .cmdClose.Left) 'lngSub2Btn_Offset
15290       arr_varCtl(C_LN_NAM, lngE) = Null
15300       arr_varCtl(C_LN_LFT, lngE) = Null
15310       arr_varCtl(C_LN_TOP, lngE) = Null
15320       arr_varCtl(C_F1_NAM, lngE) = "lngSubFrm_Width"
15330       arr_varCtl(C_F1_WDT, lngE) = .frmAccountContacts_Sub.Width 'lngSubFrm_Width
15340       arr_varCtl(C_F2_NAM, lngE) = "lngFormSub_Width"
15350       arr_varCtl(C_F2_WDT, lngE) = frmSub1.Width 'lngFormSub_Width

            ' ** 1. lngDetail_Height, lngDetailHline3_Top.
15360       lngCtls = lngCtls + 1&
15370       lngE = lngCtls - 1&
15380       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
15390       arr_varCtl(C_CNAM, lngE) = "Detail"
15400       arr_varCtl(C_TOP, lngE) = Null
15410       arr_varCtl(C_LFT, lngE) = Null
15420       arr_varCtl(C_WDT, lngE) = Null
15430       arr_varCtl(C_HGT, lngE) = .Detail.Height 'lngDetail_Height
15440       arr_varCtl(C_L1_NAM, lngE) = Null
15450       arr_varCtl(C_L1_LFT, lngE) = Null
15460       arr_varCtl(C_L2_NAM, lngE) = Null
15470       arr_varCtl(C_L2_LFT, lngE) = Null
15480       arr_varCtl(C_LN_NAM, lngE) = "lngDetailHline3_Top"
15490       arr_varCtl(C_LN_LFT, lngE) = Null
15500       arr_varCtl(C_LN_TOP, lngE) = .Detail_hline03.Top 'lngDetailHline3_Top
15510       arr_varCtl(C_F1_NAM, lngE) = Null
15520       arr_varCtl(C_F1_WDT, lngE) = Null
15530       arr_varCtl(C_F2_NAM, lngE) = Null
15540       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 2. lngClose_Left.
15550       lngCtls = lngCtls + 1&
15560       lngE = lngCtls - 1&
15570       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
15580       arr_varCtl(C_CNAM, lngE) = "cmdClose"
15590       arr_varCtl(C_TOP, lngE) = Null
15600       arr_varCtl(C_LFT, lngE) = .cmdClose.Left 'lngClose_Left
15610       arr_varCtl(C_WDT, lngE) = Null
15620       arr_varCtl(C_HGT, lngE) = Null
15630       arr_varCtl(C_L1_NAM, lngE) = Null
15640       arr_varCtl(C_L1_LFT, lngE) = Null
15650       arr_varCtl(C_L2_NAM, lngE) = Null
15660       arr_varCtl(C_L2_LFT, lngE) = Null
15670       arr_varCtl(C_LN_NAM, lngE) = Null
15680       arr_varCtl(C_LN_LFT, lngE) = Null
15690       arr_varCtl(C_LN_TOP, lngE) = Null
15700       arr_varCtl(C_F1_NAM, lngE) = Null
15710       arr_varCtl(C_F1_WDT, lngE) = Null
15720       arr_varCtl(C_F2_NAM, lngE) = Null
15730       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 3. lngAdd_Left.
15740       lngCtls = lngCtls + 1&
15750       lngE = lngCtls - 1&
15760       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
15770       arr_varCtl(C_CNAM, lngE) = "cmdAdd"
15780       arr_varCtl(C_TOP, lngE) = Null
15790       arr_varCtl(C_LFT, lngE) = .cmdAdd.Left 'lngAdd_Left
15800       arr_varCtl(C_WDT, lngE) = Null
15810       arr_varCtl(C_HGT, lngE) = Null
15820       arr_varCtl(C_L1_NAM, lngE) = Null
15830       arr_varCtl(C_L1_LFT, lngE) = Null
15840       arr_varCtl(C_L2_NAM, lngE) = Null
15850       arr_varCtl(C_L2_LFT, lngE) = Null
15860       arr_varCtl(C_LN_NAM, lngE) = Null
15870       arr_varCtl(C_LN_LFT, lngE) = Null
15880       arr_varCtl(C_LN_TOP, lngE) = Null
15890       arr_varCtl(C_F1_NAM, lngE) = Null
15900       arr_varCtl(C_F1_WDT, lngE) = Null
15910       arr_varCtl(C_F2_NAM, lngE) = Null
15920       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 4. lngDelete_Left.
15930       lngCtls = lngCtls + 1&
15940       lngE = lngCtls - 1&
15950       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
15960       arr_varCtl(C_CNAM, lngE) = "cmdDelete"
15970       arr_varCtl(C_TOP, lngE) = Null
15980       arr_varCtl(C_LFT, lngE) = .cmdDelete.Left 'lngDelete_Left
15990       arr_varCtl(C_WDT, lngE) = Null
16000       arr_varCtl(C_HGT, lngE) = Null
16010       arr_varCtl(C_L1_NAM, lngE) = Null
16020       arr_varCtl(C_L1_LFT, lngE) = Null
16030       arr_varCtl(C_L2_NAM, lngE) = Null
16040       arr_varCtl(C_L2_LFT, lngE) = Null
16050       arr_varCtl(C_LN_NAM, lngE) = Null
16060       arr_varCtl(C_LN_LFT, lngE) = Null
16070       arr_varCtl(C_LN_TOP, lngE) = Null
16080       arr_varCtl(C_F1_NAM, lngE) = Null
16090       arr_varCtl(C_F1_WDT, lngE) = Null
16100       arr_varCtl(C_F2_NAM, lngE) = Null
16110       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 5. lngUpdate_Left.
16120       lngCtls = lngCtls + 1&
16130       lngE = lngCtls - 1&
16140       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
16150       arr_varCtl(C_CNAM, lngE) = "cmdUpdate"
16160       arr_varCtl(C_TOP, lngE) = Null
16170       arr_varCtl(C_LFT, lngE) = .cmdUpdate.Left 'lngUpdate_Left
16180       arr_varCtl(C_WDT, lngE) = Null
16190       arr_varCtl(C_HGT, lngE) = Null
16200       arr_varCtl(C_L1_NAM, lngE) = Null
16210       arr_varCtl(C_L1_LFT, lngE) = Null
16220       arr_varCtl(C_L2_NAM, lngE) = Null
16230       arr_varCtl(C_L2_LFT, lngE) = Null
16240       arr_varCtl(C_LN_NAM, lngE) = Null
16250       arr_varCtl(C_LN_LFT, lngE) = Null
16260       arr_varCtl(C_LN_TOP, lngE) = Null
16270       arr_varCtl(C_F1_NAM, lngE) = Null
16280       arr_varCtl(C_F1_WDT, lngE) = Null
16290       arr_varCtl(C_F2_NAM, lngE) = Null
16300       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 6. lngReset_Left, lngReset_Top.
16310       lngCtls = lngCtls + 1&
16320       lngE = lngCtls - 1&
16330       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
16340       arr_varCtl(C_CNAM, lngE) = "cmdReset"
16350       arr_varCtl(C_TOP, lngE) = .cmdReset.Top 'lngReset_Top
16360       arr_varCtl(C_LFT, lngE) = .cmdReset.Left 'lngReset_Left
16370       arr_varCtl(C_WDT, lngE) = Null
16380       arr_varCtl(C_HGT, lngE) = Null
16390       arr_varCtl(C_L1_NAM, lngE) = Null
16400       arr_varCtl(C_L1_LFT, lngE) = Null
16410       arr_varCtl(C_L2_NAM, lngE) = Null
16420       arr_varCtl(C_L2_LFT, lngE) = Null
16430       arr_varCtl(C_LN_NAM, lngE) = Null
16440       arr_varCtl(C_LN_LFT, lngE) = Null
16450       arr_varCtl(C_LN_TOP, lngE) = Null
16460       arr_varCtl(C_F1_NAM, lngE) = Null
16470       arr_varCtl(C_F1_WDT, lngE) = Null
16480       arr_varCtl(C_F2_NAM, lngE) = Null
16490       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 7. lngShowAcctNum_Left, lngShowAcctNumBox_Top.
16500       lngCtls = lngCtls + 1&
16510       lngE = lngCtls - 1&
16520       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
16530       arr_varCtl(C_CNAM, lngE) = "chkShowAcctNum"
16540       arr_varCtl(C_TOP, lngE) = Null
16550       arr_varCtl(C_LFT, lngE) = .chkShowAcctNum.Left 'lngShowAcctNum_Left
16560       arr_varCtl(C_WDT, lngE) = Null
16570       arr_varCtl(C_HGT, lngE) = Null
16580       arr_varCtl(C_L1_NAM, lngE) = "chkShowAcctNum_lbl"
16590       arr_varCtl(C_L1_LFT, lngE) = .chkShowAcctNum_lbl.Left
16600       arr_varCtl(C_L2_NAM, lngE) = Null
16610       arr_varCtl(C_L2_LFT, lngE) = Null
16620       arr_varCtl(C_LN_NAM, lngE) = "chkShowAcctNum_box"
16630       arr_varCtl(C_LN_LFT, lngE) = Null
16640       arr_varCtl(C_LN_TOP, lngE) = .chkShowAcctNum_box.Top 'lngShowAcctNumBox_Top
16650       arr_varCtl(C_F1_NAM, lngE) = Null
16660       arr_varCtl(C_F1_WDT, lngE) = Null
16670       arr_varCtl(C_F2_NAM, lngE) = Null
16680       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 8. lngEnableCountry_Left, lngChkBoxLbl_Offset.
16690       lngCtls = lngCtls + 1&
16700       lngE = lngCtls - 1&
16710       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
16720       arr_varCtl(C_CNAM, lngE) = "chkEnableCountry_Wide"
16730       arr_varCtl(C_TOP, lngE) = Null
16740       arr_varCtl(C_LFT, lngE) = .chkEnableCountry_Wide.Left 'lngEnableCountry_Left
16750       arr_varCtl(C_WDT, lngE) = Null
16760       arr_varCtl(C_HGT, lngE) = Null
16770       arr_varCtl(C_L1_NAM, lngE) = "chkEnableCountry_Wide_lbl"
16780       arr_varCtl(C_L1_LFT, lngE) = .chkEnableCountry_Wide_lbl.Left
16790       arr_varCtl(C_L2_NAM, lngE) = "lngChkBoxLbl_Offset"
16800       arr_varCtl(C_L2_LFT, lngE) = (.chkEnableCountry_Wide_lbl.Left - .chkEnableCountry_Wide.Left) 'lngChkBoxLbl_Offset
16810       arr_varCtl(C_LN_NAM, lngE) = Null
16820       arr_varCtl(C_LN_LFT, lngE) = Null
16830       arr_varCtl(C_LN_TOP, lngE) = Null
16840       arr_varCtl(C_F1_NAM, lngE) = Null
16850       arr_varCtl(C_F1_WDT, lngE) = Null
16860       arr_varCtl(C_F2_NAM, lngE) = Null
16870       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 9. lngPageOf_Left.
16880       lngCtls = lngCtls + 1&
16890       lngE = lngCtls - 1&
16900       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
16910       arr_varCtl(C_CNAM, lngE) = "chkPageOf"
16920       arr_varCtl(C_TOP, lngE) = Null
16930       arr_varCtl(C_LFT, lngE) = .chkPageOf.Left 'lngPageOf_Left
16940       arr_varCtl(C_WDT, lngE) = Null
16950       arr_varCtl(C_HGT, lngE) = Null
16960       arr_varCtl(C_L1_NAM, lngE) = Null
16970       arr_varCtl(C_L1_LFT, lngE) = Null
16980       arr_varCtl(C_L2_NAM, lngE) = Null
16990       arr_varCtl(C_L2_LFT, lngE) = Null
17000       arr_varCtl(C_LN_NAM, lngE) = Null
17010       arr_varCtl(C_LN_LFT, lngE) = Null
17020       arr_varCtl(C_LN_TOP, lngE) = Null
17030       arr_varCtl(C_F1_NAM, lngE) = Null
17040       arr_varCtl(C_F1_WDT, lngE) = Null
17050       arr_varCtl(C_F2_NAM, lngE) = Null
17060       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 10. lngAlwaysUpdate_Left, lngAlwaysUpdateLbl_Left.
17070       lngCtls = lngCtls + 1&
17080       lngE = lngCtls - 1&
17090       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
17100       arr_varCtl(C_CNAM, lngE) = "chkAlwaysUpdate"
17110       arr_varCtl(C_TOP, lngE) = Null
17120       arr_varCtl(C_LFT, lngE) = .chkAlwaysUpdate.Left 'lngAlwaysUpdate_Left
17130       arr_varCtl(C_WDT, lngE) = Null
17140       arr_varCtl(C_HGT, lngE) = Null
17150       arr_varCtl(C_L1_NAM, lngE) = "chkAlwaysUpdate_lbl"
17160       arr_varCtl(C_L1_LFT, lngE) = .chkAlwaysUpdate_lbl.Left 'lngAlwaysUpdateLbl_Left
17170       arr_varCtl(C_L2_NAM, lngE) = "chkAlwaysUpdate_lbl2"
17180       arr_varCtl(C_L2_LFT, lngE) = .chkAlwaysUpdate_lbl2.Left
17190       arr_varCtl(C_LN_NAM, lngE) = Null
17200       arr_varCtl(C_LN_LFT, lngE) = Null
17210       arr_varCtl(C_LN_TOP, lngE) = Null
17220       arr_varCtl(C_F1_NAM, lngE) = Null
17230       arr_varCtl(C_F1_WDT, lngE) = Null
17240       arr_varCtl(C_F2_NAM, lngE) = Null
17250       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 11. lngShortcutLbl_Top
17260       lngCtls = lngCtls + 1&
17270       lngE = lngCtls - 1&
17280       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
17290       arr_varCtl(C_CNAM, lngE) = "ShortcutMenu_lbl"
17300       arr_varCtl(C_TOP, lngE) = .ShortcutMenu_lbl.Top 'lngShortcutLbl_Top
17310       arr_varCtl(C_LFT, lngE) = Null
17320       arr_varCtl(C_WDT, lngE) = Null
17330       arr_varCtl(C_HGT, lngE) = Null
17340       arr_varCtl(C_L1_NAM, lngE) = Null
17350       arr_varCtl(C_L1_LFT, lngE) = Null
17360       arr_varCtl(C_L2_NAM, lngE) = Null
17370       arr_varCtl(C_L2_LFT, lngE) = Null
17380       arr_varCtl(C_LN_NAM, lngE) = Null
17390       arr_varCtl(C_LN_LFT, lngE) = Null
17400       arr_varCtl(C_LN_TOP, lngE) = Null
17410       arr_varCtl(C_F1_NAM, lngE) = Null
17420       arr_varCtl(C_F1_WDT, lngE) = Null
17430       arr_varCtl(C_F2_NAM, lngE) = Null
17440       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 12. lngOpgViewBox2_Top.
17450       lngCtls = lngCtls + 1&
17460       lngE = lngCtls - 1&
17470       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
17480       arr_varCtl(C_CNAM, lngE) = "opgView_box2"
17490       arr_varCtl(C_TOP, lngE) = .opgView_box2.Top 'lngOpgViewBox2_Top
17500       arr_varCtl(C_LFT, lngE) = Null
17510       arr_varCtl(C_WDT, lngE) = Null
17520       arr_varCtl(C_HGT, lngE) = Null
17530       arr_varCtl(C_L1_NAM, lngE) = Null
17540       arr_varCtl(C_L1_LFT, lngE) = Null
17550       arr_varCtl(C_L2_NAM, lngE) = Null
17560       arr_varCtl(C_L2_LFT, lngE) = Null
17570       arr_varCtl(C_LN_NAM, lngE) = Null
17580       arr_varCtl(C_LN_LFT, lngE) = Null
17590       arr_varCtl(C_LN_TOP, lngE) = Null
17600       arr_varCtl(C_F1_NAM, lngE) = Null
17610       arr_varCtl(C_F1_WDT, lngE) = Null
17620       arr_varCtl(C_F2_NAM, lngE) = Null
17630       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 13. lngFormatBox_Top.
17640       lngCtls = lngCtls + 1&
17650       lngE = lngCtls - 1&
17660       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
17670       arr_varCtl(C_CNAM, lngE) = "opgFormat_box"
17680       arr_varCtl(C_TOP, lngE) = .opgFormat_box.Top 'lngFormatBox_Top
17690       arr_varCtl(C_LFT, lngE) = Null
17700       arr_varCtl(C_WDT, lngE) = Null
17710       arr_varCtl(C_HGT, lngE) = Null
17720       arr_varCtl(C_L1_NAM, lngE) = Null
17730       arr_varCtl(C_L1_LFT, lngE) = Null
17740       arr_varCtl(C_L2_NAM, lngE) = Null
17750       arr_varCtl(C_L2_LFT, lngE) = Null
17760       arr_varCtl(C_LN_NAM, lngE) = Null
17770       arr_varCtl(C_LN_LFT, lngE) = Null
17780       arr_varCtl(C_LN_TOP, lngE) = Null
17790       arr_varCtl(C_F1_NAM, lngE) = Null
17800       arr_varCtl(C_F1_WDT, lngE) = Null
17810       arr_varCtl(C_F2_NAM, lngE) = Null
17820       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 14. lngPrintBox_Top, lngPrintBox_Width, lngPrintBox_Margin.
17830       lngCtls = lngCtls + 1&
17840       lngE = lngCtls - 1&
17850       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
17860       arr_varCtl(C_CNAM, lngE) = "opgPrint_box"
17870       arr_varCtl(C_TOP, lngE) = .opgPrint_box.Top 'lngPrintBox_Top
17880       arr_varCtl(C_LFT, lngE) = Null
17890       arr_varCtl(C_WDT, lngE) = .opgPrint_box.Width 'lngPrintBox_Width
17900       arr_varCtl(C_HGT, lngE) = Null
17910       arr_varCtl(C_L1_NAM, lngE) = Null
17920       arr_varCtl(C_L1_LFT, lngE) = Null
17930       arr_varCtl(C_L2_NAM, lngE) = "lngPrintBox_Margin"
17940       arr_varCtl(C_L2_LFT, lngE) = ((.opgPrint_box.Left + .opgPrint_box.Width) - (.opgPrintEmail_box.Left + .opgPrintEmail_box.Width)) 'lngPrintBox_Margin
17950       arr_varCtl(C_LN_NAM, lngE) = Null
17960       arr_varCtl(C_LN_LFT, lngE) = Null
17970       arr_varCtl(C_LN_TOP, lngE) = Null
17980       arr_varCtl(C_F1_NAM, lngE) = Null
17990       arr_varCtl(C_F1_WDT, lngE) = Null
18000       arr_varCtl(C_F2_NAM, lngE) = Null
18010       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 15. lngAccountNo_Left.
18020       lngCtls = lngCtls + 1&
18030       lngE = lngCtls - 1&
18040       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
18050       arr_varCtl(C_CNAM, lngE) = "frmSub1.accountno"
18060       arr_varCtl(C_TOP, lngE) = Null
18070       arr_varCtl(C_LFT, lngE) = frmSub1.accountno.Left 'lngAccountNo_Left
18080       arr_varCtl(C_WDT, lngE) = Null
18090       arr_varCtl(C_HGT, lngE) = Null
18100       arr_varCtl(C_L1_NAM, lngE) = Null
18110       arr_varCtl(C_L1_LFT, lngE) = Null
18120       arr_varCtl(C_L2_NAM, lngE) = Null
18130       arr_varCtl(C_L2_LFT, lngE) = Null
18140       arr_varCtl(C_LN_NAM, lngE) = Null
18150       arr_varCtl(C_LN_LFT, lngE) = Null
18160       arr_varCtl(C_LN_TOP, lngE) = Null
18170       arr_varCtl(C_F1_NAM, lngE) = Null
18180       arr_varCtl(C_F1_WDT, lngE) = Null
18190       arr_varCtl(C_F2_NAM, lngE) = Null
18200       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 16. lngShortName_Left, lngShortName_Width, lngShortName_Diff.
18210       lngCtls = lngCtls + 1&
18220       lngE = lngCtls - 1&
18230       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
18240       arr_varCtl(C_CNAM, lngE) = "frmSub1.shortname"
18250       arr_varCtl(C_TOP, lngE) = Null
18260       arr_varCtl(C_LFT, lngE) = frmSub1.shortname.Left 'lngShortName_Left
18270       arr_varCtl(C_WDT, lngE) = frmSub1.shortname.Width 'lngShortName_Width
18280       arr_varCtl(C_HGT, lngE) = Null
18290       arr_varCtl(C_L1_NAM, lngE) = "lngShortName_Diff"
18300       arr_varCtl(C_L1_LFT, lngE) = (frmSub1.shortname.Width - frmSub1.accountno.Width) 'lngShortName_Diff
18310       arr_varCtl(C_L2_NAM, lngE) = Null
18320       arr_varCtl(C_L2_LFT, lngE) = Null
18330       arr_varCtl(C_LN_NAM, lngE) = Null
18340       arr_varCtl(C_LN_LFT, lngE) = Null
18350       arr_varCtl(C_LN_TOP, lngE) = Null
18360       arr_varCtl(C_F1_NAM, lngE) = Null
18370       arr_varCtl(C_F1_WDT, lngE) = Null
18380       arr_varCtl(C_F2_NAM, lngE) = Null
18390       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 17. lngConNum_Left.
18400       lngCtls = lngCtls + 1&
18410       lngE = lngCtls - 1&
18420       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
18430       arr_varCtl(C_CNAM, lngE) = "frmSub1.Contact_Number"
18440       arr_varCtl(C_TOP, lngE) = Null
18450       arr_varCtl(C_LFT, lngE) = frmSub1.Contact_Number.Left 'lngConNum_Left
18460       arr_varCtl(C_WDT, lngE) = Null
18470       arr_varCtl(C_HGT, lngE) = Null
18480       arr_varCtl(C_L1_NAM, lngE) = Null
18490       arr_varCtl(C_L1_LFT, lngE) = Null
18500       arr_varCtl(C_L2_NAM, lngE) = Null
18510       arr_varCtl(C_L2_LFT, lngE) = Null
18520       arr_varCtl(C_LN_NAM, lngE) = Null
18530       arr_varCtl(C_LN_LFT, lngE) = Null
18540       arr_varCtl(C_LN_TOP, lngE) = Null
18550       arr_varCtl(C_F1_NAM, lngE) = Null
18560       arr_varCtl(C_F1_WDT, lngE) = Null
18570       arr_varCtl(C_F2_NAM, lngE) = Null
18580       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 18. lngConName_Left.
18590       lngCtls = lngCtls + 1&
18600       lngE = lngCtls - 1&
18610       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
18620       arr_varCtl(C_CNAM, lngE) = "frmSub1.Contact_Name"
18630       arr_varCtl(C_TOP, lngE) = Null
18640       arr_varCtl(C_LFT, lngE) = frmSub1.Contact_Name.Left 'lngConName_Left
18650       arr_varCtl(C_WDT, lngE) = Null
18660       arr_varCtl(C_HGT, lngE) = Null
18670       arr_varCtl(C_L1_NAM, lngE) = Null
18680       arr_varCtl(C_L1_LFT, lngE) = Null
18690       arr_varCtl(C_L2_NAM, lngE) = Null
18700       arr_varCtl(C_L2_LFT, lngE) = Null
18710       arr_varCtl(C_LN_NAM, lngE) = Null
18720       arr_varCtl(C_LN_LFT, lngE) = Null
18730       arr_varCtl(C_LN_TOP, lngE) = Null
18740       arr_varCtl(C_F1_NAM, lngE) = Null
18750       arr_varCtl(C_F1_WDT, lngE) = Null
18760       arr_varCtl(C_F2_NAM, lngE) = Null
18770       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 19. lngLocContact_Width, lngLocContact_Diff.
18780       lngCtls = lngCtls + 1&
18790       lngE = lngCtls - 1&
18800       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
18810       arr_varCtl(C_CNAM, lngE) = "frmSub2.opgLocContact"
18820       arr_varCtl(C_TOP, lngE) = Null
18830       arr_varCtl(C_LFT, lngE) = Null
18840       arr_varCtl(C_WDT, lngE) = (frmSub2.Contact_City.Left - frmSub2.opgLocContact.Left) 'lngLocContact_Width
18850       arr_varCtl(C_HGT, lngE) = Null
18860       arr_varCtl(C_L1_NAM, lngE) = "lngLocContact_Diff"
18870       arr_varCtl(C_L1_LFT, lngE) = (frmSub1.Contact_State.Left - frmSub1.opgLocContact.Left) 'lngLocContact_Diff
18880       arr_varCtl(C_L2_NAM, lngE) = "frmSub1.opgLocContact"
18890       arr_varCtl(C_L2_LFT, lngE) = Null
18900       arr_varCtl(C_LN_NAM, lngE) = Null
18910       arr_varCtl(C_LN_LFT, lngE) = Null
18920       arr_varCtl(C_LN_TOP, lngE) = Null
18930       arr_varCtl(C_F1_NAM, lngE) = Null
18940       arr_varCtl(C_F1_WDT, lngE) = Null
18950       arr_varCtl(C_F2_NAM, lngE) = Null
18960       arr_varCtl(C_F2_WDT, lngE) = Null

            ' ** 20. lngCountry_Width, lngCountry_Diff.
18970       lngCtls = lngCtls + 1&
18980       lngE = lngCtls - 1&
18990       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
19000       arr_varCtl(C_CNAM, lngE) = "frmSub2.Contact_Country"
19010       arr_varCtl(C_TOP, lngE) = Null
19020       arr_varCtl(C_LFT, lngE) = Null
19030       arr_varCtl(C_WDT, lngE) = (frmSub2.Contact_Phone1.Left - frmSub2.Contact_Country.Left) 'lngCountry_Width
19040       arr_varCtl(C_HGT, lngE) = Null
19050       arr_varCtl(C_L1_NAM, lngE) = "lngCountry_Diff"
19060       arr_varCtl(C_L1_LFT, lngE) = (frmSub1.Contact_Phone1.Left - frmSub1.Contact_Country.Left) 'lngCountry_Diff
19070       arr_varCtl(C_L2_NAM, lngE) = "frmSub1.Contact_Country"
19080       arr_varCtl(C_L2_LFT, lngE) = Null
19090       arr_varCtl(C_LN_NAM, lngE) = Null
19100       arr_varCtl(C_LN_LFT, lngE) = Null
19110       arr_varCtl(C_LN_TOP, lngE) = Null
19120       arr_varCtl(C_F1_NAM, lngE) = Null
19130       arr_varCtl(C_F1_WDT, lngE) = Null
19140       arr_varCtl(C_F2_NAM, lngE) = Null
19150       arr_varCtl(C_F2_WDT, lngE) = Null

19160       .CtlArraySet arr_varCtl  ' ** Form Procedure: frmAccountContacts.

19170     End If

19180   End With

EXITP:
19190   Exit Sub

ERRH:
19200   Select Case ERR.Number
        Case Else
19210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19220   End Select
19230   Resume EXITP

End Sub

Public Sub EnableCountry_SetFrmWidth_AC(lngFrm_Top_Orig As Long, lngFrm_Left_Orig As Long, lngFrm_Width_Orig As Long, lngFrm_Height_Orig As Long, lngForm_Width As Long, frm As Access.Form)
' ** Form opens at the accountno-only width, with country.

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "EnableCountry_SetFrmWidth_AC"

        Dim lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
        Dim blnAcctNo As Boolean, blnShortName As Boolean
        Dim lngNewForm_Width As Long, lngNewFrm_Width As Long, lngCtlSep As Long, lngOffset As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long, lngTmp05 As Long

19310   With frm

19320     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
19330       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
19340     End If

19350     blnAcctNo = .chkShowAcctNum
19360     blnShortName = .chkShowShortName
19370     lngCtlSep = (4& * lngTpp)
19380     lngOffset = (.chkEnableCountry_alt_line.Left - arr_varCtl(C_LFT, 8))

19390     GetFormDimensions frm, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

19400     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
19410     lngMonitorNum = 1&: lngTmp05 = 0&
19420     EnumMonitors frm  ' ** Module Function: modMonitorFuncs.
19430     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
19440     Select Case .opgView
          Case .opgView_optWide.OptionValue

19450       Select Case .chkEnableCountry_Wide
            Case True
              ' ** For Wide, with country.

19460         If blnAcctNo = True And blnShortName = True Then
19470           lngNewFrm_Width = ((lngFrm_Width_Orig + arr_varCtl(C_WDT, 16) + lngCtlSep) + (arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))) + lngTpp
19480           lngNewForm_Width = ((lngForm_Width + arr_varCtl(C_WDT, 16) + lngCtlSep) + (arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19)))
19490         Else
19500           If blnAcctNo = True Then
19510             lngNewFrm_Width = ((lngFrm_Width_Orig) + (arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))) + lngTpp
19520             lngNewForm_Width = ((lngForm_Width) + (arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19)))
19530           ElseIf blnShortName = True Then
19540             lngNewFrm_Width = ((lngFrm_Width_Orig + arr_varCtl(C_L1_LFT, 16)) + (arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))) + lngTpp
19550             lngNewForm_Width = ((lngForm_Width + arr_varCtl(C_L1_LFT, 16)) + (arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19)))
19560           End If
19570         End If

19580         If lngFrm_Width <> lngNewFrm_Width Then

19590           .Width = lngNewForm_Width
19600           lngTmp01 = lngFrm_Left_Orig - ((lngNewFrm_Width - lngFrm_Width_Orig) / 2&)

19610           If lngMonitorNum = 1& Then lngTmp05 = lngFrm_Top_Orig
19620           DoCmd.SelectObject acForm, frm.Name, False
19630           DoEvents
19640           DoCmd.MoveSize lngTmp01, lngTmp05, lngNewFrm_Width, lngFrm_Height_Orig  'lngFrm_Top_Orig
19650           If lngMonitorNum > 1& Then
19660             LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
19670           End If

19680           lngTmp02 = (lngNewForm_Width - lngForm_Width)
19690           .frmAccountContacts_Sub.Width = (arr_varCtl(C_F1_WDT, 0) + lngTmp02) + lngTpp
19700           .frmAccountContacts_Sub_box.Width = (.frmAccountContacts_Sub.Width + (2& * lngTpp))
19710           .Nav_Sub1_box01.Width = .frmAccountContacts_Sub.Width

19720           .cmdClose.Left = (arr_varCtl(C_LFT, 2) + lngTmp02) + lngTpp
19730           .chkAlwaysUpdate_lbl.Left = (arr_varCtl(C_L1_LFT, 10) + lngTmp02) + lngTpp
19740           .chkAlwaysUpdate_lbl2.Left = .chkAlwaysUpdate_lbl.Left
19750           .chkAlwaysUpdate_lbl2_dim_hi.Left = (.chkAlwaysUpdate_lbl2.Left + lngTpp)
19760           .chkAlwaysUpdate.Left = (arr_varCtl(C_LFT, 10) + lngTmp02) + lngTpp
19770           .cmdUpdate.Left = (arr_varCtl(C_LFT, 5) + lngTmp02) + lngTpp
19780           .cmdDelete.Left = (arr_varCtl(C_LFT, 4) + lngTmp02) + lngTpp
19790           .cmdAdd.Left = (arr_varCtl(C_LFT, 3) + lngTmp02) + lngTpp
19800           .Header_vline01.Left = lngNewForm_Width + lngTpp
19810           .Header_vline02.Left = lngNewForm_Width + lngTpp
19820           .Detail_vline01.Left = lngNewForm_Width + lngTpp
19830           .Detail_vline02.Left = lngNewForm_Width + lngTpp
19840           .Detail_vline03.Left = lngNewForm_Width + lngTpp
19850           .Detail_vline04.Left = lngNewForm_Width + lngTpp
19860           .Footer_vline01.Left = lngNewForm_Width + lngTpp
19870           .Footer_vline02.Left = lngNewForm_Width + lngTpp
19880           .Header_hline01.Width = lngNewForm_Width + lngTpp
19890           .Header_hline02.Width = lngNewForm_Width + lngTpp
19900           .Detail_hline01.Width = lngNewForm_Width + lngTpp
19910           .Detail_hline02.Width = lngNewForm_Width + lngTpp
19920           .Detail_hline03.Width = lngNewForm_Width + lngTpp
19930           .Detail_hline04.Width = lngNewForm_Width + lngTpp
19940           .Footer_hline01.Width = lngNewForm_Width + lngTpp
19950           .Footer_hline02.Width = lngNewForm_Width + lngTpp

19960           .Width = lngNewForm_Width

19970         End If

19980       Case False
              ' ** Wide, without country.

19990         If blnAcctNo = True And blnShortName = True Then
20000           lngNewFrm_Width = (lngFrm_Width_Orig + arr_varCtl(C_WDT, 16) + lngCtlSep)
20010           lngNewForm_Width = (lngForm_Width + arr_varCtl(C_WDT, 16) + lngCtlSep)
20020         Else
20030           If blnAcctNo = True Then
20040             lngNewFrm_Width = lngFrm_Width_Orig
20050             lngNewForm_Width = lngForm_Width
20060           ElseIf blnShortName = True Then
20070             lngNewFrm_Width = (lngFrm_Width_Orig + arr_varCtl(C_L1_LFT, 16))
20080             lngNewForm_Width = (lngForm_Width + arr_varCtl(C_L1_LFT, 16))
20090           End If
20100         End If

20110         If lngFrm_Width <> lngNewFrm_Width Then
                ' ** So this is skipped on opening.

20120           .Width = lngNewForm_Width
20130           lngTmp01 = lngFrm_Left_Orig - ((lngNewFrm_Width - lngFrm_Width_Orig) / 2&)
20140           lngTmp02 = (lngForm_Width - lngNewForm_Width)
20150           .frmAccountContacts_Sub.Width = (arr_varCtl(C_F1_WDT, 0) - lngTmp02)
20160           .frmAccountContacts_Sub_box.Width = (.frmAccountContacts_Sub.Width + (2& * lngTpp))
20170           .Nav_Sub1_box01.Width = .frmAccountContacts_Sub.Width

20180           .cmdClose.Left = (arr_varCtl(C_LFT, 2) - lngTmp02)
20190           .cmdAdd.Left = (arr_varCtl(C_LFT, 3) - lngTmp02)
20200           .cmdDelete.Left = (arr_varCtl(C_LFT, 4) - lngTmp02)
20210           .cmdUpdate.Left = (arr_varCtl(C_LFT, 5) - lngTmp02)
20220           .chkAlwaysUpdate.Left = (arr_varCtl(C_LFT, 10) - lngTmp02)
20230           .chkAlwaysUpdate_lbl.Left = (arr_varCtl(C_L1_LFT, 10) - lngTmp02)
20240           .chkAlwaysUpdate_lbl2.Left = .chkAlwaysUpdate_lbl.Left
20250           .chkAlwaysUpdate_lbl2_dim_hi.Left = (.chkAlwaysUpdate_lbl2.Left + lngTpp)
20260           .Header_vline01.Left = lngNewForm_Width
20270           .Header_vline02.Left = lngNewForm_Width
20280           .Detail_vline01.Left = lngNewForm_Width
20290           .Detail_vline02.Left = lngNewForm_Width
20300           .Detail_vline03.Left = lngNewForm_Width
20310           .Detail_vline04.Left = lngNewForm_Width
20320           .Footer_vline01.Left = lngNewForm_Width
20330           .Footer_vline02.Left = lngNewForm_Width
20340           .Header_hline01.Width = lngNewForm_Width
20350           .Header_hline02.Width = lngNewForm_Width
20360           .Detail_hline01.Width = lngNewForm_Width
20370           .Detail_hline02.Width = lngNewForm_Width
20380           .Detail_hline03.Width = lngNewForm_Width
20390           .Detail_hline04.Width = lngNewForm_Width
20400           .Footer_hline01.Width = lngNewForm_Width
20410           .Footer_hline02.Width = lngNewForm_Width

20420           .Width = lngNewForm_Width

20430           If lngMonitorNum = 1& Then lngTmp05 = lngFrm_Top_Orig
20440           DoCmd.SelectObject acForm, frm.Name, False
20450           DoEvents
20460           DoCmd.MoveSize lngTmp01, lngTmp05, lngNewFrm_Width, lngFrm_Height_Orig  'lngFrm_Top_Orig
20470           If lngMonitorNum > 1& Then
20480             LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
20490           End If

20500         End If

20510       End Select

20520     Case .opgView_optCompact.OptionValue

20530       Select Case .chkEnableCountry_Compact
            Case True
              ' ** For Compact, design View set with Country.
20540         lngNewFrm_Width = (lngFrm_Width_Orig - arr_varCtl(C_L1_LFT, 0))
              ' ** Don't adjust form Width when working with Compact.
20550         If lngFrm_Width <> lngNewFrm_Width Then
20560           lngTmp01 = (arr_varCtl(C_F1_WDT, 0) - arr_varCtl(C_L1_LFT, 0))
20570           If .frmAccountContacts_Sub2.Width <> lngTmp01 Then
20580             .frmAccountContacts_Sub2.Width = lngTmp01
20590             .frmAccountContacts_Sub2_box.Width = (.frmAccountContacts_Sub2.Width + (2& * lngTpp))
20600           End If
20610           .cmdClose.Left = (arr_varCtl(C_LFT, 2) - arr_varCtl(C_L1_LFT, 0))
20620           .cmdAdd.Left = (arr_varCtl(C_LFT, 3) - arr_varCtl(C_L1_LFT, 0))
20630           .cmdDelete.Left = (arr_varCtl(C_LFT, 4) - arr_varCtl(C_L1_LFT, 0))
20640           .cmdUpdate.Left = (arr_varCtl(C_LFT, 5) - arr_varCtl(C_L1_LFT, 0))
20650           .chkAlwaysUpdate.Left = (arr_varCtl(C_LFT, 10) - arr_varCtl(C_L1_LFT, 0))
20660           .chkAlwaysUpdate_lbl.Left = (arr_varCtl(C_L1_LFT, 10) - arr_varCtl(C_L1_LFT, 0))
20670           .chkAlwaysUpdate_lbl2.Left = .chkAlwaysUpdate_lbl.Left
20680           .chkAlwaysUpdate_lbl2_dim_hi.Left = (.chkAlwaysUpdate_lbl2.Left + lngTpp)
20690           If lngNewFrm_Width > lngFrm_Width_Orig Then
20700             lngTmp01 = ((lngNewFrm_Width - lngFrm_Width_Orig) / 2)
20710             lngTmp01 = (lngFrm_Left_Orig - lngTmp01)
20720           Else
20730             lngTmp01 = ((lngFrm_Width_Orig - lngNewFrm_Width) / 2)
20740             lngTmp01 = (lngFrm_Left_Orig + lngTmp01)
20750           End If
20760           If lngMonitorNum = 1& Then lngTmp05 = lngFrm_Top
20770           DoCmd.SelectObject acForm, frm.Name, False
20780           DoEvents
20790           DoCmd.MoveSize lngTmp01, lngTmp05, lngNewFrm_Width, lngFrm_Height  'lngFrm_Top
20800           If lngMonitorNum > 1& Then
20810             LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
20820           End If
20830         End If

20840       Case False
20850         lngNewFrm_Width = (((lngFrm_Width_Orig - arr_varCtl(C_L1_LFT, 0)) - (arr_varCtl(C_WDT, 19) + arr_varCtl(C_WDT, 20))) - lngTpp)
20860         If lngFrm_Width <> lngNewFrm_Width Then
20870           lngTmp01 = (((arr_varCtl(C_F1_WDT, 0) - arr_varCtl(C_L1_LFT, 0)) - (arr_varCtl(C_WDT, 19) + arr_varCtl(C_WDT, 20))) - lngTpp)
20880           If .frmAccountContacts_Sub2.Width <> lngTmp01 Then
20890             .frmAccountContacts_Sub2.Width = lngTmp01
20900             .frmAccountContacts_Sub2_box.Width = (.frmAccountContacts_Sub2.Width + (2& * lngTpp))
20910           End If
20920           lngTmp01 = (.chkAlwaysUpdate_lbl.Left - .cmdClose.Left)
20930           lngTmp02 = (.cmdClose.Left - .chkAlwaysUpdate.Left)
20940           lngTmp03 = (.cmdClose.Left - .cmdUpdate.Left)
20950           lngTmp04 = (.cmdUpdate.Left - .cmdDelete.Left)
20960           .cmdClose.Left = ((.frmAccountContacts_Sub2.Left + .frmAccountContacts_Sub2.Width) - arr_varCtl(C_L2_LFT, 0))
20970           .cmdUpdate.Left = (.cmdClose.Left - lngTmp03)
20980           .cmdDelete.Left = (.cmdUpdate.Left - lngTmp04)
20990           .cmdAdd.Left = (.cmdDelete.Left - lngTmp04)
21000           .chkAlwaysUpdate.Left = (.cmdClose.Left - lngTmp02)
21010           .chkAlwaysUpdate_lbl.Left = (.cmdClose.Left + lngTmp01)
21020           .chkAlwaysUpdate_lbl2.Left = .chkAlwaysUpdate_lbl.Left
21030           .chkAlwaysUpdate_lbl2_dim_hi.Left = (.chkAlwaysUpdate_lbl2.Left + lngTpp)
21040           lngTmp01 = ((lngFrm_Width_Orig - lngNewFrm_Width) / 2)
21050           lngTmp01 = (lngFrm_Left_Orig + lngTmp01)
21060           If lngMonitorNum = 1& Then lngTmp05 = lngFrm_Top
21070           DoCmd.SelectObject acForm, frm.Name, False
21080           DoEvents
21090           DoCmd.MoveSize lngTmp01, lngTmp05, lngNewFrm_Width, lngFrm_Height  'lngFrm_Top
21100           If lngMonitorNum > 1& Then
21110             LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
21120           End If
21130         End If

21140       End Select

21150     End Select

21160   End With

EXITP:
21170   Exit Sub

ERRH:
21180   Select Case ERR.Number
        Case Else
21190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21200   End Select
21210   Resume EXITP

End Sub

Public Sub AcctNoShort_Set(strProc As String, blnShow As Boolean, lngFrm_Top_Orig As Long, lngFrm_Left_Orig As Long, lngFrm_Width_Orig As Long, lngFrm_Height_Orig As Long, lngForm_Width As Long, frm As Access.Form)
' ** Form opens at the accountno-only width.
' ** These adjustments only apply to the Wide view.

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "AcctNoShort_Set"

        Dim strOption As String, strSortNow As String
        Dim blnAcctNo As Boolean, blnShortName As Boolean, blnEnableCountry As Boolean
        Dim lngCurFormSub_Width As Long, lngNewFormSub_Width As Long, lngSortLbl_Width As Long
        Dim lngCtlSep As Long, lngNewForm_Width As Long, lngCurWidth As Long
        Dim strCtlName As String
        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String, lngTmp02 As Long, lngTmp03 As Long, intTmp04 As Integer
        Dim intX As Integer

21310   With frm

21320     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
21330       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
21340     End If

21350     lngCtlSep = (4& * lngTpp)
21360     lngCurWidth = .Width
21370     blnSortHere = False: blnResort = False
21380     strCtlName = vbNullString

21390     With .frmAccountContacts_Sub.Form
21400       lngRecsCur = .RecCnt  ' ** Form Function: frmAccountContacts_Sub.
21410       lngCurFormSub_Width = .Width
21420       strSortNow = .SortNow_Get  ' ** Form Function: frmAccountContacts_Sub.
21430       lngSortLbl_Width = .Sort_lbl.Width
21440       .Sort_line.Visible = False
21450       .Sort_lbl.Visible = False
21460     End With

21470     intPos01 = InStr(strProc, "_AfterUpdate")
21480     strOption = Left(strProc, (intPos01 - 1))
21490     Select Case strOption
          Case "chkShowAcctNum"
21500       Select Case blnShow
            Case True
21510         blnAcctNo = True
21520         blnShortName = .chkShowShortName
21530       Case False
21540         blnAcctNo = False
21550         blnShortName = .chkShowShortName
21560         If blnShortName = False Then
21570           blnShortName = True
21580           .chkShowShortName = True
21590           .chkShowShortName_lbl.FontBold = True
21600         End If
21610       End Select
21620     Case "chkShowShortName"
21630       Select Case blnShow
            Case True
21640         blnShortName = True
21650         blnAcctNo = .chkShowAcctNum
21660       Case False
21670         blnShortName = False
21680         blnAcctNo = .chkShowAcctNum
21690         If blnAcctNo = False Then
21700           blnAcctNo = True
21710           .chkShowAcctNum = True
21720           .chkShowAcctNum_lbl.FontBold = True
21730         End If
21740       End Select
21750     End Select

21760     blnEnableCountry = .chkEnableCountry_Wide

21770     If blnAcctNo = True And blnShortName = True Then
            ' ** This will be the window at its widest.
21780       With .frmAccountContacts_Sub.Form

              ' ** lngFormSub_Width is accountno only, with Country.
21790         Select Case blnEnableCountry
              Case True
21800           lngNewFormSub_Width = (arr_varCtl(C_F2_WDT, 0) + arr_varCtl(C_WDT, 16) + lngCtlSep)
21810         Case False
21820           lngNewFormSub_Width = ((arr_varCtl(C_F2_WDT, 0) + arr_varCtl(C_WDT, 16) + lngCtlSep) - (arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))) - lngTpp
21830         End Select
21840         .Width = lngNewFormSub_Width

21850         .shortname.Left = arr_varCtl(C_LFT, 16)

21860         If .Sort_line.Left = .Contact_Name_lbl.Left Then blnSortHere = True
21870         .Contact_Number.Left = ((.shortname.Left + .shortname.Width) + (2& * lngTpp))
21880         .Contact_Name.Left = ((.Contact_Number.Left + .Contact_Number.Width) + lngCtlSep)
21890         .Contact_Name_LastFirst.Left = .Contact_Name.Left
21900         If blnSortHere = True Then
21910           strCtlName = "Contact_Name_lbl"
21920         End If
21930         blnSortHere = False
21940         If .Sort_line.Left = .Contact_Address1_lbl.Left Then blnSortHere = True
21950         .Contact_Address1.Left = ((.Contact_Name.Left + .Contact_Name.Width) + lngCtlSep)
21960         If blnSortHere = True Then
21970           strCtlName = "Contact_Address1_lbl"
21980         End If
21990         blnSortHere = False
22000         If .Sort_line.Left = .Contact_Address2_lbl.Left Then blnSortHere = True
22010         .Contact_Address2.Left = ((.Contact_Address1.Left + .Contact_Address1.Width) + lngCtlSep)
22020         If blnSortHere = True Then
22030           strCtlName = "Contact_Address2_lbl"
22040         End If
22050         blnSortHere = False
22060         If .Sort_line.Left = .Contact_City_lbl.Left Then blnSortHere = True
22070         .Contact_City.Left = ((.Contact_Address2.Left + .Contact_Address2.Width) + lngCtlSep)
22080         If blnSortHere = True Then
22090           strCtlName = "Contact_City_lbl"
22100         End If
22110         blnSortHere = False

22120         Select Case blnEnableCountry
              Case True
22130           If .Sort_line.Left = .opgLocContact_lbl.Left Then blnSortHere = True
22140           If ((.Contact_City.Left + .Contact_City.Width) + lngCtlSep) > .opgLocContact.Left Then
22150             lngTmp03 = (((.Contact_City.Left + .Contact_City.Width) + lngCtlSep) - .opgLocContact.Left)
22160             intTmp04 = (lngTmp03 / lngTpp)
22170             For intX = 1 To intTmp04
22180               .opgLocContact.Left = (.opgLocContact.Left + lngTpp)
22190               .opgLocContact_optOther_lbl.Left = (.opgLocContact_optOther_lbl.Left + lngTpp)
22200               .opgLocContact_optOther.Left = (.opgLocContact_optOther.Left + lngTpp)
22210               .opgLocContact_optUSA_lbl.Left = (.opgLocContact_optUSA_lbl.Left + lngTpp)
22220               .opgLocContact_optUSA.Left = (.opgLocContact_optUSA.Left + lngTpp)
22230             Next  ' ** intX.
22240           Else
22250             lngTmp03 = (.opgLocContact.Left - ((.Contact_City.Left + .Contact_City.Width) + lngCtlSep))
22260             intTmp04 = (lngTmp03 / lngTpp)
22270             For intX = 1 To intTmp04
22280               .opgLocContact.Left = (.opgLocContact.Left - lngTpp)
22290               .opgLocContact_optOther_lbl.Left = (.opgLocContact_optOther_lbl.Left - lngTpp)
22300               .opgLocContact_optOther.Left = (.opgLocContact_optOther.Left - lngTpp)
22310               .opgLocContact_optUSA_lbl.Left = (.opgLocContact_optUSA_lbl.Left - lngTpp)
22320               .opgLocContact_optUSA.Left = (.opgLocContact_optUSA.Left - lngTpp)
22330             Next  ' ** intX.
22340           End If
22350           .opgLocContact_box.Left = .opgLocContact.Left
22360           .opgLocContact_optOther_box.Left = (.opgLocContact_optOther.Left - (5& * lngTpp))
22370           .opgLocContact_optOther_lbl_on.Left = .opgLocContact_optOther_lbl.Left
22380           .opgLocContact_optOther_lbl_off.Left = .opgLocContact_optOther_lbl.Left
22390           .opgLocContact_optUSA_box.Left = (.opgLocContact_optUSA.Left - (5& * lngTpp))
22400           .opgLocContact_optUSA_lbl_on.Left = .opgLocContact_optUSA_lbl.Left
22410           .opgLocContact_optUSA_lbl_off.Left = .opgLocContact_optUSA_lbl.Left
22420           .opgLocContact_lbl.Left = (.opgLocContact.Left - lngTpp)
22430           .opgLocContact_lbl_dim_hi.Left = (.opgLocContact_lbl.Left + lngTpp)
22440           .opgLocContact_lbl_line.Left = .opgLocContact_lbl.Left
22450           .opgLocContact_lbl_line_dim_hi.Left = (.opgLocContact_lbl_line.Left + lngTpp)
22460         Case False
                ' ** I guess we don't need to move anything.
22470         End Select
22480         blnSortHere = False

22490         If .Sort_line.Left = .Contact_State_lbl.Left Then blnSortHere = True
22500         Select Case blnEnableCountry
              Case True
22510           .Contact_State.Left = ((.opgLocContact.Left + .opgLocContact.Width) + lngCtlSep)
22520         Case False
22530           .Contact_State.Left = ((.Contact_City.Left + .Contact_City.Width) + lngCtlSep)
22540         End Select
22550         .Contact_State_box.Left = (.Contact_State.Left - lngTpp)
22560         If blnSortHere = True Then
22570           strCtlName = "Contact_State_lbl"
22580         End If
22590         blnSortHere = False
22600         If .Sort_line.Left = .Contact_Zip_lbl.Left Then blnSortHere = True
22610         .Contact_Zip.Left = ((.Contact_State.Left + .Contact_State.Width) + lngCtlSep)
22620         .Contact_Zip_Format.Left = .Contact_Zip.Left
22630         .Contact_Zip_box.Left = (.Contact_Zip.Left - lngTpp)
22640         If blnSortHere = True Then
22650           strCtlName = "Contact_Zip_lbl"
22660         End If
22670         blnSortHere = False

22680         Select Case blnEnableCountry
              Case True
22690           If .Sort_line.Left = .Contact_Country_lbl.Left Then blnSortHere = True
22700           .Contact_Country.Left = ((.Contact_Zip.Left + .Contact_Zip.Width) + lngCtlSep)
22710           .Contact_Country_box.Left = (.Contact_Country.Left - lngTpp)
22720           If blnSortHere = True Then
22730             strCtlName = "Contact_Country_lbl"
22740           End If
22750           blnSortHere = False
22760           If .Sort_line.Left = .Contact_PostalCode_lbl.Left Then blnSortHere = True
22770           .Contact_PostalCode.Left = ((.Contact_Country.Left + .Contact_Country.Width) + lngCtlSep)
22780           .Contact_PostalCode_box.Left = (.Contact_PostalCode.Left - lngTpp)
22790           If blnSortHere = True Then
22800             strCtlName = "Contact_PostalCode_lbl"
22810           End If
22820           blnSortHere = False
22830           If .Sort_line.Left = .Contact_Phone1_lbl.Left Then blnSortHere = True
22840           .Contact_Phone1.Left = ((.Contact_PostalCode.Left + .Contact_PostalCode.Width) + lngCtlSep)
22850           .Contact_Phone1_Format.Left = .Contact_Phone1.Left
22860           If blnSortHere = True Then
22870             strCtlName = "Contact_Phone1_lbl"
22880           End If
22890           blnSortHere = False
22900         Case False
22910           If .Sort_line.Left = .Contact_Phone1_lbl.Left Then blnSortHere = True
22920           .Contact_Phone1.Left = ((.Contact_Zip.Left + .Contact_Zip.Width) + lngCtlSep)
22930           .Contact_Phone1_Format.Left = .Contact_Phone1.Left
22940           If blnSortHere = True Then
22950             strCtlName = "Contact_Phone1_lbl"
22960           End If
22970           blnSortHere = False
22980         End Select

22990         If .Sort_line.Left = .Contact_Phone2_lbl.Left Then blnSortHere = True
23000         .Contact_Phone2.Left = ((.Contact_Phone1.Left + .Contact_Phone1.Width) + lngCtlSep)
23010         .Contact_Phone2_Format.Left = .Contact_Phone2.Left
23020         If blnSortHere = True Then
23030           strCtlName = "Contact_Phone2_lbl"
23040         End If
23050         blnSortHere = False
23060         If .Sort_line.Left = .Contact_Fax_lbl.Left Then blnSortHere = True
23070         .Contact_Fax.Left = ((.Contact_Phone2.Left + .Contact_Phone2.Width) + lngCtlSep)
23080         .Contact_Fax_Format.Left = .Contact_Fax.Left
23090         If blnSortHere = True Then
23100           strCtlName = "Contact_Fax_lbl"
23110         End If
23120         blnSortHere = False
23130         If .Sort_line.Left = .Contact_Email_lbl.Left Then blnSortHere = True
23140         .Contact_Email.Left = ((.Contact_Fax.Left + .Contact_Fax.Width) + lngCtlSep)
23150         .Contact_Email_cmd.Left = (.Contact_Email.Left - (3& * lngTpp))
23160         If blnSortHere = True Then
23170           strCtlName = "Contact_Email_lbl"
23180         End If
23190         blnSortHere = False

              ' ** If it was on with shortname only showing, leave it there.
23200         intPos01 = InStr(strSortNow, " ")
23210         If intPos01 > 0 Then strTmp01 = Left(strSortNow, intPos01) Else strTmp01 = strSortNow
23220         intPos01 = InStr(strTmp01, "shortname")
23230         If .Sort_line.Left = .shortname_lbl.Left And intPos01 > 0 Then blnSortHere = True
23240         .shortname_lbl.Left = .shortname.Left
23250         .shortname_lbl_dim_hi.Left = (.shortname_lbl.Left + lngTpp)
23260         .shortname_lbl_line.Left = .shortname.Left
23270         .shortname_lbl_Line_dim_hi.Left = (.shortname_lbl_line.Left + lngTpp)
23280         If blnSortHere = True Then
23290           .Sort_lbl.Left = ((.shortname_lbl.Left + .shortname_lbl.Width) - lngSortLbl_Width)
23300           .Sort_line.Left = .shortname_lbl.Left
23310         End If
23320         blnSortHere = False

              ' ** If it was on with accountno only showing, leave it there.
23330         intPos01 = InStr(strSortNow, " ")
23340         If intPos01 > 0 Then strTmp01 = Left(strSortNow, intPos01) Else strTmp01 = strSortNow
23350         intPos01 = InStr(strTmp01, "accountno")
23360         If .Sort_line.Left = .accountno_lbl.Left And intPos01 > 0 Then blnSortHere = True
23370         .accountno_lbl.Width = .accountno.Width
23380         .accountno_lbl_dim_hi.Width = .accountno_lbl.Width
23390         .accountno_lbl_line.Width = (.accountno_lbl.Width + lngTpp)
23400         .accountno_lbl_line_dim_hi.Width = .accountno_lbl_line.Width
23410         If blnSortHere = True Then
23420           .Sort_lbl.Left = (((.accountno_lbl.Left + .accountno_lbl.Width) - lngSortLbl_Width) + (5& * lngTpp))
23430           .Sort_line.Width = (.accountno_lbl.Width + lngTpp)
23440           .Sort_line.Left = .accountno_lbl.Left
23450         End If
23460         blnSortHere = False
23470         .shortname_lbl.Width = ((.shortname.Width + .Contact_Number.Width) + (2& * lngTpp))
23480         .shortname_lbl_dim_hi.Width = .shortname_lbl.Width
23490         .shortname_lbl_line.Width = (.shortname_lbl.Width + lngTpp)
23500         .shortname_lbl_Line_dim_hi.Width = .shortname_lbl_line.Width

23510         .accountno.Visible = True
23520         .shortname.Visible = True
23530         .accountno_lbl.Visible = True
23540         .shortname_lbl.Visible = True
23550         .accountno_lbl_line.Visible = True
23560         .shortname_lbl_line.Visible = True

23570         If lngRecsCur = 0& Then
23580           .accountno_lbl_dim_hi.Visible = True
23590           .shortname_lbl_dim_hi.Visible = True
23600           .accountno_lbl_line_dim_hi.Visible = True
23610           .shortname_lbl_Line_dim_hi.Visible = True
23620         Else
23630           .accountno_lbl_dim_hi.Visible = False
23640           .shortname_lbl_dim_hi.Visible = False
23650           .accountno_lbl_line_dim_hi.Visible = False
23660           .shortname_lbl_Line_dim_hi.Visible = False
23670         End If

23680         Select Case blnEnableCountry
              Case True
23690           .Width = lngNewFormSub_Width - lngTpp
23700         Case False
23710           lngNewFormSub_Width = (((arr_varCtl(C_F2_WDT, 0) + arr_varCtl(C_WDT, 16) + lngCtlSep) - (arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))) - lngTpp)
23720           .Width = lngNewFormSub_Width
23730         End Select

23740       End With

23750       lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
23760       lngMonitorNum = 1&: lngTmp02 = 0&
23770       EnumMonitors frm  ' ** Module Function: modMonitorFuncs.
23780       If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

            ' ** lngSubFrm_Width is the subform control on this form.
23790       Select Case blnEnableCountry
            Case True
              ' ** Original width, plus ShortName and Country.
23800         DoCmd.SelectObject acForm, frm.Name, False
23810         If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top_Orig
23820         DoCmd.SelectObject acForm, frm.Name, False
23830         DoEvents
23840         DoCmd.MoveSize (lngFrm_Left_Orig - ((arr_varCtl(C_WDT, 16) + lngCtlSep + arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19)) / 2)), lngTmp02, _
                (lngFrm_Width_Orig + arr_varCtl(C_WDT, 16) + lngCtlSep + arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19)), lngFrm_Height_Orig  'lngFrm_Top_Orig
23850         If lngMonitorNum > 1& Then
23860           LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
23870         End If
23880         lngNewForm_Width = (lngForm_Width + arr_varCtl(C_WDT, 16) + lngCtlSep + arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))
23890         .Width = lngNewForm_Width
23900         .frmAccountContacts_Sub.Width = (arr_varCtl(C_F1_WDT, 0) + arr_varCtl(C_WDT, 16) + lngCtlSep + arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))
23910         .AcctNoShort_Move lngNewForm_Width  ' ** Form Procedure: frmAccountContacts.
23920       Case False
              ' ** Original width, plus ShortName.
23930         DoCmd.SelectObject acForm, frm.Name, False
23940         If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top_Orig
23950         DoCmd.SelectObject acForm, frm.Name, False
23960         DoEvents
23970         DoCmd.MoveSize (lngFrm_Left_Orig - ((arr_varCtl(C_WDT, 16) + lngCtlSep) / 2)), lngTmp02, _
                (lngFrm_Width_Orig + arr_varCtl(C_WDT, 16) + lngCtlSep), lngFrm_Height_Orig  'lngFrm_Top_Orig
23980         If lngMonitorNum > 1& Then
23990           LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
24000         End If
24010         lngNewForm_Width = (lngForm_Width + arr_varCtl(C_WDT, 16) + lngCtlSep)
24020         .Width = lngNewForm_Width
24030         .frmAccountContacts_Sub.Width = (arr_varCtl(C_F1_WDT, 0) + arr_varCtl(C_WDT, 16) + lngCtlSep)
24040         .AcctNoShort_Move lngNewForm_Width  ' ** Form Procedure: frmAccountContacts.
24050       End Select
24060       .frmAccountContacts_Sub_box.Width = (.frmAccountContacts_Sub.Width + (2& * lngTpp))
24070       .Nav_Sub1_box01.Width = .frmAccountContacts_Sub.Width
24080       .Width = lngNewForm_Width
24090     Else
24100       If blnAcctNo = True Then
              ' ** This is the opening width, with the window at its narrowest, if it's without Country.
24110         With .frmAccountContacts_Sub.Form

24120           Select Case blnEnableCountry
                Case True
                  ' ** lngFormSub_Width is accountno only, with Country.
24130             lngNewFormSub_Width = (arr_varCtl(C_F2_WDT, 0) - lngTpp)
24140           Case False
24150             lngNewFormSub_Width = ((arr_varCtl(C_F2_WDT, 0) - (arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))) - lngTpp)
24160           End Select
24170           .Width = lngNewFormSub_Width

24180           If .Sort_line.Left = .Contact_Name_lbl.Left Then blnSortHere = True
24190           .Contact_Number.Left = ((.accountno.Left + .accountno.Width) + (2& * lngTpp))
24200           .Contact_Name.Left = (.Contact_Number.Left + (arr_varCtl(C_LFT, 18) - arr_varCtl(C_LFT, 17)))
24210           .Contact_Name_LastFirst.Left = .Contact_Name.Left
24220           If blnSortHere = True Then
24230             strCtlName = "Contact_Name_lbl"
24240           End If
24250           blnSortHere = False
24260           If .Sort_line.Left = .Contact_Address1_lbl.Left Then blnSortHere = True
24270           .Contact_Address1.Left = ((.Contact_Name.Left + .Contact_Name.Width) + lngCtlSep)
24280           If blnSortHere = True Then
24290             strCtlName = "Contact_Address1_lbl"
24300           End If
24310           blnSortHere = False
24320           If .Sort_line.Left = .Contact_Address2_lbl.Left Then blnSortHere = True
24330           .Contact_Address2.Left = ((.Contact_Address1.Left + .Contact_Address1.Width) + lngCtlSep)
24340           If blnSortHere = True Then
24350             strCtlName = "Contact_Address2_lbl"
24360           End If
24370           blnSortHere = False
24380           If .Sort_line.Left = .Contact_City_lbl.Left Then blnSortHere = True
24390           .Contact_City.Left = ((.Contact_Address2.Left + .Contact_Address2.Width) + lngCtlSep)
24400           If blnSortHere = True Then
24410             strCtlName = "Contact_City_lbl"
24420           End If
24430           blnSortHere = False

24440           Select Case blnEnableCountry
                Case True
24450             If .Sort_line.Left = .opgLocContact_lbl.Left Then blnSortHere = True
24460             lngTmp03 = ((.Contact_City.Left + .Contact_City.Width) + lngCtlSep)
24470             If lngTmp03 > .opgLocContact.Left Then
24480               lngTmp03 = (lngTmp03 - .opgLocContact.Left)
24490               intTmp04 = (lngTmp03 / lngTpp)
24500               For intX = 1 To intTmp04
24510                 .opgLocContact.Left = (.opgLocContact.Left + lngTpp)
24520                 .opgLocContact_optOther_lbl.Left = (.opgLocContact_optOther_lbl.Left + lngTpp)
24530                 .opgLocContact_optOther.Left = (.opgLocContact_optOther.Left + lngTpp)
24540                 .opgLocContact_optUSA_lbl.Left = (.opgLocContact_optUSA_lbl.Left + lngTpp)
24550                 .opgLocContact_optUSA.Left = (.opgLocContact_optUSA.Left + lngTpp)
24560               Next  ' ** intX.
24570             ElseIf lngTmp03 < .opgLocContact.Left Then
24580               lngTmp03 = (.opgLocContact.Left - lngTmp03)
24590               intTmp04 = (lngTmp03 / lngTpp)
24600               For intX = 1 To intTmp04
24610                 .opgLocContact.Left = (.opgLocContact.Left - lngTpp)
24620                 .opgLocContact_optOther_lbl.Left = (.opgLocContact_optOther_lbl.Left - lngTpp)
24630                 .opgLocContact_optOther.Left = (.opgLocContact_optOther.Left - lngTpp)
24640                 .opgLocContact_optUSA_lbl.Left = (.opgLocContact_optUSA_lbl.Left - lngTpp)
24650                 .opgLocContact_optUSA.Left = (.opgLocContact_optUSA.Left - lngTpp)
24660               Next  ' ** intX.
24670             End If
24680             .opgLocContact_box.Left = .opgLocContact.Left
24690             .opgLocContact_optOther_box.Left = (.opgLocContact_optOther.Left - (5& * lngTpp))
24700             .opgLocContact_optOther_lbl_on.Left = .opgLocContact_optOther_lbl.Left
24710             .opgLocContact_optOther_lbl_off.Left = .opgLocContact_optOther_lbl.Left
24720             .opgLocContact_optUSA_box.Left = (.opgLocContact_optUSA.Left - (5& * lngTpp))
24730             .opgLocContact_optUSA_lbl_on.Left = .opgLocContact_optUSA_lbl.Left
24740             .opgLocContact_optUSA_lbl_off.Left = .opgLocContact_optUSA_lbl.Left
24750             .opgLocContact_lbl.Left = (.opgLocContact.Left - lngTpp)
24760             .opgLocContact_lbl_dim_hi.Left = (.opgLocContact_lbl.Left + lngTpp)
24770             .opgLocContact_lbl_line.Left = .opgLocContact_lbl.Left
24780             .opgLocContact_lbl_line_dim_hi.Left = .opgLocContact_lbl_line.Left
24790             If blnSortHere = True Then
24800               .Sort_line.Left = .opgLocContact_lbl.Left
24810               .Sort_lbl.Left = ((.opgLocContact_lbl.Left + .opgLocContact_lbl.Width) - lngSortLbl_Width)
24820             End If
24830           Case False
                  ' ** Nothing to do.
24840           End Select
24850           blnSortHere = False

24860           If .Sort_line.Left = .Contact_State_lbl.Left Then blnSortHere = True
24870           Select Case blnEnableCountry
                Case True
24880             .Contact_State.Left = ((.opgLocContact.Left + .opgLocContact.Width) + lngCtlSep)
24890           Case False
24900             .Contact_State.Left = ((.Contact_City.Left + .Contact_City.Width) + lngCtlSep)
24910           End Select
24920           .Contact_State_box.Left = (.Contact_State.Left - lngTpp)
24930           If blnSortHere = True Then
24940             strCtlName = "Contact_State_lbl"
24950           End If
24960           blnSortHere = False
24970           If .Sort_line.Left = .Contact_Zip_lbl.Left Then blnSortHere = True
24980           .Contact_Zip.Left = ((.Contact_State.Left + .Contact_State.Width) + lngCtlSep)
24990           .Contact_Zip_Format.Left = .Contact_Zip.Left
25000           .Contact_Zip_box.Left = (.Contact_Zip.Left - lngTpp)
25010           If blnSortHere = True Then
25020             strCtlName = "Contact_Zip_lbl"
25030           End If
25040           blnSortHere = False

25050           Select Case blnEnableCountry
                Case True
25060             If .Sort_line.Left = .Contact_Country_lbl.Left Then blnSortHere = True
25070             .Contact_Country.Left = ((.Contact_Zip.Left + .Contact_Zip.Width) + lngCtlSep)
25080             .Contact_Country_box.Left = (.Contact_Country.Left - lngTpp)
25090             If blnSortHere = True Then
25100               strCtlName = "Contact_Country_lbl"
25110             End If
25120             blnSortHere = False
25130             If .Sort_line.Left = .Contact_PostalCode_lbl.Left Then blnSortHere = True
25140             .Contact_PostalCode.Left = ((.Contact_Country.Left + .Contact_Country.Width) + lngCtlSep)
25150             .Contact_PostalCode_box.Left = (.Contact_PostalCode.Left - lngTpp)
25160             If blnSortHere = True Then
25170               strCtlName = "Contact_PostalCode_lbl"
25180             End If
25190             blnSortHere = False
25200             If .Sort_line.Left = .Contact_Phone1_lbl.Left Then blnSortHere = True
25210             .Contact_Phone1.Left = ((.Contact_PostalCode.Left + .Contact_PostalCode.Width) + lngCtlSep)
25220             .Contact_Phone1_Format.Left = .Contact_Phone1.Left
25230             If blnSortHere = True Then
25240               strCtlName = "Contact_Phone1_lbl"
25250             End If
25260             blnSortHere = False
25270           Case False
25280             If .Sort_line.Left = .Contact_Phone1_lbl.Left Then blnSortHere = True
25290             .Contact_Phone1.Left = ((.Contact_Zip.Left + .Contact_Zip.Width) + lngCtlSep)
25300             .Contact_Phone1_Format.Left = .Contact_Phone1.Left
25310             If blnSortHere = True Then
25320               strCtlName = "Contact_Phone1_lbl"
25330             End If
25340             blnSortHere = False
25350           End Select

25360           If .Sort_line.Left = .Contact_Phone2_lbl.Left Then blnSortHere = True
25370           .Contact_Phone2.Left = ((.Contact_Phone1.Left + .Contact_Phone1.Width) + lngCtlSep)
25380           .Contact_Phone2_Format.Left = .Contact_Phone2.Left
25390           If blnSortHere = True Then
25400             strCtlName = "Contact_Phone2_lbl"
25410           End If
25420           blnSortHere = False
25430           If .Sort_line.Left = .Contact_Fax_lbl.Left Then blnSortHere = True
25440           .Contact_Fax.Left = ((.Contact_Phone2.Left + .Contact_Phone2.Width) + lngCtlSep)
25450           .Contact_Fax_Format.Left = .Contact_Fax.Left
25460           If blnSortHere = True Then
25470             strCtlName = "Contact_Fax_lbl"
25480           End If
25490           blnSortHere = False
25500           If .Sort_line.Left = .Contact_Email_lbl.Left Then blnSortHere = True
25510           .Contact_Email.Left = ((.Contact_Fax.Left + .Contact_Fax.Width) + lngCtlSep)
25520           .Contact_Email_cmd.Left = (.Contact_Email.Left - (3& * lngTpp))
25530           If blnSortHere = True Then
25540             strCtlName = "Contact_Email_lbl"
25550           End If
25560           blnSortHere = False

25570           intPos01 = InStr(strSortNow, " ")
25580           If intPos01 > 0 Then strTmp01 = Left(strSortNow, intPos01) Else strTmp01 = strSortNow
25590           intPos01 = InStr(strTmp01, "accountno")
25600           If .Sort_line.Left = .accountno_lbl.Left And intPos01 > 0 Then blnSortHere = True
25610           .accountno_lbl.Width = ((.accountno.Width + .Contact_Number.Width) + (2& * lngTpp))
25620           .accountno_lbl_dim_hi.Width = .accountno_lbl.Width
25630           .accountno_lbl_line.Width = (.accountno_lbl.Width + lngTpp)
25640           .accountno_lbl_line_dim_hi.Width = .accountno_lbl_line.Width
25650           If blnSortHere = True Then
25660             .Sort_lbl.Left = ((.accountno_lbl.Left + .accountno_lbl.Width) - lngSortLbl_Width)
25670             .Sort_line.Width = (.accountno_lbl.Width + lngTpp)
25680             .Sort_line.Left = .accountno_lbl.Left
25690           End If
25700           blnSortHere = False

                ' ** If it was on with shortname only or both showing, move it to accountno.
25710           intPos01 = InStr(strSortNow, " ")
25720           If intPos01 > 0 Then strTmp01 = Left(strSortNow, intPos01) Else strTmp01 = strSortNow
25730           intPos01 = InStr(strTmp01, "shortname")
25740           If .Sort_line.Left = .shortname_lbl.Left And intPos01 > 0 Then blnSortHere = True
25750           .shortname.Visible = False
25760           .shortname_lbl.Visible = False
25770           .shortname_lbl_line.Visible = False
25780           If blnSortHere = True Then
25790             blnResort = True
25800             .Sort_lbl.Visible = False
25810             .Sort_lbl.Left = 0&
25820             .Sort_line.Visible = False
25830             .Sort_line.Left = 0&
25840           End If
25850           blnSortHere = False
25860           .accountno.Visible = True
25870           .accountno_lbl.Visible = True
25880           .accountno_lbl_line.Visible = True

25890           If lngRecsCur = 0& Then
25900             .accountno_lbl_dim_hi.Visible = True
25910             .shortname_lbl_dim_hi.Visible = False
25920             .accountno_lbl_line_dim_hi.Visible = True
25930             .shortname_lbl_Line_dim_hi.Visible = False
25940           Else
25950             .accountno_lbl_dim_hi.Visible = False
25960             .accountno_lbl_line_dim_hi.Visible = False
25970             .shortname_lbl_dim_hi.Visible = False
25980             .shortname_lbl_Line_dim_hi.Visible = False
25990           End If

26000           Select Case blnEnableCountry
                Case True
26010             .Width = (lngNewFormSub_Width - lngTpp)
26020           Case False
26030             lngNewFormSub_Width = ((arr_varCtl(C_F2_WDT, 0) - (arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))) - lngTpp)
26040             .Width = lngNewFormSub_Width
26050           End Select

26060           If blnResort = True Then
26070             .SortNow "accountno_lbl_DblClick"  ' ** Form Procedure: frmAccountContacts_Sub.
26080           End If

26090         End With

              ' ** lngSubFrm_Width is the subform control on this form.
26100         Select Case blnEnableCountry
              Case True
                ' ** Original width, plus Country.
26110           DoCmd.SelectObject acForm, frm.Name, False
26120           If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top_Orig
26130           DoCmd.SelectObject acForm, frm.Name, False
26140           DoEvents
26150           DoCmd.MoveSize (lngFrm_Left_Orig - ((arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19)) / 2&)), lngTmp02, _
                  (lngFrm_Width_Orig + (arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))), lngFrm_Height_Orig  'lngFrm_Top_Orig
26160           If lngMonitorNum > 1& Then
26170             LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
26180           End If
26190           lngNewForm_Width = (lngForm_Width + arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))
26200           .Width = lngNewForm_Width
26210           .frmAccountContacts_Sub.Width = (arr_varCtl(C_F1_WDT, 0) + arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))
26220           .AcctNoShort_Move lngNewForm_Width  ' ** Form Procedure: frmAccountContacts.
26230         Case False
                ' ** Original, narrowest width.
26240           DoCmd.SelectObject acForm, frm.Name, False
26250           If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top_Orig
26260           DoCmd.SelectObject acForm, frm.Name, False
26270           DoEvents
26280           DoCmd.MoveSize lngFrm_Left_Orig, lngTmp02, lngFrm_Width_Orig, lngFrm_Height_Orig  'lngFrm_Top_Orig
26290           If lngMonitorNum > 1& Then
26300             LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
26310           End If
26320           lngNewForm_Width = lngForm_Width
26330           .Width = lngNewForm_Width
26340           .frmAccountContacts_Sub.Width = arr_varCtl(C_F1_WDT, 0)
26350           .AcctNoShort_Move lngNewForm_Width  ' ** Form Procedure: frmAccountContacts.
26360         End Select
26370         .frmAccountContacts_Sub_box.Width = (.frmAccountContacts_Sub.Width + (2& * lngTpp))
26380         .Nav_Sub1_box01.Width = .frmAccountContacts_Sub.Width
26390         .Width = lngNewForm_Width

26400       ElseIf blnShortName = True Then
              ' ** Window should be at an intermediate width.
26410         With .frmAccountContacts_Sub.Form

26420           Select Case blnEnableCountry
                Case True
                  ' ** lngFormSub_Width is accountno only, with Country.
26430             lngNewFormSub_Width = ((arr_varCtl(C_F2_WDT, 0) + arr_varCtl(C_L1_LFT, 16)) - lngTpp)
26440           Case False
26450             lngNewFormSub_Width = (((arr_varCtl(C_F2_WDT, 0) + arr_varCtl(C_L1_LFT, 16)) - (arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))) - lngTpp)
26460           End Select
26470           .Width = lngNewFormSub_Width

26480           .shortname.Left = arr_varCtl(C_LFT, 15)
26490           If .Sort_line.Left = .Contact_Name_lbl.Left Then blnSortHere = True
26500           .Contact_Number.Left = ((.shortname.Left + .shortname.Width) + (2& * lngTpp))
26510           .Contact_Name.Left = ((.Contact_Number.Left + .Contact_Number.Width) + lngCtlSep)
26520           .Contact_Name_LastFirst.Left = .Contact_Name.Left
26530           If blnSortHere = True Then
26540             strCtlName = "Contact_Name_lbl"
26550           End If
26560           blnSortHere = False
26570           If .Sort_line.Left = .Contact_Address1_lbl.Left Then blnSortHere = True
26580           .Contact_Address1.Left = ((.Contact_Name.Left + .Contact_Name.Width) + lngCtlSep)
26590           If blnSortHere = True Then
26600             strCtlName = "Contact_Address1_lbl"
26610           End If
26620           blnSortHere = False
26630           If .Sort_line.Left = .Contact_Address2_lbl.Left Then blnSortHere = True
26640           .Contact_Address2.Left = ((.Contact_Address1.Left + .Contact_Address1.Width) + lngCtlSep)
26650           If blnSortHere = True Then
26660             strCtlName = "Contact_Address2_lbl"
26670           End If
26680           blnSortHere = False
26690           If .Sort_line.Left = .Contact_City_lbl.Left Then blnSortHere = True
26700           .Contact_City.Left = ((.Contact_Address2.Left + .Contact_Address2.Width) + lngCtlSep)
26710           If blnSortHere = True Then
26720             strCtlName = "Contact_City_lbl"
26730           End If
26740           blnSortHere = False

26750           Select Case blnEnableCountry
                Case True
26760             If .Sort_line.Left = .opgLocContact_lbl.Left Then blnSortHere = True
26770             lngTmp03 = ((.Contact_City.Left + .Contact_City.Width) + lngCtlSep)
26780             If lngTmp03 > .opgLocContact.Left Then
26790               lngTmp03 = (lngTmp03 - .opgLocContact.Left)
26800               intTmp04 = (lngTmp03 / lngTpp)
26810               For intX = 1 To intTmp04
26820                 .opgLocContact.Left = (.opgLocContact.Left + lngTpp)
26830                 .opgLocContact_optOther_lbl.Left = (.opgLocContact_optOther_lbl.Left + lngTpp)
26840                 .opgLocContact_optOther.Left = (.opgLocContact_optOther.Left + lngTpp)
26850                 .opgLocContact_optUSA_lbl.Left = (.opgLocContact_optUSA_lbl.Left + lngTpp)
26860                 .opgLocContact_optUSA.Left = (.opgLocContact_optUSA.Left + lngTpp)
26870               Next  ' ** intX.
26880             ElseIf lngTmp03 < .opgLocContact.Left Then
26890               lngTmp03 = (.opgLocContact.Left - lngTmp03)
26900               intTmp04 = (lngTmp03 / lngTpp)
26910               For intX = 1 To intTmp04
26920                 .opgLocContact.Left = (.opgLocContact.Left - lngTpp)
26930                 .opgLocContact_optOther_lbl.Left = (.opgLocContact_optOther_lbl.Left - lngTpp)
26940                 .opgLocContact_optOther.Left = (.opgLocContact_optOther.Left - lngTpp)
26950                 .opgLocContact_optUSA_lbl.Left = (.opgLocContact_optUSA_lbl.Left - lngTpp)
26960                 .opgLocContact_optUSA.Left = (.opgLocContact_optUSA.Left - lngTpp)
26970               Next  ' ** intX.
26980             End If
26990             .opgLocContact_box.Left = .opgLocContact.Left
27000             .opgLocContact_optOther_box.Left = (.opgLocContact_optOther.Left - (5& * lngTpp))
27010             .opgLocContact_optOther_lbl_on.Left = .opgLocContact_optOther_lbl.Left
27020             .opgLocContact_optOther_lbl_off.Left = .opgLocContact_optOther_lbl.Left
27030             .opgLocContact_optUSA_box.Left = (.opgLocContact_optUSA.Left - (5& * lngTpp))
27040             .opgLocContact_optUSA_lbl_on.Left = .opgLocContact_optUSA_lbl.Left
27050             .opgLocContact_optUSA_lbl_off.Left = .opgLocContact_optUSA_lbl.Left
27060             .opgLocContact_lbl.Left = (.opgLocContact.Left - lngTpp)
27070             .opgLocContact_lbl_dim_hi.Left = (.opgLocContact_lbl.Left + lngTpp)
27080             .opgLocContact_lbl_line.Left = .opgLocContact_lbl.Left
27090             .opgLocContact_lbl_line_dim_hi.Left = .opgLocContact_lbl_line.Left
27100             If blnSortHere = True Then
27110               .Sort_line.Left = .opgLocContact_lbl.Left
27120               .Sort_lbl.Left = ((.opgLocContact_lbl.Left + .opgLocContact_lbl.Width) - lngSortLbl_Width)
27130             End If
27140           Case False
                  ' ** Nothing to do.
27150           End Select
27160           blnSortHere = False

27170           If .Sort_line.Left = .Contact_State_lbl.Left Then blnSortHere = True
27180           Select Case blnEnableCountry
                Case True
27190             .Contact_State.Left = ((.opgLocContact.Left + .opgLocContact.Width) + lngCtlSep)
27200           Case False
27210             .Contact_State.Left = ((.Contact_City.Left + .Contact_City.Width) + lngCtlSep)
27220           End Select
27230           .Contact_State_box.Left = (.Contact_State.Left - lngTpp)
27240           If blnSortHere = True Then
27250             strCtlName = "Contact_State_lbl"
27260           End If
27270           blnSortHere = False
27280           If .Sort_line.Left = .Contact_Zip_lbl.Left Then blnSortHere = True
27290           .Contact_Zip.Left = ((.Contact_State.Left + .Contact_State.Width) + lngCtlSep)
27300           .Contact_Zip_Format.Left = .Contact_Zip.Left
27310           .Contact_Zip_box.Left = (.Contact_Zip.Left - lngTpp)
27320           If blnSortHere = True Then
27330             strCtlName = "Contact_Zip_lbl"
27340           End If
27350           blnSortHere = False

27360           Select Case blnEnableCountry
                Case True
27370             If .Sort_line.Left = .Contact_Country_lbl.Left Then blnSortHere = True
27380             .Contact_Country.Left = ((.Contact_Zip.Left + .Contact_Zip.Width) + lngCtlSep)
27390             .Contact_Country_box.Left = (.Contact_Country.Left - lngTpp)
27400             If blnSortHere = True Then
27410               strCtlName = "Contact_Country_lbl"
27420             End If
27430             blnSortHere = False
27440             If .Sort_line.Left = .Contact_PostalCode_lbl.Left Then blnSortHere = True
27450             .Contact_PostalCode.Left = ((.Contact_Country.Left + .Contact_Country.Width) + lngCtlSep)
27460             .Contact_PostalCode_box.Left = (.Contact_PostalCode.Left - lngTpp)
27470             If blnSortHere = True Then
27480               strCtlName = "Contact_PostalCode_lbl"
27490             End If
27500             blnSortHere = False
27510             If .Sort_line.Left = .Contact_Phone1_lbl.Left Then blnSortHere = True
27520             .Contact_Phone1.Left = ((.Contact_PostalCode.Left + .Contact_PostalCode.Width) + lngCtlSep)
27530             .Contact_Phone1_Format.Left = .Contact_Phone1.Left
27540             If blnSortHere = True Then
27550               strCtlName = "Contact_Phone1_lbl"
27560             End If
27570             blnSortHere = False
27580           Case False
27590             If .Sort_line.Left = .Contact_Phone1_lbl.Left Then blnSortHere = True
27600             .Contact_Phone1.Left = ((.Contact_Zip.Left + .Contact_Zip.Width) + lngCtlSep)
27610             .Contact_Phone1_Format.Left = .Contact_Phone1.Left
27620             If blnSortHere = True Then
27630               strCtlName = "Contact_Phone1_lbl"
27640             End If
27650             blnSortHere = False
27660           End Select

27670           If .Sort_line.Left = .Contact_Phone2_lbl.Left Then blnSortHere = True
27680           .Contact_Phone2.Left = ((.Contact_Phone1.Left + .Contact_Phone1.Width) + lngCtlSep)
27690           .Contact_Phone2_Format.Left = .Contact_Phone2.Left
27700           If blnSortHere = True Then
27710             strCtlName = "Contact_Phone2_lbl"
27720           End If
27730           blnSortHere = False
27740           If .Sort_line.Left = .Contact_Fax_lbl.Left Then blnSortHere = True
27750           .Contact_Fax.Left = ((.Contact_Phone2.Left + .Contact_Phone2.Width) + lngCtlSep)
27760           .Contact_Fax_Format.Left = .Contact_Fax.Left
27770           If blnSortHere = True Then
27780             strCtlName = "Contact_Fax_lbl"
27790           End If
27800           blnSortHere = False
27810           If .Sort_line.Left = .Contact_Email_lbl.Left Then blnSortHere = True
27820           .Contact_Email.Left = ((.Contact_Fax.Left + .Contact_Fax.Width) + lngCtlSep)
27830           .Contact_Email_cmd.Left = (.Contact_Email.Left - (3& * lngTpp))
27840           If blnSortHere = True Then
27850             strCtlName = "Contact_Email_lbl"
27860           End If
27870           blnSortHere = False

27880           intPos01 = InStr(strSortNow, " ")
27890           If intPos01 > 0 Then strTmp01 = Left(strSortNow, intPos01) Else strTmp01 = strSortNow
27900           intPos01 = InStr(strTmp01, "shortname")
27910           If .Sort_line.Left = .shortname_lbl.Left And intPos01 > 0 Then blnSortHere = True
27920           .shortname_lbl.Left = .shortname.Left
27930           .shortname_lbl_dim_hi.Left = (.shortname_lbl.Left + lngTpp)
27940           .shortname_lbl_line.Left = .shortname.Left
27950           .shortname_lbl_Line_dim_hi.Left = (.shortname_lbl_line.Left + lngTpp)
27960           .shortname_lbl.Width = ((.shortname.Width + .Contact_Number.Width) + (2& * lngTpp))
27970           .shortname_lbl_dim_hi.Width = .shortname_lbl.Width
27980           .shortname_lbl_line.Width = (.shortname_lbl.Width + lngTpp)
27990           .shortname_lbl_Line_dim_hi.Width = .shortname_lbl_line.Width
28000           If blnSortHere = True Then
28010             .Sort_lbl.Left = ((.shortname_lbl.Left + .shortname_lbl.Width) - lngSortLbl_Width)
28020             .Sort_line.Left = .shortname_lbl.Left
28030           End If
28040           blnSortHere = False

                ' ** If it was on with accountno only or both showing, move it to shortname.
28050           intPos01 = InStr(strSortNow, " ")
28060           If intPos01 > 0 Then strTmp01 = Left(strSortNow, intPos01) Else strTmp01 = strSortNow
28070           intPos01 = InStr(strTmp01, "accountno")
28080           If .Sort_line.Left = .accountno_lbl.Left And intPos01 > 0 Then blnSortHere = True
28090           .accountno.Visible = False
28100           .accountno_lbl.Visible = False
28110           .accountno_lbl_line.Visible = False
28120           If blnSortHere = True Then
28130             blnResort = True
28140             .Sort_lbl.Visible = False
28150             .Sort_lbl.Left = 0&
28160             .Sort_line.Visible = False
28170             .Sort_line.Left = 0&
28180           End If
28190           blnSortHere = False
28200           .shortname.Visible = True
28210           .shortname_lbl.Visible = True
28220           .shortname_lbl_line.Visible = True

28230           If lngRecsCur = 0& Then
28240             .accountno_lbl_dim_hi.Visible = False
28250             .shortname_lbl_dim_hi.Visible = True
28260             .accountno_lbl_line_dim_hi.Visible = False
28270             .shortname_lbl_Line_dim_hi.Visible = True
28280           Else
28290             .accountno_lbl_dim_hi.Visible = False
28300             .shortname_lbl_dim_hi.Visible = False
28310             .accountno_lbl_line_dim_hi.Visible = False
28320             .shortname_lbl_Line_dim_hi.Visible = False
28330           End If

28340           Select Case blnEnableCountry
                Case True
28350             .Width = (lngNewFormSub_Width - lngTpp)
28360           Case False
28370             lngNewFormSub_Width = (((arr_varCtl(C_F2_WDT, 0) + arr_varCtl(C_L1_LFT, 16)) - (arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))) - lngTpp)
28380             .Width = lngNewFormSub_Width
28390           End Select

28400           If blnResort = True Then
28410             .SortNow "shortname_lbl_DblClick"  ' ** Form Procedure: frmAccountContacts_Sub.
28420           End If

28430         End With

              ' ** lngSubFrm_Width is the subform control on this form.
28440         Select Case blnEnableCountry
              Case True
                ' ** Original width, plus Country and ShortName difference.
28450           DoCmd.SelectObject acForm, frm.Name, False
28460           If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top_Orig
28470           DoCmd.SelectObject acForm, frm.Name, False
28480           DoEvents
28490           DoCmd.MoveSize (lngFrm_Left_Orig - ((arr_varCtl(C_L1_LFT, 16) + arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19)) / 2&)), lngTmp02, _
                  (lngFrm_Width_Orig + arr_varCtl(C_L1_LFT, 16) + arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19)), lngFrm_Height_Orig  'lngFrm_Top_Orig
28500           If lngMonitorNum > 1& Then
28510             LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
28520           End If
28530           lngNewForm_Width = (lngForm_Width + arr_varCtl(C_L1_LFT, 16) + arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))
28540           .Width = lngNewForm_Width
28550           .frmAccountContacts_Sub.Width = (arr_varCtl(C_F1_WDT, 0) + arr_varCtl(C_L1_LFT, 16) + arr_varCtl(C_L1_LFT, 20) + arr_varCtl(C_L1_LFT, 19))
28560           .AcctNoShort_Move lngNewForm_Width  ' ** Form Procedure: frmAccountContacts.
28570         Case False
                ' ** Original width, plus ShortName difference.
28580           DoCmd.SelectObject acForm, frm.Name, False
28590           If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top_Orig
28600           DoCmd.SelectObject acForm, frm.Name, False
28610           DoEvents
28620           DoCmd.MoveSize (lngFrm_Left_Orig - (arr_varCtl(C_L1_LFT, 16) / 2&)), lngTmp02, _
                  (lngFrm_Width_Orig + arr_varCtl(C_L1_LFT, 16)), lngFrm_Height_Orig  'lngFrm_Top_Orig
28630           If lngMonitorNum > 1& Then
28640             LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
28650           End If
28660           lngNewForm_Width = (lngForm_Width + arr_varCtl(C_L1_LFT, 16))
28670           .Width = lngNewForm_Width
28680           .frmAccountContacts_Sub.Width = (arr_varCtl(C_F1_WDT, 0) + arr_varCtl(C_L1_LFT, 16))
28690           .AcctNoShort_Move lngNewForm_Width  ' ** Form Procedure: frmAccountContacts.
28700         End Select
28710         .frmAccountContacts_Sub_box.Width = (.frmAccountContacts_Sub.Width + (2& * lngTpp))
28720         .Nav_Sub1_box01.Width = .frmAccountContacts_Sub.Width
28730         .Width = lngNewForm_Width

28740       End If
28750     End If

          ' ** It seems all the labels are being done here!
28760     With .frmAccountContacts_Sub.Form

28770       .Contact_Name_lbl.Left = .Contact_Name.Left
28780       .Contact_Name_lbl_dim_hi.Left = (.Contact_Name_lbl.Left + lngTpp)
28790       .Contact_Address1_lbl.Left = .Contact_Address1.Left
28800       .Contact_Address1_lbl_dim_hi.Left = (.Contact_Address1_lbl.Left + lngTpp)
28810       .Contact_Address2_lbl.Left = .Contact_Address2.Left
28820       .Contact_Address2_lbl_dim_hi.Left = (.Contact_Address2_lbl.Left + lngTpp)
28830       .Contact_City_lbl.Left = .Contact_City.Left
28840       .Contact_City_lbl_dim_hi.Left = (.Contact_City_lbl.Left + lngTpp)
28850       .Contact_State_lbl.Left = .Contact_State.Left
28860       .Contact_State_lbl_dim_hi.Left = (.Contact_State_lbl.Left + lngTpp)
28870       .Contact_Zip_lbl.Left = .Contact_Zip.Left
28880       .Contact_Zip_lbl_dim_hi.Left = (.Contact_Zip_lbl.Left + lngTpp)
28890       Select Case blnEnableCountry
            Case True
28900         .Contact_Country_lbl.Left = .Contact_Country.Left
28910         .Contact_Country_lbl_dim_hi.Left = (.Contact_Country_lbl.Left + lngTpp)
28920         .Contact_PostalCode_lbl.Left = .Contact_PostalCode.Left
28930         .Contact_PostalCode_lbl_dim_hi.Left = (.Contact_PostalCode_lbl.Left + lngTpp)
28940       Case False
              ' ** Nothing.
28950       End Select
28960       .Contact_Phone1_lbl.Left = .Contact_Phone1.Left
28970       .Contact_Phone1_lbl_dim_hi.Left = (.Contact_Phone1_lbl.Left + lngTpp)
28980       .Contact_Phone2_lbl.Left = .Contact_Phone2.Left
28990       .Contact_Phone2_lbl_dim_hi.Left = (.Contact_Phone2_lbl.Left + lngTpp)
29000       .Contact_Fax_lbl.Left = .Contact_Fax.Left
29010       .Contact_Fax_lbl_dim_hi.Left = (.Contact_Fax_lbl.Left + lngTpp)
29020       .Contact_Email_lbl.Left = .Contact_Email.Left
29030       .Contact_Email_lbl2.Left = ((.Contact_Email_lbl.Left + .Contact_Email_lbl.Width) - .Contact_Email_lbl2.Width)
29040       .Contact_Email_lbl_dim_hi.Left = (.Contact_Email_lbl.Left + lngTpp)

29050       .Contact_Name_lbl_line.Left = .Contact_Name.Left
29060       .Contact_Name_lbl_line_dim_hi.Left = (.Contact_Name_lbl_line.Left + lngTpp)
29070       .Contact_Address1_lbl_line.Left = .Contact_Address1.Left
29080       .Contact_Address1_lbl_line_dim_hi.Left = (.Contact_Address1_lbl_line.Left + lngTpp)
29090       .Contact_Address2_lbl_line.Left = .Contact_Address2.Left
29100       .Contact_Address2_lbl_line_dim_hi.Left = (.Contact_Address2_lbl_line.Left + lngTpp)
29110       .Contact_City_lbl_line.Left = .Contact_City.Left
29120       .Contact_City_lbl_line_dim_hi.Left = (.Contact_City_lbl_line.Left + lngTpp)
29130       .Contact_State_lbl_line.Left = .Contact_State.Left
29140       .Contact_State_lbl_line_dim_hi.Left = (.Contact_State_lbl_line.Left + lngTpp)
29150       .Contact_Zip_lbl_line.Left = .Contact_Zip.Left
29160       .Contact_Zip_lbl_line_dim_hi.Left = (.Contact_Zip_lbl_line.Left + lngTpp)
29170       Select Case blnEnableCountry
            Case True
29180         .Contact_Country_lbl_line.Left = .Contact_Country.Left
29190         .Contact_Country_lbl_line_dim_hi.Left = (.Contact_Country_lbl_line.Left + lngTpp)
29200         .Contact_PostalCode_lbl_line.Left = .Contact_PostalCode.Left
29210         .Contact_PostalCode_lbl_line_dim_hi.Left = (.Contact_PostalCode_lbl_line.Left + lngTpp)
29220       Case False
              ' ** Nothing.
29230       End Select
29240       .Contact_Phone1_lbl_line.Left = .Contact_Phone1.Left
29250       .Contact_Phone1_lbl_line_dim_hi.Left = (.Contact_Phone1_lbl_line.Left + lngTpp)
29260       .Contact_Phone2_lbl_line.Left = .Contact_Phone2.Left
29270       .Contact_Phone2_lbl_line_dim_hi.Left = (.Contact_Phone2_lbl_line.Left + lngTpp)
29280       .Contact_Fax_lbl_line.Left = .Contact_Fax.Left
29290       .Contact_Fax_lbl_line_dim_hi.Left = (.Contact_Fax_lbl_line.Left + lngTpp)
29300       .Contact_Email_lbl_line.Left = .Contact_Email.Left
29310       .Contact_Email_lbl_line_dim_hi.Left = (.Contact_Email_lbl_line.Left + lngTpp)

            ' ** Shortname and accountno have already been moved.
29320       If strCtlName <> vbNullString Then
29330         .Sort_line.Left = .Controls(strCtlName).Left
29340         .Sort_lbl.Left = ((.Controls(strCtlName).Left + .Controls(strCtlName).Width) - lngSortLbl_Width)
29350       End If
29360       .Sort_line.Visible = True
29370       .Sort_lbl.Visible = True

29380       If lngNewFormSub_Width < lngCurFormSub_Width Then
29390         .Width = lngNewFormSub_Width
29400       End If

29410     End With

29420   End With

EXITP:
29430   Exit Sub

ERRH:
29440   Select Case ERR.Number
        Case Else
29450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29460   End Select
29470   Resume EXITP

End Sub
