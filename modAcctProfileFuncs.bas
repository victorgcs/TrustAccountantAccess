Attribute VB_Name = "modAcctProfileFuncs"
Option Compare Database
Option Explicit

'VGC 09/09/2017: CHANGES!

Private Const THIS_NAME As String = "modAcctProfileFuncs"
' **

Public Sub SysAcct_Set_AP(blnSysAcct As Boolean, blnViewOnly As Boolean, CLR_DISABLED_FG As Long, CLR_DISABLED_BG As Long, frm As Access.Form)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "SysAcct_Set_AP"

        Dim lngTmp01 As Long

110     With frm
120       Select Case blnSysAcct
          Case True
130         .accountno.Locked = True
140         .accountno.ForeColor = CLR_DISABLED_FG
150         .accountno.BackColor = CLR_DISABLED_BG
160         .related_accountno.Enabled = False
170         .related_accountno.Locked = False
180         .related_accountno.BorderColor = WIN_CLR_DISR
190         .related_accountno.BackStyle = acBackStyleTransparent
200         .related_accountno_lbl.BackStyle = acBackStyleTransparent
210         .cmdChoose.Enabled = False
220         .cmdChoose_lbl2.ForeColor = WIN_CLR_DISF
230         .cmdChoose_lbl2_dim_hi.Visible = True
240         .shortname.Locked = True
250         .shortname.ForeColor = CLR_DISABLED_FG
260         .shortname.BackColor = CLR_DISABLED_BG
270         .legalname.Locked = True
280         .legalname.ForeColor = CLR_DISABLED_FG
290         .legalname.BackColor = CLR_DISABLED_BG
300         .CaseNum.Enabled = False
310         .CaseNum.BorderColor = WIN_CLR_DISR
320         .CaseNum.BackStyle = acBackStyleTransparent
330         .CaseNum_lbl.BackStyle = acBackStyleTransparent
340         .accounttype.Enabled = False
350         .accounttype.Locked = True
360         .accounttype.ForeColor = CLR_DISABLED_FG
370         .accounttype.BackColor = CLR_DISABLED_BG
380         .description.Enabled = True  ' ** I know this seems weird, but it's locked.
390         .description.TabStop = True
400         .tin.Enabled = False
410         .tin.BorderColor = WIN_CLR_DISR
420         .tin.BackStyle = acBackStyleTransparent
430         .tin_lbl.BackStyle = acBackStyleTransparent
440         .documentdate.Enabled = False
450         .documentdate.BorderColor = WIN_CLR_DISR
460         .documentdate.BackStyle = acBackStyleTransparent
470         .documentdate_lbl.BackStyle = acBackStyleTransparent
480         .amendments.Enabled = False
490         .amendments.BorderColor = WIN_CLR_DISR
500         .amendments.BackStyle = acBackStyleTransparent
510         .amendments_lbl.BackStyle = acBackStyleTransparent
520         .appointmentdate.Enabled = False
530         .appointmentdate.BorderColor = WIN_CLR_DISR
540         .appointmentdate.BackStyle = acBackStyleTransparent
550         .appointmentdate_lbl.BackStyle = acBackStyleTransparent
560         .cotrustee.Enabled = False
570         .cotrustee.BorderColor = WIN_CLR_DISR
580         .cotrustee.BackStyle = acBackStyleTransparent
590         .cotrustee_lbl.BackStyle = acBackStyleTransparent
600         .discretion.Enabled = False
610         .discretion.BorderColor = WIN_CLR_DISR
620         .discretion.BackStyle = acBackStyleTransparent
630         .discretion_lbl.BackStyle = acBackStyleTransparent
640         .dateclosed.Enabled = False
650         .dateclosed.BorderColor = WIN_CLR_DISR
660         .dateclosed.BackStyle = acBackStyleTransparent
670         .dateclosed_lbl.BackStyle = acBackStyleTransparent
680         .courtsupervised.Enabled = False
690         .courtsupervised.BorderColor = WIN_CLR_DISR
700         .courtsupervised.BackStyle = acBackStyleTransparent
710         .courtsupervised_lbl.BackStyle = acBackStyleTransparent
720         .jurisdiction.Enabled = False
730         .jurisdiction.BorderColor = WIN_CLR_DISR
740         .jurisdiction.BackStyle = acBackStyleTransparent
750         .jurisdiction_lbl.BackStyle = acBackStyleTransparent
760         .county.Enabled = False
770         .county.BorderColor = WIN_CLR_DISR
780         .county.BackStyle = acBackStyleTransparent
790         .county_lbl.BackStyle = acBackStyleTransparent
800         .cmdStatementFreq.Enabled = True  ' ** This stays available.
810         .cmdReviewFreq.Enabled = False
820         .cmdReviewFreq_lbl.BackStyle = acBackStyleTransparent
830         .opgFeeFreq_box_dim.Visible = True
840         .opgFeeFreq_box.Visible = False
850         .opgFeeFreq_opt_box.Visible = False
860         .opgFeeFreq_lbl.ForeColor = WIN_CLR_DISF
870         .opgFeeFreq_lbl.BackStyle = acBackStyleTransparent
880         .opgFeeFreq_lbl_dim_hi.Visible = True
890         .opgFeeFreq_hline03.Visible = True
900         .opgFeeFreq_optMonth.Enabled = False
910         .opgFeeFreq_optMonth = False
920         .opgFeeFreq_optQuarter.Enabled = False
930         .opgFeeFreq_optQuarter = False
940         .opgFeeFreq_optSemi.Enabled = False
950         .opgFeeFreq_optSemi = False
960         .opgFeeFreq_optAnnual.Enabled = False
970         .opgFeeFreq_optAnnual = False
980         .Schedule_ID.Enabled = False
990         .Schedule_ID.BorderColor = WIN_CLR_DISR
1000        .Schedule_ID.BackStyle = acBackStyleTransparent
1010        .Schedule_ID_lbl.BackStyle = acBackStyleTransparent
1020        .adminno.Enabled = False
1030        .adminno.BorderColor = WIN_CLR_DISR
1040        .adminno.BackStyle = acBackStyleTransparent
1050        .adminno_lbl.BackStyle = acBackStyleTransparent
1060        .investmentobj.Enabled = False
1070        .investmentobj.BorderColor = WIN_CLR_DISR
1080        .investmentobj.BackStyle = acBackStyleTransparent
1090        .investmentobj_lbl.BackStyle = acBackStyleTransparent
1100        .cmbSweep.Enabled = False
1110        .cmbSweep.BorderColor = WIN_CLR_DISR
1120        .cmbSweep.BackStyle = acBackStyleTransparent
1130        .cmbSweep_lbl.BackStyle = acBackStyleTransparent
1140        .Bank_Name.Enabled = False
1150        .Bank_Name.BorderColor = WIN_CLR_DISR
1160        .Bank_Name.BackStyle = acBackStyleTransparent
1170        .Bank_Name_lbl.BackStyle = acBackStyleTransparent
1180        .Bank_City.Enabled = False
1190        .Bank_City.BorderColor = WIN_CLR_DISR
1200        .Bank_City.BackStyle = acBackStyleTransparent
1210        .Bank_City_lbl.BackStyle = acBackStyleTransparent
1220        .opgLoc.Enabled = False
1230        .Bank_State.Enabled = False
1240        .Bank_State.BorderColor = WIN_CLR_DISR
1250        .Bank_State.BackStyle = acBackStyleTransparent
1260        .Bank_State_lbl.BackStyle = acBackStyleTransparent
1270        .Bank_Country.Enabled = False
1280        .Bank_Country.BorderColor = WIN_CLR_DISR
1290        .Bank_Country.BackStyle = acBackStyleTransparent
1300        .Bank_Country_lbl.BackStyle = acBackStyleTransparent
1310        .LastCheckNum.Enabled = False
1320        .LastCheckNum.BorderColor = WIN_CLR_DISR
1330        .LastCheckNum.BackStyle = acBackStyleTransparent
1340        .LastCheckNum_lbl.BackStyle = acBackStyleTransparent
1350        .Bank_RoutingNumber.Enabled = False
1360        .Bank_RoutingNumber.BorderColor = WIN_CLR_DISR
1370        .Bank_RoutingNumber.BackStyle = acBackStyleTransparent
1380        .Bank_RoutingNumber_lbl.BackStyle = acBackStyleTransparent
1390        .Bank_AccountNumber.Enabled = False
1400        .Bank_AccountNumber.BorderColor = WIN_CLR_DISR
1410        .Bank_AccountNumber.BackStyle = acBackStyleTransparent
1420        .Bank_AccountNumber_lbl.BackStyle = acBackStyleTransparent
1430        .FedIFNum1.Enabled = False
1440        .FedIFNum1.BorderColor = WIN_CLR_DISR
1450        .FedIFNum1.BackStyle = acBackStyleTransparent
1460        .FedIFNum1_lbl.BackStyle = acBackStyleTransparent
1470        .FedIFNum2.Enabled = False
1480        .FedIFNum2.BorderColor = WIN_CLR_DISR
1490        .FedIFNum2.BackStyle = acBackStyleTransparent
1500        .cmbAssets.Enabled = False
1510        .cmbAssets.BorderColor = WIN_CLR_DISR
1520        .cmbAssets.BackStyle = acBackStyleTransparent
1530        .cmbAssets_lbl.BackStyle = acBackStyleTransparent
1540        .tabContacts = .tabContacts_pgNotes.PageIndex
1550        .tabContacts_pgContact1.Enabled = False
1560        .tabContacts_pgContact2.Enabled = False
1570        .tabContacts_pgNotes.Enabled = True  ' ** This stays available.
1580        .Contact1_Name.Enabled = False
1590        .Contact1_Name.BorderColor = WIN_CLR_DISR
1600        .Contact1_Name.BackStyle = acBackStyleTransparent
1610        .Contact1_Name_lbl.BackStyle = acBackStyleTransparent
1620        .Contact1_Address1.Enabled = False
1630        .Contact1_Address1.BorderColor = WIN_CLR_DISR
1640        .Contact1_Address1.BackStyle = acBackStyleTransparent
1650        .Contact1_Address1_lbl.BackStyle = acBackStyleTransparent
1660        .Contact1_Address2.Enabled = False
1670        .Contact1_Address2.BorderColor = WIN_CLR_DISR
1680        .Contact1_Address2.BackStyle = acBackStyleTransparent
1690        .Contact1_Address2_lbl.BackStyle = acBackStyleTransparent
1700        .Contact1_City.Enabled = False
1710        .Contact1_City.BorderColor = WIN_CLR_DISR
1720        .Contact1_City.BackStyle = acBackStyleTransparent
1730        .Contact1_City_lbl.BackStyle = acBackStyleTransparent
1740        .opgLocContact1.Enabled = False
1750        .Contact1_State.Enabled = False
1760        .Contact1_State.BorderColor = WIN_CLR_DISR
1770        .Contact1_State.BackStyle = acBackStyleTransparent
1780        .Contact1_State_lbl.BackStyle = acBackStyleTransparent
1790        .Contact1_Zip.Enabled = False
1800        .Contact1_Zip.BorderColor = WIN_CLR_DISR
1810        .Contact1_Zip.BackStyle = acBackStyleTransparent
1820        .Contact1_Zip_lbl.BackStyle = acBackStyleTransparent
1830        .Contact1_Country.Enabled = False
1840        .Contact1_Country.BorderColor = WIN_CLR_DISR
1850        .Contact1_Country.BackStyle = acBackStyleTransparent
1860        .Contact1_Country_lbl.BackStyle = acBackStyleTransparent
1870        .Contact1_PostalCode.Enabled = False
1880        .Contact1_PostalCode.BorderColor = WIN_CLR_DISR
1890        .Contact1_PostalCode.BackStyle = acBackStyleTransparent
1900        .Contact1_PostalCode_lbl.BackStyle = acBackStyleTransparent
1910        .Contact1_Phone1.Enabled = False
1920        .Contact1_Phone1.BorderColor = WIN_CLR_DISR
1930        .Contact1_Phone1.BackStyle = acBackStyleTransparent
1940        .Contact1_Phone1_lbl.BackStyle = acBackStyleTransparent
1950        .Contact1_Phone2.Enabled = False
1960        .Contact1_Phone2.BorderColor = WIN_CLR_DISR
1970        .Contact1_Phone2.BackStyle = acBackStyleTransparent
1980        .Contact1_Phone2_lbl.BackStyle = acBackStyleTransparent
1990        .Contact1_Fax.Enabled = False
2000        .Contact1_Fax.BorderColor = WIN_CLR_DISR
2010        .Contact1_Fax.BackStyle = acBackStyleTransparent
2020        .Contact1_Fax_lbl.BackStyle = acBackStyleTransparent
2030        .Contact1_Email.Enabled = False
2040        .Contact1_Email.BorderColor = WIN_CLR_DISR
2050        .Contact1_Email.BackStyle = acBackStyleTransparent
2060        .Contact1_Email_lbl.BackStyle = acBackStyleTransparent
2070        .Contact2_Name.Enabled = False
2080        .Contact2_Name.BorderColor = WIN_CLR_DISR
2090        .Contact2_Name.BackStyle = acBackStyleTransparent
2100        .Contact2_Name_lbl.BackStyle = acBackStyleTransparent
2110        .Contact2_Address1.Enabled = False
2120        .Contact2_Address1.BorderColor = WIN_CLR_DISR
2130        .Contact2_Address1.BackStyle = acBackStyleTransparent
2140        .Contact2_Address1_lbl.BackStyle = acBackStyleTransparent
2150        .Contact2_Address2.Enabled = False
2160        .Contact2_Address2.BorderColor = WIN_CLR_DISR
2170        .Contact2_Address2.BackStyle = acBackStyleTransparent
2180        .Contact2_Address2_lbl.BackStyle = acBackStyleTransparent
2190        .Contact2_City.Enabled = False
2200        .Contact2_City.BorderColor = WIN_CLR_DISR
2210        .Contact2_City.BackStyle = acBackStyleTransparent
2220        .Contact2_City_lbl.BackStyle = acBackStyleTransparent
2230        .opgLocContact2.Enabled = False
2240        .Contact2_State.Enabled = False
2250        .Contact2_State.BorderColor = WIN_CLR_DISR
2260        .Contact2_State.BackStyle = acBackStyleTransparent
2270        .Contact2_State_lbl.BackStyle = acBackStyleTransparent
2280        .Contact2_Zip.Enabled = False
2290        .Contact2_Zip.BorderColor = WIN_CLR_DISR
2300        .Contact2_Zip.BackStyle = acBackStyleTransparent
2310        .Contact2_Zip_lbl.BackStyle = acBackStyleTransparent
2320        .Contact2_Country.Enabled = False
2330        .Contact2_Country.BorderColor = WIN_CLR_DISR
2340        .Contact2_Country.BackStyle = acBackStyleTransparent
2350        .Contact2_Country_lbl.BackStyle = acBackStyleTransparent
2360        .Contact2_PostalCode.Enabled = False
2370        .Contact2_PostalCode.BorderColor = WIN_CLR_DISR
2380        .Contact2_PostalCode.BackStyle = acBackStyleTransparent
2390        .Contact2_PostalCode_lbl.BackStyle = acBackStyleTransparent
2400        .Contact2_Phone1.Enabled = False
2410        .Contact2_Phone1.BorderColor = WIN_CLR_DISR
2420        .Contact2_Phone1.BackStyle = acBackStyleTransparent
2430        .Contact2_Phone1_lbl.BackStyle = acBackStyleTransparent
2440        .Contact2_Phone2.Enabled = False
2450        .Contact2_Phone2.BorderColor = WIN_CLR_DISR
2460        .Contact2_Phone2.BackStyle = acBackStyleTransparent
2470        .Contact2_Phone2_lbl.BackStyle = acBackStyleTransparent
2480        .Contact2_Fax.Enabled = False
2490        .Contact2_Fax.BorderColor = WIN_CLR_DISR
2500        .Contact2_Fax.BackStyle = acBackStyleTransparent
2510        .Contact2_Fax_lbl.BackStyle = acBackStyleTransparent
2520        .Contact2_Email.Enabled = False
2530        .Contact2_Email.BorderColor = WIN_CLR_DISR
2540        .Contact2_Email.BackStyle = acBackStyleTransparent
2550        .Contact2_Email_lbl.BackStyle = acBackStyleTransparent
2560      Case False
2570        If blnViewOnly = False Then
2580          .accountno.Locked = False
2590        End If
2600        .accountno.ForeColor = CLR_BLK
2610        .accountno.BackColor = CLR_WHT
2620        .related_accountno.Enabled = True
2630        .related_accountno.Locked = True
2640        .related_accountno.Visible = True
2650        .related_accountno.BorderColor = CLR_LTBLU2
2660        .related_accountno.BackStyle = acBackStyleNormal
2670        .related_accountno_lbl.BackStyle = acBackStyleNormal
2680        .cmdChoose.Enabled = True
2690        .cmdChoose_lbl2.ForeColor = CLR_VDKGRY
2700        .cmdChoose_lbl2_dim_hi.Visible = False
2710        If blnViewOnly = False Then
2720          .shortname.Locked = False
2730        End If
2740        .shortname.ForeColor = CLR_BLK
2750        .shortname.BackColor = CLR_WHT
2760        If blnViewOnly = False Then
2770          .legalname.Locked = False
2780        End If
2790        .legalname.ForeColor = CLR_BLK
2800        .legalname.BackColor = CLR_WHT
2810        .CaseNum.Enabled = True
2820        If blnViewOnly = True Then
2830          .CaseNum.Locked = True
2840        End If
2850        .CaseNum.BorderColor = CLR_LTBLU2
2860        .CaseNum.BackStyle = acBackStyleNormal
2870        .CaseNum_lbl.BackStyle = acBackStyleNormal
2880        .accounttype.Enabled = True
2890        If blnViewOnly = False Then
2900          .accounttype.Locked = False
2910        End If
2920        .accounttype.ForeColor = CLR_BLK
2930        .accounttype.BackColor = CLR_WHT
2940        .description.Enabled = False
2950        .description.TabStop = False
2960        .tin.Enabled = True
2970        .tin.BorderColor = CLR_LTBLU2
2980        .tin.BackStyle = acBackStyleNormal
2990        .tin_lbl.BackStyle = acBackStyleNormal
3000        .documentdate.Enabled = True
3010        .documentdate.BorderColor = CLR_LTBLU2
3020        .documentdate.BackStyle = acBackStyleNormal
3030        .documentdate_lbl.BackStyle = acBackStyleNormal
3040        .amendments.Enabled = True
3050        .amendments.BorderColor = CLR_LTBLU2
3060        .amendments.BackStyle = acBackStyleNormal
3070        .amendments_lbl.BackStyle = acBackStyleNormal
3080        .appointmentdate.Enabled = True
3090        .appointmentdate.BorderColor = CLR_LTBLU2
3100        .appointmentdate.BackStyle = acBackStyleNormal
3110        .appointmentdate_lbl.BackStyle = acBackStyleNormal
3120        .cotrustee.Enabled = True
3130        .cotrustee.BorderColor = CLR_LTBLU2
3140        .cotrustee.BackStyle = acBackStyleNormal
3150        .cotrustee_lbl.BackStyle = acBackStyleNormal
3160        .discretion.Enabled = True
3170        .discretion.BorderColor = CLR_LTBLU2
3180        .discretion.BackStyle = acBackStyleNormal
3190        .discretion_lbl.BackStyle = acBackStyleNormal
            ' ** This causes a problem for a new record on some machines.
3200  On Error Resume Next
3210        lngTmp01 = .ActiveAssets
3220        If ERR <> 0 Then
3230  On Error GoTo ERRH
3240          lngTmp01 = 0&
3250        Else
3260  On Error GoTo ERRH
3270        End If
3280        If lngTmp01 > 0 Then
3290          .dateclosed.Enabled = False
3300          .dateclosed.BorderColor = WIN_CLR_DISR
3310          .dateclosed.BackStyle = acBackStyleTransparent
3320          .dateclosed_lbl.BackStyle = acBackStyleTransparent
3330        Else
3340          .dateclosed.Enabled = True
3350          .dateclosed.BorderColor = CLR_LTBLU2
3360          .dateclosed.BackStyle = acBackStyleNormal
3370          .dateclosed_lbl.BackStyle = acBackStyleNormal
3380        End If
3390        If IsNull(.dateclosed) = False Then
3400          .dateclosed.Enabled = True
3410          .dateclosed.BorderColor = CLR_LTBLU2
3420          .dateclosed.BackStyle = acBackStyleNormal
3430          .dateclosed_lbl.BackStyle = acBackStyleNormal
3440        End If
3450        .courtsupervised.Enabled = True
3460        .courtsupervised.BorderColor = CLR_LTBLU2
3470        .courtsupervised.BackStyle = acBackStyleNormal
3480        .courtsupervised_lbl.BackStyle = acBackStyleNormal
3490        .jurisdiction.Enabled = True
3500        .jurisdiction.BorderColor = CLR_LTBLU2
3510        .jurisdiction.BackStyle = acBackStyleNormal
3520        .jurisdiction_lbl.BackStyle = acBackStyleNormal
3530        .county.Enabled = True
3540        .county.BorderColor = CLR_LTBLU2
3550        .county.BackStyle = acBackStyleNormal
3560        .county_lbl.BackStyle = acBackStyleNormal
3570        .cmdStatementFreq.Enabled = True  ' ** This stays available.
3580        .cmdReviewFreq.Enabled = True
3590        .cmdReviewFreq_lbl.BackStyle = acBackStyleNormal
3600        .opgFeeFreq_box.Visible = True
3610        .opgFeeFreq_box_dim.Visible = False
3620        .opgFeeFreq_lbl.ForeColor = CLR_WHT
3630        .opgFeeFreq_lbl.BackStyle = acBackStyleNormal
3640        .opgFeeFreq_lbl_dim_hi.Visible = False
3650        .opgFeeFreq_hline03.Visible = False
3660        .opgFeeFreq_optMonth.Enabled = True
3670        .opgFeeFreq_optMonth = False
3680        .opgFeeFreq_optQuarter.Enabled = True
3690        .opgFeeFreq_optQuarter = False
3700        .opgFeeFreq_optSemi.Enabled = True
3710        .opgFeeFreq_optSemi = False
3720        .opgFeeFreq_optAnnual.Enabled = True
3730        .opgFeeFreq_optAnnual = False
3740        .opgFeeFreq_opt_box.Visible = False
3750        If IsNull(.feeFrequency) = False Then
3760          Select Case .feeFrequency
              Case 1
3770            .opgFeeFreq_optMonth = True
3780            .opgFeeFreq_opt_box.Left = .opgFeeFreq_optMonth_lbl.Left
3790            .opgFeeFreq_opt_box.Visible = True
3800          Case 2
3810            .opgFeeFreq_optQuarter = True
3820            .opgFeeFreq_opt_box.Left = .opgFeeFreq_optQuarter_lbl.Left
3830            .opgFeeFreq_opt_box.Visible = True
3840          Case 3
3850            .opgFeeFreq_optSemi = True
3860            .opgFeeFreq_opt_box.Left = .opgFeeFreq_optSemi_lbl.Left
3870            .opgFeeFreq_opt_box.Visible = True
3880          Case 4
3890            .opgFeeFreq_optAnnual = True
3900            .opgFeeFreq_opt_box.Left = .opgFeeFreq_optAnnual_lbl.Left
3910            .opgFeeFreq_opt_box.Visible = True
3920          End Select
3930        End If
3940        .Schedule_ID.Enabled = True
3950        .Schedule_ID.BorderColor = CLR_LTBLU2
3960        .Schedule_ID.BackStyle = acBackStyleNormal
3970        .Schedule_ID_lbl.BackStyle = acBackStyleNormal
3980        .adminno.Enabled = True
3990        .adminno.BorderColor = CLR_LTBLU2
4000        .adminno.BackStyle = acBackStyleNormal
4010        .adminno_lbl.BackStyle = acBackStyleNormal
4020        .investmentobj.Enabled = True
4030        .investmentobj.BorderColor = CLR_LTBLU2
4040        .investmentobj.BackStyle = acBackStyleNormal
4050        .investmentobj_lbl.BackStyle = acBackStyleNormal
4060        .cmbSweep.Enabled = True
4070        .cmbSweep.BorderColor = CLR_LTBLU2
4080        .cmbSweep.BackStyle = acBackStyleNormal
4090        .cmbSweep_lbl.BackStyle = acBackStyleNormal
4100        Select Case .account_SWEEP
            Case True
4110          .cmbSweep.SetFocus
4120          Select Case blnViewOnly
              Case True
4130            .cmbSweep.Enabled = True
4140            .cmbSweep.Value = "Yes"
4150            .cmbSweep.text = "Yes"
4160            .cmbSweep.Enabled = False
4170            .cmbSweep.BorderColor = WIN_CLR_DISR
4180            .cmbSweep.BackStyle = acBackStyleTransparent
4190            .cmbSweep_lbl.BackStyle = acBackStyleTransparent
4200          Case False
4210            .cmbSweep.Value = "Yes"
4220            .cmbSweep.text = "Yes"
4230          End Select
4240        Case False
4250          .cmbSweep.SetFocus
4260          Select Case blnViewOnly
              Case True
4270            .cmbSweep.Enabled = True
4280            .cmbSweep.Value = "No"
4290            .cmbSweep.text = "No"
4300            .cmbSweep.Enabled = False
4310            .cmbSweep.BorderColor = WIN_CLR_DISR
4320            .cmbSweep.BackStyle = acBackStyleTransparent
4330            .cmbSweep_lbl.BackStyle = acBackStyleTransparent
4340          Case False
4350            .cmbSweep.Value = "No"
4360            .cmbSweep.text = "No"
4370          End Select
4380        End Select
4390        .Bank_Name.Enabled = True
4400        .Bank_Name.BorderColor = CLR_LTBLU2
4410        .Bank_Name.BackStyle = acBackStyleNormal
4420        .Bank_Name_lbl.BackStyle = acBackStyleNormal
4430        .Bank_City.Enabled = True
4440        .Bank_City.BorderColor = CLR_LTBLU2
4450        .Bank_City.BackStyle = acBackStyleNormal
4460        .Bank_City_lbl.BackStyle = acBackStyleNormal
4470        .opgLoc.Enabled = True
4480        Select Case .opgLoc
            Case .opgLoc_optUSA.OptionValue
4490          .Bank_State.Enabled = True
4500          .Bank_State.BorderColor = CLR_LTBLU2
4510          .Bank_State.BackStyle = acBackStyleNormal
4520          .Bank_State_lbl.BackStyle = acBackStyleNormal
4530          .Bank_Country.Enabled = False
4540          .Bank_Country.BorderColor = WIN_CLR_DISR
4550          .Bank_Country.BackStyle = acBackStyleTransparent
4560          .Bank_Country_lbl.BackStyle = acBackStyleTransparent
4570        Case .opgLoc_optOther.OptionValue
4580          .Bank_State.Enabled = False
4590          .Bank_State.BorderColor = WIN_CLR_DISR
4600          .Bank_State.BackStyle = acBackStyleTransparent
4610          .Bank_State_lbl.BackStyle = acBackStyleTransparent
4620          .Bank_Country.Enabled = True
4630          .Bank_Country.BorderColor = CLR_LTBLU2
4640          .Bank_Country.BackStyle = acBackStyleNormal
4650          .Bank_Country_lbl.BackStyle = acBackStyleNormal
4660        End Select
4670        .LastCheckNum.Enabled = True
4680        .LastCheckNum.BorderColor = CLR_LTBLU2
4690        .LastCheckNum.BackStyle = acBackStyleNormal
4700        .LastCheckNum_lbl.BackStyle = acBackStyleNormal
4710        .Bank_RoutingNumber.Enabled = True
4720        .Bank_RoutingNumber.BorderColor = CLR_LTBLU2
4730        .Bank_RoutingNumber.BackStyle = acBackStyleNormal
4740        .Bank_RoutingNumber_lbl.BackStyle = acBackStyleNormal
4750        .Bank_AccountNumber.Enabled = True
4760        .Bank_AccountNumber.BorderColor = CLR_LTBLU2
4770        .Bank_AccountNumber.BackStyle = acBackStyleNormal
4780        .Bank_AccountNumber_lbl.BackStyle = acBackStyleNormal
4790        .FedIFNum1.Enabled = True
4800        .FedIFNum1.BorderColor = CLR_LTBLU2
4810        .FedIFNum1.BackStyle = acBackStyleNormal
4820        .FedIFNum1_lbl.BackStyle = acBackStyleNormal
4830        .FedIFNum2.Enabled = True
4840        .FedIFNum2.BackStyle = acBackStyleNormal
4850        .FedIFNum2.BorderColor = CLR_LTBLU2
4860        .cmbAssets.Enabled = True
4870        .cmbAssets.BorderColor = CLR_LTBLU2
4880        .cmbAssets.BackStyle = acBackStyleNormal
4890        .cmbAssets_lbl.BackStyle = acBackStyleNormal
4900        .tabContacts_pgContact1.Enabled = True
4910        .tabContacts_pgContact2.Enabled = True
4920        .tabContacts_pgNotes.Enabled = True  ' ** This stays available.
4930        .tabContacts = .tabContacts_pgContact1.PageIndex
4940        .Contact1_Name.Enabled = True
4950        .Contact1_Name.BorderColor = CLR_LTBLU2
4960        .Contact1_Name.BackStyle = acBackStyleNormal
4970        .Contact1_Name_lbl.BackStyle = acBackStyleNormal
4980        .Contact1_Address1.Enabled = True
4990        .Contact1_Address1.BorderColor = CLR_LTBLU2
5000        .Contact1_Address1.BackStyle = acBackStyleNormal
5010        .Contact1_Address1_lbl.BackStyle = acBackStyleNormal
5020        .Contact1_Address2.Enabled = True
5030        .Contact1_Address2.BorderColor = CLR_LTBLU2
5040        .Contact1_Address2.BackStyle = acBackStyleNormal
5050        .Contact1_Address2_lbl.BackStyle = acBackStyleNormal
5060        .Contact1_City.Enabled = True
5070        .Contact1_City.BorderColor = CLR_LTBLU2
5080        .Contact1_City.BackStyle = acBackStyleNormal
5090        .Contact1_City_lbl.BackStyle = acBackStyleNormal
5100        .opgLocContact1.Enabled = True
5110        Select Case .opgLocContact1
            Case .opgLocContact1_optUSA.OptionValue
5120          .Contact1_State.Enabled = True
5130          .Contact1_State.BorderColor = CLR_LTBLU2
5140          .Contact1_State.BackStyle = acBackStyleNormal
5150          .Contact1_State_lbl.BackStyle = acBackStyleNormal
5160          .Contact1_Zip.Enabled = True
5170          .Contact1_Zip.BorderColor = CLR_LTBLU2
5180          .Contact1_Zip.BackStyle = acBackStyleNormal
5190          .Contact1_Zip_lbl.BackStyle = acBackStyleNormal
5200          .Contact1_Country.Enabled = False
5210          .Contact1_Country.BorderColor = WIN_CLR_DISR
5220          .Contact1_Country.BackStyle = acBackStyleTransparent
5230          .Contact1_Country_lbl.BackStyle = acBackStyleTransparent
5240          .Contact1_PostalCode.Enabled = False
5250          .Contact1_PostalCode.BorderColor = WIN_CLR_DISR
5260          .Contact1_PostalCode.BackStyle = acBackStyleTransparent
5270          .Contact1_PostalCode_lbl.BackStyle = acBackStyleTransparent
5280        Case .opgLocContact1_optOther.OptionValue
5290          .Contact1_State.Enabled = False
5300          .Contact1_State.BorderColor = WIN_CLR_DISR
5310          .Contact1_State.BackStyle = acBackStyleTransparent
5320          .Contact1_State_lbl.BackStyle = acBackStyleTransparent
5330          .Contact1_Zip.Enabled = False
5340          .Contact1_Zip.BorderColor = WIN_CLR_DISR
5350          .Contact1_Zip.BackStyle = acBackStyleTransparent
5360          .Contact1_Zip_lbl.BackStyle = acBackStyleTransparent
5370          .Contact1_Country.Enabled = True
5380          .Contact1_Country.BorderColor = CLR_LTBLU2
5390          .Contact1_Country.BackStyle = acBackStyleNormal
5400          .Contact1_Country_lbl.BackStyle = acBackStyleNormal
5410          .Contact1_PostalCode.Enabled = True
5420          .Contact1_PostalCode.BorderColor = CLR_LTBLU2
5430          .Contact1_PostalCode.BackStyle = acBackStyleNormal
5440          .Contact1_PostalCode_lbl.BackStyle = acBackStyleNormal
5450        End Select
5460        .Contact1_Phone1.Enabled = True
5470        .Contact1_Phone1.BorderColor = CLR_LTBLU2
5480        .Contact1_Phone1.BackStyle = acBackStyleNormal
5490        .Contact1_Phone1_lbl.BackStyle = acBackStyleNormal
5500        .Contact1_Phone2.Enabled = True
5510        .Contact1_Phone2.BorderColor = CLR_LTBLU2
5520        .Contact1_Phone2.BackStyle = acBackStyleNormal
5530        .Contact1_Phone2_lbl.BackStyle = acBackStyleNormal
5540        .Contact1_Fax.Enabled = True
5550        .Contact1_Fax.BorderColor = CLR_LTBLU2
5560        .Contact1_Fax.BackStyle = acBackStyleNormal
5570        .Contact1_Fax_lbl.BackStyle = acBackStyleNormal
5580        .Contact1_Email.Enabled = True
5590        .Contact1_Email.BorderColor = CLR_LTBLU2
5600        .Contact1_Email.BackStyle = acBackStyleNormal
5610        .Contact1_Email_lbl.BackStyle = acBackStyleNormal
5620        .Contact2_Name.Enabled = True
5630        .Contact2_Name.BorderColor = CLR_LTBLU2
5640        .Contact2_Name.BackStyle = acBackStyleNormal
5650        .Contact2_Name_lbl.BackStyle = acBackStyleNormal
5660        .Contact2_Address1.Enabled = True
5670        .Contact2_Address1.BorderColor = CLR_LTBLU2
5680        .Contact2_Address1.BackStyle = acBackStyleNormal
5690        .Contact2_Address1_lbl.BackStyle = acBackStyleNormal
5700        .Contact2_Address2.Enabled = True
5710        .Contact2_Address2.BorderColor = CLR_LTBLU2
5720        .Contact2_Address2.BackStyle = acBackStyleNormal
5730        .Contact2_Address2_lbl.BackStyle = acBackStyleNormal
5740        .Contact2_City.Enabled = True
5750        .Contact2_City.BorderColor = CLR_LTBLU2
5760        .Contact2_City.BackStyle = acBackStyleNormal
5770        .Contact2_City_lbl.BackStyle = acBackStyleNormal
5780        .opgLocContact2.Enabled = True
5790        Select Case .opgLocContact1
            Case .opgLocContact1_optUSA.OptionValue
5800          .Contact2_State.Enabled = True
5810          .Contact2_State.BorderColor = CLR_LTBLU2
5820          .Contact2_State.BackStyle = acBackStyleNormal
5830          .Contact2_State_lbl.BackStyle = acBackStyleNormal
5840          .Contact2_Zip.Enabled = True
5850          .Contact2_Zip.BorderColor = CLR_LTBLU2
5860          .Contact2_Zip.BackStyle = acBackStyleNormal
5870          .Contact2_Zip_lbl.BackStyle = acBackStyleNormal
5880          .Contact2_Country.Enabled = False
5890          .Contact2_Country.BorderColor = WIN_CLR_DISR
5900          .Contact2_Country.BackStyle = acBackStyleTransparent
5910          .Contact2_Country_lbl.BackStyle = acBackStyleTransparent
5920          .Contact2_PostalCode.Enabled = False
5930          .Contact2_PostalCode.BorderColor = WIN_CLR_DISR
5940          .Contact2_PostalCode.BackStyle = acBackStyleTransparent
5950          .Contact2_PostalCode_lbl.BackStyle = acBackStyleTransparent
5960        Case .opgLocContact1_optOther.OptionValue
5970          .Contact2_State.Enabled = False
5980          .Contact2_State.BorderColor = WIN_CLR_DISR
5990          .Contact2_State.BackStyle = acBackStyleTransparent
6000          .Contact2_State_lbl.BackStyle = acBackStyleTransparent
6010          .Contact2_Zip.Enabled = False
6020          .Contact2_Zip.BorderColor = WIN_CLR_DISR
6030          .Contact2_Zip.BackStyle = acBackStyleTransparent
6040          .Contact2_Zip_lbl.BackStyle = acBackStyleTransparent
6050          .Contact2_Country.Enabled = True
6060          .Contact2_Country.BorderColor = CLR_LTBLU2
6070          .Contact2_Country.BackStyle = acBackStyleNormal
6080          .Contact2_Country_lbl.BackStyle = acBackStyleNormal
6090          .Contact2_PostalCode.Enabled = True
6100          .Contact2_PostalCode.BorderColor = CLR_LTBLU2
6110          .Contact2_PostalCode.BackStyle = acBackStyleNormal
6120          .Contact2_PostalCode_lbl.BackStyle = acBackStyleNormal
6130        End Select
6140        .Contact2_Phone1.Enabled = True
6150        .Contact2_Phone1.BorderColor = CLR_LTBLU2
6160        .Contact2_Phone1.BackStyle = acBackStyleNormal
6170        .Contact2_Phone1_lbl.BackStyle = acBackStyleNormal
6180        .Contact2_Phone2.Enabled = True
6190        .Contact2_Phone2.BorderColor = CLR_LTBLU2
6200        .Contact2_Phone2.BackStyle = acBackStyleNormal
6210        .Contact2_Phone2_lbl.BackStyle = acBackStyleNormal
6220        .Contact2_Fax.Enabled = True
6230        .Contact2_Fax.BorderColor = CLR_LTBLU2
6240        .Contact2_Fax.BackStyle = acBackStyleNormal
6250        .Contact2_Fax_lbl.BackStyle = acBackStyleNormal
6260        .Contact2_Email.Enabled = True
6270        .Contact2_Email.BorderColor = CLR_LTBLU2
6280        .Contact2_Email.BackStyle = acBackStyleNormal
6290        .Contact2_Email_lbl.BackStyle = acBackStyleNormal
6300      End Select
6310    End With

EXITP:
6320    Exit Sub

ERRH:
6330    Select Case ERR.Number
        Case 2135  ' ** This property is read-only and can't be set.
          ' ** I think I've handled this with the blnViewOnly variable.
6340    Case Else
6350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6360    End Select
6370    Resume EXITP

End Sub
