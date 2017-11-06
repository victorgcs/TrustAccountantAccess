VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountContacts"

'VGC 09/23/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   AcctNoShort_Set()
' ##   EnableCountry_SetFrmWidth_AC()
' #######################################

' #########################
' ## Use VBA_RenumErrh().  39030
' #########################

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd}
' **   Enable Country:   B {chkEnableCountry_Wide, chkEnableCountry_Compact}
' **   Close:            C {cmdClose}
' **   View Wide:        D {opgView_optWide}
' **   Reset:            E {cmdReset}
' **   Excel:            L {cmdExcel}
' **   View Compact:     M {opgView_optCompact}
' **   Account Num:      N {chkShowAcctNum}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Short Name:       S {chkShowShortName}
' **   Delete:           T {cmdDelete}
' **   Update:           U {cmdUpdate}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}
' **   Always Update:    Y {chkAlwaysUpdate}

' ** Shortcut Ctrl keys responsive from this form:
' **   Send Email:       E {Contact_Email_cmd on frmAccountContacts_Sub2}
' **   Page Count:       G {chkPageOf}
' **   Other:            O {opgLocContact_optOther on frmAccountContacts_Sub, frmAccountContacts_Sub2}
' **   Save:             S {cmdSave on frmAccountContacts_Sub}
' **   USA:              U {opgLocContact_optUSA on frmAccountContacts_Sub, frmAccountContacts_Sub2}

' ******************************************
' ** opgFormat and opgPrint shortcut keys:
' ******************************************

' ** Shortcut Ctrl keys responsive from this form:
' **   Zip Unfrmat:      D {opgFormatZip_optUnformatted}
' **   Zip Frmat:        F {opgFormatZip_optFormatted}
' **   Name Last/First:  I {opgFormatName_optLastFirst}
' **   Phone Unfrmat:    M {opgFormatPhone_optUnformatted}
' **   Phone Frmat:      R {opgFormatPhone_optFormatted}
' **   Name As Written:  T {opgFormatName_optAsWritten}

' ** Shortcut Ctrl-Alt keys responsive from this form:
' **   CSZ Separate:     A {opgPrintCSZ_optSeparate, opgPrintCSZCP_optSeparate}
' **   CSZ Combined:     B {opgPrintCSZ_optCombined, opgPrintCSZCP_optCombined}
' **   Addr Combined:    C {opgPrintAddress_optCombined}
' **   Addr None:        N {opgPrintAddress_optNone}
' **   CSZ None:         O {opgPrintCSZ_optNone, opgPrintCSZCP_optNone}
' **   Addr Separate:    S {opgPrintAddress_optSeparate}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Phone Combined:   C {opgPrintPhone_optCombined}
' **   Email None:       E {opgPrintEmail_optNone}
' **   Phone None:       N {opgPrintPhone_optNone}
' **   Fax None:         O {opgPrintFax_optNone}
' **   Fax Print:        P {opgPrintFax_optPrint}
' **   Phone Separate:   S {opgPrintPhone_optSeparate}
' **   Email Print:      T {opgPrintEmail_optPrint}
' **   Phone 1 Only:     Y {opgPrintPhone_opt1Only}

' ** Array: arr_varCtl().
Private lngCtls As Long, arr_varCtl As Variant
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

Private blnReset_Focus As Boolean, blnReset_MouseDown As Boolean

Private lngFrm_Left_Orig As Long, lngFrm_Top_Orig As Long, lngFrm_Width_Orig As Long, lngFrm_Height_Orig As Long
Private lngForm_Width As Long, lngWinSub_Diff As Long
Private blnFromClose As Boolean, blnCountry As Boolean
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim frmSub1 As Access.Form, frmSub2 As Access.Form
        Dim blnRetVal As Boolean
        Dim lngTmp01 As Long

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True: blnFromClose = False

150           strCallingForm = Trim(.OpenArgs)

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170           blnReset_Focus = False: blnReset_MouseDown = False: blnCountry = False

              ' ** Variables are fed empty, then populated ByRef.
180           GetFormDimensions Me, lngFrm_Left_Orig, lngFrm_Top_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig  ' ** Module Function: modWindowFunctions.

190           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
200           lngMonitorNum = 1&: lngTmp01 = 0&
210           EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
220           If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
230           If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top_Orig
240           DoCmd.SelectObject acForm, THIS_NAME, False
250           DoEvents
260           DoCmd.MoveSize lngFrm_Left_Orig, lngTmp01, lngFrm_Width_Orig, lngFrm_Height_Orig  'lngFrm_Top_Orig
270           FrmDimSave THIS_NAME, lngFrm_Left_Orig, lngFrm_Top_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig  ' ** Module Function: modMonitorFuncs.
280           If lngMonitorNum > 1& Then
290             LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
300           End If

              ' ** Just get these out of the way.
310           .form_width_line.Left = 0&
320           .detail_height_line.Top = 0&
330           .cmdReset_lbl.Top = 0&
340           .cmdReset_lbl.Left = 0&

350           Set frmSub1 = .frmAccountContacts_Sub.Form
360           Set frmSub2 = .frmAccountContacts_Sub2.Form

370           CtlArrayLoad Me  ' ** Module Procedure: modAccountContacts.

              ' ** Form opens at the accountno-only width.
380           lngWinSub_Diff = (lngFrm_Width_Orig - .frmAccountContacts_Sub.Width)
390           lngForm_Width = .Width

400           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
410           DoEvents

420           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
430           SetNav_Access2007 Me  ' ** Module Procedure: modXAccess_07_10_Funcs.
              'SystemColor_Move Me  ' ** Module Procedure: modSystemColorFuncs.

              ' ** Save the preferences.
440           .chkShowAcctNum_Last = .chkShowAcctNum
450           .chkShowShortName_Last = .chkShowShortName
460           .chkEnableCountry_Wide_Last = .chkEnableCountry_Wide
470           .chkEnableCountry_Compact_Last = .chkEnableCountry_Compact

480           chkShowAcctNum_AfterUpdate  ' ** Procedure: Below.
490           chkShowShortName_AfterUpdate  ' ** Procedure: Below.
500           chkEnableCountry_Wide_AfterUpdate  ' ** Procedure: Below.
510           chkEnableCountry_Compact_AfterUpdate  ' ** Procedure: Below.  ' ** This sets it on opening.
520           chkPageOf_AfterUpdate  ' ** Procedure: Below.
530           chkAlwaysUpdate_AfterUpdate  ' ** Procedure: Below.

540           opgView_AfterUpdate  ' ** Procedure: Below.

550           opgFormatName_AfterUpdate  ' ** Procedure: Below.
560           opgFormatZip_AfterUpdate  ' ** Procedure: Below.
570           opgFormatPhone_AfterUpdate  ' ** Procedure: Below.
580           opgPrintAddress_AfterUpdate  ' ** Procedure: Below.
590           opgPrintCSZ_AfterUpdate  ' ** Procedure: Below.
600           opgPrintCSZCP_AfterUpdate  ' ** Procedure: Below.
610           opgPrintPhone_AfterUpdate  ' ** Procedure: Below.
620           opgPrintFax_AfterUpdate  ' ** Procedure: Below.
630           opgPrintEmail_AfterUpdate  ' ** Procedure: Below.

640           .FocusHolder.SetFocus

650           blnIsOpen = False

660         Else
670           Cancel = -1
680         End If
690       Else
700         Cancel = -1
710       End If
720     End With

730     If Cancel = -1 Then
740       Beep
750       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
760     End If

EXITP:
770     Set frmSub1 = Nothing
780     Set frmSub2 = Nothing
790     Exit Sub

ERRH:
800     DoCmd.Hourglass False
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub Form_Load()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim frmSub1 As Access.Form, frmSub2 As Access.Form
        Dim varTmp00 As Variant

810     With Me

820       Set dbs = CurrentDb
          ' ** Empty tmpXAdmin_Account_02.
830       Set qdf = dbs.QueryDefs("qryAccount_Contacts_18_01")
840       qdf.Execute
850       Set qdf = Nothing
860       DoEvents
          ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
          ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
          ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
          ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
          ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
          ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
          ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
          ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpXAdmin_Account_02.
870       Set qdf = dbs.QueryDefs("qryAccount_Contacts_18_02")
880       qdf.Execute
890       Set qdf = Nothing
900       DoEvents
910       dbs.Close
920       Set dbs = Nothing
930       DoEvents

940       Set frmSub1 = .frmAccountContacts_Sub.Form
950       Set frmSub2 = .frmAccountContacts_Sub2.Form
960       frmSub1.SortNow THIS_PROC   ' ** Form Procedure: frmAccountContacts_Sub.
970       frmSub2.SortNow THIS_PROC   ' ** Form Procedure: frmAccountContacts_Sub2.
980       DoEvents
          ' ** tblAccount_Contact_Staging, linked to Account.
990       varTmp00 = DCount("*", "qryAccount_Contacts_01_01")
1000      If IsNull(varTmp00) = True Then
1010        lngRecsCur = 0&
1020      Else
1030        lngRecsCur = varTmp00
1040      End If
1050      Select Case .opgView
          Case .opgView_optCompact.OptionValue
1060        blnCountry = .chkEnableCountry_Compact
1070      Case .opgView_optWide.OptionValue
1080        blnCountry = .chkEnableCountry_Wide
1090      End Select
1100      If lngRecsCur > 0& Then
1110        .cmdPreview.Enabled = True
1120        .cmdPrint.Enabled = True
1130        .cmdWord.Enabled = True
      #If NoExcel Then
1140        .cmdExcel.Enabled = False
      #Else
1150        .cmdExcel.Enabled = True
      #End If
1160        .cmdDelete.Enabled = True
1170        .cmdUpdate.Enabled = True
1180        .opgFormat_lbl.ForeColor = CLR_VDKGRY
1190        .opgFormat_lbl_dim_hi.Visible = False
1200        .opgFormatName.Enabled = True
1210        .opgFormatZip.Enabled = True
1220        .opgFormatPhone.Enabled = True
1230        .opgPrint_lbl.ForeColor = CLR_VDKGRY
1240        .opgPrint_lbl_dim_hi.Visible = False
1250        .opgPrintAddress.Enabled = True
1260        .opgPrintCSZ.Enabled = True
1270        .opgPrintCSZCP.Enabled = True
1280        .opgPrintPhone.Enabled = True
1290        .opgPrintFax.Enabled = True
1300        .opgPrintEmail.Enabled = True
1310        .chkShowAcctNum.Enabled = True
1320        .chkShowShortName.Enabled = True
1330        .chkEnableCountry_Wide.Enabled = True
1340        .chkEnableCountry_Compact.Enabled = True
1350        .chkPageOf.Enabled = True
1360        .chkAlwaysUpdate.Enabled = True
1370        .chkAlwaysUpdate_lbl2.ForeColor = CLR_VDKGRY
1380        .chkAlwaysUpdate_lbl2_dim_hi.Visible = False
1390        With frmSub1
1400          .accountno_lbl.ForeColor = CLR_DKGRY2
1410          .accountno_lbl_dim_hi.Visible = False
1420          .accountno_lbl_line.BorderColor = CLR_DKGRY
1430          .accountno_lbl_line_dim_hi.Visible = False
1440          .shortname_lbl.ForeColor = CLR_DKGRY2
1450          .shortname_lbl_dim_hi.Visible = False
1460          .shortname_lbl_line.BorderColor = CLR_DKGRY
1470          .shortname_lbl_Line_dim_hi.Visible = False
1480          .Contact_Name_lbl.ForeColor = CLR_DKGRY2
1490          .Contact_Name_lbl_dim_hi.Visible = False
1500          .Contact_Name_lbl_line.BorderColor = CLR_DKGRY
1510          .Contact_Name_lbl_line_dim_hi.Visible = False
1520          .Contact_Address1_lbl.ForeColor = CLR_DKGRY2
1530          .Contact_Address1_lbl_dim_hi.Visible = False
1540          .Contact_Address1_lbl_line.BorderColor = CLR_DKGRY
1550          .Contact_Address1_lbl_line_dim_hi.Visible = False
1560          .Contact_Address2_lbl.ForeColor = CLR_DKGRY2
1570          .Contact_Address2_lbl_dim_hi.Visible = False
1580          .Contact_Address2_lbl_line.BorderColor = CLR_DKGRY
1590          .Contact_Address2_lbl_line_dim_hi.Visible = False
1600          .Contact_City_lbl.ForeColor = CLR_DKGRY2
1610          .Contact_City_lbl_dim_hi.Visible = False
1620          .Contact_City_lbl_line.BorderColor = CLR_DKGRY
1630          .Contact_City_lbl_line_dim_hi.Visible = False
1640          .opgLocContact_lbl.ForeColor = CLR_DKGRY2
1650          .opgLocContact_lbl_dim_hi.Visible = False
1660          .opgLocContact_lbl_line.BorderColor = CLR_DKGRY
1670          .opgLocContact_lbl_line_dim_hi.Visible = False
1680          .Contact_State_lbl.ForeColor = CLR_DKGRY2
1690          .Contact_State_lbl_dim_hi.Visible = False
1700          .Contact_State_lbl_line.BorderColor = CLR_DKGRY
1710          .Contact_State_lbl_line_dim_hi.Visible = False
1720          .Contact_Zip_lbl.ForeColor = CLR_DKGRY2
1730          .Contact_Zip_lbl_dim_hi.Visible = False
1740          .Contact_Zip_lbl_line.BorderColor = CLR_DKGRY
1750          .Contact_Zip_lbl_line_dim_hi.Visible = False
1760          .Contact_Country_lbl.ForeColor = CLR_DKGRY2
1770          .Contact_Country_lbl_dim_hi.Visible = False
1780          .Contact_Country_lbl_line.BorderColor = CLR_DKGRY
1790          .Contact_Country_lbl_line_dim_hi.Visible = False
1800          .Contact_PostalCode_lbl.ForeColor = CLR_DKGRY2
1810          .Contact_PostalCode_lbl_dim_hi.Visible = False
1820          .Contact_PostalCode_lbl_line.BorderColor = CLR_DKGRY
1830          .Contact_PostalCode_lbl_line_dim_hi.Visible = False
1840          .Contact_Phone1_lbl.ForeColor = CLR_DKGRY2
1850          .Contact_Phone1_lbl_dim_hi.Visible = False
1860          .Contact_Phone1_lbl_line.BorderColor = CLR_DKGRY
1870          .Contact_Phone1_lbl_line_dim_hi.Visible = False
1880          .Contact_Phone2_lbl.ForeColor = CLR_DKGRY2
1890          .Contact_Phone2_lbl_dim_hi.Visible = False
1900          .Contact_Phone2_lbl_line.BorderColor = CLR_DKGRY
1910          .Contact_Phone2_lbl_line_dim_hi.Visible = False
1920          .Contact_Fax_lbl.ForeColor = CLR_DKGRY2
1930          .Contact_Fax_lbl_dim_hi.Visible = False
1940          .Contact_Fax_lbl_line.BorderColor = CLR_DKGRY
1950          .Contact_Fax_lbl_line_dim_hi.Visible = False
1960          .Contact_Email_lbl.ForeColor = CLR_DKGRY2
1970          .Contact_Email_lbl_dim_hi.Visible = False
1980          .Contact_Email_lbl_line.BorderColor = CLR_DKGRY
1990          .Contact_Email_lbl_line_dim_hi.Visible = False
2000        End With
2010        With frmSub2
2020          .accountno1_lbl.ForeColor = CLR_DKGRY2
2030          .accountno1_lbl_dim_hi.Visible = False
2040          .accountno1_lbl_line.BorderColor = CLR_DKGRY
2050          .accountno1_lbl_line_dim_hi.Visible = False
2060          .Contact_Number_lbl.ForeColor = CLR_DKGRY2
2070          .Contact_Number_lbl_dim_hi.Visible = False
2080          .Contact_Number_lbl_line.BorderColor = CLR_DKGRY
2090          .Contact_Number_lbl_line_dim_hi.Visible = False
2100          .Contact_Name_lbl.ForeColor = CLR_DKGRY2
2110          .Contact_Name_lbl_dim_hi.Visible = False
2120          .Contact_Name_lbl_line.BorderColor = CLR_DKGRY
2130          .Contact_Name_lbl_line_dim_hi.Visible = False
2140          .Contact_Address1_lbl.ForeColor = CLR_DKGRY2
2150          .Contact_Address1_lbl_dim_hi.Visible = False
2160          .Contact_Address1_lbl_line.BorderColor = CLR_DKGRY
2170          .Contact_Address1_lbl_line_dim_hi.Visible = False
2180          .opgLocContact_lbl.ForeColor = CLR_DKGRY2
2190          .opgLocContact_lbl_dim_hi.Visible = False
2200          .opgLocContact_lbl2.ForeColor = CLR_DKGRY2
2210          .opgLocContact_lbl2_dim_hi.Visible = False
2220          .opgLocContact_lbl_line.BorderColor = CLR_DKGRY
2230          .opgLocContact_lbl_line_dim_hi.Visible = False
2240          .Contact_City_lbl.ForeColor = CLR_DKGRY2
2250          .Contact_City_lbl_dim_hi.Visible = False
2260          .Contact_City_lbl_line.BorderColor = CLR_DKGRY
2270          .Contact_City_lbl_line_dim_hi.Visible = False
2280          .Contact_Country_lbl.ForeColor = CLR_DKGRY2
2290          .Contact_Country_lbl_dim_hi.Visible = False
2300          .Contact_Country_lbl_line.BorderColor = CLR_DKGRY
2310          .Contact_Country_lbl_line_dim_hi.Visible = False
2320          .Contact_Phone1_lbl.ForeColor = CLR_DKGRY2
2330          .Contact_Phone1_lbl_dim_hi.Visible = False
2340          .Contact_Phone1_lbl_line.BorderColor = CLR_DKGRY
2350          .Contact_Phone1_lbl_line_dim_hi.Visible = False
2360          .Contact_Email_lbl.ForeColor = CLR_DKGRY2
2370          .Contact_Email_lbl_dim_hi.Visible = False
2380          .Contact_Email_lbl_line.BorderColor = CLR_DKGRY
2390          .Contact_Email_lbl_line_dim_hi.Visible = False
2400        End With
2410      Else  ' ** lngRecsCur = 0&
2420        .cmdAdd.SetFocus
2430        .cmdPreview.Enabled = False
2440        .cmdPrint.Enabled = False
2450        .cmdWord.Enabled = False
      #If NoExcel Then
2460        .cmdExcel.Enabled = False
      #Else
2470        .cmdExcel.Enabled = False
      #End If
2480        .cmdDelete.Enabled = False
2490        .cmdUpdate.Enabled = False
2500        .opgFormat_lbl.ForeColor = WIN_CLR_DISF
2510        .opgFormat_lbl_dim_hi.Visible = True
2520        .opgFormatName.Enabled = False
2530        .opgFormatZip.Enabled = False
2540        .opgFormatPhone.Enabled = False
2550        .opgPrint_lbl.ForeColor = WIN_CLR_DISF
2560        .opgPrint_lbl_dim_hi.Visible = True
2570        .opgPrintAddress.Enabled = False
2580        .opgPrintCSZ.Enabled = False
2590        .opgPrintCSZCP.Enabled = False
2600        .opgPrintPhone.Enabled = False
2610        .opgPrintFax.Enabled = False
2620        .opgPrintEmail.Enabled = False
2630        .cmdReset.Enabled = False
2640        .cmdReset_raised_img_dis.Visible = True
2650        .cmdReset_raised_img.Visible = False
2660        .cmdReset_raised_semifocus_dots_img.Visible = False
2670        .cmdReset_raised_focus_img.Visible = False
2680        .cmdReset_raised_focus_dots_img.Visible = False
2690        .cmdReset_sunken_focus_dots_img.Visible = False
2700        .chkShowAcctNum.Enabled = False
2710        .chkShowShortName.Enabled = False
2720        .chkEnableCountry_Wide.Enabled = False
2730        .chkEnableCountry_Compact.Enabled = False
2740        .chkPageOf.Enabled = False
2750        .chkAlwaysUpdate.Enabled = False
2760        .chkAlwaysUpdate_lbl2.ForeColor = WIN_CLR_DISF
2770        .chkAlwaysUpdate_lbl2_dim_hi.Visible = True
2780        With frmSub1
2790          .accountno_lbl.ForeColor = WIN_CLR_DISF
2800          .accountno_lbl_dim_hi.Visible = True
2810          .accountno_lbl_line.BorderColor = WIN_CLR_DISR
2820          .accountno_lbl_line_dim_hi.Visible = True
2830          .shortname_lbl.ForeColor = WIN_CLR_DISF
2840          .shortname_lbl_dim_hi.Visible = True
2850          .shortname_lbl_line.BorderColor = WIN_CLR_DISR
2860          .shortname_lbl_Line_dim_hi.Visible = True
2870          .Contact_Name_lbl.ForeColor = WIN_CLR_DISF
2880          .Contact_Name_lbl_dim_hi.Visible = True
2890          .Contact_Name_lbl_line.BorderColor = WIN_CLR_DISR
2900          .Contact_Name_lbl_line_dim_hi.Visible = True
2910          .Contact_Address1_lbl.ForeColor = WIN_CLR_DISF
2920          .Contact_Address1_lbl_dim_hi.Visible = True
2930          .Contact_Address1_lbl_line.BorderColor = WIN_CLR_DISR
2940          .Contact_Address1_lbl_line_dim_hi.Visible = True
2950          .Contact_Address2_lbl.ForeColor = WIN_CLR_DISF
2960          .Contact_Address2_lbl_dim_hi.Visible = True
2970          .Contact_Address2_lbl_line.BorderColor = WIN_CLR_DISR
2980          .Contact_Address2_lbl_line_dim_hi.Visible = True
2990          .Contact_City_lbl.ForeColor = WIN_CLR_DISF
3000          .Contact_City_lbl_dim_hi.Visible = True
3010          .Contact_City_lbl_line.BorderColor = WIN_CLR_DISR
3020          .Contact_City_lbl_line_dim_hi.Visible = True
3030          Select Case blnCountry
              Case True
3040            .opgLocContact_lbl_dim_hi.Visible = True
3050            .opgLocContact_lbl_line_dim_hi.Visible = True
3060            .Contact_Country_lbl_dim_hi.Visible = True
3070            .Contact_Country_lbl_line_dim_hi.Visible = True
3080            .Contact_PostalCode_lbl_dim_hi.Visible = True
3090            .Contact_PostalCode_lbl_line_dim_hi.Visible = True
3100          Case False
3110            .opgLocContact_lbl_dim_hi.Visible = False
3120            .opgLocContact_lbl_line_dim_hi.Visible = False
3130            .Contact_Country_lbl_dim_hi.Visible = False
3140            .Contact_Country_lbl_line_dim_hi.Visible = False
3150            .Contact_PostalCode_lbl_dim_hi.Visible = False
3160            .Contact_PostalCode_lbl_line_dim_hi.Visible = False
3170          End Select
3180          .opgLocContact_lbl.ForeColor = WIN_CLR_DISF
3190          .opgLocContact_lbl_line.BorderColor = WIN_CLR_DISR
3200          .Contact_State_lbl.ForeColor = WIN_CLR_DISF
3210          .Contact_State_lbl_dim_hi.Visible = True
3220          .Contact_State_lbl_line.BorderColor = WIN_CLR_DISR
3230          .Contact_State_lbl_line_dim_hi.Visible = True
3240          .Contact_Zip_lbl.ForeColor = WIN_CLR_DISF
3250          .Contact_Zip_lbl_dim_hi.Visible = True
3260          .Contact_Zip_lbl_line.BorderColor = WIN_CLR_DISR
3270          .Contact_Zip_lbl_line_dim_hi.Visible = True
3280          .Contact_Country_lbl.ForeColor = WIN_CLR_DISF
3290          .Contact_Country_lbl_line.BorderColor = WIN_CLR_DISR
3300          .Contact_PostalCode_lbl.ForeColor = WIN_CLR_DISF
3310          .Contact_PostalCode_lbl_line.BorderColor = WIN_CLR_DISR
3320          .Contact_Phone1_lbl.ForeColor = WIN_CLR_DISF
3330          .Contact_Phone1_lbl_dim_hi.Visible = True
3340          .Contact_Phone1_lbl_line.BorderColor = WIN_CLR_DISR
3350          .Contact_Phone1_lbl_line_dim_hi.Visible = True
3360          .Contact_Phone2_lbl.ForeColor = WIN_CLR_DISF
3370          .Contact_Phone2_lbl_dim_hi.Visible = True
3380          .Contact_Phone2_lbl_line.BorderColor = WIN_CLR_DISR
3390          .Contact_Phone2_lbl_line_dim_hi.Visible = True
3400          .Contact_Fax_lbl.ForeColor = WIN_CLR_DISF
3410          .Contact_Fax_lbl_dim_hi.Visible = True
3420          .Contact_Fax_lbl_line.BorderColor = WIN_CLR_DISR
3430          .Contact_Fax_lbl_line_dim_hi.Visible = True
3440          .Contact_Email_lbl.ForeColor = WIN_CLR_DISF
3450          .Contact_Email_lbl_dim_hi.Visible = True
3460          .Contact_Email_lbl_line.BorderColor = WIN_CLR_DISR
3470          .Contact_Email_lbl_line_dim_hi.Visible = True
3480        End With
3490        With frmSub2
3500          .accountno1_lbl.ForeColor = WIN_CLR_DISF
3510          .accountno1_lbl_dim_hi.Visible = True
3520          .accountno1_lbl_line.BorderColor = WIN_CLR_DISR
3530          .accountno1_lbl_line_dim_hi.Visible = True
3540          .Contact_Number_lbl.ForeColor = WIN_CLR_DISF
3550          .Contact_Number_lbl_dim_hi.Visible = True
3560          .Contact_Number_lbl_line.BorderColor = WIN_CLR_DISR
3570          .Contact_Number_lbl_line_dim_hi.Visible = True
3580          .Contact_Name_lbl.ForeColor = WIN_CLR_DISF
3590          .Contact_Name_lbl_dim_hi.Visible = True
3600          .Contact_Name_lbl_line.BorderColor = WIN_CLR_DISR
3610          .Contact_Name_lbl_line_dim_hi.Visible = True
3620          .Contact_Address1_lbl.ForeColor = WIN_CLR_DISF
3630          .Contact_Address1_lbl_dim_hi.Visible = True
3640          .Contact_Address1_lbl_line.BorderColor = WIN_CLR_DISR
3650          .Contact_Address1_lbl_line_dim_hi.Visible = True
3660          Select Case blnCountry
              Case True
3670            .opgLocContact_lbl_dim_hi.Visible = True
3680            .opgLocContact_lbl2_dim_hi.Visible = True
3690            .opgLocContact_lbl_line_dim_hi.Visible = True
3700            .Contact_Country_lbl_dim_hi.Visible = True
3710            .Contact_Country_lbl_line_dim_hi.Visible = True
3720          Case False
3730            .opgLocContact_lbl_dim_hi.Visible = False
3740            .opgLocContact_lbl2_dim_hi.Visible = False
3750            .opgLocContact_lbl_line_dim_hi.Visible = False
3760            .Contact_Country_lbl_dim_hi.Visible = False
3770            .Contact_Country_lbl_line_dim_hi.Visible = False
3780          End Select
3790          .opgLocContact_lbl.ForeColor = WIN_CLR_DISF
3800          .opgLocContact_lbl2.ForeColor = WIN_CLR_DISF
3810          .opgLocContact_lbl_line.BorderColor = WIN_CLR_DISR
3820          .Contact_Country_lbl.ForeColor = WIN_CLR_DISF
3830          .Contact_Country_lbl_line.BorderColor = WIN_CLR_DISR
3840          .Contact_City_lbl.ForeColor = WIN_CLR_DISF
3850          .Contact_City_lbl_dim_hi.Visible = True
3860          .Contact_City_lbl_line.BorderColor = WIN_CLR_DISR
3870          .Contact_City_lbl_line_dim_hi.Visible = True
3880          .Contact_Phone1_lbl.ForeColor = WIN_CLR_DISF
3890          .Contact_Phone1_lbl_dim_hi.Visible = True
3900          .Contact_Phone1_lbl_line.BorderColor = WIN_CLR_DISR
3910          .Contact_Phone1_lbl_line_dim_hi.Visible = True
3920          .Contact_Email_lbl.ForeColor = WIN_CLR_DISF
3930          .Contact_Email_lbl_dim_hi.Visible = True
3940          .Contact_Email_lbl_line.BorderColor = WIN_CLR_DISR
3950          .Contact_Email_lbl_line_dim_hi.Visible = True
3960        End With
3970      End If

3980      DoCmd.Hourglass False

3990    End With

EXITP:
4000    Set qdf = Nothing
4010    Set dbs = Nothing
4020    Set frmSub1 = Nothing
4030    Set frmSub2 = Nothing
4040    Exit Sub

ERRH:
100     DoCmd.Hourglass False
110     Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Ignore.
120     Case Else
130       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
140     End Select
150     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim frmSub1 As Access.Form, frmSub2 As Access.Form
        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4120    intShiftDown = (Shift And acShiftMask) > 0
4130    intAltDown = (Shift And acAltMask) > 0
4140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Enable Country:   B {chkEnableCountry_Wide, chkEnableCountry_Compact}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Zip Unfrmat:      D {opgFormatZip_optUnformatted}
        ' **   Send Email:       E {Contact_Email_cmd on frmAccountContacts_Sub2}
        ' **   Zip Frmat:        F {opgFormatZip_optFormatted}
        ' **   Page Count:       G {chkPageOf}
        ' **   Name Last/First:  I {opgFormatName_optLastFirst}
        ' **   Phone Unfrmat:    M {opgFormatPhone_optUnformatted}
        ' **   Other:            O {opgLocContact_optOther on frmAccountContacts_Sub, frmAccountContacts_Sub2}
        ' **   Phone Frmat:      R {opgFormatPhone_optFormatted}
        ' **   Save:             S {cmdSave on frmAccountContacts_Sub}
        ' **   Name As Written:  T {opgFormatName_optAsWritten}
        ' **   USA:              U {opgLocContact_optUSA on frmAccountContacts_Sub, frmAccountContacts_Sub2}

        ' ** Shortcut Ctrl-Alt keys to other forms and functionality:
        ' **   CSZ Separate:     A {opgPrintCSZ_optSeparate, opgPrintCSZCP_optSeparate}
        ' **   CSZ Combined:     B {opgPrintCSZ_optCombined, opgPrintCSZCP_optCombined}
        ' **   Addr Combined:    C {opgPrintAddress_optCombined}
        ' **   Addr None:        N {opgPrintAddress_optNone}
        ' **   CSZ None:         O {opgPrintCSZ_optNone, opgPrintCSZCP_optNone}
        ' **   Addr Separate:    S {opgPrintAddress_optSeparate}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Phone Combined:   C {opgPrintPhone_optCombined}
        ' **   Email None:       E {opgPrintEmail_optNone}
        ' **   Phone None:       N {opgPrintPhone_optNone}
        ' **   Fax None:         O {opgPrintFax_optNone}
        ' **   Fax Print:        P {opgPrintFax_optPrint}
        ' **   Phone Separate:   S {opgPrintPhone_optSeparate}
        ' **   Email Print:      T {opgPrintEmail_optPrint}
        ' **   Phone 1 Only:     Y {opgPrintPhone_opt1Only}

        ' ** Plain keys.
4150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4160      Select Case intRetVal
          Case vbKeyEscape
4170        intRetVal = 0
4180        cmdClose_Click  ' ** Procedure: Below.
4190      End Select
4200    End If

        ' ** Alt keys.
4210    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
4220      Select Case intRetVal
          Case vbKeyB
4230        With Me
4240          intRetVal = 0
4250          Select Case .opgView
              Case .opgView_optWide.OptionValue
4260            If .chkEnableCountry_Wide.Enabled = True Then
4270              .chkEnableCountry_Wide.SetFocus
4280              .chkEnableCountry_Wide = (Not .chkEnableCountry_Wide)
4290              chkEnableCountry_Wide_AfterUpdate  ' ** Procedure: Below.
4300            Else
4310              Beep
4320            End If
4330          Case .opgView_optCompact.OptionValue
4340            If .chkEnableCountry_Compact.Enabled = True Then
4350              .chkEnableCountry_Compact.SetFocus
4360              .chkEnableCountry_Compact = (Not .chkEnableCountry_Compact)
4370              chkEnableCountry_Compact_AfterUpdate  ' ** Procedure: Below.
4380            Else
4390              Beep
4400            End If
4410          End Select
4420        End With
4430      Case vbKeyX
4440        intRetVal = 0
4450        cmdClose_Click  ' ** Procedure: Below.
4460      End Select
4470    End If

        ' ** Ctrl keys.
4480    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
4490      Select Case intRetVal
          Case vbKeyD
            ' **   Zip Unfrmat:      D {opgFormatZip_optUnformatted}
4500        With Me
4510          intRetVal = 0
4520          .opgFormatZip = .opgFormatZip_optUnformatted.OptionValue
4530          .opgFormatZip.SetFocus
4540          opgFormatZip_AfterUpdate  ' ** Procedure: Below.
4550        End With
4560      Case vbKeyE
4570        With Me
4580          intRetVal = 0
4590          Select Case .opgView
              Case .opgView_optWide.OptionValue
4600            Beep
4610          Case .opgView_optCompact.OptionValue
4620            Set frmSub2 = .frmAccountContacts_Sub2.Form
4630            If frmSub2.Contact_Email_cmd.Visible = True And frmSub2.Contact_Email_cmd.Enabled = True Then
4640              .frmAccountContacts_Sub2.SetFocus
4650              frmSub2.Contact_Email_cmd.SetFocus
4660              frmSub2.Contact_Email_cmd_Click  ' ** Form Procedure: frmAccountContacts_Sub2.
4670            End If
4680          End Select
4690        End With
4700      Case vbKeyF
            ' **   Zip Frmat:        F {opgFormatZip_optFormatted}
4710        With Me
4720          intRetVal = 0
4730          .opgFormatZip = .opgFormatZip_optFormatted.OptionValue
4740          .opgFormatZip.SetFocus
4750          opgFormatZip_AfterUpdate  ' ** Procedure: Below.
4760        End With
4770      Case vbKeyG
4780        With Me
4790          intRetVal = 0
4800          If .chkPageOf.Enabled = True Then
4810            .chkPageOf.SetFocus
4820            .chkPageOf = (Not .chkPageOf)
4830            chkPageOf_AfterUpdate  ' ** Procedure: Below.
4840          End If
4850        End With
4860      Case vbKeyI
            ' **   Name Last/First:  I {opgFormatName_optLastFirst}
4870        With Me
4880          intRetVal = 0
4890          .opgFormatName = .opgFormatName_optLastFirst.OptionValue
4900          .opgFormatName.SetFocus
4910          opgFormatName_AfterUpdate  ' ** Procedure: Below.
4920        End With
4930      Case vbKeyM
            ' **   Phone Unfrmat:    M {opgFormatPhone_optUnformatted}
4940        With Me
4950          intRetVal = 0
4960          .opgFormatPhone = .opgFormatPhone_optUnformatted.OptionValue
4970          .opgFormatPhone.SetFocus
4980          opgFormatPhone_AfterUpdate  ' ** Procedure: Below.
4990        End With
5000      Case vbKeyO
5010        With Me
5020          intRetVal = 0
5030          Select Case .opgView
              Case .opgView_optWide.OptionValue
5040            Set frmSub1 = .frmAccountContacts_Sub.Form
5050            If frmSub1.opgLocContact.Visible = True And frmSub1.opgLocContact.Enabled = True Then
5060              .frmAccountContacts_Sub.SetFocus
5070              frmSub1.opgLocContact.SetFocus
5080              frmSub1.opgLocContact = frmSub1.opgLocContact_optOther.OptionValue
5090              frmSub1.opgLocContact_AfterUpdate  ' ** Form Procedure: frmAccountContacts_Sub.
5100            Else
5110              Beep
5120            End If
5130          Case .opgView_optCompact.OptionValue
5140            Set frmSub2 = .frmAccountContacts_Sub2.Form
5150            If frmSub2.opgLocContact.Visible = True And frmSub2.opgLocContact.Enabled = True Then
5160              .frmAccountContacts_Sub2.SetFocus
5170              frmSub2.opgLocContact.SetFocus
5180              frmSub2.opgLocContact = frmSub2.opgLocContact_optOther.OptionValue
5190              frmSub2.opgLocContact_AfterUpdate  ' ** Form Procedure: frmAccountContacts_Sub2.
5200            Else
5210              Beep
5220            End If
5230          End Select
5240        End With
5250      Case vbKeyR
            ' **   Phone Frmat:      R {opgFormatPhone_optFormatted}
5260        With Me
5270          intRetVal = 0
5280          .opgFormatPhone = .opgFormatPhone_optFormatted.OptionValue
5290          .opgFormatPhone.SetFocus
5300          opgFormatPhone_AfterUpdate  ' ** Procedure: Below.
5310        End With
5320      Case vbKeyS
5330        With Me
5340          intRetVal = 0
5350          Select Case .opgView
              Case .opgView_optWide.OptionValue
5360            .frmAccountContacts_Sub.SetFocus
5370            .frmAccountContacts_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAccountContacts_Sub.
5380          Case .opgView_optCompact.OptionValue
5390            .frmAccountContacts_Sub2.SetFocus
5400            .frmAccountContacts_Sub2.Form.cmdSave_Click  ' ** Form Procedure: frmAccountContacts_Sub2.
5410          End Select
5420        End With
5430      Case vbKeyT
            ' **   Name As Written:  T {opgFormatName_optAsWritten}
5440        With Me
5450          intRetVal = 0
5460          .opgFormatName = .opgFormatName_optAsWritten.OptionValue
5470          .opgFormatName.SetFocus
5480          opgFormatName_AfterUpdate  ' ** Procedure: Below.
5490        End With
5500      Case vbKeyU
5510        With Me
5520          intRetVal = 0
5530          Select Case .opgView
              Case .opgView_optWide.OptionValue
5540            Set frmSub1 = .frmAccountContacts_Sub.Form
5550            If frmSub1.opgLocContact.Visible = True And frmSub2.opgLocContact.Enabled = True Then
5560              .frmAccountContacts_Sub.SetFocus
5570              frmSub1.opgLocContact.SetFocus
5580              frmSub1.opgLocContact = frmSub1.opgLocContact_optUSA.OptionValue
5590              frmSub1.opgLocContact_AfterUpdate  ' ** Form Procedure: frmAccountContacts_Sub.
5600            Else
5610              Beep
5620            End If
5630          Case .opgView_optCompact.OptionValue
5640            Set frmSub2 = .frmAccountContacts_Sub2.Form
5650            If frmSub2.opgLocContact.Visible = True And frmSub2.opgLocContact.Enabled = True Then
5660              .frmAccountContacts_Sub2.SetFocus
5670              frmSub2.opgLocContact.SetFocus
5680              frmSub2.opgLocContact = frmSub2.opgLocContact_optUSA.OptionValue
5690              frmSub2.opgLocContact_AfterUpdate  ' ** Form Procedure: frmAccountContacts_Sub2.
5700            Else
5710              Beep
5720            End If
5730          End Select
5740        End With
5750      End Select
5760    End If

        ' ** Ctrl-Alt keys.
5770    If intCtrlDown And intAltDown And (Not intShiftDown) Then
5780      Select Case intRetVal
          Case vbKeyA
            ' **   CSZ Separate:     A {opgPrintCSZ_optSeparate, opgPrintCSZCP_optSeparate}
5790        With Me
5800          intRetVal = 0
5810          If .opgPrintCSZ.Visible = True Then
5820            .opgPrintCSZ = .opgPrintCSZ_optSeparate.OptionValue
5830            .opgPrintCSZ.SetFocus
5840            opgPrintCSZ_AfterUpdate  ' ** Procedure: Below.
5850          ElseIf .opgPrintCSZCP.Visible = True Then
5860            .opgPrintCSZCP = .opgPrintCSZCP_optSeparate.OptionValue
5870            .opgPrintCSZCP.SetFocus
5880            opgPrintCSZCP_AfterUpdate  ' ** Procedure: Below.
5890          End If
5900        End With
5910      Case vbKeyB
            ' **   CSZ Combined:     B {opgPrintCSZ_optCombined, opgPrintCSZCP_optCombined}
5920        With Me
5930          intRetVal = 0
5940          If .opgPrintCSZ.Visible = True Then
5950            .opgPrintCSZ = .opgPrintCSZ_optCombined.OptionValue
5960            .opgPrintCSZ.SetFocus
5970            opgPrintCSZ_AfterUpdate  ' ** Procedure: Below.
5980          ElseIf .opgPrintCSZCP.Visible = True Then
5990            .opgPrintCSZCP = .opgPrintCSZCP_optCombined.OptionValue
6000            .opgPrintCSZCP.SetFocus
6010            opgPrintCSZCP_AfterUpdate  ' ** Procedure: Below.
6020          End If
6030        End With
6040      Case vbKeyC
            ' **   Addr Combined:    C {opgPrintAddress_optCombined}
6050        With Me
6060          intRetVal = 0
6070          .opgPrintAddress = .opgPrintAddress_optCombined.OptionValue
6080          .opgPrintAddress.SetFocus
6090          opgPrintAddress_AfterUpdate  ' ** Procedure: Below.
6100        End With
6110      Case vbKeyN
            ' **   Addr None:        N {opgPrintAddress_optNone}
6120        With Me
6130          intRetVal = 0
6140          .opgPrintAddress = .opgPrintAddress_optNone.OptionValue
6150          .opgPrintAddress.SetFocus
6160          opgPrintAddress_AfterUpdate  ' ** Procedure: Below.
6170        End With
6180      Case vbKeyO
            ' **   CSZ None:         O {opgPrintCSZ_optNone, opgPrintCSZCP_optNone}
6190        With Me
6200          intRetVal = 0
6210          If .opgPrintCSZ.Visible = True Then
6220            .opgPrintCSZ = .opgPrintCSZ_optNone.OptionValue
6230            .opgPrintCSZ.SetFocus
6240            opgPrintCSZ_AfterUpdate  ' ** Procedure: Below.
6250          ElseIf .opgPrintCSZCP.Visible = True Then
6260            .opgPrintCSZCP = .opgPrintCSZCP_optNone.OptionValue
6270            .opgPrintCSZCP.SetFocus
6280            opgPrintCSZCP_AfterUpdate  ' ** Procedure: Below.
6290          End If
6300        End With
6310      Case vbKeyS
            ' **   Addr Separate:    S {opgPrintAddress_optSeparate}
6320        With Me
6330          intRetVal = 0
6340          .opgPrintAddress = .opgPrintAddress_optSeparate.OptionValue
6350          .opgPrintAddress.SetFocus
6360          opgPrintAddress_AfterUpdate  ' ** Procedure: Below.
6370        End With
6380      End Select
6390    End If

        ' ** Ctrl-Shift keys.
6400    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6410      Select Case intRetVal
          Case vbKeyC
            ' **   Phone Combined:   C {opgPrintPhone_optCombined}
6420        With Me
6430          intRetVal = 0
6440          .opgPrintPhone = .opgPrintPhone_optCombined.OptionValue
6450          .opgPrintPhone.SetFocus
6460          opgPrintPhone_AfterUpdate  ' ** Procedure: Below.
6470        End With
6480      Case vbKeyE
            ' **   Email None:       E {opgPrintEmail_optNone}
6490        With Me
6500          intRetVal = 0
6510          .opgPrintEmail = .opgPrintEmail_optNone.OptionValue
6520          .opgPrintEmail.SetFocus
6530          opgPrintEmail_AfterUpdate  ' ** Procedure: Below.
6540        End With
6550      Case vbKeyF
6560        With Me
6570          intRetVal = 0
6580          .FocusHolder.SetFocus
6590          DoCmd.GoToPage 1, 0
6600        End With
6610      Case vbKeyN
            ' **   Phone None:       N {opgPrintPhone_optNone}
6620        With Me
6630          intRetVal = 0
6640          .opgPrintPhone = .opgPrintPhone_optNone.OptionValue
6650          .opgPrintPhone.SetFocus
6660          opgPrintPhone_AfterUpdate  ' ** Procedure: Below.
6670        End With
6680      Case vbKeyO
            ' **   Fax None:         O {opgPrintFax_optNone}
6690        With Me
6700          intRetVal = 0
6710          .opgPrintFax = .opgPrintFax_optNone.OptionValue
6720          .opgPrintFax.SetFocus
6730          opgPrintFax_AfterUpdate  ' ** Procedure: Below.
6740        End With
6750      Case vbKeyP
            ' **   Fax Print:        P {opgPrintFax_optPrint}
6760        With Me
6770          intRetVal = 0
6780          .opgPrintFax = .opgPrintFax_optPrint.OptionValue
6790          .opgPrintFax.SetFocus
6800          opgPrintFax_AfterUpdate  ' ** Procedure: Below.
6810        End With
6820      Case vbKeyS
            ' **   Phone Separate:   S {opgPrintPhone_optSeparate}
6830        With Me
6840          intRetVal = 0
6850          .opgPrintPhone = .opgPrintPhone_optSeparate.OptionValue
6860          .opgPrintPhone.SetFocus
6870          opgPrintPhone_AfterUpdate  ' ** Procedure: Below.
6880        End With
6890      Case vbKeyT
            ' **   Email Print:      T {opgPrintEmail_optPrint}
6900        With Me
6910          intRetVal = 0
6920          .opgPrintEmail = .opgPrintEmail_optPrint.OptionValue
6930          .opgPrintEmail.SetFocus
6940          opgPrintEmail_AfterUpdate  ' ** Procedure: Below.
6950        End With
6960      Case vbKeyY
            ' **   Phone 1 Only:     Y {opgPrintPhone_opt1Only}
6970        With Me
6980          intRetVal = 0
6990          .opgPrintPhone = .opgPrintPhone_opt1Only.OptionValue
7000          .opgPrintPhone.SetFocus
7010          opgPrintPhone_AfterUpdate  ' ** Procedure: Below.
7020        End With
7030      End Select
7040    End If

EXITP:
7050    Set frmSub2 = Nothing
7060    KeyCode = intRetVal
7070    Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub Form_Timer()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

7110    With Me
7120      .TimerInterval = 0&
7130      If gblnSetFocus = True Then
7140        gblnSetFocus = False
7150        strFormName = vbNullString
7160  On Error Resume Next
7170        Do Until strFormName = THIS_NAME
7180          DoCmd.SelectObject acForm, THIS_NAME, False
7190          strFormName = Screen.ActiveForm.Name
7200        Loop
7210  On Error GoTo ERRH
7220      End If
7230      If gblnGoToReport = True Then
7240        gblnGoToReport = False
7250        If .cmdPrint.Enabled = True Then
7260          .cmdPrint.SetFocus
7270        Else
7280          .cmdClose.SetFocus
7290        End If
7300        Beep
7310        DoCmd.Hourglass False
7320        DoEvents
7330      End If
7340    End With

EXITP:
7350    Exit Sub

ERRH:
800     DoCmd.Hourglass False
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

7410    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

7420    Set dbs = CurrentDb
          ' ** Empty tmpAccount.
7430      Set qdf = dbs.QueryDefs("qryAccount_Contacts_18_01")
7440      qdf.Execute
7450    Set qdf = Nothing
7460    dbs.Close
7470    Set dbs = Nothing
7480    DoEvents

7490    gstrReturningForm = THIS_NAME
7500    gblnSetFocus = True
7510    Select Case strCallingForm
        Case "frmMenu_Maintenance"
7520      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Utility"
7530    Case Else
          ' ** Nothing at the moment.
7540    End Select

EXITP:
7550    Set qdf = Nothing
7560    Set dbs = Nothing
7570    Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

7610    With Me
7620      If .cmdReset_raised_focus_dots_img.Visible = True Or .cmdReset_raised_focus_img.Visible = True Then
7630        Select Case blnReset_Focus
            Case True
7640          .cmdReset_raised_semifocus_dots_img.Visible = True
7650          .cmdReset_raised_img.Visible = False
7660        Case False
7670          .cmdReset_raised_img.Visible = True
7680          .cmdReset_raised_semifocus_dots_img.Visible = False
7690        End Select
7700        .cmdReset_raised_focus_img.Visible = False
7710        .cmdReset_raised_focus_dots_img.Visible = False
7720        .cmdReset_sunken_focus_dots_img.Visible = False
7730        .cmdReset_raised_img_dis.Visible = False
7740      End If
7750    End With

EXITP:
7760    Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

7810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7820    Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Public Sub cmdClose_Click()

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim lngRecsChanged As Long
        Dim blnContinue As Boolean

7910    blnContinue = True

7920    With Me
7930      lngRecsCur = .frmAccountContacts_Sub.Form.RecCnt
7940      If lngRecsCur > 0& Then
7950        lngRecsChanged = Abs(.frmAccountContacts_Sub.Form.Contact_Changed_sum)
7960        .Contact_Changed_sum = lngRecsChanged
7970        Select Case .chkAlwaysUpdate
            Case True
7980          blnFromClose = True
7990          cmdUpdate_Click  ' ** Procedure: Below.
8000        Case False
8010          If lngRecsChanged > 0& Or .Has_Deletes = True Then
8020            Beep
8030            If MsgBox("You have made changes, but have not confirmed them by clicking the Update button." & vbCrLf & vbCrLf & _
                    "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved") <> vbYes Then
8040              blnContinue = False
8050            End If
8060          End If
8070        End Select
8080      End If
8090    End With

8100    If blnContinue = True Then
8110      DoCmd.Close acForm, THIS_NAME
8120    End If

EXITP:
8130    Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim frm As Access.Form
        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8220    intShiftDown = (Shift And acShiftMask) > 0
8230    intAltDown = (Shift And acAltMask) > 0
8240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8260      Select Case intRetVal
          Case vbKeyTab
8270        With Me
8280          intRetVal = 0
8290          lngRecsCur = .frmAccountContacts_Sub.Form.RecCnt  ' ** Form Function: frmAccountContacts_Sub.
8300          If lngRecsCur > 0& Then
8310            Select Case .opgView
                Case .opgView_optWide.OptionValue
8320              Set frm = .frmAccountContacts_Sub.Form
8330              .frmAccountContacts_Sub.SetFocus
8340              frm.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAccountContacts_Sub.
8350              If frm.Contact_Name.Visible = True And frm.Contact_Name.Enabled = True Then
8360                frm.Contact_Name.SetFocus
8370              ElseIf frm.Contact_Name_LastFirst.Visible = True And frm.Contact_Name_LastFirst.Enabled = True Then
8380                frm.Contact_Name_LastFirst.SetFocus
8390              End If
8400            Case .opgView_optCompact.OptionValue
8410              Set frm = .frmAccountContacts_Sub2.Form
8420              frm.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAccountContacts_Sub.
8430              If frm.Contact_Name.Visible = True And frm.Contact_Name.Enabled = True Then
8440                frm.Contact_Name.SetFocus
8450              ElseIf frm.Contact_Name_LastFirst.Visible = True And frm.Contact_Name_LastFirst.Enabled = True Then
8460                frm.Contact_Name_LastFirst.SetFocus
8470              End If
8480            End Select
8490          Else
8500            .opgView.SetFocus
8510          End If
8520        End With
8530      End Select
8540    End If

        ' ** Shift keys.
8550    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8560      Select Case intRetVal
          Case vbKeyTab
8570        With Me
8580          intRetVal = 0
8590          .chkAlwaysUpdate.SetFocus
8600        End With
8610      End Select
8620    End If

EXITP:
8630    Set frm = Nothing
8640    KeyCode = intRetVal
8650    Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub cmdPreview_Click()

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim frm As Access.Form
        Dim strDocName As String, strSortNow As String

8710    With Me
8720      Select Case .opgView
          Case .opgView_optWide.OptionValue
8730        Set frm = .frmAccountContacts_Sub.Form
8740      Case .opgView_optCompact.OptionValue
8750        Set frm = .frmAccountContacts_Sub2.Form
8760      End Select
8770      lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountContacts_Sub, frmAccountContacts_Sub2.
8780      If lngRecsCur > 0& Then
8790        gstrReportCallingForm = THIS_NAME  ' ** When not vbNullString, sets this form .Visible = False.
8800        strSortNow = frm.SortNow_Get  ' ** Form Function: frmAccountContacts_Sub, frmAccountContacts_Sub2.
8810        strDocName = "rptAccountContacts"
8820        DoCmd.OpenReport strDocName, acViewPreview, , , , THIS_NAME & "~" & strSortNow
8830        DoCmd.Maximize
8840        DoCmd.RunCommand acCmdFitToWindow
8850      Else
8860        Beep
8870      End If
8880    End With

EXITP:
8890    Set frm = Nothing
8900    Exit Sub

ERRH:
800     DoCmd.Hourglass False
810     DoCmd.Restore
820     If Me.Visible = False Then
830       Me.Visible = True
840     End If
850     Select Case ERR.Number
        Case Else
860       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
870     End Select
880     Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9020    intShiftDown = (Shift And acShiftMask) > 0
9030    intAltDown = (Shift And acAltMask) > 0
9040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9060      Select Case intRetVal
          Case vbKeyTab
9070        With Me
9080          intRetVal = 0
9090          .cmdPrint.SetFocus
9100        End With
9110      End Select
9120    End If

        ' ** Shift keys.
9130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9140      Select Case intRetVal
          Case vbKeyTab
9150        With Me
9160          intRetVal = 0
9170          .chkPageOf.SetFocus
9180        End With
9190      End Select
9200    End If

EXITP:
9210    KeyCode = intRetVal
9220    Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub cmdPrint_Click()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim frm As Access.Form
        Dim strDocName As String, strSortNow As String

9310    With Me
9320      Select Case .opgView
          Case .opgView_optWide.OptionValue
9330        Set frm = .frmAccountContacts_Sub.Form
9340      Case .opgView_optCompact.OptionValue
9350        Set frm = .frmAccountContacts_Sub2.Form
9360      End Select
9370      lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountContacts_Sub, frmAccountContacts_Sub2.
9380      If lngRecsCur > 0& Then
9390        gstrReportCallingForm = vbNullString  ' ** When vbNullString, leaves this form .Visible = True.
9400        strSortNow = frm.SortNow_Get  ' ** Form Function: frmAccountContacts_Sub.
9410        strDocName = "rptAccountContacts"
            '##GTR_Ref: rptAccountContacts
9420        DoCmd.OpenReport strDocName, acViewPreview, , , , THIS_NAME & "~" & strSortNow
9430      Else
9440        Beep
9450      End If
9460    End With

EXITP:
9470    Set frm = Nothing
9480    Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9520    intShiftDown = (Shift And acShiftMask) > 0
9530    intAltDown = (Shift And acAltMask) > 0
9540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9560      Select Case intRetVal
          Case vbKeyTab
9570        With Me
9580          intRetVal = 0
9590          .cmdWord.SetFocus
9600        End With
9610      End Select
9620    End If

        ' ** Shift keys.
9630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9640      Select Case intRetVal
          Case vbKeyTab
9650        With Me
9660          intRetVal = 0
9670          .cmdPreview.SetFocus
9680        End With
9690      End Select
9700    End If

EXITP:
9710    KeyCode = intRetVal
9720    Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub cmdWord_Click()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

9810    With Me

9820      If IsNull(.UserReportPath) = True Then
9830        strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
9840      Else
9850        strRptPath = .UserReportPath
9860      End If
9870      strRptCap = "rptAccountContacts_" & Format(Date, "yyyymmdd")

9880      strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

9890      If strRptPathFile <> vbNullString Then
9900        strRpt = "rptAccountContacts"
9910        DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
9920        .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
9930      End If

9940    End With

EXITP:
9950    Exit Sub

ERRH:
800     Select Case ERR.Number
        Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
810       Beep
820       MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
830     Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
850     End Select
860     Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10020   intShiftDown = (Shift And acShiftMask) > 0
10030   intAltDown = (Shift And acAltMask) > 0
10040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10060     Select Case intRetVal
          Case vbKeyTab
10070       With Me
10080         intRetVal = 0
10090         If .cmdExcel.Enabled Then
10100           .cmdExcel.SetFocus
10110         ElseIf .cmdAdd.Enabled = True Then
10120           .cmdAdd.SetFocus
10130         ElseIf .cmdDelete.Enabled = True Then
10140           .cmdDelete.SetFocus
10150         ElseIf .cmdUpdate.Enabled = True Then
10160           .cmdUpdate.SetFocus
10170         Else
10180           .cmdClose.SetFocus
10190         End If
10200       End With
10210     End Select
10220   End If

        ' ** Shift keys.
10230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10240     Select Case intRetVal
          Case vbKeyTab
10250       With Me
10260         intRetVal = 0
10270         .cmdPrint.SetFocus
10280       End With
10290     End Select
10300   End If

EXITP:
10310   KeyCode = intRetVal
10320   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub cmdExcel_Click()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form
        Dim strRptPath As String, strRptPathFile As String, strRptCap As String
        Dim strSortNow As String, strMacro As String
        Dim strSQL1 As String, strSQL2 As String, strSQL3 As String, strSQL4 As String
        Dim lngXclFlds As Long, arr_varXclFld As Variant
        Dim lngRptFlds As Long, arr_varRptFld As Variant
        Dim lngExpFlds As Long, arr_varExpFld As Variant
        Dim strQrySource1 As String, strQrySource2 As String, strQryName1 As String, strQryName2 As String
        Dim lngConRptID As Long, lngConVerID As Long, lngConFormatID As Long
        Dim blnContinue As Boolean, blnFound As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim lngX As Long, lngY As Long

        ' ** Array: arr_varXclFld().
        Const XF_FNAM As Integer = 0

        ' ** Array: arr_varRptFld().
        Const RF_FNAM As Integer = 0

        ' ** Array: arr_varExpFld().
        'Const EF_RID  As Integer = 0
        Const EF_VER  As Integer = 1
        'Const EF_FID  As Integer = 2
        'Const EF_RTYP As Integer = 3
        'Const EF_KEY  As Integer = 4
        'Const EF_FLDS As Integer = 5
        Const EF_FNAM As Integer = 6
        Const EF_FRMT As Integer = 7

10410   blnContinue = True

10420   With Me

10430     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
10440       DoCmd.Hourglass False
10450       msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
10460       If msgResponse <> vbRetry Then
10470         blnContinue = False
10480       End If
10490     End If

10500     If blnContinue = True Then

10510       DoCmd.Hourglass True
10520       DoEvents

10530       lngConRptID = SetupExcelFlds(THIS_NAME)  ' ** Module Function: modAccountContacts.

10540       If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
10550         .conrpt_id.Visible = True
10560         .conrpt_id = lngConRptID
10570         DoEvents
10580       End If

10590       If lngConRptID > 0& Then

10600         Select Case .opgView
              Case .opgView_optWide.OptionValue
10610           Set frm = .frmAccountContacts_Sub.Form
10620         Case .opgView_optCompact.OptionValue
10630           Set frm = .frmAccountContacts_Sub2.Form
10640         End Select
10650         strSortNow = frm.SortNow_Get  ' ** Form Function: frmAccountContacts_Sub, frmAccountContacts_Sub2.

10660         If IsNull(.UserReportPath) = True Then
10670           strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
10680         Else
10690           strRptPath = .UserReportPath
10700         End If

10710         Set dbs = CurrentDb
10720         With dbs

                ' ** qryAccount_Contacts_22a (tblAccount_Contact_Staging, linked
                ' ** to Account, all fields,), all fields, with new field names.
10730           strQryName1 = "qryAccount_Contacts_23a_01"
10740           arr_varXclFld = Qry_FldList_rel(strQryName1)  ' ** Module Function: modQueryFunctions1.
10750           lngXclFlds = UBound(arr_varXclFld, 2) + 1&
                ' *********************************************
                ' ** Array: arr_varXclFld()
                ' **
                ' **   Field  Element  Name        Constant
                ' **   =====  =======  ==========  ==========
                ' **     1       0     fld_name    XF_FNAM
                ' **
                ' *********************************************
10760           If lngXclFlds > 1& Then

                  ' ** qryAccount_Contacts_22a (tblAccount_Contact_Staging,
                  ' ** linked to Account, all fields), without Contact_ID.
10770             strQryName1 = "qryAccount_Contacts_23a_02"
10780             arr_varRptFld = Qry_FldList_rel(strQryName1)  ' ** Module Function: modQueryFunctions1.
10790             lngRptFlds = UBound(arr_varRptFld, 2) + 1&
                  ' *********************************************
                  ' ** Array: arr_varRptFld()
                  ' **
                  ' **   Field  Element  Name        Constant
                  ' **   =====  =======  ==========  ==========
                  ' **     1       0     fld_name    RF_FNAM
                  ' **
                  ' *********************************************
10800             If lngRptFlds > 1& Then

                    ' ** arr_varXclFld() has the fields names as sent to Excel, currently 23.          THESE TWO
                    ' ** arr_varRptFld() has the fields names as sent to the report, currently 23.     SHOULD MATCH!
                    ' ** arr_varExpFld() has just the fields chosen for export.

                    ' ** tblAcctCon_Report, linked to tblAcctCon_Field, by specified [rid].
10810               Set qdf = .QueryDefs("qryAccount_Contacts_23a_03")
10820               With qdf.Parameters
10830                 ![rid] = lngConRptID
10840               End With
10850               Set rst = qdf.OpenRecordset
10860               With rst
10870                 If .BOF = True And .EOF = True Then
10880                   blnContinue = False
10890                   DoCmd.Hourglass False
10900                   MsgBox "Trust Accountant is unable to generate the Excel export.", vbInformation + vbOKOnly, "Excel Export Error         1"
10910                 Else
10920                   .MoveLast
10930                   lngExpFlds = .RecordCount
10940                   .MoveFirst
10950                   arr_varExpFld = .GetRows(lngExpFlds)
                        ' **************************************************
                        ' ** Array: arr_varExpFld()
                        ' **
                        ' **   Field  Element  Name             Constant
                        ' **   =====  =======  ===============  ==========
                        ' **     1       0     conrpt_id        EF_RID
                        ' **     2       1     conver_id        EF_VER
                        ' **     3       2     confld_id        EF_FID
                        ' **     4       3     conrpt_type      EF_RTYP
                        ' **     5       4     conrptver_key    EF_KEY
                        ' **     6       5     conrpt_fields    EF_FLDS
                        ' **     7       6     confld_name      EF_FNAM
                        ' **     8       7     confmt_id04      EF_FRMT
                        ' **
                        ' **************************************************
10960                 End If
10970                 .Close
10980               End With
10990               Set rst = Nothing
11000               Set qdf = Nothing

11010               If blnContinue = True Then

11020                 lngConVerID = arr_varExpFld(EF_VER, 0)  ' ** accountno only, shortname only, or both.
11030                 lngConFormatID = arr_varExpFld(EF_FRMT, 0)  ' ** Format of Contact_Name.
                      ' ** confmt_id04 has: 'Name' - Unformatted, 'Contact Last, First' - Formatted.

11040                 strRptCap = "rptAccountContacts_" & Format(Date, "yyyymmdd")

11050                 If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
11060                   Select Case lngConVerID
                        Case 1&  ' ** AccountNumOnly.
11070                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
11080                       strRptCap = strRptCap & "_af"
11090                     Case 2&  ' ** Unformatted.
11100                       strRptCap = strRptCap & "_au"
11110                     End Select
11120                   Case 2&  ' ** ShortNameOnly.
11130                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
11140                       strRptCap = strRptCap & "_sf"
11150                     Case 2&  ' ** Unformatted.
11160                       strRptCap = strRptCap & "_su"
11170                     End Select
11180                   Case 3&  ' ** AccountNumShortName.
11190                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
11200                       strRptCap = strRptCap & "_asf"
11210                     Case 2&  ' ** Unformatted.
11220                       strRptCap = strRptCap & "_asu"
11230                     End Select
11240                   End Select
11250                 End If

11260                 DoCmd.Hourglass False

11270                 strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

11280                 If strRptPathFile <> vbNullString Then

11290                   DoCmd.Hourglass True
11300                   DoEvents

11310                   Select Case lngConVerID
                        Case 1&  ' ** AccountNumOnly.
11320                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
11330                       strQryName1 = "qryAccount_Contacts_24_02"
11340                       strQryName2 = "qryAccount_Contacts_27_02"
11350                     Case 2&  ' ** Unformatted.
11360                       strQryName1 = "qryAccount_Contacts_24_01"
11370                       strQryName2 = "qryAccount_Contacts_27_01"
11380                     End Select
11390                   Case 2&  ' ** ShortNameOnly.
11400                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
11410                       strQryName1 = "qryAccount_Contacts_24_04"
11420                       strQryName2 = "qryAccount_Contacts_27_04"
11430                     Case 2&  ' ** Unformatted.
11440                       strQryName1 = "qryAccount_Contacts_24_03"
11450                       strQryName2 = "qryAccount_Contacts_27_03"
11460                     End Select
11470                   Case 3&  ' ** AccountNumShortName.
11480                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
11490                       strQryName1 = "qryAccount_Contacts_24_06"
11500                       strQryName2 = "qryAccount_Contacts_27_06"
11510                     Case 2&  ' ** Unformatted.
11520                       strQryName1 = "qryAccount_Contacts_24_05"
11530                       strQryName2 = "qryAccount_Contacts_27_05"
11540                     End Select
11550                   End Select

                        ' ** qryAccount_Contacts_23a_01 (qryAccount_Contacts_22a (tblAccount_Contact_Staging, linked
                        ' ** to Account, all fields), all fields, with new field names), just needed fields.
11560                   Set qdf = .QueryDefs(strQryName1)
11570                   strSQL1 = qdf.SQL  ' ** This is the previous report's SQL.
11580                   Set qdf = Nothing

                        ' ** SELECT qryAccount_Contacts_23a_01.[Account Num], qryAccount_Contacts_23a_01.[Account Name],
                        ' **   qryAccount_Contacts_23a_01.[Contact #], qryAccount_Contacts_23a_01.[Contact Last, First],
                        ' **   qryAccount_Contacts_23a_01.[Address 1], qryAccount_Contacts_23a_01.[Address 2],
                        ' **   qryAccount_Contacts_23a_01.City, qryAccount_Contacts_23a_01.State,
                        ' **   qryAccount_Contacts_23a_01.[Zip +], qryAccount_Contacts_23a_01.[Phone 1 +],
                        ' **   qryAccount_Contacts_23a_01.[Phone 2 +], qryAccount_Contacts_23a_01.[Fax +],
                        ' **   qryAccount_Contacts_23a_01.Email
                        ' ** FROM qryAccount_Contacts_23a_01;

11590                   intPos01 = InStr(strSQL1, ".")
11600                   strSQL2 = Left(strSQL1, intPos01)
11610                   intPos01 = InStr(strSQL2, " ")
11620                   strQrySource1 = Trim(Mid(strSQL2, intPos01))  ' ** Query source name, with '.'.
11630                   strSQL2 = Left(strSQL2, intPos01)  ' ** Just 'SELECT '.
11640                   strSQL3 = vbNullString

11650                   For lngX = 0& To (lngExpFlds - 1&)
11660                     For lngY = 0& To (lngRptFlds - 1&)
11670                       If arr_varRptFld(RF_FNAM, lngY) = arr_varExpFld(EF_FNAM, lngX) Then
                              ' ** arr_varXclFld() should exactly match arr_varRptFld().
                              ' ** Use brackets because new field names have spaces.
11680                         strSQL2 = strSQL2 & strQrySource1 & "[" & arr_varXclFld(XF_FNAM, lngY) & "], "
11690                         If arr_varExpFld(EF_FNAM, lngX) <> "accountno" And arr_varExpFld(EF_FNAM, lngX) <> "shortname" And _
                                  arr_varExpFld(EF_FNAM, lngX) <> "Contact_Number" And arr_varExpFld(EF_FNAM, lngX) <> "Contact_Name" And _
                                  arr_varExpFld(EF_FNAM, lngX) <> "Contact_Name_LastFirst" Then
11700                           strSQL3 = strSQL3 & "'' As [" & arr_varXclFld(XF_FNAM, lngY) & "], "
11710                         End If
11720                         Exit For
11730                       End If
11740                     Next  ' ** lngY
11750                   Next  ' ** lngX.
11760                   strSQL4 = strSQL2  ' ** The output query shouldn't see the sort column.
11770                   strSQL2 = strSQL2 & "CLng(1) AS sort "
11780                   If Right(Trim(strSQL2), 1) = "," Then strSQL2 = Left(Trim(strSQL2), (Len(Trim(strSQL2)) - 1))  ' ** Strip the last comma.
11790                   If Right(Trim(strSQL4), 1) = "," Then strSQL4 = Left(Trim(strSQL4), (Len(Trim(strSQL4)) - 1))  ' ** Strip the last comma.
11800                   intPos01 = InStr(strSQL1, "FROM ")
11810                   strSQL2 = strSQL2 & " " & Mid(strSQL1, intPos01)  ' ** Add the FROM clause.
11820                   strSQL4 = strSQL4 & " " & Mid(strSQL1, intPos01)  ' ** Add the FROM clause.
11830                   intPos01 = InStr(strSQL2, "ORDER BY")
11840                   If intPos01 > 0 Then strSQL2 = Left(strSQL2, (intPos01 - 1))
11850                   intPos01 = InStr(strSQL4, "ORDER BY")
11860                   If intPos01 > 0 Then strSQL4 = Left(strSQL4, (intPos01 - 1))

11870                   Set qdf = .QueryDefs(strQryName1)
11880                   qdf.SQL = strSQL2  ' ** Put the new field list into the original query.
11890                   Set qdf = Nothing

                        ' ** The query source is different for the output query.
11900                   strTmp01 = vbNullString
11910                   Select Case lngConVerID
                        Case 1&  ' ** AccountNumOnly.
11920                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
11930                       strTmp01 = "_26_02"
11940                     Case 2&  ' ** Unformatted.
11950                       strTmp01 = "_26_01"
11960                     End Select
11970                   Case 2&  ' ** ShortNameOnly.
11980                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
11990                       strTmp01 = "_26_04"
12000                     Case 2&  ' ** Unformatted.
12010                       strTmp01 = "_26_03"
12020                     End Select
12030                   Case 3&  ' ** AccountNumShortName.
12040                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
12050                       strTmp01 = "_26_06"
12060                     Case 2&  ' ** Unformatted.
12070                       strTmp01 = "_26_05"
12080                     End Select
12090                   End Select
12100                   strSQL4 = StringReplace(strSQL4, "_23a_01", strTmp01)  ' ** Module Function: modStringFuncs.
12110                   strQrySource2 = "qryAccount_Contacts" & strTmp01 & "."

12120                   intPos01 = InStr(strSQL4, ";")
12130                   If intPos01 > 0 Then strSQL4 = Left(strSQL4, (intPos01 - 1))
12140                   strSQL4 = strSQL4 & " ORDER BY "
12150                   strSQL1 = vbNullString
12160                   strTmp01 = strSortNow
12170                   strTmp02 = vbNullString
12180                   intPos01 = InStr(strTmp01, ",")
12190                   If intPos01 > 0 Then
12200                     Do While intPos01 > 0
12210                       strTmp02 = Trim(Left(strTmp01, (intPos01 - 1)))
12220                       strTmp01 = Mid(strTmp01, (intPos01 + 1))
12230                       If Left(strTmp02, 1) = "[" Then strTmp02 = Mid(strTmp02, 2)
12240                       If Right(strTmp02, 1) = "]" Then strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
12250                       blnFound = False
12260                       For lngY = 0& To (lngRptFlds - 1&)
12270                         If arr_varRptFld(RF_FNAM, lngY) = strTmp02 Then
12280                           blnFound = True
12290                           strSQL1 = strSQL1 & strQrySource2 & "[" & arr_varXclFld(XF_FNAM, lngY) & "], "
12300                           Exit For
12310                         End If
12320                       Next
12330                       If blnFound = False Then
                              ' ** Contact_Address_Combined, Contact_CSZ, Contact_CSZ_Format, Contact_Phone_Combined,
                              ' ** and Contact_Phone_Format_Combined don't ever appear on the subform.
                              ' ** [Address], [City, State, Zip], [City, State, Zip +], [Phone], [Phone +].
12340                         If InStr(strTmp02, "Address") > 0 Then
12350                           strSQL1 = strSQL1 & strQrySource2 & "[Address], "
12360                         Else
12370                           If InStr(strTmp02, "Phone") > 0 And InStr(strTmp02, "Format") > 0 Then
12380                             strSQL1 = strSQL1 & strQrySource2 & "[Phone +], "
12390                           ElseIf InStr(strTmp02, "Phone") > 0 Then
12400                             strSQL1 = strSQL1 & strQrySource2 & "[Phone], "
12410                           Else
12420                             Stop
12430                           End If
12440                         End If
12450                       End If
12460                       intPos01 = InStr(strTmp01, ",")
12470                       If intPos01 = 0 Then
12480                         strTmp02 = Trim(strTmp01)
12490                         If Left(strTmp02, 1) = "[" Then strTmp02 = Mid(strTmp02, 2)
12500                         If Right(strTmp02, 1) = "]" Then strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
12510                         blnFound = False
12520                         For lngY = 0& To (lngRptFlds - 1&)
12530                           If arr_varRptFld(RF_FNAM, lngY) = strTmp02 Then
12540                             blnFound = True
12550                             strSQL1 = strSQL1 & strQrySource2 & "[" & arr_varXclFld(XF_FNAM, lngY) & "], "
12560                             Exit For
12570                           End If
12580                         Next
12590                         If blnFound = False Then
12600                           Stop
12610                         End If
12620                       End If
12630                     Loop
12640                     strSQL1 = Trim(strSQL1)
12650                     If Right(strSQL1, 1) = "," Then strSQL1 = Left(strSQL1, (Len(strSQL1) - 1))
12660                     strSQL1 = "[sort], " & strSQL1
12670                   Else
12680                     Stop
12690                   End If
12700                   strSQL4 = strSQL4 & strSQL1 & ";"

                        ' ** SELECT 'Period:' AS [Account Name], '' AS [Contact #],
                        ' **   'As of ' & Format(Date(),'mm/dd/yy') AS [Contact Last, First],
                        ' **   '' AS [Address 1], '' AS [Address 2], '' AS City, '' AS State, '' AS [Zip +],
                        ' **   '' AS [Phone 1 +], '' AS [Phone 2 +], CLng(-1) AS sort
                        ' ** FROM tblYesNo
                        ' ** WHERE (((tblYesNo.yn_boolean)=True));

12710                   strSQL1 = vbNullString: strSQL2 = vbNullString
12720                   Select Case lngConVerID
                        Case 1&  ' ** AccountNumOnly.
12730                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
                            ' ** Replace everything after '[Contact Last, First]'.
                            ' ** Report title.
12740                       Set qdf = .QueryDefs("qryAccount_Contacts_25a_02")
12750                     Case 2&  ' ** Unformatted.
                            ' ** Replace everything after 'Name'.
                            ' ** Report title.
12760                       Set qdf = .QueryDefs("qryAccount_Contacts_25a_01")
12770                     End Select
12780                   Case 2&  ' ** ShortNameOnly.
12790                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
                            ' ** Replace everything after '[Contact Last, First]'.
                            ' ** Report title.
12800                       Set qdf = .QueryDefs("qryAccount_Contacts_25a_04")
12810                     Case 2&  ' ** Unformatted.
                            ' ** Replace everything after 'Name'.
                            ' ** Report title.
12820                       Set qdf = .QueryDefs("qryAccount_Contacts_25a_03")
12830                     End Select
12840                   Case 3&  ' ** AccountNumShortName.
12850                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
                            ' ** Replace everything after '[Contact Last, First]'.
                            ' ** Report title.
12860                       Set qdf = .QueryDefs("qryAccount_Contacts_25a_06")
12870                     Case 2&  ' ** Unformatted.
                            ' ** Replace everything after 'Name'.
                            ' ** Report title.
12880                       Set qdf = .QueryDefs("qryAccount_Contacts_25a_05")
12890                     End Select
12900                   End Select

12910                   Select Case lngConFormatID
                        Case 1&  ' ** Formatted.
                          ' ** Report title.
12920                     strSQL1 = qdf.SQL
12930                     intPos01 = InStr(strSQL1, "First]")
12940                     If intPos01 > 0 Then
12950                       intPos01 = InStr(intPos01, strSQL1, ",")
12960                       strSQL2 = Left(strSQL1, intPos01)
12970                       strSQL2 = strSQL2 & strSQL3 & "CLng(-2) AS sort "
12980                       intPos01 = InStr(strSQL1, "FROM ")
12990                       strSQL2 = strSQL2 & Mid(strSQL1, intPos01)
13000                       qdf.SQL = strSQL2
13010                     End If
13020                     Set qdf = Nothing
13030                   Case 2&  ' ** Unformatted.
                          ' ** Report title.
13040                     strSQL1 = qdf.SQL
13050                     intPos01 = InStr(strSQL1, "[Contact #],")
13060                     If intPos01 > 0 Then
13070                       intPos01 = intPos01 + Len("[Contact #],")
13080                       intPos01 = InStr(intPos01, strSQL1, " AS ")
13090                       intPos01 = InStr(intPos01, strSQL1, ",")
13100                       strSQL2 = Left(strSQL1, intPos01)
13110                       strSQL2 = strSQL2 & strSQL3 & "CLng(-2) AS sort "
13120                       intPos01 = InStr(strSQL1, "FROM ")
13130                       strSQL2 = strSQL2 & Mid(strSQL1, intPos01)
13140                       qdf.SQL = strSQL2
13150                     End If
13160                     Set qdf = Nothing
13170                   End Select

13180                   strSQL1 = vbNullString: strSQL2 = vbNullString
13190                   Select Case lngConVerID
                        Case 1&  ' ** AccountNumOnly.
13200                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
                            ' ** Replace everything after '[Contact Last, First]'.
                            ' ** Report period.
13210                       Set qdf = .QueryDefs("qryAccount_Contacts_25b_02")
13220                     Case 2&  ' ** Unformatted.
                            ' ** Replace everything after 'Name'.
                            ' ** Report period.
13230                       Set qdf = .QueryDefs("qryAccount_Contacts_25b_01")
13240                     End Select
13250                   Case 2&  ' ** ShortNameOnly.
13260                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
                            ' ** Replace everything after '[Contact Last, First]'.
                            ' ** Report period.
13270                       Set qdf = .QueryDefs("qryAccount_Contacts_25b_04")
13280                     Case 2&  ' ** Unformatted.
                            ' ** Replace everything after 'Name'.
                            ' ** Report period.
13290                       Set qdf = .QueryDefs("qryAccount_Contacts_25b_03")
13300                     End Select
13310                   Case 3&  ' ** AccountNumShortName.
13320                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
                            ' ** Replace everything after '[Contact Last, First]'.
                            ' ** Report period.
13330                       Set qdf = .QueryDefs("qryAccount_Contacts_25b_06")
13340                     Case 2&  ' ** Unformatted.
                            ' ** Replace everything after 'Name'.
                            ' ** Report period.
13350                       Set qdf = .QueryDefs("qryAccount_Contacts_25b_05")
13360                     End Select
13370                   End Select

13380                   Select Case lngConFormatID
                        Case 1&  ' ** Formatted.
                          ' ** Report period.
13390                     strSQL1 = qdf.SQL
13400                     intPos01 = InStr(strSQL1, "First]")
13410                     If intPos01 > 0 Then
13420                       intPos01 = InStr(intPos01, strSQL1, ",")
13430                       strSQL2 = Left(strSQL1, intPos01)
13440                       strSQL2 = strSQL2 & strSQL3 & "CLng(-1) AS sort "
13450                       intPos01 = InStr(strSQL1, "FROM ")
13460                       strSQL2 = strSQL2 & Mid(strSQL1, intPos01)
13470                       qdf.SQL = strSQL2
13480                     End If
13490                     Set qdf = Nothing
13500                   Case 2&  ' ** Unformatted.
                          ' ** Report period.
13510                     strSQL1 = qdf.SQL
13520                     intPos01 = InStr(strSQL1, "[Contact #],")
13530                     If intPos01 > 0 Then
13540                       intPos01 = intPos01 + Len("[Contact #],")
13550                       intPos01 = InStr(intPos01, strSQL1, " AS ")
13560                       intPos01 = InStr(intPos01, strSQL1, ",")
13570                       strSQL2 = Left(strSQL1, intPos01)
13580                       strSQL2 = strSQL2 & strSQL3 & "CLng(-1) AS sort "
13590                       intPos01 = InStr(strSQL1, "FROM ")
13600                       strSQL2 = strSQL2 & Mid(strSQL1, intPos01)
13610                       qdf.SQL = strSQL2
13620                       Set qdf = Nothing
13630                     End If
13640                     Set qdf = Nothing
13650                   End Select

                        ' ** Finally, update the output query's field list.
13660                   Set qdf = .QueryDefs(strQryName2)
13670                   qdf.SQL = strSQL4
13680                   Set qdf = Nothing

                        ' ** Now set up the Macro.
13690                   strMacro = "mcrExcelExport_AcctCon_"
13700                   Select Case lngConVerID
                        Case 1&  ' ** AccountNumOnly.
13710                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
                            ' ** mcrExcelExport_AcctCon_02_ActNo_FirstLast.
                            ' ** strQryName2: qryAccount_Contacts_27_02.
13720                       strMacro = strMacro & "02_ActNo_FirstLast"
13730                     Case 2&  ' ** Unformatted.
                            ' ** mcrExcelExport_AcctCon_01_ActNo_Name.
                            ' ** strQryName2: qryAccount_Contacts_27_01.
13740                       strMacro = strMacro & "01_ActNo_Name"
13750                     End Select
13760                   Case 2&  ' ** ShortNameOnly.
13770                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
                            ' ** mcrExcelExport_AcctCon_04_Short_FirstLast.
                            ' ** strQryName2: qryAccount_Contacts_27_04.
13780                       strMacro = strMacro & "04_Short_FirstLast"
13790                     Case 2&  ' ** Unformatted.
                            ' ** mcrExcelExport_AcctCon_03_Short_Name.
                            ' ** strQryName2: qryAccount_Contacts_27_03.
13800                       strMacro = strMacro & "03_Short_Name"
13810                     End Select
13820                   Case 3&  ' ** AccountNumShortName.
13830                     Select Case lngConFormatID
                          Case 1&  ' ** Formatted.
                            ' ** mcrExcelExport_AcctCon_06_ActNoShort_FirstLast.
                            ' ** strQryName2: qryAccount_Contacts_27_06.
13840                       strMacro = strMacro & "06_ActNoShort_FirstLast"
13850                     Case 2&  ' ** Unformatted.
                            ' ** mcrExcelExport_AcctCon_05_ActNoShort_Name.
                            ' ** strQryName2: qryAccount_Contacts_27_05.
13860                       strMacro = strMacro & "05_ActNoShort_Name"
13870                     End Select
13880                   End Select

13890                   If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
13900                     Kill strRptPathFile
13910                   End If

                        ' ** Run the Macro.
13920                   DoCmd.RunMacro strMacro
                        'DoCmd.OutputTo acOutputQuery, strQryName2, acFormatXLS, strRptPathFile, True
                        ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                        ' ** So, it's exported to 'AccountContacts_xxx.xls', which is then renamed.
13930                   If FileExists(CurrentAppPath & LNK_SEP & "AccountContacts_xxx.xls") = True Then  ' ** Module Function: modFileUtilities.
13940                     Name (CurrentAppPath & LNK_SEP & "AccountContacts_xxx.xls") As (strRptPathFile)  ' ** Module Function: modFileUtilities.
                          ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
13950                     If Excel_AcctCon(strRptPathFile, strQryName2) = True Then  ' ** Module Function: modExcelFuncs.
13960                       DoEvents
                            ' ** Even though all references to the Excel_IncExp() objects are explicit,
                            ' ** and they're closed and quit, sometimes the process will not shut down.
                            ' ** (One suggestion is that the worksheet copy is the culprit.)
                            ' ** This, below, seems to be my only recourse.
13970                       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
13980                         EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
13990                       End If
14000                       DoEvents
                            'Shell "EXCEL.EXE " & Chr(34) & (strRptPathFile) & Chr(34), vbNormalFocus
14010                       OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
14020                     End If
14030                   ElseIf FileExists(strRptPath & LNK_SEP & "AccountContacts_xxx.xls") = True Then
14040                     Name (strRptPath & LNK_SEP & "AccountContacts_xxx.xls") As (strRptPathFile)
                          ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
14050                     If Excel_AcctCon(strRptPathFile, strQryName2) = True Then  ' ** Module Function: modExcelFuncs.
14060                       DoEvents
                            ' ** Even though all references to the Excel_IncExp() objects are explicit,
                            ' ** and they're closed and quit, sometimes the process will not shut down.
                            ' ** (One suggestion is that the worksheet copy is the culprit.)
                            ' ** This, below, seems to be my only recourse.
14070                       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
14080                         EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
14090                       End If
14100                       DoEvents
14110                       OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
14120                     End If
14130                   End If

14140                 End If  ' ** strRptPathFile.

14150               End If  ' ** blnContinue.

14160             Else
14170               blnContinue = False
14180               DoCmd.Hourglass False
14190               MsgBox "Trust Accountant is unable to generate the Excel export.", vbInformation + vbOKOnly, "Excel Export Error         2"
14200             End If  ' ** lngRptFlds.
14210           Else
14220             blnContinue = False
14230             DoCmd.Hourglass False
14240             MsgBox "Trust Accountant is unable to generate the Excel export.", vbInformation + vbOKOnly, "Excel Export Error         3"
14250           End If  ' ** lngXclFlds.
14260           .Close
14270         End With  ' ** dbs.

14280         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.

14290         DoCmd.Hourglass False

14300       Else
14310         blnContinue = False
14320         DoCmd.Hourglass False
14330         MsgBox "Trust Accountant is unable to generate the Excel export.", vbInformation + vbOKOnly, "Excel Export Error         4"
14340       End If

14350     End If  ' ** blnContinue.
14360   End With

EXITP:
14370   Set frm = Nothing
14380   Set rst = Nothing
14390   Set qdf = Nothing
14400   Set dbs = Nothing
14410   Exit Sub

ERRH:
800     DoCmd.Hourglass False
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14520   intShiftDown = (Shift And acShiftMask) > 0
14530   intAltDown = (Shift And acAltMask) > 0
14540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14560     Select Case intRetVal
          Case vbKeyTab
14570       With Me
14580         intRetVal = 0
14590         If .cmdAdd.Enabled = True Then
14600           .cmdAdd.SetFocus
14610         ElseIf .cmdDelete.Enabled = True Then
14620           .cmdDelete.SetFocus
14630         ElseIf .cmdUpdate.Enabled = True Then
14640           .cmdUpdate.SetFocus
14650         Else
14660           .chkAlwaysUpdate.SetFocus
14670         End If
14680       End With
14690     End Select
14700   End If

        ' ** Shift keys.
14710   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14720     Select Case intRetVal
          Case vbKeyTab
14730       With Me
14740         intRetVal = 0
14750         .cmdWord.SetFocus
14760       End With
14770     End Select
14780   End If

EXITP:
14790   KeyCode = intRetVal
14800   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub cmdAdd_Click()

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form
        Dim strAccountNo As String, strHasOne As String, strHasTwo As String

14910   With Me

14920     DoCmd.Hourglass True
14930     DoEvents

14940     Select Case .opgView
          Case .opgView_optWide.OptionValue
14950       Set frm = .frmAccountContacts_Sub.Form
14960     Case .opgView_optCompact.OptionValue
14970       Set frm = .frmAccountContacts_Sub2.Form
14980     End Select

14990     lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountContacts_Sub, frmAccountContacts_Sub2.
15000     If lngRecsCur > 0& Then
15010       strAccountNo = frm.accountno
15020     Else
15030       strAccountNo = "{none}"
15040     End If

15050     Set dbs = CurrentDb

15060     strHasOne = "False"
15070     strHasTwo = "False"
15080     If strAccountNo <> "{none}" Then
            ' ** qryAccount_Contacts_16 (tblAccount_Contact_Staging, with Has_One, Has_Two,
            ' ** by specified [actno]), grouped by accountno, with Min(Has_One), Min(Has_Two).
15090       Set qdf = dbs.QueryDefs("qryAccount_Contacts_17")
15100       With qdf.Parameters
15110         ![actno] = strAccountNo
15120       End With
15130       Set rst = qdf.OpenRecordset
15140       With rst
15150         If .BOF = True And .EOF = True Then
                ' ** No contacts for this account yet.
15160         Else
15170           .MoveFirst
15180           If ![Has_One] = True Then strHasOne = "True"
15190           If ![Has_Two] = True Then strHasTwo = "True"
15200         End If
15210         .Close
15220       End With
15230       Set rst = Nothing
15240       Set qdf = Nothing
15250       DoEvents
15260     End If
15270     dbs.Close
15280     Set dbs = Nothing
15290     DoEvents

15300     DoCmd.OpenForm "frmAccountContacts_Add", , , , , acDialog, THIS_NAME & "~" & strAccountNo & "~" & strHasOne & "~" & strHasTwo

15310     DoEvents

15320     lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountContacts_Sub, frmAccountContacts_Sub2.
15330     If lngRecsCur > 0& And .opgFormatName.Enabled = False Then
            ' ** First contact.
15340       .opgFormatName.Enabled = True
15350       .opgFormatZip.Enabled = True
15360       .opgFormatPhone.Enabled = True
15370       .opgPrintAddress.Enabled = True
15380       .opgPrintCSZ.Enabled = True
15390       .opgPrintCSZCP.Enabled = True
15400       .opgPrintPhone.Enabled = True
15410       .opgPrintFax.Enabled = True
15420       .opgPrintEmail.Enabled = True
15430       .chkShowAcctNum.Enabled = True
15440       .chkShowShortName.Enabled = True
15450       Select Case .opgView
            Case .opgView_optWide.OptionValue
15460         .chkEnableCountry_Wide.Enabled = True
15470       Case .opgView_optCompact.OptionValue
15480         .chkEnableCountry_Compact.Enabled = True
15490       End Select
15500       .cmdPreview.Enabled = True
15510       .cmdPrint.Enabled = True
15520       .cmdWord.Enabled = True
      #If NoExcel Then
15530       .cmdExcel.Enabled = False
      #Else
15540       .cmdExcel.Enabled = True
      #End If
15550       .cmdDelete.Enabled = True
15560       .cmdReset.Enabled = True
15570       .cmdReset_raised_img.Visible = True
15580       .cmdReset_raised_semifocus_dots_img.Visible = False
15590       .cmdReset_raised_focus_img.Visible = False
15600       .cmdReset_raised_focus_dots_img.Visible = False
15610       .cmdReset_sunken_focus_dots_img.Visible = False
15620       .cmdReset_raised_img_dis.Visible = False
15630       DoEvents
15640     End If

15650   End With

EXITP:
15660   Set frm = Nothing
15670   Set rst = Nothing
15680   Set qdf = Nothing
15690   Set dbs = Nothing
15700   Exit Sub

ERRH:
800     DoCmd.Hourglass False
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15820   intShiftDown = (Shift And acShiftMask) > 0
15830   intAltDown = (Shift And acAltMask) > 0
15840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15860     Select Case intRetVal
          Case vbKeyTab
15870       With Me
15880         intRetVal = 0
15890         If .cmdDelete.Enabled = True Then
15900           .cmdDelete.SetFocus
15910         ElseIf .cmdUpdate.Enabled = True Then
15920           .cmdUpdate.SetFocus
15930         Else
15940           .chkAlwaysUpdate.SetFocus
15950         End If
15960       End With
15970     End Select
15980   End If

        ' ** Shift keys.
15990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16000     Select Case intRetVal
          Case vbKeyTab
16010       With Me
16020         intRetVal = 0
16030         If .cmdExcel.Enabled = True Then
16040           .cmdExcel.SetFocus
16050         ElseIf .cmdWord.Enabled = True Then
16060           .cmdWord.SetFocus
16070         Else
16080           .chkPageOf.SetFocus
16090         End If
16100       End With
16110     End Select
16120   End If

EXITP:
16130   KeyCode = intRetVal
16140   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub cmdDelete_Click()

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_Click"

        Dim frm As Access.Form
        Dim strAccountNo As String, strContactName As String
        Dim lngContactID As Long
        Dim msgResponse As VbMsgBoxResult

16210   With Me
16220     Select Case .opgView
          Case .opgView_optWide.OptionValue
16230       Set frm = .frmAccountContacts_Sub.Form
16240     Case .opgView_optCompact.OptionValue
16250       Set frm = .frmAccountContacts_Sub2.Form
16260     End Select
16270     lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountContacts_Sub, frmAccountContacts_Sub2.
16280     If lngRecsCur > 0& Then
16290       strAccountNo = frm.accountno
16300       strContactName = frm.Contact_Name
16310       Beep
16320       msgResponse = MsgBox("Are you sure you want to delete this contact?" & vbCrLf & _
              "    " & strContactName & " for account " & strAccountNo, vbQuestion + vbYesNo, "Delete Contact")
16330       If msgResponse = vbYes Then
16340         lngContactID = frm.Contact_ID
16350         frm.DelRec_AC lngContactID  ' ** Form Function: frmAccountContacts_Sub, frmAccountContacts_Sub2.
16360         .Has_Deletes = True
16370       End If
16380     Else
16390       Beep
16400     End If
16410   End With

EXITP:
16420   Set frm = Nothing
16430   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16520   intShiftDown = (Shift And acShiftMask) > 0
16530   intAltDown = (Shift And acAltMask) > 0
16540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16560     Select Case intRetVal
          Case vbKeyTab
16570       With Me
16580         intRetVal = 0
16590         If .cmdUpdate.Enabled = True Then
16600           .cmdUpdate.SetFocus
16610         Else
16620           .chkAlwaysUpdate.SetFocus
16630         End If
16640       End With
16650     End Select
16660   End If

        ' ** Shift keys.
16670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16680     Select Case intRetVal
          Case vbKeyTab
16690       With Me
16700         intRetVal = 0
16710         If .cmdAdd.Enabled = True Then
16720           .cmdAdd.SetFocus
16730         ElseIf .cmdExcel.Enabled = True Then
16740           .cmdExcel.SetFocus
16750         ElseIf .cmdWord.Enabled = True Then
16760           .cmdWord.SetFocus
16770         Else
16780           .chkPageOf.SetFocus
16790         End If
16800       End With
16810     End Select
16820   End If

EXITP:
16830   KeyCode = intRetVal
16840   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub cmdUpdate_Click()

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim frmSub1 As Access.Form, frmSub2 As Access.Form
        Dim lngRecs1 As Long, lngRecs2 As Long, lngRecsChanged As Long, lngRecsCur2 As Long
        Dim lngDels As Long, arr_varDel() As Variant
        Dim blnDelAll As Boolean, blnHasDeletes As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim lngTmp01 As Long
        Dim lngX As Long, lngE As Long

        ' ** Array: arr_varDel().
        Const D_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const D_ACTNO As Integer = 0
        Const D_CNO   As Integer = 1

16910   With Me

16920     Set frmSub1 = .frmAccountContacts_Sub.Form
16930     Set frmSub2 = .frmAccountContacts_Sub2.Form

16940     lngRecsCur = frmSub1.RecCnt  ' ** Form Function: frmAccountContacts_Sub.
16950     lngRecsCur2 = frmSub2.RecCnt  ' ** Form Function: frmAccountContacts_Sub2.

16960     blnHasDeletes = .Has_Deletes
16970     If lngRecsCur > 0& Or lngRecsCur2 > 0& Or blnHasDeletes = True Then
16980       lngRecsChanged = Abs(frmSub1.Contact_Changed_sum)
16990       lngTmp01 = Abs(frmSub2.Contact_Changed_sum)
17000       If lngTmp01 > lngRecsChanged Then lngRecsChanged = lngTmp01
17010       .Contact_Changed_sum = lngRecsChanged
17020       If lngRecsChanged > 0& Or blnHasDeletes = True Then

              'IF THE CHANGE IS MADE IN ..FORMAT, DON'T FORGET TO CHANGE THE REAL ONE! AND VICE-VERSA!!
              'THIS MUST BE DONE WITH EACH AFTER_UPDATE!!!
17030         DoCmd.Hourglass True
17040         DoEvents

17050         If lngRecsChanged > 0& Then
17060           Set dbs = CurrentDb
17070           With dbs
17080             Set rst2 = .OpenRecordset("account", dbOpenDynaset, dbConsistent)
                  ' ** tblAccount_Contact_Staging, just Contact_Changed = True.
17090             Set qdf1 = .QueryDefs("qryAccount_Contacts_12")
17100             Set rst1 = qdf1.OpenRecordset
17110             With rst1
17120               .MoveLast
17130               lngRecs1 = .RecordCount
17140               .MoveFirst
17150               For lngX = 1& To lngRecs1
17160                 With rst2
17170                   .FindFirst "[accountno] = '" & rst1![accountno] & "'"
17180                   If .NoMatch = False Then
17190                     Select Case rst1![Contact_Number]
                          Case 1&
17200                       .Edit
                            'OK, HOW TO COORDINATE THIS?
                            'IF CORRESPONDING CHANGES WERE MADE WITH EACH AFTER_UPDATE,
                            'THEN ALL I HAVE TO WORRY ABOUT HERE IS THE REAL ONE!
17210                       ![Contact1_Name] = rst1![Contact_Name]
                            ' ** rst1![Contact_Name_LastFirst]
17220                       ![Contact1_Address1] = rst1![Contact_Address1]
17230                       ![Contact1_Address2] = rst1![Contact_Address2]
17240                       ![Contact1_City] = rst1![Contact_City]
17250                       Select Case rst1![opgLocContact]
                            Case 1
17260                         ![Contact1_State] = rst1![Contact_State]
17270                         ![Contact1_Zip] = rst1![Contact_Zip]
17280                       Case 2
17290                         ![Contact1_Country] = rst1![Contact_Country]
17300                         ![Contact1_PostalCode] = rst1![Contact_PostalCode]
17310                       End Select
                            ' ** rst1![Contact_Zip_Format]
17320                       ![Contact1_Phone1] = rst1![Contact_Phone1]
                            ' ** rst1![Contact_Phone1_Format]
17330                       ![Contact1_Phone2] = rst1![Contact_Phone2]
                            ' ** rst1![Contact_Phone2_Format]
17340                       ![Contact1_Fax] = rst1![Contact_Fax]
                            ' ** rst1![Contact_Fax_Format]
17350                       ![Contact1_Email] = rst1![Contact_Email]
17360                       .Update
17370                     Case 2&
17380                       .Edit
17390                       ![Contact2_Name] = rst1![Contact_Name]
                            ' ** rst1![Contact_Name_LastFirst]
17400                       ![Contact2_Address1] = rst1![Contact_Address1]
17410                       ![Contact2_Address2] = rst1![Contact_Address2]
17420                       ![Contact2_City] = rst1![Contact_City]
17430                       Select Case rst1![opgLocContact]
                            Case 1
17440                         ![Contact2_State] = rst1![Contact_State]
17450                         ![Contact2_Zip] = rst1![Contact_Zip]
17460                       Case 2
17470                         ![Contact2_Country] = rst1![Contact_Country]
17480                         ![Contact2_PostalCode] = rst1![Contact_PostalCode]
17490                       End Select
                            ' ** rst1![Contact_Zip_Format]
17500                       ![Contact2_Phone1] = rst1![Contact_Phone1]
                            ' ** rst1![Contact_Phone1_Format]
17510                       ![Contact2_Phone2] = rst1![Contact_Phone2]
                            ' ** rst1![Contact_Phone2_Format]
17520                       ![Contact2_Fax] = rst1![Contact_Fax]
                            ' ** rst1![Contact_Fax_Format]
17530                       ![Contact2_Email] = rst1![Contact_Email]
17540                       .Update
17550                     End Select
17560                   End If
17570                 End With  ' ** rst2.
17580                 If lngX < lngRecs1 Then .MoveNext
17590               Next  ' ** lngX.
17600               .Close
17610             End With  ' ** rst1.
17620             rst2.Close
17630             .Close
17640           End With  ' ** dbs.
17650           Set rst1 = Nothing
17660           Set rst2 = Nothing
17670           Set qdf1 = Nothing
17680           Set dbs = Nothing
17690         End If  ' ** lngRecsChanged.

17700         If blnHasDeletes = True Then
                ' ** If a contact was deleted, then replaced with a new one,
                ' ** it'll be caught in the first section below,
                ' ** and won't show up the the delete section.

17710           blnDelAll = False
17720           lngDels = 0&
17730           ReDim arr_varDel(D_ELEMS, 0)

17740           Set dbs = CurrentDb
17750           With dbs

                  ' ** qryAccount_Contacts_03_01 (qryAccount_Contacts_02c (Union of qryAccount_Contacts_02a
                  ' ** (Account, just Contact1), qryAccount_Contacts_02b (Account, just Contact2)),
                  ' ** with new field names), not in qryAccount_Contacts_03_02 (qryAccount_Contacts_03_01
                  ' ** (qryAccount_Contacts_02c (Union of qryAccount_Contacts_02a (Account, just Contact1),
                  ' ** qryAccount_Contacts_02b (Account, just Contact2)), with new field names),
                  ' ** just empty contacts), just good contacts.
17760             Set qdf1 = .QueryDefs("qryAccount_Contacts_03_03")
17770             Set rst1 = qdf1.OpenRecordset
17780             With rst1
17790               If .BOF = True And .EOF = True Then
                      ' ** Must have been created and deleted in this same session.
17800                 blnHasDeletes = False
17810               Else

17820                 .MoveLast
17830                 lngRecs1 = .RecordCount
17840                 .MoveFirst
                      ' ** tblAccount_Contact_Staging, just needed fields.
17850                 Set qdf2 = dbs.QueryDefs("qryAccount_Contacts_20")
17860                 Set rst2 = qdf2.OpenRecordset
17870                 With rst2
17880                   If .BOF = True And .EOF = True Then
                          ' ** All deleted.
17890                     blnDelAll = True
17900                   Else
17910                     .MoveLast
17920                     lngRecs2 = .RecordCount
17930                     .MoveFirst
17940                   End If
17950                 End With  ' ** rst2.

17960                 If blnDelAll = False Then
                        ' ** A contact that was deleted, then replaced, won't show up here.
17970                   For lngX = 1& To lngRecs1
17980                     rst2.FindFirst "[accountno] = '" & ![accountno] & "' And [Contact_Number] = " & CStr(![Contact_Number])
17990                     If rst2.NoMatch = True Then
18000                       lngDels = lngDels + 1&
18010                       lngE = lngDels - 1&
18020                       ReDim Preserve arr_varDel(D_ELEMS, lngE)
18030                       arr_varDel(D_ACTNO, lngE) = ![accountno]
18040                       arr_varDel(D_CNO, lngE) = ![Contact_Number]
18050                     End If
18060                     If lngX < lngRecs1 Then .MoveNext
18070                   Next  ' ** lngX.
18080                 End If  ' ** blnDelAll.

18090               End If  ' ** BOF, EOF.
18100               .Close
18110             End With  ' ** rst1.
18120             rst2.Close
18130             Set rst1 = Nothing
18140             Set rst2 = Nothing
18150             Set qdf1 = Nothing
18160             Set qdf2 = Nothing

18170             If blnDelAll = True Then
18180               If lngRecsCur = 0& Then  ' ** To be sure it's not a fluke.
18190                 Beep
18200                 msgResponse = MsgBox("You are about to delete all contact information from all accounts." & vbCrLf & _
                        "Are you sure you wish to proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete All Contacts")
18210                 If msgResponse = vbYes Then
                        ' ** Update Account, for all Contact info = Null.
18220                   Set qdf1 = .QueryDefs("qryAccount_Contacts_21a")
18230                   qdf1.Execute
18240                   Set qdf1 = Nothing
18250                   blnHasDeletes = False
18260                 End If
18270               End If  ' ** lngRecsCur.
18280             ElseIf lngDels > 0& Then
18290               For lngX = 0& To (lngDels - 1&)
18300                 Select Case arr_varDel(D_CNO, lngX)
                      Case 1&
                        ' ** Update Account, for Contact1 info = Null, by specified [actno].
18310                   Set qdf1 = .QueryDefs("qryAccount_Contacts_21b")
18320                   With qdf1.Parameters
18330                     ![actno] = arr_varDel(D_ACTNO, lngX)
18340                   End With
18350                   qdf1.Execute
18360                   Set qdf1 = Nothing
18370                 Case 2&
                        ' ** Update Account, for Contact2 info = Null, by specified [actno].
18380                   Set qdf1 = .QueryDefs("qryAccount_Contacts_21c")
18390                   qdf1.Execute
18400                   Set qdf1 = Nothing
18410                 End Select
18420               Next
18430               blnHasDeletes = False
18440             End If

18450             .Close
18460           End With  ' ** dbs.
18470           Set dbs = Nothing

18480         End If  ' ** blnHasDeletes.

18490         If blnFromClose = False And lngRecsChanged > 0& Then
18500           DoEvents
18510           Select Case .opgView  ' ** Only 1 of them has to be updated.
                Case .opgView_optWide.OptionValue
18520             frmSub1.ResetChanged  ' ** Form Procedure: frmAccountContacts_Sub.
18530             frmSub2.Refresh
18540             frmSub2.Requery
18550           Case .opgView_optCompact.OptionValue
18560             frmSub2.ResetChanged  ' ** Form Procedure: frmAccountContacts_Sub2.
18570             frmSub1.Refresh
18580             frmSub1.Requery
18590           End Select
18600         End If
18610         .Has_Deletes = blnHasDeletes

18620         Beep
18630         DoCmd.Hourglass False

18640       ElseIf blnFromClose = False Then
18650         Beep
18660         MsgBox "No records have been changed.", vbInformation + vbOKOnly, "Nothing To Do"
18670       End If
18680     Else
18690       Beep
18700       MsgBox "There are no records to update.", vbInformation + vbOKOnly, "Nothing To Do"
18710     End If
18720   End With

EXITP:
18730   Set frmSub1 = Nothing
18740   Set frmSub2 = Nothing
18750   Set rst1 = Nothing
18760   Set rst2 = Nothing
18770   Set qdf1 = Nothing
18780   Set qdf2 = Nothing
18790   Set dbs = Nothing
18800   Exit Sub

ERRH:
800     DoCmd.Hourglass False
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18920   intShiftDown = (Shift And acShiftMask) > 0
18930   intAltDown = (Shift And acAltMask) > 0
18940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18960     Select Case intRetVal
          Case vbKeyTab
18970       With Me
18980         intRetVal = 0
18990         .chkAlwaysUpdate.SetFocus
19000       End With
19010     End Select
19020   End If

        ' ** Shift keys.
19030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19040     Select Case intRetVal
          Case vbKeyTab
19050       With Me
19060         intRetVal = 0
19070         If .cmdDelete.Enabled = True Then
19080           .cmdDelete.SetFocus
19090         ElseIf .cmdAdd.Enabled = True Then
19100           .cmdAdd.SetFocus
19110         ElseIf .cmdExcel.Enabled = True Then
19120           .cmdExcel.SetFocus
19130         ElseIf .cmdWord.Enabled = True Then
19140           .cmdWord.SetFocus
19150         Else
19160           .chkPageOf.SetFocus
19170         End If
19180       End With
19190     End Select
19200   End If

EXITP:
19210   KeyCode = intRetVal
19220   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Public Sub opgView_AfterUpdate()

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgView_AfterUpdate"

19310   With Me

19320     DoCmd.SelectObject acForm, THIS_NAME, False
19330     .FocusHolder.SetFocus

19340     Select Case .opgView
          Case .opgView_optWide.OptionValue
19350       .opgView_optWide_lbl.FontBold = True
19360       .opgView_optCompact_lbl.FontBold = False
19370       .chkEnableCountry_Wide.Visible = True
19380       .chkEnableCountry_Compact.Visible = False
19390       Sub_View 1, lngForm_Width, lngFrm_Left_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig, _
              lngWinSub_Diff, Me  ' ** Module Procedure: modAccountContacts.
19400     Case .opgView_optCompact.OptionValue
19410       .opgView_optWide_lbl.FontBold = False
19420       .opgView_optCompact_lbl.FontBold = True
19430       .chkEnableCountry_Compact.Visible = True
19440       .chkEnableCountry_Wide.Visible = False
19450       Sub_View 2, lngForm_Width, lngFrm_Left_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig, _
              lngWinSub_Diff, Me ' ** Module Procedure: modAccountContacts.
19460     End Select
19470     .FocusHolder.SetFocus

19480   End With

EXITP:
19490   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub opgView_optWide_KeyDown(KeyCode As Integer, Shift As Integer)

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgView_optWide_KeyDown"

        Dim frm As Access.Form
        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19520   intShiftDown = (Shift And acShiftMask) > 0
19530   intAltDown = (Shift And acAltMask) > 0
19540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19570       With Me
19580         intRetVal = 0
19590         If .opgFormatName.Enabled = True Then
19600           .opgFormatName.SetFocus
19610         ElseIf .opgPrintAddress.Enabled = True Then
19620           .opgPrintAddress.SetFocus
19630         ElseIf .chkShowAcctNum.Enabled = True Then
19640           .chkShowAcctNum.SetFocus
19650         ElseIf .chkShowShortName.Enabled = True Then
19660           .chkShowShortName.SetFocus
19670         Else
19680           Select Case .opgView
                Case .opgView_optWide.OptionValue
19690             .chkEnableCountry_Wide.SetFocus
19700           Case .opgView_optCompact.OptionValue
19710             .chkEnableCountry_Compact.SetFocus
19720           End Select
19730         End If
19740       End With
19750     End Select
19760   End If

        ' ** Shift keys.
19770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19790       With Me
19800         intRetVal = 0
19810         Select Case .opgView
              Case .opgView_optWide.OptionValue
19820           Set frm = .frmAccountContacts_Sub.Form
19830         Case .opgView_optCompact.OptionValue
19840           Set frm = .frmAccountContacts_Sub2.Form
19850         End Select
19860         lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountContacts_Sub, frmAccountContacts_Sub2.
19870         If lngRecsCur > 0& Then
19880           Select Case .opgView
                Case .opgView_optWide.OptionValue
19890             .frmAccountContacts_Sub.SetFocus
19900           Case .opgView_optCompact.OptionValue
19910             .frmAccountContacts_Sub2.SetFocus
19920           End Select
19930           frm.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountContacts_Sub, frmAccountContacts_Sub2.
19940         Else
19950           .cmdClose.SetFocus
19960         End If
19970       End With
19980     End Select
19990   End If

EXITP:
20000   Set frm = Nothing
20010   KeyCode = intRetVal
20020   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgView_optCompact_KeyDown(KeyCode As Integer, Shift As Integer)

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgView_optCompact_KeyDown"

        Dim frm As Access.Form
        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20120   intShiftDown = (Shift And acShiftMask) > 0
20130   intAltDown = (Shift And acAltMask) > 0
20140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20170       With Me
20180         intRetVal = 0
20190         If .opgFormatName.Enabled = True Then
20200           .opgFormatName.SetFocus
20210         ElseIf .opgPrintAddress.Enabled = True Then
20220           .opgPrintAddress.SetFocus
20230         ElseIf .chkShowAcctNum.Enabled = True Then
20240           .chkShowAcctNum.SetFocus
20250         ElseIf .chkShowShortName.Enabled = True Then
20260           .chkShowShortName.SetFocus
20270         Else
20280           Select Case .opgView
                Case .opgView_optWide.OptionValue
20290             .chkEnableCountry_Wide.SetFocus
20300           Case .opgView_optCompact.OptionValue
20310             .chkEnableCountry_Compact.SetFocus
20320           End Select
20330         End If
20340       End With
20350     End Select
20360   End If

        ' ** Shift keys.
20370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20390       With Me
20400         intRetVal = 0
20410         Select Case .opgView
              Case .opgView_optWide.OptionValue
20420           Set frm = .frmAccountContacts_Sub.Form
20430         Case .opgView_optCompact.OptionValue
20440           Set frm = .frmAccountContacts_Sub2.Form
20450         End Select
20460         lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountContacts_Sub, frmAccountContacts_Sub2.
20470         If lngRecsCur > 0& Then
20480           Select Case .opgView
                Case .opgView_optWide.OptionValue
20490             .frmAccountContacts_Sub.SetFocus
20500           Case .opgView_optCompact.OptionValue
20510             .frmAccountContacts_Sub2.SetFocus
20520           End Select
20530           frm.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountContacts_Sub, frmAccountContacts_Sub2.
20540         Else
20550           .cmdClose.SetFocus
20560         End If
20570       End With
20580     End Select
20590   End If

EXITP:
20600   Set frm = Nothing
20610   KeyCode = intRetVal
20620   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Public Sub opgFormatName_AfterUpdate()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFormatName_AfterUpdate"

        Dim frm As Access.Form, frmSub1 As Access.Form, frmSub2 As Access.Form

20710   With Me

20720     Select Case .opgView
          Case .opgView_optWide.OptionValue
20730       Set frm = .frmAccountContacts_Sub.Form
20740     Case .opgView_optCompact.OptionValue
20750       Set frm = .frmAccountContacts_Sub2.Form
20760     End Select
20770     Set frmSub1 = .frmAccountContacts_Sub.Form
20780     Set frmSub2 = .frmAccountContacts_Sub2.Form

20790     lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountContacts_Sub, frmAccountContacts_Sub2.
20800     Select Case .opgFormatName
          Case .opgFormatName_optAsWritten.OptionValue
20810       .opgFormatName_optAsWritten_lbl.FontBold = True
20820       .opgFormatName_optLastFirst_lbl.FontBold = False
20830       If lngRecsCur > 0& Then
20840         frmSub1.FocusHolder.SetFocus
20850         frmSub2.FocusHolder.SetFocus
20860       End If
20870       frmSub1.FocusHolder.SetFocus
20880       frmSub1.Contact_Name.Enabled = True
20890       frmSub1.Contact_Name.Visible = True
20900       frmSub1.Contact_Name_LastFirst.Visible = False
20910       frmSub1.Contact_Name_LastFirst.Enabled = False
20920       frmSub1.FocusHolder.SetFocus
20930       frmSub2.Contact_Name.Enabled = True
20940       frmSub2.Contact_Name.Visible = True
20950       frmSub2.Contact_Name_LastFirst.Visible = False
20960       frmSub2.Contact_Name_LastFirst.Enabled = False
20970     Case .opgFormatName_optLastFirst.OptionValue
20980       .opgFormatName_optAsWritten_lbl.FontBold = False
20990       .opgFormatName_optLastFirst_lbl.FontBold = True
21000       If lngRecsCur > 0& Then
21010         frmSub1.FocusHolder.SetFocus
21020         frmSub2.FocusHolder.SetFocus
21030       End If
21040       frmSub1.FocusHolder.SetFocus
21050       frmSub1.Contact_Name_LastFirst.Enabled = True
21060       frmSub1.Contact_Name_LastFirst.Visible = True
21070       frmSub1.Contact_Name.Visible = False
21080       frmSub1.Contact_Name.Enabled = False
21090       frmSub2.FocusHolder.SetFocus
21100       frmSub2.Contact_Name_LastFirst.Enabled = True
21110       frmSub2.Contact_Name_LastFirst.Visible = True
21120       frmSub2.Contact_Name.Visible = False
21130       frmSub2.Contact_Name.Enabled = False
21140     End Select

21150     SetReset IsntDefault  ' ** Procedure/Function: Below.

21160   End With

EXITP:
21170   Set frm = Nothing
21180   Set frmSub1 = Nothing
21190   Set frmSub2 = Nothing
21200   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Ignore.
810     Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgFormatName_optAsWritten_KeyDown(KeyCode As Integer, Shift As Integer)

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFormatName_optAsWritten_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21320   intShiftDown = (Shift And acShiftMask) > 0
21330   intAltDown = (Shift And acAltMask) > 0
21340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21370       With Me
21380         intRetVal = 0
21390         .opgFormatZip.SetFocus
21400       End With
21410     End Select
21420   End If

        ' ** Shift keys.
21430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21450       With Me
21460         intRetVal = 0
21470         .opgView.SetFocus
21480       End With
21490     End Select
21500   End If

EXITP:
21510   KeyCode = intRetVal
21520   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgFormatName_optLastFirst_KeyDown(KeyCode As Integer, Shift As Integer)

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFormatName_optLastFirst_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21620   intShiftDown = (Shift And acShiftMask) > 0
21630   intAltDown = (Shift And acAltMask) > 0
21640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21670       With Me
21680         intRetVal = 0
21690         .opgFormatZip.SetFocus
21700       End With
21710     End Select
21720   End If

        ' ** Shift keys.
21730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21750       With Me
21760         intRetVal = 0
21770         .opgView.SetFocus
21780       End With
21790     End Select
21800   End If

EXITP:
21810   KeyCode = intRetVal
21820   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Public Sub opgFormatZip_AfterUpdate()

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFormatZip_AfterUpdate"

        Dim frm As Access.Form, frmSub1 As Access.Form, frmSub2 As Access.Form

21910   With Me

21920     Select Case .opgView
          Case .opgView_optWide.OptionValue
21930       Set frm = .frmAccountContacts_Sub.Form
21940     Case .opgView_optCompact.OptionValue
21950       Set frm = .frmAccountContacts_Sub2.Form
21960     End Select
21970     Set frmSub1 = .frmAccountContacts_Sub.Form
21980     Set frmSub2 = .frmAccountContacts_Sub2.Form

21990     lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountContacts_Sub, frmAccountContacts_Sub2.
22000     Select Case .opgFormatZip
          Case .opgFormatZip_optFormatted.OptionValue
22010       .opgFormatZip_optFormatted_lbl.FontBold = True
22020       .opgFormatZip_optUnformatted_lbl.FontBold = False
22030       If lngRecsCur > 0& Then
22040         frmSub1.FocusHolder.SetFocus
22050         frmSub2.FocusHolder.SetFocus
22060       End If
22070       frmSub1.Contact_Zip_Format.Enabled = True
22080       frmSub1.Contact_Zip_Format.Visible = True
22090       frmSub1.Contact_Zip.Visible = False
22100       frmSub1.Contact_Zip.Enabled = False
22110       frmSub2.Contact_Zip_Format.Enabled = True
22120       frmSub2.Contact_Zip_Format.Visible = True
22130       frmSub2.Contact_Zip.Visible = False
22140       frmSub2.Contact_Zip.Enabled = False
22150     Case .opgFormatZip_optUnformatted.OptionValue
22160       .opgFormatZip_optFormatted_lbl.FontBold = False
22170       .opgFormatZip_optUnformatted_lbl.FontBold = True
22180       If lngRecsCur > 0& Then
22190         frmSub1.FocusHolder.SetFocus
22200         frmSub2.FocusHolder.SetFocus
22210       End If
22220       frmSub1.Contact_Zip.Enabled = True
22230       frmSub1.Contact_Zip.Visible = True
22240       frmSub1.Contact_Zip_Format.Visible = False
22250       frmSub1.Contact_Zip_Format.Enabled = False
22260       frmSub2.Contact_Zip.Enabled = True
22270       frmSub2.Contact_Zip.Visible = True
22280       frmSub2.Contact_Zip_Format.Visible = False
22290       frmSub2.Contact_Zip_Format.Enabled = False
22300     End Select

22310     SetReset IsntDefault  ' ** Procedure/Function: Below.

22320   End With

EXITP:
22330   Set frm = Nothing
22340   Set frmSub1 = Nothing
22350   Set frmSub2 = Nothing
22360   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Ignore.
810     Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Public Sub opgFormatZip_optFormatted_KeyDown(KeyCode As Integer, Shift As Integer)

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFormatZip_optFormatted_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22420   intShiftDown = (Shift And acShiftMask) > 0
22430   intAltDown = (Shift And acAltMask) > 0
22440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22470       With Me
22480         intRetVal = 0
22490         .opgFormatPhone.SetFocus
22500       End With
22510     End Select
22520   End If

        ' ** Shift keys.
22530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22550       With Me
22560         intRetVal = 0
22570         .opgFormatName.SetFocus
22580       End With
22590     End Select
22600   End If

EXITP:
22610   KeyCode = intRetVal
22620   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgFormatZip_optUnformatted_KeyDown(KeyCode As Integer, Shift As Integer)

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFormatZip_optUnformatted_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22720   intShiftDown = (Shift And acShiftMask) > 0
22730   intAltDown = (Shift And acAltMask) > 0
22740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22770       With Me
22780         intRetVal = 0
22790         .opgFormatPhone.SetFocus
22800       End With
22810     End Select
22820   End If

        ' ** Shift keys.
22830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22850       With Me
22860         intRetVal = 0
22870         .opgFormatName.SetFocus
22880       End With
22890     End Select
22900   End If

EXITP:
22910   KeyCode = intRetVal
22920   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Public Sub opgFormatPhone_AfterUpdate()

23000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFormatPhone_AfterUpdate"

        Dim frm As Access.Form, frmSub1 As Access.Form, frmSub2 As Access.Form

23010   With Me

23020     Select Case .opgView
          Case .opgView_optWide.OptionValue
23030       Set frm = .frmAccountContacts_Sub.Form
23040     Case .opgView_optCompact.OptionValue
23050       Set frm = .frmAccountContacts_Sub2.Form
23060     End Select
23070     Set frmSub1 = .frmAccountContacts_Sub.Form
23080     Set frmSub2 = .frmAccountContacts_Sub2.Form

23090     lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountContacts_Sub, frmAccountContacts_Sub2.
23100     Select Case .opgFormatPhone
          Case .opgFormatPhone_optFormatted.OptionValue
23110       .opgFormatPhone_optFormatted_lbl.FontBold = True
23120       .opgFormatPhone_optUnformatted_lbl.FontBold = False
23130       If lngRecsCur > 0& Then
23140         frmSub1.FocusHolder.SetFocus
23150         frmSub2.FocusHolder.SetFocus
23160       End If
23170       frmSub1.Contact_Phone1_Format.Enabled = True
23180       frmSub1.Contact_Phone1_Format.Visible = True
23190       frmSub1.Contact_Phone1.Visible = False
23200       frmSub1.Contact_Phone1.Enabled = False
23210       frmSub1.Contact_Phone2_Format.Enabled = True
23220       frmSub1.Contact_Phone2_Format.Visible = True
23230       frmSub1.Contact_Phone2.Visible = False
23240       frmSub1.Contact_Phone2.Enabled = False
23250       frmSub1.Contact_Fax_Format.Enabled = True
23260       frmSub1.Contact_Fax_Format.Visible = True
23270       frmSub1.Contact_Fax.Visible = False
23280       frmSub1.Contact_Fax.Enabled = False
23290       frmSub2.Contact_Phone1_Format.Enabled = True
23300       frmSub2.Contact_Phone1_Format.Visible = True
23310       frmSub2.Contact_Phone1.Visible = False
23320       frmSub2.Contact_Phone1.Enabled = False
23330       frmSub2.Contact_Phone2_Format.Enabled = True
23340       frmSub2.Contact_Phone2_Format.Visible = True
23350       frmSub2.Contact_Phone2.Visible = False
23360       frmSub2.Contact_Phone2.Enabled = False
23370       frmSub2.Contact_Fax_Format.Enabled = True
23380       frmSub2.Contact_Fax_Format.Visible = True
23390       frmSub2.Contact_Fax.Visible = False
23400       frmSub2.Contact_Fax.Enabled = False
23410     Case .opgFormatPhone_optUnformatted.OptionValue
23420       .opgFormatPhone_optFormatted_lbl.FontBold = False
23430       .opgFormatPhone_optUnformatted_lbl.FontBold = True
23440       If lngRecsCur > 0& Then
23450         frmSub1.FocusHolder.SetFocus
23460         frmSub2.FocusHolder.SetFocus
23470       End If
23480       frmSub1.Contact_Phone1.Enabled = True
23490       frmSub1.Contact_Phone1.Visible = True
23500       frmSub1.Contact_Phone1_Format.Visible = False
23510       frmSub1.Contact_Phone1_Format.Enabled = False
23520       frmSub1.Contact_Phone2.Enabled = True
23530       frmSub1.Contact_Phone2.Visible = True
23540       frmSub1.Contact_Phone2_Format.Visible = False
23550       frmSub1.Contact_Phone2_Format.Enabled = False
23560       frmSub1.Contact_Fax.Enabled = True
23570       frmSub1.Contact_Fax.Visible = True
23580       frmSub1.Contact_Fax_Format.Visible = False
23590       frmSub1.Contact_Fax_Format.Enabled = False
23600       frmSub2.Contact_Phone1.Enabled = True
23610       frmSub2.Contact_Phone1.Visible = True
23620       frmSub2.Contact_Phone1_Format.Visible = False
23630       frmSub2.Contact_Phone1_Format.Enabled = False
23640       frmSub2.Contact_Phone2.Enabled = True
23650       frmSub2.Contact_Phone2.Visible = True
23660       frmSub2.Contact_Phone2_Format.Visible = False
23670       frmSub2.Contact_Phone2_Format.Enabled = False
23680       frmSub2.Contact_Fax.Enabled = True
23690       frmSub2.Contact_Fax.Visible = True
23700       frmSub2.Contact_Fax_Format.Visible = False
23710       frmSub2.Contact_Fax_Format.Enabled = False
23720     End Select

23730     SetReset IsntDefault  ' ** Procedure/Function: Below.

23740   End With

EXITP:
23750   Set frm = Nothing
23760   Set frmSub1 = Nothing
23770   Set frmSub2 = Nothing
23780   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Ignore.
810     Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgFormatPhone_optFormatted_KeyDown(KeyCode As Integer, Shift As Integer)

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFormatPhone_optFormatted_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23820   intShiftDown = (Shift And acShiftMask) > 0
23830   intAltDown = (Shift And acAltMask) > 0
23840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23870       With Me
23880         intRetVal = 0
23890         .opgPrintAddress.SetFocus
23900       End With
23910     End Select
23920   End If

        ' ** Shift keys.
23930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23950       With Me
23960         intRetVal = 0
23970         .opgFormatZip.SetFocus
23980       End With
23990     End Select
24000   End If

EXITP:
24010   KeyCode = intRetVal
24020   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgFormatPhone_optUnformatted_KeyDown(KeyCode As Integer, Shift As Integer)

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFormatPhone_optUnformatted_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24120   intShiftDown = (Shift And acShiftMask) > 0
24130   intAltDown = (Shift And acAltMask) > 0
24140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24170       With Me
24180         intRetVal = 0
24190         .opgPrintAddress.SetFocus
24200       End With
24210     End Select
24220   End If

        ' ** Shift keys.
24230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24250       With Me
24260         intRetVal = 0
24270         .opgFormatZip.SetFocus
24280       End With
24290     End Select
24300   End If

EXITP:
24310   KeyCode = intRetVal
24320   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgPrint_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrint_box_MouseMove"

24410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
24420   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Public Sub opgPrintAddress_AfterUpdate()

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintAddress_AfterUpdate"

24510   With Me
24520     Select Case .opgPrintAddress
          Case .opgPrintAddress_optSeparate.OptionValue
24530       .opgPrintAddress_optSeparate_lbl.FontBold = True
24540       .opgPrintAddress_optCombined_lbl.FontBold = False
24550       .opgPrintAddress_optNone_lbl.FontBold = False
24560     Case .opgPrintAddress_optCombined.OptionValue
24570       .opgPrintAddress_optSeparate_lbl.FontBold = False
24580       .opgPrintAddress_optCombined_lbl.FontBold = True
24590       .opgPrintAddress_optNone_lbl.FontBold = False
24600     Case .opgPrintAddress_optNone.OptionValue
24610       .opgPrintAddress_optSeparate_lbl.FontBold = False
24620       .opgPrintAddress_optCombined_lbl.FontBold = False
24630       .opgPrintAddress_optNone_lbl.FontBold = True
24640     End Select
24650     SetReset IsntDefault  ' ** Procedure/Function: Below.
24660   End With

EXITP:
24670   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub opgPrintAddress_optSeparate_KeyDown(KeyCode As Integer, Shift As Integer)

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintAddress_optSeparate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24720   intShiftDown = (Shift And acShiftMask) > 0
24730   intAltDown = (Shift And acAltMask) > 0
24740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24770       With Me
24780         intRetVal = 0
24790         If .opgPrintCSZ.Visible = True Then
24800           .opgPrintCSZ.SetFocus
24810         ElseIf .opgPrintCSZCP.Visible = True Then
24820           .opgPrintCSZCP.SetFocus
24830         End If
24840       End With
24850     End Select
24860   End If

        ' ** Shift keys.
24870   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24880     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24890       With Me
24900         intRetVal = 0
24910         .opgFormatPhone.SetFocus
24920       End With
24930     End Select
24940   End If

EXITP:
24950   KeyCode = intRetVal
24960   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgPrintAddress_optCombined_KeyDown(KeyCode As Integer, Shift As Integer)

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintAddress_optCombined_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25020   intShiftDown = (Shift And acShiftMask) > 0
25030   intAltDown = (Shift And acAltMask) > 0
25040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25070       With Me
25080         intRetVal = 0
25090         If .opgPrintCSZ.Visible = True Then
25100           .opgPrintCSZ.SetFocus
25110         ElseIf .opgPrintCSZCP.Visible = True Then
25120           .opgPrintCSZCP.SetFocus
25130         End If
25140       End With
25150     End Select
25160   End If

        ' ** Shift keys.
25170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25180     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25190       With Me
25200         intRetVal = 0
25210         .opgFormatPhone.SetFocus
25220       End With
25230     End Select
25240   End If

EXITP:
25250   KeyCode = intRetVal
25260   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgPrintAddress_optNone_KeyDown(KeyCode As Integer, Shift As Integer)

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintAddress_optNone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25320   intShiftDown = (Shift And acShiftMask) > 0
25330   intAltDown = (Shift And acAltMask) > 0
25340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25370       With Me
25380         intRetVal = 0
25390         If .opgPrintCSZ.Visible = True Then
25400           .opgPrintCSZ.SetFocus
25410         ElseIf .opgPrintCSZCP.Visible = True Then
25420           .opgPrintCSZCP.SetFocus
25430         End If
25440       End With
25450     End Select
25460   End If

        ' ** Shift keys.
25470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25490       With Me
25500         intRetVal = 0
25510         .opgFormatPhone.SetFocus
25520       End With
25530     End Select
25540   End If

EXITP:
25550   KeyCode = intRetVal
25560   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Public Sub opgPrintCSZ_AfterUpdate()

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintCSZ_AfterUpdate"

25610   With Me
25620     Select Case .opgPrintCSZ
          Case .opgPrintCSZ_optSeparate.OptionValue
25630       .opgPrintCSZ_optSeparate_lbl.FontBold = True
25640       .opgPrintCSZ_optCombined_lbl.FontBold = False
25650       .opgPrintCSZ_optNone_lbl.FontBold = False
25660     Case .opgPrintCSZ_optCombined.OptionValue
25670       .opgPrintCSZ_optSeparate_lbl.FontBold = False
25680       .opgPrintCSZ_optCombined_lbl.FontBold = True
25690       .opgPrintCSZ_optNone_lbl.FontBold = False
25700     Case .opgPrintCSZ_optNone.OptionValue
25710       .opgPrintCSZ_optSeparate_lbl.FontBold = False
25720       .opgPrintCSZ_optCombined_lbl.FontBold = False
25730       .opgPrintCSZ_optNone_lbl.FontBold = True
25740     End Select
25750     SetReset IsntDefault  ' ** Procedure/Function: Below.
25760   End With

EXITP:
25770   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub opgPrintCSZ_optSeparate_KeyDown(KeyCode As Integer, Shift As Integer)

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintCSZ_optSeparate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25820   intShiftDown = (Shift And acShiftMask) > 0
25830   intAltDown = (Shift And acAltMask) > 0
25840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25870       With Me
25880         intRetVal = 0
25890         .opgPrintPhone.SetFocus
25900       End With
25910     End Select
25920   End If

        ' ** Shift keys.
25930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25950       With Me
25960         intRetVal = 0
25970         .opgPrintAddress.SetFocus
25980       End With
25990     End Select
26000   End If

EXITP:
26010   KeyCode = intRetVal
26020   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgPrintCSZ_optCombined_KeyDown(KeyCode As Integer, Shift As Integer)

26100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintCSZ_optCombined_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26120   intShiftDown = (Shift And acShiftMask) > 0
26130   intAltDown = (Shift And acAltMask) > 0
26140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26170       With Me
26180         intRetVal = 0
26190         .opgPrintPhone.SetFocus
26200       End With
26210     End Select
26220   End If

        ' ** Shift keys.
26230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26250       With Me
26260         intRetVal = 0
26270         .opgPrintAddress.SetFocus
26280       End With
26290     End Select
26300   End If

EXITP:
26310   KeyCode = intRetVal
26320   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgPrintCSZ_optNone_KeyDown(KeyCode As Integer, Shift As Integer)

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintCSZ_optNone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26420   intShiftDown = (Shift And acShiftMask) > 0
26430   intAltDown = (Shift And acAltMask) > 0
26440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26470       With Me
26480         intRetVal = 0
26490         .opgPrintPhone.SetFocus
26500       End With
26510     End Select
26520   End If

        ' ** Shift keys.
26530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26550       With Me
26560         intRetVal = 0
26570         .opgPrintAddress.SetFocus
26580       End With
26590     End Select
26600   End If

EXITP:
26610   KeyCode = intRetVal
26620   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Public Sub opgPrintCSZCP_AfterUpdate()

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintCSZCP_AfterUpdate"

26710   With Me
26720     Select Case .opgPrintCSZCP
          Case .opgPrintCSZCP_optSeparate.OptionValue
26730       .opgPrintCSZCP_optSeparate_lbl.FontBold = True
26740       .opgPrintCSZCP_optCombined_lbl.FontBold = False
26750       .opgPrintCSZCP_optNone_lbl.FontBold = False
26760     Case .opgPrintCSZCP_optCombined.OptionValue
26770       .opgPrintCSZCP_optSeparate_lbl.FontBold = False
26780       .opgPrintCSZCP_optCombined_lbl.FontBold = True
26790       .opgPrintCSZCP_optNone_lbl.FontBold = False
26800     Case .opgPrintCSZCP_optNone.OptionValue
26810       .opgPrintCSZCP_optSeparate_lbl.FontBold = False
26820       .opgPrintCSZCP_optCombined_lbl.FontBold = False
26830       .opgPrintCSZCP_optNone_lbl.FontBold = True
26840     End Select
26850     SetReset IsntDefault  ' ** Procedure/Function: Below.
26860   End With

EXITP:
26870   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub opgPrintCSZCP_optSeparate_KeyDown(KeyCode As Integer, Shift As Integer)

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintCSZCP_optSeparate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26920   intShiftDown = (Shift And acShiftMask) > 0
26930   intAltDown = (Shift And acAltMask) > 0
26940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26970       With Me
26980         intRetVal = 0
26990         .opgPrintPhone.SetFocus
27000       End With
27010     End Select
27020   End If

        ' ** Shift keys.
27030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27050       With Me
27060         intRetVal = 0
27070         .opgPrintAddress.SetFocus
27080       End With
27090     End Select
27100   End If

EXITP:
27110   KeyCode = intRetVal
27120   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgPrintCSZCP_optCombined_KeyDown(KeyCode As Integer, Shift As Integer)

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintCSZCP_optCombined_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27220   intShiftDown = (Shift And acShiftMask) > 0
27230   intAltDown = (Shift And acAltMask) > 0
27240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27270       With Me
27280         intRetVal = 0
27290         .opgPrintPhone.SetFocus
27300       End With
27310     End Select
27320   End If

        ' ** Shift keys.
27330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27350       With Me
27360         intRetVal = 0
27370         .opgPrintAddress.SetFocus
27380       End With
27390     End Select
27400   End If

EXITP:
27410   KeyCode = intRetVal
27420   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgPrintCSZCP_optNone_KeyDown(KeyCode As Integer, Shift As Integer)

27500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintCSZCP_optNone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27520   intShiftDown = (Shift And acShiftMask) > 0
27530   intAltDown = (Shift And acAltMask) > 0
27540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27570       With Me
27580         intRetVal = 0
27590         .opgPrintPhone.SetFocus
27600       End With
27610     End Select
27620   End If

        ' ** Shift keys.
27630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27650       With Me
27660         intRetVal = 0
27670         .opgPrintAddress.SetFocus
27680       End With
27690     End Select
27700   End If

EXITP:
27710   KeyCode = intRetVal
27720   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Public Sub opgPrintPhone_AfterUpdate()

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintPhone_AfterUpdate"

27810   With Me
27820     Select Case .opgPrintPhone
          Case .opgPrintPhone_optSeparate.OptionValue
27830       .opgPrintPhone_optSeparate_lbl.FontBold = True
27840       .opgPrintPhone_optCombined_lbl.FontBold = False
27850       .opgPrintPhone_opt1Only_lbl.FontBold = False
27860       .opgPrintPhone_optNone_lbl.FontBold = False
27870     Case .opgPrintPhone_optCombined.OptionValue
27880       .opgPrintPhone_optSeparate_lbl.FontBold = False
27890       .opgPrintPhone_optCombined_lbl.FontBold = True
27900       .opgPrintPhone_opt1Only_lbl.FontBold = False
27910       .opgPrintPhone_optNone_lbl.FontBold = False
27920     Case .opgPrintPhone_opt1Only.OptionValue
27930       .opgPrintPhone_optSeparate_lbl.FontBold = False
27940       .opgPrintPhone_optCombined_lbl.FontBold = False
27950       .opgPrintPhone_opt1Only_lbl.FontBold = True
27960       .opgPrintPhone_optNone_lbl.FontBold = False
27970     Case .opgPrintPhone_optNone.OptionValue
27980       .opgPrintPhone_optSeparate_lbl.FontBold = False
27990       .opgPrintPhone_optCombined_lbl.FontBold = False
28000       .opgPrintPhone_opt1Only_lbl.FontBold = False
28010       .opgPrintPhone_optNone_lbl.FontBold = True
28020     End Select
28030     SetReset IsntDefault  ' ** Procedure/Function: Below.
28040   End With

EXITP:
28050   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub opgPrintPhone_optSeparate_KeyDown(KeyCode As Integer, Shift As Integer)

28100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintPhone_optSeparate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28120   intShiftDown = (Shift And acShiftMask) > 0
28130   intAltDown = (Shift And acAltMask) > 0
28140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28170       With Me
28180         intRetVal = 0
28190         .opgPrintFax.SetFocus
28200       End With
28210     End Select
28220   End If

        ' ** Shift keys.
28230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28250       With Me
28260         intRetVal = 0
28270         If .opgPrintCSZ.Visible = True Then
28280           .opgPrintCSZ.SetFocus
28290         ElseIf .opgPrintCSZCP.Visible = True Then
28300           .opgPrintCSZCP.SetFocus
28310         End If
28320       End With
28330     End Select
28340   End If

EXITP:
28350   KeyCode = intRetVal
28360   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgPrintPhone_optCombined_KeyDown(KeyCode As Integer, Shift As Integer)

28400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintPhone_optCombined_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28420   intShiftDown = (Shift And acShiftMask) > 0
28430   intAltDown = (Shift And acAltMask) > 0
28440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28470       With Me
28480         intRetVal = 0
28490         .opgPrintFax.SetFocus
28500       End With
28510     End Select
28520   End If

        ' ** Shift keys.
28530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28550       With Me
28560         intRetVal = 0
28570         If .opgPrintCSZ.Visible = True Then
28580           .opgPrintCSZ.SetFocus
28590         ElseIf .opgPrintCSZCP.Visible = True Then
28600           .opgPrintCSZCP.SetFocus
28610         End If
28620       End With
28630     End Select
28640   End If

EXITP:
28650   KeyCode = intRetVal
28660   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgPrintPhone_opt1Only_KeyDown(KeyCode As Integer, Shift As Integer)

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintPhone_opt1Only_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28720   intShiftDown = (Shift And acShiftMask) > 0
28730   intAltDown = (Shift And acAltMask) > 0
28740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28770       With Me
28780         intRetVal = 0
28790         .opgPrintFax.SetFocus
28800       End With
28810     End Select
28820   End If

        ' ** Shift keys.
28830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28850       With Me
28860         intRetVal = 0
28870         If .opgPrintCSZ.Visible = True Then
28880           .opgPrintCSZ.SetFocus
28890         ElseIf .opgPrintCSZCP.Visible = True Then
28900           .opgPrintCSZCP.SetFocus
28910         End If
28920       End With
28930     End Select
28940   End If

EXITP:
28950   KeyCode = intRetVal
28960   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgPrintPhone_optNone_KeyDown(KeyCode As Integer, Shift As Integer)

29000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintPhone_optNone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29020   intShiftDown = (Shift And acShiftMask) > 0
29030   intAltDown = (Shift And acAltMask) > 0
29040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29070       With Me
29080         intRetVal = 0
29090         .opgPrintFax.SetFocus
29100       End With
29110     End Select
29120   End If

        ' ** Shift keys.
29130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29150       With Me
29160         intRetVal = 0
29170         If .opgPrintCSZ.Visible = True Then
29180           .opgPrintCSZ.SetFocus
29190         ElseIf .opgPrintCSZCP.Visible = True Then
29200           .opgPrintCSZCP.SetFocus
29210         End If
29220       End With
29230     End Select
29240   End If

EXITP:
29250   KeyCode = intRetVal
29260   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Public Sub opgPrintFax_AfterUpdate()

29300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintFax_AfterUpdate"

29310   With Me
29320     Select Case .opgPrintFax
          Case .opgPrintFax_optNone.OptionValue
29330       .opgPrintFax_optNone_lbl.FontBold = True
29340       .opgPrintFax_optPrint_lbl.FontBold = False
29350     Case .opgPrintFax_optPrint.OptionValue
29360       .opgPrintFax_optNone_lbl.FontBold = False
29370       .opgPrintFax_optPrint_lbl.FontBold = True
29380     End Select
29390     SetReset IsntDefault  ' ** Procedure/Function: Below.
29400   End With

EXITP:
29410   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub opgPrintFax_optPrint_KeyDown(KeyCode As Integer, Shift As Integer)

29500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintFax_optPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29520   intShiftDown = (Shift And acShiftMask) > 0
29530   intAltDown = (Shift And acAltMask) > 0
29540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29570       With Me
29580         intRetVal = 0
29590         .opgPrintEmail.SetFocus
29600       End With
29610     End Select
29620   End If

        ' ** Shift keys.
29630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29650       With Me
29660         intRetVal = 0
29670         .opgPrintPhone.SetFocus
29680       End With
29690     End Select
29700   End If

EXITP:
29710   KeyCode = intRetVal
29720   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgPrintFax_optNone_KeyDown(KeyCode As Integer, Shift As Integer)

29800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintFax_optNone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29820   intShiftDown = (Shift And acShiftMask) > 0
29830   intAltDown = (Shift And acAltMask) > 0
29840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29870       With Me
29880         intRetVal = 0
29890         .opgPrintEmail.SetFocus
29900       End With
29910     End Select
29920   End If

        ' ** Shift keys.
29930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29950       With Me
29960         intRetVal = 0
29970         .opgPrintPhone.SetFocus
29980       End With
29990     End Select
30000   End If

EXITP:
30010   KeyCode = intRetVal
30020   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Public Sub opgPrintEmail_AfterUpdate()

30100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintEmail_AfterUpdate"

30110   With Me
30120     Select Case .opgPrintEmail
          Case .opgPrintEmail_optNone.OptionValue
30130       .opgPrintEmail_optNone_lbl.FontBold = True
30140       .opgPrintEmail_optPrint_lbl.FontBold = False
30150     Case .opgPrintEmail_optPrint.OptionValue
30160       .opgPrintEmail_optNone_lbl.FontBold = False
30170       .opgPrintEmail_optPrint_lbl.FontBold = True
30180     End Select
30190     SetReset IsntDefault  ' ** Procedure/Function: Below.
30200   End With

EXITP:
30210   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub opgPrintEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintEmail_MouseMove"

30310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
30320   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub opgPrintEmail_optPrint_KeyDown(KeyCode As Integer, Shift As Integer)

30400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintEmail_optPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
30420   intShiftDown = (Shift And acShiftMask) > 0
30430   intAltDown = (Shift And acAltMask) > 0
30440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30470       With Me
30480         intRetVal = 0
30490         If .cmdReset.Enabled = True Then
30500           .cmdReset.SetFocus
30510         ElseIf .chkShowAcctNum.Enabled = True Then
30520           .chkShowAcctNum.SetFocus
30530         ElseIf .chkShowShortName.Enabled = True Then
30540           .chkShowShortName.SetFocus
30550         Else
30560           Select Case .opgView
                Case .opgView_optWide.OptionValue
30570             .chkEnableCountry_Wide.SetFocus
30580           Case .opgView_optCompact.OptionValue
30590             .chkEnableCountry_Compact.SetFocus
30600           End Select
30610         End If
30620       End With
30630     End Select
30640   End If

        ' ** Shift keys.
30650   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
30660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30670       With Me
30680         intRetVal = 0
30690         .opgPrintFax.SetFocus
30700       End With
30710     End Select
30720   End If

EXITP:
30730   KeyCode = intRetVal
30740   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgPrintEmail_optPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintEmail_optPrint_MouseMove"

30810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
30820   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub opgPrintEmail_optNone_KeyDown(KeyCode As Integer, Shift As Integer)

30900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintEmail_optNone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
30920   intShiftDown = (Shift And acShiftMask) > 0
30930   intAltDown = (Shift And acAltMask) > 0
30940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30970       With Me
30980         intRetVal = 0
30990         If .cmdReset.Enabled = True Then
31000           .cmdReset.SetFocus
31010         ElseIf .chkShowAcctNum.Enabled = True Then
31020           .chkShowAcctNum.SetFocus
31030         ElseIf .chkShowShortName.Enabled = True Then
31040           .chkShowShortName.SetFocus
31050         Else
31060           Select Case .opgView
                Case .opgView_optWide.OptionValue
31070             .chkEnableCountry_Wide.SetFocus
31080           Case .opgView_optCompact.OptionValue
31090             .chkEnableCountry_Compact.SetFocus
31100           End Select
31110         End If
31120       End With
31130     End Select
31140   End If

        ' ** Shift keys.
31150   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31170       With Me
31180         intRetVal = 0
31190         .opgPrintFax.SetFocus
31200       End With
31210     End Select
31220   End If

EXITP:
31230   KeyCode = intRetVal
31240   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub opgPrintEmail_optNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

31300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrintEmail_optNone_MouseMove"

31310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
31320   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub cmdReset_GotFocus()

31400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_GotFocus"

31410   With Me
31420     blnReset_Focus = True
31430     .cmdReset_raised_semifocus_dots_img.Visible = True
31440     .cmdReset_raised_img.Visible = False
31450     .cmdReset_raised_focus_img.Visible = False
31460     .cmdReset_raised_focus_dots_img.Visible = False
31470     .cmdReset_sunken_focus_dots_img.Visible = False
31480     .cmdReset_raised_img_dis.Visible = False
31490   End With

EXITP:
31500   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub cmdReset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

31600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseDown"

31610   With Me
31620     blnReset_MouseDown = True
31630     .cmdReset_sunken_focus_dots_img.Visible = True
31640     .cmdReset_raised_img.Visible = False
31650     .cmdReset_raised_semifocus_dots_img.Visible = False
31660     .cmdReset_raised_focus_img.Visible = False
31670     .cmdReset_raised_focus_dots_img.Visible = False
31680     .cmdReset_raised_img_dis.Visible = False
31690   End With

EXITP:
31700   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub cmdReset_Click()

31800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_Click"

31810   With Me

31820     .opgFormatName = .opgFormatName.DefaultValue
31830     opgFormatName_AfterUpdate  ' ** Procedure: Above.
31840     DoEvents
31850     .opgFormatZip = .opgFormatZip.DefaultValue
31860     opgFormatZip_AfterUpdate  ' ** Procedure: Above.
31870     DoEvents
31880     .opgFormatPhone = .opgFormatPhone.DefaultValue
31890     opgFormatPhone_AfterUpdate  ' ** Procedure: Above.
31900     DoEvents
31910     .opgPrintAddress = .opgPrintAddress.DefaultValue
31920     opgPrintAddress_AfterUpdate  ' ** Procedure: Above.
31930     DoEvents
31940     .opgPrintCSZ = .opgPrintCSZ.DefaultValue
31950     opgPrintCSZ_AfterUpdate  ' ** Procedure: Above.
31960     DoEvents
31970     .opgPrintCSZCP = .opgPrintCSZCP.DefaultValue
31980     opgPrintCSZCP_AfterUpdate  ' ** Procedure: Above.
31990     DoEvents
32000     .opgPrintPhone = .opgPrintPhone.DefaultValue
32010     opgPrintPhone_AfterUpdate  ' ** Procedure: Above.
32020     DoEvents
32030     .opgPrintFax = .opgPrintFax.DefaultValue
32040     opgPrintFax_AfterUpdate  ' ** Procedure: Above.
32050     DoEvents
32060     .opgPrintEmail = .opgPrintEmail.DefaultValue
32070     opgPrintEmail_AfterUpdate  ' ** Procedure: Above.
32080     DoEvents

32090     Select Case .opgView
          Case .opgView_optWide.OptionValue
32100       .chkShowAcctNum = .chkShowAcctNum.DefaultValue
32110       chkShowAcctNum_AfterUpdate  ' ** Procedure: Below.
32120       DoEvents
32130       .chkShowShortName = .chkShowShortName.DefaultValue
32140       chkShowShortName_AfterUpdate  ' ** Procedure: Below.
32150       DoEvents
32160     Case .opgView_optCompact.OptionValue
32170       .chkShowAcctNum = True
32180       chkShowAcctNum_AfterUpdate  ' ** Procedure: Below.
32190       DoEvents
32200       .chkShowShortName = True
32210       chkShowShortName_AfterUpdate  ' ** Procedure: Below.
32220       DoEvents
32230     End Select
32240     .chkEnableCountry_Wide = .chkEnableCountry_Wide.DefaultValue
32250     chkEnableCountry_Wide_AfterUpdate  ' ** Procedure: Below.
32260     DoEvents
32270     .chkEnableCountry_Compact = .chkEnableCountry_Compact.DefaultValue
32280     chkEnableCountry_Compact_AfterUpdate  ' ** Procedure: Below.
32290     DoEvents
32300     .chkPageOf = .chkPageOf.DefaultValue
32310     chkPageOf_AfterUpdate  ' ** Procedure: Below.
32320     DoEvents

32330   End With

EXITP:
32340   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub cmdReset_KeyDown(KeyCode As Integer, Shift As Integer)

32400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

32410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
32420   intShiftDown = (Shift And acShiftMask) > 0
32430   intAltDown = (Shift And acAltMask) > 0
32440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
32450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
32460     Select Case intRetVal
          Case vbKeyTab
32470       With Me
32480         intRetVal = 0
32490         If .chkShowAcctNum.Enabled = True Then
32500           .chkShowAcctNum.SetFocus
32510         ElseIf .chkShowShortName.Enabled = True Then
32520           .chkShowShortName.SetFocus
32530         Else
32540           Select Case .opgView
                Case .opgView_optWide.OptionValue
32550             .chkEnableCountry_Wide.SetFocus
32560           Case .opgView_optCompact.OptionValue
32570             .chkEnableCountry_Compact.SetFocus
32580           End Select
32590         End If
32600       End With
32610     End Select
32620   End If

        ' ** Shift keys.
32630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
32640     Select Case intRetVal
          Case vbKeyTab
32650       With Me
32660         intRetVal = 0
32670         .opgPrintEmail.SetFocus
32680       End With
32690     End Select
32700   End If

EXITP:
32710   KeyCode = intRetVal
32720   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub cmdReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

32800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseMove"

32810   With Me
32820     If blnReset_MouseDown = False Then
32830       Select Case blnReset_Focus
            Case True
32840         .cmdReset_raised_focus_dots_img.Visible = True
32850         .cmdReset_raised_focus_img.Visible = False
32860       Case False
32870         .cmdReset_raised_focus_img.Visible = True
32880         .cmdReset_raised_focus_dots_img.Visible = False
32890       End Select
32900       .cmdReset_raised_img.Visible = False
32910       .cmdReset_raised_semifocus_dots_img.Visible = False
32920       .cmdReset_sunken_focus_dots_img.Visible = False
32930       .cmdReset_raised_img_dis.Visible = False
32940     End If
32950   End With

EXITP:
32960   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub cmdReset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

33000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseUp"

33010   With Me
33020     .cmdReset_raised_focus_dots_img.Visible = True
33030     .cmdReset_raised_img.Visible = False
33040     .cmdReset_raised_semifocus_dots_img.Visible = False
33050     .cmdReset_raised_focus_img.Visible = False
33060     .cmdReset_sunken_focus_dots_img.Visible = False
33070     .cmdReset_raised_img_dis.Visible = False
33080     blnReset_MouseDown = False
33090   End With

EXITP:
33100   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub cmdReset_LostFocus()

33200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_LostFocus"

33210   With Me
33220     .cmdReset_raised_img.Visible = True
33230     .cmdReset_raised_semifocus_dots_img.Visible = False
33240     .cmdReset_raised_focus_img.Visible = False
33250     .cmdReset_raised_focus_dots_img.Visible = False
33260     .cmdReset_sunken_focus_dots_img.Visible = False
33270     .cmdReset_raised_img_dis.Visible = False
33280     blnReset_Focus = False
33290   End With

EXITP:
33300   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Public Sub chkShowAcctNum_AfterUpdate()

33400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAcctNum_AfterUpdate"

33410   With Me
33420     If .opgView = .opgView_optWide.OptionValue Then
            ' ** Both always show in Compact mode.
33430       Select Case .chkShowAcctNum
            Case True
33440         .chkShowAcctNum_lbl.FontBold = True
33450         AcctNoShort_Set THIS_PROC, True, lngFrm_Top_Orig, lngFrm_Left_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig, lngForm_Width, Me  ' ** Module Procedure: modAccountContacts.
33460       Case False
33470         .chkShowAcctNum_lbl.FontBold = False
33480         AcctNoShort_Set THIS_PROC, False, lngFrm_Top_Orig, lngFrm_Left_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig, lngForm_Width, Me  ' ** Module Procedure: modAccountContacts.
33490       End Select
33500       SetReset IsntDefault  ' ** Procedure/Function: Below.
33510       .chkShowAcctNum_Last = .chkShowAcctNum
33520       .frmAccountContacts_Sub.SetFocus
33530     End If
33540   End With

EXITP:
33550   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub chkShowAcctNum_KeyDown(KeyCode As Integer, Shift As Integer)

33600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAcctNum_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

33610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
33620   intShiftDown = (Shift And acShiftMask) > 0
33630   intAltDown = (Shift And acAltMask) > 0
33640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
33650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
33660     Select Case intRetVal
          Case vbKeyTab
33670       With Me
33680         intRetVal = 0
33690         .chkShowShortName.SetFocus
33700       End With
33710     End Select
33720   End If

        ' ** Shift keys.
33730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
33740     Select Case intRetVal
          Case vbKeyTab
33750       With Me
33760         intRetVal = 0
33770         If .cmdReset.Enabled = True Then
33780           .cmdReset.SetFocus
33790         ElseIf .opgPrintEmail.Enabled = True Then
33800           .opgPrintEmail.SetFocus
33810         ElseIf .opgFormatPhone.Enabled = True Then
33820           .opgFormatPhone.SetFocus
33830         Else
33840           .opgView.SetFocus
33850         End If
33860       End With
33870     End Select
33880   End If

EXITP:
33890   KeyCode = intRetVal
33900   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Public Sub chkShowShortName_AfterUpdate()

34000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowShortName_AfterUpdate"

34010   With Me
34020     If .opgView = .opgView_optWide.OptionValue Then
            ' ** Both always show in Compact mode.
34030       Select Case .chkShowShortName
            Case True
34040         .chkShowShortName_lbl.FontBold = True
34050         AcctNoShort_Set THIS_PROC, True, lngFrm_Top_Orig, lngFrm_Left_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig, lngForm_Width, Me  ' ** Module Procedure: modAccountContacts.
34060       Case False
34070         .chkShowShortName_lbl.FontBold = False
34080         AcctNoShort_Set THIS_PROC, False, lngFrm_Top_Orig, lngFrm_Left_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig, lngForm_Width, Me  ' ** Module Procedure: modAccountContacts.
34090       End Select
34100       SetReset IsntDefault  ' ** Procedure/Function: Below.
34110       .chkShowShortName_Last = .chkShowShortName
34120       .frmAccountContacts_Sub.SetFocus
34130     End If
34140   End With

EXITP:
34150   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub chkShowShortName_KeyDown(KeyCode As Integer, Shift As Integer)

34200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowShortName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

34210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
34220   intShiftDown = (Shift And acShiftMask) > 0
34230   intAltDown = (Shift And acAltMask) > 0
34240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
34250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
34260     Select Case intRetVal
          Case vbKeyTab
34270       With Me
34280         intRetVal = 0
34290         Select Case .opgView
              Case .opgView_optWide.OptionValue
34300           .chkEnableCountry_Wide.SetFocus
34310         Case .opgView_optCompact.OptionValue
34320           .chkEnableCountry_Compact.SetFocus
34330         End Select
34340       End With
34350     End Select
34360   End If

        ' ** Shift keys.
34370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
34380     Select Case intRetVal
          Case vbKeyTab
34390       With Me
34400         intRetVal = 0
34410         .chkShowAcctNum.SetFocus
34420       End With
34430     End Select
34440   End If

EXITP:
34450   KeyCode = intRetVal
34460   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Public Sub chkEnableCountry_Wide_AfterUpdate()

34500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkEnableCountry_Wide_AfterUpdate"

        Dim frm As Access.Form
        Dim lngTmp01 As Long, lngTmp02 As Long

34510   With Me
34520     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
34530       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
34540     End If
34550     Select Case .opgView
          Case .opgView_optWide.OptionValue
34560       Set frm = .frmAccountContacts_Sub.Form
34570       Select Case .chkEnableCountry_Wide
            Case True
34580         .chkEnableCountry_Wide_lbl.FontBold = True
34590         frm.EnableCountry True  ' ** Form Procedure: frmAccountContacts_Sub.
34600         EnableCountry_SetFrmWidth_AC lngFrm_Top_Orig, lngFrm_Left_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig, lngForm_Width, Me  ' ** Module Procedure: modAccountContacts.
34610         lngTmp01 = ((.opgPrintCSZCP_box.Left + .opgPrintCSZCP_box.Width) + arr_varCtl(C_L2_LFT, 14))  ' ** New right-hand edge.
34620         lngTmp01 = (lngTmp01 - .opgPrint_box.Left)  ' ** New width
34630         lngTmp02 = (lngTmp01 - arr_varCtl(C_WDT, 14))  ' ** Difference.
34640         .cmdReset.Left = (arr_varCtl(C_LFT, 6) + lngTmp02)
34650         .cmdReset_raised_img.Left = .cmdReset.Left
34660         .cmdReset_raised_semifocus_dots_img.Left = .cmdReset.Left
34670         .cmdReset_raised_focus_img.Left = .cmdReset.Left
34680         .cmdReset_raised_focus_dots_img.Left = .cmdReset.Left
34690         .cmdReset_sunken_focus_dots_img.Left = .cmdReset.Left
34700         .cmdReset_raised_img_dis.Left = .cmdReset.Left
34710         .opgPrint_box.Width = lngTmp01
34720         .opgPrint_hline01.Width = (.opgPrint_box.Width - (2& * lngTpp)) '(.opgPrint_hline01.Width + lngTmp02)
34730         .opgPrint_hline02.Width = .opgPrint_hline01.Width '(.opgPrint_hline02.Width + lngTmp02)
34740         .opgPrintCSZ.Visible = False
34750         .opgPrintCSZ_box.Visible = False
34760         .opgPrintCSZCP.Visible = True
34770         .opgPrintCSZCP_box.Visible = True
34780       Case False
34790         .chkEnableCountry_Wide_lbl.FontBold = False
34800         frm.EnableCountry False  ' ** Form Procedure: frmAccountContacts_Sub.
34810         EnableCountry_SetFrmWidth_AC lngFrm_Top_Orig, lngFrm_Left_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig, lngForm_Width, Me  ' ** Module Procedure: modAccountContacts.
34820         .opgPrintCSZCP.Visible = False
34830         .opgPrintCSZCP_box.Visible = False
34840         .opgPrintCSZ.Visible = True
34850         .opgPrintCSZ_box.Visible = True
34860         lngTmp01 = (.opgPrint_box.Width - arr_varCtl(C_WDT, 14))  ' ** Difference.
34870         .opgPrint_box.Width = arr_varCtl(C_WDT, 14)
34880         .opgPrint_hline01.Width = (.opgPrint_box.Width - (2& * lngTpp)) '(.opgPrint_hline01.Width - lngTmp01)
34890         .opgPrint_hline02.Width = .opgPrint_hline01.Width '(.opgPrint_hline02.Width - lngTmp01)
34900         .cmdReset.Left = arr_varCtl(C_LFT, 6)
34910         .cmdReset_raised_img.Left = .cmdReset.Left
34920         .cmdReset_raised_semifocus_dots_img.Left = .cmdReset.Left
34930         .cmdReset_raised_focus_img.Left = .cmdReset.Left
34940         .cmdReset_raised_focus_dots_img.Left = .cmdReset.Left
34950         .cmdReset_sunken_focus_dots_img.Left = .cmdReset.Left
34960         .cmdReset_raised_img_dis.Left = .cmdReset.Left
34970       End Select
34980       .chkEnableCountry_Wide_Last = .chkEnableCountry_Wide
34990       SetReset IsntDefault  ' ** Procedure/Function: Below.
35000       lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountContacts_Sub.
35010       If lngRecsCur > 0 Then
35020         .frmAccountContacts_Sub.SetFocus
35030       End If
35040     Case .opgView_optCompact.OptionValue
            ' ** Shouldn't be here.
35050     End Select
35060   End With

EXITP:
35070   Set frm = Nothing
35080   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub chkEnableCountry_Wide_KeyDown(KeyCode As Integer, Shift As Integer)

35100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkEnableCountry_Wide_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

35110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
35120   intShiftDown = (Shift And acShiftMask) > 0
35130   intAltDown = (Shift And acAltMask) > 0
35140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
35150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
35160     Select Case intRetVal
          Case vbKeyTab
35170       With Me
35180         intRetVal = 0
35190         If .chkShowShortName.Enabled = True Then
35200           .chkShowShortName.SetFocus
35210         ElseIf .chkShowAcctNum.Enabled = True Then
35220           .chkShowAcctNum.SetFocus
35230         ElseIf .cmdReset.Enabled = True Then
35240           .cmdReset.SetFocus
35250         ElseIf .opgPrintEmail.Enabled = True Then
35260           .opgPrintEmail.SetFocus
35270         ElseIf .opgFormatPhone.Enabled = True Then
35280           .opgFormatPhone.SetFocus
35290         Else
35300           .opgView.SetFocus
35310         End If
35320       End With
35330     End Select
35340   End If

        ' ** Shift keys.
35350   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
35360     Select Case intRetVal
          Case vbKeyTab
35370       With Me
35380         intRetVal = 0
35390         .chkPageOf.SetFocus
35400       End With
35410     End Select
35420   End If

EXITP:
35430   KeyCode = intRetVal
35440   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub chkEnableCountry_Compact_AfterUpdate()

35500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkEnableCountry_Compact_AfterUpdate"

        Dim frm As Access.Form
        Dim lngTmp01 As Long, lngTmp02 As Long

35510   With Me
35520     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
35530       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
35540     End If
35550     Select Case .opgView
          Case .opgView_optWide.OptionValue
            ' ** Shouldn't be here.
35560     Case .opgView_optCompact.OptionValue
35570       Set frm = .frmAccountContacts_Sub2.Form
35580       Select Case .chkEnableCountry_Compact
            Case True
35590         .chkEnableCountry_Compact_lbl.FontBold = True
35600         frm.EnableCountry True  ' ** Form Procedure: frmAccountContacts_Sub2.
35610         EnableCountry_SetFrmWidth_AC lngFrm_Top_Orig, lngFrm_Left_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig, lngForm_Width, Me  ' ** Module Procedure: modAccountContacts.
35620         lngTmp01 = (.opgPrintCSZCP_box.Left + .opgPrintCSZCP_box.Width) + arr_varCtl(C_L2_LFT, 14)  ' ** New right-hand edge.
35630         lngTmp01 = lngTmp01 - .opgPrint_box.Left  ' ** New width
35640         lngTmp02 = lngTmp01 - arr_varCtl(C_WDT, 14) ' ** Difference.
35650         .cmdReset.Left = (arr_varCtl(C_LFT, 6) + lngTmp02)
35660         .cmdReset_raised_img.Left = .cmdReset.Left
35670         .cmdReset_raised_semifocus_dots_img.Left = .cmdReset.Left
35680         .cmdReset_raised_focus_img.Left = .cmdReset.Left
35690         .cmdReset_raised_focus_dots_img.Left = .cmdReset.Left
35700         .cmdReset_sunken_focus_dots_img.Left = .cmdReset.Left
35710         .cmdReset_raised_img_dis.Left = .cmdReset.Left
35720         .opgPrint_box.Width = lngTmp01
              ' ** I can't figure out why this messes up when done like all the others!
35730         .opgPrint_hline01.Width = (lngTmp01 - (2& * lngTpp))
35740         .opgPrint_hline02.Width = .opgPrint_hline01.Width
35750         .opgPrintCSZ.Visible = False
35760         .opgPrintCSZ_box.Visible = False
35770         .opgPrintCSZCP.Visible = True
35780         .opgPrintCSZCP_box.Visible = True
35790       Case False
35800         .chkEnableCountry_Compact_lbl.FontBold = False
35810         frm.EnableCountry False  ' ** Form Procedure: frmAccountContacts_Sub2.
35820         EnableCountry_SetFrmWidth_AC lngFrm_Top_Orig, lngFrm_Left_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig, lngForm_Width, Me  ' ** Module Procedure: modAccountContacts.
35830         .opgPrintCSZCP.Visible = False
35840         .opgPrintCSZCP_box.Visible = False
35850         .opgPrintCSZ.Visible = True
35860         .opgPrintCSZ_box.Visible = True
35870         lngTmp01 = (.opgPrint_box.Width - arr_varCtl(C_WDT, 14))  ' ** Difference.
35880         .opgPrint_hline01.Width = (.opgPrint_hline01.Width - lngTmp01)
35890         .opgPrint_hline02.Width = (.opgPrint_hline02.Width - lngTmp01)
35900         .opgPrint_box.Width = arr_varCtl(C_WDT, 14)
35910         .cmdReset.Left = arr_varCtl(C_LFT, 6)
35920         .cmdReset_raised_img.Left = .cmdReset.Left
35930         .cmdReset_raised_semifocus_dots_img.Left = .cmdReset.Left
35940         .cmdReset_raised_focus_img.Left = .cmdReset.Left
35950         .cmdReset_raised_focus_dots_img.Left = .cmdReset.Left
35960         .cmdReset_sunken_focus_dots_img.Left = .cmdReset.Left
35970         .cmdReset_raised_img_dis.Left = .cmdReset.Left
35980       End Select
35990       .chkEnableCountry_Compact_Last = .chkEnableCountry_Compact
36000       SetReset IsntDefault  ' ** Procedure/Function: Below.
36010       lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountContacts_Sub2.
36020       If lngRecsCur > 0 Then
36030         .frmAccountContacts_Sub2.SetFocus
36040       End If
36050     End Select
36060   End With

EXITP:
36070   Set frm = Nothing
36080   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub chkEnableCountry_Compact_KeyDown(KeyCode As Integer, Shift As Integer)

36100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkEnableCountry_Compact_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

36110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
36120   intShiftDown = (Shift And acShiftMask) > 0
36130   intAltDown = (Shift And acAltMask) > 0
36140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
36150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
36160     Select Case intRetVal
          Case vbKeyTab
36170       With Me
36180         intRetVal = 0
36190         If .chkShowShortName.Enabled = True Then
36200           .chkShowShortName.SetFocus
36210         ElseIf .chkShowAcctNum.Enabled = True Then
36220           .chkShowAcctNum.SetFocus
36230         ElseIf .cmdReset.Enabled = True Then
36240           .cmdReset.SetFocus
36250         ElseIf .opgPrintEmail.Enabled = True Then
36260           .opgPrintEmail.SetFocus
36270         ElseIf .opgFormatPhone.Enabled = True Then
36280           .opgFormatPhone.SetFocus
36290         Else
36300           .opgView.SetFocus
36310         End If
36320       End With
36330     End Select
36340   End If

        ' ** Shift keys.
36350   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
36360     Select Case intRetVal
          Case vbKeyTab
36370       With Me
36380         intRetVal = 0
36390         .chkPageOf.SetFocus
36400       End With
36410     End Select
36420   End If

EXITP:
36430   KeyCode = intRetVal
36440   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Public Sub chkPageOf_AfterUpdate()

36500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

36510   With Me
36520     Select Case .chkPageOf
          Case True
36530       .chkPageOf_lbl.FontBold = True
36540     Case False
36550       .chkPageOf_lbl.FontBold = False
36560     End Select
36570     SetReset IsntDefault  ' ** Procedure/Function: Below.
36580   End With

EXITP:
36590   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

36600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

36610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
36620   intShiftDown = (Shift And acShiftMask) > 0
36630   intAltDown = (Shift And acAltMask) > 0
36640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
36650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
36660     Select Case intRetVal
          Case vbKeyTab
36670       With Me
36680         intRetVal = 0
36690         If .cmdPreview.Enabled = True Then
36700           .cmdPreview.SetFocus
36710         ElseIf .cmdAdd.Enabled = True Then
36720           .cmdAdd.SetFocus
36730         ElseIf .cmdDelete.Enabled = True Then
36740           .cmdDelete.SetFocus
36750         ElseIf .cmdUpdate.Enabled = True Then
36760           .cmdUpdate.SetFocus
36770         Else
36780           .chkAlwaysUpdate.SetFocus
36790         End If
36800       End With
36810     End Select
36820   End If

        ' ** Shift keys.
36830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
36840     Select Case intRetVal
          Case vbKeyTab
36850       With Me
36860         intRetVal = 0
36870         Select Case .opgView
              Case .opgView_optWide.OptionValue
36880           .chkEnableCountry_Wide.SetFocus
36890         Case .opgView_optCompact.OptionValue
36900           .chkEnableCountry_Compact.SetFocus
36910         End Select
36920       End With
36930     End Select
36940   End If

EXITP:
36950   KeyCode = intRetVal
36960   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub chkAlwaysUpdate_AfterUpdate()

37000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAlwaysUpdate_AfterUpdate"

        Dim varTmp00 As Variant

37010   With Me
37020     Select Case .chkAlwaysUpdate
          Case True
37030       .chkAlwaysUpdate_lbl.FontBold = True
37040       .chkAlwaysUpdate_lbl2.FontBold = True
37050       .cmdUpdate.Enabled = False
37060     Case False
37070       .chkAlwaysUpdate_lbl.FontBold = False
37080       .chkAlwaysUpdate_lbl2.FontBold = False
            ' ** tblAccount_Contact_Staging, linked to Account.
37090       varTmp00 = DCount("*", "qryAccount_Contacts_01_01")
37100       If IsNull(varTmp00) = True Then
37110         lngRecsCur = 0&
37120       Else
37130         lngRecsCur = varTmp00
37140       End If
37150       If lngRecsCur > 0& Then
37160         .cmdUpdate.Enabled = True
37170       End If
37180     End Select
37190   End With

EXITP:
37200   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub chkAlwaysUpdate_KeyDown(KeyCode As Integer, Shift As Integer)

37300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAlwaysUpdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

37310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
37320   intShiftDown = (Shift And acShiftMask) > 0
37330   intAltDown = (Shift And acAltMask) > 0
37340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
37350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
37360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
37370       With Me
37380         intRetVal = 0
37390         .cmdClose.SetFocus
37400       End With
37410     End Select
37420   End If

        ' ** Shift keys.
37430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
37440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
37450       With Me
37460         intRetVal = 0
37470         If .cmdUpdate.Enabled = True Then
37480           .cmdUpdate.SetFocus
37490         ElseIf .cmdDelete.Enabled = True Then
37500           .cmdDelete.SetFocus
37510         ElseIf .cmdAdd.Enabled = True Then
37520           .cmdAdd.SetFocus
37530         ElseIf .cmdExcel.Enabled = True Then
37540           .cmdExcel.SetFocus
37550         ElseIf .cmdWord.Enabled = True Then
37560           .cmdWord.SetFocus
37570         Else
37580           .chkPageOf.SetFocus
37590         End If
37600       End With
37610     End Select
37620   End If

EXITP:
37630   KeyCode = intRetVal
37640   Exit Sub

ERRH:
800     intRetVal = 0
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub chkAlwaysUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

37700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAlwaysUpdate_MouseMove"

37710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
37720   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Public Sub AcctNoShort_Move(lngNewForm_Width As Long)

37800 On Error GoTo ERRH

        Const THIS_PROC As String = "AcctNoShort_Move"

        Dim lngOffset As Long
        Dim lngTmp01 As Long

37810   With Me

37820     .Width = lngNewForm_Width

37830     lngTmp01 = (lngNewForm_Width - lngForm_Width)
37840     lngOffset = (.chkEnableCountry_alt_line.Left - arr_varCtl(C_LFT, 8))

37850     .frmAccountContacts_Sub.Width = (arr_varCtl(C_F1_WDT, 0) + lngTmp01)
37860     .frmAccountContacts_Sub_box.Width = (.frmAccountContacts_Sub.Width + (2& * lngTpp))
37870     .Nav_Sub1_box01.Width = .frmAccountContacts_Sub.Width

37880     .cmdClose.Left = (arr_varCtl(C_LFT, 2) + lngTmp01)
37890     .chkAlwaysUpdate_lbl.Left = (arr_varCtl(C_L1_LFT, 10) + lngTmp01)
37900     .chkAlwaysUpdate_lbl2.Left = .chkAlwaysUpdate_lbl.Left
37910     .chkAlwaysUpdate_lbl2_dim_hi.Left = (.chkAlwaysUpdate_lbl2.Left + lngTpp)
37920     .chkAlwaysUpdate.Left = (arr_varCtl(C_LFT, 10) + lngTmp01)
37930     .cmdUpdate.Left = (arr_varCtl(C_LFT, 5) + lngTmp01)
37940     .cmdDelete.Left = (arr_varCtl(C_LFT, 4) + lngTmp01)
37950     .cmdAdd.Left = (arr_varCtl(C_LFT, 3) + lngTmp01)
37960     .Header_vline01.Left = lngNewForm_Width
37970     .Header_vline02.Left = lngNewForm_Width
37980     .Detail_vline01.Left = lngNewForm_Width
37990     .Detail_vline02.Left = lngNewForm_Width
38000     .Detail_vline03.Left = lngNewForm_Width
38010     .Detail_vline04.Left = lngNewForm_Width
38020     .Footer_vline01.Left = lngNewForm_Width
38030     .Footer_vline02.Left = lngNewForm_Width
38040     .Header_hline01.Width = lngNewForm_Width
38050     .Header_hline02.Width = lngNewForm_Width
38060     .Detail_hline01.Width = lngNewForm_Width
38070     .Detail_hline02.Width = lngNewForm_Width
38080     .Detail_hline03.Width = lngNewForm_Width
38090     .Detail_hline04.Width = lngNewForm_Width
38100     .Footer_hline01.Width = lngNewForm_Width
38110     .Footer_hline02.Width = lngNewForm_Width

38120     If lngNewForm_Width = lngForm_Width Then
            ' ** Nothing.
38130     Else
            ' ** Nothing.
38140     End If

38150   End With

EXITP:
38160   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Function IsntDefault() As Boolean

38200 On Error GoTo ERRH

        Const THIS_PROC As String = "IsntDefault"

        Dim blnRetVal As Boolean

38210   With Me
38220     blnRetVal = False
38230     If .opgFormatName <> .opgFormatName.DefaultValue Then
38240       blnRetVal = True
38250     ElseIf .opgFormatZip <> .opgFormatZip.DefaultValue Then
38260       blnRetVal = True
38270     ElseIf .opgFormatPhone <> .opgFormatPhone.DefaultValue Then
38280       blnRetVal = True
38290     ElseIf .opgPrintAddress <> .opgPrintAddress.DefaultValue Then
38300       blnRetVal = True
38310     ElseIf .opgPrintCSZ <> .opgPrintCSZ.DefaultValue Then
38320       blnRetVal = True
38330     ElseIf .opgPrintCSZCP <> .opgPrintCSZCP.DefaultValue Then
38340       blnRetVal = True
38350     ElseIf .opgPrintPhone <> .opgPrintPhone.DefaultValue Then
38360       blnRetVal = True
38370     ElseIf .opgPrintFax <> .opgPrintFax.DefaultValue Then
38380       blnRetVal = True
38390     ElseIf .opgPrintEmail <> .opgPrintEmail.DefaultValue Then
38400       blnRetVal = True
38410     ElseIf .chkPageOf <> CBool(.chkPageOf.DefaultValue) Then
38420       blnRetVal = True
38430     Else
38440       Select Case .opgView
            Case .opgView_optWide.OptionValue
38450         If .chkShowAcctNum <> CBool(.chkShowAcctNum.DefaultValue) Then
38460           blnRetVal = True
38470         ElseIf .chkShowShortName <> CBool(.chkShowShortName.DefaultValue) Then
38480           blnRetVal = True
38490         ElseIf .chkEnableCountry_Wide <> CBool(.chkEnableCountry_Wide.DefaultValue) Then
38500           blnRetVal = True
38510         End If
38520       Case .opgView_optCompact.OptionValue
38530         If .chkEnableCountry_Compact <> CBool(.chkEnableCountry_Compact.DefaultValue) Then
38540           blnRetVal = True
38550         End If
38560       End Select
38570     End If
38580   End With

EXITP:
38590   IsntDefault = blnRetVal
38600   Exit Function

ERRH:
800     blnRetVal = False
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Function

Private Sub SetReset(blnEnable As Boolean)

38700 On Error GoTo ERRH

        Const THIS_PROC As String = "SetReset"

38710   With Me
38720     If .cmdDelete.Enabled = False And .cmdUpdate.Enabled = False Then
            ' ** Opening with no records
38730       blnEnable = False
38740     End If
38750     Select Case blnEnable
          Case True
38760       .cmdReset.Enabled = True
38770       .cmdReset_raised_img.Visible = True
38780       .cmdReset_raised_semifocus_dots_img.Visible = False
38790       .cmdReset_raised_focus_img.Visible = False
38800       .cmdReset_raised_focus_dots_img.Visible = False
38810       .cmdReset_sunken_focus_dots_img.Visible = False
38820       .cmdReset_raised_img_dis.Visible = False
38830     Case False
38840       .FocusHolder.SetFocus
38850       .cmdReset.Enabled = False
38860       .cmdReset_raised_img_dis.Visible = True
38870       .cmdReset_raised_img.Visible = False
38880       .cmdReset_raised_semifocus_dots_img.Visible = False
38890       .cmdReset_raised_focus_img.Visible = False
38900       .cmdReset_raised_focus_dots_img.Visible = False
38910       .cmdReset_sunken_focus_dots_img.Visible = False
38920     End Select
38930   End With

EXITP:
38940   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Public Sub CtlArraySet(arr_varTmp00 As Variant)

39000 On Error GoTo ERRH

        Const THIS_PROC As String = "CtlArraySet"

39010   arr_varCtl = arr_varTmp00
39020   lngCtls = UBound(arr_varCtl, 2) + 1&

EXITP:
39030   Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub
