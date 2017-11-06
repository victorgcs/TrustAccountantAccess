VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Title"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Title"

'VGC 09/22/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Main Menu:        M {cmdMenu}
' **   Quit:             Q {cmdQuit}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Background: B {CloseBack}
' **   Close Form:       X {cmdCloseForm}

' ** Shortcut Ctrl-Alt-Shift keys responsive from this form:
' **   References:       R {ShowRefs}

' ***************************************************************************************************
' ***************************************************************************************************
' **
' ** Secret Keys (Easter Eggs):
' **   frmAccountExport:
' **     On folder header (opgExport_lbl_cmd): Ctrl+Alt+Right-Click, ShowDataSwitch().
' **     ' ** Reveals/Hides extended export options.
' **   frmAccountHideTrans_Hidden:
' **     On 'Hidden' column label: Ctrl+Alt+Right-Click, regenerate LedgerHidden table.
' **     ' ** Empties the LedgerHidden and tblLedgerHidden tables, so that they can be regenerated.
' **   frmAssetPricing_Import:
' **     On Export Type list box: Ctrl+Alt+Right-Click, priceimport_id_lbx_MouseUp.
' **     ' ** Null-out Last File Name.
' **   frmLicense:
' **     On Key button (cmdKey): Ctrl+Alt+Right-Click, open frmLicense_Edit.
' **     ' ** Opens frmLicense_Edit.
' **   frmMenu_Account:
' **     On cmdHideTransactions: Ctrl+Enter, regenerate LedgerHidden table.
' **     ' ** Empties the LedgerHidden and tblLedgerHidden tables, so that they can be regenerated.
' **   frmMenu_Utility:
' **     On cmdArchiveTransactions: Ctrl+Enter, Ctrl+Click, open frmRpt_ArchivedTransactions.
' **     ' ** Empties the LedgerHidden table, so that they can be regenerated.
' **   frmOptions:
' **     On Header ('Options'): DblClick, open frmOptions_Force.
' **     ' ** Forces Account Type prefix.
' **   frmStatementParameters:
' **     On cmdTransactionsPreview: Ctrl+Alt+Right-Click, cmdTransactionsPreview_Click.
' **     ' ** This was added to assist when users do things out of sequence. VGC: May now be moot.
' **
' ***************************************************************************************************
' ***************************************************************************************************

' ** Array: arr_varRef().
Private lngRefs As Long, arr_varRef() As Variant
Private Const R_ELEMS As Integer = 7  ' ** Array's first-element UBound().
Private Const R_RNAM    As Integer = 0
Private Const R_DSC     As Integer = 1
Private Const R_ISBROKE As Integer = 2
Private Const R_BUILTIN As Integer = 3
Private Const R_GUID    As Integer = 4
Private Const R_PATH    As Integer = 5
Private Const R_PATHERR As Integer = 6
Private Const R_OTHRERR As Integer = 7

Private intOpenVer As Integer, blnTimer1 As Boolean, blnTimer2 As Boolean, blnTimer3 As Boolean
Private strCallingForm As String, lngEvalLbl_Top As Long, lngTpp As Long
Private blnLogoAnim_Burst1 As Boolean, blnLogoAnim_Burst2 As Boolean, blnLogoAnim_Key As Boolean
Private blnInitMsg As Boolean, blnReturnFromMain As Boolean
Private intSizeChks As Integer, intConversionCheck As Integer
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim clsDevs As clsDevices, ref As Access.Reference
        Dim strMessage As String, strTestMessage As String, strFormName As String
        Dim strLicense As String, strVersion As String
        Dim intLenMsg As Integer
        Dim dblVersion As Double
        Dim blnSkip As Boolean, blnFound As Boolean, blnRetVal As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, strTmp01 As String, intTmp02 As Integer, lngTmp03 As Long, blnTmp04 As Boolean
        Dim intX As Integer, lngE As Long

110     With Me
120       If CurrentUser = "taimport" Then
130         Cancel = -1
140         Beep
150         MsgBox "You do not have authorization to use Trust Accountant." & vbCrLf & _
              "Please use the Trust Accountant administrative username and password.", _
              vbCritical + vbOKOnly, "Invalid Username"
160       Else

170         .PictureData = .Form_Menu_Title_img

180         blnInitMsg = False: blnReturnFromMain = False: blnRetVal = True
190         blnTimer1 = False: blnTimer2 = False: blnTimer3 = False

200         If gstrReturningForm = "frmMenu_Main" Then
210           blnSkip = True
220           If blnSkip = False Then
                ' ** For GTR_REF only.
230             DoCmd.OpenForm "frmMenu_Title"
240           End If
250           .InitMsg_lbl2.Visible = False
260           .InitMsg_lbl3.Visible = False
270           blnReturnFromMain = True
280         End If

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
290         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

300         lngEvalLbl_Top = .Eval_lbl.Top

310         .DevKeys_lbl1a.Visible = False  ' ** Ctrl+Shift+A :
320         .DevKeys_lbl1b.Visible = False  ' ** Load Title If Lost
330         .DevKeys_lbl2a.Visible = False  ' ** Ctrl+Shift+B :
340         .DevKeys_lbl2b.Visible = False  ' ** Close Menu Background
350         .DevKeys_lbl3a.Visible = False  ' ** Ctrl+Shift+X :
360         .DevKeys_lbl3b.Visible = False  ' ** Close Menu
370         .DevKeys_lbl4a.Visible = False  ' ** Ctrl+Alt+Shift+R :
380         .DevKeys_lbl4b.Visible = False  ' ** Show References

390         If blnReturnFromMain = False Then
              ' ** STATUS 1: Initialization Start.  ' ** OF 16 STEPS!
400           .InitMsg_lbl2.Caption = "Initializing"
410           .InitMsg_lbl2.Visible = True
420           DoEvents
430         End If  ' ** blnReturnFromMain.

440         lngTmp03 = 0
450         strFormName = vbNullString
460   On Error Resume Next
470         Do Until strFormName = THIS_NAME
480           lngTmp03 = lngTmp03 + 1&
490           DoCmd.SelectObject acForm, THIS_NAME, False
500           strFormName = Screen.ActiveForm.Name
510           If lngTmp03 > 100& Then Exit Do
520         Loop
530   On Error GoTo ERRH
540         If IsMaximized(Me) = True Then
550           DoCmd.Restore
560         End If
570         DoEvents

580         blnLogoAnim_Burst2 = False: blnLogoAnim_Key = False
590         If gstrReturningForm = vbNullString Then
600           blnLogoAnim_Burst1 = True
610         End If

            ' ** Initializing message now has its own label!
620         Select Case blnReturnFromMain
            Case True
630           .InitMsg_lbl1.Visible = False
640         Case False
650           .InitMsg_lbl1.Visible = True
660         End Select
670         DoCmd.Hourglass True
680         DoEvents
690         gstrReturningForm = vbNullString

700         If blnReturnFromMain = False Then
              ' ** STATUS 2: Database Window / Navigation Pane
710           .InitMsg_lbl2.Caption = "Database Window / Navigation Pane"
720           DoEvents
730         End If  ' ** blnReturnFromMain.

740         If blnReturnFromMain = False Then
              ' ** I'd like to start with windows not maximized, but with this form a Popup, Restore doesn't work.
              ' ** During testing, if the Database Window is visible, I can move there, Restore, and move back to here,
              ' ** but that won't work, of course, in the Release version. I suppose, then, that it really doesn't
              ' ** matter, since the users won't see anything Maximized unless I make it so.
750   On Error Resume Next
760           blnTmp04 = CurrentDb.Properties("StartupShowDBWindow")
770           If ERR.Number <> 0 Then
780   On Error GoTo ERRH
790             AppProperty_Add "StartupShowDBWindow", dbBoolean, False
800             blnTmp04 = CurrentDb.Properties("StartupShowDBWindow")
810           Else
820   On Error GoTo ERRH
830           End If
840           If blnTmp04 = True Then
850             DoCmd.SelectObject acForm, THIS_NAME, True
860             DoCmd.Restore
870             DoCmd.SelectObject acForm, THIS_NAME, False
880             DoEvents
890           End If
900         End If  ' ** blnReturnFromMain.

910         If IsNull(.OpenArgs) = False Then
920           If Trim(.OpenArgs) <> vbNullString Then
930             strCallingForm = Trim(.OpenArgs)
940           End If
950         End If

960         If gblnDev_NoErrHandle = True Then
970           MsgBox "ERROR HANDLER IS OFF!", vbExclamation + vbOKOnly, "Error Handler Off"
980   On Error GoTo 0
990         End If

1000        If blnReturnFromMain = False Then
              ' ** Gets the Navigation Pane out of there immediately!
1010          If IsAccess2007 = True Then  ' ** Module Function: modXAccess_07_10_Funcs.
1020            intOpenVer = 12
1030          ElseIf IsAccess2010 = True Then  ' ** Module Function: modXAccess_07_10_Funcs.
1040            intOpenVer = 14
1050          Else
1060            intOpenVer = 0
1070          End If
1080          Select Case intOpenVer
              Case 12
1090            SetOption_Access2007 True, THIS_PROC  ' ** Module Function: modXAccess_07_10_Funcs.
1100            DoEvents
1110          Case 14
1120            SetOption_Access2010 True  ' ** Module Function: modXAccess_07_10_Funcs.
1130            DoEvents
1140          End Select
1150        End If  ' ** blnReturnFromMain.

1160        intConversionCheck = 1: glngInstance = -1&  ' ** Default to Unnecessary : Don't run conversion loop.

            ' ** Set Demo status.
1170        If Len(TA_SEC) > Len(TA_SEC2) Then
1180          .Eval_lbl.Visible = True
1190          If .InitMsg_lbl1.Visible = False Then
1200            .InitMsg_lbl1.Visible = True
1210            DoEvents
1220          End If
1230          .Eval_lbl.Top = lngEvalLbl_Top + (10& * lngTpp)
1240        Else
1250          .Eval_lbl.Visible = False
1260        End If

1270        glngJournalForm = 0&
1280        gblnCompact = False: gblnBadSec = False
1290        gblnGoToReport = False
1300        gblnSingleUser = True  ' ** Untill proven otherwise.
1310        glngUserCntLedger = -1&

1320        If blnReturnFromMain = False Then
              ' ** STATUS 3: Multiple Instance Check
1330          .InitMsg_lbl2.Caption = "Multiple Instance Check"
1340          DoEvents
1350        End If  ' ** blnReturnFromMain.

1360        lngRefs = 0&
1370        ReDim arr_varRef(R_ELEMS, 0)

1380        If IsTAOpen = True Then  ' ** Module Function: modFileUtilities.
1390          MsgBox "There can only be one Trust Accountant session open at a time.", vbInformation + vbOKOnly, "Trust Accountant Already Open"
1400          DoCmd.Quit acQuitSaveNone
1410        Else

1420          If blnReturnFromMain = False Then

                ' ** STATUS 4: Access Options
1430            .InitMsg_lbl2.Caption = "Access Options"
1440            DoEvents

                ' ** Set the various Access options.
1450            SetOption_Run True  ' ** Module Function: modStartupFuncs.

                ' ** STATUS 5: References
1460            .InitMsg_lbl2.Caption = "References"
1470            DoEvents

1480            If gblnDev_NoErrHandle = True Then
1490  On Error GoTo 0
1500            Else
1510  On Error GoTo ERRH
1520            End If
                ' ** Check to see if references are broken.
1530            strMessage = "Missing Reference:"
1540            strTestMessage = "Reference List:"
1550            intLenMsg = Len(strMessage)
1560            blnFound = False
1570  On Error Resume Next
1580            For Each ref In Application.References
                  ' ** strMessage only gets path if broken; strTestMessage always gets path.
                  ' ** Array arr_varRef() gets everything.
1590              If ref.IsBroken Then
1600                strMessage = strMessage & vbCrLf & ref.FullPath
1610              End If
1620              strTestMessage = strTestMessage & vbCrLf & ref.FullPath
1630              lngRefs = lngRefs + 1&
1640              lngE = lngRefs - 1&
1650              ReDim Preserve arr_varRef(R_ELEMS, lngE)
1660              arr_varRef(R_RNAM, lngE) = ref.Name
                  ' ** Application.VBE.ActiveVBProject.References(1).Description  ' ** More info, 1-based.
                  ' ** Application.References(1).Name                             ' ** Less info, 1-based.
1670              arr_varRef(R_DSC, lngE) = Application.VBE.ActiveVBProject.References(lngRefs).description
1680              arr_varRef(R_ISBROKE, lngE) = ref.IsBroken
1690              arr_varRef(R_BUILTIN, lngE) = ref.BuiltIn
1700              arr_varRef(R_GUID, lngE) = ref.Guid
1710              If ERR.Number <> 0 Then
1720                blnFound = True
1730                arr_varRef(R_OTHRERR, lngE) = ERR.Number
1740  On Error Resume Next
1750              Else
1760                arr_varRef(R_OTHRERR, lngE) = CLng(0)
1770              End If
1780              arr_varRef(R_PATH, lngE) = ref.FullPath
1790              If ERR.Number <> 0 Then
1800                blnFound = True
1810                arr_varRef(R_PATHERR, lngE) = CBool(True)
1820  On Error Resume Next
1830              Else
1840                arr_varRef(R_PATHERR, lngE) = CBool(False)
1850              End If
1860            Next ref
1870            If ERR <> 0 Or blnFound = True Then
1880              If gblnDev_NoErrHandle = True Then
1890  On Error GoTo 0
1900              Else
1910  On Error GoTo ERRH
1920              End If
1930              DoCmd.Hourglass False
1940              msgResponse = MsgBox("There is a problem with Access references," & vbCrLf & _
                    "and Trust Accountant must shut down." & vbCrLf & vbCrLf & _
                    "Please contact Delta Data, Inc." & vbCrLf & vbCrLf & _
                    "Click Yes to see details before shutting down." & vbCrLf & _
                    "Click No to exit immediately.", vbCritical + vbYesNo, "Missing Reference")
1950              If msgResponse = vbYes Then
1960                ShowRefs  ' ** Procedure: Below.
1970              End If
1980              DoCmd.Quit
1990            Else
2000              If gblnDev_NoErrHandle = True Then
2010  On Error GoTo 0
2020              Else
2030  On Error GoTo ERRH
2040              End If
2050            End If

2060            If Len(strMessage) <> intLenMsg Then
2070              DoCmd.Hourglass False
2080              msgResponse = MsgBox("There is a problem with Access references," & vbCrLf & _
                    "and Trust Accountant must shut down." & vbCrLf & vbCrLf & _
                    "Please contact Delta Data, Inc." & vbCrLf & vbCrLf & _
                    "Click Yes to see details before shutting down." & vbCrLf & _
                    "Click No to exit immediately.", vbCritical + vbYesNo, "Missing Reference")
2090              If msgResponse = vbYes Then
2100                ShowRefs  ' ** Procedure: Below.
2110              End If
2120              DoCmd.Quit
2130            End If

                ' ** STATUS 6: Other Forms
2140            .InitMsg_lbl2.Caption = "Open Forms"
2150            DoEvents

2160          End If  ' ** blnReturnFromMain.

              ' ** Make sure there aren't any other forms open from backup or restore.
2170          intTmp02 = 0
2180          Do While Forms.Count > 1
2190            intTmp02 = intTmp02 + 1
2200            For intX = (Forms.Count - 1) To 0 Step -1
2210              Select Case Forms(intX).Name
                  Case THIS_NAME, "frmEULA", "frmLicense", "frmLicense_Edit", "frmMenu_Main", "frmMenu_Background"
                    ' ** These may be open on new install.
2220              Case Else
2230                DoCmd.Close acForm, Forms(intX).Name
2240                Exit For  ' ** Do them one-at-a-time.
2250              End Select
2260            Next
2270            If intTmp02 >= 5 Then Exit Do  ' ** Don't get caught in a loop.
2280          Loop

              ' ** Let's check reports, too.
2290          intTmp02 = 0
2300          Do While Reports.Count > 0
2310            intTmp02 = intTmp02 + 1
2320            DoCmd.Close acReport, Reports(0).Name, acSavePrompt
2330            If intTmp02 >= 5 Then Exit Do  ' ** Don't get caught in a loop.
2340          Loop

2350          blnFound = False

              ' ** Only on initial open.
2360          If gstrTrustDataLocation = vbNullString Then

                ' ** At this point, neither ConversionCheck() nor Backend_Update() has been run,
                ' ** therefore, InitializeTables() may run into problems if the backend isn't right.
                ' ** However, if this IS a conversion, it's got a brand-new
                ' ** backend and there should be no problem. (I wish I had one of those...)

                ' ** STATUS 7: Data Link
2370            .InitMsg_lbl2.Caption = "Data Link"
2380            DoEvents

                ' ** Relink all tables; assure backend location is found.
                ' ** Contains 22 STATUS SUB items.
2390            blnRetVal = InitializeTables  ' ** Module Function: modStartupFuncs.

2400            Select Case blnRetVal
                Case True

2410              DoCmd.Hourglass True  ' ** Make sure it's still running.
2420              DoEvents

2430              intSizeChks = SZ_OK
2440              blnFound = True  ' ** InitializeTables() completed successfully.

                  ' ** STATUX 0: ControlTipText
                  '.InitMsg_lbl2.Caption = "Control Tips"
                  '.InitMsg_lbl3.Caption = vbNullString
                  'DoEvents

                  ' ** gblnCtlTipTxtShow set in InitializeTables().
                  'Pref_ControlTipText .chkControlTipText, Me  ' ** Module Procedure: modPreferenceFuncs.

                  ' ** Check the mdb size; don't let it run if they get too big.
                  'intSizeChks = SizeChecks  ' ** Module Procedure: modStartupFuncs.
                  'If intSizeChks <> SZ_OK Then
                  '  Select Case intSizeChks
                  '  Case SZ_COMP
                  '    ' ** Send them to the compact menu.
                  '  Case Else
                  '    DoCmd.Quit
                  '  End Select
                  'End If
                  'DoCmd.Hourglass True  ' ** Make sure it's still running.
                  'DoEvents
                  'If intSizeChks = SZ_OK Or intSizeChks = SZ_COMP Then
                  'End If

                  ' ** STATUS 8: Security Sync
2450              .InitMsg_lbl2.Caption = "Security Sync"
2460              .InitMsg_lbl3.Visible = False
2470              DoEvents

                  ' ** Make sure TAImport is in all 3 tables.
                  'CheckTAImport  ' ** Module Function: modSecurityFunctions.

2480              If Security_SyncChk = False Then  ' ** Module Function: modSecurityFunctions.
2490                gblnBadSec = True
2500                .TimerInterval = 0&
2510                DoCmd.Quit
2520              End If

                  ' ** Now check again!
                  'CheckTAImport  ' ** Module Function: modSecurityFunctions.

2530            Case False
                  ' ** This may have been invoked up to 2 times already!
                  ' **   frmLicense may have invoked it.
                  ' **   InitializeTables() may have invoked it.
2540              blnRetVal = False
2550              .TimerInterval = 0&
2560              DoCmd.Quit
2570            End Select
2580          Else
2590            .cmdMenu.Visible = True
2600            .cmdMenu_box1.Visible = True
2610            .cmdMenu_box2.Visible = True
2620            .cmdQuit.Visible = True
2630            .cmdQuit_box1.Visible = True
2640            .cmdQuit_box2.Visible = True
                ' ** Initializing message now has its own label!
2650            DoEvents
2660            .cmdMenu.SetFocus
2670          End If

              ' ** If this is quitting, don't let it do anything else!
2680          If blnRetVal = True Then

2690            DoCmd.Hourglass True  ' ** Make sure it's still running.
2700            DoEvents

2710            If blnReturnFromMain = False Then

                  ' ** STATUS 9: Default Printer
2720              .InitMsg_lbl2.Caption = "Default Printer"
2730              .InitMsg_lbl3.Visible = False  ' ** To be sure.
2740              DoEvents

                  ' ** Check to see if there is a default printer.
                  ' ** If not give the user a choice of whether or not to continue.
                  ' ** VGC 1/14/08: End user got an error here, reading
                  ' **              clsDevs.CurrentDevice.Default, so message never got shown.
2750  On Error Resume Next
2760              Set clsDevs = New clsDevices
2770              blnTmp04 = clsDevs.CurrentDevice.Default
2780              If ERR.Number <> 0 Then
2790                blnTmp04 = False
2800                If gblnDev_NoErrHandle = True Then
2810  On Error GoTo 0
2820                Else
2830  On Error GoTo ERRH
2840                End If
2850              Else
2860                If gblnDev_NoErrHandle = True Then
2870  On Error GoTo 0
2880                Else
2890  On Error GoTo ERRH
2900                End If
2910              End If
2920              If blnTmp04 = False Then
2930                DoCmd.Hourglass False
2940                If MsgBox("Trust Accountant must have a default printer assigned." & vbCrLf & _
                        "Call your local Administrator if you need assistence in setting a default printer." & vbCrLf & _
                        "Do you wish to continue?", vbQuestion + vbYesNo, "Missing Default Printer") = vbNo Then
2950                  DoCmd.Quit
2960                Else
2970                  DoCmd.Hourglass True
2980                End If
2990              End If

3000              DoCmd.Hourglass True  ' ** Make sure it's still running.
3010              DoEvents

                  ' ** STATUS 10: Retrieve Version
3020              .InitMsg_lbl2.Caption = "Version"
3030              DoEvents

3040            End If  ' ** blnReturnFromMain.

3050            dblVersion = Nz(DLookup("Version", "[License Name]"), 0)
3060            strVersion = CStr(dblVersion)
3070            If InStr(strVersion, ".") > 0 Then
3080              If Len(Mid(strVersion, (InStr(strVersion, ".") + 1))) > 1 Then
3090                strVersion = Left(strVersion, InStr(strVersion, ".")) & _
                      Mid(strVersion, (InStr(strVersion, ".") + 1), 1) & "." & Mid(strVersion, (InStr(strVersion, ".") + 2))
3100              Else
3110                strVersion = strVersion & ".00"
3120              End If
3130            Else
3140              strVersion = strVersion & ".0.00"
3150            End If

3160            DoCmd.Hourglass True  ' ** Make sure it's still running.
3170            DoEvents

3180            If blnReturnFromMain = False Then
                  ' ** STATUS 11: Retrieve License
3190              .InitMsg_lbl2.Caption = "License"
3200              DoEvents
3210            End If  ' ** blnReturnFromMain.

3220            strLicense = DecodeString(IniFile_Get("License", "Firm", EncodeString("Call Delta Data, Inc., for Licensing info."), _
                  gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
3230            strLicense = "Licensed to " & strLicense & "    Version " & _
                  IIf(Left(Right(strVersion, 2), 1) = ".", strVersion & "0", strVersion)
3240            .LicenseInfo.Value = strLicense
3250            If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
                  ' ** e.g., C:\VictorGCS_Clients\TrustAccountant\NewDemo\EmptyDatabase
3260              strTmp01 = CurrentBackendPath  ' ** Module Function: modFileUtilities.
3270              If InStr(strTmp01, "\VictorGCS_Clients\TrustAccountant\") > 0 Then
3280                strTmp01 = Mid(strTmp01, (InStr(25, strTmp01, "\") + 1))
3290              End If
3300              strTmp01 = strTmp01 & "  " & AppVersion_GetDta  ' ** Module Function: modAppVersionFuncs.
3310              .DevMsg_lbl.Caption = strTmp01
3320              .DevMsg_lbl.Visible = True
3330              varTmp00 = DLookup("[CoInfo_Name]", "CompanyInformation")
3340              Select Case IsNull(varTmp00)
                  Case True
3350                .DevMsg_lbl2.Caption = "Empty Database"
3360              Case False
3370                If InStr(varTmp00, "&") > 0 Then
3380                  varTmp00 = StringReplace(CStr(varTmp00), "&", "~~") ' ** Module Function: modStringFuncs.
3390                  varTmp00 = StringReplace(CStr(varTmp00), "~~", "&&") ' ** Module Function: modStringFuncs.
3400                End If
3410                .DevMsg_lbl2.Caption = varTmp00
3420              End Select
3430              .DevMsg_lbl2.Visible = True
3440              Frm_PopUpChk  ' ** Module Function: zz_mod_MDEPrepFuncs.
3450            End If

3460            DoCmd.Hourglass True  ' ** Make sure it's still running.
3470            DoEvents

3480            If intSizeChks = SZ_OK Then
                  ' ** It's already gone through InitializeTables(), above.

3490              If blnReturnFromMain = False Then
                    ' ** STATUS 12: Conversion Check
3500                .InitMsg_lbl2.Caption = "Conversion"
3510                DoEvents
3520              End If  ' ** blnReturnFromMain.

3530              If .chkConversionCheck = False Then  ' ** Set during InitializeTables().
3540                blnSkip = True
3550                If blnSkip = False Then
                      ' ** For GTR_REF only.
3560                  DoCmd.OpenForm "frmVersion_Main", , , , , acDialog, THIS_NAME
3570                End If
3580                If blnSkip = False Then
                      ' ** For GTR_REF only.
3590                  DoCmd.OpenForm "frmUser_SecurityNotice", , , , , acDialog, THIS_NAME
3600                End If
3610                intConversionCheck = ConversionCheck  ' ** Module Function: modVersionConvertFuncs2.
                    ' ** Until it gets back from Version_Upgrade_01(), this form is in limbo.
                    ' ** As soon as it gets back, it wants to return here immediately and finish the OnOpen event.
                    ' ** If we skipped the rest, below, then made this invisible, could we
                    ' ** come back here, finish below from a Sub, then make it visible?
3620              Else
                    ' ** Once a conversion has successfully run, don't keep checking.
3630                intConversionCheck = 1
3640              End If

3650              DoCmd.Hourglass True  ' ** Make sure it's still running.
3660              DoEvents

                  ' ** Continue immediately if no conversion necessary.
3670              If intConversionCheck = 1 Then
                    ' ** STATUS 13 is here.
3680                OpenContinues True  ' ** Procedure: Below.
3690              Else
3700                If IsLoaded("frmVersion_Main", acForm) = True Then  ' ** Module Function: modFileUtilities.
                      ' ** I can't quite get it to pop back there immediately!
                      '.Visible = False  'THIS DOESN'T WORK
3710                  DoCmd.SelectObject acForm, "frmVersion_Main", False
3720                  gblnSetFocus = True
3730                End If
3740              End If

3750            Else
3760              OpenContinues False  ' ** Procedure: Below.
3770            End If

3780            DoCmd.Hourglass True  ' ** Make sure it's still running.
3790            DoEvents

3800            If blnReturnFromMain = False Then
                  ' ** STATUS 14: Password Expiration
3810              .InitMsg_lbl2.Caption = "Password Lifetime"
3820              DoEvents
3830            End If  ' ** blnReturnFromMain.

                ' ** Check password expiration.
3840            If Security_PW_Exp_Chk(Me) = False Then  ' ** Module Function: modSecurityFunctions.
3850              Cancel = -1
3860              DoCmd.Quit
3870            End If

                ' ** If a Menu key is pressed while the form is still loading
                ' ** (when the Hourglass is still on), it stores that keystroke,
                ' ** and invokes the OnClick() event, thus closing the form,
                ' ** before the code has gotten to here!
3880  On Error Resume Next
3890            .InitMsg_lbl1.Visible = False
                'If CurrentUser = "Superuser" And blnReturnFromMain = False Then  ' ** Internal Access Function: Trust Accountant login.
                '.DevKeys_lbl1a.Visible = True
                '.DevKeys_lbl1b.Visible = True
                '.DevKeys_lbl2a.Visible = True
                '.DevKeys_lbl2b.Visible = True
                '.DevKeys_lbl3a.Visible = True
                '.DevKeys_lbl3b.Visible = True
                '.DevKeys_lbl4a.Visible = True
                '.DevKeys_lbl4b.Visible = True
                'End If
3900            .Eval_lbl.Top = lngEvalLbl_Top
3910  On Error GoTo ERRH
                ' ** 2467  The expression you entered refers to an object that is closed or doesn't exist.
                ' ** frmMenu_Title
                ' ** Form_Open()
                ' ** Line: 2350

3920          End If  ' ** blnRetVal.
3930        End If  ' ** IsTAOpen().

3940        If blnRetVal = True Then
              ' ** It's a good startup, so continue.

3950          If blnFound = False Then
                ' ** InitializeTables() has been skipped; probably developer startup.
3960            If blnReturnFromMain = False Then
                  ' ** STATUS 15: Foreign Currency Check
3970              .InitMsg_lbl2.Caption = "Foreign Currency Check"
3980              DoEvents
3990            End If

4000            If blnReturnFromMain = False Then
                  ' ** Developer startup, since STATUS SUB 22 has been skipped.
4010              gblnHasForEx = HasForEx_All  ' ** Module Function: modCurrencyFuncs.
4020              HasForEx_Load  ' ** Module Procedure: modCurrencyFuncs.
4030            End If
4040          End If

4050          If blnReturnFromMain = False Then
                ' ** STATUS 16: Finished
4060            .InitMsg_lbl2.Caption = "Done"
4070            DoEvents
4080            blnTimer3 = True
4090            .TimerInterval = 50&
4100          End If  ' ** blnReturnFromMain.

              '.Repaint  ' ** This is what causes error 2467, below.
4110          DoCmd.Hourglass False

4120        End If  ' ** blnRetVal.
4130      End If  ' ** TAImport.
4140    End With

4150    blnTmp04 = False
4160    If Cancel = -1 And blnTmp04 = True Then
          ' ** Just for the record.
4170      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Allowed"
4180    End If

EXITP:
4190    Set clsDevs = Nothing
4200    Set ref = Nothing
4210    Exit Sub

ERRH:
4220    DoCmd.Hourglass False
4230    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Sometimes pushing buttons before the hourglass is off causes this when frmMenu_Main is already open.
4240    Case 2489  ' ** The object '|' isn't open.
          ' ** Ignore; may hit this when I close the window before the Timer stuff plays out.
4250    Case Else
4260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4270    End Select
4280    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

4300  On Error GoTo ERRH

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        Const THIS_PROC As String = "Form_KeyDown"

        ' ** Use bit masks to determine which key was pressed.
4310    intShiftDown = (Shift And acShiftMask) > 0
4320    intAltDown = (Shift And acAltMask) > 0
4330    intCtrlDown = (Shift And acCtrlMask) > 0

4340    intRetVal = KeyCode

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Background: B {CloseBack}
        ' **   Close Form:       X {cmdCloseForm}

        ' ** Shortcut Ctrl-Alt-Shift keys to other forms and functionality:
        ' **   References:       R {ShowRefs}

        ' ** Ctrl-Shift keys.
4350    If intCtrlDown And (Not intAltDown) And intShiftDown Then
4360      Select Case intRetVal
          Case vbKeyB
4370        intRetVal = 0
4380        CloseBack  ' ** Procedure: Below.
4390      Case vbKeyX
            ' ** Close form without moving to another.
4400        If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
4410          intRetVal = 0
4420          SetOption_Dev  ' ** Module Function: modUtilities.
4430          DoCmd.Close acForm, THIS_NAME
4440        End If
4450      End Select
4460    End If

        ' ** Ctrl-Alt-Shift keys.
4470    If intCtrlDown And intAltDown And intShiftDown Then
4480      Select Case intRetVal
          Case vbKeyR
4490        ShowRefs  ' ** Procedure: Below.
4500      End Select
4510    End If

EXITP:
4520    KeyCode = intRetVal
4530    Exit Sub

ERRH:
4540    intRetVal = 0
4550    Select Case ERR.Number
        Case Else
4560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4570    End Select
4580    Resume EXITP

End Sub

Private Sub Form_Timer()
' ** This is a check for that sporadic problem of sitting
' ** here with no buttons visible, and no way to quit.
' ** All these Timer events have become very complicated.
' ** The blnTimer variables are an attempt to prevent some
' ** processes from happening over and over again.

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim obj As Object
        Dim strFormName As String
        Dim lngTmp03 As Long

        Static lngTmp01 As Long

4610    With Me
4620      .TimerInterval = 0&
4630      If gblnSetFocus = True Then
4640        gblnSetFocus = False
4650        lngTmp03 = 0
4660        strFormName = vbNullString
4670  On Error Resume Next
4680        Do Until strFormName = THIS_NAME
4690          lngTmp03 = lngTmp03 + 1&
4700          DoCmd.SelectObject acForm, THIS_NAME, False
4710          strFormName = Screen.ActiveForm.Name
4720          If lngTmp03 > 100& Then Exit Do
4730        Loop
4740  On Error GoTo ERRH
4750        If IsMaximized(Me) = True Then  ' ** Module Function: modWindowFunctions.
4760          DoCmd.Restore
4770        End If
4780      End If
4790      If blnLogoAnim_Key = False Then
4800        If blnInitMsg = False And blnTimer1 = False Then
4810          blnTimer1 = True
4820          If .cmdMenu.Visible = False Then
4830            .cmdMenu.Visible = True
4840            .cmdMenu_box1.Visible = True
4850            .cmdMenu_box2.Visible = True
4860            .cmdQuit.Visible = True
4870            .cmdQuit_box1.Visible = True
4880            .cmdQuit_box2.Visible = True
4890          End If
              ' ** Initializing message now has its own label!
4900          .InitMsg_lbl1.Visible = False  ' ** This does go out at the end of Form_Open()!
4910          .Eval_lbl.Top = lngEvalLbl_Top
4920          If blnTimer2 = False Then
4930            If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
4940              blnTimer2 = True
4950              CmdBars_Hide True  ' ** Module Procedure: modWindowFunctions.
4960              CmdBars_Prop  ' ** Module Procedure: modWindowFunctions.
4970            End If
4980          End If  ' ** blnTimer2.
4990          If intOpenVer > 0 Then
5000            If IsNavPaneOpen = True And strCallingForm = vbNullString Then  ' ** Module Procedure: modWindowFunctions.
                  ' ** Because the NavigateTo Method is not found in Access 2000 and Access 2003,
                  ' ** using a generic object type allows it to be compiled in Access 2000.
5010              Set obj = Application
5020              obj.DoCmd.NavigateTo "acNavigationCategoryObjectType"  ' ** Not documented properly in help.
5030              DoCmd.RunCommand acCmdWindowHide
5040            End If
5050          End If
5060          blnInitMsg = True
5070          .TimerInterval = 25& '1000&
5080        Else
5090          If blnTimer3 = True Then
5100            blnInitMsg = False
5110            .InitMsg_lbl2.Visible = False
5120            .InitMsg_lbl3.Visible = False
5130            DoEvents
5140            If blnLogoAnim_Burst1 = True Then
5150              blnLogoAnim_Burst1 = False
5160              blnLogoAnim_Burst2 = True
5170            End If
5180          End If  ' ** blnTimer3.
5190        End If
5200      Else
5210        blnTimer3 = True
5220      End If
5230      If blnLogoAnim_Burst2 = True Then
5240        If blnLogoAnim_Key = True Then
5250          blnLogoAnim_Key = False
5260          lngTmp01 = 0&
5270        End If
5280        If blnTimer3 = True Then
5290          lngTmp01 = lngTmp01 + 1&
5300          If lngTmp01 < 22& Then
5310  On Error Resume Next
5320            .PictureData = .Controls("Title_Burst_" & Right("00" & CStr(lngTmp01), 2) & "_anim_img")
5330            If ERR.Number <> 0 Then
5340  On Error GoTo ERRH
                  ' ** 2192  The bitmap you specified is not in a device-independent bitmap (.dib) format.
                  ' ** 2113  The value you entered isn't valid for this field.
5350              blnLogoAnim_Burst2 = False
5360            Else
5370  On Error GoTo ERRH
5380              .TimerInterval = 50&
5390            End If
5400          Else
5410            blnLogoAnim_Burst2 = False
5420            blnTimer1 = False
5430            blnTimer2 = False
5440            blnTimer3 = False
5450          End If
5460        End If  ' ** blnTimer3.
5470      End If  ' ** blnLogoAnim_Burst2.
5480    End With

EXITP:
5490    Set obj = Nothing
5500    Exit Sub

ERRH:
5510    Me.TimerInterval = 0&
5520    Select Case ERR.Number
        Case 2192  ' ** The bitmap you specified is not in a device-independent bitmap (.dib) format.
          ' ** Just keep going.
5530      blnLogoAnim_Burst2 = False
5540    Case 2489  ' ** The object 'frmMenu_Title' isn't open.
          ' ** A little too quick to click!
5550    Case Else
5560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5570    End Select
5580    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

5610    If DataErr <> 0 Then
5620      Select Case DataErr
          Case 2052  ' ** There isn't enough free memory to update the display. Close unneeded programs and try again.
5630        Response = acDataErrContinue
5640      Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
            ' ** Sometimes during a Quit from the License page, Access gets confused.
5650        Response = acDataErrContinue
            'Case 2473  ' ** The expression On Exit you entered as the event property setting produced the following error:
            '  Response = acDataErrContinue  ' ** There was a problem referencing a property or method of the object.
5660      Case Else
5670        MsgBox "Error: " & CStr(DataErr), vbInformation + vbOKOnly, "Error"
5680        Response = acDataErrDisplay
            'If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
            'zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
5690      End Select
5700    End If

        ' ** AcDataError enumeration:
        ' **   0  acDataErrContinue  Ignore the error and continue without displaying the default Microsoft Access
        ' **                         error message. A custom error message may be displayed in place of the default
        ' **                         error message.
        ' **   1  acDataErrDisplay   Display the default Microsoft Access error message. (Default)
        ' **   2  acDataErrAdded     Don't display the default Microsoft Access error message. The entry may be
        ' **                         added to the combo box list in the NotInList event procedure. After the entry
        ' **                         is added, Microsoft Access updates the list by requerying the combo box.
        ' **                         Microsoft Access then rechecks the string against the combo box list, and saves
        ' **                         the value in the NewData argument in the field the combo box is bound to. If
        ' **                         the string is not in the list, then Microsoft Access displays an error message.

EXITP:
5710    Exit Sub

ERRH:
5720    Select Case ERR.Number
        Case Else
5730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5740    End Select
5750    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

5810    With Me
5820      SavePosition .hwnd  ' ** Module Function: modMonitorFuncs.
5830    End With

EXITP:
5840    Exit Sub

ERRH:
5850    Select Case ERR.Number
        Case Else
5860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5870    End Select
5880    Resume EXITP

End Sub

Private Sub cmdMenu_Click()

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMenu_Click"

        Dim strDocName As String

5910    DoCmd.Hourglass True
5920    DoEvents

5930    gblnSetFocus = True
5940    strDocName = "frmMenu_Main"
5950    If IsLoaded(strDocName, acForm) = False Then  ' ** Module Function: modFileUtilities.
          ' ** Quickly clicking back-and-forth sometimes leaves nothing!
5960      gblnBeenToBackup = False
5970      DoCmd.OpenForm strDocName, , , , , , THIS_NAME
5980      DoCmd.Close acForm, THIS_NAME
5990    End If

EXITP:
6000    Exit Sub

ERRH:
6010    DoCmd.Hourglass False
6020    Select Case ERR.Number
        Case Else
6030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6040    End Select
6050    Resume EXITP

End Sub

Private Sub cmdQuit_Click()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdQuit_Click"

        Dim msgResponse As VbMsgBoxResult

6110    Select Case gblnBeenToBackup
        Case True
6120      QuitNow  ' ** Module Procedure: modStartupFuncs.
6130    Case False
6140      msgResponse = MsgBox("Before you quit, it is recommended that you backup your data." & vbCrLf & vbCrLf & _
            "Would you like to do this now?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Backup Trust Data")
6150      EmptyTable_Tmp_SP  ' ** Module Procedure: modStatementParamFuncs2.
6160      If msgResponse = vbYes Then
6170        DoCmd.Close acForm, THIS_NAME
6180        gblnBeenToBackup = True
6190        DoCmd.OpenForm "frmBackupRestore", , , , , , (THIS_NAME & "~" & THIS_PROC)
6200      ElseIf msgResponse = vbNo Then
6210        QuitNow  ' ** Module Procedure: modStartupFuncs.
6220      Else
            ' ** Includes Cancel and Escape
6230      End If
6240    End Select

EXITP:
6250    Exit Sub

ERRH:
6260    Select Case ERR.Number
        Case Else
6270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6280    End Select
6290    Resume EXITP

End Sub

Private Sub cmdTitleBurst_DblClick(Cancel As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTitleBurst_DblClick"

6310    With Me
6320      If .TimerInterval = 0& Then
6330        If gblnSetFocus = False And blnInitMsg = False And blnLogoAnim_Burst1 = False And blnLogoAnim_Burst2 = False Then
              ' ** To make sure it doesn't get called while something else is going on.
6340          blnLogoAnim_Key = True
6350          blnLogoAnim_Burst2 = True
6360          .TimerInterval = 50&
6370        End If
6380      End If
6390    End With

EXITP:
6400    Exit Sub

ERRH:
6410    Select Case ERR.Number
        Case Else
6420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6430    End Select
6440    Resume EXITP

End Sub

Public Sub OpenContinues(blnRunUpdates As Boolean)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "OpenContinues"

6510    With Me

6520      DoCmd.Hourglass True  ' ** Make sure it's still running.
6530      DoEvents

6540      If blnReturnFromMain = False Then
            ' ** STATUS 13: Backend Update
6550        .InitMsg_lbl2.Caption = "Backend Update"
6560        DoEvents
6570      End If  ' ** blnReturnFromMain.

6580      If intSizeChks = SZ_OK And blnRunUpdates = True Then
6590        If blnReturnFromMain = False Then
              ' ** Do all manner of updates, upgrades, and checks.
6600          Backend_Update THIS_NAME  ' ** Module Function: modVersionConvertFuncs1
6610        End If  ' ** blnReturnFromMain.
            ' ** Initialize glngUserCntLedger.
6620        IsSingleUser True  ' ** Module Function: modSecurityFunctions.
6630      End If

6640      DoCmd.Hourglass True  ' ** Make sure it's still running.
6650      DoEvents

6660      If blnReturnFromMain = False Then

            ' ** STATUX 0: Load Time
            '.InitMsg_lbl2.Caption = "Load Time"
            'DoEvents

6670        If IsLoaded("frmLicense", acForm) = True Then  ' ** Module Function: modFileUtilities.
6680          DoCmd.SelectObject acForm, "frmLicense", False
              'Form_frmLicense.Tag = THIS_NAME  ' ** If it's already open, it's already got an OpenArgs with it's source.
              'gdatLoadEnd = Now()
              'SaveLoadTime  ' ** Module Function: modStartupFuncs.
6690          DoCmd.Close acForm, THIS_NAME
6700        Else
              'If intSizeChks = SZ_COMP Then
              '  gdatLoadEnd = Now()
              '  SaveLoadTime  ' ** Module Function: modStartupFuncs.
              '  CurrentBackendCompact True  ' ** Module Function: modFileUtilities.
              'Else
              'gdatLoadEnd = Now()
              'SaveLoadTime  ' ** Module Function: modStartupFuncs.
6710          DoCmd.SelectObject acForm, THIS_NAME, False
              'End If
6720        End If

6730        EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
            'LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.

6740      End If  ' ** blnReturnFromMain.

6750      .TimerInterval = 500&

6760      If gstrReturningForm = "frmVersion_Main" Then
6770        gstrReturningForm = vbNullString
6780        DoCmd.Hourglass False
6790      Else
6800        DoCmd.Hourglass True  ' ** Make sure it's still running.
6810        DoEvents
6820      End If

6830    End With

EXITP:
6840    Exit Sub

ERRH:
6850    DoCmd.Hourglass False
6860    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Ignore; may hit this when I close the window before the Timer stuff plays out.
6870    Case 2489  ' ** The object '|' isn't open.
          ' ** Ignore; may hit this when I close the window before the Timer stuff plays out.
6880    Case Else
6890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6900    End Select
6910    Resume EXITP

End Sub

Private Sub cmdEULA_Click()

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEULA_Click"

        Dim strMsg As String
        Dim blnRetVal As Boolean

7010    blnRetVal = True

7020    strMsg = DemoLicense_Get  ' ** Module Functions: modSecurityFunctions.
7030    If strMsg <> RET_ERR And strMsg <> "#MISSING" Then
7040      If Left(strMsg, 5) <> "#EULA" Then
7050        gdatAccept = CDate(Left(strMsg, (InStr(strMsg, "~") - 1)))
7060        gstrAccept = Mid(strMsg, (InStr(strMsg, "~") + 1))
7070      Else
7080        DoCmd.OpenForm "frmEula", acNormal, , , , acDialog, THIS_NAME
7090        DoEvents
7100        If InStr(gstrAccept, "Accept") = 0 Then
7110          blnRetVal = False
7120        End If
7130      End If
7140    Else
          ' ** There was an error getting the demo license stuff!
7150      blnRetVal = False
7160      Beep
7170      MsgBox "Your End User License Agreement is invalid." & vbCrLf & vbCrLf & _
            "Please contact Delta Data, Inc.", vbCritical + vbOKOnly, "License Invalid"
7180    End If

EXITP:
7190    Exit Sub

ERRH:
7200    Select Case ERR.Number
        Case Else
7210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7220    End Select
7230    Resume EXITP

End Sub

Private Sub ShowRefs()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowRefs"

        Dim strDocName As String

7310    strDocName = "frmReferences"
7320    DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME

EXITP:
7330    Exit Sub

ERRH:
7340    Select Case ERR.Number
        Case Else
7350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7360    End Select
7370    Resume EXITP

End Sub

Public Sub CloseBack()

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "CloseBack"

7410    If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
7420      gblnDev_NoAppBackground = True
7430      If IsLoaded("frmMenu_Background", acForm) = True Then  ' ** Module Function: modFileUtilities.
7440        DoCmd.Close acForm, "frmMenu_Background"
7450      End If
7460    End If

EXITP:
7470    Exit Sub

ERRH:
7480    Select Case ERR.Number
        Case Else
7490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7500    End Select
7510    Resume EXITP

End Sub
