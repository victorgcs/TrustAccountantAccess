Attribute VB_Name = "modReportFunctions"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modReportFunctions"

'VGC 07/24/2017: CHANGES!

' ** The GroupLevel property setting is an array,
' ** with up to 10 Group Levels, 0-9.
' ** Reports(0).GroupLevel(0).Properties.Count = 7
' **   0  ControlSource  String
' **   1  SortOrder
            ' ** The SortOrder property uses the following settings.
            ' **   True   Descending  Sorts values in descending (Z to A, 9 to 0) order.
            ' **   False  Ascending   Sorts values in ascending (A to Z, 0 to 9) order. (Default)
' **   2  GroupHeader
' **   3  GroupFooter
            ' ** The GroupHeader and GroupFooter properties use the following settings.
            ' **   True   Creates a group header or footer.
            ' **   False  Doesn't create a group header or footer. (Default)
' **   4  GroupOn
            ' ** The GroupOn property settings available for a field depend on its data type,
            ' ** as the following table shows. For an expression, all of the settings
            ' ** are available. The default setting for all data types is Each Value.
            ' **   Text
            ' **   ====
            ' **     0  Each Value         The same value in the field or expression. (Default)
            ' **     1  Prefix Characters  The same first n number of characters in the field or expression.
            ' **   Date/Time
            ' **   =========
            ' **     0  Each Value         The same value in the field or expression. (Default)
            ' **     2  Year               Dates in the same calendar year.
            ' **     3  Qtr                Dates in the same calendar quarter.
            ' **     4  Month              Dates in the same month.
            ' **     5  Week               Dates in the same week.
            ' **     6  Day                Dates on the same date.
            ' **     7  Hour               Times in the same hour.
            ' **     8  Minute             Times in the same minute.
            ' **   AutoNumber, Currency, Number
            ' **   ============================
            ' **     0  Each Value         The same value in the field or expression. (Default)
            ' **     9  Interval           Values within an interval you specify.
' **   5  GroupInterval
            ' ** The GroupInterval property specifies an interval value that records are grouped by.
            ' **   This interval differs depending on the data type and GroupOn property setting of the field
            ' **   or expression you're grouping on. For example, you can set the GroupInterval property to 1
            ' **   if you want to group records by the first character of a Text field, such as ProductName.
            ' ** The GroupInterval property settings are Long values that depend on the field's data type and
            ' **   its GroupOn property setting. The default GroupInterval setting is 1.
            ' ** You can set this property by using the Sorting And Grouping box, a macro, or Visual Basic.
            ' ** You can set the GroupInterval property only in report Design view or in the Open event procedure of a report.
            ' ** Here are examples of GroupInterval property settings for different field data types.
            ' **   All        Each value         Set to 1. (Default)
            ' **   Text       Prefix characters  Set to 3 for grouping by the first three characters in the field
            ' **                                   (for example, Chai, Chartreuse, and Chang would be grouped together).
            ' **   Date/Time  Week               Set to 2 to return data in biweekly groups.
            ' **   Date/Time  Hour               Set to 12 to return data in half-day groups.
' **   6  KeepTogether
            ' ** The KeepTogether property for a group uses the following settings.
            ' **   0  No                 Prints the group without keeping the group header, detail section, and group
            ' **                         footer on the same page. (Default)
            ' **   1  Whole Group        Prints the group header, detail section, and group footer on the same page.
            ' **   2  With First Detail  Prints the group header on a page only if it can also print the first detail record.

' ** You can use the CreateGroupLevel method to specify a field or expression on which to group or sort data in a report. Long.
' **     expression.CreateGroupLevel(ReportName, Expression, Header, Footer)
' ** expression  Required          An expression that returns one of the objects in the Applies To list.
' ** ReportName  Required String   A string expression identifying the name of the report that will contain the new group level.
' ** Expression  Required String   A string expression identifying the field or expression to sort or group on.
' ** Header      Required Integer  An Integer value that indicates a field or expression will have an associated group header.
' **                               If the header argument is True (–1), the field or expression will have a group header.
' **                               If the header argument is False (0), the field or expression won't.
' **                               You can create a header by setting the argument to True.
' ** Footer      Required Integer  An Integer value that indicates a field or expression will have an associated group footer.
' **                               If the footer argument is True (–1), the field or expression will have a group footer.
' **                               If the footer argument is False (0), the field or expression won't.
' **                               You can create a footer by setting the argument to True.

Private Type PRTMIP_STR
  strRGB As String * 28
End Type

Private Type PRTMIP_TYPE
  xLeftMargin As Long
  yTopMargin As Long
  xRightMargin As Long
  yBotMargin As Long
  fDataOnly As Long
  xWidth As Long
  yHeight As Long
  fDefaultSize As Long
  cxColumns As Long
  yColumnSpacing As Long
  xRowSpacing As Long
  rItemLayout As Long
  fFastPrint As Long
  fDatasheet As Long
End Type

Private Type DEVMODE_STR
  RGB As String * 94
End Type

Private Type DEVMODE_TYPE
  strDeviceName As String * 16
  intSpecVersion As Integer
  intDriverVersion As Integer
  intSize As Integer
  intDriverExtra As Integer
  lngFields As Long
  intOrientation As Integer
  intPaperSize As Integer
  intPaperLength As Integer
  intPaperWidth As Integer
  intScale As Integer
  intCopies As Integer
  intDefaultSource As Integer
  intPrintQuality As Integer
  intColor As Integer
  intDuplex As Integer
  intResolution As Integer
  intTTOption As Integer
  intCollate As Integer
  strFormName As String * 16
  lngPad As Long
  lngBits As Long
  lngPW As Long
  lngPH As Long
  lngDFI As Long
  lngDFr As Long
End Type

Private Type PRINTDLG_TYPE
  lStructSize As Long
  hwndOwner As Long
  hDevMode As Long
  hDevNames As Long
  hDC As Long
  flags As Long
  nFromPage As Integer
  nToPage As Integer
  nMinPage As Integer
  nMaxPage As Integer
  nCopies As Integer
  hInstance As Long
  lCustData As Long
  lpfnPrintHook As Long
  lpfnSetupHook As Long
  lpPrintTemplateName As String
  lpSetupTemplateName As String
  hPrintTemplate As Long
  hSetupTemplate As Long
End Type

' ** PRINTDLG_TYPE structure:
' **   lStructSize          The size in bytes of this structure.
' **   hwndOwner            A handle to the window opening the dialog box, if any.
' **   hDevMode             A handle to the memory block holding the information contained in a DEVMODE structure.
' **                        This data specifies information about the printer.
' **   hDevNames            A handle to the memory block holding the information contained in a DEVNAMES structure.
' **                        This data specifies the driver name, printer name, and port name(s) of the printer.
' **   hdc                  Receives either a device context or an information context (depending on the value set as
' **                        flags) to the printer the user selected.
' **   flags                Zero or more of the following flags specifying various options for creating the Print or
' **                        Print Setup dialog. Note that when PrintDlg returns, many of these flags will be set by the
' **                        function to indicate selections by the user:
' **                          PD_ALLPAGES = &H0                        Select the All Pages radio button.
' **                          PD_COLLATE = &H10                        Check the Collate check box. If this flag is set
' **                                                                   when the function returns, the user checked the box
' **                                                                   and the printer doesn't automatically support
' **                                                                   collation. If the box is checked andthe printer does
' **                                                                   support it, this flag will not be set.
' **                          PD_DISABLEPRINTTOFILE = &H80000          Disable the Print to File check box.
' **                          PD_ENABLEPRINTHOOK = &H1000              Use the hook function pointed to by lpfnPrintHook
' **                                                                   to process the Print dialog box's messages.
' **                          PD_ENABLEPRINTTEMPLATE = &H4000          Use the Print dialog box template specified by
' **                                                                   lpPrintTemplateName.
' **                          PD_ENABLEPRINTTEMPLATEHANDLE = &H10000   Use the preloaded Print dialog box template
' **                                                                   specified by hPrintTemplate.
' **                          PD_ENABLESETUPHOOK = &H2000              Use the hook function pointed to by lpfnSetupHook
' **                                                                   to process the Print Setup dialog box's messages.
' **                          PD_ENABLESETUPTEMPLATE = &H8000          Use the Print Setup dialog box template specified by
' **                                                                   lpSetupTemplateName.
' **                          PD_ENABLESETUPTEMPLATEHANDLE = &H20000   Use the preloaded Print Setup dialog box template
' **                                                                   specified by hSetupTemplate.
' **                          PD_HIDEPRINTTOFILE = &H100000            Hide the Print to File check box.
' **                          PD_NONETWORKBUTTON = &H200000            Do not display any buttons associated with the
' **                                                                   network.
' **                          PD_NOPAGENUMS = &H8                      Disable the Page Range radio button and edit boxes.
' **                          PD_NOSELECTION = &H4                     Disable the Selection radio button.
' **                          PD_NOWARNING = &H80                      Do not warn the user if there is no default printer.
' **                          PD_PAGENUMS = &H2                        Select the Page Range radio button.
' **                          PD_PRINTSETUP = &H40                     Display the Print Setup dialog box instead of the
' **                                                                   Print dialog box.
' **                          PD_PRINTTOFILE = &H20                    Select the Print to File check box.
' **                          PD_RETURNDC = &H100                      Return a device context to the selected printer as
' **                                                                   hdc.
' **                          PD_RETURNDEFAULT = &H400                 Instead of displaying either dialog box, simply load
' **                                                                   information about the default printer into hDevMode
' **                                                                   and hDevNames. For this to work, those two values
' **                                                                   must be set to 0 before calling the function.
' **                          PD_RETURNIC = &H200                      Return an information context to the selected
' **                                                                   printer as hdc.
' **                          PD_SELECTION = &H1                       Select the Selection radio button.
' **                          PD_SHOWHELP = &H800                      Display the Help button.
' **                          PD_USEDEVMODECOPIES = &H40000            Same as PD_USEDEVMODECOPIESANDCOLLATE.
' **                          PD_USEDEVMODECOPIESANDCOLLATE = &H40000  If the printer does not automatically support
' **                                                                   multiple copies or collation, disable the
' **                                                                   corresponding options in the dialog box. The number
' **                                                                   of copies to print and the collation setting will be
' **                                                                   placed into hDevMode. The information returned to
' **                                                                   this structure will specify the number of pages and
' **                                                                   the collation which the program must print with --
' **                                                                   the printer will print the copies or collate itself.
' **   nFromPage            The value entered in the From Page text box, specifying which page begin printing at.
' **   nToPage              The value entered in the To Page text box, specifying which page to stop printing at.
' **   nMinPage             The minimum allowable value for nFromPage and nToPage.
' **   nMaxPage             The maximum allowable value for nFromPage and nToPage.
' **   nCopies              The number of copies the program needs to print.
' **   hInstance            A handle to the application instance which has the desired dialog box template.
' **   lCustData            A program-defined value to pass to whichever hook function is used.
' **   lpfnPrintHook        A handle to the program-defined hook function to use to process the Print dialog box's
' **                        messages.
' **   lpfnSetupHook        A handle to the program-defined hook function to use to process the Print Setup dialog box's
' **                        messages.
' **   lpPrintTemplateName  The name of the Print dialog box template to use from the application instance specified by
' **                        hInstance.
' **   lpSetupTemplateName  The name of the Print Setup dialog box template to use from the application instance specified
' **                        by hInstance.
' **   hPrintTemplate       A handle to the preloaded Print dialog box template to use.
' **   hSetupTemplate       A handle to the preloaded Print Setup dialog box template to use

Private Type DEVNAMES_TYPE
  wDriverOffset As Integer
  wDeviceOffset As Integer
  wOutputOffset As Integer
  wDefault As Integer
  extra As String * 100
End Type

Private Type DOCINFO
  pDocName As String
  pOutputFile As String
  pDatatype As String
End Type

Private TA_cbar As Office.CommandBar, cbctlTA_StdPrint As Office.CommandBarControl
Private cbctlTA_StdPrintD As Office.CommandBarControl, cbctlTA_ZeroPrint As Office.CommandBarControl
Private cbctlTA_ZeroPrintD As Office.CommandBarControl, cbctlTA_Zoom As Office.CommandBarControl, cbctlTA_Close As Office.CommandBarControl

Private lngCourtRpts As Long, arr_varCourtRpt() As Variant

' ** DevModeOrientation enumeration:
Public Const DM_PORTRAIT  As Long = 1&
Public Const DM_LANDSCAPE As Long = 2&
' **

Public Function TAReports_Access2007(blnIsOpen As Boolean) As Boolean

100   On Error GoTo ERRH

        Const THIS_PROC As String = "TAReports_Access2007"

        Dim intOpenVer As Integer
        Dim blnRetVal As Boolean

110     blnRetVal = True

120     If IsAccess2007 = True Then  ' ** Module Function: modXAccess_07_10_Funcs.
130       intOpenVer = 12
140     ElseIf IsAccess2010 = True Then  ' ** Module Function: modXAccess_07_10_Funcs.
150       intOpenVer = 14
160     Else
170       intOpenVer = 10
180     End If
190     If intOpenVer > 0 Then
200       Select Case intOpenVer
          Case 12
210         SetReport_Access2007 blnIsOpen  ' ** Module Function: modXAccess_07_10_Funcs.
220       Case 14
230         SetReport_Access2010 blnIsOpen  ' ** Module Function: modXAccess_07_10_Funcs.
240       End Select
250     End If

EXITP:
260     TAReports_Access2007 = blnRetVal
270     Exit Function

ERRH:
280     blnRetVal = False
290     Select Case ERR.Number
        Case 2585  ' ** This action can't be carried out while processing a form or report event.
          ' ** Ignore.
300     Case Else
310       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
320     End Select
330     Resume EXITP

End Function

Public Function TAReports_Access2010(blnIsOpen As Boolean) As Boolean

400   On Error GoTo ERRH

        Const THIS_PROC As String = "TAReports_Access2010"

        Dim blnRetVal As Boolean

410     blnRetVal = True

        ' ** For now, these will be the same.
420     blnRetVal = TAReports_Access2007(blnIsOpen)  ' ** Function: Above.

EXITP:
430     TAReports_Access2010 = blnRetVal
440     Exit Function

ERRH:
450     blnRetVal = False
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
480     End Select
490     Resume EXITP

End Function

Public Function TAReports_Print1() As Boolean
' ** When a report is in Preview, and the Print button on that Preview is
' ** clicked (our TAReports CommandBar), some special handling may be needed.
' ** The Court Reports Summary may fill a temporary table once for a
' ** beginning balance, save it to a control or variable, then empty and
' ** refill it again for an ending balance. Therefore, the table's state
' ** while that report is in Preview is only the final state of that table.
' ** In order to then Print the report, that same fill, empty, fill process
' ** has to be repeated, and an open Preview, being still linked to that table,
' ** prevents it from being deleted, re-created, and/or emptied and refilled.
' ** Therefore, the Preview must be closed before carrying out the Print
' ** action invoked from the Preview CommandBar.
' ** Called by mcrPrint_CA_0, which in turn is:
' **   Invoked by Button 2 (&PrintZero) on the TAReports command bar.
' **
' *******************************************************************
' ** It appears that the only reports needing this are the CA ones!
' *******************************************************************

500   On Error GoTo ERRH

        Const THIS_PROC As String = "TAReports_Print1"

        Dim frm As Access.Form
        Dim strTmp01 As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

        ' ** Load arr_varCourtRpt() array with Court Summary and AssetList reports.
510     TAReports_LoadSpecRpts  ' ** Function: Below.

520     gblnCrtRpt_Zero = True: gblnCrtRpt_ZeroDialog = False
530     For Each frm In Forms
540       With frm
550         If Left(.Name, 19) = "frmRpt_CourtReports" Then
              ' ** frmRpt_CourtReports_CA
560           strTmp01 = Right(.Name, 2)  ' ** Get report version: CA, FL, NS, NY.
570           Exit For
580         End If
590       End With
600     Next
610     For lngX = 0& To (lngCourtRpts - 1&)
620       If IsLoaded(arr_varCourtRpt(lngX), acReport) = True Then  ' ** Module Function: modFileUtilities.
630         DoCmd.Close acReport, arr_varCourtRpt(lngX)
640       End If
650     Next
660     Select Case strTmp01
        Case "CA"
670       Forms("frmRpt_CourtReports_CA").cmdPrint00_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
680     Case "FL"
690       Forms("frmRpt_CourtReports_FL").cmdPrint00_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
700     Case "NS"
710       Forms("frmRpt_CourtReports_NS").cmdPrint00_Click  ' ** Form Procedure: frmRpt_CourtReports_NS.
          'Case "NY"
          '  Forms("frmRpt_CourtReports_NY").cmdPrint00_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
720     End Select
730     gblnCrtRpt_Zero = False

EXITP:
740     Set frm = Nothing
750     TAReports_Print1 = blnRetVal
760     Exit Function

ERRH:
770     blnRetVal = False
780     Select Case ERR.Number
        Case Else
790       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
800     End Select
810     Resume EXITP

End Function

Public Function TAReports_Print2() As Boolean
' ** Like TAReports_Print1(), above, but invokes the Print
' ** Dialog box instead of going directly to the printer.
' *******************************************************************
' ** CURRENTLY NOT WORKING! Invoking the command prints the screen!
' *******************************************************************

900   On Error GoTo ERRH

        Const THIS_PROC As String = "TAReports_Print2"

        Dim frm As Access.Form
        Dim strTmp01 As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

        ' ** Load arr_varCourtRpt() array with Court Summary and AssetList reports.
910     TAReports_LoadSpecRpts  ' ** Function: Below.

920     gblnCrtRpt_Zero = True: gblnCrtRpt_ZeroDialog = True
930     For Each frm In Forms
940       With frm
950         If Left(.Name, 19) = "frmRpt_CourtReports" Then
              ' ** frmRpt_CourtReports_CA
960           strTmp01 = Right(.Name, 2)  ' ** Get report version: CA, FL, NS, NY.
970           Exit For
980         End If
990       End With
1000    Next
1010    For lngX = 0& To (lngCourtRpts - 1&)
1020      If IsLoaded(arr_varCourtRpt(lngX), acReport) = True Then  ' ** Module Function: modFileUtilities.
1030        DoCmd.Close acReport, arr_varCourtRpt(lngX)
1040      End If
1050    Next
1060    Select Case strTmp01
        Case "CA"
1070      Forms("frmRpt_CourtReports_CA").cmdPrint00_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
1080    Case "FL"
1090      Forms("frmRpt_CourtReports_FL").cmdPrint00_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
1100    Case "NS"
1110      Forms("frmRpt_CourtReports_NS").cmdPrint00_Click  ' ** Form Procedure: frmRpt_CourtReports_NS.
          'Case "NY"
          '  Forms("frmRpt_CourtReports_NY").cmdPrint00_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
1120    End Select
1130    gblnCrtRpt_Zero = False: gblnCrtRpt_ZeroDialog = False

EXITP:
1140    Set frm = Nothing
1150    TAReports_Print2 = blnRetVal
1160    Exit Function

ERRH:
1170    blnRetVal = False
1180    Select Case ERR.Number
        Case Else
1190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
1200    End Select
1210    Resume EXITP

End Function

Private Function TAReports_LoadSpecRpts() As Boolean
' ** Loads a list of the court reports needing special handling.
' ** These are the ones that require intervention if the
' ** Print button is invoked from the Print Preview screen.
' ** The array is used by:
' **   Above:
' **     TAReports_Print1()
' **     TAReports_Print2()
' *******************************************************************
' ** It appears that the only reports needing this are the CA ones!
' *******************************************************************

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "TAReports_LoadSpecRpts"

        Dim prj As Access.CurrentProject, rpt As Access.AccessObject
        Dim lngRpts As Long
        Dim strRptName As String
        Dim intPos01 As Integer
        Dim blnRetVal As Boolean

1310    blnRetVal = True

1320    lngCourtRpts = 0&
1330    ReDim arr_varCourtRpt(0)

1340    Set prj = Application.CurrentProject
1350    With prj
1360      lngRpts = .AllReports.Count
1370      For Each rpt In .AllReports
1380        strRptName = vbNullString
1390        With rpt
1400          strRptName = .Name
1410          intPos01 = InStr(strRptName, "_")
1420          If intPos01 > 0 Then
1430            If Mid(strRptName, intPos01, 3) = "_00" Then
1440              lngCourtRpts = lngCourtRpts + 1&
1450              ReDim Preserve arr_varCourtRpt(lngCourtRpts - 1&)
1460              arr_varCourtRpt(lngCourtRpts - 1&) = strRptName
1470            End If
1480          End If
1490        End With
1500      Next
1510    End With
        'X rptCourtRptCA_00
        'rptCourtRptFL_00
        'rptCourtRptNS_00
        'rptCourtRptNY_00
        'X rptCourtRptCA_00A
        'rptCourtRptNS_00A
        'rptCourtRptNY_00A
        'X rptCourtRptCA_00B
        'rptCourtRptFL_00B
        'rptCourtRptNS_00DA

EXITP:
1520    Set rpt = Nothing
1530    Set prj = Nothing
1540    TAReports_LoadSpecRpts = blnRetVal
1550    Exit Function

ERRH:
1560    blnRetVal = False
1570    Select Case ERR.Number
        Case Else
1580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
1590    End Select
1600    Resume EXITP

End Function

Public Function TAReports_SetZero(blnOn As Boolean, Optional varReset As Variant) As Boolean
' ** Set the buttons on the TAReports Command Bar
' ** appropriate to those requiring special Handling.
' ** Zero refers to it only being for all the rptCourtRptXX_00
' ** series of reports, including CA, FL, and NS.
' ** The buttons are identifiable by their blue color scheme.
' ** Called by:
' **   frmMenu_Main
' **     Form_Open()  False
' **     Form_Unload()  False
' **   frmMenu_Report
' **     Form_Open()  False
' **     Form_Close()  False
' **   frmRpt_CourtReports_CA
' **     cmdPreview00_Click()  True
' **     cmdPrint00_Click()  False
' **     cmdPreview01_Click()  False
' **     cmdPrint01_Click()  False
' **     ...
' **     cmdPreview11_Click()  False
' **     cmdPrint11_Click()  False
' **     cmdPrintAll_Click()  False, True
' **   frmRpt_CourtReports_FL
' **     cmdPreview00_Click()  False  'True
' **     cmdPrint00_Click()  False
' **     ...
' **     cmdPreview05_Click()  False
' **     cmdPrint05_Click()  False
' **     cmdPrintAll_Click()  False, False  'True
' **   frmRpt_CourtReports_NY
' **     cmdPreview00_Click()  False
' **     cmdPrint00_Click()  False
' **     ...
' **     cmdPreview09_Click()  False
' **     cmdPrint09_Click()  False
' **     cmdPrintAll_Click()  False, False  'True, False ?
' **   rptCourtRptCA_00
' **     Report_Activate()  True
' **     Report_Close()  False
' **   rptCourtRptCA_00A
' **     Report_Activate()  True
' **     Report_Close()  False
' **   rptCourtRptFL_00
' **     Report_Activate()  False
' **   rptCourtRptFL_00B
' **     Report_Activate()  False
' *******************************************************************
' ** It appears that the only reports needing this are the CA ones!
' *******************************************************************

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "TAReports_SetZero"

        Dim blnRetVal As Boolean

1710    If IsAccess2007 = False Then  ' ** Module Function: modXAccess_07_10_Funcs.
          ' ** The custom command bar doesn't apply to Access 2007, which uses a custom Ribbon.
1720      blnRetVal = TAReports_SetVars  ' ** Function: Below.
1730      If blnRetVal = True Then
            ' ** Buttons 5 and 6, Zoom and Close, always remain visible and enabled.
1740        Select Case blnOn
            Case True  ' ** Used for Summary Court Reports,
1750          With cbctlTA_StdPrint
1760            .Enabled = False
1770            .Visible = False
1780          End With
1790          With cbctlTA_StdPrintD
1800            .Enabled = False
1810            .Visible = False
1820          End With
1830          With cbctlTA_ZeroPrint
1840            .Enabled = True
1850            .Visible = True
1860          End With
1870          With cbctlTA_ZeroPrintD
1880            .Enabled = False  'True
1890            .Visible = False  'True
1900          End With
1910          With cbctlTA_Zoom
1920            .Enabled = True
1930            .Visible = True
1940          End With
1950          With cbctlTA_Close
1960            .Enabled = True
1970            .Visible = True
1980          End With
1990        Case False  ' ** Used for non-Summary Court Reports.
2000          With cbctlTA_StdPrint
2010            .Enabled = True
2020            .Visible = True
2030          End With
2040          With cbctlTA_StdPrintD
2050            .Enabled = True
2060            .Visible = True
2070          End With
2080          With cbctlTA_ZeroPrint
2090            .Enabled = False
2100            .Visible = False
2110          End With
2120          With cbctlTA_ZeroPrintD
2130            .Enabled = False
2140            .Visible = False
2150          End With
2160          With cbctlTA_Zoom
2170            .Enabled = True
2180            .Visible = True
2190          End With
2200          With cbctlTA_Close
2210            .Enabled = True
2220            .Visible = True
2230          End With
2240        End Select
2250      End If
2260    End If  ' ** IsAccess2077().

EXITP:
2270    TAReports_SetZero = blnRetVal
2280    Exit Function

ERRH:
2290    blnRetVal = False
2300    Select Case ERR.Number
        Case Else
2310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
2320    End Select
2330    Resume EXITP

End Function

Public Function TAReports_SetAll() As Boolean
' ** Use this in the Immediate Window to make all buttons
' ** of the TAReports Command Bar visible for development.

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "TAReports_SetAll"

        Dim blnRetVal As Boolean

        ' ** Make the TAReports Command Bar visible if it isn't.
2410    If CommandBars("TAReports").Visible = False Then CommandBars("TAReports").Visible = True

2420    blnRetVal = TAReports_SetVars  ' ** Function: Below.
2430    If blnRetVal = True Then
2440      With cbctlTA_StdPrint
2450        .Enabled = True
2460        .Visible = True
2470      End With
2480      With cbctlTA_StdPrintD
2490        .Enabled = True
2500        .Visible = True
2510      End With
2520      With cbctlTA_ZeroPrint
2530        .Enabled = True
2540        .Visible = True
2550      End With
2560      With cbctlTA_ZeroPrintD
2570        .Enabled = True
2580        .Visible = True
2590      End With
2600      With cbctlTA_Zoom
2610        .Enabled = True
2620        .Visible = True
2630      End With
2640      With cbctlTA_Close
2650        .Enabled = True
2660        .Visible = True
2670      End With
2680    End If

2690    Beep

EXITP:
2700    TAReports_SetAll = blnRetVal
2710    Exit Function

ERRH:
2720    blnRetVal = False
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
2750    End Select
2760    Resume EXITP

End Function

Public Function TAReports_SetVars() As Boolean
' ** Initialize the TAReports Command Bar Control variables.

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "TAReports_SetVars"

        Dim cbctl As Office.CommandBarControl
        Dim blnTmp01 As Boolean, blnTmp02 As Boolean, blnTmp03 As Boolean, blnTmp04 As Boolean, blnTmp05 As Boolean, blnTmp06 As Boolean
        Dim blnRetVal As Boolean

2810    blnRetVal = False
2820    blnTmp01 = False: blnTmp02 = False: blnTmp03 = False: blnTmp04 = False: blnTmp05 = False: blnTmp06 = False

2830    If TA_cbar Is Nothing Then
2840      Set TA_cbar = CommandBars("TAReports")
2850    End If

2860    With TA_cbar
2870      For Each cbctl In .Controls
2880        With cbctl
2890          If .BuiltIn = True And .Caption = "&Print..." And .HelpContextId = 3843 Then
2900            blnTmp01 = True
2910            Set cbctlTA_StdPrint = TA_cbar.Controls(.index)
2920          ElseIf .BuiltIn = True And .Caption = "Pri&nt..." And .HelpContextId = 3106 Then
2930            blnTmp02 = True
2940            Set cbctlTA_StdPrintD = TA_cbar.Controls(.index)
2950          ElseIf .BuiltIn = False And .Caption = "&PrintZero" And .HelpContextId = 3843 Then
2960            blnTmp03 = True
2970            Set cbctlTA_ZeroPrint = TA_cbar.Controls(.index)
2980          ElseIf .BuiltIn = False And .Caption = "Pri&ntZeroDialog" And .HelpContextId = 3106 Then
2990            blnTmp04 = True
3000            Set cbctlTA_ZeroPrintD = TA_cbar.Controls(.index)
3010          ElseIf .BuiltIn = True And .Caption = "&Zoom:" And .HelpContextId = 3972 Then
3020            blnTmp05 = True
3030            Set cbctlTA_Zoom = TA_cbar.Controls(.index)
3040          ElseIf .BuiltIn = True And .Caption = "&Close" And .HelpContextId = 3870 Then
3050            blnTmp06 = True
3060            Set cbctlTA_Close = TA_cbar.Controls(.index)
3070          End If
3080        End With
3090      Next
3100    End With
        'FND 1: cbctlTA_StdPrint
        'FND 2: cbctlTA_StdPrintD
        'FND 3: cbctlTA_ZeroPrint
        'FND 4: cbctlTA_ZeroPrintD
        'FND 5: cbctlTA_Zoom
        'FND 6: cbctlTA_Close

3110    If blnTmp01 = False Or blnTmp02 = False Or blnTmp03 = False Or blnTmp04 = False Or blnTmp05 = False Then
3120      If blnTmp01 = False Then
3130        Set cbctlTA_StdPrint = Nothing
3140      End If
3150      If blnTmp02 = False Then
3160        Set cbctlTA_StdPrintD = Nothing
3170      End If
3180      If blnTmp03 = False Then
3190        Set cbctlTA_ZeroPrint = Nothing
3200      End If
3210      If blnTmp04 = False Then
3220        Set cbctlTA_ZeroPrintD = Nothing
3230      End If
3240      If blnTmp05 = False Then
3250        Set cbctlTA_Zoom = Nothing
3260      End If
3270      If blnTmp06 = False Then
3280        Set cbctlTA_Close = Nothing
3290      End If
3300    Else
3310      blnRetVal = True
3320    End If

        ' ** TAReports CommandBar:
        ' **   Button 1:
        ' **     Caption:        &Print...
        ' **     Shortcut Text:
        ' **     ScreenTip       &Print...
        ' **     Style:          Default Style {icon only}
        ' **     OnAction:
        ' **     Help ContextID: 3843
        ' **     HelpFile:       C:\Program Files\Microsoft Office\Office\1033\actip9.hlp
        ' **     BuiltIn:        True
        ' **     BeginGroup:     True
        ' **   Button 2:
        ' **     Caption:        &Print...
        ' **     ShortcutText:   Ctrl+P
        ' **     ScreenTip:      &Print... (Ctrl+P)
        ' **     Style:          Default Style {icon only}
        ' **     OnAction:
        ' **     HelpContextID:  3106
        ' **     HelpFile:       C:\Program Files\Microsoft Office\Office\1033\actip9.hlp
        ' **     BuiltIn:        True
        ' **     BeginGroup:     True
        ' **   Button 3:
        ' **     Caption:        &PrintZero
        ' **     ShortcutText:
        ' **     ScreenTip:      &PrintZero
        ' **     Style:          Default Style {icon only}
        ' **     OnAction:       mcrPrint_CA_0
        ' **     HelpContextID:  3843
        ' **     HelpFile:       C:\Program Files\Microsoft Office\Office\1033\actip9.hlp
        ' **     BuiltIn:        False
        ' **     BeginGroup:     False
        ' **   Button 4:
        ' **     Caption:        Pri&ntZeroDialog
        ' **     ShortcutText:   Ctrl+P
        ' **     ScreenTip:      &Print... (Ctrl+P)
        ' **     Style:          Default Style {icon only}
        ' **     OnAction:
        ' **     HelpContextID:  3106
        ' **     HelpFile:       C:\Program Files\Microsoft Office\Office\1033\actip9.hlp
        ' **     BuiltIn:        False
        ' **     BeginGroup:     False
        ' **   Button 5:
        ' **     Caption:        &Zoom:
        ' **     ShortcutText:
        ' **     ScreenTip:      &Zoom
        ' **     Style:          Show Label
        ' **     OnAction:
        ' **     HelpContextID:  3106
        ' **     HelpFile:       C:\Program Files\Microsoft Office\Office\1033\actip9.hlp
        ' **     BuiltIn:        True
        ' **     BeginGroup:     True
        ' **   Button 6:
        ' **     Caption:        &Print...
        ' **     ShortcutText:
        ' **     ScreenTip:      &Close
        ' **     Style:          Text Only (Always)
        ' **     OnAction:
        ' **     HelpContextID:  3870
        ' **     HelpFile:       C:\Program Files\Microsoft Office\Office\1033\actip9.hlp
        ' **     BuiltIn:        True
        ' **     BeginGroup:     True

EXITP:
3330    Set cbctl = Nothing
3340    TAReports_SetVars = blnRetVal
3350    Exit Function

ERRH:
3360    blnRetVal = False
3370    Select Case ERR.Number
        Case Else
3380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
3390    End Select
3400    Resume EXITP

End Function

Public Function Rpt_Orient_Get(rpt As Access.Report) As Long
' ** Returns a report's current Page Setup orientation.

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_Orient_Get"

        Dim DevString As DEVMODE_STR
        Dim DM As DEVMODE_TYPE
        Dim strDevModeExtra As String
        Dim lngRetVal As Long

3510    If IsNull(rpt.PrtDevMode) = False Then
3520      strDevModeExtra = rpt.PrtDevMode
3530      DevString.RGB = strDevModeExtra
3540      LSet DM = DevString
3550      DM.lngFields = DM.lngFields Or DM.intOrientation  ' ** Initialize fields.
3560      lngRetVal = DM.intOrientation
3570    End If

EXITP:
3580    Rpt_Orient_Get = lngRetVal
3590    Exit Function

ERRH:
3600    Select Case ERR.Number
        Case Else
3610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
3620    End Select
3630    Resume EXITP

End Function

Public Function Rpt_OrientStr_Get(rpt As Access.Report) As String
' ** Returns a report's current Page Setup orientation as a string.

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_OrientStr_Get"

        Dim DevString As DEVMODE_STR
        Dim DM As DEVMODE_TYPE
        Dim strDevModeExtra As String
        Dim strRetVal As String

3710    If IsNull(rpt.PrtDevMode) = False Then
3720      strDevModeExtra = rpt.PrtDevMode
3730      DevString.RGB = strDevModeExtra
3740      LSet DM = DevString
3750      DM.lngFields = DM.lngFields Or DM.intOrientation  ' ** Initialize fields.
3760      Select Case DM.intOrientation
          Case DM_PORTRAIT   '1
3770        strRetVal = "PORTRT"
3780      Case DM_LANDSCAPE  '2
3790        strRetVal = "LNDSCP"
3800      End Select
3810    End If

EXITP:
3820    Rpt_OrientStr_Get = strRetVal
3830    Exit Function

ERRH:
3840    Select Case ERR.Number
        Case Else
3850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
3860    End Select
3870    Resume EXITP

End Function

Public Function Rpt_Margins_Get(rpt As Access.Report, intMode As Integer) As Long
' ** Return a report's specified margin, in Twips (1440 Twips per inch).

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_Margins_Get"

        Dim PrtMipString As PRTMIP_STR
        Dim PM As PRTMIP_TYPE
        Dim lngRetVal As Long

3910    lngRetVal = 0&

3920    If intMode > 0 Then
3930      PrtMipString.strRGB = rpt.PrtMip
3940      LSet PM = PrtMipString
3950      Select Case intMode
          Case 1  ' ** Top.
3960        lngRetVal = PM.yTopMargin
3970      Case 2  ' ** Bottom.
3980        lngRetVal = PM.yBotMargin
3990      Case 3  ' ** Left.
4000        lngRetVal = PM.xLeftMargin
4010      Case 4  ' ** Right.
4020        lngRetVal = PM.xRightMargin
4030      End Select
4040    End If

EXITP:
4050    Rpt_Margins_Get = lngRetVal
4060    Exit Function

ERRH:
4070    lngRetVal = -1
4080    Select Case ERR.Number
        Case Else
4090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
4100    End Select
4110    Resume EXITP

End Function

Public Function Rpt_Margins_Set(rpt As Access.Report, intMode As Integer, lngValue As Long, Optional varQuiet As Variant) As Boolean
' ** Set a report's specified margin, in Twips (1440 Twips per inch).

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_Margins_Set"

        Dim PrtMipString As PRTMIP_STR
        Dim PM As PRTMIP_TYPE
        Dim blnRetVal As Boolean

4210    blnRetVal = True

4220    PrtMipString.strRGB = rpt.PrtMip
4230    LSet PM = PrtMipString

        ' ** Set margins.
4240    Select Case intMode
        Case 1  ' ** Top.
4250      PM.yTopMargin = 1& * lngValue
4260    Case 2  ' ** Bottom.
4270      PM.yBotMargin = 1& * lngValue
4280    Case 3  ' ** Left.
4290      PM.xLeftMargin = 1& * lngValue
4300    Case 4  ' ** Right.
4310      PM.xRightMargin = 1& * lngValue
4320    End Select

        ' ** Update property.
4330    LSet PrtMipString = PM
4340    rpt.PrtMip = PrtMipString.strRGB

4350    Select Case IsMissing(varQuiet)
        Case True
4360      Beep
4370    Case False
4380      If varQuiet = False Then
4390        Beep
4400      End If
4410    End Select

EXITP:
4420    Set rpt = Nothing
4430    Rpt_Margins_Set = blnRetVal
4440    Exit Function

ERRH:
4450    blnRetVal = False
4460    Select Case ERR.Number
        Case Else
4470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
4480    End Select
4490    Resume EXITP

End Function

Public Sub Rpt_MarginsDef_Set(ByVal strReportName As String)
' ** Set a report's Page Setup margins to the default 1 inch all around.
' ** Leaves report open in Design Mode.

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_MarginsDef_Set"

        Dim PrtMipString As PRTMIP_STR
        Dim PM As PRTMIP_TYPE
        Dim rpt As Access.Report

        ' ** Open the report.
4510    DoCmd.OpenReport strReportName, acViewDesign
4520    Set rpt = Reports(strReportName)
4530    PrtMipString.strRGB = rpt.PrtMip
4540    LSet PM = PrtMipString

        ' ** Set margins.
4550    PM.xLeftMargin = 1& * 1440&
4560    PM.yTopMargin = 1& * 1440&
4570    PM.xRightMargin = 1& * 1440&
4580    PM.yBotMargin = 1& * 1440&

        ' ** Update property.
4590    LSet PrtMipString = PM
4600    rpt.PrtMip = PrtMipString.strRGB

EXITP:
4610    Set rpt = Nothing
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
4650    End Select
4660    Resume EXITP

End Sub

Public Function AcctReview_Get(blnJan As Boolean, blnFeb As Boolean, blnMar As Boolean, blnApr As Boolean, blnMay As Boolean, blnJun As Boolean, blnJul As Boolean, blnAug As Boolean, blnSep As Boolean, blnOct As Boolean, blnNov As Boolean, blnDec As Boolean)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "AcctReview_Get"

        Dim varMonth As Variant
        Dim blnRetVal As Boolean

4710    blnRetVal = False

4720    varMonth = FormRef("frmRpt_AccountReviews")  ' ** Module Function: modQueryFunctions1.
4730    If IsNull(varMonth) = False Then
4740      Select Case varMonth
          Case "jan"
4750        blnRetVal = blnJan
4760      Case "feb"
4770        blnRetVal = blnFeb
4780      Case "mar"
4790        blnRetVal = blnMar
4800      Case "apr"
4810        blnRetVal = blnApr
4820      Case "may"
4830        blnRetVal = blnMay
4840      Case "jun"
4850        blnRetVal = blnJun
4860      Case "jul"
4870        blnRetVal = blnJul
4880      Case "aug"
4890        blnRetVal = blnAug
4900      Case "sep"
4910        blnRetVal = blnSep
4920      Case "oct"
4930        blnRetVal = blnOct
4940      Case "nov"
4950        blnRetVal = blnNov
4960      Case "dec"
4970        blnRetVal = blnDec
4980      End Select
4990    End If

EXITP:
5000    AcctReview_Get = blnRetVal
5010    Exit Function

ERRH:
5020    blnRetVal = False
5030    Select Case ERR.Number
        Case Else
5040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5050    End Select
5060    Resume EXITP

End Function

Public Function AcctName_Get(strAccountNo As String) As String
' ** Called by:
' **   rptCourtRptNY_... (all NY court reports)
' **     Report_Open()
' ** Was called 'fncAcctShortName'

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "AcctName_Get"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strWhere As String
        Dim strRetVal As String

5110    strRetVal = vbNullString

5120    strWhere = "Accountno = '" & strAccountNo & "'"
        ' ** Set the account name.
5130    Set dbs = CurrentDb
5140    Set rst = dbs.OpenRecordset("Account", dbOpenDynaset)
5150    With rst
5160      .FindFirst strWhere
5170      If .NoMatch = False Then
5180        If gblnLegalName = False Then
5190          strRetVal = rst![shortname]
5200        Else
5210          strRetVal = rst![legalname]
5220        End If
5230      End If
5240      .Close
5250    End With
5260    dbs.Close

EXITP:
5270    Set rst = Nothing
5280    Set dbs = Nothing
5290    AcctName_Get = strRetVal
5300    Exit Function

ERRH:
5310    strRetVal = vbNullString
5320    Select Case ERR.Number
        Case Else
5330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5340    End Select
5350    Resume EXITP

End Function

Public Function fncTransactionDesc(strRecurringItem As Variant, strDescription As Variant, dblRate As Variant, dblDue As Variant, strComment As Variant, Optional varCheckNum As Variant) As String
' ** Called by:
' **   modCourtReportsCA
' **     CABuildCourtReportData()
' **   modCourtReportsFL
' **     FLBuildCourtReportData()
' **   modCourtReportsNY
' **     NYBuildCourtReportData()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "fncTransactionDesc"

        Dim strTmp01 As String

5410    strTmp01 = ""
5420    If Not IsNull(strRecurringItem) Then
5430      strTmp01 = strTmp01 & Trim(strRecurringItem)
5440    End If

5450    If Not IsNull(strDescription) Then
5460      strTmp01 = strTmp01 & " " & CStr(strDescription)
5470    End If

5480    If dblRate > 0 Then
5490      strTmp01 = strTmp01 & " " & Format(dblRate, "#,##0.000%")
5500    End If

5510    If Not IsNull(dblDue) Then
5520      strTmp01 = strTmp01 & "  Due " & Format(dblDue, "mm/dd/yyyy")
5530    End If

5540    If Not IsNull(strComment) Then
5550      strTmp01 = strTmp01 & " " & strComment
5560    End If

5570    Select Case IsMissing(varCheckNum)
        Case True
          ' ** Nothing
5580    Case False
5590      If IsNull(varCheckNum) = False Then
5600        strTmp01 = strTmp01 & "~" & CStr(varCheckNum)
5610      End If
5620    End Select

5630    fncTransactionDesc = strTmp01

EXITP:
5640    Exit Function

ERRH:
5650    Select Case ERR.Number
        Case Else
5660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5670    End Select
5680    Resume EXITP

End Function

Public Function Rpt_Tax_ReceiptIncome() As Boolean
' ** Called by:
' **   frmRpt_TaxIncomeDeductions:
' **     cmdReceiptIncomeExcel_Click()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_Tax_ReceiptIncome"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim strAccountNo As String, strAcctName As String, lngAcctNumID As Long
        Dim intTaxCode As Integer, strTaxType As String, lngTaxCodeID As Long
        Dim strTotDesc As String, lngTotDescID As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

5710    blnRetVal = True

5720    Set dbs = CurrentDb
5730    With dbs

          ' ** Empty tmpReceiptsIncome.
5740      Set qdf = .QueryDefs("qryTaxReporting_12")
5750      qdf.Execute

          ' ** Append qryTaxReporting, with column headers and totals, to tmpReceiptsIncome.
5760      Set qdf = .QueryDefs("qryTaxReporting_41_08")
5770      qdf.Execute

          ' ** Receipts/Income.
5780      Set qdf = .QueryDefs("qryTaxReporting_42")
5790      Set rst = qdf.OpenRecordset
5800      With rst
5810        If .BOF = True And .EOF = True Then
              ' ** No records, so don't continue.
5820        Else
5830          .MoveLast
5840          lngRecs = .RecordCount
5850          .MoveFirst
5860          strAccountNo = vbNullString: strAcctName = vbNullString: lngAcctNumID = 0&
5870          intTaxCode = 0: strTaxType = vbNullString: lngTaxCodeID = 0&
5880          strTotDesc = vbNullString: lngTotDescID = 0&
5890          For lngX = 1& To lngRecs
5900            .Edit
5910            If lngX = 1& Then
5920              ![taxri_first] = True
5930            ElseIf lngX = lngRecs Then
5940              ![taxri_last] = True
5950            End If
5960            If ![Account Num] <> strAccountNo Then
5970              strAccountNo = ![Account Num]
5980              strAcctName = ![Name]
5990              lngAcctNumID = ![taxri_id]
6000              If IsNull(![taxri_par1]) = True Then
6010                ![taxri_par1] = 0&
6020              End If
6030              If IsNull(![taxri_par2]) = True Then
6040                ![taxri_par2] = 0&
6050              End If
6060              If IsNull(![taxri_par3]) = True Then
6070                ![taxri_par3] = 0&
6080              End If
6090              intTaxCode = ![taxcode]
6100              strTaxType = ![Tax Type]
6110              lngTaxCodeID = ![taxri_id]
6120              strTotDesc = ![totdesc]
6130              lngTotDescID = ![taxri_id]
6140            Else
6150              If IsNull(![taxcode]) = True Then
                    ' ** Account total line.
6160                ![taxri_par1] = lngAcctNumID
6170                ![taxri_par2] = lngTaxCodeID
6180                ![taxri_par3] = lngTotDescID
6190                ![Account Num] = Null
6200                ![Name] = Null
6210                ![description] = "        Total"
6220              Else
6230                If ![taxcode] <> intTaxCode Then
6240                  intTaxCode = ![taxcode]
6250                  strTaxType = ![Tax Type]
6260                  lngTaxCodeID = ![taxri_id]
6270                  strTotDesc = ![totdesc]
6280                  lngTotDescID = ![taxri_id]
6290                  ![Account Num] = Null
6300                  ![Name] = Null
6310                  If IsNull(![taxri_par1]) = True Then
6320                    ![taxri_par1] = lngAcctNumID
6330                  End If
6340                Else
6350                  If IsNull(![totdesc]) = True Then
                        ' ** Tax Type total line.
6360                    ![taxri_par1] = lngAcctNumID
6370                    ![taxri_par2] = lngTaxCodeID
6380                    ![taxri_par3] = lngTotDescID
6390                    ![Account Num] = Null
6400                    ![Name] = Null
6410                    ![Tax Type] = Null
6420                    ![description] = "    " & strTaxType & " Total"
6430                  Else
6440                    If ![totdesc] <> strTotDesc Then
6450                      strTotDesc = ![totdesc]
6460                      lngTotDescID = ![taxri_id]
6470                      ![taxri_par1] = lngAcctNumID
6480                      ![taxri_par2] = lngTaxCodeID
6490                      ![Account Num] = Null
6500                      ![Name] = Null
6510                      ![Tax Type] = Null
6520                      If IsNull(![taxri_par1]) = True Then
6530                        ![taxri_par1] = lngAcctNumID
6540                      End If
6550                    Else
6560                      If IsNull(![description]) = True Then
                            ' ** totdesc total Line.
6570                        ![taxri_par1] = lngAcctNumID
6580                        ![taxri_par2] = lngTaxCodeID
6590                        ![taxri_par3] = lngTotDescID
6600                        ![Account Num] = Null
6610                        ![Name] = Null
6620                        ![Tax Type] = Null
6630                        ![Journal Type] = Null
6640                      Else
6650                        ![Account Num] = Null
6660                        ![Name] = Null
6670                        ![taxri_par1] = lngAcctNumID
6680                        ![Tax Type] = Null
6690                        If IsNull(![taxri_par2]) = True Then
6700                          ![taxri_par2] = lngTaxCodeID
6710                        End If
6720                        If IsNull(![taxri_par3]) = True Then
6730                          ![taxri_par3] = lngTotDescID
6740                        End If
6750                      End If
6760                    End If
6770                  End If
6780                End If
6790              End If
6800            End If
6810            .Update
6820            If lngX < lngRecs Then .MoveNext
6830          Next
6840        End If
6850        .Close
6860      End With

          ' ** Append report title to tmpTaxDisbursementsDeductions.
6870      Set qdf = .QueryDefs("qryTaxReporting_43_02")
6880      qdf.Execute

          ' ** Append report period to tmpTaxDisbursementsDeductions.
6890      Set qdf = .QueryDefs("qryTaxReporting_43_04")
6900      With qdf.Parameters
6910        ![datbeg] = Format(Forms("frmRpt_TaxIncomeDeductions").DateStart, "mm/dd/yyyy")
6920        ![datEnd] = Format(Forms("frmRpt_TaxIncomeDeductions").DateEnd, "mm/dd/yyyy")
6930      End With
6940      qdf.Execute

6950    End With

EXITP:
6960    Set rst = Nothing
6970    Set qdf = Nothing
6980    Set dbs = Nothing
6990    Rpt_Tax_ReceiptIncome = blnRetVal
7000    Exit Function

ERRH:
7010    blnRetVal = False
7020    Select Case ERR.Number
        Case Else
7030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7040    End Select
7050    Resume EXITP

End Function

Public Function Rpt_Tax_DisbursementDeduction() As Boolean
' ** Called by:
' **   frmRpt_TaxIncomeDeductions:
' **     cmdDisbursementsDeductionsExcel_Click()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_Tax_DisbursementDeduction"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim strAccountNo As String, strAcctName As String, lngAcctNumID As Long
        Dim intTaxCode As Integer, strTaxType As String, lngTaxCodeID As Long
        Dim strTotDesc As String, lngTotDescID As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

7110    blnRetVal = True

7120    Set dbs = CurrentDb
7130    With dbs

          ' ** Empty tmpTaxDisbursementsDeductions.
7140      Set qdf = .QueryDefs("qryTaxReporting_13")
7150      qdf.Execute

          ' ** Append qryTaxReporting, with column headers and totals, to tmpTaxDisbursementsDeductions.
7160      Set qdf = .QueryDefs("qryTaxReporting_46_08")
7170      qdf.Execute

          ' ** Disbursements/Deductions.
7180      Set qdf = .QueryDefs("qryTaxReporting_47")
7190      Set rst = qdf.OpenRecordset
7200      With rst
7210        If .BOF = True And .EOF = True Then
              ' ** No records, so don't continue.
7220        Else
7230          .MoveLast
7240          lngRecs = .RecordCount
7250          .MoveFirst
7260          strAccountNo = vbNullString: strAcctName = vbNullString: lngAcctNumID = 0&
7270          intTaxCode = 0: strTaxType = vbNullString: lngTaxCodeID = 0&
7280          strTotDesc = vbNullString: lngTotDescID = 0&
7290          For lngX = 1& To lngRecs
7300            .Edit
7310            If lngX = 1& Then
7320              ![taxdd_first] = True
7330            ElseIf lngX = lngRecs Then
7340              ![taxdd_last] = True
7350            End If
7360            If ![Account Num] <> strAccountNo Then
7370              strAccountNo = ![Account Num]
7380              strAcctName = ![Name]
7390              lngAcctNumID = ![taxdd_id]
7400              If IsNull(![taxdd_par1]) = True Then
7410                ![taxdd_par1] = 0&
7420              End If
7430              If IsNull(![taxdd_par2]) = True Then
7440                ![taxdd_par2] = 0&
7450              End If
7460              If IsNull(![taxdd_par3]) = True Then
7470                ![taxdd_par3] = 0&
7480              End If
7490              intTaxCode = ![taxcode]
7500              strTaxType = ![Tax Type]
7510              lngTaxCodeID = ![taxdd_id]
7520              strTotDesc = ![totdesc]
7530              lngTotDescID = ![taxdd_id]
7540            Else
7550              If IsNull(![taxcode]) = True Then
                    ' ** Account total line.
7560                ![taxdd_par1] = lngAcctNumID
7570                ![taxdd_par2] = lngTaxCodeID
7580                ![taxdd_par3] = lngTotDescID
7590                ![Account Num] = Null
7600                ![Name] = Null
7610                ![description] = "        Total"
7620              Else
7630                If ![taxcode] <> intTaxCode Then
7640                  intTaxCode = ![taxcode]
7650                  strTaxType = ![Tax Type]
7660                  lngTaxCodeID = ![taxdd_id]
7670                  strTotDesc = ![totdesc]
7680                  lngTotDescID = ![taxdd_id]
7690                  ![Account Num] = Null
7700                  ![Name] = Null
7710                  If IsNull(![taxdd_par1]) = True Then
7720                    ![taxdd_par1] = lngAcctNumID
7730                  End If
7740                Else
7750                  If IsNull(![totdesc]) = True Then
                        ' ** Tax Type total line.
7760                    ![taxdd_par1] = lngAcctNumID
7770                    ![taxdd_par2] = lngTaxCodeID
7780                    ![taxdd_par3] = lngTotDescID
7790                    ![Account Num] = Null
7800                    ![Name] = Null
7810                    ![Tax Type] = Null
7820                    ![description] = "    " & strTaxType & " Total"
7830                  Else
7840                    If ![totdesc] <> strTotDesc Then
7850                      strTotDesc = ![totdesc]
7860                      lngTotDescID = ![taxdd_id]
7870                      ![taxdd_par1] = lngAcctNumID
7880                      ![taxdd_par2] = lngTaxCodeID
7890                      ![Account Num] = Null
7900                      ![Name] = Null
7910                      ![Tax Type] = Null
7920                      If IsNull(![taxdd_par1]) = True Then
7930                        ![taxdd_par1] = lngAcctNumID
7940                      End If
7950                    Else
7960                      If IsNull(![description]) = True Then
                            ' ** totdesc total Line.
7970                        ![taxdd_par1] = lngAcctNumID
7980                        ![taxdd_par2] = lngTaxCodeID
7990                        ![taxdd_par3] = lngTotDescID
8000                        ![Account Num] = Null
8010                        ![Name] = Null
8020                        ![Tax Type] = Null
8030                        ![Journal Type] = Null
8040                      Else
8050                        ![Account Num] = Null
8060                        ![Name] = Null
8070                        ![taxdd_par1] = lngAcctNumID
8080                        ![Tax Type] = Null
8090                        If IsNull(![taxdd_par2]) = True Then
8100                          ![taxdd_par2] = lngTaxCodeID
8110                        End If
8120                        If IsNull(![taxdd_par3]) = True Then
8130                          ![taxdd_par3] = lngTotDescID
8140                        End If
8150                      End If
8160                    End If
8170                  End If
8180                End If
8190              End If
8200            End If
8210            .Update
8220            If lngX < lngRecs Then .MoveNext
8230          Next
8240        End If
8250        .Close
8260      End With

          ' ** Append report title to tmpTaxDisbursementsDeductions.
8270      Set qdf = .QueryDefs("qryTaxReporting_48_02")
8280      qdf.Execute

          ' ** Append report period to tmpTaxDisbursementsDeductions.
8290      Set qdf = .QueryDefs("qryTaxReporting_48_04")
8300      With qdf.Parameters
8310        ![datbeg] = Format(Forms("frmRpt_TaxIncomeDeductions").DateStart, "mm/dd/yyyy")
8320        ![datEnd] = Format(Forms("frmRpt_TaxIncomeDeductions").DateEnd, "mm/dd/yyyy")
8330      End With
8340      qdf.Execute

8350    End With

EXITP:
8360    Set rst = Nothing
8370    Set qdf = Nothing
8380    Set dbs = Nothing
8390    Rpt_Tax_DisbursementDeduction = blnRetVal
8400    Exit Function

ERRH:
8410    blnRetVal = False
8420    Select Case ERR.Number
        Case Else
8430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8440    End Select
8450    Resume EXITP

End Function

Public Function CleanInputBox(strValue As String) As String
' ** Cleans up user data entered manually into an InputBox.

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "CleanInputBox"

        Dim intPos01 As Integer
        Dim strRetVal As String

8510    strRetVal = "0"

8520    If strValue <> vbNullString Then
          ' ** Remove dollar sign.
8530      intPos01 = InStr(strValue, "$")
8540      If intPos01 > 0 Then  ' ** $2,121.97
8550        If intPos01 = 1 Then
8560          strValue = Trim(Mid(strValue, 2))
8570        Else  ' ** ($2,121.97)
8580          strValue = Trim(Left(strValue, (intPos01 - 1)) & Mid(strValue, (intPos01 + 1)))
8590        End If
8600      End If
          ' ** Remove opening paren, and replace with minus sign.
8610      intPos01 = InStr(strValue, "(")
8620      If intPos01 > 0 Then
8630        If intPos01 = 1 Then
8640          strValue = "-" & Trim(Mid(strValue, 2))
8650        Else
8660          strValue = Trim(Left(strValue, (intPos01 - 1)) & "-" & Mid(strValue, (intPos01 + 1)))
8670        End If
8680      End If
          ' ** Remove closing paren.
8690      intPos01 = InStr(strValue, ")")
8700      If intPos01 > 0 Then
8710        If intPos01 = Len(strValue) Then
8720          strValue = Trim(Left(strValue, (Len(strValue) - 1)))
8730        Else
8740          strValue = Trim(Left(strValue, (intPos01 - 1)) & Mid(strValue, (intPos01 + 1)))
8750        End If
8760      End If
          ' ** Remove commas.
8770      intPos01 = InStr(strValue, ",")
8780      If intPos01 > 0 Then
8790        Do While intPos01 > 0
8800          strValue = Trim(Left(strValue, (intPos01 - 1)) & Mid(strValue, (intPos01 + 1)))
8810          intPos01 = InStr(strValue, ",")
8820        Loop
8830      End If
8840      If IsNumeric(strValue) = False Then strValue = "0"
8850      strValue = CStr(Val(strValue))
8860      If strValue = vbNullString Or strValue = "-" Then strValue = "0"
8870      strRetVal = strValue
8880    End If

EXITP:
8890    CleanInputBox = strRetVal
8900    Exit Function

ERRH:
8910    strRetVal = "0"
8920    Select Case ERR.Number
        Case Else
8930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8940    End Select
8950    Resume EXITP

End Function

Public Function Fnc_ckgDisplay_opt01_AccountNo_AfterUpdate() As Integer

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "Fnc_ckgDisplay_opt01_AccountNo_AfterUpdate"

        Dim Cancel As Integer

9010    Cancel = 0
9020    Form_frmRpt_Checks.ckgDisplay_opt01_AccountNo_AfterUpdate  ' ** Form Procedure: frmRpt_Checks.

EXITP:
9030    Fnc_ckgDisplay_opt01_AccountNo_AfterUpdate = Cancel
9040    Exit Function

ERRH:
9050    Cancel = -1
9060    Select Case ERR.Number
        Case Else
9070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9080    End Select
9090    Resume EXITP

End Function

Public Function Fnc_ckgDisplay_opt02_ShortName_AfterUpdate() As Integer

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "Fnc_ckgDisplay_opt02_ShortName_AfterUpdate"

        Dim Cancel As Integer

9110    Cancel = 0
9120    Form_frmRpt_Checks.ckgDisplay_opt02_ShortName_AfterUpdate  ' ** Form Procedure: frmRpt_Checks.

EXITP:
9130    Fnc_ckgDisplay_opt02_ShortName_AfterUpdate = Cancel
9140    Exit Function

ERRH:
9150    Cancel = -1
9160    Select Case ERR.Number
        Case Else
9170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9180    End Select
9190    Resume EXITP

End Function

Public Function Fnc_ckgDisplay_opt05_Payee_AfterUpdate() As Integer

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "Fnc_ckgDisplay_opt05_Payee_AfterUpdate"

        Dim Cancel As Integer

9210    Cancel = 0
9220    Form_frmRpt_Checks.ckgDisplay_opt05_Payee_AfterUpdate  ' ** Form Procedure: frmRpt_Checks.

EXITP:
9230    Fnc_ckgDisplay_opt05_Payee_AfterUpdate = Cancel
9240    Exit Function

ERRH:
9250    Cancel = -1
9260    Select Case ERR.Number
        Case Else
9270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9280    End Select
9290    Resume EXITP

End Function

Public Function Fnc_ckgDisplay_opt06_BankName_AfterUpdate() As Integer

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "Fnc_ckgDisplay_opt06_BankName_AfterUpdate"

        Dim Cancel As Integer

9310    Cancel = 0
9320    Form_frmRpt_Checks.ckgDisplay_opt06_BankName_AfterUpdate  ' ** Form Procedure: frmRpt_Checks.

EXITP:
9330    Fnc_ckgDisplay_opt06_BankName_AfterUpdate = Cancel
9340    Exit Function

ERRH:
9350    Cancel = -1
9360    Select Case ERR.Number
        Case Else
9370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9380    End Select
9390    Resume EXITP

End Function

Public Function Fnc_ckgDisplay_opt07_BankAcctNum_AfterUpdate() As Integer

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "Fnc_ckgDisplay_opt07_BankAcctNum_AfterUpdate"

        Dim Cancel As Integer

9410    Cancel = 0
9420    Form_frmRpt_Checks.ckgDisplay_opt07_BankAcctNum_AfterUpdate  ' ** Form Procedure: frmRpt_Checks.

EXITP:
9430    Fnc_ckgDisplay_opt07_BankAcctNum_AfterUpdate = Cancel
9440    Exit Function

ERRH:
9450    Cancel = -1
9460    Select Case ERR.Number
        Case Else
9470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9480    End Select
9490    Resume EXITP

End Function

Public Function Rpt_GroupLevel_Set(rpt As Access.Report, strSortOrd As String, blnDesc As Boolean) As Boolean
' ** Called by:
' **   rptChecks_Blank
' **     Report_Open()
' **   rptChecks_Preprinted
' **     Report_Open()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_GroupLevel_Set"

        Dim grp1 As Access.GroupLevel, grp2 As Access.GroupLevel
        Dim lngLvls As Long, arr_varLvl() As Variant
        Dim strControlSource As String
        Dim blnNoChange As Boolean
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varLvl().
        Const L_ELEMS As Integer = 7  ' ** Array's first-element UBound().
        Const L_LEVEL  As Integer = 0
        Const L_SOURCE As Integer = 1
        Const L_SORT   As Integer = 2
        Const L_HEAD   As Integer = 3
        Const L_FOOT   As Integer = 4
        Const L_GRPON  As Integer = 5
        Const L_INTRVL As Integer = 6
        Const L_KEEP   As Integer = 7

9510    blnRetVal = True
9520    If IsNothing(rpt) = False Then  ' ** Module Function: modUtilities.

9530      lngLvls = 0&
9540      ReDim arr_varLbl(L_ELEMS, 0)
9550      blnNoChange = False

9560      With rpt

            ' ** First, find out what the current group levels are.
9570        For lngX = 0& To 9&  ' ** Max group levels.
9580  On Error Resume Next
9590          strControlSource = .GroupLevel(lngX).ControlSource
9600          If ERR.Number = 0 Then
9610  On Error GoTo ERRH
9620            lngLvls = lngLvls + 1&
9630            lngE = lngLvls - 1&
9640            ReDim Preserve arr_varLvl(L_ELEMS, lngE)
9650            arr_varLvl(L_LEVEL, lngE) = lngX
9660            arr_varLvl(L_SOURCE, lngE) = strControlSource
9670            arr_varLvl(L_SORT, lngE) = .GroupLevel(lngX).sortOrder
9680            arr_varLvl(L_HEAD, lngE) = .GroupLevel(lngX).GroupHeader
9690            arr_varLvl(L_FOOT, lngE) = .GroupLevel(lngX).GroupFooter
9700            arr_varLvl(L_GRPON, lngE) = .GroupLevel(lngX).GroupOn
9710            arr_varLvl(L_INTRVL, lngE) = .GroupLevel(lngX).GroupInterval
9720            arr_varLvl(L_KEEP, lngE) = .GroupLevel(lngX).KeepTogether
9730          Else
9740  On Error GoTo ERRH
9750          End If
9760        Next  ' ** lngX.

9770        If lngLvls > 0& Then
9780          If arr_varLvl(L_HEAD, 0) = False And arr_varLvl(L_FOOT, 0) = False Then
9790            Set grp1 = .GroupLevel(arr_varLvl(L_LEVEL, 0))  ' ** First in array, regardless of its level.
9800            If lngLvls > 1& Then
9810              Set grp2 = .GroupLevel(arr_varLvl(L_LEVEL, 1))  ' ** Second in array, regardless of its level.

9820              Select Case strSortOrd
                  Case "Account_Number"       ' ** accountno
9830                If arr_varLvl(L_SOURCE, 0) <> "alphasort" Then
9840                  grp1.ControlSource = "alphasort"
9850                End If
9860                If arr_varLvl(L_SOURCE, 1) <> "RecurringItem" Then
9870                  grp2.ControlSource = "RecurringItem"  ' ** Payee.
9880                End If
9890              Case "Short_Name"           ' ** shortname
9900                If arr_varLvl(L_SOURCE, 0) <> "shortname" Then
9910                  grp1.ControlSource = "shortname"
9920                End If
9930                If arr_varLvl(L_SOURCE, 1) <> "RecurringItem" Then
9940                  grp2.ControlSource = "RecurringItem"  ' ** Payee.
9950                End If
9960              Case "Check_Count"          ' ** {n/a}
                    ' ** Can't sort by this.
9970                blnNoChange = True
9980              Case "Last_Check_Number"    ' ** {n/a}
                    ' ** Can't sort by this.
9990                blnNoChange = True
10000             Case "Payee"                ' ** RecurringItem
10010               If arr_varLvl(L_SOURCE, 0) <> "RecurringItem" Then
10020                 grp1.ControlSource = "RecurringItem"  ' ** Payee.
10030               End If
10040               If arr_varLvl(L_SOURCE, 1) <> "shortname" Then
10050                 grp2.ControlSource = "shortname"
10060               End If
10070             Case "Bank_Name"            ' ** Bank_Name
10080               If arr_varLvl(L_SOURCE, 0) <> "Bank_Name" Then
10090                 grp1.ControlSource = "Bank_Name"
10100               End If
10110               If arr_varLvl(L_SOURCE, 1) <> "RecurringItem" Then
10120                 grp2.ControlSource = "RecurringItem"  ' ** Payee.
10130               End If
10140             Case "Bank_Account_Number"  ' ** Bank_AccountNumber
10150               If arr_varLvl(L_SOURCE, 0) <> "Bank_AccountNumber" Then
10160                 grp1.ControlSource = "Bank_AccountNumber"
10170               End If
10180               If arr_varLvl(L_SOURCE, 1) <> "RecurringItem" Then
10190                 grp2.ControlSource = "RecurringItem"  ' ** Payee.
10200               End If
10210             End Select  ' ** strSortOrd.

10220             If blnNoChange = False Then
10230               With grp1
10240                 .sortOrder = blnDesc  ' ** False: Ascending; True: Descending.
10250                 .GroupOn = 0  ' ** Each Value.
10260                 .GroupInterval = 1  ' ** Each Value.
10270                 .KeepTogether = 0  ' ** 0: No; 1: Whole Group; 2: With First Detail.
10280               End With
10290               With grp2
10300                 .sortOrder = blnDesc  ' ** False: Ascending; True: Descending.
10310                 .GroupOn = 0  ' ** Each Value.
10320                 .GroupInterval = 1  ' ** Each Value.
10330                 .KeepTogether = 0  ' ** 0: No; 1: Whole Group; 2: With First Detail.
10340               End With
10350             End If

10360           Else
                  ' ** For now, the reports already have 2 levels.
10370             blnNoChange = True
10380           End If
10390         Else
                ' ** Maybe later.
10400           blnNoChange = True
10410         End If  ' ** Header, Footer.
10420       Else
              ' ** Would have to open in Design View to add.
10430         blnNoChange = True
10440       End If  ' ** lngLvls.

10450     End With  ' ** rpt

10460   Else
10470     blnRetVal = False
10480   End If

10490   If blnNoChange = True Then blnRetVal = False

EXITP:
10500   Set grp1 = Nothing
10510   Set grp2 = Nothing
10520   Rpt_GroupLevel_Set = blnRetVal
10530   Exit Function

ERRH:
10540   blnRetVal = False
10550   Select Case ERR.Number
        Case Else
10560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10570   End Select
10580   Resume EXITP

End Function

Public Function Rpt_Import() As Boolean

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_Import"

        'Dim wrk As DAO.Workspace, dbs As DAO.Database, rpt As Access.Report
        Dim strMDB As String
        Dim lngRpts As Long, arr_varRpt() As Variant
        Dim lngHits As Long
        Dim blnSkip As Boolean
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varRpt().
        Const R_ELEMS As Integer = 0  ' ** Array's first-element UBound().
        Const R_RNAM As Integer = 0

10610 On Error GoTo 0

10620   blnRetVal = True

10630   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
10640   DoEvents

10650   DBEngine.SystemDB = "C:\Program Files\Delta Data\Trust Accountant\Database\TrustSec.mdw"

10660   strMDB = "C:\Program Files\Delta Data\Trust Accountant\Trust_bak99_03.mdb"

10670   blnSkip = False
10680   If blnSkip = False Then

10690     lngRpts = 0&
10700     ReDim arr_varRpt(R_ELEMS, 0)

          'Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)
          'Set dbs = wrk.OpenDatabase(strMDB, False, True)  ' ** {pathfile}, {exclusive}, {read-only}

10710     lngRpts = lngRpts + 1&
10720     lngE = lngRpts - 1&
10730     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
10740     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Dividend_01_Sub_02"
10750     lngRpts = lngRpts + 1&
10760     lngE = lngRpts - 1&
10770     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
10780     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Interest_01"
10790     lngRpts = lngRpts + 1&
10800     lngE = lngRpts - 1&
10810     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
10820     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Interest_01_Sub_01"
10830     lngRpts = lngRpts + 1&
10840     lngE = lngRpts - 1&
10850     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
10860     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Interest_01_Sub_02"
10870     lngRpts = lngRpts + 1&
10880     lngE = lngRpts - 1&
10890     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
10900     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Interest_01_Sub_03"
10910     lngRpts = lngRpts + 1&
10920     lngE = lngRpts - 1&
10930     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
10940     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Interest_01_Sub_04"
10950     lngRpts = lngRpts + 1&
10960     lngE = lngRpts - 1&
10970     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
10980     arr_varRpt(R_RNAM, lngE) = "'zz_rptGlenwood_Interest_01_Sub_05"
10990     lngRpts = lngRpts + 1&
11000     lngE = lngRpts - 1&
11010     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
11020     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Deposit_01"
11030     lngRpts = lngRpts + 1&
11040     lngE = lngRpts - 1&
11050     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
11060     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Deposit_01_Sub_01"
11070     lngRpts = lngRpts + 1&
11080     lngE = lngRpts - 1&
11090     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
11100     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Deposit_01_Sub_02"
11110     lngRpts = lngRpts + 1&
11120     lngE = lngRpts - 1&
11130     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
11140     arr_varRpt(R_RNAM, lngE) = "'zz_rptGlenwood_Deposit_01_Sub_03"
11150     lngRpts = lngRpts + 1&
11160     lngE = lngRpts - 1&
11170     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
11180     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Deposit_01_Sub_04"
11190     lngRpts = lngRpts + 1&
11200     lngE = lngRpts - 1&
11210     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
11220     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Purchase_01_Sub_01"
11230     lngRpts = lngRpts + 1&
11240     lngE = lngRpts - 1&
11250     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
11260     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Purchase_01_Sub_02"
11270     lngRpts = lngRpts + 1&
11280     lngE = lngRpts - 1&
11290     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
11300     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Purchase_01_Sub_03"
11310     lngRpts = lngRpts + 1&
11320     lngE = lngRpts - 1&
11330     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
11340     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Purchase_01_Sub_04"
11350     lngRpts = lngRpts + 1&
11360     lngE = lngRpts - 1&
11370     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
11380     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Purchase_01_Sub_05"
11390     lngRpts = lngRpts + 1&
11400     lngE = lngRpts - 1&
11410     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
11420     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Purchase_01_Sub_06"
11430     lngRpts = lngRpts + 1&
11440     lngE = lngRpts - 1&
11450     ReDim Preserve arr_varRpt(R_ELEMS, lngE)
11460     arr_varRpt(R_RNAM, lngE) = "zz_rptGlenwood_Purchase_01"

11470     Debug.Print "'RPTS: " & CStr(lngRpts)
11480     DoEvents

11490     lngHits = 0&
11500     If lngRpts > 0& Then
11510       For lngX = 0& To (lngRpts - 1&)
11520         DoCmd.TransferDatabase acImport, "Microsoft Access", strMDB, acReport, arr_varRpt(R_RNAM, lngX), arr_varRpt(R_RNAM, lngX)
11530         lngHits = lngHits + 1&
11540       Next  ' ** lngX.
11550     End If  ' ** lngRpts.

11560   End If  ' ** blnSkip.

11570   Debug.Print "'RPTS COPIED: " & CStr(lngHits)
11580   DoEvents

11590   Beep
11600   Debug.Print "'DONE!"

EXITP:
        'Set rpt = Nothing
        'Set dbs = Nothing
        'Set wrk = Nothing
11610   Rpt_Import = blnRetVal
11620   Exit Function

ERRH:
11630   blnRetVal = False
11640   Select Case ERR.Number
        Case Else
11650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
11660   End Select
11670   Resume EXITP

End Function

Public Function Rpt_RibbonAdd() As Boolean

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_RibbonAdd"

        Dim dbs As DAO.Database, rst As DAO.Recordset, cntr As DAO.Container, doc As DAO.Document, rpt As Access.Report
        Dim lngRpts As Long, arr_varRpt() As Variant
        Dim blnChanged As Boolean, lngChanges As Long
        Dim varTmp00 As Variant
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varRpt().
        Const R_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const R_NAM As Integer = 0
        Const R_MOD As Integer = 1
        Const R_CID As Integer = 2

11710   blnRetVal = True

11720   lngRpts = 0&
11730   ReDim arr_varRpt(R_ELEMS, 0)

11740   Set dbs = CurrentDb
11750   With dbs

11760     Set cntr = .Containers("Reports")
11770     With cntr
11780       For Each doc In .Documents
11790         lngRpts = lngRpts + 1&
11800         lngE = lngRpts - 1&
11810         ReDim Preserve arr_varRpt(R_ELEMS, lngE)
11820         arr_varRpt(R_NAM, lngE) = doc.Name
11830         arr_varRpt(R_MOD, lngE) = ("Report_" & doc.Name)
11840         varTmp00 = DLookup("[vbcom_id]", "tblVBComponent", "[vbcom_name] = '" & ("Report_" & doc.Name) & "'")
11850         If IsNull(varTmp00) = False Then
11860           arr_varRpt(R_CID, lngE) = CLng(varTmp00)
11870         Else
11880           arr_varRpt(R_CID, lngE) = CLng(0)
11890           Debug.Print "'NOT FOUND! " & doc.Name
11900         End If
11910       Next
11920     End With

11930     lngChanges = 0&
11940     For lngX = 0& To (lngRpts - 1&)
11950       blnChanged = False
11960       DoCmd.OpenReport arr_varRpt(R_NAM, lngX), acViewDesign, , , acHidden
11970       Set rpt = Reports(arr_varRpt(R_NAM, lngX))
11980       With rpt
11990         If IsNull(.Tag) = True Then
12000           If IsNull(.RibbonName) = True Then
12010             .RibbonName = "TAReports"
12020             blnChanged = True
12030             lngChanges = lngChanges + 1&
12040           Else
12050             If Trim(.RibbonName) = vbNullString Then
12060               .RibbonName = "TAReports"
12070               blnChanged = True
12080               lngChanges = lngChanges + 1&
12090             Else
12100               If .RibbonName <> "TAReports" Then
12110                 .RibbonName = "TAReports"
12120                 blnChanged = True
12130                 lngChanges = lngChanges + 1&
12140               End If
12150             End If
12160           End If
12170         Else
12180           If Trim(.Tag) = vbNullString Then
12190             If IsNull(.RibbonName) = True Then
12200               .RibbonName = "TAReports"
12210               blnChanged = True
12220               lngChanges = lngChanges + 1&
12230             Else
12240               If Trim(.RibbonName) = vbNullString Then
12250                 .RibbonName = "TAReports"
12260                 blnChanged = True
12270                 lngChanges = lngChanges + 1&
12280               Else
12290                 If .RibbonName <> "TAReports" Then
12300                   .RibbonName = "TAReports"
12310                   blnChanged = True
12320                   lngChanges = lngChanges + 1&
12330                 End If
12340               End If
12350             End If
12360           Else
12370             If InStr(.Tag, "Is Sub") = 0 Then
12380               If IsNull(.RibbonName) = True Then
12390                 .RibbonName = "TAReports"
12400                 blnChanged = True
12410                 lngChanges = lngChanges + 1&
12420               Else
12430                 If Trim(.RibbonName) = vbNullString Then
12440                   .RibbonName = "TAReports"
12450                   blnChanged = True
12460                   lngChanges = lngChanges + 1&
12470                 Else
12480                   If .RibbonName <> "TAReports" Then
12490                     .RibbonName = "TAReports"
12500                     blnChanged = True
12510                     lngChanges = lngChanges + 1&
12520                   End If
12530                 End If
12540               End If
12550             End If
12560           End If
12570         End If
12580       End With
12590       If blnChanged = True Then
12600         DoCmd.Close acReport, arr_varRpt(R_NAM, lngX), acSaveYes
12610       Else
12620         DoCmd.Close acReport, arr_varRpt(R_NAM, lngX), acSaveNo
12630       End If
12640       Set rpt = Nothing
12650     Next

12660     .Close
12670   End With  ' ** dbs.

12680   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

12690   Debug.Print "'CHANGED: " & CStr(lngChanges)
12700   Debug.Print "'DONE!  " & THIS_PROC & "()" & "  " & blnRetVal

12710   Beep

EXITP:
12720   Set rpt = Nothing
12730   Set doc = Nothing
12740   Set cntr = Nothing
12750   Set rst = Nothing
12760   Set dbs = Nothing
12770   Rpt_RibbonAdd = blnRetVal
12780   Exit Function

ERRH:
12790   blnRetVal = False
12800   Select Case ERR.Number
        Case Else
12810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12820   End Select
12830   Resume EXITP

End Function

Public Function Rpt_SearchZOrder()
' ** This will attempt to pile the controls in their proper z-order, back to front!

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_SearchZOrder"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form
        Dim lngCtls As Long, arr_varCtl As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varCtl().
        'Const C_DID  As Integer = 0
        'Const C_DNAM As Integer = 1
        'Const C_FID  As Integer = 2
        'Const C_FNAM As Integer = 3
        'Const C_CID  As Integer = 4
        Const C_CNAM As Integer = 5
        'Const C_TYP  As Integer = 6
        'Const C_IS_Z As Integer = 7
        'Const C_ORD1 As Integer = 8
        'Const C_ORD2 As Integer = 9

12910   blnRetVal = True

12920   Set dbs = CurrentDb
12930   With dbs
          ' ** tblForm_Control, just frmReportList z-order controls, sorted, by specified CurrentAppName().
12940     Set qdf = .QueryDefs("zz_qry_Report_List_03")
12950     Set rst = qdf.OpenRecordset
12960     With rst
12970       .MoveLast
12980       lngCtls = .RecordCount
12990       .MoveFirst
13000       arr_varCtl = .GetRows(lngCtls)
            ' *************************************************
            ' ** Array: arr_varCtl():
            ' **
            ' **   Field  Element  Name            Constant
            ' **   =====  =======  ==============  ==========
            ' **     1       0     dbs_id          C_DID
            ' **     2       1     dbs_name        C_DNAM
            ' **     3       2     frm_id          C_FID
            ' **     4       3     frm_name        C_FNAM
            ' **     5       4     ctl_id          C_CID
            ' **     6       5     ctl_name        C_CNAM
            ' **     7       6     ctltype_type    C_TYP
            ' **     8       7     ZCtlx           C_IS_Z
            ' **     9       8     ZOrd1           C_ORD1
            ' **    10       9     ZOrd2           C_ORD2
            ' **
            ' *************************************************
13010       .Close
13020     End With  ' ** rst.
13030     Set rst = Nothing
13040     Set qdf = Nothing
13050     .Close
13060   End With  ' ** dbs.
13070   Set dbs = Nothing

13080   If IsLoaded("frmReportList", acForm) = False Then  ' ** Module Function: modFileUtilities.
13090     DoCmd.OpenForm "frmReportList", acDesign
13100   End If

        ' ** frmReportList should already be open in Design View.
13110   Set frm = Forms(0)
13120   With frm

          ' ** Get us back on the front side.
13130     DoCmd.SelectObject acForm, .Name, False

13140     For lngX = 0& To (lngCtls - 1&)
13150       .Controls(arr_varCtl(C_CNAM, lngX)).InSelection = True
13160       DoCmd.RunCommand acCmdBringToFront
13170       DoEvents
13180       DoCmd.RunCommand acCmdBringToFront
13190       DoEvents
13200       .Controls(arr_varCtl(C_CNAM, lngX)).InSelection = False
13210     Next  ' ** lngX.

13220   End With  ' ** frm
13230   Set frm = Nothing

        'REPORT LIST Z-ORDER:
        '====================
        'IN THIS ORDER, BACK TO FRONT!
        'SO, EITHER START AT THE END, AND
        'SEND TO BACK LAST TO FIRST, OR
        'START AT TOP, AND BRING TO FRONT,
        'FIRST TO LAST!

        '.Controls(strTmp01).InSelection = True
        'DoCmd.RunCommand acCmdSendToBack
        'DoEvents
        '.Controls(strTmp01).InSelection = False

        '.Controls(strTmp01).InSelection = True
        'DoCmd.RunCommand acCmdBringToFront
        'DoEvents
        '.Controls(strTmp01).InSelection = False

        'FILE FOLDER PIECES:
        'Location_box2
        'Location_box
        'Location_hline01
        'Location_hline02
        'Location_hline03
        'Location_vline01
        'Location_vline02
        'Location_vline03
        'Location_vline04
        'Location_hline05 (etched above btn)
        'Location_hline06
        'Location_lbl
        'cmdLocationGo

        'COLOR BOXES:
        'cmbMenu1_box1
        'cmbMenu1_box2
        'cmbMenu2_box1
        'cmbMenu2_box2
        'cmbMenu3_box1
        'cmbMenu3_box2

        'CONNECTING LINES:
        'MENU1:
        'Location_arrow1a (down line)
        'Location_arrow1c (right arrow)
        'BTN1:
        'Location_arrow1d (down line)
        'Location_arrow1e (down arrow)
        'MENU2:
        'Location_arrow2a (down line)
        'Location_arrow2c (right arrow)
        'BTN2:
        'Location_arrow2d (down line)
        'Location_arrow2e (down arrow)
        'MENU3:
        'Location_arrow3a (down line)
        'Location_arrow3c (right arrow)
        'VIA1:
        'Location_arrow4a (down line)
        'Location_arrow4c (right arrow)
        'VIA2:
        'Location_arrow5a (down line)
        'Location_arrow5c (right arrow)

        'MISC LABELS:
        'ArchivedTrans_lbl1
        'ArchivedTrans_lbl2
        'AssetHistory_lbl
        'AssetPricing_lbl
        'Pricing_lbl
        'TransactionAudit_lbl

        'COMBO BOXES:
        'VIA1:
        'cmbViaForm1/cmbViaForm1_lbl
        'cmbViaButton1_hline01 -
        ' cmbViaButton1_hline15
        'cmbViaButton1_lbl_dim_hi
        'cmbViaButton1/cmbViaButton1_lbl
        'VIA2:
        'cmbViaButton2/cmbViaButton2_lbl
        'cmbViaButton2_hline01 -
        ' cmbViaButton2_hline15
        'cmbViaForm2/cmbViaForm2_lbl
        'MENU1:
        'cmbMenu1/cmbMenu1_lbl
        'cmbButton1_hline01 -
        ' cmbButton1_hline15
        'cmbButton1_lbl_dim_hi
        'cmbButton1/cmbButton1_lbl
        'MENU2:
        'cmbMenu2/cmbMenu2_lbl
        'cmbButton2_hline01 -
        ' cmbButton2_hline15
        'cmbButton2_lbl_dim_hi
        'cmbButton2/cmbButton2_lbl
        'MENU3:
        'cmbMenu3/cmbMenu3_lbl
        'cmbButton3_hline01 -
        ' cmbButton3_hline15
        'cmbButton3/cmbButton3_lbl

EXITP:
13240   Set frm = Nothing
13250   Set rst = Nothing
13260   Set qdf = Nothing
13270   Set dbs = Nothing
13280   Rpt_SearchZOrder = blnRetVal
13290   Exit Function

ERRH:
13300   blnRetVal = False
13310   Select Case ERR.Number
        Case Else
13320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13330   End Select
13340   Resume EXITP

End Function

Public Function ReportList_Order() As Boolean

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "ReportList_Order"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

13410   blnRetVal = True

13420   Set dbs = CurrentDb
13430   With dbs
          ' ** tblReport_List_Report, sorted by rpt_caption, with rptlistrpt_datemodified.
13440     Set qdf = .QueryDefs("qryReport_List_50a")
13450     Set rst = qdf.OpenRecordset
13460     With rst
13470       If .BOF = True And .EOF = True Then
13480         Beep
13490         blnRetVal = False
13500       Else
13510         .MoveLast
13520         lngRecs = .RecordCount
13530         .MoveFirst
              ' ** First, move them all out of the way.
13540         For lngX = 1& To lngRecs
13550           .Edit
13560           ![rptlistrpt_order] = (![rptlistrpt_order] * 1000&)
13570           .Update
13580           If lngX < lngRecs Then .MoveNext
13590         Next
13600         .MoveFirst
              ' ** Then, renumber sequentially.
13610         For lngX = 1& To lngRecs
13620           .Edit
13630           ![rptlistrpt_order] = lngX
13640           ![rptlistrpt_datemodified] = Now()
13650           .Update
13660           If lngX < lngRecs Then .MoveNext
13670         Next
13680       End If
13690     End With
13700     .Close
13710   End With

EXITP:
13720   Set rst = Nothing
13730   Set qdf = Nothing
13740   Set dbs = Nothing
13750   ReportList_Order = blnRetVal
13760   Exit Function

ERRH:
13770   blnRetVal = False
13780   Select Case ERR.Number
        Case Else
13790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13800   End Select
13810   Resume EXITP

End Function

Public Function ReportList_DelJournal() As Boolean
' ** Called by:
' **   frmMap_Reinvest_DivInt_Detail:
' **     Form_Timer()
' **   frmMap_Reinvest_Rec_Detail:
' **     Form_Timer()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "ReportList_DelJournal"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

13910   blnRetVal = True

13920   Set dbs = CurrentDb
13930   With dbs
          ' ** Delete Journal, for journal_USER = 'System', by specified [jrnid].
13940     Set qdf = .QueryDefs("qryReport_List_94a")
13950     With qdf.Parameters
13960       ![jrnid] = glngTaxCode_Distribution  ' ** Borrowing this variable from the Court Reports.
13970     End With
13980     qdf.Execute
13990     Set qdf = Nothing
          ' ** Delete tblJournal_Columns, for journal_USER = 'System', by specified [jcolid].
14000     Set qdf = .QueryDefs("qryReport_List_94b")
14010     With qdf.Parameters
14020       ![jcolid] = glngTaxCode_Distribution
14030     End With
14040     qdf.Execute
14050     glngTaxCode_Distribution = 0&
14060     .Close
14070   End With

EXITP:
14080   Set rst = Nothing
14090   Set qdf = Nothing
14100   Set dbs = Nothing
14110   ReportList_DelJournal = blnRetVal
14120   Exit Function

ERRH:
14130   blnRetVal = False
14140   Select Case ERR.Number
        Case Else
14150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14160   End Select
14170   Resume EXITP

End Function

Public Function Rpt_SecRename() As Boolean

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_SecRename"

        Dim dbs As DAO.Database, cntr As DAO.Container, doc As DAO.Document, rpt As Access.Report, Sec As Access.Section
        Dim lngRpts As Long, arr_varRpt() As Variant
        Dim lngFixRpts As Long, lngFixes As Long, lngFixedRpts As Long, lngFixed As Long
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varRpt().
        Const R_ELEMS As Integer = 7  ' ** Array's first-element UBound().
        Const R_RID   As Integer = 0
        Const R_RNAM  As Integer = 1
        Const R_SIDX1 As Integer = 2
        Const R_SNAM1 As Integer = 3
        Const R_FIX1  As Integer = 4
        Const R_SIDX2 As Integer = 5
        Const R_SNAM2 As Integer = 6
        Const R_FIX2  As Integer = 7

14210 On Error GoTo 0

14220   blnRetVal = True

14230   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
14240   DoEvents

14250   Set dbs = CurrentDb
14260   With dbs

14270     lngRpts = 0&
14280     ReDim arr_varRpt(R_ELEMS, 0)

14290     Set cntr = .Containers("Reports")
14300     With cntr
14310       For Each doc In .Documents
14320         With doc
14330           lngRpts = lngRpts + 1&
14340           lngE = lngRpts - 1&
14350           ReDim Preserve arr_varRpt(R_ELEMS, lngE)
14360           arr_varRpt(R_RID, lngE) = CLng(0)
14370           arr_varRpt(R_RNAM, lngE) = .Name
14380           arr_varRpt(R_SIDX1, lngE) = Null
14390           arr_varRpt(R_SNAM1, lngE) = Null
14400           arr_varRpt(R_FIX1, lngE) = CBool(False)
14410           arr_varRpt(R_SIDX2, lngE) = Null
14420           arr_varRpt(R_SNAM2, lngE) = Null
14430           arr_varRpt(R_FIX2, lngE) = CBool(False)
14440         End With
14450       Next
14460     End With
14470     Set doc = Nothing
14480     Set cntr = Nothing

14490     .Close
14500   End With
14510   Set dbs = Nothing

14520   Debug.Print "'RPTS: " & CStr(lngRpts)
14530   DoEvents

        ' ** Make sure nothing's open.
14540   Do While Reports.Count > 0
14550     DoCmd.Close acReport, Reports(0).Name
14560     DoEvents
14570   Loop

14580   If lngRpts > 0& Then
14590     For lngX = 0& To (lngRpts - 1&)

14600       DoCmd.OpenReport arr_varRpt(R_RNAM, lngX), acViewDesign, , , acHidden
14610       Set rpt = Reports(arr_varRpt(R_RNAM, lngX))
14620       With rpt
14630 On Error Resume Next
14640         Set Sec = .Section(acPageHeader)
14650         If ERR.Number = 0 Then
14660 On Error GoTo 0
14670           With Sec
14680             arr_varRpt(R_SIDX1, lngX) = acPageHeader
14690             arr_varRpt(R_SNAM1, lngX) = .Name
14700           End With
14710           Set Sec = Nothing
14720           Set Sec = .Section(acPageFooter)
14730           With Sec
14740             arr_varRpt(R_SIDX2, lngX) = acPageFooter
14750             arr_varRpt(R_SNAM2, lngX) = .Name
14760           End With
14770           Set Sec = Nothing
14780         Else
14790 On Error GoTo 0
14800         End If
14810       End With
14820       Set rpt = Nothing
14830       DoCmd.Close acReport, arr_varRpt(R_RNAM, lngX), acSaveNo
14840     Next

14850     lngFixRpts = 0&: lngFixes = 0&
14860     For lngX = 0& To (lngRpts - 1&)
14870       If IsNull(arr_varRpt(R_SNAM1, lngX)) = False Then
14880         If InStr(arr_varRpt(R_SNAM1, lngX), "Page") > 0 And InStr(arr_varRpt(R_SNAM1, lngX), "Header") > 0 Then
14890           If InStr(arr_varRpt(R_SNAM1, lngX), "Section") = 0 Then
14900             arr_varRpt(R_FIX1, lngX) = CBool(True)
14910             lngFixes = lngFixes + 1&
14920           End If
14930         Else
14940           Stop
14950         End If
14960       End If
14970       If IsNull(arr_varRpt(R_SNAM2, lngX)) = False Then
14980         If InStr(arr_varRpt(R_SNAM2, lngX), "Page") > 0 And InStr(arr_varRpt(R_SNAM2, lngX), "Footer") > 0 Then
14990           If InStr(arr_varRpt(R_SNAM2, lngX), "Section") = 0 Then
15000             arr_varRpt(R_FIX2, lngX) = CBool(True)
15010             lngFixes = lngFixes + 1&
15020           End If
15030         Else
15040           Stop
15050         End If
15060       End If
15070       If arr_varRpt(R_FIX1, lngX) = True Or arr_varRpt(R_FIX2, lngX) = True Then
15080         lngFixRpts = lngFixRpts + 1&
15090       End If
15100     Next

15110     Debug.Print "'FIXES: " & CStr(lngFixes) & "  RPTS: " & CStr(lngFixRpts)
15120     DoEvents

15130     lngFixedRpts = 0&: lngFixed = 0&
15140     If lngFixRpts > 0& Then
15150       For lngX = 0& To (lngRpts - 1&)
15160         If arr_varRpt(R_FIX1, lngX) = True Or arr_varRpt(R_FIX2, lngX) = True Then
15170           DoCmd.OpenReport arr_varRpt(R_RNAM, lngX), acViewDesign, , , acHidden
15180           Set rpt = Reports(arr_varRpt(R_RNAM, lngX))
15190           With rpt
15200             If arr_varRpt(R_FIX1, lngX) = True Then
15210               Set Sec = .Section(acPageHeader)
15220               With Sec
15230                 .Name = "PageHeaderSection"
15240                 lngFixed = lngFixed + 1&
15250               End With
15260               Set Sec = Nothing
15270             End If
15280             If arr_varRpt(R_FIX2, lngX) = True Then
15290               Set Sec = .Section(acPageFooter)
15300               With Sec
15310                 .Name = "PageFooterSection"
15320                 lngFixed = lngFixed + 1&
15330               End With
15340               Set Sec = Nothing
15350             End If
15360           End With
15370           Set rpt = Nothing
15380           DoCmd.Close acReport, arr_varRpt(R_RNAM, lngX), acSaveYes
15390           lngFixedRpts = lngFixedRpts + 1&
15400         End If
15410       Next
15420     End If  ' ** lngFixRpts.

15430   End If  ' ** lngRpts.

15440   Debug.Print "'FIXED RPTS: " & CStr(lngFixedRpts) & "  FIXED SECS: " & CStr(lngFixed)
15450   DoEvents

15460   Beep

15470   Debug.Print "'DONE!"
15480   DoEvents

        'RPTS: 240
        'FIXES: 37  RPTS: 19
        'FIXED RPTS: 19  FIXED SECS: 37
        'DONE!

EXITP:
15490   Set Sec = Nothing
15500   Set rpt = Nothing
15510   Set doc = Nothing
15520   Set cntr = Nothing
15530   Set dbs = Nothing
15540   Rpt_SecRename = blnRetVal
15550   Exit Function

ERRH:
15560   blnRetVal = False
15570   Select Case ERR.Number
        Case Else
15580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
15590   End Select
15600   Resume EXITP

End Function

Public Function Rpt_FindVBA() As Boolean

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_FindVBA"

        Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
        Dim lngLines As Long, lngDecLines As Long
        Dim strModName As String, strRptName As String, strProcName As String, strLine As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

15710   blnRetVal = True

15720   Set vbp = Application.VBE.ActiveVBProject
15730   With vbp
15740     For Each vbc In .VBComponents
15750       With vbc
15760         strModName = .Name
15770         If Left(strModName, 7) = "Report_" Then
15780           strRptName = Mid(strModName, (InStr(strModName, "_") + 1))
15790           Set cod = .CodeModule
15800           With cod

15810             lngLines = .CountOfLines
15820             lngDecLines = .CountOfDeclarationLines

15830             For lngX = lngDecLines To lngLines
15840               strLine = .Lines(lngX, 1)
15850               If Trim(strLine) <> vbNullString Then
15860                 If Left(Trim(strLine), 1) <> "'" Then
15870                   If InStr(strLine, "SortNow_Get") > 0 Then
15880                     strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
15890                     Debug.Print "'RPT: " & strRptName
15900                     Debug.Print "'  '" & strLine & "'"
                          'strOrderBy = frmSub.SortNow_Get

15910                     Exit For
15920                   End If
15930                 End If  ' ** Remark.
15940               End If  ' ** vbNullString.
15950             Next  ' ** lngX.

15960           End With  ' ** cod.
15970         End If  ' ** Report_.
15980       End With  ' ** vbc.
15990     Next  ' ** vbc.
16000   End With  ' ** vbp

        'X rptAccountContacts
        'X rptCurrencies
        'X rptCurrencyRateHistory
        'X rptCurrencyRates
        'X rptListOfAssets_..
        'X rptListOfCountryCodes
        'X rptListOfCurrencySymbols
        'X rptListOfFeeSchedules_..

        ' In ('rptAccountContacts','rptCurrencies','rptCurrencyRateHistory','rptCurrencyRates','rptListOfCountryCodes','rptListOfCurrencySymbols')

        'strOrderBy = frmSub.SortNow_Get

        'RPT: rptAccountContacts
        '  '380           strSortNow = frm.frmAccountContacts_Sub.Form.SortNow_Get  ' ** Form Function: frmAccountContacts_Sub.'
        'RPT: rptCurrencies
        '  '390         strOrderBy = Forms(strCallingForm).frmCurrency_Sub.Form.SortNow_Get  ' ** Form Function: frmCurrency_Rate_Sub.'
        'RPT: rptCurrencyRateHistory
        '  '290         strOrderBy = frmSub.SortNow_Get  ' ** Form Function: frmCurrency_Rate_Sub.'
        'RPT: rptCurrencyRates
        '  '270         strOrderBy = Forms(strCallingForm).frmCurrency_Rate_Sub.Form.SortNow_Get  ' ** Form Function: frmCurrency_Rate_Sub.'
        'RPT: rptListOfCountryCodes
        '  '470         strOrderBy = Forms(strCallingForm).frmCountryCode_Sub.Form.SortNow_Get  ' ** Form Function: frmCountryCode_Sub.'
        'RPT: rptListOfCurrencySymbols
        '  '240         strOrderBy = Forms(strCallingForm).frmCurrency_Symbol_Sub.Form.SortNow_Get  ' ** Form Function: frmCurrency_Symbol_Sub.'
16010   Beep

EXITP:
16020   Set cod = Nothing
16030   Set vbc = Nothing
16040   Set vbp = Nothing
16050   Rpt_FindVBA = blnRetVal
16060   Exit Function

ERRH:
16070   blnRetVal = False
16080   Select Case ERR.Number
        Case Else
16090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
16100   End Select
16110   Resume EXITP

End Function

Public Function Rpt_DateLabels() As Boolean

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_DateLabels"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, rpt As Access.Report, ctl As Access.Control
        Dim lngCtls As Long, arr_varCtl As Variant
        Dim strRptName As String, strXType As String
        Dim lngEdits As Long
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim blnSkip As Boolean
        Dim lngX As Long, lngY As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varCtl().
        'Const C_DID  As Integer = 0
        'Const C_DNAM As Integer = 1
        Const C_RID  As Integer = 2
        Const C_RNAM As Integer = 3
        'Const C_CID  As Integer = 4
        Const C_CNAM As Integer = 5
        'Const C_CTYP As Integer = 6
        Const C_CAP  As Integer = 7
        'Const C_XTYP As Integer = 8

16210 On Error GoTo 0

16220   blnRetVal = True

16230   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
16240   DoEvents

16250   Do While Reports.Count > 0
16260     DoCmd.Close acReport, Reports(0).Name, acSaveYes
16270     DoEvents
16280   Loop

16290   Set dbs = CurrentDb
16300   With dbs

16310     lngEdits = 0&
16320     For lngX = 1& To 12&

16330       Select Case lngX
            Case 1
16340         Set qdf = .QueryDefs("zzz_qry_zReport_02_03_01")
16350         strXType = "D1x"
16360       Case 2
16370         Set qdf = .QueryDefs("zzz_qry_zReport_02_04_01")
16380         strXType = "D2x"
16390       Case 3
16400         Set qdf = .QueryDefs("zzz_qry_zReport_02_05_01")
16410         strXType = "AO1x"
16420       Case 4
16430         Set qdf = .QueryDefs("zzz_qry_zReport_02_06_01")
16440         strXType = "AO2x"
16450       Case 5
16460         Set qdf = .QueryDefs("zzz_qry_zReport_02_07_01")
16470         strXType = "P1x"
16480       Case 6
16490         Set qdf = .QueryDefs("zzz_qry_zReport_02_08_01")
16500         strXType = "P2x"
16510       Case 7
16520         Set qdf = .QueryDefs("zzz_qry_zReport_02_09_01")
16530         strXType = "Fx"
16540       Case 8
16550         Set qdf = .QueryDefs("zzz_qry_zReport_02_10_01")
16560         strXType = "FP1x"
16570       Case 9
16580         Set qdf = .QueryDefs("zzz_qry_zReport_02_11_01")
16590         strXType = "FP2x"
16600       Case 10
16610         Set qdf = .QueryDefs("zzz_qry_zReport_02_12_01")
16620         strXType = "MAx"
16630       Case 11
16640         Set qdf = .QueryDefs("zzz_qry_zReport_02_13_01")
16650         strXType = "Cx"
16660       Case 12
16670         Set qdf = .QueryDefs("zzz_qry_zReport_02_14_01")
16680         strXType = "Vx"
16690       End Select

16700       Set rst = qdf.OpenRecordset
16710       With rst
16720         .MoveLast
16730         lngCtls = .RecordCount
16740         .MoveFirst
16750         arr_varCtl = .GetRows(lngCtls)
              ' ************************************************
              ' ** Array: arr_varCtl().
              ' **
              ' **   Field  Element  Name           Constant
              ' **   =====  =======  =============  ==========
              ' **     1       0     dbs_id         C_DID
              ' **     2       1     dbs_name       C_DNAM
              ' **     3       2     rpt_id         C_RID
              ' **     4       3     rpt_name       C_RNAM
              ' **     5       4     ctl_id         C_CID
              ' **     6       5     ctl_name       C_CNAM
              ' **     7       6     ctltype_type   C_CTYP
              ' **     8       7     ctl_caption    C_CAP
              ' **     9       8     D1x,D2x,...    C_XTYP
              ' **
              ' ************************************************
16760         .Close
16770       End With
16780       Set rst = Nothing
16790       Set qdf = Nothing

16800       If lngCtls > 0& Then

16810         strRptName = vbNullString
16820         For lngY = 0& To (lngCtls - 1&)

16830           If arr_varCtl(C_RNAM, lngY) <> strRptName Then
16840             If strRptName <> vbNullString Then
16850               DoCmd.Close acReport, strRptName, acSaveYes
16860               Set rpt = Nothing
16870               DoEvents
16880             End If
16890             strRptName = arr_varCtl(C_RNAM, lngY)
16900             DoCmd.OpenReport strRptName, acViewDesign, , , acHidden
16910             Set rpt = Reports(strRptName)
16920           End If

16930           With rpt
16940             Set ctl = .Controls(arr_varCtl(C_CNAM, lngY))
16950             blnSkip = False
16960             If blnSkip = False Then
16970               strTmp01 = arr_varCtl(C_CAP, lngY)
16980               Select Case strXType
                    Case "D1x"  ' ** Short date.
16990                 If Len(arr_varCtl(C_CAP, lngY)) = 10 Then
17000                   Select Case arr_varCtl(C_CNAM, lngY)
                        Case "DateBeg_lbl"
17010                     ctl.Caption = "01/01/2015"
17020                   Case "DateEnd_lbl"
17030                     ctl.Caption = "12/31/2015"
17040                   End Select
17050                 Else
                        ' ** With time.
17060                   ctl.Caption = "01/01/2016 09:30 AM"
17070                 End If
17080               Case "D2x"  ' ** Long date.
17090                 If Left(strTmp01, 3) = "Jan" Then
17100                   intPos01 = InStr(strTmp01, ",")
17110                   If intPos01 > 0 Then
17120                     strTmp01 = Left(strTmp01, intPos01) & " 2016"
17130                     ctl.Caption = strTmp01
17140                   Else
17150                     Stop
17160                   End If
17170                 ElseIf Left(strTmp01, 3) = "Dec" Then
17180                   intPos01 = InStr(strTmp01, " ")
17190                   strTmp01 = Left(strTmp01, intPos01) & "2015"
17200                   ctl.Caption = strTmp01
17210                 Else
17220                   Stop
17230                 End If
17240               Case "AO1x"  ' ** 'As Of', Short date.
17250                 intPos01 = CharPos(strTmp01, 2, "/")  ' ** Module Function: modStringFuncs.
17260                 If intPos01 > 0 Then
17270                   strTmp01 = Left(strTmp01, intPos01) & "2016"
17280                   ctl.Caption = strTmp01
17290                 Else
17300                   Stop
17310                 End If
17320               Case "AO2x"  ' ** 'As Of', Long date.
17330                 intPos01 = InStr(strTmp01, ",")
17340                 If intPos01 > 0 Then
17350                   strTmp01 = Left(strTmp01, intPos01) & " 2016"
17360                   ctl.Caption = strTmp01
17370                 Else
17380                   Stop
17390                 End If
17400               Case "P1x"  ' ** 'Printed', Short date.
17410                 intPos01 = CharPos(strTmp01, 2, "/")  ' ** Module Function: modStringFuncs.
17420                 If intPos01 > 0 Then
17430                   strTmp01 = Left(strTmp01, intPos01) & "2016"
17440                   ctl.Caption = strTmp01
17450                 Else
17460                   Stop
17470                 End If
17480               Case "P2x"  ' ** 'Printed', Long date.
17490                 intPos01 = InStr(strTmp01, ",")
17500                 If intPos01 > 0 Then
17510                   strTmp01 = Left(strTmp01, intPos01) & " 2016"
17520                   ctl.Caption = strTmp01
17530                 Else
17540                   Stop
17550                 End If
17560               Case "Fx"   ' ** 'From/To'.
17570                 ctl.Caption = "From 01/01/2015 To 12/31/2015"
17580               Case "FP1x"  ' ** 'For Period', Short date.
17590                 intPos01 = CharPos(strTmp01, 2, "/")  ' ** Module Function: modStringFuncs.
17600                 If intPos01 > 0 Then
17610                   strTmp01 = Left(strTmp01, intPos01) & "2015"
17620                   ctl.Caption = strTmp01
17630                 Else
17640                   Stop
17650                 End If
17660               Case "FP2x"  ' ** 'For Period', Long date.
17670                 ctl.Caption = "For Period Ending December 31, 2015"
17680               Case "MAx"  ' ** 'Market as of'.
17690                 intPos01 = CharPos(strTmp01, 2, "/")  ' ** Module Function: modStringFuncs.
17700                 If intPos01 > 0 Then
17710                   strTmp01 = Left(strTmp01, intPos01) & "2015"
17720                   ctl.Caption = strTmp01
17730                 Else
17740                   Stop
17750                 End If
17760               Case "Cx"  ' ** 'Copyright'
17770                 ctl.Caption = "Copyright © 1998-2016 Delta Data Inc."
17780               Case "Vx"  ' ** 'Ver.'.
17790                 ctl.Caption = "Ver: 2.2.23, Rel: 05/29/2016 09:00:00"
17800               End Select
17810             End If  ' ** blnSkip.
17820           End With
17830           Set ctl = Nothing
17840           lngEdits = lngEdits + 1&

17850         Next  ' ** lngY.

17860       End If  ' ** lngCtls.

            'Stop

17870     Next  ' ** lngX.
17880     DoCmd.Close acReport, strRptName, acSaveYes
17890     DoEvents
17900     If Reports.Count > 0 Then
17910       Stop
17920       Do While Reports.Count > 0
17930         DoCmd.Close acReport, Reports(0).Name, acSaveYes
17940       Loop
17950     End If

17960     blnSkip = True
17970     If blnSkip = False Then

17980       Set rst = .OpenRecordset("tblReport", dbOpenDynaset, dbConsistent)
17990       With rst
18000         .MoveFirst
18010         For lngX = 0& To (lngCtls - 1&)
18020           .FindFirst "[dbs_id] = 1 And [rpt_name] = '" & arr_varCtl(C_RNAM, lngX) & "'"
18030           If .NoMatch = False Then
18040             If arr_varCtl(C_RID, lngX) <> ![rpt_id] Then
18050               arr_varCtl(C_RID, lngX) = ![rpt_id]
18060             End If
18070           Else
18080             Stop
18090           End If
18100         Next  ' ** lngX.
18110         .Close
18120       End With  ' ** rst.
18130       Set rst = Nothing
18140       DoEvents

18150       Set rst = .OpenRecordset("tblReport_Control", dbOpenDynaset, dbConsistent)
18160       With rst
18170         .MoveFirst
18180         For lngX = 0& To (lngCtls - 1&)
18190           .FindFirst "[dbs_id] = 1 And [rpt_id] = " & CStr(arr_varCtl(C_RID, lngX)) & " And " & _
                  "[ctl_name] = '" & arr_varCtl(C_CNAM, lngX) & "'"
18200           If .NoMatch = False Then
18210             .Edit
18220             ![ctl_caption] = arr_varCtl(C_CAP, lngX)
18230             .Update
18240           Else
18250             Stop
18260           End If
18270         Next  ' ** lngX.
18280         .Close
18290       End With  ' ** rst.
18300       Set rst = Nothing
18310       DoEvents

18320     End If  ' ** blnSkip.

18330     .Close
18340   End With
18350   Set dbs = Nothing

18360   Debug.Print "'LBLS EDITED: " & CStr(lngEdits)
18370   DoEvents

18380   Beep

18390   Debug.Print "'DONE!"
18400   DoEvents

        'LBLS EDITED: 303
        'DONE!

EXITP:
18410   Set ctl = Nothing
18420   Set rpt = Nothing
18430   Set rst = Nothing
18440   Set qdf = Nothing
18450   Set dbs = Nothing
18460   Rpt_DateLabels = blnRetVal
18470   Exit Function

ERRH:
18480   blnRetVal = False
18490   Select Case ERR.Number
        Case Else
18500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
18510   End Select
18520   Resume EXITP

End Function

Public Function Rpt_FeeSchedules() As Boolean

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_FeeSchedules"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, cntr As DAO.Container, doc As DAO.Document
        Dim rpt1 As Access.Report, ctl1 As Access.Control, rpt2 As Access.Report, ctl2 As Access.Control
        Dim lngRpts As Long, arr_varRpt As Variant
        Dim lngRetVals As Long, arr_varRetVal As Variant
        Dim strRptName As String, strCtlSource As String, strSortNow As String, strDesc As String
        Dim blnSkip As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varRpt().
        Const R_ELEMS As Integer = 18 '5  ' ** Array's first-element UBound().
        Const R_DID   As Integer = 0
        Const R_DNAM  As Integer = 1
        Const R_RID   As Integer = 2
        Const R_RNAM  As Integer = 3
        Const R_CAP   As Integer = 4
        Const R_DSC   As Integer = 5
        Const R_DSCN  As Integer = 6
        Const R_GRP0  As Integer = 7
        Const R_GRP0S As Integer = 8
        Const R_GRP1  As Integer = 9
        Const R_GRP1S As Integer = 10
        Const R_GRP2  As Integer = 11
        Const R_GRP2S As Integer = 12
        Const R_GRP3  As Integer = 13
        Const R_GRP3S As Integer = 14
        Const R_GRP4  As Integer = 15
        Const R_GRP4S As Integer = 16
        Const R_GRP5  As Integer = 17
        Const R_GRP5S As Integer = 18

18610 On Error GoTo 0

18620   blnRetVal = True

18630   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
18640   DoEvents

18650   Set dbs = CurrentDb
18660   With dbs
          ' ** tblReport, just 'rptListOfFeeSchedules_..'.
          'Set qdf = .QueryDefs("qryRpt_ListOfFeeSchedules_10")
          ' ** tblReport, just 'rptListOfFeeSchedules_..', with rptgrp_grouplevel's.
18670     Set qdf = .QueryDefs("qryRpt_ListOfFeeSchedules_11")
18680     Set rst = qdf.OpenRecordset
18690     With rst
18700       .MoveLast
18710       lngRetVals = .RecordCount
18720       .MoveFirst
18730       arr_varRetVal = .GetRows(lngRetVals)
            ' ********************************************************
            ' ** Array: arr_varRpt()
            ' **
            ' **   Field  Element  Name                   Constant
            ' **   =====  =======  =====================  ==========
            ' **     1       0     dbs_id                 R_DID
            ' **     2       1     dbs_name               R_DNAM
            ' **     3       2     rpt_id                 R_RID
            ' **     4       3     rpt_name               R_RNAM
            ' **     5       4     rpt_caption            R_CAP
            ' **     6       5     rpt_description        R_DSC
            ' **
            ' **     7       6     rpt_description_new    R_DSCN
            ' **     8       7     rptgrp_grouplevel0     R_GRP0
            ' **     9       8     rptgrp_grouplevel0s    R_GRP0S
            ' **    10       9     rptgrp_grouplevel1     R_GRP1
            ' **    11      10     rptgrp_grouplevel1s    R_GRP1S
            ' **    12      11     rptgrp_grouplevel2     R_GRP2
            ' **    13      12     rptgrp_grouplevel2s    R_GRP2S
            ' **    14      13     rptgrp_grouplevel3     R_GRP3
            ' **    15      14     rptgrp_grouplevel3s    R_GRP3S
            ' **    16      15     rptgrp_grouplevel4     R_GRP4
            ' **    17      16     rptgrp_grouplevel4s    R_GRP4S
            ' **    18      17     rptgrp_grouplevel5     R_GRP5
            ' **    19      18     rptgrp_grouplevel5s    R_GRP5S
            ' **
            ' ********************************************************
18740       .Close
18750     End With
18760     Set rst = Nothing
18770     Set qdf = Nothing
18780   End With
18790   Set dbs = Nothing

18800   lngRpts = 0&
18810   ReDim arr_varRpt(R_ELEMS, 0)

18820   For lngX = 0& To (lngRetVals - 1&)
18830     Select Case arr_varRetVal(R_RNAM, lngX)
          Case "rptListOfFeeSchedules_01a_02d", "rptListOfFeeSchedules_01d_02d", "rptListOfFeeSchedules_02a_02d", _
              "rptListOfFeeSchedules_02d_02d", "rptListOfFeeSchedules_03a_02d", "rptListOfFeeSchedules_03d_02d", _
              "rptListOfFeeSchedules_04a_02d", "rptListOfFeeSchedules_04d_02d"
18840       lngRpts = lngRpts + 1&
18850       lngE = lngRpts - 1&
18860       ReDim Preserve arr_varRpt(R_ELEMS, lngE)
18870       For lngZ = 0& To R_ELEMS
18880         arr_varRpt(lngZ, lngE) = arr_varRetVal(lngZ, lngX)
18890       Next
18900       strRptName = Left(arr_varRetVal(R_RNAM, lngX), CharPos(arr_varRetVal(R_RNAM, lngX), 2, "_"))
18910       strRptName = strRptName & "03a"
18920       lngRpts = lngRpts + 1&
18930       lngE = lngRpts - 1&
18940       ReDim Preserve arr_varRpt(R_ELEMS, lngE)
18950       arr_varRpt(R_DID, lngE) = arr_varRetVal(R_DID, lngX)
18960       arr_varRpt(R_DNAM, lngE) = arr_varRetVal(R_DNAM, lngX)
18970       arr_varRpt(R_RID, lngE) = Null
18980       arr_varRpt(R_RNAM, lngE) = strRptName
18990       arr_varRpt(R_CAP, lngE) = arr_varRetVal(R_CAP, lngX)
19000       arr_varRpt(R_DSC, lngE) = arr_varRetVal(R_DSC, lngX)
19010       strRptName = Left(strRptName, (Len(strRptName) - 1)) & "d"
19020       lngRpts = lngRpts + 1&
19030       lngE = lngRpts - 1&
19040       ReDim Preserve arr_varRpt(R_ELEMS, lngE)
19050       arr_varRpt(R_DID, lngE) = arr_varRetVal(R_DID, lngX)
19060       arr_varRpt(R_DNAM, lngE) = arr_varRetVal(R_DNAM, lngX)
19070       arr_varRpt(R_RID, lngE) = Null
19080       arr_varRpt(R_RNAM, lngE) = strRptName
19090       arr_varRpt(R_CAP, lngE) = arr_varRetVal(R_CAP, lngX)
19100       arr_varRpt(R_DSC, lngE) = arr_varRetVal(R_DSC, lngX)
19110     Case Else
19120       lngRpts = lngRpts + 1&
19130       lngE = lngRpts - 1&
19140       ReDim Preserve arr_varRpt(R_ELEMS, lngE)
19150       For lngZ = 0& To R_ELEMS
19160         arr_varRpt(lngZ, lngE) = arr_varRetVal(lngZ, lngX)
19170       Next
19180     End Select
19190   Next

19200   Debug.Print "'RPTS: " & CStr(lngRpts)
19210   DoEvents

        'For lngX = 0& To (lngRpts - 1&)
        '  Debug.Print "'" & arr_varRpt(R_RNAM, lngX)
        'Next

        'For lngX = 0& To (lngRpts - 1&)
        '  DoCmd.OpenReport arr_varRpt(R_RNAM, lngX), acViewDesign, , , acWindowNormal
        '  If (lngX + 1&) Mod 12 = 0 Then
        '    Stop
        '  End If
        'Next

19220   For lngX = 0& To (lngRpts - 1&)
19230     arr_varRpt(R_DSCN, lngX) = Null
19240     arr_varRpt(R_GRP0, lngX) = Null
19250     arr_varRpt(R_GRP0S, lngX) = Null
19260     arr_varRpt(R_GRP1, lngX) = Null
19270     arr_varRpt(R_GRP1S, lngX) = Null
19280     arr_varRpt(R_GRP2, lngX) = Null
19290     arr_varRpt(R_GRP2S, lngX) = Null
19300     arr_varRpt(R_GRP3, lngX) = Null
19310     arr_varRpt(R_GRP3S, lngX) = Null
19320     arr_varRpt(R_GRP4, lngX) = Null
19330     arr_varRpt(R_GRP4S, lngX) = Null
19340     arr_varRpt(R_GRP5, lngX) = Null
19350     arr_varRpt(R_GRP5S, lngX) = Null
19360   Next

19370   For lngX = 0& To (lngRpts - 1&)
19380     strRptName = arr_varRpt(R_RNAM, lngX)
19390     DoCmd.OpenReport strRptName, acViewDesign, , , acHidden
19400     Set rpt1 = Reports(strRptName)
19410     With rpt1
19420       For lngY = 0& To 5&
19430 On Error Resume Next
19440         strCtlSource = .GroupLevel(lngY).ControlSource
19450         If ERR.Number = 0 Then
19460 On Error GoTo 0
19470           Select Case .GroupLevel(lngY).sortOrder
                Case True
19480             strSortNow = "D"
19490           Case False
19500             strSortNow = "A"
19510           End Select
19520           Select Case lngY
                Case 0&
19530             arr_varRpt(R_GRP0, lngX) = strCtlSource
19540             arr_varRpt(R_GRP0S, lngX) = strSortNow
19550           Case 1&
19560             arr_varRpt(R_GRP1, lngX) = strCtlSource
19570             arr_varRpt(R_GRP1S, lngX) = strSortNow
19580           Case 2&
19590             arr_varRpt(R_GRP2, lngX) = strCtlSource
19600             arr_varRpt(R_GRP2S, lngX) = strSortNow
19610           Case 3&
19620             arr_varRpt(R_GRP3, lngX) = strCtlSource
19630             arr_varRpt(R_GRP3S, lngX) = strSortNow
19640           Case 4&
19650             arr_varRpt(R_GRP4, lngX) = strCtlSource
19660             arr_varRpt(R_GRP4S, lngX) = strSortNow
19670           Case 5&
19680             arr_varRpt(R_GRP5, lngX) = strCtlSource
19690             arr_varRpt(R_GRP5S, lngX) = strSortNow
19700           End Select
19710         Else
19720 On Error GoTo 0
19730         End If
19740       Next
19750     End With
19760     Set rpt1 = Nothing
19770     DoCmd.Close acReport, strRptName, acSaveNo
19780   Next

19790   blnSkip = True
19800   If blnSkip = False Then
          ' ** List of all Fee Schedules, with detail; from frmFeeSchedules.cmdPrintReport_Click(), by Schedule_ID, Rate, Amount.
19810     strDesc = arr_varRpt(R_DSC, 0)
19820     intPos01 = InStr(strDesc, " by ")
19830     strDesc = Left(strDesc, intPos01 + 3)
19840     For lngX = 0& To (lngRpts - 1&)
19850       strTmp01 = strDesc
19860       If IsNull(arr_varRpt(R_GRP0, lngX)) = False Then
19870         strTmp01 = strTmp01 & arr_varRpt(R_GRP0, lngX) & IIf(arr_varRpt(R_GRP0S, lngX) = "D", " D", vbNullString)
19880       End If
19890       If IsNull(arr_varRpt(R_GRP1, lngX)) = False Then
19900         strTmp01 = strTmp01 & ", " & arr_varRpt(R_GRP1, lngX) & IIf(arr_varRpt(R_GRP1S, lngX) = "D", " D", vbNullString)
19910       End If
19920       If IsNull(arr_varRpt(R_GRP2, lngX)) = False Then
19930         strTmp01 = strTmp01 & ", " & arr_varRpt(R_GRP2, lngX) & IIf(arr_varRpt(R_GRP2S, lngX) = "D", " D", vbNullString)
19940       End If
19950       If IsNull(arr_varRpt(R_GRP3, lngX)) = False Then
19960         strTmp01 = strTmp01 & ", " & arr_varRpt(R_GRP3, lngX) & IIf(arr_varRpt(R_GRP3S, lngX) = "D", " D", vbNullString)
19970       End If
19980       If IsNull(arr_varRpt(R_GRP4, lngX)) = False Then
19990         strTmp01 = strTmp01 & ", " & arr_varRpt(R_GRP4, lngX) & IIf(arr_varRpt(R_GRP4S, lngX) = "D", " D", vbNullString)
20000       End If
20010       If IsNull(arr_varRpt(R_GRP5, lngX)) = False Then
20020         strTmp01 = strTmp01 & ", " & arr_varRpt(R_GRP5, lngX) & IIf(arr_varRpt(R_GRP5S, lngX) = "D", " D", vbNullString)
20030       End If
20040       strTmp01 = strTmp01 & "."
20050       arr_varRpt(R_DSCN, lngX) = strTmp01
20060     Next
20070   End If  ' ** blnSkip.

        'For lngX = 0& To (lngRpts - 1&)
        '  Debug.Print "'" & arr_varRpt(R_DSCN, lngX)
        'Next

20080   Set dbs = CurrentDb
20090   With dbs
20100     Set cntr = .Containers("Reports")
20110     With cntr
20120       For lngX = 0& To (lngRpts - 1&)
20130         strRptName = arr_varRpt(R_RNAM, lngX)
20140         Set doc = .Documents(strRptName)
20150         With doc
20160           strDesc = .Properties("Description")
20170           strDesc = StringReplace(strDesc, "from frmFeeSchedules.cmdPrintReport_Click(), ", vbNullString)
20180           .Properties("Description") = strDesc
                '.Properties("Description") = arr_varRpt(R_DSCN, lngX)
20190         End With
20200         Set doc = Nothing
20210       Next
20220     End With
20230     Set cntr = Nothing
20240     .Close
20250   End With
20260   Set dbs = Nothing

20270   blnSkip = True
20280   If blnSkip = False Then
20290     For lngX = 0& To (lngRpts - 1&)
20300       Debug.Print "'" & arr_varRpt(R_RNAM, lngX) & "  ";
20310       If IsNull(arr_varRpt(R_GRP0, lngX)) = False Then
20320         Debug.Print arr_varRpt(R_GRP0, lngX) & "_" & arr_varRpt(R_GRP0S, lngX) & ";";
20330       End If
20340       If IsNull(arr_varRpt(R_GRP1, lngX)) = False Then
20350         Debug.Print arr_varRpt(R_GRP1, lngX) & "_" & arr_varRpt(R_GRP1S, lngX) & ";";
20360       End If
20370       If IsNull(arr_varRpt(R_GRP2, lngX)) = False Then
20380         Debug.Print arr_varRpt(R_GRP2, lngX) & "_" & arr_varRpt(R_GRP2S, lngX) & ";";
20390       End If
20400       If IsNull(arr_varRpt(R_GRP3, lngX)) = False Then
20410         Debug.Print arr_varRpt(R_GRP3, lngX) & "_" & arr_varRpt(R_GRP3S, lngX) & ";";
20420       End If
20430       If IsNull(arr_varRpt(R_GRP4, lngX)) = False Then
20440         Debug.Print arr_varRpt(R_GRP4, lngX) & "_" & arr_varRpt(R_GRP4S, lngX) & ";";
20450       End If
20460       If IsNull(arr_varRpt(R_GRP5, lngX)) = False Then
20470         Debug.Print arr_varRpt(R_GRP5, lngX) & "_" & arr_varRpt(R_GRP5S, lngX) & ";";
20480       End If
20490       Debug.Print
20500     Next
20510   End If  ' ** blnSkip.

20520   blnSkip = True
20530   If blnSkip = False Then
20540     Set rpt1 = Reports(0)
20550     For lngX = 1& To (lngRpts - 1&)
20560       strRptName = arr_varRpt(R_RNAM, lngX)
20570       DoCmd.OpenReport strRptName, acViewDesign, , , acHidden
20580       Set rpt2 = Reports(strRptName)
20590       With rpt2
20600         .RecordSource = rpt1.RecordSource
20610         .Schedule_ID.Width = rpt1.Schedule_ID.Width
20620         .Schedule_Name.Left = rpt1.Schedule_Name.Left
20630         .Schedule_Base.Left = rpt1.Schedule_Base.Left
20640         .Schedule_Base.Width = rpt1.Schedule_Base.Width
20650         .Schedule_Base_lbl.Left = rpt1.Schedule_Base_lbl.Left
20660         .Schedule_Base_lbl.Width = rpt1.Schedule_Base_lbl.Width
20670         .Schedule_Base_lbl_line.Left = rpt1.Schedule_Base_lbl_line.Left
20680         .Schedule_Base_lbl_line.Width = rpt1.Schedule_Base_lbl_line.Width
20690         .Schedule_Minimum.Left = rpt1.Schedule_Minimum.Left
20700         .Schedule_Minimum.Width = rpt1.Schedule_Minimum.Width
20710         .Schedule_Minimum_lbl.Left = rpt1.Schedule_Minimum_lbl.Left
20720         .Schedule_Minimum_lbl.Width = rpt1.Schedule_Minimum_lbl.Width
20730         .Schedule_Minimum_lbl_line.Left = rpt1.Schedule_Minimum_lbl_line.Left
20740         .Schedule_Minimum_lbl_line.Width = rpt1.Schedule_Minimum_lbl_line.Width
20750         .ScheduleDetail_Rate_display.Left = rpt1.ScheduleDetail_Rate_display.Left
20760         .ScheduleDetail_Rate_display.Width = rpt1.ScheduleDetail_Rate_display.Width
20770         .ScheduleDetail_Rate_display_lbl.Left = rpt1.ScheduleDetail_Rate_display_lbl.Left
20780         .ScheduleDetail_Rate_display_lbl.Width = rpt1.ScheduleDetail_Rate_display_lbl.Width
20790         .ScheduleDetail_Rate_display_lbl_line.Left = rpt1.ScheduleDetail_Rate_display_lbl_line.Left
20800         .ScheduleDetail_Rate_display_lbl_line.Width = rpt1.ScheduleDetail_Rate_display_lbl_line.Width
20810         .ScheduleDetail_Amount.Left = rpt1.ScheduleDetail_Amount.Left
20820         .ScheduleDetail_Amount.Width = rpt1.ScheduleDetail_Amount.Width
20830         .ScheduleDetail_Amount_lbl.Left = rpt1.ScheduleDetail_Amount_lbl.Left
20840         .ScheduleDetail_Amount_lbl.Width = rpt1.ScheduleDetail_Amount_lbl.Width
20850         .ScheduleDetail_Amount_lbl_line.Left = rpt1.ScheduleDetail_Amount_lbl_line.Left
20860         .ScheduleDetail_Amount_lbl_line.Width = rpt1.ScheduleDetail_Amount_lbl_line.Width
20870       End With
20880       DoCmd.Close acReport, strRptName, acSaveYes
20890       Set rpt2 = Nothing
20900     Next
20910   End If  ' ** blnSkip.

        'RPTS: 48
        'rptListOfFeeSchedules_01a_01a  Schedule_ID_A;scheddets_order_A;ScheduleDetail_Rate_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_01a_01d  Schedule_ID_A;scheddets_order_D;ScheduleDetail_Rate_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_01a_02a  Schedule_ID_A;ScheduleDetail_Rate_A;scheddets_order_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_01a_02d  Schedule_ID_A;ScheduleDetail_Rate_D;scheddets_order_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_01a_03a  Schedule_ID_A;ScheduleDetail_Amount_A;scheddets_order_A;ScheduleDetail_Rate_A;
        'rptListOfFeeSchedules_01a_03d  Schedule_ID_A;ScheduleDetail_Amount_D;scheddets_order_D;ScheduleDetail_Rate_D;

        'rptListOfFeeSchedules_01d_01a  Schedule_ID_D;scheddets_order_A;ScheduleDetail_Rate_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_01d_01d  Schedule_ID_D;scheddets_order_D;ScheduleDetail_Rate_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_01d_02a  Schedule_ID_D;ScheduleDetail_Rate_A;scheddets_order_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_01d_02d  Schedule_ID_D;ScheduleDetail_Rate_D;scheddets_order_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_01d_03a  Schedule_ID_D;ScheduleDetail_Amount_A;scheddets_order_A;ScheduleDetail_Rate_A;
        'rptListOfFeeSchedules_01d_03d  Schedule_ID_D;ScheduleDetail_Amount_D;scheddets_order_D;ScheduleDetail_Rate_D;

        'rptListOfFeeSchedules_02a_01a  Schedule_Name_A;scheddets_order_A;ScheduleDetail_Rate_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_02a_01d  Schedule_Name_A;scheddets_order_D;ScheduleDetail_Rate_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_02a_02a  Schedule_Name_A;ScheduleDetail_Rate_A;scheddets_order_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_02a_02d  Schedule_Name_A;ScheduleDetail_Rate_D;scheddets_order_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_02a_03a  Schedule_Name_A;ScheduleDetail_Amount_A;scheddets_order_A;ScheduleDetail_Rate_A;
        'rptListOfFeeSchedules_02a_03d  Schedule_Name_A;ScheduleDetail_Amount_D;scheddets_order_D;ScheduleDetail_Rate_D;

        'rptListOfFeeSchedules_02d_01a  Schedule_Name_D;scheddets_order_A;ScheduleDetail_Rate_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_02d_01d  Schedule_Name_D;scheddets_order_D;ScheduleDetail_Rate_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_02d_02a  Schedule_Name_D;ScheduleDetail_Rate_A;scheddets_order_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_02d_02d  Schedule_Name_D;ScheduleDetail_Rate_D;scheddets_order_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_02d_03a  Schedule_Name_D;ScheduleDetail_Amount_A;scheddets_order_A;ScheduleDetail_Rate_A;
        'rptListOfFeeSchedules_02d_03d  Schedule_Name_D;ScheduleDetail_Amount_D;scheddets_order_D;ScheduleDetail_Rate_D;

        'rptListOfFeeSchedules_03a_01a  Schedule_Base_A;Schedule_Name_A;scheddets_order_A;ScheduleDetail_Rate_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_03a_01d  Schedule_Base_A;Schedule_Name_A;scheddets_order_D;ScheduleDetail_Rate_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_03a_02a  Schedule_Base_A;Schedule_Name_A;ScheduleDetail_Rate_A;scheddets_order_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_03a_02d  Schedule_Base_A;Schedule_Name_A;ScheduleDetail_Rate_D;scheddets_order_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_03a_03a  Schedule_Base_A;Schedule_Name_A;ScheduleDetail_Amount_A;scheddets_order_A;ScheduleDetail_Rate_A;
        'rptListOfFeeSchedules_03a_03d  Schedule_Base_A;Schedule_Name_A;ScheduleDetail_Amount_D;scheddets_order_D;ScheduleDetail_Rate_D;

        'rptListOfFeeSchedules_03d_01a  Schedule_Base_D;Schedule_Name_D;scheddets_order_A;ScheduleDetail_Rate_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_03d_01d  Schedule_Base_D;Schedule_Name_D;scheddets_order_D;ScheduleDetail_Rate_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_03d_02a  Schedule_Base_D;Schedule_Name_D;ScheduleDetail_Rate_A;scheddets_order_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_03d_02d  Schedule_Base_D;Schedule_Name_D;ScheduleDetail_Rate_D;scheddets_order_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_03d_03a  Schedule_Base_D;Schedule_Name_D;ScheduleDetail_Amount_A;scheddets_order_A;ScheduleDetail_Rate_A;
        'rptListOfFeeSchedules_03d_03d  Schedule_Base_D;Schedule_Name_D;ScheduleDetail_Amount_D;scheddets_order_D;ScheduleDetail_Rate_D;

        'rptListOfFeeSchedules_04a_01a  Schedule_Minimum_A;Schedule_Name_A;scheddets_order_A;ScheduleDetail_Rate_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_04a_01d  Schedule_Minimum_A;Schedule_Name_A;scheddets_order_D;ScheduleDetail_Rate_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_04a_02a  Schedule_Minimum_A;Schedule_Name_A;ScheduleDetail_Rate_A;scheddets_order_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_04a_02d  Schedule_Minimum_A;Schedule_Name_A;ScheduleDetail_Rate_D;scheddets_order_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_04a_03a  Schedule_Minimum_A;Schedule_Name_A;ScheduleDetail_Amount_A;scheddets_order_A;ScheduleDetail_Rate_A;
        'rptListOfFeeSchedules_04a_03d  Schedule_Minimum_A;Schedule_Name_A;ScheduleDetail_Amount_D;scheddets_order_D;ScheduleDetail_Rate_D;

        'rptListOfFeeSchedules_04d_01a  Schedule_Minimum_D;Schedule_Name_D;scheddets_order_A;ScheduleDetail_Rate_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_04d_01d  Schedule_Minimum_D;Schedule_Name_D;scheddets_order_D;ScheduleDetail_Rate_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_04d_02a  Schedule_Minimum_D;Schedule_Name_D;ScheduleDetail_Rate_A;scheddets_order_A;ScheduleDetail_Amount_A;
        'rptListOfFeeSchedules_04d_02d  Schedule_Minimum_D;Schedule_Name_D;ScheduleDetail_Rate_D;scheddets_order_D;ScheduleDetail_Amount_D;
        'rptListOfFeeSchedules_04d_03a  Schedule_Minimum_D;Schedule_Name_D;ScheduleDetail_Amount_A;scheddets_order_A;ScheduleDetail_Rate_A;
        'rptListOfFeeSchedules_04d_03d  Schedule_Minimum_D;Schedule_Name_D;ScheduleDetail_Amount_D;scheddets_order_D;ScheduleDetail_Rate_D;
        'DONE!

20920   Beep

20930   Debug.Print "'DONE!"
20940   DoEvents

EXITP:
20950   Set ctl1 = Nothing
20960   Set ctl2 = Nothing
20970   Set rpt1 = Nothing
20980   Set rpt2 = Nothing
20990   Set doc = Nothing
21000   Set cntr = Nothing
21010   Set rst = Nothing
21020   Set qdf = Nothing
21030   Set dbs = Nothing
21040   Rpt_FeeSchedules = blnRetVal
21050   Exit Function

ERRH:
21060   blnRetVal = False
21070   Select Case ERR.Number
        Case Else
21080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
21090   End Select
21100   Resume EXITP

End Function

Public Function Rpt_CtlLabels() As Boolean

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "Rpt_CtlLabels"

        Dim rpt As Access.Report, ctl As Access.Control
        Dim blnRetVal As Boolean

21210 On Error GoTo 0

21220   blnRetVal = True

21230   Set rpt = Reports(0)

21240   With rpt
21250     For Each ctl In .Detail.Controls
21260       With ctl
21270         If .ControlType = acLabel Then
21280           If Left(.Name, 5) = "Label" Then
21290             .Name = .Parent.Name & "_lbl"
21300           End If
21310         End If
21320       End With
21330     Next
21340   End With

21350   Beep

21360   Debug.Print "'DONE!"
21370   DoEvents

EXITP:
21380   Set ctl = Nothing
21390   Set rpt = Nothing
21400   Rpt_CtlLabels = blnRetVal
21410   Exit Function

ERRH:
21420   blnRetVal = False
21430   Select Case ERR.Number
        Case Else
21440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
21450   End Select
21460   Resume EXITP

End Function
