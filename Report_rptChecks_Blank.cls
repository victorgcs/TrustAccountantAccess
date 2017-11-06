VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptChecks_Blank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptChecks_Blank"

'VGC 07/24/2017: CHANGES!

' **********************************************
' ** NOTE: Popup required for proper printing!
' **********************************************

' ** The original values are negative, from the 'Paid' transaction.
' ** Amount1:
' **   =[Amount]*-1
' ** AmountText1:
' **   =GetDollarString([Amount]*-1)

Private Const strCallingForm As String = "frmRpt_Checks"
Private Const strCallingForm2 As String = "frmRpt_Checks_MICR_Adjust"

Private strAccountNo As String
Private intMode As Integer
Private strTmpCallingForm As String
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim frm As Access.Form
        Dim intScreenFontSize As Integer
        Dim strSortOrd As String, blnDesc As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Or IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Function: modFileUtilities.
130         If IsNull(.OpenArgs) = False Then
140           strTmp01 = Trim(.OpenArgs)
150           If strTmp01 <> vbNullString Then
160             intPos01 = InStr(strTmp01, "~")
170             If intPos01 > 0 Then

                  ' ** Sort order options:
                  ' **   Account_Number
                  ' **   Short_Name
                  ' **   Payee
                  ' **   Bank_Name
                  ' **   Bank_Account_Number

180               strSortOrd = Left(strTmp01, (intPos01 - 1))
190               strTmp01 = Mid(strTmp01, (intPos01 + 1))
200               intPos01 = InStr(strTmp01, "~")
210               blnDesc = CBool(Left(strTmp01, (intPos01 - 1)))  ' ** Descending.
220               strTmp01 = Mid(strTmp01, (intPos01 + 1))
230               intPos01 = InStr(strTmp01, "~")
240               intMode = CInt(Left(strTmp01, (intPos01 - 1)))
250               strAccountNo = Mid(strTmp01, (intPos01 + 1))

260               If IsLoaded(strCallingForm, acForm) = True Then
270                 Set frm = Forms(strCallingForm)
280                 If frm.opgPayeeFont.Visible = True Then
290                   Select Case frm.opgPayeeFont
                      Case frm.opgPayeeFont_opt10Pt.OptionValue
300                     .RecurringItem1.Visible = True
310                     .RecurringItem1_12pt.Visible = False
320                   Case frm.opgPayeeFont_opt12Pt.OptionValue
330                     .RecurringItem1.Visible = False
340                     .RecurringItem1_12pt.Visible = True
350                   End Select
360                 Else
370                   .RecurringItem1.Visible = True
380                   .RecurringItem1_12pt.Visible = False
390                 End If
400                 Set frm = Nothing
410               Else
420                 .RecurringItem1.Visible = True
430                 .RecurringItem1_12pt.Visible = False
440               End If

450               If strSortOrd = "Test_Check" Then
460                 strSortOrd = "Short_Name"
470                 strTmpCallingForm = strCallingForm2
480                 .RecordSource = "qryRpt_Checks_MICR_Adjust_05"
490                 .TestCheck_lbl1a.Visible = True
500                 .TestCheck_lbl1b.Visible = True
510                 .TestCheck_lbl2a.Visible = True
520                 .TestCheck_lbl2b.Visible = True
530                 .TestCheck_lbl3a.Visible = True
540                 .TestCheck_lbl3b.Visible = True
550               Else
560                 strTmpCallingForm = strCallingForm
570                 .TestCheck_lbl1a.Visible = False
580                 .TestCheck_lbl1b.Visible = False
590                 .TestCheck_lbl2a.Visible = False
600                 .TestCheck_lbl2b.Visible = False
610                 .TestCheck_lbl3a.Visible = False
620                 .TestCheck_lbl3b.Visible = False
630               End If

640               TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

650               .Caption = "Checks_" & Format(Date, "yyyymmdd") & "_" & strAccountNo

660               blnRetVal = Rpt_GroupLevel_Set(Me, strSortOrd, blnDesc)  ' ** Module Function: modReportFunctions.

670               If gstrCo_Name = vbNullString Then
680                 CoOptions_Read  ' ** Module Function: modStartupFuncs.
690               End If

700               .txtMICR.FontName = Forms(strTmpCallingForm).MICRFontName

                  ' ** Get the stored offset from DDTrust.INI.
                  ' ** Formula is new position = store position - ((rpt font size for the field * screen offset in twips) / screen font size for the adjustment character)
710               intScreenFontSize = IniFile_Get("CheckMICROffset", "ScreenFontSize", RET_ERR, CurrentAppPath & LNK_SEP & gstrFile_INI)  ' ** Module Function: modStartupFuncs, modFileUtilities.
720               If intScreenFontSize <> 0 Then
730                 .txtMICR.Top = .txtMICR.Top - _
                      (.txtMICR.FontSize * IniFile_Get("CheckMICROffset", "Top", RET_ERR, CurrentAppPath & LNK_SEP & gstrFile_INI)) / intScreenFontSize  ' ** Module Function: modStartupFuncs, modFileUtilities.
740                 .txtMICR.Left = .txtMICR.Left - _
                      (.txtMICR.FontSize * IniFile_Get("CheckMICROffset", "left", RET_ERR, CurrentAppPath & LNK_SEP & gstrFile_INI)) / intScreenFontSize  ' ** Module Function: modStartupFuncs, modFileUtilities.
750               End If

760             Else
770               Cancel = -1
780             End If
790           Else
800             Cancel = -1
810           End If
820         Else
830           Cancel = -1
840         End If
850       Else
860         Cancel = -1
870       End If

880     End With

890     If Cancel = -1 Then
900       Beep
910       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
920     End If

EXITP:
930     Set frm = Nothing
940     Exit Sub

ERRH:
950     If Forms(strTmpCallingForm).Visible = False Then
960       Forms(strTmpCallingForm).Visible = True
970     End If
980     Select Case ERR.Number
        Case 2001  ' ** You canceled the previous operation.
990       Cancel = -1
1000    Case Else
1010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1020    End Select
1030    Resume EXITP

End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Format"

        Dim strPayeeCSZ As String

1110    With Me

          ' ** AmountText1: =GetDollarString([Amount]*-1)

1120      .txtCompanyName = Trim(gstrCo_Name)
1130      .txtAddress = Trim(gstrCo_Address1) & vbCrLf
1140      If Len(Trim(gstrCo_Address2)) > 0 Then  ' ** Check to see if second line exists.
1150        .txtAddress = .txtAddress & Trim(gstrCo_Address2) & vbCrLf
1160      End If
1170      If gstrCo_Country = vbNullString Then
1180        .txtAddress = .txtAddress & Trim(gstrCo_City & ", " & gstrCo_State & " " & FormatZip(gstrCo_Zip))  ' ** Module Function: modStringFuncs.
1190      Else
1200        .txtAddress = .txtAddress & gstrCo_City & vbCrLf & Trim(gstrCo_Country & "  " & gstrCo_PostalCode)
1210      End If

1220      If IsNull(.TxtFedIFNum1) = True Or IsNull(.txtFedFiNum2) = True Then
1230        .TxtFedIFNum_lbl.Visible = False
1240      Else
1250        If Trim(.TxtFedIFNum1) = vbNullString Or Trim(.txtFedFiNum2) = vbNullString Then
1260          .TxtFedIFNum_lbl.Visible = False
1270        End If
1280      End If

1290      .txtMICR = FormatMICRString  ' ** Function: Below.

1300      If IsNull(.txtPayeeAddress) = False Then
1310        If Trim(.txtPayeeAddress) <> vbNullString Then
1320          If IsNull(.RecurringItem_Countryx) = True And IsNull(.RecurringItem_PostalCodex) = True Then
1330            strPayeeCSZ = Trim(IIf(IsNull(.RecurringItem_Cityx) = True, vbNullString, _
                  IIf(Trim(.RecurringItem_Cityx) = vbNullString, vbNullString, .RecurringItem_Cityx & ", ")) & _
                  Nz(.RecurringItem_Statex, vbNullString) & " " & FormatZip(Nz(.RecurringItem_Zipx, vbNullString)))
1340          Else
1350            If Trim(.RecurringItem_Countryx) = vbNullString And Trim(.RecurringItem_PostalCodex) = vbNullString Then
1360              strPayeeCSZ = Trim(IIf(IsNull(.RecurringItem_Cityx) = True, vbNullString, _
                    IIf(Trim(.RecurringItem_Cityx) = vbNullString, vbNullString, .RecurringItem_Cityx & ", ")) & _
                    Nz(.RecurringItem_Statex, vbNullString) & " " & FormatZip(Nz(.RecurringItem_Zipx, vbNullString)))
1370            Else
1380              strPayeeCSZ = Trim(IIf(IsNull(.RecurringItem_Cityx) = True, vbNullString, _
                    IIf(Trim(.RecurringItem_Cityx) = vbNullString, vbNullString, .RecurringItem_Cityx & ", ")) & _
                    Nz(.RecurringItem_Countryx, vbNullString) & "  " & Nz(.RecurringItem_PostalCodex, vbNullString))
1390            End If
1400          End If
1410        Else
1420          strPayeeCSZ = vbNullString
1430        End If
1440      Else
1450        strPayeeCSZ = vbNullString
1460      End If

1470      If Trim(strPayeeCSZ) = "," Then
1480        .txtPayeeCSZ = vbNullString
1490      ElseIf Trim(strPayeeCSZ) = vbNullString Then
1500        .txtPayeeCSZ = vbNullString
1510      Else
1520        .txtPayeeCSZ = strPayeeCSZ
1530      End If

1540    End With

EXITP:
1550    Exit Sub

ERRH:
1560    If Forms(strTmpCallingForm).Visible = False Then
1570      Forms(strTmpCallingForm).Visible = True
1580    End If
1590    Select Case ERR.Number
        Case Else
1600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1610    End Select
1620    Resume EXITP

End Sub

Private Sub Report_Close()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1710    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1720    If Forms(strTmpCallingForm).Visible = False Then
1730      Forms(strTmpCallingForm).Visible = True
1740    End If

1750    If intMode = acViewPreview Then
1760      gblnSetFocus = True
1770      Forms(strTmpCallingForm).TimerInterval = 250&
1780    End If

EXITP:
1790    Exit Sub

ERRH:
1800    If Forms(strTmpCallingForm).Visible = False Then
1810      Forms(strTmpCallingForm).Visible = True
1820    End If
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Function FormatMICRString() As String

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "FormatMICRString"

        Dim intCheck As Integer
        Dim strCheck As String
        Dim intBank As Integer
        Dim strBank As String
        Dim intX As Integer
        Dim strRetVal As String

1910    strRetVal = vbNullString

1920    With Me

1930      If IsNull(.CheckNum) = True Then
1940        strCheck = String(6, "0")
1950      Else
1960        If Len(.CheckNum) < 6 Then
1970          intCheck = 6 - Len(.CheckNum)
1980          strCheck = .CheckNum
1990          For intX = 1 To intCheck
2000            strCheck = "0" & strCheck
2010          Next
2020        Else
2030          strCheck = .CheckNum
2040        End If
2050      End If

2060      If IsNull(.Bank_RoutingNumber) = True Then
2070        strBank = String(9, "0")
2080      Else
2090        If Len(.Bank_RoutingNumber) < 9 Then
2100          intBank = 9 - Len(.Bank_RoutingNumber)
2110          strBank = .Bank_RoutingNumber
2120          For intX = 1 To intBank
2130            strBank = "0" & strBank
2140          Next
2150        Else
2160          strBank = .Bank_RoutingNumber
2170        End If
2180      End If

2190      strRetVal = "C" & strCheck & "C " & "A" & CStr(Nz(strBank, vbNullString)) & "A " & _
            FormatDashes(CStr(Nz(.Bank_AccountNumber, vbNullString))) & "C"  ' ** Function: Below.

2200    End With

EXITP:
2210    FormatMICRString = strRetVal
2220    Exit Function

ERRH:
2230    If Forms(strTmpCallingForm).Visible = False Then
2240      Forms(strTmpCallingForm).Visible = True
2250    End If
2260    strRetVal = vbNullString
2270    Select Case ERR.Number
        Case Else
2280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2290    End Select
2300    Resume EXITP

End Function

Private Function FormatDashes(strAccountNum As String) As String

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "FormatDashes"

        Dim lngCnt As Long
        Dim lngLen As Long
        Dim strBuild As String
        Dim strChar As String
        Dim strRetVal As String

2410    strRetVal = vbNullString

2420    lngLen = Len(strAccountNum)
2430    If lngLen = 0 Then
2440      strRetVal = strAccountNum
2450    Else
2460      For lngCnt = 1 To lngLen
2470        strChar = Mid(strAccountNum, lngCnt, 1)
2480        If strChar = "-" Then
2490          strBuild = strBuild & "="
2500        Else
2510          strBuild = strBuild & strChar
2520        End If
2530      Next lngCnt
2540    End If

2550    strRetVal = strBuild

EXITP:
2560    FormatDashes = strRetVal
2570    Exit Function

ERRH:
2580    If Forms(strTmpCallingForm).Visible = False Then
2590      Forms(strTmpCallingForm).Visible = True
2600    End If
2610    strRetVal = strAccountNum
2620    Select Case ERR.Number
        Case Else
2630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2640    End Select
2650    Resume EXITP

End Function
