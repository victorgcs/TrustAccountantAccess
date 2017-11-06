VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAssetList_ForEx_NoData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAssetList_ForEx_NoData"

'VGC 03/04/2016: CHANGES!

Private Const strCallingForm As String = "frmStatementParameters"

Private datStartDate As Date, datEndDate As Date
Private frm As Access.Form
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim blnSkip As Boolean
        Dim varTmp00 As Variant
        Dim lngZ As Long

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
140         blnSkip = False

150         Set frm = Forms(strCallingForm)

160         Select Case frm.chkPageOf
            Case True
170           .PageNumOf.Visible = True
180           .PageNum.Visible = False
190         Case False
200           .PageNum.Visible = True
210           .PageNumOf.Visible = False
220           .PageNumOf.ControlSource = vbNullString
230         End Select

240         .Name_lbl.Caption = vbNullString

250         varTmp00 = frm.cmbAccounts
260         If IsNull(varTmp00) = False Then
270           varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
280           varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
290           .Caption = "Asset_List_" & varTmp00 & "_" & Format(CDate(frm.AssetListDate), "yymmdd")
300         Else
310           .Caption = "Asset_List_All_" & Format(CDate(frm.AssetListDate), "yymmdd")
320         End If
330         datStartDate = frm.DateStart
340         datEndDate = frm.AssetListDate

350         .AsOf_lbl.Caption = "As of " & Format(CDate(frm.AssetListDate), "mmmm dd, yyyy")
360         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
370         .AsOf.Caption = "Printed " & Format(Now(), "mmmm dd"", ""yyyy")
380         .AsOf_lbl2.Caption = "As of " & Format(CDate(frm.AssetListDate), "mmmm dd, yyyy")
390         .CoInfoBlock2.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
400         .AsOf2.Caption = "Printed " & Format(Now(), "mmmm dd"", ""yyyy")

410         If glngPrintRpts > 0& Then
420           For lngZ = 0& To (glngPrintRpts - 1&)
430             If garr_varPrintRpt(PR_ACTNO, lngZ) = gstrAccountNo Then
440               If garr_varPrintRpt(PR_ALIST, lngZ) = True Then
450                 blnSkip = True
460               Else
470                 garr_varPrintRpt(PR_ALIST, lngZ) = CBool(True)
480               End If
490               Exit For
500             End If
510           Next
520         End If

530       Else
540         Cancel = -1
550       End If
560     End With

570     If Cancel = -1 Then
580       Beep
590       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
600     ElseIf blnSkip = True Then
610       Cancel = -1
620     End If

EXITP:
630     Exit Sub

ERRH:
640     Select Case ERR.Number
        Case 2585  ' ** This action can't be carried out while processing a form or report event.
          ' ** Ignore.
650     Case Else
660       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
670     End Select
680     Resume EXITP

End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "ReportHeader_Format"

        Dim datCurrentDate As Date
        Dim varTmp00 As Variant

710     With Me

720       Select Case IsNull(frm.currentDate)
          Case True
730         varTmp00 = DMax("[currentDate]", "masterasset")
740         datCurrentDate = varTmp00
750       Case False
760         datCurrentDate = frm.currentDate
770       End Select
780       If frm.UsePriceHistory = True Then
790         .MarketAsOf_lbl.Caption = "Market as of varies"
800         .MarketAsOf_lbl2.Caption = "Market as of varies"
810       Else
820         If datCurrentDate <> datEndDate Then
830           .MarketAsOf_lbl.Caption = "Market as of " & Format(datCurrentDate, "mm/dd/yyyy")
840           .MarketAsOf_lbl2.Caption = "Market as of " & Format(datCurrentDate, "mm/dd/yyyy")
850         Else
860           .MarketAsOf_lbl.Caption = vbNullString
870           .MarketAsOf_lbl2.Caption = vbNullString
880         End If
890       End If

900     End With

EXITP:
910     Exit Sub

ERRH:
920     Select Case ERR.Number
        Case Else
930       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
940     End Select
950     Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

        Dim varTmp00 As Variant

1010    With Me
1020      Select Case frm.chkLegalName
          Case True
1030        varTmp00 = FixAmps(Nz(.legalname, vbNullString))  ' ** Module Functions: modStringFuncs.
1040        If IsNull(varTmp00) = True Then
1050          varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
1060        End If
1070        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
1080        .Name_lbl.Caption = varTmp00
1090      Case False
1100        varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
1110        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
1120        .Name_lbl.Caption = varTmp00
1130      End Select  ' ** chkLegalName.
1140      .NoData_lbl.BackColor = CLR_WHT
1150    End With

EXITP:
1160    Exit Sub

ERRH:
1170    Select Case ERR.Number
        Case Else
1180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1190    End Select
1200    Resume EXITP

End Sub

Private Sub Report_Close()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1310    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1320    Exit Sub

ERRH:
1330    Select Case ERR.Number
        Case Else
1340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1350    End Select
1360    Resume EXITP

End Sub
