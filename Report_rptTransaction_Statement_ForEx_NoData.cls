VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptTransaction_Statement_ForEx_NoData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptTransaction_Statement_ForEx_NoData"

'VGC 10/03/2015: CHANGES!

'gstrFormQuerySpec = "frmStatementParameters"

Private Const strCallingForm As String = "frmStatementParameters"

Private strThisProc As String
Private varTmp00 As Variant
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnSkip As Boolean
        Dim lngZ As Long

110     strThisProc = THIS_PROC

120     With Me
130       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

140         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
150         blnSkip = False

160         varTmp00 = .OpenArgs
170         If IsNull(varTmp00) = False Then
180           varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
190           varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
200           If CLng(datStartDate) = 0 Then
210             .Caption = "Transaction_Statement_" & varTmp00 & "_" & Format(datEndDate, "yymmdd")
220           Else
230             .Caption = "Transaction_Statement_" & varTmp00 & "_" & Format(datStartDate, "yymmdd") & "_To_" & Format(datEndDate, "yymmdd")
240           End If
250         End If

260         Select Case Forms(strCallingForm).chkPageOf
            Case True
270           .PageNumOf.Visible = True
280           .PageNum.Visible = False
290         Case False
300           .PageNum.Visible = True
310           .PageNumOf.Visible = False
320           .PageNumOf.ControlSource = vbNullString
330         End Select

340         .Name_lbl.Caption = vbNullString

350         Select Case Forms("frmStatementParameters").chkTransactions
            Case True
360           datStartDate = CDate(Forms("frmStatementParameters").TransDateStart)
370           datEndDate = CDate(Forms("frmStatementParameters").TransDateEnd)
380           .AsOf_lbl.Caption = "From " & Format(datStartDate, "mm/dd/yyyy") & " To " & Format(datEndDate, "mm/dd/yyyy")
390           .AsOf_lbl2.Caption = "From " & Format(datStartDate, "mm/dd/yyyy") & " To " & Format(datEndDate, "mm/dd/yyyy")
400           Select Case Forms("frmStatementParameters").chkArchiveOnly_Trans
              Case True
410             .AsOf.Visible = True
420             .AsOf2.Visible = True
430           Case False
440             .AsOf.Visible = False
450             .AsOf2.Visible = False
460           End Select  ' ** chkArchiveOnly_Trans.
470         Case False
480           datStartDate = 0
490           datEndDate = CDate(Forms("frmStatementParameters").DateEnd)
500           .AsOf_lbl.Caption = "For Period Ending " & Format(datEndDate, "mmmm dd, yyyy")
510           .AsOf_lbl2.Caption = "For Period Ending " & Format(datEndDate, "mmmm dd, yyyy")
520           .AsOf.Visible = False
530           .AsOf2.Visible = False
540         End Select  ' ** chkTransactions.

550         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
560         .CoInfoBlock2.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

570         Select Case Forms("frmStatementParameters").chkCombineCash
            Case True
580           .icash_lbl.Visible = False
590           .icash_lbl_line.Visible = False
600           .pcash_lbl.Visible = False
610           .pcash_lbl_line.Visible = False
620           .cash_lbl.Visible = True
630           .cash_lbl_line.Visible = True
640           .icash_usd_lbl.Visible = False
650           .icash_usd_lbl_line.Visible = False
660           .pcash_usd_lbl.Visible = False
670           .pcash_usd_lbl_line.Visible = False
680           .cash_usd_lbl.Visible = True
690           .cash_usd_lbl_line.Visible = True
700         Case False
710           .cash_lbl.Visible = False
720           .cash_lbl_line.Visible = False
730           .icash_lbl.Visible = True
740           .icash_lbl_line.Visible = True
750           .pcash_lbl.Visible = True
760           .pcash_lbl_line.Visible = True
770           .cash_usd_lbl.Visible = False
780           .cash_usd_lbl_line.Visible = False
790           .icash_usd_lbl.Visible = True
800           .icash_usd_lbl_line.Visible = True
810           .pcash_usd_lbl.Visible = True
820           .pcash_usd_lbl_line.Visible = True
830         End Select

840         If glngPrintRpts > 0& Then
850           For lngZ = 0& To (glngPrintRpts - 1&)
860             If garr_varPrintRpt(PR_ACTNO, lngZ) = gstrAccountNo Then
870               If garr_varPrintRpt(PR_TRANS, lngZ) = True Then
880                 blnSkip = True
890               Else
900                 garr_varPrintRpt(PR_TRANS, lngZ) = CBool(True)
910               End If
920               Exit For
930             End If
940           Next
950         End If

960       Else
970         Cancel = -1
980       End If
990     End With

1000    If Cancel = -1 Then
1010      Beep
1020      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
1030    ElseIf blnSkip = True Then
1040      Cancel = -1
1050    End If

EXITP:
1060    Exit Sub

ERRH:
1070    Select Case ERR.Number
        Case Else
1080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1090    End Select
1100    Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

        Dim varTmp00 As Variant

1210    strThisProc = THIS_PROC

1220    With Me
1230      Select Case Forms(strCallingForm).chkLegalName
          Case True
1240        varTmp00 = FixAmps(Nz(.legalname, vbNullString))  ' ** Module Functions: modStringFuncs.
1250        If IsNull(varTmp00) = True Then
1260          varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
1270        End If
1280        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
1290        .Name_lbl.Caption = varTmp00
1300      Case False
1310        varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
1320        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
1330        .Name_lbl.Caption = varTmp00
1340      End Select  ' ** chkLegalName.
1350    End With

EXITP:
1360    Exit Sub

ERRH:
1370    Select Case ERR.Number
        Case Else
1380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1390    End Select
1400    Resume EXITP

End Sub

Private Sub Report_Close()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1510    strThisProc = THIS_PROC

1520    If strCallingForm <> vbNullString Then
1530      gblnReportClose = True
1540      Forms(strCallingForm).TimerInterval = 250&
1550    End If

EXITP:
1560    Exit Sub

ERRH:
1570    Select Case ERR.Number
        Case Else
1580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1590    End Select
1600    Resume EXITP

End Sub
