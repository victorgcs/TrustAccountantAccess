VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptTransaction_Statement_NoData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptTransaction_Statement_NoData"

'VGC 09/25/2015: CHANGES!

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

570         If glngPrintRpts > 0& Then
580           For lngZ = 0& To (glngPrintRpts - 1&)
590             If garr_varPrintRpt(PR_ACTNO, lngZ) = gstrAccountNo Then
600               If garr_varPrintRpt(PR_TRANS, lngZ) = True Then
610                 blnSkip = True
620               Else
630                 garr_varPrintRpt(PR_TRANS, lngZ) = CBool(True)
640               End If
650               Exit For
660             End If
670           Next
680         End If

690       Else
700         Cancel = -1
710       End If
720     End With

730     If Cancel = -1 Then
740       Beep
750       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
760     ElseIf blnSkip = True Then
770       Cancel = -1
780     End If

EXITP:
790     Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

        Dim varTmp00 As Variant

910     strThisProc = THIS_PROC

920     With Me
930       Select Case Forms(strCallingForm).chkLegalName
          Case True
940         varTmp00 = FixAmps(Nz(.legalname, vbNullString))  ' ** Module Functions: modStringFuncs.
950         If IsNull(varTmp00) = True Then
960           varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
970         End If
980         If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
990         .Name_lbl.Caption = varTmp00
1000      Case False
1010        varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
1020        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
1030        .Name_lbl.Caption = varTmp00
1040      End Select  ' ** chkLegalName.
1050    End With

EXITP:
1060    Exit Sub

ERRH:
1070    Select Case ERR.Number
        Case Else
1080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1090    End Select
1100    Resume EXITP

End Sub

Private Sub Report_Close()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1210    strThisProc = THIS_PROC

1220    If strCallingForm <> vbNullString Then
1230      gblnReportClose = True
1240      Forms(strCallingForm).TimerInterval = 250&
1250    End If

EXITP:
1260    Exit Sub

ERRH:
1270    Select Case ERR.Number
        Case Else
1280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1290    End Select
1300    Resume EXITP

End Sub
