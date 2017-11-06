VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptTransactions_ByType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptTransactions_ByType"

'VGC 04/02/2015: CHANGES!

Private Const strCallingForm As String = "frmRpt_TransactionsByType"

Private blnOnce As Boolean
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strAccountNo As String, strPeriod As String
        Dim strFilter As String
        Dim strTmp01 As String
        Dim lngX As Long

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Functions: modFileUtilities.

130         If blnOnce = False Then
140           blnOnce = True
150           TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
160         End If

170         Select Case Forms(strCallingForm).chkIncludeArchive
            Case True
180           If .RecordSource <> "qryTransactions_ByType_04_02" Then
190             .RecordSource = "qryTransactions_ByType_04_02"
200           End If
210         Case False
220           If .RecordSource <> "qryTransactions_ByType_04_01" Then
230             .RecordSource = "qryTransactions_ByType_04_01"
240           End If
250         End Select

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

350         If Forms(strCallingForm).ReportMode = "Word" Then
              ' ** Add the TransDate and JournalType filtering here.
360           strFilter = "([transdate] >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And " & _
                "[transdate] <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#) And ("
370           For lngX = 1& To 12&
380             strTmp01 = FormRef(CStr(lngX))
390             If strTmp01 <> vbNullString Then
400               strFilter = strFilter & "([journaltypex] = '" & strTmp01 & "') Or "
410             End If
420           Next
430           strFilter = Trim(strFilter)
440           If Right(strFilter, 2) = "Or" Then strFilter = Left(strFilter, (Len(strFilter) - 2))
450           strFilter = Trim(strFilter)
460           strFilter = strFilter & ")"
470           .Filter = strFilter
480           .FilterOn = True
490         End If

500         .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
510         .AsOf_lbl2.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")

520         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
530         .CoInfoBlock2.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

540         strAccountNo = Forms(strCallingForm).cmbAccounts
550         Select Case Forms(strCallingForm).chkTransDateAll
            Case True
560           strPeriod = "All"
570         Case False
580           strPeriod = Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")
590         End Select

600         .Caption = "Transactions_By_Type_" & strAccountNo & "_" & strPeriod

610       Else
620         Cancel = -1
630       End If
640     End With

650     If Cancel = -1 Then
660       Beep
670       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
680     End If

EXITP:
690     Exit Sub

ERRH:
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

810     With Me
820       .Name_lbl.Caption = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
830     End With

EXITP:
840     Exit Sub

ERRH:
850     Select Case ERR.Number
        Case Else
860       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
870     End Select
880     Resume EXITP

End Sub

Private Sub GroupHeader3_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader3_Format"

910     With Me
920       .journaltypex_lbl.Caption = Forms(strCallingForm).JType_Get(.journaltypex)  ' ** Form Function: frmRpt_TransactionsByType.
930     End With

EXITP:
940     Exit Sub

ERRH:
950     Select Case ERR.Number
        Case Else
960       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
970     End Select
980     Resume EXITP

End Sub

Private Sub Report_Close()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1010    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1020    Exit Sub

ERRH:
1030    Select Case ERR.Number
        Case Else
1040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1050    End Select
1060    Resume EXITP

End Sub
