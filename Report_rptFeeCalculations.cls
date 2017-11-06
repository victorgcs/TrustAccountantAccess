VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptFeeCalculations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptFeeCalculations"

'VGC 07/03/2016: CHANGES!

Private Const strCallingForm As String = "frmFeeCalculations"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strMonth As String, strYear As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

140         Select Case Forms(strCallingForm).chkPageOf
            Case True
150           .PageNumOf.Visible = True
160           .PageNum.Visible = False
170         Case False
180           .PageNum.Visible = True
190           .PageNumOf.Visible = False
200           .PageNumOf.ControlSource = vbNullString
210         End Select

220         Select Case Forms(strCallingForm).HasForEx
            Case True
230           .ForExMsg_lbl.Visible = True
240         Case False
250           .ForExMsg_lbl.Visible = False
260         End Select

270         strMonth = Forms(strCallingForm).FeeMonth
280         strYear = Forms(strCallingForm).FeeYear

290         .AsOf.Caption = strMonth & ", " & strYear
300         .AsOf2.Caption = strMonth & ", " & strYear
310         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
320         .CoInfoBlock2.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

330         .Caption = "Fee_Calculation_" & Forms(strCallingForm).FeeMonth.Column(2) & Right(strYear, 2)

340         If gstrReportCallingForm <> vbNullString Then
350           Forms(gstrReportCallingForm).Visible = False
360         End If

370       Else
380         Cancel = -1
390       End If
400     End With

410     If Cancel = -1 Then
420       Beep
430       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
440     End If

EXITP:
450     Exit Sub

ERRH:
460     If gstrReportCallingForm <> vbNullString Then
470       DoCmd.Restore
480       Forms(gstrReportCallingForm).Visible = True
490       DoCmd.SelectObject acForm, gstrReportCallingForm, False
500     End If
510     Select Case ERR.Number
        Case Else
520       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
530     End Select
540     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

610     With Me
630       If .Page >= 2 Then
640         .Section(acPageHeader).Visible = True
650       Else
660         .Section(acPageHeader).Visible = False
670       End If
680     End With

EXITP:
690     Exit Sub

ERRH:
700     If gstrReportCallingForm <> vbNullString Then
710       DoCmd.Restore
720       Forms(gstrReportCallingForm).Visible = True
730       DoCmd.SelectObject acForm, gstrReportCallingForm, False
740     End If
750     Select Case ERR.Number
        Case Else
760       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
770     End Select
780     Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

        ' ** This may not work!
810     With Me
820       MakeTopMost .hwnd  ' ** Module Function: modWindowFunctions.
830     End With

EXITP:
840     Exit Sub

ERRH:
850     If gstrReportCallingForm <> vbNullString Then
860       DoCmd.Restore
870       Forms(gstrReportCallingForm).Visible = True
880       DoCmd.SelectObject acForm, gstrReportCallingForm, False
890     End If
900     Select Case ERR.Number
        Case Else
910       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
920     End Select
930     Resume EXITP

End Sub

Private Sub Report_Close()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1010    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1020    DoCmd.Restore

1030    If gstrReportCallingForm <> vbNullString Then
1040      Forms(gstrReportCallingForm).Visible = True
1050    End If

EXITP:
1060    Exit Sub

ERRH:
1070    If gstrReportCallingForm <> vbNullString Then
1080      DoCmd.Restore
1090      Forms(gstrReportCallingForm).Visible = True
1100      DoCmd.SelectObject acForm, gstrReportCallingForm, False
1110    End If
1120    Select Case ERR.Number
        Case Else
1130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1140    End Select
1150    Resume EXITP

End Sub
