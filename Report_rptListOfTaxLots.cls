VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfTaxLots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfTaxLots"

'VGC 03/23/2016: CHANGES!

Private strCallingForm As String
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strOrderBy As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

              ' ** Default to PageNumOf if calling form has no check box.
160           .PageNumOf.Visible = True
170           .PageNum.Visible = False

180           .AsOf.Caption = "As of " & Format(Date, "mmmm d, yyyy")
190           .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

200           Select Case strCallingForm
              Case "frmTaxLot"
210             strOrderBy = Forms(strCallingForm).frmTaxLot_Sub.Form.SortNow_Get  ' ** Form Function: frmTaxLot_Sub.
220           Case "frmJournal_Columns_TaxLot"
230             strOrderBy = Forms(strCallingForm).frmJournal_Columns_TaxLot_Sub.Form.SortNow_Get  ' ** Form Function: frmJournal_Columns_TaxLot_Sub.
240           End Select
              ' ** Potential Sorts:
              ' **   [assetdate]
              ' **   [shareface], [assetdate]
              ' **   [cost], [assetdate]
              ' **   [cost_usd], [assetdate]
              ' **   [priceperunit], [assetdate]
              ' **   [averagepriceperunit], [assetdate]
              ' **   [curr_code], [assetdate]
              ' **   [Holding_Period], [assetdate]
              ' **   [Loc_Name], [assetdate]
250           .OrderBy = strOrderBy
260           .OrderByOn = True

270         Else
280           Cancel = -1
290         End If
300       Else
310         Cancel = -1
320       End If
330     End With

340     If Cancel = -1 Then
350       Beep
360       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
370     End If

EXITP:
380     Exit Sub

ERRH:
390     If Forms(strCallingForm).Visible = False Then
400       DoCmd.Restore
410       Select Case strCallingForm
          Case "frmTaxLot"
420         Forms("frmJournal").Visible = True
430         Forms(strCallingForm).Visible = True
440       Case "frmJournal_Columns_TaxLot"
450         Forms("frmJournal_Columns").Visible = True
460         Forms(strCallingForm).Visible = True
470       End Select
480     End If
490     Select Case ERR.Number
        Case Else
500       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
510     End Select
520     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

610     With Me
620       If .IsAverage = True Then
630         .averagepriceperunit.Visible = True
640         .priceperunit.Visible = False
650       End If
660     End With

EXITP:
670     Exit Sub

ERRH:
680     If Forms(strCallingForm).Visible = False Then
690       DoCmd.Restore
700       Select Case strCallingForm
          Case "frmTaxLot"
710         Forms("frmJournal").Visible = True
720         Forms(strCallingForm).Visible = True
730       Case "frmJournal_Columns_TaxLot"
740         Forms("frmJournal_Columns").Visible = True
750         Forms(strCallingForm).Visible = True
760       End Select
770     End If
780     Select Case ERR.Number
        Case Else
790       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
800     End Select
810     Resume EXITP

End Sub

Private Sub Report_Close()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"


910     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

920     If Forms(strCallingForm).Visible = False Then
930       DoCmd.Restore
940       Select Case strCallingForm
          Case "frmTaxLot"
950         Forms("frmJournal").Visible = True
960         Forms(strCallingForm).Visible = True
970       Case "frmJournal_Columns_TaxLot"
980         Forms("frmJournal_Columns").Visible = True
990         Forms(strCallingForm).Visible = True
1000      End Select
1010    End If

EXITP:
1020    Exit Sub

ERRH:
1030    If Forms(strCallingForm).Visible = False Then
1040      DoCmd.Restore
1050      Select Case strCallingForm
          Case "frmTaxLot"
1060        Forms("frmJournal").Visible = True
1070        Forms(strCallingForm).Visible = True
1080      Case "frmJournal_Columns_TaxLot"
1090        Forms("frmJournal_Columns").Visible = True
1100        Forms(strCallingForm).Visible = True
1110      End Select
1120    End If
1130    Select Case ERR.Number
        Case Else
1140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1150    End Select
1160    Resume EXITP

End Sub
