VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptTaxLot_One_SortType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptTaxLot_One_SortType"

'VGC 05/29/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_TaxLot"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

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

220         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
230         .AsOf_lbl2.Caption = "As of " & Format(Now(), "mm/dd/yyyy")
240         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
250         .Caption = "Tax_Lots_" & gstrAccountNo & "_" & Format(Date, "yymmdd")

260       Else
270         Cancel = -1
280       End If
290     End With

300     If Cancel = -1 Then
310       Beep
320       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
330     End If

EXITP:
340     Exit Sub

ERRH:
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

400   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

410     With Me
420       If .Page >= 2 Then
430         .Section(acPageHeader).Visible = True
440       Else
450         .Section(acPageHeader).Visible = False
460       End If
470       If .IsAverage = True Then
480         .averagepriceperunit.Visible = True
490         .priceperunit.Visible = False
500       End If
510     End With

EXITP:
520     Exit Sub

ERRH:
530     Select Case ERR.Number
        Case Else
540       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
550     End Select
560     Resume EXITP

End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Format"

610     With Me
620       Select Case .IsAverage
          Case True
630         .averagepriceperunit.Visible = True
640         .priceperunit.Visible = False
650       Case False
660         .averagepriceperunit.Visible = False
670         .priceperunit.Visible = True
680       End Select
690     End With

EXITP:
700     Exit Sub

ERRH:
710     Select Case ERR.Number
        Case Else
720       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
730     End Select
740     Resume EXITP

End Sub

Private Sub Report_Close()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

810     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
820     Exit Sub

ERRH:
830     Select Case ERR.Number
        Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
850     End Select
860     Resume EXITP

End Sub
