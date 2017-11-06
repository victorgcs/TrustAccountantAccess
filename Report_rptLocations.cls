VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptLocations"

'VGC 11/11/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_Locations"
Private lngLblLine_Offset As Long, lngTpp As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
140         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

150         Select Case Forms(strCallingForm).chkPageOf
            Case True
160           .PageNumOf.Visible = True
170           .PageNum.Visible = False
180         Case False
190           .PageNum.Visible = True
200           .PageNumOf.Visible = False
210           .PageNumOf.ControlSource = vbNullString
220         End Select

230         lngLblLine_Offset = .AccountName_lbl_line.Top - .AccountName_lbl.Top

240         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
250         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
260         .AsOf_lbl2.Caption = "As of " & Format(Date, "mmmm dd, yyyy")

270         .Caption = "Locations_" & Format(Date, "yymmdd")

280       Else
290         Cancel = -1
300       End If
310     End With

320     If Cancel = -1 Then
330       Beep
340       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
350     End If

EXITP:
360     Exit Sub

ERRH:
370     Select Case ERR.Number
        Case Else
380       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
390     End Select
400     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

510     With Me
520       If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
530         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
540       End If
550       If .Page >= 2 Then
560         lngTmp01 = ((.AsOf_lbl2.Height + .AccountName_lbl.Height) + (2& * lngTpp))
570         .Section(acPageHeader).Height = lngTmp01
580         lngTmp01 = (.AsOf_lbl2.Height + (2& * lngTpp))
590         .TotalShareface_lbl.Top = lngTmp01
600         .TotalShareface_lbl_line.Top = (.TotalShareface_lbl.Top + lngLblLine_Offset)
610         .TotalCost_lbl.Top = lngTmp01
620         .TotalCost_lbl_line.Top = (.TotalCost_lbl.Top + lngLblLine_Offset)
630         .AccountName_lbl.Top = lngTmp01
640         .AccountName_lbl_line.Top = (.AccountName_lbl.Top + lngLblLine_Offset)
650         .AsOf_lbl2.Visible = True
660         .Header_lbl2.Visible = True
670         .Header_hline05.Visible = True
            '.Section(acPageHeader).Visible = True
680       Else
690         .AsOf_lbl2.Visible = False
700         .Header_lbl2.Visible = False
710         .Header_hline05.Visible = False
720         .TotalShareface_lbl.Top = 0&
730         .TotalShareface_lbl_line.Top = lngLblLine_Offset
740         .TotalCost_lbl.Top = 0&
750         .TotalCost_lbl_line.Top = lngLblLine_Offset
760         .AccountName_lbl.Top = 0&
770         .AccountName_lbl_line.Top = lngLblLine_Offset
780         .Section(acPageHeader).Height = .page_header_height_line.Height
            '.Section(acPageHeader).Visible = False
790       End If
800     End With

EXITP:
810     Exit Sub

ERRH:
820     Select Case ERR.Number
        Case Else
830       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
840     End Select
850     Resume EXITP

End Sub

Private Sub Report_Close()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

910     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
920     Exit Sub

ERRH:
930     Select Case ERR.Number
        Case Else
940       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
950     End Select
960     Resume EXITP

End Sub
