VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptLocations_Detailed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptLocations_Detailed"

'VGC 11/11/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_Locations"
Private lngLbl_Offset As Long, lngLblLine_Offset As Long, lngTpp As Long
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

230         lngLblLine_Offset = .description_lbl_line.Top - .description_lbl.Top
240         lngLbl_Offset = .description_lbl.Top

250         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
260         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
270         .AsOf_lbl2.Caption = "As of " & Format(Date, "mmmm dd, yyyy")

280         .Caption = "Locations_Detailed_" & Format(Date, "yymmdd")

290       Else
300         Cancel = -1
310       End If
320     End With

330     If Cancel = -1 Then
340       Beep
350       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
360     End If

EXITP:
370     Exit Sub

ERRH:
380     Select Case ERR.Number
        Case Else
390       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
400     End Select
410     Resume EXITP

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
560         lngTmp01 = ((.AsOf_lbl2.Height + .assetdate_lbl.Height) + (6& * lngTpp))
570         .Section(acPageHeader).Height = lngTmp01
580         lngTmp01 = (.AsOf_lbl2.Height + (2& * lngTpp))
590         .description_lbl.Top = (lngTmp01 + lngLbl_Offset)
600         .description_lbl_line.Top = (.description_lbl.Top + lngLblLine_Offset)
610         .assetdate_lbl.Top = lngTmp01
620         .assetdate_lbl_line.Top = .description_lbl_line.Top
630         .transdate_lbl.Top = lngTmp01
640         .transdate_lbl_line.Top = .description_lbl_line.Top
650         .shareface_lbl.Top = (lngTmp01 + lngLbl_Offset)
660         .shareface_lbl_line.Top = .description_lbl_line.Top
670         .cost_lbl.Top = (lngTmp01 + lngLbl_Offset)
680         .cost_lbl_line.Top = .description_lbl_line.Top
690         .Header_lbl2.Visible = True
700         .AsOf_lbl2.Visible = True
710         .Header_hline05.Visible = True
            '.Section(acPageHeader).Visible = True
720       Else
730         .Header_lbl2.Visible = False
740         .AsOf_lbl2.Visible = False
750         .Header_hline05.Visible = False
760         .description_lbl.Top = lngLbl_Offset
770         .description_lbl_line.Top = (.description_lbl.Top + lngLblLine_Offset)
780         .assetdate_lbl.Top = 0&
790         .assetdate_lbl_line.Top = .description_lbl_line.Top
800         .transdate_lbl.Top = 0&
810         .transdate_lbl_line.Top = .description_lbl_line.Top
820         .shareface_lbl.Top = lngLbl_Offset
830         .shareface_lbl_line.Top = .description_lbl_line.Top
840         .cost_lbl.Top = lngLbl_Offset
850         .cost_lbl_line.Top = .description_lbl_line.Top
860         .Section(acPageHeader).Height = (.assetdate_lbl.Height + (2& * lngTpp))
            '.Section(acPageHeader).Visible = False
870       End If
880     End With

EXITP:
890     Exit Sub

ERRH:
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

EXITP:
1020    Exit Sub

ERRH:
1030    Select Case ERR.Number
        Case Else
1040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1050    End Select
1060    Resume EXITP

End Sub
